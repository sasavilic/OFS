#!/usr/bin/env python3
#
# OFS — LibreOffice Calc extension for OFS fiscal cash register (ofs.ba)
# Copyright (C) 2026 Fortunacommerc d.o.o.
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.
#

import concurrent.futures
from datetime import datetime as _datetime
import gzip
import ipaddress
import json
import logging
import logging.handlers
import os
import platform
import re as _re
import shutil
import socket
import subprocess
import sys
import urllib.request
import urllib.error

# ──────────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────────

VAT_LABEL_DEFAULT = "E"

# Latin → Cyrillic mapping for tax labels (visually identical, different codepoints)
_LATIN_TO_CYRILLIC = {
    "A": "\u0410",  # А
    "B": "\u0412",  # В
    "C": "\u0421",  # С
    "E": "\u0415",  # Е
    "H": "\u041D",  # Н
    "K": "\u041A",  # К
    "M": "\u041C",  # М
    "O": "\u041E",  # О
    "P": "\u0420",  # Р
    "T": "\u0422",  # Т
    "X": "\u0425",  # Х
}
_CYRILLIC_TO_LATIN = {v: k for k, v in _LATIN_TO_CYRILLIC.items()}

_REGISTER_CONFIG_NODE = "/com.fortunacommerc.ofs.Settings/Register"
REGISTER_CONFIG_DEFAULTS = {"mac": "", "host": "localhost", "port": 3566, "api_key": "", "pin": ""}

# Document property ključevi sa default vrijednostima
DOC_PROPS = {
    "OFS_Cell_BuyerId":       "",
    "OFS_Range_Buyer":        "",
    "OFS_Items_FirstRow":     "",
    "OFS_Items_LastRow":      "",
    "OFS_Col_ItemName":       "",
    "OFS_Col_ItemUOM":        "",
    "OFS_Col_ItemQty":        "",
    "OFS_Col_ItemPrice":      "",
    "OFS_Col_ItemLabel":      "",
    "OFS_Col_ItemBarcode":    "",
    "OFS_Cell_TotalBase":     "",
    "OFS_Cell_TotalVAT":      "",
    "OFS_Cell_TotalWithVAT":  "",
    "OFS_Cell_InvoiceNumber": "",
    "OFS_Cell_VATPercent":    "",
}

DOC_PROPS_REQUIRED = list(DOC_PROPS.keys())

INVOICE_TYPES = [
    ("Normalni (Normal)",   "Normal"),
    ("Proforma",            "Proforma"),
    ("Kopija (Copy)",       "Copy"),
    ("Trening (Training)",  "Training"),
    ("Avansni (Advance)",   "Advance"),
]
TRANSACTION_TYPES = [
    ("Prodaja (Sale)",   "Sale"),
    ("Povrat (Refund)",  "Refund"),
]
PAYMENT_TYPES = [
    ("Virman (WireTransfer)",       "WireTransfer"),
    ("Gotovina (Cash)",             "Cash"),
    ("Kartica (Card)",              "Card"),
    ("Ček (Check)",                 "Check"),
    ("Vaučer (Voucher)",            "Voucher"),
    ("Mobilni novac (MobileMoney)", "MobileMoney"),
    ("Ostalo (Other)",              "Other"),
]


# ──────────────────────────────────────────────────────────────────────────────
# Logging
# ──────────────────────────────────────────────────────────────────────────────

_EXT_NAME = "ofs"
_LOGGING_CONFIG_NODE = "/com.fortunacommerc.ofs.Settings/Logging"


def _setup_logger():
    if platform.system() == "Windows":
        log_dir = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), _EXT_NAME)
    else:
        log_dir = os.path.join(os.path.expanduser("~"), ".config", _EXT_NAME)
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(log_dir, "extension.log")
    logger = logging.getLogger(_EXT_NAME)
    if logger.handlers:
        return logger
    handler = logging.handlers.TimedRotatingFileHandler(
        log_file, when="W0", backupCount=0, encoding="utf-8")

    def _rotator(source, dest):
        with open(source, "rb") as f_in, gzip.open(dest, "wb") as f_out:
            shutil.copyfileobj(f_in, f_out)
        os.remove(source)

    handler.rotator = _rotator
    handler.namer = lambda name: name + ".gz"

    handler.setFormatter(logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s: %(message)s"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    return logger


log = _setup_logger()

log.info("Logger ready")

import uno
import unohelper
from com.sun.star.awt import XActionListener
from com.sun.star.frame import XDispatch, XDispatchProvider
from com.sun.star.lang import XServiceInfo, XInitialization

log.info("UNO imports loaded")

# Frame aktuelnog dokumenta — postavlja se u svakom entry pointu
_g_frame = None


def _apply_log_level(ctx):
    try:
        pv = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        pv.Name = "nodepath"
        pv.Value = _LOGGING_CONFIG_NODE
        cp = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx)
        node = cp.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (pv,))
        level_name = node.getByName("LogLevel") or "INFO"
        log.setLevel(getattr(logging, level_name.upper(), logging.INFO))
    except Exception as e:
        log.warning("Could not read LogLevel from registry: %s", e)


# ──────────────────────────────────────────────────────────────────────────────
# Cell reading helpers ("A1", "B13", ...)
# ──────────────────────────────────────────────────────────────────────────────

def _col_idx(letters):
    """'A' → 0,  'B' → 1,  'AA' → 26, ..."""
    result = 0
    for ch in letters.upper():
        result = result * 26 + (ord(ch) - ord('A') + 1)
    return result - 1


def _parse_addr(addr):
    """'B13' → (col_idx=1, row_idx=12) (both 0-based)."""
    m = _re.match(r'^([A-Za-z]+)(\d+)$', addr.strip())
    if not m:
        raise ValueError(f"Invalid cell address: {addr!r}")
    return _col_idx(m.group(1)), int(m.group(2)) - 1


def read_cell_string(sheet, addr):
    """Read string from cell by address. Returns calculated value (not formula)."""
    col, row = _parse_addr(addr)
    return sheet.getCellByPosition(col, row).getString().strip()


def read_cell_number(sheet, addr):
    """Read number from cell by address. Returns calculated value (not formula)."""
    col, row = _parse_addr(addr)
    return sheet.getCellByPosition(col, row).getValue()


def read_range_strings(sheet, range_addr):
    """Read all non-empty strings from a range like 'B5:B8' or 'B5:C8'."""
    m = _re.match(r'^([A-Za-z]+)(\d+):([A-Za-z]+)(\d+)$', range_addr.strip())
    if not m:
        return []
    col_s = _col_idx(m.group(1)); row_s = int(m.group(2)) - 1
    col_e = _col_idx(m.group(3)); row_e = int(m.group(4)) - 1
    result = []
    for r in range(row_s, row_e + 1):
        for c in range(col_s, col_e + 1):
            val = sheet.getCellByPosition(c, r).getString().strip()
            if val:
                result.append(val)
    return result


# ──────────────────────────────────────────────────────────────────────────────
# Document property helpers
# ──────────────────────────────────────────────────────────────────────────────

def get_doc_property(doc, name, default=""):
    """Read custom document property. Returns default if not found."""
    try:
        return doc.DocumentProperties.UserDefinedProperties.getPropertyValue(name)
    except Exception:
        return default


def set_doc_property(doc, name, value):
    """Set custom document property. Creates if not exists."""
    user_props = doc.DocumentProperties.UserDefinedProperties
    try:
        user_props.getPropertyValue(name)
        user_props.setPropertyValue(name, value)
    except Exception:
        user_props.addProperty(name, 64, value)  # 64 = REMOVEABLE


# ──────────────────────────────────────────────────────────────────────────────
# Register config (LO Configuration Registry)
# ──────────────────────────────────────────────────────────────────────────────

def load_register_config(ctx):
    """Read register settings from LO registry. Returns dict with defaults as fallback."""
    try:
        cp = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.configuration.ConfigurationProvider", ctx)
        pv = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        pv.Name = "nodepath"
        pv.Value = _REGISTER_CONFIG_NODE
        node = cp.createInstanceWithArguments(
            "com.sun.star.configuration.ConfigurationAccess", (pv,))
        return {
            "mac":     node.getByName("Mac"),
            "host":    node.getByName("Host"),
            "port":    node.getByName("Port"),
            "api_key": node.getByName("ApiKey"),
            "pin":     node.getByName("Pin"),
        }
    except Exception:
        return dict(REGISTER_CONFIG_DEFAULTS)


def save_register_config(ctx, register_config):
    """Save register settings to LO registry."""
    cp = ctx.ServiceManager.createInstanceWithContext(
        "com.sun.star.configuration.ConfigurationProvider", ctx)
    pv = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
    pv.Name = "nodepath"
    pv.Value = _REGISTER_CONFIG_NODE
    node = cp.createInstanceWithArguments(
        "com.sun.star.configuration.ConfigurationUpdateAccess", (pv,))
    node.replaceByName("Mac", register_config["mac"])
    node.replaceByName("Host", register_config["host"])
    node.replaceByName("Port", register_config["port"])
    node.replaceByName("ApiKey", register_config["api_key"])
    node.replaceByName("Pin", register_config["pin"])
    node.commitChanges()


# ──────────────────────────────────────────────────────────────────────────────
# Cell references — load settings from document properties
# ──────────────────────────────────────────────────────────────────────────────

def load_cell_references(doc):
    """Load all OFS_* document properties. Returns dict."""
    return {key: get_doc_property(doc, key, default) for key, default in DOC_PROPS.items()}


def validate_cell_references(cell_references):
    """Raise ValueError if required settings are missing."""
    missing = [k for k in DOC_PROPS_REQUIRED if not cell_references.get(k)]
    if missing:
        raise ValueError("Ćelije fakture nisu podešene.\n\nKoristite OFS > Podešavanja dokumenta.")


# ──────────────────────────────────────────────────────────────────────────────
# Read invoice data from document
# ──────────────────────────────────────────────────────────────────────────────

def read_invoice_data(doc, sheet, cell_references):
    """
    Read invoice items from sheet using cell_references settings.
    Returns dict with all data needed for dialog and API.
    """
    # ── BuyerId ────────────────────────────────────────────────────
    buyer_id = read_cell_string(sheet, cell_references["OFS_Cell_BuyerId"]) if cell_references["OFS_Cell_BuyerId"] else ""

    # ── Podaci o kupcu ─────────────────────────────────────────────
    rng = cell_references.get("OFS_Range_Buyer", "")
    buyer_lines = read_range_strings(sheet, rng) if rng else []

    # ── Čitaj stavke ───────────────────────────────────────────────
    first_row = int(cell_references["OFS_Items_FirstRow"])
    last_row  = int(cell_references["OFS_Items_LastRow"])
    col_name  = cell_references["OFS_Col_ItemName"]
    col_uom   = cell_references["OFS_Col_ItemUOM"]
    col_qty   = cell_references["OFS_Col_ItemQty"]
    col_price = cell_references["OFS_Col_ItemPrice"]
    col_label = cell_references["OFS_Col_ItemLabel"]
    col_barcode = cell_references["OFS_Col_ItemBarcode"]

    # Izvuci samo slovo kolone iz adrese (npr. "C13" → "C")
    def col_letter(addr):
        m = _re.match(r'^([A-Za-z]+)', addr.strip())
        return m.group(1) if m else addr

    col_name_l    = col_letter(col_name)
    col_uom_l     = col_letter(col_uom) if col_uom else ""
    col_qty_l     = col_letter(col_qty)
    col_price_l   = col_letter(col_price)
    col_label_l   = col_letter(col_label) if col_label else ""
    col_barcode_l = col_letter(col_barcode)

    items = []
    for row_num in range(first_row, last_row + 1):
        naziv = read_cell_string(sheet, f"{col_name_l}{row_num}")
        if not naziv:
            break

        kol    = read_cell_number(sheet, f"{col_qty_l}{row_num}")
        cijena = read_cell_number(sheet, f"{col_price_l}{row_num}")

        uom = ""
        if col_uom_l:
            uom = read_cell_string(sheet, f"{col_uom_l}{row_num}")

        label = VAT_LABEL_DEFAULT
        if col_label_l:
            val = read_cell_string(sheet, f"{col_label_l}{row_num}")
            if val:
                label = val

        barcode = read_cell_string(sheet, f"{col_barcode_l}{row_num}")
        if not barcode:
            raise ValueError(f"Stavka u redu {row_num} nema bar kod.")

        # API name format: "<naziv> /<jed.mjere>" ako jed.mjere postoji
        api_name = f"{naziv} /{uom}" if uom else naziv

        items.append({
            "naziv":     naziv,
            "api_name":  api_name,
            "uom":       uom,
            "kolicina":  kol,
            "cijena_sa": cijena,
            "label":     label,
            "barcode":   barcode,
        })

    ukupno_bez = round(read_cell_number(sheet, cell_references["OFS_Cell_TotalBase"]), 2)
    ukupno_pdv = round(read_cell_number(sheet, cell_references["OFS_Cell_TotalVAT"]), 2)
    ukupno_sa  = round(read_cell_number(sheet, cell_references["OFS_Cell_TotalWithVAT"]), 2)
    pdv_percent = read_cell_number(sheet, cell_references["OFS_Cell_VATPercent"])

    return {
        "buyer_id":     buyer_id,
        "buyer_lines":  buyer_lines,
        "items":        items,
        "ukupno_bez":   ukupno_bez,
        "ukupno_pdv":   ukupno_pdv,
        "ukupno_sa":    ukupno_sa,
        "pdv_percent":  pdv_percent,
    }


# ──────────────────────────────────────────────────────────────────────────────
# Write invoice number back to document
# ──────────────────────────────────────────────────────────────────────────────

def write_invoice_number(doc, sheet, cell_references, invoice_number):
    """Write invoiceNumber to cell defined by OFS_Cell_InvoiceNumber."""
    addr = cell_references["OFS_Cell_InvoiceNumber"]
    if not invoice_number:
        return
    col, row = _parse_addr(addr)
    sheet.getCellByPosition(col, row).setString(str(invoice_number))


# ──────────────────────────────────────────────────────────────────────────────
# Dialog helpers
# ──────────────────────────────────────────────────────────────────────────────

def format_amount_km(v):
    """Format amount in KM with dots and commas."""
    s = f"{v:,.2f}"  # e.g. "1,053.00"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"{s} KM"


def format_number(v, int_width, decimals):
    """Format number: dots for thousands, comma for decimals, right-align integer part."""
    s = f"{v:,.{decimals}f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    int_part, dec_part = s.split(",", 1)
    return f"{int_part:>{int_width}},{dec_part}"


def wrap_text(text, width):
    """Wrap text to specified width, returns list of lines."""
    lines = []
    remaining = text
    while remaining:
        if len(remaining) <= width:
            lines.append(remaining)
            break
        else:
            # Find last space within width
            space_idx = remaining.rfind(' ', 0, width)
            if space_idx == -1:
                # No space found, break at width
                lines.append(remaining[:width])
                remaining = remaining[width:].lstrip()
            else:
                lines.append(remaining[:space_idx])
                remaining = remaining[space_idx+1:]
    return lines


def show_msgbox(ctx, message, title="OFS", msgtype="info"):
    """Display LibreOffice message box."""
    smgr = ctx.ServiceManager
    frame = _g_frame

    type_map = {"info": 1, "error": 3, "question": 2}
    btn_map  = {"info": 1, "error": 1, "question": 3}

    toolkit = smgr.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
    parent  = frame.getContainerWindow() if frame else None

    msgbox = toolkit.createMessageBox(
        parent,
        type_map.get(msgtype, 1),
        btn_map.get(msgtype, 1),
        title,
        message,
    )
    result = msgbox.execute()
    msgbox.dispose()
    return result


def _load_dialog(ctx, dialog_name):
    """Load XDL dialog from extension package, center, return dialog."""
    dp = ctx.ServiceManager.createInstanceWithContext("com.sun.star.awt.DialogProvider", ctx)
    try:
        dialog = dp.createDialog(f"vnd.sun.star.extension://com.fortunacommerc.ofs/dialogs/{dialog_name}.xdl")
        parent = _g_frame.getContainerWindow()
        toolkit = parent.getToolkit()
        dialog.createPeer(toolkit, parent)
        parent_size = parent.getSize()
        dialog_size = dialog.getSize()
        x = max(0, (parent_size.Width  - dialog_size.Width)  // 2)
        y = max(0, (parent_size.Height - dialog_size.Height) // 2)
        dialog.setPosSize(x, y, dialog_size.Width, dialog_size.Height,
                          uno.getConstantByName("com.sun.star.awt.PosSize.POS"))
        return dialog
    except Exception as e:
        # Print the full exception chain
        log.error(f"Exception: {e}")
        if hasattr(e, 'value'):
            inner = e.value
            log.error(f"Inner: {inner}")
            if hasattr(inner, 'Message'):
                log.error(f"Message: {inner.Message}")
            if hasattr(inner, 'TargetException'):
                log.error(f"Target: {inner.TargetException}")
        raise e


def show_success_dialog(ctx, response):
    """Display details of successfully fiscalized invoice."""
    invoice_number = response.get("invoiceNumber", "—")
    total          = response.get("totalAmount", 0)
    inv_type       = response.get("invoiceType", "—")
    counter        = response.get("invoiceCounter", "—")
    ver_url        = response.get("verificationUrl", "")

    prefix = ""
    if inv_type == "Training":
        prefix = "[ OBUČNI RAČUN ]\n\n"

    msg = (
        f"{prefix}"
        f"Račun uspješno fiskalizovan!\n\n"
        f"Broj računa : {invoice_number}\n"
        f"Tip         : {inv_type}\n"
        f"Brojač      : {counter}\n"
        f"Ukupno      : {format_amount_km(total)}\n"
    )
    if ver_url:
        msg += f"\nVerifikacija:\n{ver_url}"

    show_msgbox(ctx, msg, "OFS — Uspjeh", "info")


# ──────────────────────────────────────────────────────────────────────────────
# Dialog: Fiscalization confirmation
# ──────────────────────────────────────────────────────────────────────────────

SEP = "─" * 60

# Receipt column layout (60 chars total):
#   cijena: 14 chars | " x " | kol: 10 chars | " = " | iznos: 14 chars  → 44 chars
#   Left side (Naziv): 60 - 44 = 16 chars
# Totals: label 46 chars | iznos 14 chars = 60 chars
_W = 60
_AMT = f"{'Cijena':>14}   {'Kol.':>10}   {'Ukupno':>14}"   # 44 chars
_LPAD = _W - len(_AMT)                                       # 16

def build_items_text(invoice_data):
    """Build text for displaying items in dialog (60 char receipt width)."""
    lines = []

    # ── Buyer ID ──────────────────────────────────────────────────
    lines.extend(wrap_text(f"Buyer ID: {invoice_data['buyer_id'] or '(nije pronađen)'}", _W))
    lines.append(SEP)

    # ── Podaci o kupcu ─────────────────────────────────────────────
    for l in invoice_data["buyer_lines"]:
        lines.extend(wrap_text(l, _W))
    lines.append(SEP)

    # ── Stavke: header ──────────────────────────────────────────────
    lines.append(f"{'Naziv':<{_LPAD}}{_AMT}")

    # ── Stavke: rows ────────────────────────────────────────────────
    for item in invoice_data["items"]:
        # Line 1: [barcode] naziv /uom (label)
        naziv_display = item["naziv"]
        if item.get("uom"):
            naziv_display += f" /{item['uom']}"
        naziv_display += f" ({item.get('label', 'E')})"
        if item.get("barcode"):
            naziv_display = f"{item['barcode']} {naziv_display}"
        lines.extend(wrap_text(naziv_display, _W))

        # Line 2: cijena x kolicina = ukupno — right-aligned to 60
        cijena = item["cijena_sa"]
        kol    = item["kolicina"]
        iznos  = round(kol * cijena, 2)
        line2  = f"{format_number(cijena, 9, 4)} x {format_number(kol, 7, 2)} = {format_number(iznos, 11, 2)}"
        lines.append(f"{line2:>{_W}}")

    lines.append(SEP)

    # ── Totali: label 46 chars + iznos 14 chars = 60 ───────────────
    _TL = _W - 14   # 46
    pdv_pct = invoice_data["pdv_percent"]
    pdv_label = f"PDV {pdv_pct:g}%:"
    lines.append(f"{'Osnovica (bez PDV):':<{_TL}}{format_number(invoice_data['ukupno_bez'], 11, 2)}")
    lines.append(f"{pdv_label:<{_TL}}{format_number(invoice_data['ukupno_pdv'], 11, 2)}")
    lines.append(f"{'UKUPNO SA PDV:':<{_TL}}{format_number(invoice_data['ukupno_sa'], 11, 2)}")

    return "\n".join(lines)


def _parse_ref_dt(text):
    """Parse ref datum from DD.MM.YYYY HH:MM:SS to ISO format.

    Returns ISO string on success, empty string if input is empty,
    or raises ValueError on invalid format.
    """
    if not text:
        return ""
    dt = _datetime.strptime(text, "%d.%m.%Y %H:%M:%S")
    return dt.strftime("%Y-%m-%dT%H:%M:%S")


def _read_dialog_options(dialog):
    """Read current selections from confirm dialog. Returns dict."""
    tip_idx = dialog.getControl("lbTip").getSelectedItemPos()
    txn_idx = dialog.getControl("lbTxn").getSelectedItemPos()
    pay_idx = dialog.getControl("lbPay").getSelectedItemPos()
    ref_num = dialog.getControl("tfRef").getText().strip()
    ref_dt  = dialog.getControl("tfRefDt").getText().strip()

    return {
        "invoice_type":     INVOICE_TYPES[tip_idx][1],
        "transaction_type": TRANSACTION_TYPES[txn_idx][1],
        "payment_type":     PAYMENT_TYPES[pay_idx][1],
        "referent_number":  ref_num,
        "referent_dt":      ref_dt,
    }


def _show_json_preview(ctx, json_str):
    """Display JSON preview dialog with read-only text area."""
    dialog = _load_dialog(ctx, "JsonPreview")
    dialog.getControl("txtJson").setText(json_str)
    dialog.execute()
    dialog.dispose()


def show_confirm_dialog(ctx, invoice_data, register_config):
    """
    Display dialog with items and options for fiscalization.
    Returns dict or None if user cancelled.
    """
    dialog = _load_dialog(ctx, "FiscalizationConfirm")

    txt_stavke = dialog.getControl("txtStavke")
    txt_stavke.setText(build_items_text(invoice_data))
    dialog.getControl("lblUrl").getModel().setPropertyValue(
        "Label", f"Kasa: {register_config['host']}:{register_config['port']}")

    lb_tip = dialog.getControl("lbTip")
    lb_tip.addItems(tuple(d for d, _ in INVOICE_TYPES), 0)
    lb_tip.selectItemPos(0, True)

    lb_txn = dialog.getControl("lbTxn")
    lb_txn.addItems(tuple(d for d, _ in TRANSACTION_TYPES), 0)
    lb_txn.selectItemPos(0, True)

    lb_pay = dialog.getControl("lbPay")
    lb_pay.addItems(tuple(d for d, _ in PAYMENT_TYPES), 0)
    lb_pay.selectItemPos(0, True)

    class _PreviewListener(unohelper.Base, XActionListener):
        def actionPerformed(self, ev):
            opts = _read_dialog_options(dialog)
            try:
                opts["referent_dt"] = _parse_ref_dt(opts["referent_dt"])
            except ValueError:
                show_msgbox(ctx, "Neispravan format datuma.\n"
                            "Očekivani format: DD.MM.YYYY HH:MM:SS",
                            "OFS", "error")
                return
            req = build_esir_request(invoice_data, opts)
            _show_json_preview(ctx, json.dumps(req, ensure_ascii=False, indent=2))
        def disposing(self, ev):
            pass

    dialog.getControl("btnPreview").addActionListener(_PreviewListener())

    while True:
        if dialog.execute() != 1:
            dialog.dispose()
            return None

        options = _read_dialog_options(dialog)
        try:
            options["referent_dt"] = _parse_ref_dt(options["referent_dt"])
        except ValueError:
            show_msgbox(ctx, "Neispravan format datuma.\n"
                        "Očekivani format: DD.MM.YYYY HH:MM:SS",
                        "OFS", "error")
            continue

        dialog.dispose()
        return options


# ──────────────────────────────────────────────────────────────────────────────
# Send to ESIR API
# ──────────────────────────────────────────────────────────────────────────────

def build_esir_request(invoice_data, options):
    """Build JSON request for ESIR API."""
    total = round(invoice_data["ukupno_sa"], 2)

    items = []
    for item in invoice_data["items"]:
        items.append({
            "gtin":        item.get("barcode", ""),
            "name":        item["api_name"],
            "quantity":    item["kolicina"],
            "unitPrice":   round(item["cijena_sa"], 4),
            "labels":      [item.get("label", VAT_LABEL_DEFAULT)],
            "totalAmount": round(item["kolicina"] * item["cijena_sa"], 2),
        })

    req = {
        "invoiceRequest": {
            "buyerId":          invoice_data.get("buyer_id", ""),
            "invoiceType":      options["invoice_type"],
            "transactionType":  options["transaction_type"],
            "payment": [
                {
                    "paymentType": options["payment_type"],
                    "amount":      total,
                }
            ],
            "items": items,
        }
    }

    ref_num = options.get("referent_number", "")
    ref_dt  = options.get("referent_dt", "")
    if ref_num:
        req["invoiceRequest"]["referentDocumentNumber"] = ref_num
    if ref_dt:
        req["invoiceRequest"]["referentDocumentDT"] = ref_dt

    return req


def send_to_esir_api(request_body, register_config):
    """Send HTTP POST to ESIR API. Returns (success, response_dict)."""
    endpoint = f"http://{register_config['host']}:{register_config['port']}/api/invoices"
    payload = json.dumps(request_body, ensure_ascii=False).encode("utf-8")

    headers = {
        "Content-Type":   "application/json; charset=utf-8",
        "Content-Length": str(len(payload)),
        "Accept":         "application/json",
    }
    if register_config.get("api_key"):
        headers["Authorization"] = f"Bearer {register_config['api_key']}"

    req = urllib.request.Request(
        endpoint,
        data=payload,
        headers=headers,
        method="POST",
    )
    log.info(f"ESIR request: POST {endpoint} body={payload.decode('utf-8')}")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8")
            log.info(f"ESIR response: HTTP 200 body={body}")
            return True, json.loads(body)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8")
        log.info(f"ESIR response: HTTP {e.code} body={body}")
        try:
            return False, json.loads(body)
        except Exception:
            return False, {"error": str(e), "body": body}
    except urllib.error.URLError as e:
        return False, {"error": str(e.reason)}
    except Exception as e:
        return False, {"error": str(e)}


# ──────────────────────────────────────────────────────────────────────────────
# ESIR Initialization
# ──────────────────────────────────────────────────────────────────────────────

def esir_check_attention(register_config):
    """GET /api/attention — check register availability. Returns (success, error_msg)."""
    url = f"http://{register_config['host']}:{register_config['port']}/api/attention"
    headers = {"Accept": "application/json"}
    if register_config.get("api_key"):
        headers["Authorization"] = f"Bearer {register_config['api_key']}"

    req = urllib.request.Request(url, headers=headers, method="GET")
    log.info(f"ESIR request: GET {url}")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            log.info(f"ESIR response: HTTP {resp.status} (attention OK)")
            return True, None
    except urllib.error.HTTPError as e:
        log.info(f"ESIR response: HTTP {e.code} (attention)")
        return False, f"Kasa nije dostupna: HTTP {e.code}"
    except urllib.error.URLError as e:
        return False, f"Kasa nije dostupna: {e.reason}"
    except Exception as e:
        return False, f"Kasa nije dostupna: {e}"


def _probe_attention(host, port, api_key):
    """Lightweight HTTP probe of /api/attention. Returns HTTP status code or -1 on error."""
    url = f"http://{host}:{port}/api/attention"
    headers = {"Accept": "application/json"}
    if api_key:
        headers["Authorization"] = f"Bearer {api_key}"
    req = urllib.request.Request(url, headers=headers, method="GET")
    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            return resp.status
    except urllib.error.HTTPError as e:
        return e.code
    except Exception:
        return -1


def _verify_scan_results(found_hosts, port, api_key):
    """Probe /api/attention on all found hosts in parallel.

    Returns list of (ip, subnet, http_status) tuples.
    """
    if not found_hosts:
        return []

    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=10) as pool:
        future_map = {
            pool.submit(_probe_attention, ip, port, api_key): (ip, subnet)
            for ip, subnet in found_hosts
        }
        for future in concurrent.futures.as_completed(future_map):
            ip, subnet = future_map[future]
            try:
                status = future.result()
            except Exception:
                status = -1
            results.append((ip, subnet, status))

    results.sort(key=lambda x: ipaddress.IPv4Address(x[0]))
    return results


def _build_label_map(status_data):
    """Build Latin→register label map from currentTaxRates.

    If the register uses Cyrillic labels (e.g. 'Е' instead of 'E'),
    returns a dict mapping Latin letters to the Cyrillic equivalents
    that the register expects. If labels are already Latin, returns
    an empty dict (no translation needed).
    """
    label_map = {}
    tax_rates = status_data.get("currentTaxRates", {})
    for cat in tax_rates.get("taxCategories", []):
        for tr in cat.get("taxRates", []):
            reg_label = tr.get("label", "")
            if not reg_label:
                continue
            # Check if this register label is Cyrillic
            latin_equiv = _CYRILLIC_TO_LATIN.get(reg_label)
            if latin_equiv:
                label_map[latin_equiv] = reg_label
                log.debug(f"label map: {latin_equiv!r} -> {reg_label!r} (Cyrillic)")
            else:
                # Register uses Latin — map to itself (no-op, but explicit)
                log.debug(f"label map: {reg_label!r} is Latin (no translation)")
    if label_map:
        log.info(f"Register uses Cyrillic tax labels: {label_map}")
    return label_map


def esir_check_status(register_config):
    """GET /api/status — check SE status. Returns (success, needs_pin, label_map, error_msg)."""
    url = f"http://{register_config['host']}:{register_config['port']}/api/status"
    headers = {"Accept": "application/json"}
    if register_config.get("api_key"):
        headers["Authorization"] = f"Bearer {register_config['api_key']}"

    req = urllib.request.Request(url, headers=headers, method="GET")
    log.info(f"ESIR request: GET {url}")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            body = resp.read().decode("utf-8")
            log.info(f"ESIR response: HTTP {resp.status} body={body}")
            data = json.loads(body)
            gsc = data.get("gsc", [])
            needs_pin = "1300" in gsc or "1500" in gsc
            label_map = _build_label_map(data)
            return True, needs_pin, label_map, None
    except urllib.error.HTTPError as e:
        body_err = ""
        try:
            body_err = e.read().decode("utf-8")
        except Exception:
            pass
        log.info(f"ESIR response: HTTP {e.code} body={body_err}")
        return False, False, {}, f"Status provjera neuspjela: HTTP {e.code}"
    except urllib.error.URLError as e:
        return False, False, {}, f"Status provjera neuspjela: {e.reason}"
    except Exception as e:
        return False, False, {}, f"Status provjera neuspjela: {e}"


def esir_send_pin(register_config):
    """POST /api/pin — unlock SE with PIN code. Returns (success, error_msg)."""
    pin = register_config.get("pin", "")
    if not pin:
        return False, "PIN kod nije podešen u konfiguraciji kase."

    url = f"http://{register_config['host']}:{register_config['port']}/api/pin"
    payload = pin.encode("utf-8")
    headers = {
        "Content-Type": "text/plain; charset=utf-8",
        "Accept": "application/json",
    }
    if register_config.get("api_key"):
        headers["Authorization"] = f"Bearer {register_config['api_key']}"

    req = urllib.request.Request(url, data=payload, headers=headers, method="POST")
    log.info(f"ESIR request: POST {url} body=(PIN omitted)")
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            log.info(f"ESIR response: HTTP {resp.status} (pin OK)")
            return True, None
    except urllib.error.HTTPError as e:
        hints = {
            "1300": "bezbjednosni element nije prisutan",
            "2400": "LPFR nije spreman",
            "2800": "neispravan PIN",
            "2806": "neispravan PIN",
        }
        body = ""
        try:
            body = e.read().decode("utf-8")
        except Exception:
            pass
        log.info(f"ESIR response: HTTP {e.code} body={body}")
        hint = ""
        for code, desc in hints.items():
            if code in body:
                hint = f" ({desc})"
                break
        return False, f"PIN greška: HTTP {e.code}{hint}"
    except urllib.error.URLError as e:
        return False, f"PIN slanje neuspjelo: {e.reason}"
    except Exception as e:
        return False, f"PIN slanje neuspjelo: {e}"


def esir_init(register_config):
    """Initialize register: attention -> status -> pin (if needed). Returns (success, label_map, error_msg)."""
    ok, err = esir_check_attention(register_config)
    if not ok:
        log.error(f"esir_init: attention failed: {err}")
        return False, {}, err
    log.debug("esir_init: attention OK")

    ok, needs_pin, label_map, err = esir_check_status(register_config)
    if not ok:
        log.error(f"esir_init: status failed: {err}")
        return False, {}, err
    log.debug("esir_init: status OK")

    if needs_pin:
        ok, err = esir_send_pin(register_config)
        if not ok:
            log.error(f"esir_init: pin failed: {err}")
            return False, {}, err
        log.debug("esir_init: pin OK")

    return True, label_map, None


# ──────────────────────────────────────────────────────────────────────────────
# Document settings — dialog with 15 fields
# ──────────────────────────────────────────────────────────────────────────────

def show_document_settings(ctx, doc):
    """
    OFS > Podešavanja dokumenta.
    Dialog sa 15 polja za konfigurisanje referenci ćelija.
    """
    global _g_frame
    log.info("show_document_settings: started")
    try:
        _g_frame = doc.getCurrentController().getFrame()
    except Exception:
        pass

    cell_references = load_cell_references(doc)
    dialog = _load_dialog(ctx, "DocumentSettings")

    for prop_key in DOC_PROPS:
        try:
            dialog.getControl(f"tf_{prop_key}").setText(
                cell_references.get(prop_key, ""))
        except Exception:
            pass

    if dialog.execute() != 1:
        dialog.dispose()
        return

    for prop_key in DOC_PROPS:
        try:
            val = dialog.getControl(f"tf_{prop_key}").getText().strip()
            set_doc_property(doc, prop_key, val)
        except Exception:
            pass

    dialog.dispose()
    doc.setModified(True)

    show_msgbox(ctx,
        "Podešavanja ćelija sačuvana.\n"
        "Sačuvajte dokument (Ctrl+S) da biste zadržali promjene.",
        "OFS — Podešavanja", "info")


# ──────────────────────────────────────────────────────────────────────────────
# Network scanning
# ──────────────────────────────────────────────────────────────────────────────

def _get_local_subnets():
    """Detect local IPv4 subnets. Returns list of ipaddress.IPv4Network."""
    subnets = []
    try:
        if sys.platform == "linux":
            out = subprocess.check_output(
                ["ip", "-4", "-o", "addr", "show"], timeout=5
            ).decode("utf-8", errors="replace")
            for line in out.splitlines():
                # e.g. "2: eth0    inet 192.168.1.5/24 brd ..."
                m = _re.search(r'inet\s+(\d+\.\d+\.\d+\.\d+)/(\d+)', line)
                if m:
                    ip, prefix = m.group(1), int(m.group(2))
                    net = ipaddress.IPv4Network(f"{ip}/{prefix}", strict=False)
                    subnets.append(net)
        elif sys.platform == "win32":
            out = subprocess.check_output(
                ["powershell", "-Command",
                 "Get-NetIPAddress -AddressFamily IPv4 "
                 "| Select IPAddress,PrefixLength "
                 "| ConvertTo-Csv -NoTypeInformation"],
                timeout=5
            ).decode("utf-8", errors="replace")
            for line in out.splitlines()[1:]:  # skip header
                parts = line.strip().strip('"').split('","')
                if len(parts) == 2:
                    ip, prefix = parts[0], int(parts[1])
                    net = ipaddress.IPv4Network(f"{ip}/{prefix}", strict=False)
                    subnets.append(net)
    except Exception as e:
        log.warning("_get_local_subnets: command failed: %s", e)

    if not subnets:
        # Fallback: use socket to get local IPs with assumed /24
        try:
            for info in socket.getaddrinfo(socket.gethostname(), None, socket.AF_INET):
                ip = info[4][0]
                net = ipaddress.IPv4Network(f"{ip}/24", strict=False)
                subnets.append(net)
        except Exception as e:
            log.warning("_get_local_subnets: fallback failed: %s", e)

    # Filter out loopback and link-local
    filtered = []
    for net in subnets:
        if net.overlaps(ipaddress.IPv4Network("127.0.0.0/8")):
            continue
        if net.overlaps(ipaddress.IPv4Network("169.254.0.0/16")):
            continue
        filtered.append(net)
    return filtered


def _scan_port(host, port, timeout=0.3):
    """Check if TCP port is open on host. Returns True if open."""
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(timeout)
    try:
        return sock.connect_ex((str(host), port)) == 0
    except Exception:
        return False
    finally:
        sock.close()


def _scan_network(port):
    """Scan local subnets for hosts with open TCP port. Returns list of (ip_str, subnet_str)."""
    subnets = _get_local_subnets()
    log.info("_scan_network: port=%d, subnets=%s", port, subnets)
    if not subnets:
        return []

    tasks = []
    for net in subnets:
        for host in net.hosts():
            tasks.append((str(host), str(net)))

    results = []
    with concurrent.futures.ThreadPoolExecutor(max_workers=50) as pool:
        future_map = {
            pool.submit(_scan_port, ip, port): (ip, subnet)
            for ip, subnet in tasks
        }
        for future in concurrent.futures.as_completed(future_map):
            ip, subnet = future_map[future]
            try:
                if future.result():
                    results.append((ip, subnet))
            except Exception:
                pass

    results.sort(key=lambda x: ipaddress.IPv4Address(x[0]))
    log.info("_scan_network: found %d hosts", len(results))
    return results


def _get_arp_table():
    """Return dict mapping IP → MAC address from the OS ARP cache."""
    arp = {}
    try:
        if sys.platform.startswith("linux"):
            with open("/proc/net/arp", "r") as f:
                for line in f:
                    parts = line.split()
                    if len(parts) >= 4 and parts[0] != "IP":
                        ip, mac = parts[0], parts[3]
                        if mac != "00:00:00:00:00:00":
                            arp[ip] = mac
        elif sys.platform == "win32":
            import subprocess
            out = subprocess.check_output(["arp", "-a"], text=True,
                                          timeout=5)
            for line in out.splitlines():
                m = _re.match(r"\s+([\d.]+)\s+([\da-fA-F-]{17})\s+", line)
                if m:
                    ip, mac = m.group(1), m.group(2).replace("-", ":")
                    arp[ip] = mac
    except Exception:
        log.debug("_get_arp_table: failed to read ARP cache", exc_info=True)
    return arp


# ──────────────────────────────────────────────────────────────────────────────
# Register settings — system configuration
# ──────────────────────────────────────────────────────────────────────────────

def show_kasa_settings(ctx, doc=None):
    """
    OFS > Podešavanje kase.
    Dialog sa 4 polja: host, port, API ključ, PIN.
    """
    global _g_frame
    log.info("show_kasa_settings: started")
    if doc is not None:
        try:
            _g_frame = doc.getCurrentController().getFrame()
        except Exception:
            pass

    register_config = load_register_config(ctx)
    dialog = _load_dialog(ctx, "RegisterSettings")

    dialog.getControl("tf_mac").setText(str(register_config["mac"]))
    dialog.getControl("tf_host").setText(str(register_config["host"]))
    dialog.getControl("tf_port").setText(str(register_config["port"]))
    dialog.getControl("tf_api_key").setText(str(register_config["api_key"]))
    dialog.getControl("tf_pin").setText(str(register_config["pin"]))

    class _TestListener(unohelper.Base, XActionListener):
        def actionPerformed(self, ev):
            port_str = dialog.getControl("tf_port").getText().strip()
            try:
                port_val = int(port_str)
            except ValueError:
                show_msgbox(ctx,
                    f"Neispravan port: {port_str!r}\nPort mora biti cijeli broj.",
                    "OFS — Greška", "error")
                return
            temp = {
                "mac":     dialog.getControl("tf_mac").getText().strip(),
                "host":    dialog.getControl("tf_host").getText().strip(),
                "port":    port_val,
                "api_key": dialog.getControl("tf_api_key").getText().strip(),
                "pin":     dialog.getControl("tf_pin").getText().strip(),
            }
            ok, _lm, err = esir_init(temp)
            if ok:
                show_msgbox(ctx, "Kasa je dostupna i spremna.", "OFS — Test", "info")
            else:
                show_msgbox(ctx,
                    f"Greška:\n\n{err}\n\nProvjerite podešavanja i da li je ESIR pokrenut\n"
                    f"na {temp['host']}:{temp['port']}.", "OFS — Test", "error")
        def disposing(self, ev):
            pass

    dialog.getControl("btnTest").addActionListener(_TestListener())

    _scan_found = [None]  # set by _ScanListener when single device auto-configured

    class _ScanListener(unohelper.Base, XActionListener):
        def actionPerformed(self, ev):
            port_str = dialog.getControl("tf_port").getText().strip()
            try:
                port_val = int(port_str)
            except ValueError:
                show_msgbox(ctx,
                    f"Neispravan port: {port_str!r}\nPort mora biti cijeli broj.",
                    "OFS — Greška", "error")
                return
            found = _scan_network(port_val)
            if not found:
                show_msgbox(ctx,
                    f"Nije pronadjen nijedan uredjaj na portu {port_val}.",
                    "OFS — Skeniranje mreze", "info")
                return
            api_key = dialog.getControl("tf_api_key").getText().strip()
            verified = _verify_scan_results(found, port_val, api_key)
            arp = _get_arp_table()
            mac_filter = dialog.getControl("tf_mac").getText().strip().lower()
            if mac_filter:
                verified = [
                    (ip, subnet, status)
                    for ip, subnet, status in verified
                    if ip in arp and arp[ip].lower() == mac_filter
                ]
                if not verified:
                    show_msgbox(ctx,
                        f"Nije pronadjen uredjaj sa MAC adresom {mac_filter}.",
                        "OFS — Skeniranje mreze", "info")
                    return
            ok_hosts = [(ip, subnet) for ip, subnet, status in verified if status == 200]
            if len(ok_hosts) == 1:
                found_ip = ok_hosts[0][0]
                dialog.getControl("tf_host").setText(found_ip)
                if found_ip in arp:
                    dialog.getControl("tf_mac").setText(arp[found_ip])
                    mac_info = f"  [{arp[found_ip]}]"
                else:
                    mac_info = ""
                _scan_found[0] = f"{found_ip}{mac_info}:{port_val}"
                show_msgbox(ctx,
                    f"Pronadjena fiskalna kasa na {found_ip}{mac_info}.\n\n"
                    f"Polja Host i MAC adresa su azurirana.",
                    "OFS — Skeniranje mreze", "info")
            else:
                scan_dialog = _load_dialog(ctx, "ScanResults")
                lines = [f"Pronadjeni uredjaji na portu {port_val}:", ""]
                for ip, subnet, status in verified:
                    if status == 200:
                        annotation = "  —  API kljuc prihvacen"
                    elif status == 401:
                        annotation = "  —  Pogresan API kljuc"
                    elif status == 404:
                        annotation = "  —  Moguce fiskalna kasa (nepoznat API)"
                    elif status == -1:
                        annotation = ""
                    else:
                        annotation = f"  —  HTTP {status}"
                    mac_part = f"  [{arp[ip]}]" if ip in arp else ""
                    lines.append(f"  {ip}{mac_part}  ({subnet}){annotation}")
                scan_dialog.getControl("txtResults").setText("\n".join(lines))
                scan_dialog.execute()
                scan_dialog.dispose()
        def disposing(self, ev):
            pass

    dialog.getControl("btnScan").addActionListener(_ScanListener())

    if dialog.execute() != 1:
        dialog.dispose()
        return

    host_str    = dialog.getControl("tf_host").getText().strip()
    port_str    = dialog.getControl("tf_port").getText().strip()
    api_key_str = dialog.getControl("tf_api_key").getText().strip()
    pin_str     = dialog.getControl("tf_pin").getText().strip()

    try:
        port_val = int(port_str)
    except ValueError:
        dialog.dispose()
        show_msgbox(ctx,
            f"Neispravan port: {port_str!r}\nPort mora biti cijeli broj.",
            "OFS — Greška", "error")
        return

    if not host_str or not api_key_str or not pin_str:
        dialog.dispose()
        show_msgbox(ctx, "Sva polja su obavezna.", "OFS — Greška", "error")
        return

    if len(pin_str) != 4:
        dialog.dispose()
        show_msgbox(ctx, f"PIN kod mora imati tačno 4 znaka.", "OFS — Greška", "error")
        return

    mac_str     = dialog.getControl("tf_mac").getText().strip()

    new_config = {
        "mac":     mac_str,
        "host":    host_str,
        "port":    port_val,
        "api_key": api_key_str,
        "pin":     pin_str,
    }
    dialog.dispose()

    try:
        save_register_config(ctx, new_config)
    except Exception as e:
        show_msgbox(ctx, f"Greška pri snimanju:\n\n{e}", "OFS — Greška", "error")
        return

    if _scan_found[0]:
        show_msgbox(ctx,
            f"Pronadjena fiskalna kasa na {_scan_found[0]}.\n\n"
            f"Podesavanja su sacuvana.",
            "OFS — Skeniranje mreze", "info")
    else:
        show_msgbox(ctx,
            f"Podešavanja kase sačuvana.\n\nHost: {new_config['host']}\nPort: {new_config['port']}",
            "OFS — Podešavanja", "info")


# ──────────────────────────────────────────────────────────────────────────────
# Main entry point
# ──────────────────────────────────────────────────────────────────────────────

def send_to_ofs(ctx, doc):
    """
    OFS > Fiskalizacija — glavni entry point.
    """
    global _g_frame
    log.info("send_to_ofs: started")
    try:
        _g_frame = doc.getCurrentController().getFrame()
    except Exception:
        pass

    # ── Dohvati aktivni sheet ──────────────────────────────────────
    try:
        sheet = doc.Sheets.getByIndex(0)
    except Exception as e:
        log.error(f"send_to_ofs: pristup dokumentu: {e}")
        show_msgbox(ctx, f"Greška: Nije moguće pristupiti dokumentu.\n\n{e}", "OFS — Greška", "error")
        return

    # ── Load register config ──────────────────────────────────────
    register_config = load_register_config(ctx)
    log.debug(f"register config: {register_config['host']}:{register_config['port']}")

    # ── Load and validate cell references ──────────────────────────
    try:
        cell_references = load_cell_references(doc)
        validate_cell_references(cell_references)
    except ValueError as e:
        log.error(f"send_to_ofs: cell references error: {e}")
        show_msgbox(ctx, str(e), "OFS — Greška", "error")
        return

    # ── Read invoice data ──────────────────────────────────────────
    try:
        invoice_data = read_invoice_data(doc, sheet, cell_references)
    except Exception as e:
        log.error(f"send_to_ofs: čitanje dokumenta: {e}")
        show_msgbox(ctx, f"Greška pri čitanju dokumenta:\n\n{e}", "OFS — Greška", "error")
        return

    log.info(f"invoice items: {len(invoice_data['items'])}")

    if not invoice_data["items"]:
        show_msgbox(ctx,
            "Nema stavki za slanje.\n\n"
            "Provjerite da su reference ćelija ispravno podešene\n"
            "u OFS > Podešavanja dokumenta.",
            "OFS — Upozorenje", "info")
        return

    # ── Confirmation dialog ───────────────────────────────────────
    try:
        options = show_confirm_dialog(ctx, invoice_data, register_config)
    except Exception as e:
        log.error(f"send_to_ofs: dialog greška: {e}")
        show_msgbox(ctx, f"Greška pri prikazivanju dijaloga:\n\n{e}", "OFS — Greška", "error")
        return

    if options is None:
        return  # User cancelled

    # ── ESIR initialization ────────────────────────────────────────
    init_ok, label_map, init_err = esir_init(register_config)
    if not init_ok:
        show_msgbox(ctx,
            f"Greška pri inicijalizaciji kase:\n\n{init_err}\n\n"
            f"Provjerite podešavanja kase (OFS > Podešavanje kase)\n"
            f"i da li je ESIR pokrenut na {register_config['host']}:{register_config['port']}.",
            "OFS — Greška", "error")
        return

    # ── Translate tax labels if register uses Cyrillic ─────────────
    if label_map:
        for item in invoice_data["items"]:
            orig = item.get("label", VAT_LABEL_DEFAULT)
            item["label"] = label_map.get(orig, orig)

    # ── Izgradi i pošalji request ──────────────────────────────────
    try:
        request_body = build_esir_request(invoice_data, options)
    except Exception as e:
        log.error(f"send_to_ofs: build request greška: {e}")
        show_msgbox(ctx, f"Greška pri pripremi podataka:\n\n{e}", "OFS — Greška", "error")
        return

    log.debug(f"ESIR request: {json.dumps(request_body)}")
    success, response = send_to_esir_api(request_body, register_config)

    if success:
        invoice_number = response.get("invoiceNumber", "")
        log.info(f"ESIR success: invoice_number={invoice_number}")
        show_success_dialog(ctx, response)
        # Write invoice number to cell
        if invoice_number:
            try:
                write_invoice_number(doc, sheet, cell_references, invoice_number)
            except Exception:
                pass  # Don't abort if write fails
    else:
        error_msg = response.get("error", json.dumps(response, ensure_ascii=False))
        log.error(f"ESIR error: {error_msg}")
        show_msgbox(ctx,
            f"Greška pri slanju na fiskalnu kasu:\n\n{error_msg}\n\n"
            f"Provjerite podešavanja kase (OFS > Podešavanje kase)\n"
            f"i da li je ESIR pokrenut na {register_config['host']}:{register_config['port']}.",
            "OFS — Greška", "error")


# ──────────────────────────────────────────────────────────────────────────────
# UNO ProtocolHandler
# ──────────────────────────────────────────────────────────────────────────────

class OFSDispatch(unohelper.Base, XDispatch):
    def __init__(self, context, frame, command):
        self.context = context
        self.frame = frame
        self.command = command

    def dispatch(self, url, args):
        global _g_frame
        try:
            ctx = self.context
            _g_frame = self.frame
            doc = self.frame.getController().getModel()
            if self.command == "send_to_ofs":
                send_to_ofs(ctx, doc)
            elif self.command == "show_document_settings":
                show_document_settings(ctx, doc)
            elif self.command == "show_kasa_settings":
                show_kasa_settings(ctx, doc)
            else:
                log.warning(f"Unknown command: {self.command}")
        except Exception as e:
            log.error(f"Dispatch [{self.command}] failed: {e}", exc_info=True)

    def addStatusListener(self, listener, url):
        pass

    def removeStatusListener(self, listener, url):
        pass


class OFSProtocolHandler(unohelper.Base, XDispatchProvider, XServiceInfo, XInitialization):
    def __init__(self, context):
        self.context = context
        self.frame = None
        _apply_log_level(context)
        log.info("OFSProtocolHandler initialized")

    def initialize(self, args):
        if args:
            self.frame = args[0]

    def queryDispatch(self, url, target_frame, search_flags):
        if url.Protocol == "vnd.fortunacommerc.ofs:":
            return OFSDispatch(self.context, self.frame, url.Path)
        return None

    def queryDispatches(self, requests):
        return [self.queryDispatch(r.FeatureURL, r.FrameName, r.SearchFlags)
                for r in requests]

    def getImplementationName(self):
        return "com.fortunacommerc.ofs.ProtocolHandler"

    def supportsService(self, name):
        return name == "com.sun.star.frame.ProtocolHandler"

    def getSupportedServiceNames(self):
        return ("com.sun.star.frame.ProtocolHandler",)


g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    OFSProtocolHandler,
    "com.fortunacommerc.ofs.ProtocolHandler",
    ("com.sun.star.frame.ProtocolHandler",))
