# OFS — LibreOffice Extension

## Pregled

LibreOffice Calc extension za slanje faktura na OFS fiskalnu kasu (ofs.ba, Republika Srpska).
Čita aktivni ODS dokument, prikazuje dialog za potvrdu, i šalje request na ESIR API.
Reference ćelija se konfigurišu per-dokument; podešavanja kase se čuvaju u LO registry.

## Fajlovi

| Fajl | Opis |
|------|------|
| `python/ofs.py` | Sva logika — čitanje fakture, UNO dijalozi, HTTP slanje, mrežno skeniranje |
| `dialogs/*.xdl` | XDL dijalozi: `FiscalizationConfirm`, `DocumentSettings`, `RegisterSettings`, `ScanResults`, `JsonPreview` |
| `Addons.xcu` | Definicija menija: OFS > Fiskalizacija \| Podešavanja dokumenta \| Podešavanje kase; context restriction po tipu dokumenta |
| `ProtocolHandler.xcu` | Registracija `vnd.fortunacommerc.ofs:*` protokola sa LO |
| `META-INF/manifest.xml` | Registracija fajlova u extension paketu |
| `OFS.xcs` | Schema za LO registry: Kasa i Logging grupe |
| `description.xml` | Identitet: `com.fortunacommerc.ofs` v2.0.0 |
| `python/mock_esir.py` | Standalone mock ESIR server za razvoj (nije u .oxt) |
| `Makefile` | `make build` → `OFS.oxt` / `make install` → `unopkg add` |

## Build & instalacija

```bash
make build    # kreira OFS.oxt
make install  # build + unopkg remove/add (restart LO)
make clean    # briše .oxt
```

`ofs.py` je pakovan unutar `.oxt` kao UNO komponenta — ne kopira se u `Scripts/python`.

## Podešavanje kase (sistemsko)

Čuva se u LO registry: `/com.fortunacommerc.ofs.Settings/Register`. Postavljanje: **OFS > Podešavanje kase**.

| Polje | Default | Opis |
|-------|---------|------|
| `Mac` | `""` | MAC adresa kase (za filtriranje pri skeniranju) |
| `Host` | `localhost` | ESIR hostname |
| `Port` | `3566` | ESIR port |
| `ApiKey` | `""` | Bearer token za autentifikaciju |
| `Pin` | `""` | PIN kod operatera |

## Podešavanja dokumenta (per-dokument)

Sve reference ćelija se čuvaju kao **custom document properties** (File > Properties > Custom Properties).
Postavljanje: **OFS > Podešavanja dokumenta** — dialog sa 15 polja u 5 sekcija.

### Document properties — ključevi

| Property | Opis | Obavezno |
|----------|------|----------|
| `OFS_Cell_BuyerId` | Ćelija sa BuyerID (format `VP:xxx`) | Da |
| `OFS_Range_Buyer` | Opseg ćelija sa podacima kupca (npr. `B5:B8`); sve neprazne ćelije prikazuju se u pregledu, svaka u novom redu | Ne |
| `OFS_Items_FirstRow` | Početni red stavki (broj, npr. `13`) | Da |
| `OFS_Items_LastRow` | Krajnji red stavki (broj, npr. `27`) | Da |
| `OFS_Col_ItemName` | Kolona naziva artikla (npr. `C`) | Da |
| `OFS_Col_ItemUOM` | Kolona jed. mjere (opciono) | Ne |
| `OFS_Col_ItemQty` | Kolona količine (npr. `E`) | Da |
| `OFS_Col_ItemPrice` | Kolona jed. cijene sa PDV (npr. `H`) | Da |
| `OFS_Col_ItemLabel` | Kolona poreske oznake (opciono, default `E`) | Ne |
| `OFS_Col_ItemBarcode` | Kolona bar koda | Da |
| `OFS_Cell_TotalBase` | Ćelija ukupno bez PDV (npr. `J29`) | Da |
| `OFS_Cell_TotalVAT` | Ćelija iznos PDV (npr. `J30`) | Da |
| `OFS_Cell_TotalWithVAT` | Ćelija ukupno sa PDV (npr. `J31`) | Da |
| `OFS_Cell_VATPercent` | Ćelija sa procentom PDV-a (npr. `17`) | Da |
| `OFS_Cell_InvoiceNumber` | Ćelija za upis broja računa nakon fiskalizacije (opciono) | Ne |

**Detekcija kraja stavki**: prva prazna ćelija u koloni `OFS_Col_ItemName`.
**API name format**: `<naziv> /<jed.mjere>` ako jed. mjere postoji, inače samo `<naziv>`.
**VAT label default**: `"E"` — koristi se ako `OFS_Col_ItemLabel` nije podešen ili je ćelija prazna.

## Meni struktura

OFS meni vidljiv u `SpreadsheetDocument` i `StartModule` (Start Center prikazuje samo Podešavanje kase).

```
OFS
├── Inicijalizacija        [SpreadsheetDocument + StartModule]
├── Fiskalizacija          [samo SpreadsheetDocument]
├── Podešavanja dokumenta  [samo SpreadsheetDocument]
└── Podešavanje kase       [SpreadsheetDocument + StartModule]
```

LO ne renderuje addon meni separatore — ne koristiti.

## ESIR API

```
POST http://<host>:<port>/api/invoices
Content-Type: application/json
Authorization: Bearer <api_key>
```

Request body omotan u `"invoiceRequest": { ... }`.

| Polje | Vrijednosti |
|-------|-------------|
| `invoiceType` | `Normal` \| `Proforma` \| `Copy` \| `Training` \| `Advance` |
| `transactionType` | `Sale` \| `Refund` |
| `payment[].paymentType` | `WireTransfer` \| `Cash` \| `Card` \| `Check` \| `Voucher` \| `MobileMoney` \| `Other` |
| `items[].labels` | Per-item poreska oznaka (default `["E"]`) |
| `items[].gtin` | Per-item bar kod (opciono) |
| `buyerId` | format `"VP:<JIB>"` |
| `referentDocumentNumber` | Ref. broj (za Copy/Refund, opciono — šalje se samo ako nije prazan) |
| `referentDocumentDT` | Ref. datum (za Copy/Refund, opciono — šalje se samo ako nije prazan); korisnik unosi `DD.MM.YYYY HH:MM:SS`, `_parse_ref_dt()` konvertuje u ISO format |

**Ne šaljemo** `dateAndTimeOfIssue`, `cashier`, `buyerCostCenterId` — kasa ih sama popunjava.

Na uspjeh: `invoiceNumber` iz response-a se upisuje u ćeliju `OFS_Cell_InvoiceNumber` (ako je podešena).

**Ćirilične poreske oznake**: ESIR može koristiti ćirilična slova za poreske oznake (npr. `Е` umjesto `E`,
`А` umjesto `A`, `К` umjesto `K`). Tokom `esir_init()`, `currentTaxRates` iz `/api/status` se analizira.
Ako su oznake ćirilične, automatski se prevode iz latiničnih (iz dokumenta) u ćirilične prije slanja.

## ESIR Inicijalizacija

Prije slanja računa, `esir_init()` automatski izvršava tri koraka:

| Korak | Endpoint | Metoda | Opis |
|-------|----------|--------|------|
| 1 | `/api/attention` | GET | Provjera dostupnosti kase |
| 2 | `/api/status` | GET | Provjera statusa SE; `"1300"` ili `"1500"` u `gsc` znači da je zaključan |
| 3 | `/api/pin` | POST | Otključavanje SE PIN kodom (samo ako korak 2 kaže da treba) |

**Napomena**: `/api/pin` koristi `Content-Type: text/plain` — tijelo je samo PIN string, ne JSON.

Status kodovi u `/api/pin` odgovoru: `0100`=OK, `1300`=nema SE, `2400`=LPFR ne spreman, `2800`/`2806`=neispravan PIN.

Inicijalizacija je tiha na uspjeh, poruka se prikazuje samo na grešku. Dugme "Testiraj" u podešavanjima kase poziva istu `esir_init()` sa trenutnim (nesačuvanim) vrijednostima iz dijaloga.

## Mock ESIR Server

`python/mock_esir.py` — standalone mock za razvoj i testiranje. Nije uključen u `.oxt`.

```bash
python3 python/mock_esir.py                    # default: port=3566, key=test, pin=1111
python3 python/mock_esir.py --locked           # SE zaključan na startu
python3 python/mock_esir.py --cyrillic         # ćirilične poreske oznake (А, Е, К)
python3 python/mock_esir.py --port 4000 --key mykey --pin 9999
```

Endpointi: `/api/attention`, `/api/status`, `/api/pin`, `/api/invoices`.
Invoice brojevi: `MOCK-000001`, `MOCK-000002`, ...

## UNO ProtocolHandler komande

Komande se pozivaju putem `vnd.fortunacommerc.ofs:<command>` protokola.
`OFSProtocolHandler` prima frame u `initialize()`, prosljeđuje ga `OFSDispatch`,
koji poziva odgovarajuću Python funkciju sa `(ctx, doc)`.

### `esir_initialize(ctx)`
Entry point za OFS > Inicijalizacija. Čita sačuvanu konfiguraciju kase (`load_register_config`),
poziva `esir_init()` (attention → status → pin), i prikazuje rezultat korisniku putem `show_msgbox`.

### `send_to_ofs(ctx, doc)`
Glavni entry point — OFS > Fiskalizacija. Tok:
1. `load_register_config(ctx)` — čita kasa config iz LO registry
2. `load_cell_references(doc)` + `validate_cell_references()` — čita i validira doc properties
3. `read_invoice_data(doc, sheet, cell_references)` — čita stavke po konfigurisanim referencama
4. `show_confirm_dialog(ctx, invoice_data, register_config)` — UNO dialog: tip računa, transakcija, plaćanje, ref. polja
5. `esir_init(register_config)` — inicijalizacija: attention → status → pin (tiha na uspjeh)
6. `build_esir_request()` + `send_to_esir_api()` — slanje na ESIR
7. Na uspjeh: `show_success_dialog()` + `write_invoice_number()` (upis broja nazad u dokument)

**`build_items_text()` format** (monospace, širina 60 znakova):
- Header red: `Naziv` (ljevo, 16ch) + `Cijena` (14ch) + `Kol.` (10ch) + `Ukupno` (14ch) desno
- Per-stavka: red 1 = barcode + naziv + /uom + (label); red 2 = cijena x količina = ukupno (right-aligned)
- Totali: label (46ch) + `format_number(v, 11, 2)` (14ch), bez valute

**JSON Preview**: dugme "Pregled" u confirm dialogu prikazuje `JsonPreview` dialog sa read-only prikazom ESIR request JSON-a.

### `show_document_settings(ctx, doc)`
UNO dialog sa 15 polja u 5 sekcija za konfigurisanje referenci ćelija.
Svako polje je label + text input, pre-populated iz sačuvanih doc properties.
Na Sačuvaj: iteracija svih polja, `set_doc_property()` za svako, `doc.setModified(True)`.

### `show_kasa_settings(ctx, doc=None)`
UNO dialog sa 5 polja (MAC, host, port, API ključ, PIN) + dugmad "Testiraj" i "Skeniraj".
Pre-populated iz `load_register_config()`. Validira port (integer), obaveznost polja, PIN dužinu (4 znaka).
Na Sačuvaj: `save_register_config()`.
Testiraj: čita nesačuvane vrijednosti iz dijaloga, poziva `esir_init()` i prikazuje rezultat.
Skeniraj: pokreće mrežno skeniranje (vidi sekciju Mrežno skeniranje).

## Mrežno skeniranje

Dugme "Skeniraj" u podešavanjima kase skenira lokalne mreže za uređaje sa otvorenim ESIR portom.

**Tok**:
1. `_get_local_subnets()` — detektuje lokalne IPv4 podmreže (`ip -4 -o addr show` na Linuxu, PowerShell na Windowsu, fallback na `socket` sa /24)
2. `_scan_network(port)` — TCP port scan svih hostova u podmrežama (50 threadova, timeout 0.3s)
3. `_verify_scan_results()` — HTTP probe `/api/attention` na pronađenim hostovima (10 threadova)
4. `_get_arp_table()` — čita ARP cache za MAC adrese (`/proc/net/arp` na Linuxu)

**MAC filtriranje**: ako je MAC polje popunjeno, rezultati se filtriraju po ARP tabeli.
**Jedan rezultat**: automatski popunjava Host i MAC polja u dialogu.
**Više rezultata**: prikazuje `ScanResults` dialog sa listom uređaja i statusima (200=OK, 401=pogrešan ključ, 404=nepoznat API).

## Logging

Log fajl: `~/.config/ofs/extension.log` (TimedRotatingFileHandler, sedmična rotacija `W0`, gzip kompresija, bez backupCount limita).
Logger se inicijalizuje **prije** UNO importa — tako se greške pri uvozu loguju.
Log level se čita iz LO registry: `/com.fortunacommerc.ofs.Settings/Logging/LogLevel` (default `INFO`).
Primjenjuje se u `OFSProtocolHandler.__init__()` → `_apply_log_level(ctx)`.
Config node konstanta: `_LOGGING_CONFIG_NODE`.

## LibreOffice API napomene

- `getString()` / `getValue()` — uvijek vraćaju izračunatu vrijednost, nikad formulu
- `getCellByPosition(col, row)` — oba indeksa su 0-bazirani
- `read_cell_string(sheet, "B13")` / `read_cell_number(sheet, "B13")` — helperi za čitanje po adresi ćelije
- `format_number(v, int_width, decimals)` — tačke za hiljade, zarez za decimale, right-align integer dio
- `format_amount_km(v)` — `format_number(v, 9, 2)` + ` KM` sufiks
- `createMessageBox(parent, type, buttons, title, message)` — **5 parametara** (LO 26.x)
  - Starija verzija imala 6 (sa Rectangle parametrom) — **ne koristiti**
- `_g_frame` — module-level varijabla; postavlja se na početku svakog entry pointa
- `_load_dialog(ctx, name)` — učitava XDL dialog iz .oxt paketa, kreira peer, centrira; vraća dialog
- `ConfigurationUpdateAccess` implementira `XNameReplace` → koristiti `replaceByName()`, **ne** `setByName()`
- Protocol URL format u Addons.xcu: `vnd.fortunacommerc.ofs:<command>`
- Ne koristiti emotikone/unicode simbole (★, ✓, itd.) u string porukama dijaloga — LO font rendering može biti nepouzdan

## Okruženje

- LibreOffice 26.2+ (dozvoljena novija bugfix verzija), Python 3.12.12 (ugrađen u LibreOffice, ne sistemski), Linux (Ubuntu)
- Bez eksternih Python zavisnosti — samo stdlib + UNO bindings