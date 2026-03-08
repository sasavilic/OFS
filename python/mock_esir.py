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
"""
Mock ESIR Server za razvoj i testiranje OFS ekstenzije.

Simulira ESIR API endpointe:
  GET  /api/attention  — provjera dostupnosti
  GET  /api/status     — status bezbjednosnog elementa (SE)
  POST /api/pin        — otključavanje SE PIN kodom
  POST /api/invoices   — mock fiskalizacija

Pokretanje:
  python3 mock_esir.py                    # default: port=3566, key=test, pin=1111
  python3 mock_esir.py --locked           # SE zaključan na startu
  python3 mock_esir.py --port 4000 --key mykey --pin 9999
"""

import argparse
import json
import datetime
from http.server import HTTPServer, BaseHTTPRequestHandler


class ESIRHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        print(f"[ESIR] {self.command} {self.path} -> {args[1] if len(args) > 1 else args[0]}")

    def _check_auth(self):
        expected = self.server.api_key
        if not expected:
            return True
        auth = self.headers.get("Authorization", "")
        if auth == f"Bearer {expected}":
            return True
        self._send_json(401, {"error": "Unauthorized"})
        return False

    def _send_json(self, code, data):
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _send_text(self, code, text):
        body = text.encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        if self.path == "/api/attention":
            self._handle_attention()
        elif self.path == "/api/status":
            self._handle_status()
        else:
            self._send_json(404, {"error": "Not found"})

    def do_POST(self):
        if self.path == "/api/pin":
            self._handle_pin()
        elif self.path == "/api/invoices":
            self._handle_invoices()
        else:
            self._send_json(404, {"error": "Not found"})

    def _handle_attention(self):
        if not self._check_auth():
            return
        self._send_json(200, {"status": "OK"})

    def _handle_status(self):
        if not self._check_auth():
            return
        gsc = ["1300"] if self.server.locked else []
        if self.server.cyrillic:
            labels = {
                "bez_pdv": "\u041A",    # К (Cyrillic)
                "nije_pdv": "\u0410",   # А (Cyrillic)
                "pdv": "\u0415",        # Е (Cyrillic)
            }
        else:
            labels = {"bez_pdv": "K", "nije_pdv": "A", "pdv": "E"}
        tax_rates = {
            "groupId": 2,
            "taxCategories": [
                {
                    "categoryType": 0, "name": "Bez PDV", "orderId": 1,
                    "taxRates": [{"label": labels["bez_pdv"], "rate": 0.0}],
                },
                {
                    "categoryType": 0, "name": "Nije u PDV", "orderId": 2,
                    "taxRates": [{"label": labels["nije_pdv"], "rate": 0.0}],
                },
                {
                    "categoryType": 0, "name": "PDV", "orderId": 3,
                    "taxRates": [{"label": labels["pdv"], "rate": 17.0}],
                },
            ],
            "validFrom": "2024-07-01T00:00:00.000+02:00",
        }
        self._send_json(200, {
            "sdcDateTime": datetime.datetime.now().isoformat(),
            "gsc": gsc,
            "uid": "MOCK-SE-001",
            "currentTaxRates": tax_rates,
        })

    def _handle_pin(self):
        if not self._check_auth():
            return
        length = int(self.headers.get("Content-Length", 0))
        pin = self.rfile.read(length).decode("utf-8").strip()

        if pin == self.server.pin:
            self.server.locked = False
            self._send_text(200, "0100")
        else:
            self._send_text(400, "2800")

    def _handle_invoices(self):
        if not self._check_auth():
            return

        if self.server.locked:
            self._send_json(400, {"error": "SE zaključan, potrebno otključavanje PIN kodom (1300)"})
            return

        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length).decode("utf-8")

        try:
            data = json.loads(body)
            print("[ESIR] Body: \n", json.dumps(data, indent=2))
        except json.JSONDecodeError:
            self._send_json(400, {"error": "Invalid JSON"})
            return

        self.server.invoice_counter += 1
        inv_req = data.get("invoiceRequest", {})
        total = 0
        for p in inv_req.get("payment", []):
            total += p.get("amount", 0)

        invoice_number = f"MOCK-{self.server.invoice_counter:06d}"

        self._send_json(200, {
            "invoiceNumber": invoice_number,
            "invoiceType": inv_req.get("invoiceType", "Normal"),
            "transactionType": inv_req.get("transactionType", "Sale"),
            "totalAmount": round(total, 2),
            "invoiceCounter": f"{self.server.invoice_counter}/{self.server.invoice_counter}NS",
            "verificationUrl": f"https://mock.ofs.ba/v/{invoice_number}",
            "sdcDateTime": datetime.datetime.now().isoformat(),
        })


def main():
    parser = argparse.ArgumentParser(description="Mock ESIR Server")
    parser.add_argument("--port", type=int, default=3566, help="Port (default: 3566)")
    parser.add_argument("--key", default="test", help="API key (default: test)")
    parser.add_argument("--pin", default="1111", help="PIN kod (default: 1111)")
    parser.add_argument("--locked", action="store_true", help="Start sa zaključanim SE")
    parser.add_argument("--cyrillic", action="store_true", help="Koristi ćirilične poreske oznake (А, Е, К)")
    args = parser.parse_args()

    server = HTTPServer(("0.0.0.0", args.port), ESIRHandler)
    server.api_key = args.key
    server.pin = args.pin
    server.locked = args.locked
    server.cyrillic = args.cyrillic
    server.invoice_counter = 0

    status = "ZAKLJUČAN" if args.locked else "otključan"
    labels_info = "ćirilica (А, Е, К)" if args.cyrillic else "latinica (A, E, K)"
    print(f"[ESIR] Mock server pokrenut na portu {args.port}")
    print(f"[ESIR] API key: {args.key}")
    print(f"[ESIR] PIN: {args.pin}")
    print(f"[ESIR] SE status: {status}")
    print(f"[ESIR] Poreske oznake: {labels_info}")
    print(f"[ESIR] Ctrl+C za zaustavljanje")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[ESIR] Server zaustavljen.")
        server.server_close()


if __name__ == "__main__":
    main()
