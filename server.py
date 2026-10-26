"""Máy chủ web nhỏ (chỉ dùng thư viện chuẩn + lxml/openpyxl).

Chạy:  python3 server.py [cổng]     rồi mở http://localhost:8000
"""
import json
import mimetypes
import re
import sys
import traceback
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, unquote, urlparse

import exporter
import importer

ROOT = Path(__file__).parent
STATIC = ROOT / "static"
MAX_BODY = 40 * 1024 * 1024


def safe_name(state):
    h = state.get("header") or {}
    base = "BOM_%s_Muc%s" % (h.get("code") or "SP", h.get("muc") or "")
    return re.sub(r"[^\w.\-()]+", "_", base) + ".xlsm"


class Handler(BaseHTTPRequestHandler):
    def log_message(self, fmt, *args):
        sys.stderr.write("%s - %s\n" % (self.address_string(), fmt % args))

    def _send(self, code, body, ctype="application/json; charset=utf-8", headers=None):
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        for k, v in (headers or {}).items():
            self.send_header(k, v)
        self.end_headers()
        self.wfile.write(body)

    def _json(self, code, obj):
        self._send(code, json.dumps(obj, ensure_ascii=False).encode("utf-8"))

    def _body(self):
        n = int(self.headers.get("Content-Length") or 0)
        if n <= 0 or n > MAX_BODY:
            raise ValueError("Dữ liệu gửi lên rỗng hoặc quá lớn")
        return self.rfile.read(n)

    def do_GET(self):
        path = unquote(urlparse(self.path).path)
        if path == "/api/lists":
            return self._send(200, (ROOT / "data" / "lists.json").read_bytes())
        if path == "/":
            path = "/index.html"
        f = (STATIC / path.lstrip("/")).resolve()
        if STATIC.resolve() not in f.parents or not f.is_file():
            return self._json(404, {"error": "Không tìm thấy"})
        ctype = mimetypes.guess_type(f.name)[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype.endswith("javascript") or ctype.endswith("json"):
            ctype += "; charset=utf-8"
        self._send(200, f.read_bytes(), ctype)

    def do_POST(self):
        path = urlparse(self.path).path
        try:
            if path == "/api/export":
                state = json.loads(self._body().decode("utf-8"))
                data = exporter.export_xlsm(state)
                name = safe_name(state)
                return self._send(200, data, "application/vnd.ms-excel.sheet.macroEnabled.12", {
                    "Content-Disposition": "attachment; filename*=UTF-8''%s" % quote(name)})
            if path == "/api/import":
                return self._json(200, importer.import_xlsm(self._body()))
            return self._json(404, {"error": "Không tìm thấy"})
        except (exporter.ExportError, ValueError) as e:
            return self._json(400, {"error": str(e)})
        except Exception as e:  # noqa: BLE001 - trả lỗi rõ ràng cho giao diện
            traceback.print_exc()
            return self._json(500, {"error": "Lỗi máy chủ: %s" % e})


def main():
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    srv = ThreadingHTTPServer(("0.0.0.0", port), Handler)
    print("BOM web chạy tại http://localhost:%d" % port)
    srv.serve_forever()


if __name__ == "__main__":
    main()
