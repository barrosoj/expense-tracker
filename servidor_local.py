"""Local server for the app on this computer.

Serves the app at http://127.0.0.1:8080 (only reachable from this computer) and lets the app
save the expense Excel files it generates into the "Despesas" folder:
  - "Despesas AAAA.xlsx"        (monthly expense map for a year)
  - "Combustiveis AAAA-MM.xlsx" (fleet-card fuel purchases for a month)
The templates (Despesas.xlsx, Combustiveis.xlsx) are only read, never overwritten. Before a
file is replaced, the previous version is kept in "Despesas/.copias" (last 5 of each file).

Run:  python servidor_local.py          (add --abrir to open the browser)
"""
import http.server
import json
import os
import re
import sys
import time
import webbrowser
from urllib.parse import unquote, urlsplit

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(ROOT, "Despesas")
BACKUP_DIR = os.path.join(OUT_DIR, ".copias")
HOST = "127.0.0.1"
PORT = int(os.environ.get("PORT", "8080"))
OUTPUT_NAME = re.compile(r"^(Despesas \d{4}|Combustiveis \d{4}-\d{2})\.xlsx$")
MAX_BYTES = 25 * 1024 * 1024
KEEP_BACKUPS = 5


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()

    def log_message(self, fmt, *args):
        sys.stderr.write("%s %s\n" % (time.strftime("%H:%M:%S"), fmt % args))

    def _json(self, status, payload):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _same_origin(self):
        # Browsers always send Origin on PUT; only the app itself (this server) may write.
        origin = self.headers.get("Origin", "")
        return origin in ("http://%s:%d" % (HOST, PORT), "http://localhost:%d" % PORT)

    def do_GET(self):
        if urlsplit(self.path).path == "/api/estado":
            return self._json(200, {"gravarExcel": True, "pasta": "Despesas"})
        return super().do_GET()

    def do_PUT(self):
        if not self._same_origin():
            return self._json(403, {"erro": "Pedido recusado."})
        path = unquote(urlsplit(self.path).path)
        m = re.match(r"^/Despesas/([^/\\]+)$", path)
        name = m.group(1) if m else ""
        if not OUTPUT_NAME.match(name):
            return self._json(400, {"erro": "Nome de ficheiro não permitido."})
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0 or length > MAX_BYTES:
            return self._json(400, {"erro": "Ficheiro vazio ou demasiado grande."})
        data = self.rfile.read(length)
        if data[:2] != b"PK":
            return self._json(400, {"erro": "O conteúdo não é um ficheiro Excel."})

        os.makedirs(OUT_DIR, exist_ok=True)
        target = os.path.join(OUT_DIR, name)
        tmp = target + ".tmp"
        try:
            if os.path.exists(target):
                os.makedirs(BACKUP_DIR, exist_ok=True)
                stamp = time.strftime("%Y%m%d-%H%M%S")
                with open(target, "rb") as src, open(os.path.join(BACKUP_DIR, "%s %s.xlsx" % (name[:-5], stamp)), "wb") as dst:
                    dst.write(src.read())
                old = sorted(f for f in os.listdir(BACKUP_DIR) if f.startswith(name[:-5] + " "))
                for f in old[:-KEEP_BACKUPS]:
                    os.remove(os.path.join(BACKUP_DIR, f))
            with open(tmp, "wb") as f:
                f.write(data)
            os.replace(tmp, target)
        except PermissionError:
            if os.path.exists(tmp):
                os.remove(tmp)
            return self._json(423, {"erro": "O ficheiro %s está aberto no Excel. Fecha-o para a app o poder atualizar." % name})
        except OSError as e:
            if os.path.exists(tmp):
                os.remove(tmp)
            return self._json(500, {"erro": "Não foi possível gravar %s: %s" % (name, e)})
        return self._json(200, {"gravado": name})


def main():
    server = http.server.ThreadingHTTPServer((HOST, PORT), Handler)
    url = "http://%s:%d" % (HOST, PORT)
    print("App em %s  (Ctrl+C para parar)" % url)
    print("Os Excel de despesas são gravados em %s" % OUT_DIR)
    if "--abrir" in sys.argv:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
