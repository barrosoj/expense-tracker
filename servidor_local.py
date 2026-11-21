"""Local server for the app on this computer (for testing; the app itself runs on GitHub Pages).

Serves the app at http://127.0.0.1:8080, only reachable from this computer. Read-only: the
Excel files now live in Google Drive, so this server never writes anything to disk.
Hardening:
  - only answers requests addressed to 127.0.0.1 / localhost (blocks DNS-rebinding pages);
  - only GET/HEAD; hidden paths (.git, .claude…) and source/config files are never served;
  - security headers on every response.

Run:  python servidor_local.py          (add --abrir to open the browser)
"""
import http.server
import os
import posixpath
import sys
import time
import webbrowser
from urllib.parse import unquote, urlsplit

ROOT = os.path.dirname(os.path.abspath(__file__))
HOST = "127.0.0.1"
PORT = int(os.environ.get("PORT", "8080"))
ALLOWED_HOSTS = {"127.0.0.1:%d" % PORT, "localhost:%d" % PORT}
BLOCKED_EXT = (".py", ".cmd", ".sh", ".sql", ".md", ".xlsx", ".xls", ".csv", ".json")


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        super().end_headers()

    def log_message(self, fmt, *args):
        sys.stderr.write("%s %s\n" % (time.strftime("%H:%M:%S"), fmt % args))

    def list_directory(self, path):
        self.send_error(404)
        return None

    def _allowed(self):
        if self.headers.get("Host", "") not in ALLOWED_HOSTS:
            return False
        path = posixpath.normpath(unquote(urlsplit(self.path).path))
        parts = [p for p in path.split("/") if p]
        if any(p.startswith(".") for p in parts):
            return False
        # manifest.webmanifest is fine; data and source files are not.
        return not path.lower().endswith(BLOCKED_EXT)

    def do_GET(self):
        if not self._allowed():
            return self.send_error(404)
        return super().do_GET()

    def do_HEAD(self):
        if not self._allowed():
            return self.send_error(404)
        return super().do_HEAD()


def main():
    server = http.server.ThreadingHTTPServer((HOST, PORT), Handler)
    url = "http://%s:%d" % (HOST, PORT)
    print("App em %s  (Ctrl+C para parar)" % url)
    if "--abrir" in sys.argv:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
