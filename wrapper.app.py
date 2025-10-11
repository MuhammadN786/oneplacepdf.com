# app.py — serve BOTH apps; rewrite any external "Create QR Code" link to /qr/create

import os, re, importlib.util
from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.serving import run_simple

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def _load_app(module_name: str, file_path: str):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot load module from {file_path}")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    if not hasattr(mod, "app"):
        raise RuntimeError(f"`app` not found in {file_path}. Ensure it defines `app = Flask(__name__)`.")
    return mod.app

def _first_existing(paths):
    for p in paths:
        if p and os.path.isfile(p):
            return p
    return None

# --- HTML filter: rewrite any external QR link to internal /qr/create ---
class _RewriteQrLink:
    """
    Rewrites anchors like:
      href="https://oneplacepdf-com-qr-generator.onrender.com/create"
      href="http://oneplacepdf-com-qr-generator.onrender.com/create"
    to:
      href="/qr/create"
    """
    _HREF_RE = re.compile(
        rb'href\s*=\s*"(https?://oneplacepdf-com-qr-generator\.onrender\.com/create)"',
        re.IGNORECASE
    )
    _REPLACEMENT = b'href="/qr/create"'

    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
            captured = {}
            def _cap_start(status, headers, exc_info=None):
                captured["status"] = status
                captured["headers"] = headers
                captured["exc_info"] = exc_info
                return lambda _: None

            app_iter = self.wsgi_app(environ, _cap_start)
            headers = captured.get("headers", [])
            ctype = next((v for k, v in headers if k.lower() == "content-type"), "")
            if not (isinstance(ctype, str) and "text/html" in ctype.lower()):
                start_response(captured["status"], headers, captured["exc_info"])
                for chunk in app_iter:
                    yield chunk
                return

            body = b"".join(app_iter)
            body = self._HREF_RE.sub(self._REPLACEMENT, body)

            new_headers, blen = [], str(len(body))
            for k, v in headers:
                if k.lower() == "content-length":
                    new_headers.append((k, blen))
                else:
                    new_headers.append((k, v))

            start_response(captured["status"], new_headers, captured["exc_info"])
            yield body

# You can override these via Render env vars if needed:
# MAIN_APP_PATH=/full/path/to/your/main app.py
# QR_APP_PATH=/full/path/to/your/qr app.py
MAIN_APP_PATH = _first_existing([
    os.environ.get("MAIN_APP_PATH"),
    os.path.join(BASE_DIR, "main_app", "app.py"),
])
QR_APP_PATH = _first_existing([
    os.environ.get("QR_APP_PATH"),
    os.path.join(BASE_DIR, "qr_app", "app.py"),
])

if not MAIN_APP_PATH:
    raise FileNotFoundError("Main app not found. Put it at ./main_app/app.py or set env MAIN_APP_PATH.")
if not QR_APP_PATH:
    raise FileNotFoundError("QR app not found. Put it at ./qr_app/app.py or set env QR_APP_PATH.")

main_app = _load_app("main_app_module", MAIN_APP_PATH)
qr_app   = _load_app("qr_app_module",   QR_APP_PATH)

# Wrap main app to rewrite external QR links
main_app = _RewriteQrLink(main_app)

# Mount: main at "/", QR at "/qr"
app = DispatcherMiddleware(main_app, { "/qr": qr_app })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    run_simple("0.0.0.0", port, app, use_reloader=False, use_debugger=False)
