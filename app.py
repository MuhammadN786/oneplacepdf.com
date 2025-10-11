# app.py — serve BOTH apps; strip the broken QR tab from the main homepage (no changes to either app)

import os
import re
import importlib.util
from werkzeug.middleware.dispatcher import DispatcherMiddleware
from werkzeug.serving import run_simple

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def _load_app(module_name: str, file_path: str):
    """Load a module from file_path and return its Flask `app` object (unchanged)."""
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

# --- HTML filter: remove the QR tab button on the main homepage only ---
class _StripQrTab:
    """
    Removes the tab button:
      <button class="tab-btn" data-target="qr">...</button>
    from the main app's homepage HTML.
    """
    _BTN_RE = re.compile(
        rb'<button[^>]*\bclass="[^"]*\btab-btn\b[^"]*"[^>]*\bdata-target="qr"[^>]*>.*?</button>',
        re.IGNORECASE | re.DOTALL
    )

    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        is_home = environ.get("REQUEST_METHOD") == "GET" and environ.get("PATH_INFO", "/") in ("/", "")
        captured = {}

        def _cap_start(status, headers, exc_info=None):
            captured["status"] = status
            captured["headers"] = headers
            captured["exc_info"] = exc_info
            return lambda _: None

        app_iter = self.wsgi_app(environ, _cap_start)

        headers = captured.get("headers", [])
        ctype = next((v for k, v in headers if k.lower() == "content-type"), "")
        if not (is_home and isinstance(ctype, str) and "text/html" in ctype.lower()):
            start_response(captured["status"], headers, captured["exc_info"])
            for chunk in app_iter:
                yield chunk
            return

        body = b"".join(app_iter)
        body = self._BTN_RE.sub(b"", body)  # remove the QR tab button

        # fix Content-Length if present
        new_headers = []
        blen = str(len(body))
        for k, v in headers:
            if k.lower() == "content-length":
                new_headers.append((k, blen))
            else:
                new_headers.append((k, v))

        start_response(captured["status"], new_headers, captured["exc_info"])
        yield body

# You can override these via Render environment variables if needed:
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

# Load both apps exactly as they are
main_app = _load_app("main_app_module", MAIN_APP_PATH)
qr_app   = _load_app("qr_app_module",   QR_APP_PATH)

# Wrap main app to strip the QR tab from the homepage
main_app = _StripQrTab(main_app)

# Mount: main at "/", QR at "/qr"
app = DispatcherMiddleware(main_app, {
    "/qr": qr_app,
})

# Local dev convenience: `python app.py`
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    print(f"[wrapper] MAIN from: {MAIN_APP_PATH}")
    print(f"[wrapper] QR   from: {QR_APP_PATH}")
    print(f"[wrapper] http://0.0.0.0:{port}  (main=/ , qr=/qr)")
    run_simple("0.0.0.0", port, app, use_reloader=False, use_debugger=False)
