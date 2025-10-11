# app.py — serve BOTH of your existing apps in one Render service:
# - Main app at "/"
# - QR app at "/qr"
# No changes to either app's code. The wrapper also rewrites the broken QR tab into a link.

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
    raise FileNotFoundError(
        "Main app not found. Put it at ./main_app/app.py or set env MAIN_APP_PATH to its file path."
    )
if not QR_APP_PATH:
    raise FileNotFoundError(
        "QR app not found. Put it at ./qr_app/app.py or set env QR_APP_PATH to its file path."
    )

# --- HTML filter to fix the non-working bottom QR tab without editing the main app ---
class HtmlFixQrTab:
    """
    On GET / (homepage) HTML, replace the QR tab button
    <button class="tab-btn" data-target="qr">...</button>
    with a real link to /qr/create so it works.
    """
    # match the whole button (class contains tab-btn AND data-target="qr")
    TAB_BTN_RE = re.compile(
        rb'<button[^>]*\bclass="[^"]*\btab-btn\b[^"]*"[^>]*\bdata-target="qr"[^>]*>.*?</button>',
        re.DOTALL | re.IGNORECASE
    )
    # replacement anchor
    REPLACEMENT = b'<a class="tab-btn" href="/qr/create">Create QR Code</a>'

    def __init__(self, wsgi_app):
        self.wsgi_app = wsgi_app

    def __call__(self, environ, start_response):
        is_home = (
            environ.get("REQUEST_METHOD") == "GET" and
            environ.get("PATH_INFO", "/") in ("/", "")
        )
        captured = {}

        def _cap_start(status, headers, exc_info=None):
            captured["status"] = status
            captured["headers"] = headers
            captured["exc_info"] = exc_info
            return lambda x: None

        app_iter = self.wsgi_app(environ, _cap_start)

        headers = captured.get("headers", [])
        ctype = next((v for k, v in headers if k.lower() == "content-type"), "")
        is_html = isinstance(ctype, str) and "text/html" in ctype.lower()

        if not (is_home and is_html):
            start_response(captured["status"], headers, captured["exc_info"])
            for chunk in app_iter:
                yield chunk
            return

        body = b"".join(app_iter)
        body = self.TAB_BTN_RE.sub(self.REPLACEMENT, body)  # <-- rewrite the bottom tab

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

# Load both apps exactly as they are
main_app = _load_app("main_app_module", MAIN_APP_PATH)
qr_app   = _load_app("qr_app_module",   QR_APP_PATH)

# Wrap main app to fix the bottom QR tab
fixed_main = HtmlFixQrTab(main_app)

# Mount: main at "/", QR at "/qr"
app = DispatcherMiddleware(fixed_main, {
    "/qr": qr_app,
})

# Local dev convenience: `python app.py`
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    print(f"[wrapper] MAIN from: {MAIN_APP_PATH}")
    print(f"[wrapper] QR   from: {QR_APP_PATH}")
    print(f"[wrapper] http://0.0.0.0:{port}  (main=/ , qr=/qr)")
    run_simple("0.0.0.0", port, app, use_reloader=False, use_debugger=False)
