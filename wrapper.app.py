# app.py — wrapper
# - Mounts main app at "/"
# - Mounts QR app at "/qr"
# - Rewrites hard-coded external QR link to /qr/create
# - Allows indexing on /editor by replacing/adding robots meta tag

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

# --- Middleware: Rewrite external QR link to /qr/create ----------------------
class _RewriteQrLink:
    _HREF_RE = re.compile(
        rb'href\s*=\s*"(https?://oneplacepdf-com-qr-generator\.onrender\.com/create)"',
        re.IGNORECASE
    )
    _REPLACEMENT = b'href="/qr/create"'
    def __init__(self, wsgi_app): self.wsgi_app = wsgi_app
    def __call__(self, environ, start_response):
        captured = {}
        def cap(status, headers, exc=None):
            captured["status"], captured["headers"], captured["exc"] = status, headers, exc
            return lambda _:_  # swallow write
        it = self.wsgi_app(environ, cap)
        headers = captured.get("headers", [])
        ctype = next((v for k,v in headers if k.lower()=="content-type"), "")
        if "text/html" not in str(ctype).lower():
            start_response(captured["status"], headers, captured["exc"])
            for chunk in it: yield chunk
            return
        body = b"".join(it)
        body = self._HREF_RE.sub(self._REPLACEMENT, body)
        new_headers = [(k, (str(len(body)) if k.lower()=="content-length" else v)) for k,v in headers]
        start_response(captured["status"], new_headers, captured["exc"])
        yield body

# --- Middleware: Force index,follow on /editor only --------------------------
class _AllowEditorIndexing:
    # replace existing noindex meta if present
    _NOINDEX_RE = re.compile(
        rb'<meta\s+name=["\']robots["\']\s+content=["\']\s*noindex\s*,\s*nofollow\s*["\']\s*/?>',
        re.IGNORECASE
    )
    # if no robots tag exists, inject one before </head>
    _HEAD_CLOSE_RE = re.compile(rb'</head\s*>', re.IGNORECASE)
    _ROBOTS_ALLOW = b'<meta name="robots" content="index,follow" />'
    def __init__(self, wsgi_app): self.wsgi_app = wsgi_app
    def __call__(self, environ, start_response):
        if environ.get("PATH_INFO") != "/editor":
            return self.wsgi_app(environ, start_response)
        captured = {}
        def cap(status, headers, exc=None):
            captured["status"], captured["headers"], captured["exc"] = status, headers, exc
            return lambda _:_ 
        it = self.wsgi_app(environ, cap)
        headers = captured.get("headers", [])
        ctype = next((v for k,v in headers if k.lower()=="content-type"), "")
        if "text/html" not in str(ctype).lower():
            start_response(captured["status"], headers, captured["exc"])
            for chunk in it: yield chunk
            return
        body = b"".join(it)
        # replace noindex -> index
        new_body = self._NOINDEX_RE.sub(self._ROBOTS_ALLOW, body)
        if new_body == body:
            # no robots tag found—inject index,follow before </head>
            new_body = self._HEAD_CLOSE_RE.sub(self._ROBOTS_ALLOW + b'\n</head>', body, count=1)
        body = new_body
        new_headers = [(k, (str(len(body)) if k.lower()=="content-length" else v)) for k,v in headers]
        start_response(captured["status"], new_headers, captured["exc"])
        yield body

# --- Paths (env overrides supported) -----------------------------------------
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

# --- Compose -----------------------------------------------------------------
main_app = _load_app("main_app_module", MAIN_APP_PATH)
qr_app   = _load_app("qr_app_module",   QR_APP_PATH)

# order matters: rewrite QR links, then allow /editor indexing
main_app = _AllowEditorIndexing(_RewriteQrLink(main_app))

# Mount: main at "/", QR at "/qr"
app = DispatcherMiddleware(main_app, { "/qr": qr_app })

if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    run_simple("0.0.0.0", port, app, use_reloader=False, use_debugger=False)
