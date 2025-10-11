# app.py — serve BOTH of your existing apps in one Render service:
# - Main app at "/"
# - QR app at "/qr"
# No changes to either app's code.

import os
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

# Load both apps exactly as they are
main_app = _load_app("main_app_module", MAIN_APP_PATH)
qr_app   = _load_app("qr_app_module",   QR_APP_PATH)

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
