import io, os, secrets, json, datetime as dt
from slugify import slugify
from PIL import Image
from flask import (
    Flask, request, redirect, url_for, render_template_string,
    send_file, send_from_directory, flash
)
from flask_sqlalchemy import SQLAlchemy
import qrcode
import qrcode.image.svg as qsvg

# ── App & DB config ──────────────────────────────────────────────────────────
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", secrets.token_hex(16))
app.config["MAX_CONTENT_LENGTH"] = int(os.environ.get("MAX_CONTENT_LENGTH", 50 * 1024 * 1024))

DB_URL = os.environ.get("DATABASE_URL")
if DB_URL and DB_URL.startswith("postgres://"):
    DB_URL = DB_URL.replace("postgres://", "postgresql://", 1)
app.config["SQLALCHEMY_DATABASE_URI"] = DB_URL or "sqlite:///oneplaceqr.db"
app.config["SQLALCHEMY_ENGINE_OPTIONS"] = {"pool_pre_ping": True}
app.config["SQLALCHEMY_TRACK_MODIFICATIONS"] = False

UPLOAD_DIR = os.environ.get("UPLOAD_DIR", "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

db = SQLAlchemy(app)
DEFAULT_OWNER_EMAIL = os.environ.get("DEFAULT_OWNER_EMAIL", "free@oneplaceqr.local")

# ── Models ───────────────────────────────────────────────────────────────────
class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    email = db.Column(db.String(255), unique=True, nullable=False)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)

class QRCode(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    owner_id = db.Column(db.Integer, db.ForeignKey("user.id"), nullable=False)
    name = db.Column(db.String(200), nullable=False)
    slug = db.Column(db.String(200), unique=True, nullable=False, index=True)
    qr_type = db.Column(db.String(50), nullable=False)
    data_json = db.Column(db.Text, nullable=False)
    fg = db.Column(db.String(20), default="#000000")
    bg = db.Column(db.String(20), default="#ffffff")
    logo_path = db.Column(db.String(400))
    mode = db.Column(db.String(20), default="redirect")  # redirect | landing
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)

class ScanEvent(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    code_id = db.Column(db.Integer, db.ForeignKey("qr_code.id"), nullable=False)
    ts = db.Column(db.DateTime, default=dt.datetime.utcnow, index=True)
    ip = db.Column(db.String(64))
    ua = db.Column(db.Text)
    referer = db.Column(db.Text)
    path = db.Column(db.String(255))

with app.app_context():
    db.create_all()
    owner = User.query.filter_by(email=DEFAULT_OWNER_EMAIL).first()
    if not owner:
        owner = User(email=DEFAULT_OWNER_EMAIL)
        db.session.add(owner); db.session.commit()
    DEFAULT_OWNER_ID = owner.id

# ── Templates ────────────────────────────────────────────────────────────────
BASE = """
<!doctype html><html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>{{ title or "OnePlaceQR" }}</title>
<style>
:root{--fg:#0b0b0b;--bg:#fafafa;--card:#fff;--bd:#ddd;--muted:#666;--brand:#111}
*{box-sizing:border-box} body{margin:0;font-family:system-ui,Arial;background:var(--bg);color:var(--fg)}
header{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;background:#fff;border-bottom:1px solid var(--bd)}
nav a{margin-left:14px;text-decoration:none;color:var(--fg)}
.container{max-width:1100px;margin:24px auto;padding:0 16px}
.card{background:var(--card);border:1px solid var(--bd);border-radius:12px;padding:14px}
input,select,textarea{width:100%;padding:10px;border:1px solid #ccc;border-radius:10px}
button{padding:10px 14px;border:0;border-radius:10px;background:var(--brand);color:#fff;cursor:pointer}
.muted{color:var(--muted)} .row{display:flex;gap:10px;flex-wrap:wrap}
.hint{font-size:12px;color:#666;margin-top:4px}
img.preview{max-width:100%;height:auto;border-radius:12px;border:1px solid #eee}
</style>
</head><body>
<header>
  <nav>
    <a href="{{ url_for('create') }}">Create QR Code</a>
    <a href="{{ url_for('about') }}">How it works</a>
  </nav>
</header>
<div class="container">
  {% with messages = get_flashed_messages() %}
    {% if messages %}<div class="card">{% for m in messages %}<div>{{ m }}</div>{% endfor %}</div>{% endif %}
  {% endwith %}
  {{ body|safe }}
</div>
</body></html>
"""

CREATE = """
<div class="card">
  <h2>Create your QR code</h2>
  <ol class="muted" style="margin-top:6px"><li>Choose content</li><li>Customize</li><li>Download</li></ol>
</div>

<form class="card" method="POST" enctype="multipart/form-data">
  <div class="row">
    <div style="flex:2;min-width:280px">
      <label>QR Name</label>
      <input name="name" placeholder="Campaign / Menu / Profile" required>

      <label style="margin-top:10px">QR Type</label>
      <select name="qr_type" id="qr_type" required>
        {% for v,label in types %}<option value="{{ v }}" {{ "selected" if v==qr_type else "" }}>{{ label }}</option>{% endfor %}
      </select>

      <label style="margin-top:10px">URL (optional)</label>
      <input name="url" placeholder="https://example.com/your-content">
      <div class="hint">Paste a URL or upload a file below.</div>

      <!-- ✅ ALWAYS VISIBLE FILE PICKER -->
      <label style="margin-top:10px">Upload file </label>
      <input type="file" name="file" id="file_input">

      <!-- Images extras (when Images is selected) -->
      <div id="images-extra" style="display:none; margin-top:10px">
        <label>Image URLs (comma-separated, optional)</label>
        <textarea name="images" rows="3" placeholder="https://...jpg, https://...png"></textarea>
        <label style="margin-top:8px">Upload multiple images (optional)</label>
        <input type="file" name="file_images" id="file_images" accept="image/*" multiple>
      </div>

      <label style="margin-top:10px">Mode</label>
      <select name="mode">
        <option value="redirect">Dynamic redirect</option>
        <option value="landing">Hosted landing page</option>
      </select>
    </div>

    <div style="flex:1;min-width:260px">
      <label>Foreground (hex)</label><input name="fg" value="#000000">
      <label>Background (hex)</label><input name="bg" value="#ffffff">
      <label>Center Logo (optional)</label>
      <input type="file" name="logo" accept="image/*">
      <div class="hint">PNG/JPG/SVG. Small square works best.</div>
    </div>
  </div>

  <div style="margin-top:12px"><button type="submit">Save & Generate</button></div>
</form>

<script>
  const typeSel = document.getElementById('qr_type');
  const fileInput = document.getElementById('file_input');
  const imagesExtra = document.getElementById('images-extra');

  function updateAccept() {
    const t = typeSel.value;
    imagesExtra.style.display = (t === 'images') ? '' : 'none';
    if (t === 'pdf')        fileInput.setAttribute('accept','application/pdf');
    else if (t === 'video') fileInput.setAttribute('accept','video/*');
    else if (t === 'mp3')   fileInput.setAttribute('accept','audio/*');
    else if (t === 'images')fileInput.setAttribute('accept','image/*');
    else                    fileInput.removeAttribute('accept'); // any file
  }
  typeSel.addEventListener('change', updateAccept);
  updateAccept();
</script>
"""

ABOUT = """
<div class="card">
  <h2>How it works</h2>
  <ul>
    <li>Use a URL <b>or</b> the always-visible <em>Upload file</em> to host content.</li>
    <li>For <b>Images</b>, you can add multiple URL images and/or upload multiple images.</li>
    <li>Customize colors and an optional center logo.</li>
    <li>Click <em>Save & Generate</em>, then download PNG/JPG/WEBP/SVG/EPS/PDF.</li>
  </ul>
</div>
"""

VIEW = """
<div class="card">
  <h3>{{ code.name }}</h3>
  <p class="muted">{{ code.qr_type }} · {{ code.mode }} · slug: {{ code.slug }}</p>
  <img class="preview" src="{{ url_for('download', slug=code.slug, fmt='png') }}" alt="QR">
  <div style="margin-top:10px">
    <a href="{{ url_for('download', slug=code.slug, fmt='png') }}">PNG</a> ·
    <a href="{{ url_for('download', slug=code.slug, fmt='jpg') }}">JPG</a> ·
    <a href="{{ url_for('download', slug=code.slug, fmt='webp') }}">WEBP</a> ·
    <a href="{{ url_for('download', slug=code.slug, fmt='svg') }}">SVG</a> ·
    <a href="{{ url_for('download', slug=code.slug, fmt='eps') }}">EPS</a> ·
    <a href="{{ url_for('download', slug=code.slug, fmt='pdf') }}">PDF</a> ·
    <a href="{{ url_for('scan', slug=code.slug) }}" target="_blank">Scan URL</a>
  </div>
</div>
"""

LANDING = """
<div class="container" style="max-width:720px;margin:24px auto">
  <div class="card"><h1>{{ title }}</h1><div class="muted">{{ subtitle }}</div><div style="margin-top:10px">{{ body|safe }}</div></div>
</div>
"""

# ── Upload helpers & QR generation ───────────────────────────────────────────
ALLOWED_IMG   = {"png","jpg","jpeg","gif","webp","bmp","svg"}
ALLOWED_PDF   = {"pdf"}
ALLOWED_AUDIO = {"mp3","wav","m4a","aac","ogg"}
ALLOWED_VIDEO = {"mp4","webm","mov","mkv"}
ALLOWED_DOCS  = {"txt","csv","md","doc","docx","ppt","pptx","xls","xlsx","rtf"}
ALLOWED_ARCH  = {"zip","rar","7z","gz","tar"}
ALLOWED_MISC  = {"apk"}
ALLOWED_ANY   = ALLOWED_IMG | ALLOWED_PDF | ALLOWED_AUDIO | ALLOWED_VIDEO | ALLOWED_DOCS | ALLOWED_ARCH | ALLOWED_MISC

def _safe_ext(fn): return os.path.splitext(fn)[1].lower().replace(".", "")

@app.route("/u/<path:filename>")
def uploads(filename):
    return send_from_directory(UPLOAD_DIR, filename, as_attachment=False)

def save_upload(file_storage, allowed_ext):
    if not file_storage or file_storage.filename == "": return None
    ext = _safe_ext(file_storage.filename)
    if ext not in allowed_ext: raise ValueError(f"Unsupported file type: .{ext}")
    fname = f"{secrets.token_hex(8)}.{ext}"
    path = os.path.join(UPLOAD_DIR, fname)
    file_storage.save(path)
    # Downscale large raster images
    try:
        if ext in ALLOWED_IMG and ext != "svg":
            im = Image.open(path); im.thumbnail((3000, 3000)); im.save(path)
    except Exception:
        pass
    return url_for("uploads", filename=fname, _external=True)

def save_multi(files, allowed_ext):
    return [save_upload(f, allowed_ext) for f in files if f and f.filename]

def qr_png_pil(data, fg="#000000", bg="#ffffff", logo_path=None):
    qr = qrcode.QRCode(version=None, error_correction=qrcode.constants.ERROR_CORRECT_M, box_size=10, border=4)
    qr.add_data(data); qr.make(fit=True)
    img = qr.make_image(fill_color=fg, back_color=bg).convert("RGB")
    if logo_path:
        try:
            logo = Image.open(logo_path).convert("RGBA")
            qr_w, qr_h = img.size
            target = int(qr_w * 0.2)
            logo.thumbnail((target, target))
            pad = Image.new("RGBA", logo.size, (255,255,255,255))
            pad.paste(logo, (0,0), logo)
            img.paste(pad.convert("RGB"), ((qr_w - logo.width)//2, (qr_h - logo.height)//2))
        except Exception:
            pass
    return img

def qr_svg_text(data):
    img = qrcode.make(data, image_factory=qsvg.SvgImage)
    buf = io.BytesIO(); img.save(buf); return buf.getvalue().decode()

# ── Routes ───────────────────────────────────────────────────────────────────
@app.route("/")
def home():
    return redirect(url_for("create"))

def render_page(title, body_html):
    return render_template_string(BASE, title=title, body=body_html)

@app.route("/create", methods=["GET","POST"])
def create():
    qr_type = request.args.get("qr_type","website")
    if request.method == "POST":
        name = request.form["name"].strip()
        qr_type = request.form["qr_type"]

        # Generic file (always visible)
        up_url = None
        uploaded_file = request.files.get("file")
        if uploaded_file and uploaded_file.filename:
            up_url = save_upload(uploaded_file, ALLOWED_ANY)

        # Extra multi-image field for Images
        images_extra = save_multi(request.files.getlist("file_images"), ALLOWED_IMG) if request.files.getlist("file_images") else []

        url_val = (request.form.get("url") or "").strip() or None

        # Build payload per type
        def payload_for(t: str):
            if t == "images":
                url_list = [u.strip() for u in (request.form.get("images","").split(",")) if u.strip()]
                gen_is_img = False
                if up_url and uploaded_file:
                    gen_is_img = _safe_ext(uploaded_file.filename) in ALLOWED_IMG
                imgs = (url_list or []) + images_extra + ([up_url] if (up_url and gen_is_img) else [])
                return {"images": imgs}
            if t in ("website","pdf","video","app","facebook","instagram","whatsapp","coupon","mp3"):
                return {"url": up_url or url_val}
            if t in ("links","social"):
                rows = [l.strip() for l in request.form.get("links","").splitlines() if l.strip()]
                links = []
                for row in rows:
                    if "|" in row:
                        label, u = row.split("|",1)
                        links.append({"label": label.strip(), "url": u.strip()})
                return {"links": links}
            if t == "business":
                return {"name": request.form.get("biz_name"),
                        "website": request.form.get("website"),
                        "desc": request.form.get("desc")}
            if t == "wifi":
                return {"ssid": request.form.get("ssid"),
                        "password": request.form.get("password"),
                        "security": request.form.get("security","WPA")}
            if t == "vcard":
                fn, title, email, phone, org, website = (
                    request.form.get("fn",""), request.form.get("title",""),
                    request.form.get("email",""), request.form.get("phone",""),
                    request.form.get("org",""), request.form.get("website","")
                )
                vcard = f"""BEGIN:VCARD
VERSION:3.0
N:{fn}
FN:{fn}
TITLE:{title}
ORG:{org}
TEL;TYPE=CELL:{phone}
EMAIL;TYPE=INTERNET:{email}
URL:{website}
END:VCARD"""
                return {"fn": fn, "title": title, "email": email, "phone": phone, "org": org,
                        "website": website, "vcard_text": vcard}
            return {"url": up_url or url_val}

        try:
            payload = payload_for(qr_type)
        except ValueError as e:
            flash(str(e)); return redirect(url_for("create", qr_type=qr_type))

        fg = request.form.get("fg","#000000"); bg = request.form.get("bg","#ffffff")
        mode = request.form.get("mode","redirect")

        # Optional center logo overlay
        logo_path = None
        if request.files.get("logo") and request.files["logo"].filename:
            lf = request.files["logo"]; ext = _safe_ext(lf.filename)
            if ext in ALLOWED_IMG:
                fname = f"logo_{secrets.token_hex(8)}.{ext}"
                logo_path = os.path.join(UPLOAD_DIR, fname); lf.save(logo_path)
                try:
                    if ext != "svg":
                        im = Image.open(logo_path); im.thumbnail((800,800)); im.save(logo_path)
                except Exception: pass
            else:
                flash("Unsupported logo type.")

        base_slug = slugify(name) or secrets.token_hex(3)
        slug = base_slug; i = 1
        while QRCode.query.filter_by(slug=slug).first():
            i += 1; slug = f"{base_slug}-{i}"

        qr_row = QRCode(owner_id=DEFAULT_OWNER_ID, name=name, slug=slug, qr_type=qr_type,
                        data_json=json.dumps(payload), fg=fg, bg=bg, logo_path=logo_path, mode=mode)
        db.session.add(qr_row); db.session.commit()
        return redirect(url_for("view_code", slug=slug))

    types = [("website","Website"), ("pdf","PDF"), ("images","Images"), ("video","Video"),
             ("wifi","WiFi"), ("business","Business"), ("vcard","vCard"), ("mp3","MP3"),
             ("app","Apps"), ("links","List of Links"), ("coupon","Coupon"),
             ("facebook","Facebook"), ("instagram","Instagram"), ("social","Social Media"),
             ("whatsapp","WhatsApp")]
    return render_template_string(
        BASE, title="Create QR Code",
        body=render_template_string(CREATE, types=types, qr_type=qr_type)
    )

@app.route("/about")
def about():
    return render_template_string(BASE, title="How it works", body=ABOUT)

@app.route("/code/<slug>")
def view_code(slug):
    code = QRCode.query.filter_by(slug=slug, owner_id=DEFAULT_OWNER_ID).first_or_404()
    return render_template_string(BASE, title=code.name, body=render_template_string(VIEW, code=code))

# ── Scan / Landing / Download ────────────────────────────────────────────────
def log_scan(code_id):
    try:
        db.session.add(ScanEvent(
            code_id=code_id,
            ip=request.headers.get("X-Forwarded-For", request.remote_addr),
            ua=request.headers.get("User-Agent"),
            referer=request.headers.get("Referer"),
            path=request.path
        )); db.session.commit()
    except Exception:
        db.session.rollback()

def code_target(qr: QRCode):
    payload = json.loads(qr.data_json)
    if qr.mode == "landing": return url_for("landing_page", slug=qr.slug, _external=True)
    t = qr.qr_type
    if t in ("website","pdf","video","app","facebook","instagram","whatsapp","coupon","mp3"):
        return payload.get("url")
    if t == "links":
        links = payload.get("links", []); return links[0]["url"] if links else "#"
    if t == "images":
        imgs = payload.get("images", []); return imgs[0] if imgs else "#"
    if t == "business":
        return payload.get("website") or "#"
    if t == "wifi":
        s, p, sec = payload.get("ssid",""), payload.get("password",""), payload.get("security","WPA")
        return f"WIFI:T:{sec};S:{s};P:{p};;"
    if t == "vcard":
        return payload.get("vcard_text")
    return payload.get("url") or "#"

@app.route("/s/<slug>")
def scan(slug):
    qr = QRCode.query.filter_by(slug=slug).first_or_404()
    target = code_target(qr); log_scan(qr.id)
    if target and (target.startswith("WIFI:") or target.startswith("BEGIN:VCARD")):
        return f"<pre>{target}</pre>"
    if qr.mode == "landing": return redirect(url_for("landing_page", slug=slug))
    return redirect(target or "#")

@app.route("/l/<slug>")
def landing_page(slug):
    qr = QRCode.query.filter_by(slug=slug).first_or_404()
    payload = json.loads(qr.data_json); log_scan(qr.id)
    t = qr.qr_type; title, subtitle = qr.name, t; body = "<div class='muted'>No content</div>"
    if t in ("website","pdf","video","app","facebook","instagram","whatsapp","coupon","mp3"):
        url = payload.get("url","#")
        body = f"<p><a href='{url}' target='_blank'>{url}</a></p>"
        if t=="mp3": body += f"<audio controls src='{url}' style='width:100%'></audio>"
        if t=="video" and ('youtube' in url.lower() or url.lower().endswith(('.mp4','.webm','.mov','.mkv'))):
            body += f"<div style='margin-top:10px'><video controls style='width:100%'><source src='{url}'></video></div>"
        if t=="pdf" and url.lower().endswith('.pdf'):
            body += f"<iframe src='{url}' style='width:100%;height:70vh;border:1px solid #eee;border-radius:10px'></iframe>"
    elif t=="images":
        imgs = payload.get("images",[]); body = "<div class='row'>" + "".join([f"<img src='{u}' style='max-width:220px;border-radius:10px;border:1px solid #eee'>" for u in imgs]) + "</div>"
    elif t in ("links","social"):
        links = payload.get("links",[]); body = "<ul>" + "".join([f"<li><a href='{l['url']}' target='_blank'>{l['label']}</a></li>" for l in links]) + "</ul>"
    elif t=="business":
        bn, desc, web = payload.get("name","Business"), payload.get("desc",""), payload.get("website","#")
        body = f"<h3>{bn}</h3><p>{desc}</p><p><a href='{web}' target='_blank'>{web}</a></p>"
    elif t=="wifi":
        s, p, sec = payload.get("ssid",""), payload.get("password",""), payload.get("security","WPA"); body = f"<pre>WIFI:T:{sec};S:{s};P:{p};;</pre>"
    elif t=="vcard":
        body = f"<pre>{payload.get('vcard_text','')}</pre>"
    return render_template_string(LANDING, title=title, subtitle=subtitle, body=body)

def _qr_bytes(qr: QRCode, fmt: str):
    data = url_for("scan", slug=qr.slug, _external=True)
    if fmt.lower() == "svg":
        svg = qr_svg_text(data); return svg.encode("utf-8"), "image/svg+xml"
    img = qr_png_pil(data, fg=qr.fg, bg=qr.bg, logo_path=qr.logo_path)
    buf = io.BytesIO(); f = fmt.lower()
    if f == "jpg":  img.save(buf, "JPEG", quality=95);  return buf.getvalue(), "image/jpeg"
    if f == "webp": img.save(buf, "WEBP");              return buf.getvalue(), "image/webp"
    if f == "pdf":  img.save(buf, "PDF");               return buf.getvalue(), "application/pdf"
    if f == "eps":  img.save(buf, "EPS");               return buf.getvalue(), "application/postscript"
    img.save(buf, "PNG");                                return buf.getvalue(), "image/png"

@app.route("/d/<slug>.<fmt>")
def download(slug, fmt):
    qr = QRCode.query.filter_by(slug=slug).first_or_404()
    payload, mime = _qr_bytes(qr, fmt)
    return send_file(io.BytesIO(payload), mimetype=mime, as_attachment=True,
                     download_name=f"{qr.slug}.{fmt.lower()}")

@app.route("/health")
def health(): return "ok", 200

# Allow embedding in your main site
@app.after_request
def csp(resp):
    resp.headers["Content-Security-Policy"] = (
        "default-src 'self'; img-src 'self' data:; media-src 'self' data:; "
        "style-src 'self' 'unsafe-inline'; script-src 'self'; "
        "frame-ancestors 'self' https://oneplacepdf.com https://www.oneplacepdf.com"
    )
    try: resp.headers.pop("X-Frame-Options", None)
    except Exception: pass
    return resp

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8000")), debug=True)

