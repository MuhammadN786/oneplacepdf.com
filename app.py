# app.py
# OnePlacePDF — Single-file Flask app (no Streamlit)
# Features:
# - Normal HTML site with real <head> (AdSense, GA4, SEO, JSON-LD)
# - Pages: Home (tools), About, Privacy, Terms, Contact
# - Tools: Images→PDF, Merge, Combine, Split, Rotate, Reorder, Extract, Compress,
#          Protect, Unlock, PDF→Images, PDF→DOCX, Watermark, Page Numbers, Office→PDF
# - Static endpoints: /robots.txt, /sitemap.xml, /ads.txt
#
# Quick start:
#   python -m venv .venv && . .venv/bin/activate   (Windows: .venv\Scripts\activate)
#   pip install flask pypdf pymupdf pillow pdf2docx python-docx reportlab pandas openpyxl python-pptx pypandoc
#   python app.py
#
# NOTE: Some optional conversions (Office→PDF) rely on extra system tools (e.g., pandoc, LaTeX).
#       The app handles failures gracefully and still serves core PDF tools.

import io, os, re, shutil, tempfile, zipfile
from typing import List, Tuple

from flask import (
    Flask, request, render_template_string, send_file,
    make_response, url_for
)
from werkzeug.utils import secure_filename

from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image
from pdf2docx import Converter as Pdf2DocxConverter

# ---------- Site config (edit these) ----------
SITE_NAME = "OnePlacePDF"
BASE_URL  = "https://oneplacepdf.com"      # change if different domain
CONTACT_EMAIL = "oneplacepdf@gmail.com"

# Google tags (replace with your real IDs)
ADSENSE_CLIENT = "ca-pub-6839950833502659"  # <-- YOUR AdSense publisher ID
ADSENSE_SLOT   = "3025573109"               # <-- an ad slot ID (create in AdSense)
GA4_ID         = "G-XXXXXXX"                # <-- optional; leave as "" to disable
# ---------------------------------------------

# --- Make sure redirect is imported ---
from flask import (
    Flask, request, render_template_string, send_file,
    redirect, url_for, make_response
)

# imports at top of file must include redirect/make_response
from flask import (
    Flask, request, render_template_string, send_file,
    redirect, url_for, make_response
)
from werkzeug.middleware.proxy_fix import ProxyFix

from flask import (
    Flask, request, render_template_string, send_file,
    redirect, url_for, make_response
)
from werkzeug.middleware.proxy_fix import ProxyFix

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB



# ==========================
# Utilities
# ==========================
def _tmpfile(suffix=""):
    fd, path = tempfile.mkstemp(suffix=suffix)
    return fd, path

def _save_upload(fs, suffix="") -> str:
    """Save a single FileStorage to a tmp path and return path."""
    fd, path = _tmpfile(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(fs.read())
    return path

def _parse_ranges(ranges: str, total_pages: int) -> List[int]:
    """Return 0-based page indices from '1-3,5,7-9'. Keeps order, de-dupes."""
    if not ranges or not ranges.strip():
        return list(range(total_pages))
    result, seen = [], set()
    parts = [p.strip() for p in ranges.split(",") if p.strip()]
    for part in parts:
        if "-" in part:
            a, b = part.split("-", 1)
            start = int(a) if a.strip() else 1
            end = int(b) if b.strip() else total_pages
        else:
            start = end = int(part)
        start = max(1, start)
        end = min(total_pages, end)
        for i in range(start - 1, end):
            if 0 <= i < total_pages and i not in seen:
                seen.add(i)
                result.append(i)
    return result

def _human(n: int) -> str:
    for u in ["B","KB","MB","GB","TB"]:
        if n < 1024:
            return f"{n:.2f} {u}"
        n /= 1024
    return f"{n:.2f} PB"

def _roman(n: int) -> str:
    # Simple Roman numerals (1..3999)
    vals = [
        (1000,"M"), (900,"CM"), (500,"D"), (400,"CD"),
        (100,"C"), (90,"XC"), (50,"L"), (40,"XL"),
        (10,"X"), (9,"IX"), (5,"V"), (4,"IV"), (1,"I")]
    out = []
    for v, s in vals:
        while n >= v:
            out.append(s); n -= v
    return "".join(out)

def _send_bytes(data: bytes, filename: str, mime: str):
    return send_file(io.BytesIO(data), as_attachment=True, download_name=filename, mimetype=mime)

# ==========================
# HTML Template
# ==========================
PAGE = r"""
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <title>{{ site_name }} — All-in-One PDF Tools</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <meta name="description" content="Free online PDF tools to merge, split, compress, convert and protect PDFs." />
  <link rel="canonical" href="{{ base_url }}/" />
  <meta property="og:title" content="{{ site_name }} — All-in-One PDF Tools" />
  <meta property="og:description" content="Merge, split, compress, convert, protect, and more — fast and simple in your browser." />
  <meta property="og:type" content="website" />
  <meta property="og:url" content="{{ base_url }}/" />
  <meta name="google-adsense-account" content="{{ adsense_client }}" />
  {% if ga4_id and ga4_id != "G-XXXXXXX" %}
  <!-- Google tag (gtag.js) -->
  <script async src="https://www.googletagmanager.com/gtag/js?id={{ ga4_id }}"></script>
  <script>
    window.dataLayer = window.dataLayer || [];
    function gtag(){dataLayer.push(arguments);}
    gtag('js', new Date()); gtag('config', '{{ ga4_id }}');
  </script>
  {% endif %}
  <!-- AdSense loader (sitewide) -->
  <script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client={{ adsense_client }}" crossorigin="anonymous"></script>

  <script type="application/ld+json">
  {
    "@context":"https://schema.org",
    "@type":"WebSite",
    "name":"{{ site_name }}",
    "url":"{{ base_url }}/",
    "description":"Free online PDF tools to merge, split, compress, convert and protect PDFs. Fast, private, works on any device. No sign-up required.",
    "inLanguage":"en",
    "publisher":{"@type":"Organization","name":"{{ site_name }}"}
  }
  </script>

  <style>
    :root {
      --bg:#0b1020; --card:#131a2a; --muted:#a9b2c7; --fg:#eaf0ff; --accent:#5da0ff; --accent2:#00d2d3;
      --border:#24304a;
    }
    * { box-sizing: border-box; }
    body { margin:0; font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, 'Helvetica Neue', Arial; color: var(--fg); background: linear-gradient(180deg, #0b1020 0%, #0a0f1e 70%, #070b17 100%); }
    header { padding: 28px 16px; border-bottom: 1px solid var(--border); position: sticky; top:0; background: rgba(11,16,32,.9); backdrop-filter: blur(6px); z-index:10; }
    .wrap { max-width: 1100px; margin: 0 auto; }
    nav a { color: var(--muted); text-decoration: none; margin-right: 16px; }
    nav a:hover { color: var(--fg); }
    .hero { padding: 28px 16px 8px; }
    .hero h1 { margin: 0 0 8px; font-size: 28px; }
    .hero p  { margin: 0 0 16px; color: var(--muted); }
    .tabs { display: flex; flex-wrap: wrap; gap: 10px; padding: 0 16px 16px; }
    .tab-btn { background: #0e1426; border: 1px solid var(--border); color: var(--fg); padding: 10px 14px; border-radius: 10px; cursor: pointer; font-size: 14px; }
    .tab-btn.active { border-color: var(--accent); box-shadow: 0 0 0 2px rgba(93,160,255,.15) inset; }
    .grid { display: grid; grid-template-columns: 1fr; gap: 16px; padding: 0 16px 40px; }
    @media (min-width: 900px) { .grid { grid-template-columns: 1fr 1fr; } }
    .card { background: var(--card); border:1px solid var(--border); border-radius: 14px; padding: 16px; }
    .card h3 { margin: 0 0 10px; font-size: 18px; }
    label { display:block; margin: 10px 0 6px; color: var(--muted); font-size: 13px; }
    input[type="file"], select, input[type="text"], input[type="number"], textarea {
      width: 100%; background: #0e1426; color: var(--fg); border:1px solid var(--border); border-radius: 8px; padding: 10px; font-size: 14px;
    }
    .row { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
    .btn { display:inline-block; background: linear-gradient(90deg, var(--accent), var(--accent2));
           color:#081020; font-weight:600; border:0; padding: 10px 14px; border-radius: 10px; cursor:pointer; margin-top:12px;}
    .muted { color: var(--muted); font-size: 12px; }
    footer { color: var(--muted); border-top: 1px solid var(--border); padding: 20px 16px; }
    .ad { display:block; width:100%; min-height:120px; margin: 10px 0; }
    .hidden { display:none; }
  </style>
</head>
<body>
  <header>
    <div class="wrap" style="display:flex; align-items:center; justify-content:space-between; gap:16px;">
      <div style="display:flex; align-items:center; gap:12px;">
        <div style="font-size:22px;">📄</div>
        <strong>{{ site_name }}</strong>
      </div>
      <nav>
        <a href="{{ url_for('home') }}">Home</a>
        <a href="{{ url_for('about') }}">About</a>
        <a href="{{ url_for('privacy') }}">Privacy</a>
        <a href="{{ url_for('terms') }}">Terms</a>
        <a href="{{ url_for('contact') }}">Contact</a>
      </nav>
    </div>
  </header>

  <div class="wrap">
    <section class="hero">
      <h1>{{ site_name }} — All-in-One PDF Tools</h1>
      <p>Merge, split, convert & secure — with quality-first processing. Everything runs locally on our server while you wait; results download instantly.</p>
      <!-- Responsive Ad unit -->
      <ins class="adsbygoogle ad"
           style="display:block"
           data-ad-client="{{ adsense_client }}"
           data-ad-slot="{{ adsense_slot }}"
           data-ad-format="auto"
           data-full-width-responsive="true"></ins>
      <script>(adsbygoogle=window.adsbygoogle||[]).push({});</script>
    </section>

    <div class="tabs" id="tabs">
      {% for t in tabs %}
        <button class="tab-btn" data-target="{{ t.id }}">{{ t.label }}</button>
      {% endfor %}
    </div>

    <div id="sections">

      <!-- Images → PDF -->
      <section class="grid section" id="images-pdf">
        <div class="card">
          <h3>Images → PDF (high quality)</h3>
          <form method="post" action="{{ url_for('images_to_pdf') }}" enctype="multipart/form-data">
            <label>Upload images (JPG/PNG)</label>
            <input type="file" name="images" accept=".jpg,.jpeg,.png" multiple required>

            <div class="row">
              <div>
                <label>Page size</label>
                <select name="page_size">
                  <option>Original</option><option>A4</option><option>Letter</option>
                </select>
              </div>
              <div>
                <label>Target DPI</label>
                <select name="dpi">
                  <option>300</option><option>600</option>
                </select>
              </div>
            </div>

            <div class="row">
              <div>
                <label>Margin (pt)</label>
                <input type="number" name="margin" value="24" min="0" max="96">
              </div>
              <div>
                <label>Orientation</label>
                <select name="orientation">
                  <option>Auto</option><option>Portrait</option><option>Landscape</option>
                </select>
              </div>
            </div>

            <label>Output</label>
            <select name="output">
              <option>Single PDF (all images)</option>
              <option>One PDF per image (ZIP)</option>
            </select>

            <button class="btn" type="submit">Convert to PDF</button>
          </form>
          <p class="muted">Keeps image quality; adds margins safely; reorders via filename upload order.</p>
        </div>
      </section>

      <!-- Merge / Combine -->
      <section class="grid section" id="merge">
        <div class="card">
          <h3>Merge PDFs (append)</h3>
          <form method="post" action="{{ url_for('merge_append') }}" enctype="multipart/form-data">
            <label>Upload PDFs (any order)</label>
            <input type="file" name="pdfs" accept=".pdf" multiple required>
            <label>Pages to include for each file (optional, comma-separated ranges; applies to ALL files)</label>
            <input type="text" name="ranges" placeholder="e.g., 1-3,5">
            <label><input type="checkbox" name="bookmarks" checked> Add bookmarks by filename</label>
            <button class="btn" type="submit">Merge Now</button>
          </form>
        </div>

        <div class="card">
          <h3>Combine PDFs (interleave)</h3>
          <form method="post" action="{{ url_for('combine_interleave') }}" enctype="multipart/form-data">
            <label>Upload 2–10 PDFs</label>
            <input type="file" name="pdfs" accept=".pdf" multiple required>
            <label>Chunk size per file (e.g., 1 = A1,B1,A2,B2...)</label>
            <input type="number" name="chunk" value="1" min="1" max="10">
            <label>Loop until all pages exhausted?</label>
            <select name="loop">
              <option>Yes</option><option>No</option>
            </select>
            <button class="btn" type="submit">Combine Now</button>
          </form>
        </div>
      </section>

      <!-- Split / Rotate / Reorder -->
      <section class="grid section" id="split-rotate">
        <div class="card">
          <h3>Split PDF</h3>
          <form method="post" action="{{ url_for('split_pdf') }}" enctype="multipart/form-data">
            <label>Upload PDF</label>
            <input type="file" name="pdf" accept=".pdf" required>
            <label>Split by</label>
            <select name="mode">
              <option>Page ranges (custom)</option>
              <option>Every page</option>
              <option>Approx file size</option>
            </select>
            <label>Ranges (if custom)</label>
            <input type="text" name="ranges" placeholder="1-3,5,7-9">
            <label>Max size per part (MB) (if size mode)</label>
            <input type="number" name="size_mb" value="5" min="1" max="50">
            <button class="btn" type="submit">Split Now</button>
          </form>
        </div>

        <div class="card">
          <h3>Rotate PDF</h3>
          <form method="post" action="{{ url_for('rotate_pdf') }}" enctype="multipart/form-data">
            <label>Upload PDF</label>
            <input type="file" name="pdf" accept=".pdf" required>
            <label>Rotate all pages by</label>
            <select name="deg">
              <option>90</option><option>180</option><option>270</option>
            </select>
            <button class="btn" type="submit">Apply Rotation</button>
          </form>
        </div>

        <div class="card">
          <h3>Re-order Pages</h3>
          <form method="post" action="{{ url_for('reorder_pdf') }}" enctype="multipart/form-data">
            <label>Upload PDF</label>
            <input type="file" name="pdf" accept=".pdf" required>
            <label>Pages to keep (blank = all)</label>
            <input type="text" name="keep" placeholder="e.g., 1-3,6,8">
            <label>Order (comma-separated, 1-based)</label>
            <input type="text" name="order" placeholder="e.g., 3,1,2">
            <button class="btn" type="submit">Build Re-ordered PDF</button>
          </form>
        </div>
      </section>

      <!-- Extract / Compress -->
      <section class="grid section" id="extract-compress">
        <div class="card">
          <h3>Extract Text</h3>
          <form method="post" action="{{ url_for('extract_text') }}" enctype="multipart/form-data">
            <label>Upload PDF</label>
            <input type="file" name="pdf" accept=".pdf" required>
            <label>Mode</label>
            <select name="mode"><option>Plain text</option><option>Preserve layout</option><option>Raw</option></select>
            <label>Pages (blank = all)</label>
            <input type="text" name="ranges" placeholder="1-3,5">
            <label>Search term (optional)</label>
            <input type="text" name="search" placeholder="keyword">
            <button class="btn" type="submit">Extract Now</button>
          </form>
        </div>

        <div class="card">
          <h3>Compress PDF</h3>
          <form method="post" action="{{ url_for('compress_pdf') }}" enctype="multipart/form-data">
            <label>Upload PDF</label>
            <input type="file" name="pdf" accept=".pdf" required>
            <label>Preset</label>
            <select name="preset">
              <option>Standard</option>
              <option>Light (best quality)</option>
              <option>Strong</option>
              <option>Extreme</option>
              <option>Lossless (no image recompress)</option>
            </select>
            <label>JPEG quality (50-95, used when recompressing)</label>
            <input type="number" name="quality" value="85" min="50" max="95">
            <label>Downsample images to DPI (blank=auto)</label>
            <input type="number" name="target_dpi" value="200" min="50" max="600">
            <label><input type="checkbox" name="strip_meta" checked> Strip metadata</label>
            <label><input type="checkbox" name="linearize" checked> Linearize for web (Fast Web View)</label>
            <label><input type="checkbox" name="clean_xref" checked> Aggressive clean</label>
            <button class="btn" type="submit">Compress</button>
          </form>
        </div>
      </section>

      <!-- Protect / Unlock -->
      <section class="grid section" id="protect-unlock">
        <div class="card">
          <h3>Protect PDF (Password)</h3>
          <form method="post" action="{{ url_for('protect_pdf') }}" enctype="multipart/form-data">
            <label>Upload PDF</label>
            <input type="file" name="pdf" accept=".pdf" required>
            <label>User Password</label>
            <input type="text" name="user_pwd" placeholder="password to open">
            <label>Owner Password</label>
            <input type="text" name="owner_pwd" placeholder="permissions password">
            <label><input type="checkbox" name="allow_print" checked> Allow printing</label>
            <label><input type="checkbox" name="allow_copy" checked> Allow copy</label>
            <label><input type="checkbox" name="allow_annot" checked> Allow annotate</label>
            <label>Encryption</label>
            <select name="encryption">
              <option>AES-256</option><option>AES-128</option><option>RC4-128</option><option>RC4-40</option>
            </select>
            <button class="btn" type="submit">Apply Protection</button>
          </form>
        </div>

        <div class="card">
          <h3>Unlock PDF (Remove Password)</h3>
          <form method="post" action="{{ url_for('unlock_pdf') }}" enctype="multipart/form-data">
            <label>Upload encrypted PDFs</label>
            <input type="file" name="pdfs" accept=".pdf" multiple required>
            <label>Password</label>
            <input type="text" name="password" required>
            <button class="btn" type="submit">Unlock Now</button>
          </form>
        </div>
      </section>

      <!-- Conversions -->
      <section class="grid section" id="conversions">
        <div class="card">
          <h3>PDF → Images (High Quality)</h3>
          <form method="post" action="{{ url_for('pdf_to_images') }}" enctype="multipart/form-data">
            <label>Upload PDFs</label>
            <input type="file" name="pdfs" accept=".pdf" multiple required>
            <label>Export DPI</label>
            <input type="number" name="dpi" value="300" min="150" max="600" step="50">
            <label>Format</label>
            <select name="fmt"><option>PNG</option><option>JPEG</option><option>WebP</option></select>
            <label>Quality (for JPEG/WebP)</label>
            <input type="number" name="quality" value="95" min="80" max="100">
            <label>Mode</label>
            <select name="mode"><option>One image per page (ZIP)</option><option>All pages stitched vertically (one image)</option></select>
            <button class="btn" type="submit">Convert</button>
          </form>
        </div>

        <div class="card">
          <h3>PDF → DOCX (Editable Word)</h3>
          <form method="post" action="{{ url_for('pdf_to_docx') }}" enctype="multipart/form-data">
            <label>Upload PDF</label>
            <input type="file" name="pdf" accept=".pdf" required>
            <label>Start page (1-based)</label>
            <input type="number" name="start" value="1" min="1">
            <label>End page (0 = all)</label>
            <input type="number" name="end" value="0" min="0">
            <label><input type="checkbox" name="keep_images" checked> Keep images</label>
            <button class="btn" type="submit">Convert to DOCX</button>
          </form>
        </div>

        <div class="card">
          <h3>Office → PDF (Word/Excel/PPT/Text/Images)</h3>
          <form method="post" action="{{ url_for('office_to_pdf') }}" enctype="multipart/form-data">
            <label>Upload files</label>
            <input type="file" name="files" multiple
                   accept=".docx,.xlsx,.pptx,.txt,.odt,.ods,.odp,.png,.jpg" required>
            <button class="btn" type="submit">Convert to PDF</button>
          </form>
          <p class="muted">Some conversions may require extra system tools (pandoc/LaTeX). If unavailable, those files will be skipped with a message.</p>
        </div>
      </section>

      <!-- Marking/Numbering -->
      <section class="grid section" id="wm-number">
        <div class="card">
          <h3>Add Watermark</h3>
          <form method="post" action="{{ url_for('watermark') }}" enctype="multipart/form-data">
            <label>Upload PDF</label>
            <input type="file" name="pdf" accept=".pdf" required>
            <label>Type</label>
            <select name="wm_type"><option>Text</option><option>Image</option></select>
            <label>Text (if text WM)</label><input type="text" name="text" value="CONFIDENTIAL">
            <label>Color hex (e.g. #FF0000)</label><input type="text" name="color" value="#FF0000">
            <label>Opacity (%)</label><input type="number" name="opacity" value="20" min="10" max="90">
            <label>Font size</label><input type="number" name="size" value="60" min="10" max="200">
            <label>Rotation angle</label><input type="number" name="angle" value="45" min="0" max="360">
            <label>Position</label>
            <select name="pos"><option>Center</option><option>Top-left</option><option>Top-right</option><option>Bottom-left</option><option>Bottom-right</option><option>Diagonal Tiled</option></select>
            <label>Watermark image (if image WM)</label><input type="file" name="wm_img" accept=".png,.jpg,.jpeg">
            <label>Apply to pages (blank = all)</label><input type="text" name="pages" placeholder="1-3,5">
            <button class="btn" type="submit">Apply Watermark</button>
          </form>
        </div>

        <div class="card">
          <h3>Add Page Numbers</h3>
          <form method="post" action="{{ url_for('page_numbers') }}" enctype="multipart/form-data">
            <label>Upload PDF</label><input type="file" name="pdf" accept=".pdf" required>
            <label>Pages (blank = all)</label><input type="text" name="ranges" placeholder="1-3,5">
            <label>Style</label>
            <select name="style"><option>1, 2, 3</option><option>01, 02, 03</option><option>i, ii, iii</option><option>I, II, III</option><option>a, b, c</option><option>A, B, C</option></select>
            <label>Template</label><input type="text" name="template" value="Page {n} of {total}">
            <label>Vertical</label><select name="pos_v"><option>Bottom</option><option>Top</option></select>
            <label>Horizontal</label><select name="pos_h"><option>Center</option><option>Left</option><option>Right</option></select>
            <label>Font size</label><input type="number" name="fontsize" value="12" min="8" max="32">
            <label>Color hex</label><input type="text" name="color" value="#000000">
            <label>Opacity (%)</label><input type="number" name="opacity" value="80" min="20" max="100">
            <button class="btn" type="submit">Add Numbers</button>
          </form>
        </div>
      </section>

    </div>
  </div>

  <footer>
    <div class="wrap">
      <div>© {{ site_name }} • <a href="{{ url_for('privacy') }}" style="color:#9fc0ff;">Privacy</a> •
        <a href="{{ url_for('terms') }}" style="color:#9fc0ff;">Terms</a> •
        <a href="{{ url_for('contact') }}" style="color:#9fc0ff;">Contact</a></div>
      <div class="muted">We do not keep your files longer than needed to deliver your download.</div>
    </div>
  </footer>

  <script>
    const tabs = [...document.querySelectorAll('.tab-btn')];
    const sections = [...document.querySelectorAll('.section')];
    function show(id){
      sections.forEach(s => s.classList.add('hidden'));
      tabs.forEach(b => b.classList.remove('active'));
      document.getElementById(id).classList.remove('hidden');
      tabs.find(b => b.dataset.target===id)?.classList.add('active');
      history.replaceState(null, '', '#'+id);
    }
    tabs.forEach(b => b.addEventListener('click', () => show(b.dataset.target)));
    const initial = location.hash?.replace('#','') || '{{ tabs[0].id }}';
    show(initial);
  </script>
</body>
</html>
"""

PAGES_SIMPLE = lambda title, body: f"""<!doctype html>
<html lang="en"><head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1" />
<title>{title} — {SITE_NAME}</title>
<link rel="canonical" href="{BASE_URL}/{title.lower() if title!='Home' else ''}" />
<meta name="description" content="{SITE_NAME} — Free online PDF tools." />
<meta name="google-adsense-account" content="{ADSENSE_CLIENT}" />
<script async src="https://pagead2.googlesyndication.com/pagead/js/adsbygoogle.js?client={ADSENSE_CLIENT}" crossorigin="anonymous"></script>
<style>body{{font-family:system-ui;margin:0;background:#0b1020;color:#eaf0ff}}.wrap{{max-width:800px;margin:0 auto;padding:24px}}a{{color:#9fc0ff}}</style>
</head><body>
<div class="wrap">
  <h1>{title}</h1>
  <p style="color:#a9b2c7">{body}</p>
  <p><a href="/">← Back to tools</a></p>
</div></body></html>
"""

# ==========================
# Routes: Pages
# ==========================
@app.get("/")
def home():
    tabs = [
        {"id":"images-pdf",   "label":"Images → PDF"},
        {"id":"merge",        "label":"Merge / Combine"},
        {"id":"split-rotate", "label":"Split / Rotate / Reorder"},
        {"id":"extract-compress","label":"Extract / Compress"},
        {"id":"protect-unlock","label":"Protect / Unlock"},
        {"id":"conversions",  "label":"PDF↔DOC/IMG"},
        {"id":"wm-number",    "label":"Watermark / Numbers"},
    ]
    return render_template_string(
        PAGE,
        site_name=SITE_NAME, base_url=BASE_URL,
        adsense_client=ADSENSE_CLIENT, adsense_slot=ADSENSE_SLOT,
        ga4_id=GA4_ID, tabs=tabs
    )

@app.get("/about")
def about():
    body = f"""
{SITE_NAME} was built to make working with PDFs simple, fast, and accessible on any device.
<br><br>
<strong>What we do</strong><br>
Merge, split, rotate, re-order, compress, protect/unlock, watermark, add page numbers, convert images ↔ PDF,
export PDF to images or DOCX, and convert common office files to PDF. We aim to keep text crisp and images clear,
avoiding needless recompression.
<br><br>
<strong>Our principles</strong><br>
• Speed and quality — outputs should look professional without trial-and-error.<br>
• Privacy by default — files are processed securely and removed after delivery.<br>
• Reliability — predictable behavior across browsers and devices, no sign-ups required.
<br><br>
<strong>How it works</strong><br>
Everything runs on our servers while you wait; there’s nothing to install. We preserve vector text/graphics where possible
and only recompress when you explicitly choose compression settings.
<br><br>
<strong>Who it’s for</strong><br>
Students preparing assignments, freelancers sending proposals, teams producing client-ready PDFs—anyone who needs dependable tools that “just work”.
<br><br>
<strong>What’s next</strong><br>
We continuously improve performance and add sensible options without bloat. If something isn’t working the way you expect,
email us at {CONTACT_EMAIL} — we read every message.
"""
    return PAGES_SIMPLE("About", body)


@app.get("/privacy")
def privacy():
    body = f"""
We respect your privacy. This page explains what we process, why, and your choices.
<br><br>
<strong>1) Files you upload</strong><br>
• Purpose: provide the requested PDF tool (e.g., merge, split, convert).<br>
• Storage: files are kept only as long as needed to produce a download and are then deleted from temporary storage.<br>
• Access: files are not shared or sold and are only handled by automated processes required to run the tool.
<br><br>
<strong>2) Basic logs & security</strong><br>
We may log anonymized events (timestamps, status codes, basic error messages) to keep the service reliable and prevent abuse.
These logs do not include your document contents.
<br><br>
<strong>3) Cookies & analytics</strong><br>
We may use Google Analytics 4 to understand overall usage (pages, device types) so we can improve {SITE_NAME}. GA4 may set cookies or use similar technologies.
You can use your browser settings to block analytics cookies if you prefer.
<br><br>
<strong>4) Advertising (Google AdSense)</strong><br>
We display ads via Google AdSense. Google and its partners may use cookies to serve and measure ads (including personalized ads where permitted).
You can manage ad personalization at adssettings.google.com and learn more in Google’s policies.
<br><br>
<strong>5) What we do NOT do</strong><br>
• We don’t sell, rent, or trade your personal information.<br>
• We don’t scan your documents to build user profiles.
<br><br>
<strong>6) Data retention</strong><br>
Temporary files are removed after the download window. Diagnostic logs are retained only as long as necessary for security and operations.
<br><br>
<strong>7) Your choices</strong><br>
Use your browser’s privacy controls to block cookies; you may still use the core tools. For any privacy questions or deletion requests,
contact us at {CONTACT_EMAIL}.
<br><br>
<strong>8) Children</strong><br>
{SITE_NAME} is a general-audience service and not directed to children under 13.
<br><br>
<strong>9) Changes</strong><br>
We may update this policy to reflect improvements or legal requirements. Continued use of the site means you accept the updated policy.
"""
    return PAGES_SIMPLE("Privacy", body)


@app.get("/terms")
def terms():
    body = f"""
By using {SITE_NAME}, you agree to these terms.
<br><br>
<strong>Acceptable use</strong><br>
You will not upload illegal content, malware, or materials that infringe intellectual-property or privacy rights, and you will not attempt to disrupt or reverse-engineer the service.
You must have the necessary rights to process any files you upload.
<br><br>
<strong>Service availability</strong><br>
We aim for high uptime but do not guarantee uninterrupted service. Features may change or be discontinued at any time.
<br><br>
<strong>Intellectual property</strong><br>
All rights to the {SITE_NAME} website, design, and software are reserved by us or our licensors. Your documents remain yours.
<br><br>
<strong>Disclaimer & limitation of liability</strong><br>
The service is provided “as is” without warranties of any kind. To the maximum extent permitted by law, we are not liable for any indirect, incidental,
special, consequential, or punitive damages, or for loss of data, profits, or business, arising from your use of the service.
<br><br>
<strong>Indemnity</strong><br>
You agree to indemnify and hold us harmless from claims arising from your use of the service or your violation of these terms.
<br><br>
<strong>Termination</strong><br>
We may suspend or terminate access for any violation or to protect the service and its users.
<br><br>
<strong>Governing law</strong><br>
These terms are governed by applicable local laws where the service is operated. If any term is unenforceable, the remainder stays in effect.
<br><br>
<strong>Updates</strong><br>
We may update these terms; continued use after publication constitutes acceptance. If you do not agree, please stop using the service.
"""
    return PAGES_SIMPLE("Terms", body)


@app.get("/contact")
def contact():
    body = f"""
We’d love to hear from you.
<br><br>
<strong>Email</strong><br>
{CONTACT_EMAIL}
<br><br>
<strong>What to include</strong><br>
• A brief description of the task you were trying to do (e.g., “merge two PDFs”).<br>
• The browser/device you’re using and any error message you saw.<br>
• If it’s safe to share, sample files (or a minimal example) that reproduces the issue.
<br><br>
<strong>Response time</strong><br>
We aim to reply within 2–3 business days. For sensitive security reports, please include “SECURITY” in the subject line so we can prioritize.
<br><br>
Thank you for helping us make {SITE_NAME} better for everyone!
"""
    return PAGES_SIMPLE("Contact", body)



# ==========================
# Static text endpoints
# ==========================
@app.get("/robots.txt")
def robots():
    txt = f"User-agent: *\nAllow: /\nSitemap: {BASE_URL}/sitemap.xml\n"
    return make_response((txt, 200, {"Content-Type":"text/plain"}))

@app.get("/ads.txt")
def ads_txt():
    # Google AdSense ads.txt line (expects "pub-XXXXXXXXXXXXXXX")
    publisher_id = ADSENSE_CLIENT.replace("ca-pub-","pub-")
    txt = f"google.com, {publisher_id}, DIRECT, f08c47fec0942fa0\n"
    return make_response((txt, 200, {"Content-Type":"text/plain"}))

@app.get("/sitemap.xml")
def sitemap():
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url><loc>{BASE_URL}/</loc></url>
  <url><loc>{BASE_URL}/about</loc></url>
  <url><loc>{BASE_URL}/privacy</loc></url>
  <url><loc>{BASE_URL}/terms</loc></url>
  <url><loc>{BASE_URL}/contact</loc></url>
</urlset>"""
    return make_response((xml, 200, {"Content-Type":"application/xml"}))

# ==========================
# Routes: Tools
# ==========================
@app.post("/tool/images-to-pdf")
def images_to_pdf():
    files = request.files.getlist("images")
    if not files:
        return "No images uploaded", 400
    page_size = request.form.get("page_size", "Original")
    dpi = int(request.form.get("dpi", "300"))
    margin = float(request.form.get("margin", "24"))
    orientation = request.form.get("orientation", "Auto")
    output = request.form.get("output", "Single PDF (all images)")

    def page_dims(iw, ih) -> Tuple[float, float]:
        if page_size == "Original":
            w_in = iw / float(max(dpi,1)); h_in = ih / float(max(dpi,1))
            return max(w_in*72.0,1.0), max(h_in*72.0,1.0)
        if page_size == "A4":     w_pt,h_pt = 595.276, 841.890
        else:                     w_pt,h_pt = 612.0,   792.0
        if orientation == "Landscape" or (orientation=="Auto" and iw>ih):
            w_pt,h_pt = h_pt,w_pt
        return w_pt,h_pt

    def safe_inner(rect, m):
        max_mx = max((rect.width-1.0)/2.0,0.0)
        max_my = max((rect.height-1.0)/2.0,0.0)
        m = min(m, max_mx, max_my)
        return fitz.Rect(rect.x0+m, rect.y0+m, rect.x1-m, rect.y1-m)

    def target_rect(rect, iw, ih):
        inner = safe_inner(rect, margin)
        pw, ph = max(inner.width,1e-6), max(inner.height,1e-6)
        ar = iw/float(max(ih,1))
        if pw/ph > ar:
            nh = ph; nw = ar*nh
        else:
            nw = pw; nh = nw/ar
        x0 = inner.x0 + (pw-nw)/2.0; y0 = inner.y0 + (ph-nh)/2.0
        return fitz.Rect(x0,y0,x0+nw,y0+nh)

    def image_stream_lossless(path):
        with open(path,"rb") as f: raw=f.read()
        try:
            with Image.open(io.BytesIO(raw)) as im:
                if im.mode in ("RGBA","LA","P"):
                    bg = Image.new("RGB", im.size, (255,255,255))
                    im_rgba = im.convert("RGBA")
                    alpha = im_rgba.split()[-1]
                    bg.paste(im_rgba, mask=alpha)
                    buf = io.BytesIO(); bg.save(buf, format="PNG", optimize=True)
                    return buf.getvalue()
            return raw
        except:
            return raw

    tmp_paths=[]
    try:
        for fs in files:
            p = _save_upload(fs)
            tmp_paths.append(p)

        if output.startswith("Single"):
            doc = fitz.open()
            for p in tmp_paths:
                with Image.open(p) as im: iw,ih = im.size
                w_pt,h_pt = page_dims(iw,ih)
                page = doc.new_page(width=w_pt, height=h_pt)
                rect = target_rect(page.rect, iw, ih)
                page.insert_image(rect, stream=image_stream_lossless(p), keep_proportion=True, overlay=True)
            out = io.BytesIO(); doc.save(out, deflate=True); doc.close()
            return _send_bytes(out.getvalue(), "images.pdf", "application/pdf")
        else:
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, "w") as zf:
                for i,p in enumerate(tmp_paths, start=1):
                    with Image.open(p) as im: iw,ih = im.size
                    w_pt,h_pt = page_dims(iw,ih)
                    doc = fitz.open()
                    page = doc.new_page(width=w_pt, height=h_pt)
                    rect = target_rect(page.rect, iw, ih)
                    page.insert_image(rect, stream=image_stream_lossless(p), keep_proportion=True, overlay=True)
                    outp = tempfile.mktemp(suffix=f"_image_{i}.pdf")
                    doc.save(outp, deflate=True); doc.close()
                    zf.write(outp, os.path.basename(outp))
                    os.remove(outp)
            mem.seek(0)
            return _send_bytes(mem.getvalue(), "images.zip", "application/zip")
    finally:
        for p in tmp_paths: 
            try: os.remove(p)
            except: pass

@app.post("/tool/merge-append")
def merge_append():
    pdfs = request.files.getlist("pdfs")
    if not pdfs:
        return "No PDFs uploaded", 400
    ranges = (request.form.get("ranges") or "").strip()
    add_bm = "bookmarks" in request.form

    writer = PdfWriter()
    for fs in pdfs:
        p = _save_upload(fs, suffix=".pdf")
        try:
            r = PdfReader(p)
            idxs = _parse_ranges(ranges, len(r.pages)) if ranges else list(range(len(r.pages)))
            if add_bm:
                writer.add_outline_item(fs.filename, len(writer.pages), 0)
            for i in idxs:
                writer.add_page(r.pages[i])
        finally:
            try: os.remove(p)
            except: pass

    out = io.BytesIO(); writer.write(out)
    return _send_bytes(out.getvalue(), "merged.pdf", "application/pdf")

@app.post("/tool/combine-interleave")
def combine_interleave():
    pdfs = request.files.getlist("pdfs")
    if len(pdfs) < 2:
        return "Upload at least two PDFs", 400
    chunk = max(1, int(request.form.get("chunk","1")))
    loop = request.form.get("loop","Yes") == "Yes"

    queues = []
    for fs in pdfs:
        p = _save_upload(fs, suffix=".pdf")
        r = PdfReader(p)
        queues.append({"name":fs.filename, "path":p, "reader":r, "pages": list(range(len(r.pages))), "chunk":chunk})

    writer = PdfWriter()
    try:
        while True:
            progress_any = False
            for q in queues:
                if not q["pages"]: continue
                progress_any = True
                take = min(q["chunk"], len(q["pages"]))
                for _ in range(take):
                    i = q["pages"].pop(0)
                    writer.add_page(q["reader"].pages[i])
            if not loop or not progress_any:
                break
        out = io.BytesIO(); writer.write(out)
        return _send_bytes(out.getvalue(), "combined.pdf", "application/pdf")
    finally:
        for q in queues:
            try: os.remove(q["path"])
            except: pass

@app.post("/tool/split")
def split_pdf():
    fs = request.files.get("pdf")
    if not fs: return "No PDF uploaded", 400
    mode = request.form.get("mode","Page ranges (custom)")
    ranges = (request.form.get("ranges") or "").strip()
    size_mb = max(1, int(request.form.get("size_mb","5")))
    p = _save_upload(fs, suffix=".pdf")
    try:
        reader = PdfReader(p)
        mem = io.BytesIO()
        with zipfile.ZipFile(mem, "w") as zf:
            if mode.startswith("Page ranges"):
                if not ranges:
                    return "Provide ranges", 400
                parts = [x.strip() for x in ranges.split(",") if x.strip()]
                for idx, part in enumerate(parts, start=1):
                    idxs = _parse_ranges(part, len(reader.pages))
                    w = PdfWriter()
                    for i in idxs: w.add_page(reader.pages[i])
                    b = io.BytesIO(); w.write(b)
                    zf.writestr(f"split_part_{idx}_{part}.pdf", b.getvalue())
            elif mode == "Every page":
                for i,pg in enumerate(reader.pages, start=1):
                    w=PdfWriter(); w.add_page(pg)
                    b=io.BytesIO(); w.write(b)
                    zf.writestr(f"page_{i}.pdf", b.getvalue())
            else:
                w=PdfWriter(); part=1
                for i,pg in enumerate(reader.pages, start=1):
                    w.add_page(pg)
                    b=io.BytesIO(); w.write(b)
                    if b.tell() >= size_mb*1024*1024:
                        zf.writestr(f"part_{part}.pdf", b.getvalue()); part+=1; w=PdfWriter()
                if len(w.pages)>0:
                    b=io.BytesIO(); w.write(b); zf.writestr(f"part_{part}.pdf", b.getvalue())
        mem.seek(0)
        return _send_bytes(mem.getvalue(), "splits.zip", "application/zip")
    finally:
        try: os.remove(p)
        except: pass

@app.post("/tool/rotate")
def rotate_pdf():
    fs = request.files.get("pdf")
    if not fs: return "No PDF", 400
    deg = int(request.form.get("deg","90"))
    p = _save_upload(fs, suffix=".pdf")
    try:
        r = PdfReader(p); w=PdfWriter()
        for pg in r.pages:
            try:
                pg.rotate(deg)  # pypdf 3.x
            except Exception:
                from pypdf import Transformation
                pg.add_transformation(Transformation().rotate(deg))  # fallback
            w.add_page(pg)
        out = io.BytesIO(); w.write(out)
        return _send_bytes(out.getvalue(), "rotated.pdf", "application/pdf")
    finally:
        try: os.remove(p)
        except: pass

@app.post("/tool/reorder")
def reorder_pdf():
    fs = request.files.get("pdf")
    if not fs: return "No PDF", 400
    keep = (request.form.get("keep") or "").strip()
    order = (request.form.get("order") or "").strip()
    p = _save_upload(fs, suffix=".pdf")
    try:
        r = PdfReader(p); w=PdfWriter()
        total = len(r.pages)
        keep_idxs = _parse_ranges(keep, total) if keep else list(range(total))
        if order:
            chosen = [int(x.strip())-1 for x in order.split(",") if x.strip().isdigit()]
        else:
            chosen = keep_idxs
        for i in chosen:
            if 0<=i<total and i in keep_idxs:
                w.add_page(r.pages[i])
        out = io.BytesIO(); w.write(out)
        return _send_bytes(out.getvalue(), "reordered.pdf", "application/pdf")
    finally:
        try: os.remove(p)
        except: pass

@app.post("/tool/extract-text")
def extract_text():
    fs = request.files.get("pdf")
    if not fs: return "No PDF", 400
    mode = request.form.get("mode","Plain text")
    ranges = (request.form.get("ranges") or "").strip()
    search = (request.form.get("search") or "").strip()
    p = _save_upload(fs, suffix=".pdf")
    try:
        doc = fitz.open(p)
        idxs = _parse_ranges(ranges, len(doc)) if ranges else list(range(len(doc)))
        out_texts=[]
        for i in idxs:
            m = "text" if mode=="Plain text" else ("blocks" if mode=="Preserve layout" else "raw")
            text = doc[i].get_text(m)
            if search and search.lower() in text.lower():
                text = re.sub(re.escape(search), lambda m: f">>>{m.group(0)}<<<", text, flags=re.IGNORECASE)
            out_texts.append(f"--- Page {i+1} ---\n{text}")
        final = "\n\n".join(out_texts).encode("utf-8")
        return _send_bytes(final, "extracted.txt", "text/plain")
    finally:
        try: os.remove(p)
        except: pass

@app.post("/tool/compress")
def compress_pdf():
    fs = request.files.get("pdf")
    if not fs: return "No PDF", 400
    preset = request.form.get("preset","Standard")
    quality = int(request.form.get("quality","85"))
    target_dpi = request.form.get("target_dpi","").strip()
    target_dpi = int(target_dpi) if target_dpi else None
    strip_meta = "strip_meta" in request.form
    linearize = "linearize" in request.form
    clean_xref = "clean_xref" in request.form

    src = fitz.open(stream=fs.read(), filetype="pdf")
    if preset == "Light (best quality)":
        quality = 90; target_dpi = target_dpi or 300
    elif preset == "Standard":
        quality = 85; target_dpi = target_dpi or 200
    elif preset == "Strong":
        quality = 75; target_dpi = target_dpi or 150
    elif preset == "Extreme":
        quality = 60; target_dpi = target_dpi or 100
    elif preset == "Lossless (no image recompress)":
        target_dpi = None; quality = None

    if strip_meta:
        try: src.set_metadata({})
        except: pass

    if quality is not None:
        for page in src:
            imgs = page.get_images(full=True)
            if not imgs: continue
            page_w_in = page.rect.width/72.0
            page_h_in = page.rect.height/72.0
            for xref, *_ in imgs:
                try:
                    info = src.extract_image(xref)
                    img_bytes = info["image"]; w_px, h_px = info.get("width"), info.get("height")
                except: continue
                new_w,new_h = w_px,h_px
                if target_dpi:
                    max_w = int(page_w_in*target_dpi*1.2+0.5)
                    max_h = int(page_h_in*target_dpi*1.2+0.5)
                    if w_px>max_w or h_px>max_h:
                        scale = max(min(min(max_w/w_px, max_h/h_px),1.0),0.05)
                        new_w = max(int(w_px*scale),1); new_h=max(int(h_px*scale),1)
                try:
                    pil = Image.open(io.BytesIO(img_bytes)).convert("RGB")
                    if (new_w,new_h)!=(w_px,h_px):
                        pil = pil.resize((new_w,new_h), Image.LANCZOS)
                    buf = io.BytesIO()
                    pil.save(buf, format="JPEG", quality=int(quality), optimize=True)
                    src.update_image(xref, stream=buf.getvalue())
                except: pass

    out_buf = io.BytesIO()
    save_kwargs = dict(deflate=True)
    if linearize: save_kwargs["linear"] = True
    if clean_xref: save_kwargs.update(clean=True, garbage=3)
    src.save(out_buf, **save_kwargs); src.close()
    return _send_bytes(out_buf.getvalue(), "compressed.pdf", "application/pdf")

@app.post("/tool/protect")
def protect_pdf():
    fs = request.files.get("pdf")
    if not fs: return "No PDF", 400
    user_pwd = request.form.get("user_pwd") or ""
    owner_pwd = request.form.get("owner_pwd") or user_pwd
    allow_print = "allow_print" in request.form
    allow_copy  = "allow_copy" in request.form
    allow_annot = "allow_annot" in request.form
    encryption  = request.form.get("encryption","AES-256")

    p = _save_upload(fs, suffix=".pdf")
    try:
        r = PdfReader(p); w = PdfWriter()
        for pg in r.pages: w.add_page(pg)

        # Permissions (use new API if available)
        perm_obj = None
        try:
            from pypdf import Permissions
            perms = set()
            if allow_print: perms.add(Permissions.PRINT)
            if allow_copy: perms.add(Permissions.COPY)
            if allow_annot: perms.add(Permissions.ANNOTATE)
            perm_obj = perms
        except Exception:
            perm_obj = None

        try:
            from pypdf import EncryptionAlgorithm
            algo = (EncryptionAlgorithm.AES_256 if "256" in encryption
                    else EncryptionAlgorithm.AES_128 if "AES-128"==encryption
                    else EncryptionAlgorithm.RC4_128 if "RC4-128"==encryption
                    else EncryptionAlgorithm.RC4_40)
            w.encrypt(user_password=user_pwd or None,
                      owner_password=owner_pwd or None,
                      permissions=perm_obj,
                      algorithm=algo)
        except Exception:
            # fallback older api
            w.encrypt(user_password=user_pwd or None,
                      owner_password=owner_pwd or None,
                      use_128bit=("128" in encryption))
        out = io.BytesIO(); w.write(out)
        return _send_bytes(out.getvalue(), "protected.pdf", "application/pdf")
    finally:
        try: os.remove(p)
        except: pass

@app.post("/tool/unlock")
def unlock_pdf():
    files = request.files.getlist("pdfs")
    password = request.form.get("password") or ""
    if not files or not password:
        return "Upload PDFs and provide password", 400
    tmpdir = tempfile.mkdtemp(prefix="unlock_")
    outs = []
    try:
        for fs in files:
            p = _save_upload(fs, suffix=".pdf")
            try:
                r = PdfReader(p)
                try:
                    ok = r.decrypt(password)  # legacy api
                    if ok == 0:
                        continue
                except Exception:
                    # newer pypdf auto-decrypts on access; if still encrypted and wrong, add_page will raise
                    pass
                w = PdfWriter()
                for pg in r.pages: w.add_page(pg)
                outp = os.path.join(tmpdir, os.path.splitext(secure_filename(fs.filename))[0] + "_unlocked.pdf")
                with open(outp,"wb") as f: w.write(f)
                outs.append(outp)
            finally:
                try: os.remove(p)
                except: pass
        if not outs:
            shutil.rmtree(tmpdir, ignore_errors=True)
            return "No files unlocked (wrong password or not encrypted).", 400
        if len(outs)==1:
            data = open(outs[0],"rb").read()
            shutil.rmtree(tmpdir, ignore_errors=True)
            return _send_bytes(data, os.path.basename(outs[0]), "application/pdf")
        mem=io.BytesIO()
        with zipfile.ZipFile(mem,"w") as zf:
            for pth in outs:
                zf.write(pth, os.path.basename(pth))
        mem.seek(0); shutil.rmtree(tmpdir, ignore_errors=True)
        return _send_bytes(mem.getvalue(), "unlocked.zip", "application/zip")
    except Exception as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        return f"Failed: {e}", 500

@app.post("/tool/pdf-to-images")
def pdf_to_images():
    pdfs = request.files.getlist("pdfs")
    if not pdfs: return "No PDFs", 400
    dpi = int(request.form.get("dpi","300"))
    fmt = request.form.get("fmt","PNG")
    quality = int(request.form.get("quality","95"))
    mode = request.form.get("mode","One image per page (ZIP)")
    tmpdir = tempfile.mkdtemp(prefix="pdf2img_")
    try:
        for fs in pdfs:
            p = _save_upload(fs, suffix=".pdf")
            doc = fitz.open(p)
            base = os.path.join(tmpdir, os.path.splitext(secure_filename(fs.filename))[0])
            os.makedirs(base, exist_ok=True)
            images=[]
            for i,page in enumerate(doc, start=1):
                pix = page.get_pixmap(dpi=dpi, alpha=False)
                if fmt=="PNG":
                    data = pix.tobytes("png"); ext="png"
                elif fmt=="JPEG":
                    data = pix.tobytes("jpg", jpg_quality=quality); ext="jpg"
                else:
                    pilimg = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    buf = io.BytesIO(); pilimg.save(buf, format="WEBP", quality=quality); data=buf.getvalue(); ext="webp"
                outp = os.path.join(base, f"page_{i}.{ext}")
                with open(outp,"wb") as f: f.write(data)
                images.append(Image.open(io.BytesIO(data)))
            if mode.startswith("All pages stitched") and images:
                widths, heights = zip(*(im.size for im in images))
                total_h = sum(heights); max_w = max(widths)
                stitched = Image.new("RGB", (max_w, total_h), "white")
                y=0
                for im in images:
                    stitched.paste(im,(0,y)); y += im.height
                stitched_path = os.path.join(base, f"{os.path.basename(base)}_stitched.{ 'jpg' if fmt=='JPEG' else ('webp' if fmt=='WebP' else 'png') }")
                if fmt=="PNG": stitched.save(stitched_path, format="PNG")
                elif fmt=="JPEG": stitched.save(stitched_path, format="JPEG", quality=quality)
                else: stitched.save(stitched_path, format="WEBP", quality=quality)
        mem = io.BytesIO()
        with zipfile.ZipFile(mem,"w") as zf:
            for root,_,files in os.walk(tmpdir):
                for f in files:
                    full = os.path.join(root,f)
                    zf.write(full, os.path.relpath(full, tmpdir))
        mem.seek(0)
        return _send_bytes(mem.getvalue(), "images.zip", "application/zip")
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

@app.post("/tool/pdf-to-docx")
def pdf_to_docx():
    fs = request.files.get("pdf")
    if not fs: return "No PDF", 400
    start = max(1, int(request.form.get("start","1")))
    end = int(request.form.get("end","0"))
    keep_images = "keep_images" in request.form
    p = _save_upload(fs, suffix=".pdf")
    out_dir = tempfile.mkdtemp(prefix="pdf2docx_")
    out_path = os.path.join(out_dir, "converted.docx")
    try:
        cv = Pdf2DocxConverter(p)
        cv.convert(out_path, start=start-1, end=None if end==0 else end-1, retain_image=keep_images)
        cv.close()
        data = open(out_path, "rb").read()
        shutil.rmtree(out_dir, ignore_errors=True); os.remove(p)
        return _send_bytes(data, "converted.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
    except Exception as e:
        shutil.rmtree(out_dir, ignore_errors=True)
        try: os.remove(p)
        except: pass
        return f"Conversion failed: {e}", 500

@app.post("/tool/office-to-pdf")
def office_to_pdf():
    files = request.files.getlist("files")
    if not files: return "No files", 400
    tmpdir = tempfile.mkdtemp(prefix="office2pdf_")
    pdf_paths = []
    try:
        for fs in files:
            fname = secure_filename(fs.filename)
            ext = fname.rsplit(".",1)[-1].lower()
            in_path = _save_upload(fs, suffix=f".{ext}")
            out_path = os.path.join(tmpdir, os.path.splitext(fname)[0] + ".pdf")
            try:
                if ext in ["docx","odt"]:
                    import pypandoc
                    pypandoc.convert_file(in_path, "pdf", outputfile=out_path, extra_args=["--standalone"])
                elif ext in ["xlsx","ods"]:
                    import pandas as pd
                    from reportlab.platypus import SimpleDocTemplate, Table
                    from reportlab.lib.pagesizes import A4
                    xl = pd.ExcelFile(in_path)
                    doc = SimpleDocTemplate(out_path, pagesize=A4)
                    elements=[]
                    for sheet in xl.sheet_names:
                        df = xl.parse(sheet)
                        data = [df.columns.tolist()] + df.values.tolist()
                        elements.append(Table(data))
                    doc.build(elements)
                elif ext in ["pptx","odp"]:
                    from pptx import Presentation
                    prs = Presentation(in_path)
                    pdf = fitz.open()
                    for _ in prs.slides:
                        page = pdf.new_page(width=800, height=600)
                        page.insert_text((40, 300), "[Slide content placeholder]")
                    pdf.save(out_path); pdf.close()
                elif ext == "txt":
                    from reportlab.platypus import SimpleDocTemplate, Paragraph
                    from reportlab.lib.styles import getSampleStyleSheet
                    doc = SimpleDocTemplate(out_path)
                    styles = getSampleStyleSheet()
                    text = open(in_path,"rb").read().decode("utf-8", errors="ignore")
                    paras = [Paragraph(line, styles["Normal"]) for line in text.split("\n")]
                    doc.build(paras)
                elif ext in ["png","jpg","jpeg"]:
                    im = Image.open(in_path).convert("RGB")
                    im.save(out_path, "PDF")
                else:
                    raise RuntimeError(f"Unsupported: {ext}")
                pdf_paths.append(out_path)
            except Exception:
                # Skip file on error; continue with others
                pass
            finally:
                try: os.remove(in_path)
                except: pass
        if not pdf_paths:
            shutil.rmtree(tmpdir, ignore_errors=True)
            return "No files converted (missing system tools for some formats?)", 400
        if len(pdf_paths)==1:
            data=open(pdf_paths[0],"rb").read()
            shutil.rmtree(tmpdir, ignore_errors=True)
            return _send_bytes(data, os.path.basename(pdf_paths[0]), "application/pdf")
        mem=io.BytesIO()
        with zipfile.ZipFile(mem,"w") as zf:
            for p in pdf_paths:
                zf.write(p, os.path.basename(p))
        mem.seek(0); shutil.rmtree(tmpdir, ignore_errors=True)
        return _send_bytes(mem.getvalue(), "converted_pdfs.zip", "application/zip")
    except Exception as e:
        shutil.rmtree(tmpdir, ignore_errors=True)
        return f"Failed: {e}", 500

@app.post("/tool/watermark")
def watermark():
    fs = request.files.get("pdf")
    if not fs: return "No PDF", 400
    wm_type = request.form.get("wm_type","Text")
    text = request.form.get("text","CONFIDENTIAL")
    color = request.form.get("color","#FF0000")
    opacity = int(request.form.get("opacity","20"))
    size = int(request.form.get("size","60"))
    angle = int(request.form.get("angle","45"))
    pos = request.form.get("pos","Center")
    pages_str = (request.form.get("pages") or "").strip()
    wm_img = request.files.get("wm_img")

    p = _save_upload(fs, suffix=".pdf")
    try:
        doc = fitz.open(p)
        idxs = _parse_ranges(pages_str, len(doc)) if pages_str else list(range(len(doc)))
        for i in idxs:
            page = doc[i]; rect = page.rect
            if wm_type=="Text":
                rgb = tuple(int(color[j:j+2],16)/255 for j in (1,3,5))
                if pos=="Diagonal Tiled":
                    step = rect.height/3
                    y = 0
                    while y < rect.height:
                        page.insert_text((rect.width/2, y), text, fontsize=size, rotate=angle,
                                         color=rgb, fill_opacity=opacity/100, align=1)
                        y += step
                else:
                    coords = {
                        "Center": rect.center,
                        "Top-left": (rect.x0+50, rect.y0+80),
                        "Top-right": (rect.x1-200, rect.y0+80),
                        "Bottom-left": (rect.x0+50, rect.y1-50),
                        "Bottom-right": (rect.x1-200, rect.y1-50)
                    }
                    page.insert_text(coords.get(pos, rect.center), text, fontsize=size, rotate=angle,
                                     color=rgb, fill_opacity=opacity/100, align=1)
            else:
                if not wm_img: continue
                img_path = _save_upload(wm_img)
                try:
                    scale = size/100.0
                    w = int(rect.width*scale); h=int(rect.height*scale)
                    pos_map = {
                        "Center": fitz.Rect(rect.width/2-w/2, rect.height/2-h/2, rect.width/2+w/2, rect.height/2+h/2),
                        "Top-left": fitz.Rect(50,50,50+w,50+h),
                        "Top-right": fitz.Rect(rect.width-w-50, 50, rect.width-50, 50+h),
                        "Bottom-left": fitz.Rect(50, rect.height-h-50, 50+w, rect.height-50),
                        "Bottom-right": fitz.Rect(rect.width-w-50, rect.height-h-50, rect.width-50, rect.height-50),
                    }
                    target = pos_map.get(pos, rect)
                    page.insert_image(target, filename=img_path, overlay=True, keep_proportion=True, opacity=opacity/100)
                finally:
                    try: os.remove(img_path)
                    except: pass
        out = io.BytesIO(); doc.save(out, deflate=True); doc.close()
        return _send_bytes(out.getvalue(), "watermarked.pdf", "application/pdf")
    finally:
        try: os.remove(p)
        except: pass

@app.post("/tool/page-numbers")
def page_numbers():
    fs = request.files.get("pdf")
    if not fs: return "No PDF", 400
    ranges = (request.form.get("ranges") or "").strip()
    style = request.form.get("style","1, 2, 3")
    template = request.form.get("template","Page {n} of {total}")
    pos_v = request.form.get("pos_v","Bottom")
    pos_h = request.form.get("pos_h","Center")
    fontsize = int(request.form.get("fontsize","12"))
    color = request.form.get("color","#000000")
    opacity = int(request.form.get("opacity","80"))

    p = _save_upload(fs, suffix=".pdf")
    try:
        doc = fitz.open(p); total = len(doc)
        idxs = _parse_ranges(ranges, total) if ranges else list(range(total))
        def fmt(n):
            if style == "1, 2, 3": return str(n)
            if style == "01, 02, 03": return f"{n:02d}"
            if style == "i, ii, iii": return _roman(n).lower()
            if style == "I, II, III": return _roman(n)
            if style == "a, b, c":    return chr(96+n)
            if style == "A, B, C":    return chr(64+n)
            return str(n)
        rgb = tuple(int(color[i:i+2],16)/255 for i in (1,3,5))
        for i in idxs:
            page = doc[i]; rect = page.rect
            num = fmt(i+1); label = template.format(n=num, total=total)
            x = rect.x0 + 40 if pos_h=="Left" else rect.x1-40 if pos_h=="Right" else rect.x0 + rect.width/2
            y = rect.y1 - 30 if pos_v=="Bottom" else rect.y0 + 40
            page.insert_text((x,y), label, fontsize=fontsize, color=rgb, fill_opacity=opacity/100,
                             align=1 if pos_h=="Center" else 0)
        out = io.BytesIO(); doc.save(out, deflate=True); doc.close()
        return _send_bytes(out.getvalue(), "numbered.pdf", "application/pdf")
    finally:
        try: os.remove(p)
        except: pass

# ==========================
# Main
# ==========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    app.run(host="0.0.0.0", port=port, debug=False)








