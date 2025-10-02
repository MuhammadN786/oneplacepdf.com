# OnePlacePDF — best-quality build
# High-quality PDF toolkit with careful handling to avoid needless recompression.
# Tabs: Images→PDF, Merge, Split, Rotate, Re-order, Extract Text, Edit, Compress,
#       Protect, Unlock, PDF→Images, PDF→DOCX, Watermark, Page Numbers, Office→PDF
#
# pip install -r requirements.txt
# streamlit run app.py

import io, os, re, shutil, subprocess, tempfile, zipfile, uuid
from typing import List

import streamlit as st
from pypdf import PdfReader, PdfWriter
import fitz  # PyMuPDF
from PIL import Image
from pdf2docx import Converter as Pdf2DocxConverter

# ============================
# Utilities (quality-focused)
# ============================

def _save_upload(uf, suffix=""):
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(uf.read())
    return path

def _save_uploads(ufs, suffix=""):
    return [_save_upload(u, suffix=suffix) for u in ufs]

def _to_bytes(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()

def _download(label: str, data: bytes, filename: str, mime="application/octet-stream"):
    st.download_button(label, data=data, file_name=filename, mime=mime)

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

def _which_gs():
    for cand in ("gs", "gswin64c", "gswin32c"):
        if shutil.which(cand):
            return cand
    return None

def _which_soffice():
    return shutil.which("soffice")

# ==================================
# Streamlit page (no global footer)
# ==================================

st.set_page_config(page_title="OnePlacePDF — High Quality PDF Tools", page_icon="📄", layout="wide")
st.title("OnePlacePDF")
st.caption("Merge, split, convert & secure — with quality-first processing.")

tabs = st.tabs([
    "Images → PDF", "Merge", "Split", "Rotate", "Re-order", "Extract Text",
    "Edit", "Compress", "Protect", "Unlock",
    "PDF → Images", "PDF → DOCX", "Watermark", "Page Numbers", "Office → PDF"
])

# -------------------------
# Images → PDF (lossless)
# -------------------------
# Approach: embed original image bytes into a PDF page sized to the image.
# PNG/JPEG are inserted directly; non-RGB modes are converted to PNG (lossless).
with tabs[0]:
    st.subheader("Images → PDF (no unnecessary recompression)")
    imgs = st.file_uploader("Upload images (JPG/PNG, ordered)", type=["jpg","jpeg","png"], accept_multiple_files=True)
    keep_orientation = st.checkbox("Respect EXIF orientation (rotate as needed)", value=True)
    if imgs and st.button("Create PDF"):
        tmpdir = tempfile.mkdtemp(prefix="imgpdf_")
        try:
            # Normalize modes but keep lossless whenever possible
            norm_paths = []
            for uf in imgs:
                p = _save_upload(uf)
                try:
                    im = Image.open(p)
                    if keep_orientation:
                        try:
                            im = ImageOps.exif_transpose(im)  # auto-rotate if EXIF
                        except Exception:
                            pass
                    if im.mode not in ("RGB", "L"):  # convert to RGB to avoid weirdness
                        im = im.convert("RGB")
                    # Save as PNG (lossless) to ensure stable insert
                    outp = os.path.join(tmpdir, os.path.basename(uf.name) + ".png")
                    im.save(outp, "PNG", optimize=True)
                    norm_paths.append(outp)
                finally:
                    try:
                        os.remove(p)
                    except Exception:
                        pass

            # Build PDF with pages matching image pixel dimensions (1 px = 1 pt)
            doc = fitz.open()
            for imgp in norm_paths:
                with open(imgp, "rb") as fh:
                    img_bytes = fh.read()
                img = Image.open(imgp)
                w, h = img.size
                page = doc.new_page(width=float(w), height=float(h))
                page.insert_image(page.rect, stream=img_bytes, keep_proportion=False)
            out = io.BytesIO()
            doc.save(out, deflate=True)  # compress structure, not image pixels
            doc.close()
            _download("Download images.pdf", out.getvalue(), "images.pdf", "application/pdf")
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

# -----------
# Merge PDFs
# -----------
with tabs[1]:
    st.subheader("Merge PDFs")
    pdfs = st.file_uploader("Upload PDFs (in order)", type=["pdf"], accept_multiple_files=True)
    if pdfs and st.button("Merge Now"):
        writer = PdfWriter()
        for p in _save_uploads(pdfs, suffix=".pdf"):
            r = PdfReader(p)
            for pg in r.pages:
                writer.add_page(pg)
        out = io.BytesIO()
        writer.write(out)
        _download("Download merged.pdf", out.getvalue(), "merged.pdf", "application/pdf")

# ----------
# Split PDF
# ----------
with tabs[2]:
    st.subheader("Split by page ranges")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="split")
    ranges = st.text_input("Ranges (e.g., 1-3,5,7-9). Leave blank for one file (all pages).")
    if pdf and st.button("Split Now"):
        path = _save_upload(pdf, suffix=".pdf")
        r = PdfReader(path)
        total = len(r.pages)
        out_zip = io.BytesIO()
        with zipfile.ZipFile(out_zip, "w") as zf:
            if not ranges.strip():
                w = PdfWriter()
                for pg in r.pages:
                    w.add_page(pg)
                b = io.BytesIO(); w.write(b)
                zf.writestr("split_all.pdf", b.getvalue())
            else:
                # Keep user's comma-separated intent as separate outputs
                parts = [p.strip() for p in ranges.split(",") if p.strip()]
                for idx, part in enumerate(parts, start=1):
                    ids = _parse_ranges(part, total)
                    if not ids: 
                        continue
                    w = PdfWriter()
                    for i in ids:
                        w.add_page(r.pages[i])
                    b = io.BytesIO(); w.write(b)
                    zf.writestr(f"split_part_{idx}_{part.replace(' ','')}.pdf", b.getvalue())
        out_zip.seek(0)
        _download("Download splits.zip", out_zip.getvalue(), "splits.zip", "application/zip")

# ----------
# Rotate
# ----------
with tabs[3]:
    st.subheader("Rotate pages")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="rotate")
    deg = st.selectbox("Rotate by", [90, 180, 270], index=0)
    pages = st.text_input("Pages to rotate (e.g., 2,5-7). Blank = all.")
    if pdf and st.button("Rotate Now"):
        path = _save_upload(pdf, suffix=".pdf")
        r = PdfReader(path)
        which = set(_parse_ranges(pages, len(r.pages))) if pages.strip() else set(range(len(r.pages)))
        w = PdfWriter()
        for i, pg in enumerate(r.pages):
            if i in which:
                pg.rotate(deg)
            w.add_page(pg)
        out = io.BytesIO(); w.write(out)
        _download("Download rotated.pdf", out.getvalue(), "rotated.pdf", "application/pdf")

# ------------
# Re-order
# ------------
with tabs[4]:
    st.subheader("Re-order pages")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="reorder")
    order = st.text_input("New order (e.g., 3,1,2 or 1-3,5,4). Omitted pages are removed; duplicates allowed.")
    if pdf and order and st.button("Re-order Now"):
        path = _save_upload(pdf, suffix=".pdf")
        r = PdfReader(path)
        idxs = _parse_ranges(order, len(r.pages))
        if not idxs:
            st.error("No valid pages selected.")
        else:
            w = PdfWriter()
            for i in idxs:
                w.add_page(r.pages[i])
            out = io.BytesIO(); w.write(out)
            _download("Download reordered.pdf", out.getvalue(), "reordered.pdf", "application/pdf")

# ----------------
# Extract Text
# ----------------
with tabs[5]:
    st.subheader("Extract text")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="extract")
    pages = st.text_input("Pages (blank = all)")
    if pdf and st.button("Extract Now"):
        path = _save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)
        idxs = _parse_ranges(pages, len(doc)) if pages.strip() else list(range(len(doc)))
        chunks = []
        for i in idxs:
            chunks.append(f"--- Page {i+1} ---\n{doc[i].get_text('text')}")
        text = "\n\n".join(chunks)
        st.text_area("Text", value=text, height=400)
        st.download_button("Download extracted.txt", data=text, file_name="extracted.txt")

# -----------
# Edit
# -----------
with tabs[6]:
    st.subheader("Edit: keep or delete pages")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="edit")
    keep = st.text_input("Keep ONLY pages (e.g., 1-3,7). Leave blank to use Delete.")
    delete = st.text_input("Delete pages (e.g., 4-6). Ignored if 'Keep ONLY' is set.")
    if pdf and st.button("Apply Edit"):
        path = _save_upload(pdf, suffix=".pdf")
        r = PdfReader(path); w = PdfWriter()
        if keep.strip():
            idxs = _parse_ranges(keep, len(r.pages))
        else:
            dels = set(_parse_ranges(delete, len(r.pages))) if delete.strip() else set()
            idxs = [i for i in range(len(r.pages)) if i not in dels]
        if not idxs:
            st.error("No pages to keep.")
        else:
            for i in idxs:
                w.add_page(r.pages[i])
            out = io.BytesIO(); w.write(out)
            _download("Download edited.pdf", out.getvalue(), "edited.pdf", "application/pdf")

# --------------
# Compress (GS)
# --------------
with tabs[7]:
    st.subheader("Compress (Ghostscript)")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="compress")
    preset = st.selectbox("Quality preset", ["/prepress (highest quality)", "/printer", "/ebook", "/screen (smallest)", "/default"], index=0)
    if pdf and st.button("Compress Now"):
        gs = _which_gs()
        if not gs:
            st.error("Ghostscript not found on this system.")
        else:
            in_path = _save_upload(pdf, suffix=".pdf")
            out_dir = tempfile.mkdtemp(prefix="gs_")
            out_path = os.path.join(out_dir, "compressed.pdf")
            try:
                setting = preset.split(" ")[0]
                cmd = [
                    gs, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.6",
                    f"-dPDFSETTINGS={setting}",
                    "-dNOPAUSE", "-dQUIET", "-dBATCH",
                    f"-sOutputFile={out_path}", in_path
                ]
                subprocess.run(cmd, check=True)
                _download("Download compressed.pdf", _to_bytes(out_path), "compressed.pdf", "application/pdf")
            except Exception:
                st.error("Compression failed.")
            finally:
                shutil.rmtree(out_dir, ignore_errors=True)

# -------------
# Protect
# -------------
with tabs[8]:
    st.subheader("Protect (encrypt with password)")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="protect")
    user_pwd = st.text_input("Password", type="password")
    owner_pwd = st.text_input("Owner password (optional)", type="password")
    if pdf and user_pwd and st.button("Protect Now"):
        path = _save_upload(pdf, suffix=".pdf")
        r = PdfReader(path); w = PdfWriter()
        for p in r.pages: w.add_page(p)
        # Use positional args for widest pypdf compatibility
        w.encrypt(user_pwd, owner_pwd or user_pwd)
        out = io.BytesIO(); w.write(out)
        _download("Download protected.pdf", out.getvalue(), "protected.pdf", "application/pdf")

# -----------
# Unlock
# -----------
with tabs[9]:
    st.subheader("Unlock (requires current password)")
    pdf = st.file_uploader("Upload protected PDF", type=["pdf"], key="unlock")
    pwd = st.text_input("Current password", type="password")
    if pdf and pwd and st.button("Unlock Now"):
        path = _save_upload(pdf, suffix=".pdf")
        r = PdfReader(path)
        try:
            if getattr(r, "is_encrypted", False):
                res = r.decrypt(pwd)
                # pypdf may return 0 on failure instead of raising
                if res == 0:
                    raise ValueError("Incorrect password.")
        except Exception:
            st.error("Incorrect password or unsupported encryption.")
        else:
            w = PdfWriter()
            for p in r.pages: w.add_page(p)
            out = io.BytesIO(); w.write(out)
            _download("Download unlocked.pdf", out.getvalue(), "unlocked.pdf", "application/pdf")

# -------------------
# PDF → Images (HQ)
# -------------------
with tabs[10]:
    st.subheader("PDF → Images (PNG/JPEG) in ZIP")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pdf2img")
    fmt = st.selectbox("Image format", ["PNG (lossless, larger)", "JPEG (smaller, lossy)"], index=0)
    dpi = st.slider("DPI", min_value=72, max_value=600, value=300, step=6)
    jpg_quality = st.slider("JPEG quality", min_value=60, max_value=100, value=95, step=1, disabled=("JPEG" not in fmt))
    if pdf and st.button("Convert"):
        path = _save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)
        mem = io.BytesIO()
        with zipfile.ZipFile(mem, "w") as z:
            for i, page in enumerate(doc, start=1):
                pix = page.get_pixmap(dpi=dpi)
                if fmt.startswith("PNG"):
                    z.writestr(f"page_{i:03d}.png", pix.tobytes("png"))
                else:
                    z.writestr(f"page_{i:03d}.jpg", pix.tobytes("jpg", jpg_quality=jpg_quality))
        mem.seek(0)
        _download("Download pages.zip", mem.getvalue(), "pages.zip", "application/zip")

# ----------------
# PDF → DOCX
# ----------------
with tabs[11]:
    st.subheader("PDF → DOCX (best-effort)")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pdf2docx")
    start = st.number_input("Start page (1-based)", min_value=1, value=1)
    end = st.number_input("End page (0 = all)", min_value=0, value=0)
    if pdf and st.button("Convert"):
        in_path = _save_upload(pdf, suffix=".pdf")
        out_dir = tempfile.mkdtemp(prefix="pdf2docx_")
        out_path = os.path.join(out_dir, "converted.docx")
        try:
            cv = Pdf2DocxConverter(in_path)
            if end == 0:
                cv.convert(out_path, start=start-1, end=None)
            else:
                cv.convert(out_path, start=start-1, end=end-1)
            cv.close()
            _download("Download converted.docx", _to_bytes(out_path), "converted.docx",
                      "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception:
            st.error("Conversion failed. (Scanned/complex PDFs are hard for open-source tools.)")
        finally:
            shutil.rmtree(out_dir, ignore_errors=True)

# -------------
# Watermark
# -------------
with tabs[12]:
    st.subheader("Add text watermark")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="wm")
    text = st.text_input("Watermark text", "CONFIDENTIAL")
    opacity = st.slider("Opacity", 10, 90, 20)
    scale = st.slider("Size (relative)", 20, 200, 100)
    angle = st.slider("Angle", 0, 360, 45)
    if pdf and text and st.button("Apply Watermark"):
        path = _save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)
        for page in doc:
            rect = page.rect
            # heuristic font size for large diagonal watermark
            fontsize = max(12, (rect.width + rect.height) * (scale / 600))
            page.insert_text(
                rect.center, text,
                fontsize=fontsize, rotate=angle,
                color=(0, 0, 0), fill_opacity=opacity/100.0,
                render_mode=0, overlay=True, align=1
            )
        out = io.BytesIO(); doc.save(out); doc.close()
        _download("Download watermarked.pdf", out.getvalue(), "watermarked.pdf", "application/pdf")

# ----------------
# Page Numbers
# ----------------
with tabs[13]:
    st.subheader("Add page numbers (footer)")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pnums")
    position = st.selectbox("Position", ["Left","Center","Right"], index=1)
    startnum = st.number_input("Start from number", min_value=1, value=1)
    fontsize = st.slider("Font size (pt)", 8, 24, 11)
    if pdf and st.button("Add Numbers"):
        path = _save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)
        for i, page in enumerate(doc, start=0):
            num = f"{startnum + i}"
            rect = page.rect
            y = rect.y1 - 18
            if position == "Left":
                x, align = rect.x0 + 36, 0
            elif position == "Center":
                x, align = rect.x0 + rect.width/2, 1
            else:
                x, align = rect.x1 - 36, 2
            page.insert_text((x, y), num, fontsize=fontsize, color=(0,0,0), align=align)
        out = io.BytesIO(); doc.save(out); doc.close()
        _download("Download numbered.pdf", out.getvalue(), "numbered.pdf", "application/pdf")

# ----------------
# Office → PDF
# ----------------
with tabs[14]:
    st.subheader("Office → PDF (via LibreOffice)")
    sof = _which_soffice()
    if not sof:
        st.warning("LibreOffice (soffice) not found on this system. This tab requires it.")
    subt = st.tabs(["Word → PDF", "Excel → PDF", "PowerPoint → PDF"])

    def convert_office(uploaded, label):
        if not sof:
            st.error("LibreOffice is not available.")
            return
        in_path = _save_upload(uploaded)
        out_dir = tempfile.mkdtemp(prefix="soffice_")
        try:
            # Use dedicated temporary user profile to avoid profile locks
            profile = f"file:///{out_dir.replace(os.sep, '/')}/lo_profile_{uuid.uuid4().hex}"
            cmd = [
                sof, "--headless",
                f"-env:UserInstallation={profile}",
                "--convert-to", "pdf",
                "--outdir", out_dir,
                in_path
            ]
            subprocess.run(cmd, check=True)
            base = os.path.splitext(os.path.basename(uploaded.name))[0] + ".pdf"
            out_path = os.path.join(out_dir, base)
            if os.path.exists(out_path):
                _download(f"Download {label}", _to_bytes(out_path), base, "application/pdf")
            else:
                st.error("Conversion failed.")
        except Exception:
            st.error("LibreOffice conversion failed.")
        finally:
            shutil.rmtree(out_dir, ignore_errors=True)

    # Word
    with subt[0]:
        doc = st.file_uploader("Upload Word (.doc/.docx)", type=["doc","docx"])
        if doc and st.button("Convert Word → PDF"):
            convert_office(doc, "Word PDF")

    # Excel
    with subt[1]:
        xls = st.file_uploader("Upload Excel (.xls/.xlsx)", type=["xls","xlsx"])
        if xls and st.button("Convert Excel → PDF"):
            convert_office(xls, "Excel PDF")

    # PowerPoint
    with subt[2]:
        ppt = st.file_uploader("Upload PowerPoint (.ppt/.pptx)", type=["ppt","pptx"])
        if ppt and st.button("Convert PPT → PDF"):
            convert_office(ppt, "PowerPoint PDF")
