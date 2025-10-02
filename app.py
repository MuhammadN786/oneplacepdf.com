# app.py — OnePlacePDF (recreated)
# Streamlit toolkit to merge, split, rotate, reorder, extract text, edit (delete/keep),
# compress, protect/unlock, convert PDF↔images, PDF→DOCX, add watermark text,
# add page numbers, Office→PDF.
#
# pip install: streamlit pypdf pymupdf Pillow img2pdf pdf2docx
# system (optional): ghostscript, libreoffice

import io, os, re, tempfile, subprocess, zipfile, shutil
from typing import List, Tuple

import streamlit as st
from pypdf import PdfReader, PdfWriter
from PIL import Image, ImageDraw, ImageFont
import img2pdf
import fitz  # PyMuPDF
from pdf2docx import Converter as Pdf2DocxConverter

# ---------- General helpers ----------
def save_upload(uf, suffix=""):
    fd, path = tempfile.mkstemp(suffix=suffix)
    with os.fdopen(fd, "wb") as f:
        f.write(uf.read())
    return path

def save_uploads(ufs, suffix=""):
    return [save_upload(u, suffix=suffix) for u in ufs]

def to_bytes(path: str) -> bytes:
    with open(path, "rb") as f:
        return f.read()

def parse_ranges(ranges: str, total_pages: int) -> List[int]:
    """Return 0-based page indices from a string like '1-3,5,7-9'."""
    result = []
    if not ranges.strip():
        return list(range(total_pages))
    parts = [p.strip() for p in ranges.split(",") if p.strip()]
    for p in parts:
        if "-" in p:
            a, b = p.split("-", 1)
            start = int(a) if a else 1
            end = int(b) if b else total_pages
        else:
            start = int(p)
            end = start
        start = max(1, start)
        end = min(total_pages, end)
        result.extend(list(range(start - 1, end)))
    # de-dup but keep order
    seen = set()
    ordered = []
    for i in result:
        if 0 <= i < total_pages and i not in seen:
            seen.add(i)
            ordered.append(i)
    return ordered

def download_button(label, data, filename, mime="application/octet-stream"):
    st.download_button(label, data=data, file_name=filename, mime=mime)

st.set_page_config(page_title="OnePlacePDF — Edit Any PDF in One Place", page_icon="📄", layout="wide")
st.title("OnePlacePDF — Edit Any PDF in One Place")
st.caption("Merge, split, compress, sign & convert — fast and private.")

tabs = st.tabs([
    "Images → PDF", "Merge", "Split", "Rotate", "Re-order", "Extract Text",
    "Edit", "Compress", "Protect (Password)", "Unlock (Password)",
    "PDF → Images", "PDF → DOCX", "Watermark (Text)", "Page Numbers", "Office ↔ PDF"
])

# ---------- Images → PDF ----------
with tabs[0]:
    st.subheader("Convert Images to a single PDF")
    imgs = st.file_uploader("Upload JPG/PNG (ordered)", type=["jpg","jpeg","png"], accept_multiple_files=True)
    if imgs and st.button("Create PDF"):
        paths = save_uploads(imgs)
        tmpdir = tempfile.mkdtemp(prefix="img2pdf_")
        jpgs = []
        try:
            for p in paths:
                with Image.open(p) as im:
                    if im.mode not in ("RGB","L"):
                        im = im.convert("RGB")
                    outp = os.path.join(tmpdir, os.path.basename(p) + ".jpg")
                    im.save(outp, "JPEG", quality=95)
                    jpgs.append(outp)
            out_pdf = os.path.join(tmpdir, "images.pdf")
            with open(out_pdf, "wb") as f:
                f.write(img2pdf.convert(jpgs))
            download_button("Download images.pdf", to_bytes(out_pdf), "images.pdf", "application/pdf")
        finally:
            shutil.rmtree(tmpdir, ignore_errors=True)

# ---------- Merge ----------
with tabs[1]:
    st.subheader("Merge PDFs")
    pdfs = st.file_uploader("Upload PDFs (in order)", type=["pdf"], accept_multiple_files=True)
    if pdfs and st.button("Merge Now"):
        writer = PdfWriter()
        paths = save_uploads(pdfs, suffix=".pdf")
        for p in paths:
            r = PdfReader(p)
            for pg in r.pages:
                writer.add_page(pg)
        out = io.BytesIO()
        writer.write(out)
        download_button("Download merged.pdf", out.getvalue(), "merged.pdf", "application/pdf")

# ---------- Split ----------
with tabs[2]:
    st.subheader("Split PDF by ranges")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="splitpdf")
    ranges = st.text_input("Page ranges (e.g., 1-3,5,7-9) • leave blank for all")
    if pdf and st.button("Split Now"):
        path = save_upload(pdf, suffix=".pdf")
        reader = PdfReader(path)
        idxs = parse_ranges(ranges, len(reader.pages)) or list(range(len(reader.pages)))
        # create one output per contiguous run
        out_zip = io.BytesIO()
        with zipfile.ZipFile(out_zip, "w") as zf:
            if not ranges.strip():
                # just save as is
                w = PdfWriter()
                for pg in reader.pages:
                    w.add_page(pg)
                b = io.BytesIO(); w.write(b)
                zf.writestr("split_all.pdf", b.getvalue())
            else:
                # group into segments based on commas/ranges order
                segments = []
                # rebuild from text to preserve intended segments
                parts = [p.strip() for p in ranges.split(",") if p.strip()]
                for pi, ptxt in enumerate(parts, start=1):
                    if "-" in ptxt:
                        a,b = ptxt.split("-",1)
                        start = int(a) if a else 1
                        end = int(b) if b else len(reader.pages)
                        segment = list(range(start-1, min(end, len(reader.pages))))
                    else:
                        v = int(ptxt); segment = [v-1]
                    segment = [i for i in segment if 0 <= i < len(reader.pages)]
                    if not segment: 
                        continue
                    w = PdfWriter()
                    for i in segment:
                        w.add_page(reader.pages[i])
                    b = io.BytesIO(); w.write(b)
                    name = f"split_part_{pi}_{ptxt.replace(' ','')}.pdf"
                    zf.writestr(name, b.getvalue())
        out_zip.seek(0)
        download_button("Download splits.zip", out_zip.getvalue(), "splits.zip", "application/zip")

# ---------- Rotate ----------
with tabs[3]:
    st.subheader("Rotate pages")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="rotatepdf")
    deg = st.selectbox("Rotate by", [90, 180, 270], index=0)
    pages = st.text_input("Which pages? (e.g., 2,5-7; blank = all)")
    if pdf and st.button("Rotate Now"):
        path = save_upload(pdf, suffix=".pdf")
        r = PdfReader(path); w = PdfWriter()
        rotate_idxs = set(parse_ranges(pages, len(r.pages))) if pages.strip() else set(range(len(r.pages)))
        for i, pg in enumerate(r.pages):
            if i in rotate_idxs:
                pg.rotate(deg)
            w.add_page(pg)
        out = io.BytesIO(); w.write(out)
        download_button("Download rotated.pdf", out.getvalue(), "rotated.pdf", "application/pdf")

# ---------- Re-order ----------
with tabs[4]:
    st.subheader("Re-order pages")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="reorderpdf")
    order = st.text_input("New order (e.g., 3,1,2 or 1-3,5,4). Omitted pages are removed; duplicates allowed.")
    if pdf and order and st.button("Re-order Now"):
        path = save_upload(pdf, suffix=".pdf")
        r = PdfReader(path); w = PdfWriter()
        idxs = parse_ranges(order, len(r.pages))
        for i in idxs:
            w.add_page(r.pages[i])
        out = io.BytesIO(); w.write(out)
        download_button("Download reordered.pdf", out.getvalue(), "reordered.pdf", "application/pdf")

# ---------- Extract Text ----------
with tabs[5]:
    st.subheader("Extract text from PDF")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="extractpdf")
    pages = st.text_input("Pages to extract (blank = all)")
    if pdf and st.button("Extract"):
        path = save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)
        idxs = parse_ranges(pages, len(doc)) if pages.strip() else list(range(len(doc)))
        text = []
        for i in idxs:
            text.append(f"--- Page {i+1} ---\n{doc[i].get_text()}")
        st.text_area("Text", value="\n\n".join(text), height=400)
        st.download_button("Download .txt", data="\n\n".join(text), file_name="extracted.txt")

# ---------- Edit (Delete/Keep) ----------
with tabs[6]:
    st.subheader("Edit: Delete / Keep specific pages")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="editpdf")
    keep = st.text_input("Keep only pages (e.g., 1-3,7). Leave blank to delete pages below.")
    delete = st.text_input("Delete pages (e.g., 4-6). Ignored if 'Keep only' is set.")
    if pdf and st.button("Apply Edit"):
        path = save_upload(pdf, suffix=".pdf")
        r = PdfReader(path); w = PdfWriter()
        if keep.strip():
            idxs = set(parse_ranges(keep, len(r.pages)))
        else:
            del_idxs = set(parse_ranges(delete, len(r.pages))) if delete.strip() else set()
            idxs = [i for i in range(len(r.pages)) if i not in del_idxs]
        if not idxs:
            st.error("No pages to keep.")
        else:
            for i in (idxs if isinstance(idxs, list) else sorted(list(idxs))):
                w.add_page(r.pages[i])
            out = io.BytesIO(); w.write(out)
            download_button("Download edited.pdf", out.getvalue(), "edited.pdf", "application/pdf")

# ---------- Compress ----------
with tabs[7]:
    st.subheader("Compress PDF (Ghostscript)")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="compresspdf")
    preset_label = st.selectbox("Quality preset", ["/screen (smallest)", "/ebook", "/printer", "/prepress", "/default"], index=1)
    if pdf and st.button("Compress Now"):
        preset = preset_label.split(" ")[0]
        in_path = save_upload(pdf, suffix=".pdf")
        out_dir = tempfile.mkdtemp(prefix="gs_"); out_path = os.path.join(out_dir, "compressed.pdf")
        try:
            cmd = [
                "gs", "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={preset}", "-dNOPAUSE", "-dQUIET", "-dBATCH",
                f"-sOutputFile={out_path}", in_path
            ]
            subprocess.run(cmd, check=True)
            download_button("Download compressed.pdf", to_bytes(out_path), "compressed.pdf", "application/pdf")
        except Exception as e:
            st.error("Ghostscript failed or not installed in this environment.")
        finally:
            shutil.rmtree(out_dir, ignore_errors=True)

# ---------- Protect ----------
with tabs[8]:
    st.subheader("Protect (encrypt) PDF with a password")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="protectpdf")
    user_pwd = st.text_input("Password", type="password")
    owner_pwd = st.text_input("Owner password (optional)", type="password")
    if pdf and user_pwd and st.button("Protect Now"):
        path = save_upload(pdf, suffix=".pdf")
        r = PdfReader(path); w = PdfWriter()
        for p in r.pages: w.add_page(p)
        w.encrypt(user_password=user_pwd, owner_password=owner_pwd or user_pwd)
        out = io.BytesIO(); w.write(out)
        download_button("Download protected.pdf", out.getvalue(), "protected.pdf", "application/pdf")

# ---------- Unlock ----------
with tabs[9]:
    st.subheader("Unlock (remove password) — you must know the current password")
    pdf = st.file_uploader("Upload protected PDF", type=["pdf"], key="unlockpdf")
    pwd = st.text_input("Current password", type="password")
    if pdf and pwd and st.button("Unlock Now"):
        path = save_upload(pdf, suffix=".pdf")
        r = PdfReader(path)
        try:
            if r.is_encrypted:
                r.decrypt(pwd)
        except Exception:
            st.error("Incorrect password or unsupported encryption.")
        else:
            w = PdfWriter()
            for p in r.pages: w.add_page(p)
            out = io.BytesIO(); w.write(out)
            download_button("Download unlocked.pdf", out.getvalue(), "unlocked.pdf", "application/pdf")

# ---------- PDF → Images ----------
with tabs[10]:
    st.subheader("PDF to Images (PNG) → ZIP")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pdf2img")
    dpi = st.slider("DPI (resolution)", 72, 300, 150, step=6)
    if pdf and st.button("Convert"):
        path = save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)
        mem = io.BytesIO()
        with zipfile.ZipFile(mem, "w") as z:
            for i, page in enumerate(doc, start=1):
                pix = page.get_pixmap(dpi=dpi)
                z.writestr(f"page_{i:03d}.png", pix.tobytes("png"))
        mem.seek(0)
        download_button("Download pages.zip", mem.getvalue(), "pages.zip", "application/zip")

# ---------- PDF → DOCX ----------
with tabs[11]:
    st.subheader("Convert PDF → DOCX (Word)")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pdf2docx")
    start = st.number_input("Start page (1-based)", min_value=1, value=1)
    end = st.number_input("End page (0 = all)", min_value=0, value=0)
    if pdf and st.button("Convert to DOCX"):
        in_path = save_upload(pdf, suffix=".pdf")
        out_dir = tempfile.mkdtemp(prefix="pdf2docx_")
        out_path = os.path.join(out_dir, "converted.docx")
        try:
            cv = Pdf2DocxConverter(in_path)
            if end == 0:
                cv.convert(out_path, start=start-1, end=None)
            else:
                cv.convert(out_path, start=start-1, end=end-1)
            cv.close()
            download_button("Download converted.docx", to_bytes(out_path), "converted.docx",
                            "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception:
            st.error("Conversion failed (pdf2docx may not perfectly convert scanned or complex PDFs).")
        finally:
            shutil.rmtree(out_dir, ignore_errors=True)

# ---------- Watermark (Text) ----------
with tabs[12]:
    st.subheader("Add text watermark")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="wm_pdf")
    text = st.text_input("Watermark text", "CONFIDENTIAL")
    opacity = st.slider("Opacity", 10, 90, 20)
    scale = st.slider("Size (relative)", 20, 200, 100)
    angle = st.slider("Angle", 0, 360, 45)
    if pdf and text and st.button("Apply Watermark"):
        path = save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)
        for page in doc:
            rect = page.rect
            fontsize = rect.width * (scale/100) / len(text) * 2.0
            page.insert_text(
                rect.center, text,
                fontsize=fontsize, rotate=angle,
                color=(0,0,0),  # black
                fill_opacity=opacity/100.0,
                render_mode=0,
                overlay=True,
                align=1
            )
        out = io.BytesIO(); doc.save(out); doc.close()
        download_button("Download watermarked.pdf", out.getvalue(), "watermarked.pdf", "application/pdf")

# ---------- Page Numbers ----------
with tabs[13]:
    st.subheader("Add page numbers (footer)")
    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pnums_pdf")
    position = st.selectbox("Position", ["Left","Center","Right"], index=1)
    startnum = st.number_input("Start from number", min_value=1, value=1)
    if pdf and st.button("Add Numbers"):
        path = save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)
        for i, page in enumerate(doc, start=0):
            num = f"{startnum + i}"
            rect = page.rect
            y = rect.y1 - 20  # 20pt up from bottom
            if position == "Left":
                x = rect.x0 + 36
                align = 0
            elif position == "Center":
                x = rect.x0 + rect.width/2
                align = 1
            else:
                x = rect.x1 - 36
                align = 2
            page.insert_text((x, y), num, fontsize=10, color=(0,0,0), align=align)
        out = io.BytesIO(); doc.save(out); doc.close()
        download_button("Download numbered.pdf", out.getvalue(), "numbered.pdf", "application/pdf")

# ---------- Office ↔ PDF ----------
with tabs[14]:
    st.subheader("Office → PDF (LibreOffice headless)")
    office = st.file_uploader("Upload DOCX/XLSX/PPTX", type=["doc","docx","xls","xlsx","ppt","pptx"])
    col1, col2 = st.columns(2)
    with col1:
        if office and st.button("Convert to PDF"):
            in_path = save_upload(office)
            out_dir = tempfile.mkdtemp(prefix="soffice_")
            try:
                cmd = ["soffice","--headless","--convert-to","pdf","--outdir",out_dir, in_path]
                subprocess.run(cmd, check=True)
                base = os.path.splitext(os.path.basename(office.name))[0] + ".pdf"
                out_path = os.path.join(out_dir, base)
                if not os.path.exists(out_path):
                    st.error("LibreOffice failed to convert.")
                else:
                    download_button(f"Download {base}", to_bytes(out_path), base, "application/pdf")
            except Exception:
                st.error("LibreOffice not available in this environment.")
            finally:
                shutil.rmtree(out_dir, ignore_errors=True)

    st.divider()
    st.subheader("PDF → Office notice")
    st.caption("Direct PDF→DOCX is provided in the **PDF → DOCX** tab above.")

st.markdown("---")
st.caption("If a PDF is password-protected, unlock it first using the Unlock tab. Files are processed in memory and discarded.")
