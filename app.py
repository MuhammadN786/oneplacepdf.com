# OnePlacePDF — Edit Any PDF in One Place
# UI: single "Edit" tab with sub-tools (Delete, Insert, Text, Stamp, Redact/Highlight, Signature)

import os, io, shutil, subprocess, tempfile, zipfile
import streamlit as st
from PIL import Image
import img2pdf
from pypdf import PdfWriter, PdfReader
import fitz  # PyMuPDF
import pikepdf  # for unlocking protected PDFs
import pandas as pd

# ---- Page config / branding ----
st.set_page_config(
    page_title="OnePlacePDF — Edit Any PDF in One Place",  # browser tab text (SEO-friendly)
    page_icon="📄",                                        # or "assets/favicon.png"
    layout="wide",
    initial_sidebar_state="collapsed",
)

# (Optional) hide Streamlit's deploy toolbar to keep UI clean
st.markdown("<style>[data-testid='stToolbar']{visibility:hidden;height:0}</style>", unsafe_allow_html=True)

# App title + tagline
st.title("OnePlacePDF")
st.caption("Merge, split, compress, sign & convert — fast and private.")


# ---------------- Helpers ----------------
def read_pdf(file):
    """Open a PDF with pypdf and block encrypted ones (use Unlock tab)."""
    try:
        reader = PdfReader(file)
        if reader.is_encrypted:
            st.error("This PDF is password-protected. Use the Unlock tab first.")
            return None
        return reader
    except Exception as e:
        st.error(f"Could not read PDF: {e}")
        return None

def parse_page_spec(spec: str, total_pages: int):
    """Parse '1-3,5,7-9' -> list of 0-based page indices."""
    pages, spec = [], (spec or "").replace(" ", "")
    if not spec:
        return []
    for part in spec.split(","):
        if "-" in part:
            a, b = part.split("-", 1)
            if a.isdigit() and b.isdigit():
                start = max(1, int(a)); end = min(total_pages, int(b))
                if start <= end:
                    pages += list(range(start-1, end))
        elif part.isdigit():
            n = int(part)
            if 1 <= n <= total_pages:
                pages.append(n-1)
    seen, out = set(), []
    for p in pages:
        if p not in seen:
            out.append(p); seen.add(p)
    return out

def find_ghostscript():
    """Find Ghostscript on Windows/Linux/Mac."""
    for name in ["gswin64c", "gswin32c", "gs"]:
        p = shutil.which(name)
        if p:
            return p
    # common Windows location if PATH not set
    root = r"C:\Program Files\gs"
    if os.path.isdir(root):
        for r, _, files in os.walk(root):
            if "gswin64c.exe" in files:
                return os.path.join(r, "gswin64c.exe")
            if "gswin32c.exe" in files:
                return os.path.join(r, "gswin32c.exe")
    return None

def compress_with_gs(input_bytes: bytes, quality: str = "/ebook") -> bytes:
    exe = find_ghostscript()
    if not exe:
        raise RuntimeError("Ghostscript not found. Install Ghostscript and restart the app.")
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f_in, \
         tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f_out:
        f_in.write(input_bytes); f_in.flush()
        args = [exe, "-sDEVICE=pdfwrite", "-dCompatibilityLevel=1.4",
                f"-dPDFSETTINGS={quality}", "-dNOPAUSE", "-dQUIET", "-dBATCH",
                f"-sOutputFile={f_out.name}", f_in.name]
        proc = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if proc.returncode != 0:
            raise RuntimeError("Ghostscript failed.")
        return open(f_out.name, "rb").read()
def soffice_convert(input_bytes: bytes, in_suffix: str, out_fmt: str = "pdf") -> bytes:
    """Convert Office files with LibreOffice headless (docx/xlsx/pptx -> pdf, or pdf -> docx/xlsx via other fmt)."""
    with tempfile.TemporaryDirectory() as d:
        in_path  = os.path.join(d, f"in{in_suffix}")
        out_path = os.path.join(d, f"out.{out_fmt}")
        with open(in_path, "wb") as f:
            f.write(input_bytes)
        cmd = ["soffice", "--headless", "--convert-to", out_fmt, "--outdir", d, in_path]
        r = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        if r.returncode != 0 or not os.path.exists(out_path):
            raise RuntimeError("LibreOffice conversion failed.")
        with open(out_path, "rb") as f:
            return f.read()

# --------------- MAIN TABS ---------------
# ----- TOP-LEVEL NAV -----
tab_img2pdf, tab_merge, tab_split, tab_rotate, tab_reorder, tab_extract, tab_edit, \
tab_compress, tab_protect, tab_unlock, tab_pdf2img, tab_pdf2docx, tab_watermark, \
tab_pagenum, tab_office = st.tabs([
    "Images → PDF", "Merge", "Split", "Rotate", "Re-order", "Extract Text", "Edit",
    "Compress", "Protect (Password)", "Unlock (Password)", "PDF → Images",
    "PDF → DOCX", "Watermark (Text)", "Page Numbers", "Office ↔ PDF"       # 👈 NEW
])

# --- Images → PDF ---
with tab_img2pdf:
    st.subheader("JPG/PNG → PDF")
    files = st.file_uploader("Choose images", type=["jpg", "jpeg", "png"],
                             accept_multiple_files=True, key="img_uploader")
    quality = st.slider("JPEG quality", 60, 100, 90)
    if files and st.button("Convert to PDF", key="btn_img2pdf"):
        imgs = []
        for f in files:
            im = Image.open(f)
            if im.mode != "RGB": im = im.convert("RGB")
            b = io.BytesIO(); im.save(b, format="JPEG", quality=quality)
            imgs.append(b.getvalue())
        pdf_bytes = img2pdf.convert(imgs)
        st.success("Done! Download below.")
        st.download_button("⬇️ images-to-pdf.pdf", data=pdf_bytes,
                           file_name="images-to-pdf.pdf", mime="application/pdf")

# --- Merge ---
with tab_merge:
    st.subheader("Merge PDFs")
    pdfs = st.file_uploader("Choose PDF files", type=["pdf"],
                            accept_multiple_files=True, key="merge_files")
    if pdfs and st.button("Merge", key="btn_merge"):
        w = PdfWriter()
        for f in pdfs:
            r = read_pdf(f)
            if not r: st.stop()
            for p in r.pages: w.add_page(p)
        out = io.BytesIO(); w.write(out); out.seek(0)
        st.success("Merged!")
        st.download_button("⬇️ merged.pdf", out.getvalue(), file_name="merged.pdf", mime="application/pdf")

# --- Split ---
with tab_split:
    st.subheader("Split PDF by pages")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="split_pdf")
    if f:
        r = read_pdf(f)
        if r:
            total = len(r.pages)
            st.write(f"Total pages: **{total}**")
            spec = st.text_input("Pages to keep (e.g., 1-3,5,7-9)")
            if st.button("Create new PDF", key="btn_split"):
                idxs = parse_page_spec(spec, total)
                if not idxs:
                    st.warning("Enter a list like 1-3,5,7-9.")
                else:
                    w = PdfWriter()
                    for i in idxs: w.add_page(r.pages[i])
                    out = io.BytesIO(); w.write(out); out.seek(0)
                    st.success("Created!")
                    st.download_button("⬇️ split.pdf", out.getvalue(),
                                       file_name="split.pdf", mime="application/pdf")

# --- Rotate ---
with tab_rotate:
    st.subheader("Rotate pages")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="rotate_pdf")
    angle = st.selectbox("Rotate by", [90, 180, 270], index=0)
    scope = st.radio("Which pages?", ["All pages", "Specific pages"], index=0, horizontal=True)
    spec = st.text_input("Pages (e.g., 1,3-5)") if scope != "All pages" else ""
    if f and st.button("Apply rotation", key="btn_rotate"):
        r = read_pdf(f)
        if r:
            w = PdfWriter(); total = len(r.pages)
            targets = list(range(total)) if scope == "All pages" else parse_page_spec(spec, total)
            if scope != "All pages" and not targets:
                st.warning("Enter valid pages.")
            else:
                for i, p in enumerate(r.pages):
                    try: p.rotate(angle)
                    except Exception: pass
                    w.add_page(p)
                out = io.BytesIO(); w.write(out); out.seek(0)
                st.success("Rotated!")
                st.download_button("⬇️ rotated.pdf", out.getvalue(),
                                   file_name="rotated.pdf", mime="application/pdf")

# --- Re-order ---
with tab_reorder:
    st.subheader("Re-order pages")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="reorder_pdf")
    order_text = st.text_input("New order (comma-separated, 1-based). Example: 3,1,2")
    if f and st.button("Re-order", key="btn_reorder"):
        r = read_pdf(f)
        if r:
            total = len(r.pages)
            wanted = parse_page_spec(order_text, total)
            if len(wanted) != total:
                st.warning(f"List exactly {total} unique page numbers (e.g., 3,1,2).")
            else:
                w = PdfWriter()
                for idx in wanted: w.add_page(r.pages[idx])
                out = io.BytesIO(); w.write(out); out.seek(0)
                st.success("Re-ordered!")
                st.download_button("⬇️ reordered.pdf", out.getvalue(),
                                   file_name="reordered.pdf", mime="application/pdf")

# --- Extract Text ---
with tab_extract:
    st.subheader("Extract text")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="extract_pdf")
    if f and st.button("Extract", key="btn_extract"):
        r = read_pdf(f)
        if r:
            chunks = []
            for i, p in enumerate(r.pages, start=1):
                try: t = p.extract_text() or ""
                except Exception: t = ""
                chunks.append(f"--- Page {i} ---\n{t}\n")
            txt = "\n".join(chunks)
            st.success("Extracted!")
            st.download_button("⬇️ extracted-text.txt", txt.encode("utf-8"),
                               file_name="extracted-text.txt", mime="text/plain")

# ---------------- Edit (all tools) ----------------
with tab_edit:
    st.subheader("Edit PDF")
    base_pdf = st.file_uploader("Upload a PDF to edit", type=["pdf"], key="edit_pdf")
    tool = st.selectbox("Choose an action", [
        "Delete Pages", "Insert Pages", "Add Text",
        "Stamp Image/Logo", "Find & Redact / Highlight", "Signature"
    ], index=0)
    st.divider()

    if not base_pdf:
        st.info("Upload a PDF above to use the edit tools.")
    else:
        # ---- Delete Pages ----
        if tool == "Delete Pages":
            r = PdfReader(base_pdf)
            total = len(r.pages)
            st.caption(f"Total pages: **{total}**")
            spec = st.text_input("Pages to delete (e.g., 2,4-5)")
            if st.button("Delete pages", key="btn_edit_delete"):
                to_del = set(parse_page_spec(spec, total))
                if not to_del:
                    st.warning("Enter pages like `2,4-5`.")
                else:
                    w = PdfWriter()
                    for i, p in enumerate(r.pages):
                        if i not in to_del:
                            w.add_page(p)
                    out = io.BytesIO(); w.write(out); out.seek(0)
                    st.success("Pages deleted.")
                    st.download_button("⬇️ deleted-pages.pdf", out.getvalue(),
                                       file_name="deleted-pages.pdf", mime="application/pdf")

        # ---- Insert Pages ----
        elif tool == "Insert Pages":
            to_insert = st.file_uploader("PDF to insert (all its pages)", type=["pdf"], key="edit_insert_pdf")
            pos = st.number_input("Insert BEFORE page (1 … N+1)", min_value=1, value=1, step=1)
            if to_insert and st.button("Insert", key="btn_edit_insert"):
                rb = PdfReader(base_pdf); ri = PdfReader(to_insert)
                N = len(rb.pages); p = max(1, min(int(pos), N + 1)) - 1
                w = PdfWriter()
                for i in range(0, p): w.add_page(rb.pages[i])
                for pg in ri.pages: w.add_page(pg)
                for i in range(p, N): w.add_page(rb.pages[i])
                out = io.BytesIO(); w.write(out); out.seek(0)
                st.success(f"Inserted before page {p+1}.")
                st.download_button("⬇️ inserted.pdf", out.getvalue(),
                                   file_name="inserted.pdf", mime="application/pdf")

        # ---- Add Text ----
        elif tool == "Add Text":
            text = st.text_input("Text", "Sample text")
            font_size = st.slider("Font size", 8, 72, 18)
            place = st.selectbox("Position", [
                "Top-Left","Top-Center","Top-Right","Center",
                "Bottom-Left","Bottom-Center","Bottom-Right"])
            scope = st.radio("Apply to", ["All pages", "Specific pages"], horizontal=True)
            spec = st.text_input("Pages (e.g., 1,3-4)") if scope != "All pages" else ""
            if st.button("Add text", key="btn_edit_text"):
                data = base_pdf.read()
                doc = fitz.open(stream=data, filetype="pdf")
                total = len(doc)
                targets = list(range(total)) if scope == "All pages" else parse_page_spec(spec, total)
                if scope != "All pages" and not targets:
                    st.warning("Enter valid pages.")
                else:
                    m = 24
                    for i in targets:
                        page = doc[i]; w, h = page.rect.width, page.rect.height
                        boxes = {
                            "Top-Left":        fitz.Rect(m, m, w/2, m+40),
                            "Top-Center":      fitz.Rect(0, m, w, m+40),
                            "Top-Right":       fitz.Rect(w/2, m, w-m, m+40),
                            "Center":          fitz.Rect(0, h/2-20, w, h/2+20),
                            "Bottom-Left":     fitz.Rect(m, h-40-m, w/2, h-m),
                            "Bottom-Center":   fitz.Rect(0, h-40-m, w, h-m),
                            "Bottom-Right":    fitz.Rect(w/2, h-40-m, w-m, h-m),
                        }
                        page.insert_textbox(boxes[place], text, fontsize=font_size, color=(0,0,0),
                                            align=1 if "Center" in place else 0)
                    out = io.BytesIO(); doc.save(out); doc.close(); out.seek(0)
                    st.success("Text added.")
                    st.download_button("⬇️ text-added.pdf", out.getvalue(),
                                       file_name="text-added.pdf", mime="application/pdf")

        # ---- Stamp Image/Logo ----
        elif tool == "Stamp Image/Logo":
            img_file = st.file_uploader("Image (PNG/JPG)", type=["png","jpg","jpeg"], key="edit_stamp_img")
            scale = st.slider("Width as % of page", 5, 60, 20)
            place = st.selectbox("Position", ["Top-Left","Top-Right","Center","Bottom-Left","Bottom-Right"])
            scope = st.radio("Apply to", ["All pages", "Specific pages"], horizontal=True)
            spec = st.text_input("Pages (e.g., 1,4-5)") if scope != "All pages" else ""
            if img_file and st.button("Stamp image", key="btn_edit_stamp"):
                data = base_pdf.read()
                doc = fitz.open(stream=data, filetype="pdf")
                img = Image.open(img_file).convert("RGBA")
                b = io.BytesIO(); img.save(b, format="PNG"); img_bytes = b.getvalue()
                total = len(doc)
                targets = list(range(total)) if scope == "All pages" else parse_page_spec(spec, total)
                if scope != "All pages" and not targets:
                    st.warning("Enter valid pages.")
                else:
                    m = 20; iw, ih = img.size
                    for i in targets:
                        page = doc[i]; w, h = page.rect.width, page.rect.height
                        tw = (scale/100.0) * w; th = tw * (ih/iw)
                        x0 = m if "Left" in place else (w - tw - m if "Right" in place else (w - tw)/2)
                        y0 = m if "Top" in place else (h - th - m if "Bottom" in place else (h - th)/2)
                        rect = fitz.Rect(x0, y0, x0+tw, y0+th)
                        page.insert_image(rect, stream=img_bytes, keep_proportion=True)
                    out = io.BytesIO(); doc.save(out); doc.close(); out.seek(0)
                    st.success("Image stamped.")
                    st.download_button("⬇️ stamped.pdf", out.getvalue(),
                                       file_name="stamped.pdf", mime="application/pdf")

        # ---- Find & Redact / Highlight ----
        elif tool == "Find & Redact / Highlight":
            query = st.text_input("Text to find (exact match)")
            action = st.radio("Action", ["Redact (black box)", "Highlight"], horizontal=True)
            scope = st.radio("Apply to", ["All pages", "Specific pages"], horizontal=True)
            spec = st.text_input("Pages (e.g., 1,2-3)") if scope != "All pages" else ""
            if query and st.button("Apply", key="btn_edit_redact"):
                data = base_pdf.read()
                doc = fitz.open(stream=data, filetype="pdf")
                total = len(doc)
                targets = list(range(total)) if scope == "All pages" else parse_page_spec(spec, total)
                if scope != "All pages" and not targets:
                    st.warning("Enter valid pages.")
                else:
                    hits = 0
                    for i in targets:
                        page = doc[i]
                        for r in (page.search_for(query) or []):
                            hits += 1
                            if action.startswith("Redact"):
                                page.add_redact_annot(r, fill=(0,0,0))
                            else:
                                page.add_highlight_annot(r)
                    if action.startswith("Redact"):
                        doc.apply_redactions()
                    out = io.BytesIO(); doc.save(out); doc.close(); out.seek(0)
                    st.success(f"{action.split()[0]}ed {hits} occurrence(s).")
                    st.download_button("⬇️ edited.pdf", out.getvalue(),
                                       file_name="edited.pdf", mime="application/pdf")

        # ---- Signature ----
        else:
            import datetime
            sig = st.file_uploader("Signature image (PNG/JPG)", type=["png","jpg","jpeg"], key="edit_sig_img")
            place = st.selectbox("Position", ["Bottom-Right","Bottom-Left","Top-Right","Top-Left","Center"])
            width_pct = st.slider("Signature width (% of page)", 5, 40, 20)
            mode = st.radio("Apply to", ["Last page", "All pages", "Specific pages"], horizontal=True)
            pages_spec = st.text_input("Pages (e.g., 1,3-4)") if mode == "Specific pages" else ""
            printed = st.text_input("Printed name (optional)", "")
            add_date = st.checkbox("Add date", value=True)
            if sig and st.button("Place signature", key="btn_edit_signature"):
                data = base_pdf.read()
                doc = fitz.open(stream=data, filetype="pdf"); N = len(doc)
                img = Image.open(sig).convert("RGBA")
                b = io.BytesIO(); img.save(b, format="PNG"); sig_bytes = b.getvalue()
                iw, ih = img.size
                if mode == "All pages":
                    targets = list(range(N))
                elif mode == "Specific pages":
                    targets = parse_page_spec(pages_spec, N)
                    if not targets: st.warning("Enter valid pages like 1,3-5."); st.stop()
                else:
                    targets = [N-1]
                m = 18; placed = 0
                for i in targets:
                    page = doc[i]; w, h = page.rect.width, page.rect.height
                    tw = (width_pct/100.0)*w; th = tw*(ih/iw)
                    x0 = w - tw - m if "Right" in place else (m if "Left" in place else (w - tw)/2)
                    y0 = h - th - m if place.startswith("Bottom") else (m if place.startswith("Top") else (h - th)/2)
                    rect = fitz.Rect(x0, y0, x0+tw, y0+th)
                    page.insert_image(rect, stream=sig_bytes, keep_proportion=True)
                    line = printed or ""
                    if add_date:
                        today = datetime.date.today().strftime("%Y-%m-%d")
                        line = (line + "  " if line else "") + today
                    if line:
                        below = fitz.Rect(rect.x0, rect.y1+6, rect.x1, rect.y1+28)
                        above = fitz.Rect(rect.x0, rect.y0-28, rect.x1, rect.y0-6)
                        target = below if below.y1 <= h - m else above
                        page.insert_textbox(target, line, fontsize=10, color=(0,0,0), align=1)
                    placed += 1
                out = io.BytesIO(); doc.save(out); doc.close(); out.seek(0)
                st.success(f"Signature placed on {placed} page(s).")
                st.download_button("⬇️ signed.pdf", out.getvalue(),
                                   file_name="signed.pdf", mime="application/pdf")

# --- Compress ---
with tab_compress:
    st.subheader("Compress PDF (reduce size)")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="compress_pdf")
    preset = st.selectbox("Quality", [("/screen","Smallest"), ("/ebook","Balanced"), ("/prepress","High quality")],
                          index=1, format_func=lambda x: x[1])[0]
    gs_path = find_ghostscript()
    if not gs_path:
        st.info("Ghostscript not found. Install it to enable compression, then restart the app.")
    if f and gs_path and st.button("Compress", key="btn_compress"):
        try:
            out_bytes = compress_with_gs(f.read(), quality=preset)
            st.success("Compressed!")
            st.download_button("⬇️ compressed.pdf", out_bytes, file_name="compressed.pdf", mime="application/pdf")
        except Exception as e:
            st.error(str(e))

# --- Protect (set password) ---
with tab_protect:
    st.subheader("Protect PDF with password")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="protect_pdf")
    pwd = st.text_input("Password to set", type="password")
    if f and pwd and st.button("Protect", key="btn_protect"):
        try:
            r = PdfReader(f); w = PdfWriter()
            for p in r.pages: w.add_page(p)
            out = io.BytesIO(); w.encrypt(user_password=pwd, owner_password=pwd)
            w.write(out); out.seek(0)
            st.success("Encrypted!")
            st.download_button("⬇️ protected.pdf", out.getvalue(), file_name="protected.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"Failed: {e}")

# --- Unlock (needs current password) ---
with tab_unlock:
    st.subheader("Unlock PDF (remove password)")
    f = st.file_uploader("Upload a password-protected PDF", type=["pdf"], key="unlock_pdf")
    pwd = st.text_input("Current password", type="password")
    if f and pwd and st.button("Unlock", key="btn_unlock"):
        try:
            pdf = pikepdf.open(f, password=pwd)
            out = io.BytesIO(); pdf.save(out); out.seek(0)
            st.success("Unlocked!")
            st.download_button("⬇️ unlocked.pdf", out.getvalue(), file_name="unlocked.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"Wrong password or error: {e}")

# --- PDF → Images (PyMuPDF; no Poppler) ---
with tab_pdf2img:
    st.subheader("PDF → Images (PNG)")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="pdf_to_img")
    dpi = st.slider("DPI (quality)", 72, 300, 150)
    if f and st.button("Convert", key="btn_pdf2img"):
        data = f.read()
        doc = fitz.open(stream=data, filetype="pdf")
        zbuf = io.BytesIO()
        with zipfile.ZipFile(zbuf, "w", zipfile.ZIP_DEFLATED) as z:
            for i, page in enumerate(doc, start=1):
                zoom = dpi / 72.0
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                z.writestr(f"page-{i:02d}.png", pix.tobytes("png"))
        n = len(doc); doc.close(); zbuf.seek(0)
        st.success(f"Converted {n} page(s).")
        st.download_button("⬇️ pages.zip", zbuf.getvalue(), file_name="pdf-pages.zip", mime="application/zip")

# --- PDF → DOCX ---
with tab_pdf2docx:
    st.subheader("Convert PDF → DOCX (Word)")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="pdf_to_docx")
    c1, c2 = st.columns(2)
    start_page = c1.number_input("Start page (1-based)", min_value=1, value=1, step=1)
    end_page = c2.number_input("End page (0 = all)", min_value=0, value=0, step=1)
    st.caption("If the PDF is password-protected, unlock it first using the Unlock tab.")
    if f and st.button("Convert to DOCX", key="btn_pdf2docx"):
        try:
            from pdf2docx import Converter
        except Exception:
            st.error("Missing package `pdf2docx`. Install with:  pip install pdf2docx")
            st.stop()
        pdf_bytes = f.getvalue()
        fi = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
        fo = tempfile.NamedTemporaryFile(suffix=".docx", delete=False)
        fi.write(pdf_bytes); fi.flush()
        pdf_path, docx_path = fi.name, fo.name
        fi.close(); fo.close()
        start = max(0, int(start_page) - 1)
        end = None if int(end_page) == 0 else max(start, int(end_page) - 1)
        try:
            with st.spinner("Converting…"):
                cv = Converter(pdf_path); cv.convert(docx_path, start=start, end=end); cv.close()
            with open(docx_path, "rb") as fh:
                st.success("Converted! Download below.")
                st.download_button("⬇️ pdf-to-docx.docx", fh.read(),
                                   file_name="pdf-to-docx.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        except Exception as e:
            st.error(f"Conversion failed: {e}")
        finally:
            for p in (pdf_path, docx_path):
                try: os.remove(p)
                except: pass

# --- Watermark (Text) ---
with tab_watermark:
    st.subheader("Add a text watermark")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="wm_pdf")
    text = st.text_input("Watermark text", "CONFIDENTIAL")
    font_size = st.slider("Font size", 24, 120, 48)
    angle = st.slider("Angle (degrees)", 0, 90, 45)
    color = (0.7, 0.7, 0.7)  # light gray
    if f and st.button("Apply watermark", key="btn_wm"):
        data = f.read()
        doc = fitz.open(stream=data, filetype="pdf")
        for page in doc:
            w, h = page.rect.width, page.rect.height
            rect = fitz.Rect(0, 0, w, h)
            page.insert_textbox(rect, text, fontsize=font_size, color=color, rotate=angle, align=1)
        out = io.BytesIO(); doc.save(out); doc.close(); out.seek(0)
        st.success("Watermark applied!")
        st.download_button("⬇️ watermarked.pdf", out.getvalue(), file_name="watermarked.pdf", mime="application/pdf")

# --- Page Numbers ---
with tab_pagenum:
    st.subheader("Add page numbers")
    f = st.file_uploader("Upload a PDF", type=["pdf"], key="pn_pdf")
    style = st.selectbox("Style", ["1, 2, 3…", "1 / N"], index=1)
    font_size = st.slider("Font size", 10, 24, 12)
    if f and st.button("Add numbers", key="btn_pn"):
        data = f.read()
        doc = fitz.open(stream=data, filetype="pdf")
        total = len(doc)
        for i, page in enumerate(doc, start=1):
            w, h = page.rect.width, page.rect.height
            txt = f"{i}" if style == "1, 2, 3…" else f"{i} / {total}"
            rect = fitz.Rect(0, h - 36, w, h - 12)  # bottom-center
            page.insert_textbox(rect, txt, fontsize=font_size, color=(0, 0, 0), align=1)
        out = io.BytesIO(); doc.save(out); doc.close(); out.seek(0)
        st.success("Page numbers added!")
        st.download_button("⬇️ numbered.pdf", out.getvalue(), file_name="numbered.pdf", mime="application/pdf")
# Add to your tabs list names if needed:
# ... , "Word → PDF", "PowerPoint → PDF", "Excel → PDF", "PDF → Excel (tables)"

tab_word, tab_ppt, tab_xls, tab_pdf2xls = st.tabs(["Word → PDF", "PowerPoint → PDF", "Excel → PDF", "PDF → Excel (tables)"])

# ----- NEW: Office converters live only inside this tab -----
with tab_office:
    st.subheader("Office ↔ PDF")
    t_word, t_ppt, t_xls, t_pdf2xls = st.tabs(["Word → PDF", "PowerPoint → PDF",
                                               "Excel → PDF", "PDF → Excel (tables)"])

    # Word → PDF
    with t_word:
        st.subheader("Convert Word (DOC/DOCX) → PDF")
        f = st.file_uploader("Upload DOC or DOCX", type=["doc", "docx", "rtf", "odt"], key="doc2pdf")
        if f and st.button("Convert to PDF", key="btn_doc2pdf"):
            suffix = os.path.splitext(f.name)[1] or ".docx"
            try:
                pdf_bytes = soffice_convert(f.read(), suffix, out_fmt="pdf")
                st.success("Converted!")
                st.download_button("⬇️ download.pdf", pdf_bytes, "converted.pdf", "application/pdf")
            except Exception as e:
                st.error(f"Conversion failed: {e}")

    # PowerPoint → PDF
    with t_ppt:
        st.subheader("Convert PowerPoint (PPT/PPTX) → PDF")
        f = st.file_uploader("Upload PPT or PPTX", type=["ppt", "pptx", "odp"], key="ppt2pdf")
        if f and st.button("Convert to PDF", key="btn_ppt2pdf"):
            suffix = os.path.splitext(f.name)[1] or ".pptx"
            try:
                pdf_bytes = soffice_convert(f.read(), suffix, out_fmt="pdf")
                st.success("Converted!")
                st.download_button("⬇️ slides.pdf", pdf_bytes, "slides.pdf", "application/pdf")
            except Exception as e:
                st.error(f"Conversion failed: {e}")

    # Excel → PDF
    with t_xls:
        st.subheader("Convert Excel (XLS/XLSX) → PDF")
        f = st.file_uploader("Upload XLS or XLSX", type=["xls", "xlsx", "ods", "csv"], key="xls2pdf")
        if f and st.button("Convert to PDF", key="btn_xls2pdf"):
            suffix = os.path.splitext(f.name)[1] or ".xlsx"
            try:
                pdf_bytes = soffice_convert(f.read(), suffix, out_fmt="pdf")
                st.success("Converted!")
                st.download_button("⬇️ workbook.pdf", pdf_bytes, "workbook.pdf", "application/pdf")
            except Exception as e:
                st.error(f"Conversion failed: {e}")

    # PDF → Excel (tables)
    with t_pdf2xls:
        st.subheader("Extract tables: PDF → Excel (XLSX)")
        one_pdf = st.file_uploader("Upload a PDF with tables", type=["pdf"], key="pdf2xls")
        flavor = st.selectbox("Detection mode", ["lattice (lines)", "stream (no lines)"], index=0)
        pages = st.text_input("Pages (e.g., 1,3,5 or 1-4)", "all")
        if one_pdf and st.button("Extract tables", key="btn_pdf2xls"):
            try:
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as ftmp:
                    ftmp.write(one_pdf.read()); pdf_path = ftmp.name
                mode = "lattice" if flavor.startswith("lattice") else "stream"
                tables = camelot.read_pdf(pdf_path, pages=pages, flavor=mode)
                if tables.n == 0:
                    st.warning("No tables detected. Try switching detection mode or page range.")
                else:
                    xbuf = io.BytesIO()
                    with pd.ExcelWriter(xbuf, engine="openpyxl") as writer:
                        for i, t in enumerate(tables):
                            t.df.to_excel(writer, index=False, sheet_name=f"Table{i+1}")
                    st.success(f"Extracted {tables.n} table(s).")
                    st.download_button("⬇️ tables.xlsx", xbuf.getvalue(), "tables.xlsx",
                                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            except Exception as e:
                st.error(f"Extraction failed: {e}")


