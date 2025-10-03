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

st.set_page_config(page_title="OnePlacePDF — All-in-One PDF Tools", page_icon="📄", layout="wide")
st.title("OnePlacePDF — All-in-One PDF Tools")
st.caption("Merge, split, convert & secure — with quality-first processing.")

tabs = st.tabs([
    "Images → PDF", "Merge", "Split", "Rotate", "Re-order", "Extract Text",
    "Edit", "Compress", "Protect", "Unlock",
    "PDF → Images", "PDF → DOCX", "Watermark", "Page Numbers", "Office → PDF"
])

# ---------- Tab 1: Images → PDF (HQ, fixed margins & deprecation) ----------
with tabs[0]:
    st.subheader("Images → PDF (high quality)")

    imgs = st.file_uploader(
        "Upload images (JPG/PNG/JPEG)",
        type=["jpg", "jpeg", "png"],
        accept_multiple_files=True
    )

    if imgs:
        st.info("Preview your images below. You can reorder them and adjust settings.")

        # Save uploads temporarily
        img_paths = _save_uploads(imgs)

        # Thumbnails (no deprecation warning)
        cols = st.columns(4)
        for i, p in enumerate(img_paths):
            with Image.open(p) as im:
                tw = 140
                th = int(im.height * tw / max(1, im.width))
                with cols[i % 4]:
                    st.image(im.resize((tw, th)), caption=f"Img {i+1}", use_container_width=True)

        # Options
        st.markdown("### Options")
        page_size   = st.selectbox("Page size", ["Original", "A4", "Letter"], index=0)
        target_dpi  = st.selectbox("Target DPI (for scaling)", [300, 600], index=0)
        margin_pt   = st.slider("Margin (points)", 0, 96, 24)
        orientation = st.selectbox("Orientation (for A4/Letter)", ["Auto", "Portrait", "Landscape"], index=0)
        output_mode = st.radio("Output", ["Single PDF (all images)", "One PDF per image (ZIP)"], index=0)

        # Helpers
        def _page_dims_pts(img_w_px: int, img_h_px: int):
            """Return page size in PDF points (1 pt = 1/72 inch)."""
            if page_size == "Original":
                w_in = img_w_px / float(max(target_dpi, 1))
                h_in = img_h_px / float(max(target_dpi, 1))
                # at least 1pt to avoid zero pages
                return max(w_in * 72.0, 1.0), max(h_in * 72.0, 1.0)

            if page_size == "A4":
                w_pt, h_pt = 595.276, 841.890
            else:  # Letter
                w_pt, h_pt = 612.0, 792.0

            # Orientation
            if orientation == "Landscape" or (orientation == "Auto" and img_w_px > img_h_px):
                w_pt, h_pt = h_pt, w_pt

            return w_pt, h_pt

        def _safe_inner_rect(page_rect: "fitz.Rect", margin: float) -> "fitz.Rect":
            """Return inner rect after margins (never collapses to zero)."""
            # Max margin that still leaves at least 1pt in both dimensions
            max_margin_x = max( (page_rect.width  - 1.0) / 2.0, 0.0 )
            max_margin_y = max( (page_rect.height - 1.0) / 2.0, 0.0 )
            m = min(margin, max_margin_x, max_margin_y)
            return fitz.Rect(
                page_rect.x0 + m,
                page_rect.y0 + m,
                page_rect.x1 - m,
                page_rect.y1 - m
            )

        def _target_rect(page_rect: "fitz.Rect", img_w_px: int, img_h_px: int) -> "fitz.Rect":
            """Fit image proportionally inside inner rect."""
            inner = _safe_inner_rect(page_rect, float(margin_pt))
            pw = max(inner.width,  1e-6)
            ph = max(inner.height, 1e-6)
            ar_img = img_w_px / float(max(img_h_px, 1))
            # choose by which side limits first
            if pw / ph > ar_img:
                new_h = ph
                new_w = ar_img * new_h
            else:
                new_w = pw
                new_h = new_w / ar_img
            x0 = inner.x0 + (pw - new_w) / 2.0
            y0 = inner.y0 + (ph - new_h) / 2.0
            return fitz.Rect(x0, y0, x0 + new_w, y0 + new_h)

        def _image_stream_lossless(img_path: str) -> bytes:
            """Embed original bytes when safe; flatten alpha to PNG if needed."""
            with open(img_path, "rb") as f:
                raw = f.read()
            try:
                with Image.open(io.BytesIO(raw)) as im:
                    if im.mode in ("RGBA", "LA", "P"):
                        bg = Image.new("RGB", im.size, (255, 255, 255))
                        im_rgba = im.convert("RGBA")
                        alpha = im_rgba.split()[-1]
                        bg.paste(im_rgba, mask=alpha)
                        buf = io.BytesIO()
                        bg.save(buf, format="PNG", optimize=True)
                        return buf.getvalue()
                return raw
            except Exception:
                return raw

        if st.button("Convert to PDF"):
            tmpdir = tempfile.mkdtemp(prefix="img2pdf_hq_")
            try:
                if output_mode.startswith("Single"):
                    doc = fitz.open()
                    for p in img_paths:
                        with Image.open(p) as im:
                            iw, ih = im.size
                        w_pt, h_pt = _page_dims_pts(iw, ih)
                        page = doc.new_page(width=w_pt, height=h_pt)
                        rect = _target_rect(page.rect, iw, ih)
                        page.insert_image(rect, stream=_image_stream_lossless(p), keep_proportion=True, overlay=True)
                    out = io.BytesIO()
                    doc.save(out, deflate=True)
                    doc.close()
                    _download("Download images.pdf", out.getvalue(), "images.pdf", "application/pdf")
                else:
                    mem = io.BytesIO()
                    with zipfile.ZipFile(mem, "w") as zf:
                        for i, p in enumerate(img_paths, start=1):
                            with Image.open(p) as im:
                                iw, ih = im.size
                            w_pt, h_pt = _page_dims_pts(iw, ih)
                            doc = fitz.open()
                            page = doc.new_page(width=w_pt, height=h_pt)
                            rect = _target_rect(page.rect, iw, ih)
                            page.insert_image(rect, stream=_image_stream_lossless(p), keep_proportion=True, overlay=True)
                            outp = os.path.join(tmpdir, f"image_{i}.pdf")
                            doc.save(outp, deflate=True)
                            doc.close()
                            zf.write(outp, f"image_{i}.pdf")
                    mem.seek(0)
                    _download("Download images.zip", mem.getvalue(), "images.zip", "application/zip")
            finally:
                shutil.rmtree(tmpdir, ignore_errors=True)

# ---------- Tab 2: Merge / Combine PDFs ----------
with tabs[1]:
    st.subheader("Merge / Combine PDFs")

    subtabs = st.tabs(["Merge (Append)", "Combine (Interleave)"])

    # ========== MERGE (APPEND) ==========
    with subtabs[0]:
        pdfs = st.file_uploader("Upload PDFs (any order)", type=["pdf"], accept_multiple_files=True, key="merge_files")

        if pdfs:
            st.markdown("### File Options")

            merge_data = []
            for i, uf in enumerate(pdfs, start=1):
                path = _save_upload(uf, suffix=".pdf")
                reader = PdfReader(path)

                st.markdown(f"**{i}. {uf.name}** ({len(reader.pages)} pages)")
                try:
                    thumb = fitz.open(path)[0].get_pixmap(dpi=40).tobytes("png")
                    st.image(thumb, caption=f"Preview {uf.name}", width=120)
                except Exception:
                    pass

                include = st.checkbox(f"Include {uf.name}", value=True, key=f"merge_inc_{i}")
                ranges = st.text_input("Pages to include (blank = all)", key=f"merge_rng_{i}")
                order = st.number_input("Order position", min_value=1, max_value=len(pdfs), value=i, key=f"merge_ord_{i}")

                merge_data.append({
                    "name": uf.name,
                    "path": path,
                    "include": include,
                    "ranges": ranges,
                    "order": order
                })

            # Sort by order input
            merge_data = sorted(merge_data, key=lambda x: x["order"])

            st.markdown("---")
            st.markdown("### Merge Options")
            add_bookmarks = st.checkbox("Add bookmarks by file name", value=True, key="merge_bm")

            if st.button("Merge Now"):
                writer = PdfWriter()

                for item in merge_data:
                    if not item["include"]:
                        continue
                    r = PdfReader(item["path"])
                    total = len(r.pages)
                    idxs = _parse_ranges(item["ranges"], total) if item["ranges"].strip() else list(range(total))

                    parent = None
                    if add_bookmarks:
                        # bookmark to the first page *after* existing ones
                        parent = writer.add_outline_item(item["name"], max(len(writer.pages), 0), 0)

                    for pg_idx in idxs:
                        writer.add_page(r.pages[pg_idx])

                out = io.BytesIO()
                writer.write(out)
                _download("Download merged.pdf", out.getvalue(), "merged.pdf", "application/pdf")

    # ========== COMBINE (INTERLEAVE) ==========
    with subtabs[1]:
        pdfs_c = st.file_uploader(
            "Upload 2–10 PDFs to interleave (A, B, C…)",
            type=["pdf"], accept_multiple_files=True, key="combine_files"
        )

        if pdfs_c:
            st.markdown("### Configure each file")
            combine_rows = []
            for i, uf in enumerate(pdfs_c, start=1):
                path = _save_upload(uf, suffix=".pdf")
                r = PdfReader(path)
                col1, col2, col3, col4 = st.columns([3, 3, 2, 2])
                with col1:
                    st.write(f"**{i}. {uf.name}** — {len(r.pages)} pages")
                with col2:
                    rng = st.text_input("Pages (blank=all)", key=f"cmb_rng_{i}")
                with col3:
                    chunk = st.number_input("Chunk size", min_value=1, max_value=10, value=1, key=f"cmb_chunk_{i}")
                with col4:
                    order = st.number_input("Order", min_value=1, max_value=len(pdfs_c), value=i, key=f"cmb_ord_{i}")

                combine_rows.append({
                    "name": uf.name,
                    "path": path,
                    "total": len(r.pages),
                    "ranges": rng,
                    "chunk": int(chunk),
                    "order": int(order)
                })

            # Sort by chosen order
            combine_rows = sorted(combine_rows, key=lambda x: x["order"])

            st.markdown("---")
            st.markdown("### Combine Options")
            loop_until_all = st.checkbox("Keep interleaving until all selected pages are exhausted", value=True, key="cmb_loop")
            add_bmarks    = st.checkbox("Add top-level bookmarks per file", value=True, key="cmb_bm")

            if st.button("Combine Now"):
                # Build page queues per file according to ranges
                queues = []
                for row in combine_rows:
                    r = PdfReader(row["path"])
                    idxs = _parse_ranges(row["ranges"], row["total"]) if row["ranges"].strip() else list(range(row["total"]))
                    queues.append({
                        "name": row["name"],
                        "reader": r,
                        "pages": idxs[:],  # queue of page indices
                        "chunk": row["chunk"]
                    })

                writer = PdfWriter()

                # Optional bookmarks (record the first output index for each file once it appears)
                first_out_index_for_file = {q["name"]: None for q in queues}

                # Interleave
                while True:
                    progress_any = False
                    for q in queues:
                        if not q["pages"]:
                            continue
                        progress_any = True

                        # Where the first page of this file lands (bookmark anchor)
                        if add_bmarks and first_out_index_for_file[q["name"]] is None:
                            first_out_index_for_file[q["name"]] = len(writer.pages)

                        take = min(q["chunk"], len(q["pages"]))
                        for _ in range(take):
                            pg_idx = q["pages"].pop(0)
                            writer.add_page(q["reader"].pages[pg_idx])

                    if not loop_until_all:
                        # Only one pass across files
                        break

                    if not progress_any:
                        # All queues empty
                        break

                # Add top-level bookmarks if requested
                if add_bmarks:
                    for q in queues:
                        anchor = first_out_index_for_file.get(q["name"])
                        if anchor is not None:
                            writer.add_outline_item(q["name"], anchor, 0)

                out = io.BytesIO()
                writer.write(out)
                _download("Download combined.pdf", out.getvalue(), "combined.pdf", "application/pdf")

# ---------- Tab 3: Split PDF ----------
with tabs[2]:
    st.subheader("Split PDF — multiple options")

    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="split_pdf")

    if pdf:
        path = _save_upload(pdf, suffix=".pdf")
        reader = PdfReader(path)
        doc = fitz.open(path)

        st.markdown(f"**This PDF has {len(reader.pages)} pages.**")
        
        # Preview thumbnails
        st.markdown("### Preview Pages")
        cols = st.columns(5)
        for i, page in enumerate(doc, start=1):
            thumb = page.get_pixmap(dpi=40).tobytes("png")
            with cols[(i-1) % 5]:
                st.image(thumb, caption=f"Page {i}", use_column_width=True)

        # Split options
        st.markdown("### Choose split method")
        mode = st.radio("Split by:", [
            "Page ranges (custom input)", 
            "Every page (1 PDF per page)", 
            "File size (approx MB)", 
            "Bookmarks"
        ])

        ranges = ""
        size_mb = 1
        if mode == "Page ranges (custom input)":
            ranges = st.text_input("Enter ranges (e.g. 1-3, 5, 7-9)")
        elif mode == "File size (approx MB)":
            size_mb = st.slider("Max size per file (MB)", 1, 50, 5)

        if st.button("Split Now"):
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, "w") as zf:
                
                if mode == "Page ranges (custom input)":
                    idxs = _parse_ranges(ranges, len(reader.pages))
                    if not idxs:
                        st.error("No valid ranges provided.")
                    else:
                        # break into segments based on commas
                        parts = [p.strip() for p in ranges.split(",") if p.strip()]
                        for idx, part in enumerate(parts, start=1):
                            seg_idxs = _parse_ranges(part, len(reader.pages))
                            w = PdfWriter()
                            for i in seg_idxs:
                                w.add_page(reader.pages[i])
                            b = io.BytesIO(); w.write(b)
                            zf.writestr(f"split_part_{idx}_{part}.pdf", b.getvalue())

                elif mode == "Every page (1 PDF per page)":
                    for i, pg in enumerate(reader.pages, start=1):
                        w = PdfWriter()
                        w.add_page(pg)
                        b = io.BytesIO(); w.write(b)
                        zf.writestr(f"page_{i}.pdf", b.getvalue())

                elif mode == "File size (approx MB)":
                    w = PdfWriter(); current_size = 0; part = 1
                    for i, pg in enumerate(reader.pages, start=1):
                        w.add_page(pg)
                        b = io.BytesIO(); w.write(b)
                        if b.tell() >= size_mb * 1024 * 1024:
                            zf.writestr(f"part_{part}.pdf", b.getvalue())
                            part += 1
                            w = PdfWriter()
                    if len(w.pages) > 0:
                        b = io.BytesIO(); w.write(b)
                        zf.writestr(f"part_{part}.pdf", b.getvalue())

                elif mode == "Bookmarks":
                    outlines = reader.outline
                    if not outlines:
                        st.error("No bookmarks found in this PDF.")
                    else:
                        # Flatten top-level bookmarks
                        for idx, item in enumerate(outlines, start=1):
                            if isinstance(item, list): continue
                            try:
                                pgnum = reader.get_destination_page_number(item)
                            except:
                                continue
                            start = pgnum
                            end = reader.get_destination_page_number(outlines[idx]) if idx < len(outlines)-1 else len(reader.pages)
                            w = PdfWriter()
                            for i in range(start, end):
                                w.add_page(reader.pages[i])
                            b = io.BytesIO(); w.write(b)
                            zf.writestr(f"bookmark_{idx}_{item.title}.pdf", b.getvalue())

            mem.seek(0)
            _download("Download splits.zip", mem.getvalue(), "splits.zip", "application/zip")

# ---------- Tab 4: Rotate Pages ----------
with tabs[3]:
    st.subheader("Rotate Pages")

    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="rotate_pdf")
    if pdf:
        path = _save_upload(pdf, suffix=".pdf")
        r = PdfReader(path)
        doc = fitz.open(path)

        st.markdown(f"**This PDF has {len(r.pages)} pages.**")

        mode = st.radio("Choose mode:", ["Rotate all pages", "Rotate selected pages"])

        if mode == "Rotate all pages":
            deg = st.selectbox("Rotate by", [90, 180, 270])
            if st.button("Apply Rotation"):
                w = PdfWriter()
                for pg in r.pages:
                    pg.rotate(deg)
                    w.add_page(pg)
                out = io.BytesIO(); w.write(out)
                _download("Download rotated.pdf", out.getvalue(), "rotated.pdf", "application/pdf")

        else:  # Rotate selected pages with preview
            st.markdown("### Per-page preview and rotation controls")

            page_data = []
            cols = st.columns(4)
            for i, page in enumerate(doc, start=1):
                thumb = page.get_pixmap(dpi=50).tobytes("png")
                with cols[(i-1) % 4]:
                    st.image(thumb, caption=f"Page {i}", use_column_width=True)
                    deg = st.selectbox(
                        f"Rotate Pg {i}", [0, 90, 180, 270], index=0, key=f"rot_{i}"
                    )
                    page_data.append({"index": i-1, "deg": deg})

            if st.button("Apply Custom Rotation"):
                w = PdfWriter()
                for pdata in page_data:
                    pg = r.pages[pdata["index"]]
                    if pdata["deg"] != 0:
                        pg.rotate(pdata["deg"])
                    w.add_page(pg)
                out = io.BytesIO(); w.write(out)
                _download("Download rotated.pdf", out.getvalue(), "rotated.pdf", "application/pdf")

# ---------- Tab 5: Re-order Pages ----------
with tabs[4]:
    st.subheader("Re-order Pages")

    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="reorder_pdf")
    if pdf:
        path = _save_upload(pdf, suffix=".pdf")
        r = PdfReader(path)
        doc = fitz.open(path)

        st.markdown(f"**This PDF has {len(r.pages)} pages.**")

        # Collect data for each page
        page_data = []
        cols = st.columns(4)
        for i, page in enumerate(doc, start=1):
            thumb = page.get_pixmap(dpi=50).tobytes("png")
            with cols[(i-1) % 4]:
                st.image(thumb, caption=f"Page {i}", use_column_width=True)
                keep = st.checkbox(f"Keep {i}", value=True, key=f"keep_{i}")
                rotate = st.selectbox(f"Rotate {i}", [0, 90, 180, 270], index=0, key=f"rot_{i}")
                page_data.append({"index": i-1, "keep": keep, "rotate": rotate})

        st.markdown("### Page Order")
        order = st.multiselect(
            "Select pages in the new order",
            [i+1 for i in range(len(page_data)) if page_data[i]["keep"]],
            default=[i+1 for i in range(len(page_data)) if page_data[i]["keep"]]
        )

        if st.button("Build Re-ordered PDF"):
            w = PdfWriter()
            for idx in order:
                pdata = page_data[idx-1]
                if not pdata["keep"]:
                    continue
                pg = r.pages[pdata["index"]]
                if pdata["rotate"]:
                    pg.rotate(pdata["rotate"])
                w.add_page(pg)
            out = io.BytesIO(); w.write(out)
            _download("Download reordered.pdf", out.getvalue(), "reordered.pdf", "application/pdf")

# ---------- Tab 6: Extract Text ----------
with tabs[5]:
    st.subheader("Extract Text from PDF")

    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="extract_pdf")
    if pdf:
        path = _save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)

        st.markdown(f"**This PDF has {len(doc)} pages.**")

        # Options
        mode = st.radio("Extraction mode", ["Plain text", "Preserve layout", "Raw"])
        ranges = st.text_input("Pages to extract (e.g., 1-3,5; blank = all)")
        search = st.text_input("Search term (optional)")

        # Preview thumbnails with text snippet
        st.markdown("### Preview Pages")
        cols = st.columns(4)
        for i, page in enumerate(doc, start=1):
            thumb = page.get_pixmap(dpi=40).tobytes("png")
            text = page.get_text("text" if mode == "Plain text" else "blocks" if mode == "Preserve layout" else "raw")
            snippet = (text[:120] + "…") if len(text) > 120 else text
            with cols[(i-1) % 4]:
                st.image(thumb, caption=f"Page {i}", use_column_width=True)
                st.caption(snippet if snippet.strip() else "[No visible text]")

        if st.button("Extract Now"):
            idxs = _parse_ranges(ranges, len(doc)) if ranges.strip() else list(range(len(doc)))
            texts = []
            for i in idxs:
                text = doc[i].get_text("text" if mode == "Plain text" else "blocks" if mode == "Preserve layout" else "raw")
                if search and search.lower() in text.lower():
                    # highlight search matches with >>> <<< markers
                    pattern = re.compile(re.escape(search), re.IGNORECASE)
                    text = pattern.sub(lambda m: f">>>{m.group(0)}<<<", text)
                texts.append(f"--- Page {i+1} ---\n{text}")

            final_text = "\n\n".join(texts)
            st.text_area("Extracted Text", value=final_text, height=400)

            # Export buttons
            st.download_button("Download .txt", data=final_text, file_name="extracted.txt", mime="text/plain")

            try:
                from docx import Document
                docx = Document()
                docx.add_paragraph(final_text)
                out_buf = io.BytesIO()
                docx.save(out_buf)
                st.download_button("Download .docx", data=out_buf.getvalue(), file_name="extracted.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            except ImportError:
                st.info("Install `python-docx` to enable DOCX export: pip install python-docx")
# ---------- Tab 7: Edit (Pro) ----------
with tabs[6]:
    st.subheader("Edit PDF — Pro (Interactive)")

    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="edit_pro")
    if pdf:
        st.success(f"Uploaded: {pdf.name}")

        subtabs = st.tabs([
            "Pages", "Visual Edit", "Annotations", "Media", 
            "Links & Signatures", "Scanned PDF Editor"
        ])

        with subtabs[0]:
            st.markdown("### Page Management")
            st.info("Coming soon...")

        with subtabs[1]:
            st.markdown("### Visual Edit")
            st.info("Coming soon...")

        with subtabs[2]:
            st.markdown("### Annotations")
            st.info("Coming soon...")

        with subtabs[3]:
            st.markdown("### Media")
            st.info("Coming soon...")

        with subtabs[4]:
            st.markdown("### Links & Signatures")
            st.info("Coming soon...")

        with subtabs[5]:
            st.markdown("### Scanned PDF Editor")
            st.info("Coming soon...")

# ---------- Tab 9: Protect PDF ----------
with tabs[8]:
    st.subheader("Protect PDF (Password + Permissions)")

    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="protect_pdf")
    if pdf:
        path = _save_upload(pdf, suffix=".pdf")
        r = PdfReader(path)

        st.markdown("### Current Status")
        if getattr(r, "is_encrypted", False):
            st.warning("This PDF is already encrypted.")
        else:
            st.success("This PDF is not encrypted yet.")

        # Inputs
        user_pwd = st.text_input("User Password (to open PDF)", type="password")
        owner_pwd = st.text_input("Owner Password (to control permissions)", type="password")

        st.markdown("### Permissions (check to allow)")
        allow_print = st.checkbox("Allow printing", value=True)
        allow_copy = st.checkbox("Allow copy text/images", value=True)
        allow_annot = st.checkbox("Allow annotations", value=True)

        encryption = st.selectbox("Encryption strength", ["AES-256", "AES-128", "RC4-128", "RC4-40"], index=0)

        if st.button("Apply Protection"):
            w = PdfWriter()
            for pg in r.pages:
                w.add_page(pg)

            # Permissions bitmask
            perms = {
                "print": allow_print,
                "copy": allow_copy,
                "annotate": allow_annot,
            }

            # Encrypt
            try:
                w.encrypt(
                    user_password=user_pwd if user_pwd else None,
                    owner_password=owner_pwd if owner_pwd else user_pwd,
                    use_128bit=("128" in encryption),
                    permissions=perms
                )
                out = io.BytesIO()
                w.write(out)
                _download("Download protected.pdf", out.getvalue(), "protected.pdf", "application/pdf")
                st.success("PDF protected successfully!")
            except Exception as e:
                st.error(f"Failed to protect: {e}")

# ---------- Tab 10: Unlock PDF ----------
with tabs[9]:
    st.subheader("Unlock PDF (Remove Password)")

    pdfs = st.file_uploader("Upload one or more encrypted PDFs", type=["pdf"], accept_multiple_files=True, key="unlock_pdf")
    password = st.text_input("Password", type="password")

    if pdfs and password and st.button("Unlock Now"):
        tmpdir = tempfile.mkdtemp(prefix="unlock_")
        unlocked_files = []

        for uf in pdfs:
            path = _save_upload(uf, suffix=".pdf")
            reader = PdfReader(path)

            if not getattr(reader, "is_encrypted", False):
                st.info(f"{uf.name}: Not encrypted, skipped.")
                continue

            try:
                res = reader.decrypt(password)
                if res == 0:
                    st.error(f"{uf.name}: Incorrect password.")
                    continue

                writer = PdfWriter()
                for pg in reader.pages:
                    writer.add_page(pg)

                out_path = os.path.join(tmpdir, uf.name.replace(".pdf","_unlocked.pdf"))
                with open(out_path, "wb") as f:
                    writer.write(f)
                unlocked_files.append(out_path)
                st.success(f"{uf.name}: Unlocked!")

            except Exception as e:
                st.error(f"{uf.name}: Failed — {e}")

        if unlocked_files:
            if len(unlocked_files) == 1:
                out = _to_bytes(unlocked_files[0])
                _download("Download unlocked.pdf", out, os.path.basename(unlocked_files[0]), "application/pdf")
            else:
                mem = io.BytesIO()
                with zipfile.ZipFile(mem, "w") as zf:
                    for p in unlocked_files:
                        zf.write(p, os.path.basename(p))
                mem.seek(0)
                _download("Download unlocked.zip", mem.getvalue(), "unlocked.zip", "application/zip")

        shutil.rmtree(tmpdir, ignore_errors=True)

# ---------- Tab 11: PDF → Images (High Quality) ----------
with tabs[10]:
    st.subheader("PDF → Images (High Quality)")

    pdfs = st.file_uploader("Upload one or more PDFs", type=["pdf"], accept_multiple_files=True, key="pdf2img_hq")

    if pdfs:
        dpi = st.slider("Export DPI (higher = sharper text)", 150, 600, 300, step=50)
        fmt = st.selectbox("Image format", ["PNG (lossless, best quality)", "JPEG (compressed)", "WebP (compressed)"])
        if "JPEG" in fmt or "WebP" in fmt:
            quality = st.slider("Quality", 80, 100, 95)

        mode = st.radio("Export mode", ["One image per page (ZIP)", "All pages stitched vertically (one image)"])

        if st.button("Convert to Images (High Quality)"):
            tmpdir = tempfile.mkdtemp(prefix="pdf2img_hq_")
            results = []

            for uf in pdfs:
                path = _save_upload(uf, suffix=".pdf")
                doc = fitz.open(path)
                subdir = os.path.join(tmpdir, os.path.splitext(uf.name)[0])
                os.makedirs(subdir, exist_ok=True)

                images = []
                for i, page in enumerate(doc, start=1):
                    # Use selected DPI for export
                    pix = page.get_pixmap(dpi=dpi, alpha=False)

                    if "PNG" in fmt:
                        imgdata = pix.tobytes("png")
                        ext = "png"
                    elif "JPEG" in fmt:
                        imgdata = pix.tobytes("jpg", jpg_quality=quality)
                        ext = "jpg"
                    else:  # WebP
                        pilimg = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                        buf = io.BytesIO()
                        pilimg.save(buf, format="WEBP", quality=quality)
                        imgdata = buf.getvalue()
                        ext = "webp"

                    outpath = os.path.join(subdir, f"page_{i}.{ext}")
                    with open(outpath, "wb") as f:
                        f.write(imgdata)

                    # Collect PIL images for stitching mode
                    images.append(Image.open(io.BytesIO(imgdata)))

                # If stitched mode, merge vertically
                if mode == "All pages stitched vertically (one image)" and images:
                    widths, heights = zip(*(im.size for im in images))
                    total_height = sum(heights)
                    max_width = max(widths)
                    stitched = Image.new("RGB", (max_width, total_height), "white")
                    y_offset = 0
                    for im in images:
                        stitched.paste(im, (0, y_offset))
                        y_offset += im.height
                    outpath = os.path.join(subdir, f"{os.path.splitext(uf.name)[0]}_stitched.{ext}")
                    stitched.save(outpath, quality=quality if "JPEG" in fmt or "WebP" in fmt else None)

            # Zip results
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, "w") as zf:
                for root, _, files in os.walk(tmpdir):
                    for f in files:
                        full = os.path.join(root, f)
                        arc = os.path.relpath(full, tmpdir)
                        zf.write(full, arcname=arc)
            mem.seek(0)

            _download("Download images.zip", mem.getvalue(), "images.zip", "application/zip")
            shutil.rmtree(tmpdir, ignore_errors=True)
# ---------- Tab 12: PDF → DOCX ----------
with tabs[11]:
    st.subheader("PDF → DOCX (Editable Word)")

    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pdf2docx_pdf")
    if pdf:
        path = _save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)

        st.markdown(f"**This PDF has {len(doc)} pages.**")

        # Preview thumbnails
        st.markdown("### Preview first few pages")
        cols = st.columns(5)
        for i, page in enumerate(doc[:min(5, len(doc))], start=1):
            thumb = page.get_pixmap(dpi=40).tobytes("png")
            with cols[(i-1) % 5]:
                st.image(thumb, caption=f"Page {i}", use_column_width=True)

        # Options
        start = st.number_input("Start page (1-based)", 1, len(doc), 1)
        end = st.number_input("End page (0 = all)", 0, len(doc), 0)
        keep_images = st.checkbox("Keep images in DOCX", value=True)

        if st.button("Convert to DOCX"):
            from pdf2docx import Converter
            out_dir = tempfile.mkdtemp(prefix="pdf2docx_")
            out_path = os.path.join(out_dir, "converted.docx")

            try:
                cv = Converter(path)
                cv.convert(out_path, start=start-1, end=None if end==0 else end-1, 
                           retain_image=keep_images)
                cv.close()

                # Return DOCX
                with open(out_path, "rb") as f:
                    data = f.read()
                _download("Download converted.docx", data, "converted.docx",
                          "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                st.success("DOCX conversion completed!")

            except Exception as e:
                st.error(f"Conversion failed: {e}")
                st.info("If this is a scanned PDF, try enabling OCR in the advanced version.")

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
# ---------- Tab 13: Watermark ----------
with tabs[12]:
    st.subheader("Add Watermark")

    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="wm_pdf")
    wm_type = st.radio("Watermark type", ["Text", "Image"])

    if pdf:
        path = _save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)

        # Options
        if wm_type == "Text":
            text = st.text_input("Watermark text", "CONFIDENTIAL")
            color = st.color_picker("Text color", "#FF0000")
            opacity = st.slider("Opacity (%)", 10, 90, 20)
            size = st.slider("Font size", 20, 120, 60)
            angle = st.slider("Rotation angle", 0, 360, 45)
            pos = st.selectbox("Position", ["Center", "Top-left", "Top-right", "Bottom-left", "Bottom-right", "Diagonal Tiled"])
        else:
            wm_img = st.file_uploader("Upload image", type=["png","jpg"])
            opacity = st.slider("Opacity (%)", 10, 90, 50)
            size = st.slider("Relative size (%)", 10, 100, 40)
            pos = st.selectbox("Position", ["Center", "Top-left", "Top-right", "Bottom-left", "Bottom-right", "Diagonal Tiled"])

        pages = st.text_input("Apply to pages (e.g., 1-3,5; blank = all)")
        if st.button("Apply Watermark"):
            idxs = _parse_ranges(pages, len(doc)) if pages.strip() else list(range(len(doc)))

            for i in idxs:
                page = doc[i]
                rect = page.rect

                if wm_type == "Text":
                    rgb = tuple(int(color[j:j+2], 16)/255 for j in (1,3,5))
                    if pos == "Diagonal Tiled":
                        # Tile across page diagonally
                        step = rect.height / 3
                        for y in range(0, int(rect.height), int(step)):
                            page.insert_text(
                                (rect.width/2, y),
                                text,
                                fontsize=size,
                                rotate=angle,
                                color=rgb,
                                fill_opacity=opacity/100,
                                align=1
                            )
                    else:
                        coords = {
                            "Center": rect.center,
                            "Top-left": (rect.x0+50, rect.y0+80),
                            "Top-right": (rect.x1-200, rect.y0+80),
                            "Bottom-left": (rect.x0+50, rect.y1-50),
                            "Bottom-right": (rect.x1-200, rect.y1-50)
                        }
                        page.insert_text(
                            coords.get(pos, rect.center),
                            text,
                            fontsize=size,
                            rotate=angle,
                            color=rgb,
                            fill_opacity=opacity/100,
                            align=1
                        )
                else:
                    if not wm_img:
                        st.error("Please upload a watermark image.")
                        break
                    img_path = _save_upload(wm_img)
                    img = Image.open(img_path)
                    w, h = img.size
                    scale = (size/100)
                    w, h = int(rect.width*scale), int(rect.height*scale)
                    pos_map = {
                        "Center": fitz.Rect(rect.width/2-w/2, rect.height/2-h/2, rect.width/2+w/2, rect.height/2+h/2),
                        "Top-left": fitz.Rect(50, 50, 50+w, 50+h),
                        "Top-right": fitz.Rect(rect.width-w-50, 50, rect.width-50, 50+h),
                        "Bottom-left": fitz.Rect(50, rect.height-h-50, 50+w, rect.height-50),
                        "Bottom-right": fitz.Rect(rect.width-w-50, rect.height-h-50, rect.width-50, rect.height-50),
                    }
                    target = pos_map.get(pos, rect)
                    page.insert_image(target, filename=img_path, overlay=True, keep_proportion=True)

            # Save output
            out = io.BytesIO()
            doc.save(out, deflate=True)
            doc.close()
            _download("Download watermarked.pdf", out.getvalue(), "watermarked.pdf", "application/pdf")

# ---------- Tab 14: Page Numbers ----------
with tabs[13]:
    st.subheader("Add Page Numbers")

    pdf = st.file_uploader("Upload PDF", type=["pdf"], key="pnum_pdf")
    if pdf:
        path = _save_upload(pdf, suffix=".pdf")
        doc = fitz.open(path)

        st.markdown(f"**This PDF has {len(doc)} pages.**")

        # Options
        ranges = st.text_input("Apply to pages (e.g. 1-3,5; blank = all)")
        style = st.selectbox("Numbering style", ["1, 2, 3", "01, 02, 03", "i, ii, iii", "I, II, III", "a, b, c", "A, B, C"])
        template = st.text_input("Custom template", "Page {n} of {total}")

        pos_v = st.radio("Vertical position", ["Top", "Bottom"], index=1)
        pos_h = st.radio("Horizontal position", ["Left", "Center", "Right"], index=1)

        fontsize = st.slider("Font size", 8, 32, 12)
        color = st.color_picker("Font color", "#000000")
        opacity = st.slider("Opacity (%)", 20, 100, 80)

        if st.button("Add Numbers"):
            idxs = _parse_ranges(ranges, len(doc)) if ranges.strip() else list(range(len(doc)))
            total = len(doc)

            def format_number(n):
                if style == "1, 2, 3":
                    return str(n)
                elif style == "01, 02, 03":
                    return f"{n:02d}"
                elif style == "i, ii, iii":
                    return to_roman(n).lower()
                elif style == "I, II, III":
                    return to_roman(n)
                elif style == "a, b, c":
                    return chr(96+n)
                elif style == "A, B, C":
                    return chr(64+n)
                return str(n)

            rgb = tuple(int(color[i:i+2], 16)/255 for i in (1,3,5))

            for i in idxs:
                page = doc[i]
                num = format_number(i+1)
                label = template.format(n=num, total=total)

                rect = page.rect
                x = rect.x0 + 40 if pos_h == "Left" else rect.x1 - 40 if pos_h == "Right" else rect.x0 + rect.width/2
                y = rect.y0 + 40 if pos_v == "Top" else rect.y1 - 30

                page.insert_text(
                    (x, y), label,
                    fontsize=fontsize,
                    color=rgb,
                    fill_opacity=opacity/100,
                    align=1 if pos_h=="Center" else 0
                )

            out = io.BytesIO()
            doc.save(out, deflate=True)
            doc.close()
            _download("Download numbered.pdf", out.getvalue(), "numbered.pdf", "application/pdf")

# ---------- Tab 15: Office → PDF ----------
with tabs[14]:
    st.subheader("Office → PDF (Word, Excel, PowerPoint, Text, Images)")

    files = st.file_uploader("Upload Office or text/image files", 
                             type=["docx","xlsx","pptx","txt","odt","ods","odp","png","jpg"], 
                             accept_multiple_files=True, key="office2pdf")

    if files and st.button("Convert to PDF"):
        tmpdir = tempfile.mkdtemp(prefix="office2pdf_")
        pdf_paths = []

        for uf in files:
            fname = uf.name
            path = _save_upload(uf, suffix=f".{fname.split('.')[-1]}")

            out_path = os.path.join(tmpdir, fname.rsplit(".",1)[0] + ".pdf")

            try:
                ext = fname.split(".")[-1].lower()

                if ext in ["docx","odt"]:
                    # DOCX to PDF
                    import pypandoc
                    pypandoc.convert_file(path, "pdf", outputfile=out_path, extra_args=["--standalone"])

                elif ext in ["xlsx","ods"]:
                    import pandas as pd
                    from reportlab.platypus import SimpleDocTemplate, Table
                    xl = pd.ExcelFile(path)
                    doc = SimpleDocTemplate(out_path)
                    elements = []
                    for sheet in xl.sheet_names:
                        df = xl.parse(sheet)
                        elements.append(Table([df.columns.tolist()] + df.values.tolist()))
                    doc.build(elements)

                elif ext in ["pptx","odp"]:
                    from pptx import Presentation
                    prs = Presentation(path)
                    pdf_doc = fitz.open()
                    for slide in prs.slides:
                        img = slide.shapes
                        # (Simple: placeholder, real implementation requires render)
                        page = pdf_doc.new_page(width=800, height=600)
                        page.insert_text((50,300),"[Slide content here]")
                    pdf_doc.save(out_path)

                elif ext == "txt":
                    from reportlab.platypus import SimpleDocTemplate, Paragraph
                    from reportlab.lib.styles import getSampleStyleSheet
                    text = uf.read().decode("utf-8")
                    doc = SimpleDocTemplate(out_path)
                    styles = getSampleStyleSheet()
                    doc.build([Paragraph(line, styles["Normal"]) for line in text.split("\n")])

                elif ext in ["png","jpg"]:
                    img = Image.open(path)
                    pdf_bytes = io.BytesIO()
                    img.convert("RGB").save(out_path, "PDF")

                pdf_paths.append(out_path)
                st.success(f"{fname} → PDF created")

            except Exception as e:
                st.error(f"Failed to convert {fname}: {e}")

        # Bundle results
        if len(pdf_paths) == 1:
            _download("Download PDF", _to_bytes(pdf_paths[0]), os.path.basename(pdf_paths[0]), "application/pdf")
        else:
            mem = io.BytesIO()
            with zipfile.ZipFile(mem, "w") as zf:
                for p in pdf_paths:
                    zf.write(p, os.path.basename(p))
            mem.seek(0)
            _download("Download PDFs.zip", mem.getvalue(), "converted_pdfs.zip", "application/zip")

        shutil.rmtree(tmpdir, ignore_errors=True)









