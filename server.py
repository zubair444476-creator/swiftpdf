import os
import io
import zipfile
import uuid
import base64

from flask import Flask, request, send_file, jsonify, render_template
from werkzeug.utils import secure_filename

import fitz  # pymupdf
from PIL import Image

from docx import Document
from docx.shared import Pt

from pptx import Presentation
from pptx.util import Emu

from openpyxl import Workbook, load_workbook

from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.pdfgen import canvas as pdf_canvas

app = Flask(__name__, static_folder="static", static_url_path="")

MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50 MB
app.config["MAX_CONTENT_LENGTH"] = MAX_CONTENT_LENGTH


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def get_uploaded_files(field_name="files"):
    files = request.files.getlist(field_name)
    if not files:
        single = request.files.get("file")
        if single:
            files = [single]
    return [f for f in files if f and f.filename]


def send_bytes(data: bytes, filename: str, mimetype: str):
    return send_file(
        io.BytesIO(data),
        mimetype=mimetype,
        as_attachment=True,
        download_name=filename,
    )


def parse_page_spec(spec: str, page_count: int):
    """Parse a 1-indexed page spec like '1,3,5-7' into a 0-indexed list,
    preserving the order given (so it can also be used to reorder pages)."""
    result = []
    spec = (spec or "").strip()
    if not spec:
        return list(range(page_count))

    for chunk in spec.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            start_s, end_s = chunk.split("-", 1)
            start, end = int(start_s), int(end_s)
            step = 1 if end >= start else -1
            for p in range(start, end + step, step):
                if 1 <= p <= page_count:
                    result.append(p - 1)
        else:
            p = int(chunk)
            if 1 <= p <= page_count:
                result.append(p - 1)
    return result


# ---------------------------------------------------------------------------
# Static frontend
# ---------------------------------------------------------------------------

@app.route("/")
def index():
    return app.send_static_file("index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


# ---------------------------------------------------------------------------
# Page thumbnails (used by the visual page picker in the frontend)
# ---------------------------------------------------------------------------

@app.route("/api/page-thumbnails", methods=["POST"])
def page_thumbnails():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400

    src = files[0]
    data = src.read()

    zoom = 0.4          # ~58 DPI — small enough to be fast, big enough to read
    mat = fitz.Matrix(zoom, zoom)
    thumbs = []

    with fitz.open(stream=data, filetype="pdf") as doc:
        page_count = len(doc)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=mat)
            png_bytes = pix.tobytes("png")
            b64 = base64.b64encode(png_bytes).decode("ascii")
            thumbs.append({
                "page": i + 1,           # 1-indexed
                "width": pix.width,
                "height": pix.height,
                "dataUrl": f"data:image/png;base64,{b64}",
            })

    return jsonify({"pageCount": page_count, "thumbnails": thumbs})


# ---------------------------------------------------------------------------
# Merge PDFs
# ---------------------------------------------------------------------------

@app.route("/api/merge", methods=["POST"])
def merge_pdfs():
    files = get_uploaded_files()
    if len(files) < 2:
        return jsonify({"error": "Upload at least 2 PDF files to merge."}), 400

    merged = fitz.open()
    try:
        for f in files:
            data = f.read()
            with fitz.open(stream=data, filetype="pdf") as doc:
                merged.insert_pdf(doc)

        out_bytes = merged.tobytes()
    finally:
        merged.close()

    return send_bytes(out_bytes, "merged.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Split a PDF into individual pages (returned as a zip)
# ---------------------------------------------------------------------------

@app.route("/api/split", methods=["POST"])
def split_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to split."}), 400

    src = files[0]
    data = src.read()
    base_name = os.path.splitext(secure_filename(src.filename))[0] or "page"

    zip_buffer = io.BytesIO()
    with fitz.open(stream=data, filetype="pdf") as doc:
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i in range(len(doc)):
                single = fitz.open()
                single.insert_pdf(doc, from_page=i, to_page=i)
                page_bytes = single.tobytes()
                single.close()
                zf.writestr(f"{base_name}_page_{i + 1}.pdf", page_bytes)

    zip_buffer.seek(0)
    return send_bytes(zip_buffer.getvalue(), f"{base_name}_split.zip", "application/zip")


# ---------------------------------------------------------------------------
# Rotate a PDF (all pages) by 90/180/270 degrees
# ---------------------------------------------------------------------------

@app.route("/api/rotate", methods=["POST"])
def rotate_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to rotate."}), 400

    try:
        angle = int(request.form.get("angle", 90))
    except ValueError:
        angle = 90
    angle = angle % 360

    src = files[0]
    data = src.read()

    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            page.set_rotation((page.rotation + angle) % 360)
        out_bytes = doc.tobytes()

    return send_bytes(out_bytes, "rotated.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Compress a PDF (re-save with garbage collection + deflate, downsample images)
# ---------------------------------------------------------------------------

@app.route("/api/compress", methods=["POST"])
def compress_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to compress."}), 400

    quality = request.form.get("quality", "recommended")
    jpeg_quality = {"low": 40, "recommended": 60, "high": 80}.get(quality, 60)
    max_dim = {"low": 800, "recommended": 1200, "high": 1600}.get(quality, 1200)

    src = files[0]
    data = src.read()

    doc = fitz.open(stream=data, filetype="pdf")
    try:
        for page in doc:
            for img in page.get_images(full=True):
                xref = img[0]
                try:
                    base = doc.extract_image(xref)
                    img_bytes = base["image"]
                    pil_img = Image.open(io.BytesIO(img_bytes)).convert("RGB")

                    w, h = pil_img.size
                    scale = min(1.0, max_dim / max(w, h))
                    if scale < 1.0:
                        pil_img = pil_img.resize(
                            (max(1, int(w * scale)), max(1, int(h * scale))),
                            Image.LANCZOS,
                        )

                    buf = io.BytesIO()
                    pil_img.save(buf, format="JPEG", quality=jpeg_quality, optimize=True)
                    doc.update_stream(xref, buf.getvalue())
                except Exception:
                    # Not all images can be re-encoded (masks, CMYK, etc.) - skip those.
                    continue

        out_bytes = doc.tobytes(deflate=True, garbage=4)
    finally:
        doc.close()

    return send_bytes(out_bytes, "compressed.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Images -> PDF
# ---------------------------------------------------------------------------

@app.route("/api/images-to-pdf", methods=["POST"])
def images_to_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload one or more images."}), 400

    pil_images = []
    for f in files:
        img = Image.open(f.stream).convert("RGB")
        pil_images.append(img)

    buf = io.BytesIO()
    first, rest = pil_images[0], pil_images[1:]
    first.save(buf, format="PDF", save_all=True, append_images=rest)
    buf.seek(0)

    return send_bytes(buf.getvalue(), "images.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# PDF -> Images (returned as a zip of PNGs)
# ---------------------------------------------------------------------------

@app.route("/api/pdf-to-images", methods=["POST"])
def pdf_to_images():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to convert."}), 400

    src = files[0]
    data = src.read()
    base_name = os.path.splitext(secure_filename(src.filename))[0] or "page"

    zoom = 2.0  # ~144 DPI
    mat = fitz.Matrix(zoom, zoom)

    zip_buffer = io.BytesIO()
    with fitz.open(stream=data, filetype="pdf") as doc:
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, page in enumerate(doc):
                pix = page.get_pixmap(matrix=mat)
                zf.writestr(f"{base_name}_page_{i + 1}.png", pix.tobytes("png"))

    zip_buffer.seek(0)
    return send_bytes(zip_buffer.getvalue(), f"{base_name}_images.zip", "application/zip")


# ---------------------------------------------------------------------------
# PDF -> plain text (simple extraction)
# ---------------------------------------------------------------------------

@app.route("/api/pdf-to-text", methods=["POST"])
def pdf_to_text():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to convert."}), 400

    src = files[0]
    data = src.read()
    base_name = os.path.splitext(secure_filename(src.filename))[0] or "document"

    text_parts = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            text_parts.append(page.get_text())

    full_text = "\n\n".join(text_parts).encode("utf-8")
    return send_bytes(full_text, f"{base_name}.txt", "text/plain")


# ---------------------------------------------------------------------------
# PDF -> Word (.docx)  — rich conversion: headings, bold, color, tables
# ---------------------------------------------------------------------------

def _pdf_int_to_rgb(color_int):
    from docx.shared import RGBColor
    return RGBColor((color_int >> 16) & 0xFF, (color_int >> 8) & 0xFF, color_int & 0xFF)


def _set_cell_bg(cell, hex_fill):
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def _render_rich_spans(paragraph, spans):
    for span in spans:
        text = span["text"]
        if not text:
            continue
        run = paragraph.add_run(text)
        run.bold   = bool(span["flags"] & 16)
        run.italic = bool(span["flags"] & 2)
        run.font.size = Pt(max(6, span["size"]))
        if span["color"]:
            run.font.color.rgb = _pdf_int_to_rgb(span["color"])


def _convert_page_to_docx(page, document):
    from docx.shared import Pt, RGBColor, Cm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    # ── Detect tables ────────────────────────────────────────────────────────
    table_rects = []
    page_tables = []   # list of (fitz.Rect, [[cell,...], ...], ncols)
    try:
        found = page.find_tables()
        for ft in found.tables:
            raw = ft.extract()
            if len(raw) < 2:
                continue
            ncols = max(len(r) for r in raw)
            if ncols < 2:
                continue
            flat = [c for row in raw for c in row if c and str(c).strip()]
            if len(flat) < 4:
                continue
            rect = fitz.Rect(ft.bbox)
            table_rects.append(rect)
            padded = [(list(r) + [""] * ncols)[:ncols] for r in raw]
            page_tables.append((rect, padded, ncols))
    except Exception:
        pass

    def in_table(bbox):
        r = fitz.Rect(bbox)
        return any(r.intersects(tr) for tr in table_rects)

    # ── Collect text blocks outside tables ───────────────────────────────────
    raw_blocks = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)["blocks"]
    text_blocks = [b for b in raw_blocks if b.get("type") == 0 and not in_table(b["bbox"])]
    text_blocks.sort(key=lambda b: (round(b["bbox"][1] / 4) * 4, b["bbox"][0]))

    # ── Interleave text + tables in reading order ────────────────────────────
    items = [(b["bbox"][1], "text", b) for b in text_blocks]
    for (rect, rows, ncols) in page_tables:
        items.append((rect.y0, "table", (rows, ncols)))
    items.sort(key=lambda x: x[0])

    for (_, kind, data) in items:

        if kind == "text":
            block = data
            for line in block["lines"]:
                spans = [s for s in line["spans"] if s["text"].strip()]
                if not spans:
                    continue
                full_text = "".join(s["text"] for s in spans).strip()
                if not full_text:
                    continue
                first = spans[0]
                size  = first["size"]
                color = first["color"]

                if size >= 26:
                    p = document.add_paragraph(style="Heading 1")
                    run = p.add_run(full_text)
                    run.bold = True
                    run.font.size = Pt(22)
                    if color:
                        run.font.color.rgb = _pdf_int_to_rgb(color)
                elif size >= 13:
                    p = document.add_paragraph(style="Heading 2")
                    run = p.add_run(full_text)
                    run.bold = True
                    run.font.size = Pt(13)
                    if color:
                        run.font.color.rgb = _pdf_int_to_rgb(color)
                else:
                    p = document.add_paragraph()
                    p.paragraph_format.space_after  = Pt(1)
                    p.paragraph_format.space_before = Pt(0)
                    _render_rich_spans(p, spans)

        elif kind == "table":
            rows, ncols = data

            # PyMuPDF often splits "1.0 Rest shelter" across col0 ("1.0 Re") and
            # col1 ("st shelter available"). Detect this: if EVERY col0 value is
            # short (≤ 8 chars) AND col1 carries the rest, merge them.
            col0_vals = [(r[0] or "").strip() for r in rows]
            col1_vals = [(r[1] or "").strip() for r in rows]
            # Check: col0 short AND (col0+col1 concatenated looks like one word)
            split_detected = (
                ncols >= 3
                and all(len(v) <= 8 for v in col0_vals)
                and any(
                    v and col1_vals[i] and not v[-1].isspace()
                    and not col1_vals[i][0].isupper()
                    for i, v in enumerate(col0_vals)
                )
            )

            # Fallback: if col0 max len is ≤ 6 (just a row number fragment), merge
            if not split_detected and ncols >= 3:
                split_detected = all(len(v) <= 6 for v in col0_vals)

            if split_detected:
                dcols = ncols - 1
                t = document.add_table(rows=len(rows), cols=dcols)
                t.style = "Table Grid"
                for ri, row in enumerate(rows):
                    c0 = (row[0] or "").rstrip()
                    c1 = (row[1] or "").lstrip()
                    # Join without space when split is mid-word (col0 doesn't end with space/digit)
                    if c0 and c1 and not c0[-1].isspace() and c1 and not c1[0].isupper() and not c1[0].isdigit():
                        merged = c0 + c1
                    else:
                        merged = (c0 + " " + c1).strip()
                    display = [merged] + [row[c] or "" for c in range(2, ncols)]
                    for ci, txt in enumerate(display[:dcols]):
                        cell = t.cell(ri, ci)
                        cell.text = ""
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_after = Pt(1)
                        run = p.add_run(txt)
                        if ri == 0:
                            run.bold = True
                            run.font.size = Pt(9)
                            _set_cell_bg(cell, "1859A9")
                            run.font.color.rgb = RGBColor(255, 255, 255)
                        else:
                            run.font.size = Pt(9)
                            if ri % 2 == 0:
                                _set_cell_bg(cell, "EEF3FA")
            else:
                t = document.add_table(rows=len(rows), cols=ncols)
                t.style = "Table Grid"
                for ri, row in enumerate(rows):
                    for ci in range(ncols):
                        cell = t.cell(ri, ci)
                        cell.text = ""
                        p = cell.paragraphs[0]
                        p.paragraph_format.space_after = Pt(1)
                        run = p.add_run(row[ci] or "")
                        if ri == 0:
                            run.bold = True
                            run.font.size = Pt(9)
                            _set_cell_bg(cell, "1859A9")
                            run.font.color.rgb = RGBColor(255, 255, 255)
                        else:
                            run.font.size = Pt(9)
                            if ri % 2 == 0:
                                _set_cell_bg(cell, "EEF3FA")

            document.add_paragraph()  # breathing room after table


@app.route("/api/pdf-to-word", methods=["POST"])
def pdf_to_word():
    from docx.shared import Cm

    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to convert."}), 400

    src = files[0]
    data = src.read()
    base_name = os.path.splitext(secure_filename(src.filename))[0] or "document"

    document = Document()

    # Tighter margins to match typical PDF layout
    for section in document.sections:
        section.left_margin   = Cm(1.8)
        section.right_margin  = Cm(1.8)
        section.top_margin    = Cm(1.5)
        section.bottom_margin = Cm(1.5)

    with fitz.open(stream=data, filetype="pdf") as doc:
        for page_index, page in enumerate(doc):
            _convert_page_to_docx(page, document)
            if page_index < len(doc) - 1:
                document.add_page_break()

    buf = io.BytesIO()
    document.save(buf)
    buf.seek(0)

    return send_bytes(
        buf.getvalue(),
        f"{base_name}.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


# ---------------------------------------------------------------------------
# PDF -> PowerPoint (.pptx) - each page becomes a full-slide image
# ---------------------------------------------------------------------------

@app.route("/api/pdf-to-pptx", methods=["POST"])
def pdf_to_pptx():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to convert."}), 400

    src = files[0]
    data = src.read()
    base_name = os.path.splitext(secure_filename(src.filename))[0] or "presentation"

    zoom = 2.0
    mat = fitz.Matrix(zoom, zoom)

    prs = Presentation()

    with fitz.open(stream=data, filetype="pdf") as doc:
        if len(doc) == 0:
            return jsonify({"error": "The PDF has no pages."}), 400

        # Size the slides to match the first page's aspect ratio
        first_rect = doc[0].rect
        prs.slide_width = Emu(int(first_rect.width * 12700))
        prs.slide_height = Emu(int(first_rect.height * 12700))
        blank_layout = prs.slide_layouts[6]

        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            img_bytes = pix.tobytes("png")

            slide = prs.slides.add_slide(blank_layout)
            slide.shapes.add_picture(
                io.BytesIO(img_bytes),
                Emu(0),
                Emu(0),
                width=prs.slide_width,
                height=prs.slide_height,
            )

    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)

    return send_bytes(
        buf.getvalue(),
        f"{base_name}.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


# ---------------------------------------------------------------------------
# PDF -> Excel (.xlsx) - extracts detected tables, falls back to raw text lines
# ---------------------------------------------------------------------------

@app.route("/api/pdf-to-excel", methods=["POST"])
def pdf_to_excel():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to convert."}), 400

    src = files[0]
    data = src.read()
    base_name = os.path.splitext(secure_filename(src.filename))[0] or "spreadsheet"

    wb = Workbook()
    wb.remove(wb.active)  # start with no sheets, add one per page below

    with fitz.open(stream=data, filetype="pdf") as doc:
        for page_index, page in enumerate(doc):
            sheet_name = f"Page {page_index + 1}"[:31]
            ws = wb.create_sheet(title=sheet_name)

            wrote_table = False
            try:
                tables = page.find_tables()
                if tables and tables.tables:
                    row_cursor = 1
                    for table in tables.tables:
                        extracted = table.extract()
                        for row in extracted:
                            for col_index, cell_value in enumerate(row, start=1):
                                ws.cell(
                                    row=row_cursor,
                                    column=col_index,
                                    value=cell_value if cell_value is not None else "",
                                )
                            row_cursor += 1
                        row_cursor += 1  # blank row between tables
                        wrote_table = True
            except Exception:
                wrote_table = False

            if not wrote_table:
                # Fall back to dumping each line of text into column A
                text = page.get_text()
                for row_index, line in enumerate(text.split("\n"), start=1):
                    if line.strip():
                        ws.cell(row=row_index, column=1, value=line)

    if not wb.sheetnames:
        wb.create_sheet(title="Sheet1")

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_bytes(
        buf.getvalue(),
        f"{base_name}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ---------------------------------------------------------------------------
# Remove pages
# ---------------------------------------------------------------------------

@app.route("/api/remove-pages", methods=["POST"])
def remove_pages():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    pages_spec = request.form.get("pages", "")
    if not pages_spec.strip():
        return jsonify({"error": "Tell me which pages to remove, e.g. 2,4-5"}), 400

    src = files[0]
    data = src.read()

    with fitz.open(stream=data, filetype="pdf") as doc:
        to_remove = sorted(set(parse_page_spec(pages_spec, len(doc))), reverse=True)
        if not to_remove:
            return jsonify({"error": "No valid page numbers found in that range."}), 400
        for idx in to_remove:
            doc.delete_page(idx)
        if len(doc) == 0:
            return jsonify({"error": "That would remove every page. Leave at least one."}), 400
        out_bytes = doc.tobytes()

    return send_bytes(out_bytes, "pages_removed.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Extract / reorder pages ("Organize PDF")
# ---------------------------------------------------------------------------

@app.route("/api/extract-pages", methods=["POST"])
def extract_pages():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    pages_spec = request.form.get("pages", "")
    if not pages_spec.strip():
        return jsonify({"error": "Tell me which pages to keep and in what order, e.g. 3,1,2"}), 400

    src = files[0]
    data = src.read()

    with fitz.open(stream=data, filetype="pdf") as doc:
        page_order = parse_page_spec(pages_spec, len(doc))
        if not page_order:
            return jsonify({"error": "No valid page numbers found in that range."}), 400

        new_doc = fitz.open()
        for idx in page_order:
            new_doc.insert_pdf(doc, from_page=idx, to_page=idx)
        out_bytes = new_doc.tobytes()
        new_doc.close()

    return send_bytes(out_bytes, "organized.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Add page numbers
# ---------------------------------------------------------------------------

@app.route("/api/add-page-numbers", methods=["POST"])
def add_page_numbers():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400

    try:
        start = int(request.form.get("start", 1))
    except ValueError:
        start = 1
    position = request.form.get("position", "bottom-center")

    src = files[0]
    data = src.read()

    with fitz.open(stream=data, filetype="pdf") as doc:
        for i, page in enumerate(doc):
            label = str(start + i)
            rect = page.rect
            margin = 24
            if position == "bottom-left":
                point = fitz.Point(margin, rect.height - margin)
            elif position == "bottom-right":
                point = fitz.Point(rect.width - margin - 20, rect.height - margin)
            elif position == "top-center":
                point = fitz.Point(rect.width / 2 - 8, margin)
            else:  # bottom-center
                point = fitz.Point(rect.width / 2 - 8, rect.height - margin)

            page.insert_text(point, label, fontsize=11, color=(0, 0, 0))

        out_bytes = doc.tobytes()

    return send_bytes(out_bytes, "numbered.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Watermark PDF (diagonal repeated text)
# ---------------------------------------------------------------------------

@app.route("/api/watermark", methods=["POST"])
def watermark_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400

    text = request.form.get("text", "CONFIDENTIAL").strip() or "CONFIDENTIAL"
    try:
        opacity = float(request.form.get("opacity", 0.3))
    except ValueError:
        opacity = 0.3
    opacity = max(0.05, min(opacity, 1.0))

    src = files[0]
    data = src.read()

    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            rect = page.rect
            page.insert_text(
                fitz.Point(rect.width * 0.15, rect.height * 0.55),
                text,
                fontsize=max(24, int(rect.width / 12)),
                rotate=45,
                color=(0.6, 0.6, 0.6),
                fill_opacity=opacity,
                overlay=True,
            )
        out_bytes = doc.tobytes()

    return send_bytes(out_bytes, "watermarked.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Protect PDF (add a password)
# ---------------------------------------------------------------------------

@app.route("/api/protect", methods=["POST"])
def protect_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    password = request.form.get("password", "").strip()
    if not password:
        return jsonify({"error": "Enter a password to protect the PDF with."}), 400

    src = files[0]
    data = src.read()

    with fitz.open(stream=data, filetype="pdf") as doc:
        out_bytes = doc.tobytes(
            encryption=fitz.PDF_ENCRYPT_AES_256,
            owner_pw=password,
            user_pw=password,
            permissions=int(
                fitz.PDF_PERM_PRINT
                | fitz.PDF_PERM_COPY
                | fitz.PDF_PERM_ANNOTATE
            ),
        )

    return send_bytes(out_bytes, "protected.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Unlock PDF (remove a known password)
# ---------------------------------------------------------------------------

@app.route("/api/unlock", methods=["POST"])
def unlock_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    password = request.form.get("password", "").strip()

    src = files[0]
    data = src.read()

    doc = fitz.open(stream=data, filetype="pdf")
    try:
        if doc.needs_pass:
            if not password:
                return jsonify({"error": "This PDF is password-protected. Enter the password."}), 400
            if not doc.authenticate(password):
                return jsonify({"error": "That password didn't work."}), 400
        out_bytes = doc.tobytes()
    finally:
        doc.close()

    return send_bytes(out_bytes, "unlocked.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Crop PDF (trim a margin off every page)
# ---------------------------------------------------------------------------

@app.route("/api/crop", methods=["POST"])
def crop_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400

    try:
        margin = float(request.form.get("margin", 36))
    except ValueError:
        margin = 36
    margin = max(0, margin)

    src = files[0]
    data = src.read()

    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            rect = page.rect
            new_rect = fitz.Rect(
                rect.x0 + margin,
                rect.y0 + margin,
                rect.x1 - margin,
                rect.y1 - margin,
            )
            if new_rect.width > 10 and new_rect.height > 10:
                page.set_cropbox(new_rect)
        out_bytes = doc.tobytes()

    return send_bytes(out_bytes, "cropped.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Word (.docx) -> PDF  (text-based conversion, not full layout fidelity)
# ---------------------------------------------------------------------------

@app.route("/api/word-to-pdf", methods=["POST"])
def word_to_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a .docx file."}), 400

    src = files[0]
    base_name = os.path.splitext(secure_filename(src.filename))[0] or "document"
    document = Document(src.stream)

    buf = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buf, pagesize=letter)
    styles = getSampleStyleSheet()
    flowables = []

    for para in document.paragraphs:
        text = para.text.strip()
        if not text:
            flowables.append(Spacer(1, 10))
            continue
        style_name = "Heading2" if para.style and "Heading" in (para.style.name or "") else "Normal"
        flowables.append(Paragraph(text.replace("&", "&amp;").replace("<", "&lt;"), styles[style_name]))
        flowables.append(Spacer(1, 6))

    if not flowables:
        flowables = [Paragraph("(No text content found in this document.)", styles["Normal"])]

    pdf_doc.build(flowables)
    buf.seek(0)

    return send_bytes(buf.getvalue(), f"{base_name}.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# PowerPoint (.pptx) -> PDF (one page per slide, text content only)
# ---------------------------------------------------------------------------

@app.route("/api/pptx-to-pdf", methods=["POST"])
def pptx_to_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a .pptx file."}), 400

    src = files[0]
    base_name = os.path.splitext(secure_filename(src.filename))[0] or "presentation"
    prs = Presentation(src.stream)

    buf = io.BytesIO()
    c = pdf_canvas.Canvas(buf, pagesize=letter)
    width, height = letter

    for slide in prs.slides:
        y = height - 60
        c.setFont("Helvetica-Bold", 16)
        texts = []
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                texts.append(shape.text_frame.text.strip())

        if texts:
            c.drawString(50, y, texts[0][:90])
            y -= 40
            c.setFont("Helvetica", 12)
            for block in texts[1:]:
                for line in block.split("\n"):
                    if y < 50:
                        c.showPage()
                        y = height - 60
                        c.setFont("Helvetica", 12)
                    c.drawString(60, y, line[:100])
                    y -= 18
        else:
            c.setFont("Helvetica", 12)
            c.drawString(50, y, "(No text content on this slide.)")

        c.showPage()

    c.save()
    buf.seek(0)

    return send_bytes(buf.getvalue(), f"{base_name}.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Excel (.xlsx) -> PDF (one table per sheet)
# ---------------------------------------------------------------------------

@app.route("/api/excel-to-pdf", methods=["POST"])
def excel_to_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a .xlsx file."}), 400

    src = files[0]
    base_name = os.path.splitext(secure_filename(src.filename))[0] or "spreadsheet"
    wb = load_workbook(src.stream, data_only=True)

    buf = io.BytesIO()
    pdf_doc = SimpleDocTemplate(buf, pagesize=letter)
    styles = getSampleStyleSheet()
    flowables = []

    for sheet in wb.worksheets:
        flowables.append(Paragraph(sheet.title, styles["Heading2"]))
        flowables.append(Spacer(1, 8))

        rows = []
        for row in sheet.iter_rows(values_only=True):
            rows.append(["" if v is None else str(v) for v in row])

        if rows:
            # Cap columns so wide sheets don't overflow the page unreadably
            max_cols = 10
            rows = [r[:max_cols] for r in rows]
            table = Table(rows, repeatRows=1)
            table.setStyle(TableStyle([
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]))
            flowables.append(table)
        else:
            flowables.append(Paragraph("(Empty sheet)", styles["Normal"]))

        flowables.append(Spacer(1, 20))

    pdf_doc.build(flowables)
    buf.seek(0)

    return send_bytes(buf.getvalue(), f"{base_name}.pdf", "application/pdf")


# ---------------------------------------------------------------------------
# Error handlers
# ---------------------------------------------------------------------------

@app.errorhandler(413)
def too_large(_e):
    return jsonify({"error": "File too large. Max size is 50MB."}), 413


@app.errorhandler(500)
def server_error(e):
    return jsonify({"error": f"Something went wrong: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    app.run(host="0.0.0.0", port=port, debug=False)
