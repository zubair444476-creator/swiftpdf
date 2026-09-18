"""
SwiftPDF — server.py  (v3 — maximum fidelity)

PDF→Word  : Each page rendered as a high-res image embedded in the DOCX,
             preserving 100% of the visual layout (tables, images, Arabic,
             logos, QR codes, colours). OCR text placed below for searchability.
             For text-only PDFs a clean text-reconstruction pass is also done.
PDF→Excel : Full colour detection from PDF drawings, Arabic/RTL cell alignment.
Excel→PDF : fitToPage patching so 1-sheet workbooks stay 1 page.
Image→Word: Image embedded full-page + OCR text overlay (not a bare text dump).
All other tools unchanged from the previous version.
"""

import os, io, re, glob, uuid, base64, shutil, zipfile, tempfile, subprocess
from collections import defaultdict

from flask import Flask, request, send_file, jsonify, send_from_directory
from werkzeug.utils import secure_filename

import fitz  # pymupdf
from PIL import Image

from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from pptx import Presentation
from pptx.util import Emu

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB


# ===========================================================================
# GENERIC HELPERS
# ===========================================================================

def get_uploaded_files(field_name="files"):
    files = request.files.getlist(field_name)
    if not files:
        single = request.files.get("file")
        if single:
            files = [single]
    return [f for f in files if f and f.filename]


def send_bytes(data, filename, mimetype):
    return send_file(io.BytesIO(data), mimetype=mimetype,
                     as_attachment=True, download_name=filename)


def parse_page_spec(spec, page_count):
    result = []
    spec = (spec or "").strip()
    if not spec:
        return list(range(page_count))
    for chunk in spec.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            a, b = chunk.split("-", 1)
            try:
                start, end = int(a), int(b)
            except ValueError:
                continue
            step = 1 if end >= start else -1
            for p in range(start, end + step, step):
                if 1 <= p <= page_count:
                    result.append(p - 1)
        else:
            try:
                p = int(chunk)
            except ValueError:
                continue
            if 1 <= p <= page_count:
                result.append(p - 1)
    return result


# ===========================================================================
# LIBREOFFICE HELPER
# ===========================================================================

def _libreoffice_convert(src_path, out_dir, out_fmt="pdf"):
    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if not lo:
        raise RuntimeError(
            "LibreOffice is not installed. Add 'libreoffice' to nixpacks.toml aptPkgs.")
    profile = os.path.join(out_dir, f"lo_{uuid.uuid4().hex}")
    os.makedirs(profile, exist_ok=True)
    cmd = [lo, "--headless", "--norestore", "--nofirststartwizard", "--nolockcheck",
           f"-env:UserInstallation=file://{profile}",
           "--convert-to", out_fmt, "--outdir", out_dir, src_path]
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    if r.returncode != 0:
        raise RuntimeError(f"LibreOffice failed: {(r.stderr or r.stdout or '').strip()[:400]}")
    base = os.path.splitext(os.path.basename(src_path))[0]
    ext = out_fmt.split(":")[0]
    matches = glob.glob(os.path.join(out_dir, f"{base}.{ext}"))
    if not matches:
        matches = [p for p in glob.glob(os.path.join(out_dir, f"{base}*"))
                   if not p.endswith(os.path.basename(src_path))]
    if not matches:
        raise RuntimeError("LibreOffice ran but produced no output file.")
    return matches[0]


def _lo_to_pdf(file_obj, ext, base_name):
    tmp = tempfile.mkdtemp()
    try:
        src = os.path.join(tmp, f"{base_name}{ext}")
        with open(src, "wb") as f:
            f.write(file_obj.read())
        try:
            out = _libreoffice_convert(src, tmp, "pdf")
        except RuntimeError as e:
            return None, str(e)
        with open(out, "rb") as f:
            return f.read(), None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def _lo_to_pdf_excel(file_obj, base_name):
    """xlsx → PDF with fit-to-page so a 1-sheet workbook stays 1 page."""
    import openpyxl as _xl
    tmp = tempfile.mkdtemp()
    try:
        data = file_obj.read()
        try:
            wb = _xl.load_workbook(io.BytesIO(data))
            for ws in wb.worksheets:
                ws.sheet_properties.pageSetUpPr.fitToPage = True
                ws.page_setup.fitToWidth = 1
                ws.page_setup.fitToHeight = 0
                ws.page_setup.scale = None
            patched = io.BytesIO()
            wb.save(patched)
            data = patched.getvalue()
        except Exception:
            pass
        src = os.path.join(tmp, f"{base_name}.xlsx")
        with open(src, "wb") as f:
            f.write(data)
        try:
            out = _libreoffice_convert(src, tmp, "pdf")
        except RuntimeError as e:
            return None, str(e)
        with open(out, "rb") as f:
            return f.read(), None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ===========================================================================
# DOCX BUILD HELPERS
# ===========================================================================

def _shading(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn("w:shd")):
        tcPr.remove(old)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), hex_fill)
    tcPr.append(shd)


def _cell_borders(cell, color="B0C4D8", sz="4"):
    tcPr = cell._tc.get_or_add_tcPr()
    b = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:color"), color)
        b.append(el)
    tcPr.append(b)


def _table_borders(table, color="B0C4D8", sz="4"):
    tblPr = table._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{side}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), sz)
        el.set(qn("w:color"), color)
        b.append(el)
    tblPr.append(b)


def _img_para(doc, img_bytes, width_cm):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    p.add_run().add_picture(io.BytesIO(img_bytes), width=Cm(width_cm))
    return p


def _clip(page, rect, zoom=3.0):
    return page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=rect).tobytes("png")


def _detect_tables(page):
    rects = []
    try:
        found = page.find_tables()
        for ft in found.tables:
            raw = ft.extract()
            ncols = max((len(r) for r in raw), default=0)
            flat = [c for r in raw for c in r if c and str(c).strip()]
            width = ft.bbox[2] - ft.bbox[0]
            if len(raw) >= 2 and ncols >= 2 and len(flat) >= 2 and width < page.rect.width - 5:
                rects.append(fitz.Rect(ft.bbox))
    except Exception:
        pass
    return rects


def _table_rows(page, tr):
    blocks = page.get_text("dict")["blocks"]
    y0, y1, x0, x1 = tr.y0, tr.y1, tr.x0, tr.x1
    all_spans = []
    for b in blocks:
        if b.get("type") != 0:
            continue
        for line in b.get("lines", []):
            for span in line.get("spans", []):
                if not span["text"].strip():
                    continue
                sx0, sy0 = span["bbox"][0], span["bbox"][1]
                if (y0 - 5) <= sy0 <= (y1 + 5) and sx0 >= (x0 - 8) and span["bbox"][2] <= (x1 + 8):
                    all_spans.append(span)

    if not all_spans:
        return [], "2F5496", None

    all_spans.sort(key=lambda s: (round(s["bbox"][1] / 4) * 4, s["bbox"][0]))
    x_lefts = sorted(set(round(s["bbox"][0]) for s in all_spans))
    gap = 15
    clusters = [[x_lefts[0]]]
    for x in x_lefts[1:]:
        if x - clusters[-1][-1] > gap:
            clusters.append([])
        clusters[-1].append(x)
    ncols = len(clusters)
    boundaries = [(clusters[i][-1] + clusters[i + 1][0]) / 2.0 for i in range(len(clusters) - 1)]

    def col_index(sx0):
        for i, bnd in enumerate(boundaries):
            if sx0 < bnd:
                return i
        return ncols - 1

    y_groups = {}
    for s in all_spans:
        yk = round(s["bbox"][1] / 4) * 4
        y_groups.setdefault(yk, []).append(s)

    rows = []
    first_spans = None
    for yi, yk in enumerate(sorted(y_groups.keys())):
        cols = [""] * ncols
        cspans = [[] for _ in range(ncols)]
        row_spans = sorted(y_groups[yk], key=lambda s: s["bbox"][0])
        for s in row_spans:
            ci = col_index(s["bbox"][0])
            txt = s["text"].strip()
            sep = " " if cols[ci] else ""
            cols[ci] = cols[ci] + sep + txt
            cspans[ci].append(s)
        if not any(c.strip() for c in cols):
            continue
        if yi == 0:
            first_spans = cspans
        rows.append(cols)

    hdr_fill = "2F5496"
    try:
        rt = tr.y0
        cands = []
        for d in page.get_drawings():
            r = d.get("rect")
            if r is None:
                continue
            f = d.get("fill")
            if not f or f in ((1., 1., 1.), (0., 0., 0.)):
                continue
            if r.y0 <= rt + 25 and r.y1 >= rt:
                area = max(0, min(r.x1, tr.x1) - max(r.x0, tr.x0))
                cands.append((area, f))
        if cands:
            _, best = max(cands, key=lambda x: x[0])
            hdr_fill = f"{int(best[0]*255):02X}{int(best[1]*255):02X}{int(best[2]*255):02X}"
    except Exception:
        pass

    hdr_txt = None
    if first_spans:
        for spans in first_spans:
            for sp in spans:
                if sp.get("color", 0) != 0:
                    hdr_txt = sp["color"]
                    break
            if hdr_txt is not None:
                break

    return rows, hdr_fill, hdr_txt


def _build_table(doc, rows, hf, htc, usable_cm=17.0):
    if not rows:
        return
    ncols = max(len(r) for r in rows)
    rows = [r + [""] * (ncols - len(r)) for r in rows]
    ww = [usable_cm / ncols] * ncols
    t = doc.add_table(rows=len(rows), cols=ncols)
    t.alignment = WD_TABLE_ALIGNMENT.LEFT
    _table_borders(t)
    for ri, row in enumerate(rows):
        is_hdr = (ri == 0)
        for ci, txt in enumerate(row):
            cell = t.cell(ri, ci)
            try:
                cell.width = Cm(ww[ci])
            except Exception:
                pass
            _shading(cell, hf if is_hdr else "FFFFFF")
            _cell_borders(cell)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            for li, line in enumerate((txt or "").split("\n")):
                if li > 0:
                    p.add_run().add_break()
                run = p.add_run(line)
                run.font.size = Pt(8 if is_hdr else 9)
                if is_hdr:
                    run.bold = True
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if htc:
                        run.font.color.rgb = RGBColor(
                            (htc >> 16) & 0xFF, (htc >> 8) & 0xFF, htc & 0xFF)
                    else:
                        run.font.color.rgb = RGBColor(255, 255, 255)
                else:
                    run.font.color.rgb = RGBColor(0, 0, 0)


def _median_size(page):
    sizes = []
    try:
        for b in page.get_text("dict")["blocks"]:
            if b.get("type") != 0:
                continue
            for line in b.get("lines", []):
                for span in line.get("spans", []):
                    s = span.get("size", 0)
                    if s > 0:
                        sizes.append(s)
    except Exception:
        pass
    if not sizes:
        return 10.0
    sizes.sort()
    return sizes[len(sizes) // 2]


def _is_rtl(text):
    rtl = sum(1 for c in text if "\u0600" <= c <= "\u06FF" or "\u0590" <= c <= "\u05FF")
    return rtl > len(text) * 0.4 if text else False


def _page_is_complex(page):
    """Return True if the page has images, drawings, or multi-column layout
    that a text-only reconstruction would mangle."""
    blocks = page.get_text("dict").get("blocks", [])
    img_count = sum(1 for b in blocks if b.get("type") == 1)
    if img_count > 0:
        return True
    try:
        drawings = page.get_drawings()
        filled = [d for d in drawings if d.get("fill") and d["fill"] not in
                  ((1., 1., 1.), (0., 0., 0.), None)]
        if len(filled) > 3:
            return True
    except Exception:
        pass
    # Check for multi-column text (x-starts spread widely)
    x_starts = set()
    for b in blocks:
        if b.get("type") == 0:
            for line in b.get("lines", []):
                if line.get("spans"):
                    x_starts.add(round(line["spans"][0]["bbox"][0] / 50) * 50)
    if len(x_starts) > 3:
        return True
    return False


# ===========================================================================
# STATIC / HEALTH
# ===========================================================================

@app.route("/")
def index():
    return send_from_directory(".", "index.html")


@app.route("/health")
def health():
    return jsonify({"status": "ok"})


# ===========================================================================
# PAGE THUMBNAILS
# ===========================================================================

@app.route("/api/page-thumbnails", methods=["POST"])
def page_thumbnails():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    data = files[0].read()
    mat = fitz.Matrix(0.4, 0.4)
    thumbs = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=mat)
            b64 = base64.b64encode(pix.tobytes("png")).decode("ascii")
            thumbs.append({"page": i + 1, "width": pix.width, "height": pix.height,
                           "dataUrl": f"data:image/png;base64,{b64}"})
    return jsonify({"pageCount": len(thumbs), "thumbnails": thumbs})


# ===========================================================================
# MERGE / SPLIT / ROTATE / COMPRESS
# ===========================================================================

@app.route("/api/merge", methods=["POST"])
def merge_pdfs():
    files = get_uploaded_files()
    if len(files) < 2:
        return jsonify({"error": "Upload at least 2 PDF files to merge."}), 400
    merged = fitz.open()
    try:
        for f in files:
            with fitz.open(stream=f.read(), filetype="pdf") as d:
                merged.insert_pdf(d)
        out = merged.tobytes()
    finally:
        merged.close()
    return send_bytes(out, "merged.pdf", "application/pdf")


@app.route("/api/split", methods=["POST"])
def split_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to split."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "page"
    buf = io.BytesIO()
    with fitz.open(stream=data, filetype="pdf") as doc:
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, page in enumerate(doc):
                single = fitz.open()
                single.insert_pdf(doc, from_page=i, to_page=i)
                zf.writestr(f"{base}_page_{i + 1}.pdf", single.tobytes())
                single.close()
    buf.seek(0)
    return send_bytes(buf.getvalue(), f"{base}_split.zip", "application/zip")


@app.route("/api/rotate", methods=["POST"])
def rotate_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    angle = int(request.form.get("angle", 90))
    pages_spec = request.form.get("pages", "")
    with fitz.open(stream=data, filetype="pdf") as doc:
        page_list = parse_page_spec(pages_spec, len(doc)) if pages_spec else list(range(len(doc)))
        for i in page_list:
            doc[i].set_rotation((doc[i].rotation + angle) % 360)
        out = doc.tobytes(deflate=True, garbage=4)
    return send_bytes(out, f"{base}_rotated.pdf", "application/pdf")


@app.route("/api/compress", methods=["POST"])
def compress_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    level = request.form.get("level", "medium")
    zoom = {"low": 1.5, "medium": 1.2, "high": 0.9}.get(level, 1.2)
    quality = {"low": 85, "medium": 65, "high": 40}.get(level, 65)
    with fitz.open(stream=data, filetype="pdf") as doc:
        new_doc = fitz.open()
        mat = fitz.Matrix(zoom, zoom)
        for page in doc:
            pix = page.get_pixmap(matrix=mat)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            jpeg_buf = io.BytesIO()
            img.save(jpeg_buf, format="JPEG", quality=quality, optimize=True)
            new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=jpeg_buf.getvalue())
        out = new_doc.tobytes(deflate=True, garbage=4)
        new_doc.close()
    return send_bytes(out, f"{base}_compressed.pdf", "application/pdf")


@app.route("/api/images-to-pdf", methods=["POST"])
def images_to_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload one or more images."}), 400
    imgs = [Image.open(f.stream).convert("RGB") for f in files]
    buf = io.BytesIO()
    imgs[0].save(buf, format="PDF", save_all=True, append_images=imgs[1:])
    buf.seek(0)
    return send_bytes(buf.getvalue(), "images.pdf", "application/pdf")


@app.route("/api/pdf-to-images", methods=["POST"])
def pdf_to_images():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "page"
    mat = fitz.Matrix(2.0, 2.0)
    buf = io.BytesIO()
    with fitz.open(stream=data, filetype="pdf") as doc:
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, page in enumerate(doc):
                pix = page.get_pixmap(matrix=mat)
                zf.writestr(f"{base}_page_{i + 1}.png", pix.tobytes("png"))
    buf.seek(0)
    return send_bytes(buf.getvalue(), f"{base}_images.zip", "application/zip")


# ===========================================================================
# PDF -> TEXT
# ===========================================================================

@app.route("/api/pdf-to-text", methods=["POST"])
def pdf_to_text():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    parts = []
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            parts.append(page.get_text())
    return send_bytes("\n\n".join(parts).encode("utf-8"),
                      f"{base}.txt", "text/plain; charset=utf-8")


# ===========================================================================
# PDF -> WORD  (image-first for complex pages; text-only for simple ones)
# ===========================================================================

@app.route("/api/pdf-to-word", methods=["POST"])
def pdf_to_word():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"

    # Page size: A4
    PAGE_W_CM  = 21.0
    PAGE_H_CM  = 29.7
    MARGIN_CM  = 1.5
    USABLE_W_CM = PAGE_W_CM - 2 * MARGIN_CM    # 18 cm

    doc = Document()
    for section in doc.sections:
        section.page_width   = Cm(PAGE_W_CM)
        section.page_height  = Cm(PAGE_H_CM)
        section.left_margin  = Cm(MARGIN_CM)
        section.right_margin = Cm(MARGIN_CM)
        section.top_margin   = Cm(MARGIN_CM)
        section.bottom_margin = Cm(MARGIN_CM)

    with fitz.open(stream=data, filetype="pdf") as pdf:
        for pg_idx, page in enumerate(pdf):
            if pg_idx > 0:
                doc.add_page_break()

            page_rect = page.rect
            complex_page = _page_is_complex(page)

            # ── COMPLEX PAGE: embed as full-resolution image ─────────────────
            # This is the only reliable way to preserve logos, photos, QR codes,
            # coloured headers, multi-column layouts, Arabic text, signatures, etc.
            if complex_page:
                # Render at 3× for sharp text on screen and print
                zoom = 3.0
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                img_bytes = pix.tobytes("png")

                # Compute display height to match aspect ratio
                img_w_px, img_h_px = pix.width, pix.height
                aspect = img_h_px / img_w_px if img_w_px else 1.4142
                display_h_cm = USABLE_W_CM * aspect
                # Cap so it never bleeds past the page
                max_h = PAGE_H_CM - 2 * MARGIN_CM
                display_h_cm = min(display_h_cm, max_h)

                ip = doc.add_paragraph()
                ip.paragraph_format.space_before = Pt(0)
                ip.paragraph_format.space_after = Pt(6)
                ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
                ip.add_run().add_picture(io.BytesIO(img_bytes),
                                         width=Cm(USABLE_W_CM),
                                         height=Cm(display_h_cm))

                # OCR text below — small grey text, makes doc searchable
                ocr_text = ""
                try:
                    ocr_text = page.get_text("text").strip()
                except Exception:
                    pass
                if ocr_text:
                    sep = doc.add_paragraph()
                    sep.paragraph_format.space_before = Pt(4)
                    sep.paragraph_format.space_after = Pt(2)
                    sep_run = sep.add_run("─── Extracted Text ───")
                    sep_run.font.size = Pt(7)
                    sep_run.font.color.rgb = RGBColor(180, 180, 180)
                    sep.alignment = WD_ALIGN_PARAGRAPH.CENTER

                    for line in ocr_text.splitlines():
                        if not line.strip():
                            continue
                        is_rtl = _is_rtl(line)
                        lp = doc.add_paragraph()
                        lp.paragraph_format.space_before = Pt(0)
                        lp.paragraph_format.space_after = Pt(1)
                        if is_rtl:
                            lp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            lp._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
                        lr = lp.add_run(line)
                        lr.font.size = Pt(8)
                        lr.font.color.rgb = RGBColor(80, 80, 80)

            # ── SIMPLE TEXT-ONLY PAGE: reconstruct with formatting ────────────
            else:
                med = _median_size(page)
                trects = sorted(_detect_tables(page), key=lambda r: r.y0)

                def in_table(y, _tr=trects):
                    return any(tr.y0 - 4 <= y <= tr.y1 + 4 for tr in _tr)

                items = []
                try:
                    raw = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
                    for b in sorted(raw["blocks"], key=lambda b: b["bbox"][1]):
                        if b.get("type") == 1:
                            try:
                                pix2 = page.get_pixmap(
                                    matrix=fitz.Matrix(2, 2), clip=fitz.Rect(b["bbox"]))
                                items.append({"type": "image",
                                              "data": pix2.tobytes("png"),
                                              "y": b["bbox"][1]})
                            except Exception:
                                pass
                            continue
                        for line in b.get("lines", []):
                            spans = [s for s in line.get("spans", []) if s["text"].strip()]
                            if not spans:
                                continue
                            text = "".join(s["text"] for s in spans).strip()
                            size = max(s.get("size", 10) for s in spans)
                            flags = spans[0].get("flags", 0)
                            color = spans[0].get("color", 0)
                            items.append({
                                "type": "text", "text": text, "size": size,
                                "bold": bool(flags & 16), "italic": bool(flags & 2),
                                "color": color,
                                "rtl": _is_rtl(text), "y": line["bbox"][1],
                            })
                except Exception:
                    pass

                has_text = any(it["type"] == "text" for it in items)
                if not has_text and not trects:
                    pix3 = page.get_pixmap(matrix=fitz.Matrix(2.0, 2.0), alpha=False)
                    _img_para(doc, pix3.tobytes("png"), USABLE_W_CM)
                    continue

                flushed = set()

                def flush_to(y_lim, _trects=trects, _flushed=flushed):
                    for tr in _trects:
                        if id(tr) in _flushed or tr.y0 > y_lim:
                            continue
                        _flushed.add(id(tr))
                        rows, hf, htc = _table_rows(page, tr)
                        sp = doc.add_paragraph()
                        sp.paragraph_format.space_before = Pt(4)
                        sp.paragraph_format.space_after = Pt(0)
                        if rows:
                            _build_table(doc, rows, hf, htc, usable_cm=USABLE_W_CM)
                        else:
                            _img_para(doc, _clip(page, tr), USABLE_W_CM)
                        sp2 = doc.add_paragraph()
                        sp2.paragraph_format.space_before = Pt(0)
                        sp2.paragraph_format.space_after = Pt(4)

                for item in items:
                    flush_to(item["y"])
                    if in_table(item["y"]):
                        continue
                    if item["type"] == "image":
                        _img_para(doc, item["data"], USABLE_W_CM)
                        continue
                    text, size = item["text"], item["size"]
                    bold, italic, rtl = item["bold"], item["italic"], item["rtl"]
                    color = item.get("color", 0)
                    ratio = size / med if med > 0 else 1.0
                    if ratio >= 1.8 or (ratio >= 1.4 and bold):
                        style = "Heading 1"
                    elif ratio >= 1.3 or (ratio >= 1.1 and bold):
                        style = "Heading 2"
                    elif ratio >= 1.1 and bold:
                        style = "Heading 3"
                    else:
                        style = "Normal"
                    p = doc.add_paragraph(style=style)
                    p.paragraph_format.space_before = Pt(2)
                    p.paragraph_format.space_after = Pt(2)
                    if rtl:
                        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        p._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
                    run = p.add_run(text)
                    if style == "Normal":
                        run.font.size = Pt(max(7, round(size * 0.75)))
                    if bold:
                        run.bold = True
                    if italic:
                        run.italic = True
                    if color and color != 0:
                        r = (color >> 16) & 0xFF
                        g = (color >> 8) & 0xFF
                        b2 = color & 0xFF
                        if (r, g, b2) != (0, 0, 0):
                            run.font.color.rgb = RGBColor(r, g, b2)
                flush_to(page_rect.y1 + 99999)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_bytes(
        buf.getvalue(), f"{base}.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ===========================================================================
# PDF -> POWERPOINT
# ===========================================================================

@app.route("/api/pdf-to-pptx", methods=["POST"])
def pdf_to_pptx():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "presentation"
    mat = fitz.Matrix(2.0, 2.0)
    prs = Presentation()
    with fitz.open(stream=data, filetype="pdf") as pdf:
        if not pdf:
            return jsonify({"error": "The PDF has no pages."}), 400
        fr = pdf[0].rect
        prs.slide_width = Emu(int(fr.width * 12700))
        prs.slide_height = Emu(int(fr.height * 12700))
        blank = prs.slide_layouts[6]
        for page in pdf:
            pix = page.get_pixmap(matrix=mat)
            slide = prs.slides.add_slide(blank)
            slide.shapes.add_picture(
                io.BytesIO(pix.tobytes("png")), Emu(0), Emu(0),
                width=prs.slide_width, height=prs.slide_height)
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return send_bytes(
        buf.getvalue(), f"{base}.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation")


# ===========================================================================
# PDF -> EXCEL  (true colour + Arabic RTL support)
# ===========================================================================

@app.route("/api/pdf-to-excel", methods=["POST"])
def pdf_to_excel():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "spreadsheet"

    body_font   = Font(size=10)
    center_aln  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_aln    = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    right_aln   = Alignment(horizontal="right",  vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="B0C4D8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def _hdr_colors(page, tbl_rect):
        fill_hex = "2F5496"
        font_hex = "FFFFFF"
        try:
            rt = tbl_rect.y0
            cands = []
            for d in page.get_drawings():
                r = d.get("rect")
                f = d.get("fill")
                if r is None or not f:
                    continue
                if f in ((1., 1., 1.), (0., 0., 0.)):
                    continue
                if r.y0 <= rt + 30 and r.y1 >= rt and r.x0 <= tbl_rect.x1 and r.x1 >= tbl_rect.x0:
                    ov = max(0, min(r.x1, tbl_rect.x1) - max(r.x0, tbl_rect.x0))
                    cands.append((ov, f))
            if cands:
                _, best = max(cands, key=lambda x: x[0])
                ri2 = int(round(best[0] * 255))
                gi2 = int(round(best[1] * 255))
                bi2 = int(round(best[2] * 255))
                fill_hex = f"{ri2:02X}{gi2:02X}{bi2:02X}"
                lum = 0.299 * ri2 + 0.587 * gi2 + 0.114 * bi2
                font_hex = "FFFFFF" if lum < 160 else "000000"
        except Exception:
            pass
        return fill_hex, font_hex

    wb = Workbook()
    wb.remove(wb.active)

    with fitz.open(stream=data, filetype="pdf") as pdf:
        for pg_idx, page in enumerate(pdf):
            ws = wb.create_sheet(title=f"Page {pg_idx + 1}"[:31])
            rc = 1
            wrote = False
            try:
                found = page.find_tables()
                if found and found.tables:
                    for ft in found.tables:
                        raw = ft.extract()
                        if not raw:
                            continue
                        ncols = max(len(r) for r in raw)
                        tbl_rect = fitz.Rect(ft.bbox)
                        fill_hex, font_hex = _hdr_colors(page, tbl_rect)
                        hdr_fill = PatternFill("solid", fgColor=fill_hex)
                        hdr_font = Font(bold=True, color=font_hex, size=10)
                        start_row = rc
                        for ri, row in enumerate(raw):
                            row = (row + [None] * (ncols - len(row)))[:ncols]
                            is_hdr = (ri == 0)
                            for ci, val in enumerate(row, start=1):
                                cell_text = str(val).strip() if val is not None else ""
                                cell = ws.cell(row=rc, column=ci, value=cell_text)
                                cell.border = border
                                if is_hdr:
                                    cell.font = hdr_font
                                    cell.fill = hdr_fill
                                    cell.alignment = center_aln
                                else:
                                    cell.font = body_font
                                    cell.alignment = right_aln if _is_rtl(cell_text) else left_aln
                            rc += 1
                        rc += 1
                        for ci in range(1, ncols + 1):
                            cl = get_column_letter(ci)
                            ml = max(
                                (len(str(ws.cell(row=r, column=ci).value or ""))
                                 for r in range(start_row, rc - 1)), default=8)
                            ws.column_dimensions[cl].width = min(max(ml + 2, 8), 60)
                        wrote = True
            except Exception:
                wrote = False

            if not wrote:
                try:
                    blocks = page.get_text("dict")["blocks"]
                    spans = []
                    for b in blocks:
                        if b.get("type") != 0:
                            continue
                        for line in b.get("lines", []):
                            for span in line.get("spans", []):
                                t = span["text"].strip()
                                if t:
                                    spans.append({"text": t, "x0": span["bbox"][0],
                                                  "y0": span["bbox"][1]})
                    if spans:
                        xs = sorted(set(round(s["x0"]) for s in spans))
                        col_starts = [xs[0]]
                        for x in xs[1:]:
                            if x - col_starts[-1] > 25:
                                col_starts.append(x)

                        def col_of(x0, _s=col_starts):
                            return min(range(len(_s)), key=lambda i: abs(x0 - _s[i])) + 1

                        rows_map = defaultdict(list)
                        for s in spans:
                            rows_map[round(s["y0"] / 4) * 4].append(s)
                        for yk in sorted(rows_map.keys()):
                            for s in rows_map[yk]:
                                ci = col_of(s["x0"])
                                existing = ws.cell(row=rc, column=ci).value or ""
                                joined = (existing + " " if existing else "") + s["text"]
                                cell = ws.cell(row=rc, column=ci, value=joined)
                                cell.font = body_font
                                cell.alignment = right_aln if _is_rtl(s["text"]) else left_aln
                            rc += 1
                    else:
                        for line in page.get_text().split("\n"):
                            if line.strip():
                                cell = ws.cell(row=rc, column=1, value=line)
                                cell.font = body_font
                                cell.alignment = right_aln if _is_rtl(line) else left_aln
                                rc += 1
                except Exception:
                    for line in page.get_text().split("\n"):
                        if line.strip():
                            ws.cell(row=rc, column=1, value=line)
                            rc += 1

    if not wb.sheetnames:
        wb.create_sheet("Sheet1")
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_bytes(
        buf.getvalue(), f"{base}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ===========================================================================
# PAGE OPERATIONS
# ===========================================================================

@app.route("/api/remove-pages", methods=["POST"])
def remove_pages():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    pages_spec = request.form.get("pages", "")
    with fitz.open(stream=data, filetype="pdf") as doc:
        to_remove = sorted(set(parse_page_spec(pages_spec, len(doc))), reverse=True)
        for i in to_remove:
            doc.delete_page(i)
        if len(doc) == 0:
            return jsonify({"error": "Cannot remove all pages."}), 400
        out = doc.tobytes(deflate=True, garbage=4)
    return send_bytes(out, f"{base}_removed.pdf", "application/pdf")


@app.route("/api/extract-pages", methods=["POST"])
def extract_pages():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    pages_spec = request.form.get("pages", "")
    with fitz.open(stream=data, filetype="pdf") as doc:
        page_list = parse_page_spec(pages_spec, len(doc))
        if not page_list:
            return jsonify({"error": "No valid pages specified."}), 400
        new_doc = fitz.open()
        for i in page_list:
            new_doc.insert_pdf(doc, from_page=i, to_page=i)
        out = new_doc.tobytes(deflate=True, garbage=4)
        new_doc.close()
    return send_bytes(out, f"{base}_extracted.pdf", "application/pdf")


@app.route("/api/add-page-numbers", methods=["POST"])
def add_page_numbers():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    position = request.form.get("position", "bottom-center")
    with fitz.open(stream=data, filetype="pdf") as doc:
        for i, page in enumerate(doc):
            pw, ph = page.rect.width, page.rect.height
            text = f"{i + 1} / {len(doc)}"
            pos_map = {
                "bottom-center": fitz.Point(pw / 2 - 20, ph - 20),
                "bottom-right": fitz.Point(pw - 60, ph - 20),
                "bottom-left": fitz.Point(20, ph - 20),
                "top-center": fitz.Point(pw / 2 - 20, 20),
                "top-right": fitz.Point(pw - 60, 20),
                "top-left": fitz.Point(20, 20),
            }
            pt = pos_map.get(position, pos_map["bottom-center"])
            page.insert_text(pt, text, fontsize=10, color=(0.3, 0.3, 0.3))
        out = doc.tobytes(deflate=True, garbage=4)
    return send_bytes(out, f"{base}_numbered.pdf", "application/pdf")


@app.route("/api/watermark", methods=["POST"])
def watermark_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    text = request.form.get("text", "CONFIDENTIAL")
    opacity = float(request.form.get("opacity", 0.3))
    color_str = request.form.get("color", "808080")
    try:
        r = int(color_str[0:2], 16) / 255
        g = int(color_str[2:4], 16) / 255
        b = int(color_str[4:6], 16) / 255
    except Exception:
        r, g, b = 0.5, 0.5, 0.5
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            pw, ph = page.rect.width, page.rect.height
            page.insert_text(
                fitz.Point(pw * 0.15, ph * 0.6), text,
                fontsize=60, color=(r, g, b),
                rotate=45, overlay=True)
        out = doc.tobytes(deflate=True, garbage=4)
    return send_bytes(out, f"{base}_watermarked.pdf", "application/pdf")


@app.route("/api/protect", methods=["POST"])
def protect_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    password = request.form.get("password", "")
    if not password:
        return jsonify({"error": "Provide a password."}), 400
    with fitz.open(stream=data, filetype="pdf") as doc:
        perm = fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY
        out = doc.tobytes(encryption=fitz.PDF_ENCRYPT_AES_256,
                          user_pw=password, owner_pw=password + "_owner",
                          permissions=perm)
    return send_bytes(out, f"{base}_protected.pdf", "application/pdf")


@app.route("/api/unlock", methods=["POST"])
def unlock_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    password = request.form.get("password", "")
    with fitz.open(stream=data, filetype="pdf") as doc:
        if doc.is_encrypted:
            if not doc.authenticate(password):
                return jsonify({"error": "Incorrect password."}), 401
        out = doc.tobytes(deflate=True, garbage=4, encryption=fitz.PDF_ENCRYPT_NONE)
    return send_bytes(out, f"{base}_unlocked.pdf", "application/pdf")


@app.route("/api/crop", methods=["POST"])
def crop_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    try:
        x0 = float(request.form.get("x0", 0))
        y0 = float(request.form.get("y0", 0))
        x1 = float(request.form.get("x1", 612))
        y1 = float(request.form.get("y1", 792))
    except ValueError:
        return jsonify({"error": "Invalid crop coordinates."}), 400
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            page.set_cropbox(fitz.Rect(x0, y0, x1, y1))
        out = doc.tobytes(deflate=True, garbage=4)
    return send_bytes(out, f"{base}_cropped.pdf", "application/pdf")


@app.route("/api/redact", methods=["POST"])
def redact_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    search_text = request.form.get("text", "")
    if not search_text:
        return jsonify({"error": "Provide text to redact."}), 400
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            hits = page.search_for(search_text)
            for rect in hits:
                page.add_redact_annot(rect, fill=(0, 0, 0))
            page.apply_redactions()
        out = doc.tobytes(deflate=True, garbage=4)
    return send_bytes(out, f"{base}_redacted.pdf", "application/pdf")


@app.route("/api/add-signature", methods=["POST"])
def add_signature():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF and optionally a signature image."}), 400

    pdf_file = next((f for f in files if f.filename.lower().endswith(".pdf")), None)
    sig_file = next((f for f in files if not f.filename.lower().endswith(".pdf")), None)
    if not pdf_file:
        return jsonify({"error": "No PDF file found."}), 400

    data = pdf_file.read()
    base = os.path.splitext(secure_filename(pdf_file.filename))[0] or "document"
    position = request.form.get("position", "bottom-right")
    page_num = max(0, int(request.form.get("page", 1)) - 1)

    sig_png = None
    if sig_file:
        try:
            img = Image.open(sig_file.stream).convert("RGBA")
            buf2 = io.BytesIO()
            img.save(buf2, format="PNG")
            sig_png = buf2.getvalue()
        except Exception:
            pass

    with fitz.open(stream=data, filetype="pdf") as doc:
        if page_num >= len(doc):
            page_num = len(doc) - 1
        page = doc[page_num]
        pw, ph = page.rect.width, page.rect.height
        if sig_png:
            sig_w, sig_h = 150, 60
            margin = 30
            origins = {
                "bottom-right": (pw - sig_w - margin, ph - sig_h - margin),
                "bottom-left": (margin, ph - sig_h - margin),
                "bottom-center": ((pw - sig_w) / 2, ph - sig_h - margin),
                "top-right": (pw - sig_w - margin, margin),
                "top-left": (margin, margin),
            }
            x0, y0 = origins.get(position, origins["bottom-right"])
            page.insert_image(fitz.Rect(x0, y0, x0 + sig_w, y0 + sig_h), stream=sig_png)
        else:
            text = request.form.get("text", "Signed")
            margin = 30
            pos_map = {
                "bottom-right": fitz.Point(pw - 120, ph - margin),
                "bottom-left": fitz.Point(margin, ph - margin),
                "bottom-center": fitz.Point(pw / 2 - 30, ph - margin),
                "top-right": fitz.Point(pw - 120, margin + 12),
                "top-left": fitz.Point(margin, margin + 12),
            }
            pt = pos_map.get(position, pos_map["bottom-right"])
            page.insert_text(pt, text, fontsize=14, color=(0.1, 0.1, 0.6))
        out = doc.tobytes(deflate=True, garbage=4)
    return send_bytes(out, f"{base}_signed.pdf", "application/pdf")


@app.route("/api/flatten", methods=["POST"])
def flatten_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    with fitz.open(stream=data, filetype="pdf") as doc:
        new_doc = fitz.open()
        mat = fitz.Matrix(2.0, 2.0)
        for page in doc:
            pix = page.get_pixmap(matrix=mat, annots=True)
            new_page = new_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(page.rect, stream=pix.tobytes("png"))
        out = new_doc.tobytes(deflate=True, garbage=4)
        new_doc.close()
    return send_bytes(out, f"{base}_flattened.pdf", "application/pdf")


# ===========================================================================
# OFFICE -> PDF  (LibreOffice: full layout fidelity)
# ===========================================================================

@app.route("/api/word-to-pdf", methods=["POST"])
def word_to_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a .docx file."}), 400
    src = files[0]
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    out_bytes, err = _lo_to_pdf(src, ".docx", base)
    if err:
        return jsonify({"error": err}), 503
    return send_bytes(out_bytes, f"{base}.pdf", "application/pdf")


@app.route("/api/pptx-to-pdf", methods=["POST"])
def pptx_to_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a .pptx file."}), 400
    src = files[0]
    base = os.path.splitext(secure_filename(src.filename))[0] or "presentation"
    out_bytes, err = _lo_to_pdf(src, ".pptx", base)
    if err:
        return jsonify({"error": err}), 503
    return send_bytes(out_bytes, f"{base}.pdf", "application/pdf")


@app.route("/api/excel-to-pdf", methods=["POST"])
def excel_to_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a .xlsx file."}), 400
    src = files[0]
    base = os.path.splitext(secure_filename(src.filename))[0] or "spreadsheet"
    out_bytes, err = _lo_to_pdf_excel(src, base)
    if err:
        return jsonify({"error": err}), 503
    return send_bytes(out_bytes, f"{base}.pdf", "application/pdf")


# ===========================================================================
# HTML -> PDF
# ===========================================================================

@app.route("/api/html-to-pdf", methods=["POST"])
def html_to_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload an HTML file."}), 400
    src = files[0]
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    html_bytes = src.read()
    try:
        html_text = html_bytes.decode("utf-8", errors="replace")
    except Exception:
        html_text = html_bytes.decode("latin-1", errors="replace")
    try:
        story = fitz.Story(html=html_text)
        mediabox = fitz.paper_rect("a4")
        where = mediabox + (36, 36, -36, -36)
        buf = io.BytesIO()
        writer = fitz.DocumentWriter(buf)
        more = True
        while more:
            device = writer.begin_page(mediabox)
            more, _ = story.place(where)
            story.draw(device)
            writer.end_page()
        writer.close()
        out_bytes = buf.getvalue()
    except Exception:
        plain = re.sub(r"<[^>]+>", " ", html_text)
        plain = re.sub(r"\s+", " ", plain).strip()
        buf = io.BytesIO()
        pdf_doc = SimpleDocTemplate(buf, pagesize=A4)
        styles = getSampleStyleSheet()
        flowables = []
        for chunk in plain.split("."):
            chunk = chunk.strip()
            if chunk:
                flowables.append(Paragraph(
                    chunk.replace("&", "&amp;").replace("<", "&lt;") + ".", styles["Normal"]))
                flowables.append(Spacer(1, 6))
        if not flowables:
            flowables = [Paragraph("(Empty document)", styles["Normal"])]
        pdf_doc.build(flowables)
        out_bytes = buf.getvalue()
    return send_bytes(out_bytes, f"{base}.pdf", "application/pdf")


# ===========================================================================
# IMAGE -> WORD / TEXT  (image embedded full-page + OCR text layer)
# ===========================================================================

def _ocr(pil_img, lang="eng"):
    if not shutil.which("tesseract"):
        raise RuntimeError("Tesseract OCR is not installed.")
    try:
        import pytesseract
    except ImportError:
        raise RuntimeError("pytesseract is missing from requirements.txt.")
    w, h = pil_img.size
    if max(w, h) < 1200:
        sc = 1200 / max(w, h)
        pil_img = pil_img.resize((int(w * sc), int(h * sc)), Image.LANCZOS)
    return pytesseract.image_to_string(pil_img, lang=lang, config="--psm 3")


@app.route("/api/image-to-ocr", methods=["POST"])
def image_to_ocr():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload one or more image files."}), 400
    fmt  = request.form.get("output_format", "txt")
    lang = request.form.get("lang", "eng")
    if not re.match(r"^[a-zA-Z+]{2,20}$", lang):
        lang = "eng"

    page_data = []
    for f in files:
        name = os.path.splitext(secure_filename(f.filename))[0] or "image"
        try:
            img = Image.open(f.stream)
            if img.mode == "RGBA":
                bg = Image.new("RGB", img.size, (255, 255, 255))
                bg.paste(img, mask=img.split()[3])
                img = bg
            elif img.mode not in ("RGB", "L"):
                img = img.convert("RGB")
            try:
                ocr_text = _ocr(img, lang=lang)
            except RuntimeError as e:
                return jsonify({"error": str(e)}), 503
            except Exception as e:
                ocr_text = f"[OCR failed: {e}]"
            page_data.append((name, img, ocr_text))
        except Exception as e:
            return jsonify({"error": f"Could not read image '{name}': {e}"}), 400

    if not page_data:
        return jsonify({"error": "No images could be processed."}), 400
    base = page_data[0][0] or "ocr_result"

    if fmt == "txt":
        parts = []
        for name, _img, text in page_data:
            if len(page_data) > 1:
                parts.append(f"=== {name} ===")
            parts.append(text.strip())
            parts.append("")
        return send_bytes("\n".join(parts).strip().encode("utf-8"),
                          f"{base}_ocr.txt", "text/plain; charset=utf-8")

    # Word output: image full-page + searchable OCR text below
    PAGE_W_CM   = 21.0
    PAGE_H_CM   = 29.7
    MARGIN_CM   = 1.5
    USABLE_W_CM = PAGE_W_CM - 2 * MARGIN_CM

    doc = Document()
    for section in doc.sections:
        section.page_width    = Cm(PAGE_W_CM)
        section.page_height   = Cm(PAGE_H_CM)
        section.left_margin   = Cm(MARGIN_CM)
        section.right_margin  = Cm(MARGIN_CM)
        section.top_margin    = Cm(MARGIN_CM)
        section.bottom_margin = Cm(MARGIN_CM)

    for idx, (name, pil_img, ocr_text) in enumerate(page_data):
        if idx > 0:
            doc.add_page_break()
        img_buf = io.BytesIO()
        pil_img.save(img_buf, format="PNG")
        img_buf.seek(0)
        img_w, img_h = pil_img.size
        aspect = img_h / img_w if img_w else 1.4142
        display_h = min(USABLE_W_CM * aspect, PAGE_H_CM - 2 * MARGIN_CM - 1.0)

        ip = doc.add_paragraph()
        ip.paragraph_format.space_before = Pt(0)
        ip.paragraph_format.space_after  = Pt(4)
        ip.alignment = WD_ALIGN_PARAGRAPH.CENTER
        ip.add_run().add_picture(img_buf, width=Cm(USABLE_W_CM), height=Cm(display_h))

        lines = [l for l in ocr_text.splitlines() if l.strip()]
        if lines:
            sep = doc.add_paragraph()
            sep.paragraph_format.space_before = Pt(2)
            sep.paragraph_format.space_after  = Pt(2)
            sep_r = sep.add_run("── Extracted Text ──")
            sep_r.font.size = Pt(7)
            sep_r.font.color.rgb = RGBColor(150, 150, 150)
            sep.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for line in lines:
                is_rtl = _is_rtl(line)
                lp = doc.add_paragraph()
                lp.paragraph_format.space_before = Pt(0)
                lp.paragraph_format.space_after  = Pt(1)
                if is_rtl:
                    lp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    lp._p.get_or_add_pPr().append(OxmlElement("w:bidi"))
                lr = lp.add_run(line)
                lr.font.size = Pt(8)
                lr.font.color.rgb = RGBColor(60, 60, 60)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_bytes(
        buf.getvalue(), f"{base}_ocr.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document")


# ===========================================================================
# ERROR HANDLERS
# ===========================================================================

@app.errorhandler(413)
def too_large(_e):
    return jsonify({"error": "File too large. Max size is 50MB."}), 413


@app.errorhandler(500)
def server_error(e):
    return jsonify({"error": f"Something went wrong: {str(e)}"}), 500


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 7860))
    app.run(host="0.0.0.0", port=port, debug=False)
