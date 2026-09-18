"""
SwiftPDF — server.py  (complete rewrite)
All tools implemented correctly. LibreOffice handles Office<->PDF conversions
for full layout fidelity; PyMuPDF handles native PDF operations with
proper text-block / table extraction for PDF->Word/Excel/PPTX.
"""

import os
import io
import re
import glob
import uuid
import base64
import shutil
import zipfile
import tempfile
import subprocess
from collections import defaultdict

from flask import Flask, request, send_file, jsonify, send_from_directory
from werkzeug.utils import secure_filename

import fitz  # pymupdf
from PIL import Image

from docx import Document
from docx.shared import Pt, Cm, RGBColor
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
    return send_file(
        io.BytesIO(data),
        mimetype=mimetype,
        as_attachment=True,
        download_name=filename,
    )


def parse_page_spec(spec, page_count):
    """Parse '1,3,5-7' (1-indexed) into a 0-indexed list, order preserved."""
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
# LIBREOFFICE HELPER (Office -> PDF, full layout fidelity)
# ===========================================================================

def _libreoffice_convert(src_path, out_dir, out_fmt="pdf"):
    lo = shutil.which("libreoffice") or shutil.which("soffice")
    if not lo:
        raise RuntimeError(
            "LibreOffice is not installed on this server. "
            "Add 'libreoffice' to nixpacks.toml aptPkgs and redeploy."
        )
    profile = os.path.join(out_dir, f"lo_{uuid.uuid4().hex}")
    os.makedirs(profile, exist_ok=True)
    cmd = [
        lo, "--headless", "--norestore", "--nofirststartwizard", "--nolockcheck",
        f"-env:UserInstallation=file://{profile}",
        "--convert-to", out_fmt, "--outdir", out_dir, src_path,
    ]
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
    """Save an uploaded file to disk and convert it to PDF with LibreOffice."""
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


# ===========================================================================
# DOCX BUILD HELPERS (used by PDF -> Word)
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
    """Find genuine multi-row/multi-column tables (skip 1-col false positives)."""
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
    """Re-extract a table's rows via span-level clustering (more reliable
    for styled/borderless tables than the raw find_tables() output), plus
    detect the header fill colour and header text colour from the page.
    Clustering happens on individual text SPANS rather than whole text
    blocks, since PyMuPDF often merges an entire table row's cells into
    one block when they sit on the same line."""
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
                        run.font.color.rgb = RGBColor((htc >> 16) & 0xFF, (htc >> 8) & 0xFF, htc & 0xFF)
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
            thumbs.append({
                "page": i + 1, "width": pix.width, "height": pix.height,
                "dataUrl": f"data:image/png;base64,{b64}",
            })
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
            for i in range(len(doc)):
                s = fitz.open()
                s.insert_pdf(doc, from_page=i, to_page=i)
                zf.writestr(f"{base}_page_{i + 1}.pdf", s.tobytes())
                s.close()
    buf.seek(0)
    return send_bytes(buf.getvalue(), f"{base}_split.zip", "application/zip")


@app.route("/api/rotate", methods=["POST"])
def rotate_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to rotate."}), 400
    try:
        angle = int(request.form.get("angle", 90)) % 360
    except ValueError:
        angle = 90
    data = files[0].read()
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            page.set_rotation((page.rotation + angle) % 360)
        out = doc.tobytes()
    return send_bytes(out, "rotated.pdf", "application/pdf")


@app.route("/api/compress", methods=["POST"])
def compress_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file to compress."}), 400
    quality = request.form.get("quality", "recommended")
    jq = {"low": 40, "recommended": 60, "high": 80}.get(quality, 60)
    md = {"low": 800, "recommended": 1200, "high": 1600}.get(quality, 1200)
    data = files[0].read()
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        for page in doc:
            for img in page.get_images(full=True):
                xref = img[0]
                try:
                    bi = doc.extract_image(xref)
                    pil = Image.open(io.BytesIO(bi["image"])).convert("RGB")
                    w, h = pil.size
                    sc = min(1.0, md / max(w, h))
                    if sc < 1.0:
                        pil = pil.resize((max(1, int(w * sc)), max(1, int(h * sc))), Image.LANCZOS)
                    b = io.BytesIO()
                    pil.save(b, format="JPEG", quality=jq, optimize=True)
                    doc.update_stream(xref, b.getvalue())
                except Exception:
                    continue
        out = doc.tobytes(deflate=True, garbage=4)
    finally:
        doc.close()
    return send_bytes(out, "compressed.pdf", "application/pdf")


# ===========================================================================
# IMAGES <-> PDF
# ===========================================================================

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
    return send_bytes("\n\n".join(parts).encode("utf-8"), f"{base}.txt", "text/plain; charset=utf-8")


# ===========================================================================
# PDF -> WORD
# ===========================================================================

@app.route("/api/pdf-to-word", methods=["POST"])
def pdf_to_word():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"

    doc = Document()
    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
    page_w_cm = 17.0

    with fitz.open(stream=data, filetype="pdf") as pdf:
        for pg_idx, page in enumerate(pdf):
            if pg_idx > 0:
                doc.add_page_break()
            page_rect = page.rect
            med = _median_size(page)
            trects = sorted(_detect_tables(page), key=lambda r: r.y0)

            def in_table(y, _trects=trects):
                return any(tr.y0 - 4 <= y <= tr.y1 + 4 for tr in _trects)

            items = []
            try:
                raw = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
                for b in sorted(raw["blocks"], key=lambda b: b["bbox"][1]):
                    if b.get("type") == 1:
                        try:
                            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), clip=fitz.Rect(b["bbox"]))
                            items.append({"type": "image", "data": pix.tobytes("png"), "y": b["bbox"][1]})
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
                        items.append({
                            "type": "text", "text": text, "size": size,
                            "bold": bool(flags & 16), "italic": bool(flags & 2),
                            "rtl": _is_rtl(text), "y": line["bbox"][1],
                        })
            except Exception:
                pass

            has_text = any(it["type"] == "text" for it in items)
            if not has_text and not trects:
                _img_para(doc, _clip(page, page_rect, zoom=2.0), page_w_cm)
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
                        _build_table(doc, rows, hf, htc, usable_cm=page_w_cm)
                    else:
                        _img_para(doc, _clip(page, tr), page_w_cm)
                    sp2 = doc.add_paragraph()
                    sp2.paragraph_format.space_before = Pt(0)
                    sp2.paragraph_format.space_after = Pt(4)

            for item in items:
                flush_to(item["y"])
                if in_table(item["y"]):
                    continue
                if item["type"] == "image":
                    _img_para(doc, item["data"], page_w_cm)
                    continue
                text, size = item["text"], item["size"]
                bold, italic, rtl = item["bold"], item["italic"], item["rtl"]
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
            flush_to(page_rect.y1 + 99999)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_bytes(
        buf.getvalue(), f"{base}.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


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
                width=prs.slide_width, height=prs.slide_height,
            )
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    return send_bytes(
        buf.getvalue(), f"{base}.pptx",
        "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )


# ===========================================================================
# PDF -> EXCEL
# ===========================================================================

@app.route("/api/pdf-to-excel", methods=["POST"])
def pdf_to_excel():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "spreadsheet"

    wb = Workbook()
    wb.remove(wb.active)
    hdr_fill = PatternFill("solid", fgColor="2F5496")
    hdr_font = Font(bold=True, color="FFFFFF", size=10)
    body_font = Font(size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="B0C4D8")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

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
                        start_row = rc
                        for ri, row in enumerate(raw):
                            row = (row + [None] * (ncols - len(row)))[:ncols]
                            for ci, val in enumerate(row, start=1):
                                cell = ws.cell(row=rc, column=ci, value=val if val is not None else "")
                                cell.border = border
                                cell.font = hdr_font if ri == 0 else body_font
                                cell.alignment = center if ri == 0 else left
                                if ri == 0:
                                    cell.fill = hdr_fill
                            rc += 1
                        rc += 1
                        for ci in range(1, ncols + 1):
                            cl = get_column_letter(ci)
                            ml = max(
                                (len(str(ws.cell(row=r, column=ci).value or ""))
                                 for r in range(start_row, rc - 1)),
                                default=8,
                            )
                            ws.column_dimensions[cl].width = min(max(ml + 2, 8), 50)
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
                                    spans.append({"text": t, "x0": span["bbox"][0], "y0": span["bbox"][1]})
                    if spans:
                        xs = sorted(set(round(s["x0"]) for s in spans))
                        gap = 25
                        clusters = [xs[0]]
                        col_starts = [xs[0]]
                        for x in xs[1:]:
                            if x - col_starts[-1] > gap:
                                col_starts.append(x)

                        def col_of(x0, _starts=col_starts):
                            return min(range(len(_starts)), key=lambda i: abs(x0 - _starts[i])) + 1

                        rows_map = defaultdict(list)
                        for s in spans:
                            rows_map[round(s["y0"] / 4) * 4].append(s)
                        for yk in sorted(rows_map.keys()):
                            for s in rows_map[yk]:
                                ci = col_of(s["x0"])
                                existing = ws.cell(row=rc, column=ci).value or ""
                                ws.cell(row=rc, column=ci, value=(existing + " " if existing else "") + s["text"])
                            rc += 1
                    else:
                        for line in page.get_text().split("\n"):
                            if line.strip():
                                ws.cell(row=rc, column=1, value=line)
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
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ===========================================================================
# PAGE OPERATIONS
# ===========================================================================

@app.route("/api/remove-pages", methods=["POST"])
def remove_pages():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    spec = request.form.get("pages", "")
    if not spec.strip():
        return jsonify({"error": "Specify which pages to remove."}), 400
    data = files[0].read()
    with fitz.open(stream=data, filetype="pdf") as doc:
        to_rm = sorted(set(parse_page_spec(spec, len(doc))), reverse=True)
        if not to_rm:
            return jsonify({"error": "No valid page numbers given."}), 400
        for idx in to_rm:
            doc.delete_page(idx)
        if len(doc) == 0:
            return jsonify({"error": "Cannot remove all pages from the document."}), 400
        out = doc.tobytes()
    return send_bytes(out, "pages_removed.pdf", "application/pdf")


@app.route("/api/extract-pages", methods=["POST"])
def extract_pages():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    spec = request.form.get("pages", "")
    if not spec.strip():
        return jsonify({"error": "Specify which pages to keep."}), 400
    data = files[0].read()
    with fitz.open(stream=data, filetype="pdf") as doc:
        order = parse_page_spec(spec, len(doc))
        if not order:
            return jsonify({"error": "No valid page numbers given."}), 400
        new = fitz.open()
        for idx in order:
            new.insert_pdf(doc, from_page=idx, to_page=idx)
        out = new.tobytes()
        new.close()
    return send_bytes(out, "organized.pdf", "application/pdf")


# ===========================================================================
# EDIT TOOLS
# ===========================================================================

@app.route("/api/add-page-numbers", methods=["POST"])
def add_page_numbers():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    try:
        start = int(request.form.get("start", 1))
    except ValueError:
        start = 1
    pos = request.form.get("position", "bottom-center")
    data = files[0].read()
    with fitz.open(stream=data, filetype="pdf") as doc:
        for i, page in enumerate(doc):
            label = str(start + i)
            r = page.rect
            m = 24
            pts = {
                "bottom-center": fitz.Point(r.width / 2 - 8, r.height - m),
                "bottom-left": fitz.Point(m, r.height - m),
                "bottom-right": fitz.Point(r.width - m - 20, r.height - m),
                "top-center": fitz.Point(r.width / 2 - 8, m),
            }
            page.insert_text(pts.get(pos, pts["bottom-center"]), label, fontsize=11, color=(0, 0, 0))
        out = doc.tobytes()
    return send_bytes(out, "numbered.pdf", "application/pdf")


@app.route("/api/watermark", methods=["POST"])
def watermark_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    text = (request.form.get("text", "CONFIDENTIAL") or "CONFIDENTIAL").strip()
    try:
        opacity = max(0.05, min(float(request.form.get("opacity", 0.3)), 1.0))
    except ValueError:
        opacity = 0.3
    data = files[0].read()
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            r = page.rect
            pt = fitz.Point(r.width * 0.15, r.height * 0.55)
            mat = fitz.Matrix(1, 1).prerotate(45)
            page.insert_text(
                pt, text,
                fontsize=max(24, int(r.width / 12)),
                color=(0.6, 0.6, 0.6), fill_opacity=opacity, overlay=True,
                morph=(pt, mat),
            )
        out = doc.tobytes()
    return send_bytes(out, "watermarked.pdf", "application/pdf")


@app.route("/api/protect", methods=["POST"])
def protect_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    pw = (request.form.get("password", "") or "").strip()
    if not pw:
        return jsonify({"error": "Enter a password."}), 400
    data = files[0].read()
    with fitz.open(stream=data, filetype="pdf") as doc:
        out = doc.tobytes(
            encryption=fitz.PDF_ENCRYPT_AES_256, owner_pw=pw, user_pw=pw,
            permissions=int(fitz.PDF_PERM_PRINT | fitz.PDF_PERM_COPY | fitz.PDF_PERM_ANNOTATE),
        )
    return send_bytes(out, "protected.pdf", "application/pdf")


@app.route("/api/unlock", methods=["POST"])
def unlock_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    pw = (request.form.get("password", "") or "").strip()
    data = files[0].read()
    doc = fitz.open(stream=data, filetype="pdf")
    try:
        if doc.needs_pass:
            if not pw:
                return jsonify({"error": "This PDF is password-protected — enter the password."}), 400
            if not doc.authenticate(pw):
                return jsonify({"error": "Incorrect password."}), 400
        out = doc.tobytes()
    finally:
        doc.close()
    return send_bytes(out, "unlocked.pdf", "application/pdf")


@app.route("/api/crop", methods=["POST"])
def crop_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    try:
        margin = max(0, float(request.form.get("margin", 36)))
    except ValueError:
        margin = 36
    data = files[0].read()
    with fitz.open(stream=data, filetype="pdf") as doc:
        for page in doc:
            r = page.rect
            nr = fitz.Rect(r.x0 + margin, r.y0 + margin, r.x1 - margin, r.y1 - margin)
            if nr.width > 10 and nr.height > 10:
                page.set_cropbox(nr)
        out = doc.tobytes()
    return send_bytes(out, "cropped.pdf", "application/pdf")


@app.route("/api/redact", methods=["POST"])
def redact_pdf():
    files = get_uploaded_files()
    if not files:
        return jsonify({"error": "Upload a PDF file."}), 400
    terms_raw = (request.form.get("terms", "") or "").strip()
    if not terms_raw:
        return jsonify({"error": "Enter at least one word or phrase to redact."}), 400
    terms = [t.strip() for t in terms_raw.replace(",", "\n").splitlines() if t.strip()]
    src = files[0]
    data = src.read()
    base = os.path.splitext(secure_filename(src.filename))[0] or "document"
    with fitz.open(stream=data, filetype="pdf") as doc:
        total = 0
        for page in doc:
            for term in terms:
                hits = page.search_for(term, quads=True)
                total += len(hits)
                for quad in hits:
                    page.add_redact_annot(quad, fill=(0, 0, 0))
            page.apply_redactions()
        if total == 0:
            return jsonify({"error": "No matching text was found to redact."}), 400
        out = doc.tobytes(deflate=True, garbage=4)
    return send_bytes(out, f"{base}_redacted.pdf", "application/pdf")


@app.route("/api/add-signature", methods=["POST"])
def add_signature():
    pdf_file = request.files.get("file") or (request.files.getlist("files") or [None])[0]
    sig_file = request.files.get("signature")
    if not pdf_file or not pdf_file.filename:
        return jsonify({"error": "Upload a PDF file."}), 400
    if not sig_file or not sig_file.filename:
        return jsonify({"error": "Upload a signature image (PNG or JPG)."}), 400
    position = request.form.get("position", "bottom-right")
    page_target = request.form.get("page_target", "last")
    try:
        scale = max(0.05, min(float(request.form.get("scale", 0.20)), 0.5))
    except ValueError:
        scale = 0.20

    pdf_data = pdf_file.read()
    sig_img = Image.open(io.BytesIO(sig_file.read())).convert("RGBA")
    sig_buf = io.BytesIO()
    sig_img.save(sig_buf, format="PNG")
    sig_png = sig_buf.getvalue()
    base = os.path.splitext(secure_filename(pdf_file.filename))[0] or "document"

    with fitz.open(stream=pdf_data, filetype="pdf") as doc:
        n = len(doc)
        if page_target == "all":
            targets = list(range(n))
        elif page_target == "first":
            targets = [0]
        elif page_target == "last":
            targets = [n - 1]
        else:
            try:
                targets = [max(0, min(int(page_target) - 1, n - 1))]
            except ValueError:
                targets = [n - 1]

        for i in targets:
            page = doc[i]
            pw, ph = page.rect.width, page.rect.height
            sig_w = pw * scale
            ar = (sig_img.height / sig_img.width) if sig_img.width else 1
            sig_h = sig_w * ar
            margin = 20
            origins = {
                "bottom-right": (pw - sig_w - margin, ph - sig_h - margin),
                "bottom-left": (margin, ph - sig_h - margin),
                "bottom-center": ((pw - sig_w) / 2, ph - sig_h - margin),
                "top-right": (pw - sig_w - margin, margin),
                "top-left": (margin, margin),
            }
            x0, y0 = origins.get(position, origins["bottom-right"])
            page.insert_image(fitz.Rect(x0, y0, x0 + sig_w, y0 + sig_h), stream=sig_png)

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
# OFFICE -> PDF  (LibreOffice: full layout / table / image fidelity)
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
    out_bytes, err = _lo_to_pdf(src, ".xlsx", base)
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
                flowables.append(Paragraph(chunk.replace("&", "&amp;").replace("<", "&lt;") + ".", styles["Normal"]))
                flowables.append(Spacer(1, 6))
        if not flowables:
            flowables = [Paragraph("(Empty document)", styles["Normal"])]
        pdf_doc.build(flowables)
        out_bytes = buf.getvalue()

    return send_bytes(out_bytes, f"{base}.pdf", "application/pdf")


# ===========================================================================
# IMAGE OCR
# ===========================================================================

def _ocr(pil_img, lang="eng"):
    if not shutil.which("tesseract"):
        raise RuntimeError("Tesseract OCR is not installed. Add tesseract-ocr to nixpacks.toml.")
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
    fmt = request.form.get("output_format", "txt")
    lang = request.form.get("lang", "eng")
    if not re.match(r"^[a-zA-Z+]{2,20}$", lang):
        lang = "eng"

    page_texts = []
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
            page_texts.append((name, _ocr(img, lang=lang)))
        except RuntimeError as e:
            return jsonify({"error": str(e)}), 503
        except Exception as e:
            page_texts.append((name, f"[Could not read this image: {e}]"))

    if not page_texts:
        return jsonify({"error": "No images could be processed."}), 400
    base = page_texts[0][0] or "ocr_result"

    if fmt == "txt":
        parts = []
        for name, text in page_texts:
            if len(page_texts) > 1:
                parts.append(f"=== {name} ===")
            parts.append(text.strip())
            parts.append("")
        return send_bytes("\n".join(parts).strip().encode("utf-8"), f"{base}_ocr.txt", "text/plain; charset=utf-8")

    doc = Document()
    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.left_margin = Cm(2.0)
        section.right_margin = Cm(2.0)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
    for idx, (name, text) in enumerate(page_texts):
        if idx > 0:
            doc.add_page_break()
        if len(page_texts) > 1:
            doc.add_paragraph(style="Heading 1").add_run(name)
        for line in text.splitlines():
            p = doc.add_paragraph()
            run = p.add_run(line)
            run.font.size = Pt(10)
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return send_bytes(
        buf.getvalue(), f"{base}_ocr.docx",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


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
