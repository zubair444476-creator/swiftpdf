import os
import io
import zipfile
import uuid

from flask import Flask, request, send_file, jsonify, render_template
from werkzeug.utils import secure_filename

import fitz  # pymupdf
from PIL import Image

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
