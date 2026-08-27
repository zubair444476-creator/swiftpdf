FROM python:3.11-slim

# Install Tesseract OCR + language packs via apt (always works on Railway)
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-eng \
    tesseract-ocr-ara \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app files
COPY server.py .
COPY index.html .

EXPOSE 8080

CMD gunicorn server:app --bind 0.0.0.0:$PORT --timeout 120 --workers 2
