FROM python:3.11-slim

# Tools needed for features
RUN apt-get update && apt-get install -y --no-install-recommends \
    ghostscript \
    libreoffice \
    fonts-dejavu \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
EXPOSE 8501
CMD ["python","-m","streamlit","run","app.py","--server.port","8501","--server.address","0.0.0.0"]
RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    libgl1 libglib2.0-0 libsm6 libxext6 libxrender1 \
 && rm -rf /var/lib/apt/lists/*

