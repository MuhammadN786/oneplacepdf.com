FROM python:3.11-slim

# Install Ghostscript for the Compress feature
RUN apt-get update && apt-get install -y --no-install-recommends \
    ghostscript \
 && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .
EXPOSE 8501
CMD ["python","-m","streamlit","run","app.py","--server.port","8501","--server.address","0.0.0.0"]
