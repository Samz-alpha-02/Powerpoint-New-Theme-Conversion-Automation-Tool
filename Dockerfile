# ── Build stage ───────────────────────────────────────────────────────────────
FROM python:3.11-slim

# System deps for Pillow (JPEG/PNG/WebP support)
RUN apt-get update && apt-get install -y --no-install-recommends \
        libjpeg-dev \
        libpng-dev \
        libwebp-dev \
        libfreetype6-dev \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies first (layer cache)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy application code
COPY app.py .

<<<<<<< HEAD
# Render injects PORT; Streamlit reads it via --server.port
EXPOSE 8501

# Disable Streamlit's browser auto-open and file watcher (not needed in prod)
ENV STREAMLIT_SERVER_HEADLESS=true \
    STREAMLIT_SERVER_FILE_WATCHER_TYPE=none \
    STREAMLIT_BROWSER_GATHER_USAGE_STATS=false
=======
# ── Expose the port Streamlit listens on ─────────────────────────────────────
# Render overrides this via the $PORT env variable; the CMD below honours it.
EXPOSE ${PORT}
>>>>>>> 4aab632399d5d5ebe9ba9daab7f562d5b81a0412

CMD ["sh", "-c", "streamlit run app.py --server.port=${PORT:-8501} --server.address=0.0.0.0"]
