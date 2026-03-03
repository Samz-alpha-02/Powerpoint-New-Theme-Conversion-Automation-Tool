# ──────────────────────────────────────────────────────────────────────────────
# PowerPoint Rebranding Tool – Dockerfile
# Target platform: Render (https://render.com)
#
# Build:  docker build -t pptx-rebrand .
# Run:    docker run -p 8501:8501 pptx-rebrand
# ──────────────────────────────────────────────────────────────────────────────

# ── Stage 1: slim Python 3.11 base ───────────────────────────────────────────
FROM python:3.11-slim

# Keep Python output unbuffered (shows logs immediately in Render dashboard)
ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    # Streamlit reads PORT from environment; Render injects this automatically
    PORT=10000

# ── System dependencies ───────────────────────────────────────────────────────
# libxml2 / libxslt are required by lxml at runtime on slim images
RUN apt-get update && apt-get install -y --no-install-recommends \
        libxml2 \
        libxslt1.1 \
    && rm -rf /var/lib/apt/lists/*

# ── Working directory ─────────────────────────────────────────────────────────
WORKDIR /app

# ── Python dependencies (cached layer – only re-runs when requirements change) ──
COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
 && pip install --no-cache-dir -r requirements.txt

# ── Application source ────────────────────────────────────────────────────────
COPY app.py              ./app.py
COPY template_applier.py ./template_applier.py

# Copy Streamlit server config (headless, disables CORS/XSRF for container use)
COPY .streamlit/config.toml ./.streamlit/config.toml

# ── Expose the port Streamlit listens on ─────────────────────────────────────
# Render overrides this via the $PORT env variable; the CMD below honours it.
EXPOSE ${PORT}

# ── Healthcheck ───────────────────────────────────────────────────────────────
HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:${PORT}/_stcore/health')" \
    || exit 1

# ── Start command ─────────────────────────────────────────────────────────────
# $PORT is provided by Render at runtime; fallback to 10000 for local testing.
CMD streamlit run app.py \
        --server.port=${PORT} \
        --server.address=0.0.0.0 \
        --server.headless=true \
        --server.enableCORS=false \
        --server.enableXsrfProtection=false
