# Gevin Metal System — production image for Linux LAN servers
FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PORT=7861 \
    GRADIO_ANALYTICS_ENABLED=False \
    GRADIO_SERVER_NAME=0.0.0.0 \
    GRADIO_NODE_SERVER_NAME=0.0.0.0

WORKDIR /app

# System deps for some wheels / SSL
RUN apt-get update \
    && apt-get install -y --no-install-recommends ca-certificates \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY . .

# Runtime dirs (overridden by volumes in compose when present)
RUN mkdir -p data output/invoices output/reports logs templates

EXPOSE 7861

CMD ["python", "app.py"]
