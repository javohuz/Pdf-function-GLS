FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    XDG_CACHE_HOME=/tmp/.cache \
    PORT=8080

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        fontconfig \
        fonts-noto-cjk \
        libpango-1.0-0 \
        libpangocairo-1.0-0 \
        libpangoft2-1.0-0 \
        libharfbuzz0b \
        libharfbuzz-subset0 \
        libcairo2 \
        libgdk-pixbuf-2.0-0 \
        libjpeg62-turbo \
        libopenjp2-7 \
        shared-mime-info \
    && mkdir -p /tmp/.cache/fontconfig \
    && fc-cache -f \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir --upgrade pip \
    && pip install --no-cache-dir -r requirements.txt

COPY . .

CMD exec gunicorn --bind :$PORT --workers 1 --threads 8 --timeout 120 app:app
