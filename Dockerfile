FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    HOST=0.0.0.0 \
    PORT=8080 \
    PDF_ENGINE=reportlab \
    PDF_PROFILE_DEFAULT=exact

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice \
    libreoffice-writer \
    fonts-dejavu \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# התקנת גופן David לתמיכה בעברית ב-LibreOffice
RUN mkdir -p /usr/share/fonts/truetype/david && \
    cp fonts/david.ttf fonts/davidbd.ttf /usr/share/fonts/truetype/david/ && \
    fc-cache -fv

EXPOSE 8080

CMD ["gunicorn", "--bind", "0.0.0.0:8080", "--workers", "1", "--threads", "4", "--timeout", "120", "app:app"]
