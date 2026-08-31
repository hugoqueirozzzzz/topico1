FROM python:3.12-slim

# LibreOffice é necessário para a conversão PPTX -> PDF (soffice --headless)
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-impress \
    fonts-dejavu \
    fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p tmp_uploads tmp_outputs

ENV PORT=5000
EXPOSE 5000

# 1 worker + várias threads: cada geração usa o LibreOffice via subprocess,
# então não precisamos de muitos workers concorrentes num plano gratuito.
CMD ["gunicorn", "-w", "1", "--threads", "4", "-b", "0.0.0.0:5000", "--timeout", "180", "app:app"]
