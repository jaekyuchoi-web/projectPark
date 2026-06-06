FROM python:3.12-slim

# LibreOffice(headless, 출력 재계산 검증용) + 폰트
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        libreoffice-calc \
        fonts-nanum \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app ./app
COPY secret.env.example ./secret.env.example

ENV PORT=8080
EXPOSE 8080

# Cloud Run 은 $PORT(기본 8080)로 수신. 로컬은 docker run -e PORT=8000 -p 8000:8000
CMD ["sh", "-c", "uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8080}"]
