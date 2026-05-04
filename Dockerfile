FROM python:3.11-slim

# wget/curl necessarios para playwright install --with-deps
RUN apt-get update && apt-get install -y --no-install-recommends wget curl \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# instala Chromium + todas as dependencias de sistema do Playwright
RUN playwright install --with-deps chromium

COPY . .

ENV PYTHONIOENCODING=utf-8
ENV PYTHONUTF8=1

# PORT e injetado automaticamente pelo Render
CMD ["sh", "-c", "uvicorn server:app --host 0.0.0.0 --port ${PORT:-8080}"]
