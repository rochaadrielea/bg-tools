FROM python:3.12-slim

# proxy is needed for pip at BUILD time only, never at run time
ARG HTTP_PROXY=""
ARG HTTPS_PROXY=""
ARG NO_PROXY=""

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY *.py /app/
COPY migrations /app/migrations

# stamped at build time so /health tells you exactly which image is running
ARG APP_VERSION=dev
ENV APP_VERSION=$APP_VERSION
ENV WEB_DIR=/app/web

EXPOSE 8000
HEALTHCHECK --interval=20s --timeout=4s --retries=5 \
  CMD python -c "import urllib.request,sys; sys.exit(0 if urllib.request.urlopen('http://127.0.0.1:8000/health',timeout=3).status==200 else 1)"
CMD ["uvicorn","main:app","--host","0.0.0.0","--port","8000"]