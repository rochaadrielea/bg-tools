FROM python:3.12-slim

# The corporate proxy is needed for pip at BUILD time only, never at run time.
ARG HTTP_PROXY=""
ARG HTTPS_PROXY=""
ARG NO_PROXY=""

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY *.py schema.sql /app/
ENV WEB_DIR=/app/web

EXPOSE 8000
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]