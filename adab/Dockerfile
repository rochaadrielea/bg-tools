# ADAB Compare web app - reproducible container
FROM python:3.11-slim
WORKDIR /app

# corporate proxy is passed in at build time by docker-compose (build.args).
# These ARGs let pip reach PyPI through the proxy during the build.
ARG HTTP_PROXY
ARG HTTPS_PROXY
ARG NO_PROXY

COPY requirements_web.txt .
RUN pip install --no-cache-dir -r requirements_web.txt

# copy the whole tool (adab_server.py, adab_batch_compare.py, matchcore/, batch_finder.py, feedback.py)
COPY . .

EXPOSE 8000
# stays up; one version everyone connects to
CMD ["uvicorn", "adab_server:app", "--host", "0.0.0.0", "--port", "8000"]
