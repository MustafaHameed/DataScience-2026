#!/usr/bin/env bash
# ch09_build.sh -- the build cache, measured, then a sound image.
set -euo pipefail
mkdir -p ~/vsapp && cd ~/vsapp

# --- 1. A small application with real dependencies. --------------------
cat > requirements.txt <<'REQ'
flask==3.0.3
requests==2.32.3
REQ
cat > main.py <<'PY'
from flask import Flask
app = Flask(__name__)

@app.route("/")
def index():
    return "virtual systems and services\n"
PY

# --- 2. The wrong order: source copied before dependencies. -----------
cat > Dockerfile.wrong <<'DF'
FROM python:3.12-slim
WORKDIR /app
COPY . /app
RUN pip install --no-cache-dir -r requirements.txt
CMD ["python", "-m", "flask", "run", "--host=0.0.0.0"]
DF
docker build -q -f Dockerfile.wrong -t vsapp:wrong . > /dev/null

# --- 3. The right order: dependency list first, then source. ----------
cat > Dockerfile.right <<'DF'
FROM python:3.12-slim
WORKDIR /app
COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt
COPY . /app
CMD ["python", "-m", "flask", "run", "--host=0.0.0.0"]
DF
docker build -q -f Dockerfile.right -t vsapp:right . > /dev/null

# --- 4. Change one character of source and time both rebuilds. --------
sed -i 's/services/services!/' main.py
echo "=== wrong order ===" ; time docker build -q -f Dockerfile.wrong \
  -t vsapp:wrong . > /dev/null
echo "=== right order ===" ; time docker build -q -f Dockerfile.right \
  -t vsapp:right . > /dev/null

# --- 5. Look at the layers, and at what is shared. --------------------
docker history vsapp:right --format '{{.Size}}\t{{.CreatedBy}}' | head -8
docker image inspect vsapp:right \
  --format '{{json .RootFS.Layers}}' | tr ',' '\n'
# Both images share the python:3.12-slim layers: content addressing at work.

# --- 6. Multi-stage and non-root. -------------------------------------
cat > Dockerfile.good <<'DF'
FROM python:3.12-slim AS build
WORKDIR /app
COPY requirements.txt /app/
RUN pip install --no-cache-dir --target=/deps -r requirements.txt

FROM python:3.12-slim
RUN useradd --create-home --uid 10001 app
WORKDIR /app
COPY --from=build /deps /deps
COPY --chown=app:app . /app
ENV PYTHONPATH=/deps
USER app
CMD ["python", "-m", "flask", "run", "--host=0.0.0.0"]
DF
docker build -q -f Dockerfile.good -t vsapp:good . > /dev/null
docker images vsapp --format '{{.Tag}}\t{{.Size}}'

# --- 7. Prove it is not running as root, and cannot write its root. ---
docker run --rm vsapp:good id
docker run --rm --read-only vsapp:good sh -c 'touch /nope' 2>&1 | tail -1

# --- 8. A secret deleted in a later layer is still in the image. ------
printf 'FROM alpine\nRUN echo hunter2 > /k && rm /k\n' > Dockerfile.leak
docker build -q -f Dockerfile.leak -t vsapp:leak . > /dev/null
docker save vsapp:leak | tar -xO 2>/dev/null | grep -ac hunter2 || \
  echo "search the layer tarballs: the secret is in one of them"

# --- 9. Tear down. ----------------------------------------------------
docker rmi -f vsapp:wrong vsapp:right vsapp:good vsapp:leak || true
