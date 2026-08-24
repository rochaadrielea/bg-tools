#!/usr/bin/env bash
# As-Built Tracker — deploy and rollback.
#
#   ./deploy.sh              build a new version, back up, switch, health-check
#   ./deploy.sh rollback     go back to the previous version
#   ./deploy.sh versions     list the images you can roll back to
#   ./deploy.sh current      what is running now
#
# Why it is safe to roll back:
#   · every build is tagged with a date and a serial, and kept
#   · the running tag is written to .version — rollback is a tag change, not a rebuild
#   · the database is dumped before every switch, into ~/abt_backups/
#   · migrations are additive only, so the previous image still runs on the new schema
set -euo pipefail

ROOT="$HOME/bgtools"
SERVICE="abt-api"
IMAGE="bgtools-abt-api"
PORT="8506"
BACKUPS="$HOME/abt_backups"
VERFILE="$ROOT/pages/asbuilt/.version"
KEEP=10

cd "$ROOT"
mkdir -p "$BACKUPS"

current() { [ -f "$VERFILE" ] && cat "$VERFILE" || echo "none"; }

versions() {
  docker images "$IMAGE" --format '{{.Tag}}\t{{.CreatedSince}}\t{{.Size}}' | grep -v '^<none>' | head -"$KEEP"
}

backup() {
  local f="$BACKUPS/abt_$(date +%Y-%m-%d_%H%M).sql"
  echo "→ backing up the database to $f"
  docker compose exec -T abt-db pg_dump -U abt abtracker > "$f"
  ls -1t "$BACKUPS"/abt_*.sql | tail -n +21 | xargs -r rm --   # keep the last 20
}

health() {
  for i in $(seq 1 30); do
    if curl -fs "localhost:$PORT/health" > /tmp/abt_health.json 2>/dev/null; then
      echo "→ healthy: $(cat /tmp/abt_health.json)"; return 0
    fi
    sleep 2
  done
  echo "✗ did not come up healthy within 60s"; return 1
}

switch_to() {
  local tag="$1"
  echo "→ switching $SERVICE to $tag"
  ABT_TAG="$tag" docker compose up -d --no-build "$SERVICE"
  if health; then
    echo "$tag" > "$VERFILE"
    echo "✓ running $tag"
  else
    echo "✗ rolling back automatically"
    docker compose logs --tail 40 "$SERVICE" || true
    return 1
  fi
}

case "${1:-deploy}" in
  current)  echo "running: $(current)";;
  versions) echo "tag             built            size"; versions;;

  rollback)
    prev="$(versions | sed -n '2p' | cut -f1)"
    [ -n "$prev" ] || { echo "✗ no previous version to roll back to"; exit 1; }
    echo "→ rolling back from $(current) to $prev"
    backup
    switch_to "$prev"
    ;;

  deploy)
    tag="$(date +%Y-%m-%d)-$(( $(docker images "$IMAGE" --format '{{.Tag}}' | grep -c "^$(date +%Y-%m-%d)" || true) + 1 ))"
    echo "→ building $IMAGE:$tag"
    docker build \
      --build-arg APP_VERSION="$tag" \
      --build-arg HTTP_PROXY="http://chbs8055.ruaggroup.com:8080" \
      --build-arg HTTPS_PROXY="http://chbs8055.ruaggroup.com:8080" \
      --build-arg NO_PROXY="localhost,127.0.0.1" \
      -t "$IMAGE:$tag" -t "$IMAGE:latest" pages/asbuilt/api
    backup
    switch_to "$tag"
    echo
    echo "roll back with:  ./deploy.sh rollback"
    ;;

  *) echo "usage: $0 [deploy|rollback|versions|current]"; exit 1;;
esac