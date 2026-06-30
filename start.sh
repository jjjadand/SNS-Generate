#!/usr/bin/env bash
set -euo pipefail

cd "$(dirname "$0")"
export HOST="${HOST:-0.0.0.0}"
export PORT="${PORT:-8005}"
export RELOAD="${RELOAD:-true}"

python3 -m uvicorn main:app --host "$HOST" --port "$PORT" --reload
