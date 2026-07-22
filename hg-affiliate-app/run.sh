#!/usr/bin/env bash
# 운영/일반 실행 (리로드 없음 → 안정적인 단일 프로세스).
# 회계 담당자가 도구를 사용할 때는 이 스크립트를 쓰세요.
set -euo pipefail
cd "$(dirname "$0")"

PORT="${PORT:-8000}"

PY="python3"
if [ -x ".venv/bin/python" ]; then
  PY=".venv/bin/python"
fi

exec "$PY" -m uvicorn app.main:app --host 0.0.0.0 --port "$PORT"
