#!/usr/bin/env bash
# 개발용 핫리로드 실행.
# .venv 가 프로젝트 안에 있으면 watchfiles 가 의존성 변화를 감지해 무한 리로드가
# 발생하므로, app 만 감시하고 .venv/.sessions/__pycache__ 를 재귀적으로 제외한다.
#
# ※ 완전한 해결을 원하면 venv 를 프로젝트 밖에 두세요. 예:
#     python -m venv ~/.venvs/hg-affiliate && source ~/.venvs/hg-affiliate/bin/activate
#     pip install -r requirements.txt
#   이렇게 하면 자발적 리로드가 0 이 됩니다.
set -euo pipefail
cd "$(dirname "$0")"

PORT="${PORT:-8000}"

if [ -d ".venv" ]; then
  # shellcheck disable=SC1091
  source .venv/bin/activate
fi

exec uvicorn app.main:app --port "$PORT" \
  --reload \
  --reload-dir app \
  --reload-exclude '**/.venv/**' \
  --reload-exclude '**/.sessions/**' \
  --reload-exclude '**/__pycache__/**'
