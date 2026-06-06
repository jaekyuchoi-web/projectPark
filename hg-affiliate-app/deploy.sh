#!/usr/bin/env bash
# Cloud Run 배포 (로컬 Docker 빌드 → Artifact Registry 푸시 → 서비스 갱신)
set -euo pipefail
cd "$(dirname "$0")"

PROJECT="${GCP_PROJECT:-hg-affiliate}"
REGION="${GCP_REGION:-asia-southeast1}"
SERVICE="${GCP_SERVICE:-hg-affiliate-runtime}"
REPO="firebaseapphosting-images"
IMAGE="asia-southeast1-docker.pkg.dev/${PROJECT}/${REPO}/hg-affiliate"
TAG="v2-$(date +%Y%m%d-%H%M%S)"

if [[ ! -f secret.env ]]; then
  echo "오류: secret.env 가 없습니다. cp secret.env.example secret.env 후 키를 입력하세요."
  exit 1
fi
if [[ ! -f app/assets/statement_template.xlsx ]]; then
  echo "오류: app/assets/statement_template.xlsx 가 없습니다. README 를 참고해 정답 템플릿을 배치하세요."
  exit 1
fi

set -a
# shellcheck disable=SC1091
source secret.env
set +a

echo "==> Docker 빌드 (linux/amd64): ${IMAGE}:${TAG}"
docker build --platform linux/amd64 -t "${IMAGE}:${TAG}" .

echo "==> Artifact Registry 푸시"
docker push "${IMAGE}:${TAG}"

echo "==> Cloud Run 배포: ${SERVICE} (${REGION})"
gcloud run deploy "${SERVICE}" \
  --project "${PROJECT}" \
  --region "${REGION}" \
  --image "${IMAGE}:${TAG}" \
  --port 8080 \
  --memory 2Gi \
  --cpu 1 \
  --timeout 300 \
  --allow-unauthenticated \
  --set-env-vars "OPENAI_API_KEY=${OPENAI_API_KEY},OPENAI_MODEL=${OPENAI_MODEL:-gpt-4.1-mini}"

echo "==> 완료"
gcloud run services describe "${SERVICE}" --project "${PROJECT}" --region "${REGION}" \
  --format='value(status.url)'
