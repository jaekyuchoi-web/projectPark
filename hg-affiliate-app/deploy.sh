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

RUN_URL="$(gcloud run services describe "${SERVICE}" \
  --project "${PROJECT}" --region "${REGION}" \
  --format='value(status.url)' 2>/dev/null || true)"
if [[ -z "${RUN_URL}" ]]; then
  RUN_URL="https://${SERVICE}-${PROJECT}.asia-southeast1.run.app"
fi

echo "==> Cloud Run 배포: ${SERVICE} (${REGION})"
# 처리 시간/자원 제약 해제: 실제 분기 데이터(10MB+ 원장 + AI 정규화 +
# LibreOffice 재계산 2회)는 5분을 초과하므로 Cloud Run 한도를 최대치로 둔다.
#   --timeout 3600 : Cloud Run 요청 타임아웃 최대값(60분). 동기 요청의 상한.
#   --memory 16Gi / --cpu 4 : 대용량 .xls 를 문자열 DataFrame 으로 적재해도
#                             OOM 이 나지 않도록 넉넉히. 유휴 시 0으로 스케일되어
#                             평소 비용은 없고 처리 중에만 과금된다.
#   --max-instances 1 : 세션/출력 파일이 인스턴스 로컬 디스크(.sessions)에 있어
#                       다운로드 요청도 같은 인스턴스가 받아야 한다.
gcloud run deploy "${SERVICE}" \
  --project "${PROJECT}" \
  --region "${REGION}" \
  --image "${IMAGE}:${TAG}" \
  --port 8080 \
  --memory 16Gi \
  --cpu 4 \
  --cpu-boost \
  --timeout 3600 \
  --max-instances 1 \
  --allow-unauthenticated \
  --set-env-vars "OPENAI_API_KEY=${OPENAI_API_KEY},OPENAI_MODEL=${OPENAI_MODEL:-gpt-4.1-mini},PUBLIC_API_URL=${RUN_URL}"

echo "==> Firebase Hosting (https://hg-affiliate.web.app) 연결"
REPO_ROOT="$(cd "$(dirname "$0")/.." && pwd)"
if command -v firebase >/dev/null 2>&1; then
  (cd "${REPO_ROOT}" && firebase deploy --only hosting --project "${PROJECT}" --non-interactive)
else
  echo "경고: firebase CLI 가 없어 Hosting 배포를 건너뜁니다. npm i -g firebase-tools 후 재실행하세요."
fi

echo "==> 완료"
echo "Cloud Run: $(gcloud run services describe "${SERVICE}" --project "${PROJECT}" --region "${REGION}" --format='value(status.url)')"
echo "웹 주소: https://hg-affiliate.web.app"
