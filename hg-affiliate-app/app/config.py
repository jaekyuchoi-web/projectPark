"""설정 로딩 모듈.

secret.env 에서 OpenAI 키/모델을 로드한다. 키 값 자체는 절대 화면·로그에
노출하지 않으며, 외부에는 "설정 여부(boolean)"만 제공한다.
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv

# 프로젝트 루트 (이 파일 기준 한 단계 위)
BASE_DIR = Path(__file__).resolve().parent.parent
SECRET_ENV_PATH = BASE_DIR / "secret.env"

# 세션(업로드/출력) 작업 루트
SESSION_ROOT = BASE_DIR / ".sessions"
SESSION_TTL_SECONDS = 60 * 60  # 1시간 후 만료 삭제

# 업로드 슬롯 정의 (e: 입력 5종)
# 용어 정의 — "전기" = 전년도 온기(연간 결산 = 4분기).
#   같은 해 직전분기 자료는 참고할 수 있으나 전기가 아니다.
#   예) 당기 2026년 1~4분기 → 전기는 모두 2025년 온기/4Q.
SLOT_KEYS = [
    "prev_balance",      # 1. 전기 이월 소스 = 전년도 온기 계정별 잔액 명세서
    "current_ledger",    # 2. 당기 거래 내역 소스 (당기 계정별 원장, 회계연도 누적/YTD)
    "current_balance",   # 3. 당기 채권채무 잔액 검증용 소스 (당기말 잔액 명세서)
    "name_pivot",        # 4. 특관자 상호 정리 (정규화 시드)
    "prev_statement",    # 5. 전기 특관자 명세서 (양식/수식 템플릿)
]

SLOT_LABELS = {
    "prev_balance": "전기 이월 소스 (전년도 온기 계정별 잔액 명세서)",
    "current_ledger": "당기 거래 내역 소스 (당기 계정별 원장)",
    "current_balance": "당기 채권채무 잔액 검증용 소스 (당기말 잔액 명세서)",
    "name_pivot": "특관자 상호 정리 (정규화 기준/시드)",
    "prev_statement": "전기 특관자 명세서 (양식/수식 템플릿)",
}

# 허용 확장자 (모든 엑셀 형식)
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb", ".csv"}

# 대상 회사 (도메인 §4)
TARGET_COMPANY_CANONICAL = "에이치엘비글로벌㈜"
TARGET_COMPANY_ALIASES = {"HLB글로벌", "HLB글로벌(주)", "에이치엘비글로벌(주)", "에이치엘비글로벌㈜"}


@dataclass(frozen=True)
class Settings:
    """런타임 설정. api_key 는 내부에서만 사용하고 외부 노출 금지."""

    openai_api_key: str | None
    openai_model: str
    # web.app( Hosting 60초 제한) 접속 시 API 를 Cloud Run 으로 직접 호출
    public_api_url: str | None = None

    @property
    def has_openai_key(self) -> bool:
        return bool(self.openai_api_key and self.openai_api_key.strip())


def load_settings() -> Settings:
    """secret.env 를 읽어 Settings 를 만든다. 매 요청마다 호출해도 안전하다."""
    if SECRET_ENV_PATH.exists():
        load_dotenv(SECRET_ENV_PATH, override=True)

    key = os.getenv("OPENAI_API_KEY") or None
    model = os.getenv("OPENAI_MODEL") or "gpt-4.1-mini"
    api_url = (os.getenv("PUBLIC_API_URL") or "").strip() or None
    return Settings(openai_api_key=key, openai_model=model, public_api_url=api_url)
