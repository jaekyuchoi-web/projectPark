"""AI 보조 날짜 파서 테스트 (네트워크 미사용 — 주입식 fake client)."""

from __future__ import annotations

from app.config import Settings
from app.domain.period_extract import (
    _openai_parse_dates,
    _parse_date_json,
    build_ai_date_parser,
)


# ── 순수 JSON 변환 ──────────────────────────────────────────────────
def test_parse_date_json_basic():
    text = '{"이천이십육년 삼월": "2026-03", "도저히모름": null}'
    out = _parse_date_json(text, ["이천이십육년 삼월", "도저히모름"])
    assert out["이천이십육년 삼월"] == (2026, 3)
    assert out["도저히모름"] is None


def test_parse_date_json_embedded_text():
    # 모델이 앞뒤로 텍스트를 붙여도 JSON 블록을 추출
    text = '결과는 다음과 같습니다: {"2026.3": "2026-03"} 이상입니다.'
    out = _parse_date_json(text, ["2026.3"])
    assert out["2026.3"] == (2026, 3)


def test_parse_date_json_garbage_returns_empty():
    assert _parse_date_json("죄송합니다 JSON 없음", ["x"]) == {}


# ── 주입식 fake client 로 전체 경로 ─────────────────────────────────
class _FakeMessage:
    def __init__(self, content):
        self.content = content


class _FakeChoice:
    def __init__(self, content):
        self.message = _FakeMessage(content)


class _FakeResp:
    def __init__(self, content):
        self.choices = [_FakeChoice(content)]


class _FakeCompletions:
    def __init__(self, content):
        self._content = content
        self.calls = 0

    def create(self, **kwargs):
        self.calls += 1
        return _FakeResp(self._content)


class _FakeChat:
    def __init__(self, content):
        self.completions = _FakeCompletions(content)


class _FakeClient:
    def __init__(self, content):
        self.chat = _FakeChat(content)


_SETTINGS = Settings(openai_api_key="sk-test", openai_model="gpt-4.1-mini")


def test_openai_parse_dates_with_fake_client():
    fake = _FakeClient('{"이천이십육년 삼월": "2026-03", "엉터리": null}')
    out = _openai_parse_dates(["이천이십육년 삼월", "엉터리"], _SETTINGS, client=fake)
    assert out["이천이십육년 삼월"] == (2026, 3)
    assert out["엉터리"] is None


# ── 팩토리: 키 유무 ─────────────────────────────────────────────────
def test_build_ai_date_parser_no_key_returns_none():
    assert build_ai_date_parser(Settings(openai_api_key=None, openai_model="m")) is None


def test_build_ai_date_parser_with_key_returns_callable():
    parser = build_ai_date_parser(_SETTINGS)
    assert callable(parser)
