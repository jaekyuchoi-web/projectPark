"""/api/run 당기 선택 검증 (FastAPI TestClient)."""

from __future__ import annotations

from fastapi.testclient import TestClient

from app.main import app

client = TestClient(app)


def test_api_run_rejects_missing_period():
    sid = client.post("/api/session").json()["sid"]
    r = client.post("/api/run", json={"sid": sid})
    assert r.status_code == 400
    assert "분기" in r.json()["message"]  # 당기 기간 미선택 사유


def test_api_run_rejects_bad_quarter():
    sid = client.post("/api/session").json()["sid"]
    r = client.post("/api/run", json={"sid": sid, "year": 2026, "quarter": 9})
    assert r.status_code == 400
    assert "분기" in r.json()["message"]


def test_index_renders_period_inputs():
    r = client.get("/")
    assert r.status_code == 200
    assert "period-year" in r.text
    assert "period-quarter" in r.text
