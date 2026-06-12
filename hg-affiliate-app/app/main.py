"""FastAPI 진입점 및 라우트.

UX 흐름(§2): 설정 안내 → 업로드(검증) → 분류 확인 → 실행 → 다운로드.
입력/출력 '내용'은 어떤 응답에도 담지 않는다. (파일명/종류/뱃지/건수만)
"""

from __future__ import annotations

import uuid
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

from .classifier import classify
from .config import SLOT_KEYS, SLOT_LABELS, load_settings
from .domain.period_extract import parse_period
from .pipeline import run_pipeline
from .session import store
from .validation import validate_file_with_sheets

BASE_DIR = Path(__file__).resolve().parent
app = FastAPI(title="HLB글로벌 특관자 명세서 생성기", version="2.0.0")

# web.app( Hosting ) 에서 Cloud Run API 직접 호출 허용
app.add_middleware(
    CORSMiddleware,
    allow_origins=[
        "https://hg-affiliate.web.app",
        "https://hg-affiliate.firebaseapp.com",
        "http://localhost:8000",
        "http://127.0.0.1:8000",
    ],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.mount("/static", StaticFiles(directory=BASE_DIR / "static"), name="static")
templates = Jinja2Templates(directory=str(BASE_DIR / "templates"))


@app.get("/", response_class=HTMLResponse)
def index(request: Request):
    settings = load_settings()
    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "has_key": settings.has_openai_key,
            "model": settings.openai_model,
            "api_base": settings.public_api_url or "",
            "slots": [{"key": k, "label": SLOT_LABELS[k]} for k in SLOT_KEYS],
        },
    )


@app.get("/api/config")
def api_config():
    settings = load_settings()
    return {"has_key": settings.has_openai_key, "model": settings.openai_model}


@app.post("/api/session")
def api_session():
    sess = store.create()
    sess.meta.setdefault("files", {})
    sess.meta.setdefault("assign", {})
    return {"sid": sess.sid}


@app.post("/api/upload")
async def api_upload(sid: str = Form(...), file: UploadFile = File(...)):
    sess = store.get(sid)
    if sess is None:
        raise HTTPException(404, "세션을 찾을 수 없습니다. 새로고침 후 다시 시도하세요.")

    original = file.filename or "uploaded"
    file_id = uuid.uuid4().hex
    dest = sess.upload_dir / f"{file_id}{Path(original).suffix.lower()}"
    content = await file.read()
    dest.write_bytes(content)

    # 선행 검증 (a) — 워크북은 1회만 읽고 분류에 재사용
    v, sheets = validate_file_with_sheets(dest, original)
    suggested = None
    if v.ok:
        suggested = classify(dest, original, sheets=sheets).slot
    else:
        dest.unlink(missing_ok=True)

    info = {
        "file_id": file_id,
        "filename": original,
        "ext": v.ext,
        "valid": v.ok,
        "reasons": v.reasons,
        "suggested_slot": suggested,
        "path": str(dest) if v.ok else "",
    }
    sess.meta["files"][file_id] = info

    # 자동 분류로 슬롯 자동 배정(빈 슬롯일 때만)
    if v.ok and suggested and suggested not in sess.meta["assign"]:
        sess.meta["assign"][suggested] = file_id

    store.save(sess)

    # 화면에는 파일명/종류/뱃지/제안슬롯만 (내용 비표시)
    return {
        "file_id": file_id,
        "filename": original,
        "ext": v.ext,
        "valid": v.ok,
        "reasons": v.reasons,
        "suggested_slot": suggested,
        "suggested_label": SLOT_LABELS.get(suggested) if suggested else None,
        "assign": sess.meta["assign"],
    }


@app.delete("/api/file/{sid}/{file_id}")
def api_file_delete(sid: str, file_id: str):
    """업로드한 파일 1건 삭제 (디스크/메타/슬롯배정에서 제거)."""
    sess = store.get(sid)
    if sess is None:
        raise HTTPException(404, "세션을 찾을 수 없습니다.")

    info = sess.meta.get("files", {}).pop(file_id, None)
    if info is None:
        raise HTTPException(404, "삭제할 파일을 찾을 수 없습니다.")

    # 디스크 파일 제거
    path = info.get("path")
    if path:
        Path(path).unlink(missing_ok=True)

    # 이 파일을 가리키던 슬롯 배정 제거
    assign = sess.meta.get("assign", {})
    for slot in [s for s, fid in assign.items() if fid == file_id]:
        assign.pop(slot, None)

    store.save(sess)
    return {"ok": True, "file_id": file_id, "assign": assign}


@app.post("/api/assign")
async def api_assign(request: Request):
    body = await request.json()
    sid = body.get("sid")
    assign = body.get("assign", {})
    sess = store.get(sid)
    if sess is None:
        raise HTTPException(404, "세션을 찾을 수 없습니다.")
    # slot -> file_id 매핑 검증(존재하는 file 만)
    clean: dict[str, str] = {}
    for slot, file_id in assign.items():
        if slot in SLOT_KEYS and file_id and file_id in sess.meta["files"]:
            if sess.meta["files"][file_id]["valid"]:
                clean[slot] = file_id
    sess.meta["assign"] = clean
    store.save(sess)
    return {"assign": clean}


@app.post("/api/run")
async def api_run(request: Request):
    body = await request.json()
    sid = body.get("sid")
    sess = store.get(sid)
    if sess is None:
        raise HTTPException(404, "세션을 찾을 수 없습니다.")

    # 당기(년도/분기) 선택 검증 — 당기 기간 정의 및 추출의 기준
    try:
        period = parse_period(body.get("year"), body.get("quarter"))
    except ValueError as exc:
        return JSONResponse(status_code=400, content={"ok": False, "message": str(exc)})

    assign: dict[str, str] = sess.meta.get("assign", {})
    files: dict[str, dict] = sess.meta.get("files", {})

    # 필수 슬롯 점검: 최소한 당기말 잔액(검증용)은 있어야 함
    slot_paths: dict[str, Path | None] = {}
    slot_filenames: dict[str, str] = {}
    for slot in SLOT_KEYS:
        fid = assign.get(slot)
        info = files.get(fid, {}) if fid else {}
        slot_paths[slot] = Path(info["path"]) if fid and info.get("path") else None
        if fid and info.get("filename"):
            slot_filenames[slot] = info["filename"]

    if slot_paths.get("current_balance") is None:
        return JSONResponse(
            status_code=400,
            content={"ok": False, "message": "필수 입력(당기 채권채무 잔액 검증용 소스)이 없습니다. 슬롯을 지정하세요."},
        )

    settings = load_settings()
    outcome = run_pipeline(slot_paths, settings, sess.output_dir, slot_filenames, period)
    sess.outputs = dict(outcome.outputs)
    store.save(sess)

    downloads = [
        {"name": name, "url": f"/download/{sid}/{name}"}
        for name in outcome.outputs
    ]
    return {
        "ok": outcome.ok,
        "message": outcome.message,
        "error_count": outcome.error_count,
        "downloads": downloads,
    }


@app.get("/download/{sid}/{name}")
def download(sid: str, name: str):
    sess = store.get(sid)
    if sess is None or name not in sess.outputs:
        raise HTTPException(404, "다운로드할 파일을 찾을 수 없습니다.")
    path = sess.outputs[name]
    if not path.exists():
        raise HTTPException(404, "파일이 만료되었거나 삭제되었습니다.")
    return FileResponse(path, filename=path.name,
                        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.delete("/api/session/{sid}")
def api_session_delete(sid: str):
    store.delete(sid)
    return {"ok": True}
