"""세션 격리 및 만료 삭제.

업로드 파일과 출력 파일은 세션별 임시 디렉터리에서 처리하고, TTL 이 지난
세션은 자동으로 삭제한다. meta.json 으로 디스크에 영속화해 동일 인스턴스에서
메모리가 비어도 복원할 수 있다. (Cloud Run 은 max-instances=1 권장)
"""

from __future__ import annotations

import json
import shutil
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from threading import Lock

from .config import SESSION_ROOT, SESSION_TTL_SECONDS

_META_FILE = "meta.json"


@dataclass
class Session:
    sid: str
    created_at: float
    root: Path
    uploads: dict[str, Path] = field(default_factory=dict)
    outputs: dict[str, Path] = field(default_factory=dict)
    meta: dict = field(default_factory=dict)

    @property
    def upload_dir(self) -> Path:
        d = self.root / "uploads"
        d.mkdir(parents=True, exist_ok=True)
        return d

    @property
    def output_dir(self) -> Path:
        d = self.root / "outputs"
        d.mkdir(parents=True, exist_ok=True)
        return d


class SessionStore:
    """세션 저장소. 메모리 + 디스크(meta.json) 이중 저장."""

    def __init__(self) -> None:
        self._sessions: dict[str, Session] = {}
        self._lock = Lock()
        SESSION_ROOT.mkdir(parents=True, exist_ok=True)

    def create(self) -> Session:
        self._sweep()
        sid = uuid.uuid4().hex
        root = SESSION_ROOT / sid
        root.mkdir(parents=True, exist_ok=True)
        sess = Session(sid=sid, created_at=time.time(), root=root)
        sess.meta.setdefault("files", {})
        sess.meta.setdefault("assign", {})
        with self._lock:
            self._sessions[sid] = sess
        self._persist(sess)
        return sess

    def get(self, sid: str) -> Session | None:
        self._sweep()
        with self._lock:
            cached = self._sessions.get(sid)
        if cached is not None:
            return cached
        loaded = self._load_from_disk(sid)
        if loaded is None:
            return None
        with self._lock:
            self._sessions[sid] = loaded
        return loaded

    def save(self, sess: Session) -> None:
        """메타·출력 경로를 디스크에 반영."""
        with self._lock:
            self._sessions[sess.sid] = sess
        self._persist(sess)

    def delete(self, sid: str) -> None:
        with self._lock:
            sess = self._sessions.pop(sid, None)
        root = SESSION_ROOT / sid
        if sess and sess.root.exists():
            shutil.rmtree(sess.root, ignore_errors=True)
        elif root.exists():
            shutil.rmtree(root, ignore_errors=True)

    def _persist(self, sess: Session) -> None:
        payload = {
            "created_at": sess.created_at,
            "meta": sess.meta,
            "outputs": {name: str(path) for name, path in sess.outputs.items()},
        }
        (sess.root / _META_FILE).write_text(
            json.dumps(payload, ensure_ascii=False),
            encoding="utf-8",
        )

    def _load_from_disk(self, sid: str) -> Session | None:
        root = SESSION_ROOT / sid
        if not root.is_dir():
            return None
        meta_path = root / _META_FILE
        if not meta_path.is_file():
            return Session(
                sid=sid,
                created_at=root.stat().st_mtime,
                root=root,
                meta={"files": {}, "assign": {}},
            )
        try:
            data = json.loads(meta_path.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            return None
        outputs = {k: Path(v) for k, v in data.get("outputs", {}).items()}
        meta = data.get("meta") or {}
        meta.setdefault("files", {})
        meta.setdefault("assign", {})
        return Session(
            sid=sid,
            created_at=float(data.get("created_at", time.time())),
            root=root,
            outputs=outputs,
            meta=meta,
        )

    def _sweep(self) -> None:
        now = time.time()
        expired: list[str] = []
        with self._lock:
            for sid, sess in self._sessions.items():
                if now - sess.created_at > SESSION_TTL_SECONDS:
                    expired.append(sid)
            for sid in expired:
                self._sessions.pop(sid, None)
        for sid in expired:
            shutil.rmtree(SESSION_ROOT / sid, ignore_errors=True)
        if SESSION_ROOT.is_dir():
            for child in SESSION_ROOT.iterdir():
                if not child.is_dir():
                    continue
                sid = child.name
                if sid in self._sessions:
                    continue
                meta_path = child / _META_FILE
                created = child.stat().st_mtime
                if meta_path.is_file():
                    try:
                        created = float(json.loads(meta_path.read_text(encoding="utf-8")).get("created_at", created))
                    except (json.JSONDecodeError, OSError):
                        pass
                if now - created > SESSION_TTL_SECONDS:
                    shutil.rmtree(child, ignore_errors=True)


store = SessionStore()
