"""세션 격리 및 만료 삭제.

업로드 파일과 출력 파일은 세션별 임시 디렉터리에서 처리하고, TTL 이 지난
세션은 자동으로 삭제한다. (비기능 요구: 세션격리·만료 삭제)
"""

from __future__ import annotations

import shutil
import time
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from threading import Lock

from .config import SESSION_ROOT, SESSION_TTL_SECONDS


@dataclass
class Session:
    sid: str
    created_at: float
    root: Path
    # 슬롯키 -> 저장된 업로드 파일 경로
    uploads: dict[str, Path] = field(default_factory=dict)
    # 출력물 이름 -> 경로
    outputs: dict[str, Path] = field(default_factory=dict)
    # 마지막 분류/검증 결과 캐시 (화면 표시는 파일명/뱃지만)
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
    """프로세스 메모리 기반 세션 저장소 (단일 인스턴스 기준)."""

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
        with self._lock:
            self._sessions[sid] = sess
        return sess

    def get(self, sid: str) -> Session | None:
        self._sweep()
        with self._lock:
            return self._sessions.get(sid)

    def delete(self, sid: str) -> None:
        with self._lock:
            sess = self._sessions.pop(sid, None)
        if sess and sess.root.exists():
            shutil.rmtree(sess.root, ignore_errors=True)

    def _sweep(self) -> None:
        """TTL 초과 세션 정리."""
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


# 전역 단일 저장소
store = SessionStore()
