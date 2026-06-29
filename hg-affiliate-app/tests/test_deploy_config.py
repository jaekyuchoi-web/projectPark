"""배포 설정 회귀 테스트."""

from __future__ import annotations

import re
from pathlib import Path


def test_cloud_run_stays_single_instance_while_sessions_use_local_disk():
    """세션/출력 파일이 인스턴스 로컬 디스크에 있으므로 단일 인스턴스여야 한다."""
    deploy_script = Path(__file__).resolve().parents[1] / "deploy.sh"
    text = deploy_script.read_text(encoding="utf-8")

    match = re.search(r"--max-instances\s+(\d+)", text)

    assert match is not None
    assert match.group(1) == "1"
