"""pytest 부트스트랩 — 패키지 루트를 sys.path 에 올려 `import app...` 가능하게 한다."""

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))
