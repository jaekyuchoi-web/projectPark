"""오류/검토 항목 수집기.

[구분, 대상, 내용, 권고조치] 4열로 누적한다. 산출 '수치'는 담지 않는다.
"""

from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class ErrorItem:
    category: str   # 구분 (정규화/검증/산출불가/형식 등)
    target: str     # 대상 (파일/시트/거래처/계정)
    content: str    # 내용
    action: str     # 권고조치


@dataclass
class ErrorLog:
    items: list[ErrorItem] = field(default_factory=list)

    def add(self, category: str, target: str, content: str, action: str) -> None:
        self.items.append(ErrorItem(category, target, content, action))

    def extend_unmatched(self, names: list[str]) -> None:
        for n in names:
            self.add("정규화 미매칭", f"거래처: {n}", "대표 법인명 매핑에 실패했거나 시드 정규명에 포함되지 않음",
                     "특관자 상호 정리 파일에 대표명/별칭을 보완하세요.")

    @property
    def count(self) -> int:
        return len(self.items)
