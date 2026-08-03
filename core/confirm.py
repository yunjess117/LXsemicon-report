"""'확인 필요' 항목을 모듈 전체에서 일관되게 쌓기 위한 공용 유틸.

판단이 애매하거나 원본에 없는 값은 절대 추정해서 채우지 않고,
이 리스트에 사유와 함께 남긴 뒤 화면(및 필요하면 셀)에 표시한다.
"""

from dataclasses import dataclass, field


@dataclass
class ConfirmItem:
    area: str  # 예: "취합/뉴스룸_Raw Data", "AD/유튜브_캠페인", "입력 파일"
    location: str  # 예: "H23", "6월 블록 3번째 행", "파일 누락"
    reason: str  # 왜 확인이 필요한지 (어떤 값이 왜 불확실한지)

    def label(self) -> str:
        return f"⚠ 확인 필요 [{self.area} / {self.location}] {self.reason}"


@dataclass
class ConfirmLog:
    """한 번의 처리(파일 생성) 동안 쌓이는 확인 필요 항목 모음."""

    items: list[ConfirmItem] = field(default_factory=list)

    def add(self, area: str, location: str, reason: str) -> None:
        self.items.append(ConfirmItem(area=area, location=location, reason=reason))

    def extend(self, other: "ConfirmLog") -> None:
        self.items.extend(other.items)

    def __len__(self) -> int:
        return len(self.items)

    def __bool__(self) -> bool:
        return bool(self.items)

    def by_area(self) -> dict[str, list[ConfirmItem]]:
        grouped: dict[str, list[ConfirmItem]] = {}
        for item in self.items:
            grouped.setdefault(item.area, []).append(item)
        return grouped
