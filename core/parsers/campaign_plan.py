"""캠페인 플랜(xlsx) 파서 — 광고 집행 기간·매체 매칭용.

실제 6월 샘플로 확인된 형식: 시트명이 대상 월('6월' 등), 8행이 헤더
(매체, 카테고리, 주제, 유형, 광고 시작일, 광고 종료일, 집행일수, 목적, 랜딩, 캠페인).
'매체' 열은 그룹 첫 행에만 채워지고 그 아래는 비어있어 이어받는(forward-fill) 값으로 처리한다.
"""

from dataclasses import dataclass

from .base import ParseError, find_column, load_sheet_rows

HEADER_CANDIDATES = ["매체", "주제"]

COLS = {
    "media": ["매체"],
    "category": ["카테고리"],
    "topic": ["주제"],
    "type": ["유형"],
    "start": ["광고 시작일"],
    "end": ["광고 종료일"],
    "budget": ["캠페인"],
}


@dataclass
class CampaignPlanRow:
    media: str
    topic: str
    content_type: str | None
    start: object
    end: object
    budget: float | None


def parse_campaign_plan(file, sheet_name: str | None = None) -> list[CampaignPlanRow]:
    _, rows = load_sheet_rows(file, sheet_name)

    header_idx = None
    for i, row in enumerate(rows[:15]):
        texts = {str(c).strip() for c in row if c is not None}
        if set(HEADER_CANDIDATES) <= texts:
            header_idx = i
            break
    if header_idx is None:
        raise ParseError("캠페인 플랜에서 헤더 행('매체', '주제' 포함)을 찾지 못했습니다.")
    header = rows[header_idx]

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"캠페인 플랜에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    out: list[CampaignPlanRow] = []
    last_media = None
    for row in rows[header_idx + 1 :]:
        if col_idx["media"] < len(row) and row[col_idx["media"]]:
            last_media = str(row[col_idx["media"]]).strip()
        topic = row[col_idx["topic"]] if col_idx["topic"] < len(row) else None
        if not topic or not last_media:
            continue
        if last_media in ("총액", "매체"):
            continue
        out.append(
            CampaignPlanRow(
                media=last_media,
                topic=str(topic).strip(),
                content_type=(
                    str(row[col_idx["type"]]).strip()
                    if col_idx["type"] < len(row) and row[col_idx["type"]]
                    else None
                ),
                start=row[col_idx["start"]] if col_idx["start"] < len(row) else None,
                end=row[col_idx["end"]] if col_idx["end"] < len(row) else None,
                budget=row[col_idx["budget"]] if col_idx["budget"] < len(row) else None,
            )
        )
    return out
