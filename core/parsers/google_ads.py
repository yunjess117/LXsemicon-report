"""구글애즈 동영상 캠페인 리포트 파서. csv(UTF-16 탭 구분)와 xlsx 둘 다 받는다.

실제 6월 샘플로 확인된 형식: 상단 2줄(제목/기간) + 3번째 줄이 헤더
(캠페인, 노출수, TrueView 조회수, 통화 코드, TrueView 평균 CPV, TrueView 조회율, 비용).
"""

from dataclasses import dataclass

from .base import ParseError, find_column, load_any_tabular, to_number

HEADER_MARKER = "캠페인"
COLS = {
    "campaign_name": ["캠페인"],
    "impressions": ["노출수", "노출"],
    "views": ["TrueView 조회수", "조회수"],
    "cpv": ["TrueView 평균 CPV", "평균 CPV", "CPV"],
    "vtr": ["TrueView 조회율", "조회율"],
    "spend": ["비용", "지출", "지출 금액"],
}


@dataclass
class GoogleAdsRow:
    campaign_name: str
    impressions: float
    views: float
    cpv: float
    vtr: float
    spend: float


def _pct(v):
    """조회율은 '21.51%' 문자열로 오기도 하고 0.2151 비율로 오기도 한다 — % 단위로 통일."""
    n = to_number(str(v).replace("%", "") if isinstance(v, str) else v)
    if n is None:
        return None
    return n * 100 if (isinstance(v, (int, float)) and n <= 1) else n


def parse_google_ads(file) -> list[GoogleAdsRow]:
    header = data_rows = None
    for _, rows in load_any_tabular(file):
        for i, row in enumerate(rows[:8]):
            first = str(row[0]).strip() if row and row[0] is not None else ""
            if first == HEADER_MARKER:
                header, data_rows = row, rows[i + 1 :]
                break
        if header is not None:
            break
    if header is None:
        raise ParseError(
            f"구글애즈 리포트에서 헤더 행('{HEADER_MARKER}'으로 시작)을 찾지 못했습니다. "
            "구글애즈에서 내려받은 원본(csv 또는 xlsx)이 맞는지 확인해주세요."
        )

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"구글애즈 리포트에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    def cell(row, key):
        i = col_idx[key]
        return row[i] if i < len(row) else None

    out: list[GoogleAdsRow] = []
    for row in data_rows:
        name = cell(row, "campaign_name")
        if not name or not str(name).strip():
            continue
        out.append(
            GoogleAdsRow(
                campaign_name=str(name).strip(),
                impressions=to_number(cell(row, "impressions")),
                views=to_number(cell(row, "views")),
                cpv=to_number(cell(row, "cpv")),
                vtr=_pct(cell(row, "vtr")),
                spend=to_number(cell(row, "spend")),
            )
        )
    return out
