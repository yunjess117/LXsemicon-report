"""구글애즈 동영상 캠페인 리포트(CSV, UTF-16) 파서.

실제 6월 샘플로 확인된 형식: 상단 2줄(제목/기간) + 3번째 줄이 헤더
(캠페인, 노출수, TrueView 조회수, 통화 코드, TrueView 평균 CPV, TrueView 조회율, 비용).
"""

import csv
import io
from dataclasses import dataclass

from .base import ParseError

HEADER_MARKER = "캠페인"


@dataclass
class GoogleAdsRow:
    campaign_name: str
    impressions: float
    views: float
    cpv: float
    vtr: float
    spend: float


def _to_num(v):
    if v in (None, "", "-"):
        return None
    return float(str(v).replace(",", "").replace('"', "").replace("%", ""))


def parse_google_ads(file) -> list[GoogleAdsRow]:
    raw = file.read()
    text = raw.decode("utf-16") if isinstance(raw, bytes) else raw
    reader = csv.reader(io.StringIO(text), delimiter="\t")
    rows = [r for r in reader if any(c.strip() for c in r)]
    if not rows:
        # 일부 내보내기는 콤마 구분일 수 있어 재시도
        reader = csv.reader(io.StringIO(text))
        rows = [r for r in reader if any(c.strip() for c in r)]

    header_idx = None
    for i, row in enumerate(rows[:5]):
        if row and row[0].strip() == HEADER_MARKER:
            header_idx = i
            break
    if header_idx is None:
        raise ParseError(
            f"구글애즈 리포트에서 헤더 행('{HEADER_MARKER}'로 시작)을 찾지 못했습니다. "
            "UTF-16 CSV 원본이 맞는지 확인해주세요."
        )

    out: list[GoogleAdsRow] = []
    for row in rows[header_idx + 1 :]:
        if len(row) < 7 or not row[0].strip():
            continue
        try:
            out.append(
                GoogleAdsRow(
                    campaign_name=row[0].strip(),
                    impressions=_to_num(row[1]),
                    views=_to_num(row[2]),
                    cpv=_to_num(row[4]),
                    vtr=_to_num(row[5]),
                    spend=_to_num(row[6]),
                )
            )
        except ValueError:
            continue
    return out
