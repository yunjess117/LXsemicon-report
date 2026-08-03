"""메타 광고 리포트(xlsx) 파서 — 인스타그램 참여/도달/동영상 조회/트래픽 캠페인 공용.

실제 6월 샘플로 확인된 형식: 'Raw Data Report' 시트, 컬럼:
캠페인 이름, 광고 세트 이름, 시작, 종료, 결과 유형, 결과, 결과당 비용, 노출, 도달,
게시물 참여, CPC(전체), CTR(전체), CPM(1,000회 노출당 비용), 지출 금액 (KRW), 빈도, 보고 시작, 보고 종료

캠페인 이름 끝의 '_참여'/'_도달'/'_동영상 조회'/'_트래픽'으로 어느 AD 시트로 갈지 판별한다.
광고 세트 이름 == 'All' 인 행은 캠페인 전체 합계 행이라 제외하고, 개별 광고 세트 행만 쓴다.
"""

from dataclasses import dataclass

from .base import ParseError, find_column, find_sheet_by_pattern, load_sheet_rows

HEADER_CANDIDATES = ["캠페인 이름", "광고 세트 이름"]

COLS = {
    "campaign_group": ["캠페인 이름"],
    "ad_set_name": ["광고 세트 이름"],
    "start": ["시작"],
    "end": ["종료"],
    "result": ["결과"],
    "cost_per_result": ["결과당 비용"],
    "impressions": ["노출"],
    "reach": ["도달"],
    "engagement": ["게시물 참여"],
    "cpc": ["CPC(전체)"],
    "ctr": ["CTR(전체)"],
    "cpm": ["CPM(1,000회 노출당 비용)", "CPM\n(1,000회 노출당 비용)"],
    "spend": ["지출 금액 (KRW)"],
}

SUFFIX_TO_TYPE = {
    "참여": "참여",
    "도달": "도달",
    "동영상 조회": "동영상 조회",
    "트래픽": "트래픽",
}


@dataclass
class MetaAdsRow:
    campaign_group: str
    ad_set_name: str
    ad_type: str | None  # 참여/도달/동영상 조회/트래픽
    start: object
    end: object
    result: float | None
    cost_per_result: float | None
    impressions: float | None
    reach: float | None
    engagement: float | None
    cpc: float | None
    ctr: float | None
    cpm: float | None
    spend: float | None


def _classify(campaign_group: str) -> str | None:
    for suffix, t in SUFFIX_TO_TYPE.items():
        if campaign_group.endswith(suffix):
            return t
    return None


def parse_meta_ads(file) -> list[MetaAdsRow]:
    try:
        sheet_name = find_sheet_by_pattern(file, r"Raw Data Report")
    except ParseError:
        sheet_name = None
    file.seek(0)
    _, rows = load_sheet_rows(file, sheet_name)

    header_idx = None
    for i, row in enumerate(rows[:6]):
        texts = {str(c).strip() for c in row if c is not None}
        if set(HEADER_CANDIDATES) <= texts:
            header_idx = i
            break
    if header_idx is None:
        raise ParseError("메타 광고 리포트에서 헤더 행을 찾지 못했습니다 ('Raw Data Report' 시트인지 확인해주세요).")
    header = rows[header_idx]

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"메타 광고 리포트에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    out: list[MetaAdsRow] = []
    for row in rows[header_idx + 1 :]:
        group = row[col_idx["campaign_group"]]
        ad_set = row[col_idx["ad_set_name"]]
        if not group or not ad_set or str(ad_set).strip() == "All":
            continue
        group = str(group).strip()
        out.append(
            MetaAdsRow(
                campaign_group=group,
                ad_set_name=str(ad_set).strip(),
                ad_type=_classify(group),
                start=row[col_idx["start"]],
                end=row[col_idx["end"]],
                result=row[col_idx["result"]],
                cost_per_result=row[col_idx["cost_per_result"]],
                impressions=row[col_idx["impressions"]],
                reach=row[col_idx["reach"]],
                engagement=row[col_idx["engagement"]],
                cpc=row[col_idx["cpc"]],
                ctr=row[col_idx["ctr"]],
                cpm=row[col_idx["cpm"]],
                spend=row[col_idx["spend"]],
            )
        )
    return out
