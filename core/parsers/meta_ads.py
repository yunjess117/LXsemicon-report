"""메타 광고 리포트 파서(xlsx/csv) — 인스타그램 참여/도달/동영상 조회/트래픽 캠페인 공용.

실제 6월 샘플로 확인된 형식: 'Raw Data Report' 시트, 컬럼:
캠페인 이름, 광고 세트 이름, 시작, 종료, 결과 유형, 결과, 결과당 비용, 노출, 도달,
게시물 참여, CPC(전체), CTR(전체), CPM(1,000회 노출당 비용), 지출 금액 (KRW), 빈도, 보고 시작, 보고 종료

같은 파일의 'Formatted Report' 시트도 헤더가 거의 같지만, 거기서는 광고 세트 행의
'캠페인 이름'이 비어 있어(캠페인명은 All 행에만 표기) 그대로 쓰면 전부 걸러진다.
그래서 'Raw Data' 시트를 우선으로 고른다.

캠페인 이름 끝의 '_참여'/'_도달'/'_동영상 조회'/'_트래픽'으로 어느 AD 시트로 갈지 판별한다.
광고 세트 이름 == 'All' 인 행은 캠페인 전체 합계 행이라 제외하고, 개별 광고 세트 행만 쓴다.
"""

from dataclasses import dataclass

from .base import ParseError, find_column, load_any_tabular, to_number

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


def _find_header(file):
    """헤더가 있는 표를 찾되, 광고 세트 행에도 캠페인명이 채워져 있는 'Raw Data' 쪽을 우선한다."""
    fallback = None
    for sheet_name, rows in load_any_tabular(file):
        for i, row in enumerate(rows[:6]):
            texts = {str(c).strip() for c in row if c is not None}
            if not (set(HEADER_CANDIDATES) <= texts):
                continue
            found = (row, rows[i + 1 :])
            if "raw data" in str(sheet_name).lower():
                return found
            if fallback is None:
                fallback = found
            break
    return fallback if fallback else (None, None)


def parse_meta_ads(file) -> list[MetaAdsRow]:
    header, data_rows = _find_header(file)
    if header is None:
        raise ParseError(
            "메타 광고 리포트에서 헤더 행(캠페인 이름, 광고 세트 이름 포함)을 찾지 못했습니다. "
            "메타 광고 관리자에서 내려받은 원본(xlsx 또는 csv)이 맞는지 확인해주세요."
        )

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"메타 광고 리포트에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    def cell(row, key):
        i = col_idx[key]
        return row[i] if i < len(row) else None

    out: list[MetaAdsRow] = []
    for row in data_rows:
        group = cell(row, "campaign_group")
        ad_set = cell(row, "ad_set_name")
        if not group or not ad_set or str(ad_set).strip() == "All":
            continue
        group = str(group).strip()
        out.append(
            MetaAdsRow(
                campaign_group=group,
                ad_set_name=str(ad_set).strip(),
                ad_type=_classify(group),
                start=cell(row, "start"),
                end=cell(row, "end"),
                result=to_number(cell(row, "result")),
                cost_per_result=to_number(cell(row, "cost_per_result")),
                impressions=to_number(cell(row, "impressions")),
                reach=to_number(cell(row, "reach")),
                engagement=to_number(cell(row, "engagement")),
                cpc=to_number(cell(row, "cpc")),
                ctr=to_number(cell(row, "ctr")),
                cpm=to_number(cell(row, "cpm")),
                spend=to_number(cell(row, "spend")),
            )
        )
    return out
