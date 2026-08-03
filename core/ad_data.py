"""기능 2: 로우데이터 AD 파일 생성.

시트 구성: 유튜브_캠페인 / 인스타그램_참여 / 인스타그램_도달 / 인스타그램_동영상 조회 /
인스타그램_트래픽(집행 시에만) / 버즈빌_캠페인.

캠페인 유형(숏폼/롱폼 등)은 발행리스트에서, 집행 기간은 광고 원본에 이미 있으면 그대로 쓰고
없으면 캠페인 플랜에서 매칭한다. 캠페인비는 실지출을 만원(10,000원) 단위로 올림한다
(문서에 명시된 규칙을 그대로 따름 — 실제 파일 일부 수식이 천원 단위인 것과는 다름, 사용자 확인).
버즈빌_캠페인 시트는 다른 시트들과 달리 12개월 고정 캘린더 구조라 이번 버전에서는
자동화하지 않고 항상 확인 필요로 남긴다.
"""

import re
from dataclasses import dataclass, field

from openpyxl.utils import column_index_from_string
from openpyxl.workbook import Workbook

from .confirm import ConfirmLog
from .excel_block import append_month_block, extract_row_formula_template, extract_summary_formula_template, find_month_blocks
from .formatters import roundup_to_10k

DATE_PREFIX_RE = re.compile(r"^\d{6}_")


def _strip_date_prefix(name: str) -> str:
    return DATE_PREFIX_RE.sub("", name).strip()


def _norm(s) -> str:
    return "".join(str(s).split()).lower()


@dataclass
class ChannelStats:
    channel: str
    total_published: int = 0
    matched: int = 0
    unmatched_titles: list[str] = field(default_factory=list)


@dataclass
class AdDataResult:
    workbook: Workbook
    confirm: ConfirmLog
    stats: dict[str, ChannelStats]


SHEET_COLUMN_MAP = {
    "인스타그램_참여": {
        "result": "G", "cost_per_result": "H", "impressions": "I", "reach": "J",
        "engagement": "K", "cpc": "L", "ctr": "M", "cpm": "N", "spend": "O", "fee": "P",
    },
    "인스타그램_도달": {
        "result": "G", "cost_per_result": "H", "impressions": "I", "reach": "J",
        "engagement": "K", "cpc": "L", "ctr": "M", "cpm": "N", "spend": "O", "fee": "P",
    },
    "인스타그램_동영상 조회": {
        "result": "G", "cost_per_result": "H", "impressions": "I", "reach": "J",
        "engagement": "K", "cpc": "L", "ctr": "M", "cpm": "N", "cpv": "O", "vtr": "P",
        "spend": "Q", "fee": "R",
    },
    "인스타그램_트래픽": {
        "result": "G", "cost_per_result": "H", "impressions": "I", "reach": "J",
        "engagement": "K", "cpc": "L", "ctr": "M", "cpm": "N", "spend": "O",
    },
}


def _match_content_type(publish_rows: list, ad_set_name: str, log: ConfirmLog, area: str) -> str | None:
    title_guess = _strip_date_prefix(ad_set_name)
    norm_guess = _norm(title_guess)
    for row in publish_rows:
        norm_title = _norm(row.title)
        if norm_title and (norm_title in norm_guess or norm_guess in norm_title):
            return row.content_type
    log.add(area, ad_set_name, "발행리스트에서 일치하는 콘텐츠를 찾지 못해 캠페인 유형(숏폼/롱폼 등) 비움")
    return None


def _format_period(start, end) -> str | None:
    def _to_md(v):
        if v is None:
            return None
        if hasattr(v, "month"):
            return f"{v.month}/{v.day}"
        m = re.match(r"(\d{4})-(\d{1,2})-(\d{1,2})", str(v))
        if m:
            return f"{int(m.group(2))}/{int(m.group(3))}"
        return None

    s, e = _to_md(start), _to_md(end)
    if s and e:
        return f"{s} ~ {e}"
    return None


def _match_campaign_plan_period(campaign_plan_rows: list | None, media: str, ad_set_name: str) -> str | None:
    if not campaign_plan_rows:
        return None
    title_guess = _norm(_strip_date_prefix(ad_set_name))
    for row in campaign_plan_rows:
        if row.media != media:
            continue
        norm_topic = _norm(row.topic)
        if title_guess and (norm_topic in title_guess or title_guess in norm_topic):
            return _format_period(row.start, row.end)
    return None


def _force_sum_avg(ws, block, col_letter: str):
    col_idx = column_index_from_string(col_letter)
    ws.cell(row=block.total_row, column=col_idx).value = f"=SUM({col_letter}{block.data_start_row}:{col_letter}{block.data_end_row})"
    ws.cell(row=block.avg_row, column=col_idx).value = f"=AVERAGE({col_letter}{block.data_start_row}:{col_letter}{block.data_end_row})"


def build_ad_data(
    prev_wb: Workbook,
    target_year: int,
    target_month: int,
    publish_rows: list,
    google_ads_rows: list | None = None,
    meta_ads_rows: list | None = None,
    campaign_plan_rows: list | None = None,
) -> AdDataResult:
    log = ConfirmLog()
    stats: dict[str, ChannelStats] = {}
    month_label = f"{target_month}월"

    stats["유튜브_캠페인"] = ChannelStats(channel="유튜브_캠페인")
    _build_youtube_campaign(prev_wb, month_label, publish_rows, google_ads_rows, campaign_plan_rows, log, stats["유튜브_캠페인"])

    for sheet_name, ad_type in [
        ("인스타그램_참여", "참여"),
        ("인스타그램_도달", "도달"),
        ("인스타그램_동영상 조회", "동영상 조회"),
        ("인스타그램_트래픽", "트래픽"),
    ]:
        stats[sheet_name] = ChannelStats(channel=sheet_name)
        _build_meta_sheet(
            prev_wb, sheet_name, ad_type, month_label, publish_rows, meta_ads_rows, campaign_plan_rows, log, stats[sheet_name]
        )

    log.add(
        "AD/버즈빌_캠페인", "전체",
        "버즈빌_캠페인 시트는 12개월 고정 캘린더 구조라 이번 버전에서는 자동 채움 대상이 아닙니다. "
        "버즈빌 리포트를 참고해 수기로 채워주세요.",
    )

    return AdDataResult(workbook=prev_wb, confirm=log, stats=stats)


def _build_youtube_campaign(prev_wb, month_label, publish_rows, google_ads_rows, campaign_plan_rows, log, stat):
    ws = prev_wb["유튜브_캠페인"]
    if google_ads_rows is None:
        log.add("AD/유튜브_캠페인", "전체", "구글애즈 광고 리포트가 없어 이 시트를 채우지 못함")
        return
    blocks = find_month_blocks(ws, num_col="C")
    if not blocks:
        raise ValueError("유튜브_캠페인: 기존 블록을 찾지 못함")
    last = blocks[-1]
    row_tpl = extract_row_formula_template(ws, last)
    summary_tpl = extract_summary_formula_template(ws, last)
    row_tpl.pop("L", None)  # 캠페인비: 만원 단위 직접 계산

    stat.total_published = len(google_ads_rows)
    new_rows = []
    for r in google_ads_rows:
        content_type = _match_content_type(publish_rows, r.campaign_name, log, "AD/유튜브_캠페인")
        if content_type:
            # 콘텐츠 Raw Data 시트는 '영상(숏폼)'/'영상(롱폼)'으로 저장하지만, 유튜브_캠페인
            # 시트의 '캠페인 유형' 열은 실제로 '숏폼'/'롱폼'만 줄여서 쓴다(실제 파일로 확인).
            m = re.match(r"영상\((숏폼|롱폼)\)", content_type)
            if m:
                content_type = m.group(1)
        period = _match_campaign_plan_period(campaign_plan_rows, "유튜브", r.campaign_name)
        if period is None:
            log.add("AD/유튜브_캠페인", r.campaign_name, "집행 기간을 캠페인 플랜에서 찾지 못해 비움")
        cpv = round(r.spend / r.views) if r.spend is not None and r.views else None
        vtr = (r.views / r.impressions) if r.impressions else None
        fee = roundup_to_10k(r.spend)
        if content_type is not None:
            stat.matched += 1
        else:
            stat.unmatched_titles.append(r.campaign_name)
        new_rows.append(
            {
                "B": month_label,
                "D": r.campaign_name,
                "E": content_type,
                "F": period,
                "G": r.impressions,
                "H": r.views,
                "I": cpv,
                "J": vtr,
                "K": r.spend,
                "L": fee,
            }
        )
    if new_rows:
        nb = append_month_block(ws, last, new_rows, row_tpl, summary_tpl, number_col="C")
        _force_sum_avg(ws, nb, "L")


def _build_meta_sheet(prev_wb, sheet_name, ad_type, month_label, publish_rows, meta_ads_rows, campaign_plan_rows, log, stat):
    ws = prev_wb[sheet_name]
    colmap = SHEET_COLUMN_MAP[sheet_name]
    fee_col = colmap.get("fee")

    if meta_ads_rows is None:
        log.add(f"AD/{sheet_name}", "전체", "메타 광고 리포트가 없어 이 시트를 채우지 못함")
        return

    rows_for_type = [r for r in meta_ads_rows if r.ad_type == ad_type]
    if not rows_for_type:
        return  # 이번 달에 이 유형의 캠페인이 없으면 블록을 추가하지 않는다 (정상)

    blocks = find_month_blocks(ws, num_col="C")
    if not blocks:
        raise ValueError(f"{sheet_name}: 기존 블록을 찾지 못함")
    last = blocks[-1]
    row_tpl = extract_row_formula_template(ws, last)
    summary_tpl = extract_summary_formula_template(ws, last)
    if fee_col:
        row_tpl.pop(fee_col, None)

    media = "인스타그램"
    stat.total_published = len(rows_for_type)
    new_rows = []
    for r in rows_for_type:
        content_type = _match_content_type(publish_rows, r.ad_set_name, log, f"AD/{sheet_name}")
        period = _format_period(r.start, r.end) or _match_campaign_plan_period(campaign_plan_rows, media, r.ad_set_name)
        if period is None:
            log.add(f"AD/{sheet_name}", r.ad_set_name, "집행 기간을 원본/캠페인 플랜에서 찾지 못해 비움")
        if content_type is not None:
            stat.matched += 1
        else:
            stat.unmatched_titles.append(r.ad_set_name)

        row_data = {
            "B": month_label,
            "D": r.ad_set_name,
            "E": content_type,
            "F": period,
            colmap["result"]: r.result,
            colmap["cost_per_result"]: r.cost_per_result,
            colmap["impressions"]: r.impressions,
            colmap["reach"]: r.reach,
            colmap["engagement"]: r.engagement,
            colmap["cpc"]: r.cpc,
            colmap["ctr"]: r.ctr,
            colmap["cpm"]: r.cpm,
            colmap["spend"]: r.spend,
        }
        if "cpv" in colmap:
            row_data[colmap["cpv"]] = r.cost_per_result
        if "vtr" in colmap:
            row_data[colmap["vtr"]] = (r.result / r.impressions) if r.impressions else None
        if fee_col:
            row_data[fee_col] = roundup_to_10k(r.spend)
        new_rows.append(row_data)

    nb = append_month_block(ws, last, new_rows, row_tpl, summary_tpl, number_col="C")
    if fee_col:
        _force_sum_avg(ws, nb, fee_col)
