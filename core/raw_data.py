"""기능 1: 로우데이터(취합) 파일 생성.

전월 fn 워크북을 그대로 받아서, 콘텐츠 Raw Data 5개 시트(뉴스룸/유튜브/인스타그램/페이스북/블로그)
맨 아래에 이번 달 블록만 추가한다. '월간 SNS 운영 수치'와 '채널 팔로워' 시트는 자동화 난이도가
매우 높아(수식·수기 혼재, 일부 기존 오류 존재) 이번 버전에서는 건드리지 않고 그대로 보존한다.
"""

from dataclasses import dataclass, field

from openpyxl.workbook import Workbook

from .confirm import ConfirmLog
from .excel_block import append_month_block, extract_row_formula_template, extract_summary_formula_template, find_month_blocks
from .formatters import format_kr_date, hms_string_to_hmm_slot, roundup_to_10k, seconds_to_hmm_slot
from .month_filter import _parse_month_from_kr_date, filter_by_url_band, filter_youtube_by_video_id, get_prev_max_url_id
from .parsers.base import extract_youtube_video_id

SHEET_NAMES = {
    "뉴스룸": "뉴스룸_Raw Data",
    "유튜브": "유튜브_Raw Data",
    "인스타그램": "인스타그램_Raw Data",
    "페이스북": "페이스북_Raw Data",
    "블로그": "블로그_Raw Data",
}


@dataclass
class ChannelMatchStats:
    channel: str
    total_published: int = 0
    matched: int = 0
    unmatched_titles: list[str] = field(default_factory=list)


@dataclass
class RawDataResult:
    workbook: Workbook
    confirm: ConfirmLog
    stats: dict[str, ChannelMatchStats]


def _norm_title(s: str) -> str:
    return "".join(str(s).split()).lower()


def _filter_rows_by_month(rows: list, target_month: int, area: str, log: ConfirmLog) -> list:
    """제목이 여러 달에 걸쳐 재사용될 수 있으므로(예: 같은 제목이 2월과 6월에 각각 발행),
    성과 파일과 제목으로 매칭하기 전에 반드시 발행일 기준으로 이번 달 것만 먼저 걸러낸다."""
    out = []
    for row in rows:
        month = _parse_month_from_kr_date(row.date_raw)
        if month is None:
            log.add(area, row.title, "발행일을 해석하지 못해 이번 달 콘텐츠인지 판별 불가")
            continue
        if month == target_month:
            out.append(row)
    return out


def _find_campaign_fee(campaign_rows: list[tuple[str, float]] | None, title: str) -> int | None:
    """(선택) AD 최종본에서 캠페인명에 콘텐츠 제목이 포함된 행을 찾아 캠페인비를 반환한다."""
    if not campaign_rows or not title:
        return None
    norm = _norm_title(title)
    for name, fee in campaign_rows:
        if norm and norm in _norm_title(name):
            return roundup_to_10k(fee) if fee is not None else None
    return None


def build_raw_data(
    prev_wb: Workbook,
    target_year: int,
    target_month: int,
    publish_rows: list,
    instagram_rows: list | None = None,
    facebook_rows: list | None = None,
    youtube_studio_rows: list | None = None,
    newsroom_ga_rows: list | None = None,
    blog_rows: list | None = None,
    youtube_campaign_rows: list[tuple[str, float]] | None = None,
    instagram_campaign_rows: list[tuple[str, float]] | None = None,
) -> RawDataResult:
    log = ConfirmLog()
    stats: dict[str, ChannelMatchStats] = {c: ChannelMatchStats(channel=c) for c in SHEET_NAMES}

    by_channel: dict[str, list] = {c: [] for c in SHEET_NAMES}
    for row in publish_rows:
        if row.channel in by_channel:
            by_channel[row.channel].append(row)

    _build_newsroom(prev_wb, target_year, target_month, by_channel["뉴스룸"], newsroom_ga_rows, log, stats["뉴스룸"])
    _build_blog(prev_wb, target_year, target_month, by_channel["블로그"], blog_rows, log, stats["블로그"])
    _build_youtube(
        prev_wb, target_year, target_month, by_channel["유튜브"], youtube_studio_rows,
        youtube_campaign_rows, log, stats["유튜브"],
    )
    _build_instagram(
        prev_wb, target_year, target_month, by_channel["인스타그램"], instagram_rows,
        instagram_campaign_rows, log, stats["인스타그램"],
    )
    _build_facebook(prev_wb, target_year, target_month, by_channel["페이스북"], facebook_rows, log, stats["페이스북"])

    return RawDataResult(workbook=prev_wb, confirm=log, stats=stats)


def _append_generic(ws, new_rows: list[dict]):
    blocks = find_month_blocks(ws)
    if not blocks:
        raise ValueError(f"{ws.title}: 기존 블록을 찾지 못해 새 블록을 추가할 수 없습니다.")
    last = blocks[-1]
    row_tpl = extract_row_formula_template(ws, last)
    summary_tpl = extract_summary_formula_template(ws, last)
    return append_month_block(ws, last, new_rows, row_tpl, summary_tpl)


def _build_newsroom(prev_wb, year, month, publish_rows, ga_rows, log, stat):
    ws = prev_wb[SHEET_NAMES["뉴스룸"]]
    prev_max = get_prev_max_url_id(ws, "F")
    result = filter_by_url_band(publish_rows, prev_max, month, "취합/뉴스룸_Raw Data")
    log.extend(result.confirm)
    accepted = result.accepted
    stat.total_published = len(accepted)

    if ga_rows is None:
        log.add("취합/뉴스룸_Raw Data", "H열(조회수)", "뉴스룸 콘텐츠별 조회수 원본이 없어 전부 비움")
        ga_by_url = {}
        ga_by_title = {}
    else:
        ga_by_url = {r.key: r.views for r in ga_rows}
        ga_by_title = {_norm_title(r.key): r.views for r in ga_rows}

    new_rows = []
    for row in accepted:
        views = ga_by_url.get(row.url) or ga_by_title.get(_norm_title(row.title))
        if views is not None:
            stat.matched += 1
        else:
            stat.unmatched_titles.append(row.title)
        new_rows.append(
            {
                "C": format_kr_date(row.date_raw, year),
                "D": row.category or _need_category(log, "취합/뉴스룸_Raw Data", row.title),
                "E": row.title,
                "F": row.url,
                "G": row.content_type,
                "H": views,
            }
        )
    if new_rows:
        _append_generic(ws, new_rows)


def _build_blog(prev_wb, year, month, publish_rows, blog_rows, log, stat):
    ws = prev_wb[SHEET_NAMES["블로그"]]
    prev_max = get_prev_max_url_id(ws, "F")
    result = filter_by_url_band(publish_rows, prev_max, month, "취합/블로그_Raw Data")
    log.extend(result.confirm)
    accepted = result.accepted
    stat.total_published = len(accepted)

    if blog_rows is None:
        log.add("취합/블로그_Raw Data", "H~J열", "블로그 실적 원본이 없어 전부 비움")
        by_url, by_title = {}, {}
    else:
        by_url = {r.key: r for r in blog_rows}
        by_title = {_norm_title(r.key): r for r in blog_rows}

    new_rows = []
    for row in accepted:
        perf = by_url.get(row.url) or by_title.get(_norm_title(row.title))
        if perf is not None:
            stat.matched += 1
        else:
            stat.unmatched_titles.append(row.title)
        new_rows.append(
            {
                "C": format_kr_date(row.date_raw, year),
                "D": row.category or _need_category(log, "취합/블로그_Raw Data", row.title),
                "E": row.title,
                "F": row.url,
                "G": row.content_type,
                "H": perf.views if perf else None,
                "I": perf.likes if perf else None,
                "J": perf.comments if perf else None,
            }
        )
    if new_rows:
        _append_generic(ws, new_rows)


def _build_youtube(prev_wb, year, month, publish_rows, studio_rows, campaign_rows, log, stat):
    ws = prev_wb[SHEET_NAMES["유튜브"]]
    if studio_rows is None:
        log.add("취합/유튜브_Raw Data", "전체", "유튜브 스튜디오 원본이 없어 당월 발행분을 판별할 수 없음")
        return
    video_ids = {r.video_id for r in studio_rows}
    result = filter_youtube_by_video_id(publish_rows, video_ids, "취합/유튜브_Raw Data")
    log.extend(result.confirm)
    accepted = result.accepted
    stat.total_published = len(accepted)

    studio_by_id = {r.video_id: r for r in studio_rows}

    new_rows = []
    for row in accepted:
        vid = extract_youtube_video_id(row.url)
        perf = studio_by_id.get(vid)
        if perf is None:
            stat.unmatched_titles.append(row.title)
            new_rows.append(
                {
                    "C": format_kr_date(row.date_raw, year),
                    "D": row.category or _need_category(log, "취합/유튜브_Raw Data", row.title),
                    "E": row.title,
                    "F": row.url,
                    "G": row.content_type,
                }
            )
            continue
        stat.matched += 1

        is_short = row.content_type is not None and "숏폼" in row.content_type
        avg_pct = "-" if is_short else perf.avg_view_pct
        if row.content_type is None:
            log.add(
                "취합/유튜브_Raw Data", row.title,
                "발행리스트에 형태(숏폼/롱폼) 정보가 없어 평균 조회율을 '-' 처리할지 결정 불가",
            )
            avg_pct = None

        length_slot = seconds_to_hmm_slot(perf.length_sec)
        if perf.length_sec is not None and length_slot is None:
            log.add("취합/유튜브_Raw Data", row.title, f"영상 길이({perf.length_sec}초)가 24분을 넘어 h:mm 슬롯으로 표현 불가")
        duration_slot = hms_string_to_hmm_slot(perf.avg_view_duration)

        fee = _find_campaign_fee(campaign_rows, row.title)
        if fee is None and campaign_rows is not None:
            log.add("취합/유튜브_Raw Data", row.title, "AD 최종본에서 이 콘텐츠와 일치하는 캠페인을 찾지 못해 캠페인비 비움")
        elif campaign_rows is None:
            log.add("취합/유튜브_Raw Data", row.title, "AD 최종본이 없어 캠페인비 비움")

        new_rows.append(
            {
                "C": format_kr_date(row.date_raw, year),
                "D": row.category or _need_category(log, "취합/유튜브_Raw Data", row.title),
                "E": row.title,
                "F": row.url,
                "G": row.content_type,
                "H": fee,
                "I": perf.views,
                "J": length_slot,
                "K": duration_slot,
                "L": avg_pct,
                "M": perf.likes,
                "N": perf.comments,
                "O": perf.shares,
            }
        )
    if new_rows:
        _append_generic(ws, new_rows)


def _build_instagram(prev_wb, year, month, publish_rows, insta_rows, campaign_rows, log, stat):
    ws = prev_wb[SHEET_NAMES["인스타그램"]]
    if insta_rows is None:
        log.add("취합/인스타그램_Raw Data", "전체", "인스타그램 성과 원본이 없어 성과를 채울 수 없음")
        insta_by_title = {}
    else:
        insta_by_title = {_norm_title(r.title): r for r in insta_rows}

    publish_rows = _filter_rows_by_month(publish_rows, month, "취합/인스타그램_Raw Data", log)
    stat.total_published = len(publish_rows)
    new_rows = []
    for row in publish_rows:
        perf = insta_by_title.get(_norm_title(row.title))
        if perf is None:
            stat.unmatched_titles.append(row.title)
            continue  # 성과 원본에 없으면 이번 달 발행분인지 자체가 불확실 -> 교차검증 실패로 스킵
        stat.matched += 1

        fee = _find_campaign_fee(campaign_rows, row.title)
        if fee is None:
            log.add(
                "취합/인스타그램_Raw Data", row.title,
                "AD 최종본이 없거나 일치하는 캠페인을 찾지 못해 캠페인비 비움" if campaign_rows is not None else "AD 최종본이 없어 캠페인비 비움",
            )

        new_rows.append(
            {
                "C": format_kr_date(row.date_raw, year),
                "D": row.category or _need_category(log, "취합/인스타그램_Raw Data", row.title),
                "E": row.title,
                "F": row.url or perf.url,
                "G": row.content_type,
                "H": fee,
                "I": perf.reach,
                "J": perf.views,
                # K(프로필 방문)/L(웹 방문): 원본에 없는 수기 항목 - 아래에서 확인 필요 처리
                "M": perf.likes,
                "N": perf.comments,
                "O": perf.saves,
                "P": perf.shares,
            }
        )
        log.add(
            "취합/인스타그램_Raw Data", row.title,
            "프로필 방문/웹 방문은 원본에 없는 수기 입력 항목 (인스타 인사이트에서 별도 확인 필요)",
        )

    # 발행리스트에 없는데 성과 파일에만 있는 콘텐츠(발행리스트 누락 가능성)
    matched_titles = {_norm_title(r.title) for r in publish_rows}
    for perf in insta_rows or []:
        if _norm_title(perf.title) not in matched_titles:
            log.add(
                "취합/인스타그램_Raw Data", perf.title,
                "인스타그램 성과 원본에는 있지만 발행리스트에서 찾지 못함 (카테고리/유형 확인 필요)",
            )

    if new_rows:
        _append_generic(ws, new_rows)


def _build_facebook(prev_wb, year, month, publish_rows, fb_rows, log, stat):
    ws = prev_wb[SHEET_NAMES["페이스북"]]
    if fb_rows is None:
        log.add("취합/페이스북_Raw Data", "전체", "페이스북 성과 원본이 없어 성과를 채울 수 없음")
        fb_by_title = {}
    else:
        fb_by_title = {_norm_title(r.title): r for r in fb_rows}

    publish_rows = _filter_rows_by_month(publish_rows, month, "취합/페이스북_Raw Data", log)
    stat.total_published = len(publish_rows)
    new_rows = []
    for row in publish_rows:
        perf = fb_by_title.get(_norm_title(row.title))
        if perf is None:
            stat.unmatched_titles.append(row.title)
            continue
        stat.matched += 1
        new_rows.append(
            {
                "C": format_kr_date(row.date_raw, year),
                "D": row.category or _need_category(log, "취합/페이스북_Raw Data", row.title),
                "E": row.title,
                "F": row.url or perf.url,
                "G": row.content_type,
                "H": "-",  # 페이스북은 광고 집행 시에만 값이 생김 (기본 '-')
                "I": perf.reach,
                "J": perf.impressions,
                "K": perf.likes,
                "L": perf.comments,
                "M": perf.shares,
            }
        )
    if new_rows:
        _append_generic(ws, new_rows)


def _need_category(log: ConfirmLog, area: str, title: str) -> None:
    log.add(area, title, "발행리스트에 카테고리가 없어 수기 분류 필요")
    return None
