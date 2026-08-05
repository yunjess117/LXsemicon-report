"""기능 1: 로우데이터(취합) 파일 생성.

전월 fn 워크북을 그대로 받아서, 콘텐츠 Raw Data 5개 시트(뉴스룸/유튜브/인스타그램/페이스북/블로그)
맨 아래에 이번 달 블록만 추가한다. '월간 SNS 운영 수치'와 '채널 팔로워' 시트는 자동화 난이도가
매우 높아(수식·수기 혼재, 일부 기존 오류 존재) 이번 버전에서는 건드리지 않고 그대로 보존한다.
"""

import datetime
import re
from dataclasses import dataclass, field

from openpyxl.workbook import Workbook

from .confirm import ConfirmLog
from .excel_block import append_month_block, extract_row_formula_template, extract_summary_formula_template, find_month_blocks
from .formatters import format_kr_date, hms_string_to_hmm_slot, parse_month_day, parse_posted_at_date, roundup_to_10k, seconds_to_hmm_slot
from .month_filter import _parse_month_from_kr_date, filter_by_url_band, filter_youtube_by_video_id, get_prev_max_url_id
from .parsers.base import extract_url_id, extract_youtube_video_id

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


TITLE_SIMILARITY_THRESHOLD = 0.5


def _title_similarity(a: str, b: str) -> float:
    import difflib

    return difflib.SequenceMatcher(None, _norm_title(a), _norm_title(b)).ratio()


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


def _find_campaign_fee(campaign_rows: list[tuple[str, float]] | None, row) -> int | None:
    """(선택) AD 최종본에서 이 콘텐츠에 해당하는 캠페인을 찾아 캠페인비를 반환한다.
    제목 포함 관계로 먼저 보고, 안 맞으면 시리즈명+발행일이 같은지로 판단한다
    (예: 'B급 단장(공통점)' ≡ 'B급 콘텐츠(공통점)')."""
    from .ad_data import _same_content

    if not campaign_rows or not getattr(row, "title", None):
        return None
    for name, fee in campaign_rows:
        if _same_content(name, row):
            return roundup_to_10k(fee) if fee is not None else None
    return None


def make_campaign_plan_fee_matcher(plan_rows: list | None, media: str, year: int):
    """캠페인 플랜에서 '광고 시작일'이 발행일과 같은 항목을 찾아 캠페인비를 돌려주는 매처.

    캠페인 플랜의 주제(제목)는 발행리스트와 표기가 달라(리뉴얼 표시, 의역 등) 제목만으로는
    매칭이 잘 안 되므로 발행일을 기준으로 삼는다. 같은 날짜에 여러 건이면 그 안에서 제목이
    가장 비슷한 것을 고르고, 그래도 못 고르면 확인 필요로 남긴다. 한 플랜 항목은 한 번만 쓴다.
    """
    rows = [r for r in (plan_rows or []) if r.media == media]
    by_date: dict[datetime.date, list] = {}
    for r in rows:
        d = parse_posted_at_date(r.start)
        if d is not None:
            by_date.setdefault(d, []).append(r)
    used: set[int] = set()

    def match(pub_row, area: str, log: ConfirmLog) -> int | None:
        md = parse_month_day(pub_row.date_raw)
        if md is None:
            return None
        try:
            target = datetime.date(year, md[0], md[1])
        except ValueError:
            return None
        candidates = [c for c in by_date.get(target, []) if id(c) not in used]
        if not candidates:
            return None

        chosen = candidates[0] if len(candidates) == 1 else None
        if chosen is None:
            norm_title = _norm_title(pub_row.title)
            scored = [
                (c, len(set(_norm_title(c.topic)) & set(norm_title)))
                for c in candidates
            ]
            scored.sort(key=lambda x: -x[1])
            if len(scored) > 1 and scored[0][1] == scored[1][1]:
                log.add(
                    area, pub_row.title,
                    f"캠페인 플랜에서 같은 광고 시작일({md[0]}/{md[1]})의 항목이 {len(candidates)}건이라 "
                    "어느 것인지 특정하지 못해 캠페인비 비움",
                )
                return None
            chosen = scored[0][0]

        used.add(id(chosen))
        if chosen.budget is None:
            return None
        if _norm_title(chosen.topic) != _norm_title(pub_row.title):
            log.add(
                area, pub_row.title,
                f"캠페인 플랜의 광고 시작일({md[0]}/{md[1]})이 같아 캠페인비를 가져왔지만 "
                f"제목 표기가 다릅니다 (플랜: '{chosen.topic}') — 맞는지 확인해주세요",
            )
        return roundup_to_10k(chosen.budget)

    return match


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
    campaign_plan_rows: list | None = None,
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
        make_campaign_plan_fee_matcher(campaign_plan_rows, "유튜브", target_year),
    )
    _build_instagram(
        prev_wb, target_year, target_month, by_channel["인스타그램"], instagram_rows,
        instagram_campaign_rows, log, stats["인스타그램"],
        make_campaign_plan_fee_matcher(campaign_plan_rows, "인스타그램", target_year),
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
        ga_by_id = {}
    else:
        ga_by_id = {}
        for r in ga_rows:
            key = str(r.key).strip()
            # 프로토콜(https://) 없이 'news.lxsemicon.com/10287'나 '/10287'처럼 올 수도 있어,
            # URL 형식 전체가 아니라 '/' 뒤에 붙는 숫자(콘텐츠 번호)만 있으면 매칭 대상으로 삼는다.
            uid = extract_url_id(key, channel="뉴스룸") if "/" in key else None
            if uid is not None:
                ga_by_id[uid] = r.views

    new_rows = []
    for row in accepted:
        row_uid = extract_url_id(row.url, channel="뉴스룸")
        views = ga_by_id.get(row_uid) if row_uid is not None else None
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
                "G": row.content_type or CONFIRM_LABEL,
                "H": _or_confirm(views),
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
                "G": row.content_type or CONFIRM_LABEL,
                "H": _or_confirm(perf.views if perf else None),
                "I": _or_confirm(perf.likes if perf else None),
                "J": _or_confirm(perf.comments if perf else None),
            }
        )
    if new_rows:
        _append_generic(ws, new_rows)


def _build_youtube(prev_wb, year, month, publish_rows, studio_rows, campaign_rows, log, stat, plan_fee_matcher=None):
    ws = prev_wb[SHEET_NAMES["유튜브"]]
    if studio_rows is None:
        log.add("취합/유튜브_Raw Data", "전체", "유튜브 스튜디오 원본이 없어 성과를 채울 수 없음")
    video_ids = {r.video_id for r in (studio_rows or [])}
    result = filter_youtube_by_video_id(publish_rows, video_ids, month, "취합/유튜브_Raw Data")
    log.extend(result.confirm)
    accepted = result.accepted
    stat.total_published = len(accepted)

    studio_by_id = {r.video_id: r for r in (studio_rows or [])}

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
                    "G": row.content_type or CONFIRM_LABEL,
                    "H": CONFIRM_LABEL,
                    "I": CONFIRM_LABEL,
                    "J": CONFIRM_LABEL,
                    "K": CONFIRM_LABEL,
                    "L": CONFIRM_LABEL,
                    "M": CONFIRM_LABEL,
                    "N": CONFIRM_LABEL,
                    "O": CONFIRM_LABEL,
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

        fee = _find_campaign_fee(campaign_rows, row)
        if fee is None and plan_fee_matcher is not None:
            fee = plan_fee_matcher(row, "취합/유튜브_Raw Data", log)
        if fee is None:
            log.add("취합/유튜브_Raw Data", row.title, "AD 최종본·캠페인 플랜 어디에서도 캠페인비를 찾지 못해 비움")

        new_rows.append(
            {
                "C": format_kr_date(row.date_raw, year),
                "D": row.category or _need_category(log, "취합/유튜브_Raw Data", row.title),
                "E": row.title,
                "F": row.url,
                "G": row.content_type or CONFIRM_LABEL,
                "H": _or_confirm(fee),
                "I": _or_confirm(perf.views),
                "J": _or_confirm(length_slot),
                "K": _or_confirm(duration_slot),
                "L": _or_confirm(avg_pct),
                "M": _or_confirm(perf.likes),
                "N": _or_confirm(perf.comments),
                "O": _or_confirm(perf.shares),
            }
        )
    if new_rows:
        _append_generic(ws, new_rows)


def _build_instagram(prev_wb, year, month, publish_rows, insta_rows, campaign_rows, log, stat, plan_fee_matcher=None):
    ws = prev_wb[SHEET_NAMES["인스타그램"]]
    publish_rows = _filter_rows_by_month(publish_rows, month, "취합/인스타그램_Raw Data", log)
    stat.total_published = len(publish_rows)

    if insta_rows is None:
        log.add("취합/인스타그램_Raw Data", "전체", "인스타그램 성과 원본이 없어 성과를 채울 수 없음")
    matched = match_perf_rows(publish_rows, insta_rows, year, "취합/인스타그램_Raw Data", log)

    new_rows = []
    for row in publish_rows:
        perf = matched.get(id(row))
        if perf is not None:
            stat.matched += 1
        else:
            # 발행리스트에 있는 이번 달 발행건은 성과를 못 찾아도 반드시 행을 추가하고,
            # 성과 칸만 '확인 필요'로 남긴다.
            stat.unmatched_titles.append(row.title)
            if insta_rows is not None:
                log.add(
                    "취합/인스타그램_Raw Data", row.title,
                    "성과 원본에서 이 콘텐츠를 찾지 못해 성과 지표를 '확인 필요'로 남김",
                )

        fee = _find_campaign_fee(campaign_rows, row)
        if fee is None and plan_fee_matcher is not None:
            fee = plan_fee_matcher(row, "취합/인스타그램_Raw Data", log)
        if fee is None:
            log.add(
                "취합/인스타그램_Raw Data", row.title,
                "AD 최종본·캠페인 플랜 어디에서도 캠페인비를 찾지 못해 비움",
            )

        new_rows.append(
            {
                "C": _publish_date_str(row, perf, year),
                "D": row.category or _need_category(log, "취합/인스타그램_Raw Data", row.title),
                "E": row.title,
                "F": row.url or (perf.url if perf else None) or CONFIRM_LABEL,
                "G": row.content_type or CONFIRM_LABEL,
                "H": _or_confirm(fee),
                "I": _or_confirm(perf.reach if perf else None),
                "J": _or_confirm(perf.views if perf else None),
                "K": CONFIRM_LABEL,  # 프로필 방문: 원본에 없는 수기 입력 항목
                "L": CONFIRM_LABEL,  # 웹 방문: 원본에 없는 수기 입력 항목
                "M": _or_confirm(perf.likes if perf else None),
                "N": _or_confirm(perf.comments if perf else None),
                "O": _or_confirm(perf.saves if perf else None),
                "P": _or_confirm(perf.shares if perf else None),
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
    publish_rows = _filter_rows_by_month(publish_rows, month, "취합/페이스북_Raw Data", log)
    stat.total_published = len(publish_rows)

    if fb_rows is None:
        log.add("취합/페이스북_Raw Data", "전체", "페이스북 성과 원본이 없어 성과를 채울 수 없음")
    matched = match_perf_rows(publish_rows, fb_rows, year, "취합/페이스북_Raw Data", log)

    new_rows = []
    for row in publish_rows:
        perf = matched.get(id(row))
        if perf is not None:
            stat.matched += 1
        else:
            # 발행리스트에 있는 이번 달 발행건은 성과를 못 찾아도 반드시 행을 추가하고,
            # 성과 칸만 '확인 필요'로 남긴다.
            stat.unmatched_titles.append(row.title)
            if fb_rows is not None:
                log.add(
                    "취합/페이스북_Raw Data", row.title,
                    "성과 원본에서 이 콘텐츠를 찾지 못해 성과 지표를 '확인 필요'로 남김",
                )

        new_rows.append(
            {
                "C": _publish_date_str(row, perf, year),
                "D": row.category or _need_category(log, "취합/페이스북_Raw Data", row.title),
                "E": row.title,
                "F": row.url or (perf.url if perf else None) or CONFIRM_LABEL,
                "G": row.content_type or CONFIRM_LABEL,
                "H": "-",  # 페이스북은 광고 집행 시에만 값이 생김 (기본 '-')
                "I": _or_confirm(perf.reach if perf else None),
                "J": _or_confirm(perf.impressions if perf else None),
                "K": _or_confirm(perf.likes if perf else None),
                "L": _or_confirm(perf.comments if perf else None),
                "M": _or_confirm(perf.shares if perf else None),
            }
        )
    if new_rows:
        _append_generic(ws, new_rows)


CONFIRM_LABEL = "⚠ 확인 필요"


def _or_confirm(value):
    return value if value is not None else CONFIRM_LABEL


def _norm_url(u) -> str | None:
    """링크 비교용 정규화 — 프로토콜/www/뒤쪽 슬래시/쿼리스트링(?igsh=, ?si= 등)을 떼어낸다."""
    if not isinstance(u, str):
        return None
    s = u.strip()
    if not s:
        return None
    s = s.split("?")[0].split("#")[0]
    s = re.sub(r"^https?://", "", s, flags=re.IGNORECASE)
    s = re.sub(r"^www\.", "", s, flags=re.IGNORECASE)
    return s.rstrip("/").lower() or None


def _build_date_index(perf_rows: list) -> dict[datetime.date, list]:
    idx: dict[datetime.date, list] = {}
    for r in perf_rows:
        d = parse_posted_at_date(getattr(r, "posted_at", None))
        if d is not None:
            idx.setdefault(d, []).append(r)
    return idx


def match_perf_rows(
    publish_rows: list, perf_rows: list | None, year: int, area: str, log: ConfirmLog,
    date_tolerance_days: int = 1,
) -> dict[int, object]:
    """발행리스트 행들에 성과 원본을 한꺼번에 매칭해 {id(발행행): 성과행} 을 돌려준다.

    발행일이 정확히 같은 것을 먼저 전부 배정한 뒤에, 남은 행들만 ±1일 허용으로 재시도한다
    (메타 성과 데이터는 태평양시 기준이라 하루 오차가 날 수 있음). 이 순서가 중요한 이유는,
    발행리스트에 연도 표기가 없어 작년 동월 콘텐츠가 섞여 있는데 한 행씩 순서대로 ±1일을
    적용하면 작년 행이 올해 행의 데이터를 먼저 가로채기 때문이다.

    같은 날짜에 후보가 둘 이상이면 제목으로 좁히고, 그래도 안 되면 매칭하지 않는다.
    한 성과 원본 행은 한 번만 배정된다.
    """
    perf_rows = perf_rows or []
    by_title = {_norm_title(r.title): r for r in perf_rows}
    by_date = _build_date_index(perf_rows)
    used: set[int] = set()
    result: dict[int, object] = {}

    # 0단계: 고유 링크가 같으면 가장 확실한 근거이므로 날짜·제목보다 먼저 배정한다.
    by_url: dict[str, list] = {}
    for r in perf_rows:
        u = _norm_url(getattr(r, "url", None))
        if u:
            by_url.setdefault(u, []).append(r)
    for row in publish_rows:
        u = _norm_url(getattr(row, "url", None))
        if not u:
            continue
        candidates = [c for c in by_url.get(u, []) if id(c) not in used]
        if len(candidates) == 1:
            used.add(id(candidates[0]))
            result[id(row)] = candidates[0]

    def center_of(row):
        md = parse_month_day(row.date_raw)
        if md is None:
            return None
        try:
            return datetime.date(year, md[0], md[1])
        except ValueError:
            return None

    def pick(row, candidates, tolerance_note: bool):
        """날짜가 맞는 후보 중에서 고른다. 발행리스트에는 연도 표기가 없어 작년 동월 콘텐츠가
        섞여 있으므로, 날짜만 같고 제목이 전혀 다른 것은 배정하지 않는다(작년 행이 올해 데이터를
        가로채는 것을 막기 위함). 제목 표기가 조금 달라도(공백/리뉴얼 표시 등) 충분히 비슷하면 허용."""
        similar = [c for c in candidates if _title_similarity(c.title, row.title) >= TITLE_SIMILARITY_THRESHOLD]
        if not similar:
            return None
        if len(similar) == 1:
            return similar[0]
        exact = [c for c in similar if _norm_title(c.title) == _norm_title(row.title)]
        if len(exact) == 1:
            return exact[0]
        if tolerance_note:
            log.add(
                area, row.title,
                f"발행일 근처에 제목이 비슷한 성과 원본 후보가 {len(similar)}건이라 특정하지 못해 매칭하지 않음",
            )
        return None

    # 1단계: 발행일이 정확히 일치하고 제목까지 같은 것부터 배정한다.
    # (제목까지 보는 이유: 같은 날 발행된 콘텐츠가 둘인데 한쪽이 성과 원본에서는 하루 전으로
    #  잡혀 있으면, 제목을 안 보고 배정할 경우 두 콘텐츠의 데이터가 서로 뒤바뀐다.)
    for row in publish_rows:
        center = center_of(row)
        if center is None:
            continue
        exact = [
            c for c in by_date.get(center, [])
            if id(c) not in used and _norm_title(c.title) == _norm_title(row.title)
        ]
        if len(exact) == 1:
            used.add(id(exact[0]))
            result[id(row)] = exact[0]

    # 2단계: 발행일만 정확히 일치하는 것(제목은 달라도 됨)
    for row in publish_rows:
        if id(row) in result:
            continue
        center = center_of(row)
        if center is None:
            continue
        candidates = [c for c in by_date.get(center, []) if id(c) not in used]
        perf = pick(row, candidates, tolerance_note=False)
        if perf is not None:
            used.add(id(perf))
            result[id(row)] = perf

    # 3단계: 남은 행만 ±N일 허용
    for row in publish_rows:
        if id(row) in result:
            continue
        center = center_of(row)
        if center is None:
            continue
        candidates, seen = [], set()
        for delta in range(-date_tolerance_days, date_tolerance_days + 1):
            if delta == 0:
                continue
            for c in by_date.get(center + datetime.timedelta(days=delta), []):
                if id(c) not in used and id(c) not in seen:
                    candidates.append(c)
                    seen.add(id(c))
        perf = pick(row, candidates, tolerance_note=True)
        if perf is not None:
            used.add(id(perf))
            result[id(row)] = perf

    # 4단계: 날짜로 전혀 못 찾은 행은 제목으로만 시도(성과 원본에 게시 시각이 없는 경우 등)
    for row in publish_rows:
        if id(row) in result:
            continue
        cand = by_title.get(_norm_title(row.title))
        if cand is not None and id(cand) not in used:
            used.add(id(cand))
            result[id(row)] = cand

    # 5단계: 제목을 아예 비교할 수 없는 원본일 때의 마지막 수단.
    # 플랫폼에서 바로 받은 페이스북/인스타그램 원본은 제목 칸에 캡션이 들어 있고 링크 형태도
    # 발행리스트와 달라(share/p/... vs posts/pfbid...) 링크·제목이 모두 무용지물이다.
    # 이때는 발행일이 같고 후보가 하나뿐인 경우에만 배정하되, 발행리스트에 연도 표기가 없어
    # 작년 동월 행이 앞쪽에 섞여 있으므로 '나중에 추가된 행'부터(역순) 가져가게 한다.
    for row in reversed(publish_rows):
        if id(row) in result:
            continue
        center = center_of(row)
        if center is None:
            continue
        for delta in (0, -1, 1):
            candidates = [
                c for c in by_date.get(center + datetime.timedelta(days=delta), [])
                if id(c) not in used
            ]
            if len(candidates) == 1:
                used.add(id(candidates[0]))
                result[id(row)] = candidates[0]
                log.add(
                    area, row.title,
                    "성과 원본에 비교할 제목·링크가 없어 발행일만으로 매칭했습니다 — 맞는지 확인해주세요",
                )
                break

    return result


def _publish_date_str(row, perf, year: int) -> str:
    """발행리스트의 발행일자를 기준으로 쓴다. 성과 원본(메타)의 게시 시각은 태평양시(PT) 기준이라
    한국 시간 기준 발행일과 하루 어긋나 보일 수 있으므로 기준으로 삼지 않는다 — 발행리스트가
    없을 때만(이 경우는 애초에 이 함수가 호출되지 않지만, 방어적으로) 성과 원본 값으로 보완한다."""
    if row.date_raw is not None:
        return format_kr_date(row.date_raw, year)
    posted_date = parse_posted_at_date(getattr(perf, "posted_at", None))
    if posted_date is not None:
        return format_kr_date(posted_date, year)
    return ""


def _need_category(log: ConfirmLog, area: str, title: str) -> None:
    log.add(area, title, "발행리스트에 카테고리가 없어 수기 분류 필요")
    return None
