"""'이번 달(당해 연·월) 콘텐츠만 사용' 판별 로직.

발행리스트에는 해당 연도 발행분만 들어 있다는 전제로, 발행일의 월이 이번 달과 같으면
무조건 채택한다. 로우데이터 시트에는 발행리스트에 있는 이번 달 발행건이 빠짐없이 들어가야
하고, 성과 원본에서 값을 못 찾은 칸만 '확인 필요'로 남기는 것이 원칙이다.

URL 번호 대역(뉴스룸/블로그)과 유튜브 video_id는 이제 제외 기준이 아니라 '이상 징후 알림'
용도로만 쓴다 — 예전에는 발행리스트에 작년 동월 콘텐츠가 섞여 있어 이것들로 연도를
구분해야 했지만, 지금은 그런 행이 없기 때문이다.
"""

import re
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string
from openpyxl.worksheet.worksheet import Worksheet

from .confirm import ConfirmLog
from .excel_block import MonthBlock, find_month_blocks
from .parsers.base import extract_url_id


def get_prev_max_url_id(ws: Worksheet, url_col: str, blocks: list[MonthBlock] | None = None) -> int | None:
    """전월 fn 시트의 마지막 블록에서 URL 열의 최대 번호를 읽는다."""
    blocks = blocks or find_month_blocks(ws)
    if not blocks:
        return None
    last = blocks[-1]
    col_idx = column_index_from_string(url_col)
    max_id = None
    for r in range(last.data_start_row, last.data_end_row + 1):
        url = ws.cell(row=r, column=col_idx).value
        uid = extract_url_id(url, channel="") if url else None
        if uid is not None:
            n = int(uid)
            max_id = n if max_id is None else max(max_id, n)
    return max_id


def _parse_month_from_kr_date(date_raw) -> int | None:
    """'7/1(화)' 같은 문자열 또는 datetime에서 월(1~12)만 뽑는다."""
    if date_raw is None:
        return None
    if hasattr(date_raw, "month"):
        return date_raw.month
    m = re.match(r"\s*(\d{1,2})\s*/\s*\d{1,2}", str(date_raw))
    if m:
        month = int(m.group(1))
        if 1 <= month <= 12:
            return month
    return None


@dataclass
class MonthFilterResult:
    accepted: list  # PublishListRow 목록 (채택)
    confirm: ConfirmLog


def filter_by_url_band(
    rows: list,
    prev_max_id: int | None,
    target_month: int,
    area_label: str,
) -> MonthFilterResult:
    """뉴스룸/블로그: 발행일의 월이 이번 달이면 모두 채택한다.

    URL 번호가 전월 최대치보다 작으면 예전 콘텐츠가 섞였을 수 있다는 뜻이므로 알림만 남기고,
    행 자체는 그대로 추가한다(발행리스트에 있는 이번 달 발행건은 빠뜨리지 않는다)."""
    log = ConfirmLog()
    accepted = []
    for row in rows:
        if _parse_month_from_kr_date(row.date_raw) != target_month:
            continue

        accepted.append(row)

        uid = extract_url_id(row.url, channel="")
        if uid is not None and prev_max_id is not None and int(uid) <= prev_max_id:
            log.add(
                area_label,
                row.title or row.url,
                f"URL 번호({uid})가 전월 최대치({prev_max_id})보다 작습니다 — "
                "예전에 발행한 콘텐츠가 아닌지 확인해주세요(행은 그대로 추가했습니다)",
            )
    return MonthFilterResult(accepted=accepted, confirm=log)


def filter_youtube_by_video_id(
    rows: list,
    studio_video_ids: set[str],
    target_month: int,
    area_label: str,
) -> MonthFilterResult:
    """유튜브: 발행일의 월이 이번 달이면 모두 채택한다.
    스튜디오 export에 video_id가 없으면 성과만 못 채우는 것이므로 알림만 남긴다."""
    from .parsers.base import extract_youtube_video_id

    log = ConfirmLog()
    accepted = []
    for row in rows:
        if _parse_month_from_kr_date(row.date_raw) != target_month:
            continue

        accepted.append(row)

        vid = extract_youtube_video_id(row.url)
        if vid is None:
            log.add(area_label, row.title or row.url, "URL에서 video_id를 추출하지 못해 성과를 채우지 못함")
        elif studio_video_ids and vid not in studio_video_ids:
            log.add(
                area_label, row.title or row.url,
                "유튜브 스튜디오 원본에 이 영상이 없어 성과를 '확인 필요'로 남김",
            )
    return MonthFilterResult(accepted=accepted, confirm=log)
