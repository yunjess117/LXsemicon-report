"""'이번 달(당해 연·월) 콘텐츠만 사용' 판별 로직.

지침 문서의 "URL 번호 대역"은 하드코딩된 상수가 아니라, 발행리스트에 연도가 없어서(예: '7/1(화)')
URL 번호가 오름차순으로 매겨지는 성질을 이용해 연도를 구분하는 것이다. 그래서 이 모듈은 번호 대역을
코드에 박아두지 않고, 매번 업로드된 "전월 fn 파일"의 마지막 블록에서 그 채널의 최대 URL 번호를
읽어와 그 값을 기준(floor)으로 삼는다.

판별이 애매한 경우(번호는 기준을 넘는데 날짜 월이 다르다, 반대로 날짜 월은 맞는데 번호가 기준
이하다 등)는 자동으로 포함하거나 제외하지 않고 확인 필요로 남긴다.
"""

import re
from dataclasses import dataclass

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
    from openpyxl.utils import column_index_from_string

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
    """뉴스룸/블로그: URL 번호가 전월 최대치보다 크고, 발행일의 월이 이번 달과 같은 것만 채택."""
    log = ConfirmLog()
    accepted = []
    for row in rows:
        uid = extract_url_id(row.url, channel="")
        month = _parse_month_from_kr_date(row.date_raw)

        if prev_max_id is None:
            log.add(area_label, row.url, "전월 fn 파일에서 이 채널의 최대 URL 번호를 읽지 못해 연·월 판별 불가")
            continue
        if uid is None:
            log.add(area_label, row.url, "URL에서 번호를 추출하지 못해 연·월 판별 불가")
            continue

        uid_n = int(uid)
        band_ok = uid_n > prev_max_id
        month_ok = month == target_month

        if band_ok and month_ok:
            accepted.append(row)
        elif band_ok and not month_ok:
            log.add(
                area_label,
                row.url,
                f"URL 번호({uid_n})는 전월 최대치({prev_max_id})보다 크지만 발행일의 월({month})이 "
                f"이번 달({target_month})과 달라 자동 포함하지 않음",
            )
        elif month_ok and not band_ok:
            log.add(
                area_label,
                row.url,
                f"발행일의 월은 이번 달과 같지만 URL 번호({uid_n})가 전월 최대치({prev_max_id}) "
                "이하라 작년 동월 콘텐츠일 수 있어 자동 포함하지 않음",
            )
        # 둘 다 불일치하면 이번 달 콘텐츠가 아닌 것으로 보고 조용히 제외 (확인 필요 아님)
    return MonthFilterResult(accepted=accepted, confirm=log)


def filter_youtube_by_video_id(
    rows: list,
    studio_video_ids: set[str],
    area_label: str,
) -> MonthFilterResult:
    """유튜브: URL로 연·월을 구분할 수 없으므로, 스튜디오 export에 video_id가 있는 것만 채택."""
    from .parsers.base import extract_youtube_video_id

    log = ConfirmLog()
    accepted = []
    for row in rows:
        vid = extract_youtube_video_id(row.url)
        if vid is None:
            log.add(area_label, row.url, "URL에서 video_id를 추출하지 못함")
            continue
        if vid in studio_video_ids:
            accepted.append(row)
        # 스튜디오 export에 없는 video_id는 이번 달 발행분이 아닌 것으로 보고 조용히 제외
    return MonthFilterResult(accepted=accepted, confirm=log)
