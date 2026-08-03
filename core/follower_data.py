"""채널 팔로워 시트 — 일별 로그 자동 채움.

담당자가 별도로 운영하는 "SNS 채널 Daily data" 워크북은 fn 파일의 '채널 팔로워' 시트와
행 번호까지 동일한 일별 로그 구조를 공유한다(같은 날짜가 같은 행 번호에 위치). 이 성질을
이용해, daily 파일에서 이번 달 날짜에 해당하는 실측값만 fn 파일의 같은 행에 그대로 옮겨 채운다.

'월간 SNS 운영 수치'/'채널 팔로워'의 월간 요약 표(상단)는 채널마다 다른(때로는 깨진) 수식
패턴이 섞여 있어 자동 채움 위험이 크므로 건드리지 않고, 일별 로그만 안전하게 채운다.
"""

import datetime

from openpyxl.workbook import Workbook

from .confirm import ConfirmLog

SHEET = "채널 팔로워"
DATE_COL = 3  # C
MONTH_COL = 2  # B

# 값 열(그대로 복사) : 열문자 -> daily 파일에서의 같은 열문자 (구조가 동일하므로 1:1)
VALUE_COLS = ["D", "G", "J", "M", "P", "S", "U", "W"]
# 각 값 열 바로 오른쪽의 '증감' 열
DIFF_COLS = {"D": "E", "G": "H", "J": "K", "M": "N", "P": "Q", "S": "T", "U": "V", "W": "X"}

from openpyxl.utils import column_index_from_string


def _col(letter):
    return column_index_from_string(letter)


def fill_daily_followers(
    prev_wb: Workbook,
    daily_wb: Workbook,
    target_year: int,
    target_month: int,
    log: ConfirmLog,
) -> int:
    """daily_wb의 이번 달 날짜 행을 prev_wb의 '채널 팔로워' 일별 로그에 채운다.
    채운 행 수를 반환한다."""
    if SHEET not in prev_wb.sheetnames or SHEET not in daily_wb.sheetnames:
        log.add("취합/채널 팔로워", "전체", "채널 팔로워 시트를 찾지 못해 일별 로그를 채우지 못함")
        return 0

    fn_ws = prev_wb[SHEET]
    daily_ws = daily_wb[SHEET]

    # daily 파일에서 날짜 -> 행 번호 매핑 (일별 로그 영역만: MONTH_COL에 '월' 텍스트가 있는 행)
    daily_by_date: dict[datetime.date, int] = {}
    for row in daily_ws.iter_rows(min_row=30):
        month_cell = row[MONTH_COL - 1]
        date_cell = row[DATE_COL - 1]
        if not (isinstance(month_cell.value, str) and month_cell.value.endswith("월")):
            continue
        if isinstance(date_cell.value, datetime.datetime):
            d = date_cell.value.date()
        elif isinstance(date_cell.value, datetime.date):
            d = date_cell.value
        else:
            continue
        daily_by_date[d] = date_cell.row

    # fn 파일에서 날짜 -> 행 번호 매핑 (동일한 방식)
    fn_by_date: dict[datetime.date, int] = {}
    for row in fn_ws.iter_rows(min_row=30):
        month_cell = row[MONTH_COL - 1]
        date_cell = row[DATE_COL - 1]
        if not (isinstance(month_cell.value, str) and month_cell.value.endswith("월")):
            continue
        if isinstance(date_cell.value, datetime.datetime):
            d = date_cell.value.date()
        elif isinstance(date_cell.value, datetime.date):
            d = date_cell.value
        else:
            continue
        fn_by_date[d] = date_cell.row

    # 이번 달 말일까지의 데이터를 전부 그대로 옮긴다(이전 달이 비어 있었어도 함께 채워진다).
    if target_month == 12:
        month_end = datetime.date(target_year, 12, 31)
    else:
        month_end = datetime.date(target_year, target_month + 1, 1) - datetime.timedelta(days=1)

    filled = 0
    for d in sorted(daily_by_date):
        if d > month_end:
            continue
        src_row = daily_by_date[d]
        dst_row = fn_by_date.get(d)
        if dst_row is None:
            log.add(
                "취합/채널 팔로워", d.isoformat(),
                "fn 파일의 일별 로그에 이 날짜 행이 없어(템플릿 미리 생성 안 됨) 건너뜀 — 수기 확인 필요",
            )
            continue

        any_value = False
        for vcol in VALUE_COLS:
            for col in (vcol, DIFF_COLS[vcol]):
                src_val = daily_ws.cell(row=src_row, column=_col(col)).value
                if src_val in (None, ""):
                    continue
                fn_ws.cell(row=dst_row, column=_col(col)).value = src_val
                if col == vcol and src_val != "-":
                    any_value = True

        if any_value:
            filled += 1

    if filled == 0:
        log.add(
            "취합/채널 팔로워", "전체",
            "Daily data 파일에서 이번 달 말일까지의 실측값을 찾지 못함",
        )

    log.add(
        "취합/채널 팔로워", "월간 요약표(상단)",
        "일별 로그만 자동으로 채웠습니다. 상단 월간 요약표(Monthly Data)는 수식이 채널마다 달라 "
        "자동 채움 대상이 아니므로 직접 채워주세요.",
    )
    return filled
