"""fn 파일 서식에 맞춘 값 변환 도우미들 (요일 표기, 초→h:mm 슬롯, 캠페인비 반올림)."""

import datetime
import re

_WEEKDAYS_KR = ["월", "화", "수", "목", "금", "토", "일"]


def format_kr_date(date_raw, fallback_year: int | None = None) -> str:
    """'1/8(목)' 형식 문자열로 통일한다. datetime이면 요일을 계산해서 만들고,
    이미 그 형식의 문자열이면 그대로 둔다."""
    if isinstance(date_raw, datetime.datetime):
        d = date_raw.date()
    elif isinstance(date_raw, datetime.date):
        d = date_raw
    else:
        s = str(date_raw).strip()
        # 이미 'M/D(요일)' 형식이면 그대로 사용
        if re.match(r"^\d{1,2}/\d{1,2}\(.+\)$", s):
            return s
        # 'M/D' 형식이면 요일만 붙여준다(연도가 있어야 요일 계산 가능하므로 fallback_year 필요)
        m = re.match(r"^(\d{1,2})/(\d{1,2})$", s)
        if m and fallback_year:
            d = datetime.date(fallback_year, int(m.group(1)), int(m.group(2)))
        else:
            return s  # 형식을 모르면 원본 그대로 두고 상위에서 확인 필요 처리
    weekday = _WEEKDAYS_KR[d.weekday()]
    return f"{d.month}/{d.day}({weekday})"


def seconds_to_hmm_slot(total_seconds) -> datetime.time | None:
    """98초 -> time(1,38) 처럼, '분:초'를 h:mm 서식 슬롯에 담는다(문서 규칙 6).
    분이 24 이상이면 datetime.time 표현이 불가능하므로 None을 반환해 호출측이 확인 필요로 처리하게 한다."""
    if total_seconds is None:
        return None
    try:
        total = int(round(float(total_seconds)))
    except (TypeError, ValueError):
        return None
    minutes, seconds = divmod(total, 60)
    if minutes >= 24:
        return None
    return datetime.time(hour=minutes, minute=seconds)


def hms_string_to_hmm_slot(hms: str) -> datetime.time | None:
    """'0:00:40' (H:MM:SS) 형식의 평균 시청 지속 시간을 초로 바꾼 뒤 h:mm 슬롯으로 변환."""
    if not hms:
        return None
    parts = str(hms).strip().split(":")
    try:
        parts = [int(p) for p in parts]
    except ValueError:
        return None
    if len(parts) == 3:
        h, m, s = parts
    elif len(parts) == 2:
        h, m, s = 0, parts[0], parts[1]
    else:
        return None
    total_seconds = h * 3600 + m * 60 + s
    return seconds_to_hmm_slot(total_seconds)


def parse_posted_at_date(posted_at) -> datetime.date | None:
    """성과 원본의 게시 시간('06/24/2026 22:12' 같은 MM/DD/YYYY HH:MM 문자열 또는 datetime)에서
    날짜만 뽑는다. 발행리스트의 '예정일'보다 실제 게시된 이 값이 더 정확한 경우가 있다
    (예: 야간 게시로 발행리스트 예정일과 하루 차이가 나는 경우 실제 fn 파일은 이 값을 따름)."""
    if posted_at is None:
        return None
    if isinstance(posted_at, datetime.datetime):
        return posted_at.date()
    if isinstance(posted_at, datetime.date):
        return posted_at
    s = str(posted_at).strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})", s)  # MM/DD/YYYY ...
    if m:
        month, day, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
        try:
            return datetime.date(year, month, day)
        except ValueError:
            return None
    return None


def parse_month_day(date_raw) -> tuple[int, int] | None:
    """발행리스트('6/2(화)' 문자열 또는 datetime)나 성과 원본 게시시간
    ('06/02/2026 01:03' 같은 MM/DD/YYYY 문자열)에서 (월, 일)만 뽑아 서로 비교할 수 있게 한다."""
    if date_raw is None:
        return None
    if isinstance(date_raw, datetime.datetime):
        return date_raw.month, date_raw.day
    if isinstance(date_raw, datetime.date):
        return date_raw.month, date_raw.day
    s = str(date_raw).strip()
    m = re.match(r"^(\d{1,2})/(\d{1,2})/\d{4}", s)  # MM/DD/YYYY ...
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.match(r"^(\d{1,2})/(\d{1,2})", s)  # M/D(요일) 또는 M/D
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


def roundup_to_10k(amount) -> int | None:
    """캠페인비 = 실지출을 만원(10,000원) 단위로 올림. 179,999 -> 180,000."""
    if amount is None:
        return None
    try:
        amount = float(amount)
    except (TypeError, ValueError):
        return None
    import math

    return int(math.ceil(amount / 10_000) * 10_000)
