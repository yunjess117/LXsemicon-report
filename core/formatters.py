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
