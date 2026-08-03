"""파서 공용 유틸.

원본 파일 형식이 확인된 것(인스타/페북/유튜브 스튜디오)은 컬럼명을 그대로 하드코딩해서 매칭하고,
형식이 확인되지 않은 것(발행리스트/블로그/버즈빌/캠페인 플랜)은 여기 있는 '유연 헤더 탐색' 도구로
후보 이름들을 찾되, 못 찾으면 절대 추정하지 않고 명확한 에러를 낸다.
"""

import re

import openpyxl

URL_RE = re.compile(r"https?://\S+")

CHANNEL_DOMAINS = {
    "뉴스룸": ("news.lxsemicon.com",),
    "유튜브": ("youtube.com", "youtu.be"),
    "인스타그램": ("instagram.com",),
    "페이스북": ("facebook.com", "fb.com"),
    "블로그": ("blog.naver.com",),
}


class ParseError(Exception):
    """원본 파일 형식을 인식하지 못했을 때 — 절대 임의로 매핑하지 않고 여기서 멈춘다."""


def classify_channel_by_url(url: str) -> str | None:
    if not isinstance(url, str):
        return None
    for channel, domains in CHANNEL_DOMAINS.items():
        if any(domain in url for domain in domains):
            return channel
    return None


def extract_url_id(url: str, channel: str) -> str | None:
    """뉴스룸/블로그 URL 끝의 콘텐츠 번호(정렬·대역 비교용)를 뽑는다."""
    if not isinstance(url, str):
        return None
    m = re.search(r"(\d+)/?$", url.strip())
    return m.group(1) if m else None


def extract_youtube_video_id(url: str) -> str | None:
    if not isinstance(url, str):
        return None
    patterns = [
        r"youtu\.be/([A-Za-z0-9_-]{6,})",
        r"youtube\.com/shorts/([A-Za-z0-9_-]{6,})",
        r"[?&]v=([A-Za-z0-9_-]{6,})",
    ]
    for p in patterns:
        m = re.search(p, url)
        if m:
            return m.group(1)
    return None


def find_header_row(rows: list[list], candidates: list[str], max_scan: int = 15) -> int | None:
    """상단 max_scan행 안에서 candidates 중 하나라도 포함된 행을 헤더 행으로 추정."""
    for i, row in enumerate(rows[:max_scan]):
        texts = {str(c).strip() for c in row if c is not None}
        if texts & set(candidates):
            return i
    return None


def find_column(header: list, candidates: list[str]) -> int | None:
    """헤더 행에서 candidates 중 하나와 정확히 일치하는 열의 인덱스를 찾는다."""
    normalized = [str(h).strip() if h is not None else "" for h in header]
    for cand in candidates:
        if cand in normalized:
            return normalized.index(cand)
    return None


def load_sheet_rows(file, sheet_name: str | None = None) -> tuple[str, list[list]]:
    """xlsx 파일을 열어 (시트명, 행 리스트) 반환. sheet_name이 없으면 첫 시트."""
    wb = openpyxl.load_workbook(file, data_only=True)
    name = sheet_name or wb.sheetnames[0]
    ws = wb[name]
    rows = [[c.value for c in row] for row in ws.iter_rows()]
    return name, rows


def find_sheet_by_pattern(file, name_regex: str) -> str:
    wb = openpyxl.load_workbook(file, data_only=True, read_only=True)
    for name in wb.sheetnames:
        if re.search(name_regex, name):
            return name
    raise ParseError(
        f"'{name_regex}' 패턴에 맞는 시트를 찾지 못했습니다. 시트 목록: {wb.sheetnames}"
    )
