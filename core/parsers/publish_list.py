"""콘텐츠 발행리스트 파서.

실제 샘플로 확인된 형식: 채널별로 시트가 나뉘어 있다(뉴스룸/블로그/페이스북/인스타그램/유튜브).
시트 이름 자체가 채널이므로 그걸 그대로 쓰고, 시트마다 헤더 행(NO/발행일/카테고리/콘텐츠 주제/
URL/유형)의 정확한 위치를 개별적으로 찾는다. 필수 정보를 찾지 못한 시트는 절대 추정하지 않고
확인 필요로 남긴다(호출 측에 알 수 있도록 로그를 남기고 그 시트는 건너뜀).
"""

from dataclasses import dataclass

import openpyxl

from .base import URL_RE, ParseError, classify_channel_by_url, find_column

DATE_CANDIDATES = ["발행일", "게시일", "날짜", "게시일자", "발행일자"]
CATEGORY_CANDIDATES = ["카테고리", "구분"]
TITLE_CANDIDATES = ["콘텐츠 주제", "콘텐츠명", "제목", "콘텐츠", "주제"]
TYPE_CANDIDATES = ["유형", "형태", "콘텐츠 유형", "콘텐츠 형태"]
CHANNEL_CANDIDATES = ["채널", "플랫폼"]

KNOWN_CHANNELS = ["뉴스룸", "유튜브", "인스타그램", "페이스북", "블로그"]


@dataclass
class PublishListRow:
    channel: str
    date_raw: object
    category: str | None
    title: str
    url: str
    content_type: str | None


def _find_url_column(rows: list[list]) -> int | None:
    n_cols = max((len(r) for r in rows[:50]), default=0)
    best_col, best_hits = None, 0
    for col in range(n_cols):
        hits = 0
        checked = 0
        for row in rows[:50]:
            if col >= len(row):
                continue
            v = row[col]
            if v is None:
                continue
            checked += 1
            if isinstance(v, str) and URL_RE.match(v.strip()):
                hits += 1
        if checked and hits / checked > 0.6 and hits > best_hits:
            best_col, best_hits = col, hits
    return best_col


def _parse_sheet_rows(rows: list[list], default_channel: str | None) -> list[PublishListRow]:
    """한 시트(=보통 한 채널) 안의 행들을 파싱한다. 여러 채널 블록이 섞여 있어도(예: 시트
    안에서 '뉴스룸' 같은 라벨이 다시 나오는 경우는 없다고 가정) 헤더가 반복되지 않는 한
    전체를 한 번에 처리한다."""
    header_idx = None
    for i, row in enumerate(rows[:15]):
        texts = {str(c).strip() for c in row if c is not None}
        if texts & set(TITLE_CANDIDATES) or texts & set(DATE_CANDIDATES):
            header_idx = i
            break
    if header_idx is None:
        raise ParseError(
            "헤더 행을 찾지 못했습니다. "
            f"다음 이름들 중 하나라도 포함된 행이 있어야 합니다: {DATE_CANDIDATES + TITLE_CANDIDATES}"
        )
    header = rows[header_idx]
    data_rows = rows[header_idx + 1 :]

    url_col = _find_url_column(data_rows)
    if url_col is None:
        raise ParseError("URL이 들어있는 열을 찾지 못했습니다.")

    date_col = find_column(header, DATE_CANDIDATES)
    category_col = find_column(header, CATEGORY_CANDIDATES)
    title_col = find_column(header, TITLE_CANDIDATES)
    type_col = find_column(header, TYPE_CANDIDATES)
    channel_col = find_column(header, CHANNEL_CANDIDATES)

    missing = []
    if date_col is None:
        missing.append(f"발행일 (후보: {DATE_CANDIDATES})")
    if title_col is None:
        missing.append(f"콘텐츠 제목 (후보: {TITLE_CANDIDATES})")
    if missing:
        raise ParseError("다음 필수 열을 찾지 못했습니다:\n- " + "\n- ".join(missing))

    out: list[PublishListRow] = []
    for row in data_rows:
        if url_col >= len(row):
            continue
        url = row[url_col]
        if not url or not isinstance(url, str) or not URL_RE.match(url.strip()):
            continue
        channel = None
        if channel_col is not None and channel_col < len(row) and row[channel_col]:
            channel = str(row[channel_col]).strip()
        elif default_channel is not None:
            channel = default_channel
        else:
            channel = classify_channel_by_url(url)
        if channel is None:
            continue

        out.append(
            PublishListRow(
                channel=channel,
                date_raw=row[date_col] if date_col is not None and date_col < len(row) else None,
                category=(
                    str(row[category_col]).strip()
                    if category_col is not None and category_col < len(row) and row[category_col]
                    else None
                ),
                title=str(row[title_col]).strip() if title_col < len(row) and row[title_col] else "",
                url=url.strip(),
                content_type=(
                    str(row[type_col]).strip()
                    if type_col is not None and type_col < len(row) and row[type_col]
                    else None
                ),
            )
        )
    return out


def parse_publish_list(file) -> list[PublishListRow]:
    wb = openpyxl.load_workbook(file, data_only=True)

    # 실제 파일은 채널별로 시트가 나뉘어 있다(뉴스룸/블로그/페이스북/인스타그램/유튜브).
    # 시트 이름이 알려진 채널명을 포함하면 그 시트 전체를 그 채널로 간주해서 파싱한다.
    channel_sheets = [(name, ch) for name in wb.sheetnames for ch in KNOWN_CHANNELS if ch in name]

    out: list[PublishListRow] = []
    errors: list[str] = []

    if channel_sheets:
        for sheet_name, channel in channel_sheets:
            ws = wb[sheet_name]
            rows = [[c.value for c in row] for row in ws.iter_rows()]
            try:
                out.extend(_parse_sheet_rows(rows, default_channel=channel))
            except ParseError as e:
                errors.append(f"'{sheet_name}' 시트: {e}")
    else:
        # 채널별 시트 구분이 없는 단일 시트 형식 (URL 도메인으로 채널 판별)
        ws = wb[wb.sheetnames[0]]
        rows = [[c.value for c in row] for row in ws.iter_rows()]
        try:
            out.extend(_parse_sheet_rows(rows, default_channel=None))
        except ParseError as e:
            errors.append(f"'{ws.title}' 시트: {e}")

    if not out and errors:
        raise ParseError("콘텐츠 발행리스트를 읽지 못했습니다:\n- " + "\n- ".join(errors))
    return out
