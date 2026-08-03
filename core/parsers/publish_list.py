"""콘텐츠 발행리스트 파서.

⚠ 이 파일은 실제 샘플을 아직 확보하지 못한 상태로 작성되었다. 컬럼명을 하드코딩하는 대신
후보 이름들을 넓게 스캔하고, URL 열은 이름이 아니라 '셀 값이 실제로 URL인지'로 탐지한다
(어느 채널 URL인지는 도메인으로 자동 분류되므로 채널 열이 따로 없어도 동작한다).
필수 정보를 찾지 못하면 절대 추정하지 않고 ParseError로 무엇을 못 찾았는지 알린다.
"""

from dataclasses import dataclass

from .base import URL_RE, ParseError, classify_channel_by_url, find_column, load_sheet_rows

DATE_CANDIDATES = ["발행일", "게시일", "날짜", "게시일자", "발행일자"]
CATEGORY_CANDIDATES = ["카테고리", "구분"]
TITLE_CANDIDATES = ["콘텐츠 주제", "콘텐츠명", "제목", "콘텐츠", "주제"]
TYPE_CANDIDATES = ["유형", "형태", "콘텐츠 유형", "콘텐츠 형태"]
CHANNEL_CANDIDATES = ["채널", "플랫폼"]


@dataclass
class PublishListRow:
    channel: str
    date_raw: object
    category: str | None
    title: str
    url: str
    content_type: str | None


def _find_url_column(header: list, rows: list[list]) -> int | None:
    n_cols = len(header)
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


def parse_publish_list(file) -> list[PublishListRow]:
    _, rows = load_sheet_rows(file)
    if not rows:
        raise ParseError("콘텐츠 발행리스트 파일이 비어 있습니다.")

    header_idx = None
    for i, row in enumerate(rows[:15]):
        texts = {str(c).strip() for c in row if c is not None}
        if texts & set(TITLE_CANDIDATES) or texts & set(DATE_CANDIDATES):
            header_idx = i
            break
    if header_idx is None:
        raise ParseError(
            "콘텐츠 발행리스트에서 헤더 행을 찾지 못했습니다. "
            f"다음 이름들 중 하나라도 포함된 행이 있어야 합니다: "
            f"{DATE_CANDIDATES + TITLE_CANDIDATES}"
        )
    header = rows[header_idx]
    data_rows = rows[header_idx + 1 :]

    url_col = _find_url_column(header, data_rows)
    if url_col is None:
        raise ParseError(
            "콘텐츠 발행리스트에서 URL이 들어있는 열을 찾지 못했습니다. "
            "각 행에 콘텐츠 URL(뉴스룸/유튜브/인스타/페북/블로그 링크)이 있어야 합니다."
        )

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
        raise ParseError(
            "콘텐츠 발행리스트에서 다음 필수 열을 찾지 못했습니다:\n- " + "\n- ".join(missing)
        )

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
        else:
            channel = classify_channel_by_url(url)
        if channel is None:
            continue  # 어느 채널인지 판별 불가 -> 이 행은 스킵 (호출 측에서 URL 매칭 실패로 확인 필요 처리)

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
