"""인스타그램 성과 원본 (메타 콘텐츠 인사이트 xlsx) 파서.

실제 6월 샘플로 확인된 형식:
  - 시트명이 날짜로 가변 (예: 'Jun-01-2026_Jun-30-2026_1032429') -> 정규식으로 탐색
  - 컬럼: NO, 발행리스트(제목), 설명, 게시 시간, 고유 링크, 게시물 유형,
          도달, 조회, 좋아요, 댓글, 저장, 공유, 팔로우
"""

from dataclasses import dataclass

from .base import ParseError, find_column, find_sheet_by_pattern, load_sheet_rows

HEADER_CANDIDATES = ["NO", "발행리스트"]

COLS = {
    "title": ["발행리스트"],
    "posted_at": ["게시 시간"],
    "url": ["고유 링크"],
    "post_type": ["게시물 유형"],
    "reach": ["도달"],
    "views": ["조회"],
    "likes": ["좋아요"],
    "comments": ["댓글"],
    "saves": ["저장"],
    "shares": ["공유"],
}


@dataclass
class InstagramRow:
    title: str
    url: str | None
    reach: float | None
    views: float | None
    likes: float | None
    comments: float | None
    saves: float | None
    shares: float | None


def parse_instagram(file) -> list[InstagramRow]:
    sheet_name = find_sheet_by_pattern(file, r"^[A-Za-z]{3}-\d{2}-\d{4}_[A-Za-z]{3}-\d{2}-\d{4}")
    file.seek(0)
    _, rows = load_sheet_rows(file, sheet_name)

    header_idx = None
    for i, row in enumerate(rows[:10]):
        texts = {str(c).strip() for c in row if c is not None}
        if set(HEADER_CANDIDATES) <= texts:
            header_idx = i
            break
    if header_idx is None:
        raise ParseError(
            f"인스타그램 성과 원본에서 헤더 행(NO, 발행리스트 포함)을 찾지 못했습니다. "
            f"시트: {sheet_name}"
        )
    header = rows[header_idx]

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"인스타그램 성과 원본에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    out: list[InstagramRow] = []
    for row in rows[header_idx + 1 :]:
        title = row[col_idx["title"]]
        if not title:
            continue
        out.append(
            InstagramRow(
                title=str(title).strip(),
                url=row[col_idx["url"]],
                reach=row[col_idx["reach"]],
                views=row[col_idx["views"]],
                likes=row[col_idx["likes"]],
                comments=row[col_idx["comments"]],
                saves=row[col_idx["saves"]],
                shares=row[col_idx["shares"]],
            )
        )
    return out
