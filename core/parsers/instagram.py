"""인스타그램 성과 원본 (메타 콘텐츠 인사이트) 파서. xlsx와 csv 둘 다 받는다.

실제 6월 샘플로 확인된 형식:
  - xlsx는 시트명이 날짜로 가변 (예: 'Jun-01-2026_Jun-30-2026_1032429')
  - 컬럼: NO, 발행리스트(제목), 설명, 게시 시간, 고유 링크, 게시물 유형,
          도달, 조회, 좋아요, 댓글, 저장, 공유, 팔로우
"""

from dataclasses import dataclass

from .base import ParseError, find_column, find_table_with_header, to_number

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
    posted_at: object
    reach: float | None
    views: float | None
    likes: float | None
    comments: float | None
    saves: float | None
    shares: float | None


def parse_instagram(file) -> list[InstagramRow]:
    header, data_rows = find_table_with_header(file, HEADER_CANDIDATES)
    if header is None:
        raise ParseError(
            "인스타그램 성과 원본에서 헤더 행(NO, 발행리스트 포함)을 찾지 못했습니다. "
            "메타 콘텐츠 인사이트 내보내기 파일(xlsx 또는 csv)이 맞는지 확인해주세요."
        )

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"인스타그램 성과 원본에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    def cell(row, key):
        i = col_idx[key]
        return row[i] if i < len(row) else None

    out: list[InstagramRow] = []
    for row in data_rows:
        title = cell(row, "title")
        if not title:
            continue
        out.append(
            InstagramRow(
                title=str(title).strip(),
                url=cell(row, "url") or None,
                posted_at=cell(row, "posted_at"),
                reach=to_number(cell(row, "reach")),
                views=to_number(cell(row, "views")),
                likes=to_number(cell(row, "likes")),
                comments=to_number(cell(row, "comments")),
                saves=to_number(cell(row, "saves")),
                shares=to_number(cell(row, "shares")),
            )
        )
    return out
