"""페이스북 성과 원본 (페이스북 인사이트) 파서. csv(UTF-8 BOM)와 xlsx 둘 다 받는다.

실제 6월 샘플로 확인된 컬럼:
  NO, 발행리스트, 제목, 게시 시간, 고유 링크, 게시물 유형, 조회수, 일반 게시물의 조회수,
  홍보 게시물의 조회수, 도달, 일반 게시물의 도달, 홍보 게시물의 도달, ...,
  "공감, 댓글 및 공유", 공감, 댓글, 공유

실제 fn 파일 숫자와 대조해서 확정한 매핑:
  '조회수' 열 -> fn의 '총노출'
  '도달'   열 -> fn의 '총도달'
  '공감'/'댓글'/'공유' -> 그대로 좋아요/댓글/공유
"""

from dataclasses import dataclass

from .base import ParseError, find_column, find_table_with_header, to_number

HEADER_CANDIDATES = ["NO", "발행리스트"]

COLS = {
    "title": ["발행리스트"],
    "posted_at": ["게시 시간"],
    "url": ["고유 링크"],
    "impressions": ["조회수"],  # fn의 '총노출'
    "reach": ["도달"],  # fn의 '총도달'
    "likes": ["공감"],
    "comments": ["댓글"],
    "shares": ["공유"],
}


@dataclass
class FacebookRow:
    title: str
    url: str | None
    posted_at: object
    impressions: float | None
    reach: float | None
    likes: float | None
    comments: float | None
    shares: float | None


def parse_facebook(file) -> list[FacebookRow]:
    header, data_rows = find_table_with_header(file, HEADER_CANDIDATES)
    if header is None:
        raise ParseError(
            "페이스북 성과 원본에서 헤더 행(NO, 발행리스트 포함)을 찾지 못했습니다. "
            "페이스북 인사이트 내보내기 파일(csv 또는 xlsx)이 맞는지 확인해주세요."
        )

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"페이스북 성과 원본에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    def cell(row, key):
        i = col_idx[key]
        return row[i] if i < len(row) else None

    out: list[FacebookRow] = []
    for row in data_rows:
        title = cell(row, "title")
        if not title:
            continue
        out.append(
            FacebookRow(
                title=str(title).strip(),
                url=cell(row, "url") or None,
                posted_at=cell(row, "posted_at"),
                impressions=to_number(cell(row, "impressions")),
                reach=to_number(cell(row, "reach")),
                likes=to_number(cell(row, "likes")),
                comments=to_number(cell(row, "comments")),
                shares=to_number(cell(row, "shares")),
            )
        )
    return out
