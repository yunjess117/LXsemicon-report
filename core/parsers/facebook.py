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

# 손으로 정리한 형식과 페이스북에서 바로 내려받은 원본 모두 가진 최소 열
HEADER_CANDIDATES = ["고유 링크", "도달"]

COLS = {
    "posted_at": ["게시 시간"],
    "url": ["고유 링크"],
    "impressions": ["조회수"],  # fn의 '총노출'
    "reach": ["도달"],  # fn의 '총도달'
    "likes": ["공감"],
    "comments": ["댓글"],
    "shares": ["공유"],
}
TITLE_CANDIDATES = ["발행리스트", "제목"]
DESC_CANDIDATES = ["설명"]


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


def _first_line(text) -> str:
    """제목/설명에서 첫 줄만 뽑는다. 플랫폼 원본은 제목 칸에도 캡션 전체(줄바꿈 포함)가
    들어 있어서, 그대로 두면 제목 비교가 안 되므로 첫 줄로 줄인다."""
    if not text:
        return ""
    for line in str(text).splitlines():
        s = line.strip().strip("⠀").strip()
        if s:
            return s[:100]
    return ""


def parse_facebook(file) -> list[FacebookRow]:
    header, data_rows = find_table_with_header(file, HEADER_CANDIDATES)
    if header is None:
        raise ParseError(
            "페이스북 성과 원본에서 헤더 행('고유 링크', '도달' 포함)을 찾지 못했습니다. "
            "페이스북 인사이트 내보내기 파일(csv 또는 xlsx)이 맞는지 확인해주세요."
        )

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"페이스북 성과 원본에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx
    title_idx = find_column(header, TITLE_CANDIDATES)
    desc_idx = find_column(header, DESC_CANDIDATES)

    def cell(row, key):
        i = col_idx[key]
        return row[i] if i < len(row) else None

    def at(row, i):
        return row[i] if (i is not None and i < len(row)) else None

    out: list[FacebookRow] = []
    for row in data_rows:
        url = cell(row, "url")
        title = _first_line(at(row, title_idx)) or _first_line(at(row, desc_idx))
        if not url and not title:
            continue
        out.append(
            FacebookRow(
                title=str(title).strip(),
                url=url or None,
                posted_at=cell(row, "posted_at"),
                impressions=to_number(cell(row, "impressions")),
                reach=to_number(cell(row, "reach")),
                likes=to_number(cell(row, "likes")),
                comments=to_number(cell(row, "comments")),
                shares=to_number(cell(row, "shares")),
            )
        )
    return out
