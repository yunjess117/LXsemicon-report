"""페이스북 성과 원본 (페이스북 인사이트 CSV, UTF-8 BOM) 파서.

실제 6월 샘플로 확인된 컬럼:
  NO, 발행리스트, 제목, 게시 시간, 고유 링크, 게시물 유형, 조회수, 일반 게시물의 조회수,
  홍보 게시물의 조회수, 도달, 일반 게시물의 도달, 홍보 게시물의 도달, ...,
  "공감, 댓글 및 공유", 공감, 댓글, 공유

실제 fn 파일 숫자와 대조해서 확정한 매핑:
  '조회수' 열 -> fn의 '총노출'
  '도달'   열 -> fn의 '총도달'
  '공감'/'댓글'/'공유' -> 그대로 좋아요/댓글/공유
"""

import csv
import io
from dataclasses import dataclass

from .base import ParseError, find_column

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
    raw = file.read()
    if isinstance(raw, bytes):
        text = raw.decode("utf-8-sig")
    else:
        text = raw
    reader = csv.reader(io.StringIO(text))
    rows = list(reader)
    if not rows:
        raise ParseError("페이스북 성과 원본 CSV가 비어 있습니다.")

    header = rows[0]
    if not (set(HEADER_CANDIDATES) <= {h.strip() for h in header}):
        raise ParseError(
            f"페이스북 성과 원본 CSV의 첫 행이 예상한 헤더(NO, 발행리스트 포함)가 아닙니다: {header}"
        )

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"페이스북 성과 원본에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx

    def to_num(v):
        if v in (None, "", "-"):
            return None
        try:
            return float(str(v).replace(",", ""))
        except ValueError:
            return None

    out: list[FacebookRow] = []
    for row in rows[1:]:
        if len(row) <= max(col_idx.values()):
            continue
        title = row[col_idx["title"]]
        if not title:
            continue
        out.append(
            FacebookRow(
                title=str(title).strip(),
                url=row[col_idx["url"]] or None,
                posted_at=row[col_idx["posted_at"]],
                impressions=to_num(row[col_idx["impressions"]]),
                reach=to_num(row[col_idx["reach"]]),
                likes=to_num(row[col_idx["likes"]]),
                comments=to_num(row[col_idx["comments"]]),
                shares=to_num(row[col_idx["shares"]]),
            )
        )
    return out
