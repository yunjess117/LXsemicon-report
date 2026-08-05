"""인스타그램 성과 원본 파서. 두 가지 형식을 모두 받는다(xlsx/csv 무관).

  (A) 손으로 정리한 형식: NO, 발행리스트(제목), 설명, 게시 시간, 고유 링크, 게시물 유형,
      도달, 조회, 좋아요, 댓글, 저장, 공유, 팔로우
  (B) 메타에서 바로 내려받은 원본: 게시물 ID, 계정 ID, 계정 사용자 이름, 계정 이름, 설명,
      기간(초), 게시 시간, 고유 링크, 게시물 유형, 데이터 댓글, 날짜, 조회, 도달,
      좋아요, 공유, 팔로우, 댓글, 저장

(B)에는 제목 열이 아예 없다. 매칭은 고유 링크·게시일이 기준이라 제목이 없어도 되고,
제목이 필요한 자리(확인 필요 안내 등)에는 설명 첫 줄을 대신 쓴다.
"""

from dataclasses import dataclass

from .base import ParseError, find_column, find_table_with_header, to_number

# 두 형식 모두 가진 최소 열 — 이걸로 헤더 행을 찾는다.
HEADER_CANDIDATES = ["고유 링크", "도달"]

COLS = {
    "posted_at": ["게시 시간"],
    "url": ["고유 링크"],
    "post_type": ["게시물 유형"],
    "reach": ["도달"],
    "views": ["조회", "조회수"],
    "likes": ["좋아요"],
    "comments": ["댓글"],
    "saves": ["저장"],
    "shares": ["공유"],
}
TITLE_CANDIDATES = ["발행리스트", "제목"]   # (B)에는 없을 수 있음
DESC_CANDIDATES = ["설명"]


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


def parse_instagram(file) -> list[InstagramRow]:
    header, data_rows = find_table_with_header(file, HEADER_CANDIDATES)
    if header is None:
        raise ParseError(
            "인스타그램 성과 원본에서 헤더 행('고유 링크', '도달' 포함)을 찾지 못했습니다. "
            "메타 콘텐츠 인사이트 내보내기 파일(xlsx 또는 csv)이 맞는지 확인해주세요."
        )

    col_idx = {}
    for key, candidates in COLS.items():
        idx = find_column(header, candidates)
        if idx is None:
            raise ParseError(f"인스타그램 성과 원본에서 '{candidates[0]}' 열을 찾지 못했습니다.")
        col_idx[key] = idx
    title_idx = find_column(header, TITLE_CANDIDATES)
    desc_idx = find_column(header, DESC_CANDIDATES)

    def cell(row, key):
        i = col_idx[key]
        return row[i] if i < len(row) else None

    def at(row, i):
        return row[i] if (i is not None and i < len(row)) else None

    out: list[InstagramRow] = []
    for row in data_rows:
        url = cell(row, "url")
        title = _first_line(at(row, title_idx)) or _first_line(at(row, desc_idx))
        # 제목이 없어도 링크만 있으면 매칭에 쓸 수 있으므로 버리지 않는다.
        if not url and not title:
            continue
        out.append(
            InstagramRow(
                title=str(title).strip(),
                url=url or None,
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
