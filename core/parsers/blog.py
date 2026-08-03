"""블로그 실적 원본(네이버 블로그 통계) 파서. xlsx와 csv 둘 다 받는다.

⚠ 실제 샘플을 아직 확보하지 못했다. fn 파일 기준 필요한 값은 콘텐츠별 조회수/공감/댓글이므로,
그 셋을 담은 표를 유연하게 탐색한다. 못 찾으면 무엇을 찾고 있었는지 명시한 에러를 낸다.
"""

from dataclasses import dataclass

from .base import ParseError, find_column, load_any_tabular

KEY_CANDIDATES = ["URL", "콘텐츠 URL", "제목", "글 제목", "콘텐츠 주제"]
VIEW_CANDIDATES = ["조회수", "조회 수"]
LIKE_CANDIDATES = ["공감", "공감수", "좋아요"]
COMMENT_CANDIDATES = ["댓글", "댓글수"]


@dataclass
class BlogRow:
    key: str
    views: float | None
    likes: float | None
    comments: float | None


def _to_num(v):
    """csv로 읽으면 셀 값이 전부 문자열이라('1,234') 숫자로 바꿔준다. xlsx는 이미 숫자로 오므로 그대로 통과."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return v
    try:
        return float(str(v).replace(",", "").strip())
    except ValueError:
        return None


def parse_blog(file) -> list[BlogRow]:
    tables = load_any_tabular(file)
    for _, rows in tables:
        rows = rows[:500]
        for i, row in enumerate(rows[:10]):
            key_col = find_column(row, KEY_CANDIDATES)
            view_col = find_column(row, VIEW_CANDIDATES)
            if key_col is None or view_col is None:
                continue
            like_col = find_column(row, LIKE_CANDIDATES)
            comment_col = find_column(row, COMMENT_CANDIDATES)
            data_rows = rows[i + 1 :]
            out: list[BlogRow] = []
            for r in data_rows:
                if key_col >= len(r) or not r[key_col]:
                    continue
                out.append(
                    BlogRow(
                        key=str(r[key_col]).strip(),
                        views=_to_num(r[view_col]) if view_col is not None and view_col < len(r) else None,
                        likes=_to_num(r[like_col]) if like_col is not None and like_col < len(r) else None,
                        comments=_to_num(r[comment_col])
                        if comment_col is not None and comment_col < len(r)
                        else None,
                    )
                )
            if out:
                return out
    raise ParseError(
        "블로그 실적 원본에서 콘텐츠별 표를 찾지 못했습니다. "
        f"'{KEY_CANDIDATES}' 중 하나와 '{VIEW_CANDIDATES}' 중 하나가 같은 표에 있어야 합니다."
    )
