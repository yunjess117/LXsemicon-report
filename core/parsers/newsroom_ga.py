"""뉴스룸 콘텐츠별 조회수 원본(GA) 파서. xlsx와 csv 둘 다 받는다.

⚠ 실제로 확보한 6월 샘플("...pv...(네이버애널리틱스).xlsx")은 사이트 전체의 '일별' 페이지뷰
집계였고, fn 파일이 필요로 하는 '콘텐츠(URL)별' 조회수가 아니었다. 이 파서는 콘텐츠별 표
(URL 또는 제목 열 + 조회수류 숫자 열)가 있는 시트/표를 찾아보고, 없으면 명확하게 실패해서
"확인 필요"로 넘어가게 한다 — 절대 일별 합계를 콘텐츠 단위로 나눠 추정하지 않는다.
"""

import re
from dataclasses import dataclass

from .base import ParseError, find_column, load_any_tabular

PATH_ID_RE = re.compile(r"^/?\d+/?$")

VIEW_CANDIDATES = ["조회수", "페이지뷰", "PV", "조회 수"]
URL_OR_TITLE_CANDIDATES = ["URL", "콘텐츠 URL", "페이지", "제목", "콘텐츠명", "콘텐츠 주제"]


@dataclass
class NewsroomGaRow:
    key: str  # URL 또는 제목 (raw_data.py에서 URL 뒷번호로 우선 매칭, 없으면 제목 매칭)
    views: float


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


def _find_path_id_columns(rows: list[list]) -> tuple[int, int] | None:
    """헤더 이름이 후보 목록과 안 맞을 때를 대비한 보강 탐지: '/10287'처럼 URL 경로+번호로
    보이는 값이 대부분인 열과, 그 옆에 숫자가 대부분인 열이 있으면 (키 열, 조회수 열)로 본다.
    실제로 확인된 GA 내보내기 중에는 헤더 없이 이 두 열만 있는 형식이 있었다."""
    if not rows:
        return None
    n_cols = max((len(r) for r in rows[:50]), default=0)

    def ratio(col, pattern_check):
        checked = hits = 0
        for r in rows[:200]:
            if col >= len(r) or r[col] in (None, ""):
                continue
            checked += 1
            if pattern_check(r[col]):
                hits += 1
        return (hits / checked) if checked else 0

    id_cols = [c for c in range(n_cols) if ratio(c, lambda v: bool(PATH_ID_RE.match(str(v).strip()))) > 0.6]
    if not id_cols:
        return None
    for key_col in id_cols:
        for view_col in range(n_cols):
            if view_col == key_col:
                continue
            if ratio(view_col, lambda v: _to_num(v) is not None) > 0.6:
                return key_col, view_col
    return None


def parse_newsroom_ga(file) -> list[NewsroomGaRow]:
    tables = load_any_tabular(file)
    for _, rows in tables:
        rows = rows[:500]
        for i, row in enumerate(rows[:10]):
            header = row
            key_col = find_column(header, URL_OR_TITLE_CANDIDATES)
            view_col = find_column(header, VIEW_CANDIDATES)
            if key_col is None or view_col is None:
                continue
            data_rows = rows[i + 1 :]
            out: list[NewsroomGaRow] = []
            for r in data_rows:
                if key_col >= len(r) or view_col >= len(r):
                    continue
                key = r[key_col]
                views = _to_num(r[view_col])
                if key is None or views is None:
                    continue
                out.append(NewsroomGaRow(key=str(key).strip(), views=views))
            if out:
                return out

        # 헤더 이름으로 못 찾았으면, 값 패턴(경로+번호 / 숫자)으로 다시 시도
        found = _find_path_id_columns(rows)
        if found:
            key_col, view_col = found
            out = []
            for r in rows:
                if key_col >= len(r) or view_col >= len(r):
                    continue
                key, raw_view = r[key_col], r[view_col]
                if not key or not PATH_ID_RE.match(str(key).strip()):
                    continue
                views = _to_num(raw_view)
                if views is None:
                    continue
                out.append(NewsroomGaRow(key=str(key).strip(), views=views))
            if out:
                return out
    raise ParseError(
        "뉴스룸 콘텐츠별 조회수 원본에서 '콘텐츠(URL/제목)별 조회수' 표를 찾지 못했습니다. "
        "업로드한 파일이 사이트 전체 일별 합계(pv)만 있는 경우 이 값으로는 채울 수 없습니다 — "
        f"'{URL_OR_TITLE_CANDIDATES}' 중 하나와 '{VIEW_CANDIDATES}' 중 하나가 같은 표에 있는 "
        "콘텐츠별(URL별) 리포트가 필요합니다."
    )
