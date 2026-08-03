"""PPT 표(콘텐츠 운영 내역 등)의 행 개수를 이번 달 항목 수에 맞춰 늘리거나 줄이고,
셀 텍스트를 채우는 공용 엔진. core/excel_block.py와 같은 발상 — 서식은 기존 표의
데이터 행(템플릿 행)에서 그대로 복제하고, 텍스트만 바꿔 끼운다.
"""

import copy

from pptx.oxml.ns import qn
from pptx.table import Table


def find_table_by_header(slides, header_must_include: list[str], ncols: int | None = None) -> Table | None:
    """모든 슬라이드를 훑어서, 첫 행(헤더)에 header_must_include의 모든 문자열이
    (부분 일치로) 포함된 첫 번째 표를 찾는다. ncols를 주면 열 개수도 맞아야 한다
    (예: 뉴스룸/블로그 표처럼 헤더가 비슷한 표를 구분하기 위해)."""
    for slide in slides:
        for shape in slide.shapes:
            if not shape.has_table:
                continue
            tbl = shape.table
            if ncols is not None and len(tbl.columns) != ncols:
                continue
            header_texts = [c.text.strip() for c in tbl.rows[0].cells]
            joined = " ".join(header_texts)
            if all(any(req in h for h in header_texts) or req in joined for req in header_must_include):
                return tbl
    return None


def _tr_list(tbl: Table):
    return tbl._tbl.findall(qn("a:tr"))


def set_cell_text(cell, text: str) -> None:
    """셀의 첫 번째 run 서식은 유지한 채 텍스트만 바꾼다. 나머지 run/문단은 지운다."""
    tf = cell.text_frame
    if not tf.paragraphs:
        cell.text = text
        return
    para = tf.paragraphs[0]
    if not para.runs:
        para.text = text
    else:
        para.runs[0].text = text
        for extra in para.runs[1:]:
            extra._r.getparent().remove(extra._r)
    for extra_para in tf.paragraphs[1:]:
        extra_para._p.getparent().remove(extra_para._p)


def sync_table_rows(
    tbl: Table,
    n_data_rows: int,
    header_rows: int = 1,
    footer_rows: int = 2,
    template_data_row_idx: int | None = None,
) -> None:
    """표의 데이터 행 개수를 n_data_rows에 맞춘다(header_rows/footer_rows는 그대로 두고
    그 사이만 늘리거나 줄인다). template_data_row_idx가 없으면 header 바로 다음 행을 서식
    템플릿으로 쓴다."""
    trs = _tr_list(tbl)
    current_data_rows = len(trs) - header_rows - footer_rows
    if current_data_rows < 0:
        raise ValueError("표 구조가 예상(헤더+데이터+푸터)과 다릅니다.")
    template_idx = template_data_row_idx if template_data_row_idx is not None else header_rows

    if n_data_rows == current_data_rows:
        return

    if n_data_rows > current_data_rows:
        template_tr = trs[template_idx]
        footer_start = header_rows + current_data_rows  # 늘리기 전 첫 footer 행의 인덱스
        insert_before = trs[footer_start] if footer_start < len(trs) else None
        for _ in range(n_data_rows - current_data_rows):
            new_tr = copy.deepcopy(template_tr)
            if insert_before is not None:
                insert_before.addprevious(new_tr)
            else:
                tbl._tbl.append(new_tr)
    else:
        # 데이터 행이 줄어드는 경우: 마지막 데이터 행부터 지운다
        for i in range(current_data_rows - n_data_rows):
            idx = header_rows + current_data_rows - 1 - i
            tr = trs[idx]
            tr.getparent().remove(tr)


def fill_data_rows(tbl: Table, rows: list[list[str]], header_rows: int = 1) -> None:
    """header_rows 바로 다음부터 rows의 내용을 순서대로 채운다.
    표 행 수는 이미 sync_table_rows로 맞춰져 있어야 한다."""
    for i, row_values in enumerate(rows):
        r = header_rows + i
        for c, value in enumerate(row_values):
            if c >= len(tbl.columns):
                break
            set_cell_text(tbl.cell(r, c), value)


def fill_footer_row(tbl: Table, row_idx: int, values: dict[int, str]) -> None:
    for col, value in values.items():
        set_cell_text(tbl.cell(row_idx, col), value)
