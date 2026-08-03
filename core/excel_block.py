"""전월 fn 엑셀의 '월별 블록'(데이터 행 + 합계 행 + 평균 행) 구조를 읽어서
그 서식·수식 패턴을 그대로 복제해 새 달 블록을 맨 아래에 추가하는 공용 엔진.

콘텐츠 Raw Data 5개 시트(뉴스룸/유튜브/인스타그램/페이스북/블로그)가 전부 이 구조를 공유한다:
  - B열에 1,2,3...으로 번호가 매겨진 데이터 행들
  - 그 다음 '합계' 행 (B:F 병합, 각 지표 열에 =SUM(범위))
  - 그 다음 '평균' 행 (B:F 병합, 각 지표 열에 =AVERAGE(범위))
  - 위 3종류가 월마다 반복되며 아래로 계속 쌓인다

수식·서식은 절대 새로 '생각해서' 만들지 않고, 반드시 직전 블록에서 그대로 읽어와 복제한다.
"""

import copy
import re
from dataclasses import dataclass

from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

NUM_COL = "B"


@dataclass
class MonthBlock:
    data_start_row: int
    data_end_row: int  # inclusive
    total_row: int
    avg_row: int

    @property
    def n_rows(self) -> int:
        return self.data_end_row - self.data_start_row + 1


def find_month_blocks(ws: Worksheet, header_row: int = 4) -> list[MonthBlock]:
    """B열을 훑어서 (숫자 데이터행... , '합계', '평균') 블록들을 찾는다."""
    blocks: list[MonthBlock] = []
    row = header_row + 2  # 헤더가 보통 4:5 두 줄 병합이므로 그 다음부터 시작
    max_row = ws.max_row
    data_start = None
    while row <= max_row + 1:
        val = ws.cell(row=row, column=column_index_from_string(NUM_COL)).value
        if isinstance(val, (int, float)) and not isinstance(val, bool):
            if data_start is None:
                data_start = row
            row += 1
            continue
        if val == "합계" and data_start is not None:
            total_row = row
            avg_row = row + 1
            avg_val = ws.cell(row=avg_row, column=column_index_from_string(NUM_COL)).value
            if avg_val != "평균":
                raise ValueError(
                    f"{ws.title}: {total_row}행이 '합계'인데 바로 다음 행({avg_row})이 "
                    f"'평균'이 아니라 {avg_val!r} 입니다. 예상한 블록 구조와 다릅니다."
                )
            blocks.append(
                MonthBlock(
                    data_start_row=data_start,
                    data_end_row=total_row - 1,
                    total_row=total_row,
                    avg_row=avg_row,
                )
            )
            data_start = None
            row = avg_row + 1
            continue
        # 숫자도 아니고 '합계'도 아닌 행을 만나면(빈 행 등) 종료
        break
    return blocks


def _last_used_column(ws: Worksheet, row: int) -> int:
    max_col = 1
    for cell in ws[row]:
        if cell.value is not None:
            max_col = max(max_col, cell.column)
    return max_col


def _normalize_formula(formula: str, row: int) -> str:
    """수식 안의 이 행 관련 셀 참조(같은 행 번호)를 {row} 플레이스홀더로 치환해
    다른 블록의 수식과 '구조'를 비교할 수 있게 만든다. 다른 행 번호(다른 블록 참조 등)는 건드리지 않는다."""
    return re.sub(rf"(?<=[A-Z]){row}(?![0-9])", "{row}", formula)


def extract_row_formula_template(ws: Worksheet, block: MonthBlock) -> dict[str, str]:
    """데이터 행 안에서 셀 자체가 수식인 열(예: 유튜브 총반응 =SUM(M6:O6))을 찾아
    {열문자: 정규화된 수식템플릿} 형태로 반환한다. 첫 데이터 행을 기준으로 판단한다."""
    template: dict[str, str] = {}
    row = block.data_start_row
    max_col = _last_used_column(ws, row)
    for col_idx in range(1, max_col + 1):
        cell = ws.cell(row=row, column=col_idx)
        if isinstance(cell.value, str) and cell.value.startswith("="):
            col_letter = get_column_letter(col_idx)
            template[col_letter] = _normalize_formula(cell.value, row)
    return template


def extract_summary_formula_template(ws: Worksheet, block: MonthBlock) -> dict[str, dict[str, str]]:
    """합계/평균 행의 수식을 {열문자: 원본수식} 형태로 그대로 추출한다(치환은 append 시 수행)."""
    result: dict[str, dict[str, str]] = {"total": {}, "avg": {}}
    max_col = max(_last_used_column(ws, block.total_row), _last_used_column(ws, block.avg_row))
    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        tcell = ws.cell(row=block.total_row, column=col_idx)
        acell = ws.cell(row=block.avg_row, column=col_idx)
        if isinstance(tcell.value, str) and tcell.value.startswith("="):
            result["total"][col_letter] = tcell.value
        elif tcell.value not in (None, ""):
            result["total"][col_letter] = tcell.value  # '-' 같은 리터럴 (예: 페북 캠페인비 열)
        if isinstance(acell.value, str) and acell.value.startswith("="):
            result["avg"][col_letter] = acell.value
        elif acell.value not in (None, ""):
            result["avg"][col_letter] = acell.value
    return result


def _formula_shape(formula: str, total_row: int) -> str:
    """합계/평균 수식에서 이 블록의 데이터 범위 행 번호만 지우고 '모양'만 남긴다.
    예: '=SUM(H6:H17)' (total_row=18) -> '=SUM(H{r}:H{r})' 식으로 정규화해 블록간 비교."""
    return re.sub(r"\d+", "#", formula)


def verify_formula_consistency(
    ws: Worksheet, blocks: list[MonthBlock], check_last_n: int = 3
) -> list[str]:
    """최근 N개 블록의 합계/평균 수식 '모양'이 서로 같은지 검사한다.
    1월 블록은 하드코딩이 섞여있을 수 있으므로 기준에서 제외하고, 최근 블록들끼리만 비교한다."""
    problems: list[str] = []
    recent = blocks[-check_last_n:] if len(blocks) >= 2 else blocks
    if len(recent) < 2:
        return problems

    shapes_per_block = []
    for b in recent:
        tpl = extract_summary_formula_template(ws, b)
        shape = {
            "total": {c: _formula_shape(str(v), b.total_row) for c, v in tpl["total"].items() if isinstance(v, str) and v.startswith("=")},
            "avg": {c: _formula_shape(str(v), b.avg_row) for c, v in tpl["avg"].items() if isinstance(v, str) and v.startswith("=")},
        }
        shapes_per_block.append(shape)

    base = shapes_per_block[-1]  # 가장 최근(전월) 블록을 기준으로
    for kind in ("total", "avg"):
        for col, shape in base[kind].items():
            for other in shapes_per_block[:-1]:
                other_shape = other[kind].get(col)
                if other_shape is not None and other_shape != shape:
                    problems.append(
                        f"{ws.title} {kind}행 {col}열 수식 구조가 최근 블록들과 다릅니다 "
                        f"(기준: {shape}, 다른 블록: {other_shape})"
                    )
    return problems


def _copy_style(src_cell, dst_cell) -> None:
    if src_cell.has_style:
        dst_cell.font = copy.copy(src_cell.font)
        dst_cell.border = copy.copy(src_cell.border)
        dst_cell.fill = copy.copy(src_cell.fill)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy.copy(src_cell.protection)
        dst_cell.alignment = copy.copy(src_cell.alignment)


def append_month_block(
    ws: Worksheet,
    prev_block: MonthBlock,
    new_rows: list[dict[str, object]],
    row_formula_template: dict[str, str],
    summary_template: dict[str, dict[str, str]],
    merge_cols: tuple[str, str] = ("B", "F"),
) -> MonthBlock:
    """prev_block 바로 아래에 새 달 블록을 추가한다.

    new_rows: 각 원소는 {열문자: 값} 딕셔너리. 'B'(번호)는 자동으로 1부터 채워지므로 생략 가능.
    row_formula_template: extract_row_formula_template() 결과 (데이터 행마다 반복되는 수식 열).
    summary_template: extract_summary_formula_template() 결과 (직전 블록의 합계/평균 수식).
    """
    start_row = prev_block.avg_row + 1
    n = len(new_rows)

    style_data_row = prev_block.data_start_row
    style_total_row = prev_block.total_row
    style_avg_row = prev_block.avg_row

    max_col = max(
        _last_used_column(ws, style_data_row),
        _last_used_column(ws, style_total_row),
    )

    # 데이터 행 작성
    for i, row_data in enumerate(new_rows):
        r = start_row + i
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            src_cell = ws.cell(row=style_data_row, column=col_idx)
            dst_cell = ws.cell(row=r, column=col_idx)
            _copy_style(src_cell, dst_cell)
            if col_letter == NUM_COL and NUM_COL not in row_data:
                dst_cell.value = i + 1
            elif col_letter in row_formula_template:
                dst_cell.value = row_formula_template[col_letter].replace("{row}", str(r))
            elif col_letter in row_data:
                dst_cell.value = row_data[col_letter]
            # 그 외(패턴에도 없고 값도 없는 열)는 비워둔다

    data_end_row = start_row + n - 1
    total_row = data_end_row + 1
    avg_row = total_row + 1

    for kind, row_num, style_row in (("total", total_row, style_total_row), ("avg", avg_row, style_avg_row)):
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            src_cell = ws.cell(row=style_row, column=col_idx)
            dst_cell = ws.cell(row=row_num, column=col_idx)
            _copy_style(src_cell, dst_cell)
            if col_letter == NUM_COL:
                dst_cell.value = "합계" if kind == "total" else "평균"
                continue
            tpl = summary_template.get(kind, {}).get(col_letter)
            if tpl is None:
                continue
            if isinstance(tpl, str) and tpl.startswith("="):
                # 이전 블록 범위(예: H6:H17)를 새 블록 범위(예: H{start}:H{end})로 치환
                new_formula = re.sub(
                    rf"([A-Z]+){prev_block.data_start_row}:([A-Z]+){prev_block.data_end_row}",
                    lambda m: f"{m.group(1)}{start_row}:{m.group(2)}{data_end_row}",
                    tpl,
                )
                dst_cell.value = new_formula
            else:
                dst_cell.value = tpl  # 리터럴 값(예: '-')은 그대로 복제

    ws.merge_cells(f"{merge_cols[0]}{total_row}:{merge_cols[1]}{total_row}")
    ws.merge_cells(f"{merge_cols[0]}{avg_row}:{merge_cols[1]}{avg_row}")

    return MonthBlock(data_start_row=start_row, data_end_row=data_end_row, total_row=total_row, avg_row=avg_row)
