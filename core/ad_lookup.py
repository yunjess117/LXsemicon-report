"""(선택) 업로드된 AD 최종본에서 '이번 달 캠페인명 -> 캠페인비'를 뽑아주는 헬퍼.

기능 2(AD 파일 자동 생성)는 아직 만들지 않았지만, 기능 1의 캠페인비 매칭은 이미 완성된
AD 최종본이 있으면 그걸 참고하도록 요청서에 명시돼 있어 최소한으로 지원한다.
"""

from openpyxl.workbook import Workbook

FEE_HEADER_CANDIDATES = {"캠페인비", "켐페인비"}  # 실제 6월 파일에 오타(켐페인비)가 있어 둘 다 인식


def _find_fee_column(ws, header_row: int = 4) -> int | None:
    for cell in ws[header_row]:
        if cell.value and str(cell.value).strip() in FEE_HEADER_CANDIDATES:
            return cell.column
    return None


def extract_campaign_fee_rows(wb: Workbook, sheet_name: str, month_label: str) -> list[tuple[str, float]]:
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    fee_col = _find_fee_column(ws)
    if fee_col is None:
        return []
    out = []
    for row in ws.iter_rows(min_row=5):
        month_val = row[1].value  # B열
        if month_val != month_label:
            continue
        name = row[3].value  # D열 (캠페인명)
        fee_cell = ws.cell(row=row[0].row, column=fee_col)
        fee = fee_cell.value
        if name and isinstance(fee, (int, float)):
            out.append((str(name).strip(), float(fee)))
    return out
