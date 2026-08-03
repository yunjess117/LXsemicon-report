"""기능 3: 월간보고서 PPT 생성.

전월 PPT를 그대로 템플릿으로 열어(31장 고정 구성 유지), '콘텐츠 운영 내역' 표처럼
완성된 로우데이터·AD 엑셀 값으로 직접 채울 수 있는 표만 이번 달 데이터로 교체한다.
네이티브 차트, 3개월 추이형 KPI 표, 스크린샷 기반 수기 데이터(방문유형/체류시간/성별연령/
키워드 등)는 값을 만들어내지 않고 체크리스트에 '확인 필요'로 남긴다 — 슬라이드 자체를
잘못 건드려 깨뜨리는 것보다 안전하다.
"""

import datetime
from dataclasses import dataclass, field

from pptx import Presentation

from .confirm import ConfirmLog
from .excel_block import find_month_blocks
from .ppt_table import fill_data_rows, fill_footer_row, find_table_by_header, sync_table_rows


def _norm(s) -> str:
    return "".join(str(s).split()).lower()


def _fmt_num(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (int, float)):
        return f"{v:,}"
    return str(v)


def _fmt_time_slot(v) -> str:
    if isinstance(v, datetime.time):
        return f"{v.hour}:{v.minute:02d}"
    if v in (None, ""):
        return ""
    return str(v)


def _fmt_pct(v) -> str:
    if v in (None, "", "-"):
        return "-"
    try:
        return f"{float(v):.1f}%"
    except (TypeError, ValueError):
        return str(v)


def _round_half_up(x: float) -> int:
    """PPT 원본은 Python 기본 round()의 은행반올림이 아니라 일반적인 반올림(0.5는 올림)을 쓴다."""
    import math

    return math.floor(x + 0.5) if x >= 0 else -math.floor(-x + 0.5)


def _time_slot_to_seconds(v) -> int | None:
    if isinstance(v, datetime.time):
        return v.hour * 60 + v.minute  # h:mm 슬롯은 실제로 '분:초'를 담고 있음 (규칙 6)
    return None


def _avg_time_slot(values: list) -> datetime.time | None:
    secs = [s for s in (_time_slot_to_seconds(v) for v in values) if s is not None]
    if not secs:
        return None
    avg = _round_half_up(sum(secs) / len(secs))
    m, s = divmod(avg, 60)
    if m >= 24:
        return None
    return datetime.time(hour=m, minute=s)


def _read_block_values(ws, block, cols: dict[str, str]) -> list[dict]:
    """block의 각 데이터 행에서 cols(별칭 -> 열문자)에 해당하는 값을 읽어온다.
    수식 셀(=로 시작하는 문자열)은 openpyxl이 캐시된 계산값을 갖고 있지 않을 수 있어
    None으로 취급한다 — 필요한 합계는 호출측에서 원본 값으로 직접 계산한다."""
    out = []
    for r in range(block.data_start_row, block.data_end_row + 1):
        row = {}
        for alias, col_letter in cols.items():
            cell = ws.cell(row=r, column=_col_idx(col_letter))
            v = cell.value
            if isinstance(v, str) and v.startswith("="):
                v = None
            row[alias] = v
        out.append(row)
    return out


def _col_idx(letter: str) -> int:
    from openpyxl.utils import column_index_from_string

    return column_index_from_string(letter)


def _find_ad_fee_and_views(ad_wb, sheet_name: str, title: str) -> tuple[float | None, float | None]:
    """AD 워크북의 마지막(=이번 달) 블록에서 캠페인명에 콘텐츠 제목이 들어간 행을 찾아
    (캠페인 조회수/결과, 캠페인비)를 반환한다."""
    if ad_wb is None or sheet_name not in ad_wb.sheetnames:
        return None, None
    ws = ad_wb[sheet_name]
    blocks = find_month_blocks(ws, num_col="C")
    if not blocks:
        return None, None
    last = blocks[-1]
    norm_title = _norm(title)
    result_col = "H" if sheet_name == "유튜브_캠페인" else "G"
    fee_col = "L" if sheet_name == "유튜브_캠페인" else "P"
    for r in range(last.data_start_row, last.data_end_row + 1):
        name = ws.cell(row=r, column=4).value  # D열: 캠페인명
        if name and norm_title and norm_title in _norm(name):
            result = ws.cell(row=r, column=_col_idx(result_col)).value
            fee = ws.cell(row=r, column=_col_idx(fee_col)).value
            return result, fee
    return None, None


@dataclass
class SlideFillResult:
    table_label: str
    filled: bool
    n_rows: int = 0
    reason: str | None = None


@dataclass
class PptReportResult:
    presentation: Presentation
    confirm: ConfirmLog
    filled_tables: list[SlideFillResult] = field(default_factory=list)


CONTENT_SHEET = "콘텐츠 Raw Data"  # (참고용, 실제 시트명은 채널별로 다름)


def build_ppt_report(
    prev_ppt: Presentation,
    raw_wb,
    ad_wb,
    target_year: int,
    target_month: int,
) -> PptReportResult:
    log = ConfirmLog()
    filled: list[SlideFillResult] = []
    slides = prev_ppt.slides

    _fill_newsroom(slides, raw_wb, log, filled)
    _fill_youtube(slides, raw_wb, ad_wb, log, filled)
    _fill_instagram(slides, raw_wb, ad_wb, log, filled)
    _fill_facebook(slides, raw_wb, log, filled)
    _fill_blog(slides, raw_wb, log, filled)
    _update_title_slide(slides, target_year, target_month, log)

    log.add(
        "PPT/전체", "차트·수기 데이터 슬라이드",
        "KPI 3개월 추이 표, 네이티브 차트(구독자 추이, 성별/연령 등), 방문유형·체류시간·"
        "트래픽소스·키워드 등 스크린샷 기반 수기 데이터 슬라이드는 이번 버전에서 자동 채움 대상이 "
        "아닙니다. 전월 내용이 그대로 남아있으니 슬라이드를 열어 직접 교체해주세요.",
    )
    return PptReportResult(presentation=prev_ppt, confirm=log, filled_tables=filled)


def _last_block_rows(raw_wb, sheet_name, cols):
    if raw_wb is None or sheet_name not in raw_wb.sheetnames:
        return None
    ws = raw_wb[sheet_name]
    blocks = find_month_blocks(ws)
    if not blocks:
        return None
    return _read_block_values(ws, blocks[-1], cols)


def _fill_newsroom(slides, raw_wb, log, filled):
    label = "뉴스룸 콘텐츠 운영 내역"
    rows = _last_block_rows(raw_wb, "뉴스룸_Raw Data", {"date": "C", "cat": "D", "title": "E", "type": "G", "views": "H"})
    if rows is None:
        filled.append(SlideFillResult(label, False, reason="취합 엑셀에 뉴스룸_Raw Data 시트가 없음"))
        return
    tbl = find_table_by_header(slides, ["발행일", "콘텐츠 제목", "조회수"], ncols=6)
    if tbl is None:
        filled.append(SlideFillResult(label, False, reason="PPT에서 해당 표를 찾지 못함"))
        return
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(
        tbl,
        [[str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""), str(r["type"] or ""), _fmt_num(r["views"])] for i, r in enumerate(rows)],
    )
    total = sum(r["views"] or 0 for r in rows)
    avg = _round_half_up(total / len(rows)) if rows else 0
    fill_footer_row(tbl, 1 + len(rows), {0: "TOTAL", 5: _fmt_num(total)})
    fill_footer_row(tbl, 1 + len(rows) + 1, {0: "평균", 5: _fmt_num(avg)})
    filled.append(SlideFillResult(label, True, n_rows=len(rows)))


def _fill_youtube(slides, raw_wb, ad_wb, log, filled):
    label = "유튜브 콘텐츠 운영 내역"
    rows = _last_block_rows(
        raw_wb, "유튜브_Raw Data",
        {"date": "C", "cat": "D", "title": "E", "views": "I", "length": "J", "avg_dur": "K", "avg_pct": "L", "likes": "M", "comments": "N", "shares": "O", "fee": "H"},
    )
    if rows is None:
        filled.append(SlideFillResult(label, False, reason="취합 엑셀에 유튜브_Raw Data 시트가 없음"))
        return
    tbl = find_table_by_header(slides, ["발행일", "콘텐츠 제목", "영상길이"])
    if tbl is None:
        filled.append(SlideFillResult(label, False, reason="PPT에서 해당 표를 찾지 못함"))
        return

    table_rows = []
    total_views = total_campaign_views = total_react = total_fee = 0
    lengths, avg_durs = [], []
    any_fee_missing = False
    for i, r in enumerate(rows):
        campaign_views, ad_fee = _find_ad_fee_and_views(ad_wb, "유튜브_캠페인", r["title"] or "")
        fee = r["fee"] if r["fee"] is not None else ad_fee
        reaction = (r["likes"] or 0) + (r["comments"] or 0) + (r["shares"] or 0)
        total_views += r["views"] or 0
        total_campaign_views += campaign_views or 0
        total_react += reaction
        total_fee += fee or 0
        if fee is None:
            any_fee_missing = True
        lengths.append(r["length"])
        avg_durs.append(r["avg_dur"])
        table_rows.append(
            [
                str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""),
                _fmt_num(r["views"]), _fmt_num(campaign_views), _fmt_time_slot(r["length"]),
                _fmt_time_slot(r["avg_dur"]), _fmt_pct(r["avg_pct"]),
                _fmt_num(reaction), _fmt_num(fee),
            ]
        )
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(tbl, table_rows)
    n = len(rows) or 1
    avg_views = _round_half_up(total_views / n) if rows else 0
    avg_campaign_views = _round_half_up(total_campaign_views / n) if rows else 0
    avg_react = _round_half_up(total_react / n) if rows else 0
    avg_length = _avg_time_slot(lengths)
    avg_dur = _avg_time_slot(avg_durs)
    fill_footer_row(
        tbl, 1 + len(rows),
        {0: "TOTAL", 4: _fmt_num(total_views), 5: _fmt_num(total_campaign_views), 8: "", 9: _fmt_num(total_react), 10: _fmt_num(total_fee) if not any_fee_missing else ""},
    )
    fill_footer_row(
        tbl, 1 + len(rows) + 1,
        {0: "평균", 4: _fmt_num(avg_views), 5: _fmt_num(avg_campaign_views), 6: _fmt_time_slot(avg_length), 7: _fmt_time_slot(avg_dur), 9: _fmt_num(avg_react)},
    )
    filled.append(SlideFillResult(label, True, n_rows=len(rows)))
    log.add("PPT/" + label, "캠페인/조회수·평균조회율", "AD 매칭이 안 된 항목은 비어 있을 수 있어 직접 확인 필요")
    log.add(
        "PPT/" + label, "TOTAL행 평균조회율",
        "쇼츠·시흥캠 영상 제외 가중평균 수기 집계값이라 자동 계산하지 않고 비웠습니다. 직접 채워주세요.",
    )


def _fill_instagram(slides, raw_wb, ad_wb, log, filled):
    label = "인스타그램 콘텐츠 운영 내역"
    rows = _last_block_rows(
        raw_wb, "인스타그램_Raw Data",
        {"date": "C", "cat": "D", "title": "E", "type": "G", "reach": "I", "likes": "M", "comments": "N", "saves": "O", "shares": "P", "fee": "H"},
    )
    if rows is None:
        filled.append(SlideFillResult(label, False, reason="취합 엑셀에 인스타그램_Raw Data 시트가 없음"))
        return
    tbl = find_table_by_header(slides, ["발행일", "콘텐츠 제목", "저장", "공유"])
    if tbl is None:
        filled.append(SlideFillResult(label, False, reason="PPT에서 해당 표를 찾지 못함"))
        return

    table_rows = []
    sums = [0, 0, 0, 0, 0]  # reach, likes, comments, saves, shares
    for i, r in enumerate(rows):
        vals = [r["reach"] or 0, r["likes"] or 0, r["comments"] or 0, r["saves"] or 0, r["shares"] or 0]
        sums = [a + b for a, b in zip(sums, vals)]
        table_rows.append(
            [str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""), str(r["type"] or ""),
             _fmt_num(r["reach"]), _fmt_num(r["likes"]), _fmt_num(r["comments"]), _fmt_num(r["saves"]), _fmt_num(r["shares"]), _fmt_num(r["fee"])]
        )
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(tbl, table_rows)
    n = len(rows) or 1
    fill_footer_row(tbl, 1 + len(rows), {0: "TOTAL", 5: _fmt_num(sums[0]), 6: _fmt_num(sums[1]), 7: _fmt_num(sums[2]), 8: _fmt_num(sums[3]), 9: _fmt_num(sums[4])})
    fill_footer_row(tbl, 1 + len(rows) + 1, {0: "평균", 5: _fmt_num(_round_half_up(sums[0] / n)), 6: _fmt_num(_round_half_up(sums[1] / n)), 7: _fmt_num(_round_half_up(sums[2] / n)), 8: _fmt_num(_round_half_up(sums[3] / n)), 9: _fmt_num(_round_half_up(sums[4] / n))})
    filled.append(SlideFillResult(label, True, n_rows=len(rows)))


def _fill_facebook(slides, raw_wb, log, filled):
    label = "페이스북 콘텐츠 운영 내역"
    rows = _last_block_rows(
        raw_wb, "페이스북_Raw Data",
        {"date": "C", "cat": "D", "title": "E", "type": "G", "reach": "I", "likes": "K", "comments": "L", "shares": "M", "fee": "H"},
    )
    if rows is None:
        filled.append(SlideFillResult(label, False, reason="취합 엑셀에 페이스북_Raw Data 시트가 없음"))
        return
    tbl = find_table_by_header(slides, ["발행일", "콘텐츠 제목", "공유수", "좋아요수"])
    if tbl is None:
        filled.append(SlideFillResult(label, False, reason="PPT에서 해당 표를 찾지 못함"))
        return

    table_rows = []
    sums = [0, 0, 0, 0]  # reach, shares, likes, comments
    for i, r in enumerate(rows):
        vals = [r["reach"] or 0, r["shares"] or 0, r["likes"] or 0, r["comments"] or 0]
        sums = [a + b for a, b in zip(sums, vals)]
        table_rows.append(
            [str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""), str(r["type"] or ""),
             _fmt_num(r["reach"]), _fmt_num(r["shares"]), _fmt_num(r["likes"]), _fmt_num(r["comments"]), str(r["fee"] or "-")]
        )
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(tbl, table_rows)
    n = len(rows) or 1
    fill_footer_row(tbl, 1 + len(rows), {0: "TOTAL", 5: _fmt_num(sums[0]), 6: _fmt_num(sums[1]), 7: _fmt_num(sums[2]), 8: _fmt_num(sums[3])})
    fill_footer_row(tbl, 1 + len(rows) + 1, {0: "평균", 5: _fmt_num(_round_half_up(sums[0] / n)), 6: _fmt_num(_round_half_up(sums[1] / n)), 7: _fmt_num(_round_half_up(sums[2] / n)), 8: _fmt_num(_round_half_up(sums[3] / n))})
    filled.append(SlideFillResult(label, True, n_rows=len(rows)))


def _fill_blog(slides, raw_wb, log, filled):
    label = "블로그 콘텐츠 운영 내역"
    rows = _last_block_rows(raw_wb, "블로그_Raw Data", {"date": "C", "cat": "D", "title": "E", "type": "G", "views": "H", "likes": "I", "comments": "J"})
    if rows is None:
        filled.append(SlideFillResult(label, False, reason="취합 엑셀에 블로그_Raw Data 시트가 없음"))
        return
    tbl = find_table_by_header(slides, ["발행일", "콘텐츠 제목", "총 반응"], ncols=7)
    if tbl is None:
        filled.append(SlideFillResult(label, False, reason="PPT에서 해당 표를 찾지 못함"))
        return

    table_rows = []
    total_views = total_react = 0
    for i, r in enumerate(rows):
        reaction = (r["likes"] or 0) + (r["comments"] or 0)
        total_views += r["views"] or 0
        total_react += reaction
        table_rows.append([str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""), str(r["type"] or ""), _fmt_num(r["views"]), _fmt_num(reaction)])
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(tbl, table_rows)
    n = len(rows) or 1
    fill_footer_row(tbl, 1 + len(rows), {0: "TOTAL", 5: _fmt_num(total_views), 6: _fmt_num(total_react)})
    fill_footer_row(tbl, 1 + len(rows) + 1, {0: "평균", 5: _fmt_num(_round_half_up(total_views / n)), 6: _fmt_num(_round_half_up(total_react / n))})
    filled.append(SlideFillResult(label, True, n_rows=len(rows)))


def _update_title_slide(slides, target_year: int, target_month: int, log: ConfirmLog):
    """표지 슬라이드의 '20XX년 X월' 텍스트만 갱신한다. 실제 템플릿은 '2026'/'년 '/'5'/'월'처럼
    연도·월이 서로 다른 run에 나뉘어 있어, run 하나짜리 정규식으로는 못 찾는다 — 문단 전체
    텍스트로 패턴을 확인한 뒤, 순수 숫자로만 된 run을 연도/월로 보고 그 run만 바꾼다."""
    import re

    pattern = re.compile(r"\d{4}\s*년\s*\d{1,2}\s*월")
    replaced = False
    first_slide = slides[0]
    for shape in first_slide.shapes:
        if not shape.has_text_frame:
            continue
        for para in shape.text_frame.paragraphs:
            joined = "".join(r.text for r in para.runs)
            if not pattern.search(joined):
                continue
            numeric_runs = [r for r in para.runs if r.text.strip().isdigit()]
            year_run = next((r for r in numeric_runs if len(r.text.strip()) == 4), None)
            month_run = next((r for r in numeric_runs if r is not year_run and len(r.text.strip()) <= 2), None)
            if year_run is not None and month_run is not None:
                year_run.text = year_run.text.replace(year_run.text.strip(), str(target_year))
                month_run.text = month_run.text.replace(month_run.text.strip(), str(target_month))
                replaced = True
    if not replaced:
        log.add("PPT/표지", "제목 슬라이드", "표지의 '연도/월' 텍스트를 자동으로 찾지 못해 직접 수정 필요")
