"""Summary 슬라이드(5P · 6P) 채우기.

5P  ■ KPI 달성률         — 26년 KPI / N월 기준 누적 / 달성률
    ■ SNS 채널별 월 구독자 추이(최근 3개월) — 막대 차트 3개 + 월별 증감 표
6P  ■ 운영 요약           — 채널별 전월/당월 + 전월 대비(발행수·조회수/도달수·누적 구독자수·평균 조회율)

이 값들은 일부만 로우데이터에서 계산할 수 있고(누적 조회수·도달수, 팔로워, 발행수 등)
평균 조회율처럼 수기 집계값도 섞여 있다. 그래서 화면에서 표로 '입력'받는 구조로 만든다:
  1) read_summary_template(): 전월 PPT의 5P·6P에서 KPI 목표·전월 값·최근 추이를 읽는다.
  2) auto_summary_values(): 로우데이터(취합)에서 계산 가능한 이번 달 값을 구한다.
  3) default_inputs(): 1)+2)로 입력 표의 기본값을 만든다(사용자가 화면에서 고친다).
  4) fill_summary_slides(): 사용자가 확정한 값을 PPT 표·차트에 쓰고, 증감·달성률은 직접 계산한다.
값이 비어 있으면 임의로 채우지 않고 '⚠ 확인 필요'로 남긴다.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pptx.chart.data import CategoryChartData

from .ppt_table import find_row_by_label, set_cell_text

CONFIRM_LABEL = "⚠ 확인 필요"

CHANNELS = ["뉴스룸", "유튜브", "인스타그램", "페이스북", "블로그"]

# 5P 월별 증감 표: 채널별 (값, 증감, 증감율) 3칸씩, 1열부터
TREND_VALUE_COL = {ch: 1 + 3 * i for i, ch in enumerate(CHANNELS)}
# 6P 운영 요약 표: 채널별 (전월, 당월, 전월 대비) 3칸씩, 1열부터
COMPARE_COL = {ch: 1 + 3 * i for i, ch in enumerate(CHANNELS)}

# 6P 표의 행(첫 열 라벨 앞부분으로 찾는다)
COMPARE_ROWS = ["콘텐츠 발행수", "총 조회수 및 도달수", "누적 구독자수", "평균 조회율"]
RATE_ROW = "평균 조회율"

# 5P KPI 표의 열 라벨 -> 로우데이터에서 이번 달 값을 구하는 방법
KPI_RATE_LABEL = "유튜브(평균조회율)"

COLOR_UP = "FF0000"
COLOR_DOWN = "0070C0"
COLOR_FLAT = "000000"


# ── 숫자 도우미 ──────────────────────────────────────────────────────────────
def _num(v) -> float | None:
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        if v != v:  # NaN (data_editor 빈칸)
            return None
        return float(v)
    s = str(v).strip().replace(",", "").replace(" ", "").replace("%", "")
    if not s or s in ("-", "—", "–") or "확인필요" in s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _round_half_up(x: float) -> int:
    import math

    return math.floor(x + 0.5) if x >= 0 else -math.floor(-x + 0.5)


def _fmt_int(v) -> str:
    n = _num(v)
    return CONFIRM_LABEL if n is None else f"{_round_half_up(n):,}"


def _fmt_rate(v) -> str:
    """평균 조회율: 52.62 -> '52.62%'."""
    n = _num(v)
    return CONFIRM_LABEL if n is None else f"{n:.2f}%"


def _color_for(diff: float | None) -> str | None:
    if diff is None:
        return None
    if diff > 0:
        return COLOR_UP
    if diff < 0:
        return COLOR_DOWN
    return COLOR_FLAT


def _ym_label(year: int, month: int) -> str:
    return f"{year % 100}년 {month}월"


def _shift_month(year: int, month: int, delta: int) -> tuple[int, int]:
    idx = year * 12 + (month - 1) + delta
    return idx // 12, idx % 12 + 1


def _rate_from_cell(text: str) -> float | None:
    """'52.62%' -> 52.62 / '0.5262' -> 52.62."""
    n = _num(text)
    if n is None:
        return None
    return n * 100 if n <= 1 and "%" not in str(text) else n


# ── 1) 전월 PPT 읽기 ─────────────────────────────────────────────────────────
@dataclass
class SummaryTemplate:
    kpi_slide: int | None = None
    kpi_table: object = None
    kpi_labels: list[str] = field(default_factory=list)       # KPI 표 1행(채널) 라벨
    kpi_target: dict[str, float | None] = field(default_factory=dict)
    kpi_prev_cum: dict[str, float | None] = field(default_factory=dict)

    trend_table: object = None
    trend_values: list[dict[str, float | None]] = field(default_factory=list)  # 3개월(오래된 순)

    compare_slide: int | None = None
    compare_table: object = None
    compare_prev: dict[str, dict[str, float | None]] = field(default_factory=dict)  # 행 -> 채널 -> 전월 PPT의 '당월' 값

    charts: list = field(default_factory=list)  # 5P 막대 차트들

    @property
    def found_kpi(self) -> bool:
        return self.kpi_table is not None

    @property
    def found_trend(self) -> bool:
        return self.trend_table is not None

    @property
    def found_compare(self) -> bool:
        return self.compare_table is not None


def _slide_text(slide) -> str:
    return " ".join(sh.text_frame.text for sh in slide.shapes if sh.has_text_frame)


def read_summary_template(prs) -> SummaryTemplate:
    t = SummaryTemplate()
    for i, slide in enumerate(prs.slides, start=1):
        for shape in slide.shapes:
            if shape.has_chart and "KPI 달성률" in _slide_text(slide):
                t.charts.append(shape.chart)
            if not shape.has_table:
                continue
            tbl = shape.table
            first = tbl.cell(0, 0).text.strip()
            if "KPI 달성률" in first and len(tbl.rows) >= 5:
                t.kpi_slide, t.kpi_table = i, tbl
                t.kpi_labels = [tbl.cell(1, c).text.strip() for c in range(1, len(tbl.columns))]
                for c, label in enumerate(t.kpi_labels, start=1):
                    if label == KPI_RATE_LABEL:
                        t.kpi_target[label] = _rate_from_cell(tbl.cell(2, c).text)
                        t.kpi_prev_cum[label] = _rate_from_cell(tbl.cell(3, c).text)
                    else:
                        t.kpi_target[label] = _num(tbl.cell(2, c).text)
                        t.kpi_prev_cum[label] = _num(tbl.cell(3, c).text)
            elif first == "월" and len(tbl.columns) >= 16 and len(tbl.rows) >= 5:
                t.kpi_slide = t.kpi_slide or i
                t.trend_table = tbl
                t.trend_values = [
                    {ch: _num(tbl.cell(r, TREND_VALUE_COL[ch]).text) for ch in CHANNELS}
                    for r in range(2, len(tbl.rows))
                ]
            elif len(tbl.columns) >= 16 and find_row_by_label(tbl, "콘텐츠 발행수") is not None:
                t.compare_slide, t.compare_table = i, tbl
                for label in COMPARE_ROWS:
                    ri = find_row_by_label(tbl, label)
                    if ri is None:
                        continue
                    row = {}
                    for ch in CHANNELS:
                        cell = tbl.cell(ri, COMPARE_COL[ch] + 1).text  # 전월 PPT의 '당월' 칸
                        row[ch] = _rate_from_cell(cell) if label == RATE_ROW else _num(cell)
                    t.compare_prev[label] = row
    return t


# ── 2) 로우데이터에서 계산 가능한 이번 달 값 ─────────────────────────────────
def auto_summary_values(raw_wb, month: int) -> dict:
    """로우데이터(취합)에서 이번 달 값을 계산한다. 못 구한 값은 None.

    실제 9월 FN 보고서와 대조해 맞춘 기준:
      - 유튜브 누적 조회수 / 인스타 누적 도달수 = '월간 SNS 운영 수치' 1월~당월 합
      - 구독자·팔로워 = '채널 팔로워' 월간 요약표(뉴스룸은 구글 세션수)
      - 6P 총 조회수 및 도달수: 뉴스룸·페이스북 = 콘텐츠 시트 합계, 유튜브·인스타·블로그 = 월간 수치
    """
    from . import ppt_report as pr

    out = {"followers": {}, "kpi_cum": {}, "count": {}, "views": {}}
    if raw_wb is None:
        return out

    follower_keys = {
        "뉴스룸": "뉴스룸_세션수",
        "유튜브": "유튜브_구독자",
        "인스타그램": "인스타_팔로워",
        "페이스북": "페북_팔로워",
        "블로그": "블로그_이웃수",
    }
    for ch, key in follower_keys.items():
        out["followers"][ch] = pr._follower_value(raw_wb, key, month)

    def cum(key):
        series = [pr._monthly_value(raw_wb, key, m) for m in range(1, month + 1)]
        return None if any(v is None for v in series) else sum(series)

    out["kpi_cum"] = {
        "유튜브(조회수)": cum("유튜브_조회수"),
        "유튜브(구독자수)": out["followers"]["유튜브"],
        KPI_RATE_LABEL: None,  # 쇼츠·시흥캠·보도자료 제외 누적 수기 집계값
        "인스타그램(도달수)": cum("인스타_도달수"),
        "인스타그램(팔로워수)": out["followers"]["인스타그램"],
    }

    view_cols = {"뉴스룸": "H", "유튜브": "I", "인스타그램": "I", "페이스북": "I", "블로그": "H"}
    for ch in CHANNELS:
        ws, blocks = pr._content_blocks(raw_wb, ch)
        if ws is None or not blocks:
            out["count"][ch] = None
            continue
        out["count"][ch] = pr._block_metric(ws, blocks[-1], {"views": view_cols[ch]}, "count")
        if ch in ("뉴스룸", "페이스북"):
            out["views"][ch] = pr._block_metric(ws, blocks[-1], {"views": view_cols[ch]}, "views")

    out["views"]["유튜브"] = pr._monthly_value(raw_wb, "유튜브_조회수", month)
    out["views"]["인스타그램"] = pr._monthly_value(raw_wb, "인스타_도달수", month)
    out["views"]["블로그"] = pr._monthly_value(raw_wb, "블로그_조회수", month)
    return out


# ── 3) 입력 표 기본값 ────────────────────────────────────────────────────────
@dataclass
class SummaryInputs:
    """화면에서 확정한 값. 비어 있는(None) 칸은 PPT에 '확인 필요'로 들어간다."""

    kpi_target: dict[str, float | None]
    kpi_cum: dict[str, float | None]
    # 최근 3개월(오래된 순) 채널별 값. trend_base는 첫 달의 증감 계산용(그 전달 값).
    trend_labels: list[str]
    trend_values: list[dict[str, float | None]]
    trend_base: dict[str, float | None]
    # 행 -> 채널 -> (전월, 당월)
    compare: dict[str, dict[str, tuple[float | None, float | None]]]


def default_inputs(t: SummaryTemplate, auto: dict, year: int, month: int) -> SummaryInputs:
    kpi_labels = t.kpi_labels or list(auto.get("kpi_cum", {}).keys())
    kpi_target = {k: t.kpi_target.get(k) for k in kpi_labels}
    kpi_cum = {k: auto.get("kpi_cum", {}).get(k) for k in kpi_labels}

    # 추이: 전월 PPT의 3개월 중 뒤 2개월 + 이번 달. 첫 달 증감 기준은 전월 PPT의 맨 앞 달.
    tv = t.trend_values
    base = tv[0] if len(tv) >= 3 else {ch: None for ch in CHANNELS}
    kept = tv[-2:] if len(tv) >= 2 else [{ch: None for ch in CHANNELS}] * 2
    cur = {ch: auto.get("followers", {}).get(ch) for ch in CHANNELS}
    labels = [_ym_label(*_shift_month(year, month, d)) for d in (-2, -1, 0)]

    rate_cur = kpi_cum.get(KPI_RATE_LABEL)
    compare = {}
    for row in COMPARE_ROWS:
        prev_row = t.compare_prev.get(row, {})
        compare[row] = {}
        for ch in CHANNELS:
            if row == "콘텐츠 발행수":
                c = auto.get("count", {}).get(ch)
            elif row == "총 조회수 및 도달수":
                c = auto.get("views", {}).get(ch)
            elif row == "누적 구독자수":
                c = cur.get(ch)
            else:
                c = rate_cur if ch == "유튜브" else None
            compare[row][ch] = (prev_row.get(ch), c)

    return SummaryInputs(
        kpi_target=kpi_target, kpi_cum=kpi_cum,
        trend_labels=labels, trend_values=[dict(kept[0]), dict(kept[1]), cur], trend_base=dict(base),
        compare=compare,
    )


# ── 4) PPT에 쓰기 ───────────────────────────────────────────────────────────
def _pct_text(prev, cur, zero: str) -> tuple[str, float | None]:
    p, c = _num(prev), _num(cur)
    if p is None or c is None:
        return CONFIRM_LABEL, None
    if p == 0:
        return ("-", 0.0) if c == 0 else (CONFIRM_LABEL, None)
    diff = (c - p) / abs(p) * 100
    if round(diff, 1) == 0:
        return zero, 0.0
    return f"{diff:.1f}%", diff


def _fill_kpi(tbl, inp: SummaryInputs, month: int, todo: list) -> None:
    label_row = find_row_by_label(tbl, "기준 누적") or 3
    set_cell_text(tbl.cell(label_row, 0), f"{month}월 기준 누적")
    for c in range(1, len(tbl.columns)):
        label = tbl.cell(1, c).text.strip()
        target, cum = inp.kpi_target.get(label), inp.kpi_cum.get(label)
        if label == KPI_RATE_LABEL:
            set_cell_text(tbl.cell(2, c), _fmt_rate(target))
            set_cell_text(tbl.cell(3, c), _fmt_rate(cum))
            set_cell_text(tbl.cell(4, c), "-")
        else:
            set_cell_text(tbl.cell(2, c), _fmt_int(target))
            set_cell_text(tbl.cell(3, c), _fmt_int(cum))
            t_n, c_n = _num(target), _num(cum)
            if t_n and c_n is not None:
                set_cell_text(tbl.cell(4, c), f"{c_n / t_n * 100:.1f}%", color=COLOR_UP)
            else:
                set_cell_text(tbl.cell(4, c), CONFIRM_LABEL)
        if _num(target) is None or _num(cum) is None:
            todo.append(f"KPI {label}")


def _fill_trend(tbl, inp: SummaryInputs, todo: list) -> None:
    prev_row = inp.trend_base
    for i, (label, values) in enumerate(zip(inp.trend_labels, inp.trend_values)):
        r = 2 + i
        if r >= len(tbl.rows):
            break
        set_cell_text(tbl.cell(r, 0), label)
        for ch in CHANNELS:
            c0 = TREND_VALUE_COL[ch]
            cur, prev = _num(values.get(ch)), _num(prev_row.get(ch))
            set_cell_text(tbl.cell(r, c0), _fmt_int(cur))
            if cur is None or prev is None:
                set_cell_text(tbl.cell(r, c0 + 1), CONFIRM_LABEL)
                set_cell_text(tbl.cell(r, c0 + 2), CONFIRM_LABEL, color=COLOR_FLAT)
                todo.append(f"구독자 추이 {label} {ch}")
                continue
            diff = cur - prev
            set_cell_text(tbl.cell(r, c0 + 1), f"{_round_half_up(diff):,}")
            text, d = _pct_text(prev, cur, zero="-")
            set_cell_text(tbl.cell(r, c0 + 2), text, color=_color_for(d) if d is not None else COLOR_FLAT)
        prev_row = values


def _fill_charts(charts, inp: SummaryInputs) -> int:
    """5P 막대 차트 3개(뉴스룸 / 유튜브 / 인스타·페북·블로그)의 시리즈를 최근 3개월로 바꾼다."""
    n = 0
    for chart in charts:
        try:
            cats = [str(c) for c in chart.plots[0].categories]
        except Exception:
            continue
        if not cats or not all(c in CHANNELS for c in cats):
            continue
        data = CategoryChartData()
        data.categories = cats
        ok = True
        for label, values in zip(inp.trend_labels, inp.trend_values):
            vals = [_num(values.get(c)) for c in cats]
            if any(v is None for v in vals):
                ok = False
            data.add_series(label, vals)
        if ok:
            chart.replace_data(data)
            n += 1
    return n


def _fill_compare(tbl, inp: SummaryInputs, prev_m: int, month: int, todo: list) -> None:
    for ch in CHANNELS:
        c0 = COMPARE_COL[ch]
        set_cell_text(tbl.cell(1, c0), f"{prev_m}월")
        set_cell_text(tbl.cell(1, c0 + 1), f"{month}월")
    for row in COMPARE_ROWS:
        ri = find_row_by_label(tbl, row)
        if ri is None:
            continue
        for ch in CHANNELS:
            c0 = COMPARE_COL[ch]
            prev, cur = inp.compare.get(row, {}).get(ch, (None, None))
            if row == RATE_ROW:
                if ch != "유튜브":
                    for k in range(3):
                        set_cell_text(tbl.cell(ri, c0 + k), "-")
                    continue
                p, c = _num(prev), _num(cur)
                set_cell_text(tbl.cell(ri, c0), _fmt_rate(p))
                set_cell_text(tbl.cell(ri, c0 + 1), _fmt_rate(c))
                if p is None or c is None:
                    set_cell_text(tbl.cell(ri, c0 + 2), CONFIRM_LABEL, color=COLOR_FLAT)
                    todo.append(f"운영 요약 {row} {ch}")
                else:
                    d = c - p  # 비율 지표는 %p 차이
                    set_cell_text(tbl.cell(ri, c0 + 2), f"{d:.1f}%", color=_color_for(round(d, 1)))
                continue
            set_cell_text(tbl.cell(ri, c0), _fmt_int(prev))
            set_cell_text(tbl.cell(ri, c0 + 1), _fmt_int(cur))
            text, d = _pct_text(prev, cur, zero="0.0%")
            set_cell_text(tbl.cell(ri, c0 + 2), text, color=_color_for(d) if d is not None else COLOR_FLAT)
            if d is None:
                todo.append(f"운영 요약 {row} {ch}")


def fill_summary_slides(prs, inp: SummaryInputs, year: int, month: int, rec) -> None:
    """5P·6P를 inp 값으로 채우고 rec(_Recorder)에 결과를 남긴다."""
    t = read_summary_template(prs)
    prev_m = _shift_month(year, month, -1)[1]

    if t.found_kpi:
        todo: list[str] = []
        _fill_kpi(t.kpi_table, inp, month, todo)
        if todo:
            rec.todo(t.kpi_slide, "Summary", "KPI 달성률 표", f"{len(todo)}칸 비어 있음",
                     "입력하지 않은 값: " + ", ".join(todo))
        else:
            rec.ok(t.kpi_slide, "Summary", "KPI 달성률 표", f"{month}월 기준 누적 + 달성률")
    else:
        rec.todo(None, "Summary", "KPI 달성률 표", "PPT에서 표를 찾지 못함")

    if t.found_trend:
        todo = []
        _fill_trend(t.trend_table, inp, todo)
        n_charts = _fill_charts(t.charts, inp)
        if todo:
            rec.todo(t.kpi_slide, "Summary", "채널별 구독자 추이 표", f"{len(todo)}칸 비어 있음",
                     "입력하지 않은 값: " + ", ".join(todo))
        else:
            rec.ok(t.kpi_slide, "Summary", "채널별 구독자 추이 표", f"{inp.trend_labels[0]}~{inp.trend_labels[-1]} + 증감")
        if n_charts:
            rec.ok(t.kpi_slide, "Summary", "구독자 추이 차트", f"{n_charts}개 차트 데이터 교체")
        if n_charts < 3:
            rec.todo(t.kpi_slide, "Summary", "구독자 추이 차트", "일부 차트는 전월 그대로",
                     "값이 비어 있거나 차트를 찾지 못해 일부 구독자 추이 차트는 전월 데이터 그대로입니다.")
    else:
        rec.todo(None, "Summary", "채널별 구독자 추이 표", "PPT에서 표를 찾지 못함")

    # 6P 유튜브 평균 조회율(당월)을 비워두면 5P KPI에 입력한 누적 평균조회율을 그대로 쓴다
    # (실제 보고서에서 두 값은 같은 수기 집계값이다 — 9월: 52.62% / 52.62%).
    yt_prev, yt_cur = inp.compare.get(RATE_ROW, {}).get("유튜브", (None, None))
    if _num(yt_cur) is None and _num(inp.kpi_cum.get(KPI_RATE_LABEL)) is not None:
        inp.compare.setdefault(RATE_ROW, {})["유튜브"] = (yt_prev, inp.kpi_cum[KPI_RATE_LABEL])

    if t.found_compare:
        todo = []
        _fill_compare(t.compare_table, inp, prev_m, month, todo)
        if todo:
            rec.todo(t.compare_slide, "Summary", "운영 요약(전월 비교) 표", f"{len(todo)}칸 비어 있음",
                     "입력하지 않은 값: " + ", ".join(todo))
        else:
            rec.ok(t.compare_slide, "Summary", "운영 요약(전월 비교) 표", f"{prev_m}월/{month}월 + 전월 대비")
    else:
        rec.todo(None, "Summary", "운영 요약(전월 비교) 표", "PPT에서 표를 찾지 못함")
