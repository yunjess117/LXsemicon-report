"""기능 3: 월간보고서 PPT 생성.

전월 PPT를 그대로 템플릿으로 열어(31장 고정 구성 유지) 아래 표들을 이번 달 데이터로 교체한다.
  - 운영 개요 > 유입데이터 분석 (전월/당월 + 전월 대비 증감률, 뉴스룸은 누적 평균까지)
  - 콘텐츠 운영 내역 (콘텐츠 제목에 발행 URL 하이퍼링크)
  - 캠페인 운영 개요 > 캠페인 운영 요약 (유튜브 동영상조회 / 인스타 참여·도달·동영상조회)

네이티브 차트, 스크린샷 기반 수기 데이터(방문유형/체류시간/성별연령/키워드 등)는 값을
만들어내지 않고 '확인 필요'로 남긴다 — 슬라이드를 잘못 건드려 깨뜨리는 것보다 안전하다.
"""

import datetime
import re
from dataclasses import dataclass, field

from openpyxl.utils import column_index_from_string
from pptx import Presentation

from .confirm import ConfirmLog
from .excel_block import find_month_blocks
from .ppt_table import fill_data_rows, fill_footer_row, find_row_by_label, find_table, set_cell_text, sync_table_rows

CONFIRM_LABEL = "⚠ 확인 필요"

# '월간 SNS 운영 수치' 시트: 26년 1월이 AF열이고 월마다 한 칸씩 오른쪽으로 간다.
MONTHLY_SHEET = "월간 SNS 운영 수치"
MONTHLY_FIRST_COL = column_index_from_string("AF")
MONTHLY_ROWS = {
    "뉴스룸_방문횟수": 4,
    "뉴스룸_세션수": 5,
    "유튜브_조회수": 6,
    "유튜브_구독자증감": 7,
    "유튜브_평균조회율": 8,
    "인스타_도달수": 9,
    "인스타_팔로워증감": 10,
    "페북_도달수": 11,
    "블로그_조회수": 12,
    "블로그_방문횟수": 13,
    "블로그_이웃증감": 14,
}

# '채널 팔로워' 시트 상단 월간 요약표: 팔로워/구독자 '절대값'은 여기서만 얻을 수 있다.
FOLLOWER_SHEET = "채널 팔로워"
FOLLOWER_COLS = {
    "뉴스룸_방문횟수": "D",
    "뉴스룸_세션수": "G",
    "유튜브_구독자": "J",
    "인스타_팔로워": "M",
    "페북_팔로워": "P",
    "블로그_이웃수": "S",
}


def _norm(s) -> str:
    return "".join(str(s).split()).lower()


def _num(v) -> float | None:
    """계산에 쓸 수 있는 값만 숫자로 돌려주고, 그 외는 None.

    로우데이터 셀에는 숫자만 들어있지 않다 — 채우지 못한 칸은 '⚠ 확인 필요',
    광고 미집행은 '-', 수기 입력·csv 유래 값은 '1,234'처럼 콤마나 공백이 섞인 문자열일 수
    있다. 이런 값을 숫자처럼 더하면 TypeError가 나므로 여기서 한 번에 걸러낸다."""
    if v is None or isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(",", "").replace(" ", "").replace("%", "")
    if not s or s in ("-", "—", "–") or CONFIRM_LABEL.replace(" ", "") in s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _round_half_up(x: float) -> int:
    """PPT 원본은 Python 기본 round()의 은행반올림이 아니라 일반적인 반올림(0.5는 올림)을 쓴다."""
    import math

    return math.floor(x + 0.5) if x >= 0 else -math.floor(-x + 0.5)


def _fmt_num(v) -> str:
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (int, float)):
        return f"{v:,}"
    return str(v)


def _fmt_int(v) -> str:
    """소수점을 버리고 정수로 표시(CPV·CPM·결과당 비용 등)."""
    n = _num(v)
    return CONFIRM_LABEL if n is None else f"{_round_half_up(n):,}"


def _agg(values: list) -> tuple[str, str]:
    """열 하나의 (합계 표시값, 평균 표시값)을 만든다.

    하나라도 숫자로 읽지 못하면(확인 필요/빈칸 등) 합계·평균을 임의로 계산하지 않고
    둘 다 '확인 필요'로 남긴다 — 일부 행이 빠진 합계를 그대로 보고서에 실으면
    맞아 보이지만 틀린 숫자가 되기 때문이다."""
    nums = []
    for v in values:
        n = _num(v)
        if n is None:
            return CONFIRM_LABEL, CONFIRM_LABEL
        nums.append(n)
    if not nums:
        return "", ""
    total = sum(nums)
    return _fmt_num(total), _fmt_num(_round_half_up(total / len(nums)))


def _sum(values: list) -> float | None:
    """합계를 숫자로 반환. 하나라도 못 읽으면 None."""
    total = 0.0
    for v in values:
        n = _num(v)
        if n is None:
            return None
        total += n
    return total


def _row_total(*values) -> float | str | None:
    """한 행 안에서 여러 열을 더한 값(총 반응 등). 하나라도 못 읽으면 '확인 필요'."""
    nums = []
    for v in values:
        n = _num(v)
        if n is None:
            return CONFIRM_LABEL
        nums.append(n)
    return sum(nums)


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


def _col_idx(letter: str) -> int:
    return column_index_from_string(letter)


def _read_block_values(ws, block, cols: dict[str, str]) -> list[dict]:
    """block의 각 데이터 행에서 cols(별칭 -> 열문자)에 해당하는 값을 읽어온다.
    수식 셀(=로 시작하는 문자열)은 openpyxl이 캐시된 계산값을 갖고 있지 않을 수 있어
    None으로 취급한다 — 필요한 합계는 호출측에서 원본 값으로 직접 계산한다."""
    out = []
    for r in range(block.data_start_row, block.data_end_row + 1):
        row = {}
        for alias, col_letter in cols.items():
            v = ws.cell(row=r, column=_col_idx(col_letter)).value
            if isinstance(v, str) and v.startswith("="):
                v = None
            row[alias] = v
        out.append(row)
    return out


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
            return (
                ws.cell(row=r, column=_col_idx(result_col)).value,
                ws.cell(row=r, column=_col_idx(fee_col)).value,
            )
    return None, None


@dataclass
class FillEntry:
    """결과 화면 '채운 내역' 한 줄."""

    slide: int | None
    channel: str
    target: str
    detail: str
    ok: bool

    @property
    def slide_label(self) -> str:
        return f"{self.slide}P" if self.slide else "-"


@dataclass
class PptReportResult:
    presentation: Presentation
    confirm: ConfirmLog
    entries: list[FillEntry] = field(default_factory=list)

    @property
    def n_filled(self) -> int:
        return sum(1 for e in self.entries if e.ok)

    @property
    def n_todo(self) -> int:
        return sum(1 for e in self.entries if not e.ok)

    def by_channel(self) -> dict[str, list[FillEntry]]:
        grouped: dict[str, list[FillEntry]] = {}
        for e in self.entries:
            grouped.setdefault(e.channel, []).append(e)
        return grouped


class _Recorder:
    """채운 내역과 확인 필요 항목을 함께 기록한다."""

    def __init__(self, log: ConfirmLog):
        self.log = log
        self.entries: list[FillEntry] = []

    def ok(self, slide, channel, target, detail):
        self.entries.append(FillEntry(slide, channel, target, detail, True))

    def todo(self, slide, channel, target, detail, reason: str | None = None):
        self.entries.append(FillEntry(slide, channel, target, detail, False))
        self.log.add(f"PPT/{channel}", target, reason or detail)


# ── 채널 정의 ────────────────────────────────────────────────────────────────
CHANNELS = ["뉴스룸", "유튜브", "인스타그램", "페이스북", "블로그"]
SLIDE_KEYWORD = {
    "뉴스룸": "NEWSROOM",
    "유튜브": "YOUTUBE",
    "인스타그램": "INSTAGRAM",
    "페이스북": "FACEBOOK",
    "블로그": "NAVER BLOG",
}
RAW_SHEET = {
    "뉴스룸": "뉴스룸_Raw Data",
    "유튜브": "유튜브_Raw Data",
    "인스타그램": "인스타그램_Raw Data",
    "페이스북": "페이스북_Raw Data",
    "블로그": "블로그_Raw Data",
}


def build_ppt_report(prev_ppt: Presentation, raw_wb, ad_wb, target_year: int, target_month: int) -> PptReportResult:
    log = ConfirmLog()
    rec = _Recorder(log)
    slides = prev_ppt.slides

    _fill_inflow_tables(slides, raw_wb, target_month, rec)
    _fill_newsroom(slides, raw_wb, rec)
    _fill_youtube(slides, raw_wb, ad_wb, rec)
    _fill_instagram(slides, raw_wb, ad_wb, rec)
    _fill_facebook(slides, raw_wb, rec)
    _fill_blog(slides, raw_wb, rec)
    _fill_campaign_tables(slides, ad_wb, rec)
    _update_title_slide(slides, target_year, target_month, rec)

    rec.todo(
        None, "공통", "차트·수기 데이터 슬라이드",
        "네이티브 차트/성별·연령/트래픽소스/키워드 등은 전월 내용 그대로",
        "KPI 3개월 추이 표, 네이티브 차트(구독자 추이, 성별/연령 등), 방문유형·체류시간·"
        "트래픽소스·키워드 등 스크린샷 기반 수기 데이터는 자동 채움 대상이 아닙니다. "
        "전월 내용이 그대로 남아있으니 슬라이드를 열어 직접 교체해주세요.",
    )
    return PptReportResult(presentation=prev_ppt, confirm=log, entries=rec.entries)


# ── ① 운영 개요 > 유입데이터 분석 ────────────────────────────────────────────
def _monthly_value(raw_wb, key: str, month: int):
    """'월간 SNS 운영 수치' 시트에서 해당 월 값을 읽는다."""
    if raw_wb is None or MONTHLY_SHEET not in raw_wb.sheetnames or key not in MONTHLY_ROWS:
        return None
    ws = raw_wb[MONTHLY_SHEET]
    v = ws.cell(row=MONTHLY_ROWS[key], column=MONTHLY_FIRST_COL + month - 1).value
    return _num(v)


def _monthly_series(raw_wb, key: str, upto_month: int) -> list[float]:
    """1월부터 upto_month까지 값들(빈 칸 제외) — 누적 평균 계산용."""
    out = []
    for m in range(1, upto_month + 1):
        v = _monthly_value(raw_wb, key, m)
        if v is not None:
            out.append(v)
    return out


def _follower_value(raw_wb, key: str, month: int):
    """'채널 팔로워' 시트 상단 월간 요약표에서 해당 월의 절대값을 읽는다."""
    if raw_wb is None or FOLLOWER_SHEET not in raw_wb.sheetnames or key not in FOLLOWER_COLS:
        return None
    ws = raw_wb[FOLLOWER_SHEET]
    want = f"{month}월"
    for r in range(1, min(ws.max_row, 40) + 1):
        label = ws.cell(row=r, column=3).value  # C열
        if isinstance(label, str) and label.strip() == want:
            return _num(ws.cell(row=r, column=_col_idx(FOLLOWER_COLS[key])).value)
    return None


def _content_blocks(raw_wb, channel: str):
    sheet = RAW_SHEET[channel]
    if raw_wb is None or sheet not in raw_wb.sheetnames:
        return None, []
    ws = raw_wb[sheet]
    return ws, find_month_blocks(ws)


def _block_metric(ws, block, cols: dict[str, str], kind: str):
    """한 달 블록에서 콘텐츠 수량/조회수/반응수를 계산한다."""
    rows = _read_block_values(ws, block, cols)
    if kind == "count":
        return float(len(rows))
    if kind == "views":
        return _sum([r["views"] for r in rows])
    if kind == "reaction":
        vals = []
        for r in rows:
            vals.extend(r[k] for k in cols if k not in ("views",))
        return _sum(vals)
    return None


# 증감률 글자색 — 원본 PPT가 쓰던 값 그대로(상승 빨강 / 하락 파랑)
CHANGE_UP = "FF0000"
CHANGE_DOWN = "0070C0"
CHANGE_NONE = "808080"  # '확인 필요'가 템플릿의 파란색을 물려받아 하락처럼 보이지 않도록


def _change_color(text: str) -> str | None:
    if text.startswith("▲"):
        return CHANGE_UP
    if text.startswith("▼"):
        return CHANGE_DOWN
    return None


def _pct_change(prev, cur) -> str:
    """전월 대비 증감률. 원본 표기(▲/▼, 소수 1자리)를 그대로 따른다."""
    if prev is None or cur is None or prev == 0:
        return CONFIRM_LABEL
    diff = (cur - prev) / abs(prev) * 100
    if abs(diff) < 0.05:
        return "-"
    return f"{'▲' if diff > 0 else '▼'}{abs(diff):.1f}%"


def _pp_change(prev, cur) -> str:
    """비율 지표(평균 조회율)는 비율의 비율이 아니라 '퍼센트포인트' 차이로 쓴다."""
    if prev is None or cur is None:
        return CONFIRM_LABEL
    diff = cur - prev
    if abs(diff) < 0.005:
        return "-"
    return f"{'▲' if diff > 0 else '▼'}{abs(diff):.2f}%"


def _youtube_avg_view_rate(raw_wb, block_index: int):
    """유튜브 평균 조회율은 수기 집계값이라 블록의 합계 행 L열에 그대로 들어있다."""
    ws, blocks = _content_blocks(raw_wb, "유튜브")
    if ws is None or len(blocks) <= abs(block_index):
        return None
    return _num(ws.cell(row=blocks[block_index].total_row, column=_col_idx("L")).value)


def _inflow_metrics(raw_wb, month: int) -> dict[str, list[tuple[str, object, object, str]]]:
    """채널별 (지표명, 전월값, 당월값, 증감 표기방식) 목록을 만든다."""
    prev_m = month - 1 if month > 1 else 12

    def content_metric(channel, kind, cols):
        ws, blocks = _content_blocks(raw_wb, channel)
        if ws is None or len(blocks) < 1:
            return None, None
        cur = _block_metric(ws, blocks[-1], cols, kind)
        prev = _block_metric(ws, blocks[-2], cols, kind) if len(blocks) >= 2 else None
        return prev, cur

    NEWS_COLS = {"views": "H"}
    BLOG_COLS = {"views": "H"}
    YT_COLS = {"views": "I", "likes": "M", "comments": "N", "shares": "O"}
    IG_COLS = {"views": "I", "likes": "M", "comments": "N", "saves": "O", "shares": "P"}
    FB_COLS = {"views": "I", "likes": "K", "comments": "L", "shares": "M"}

    def pair(key):
        return _monthly_value(raw_wb, key, prev_m), _monthly_value(raw_wb, key, month)

    def fpair(key):
        return _follower_value(raw_wb, key, prev_m), _follower_value(raw_wb, key, month)

    metrics: dict[str, list] = {}

    p, c = pair("뉴스룸_방문횟수")
    p2, c2 = pair("뉴스룸_세션수")
    nv_p, nv_c = content_metric("뉴스룸", "views", NEWS_COLS)
    nc_p, nc_c = content_metric("뉴스룸", "count", NEWS_COLS)
    metrics["뉴스룸"] = [
        ("방문횟수", p, c, "pct", "뉴스룸_방문횟수"),
        ("페이지뷰", None, None, "pct", None),
        ("세션", p2, c2, "pct", "뉴스룸_세션수"),
        ("콘텐츠 조회수", nv_p, nv_c, "pct", None),
        ("콘텐츠 수량", nc_p, nc_c, "pct", None),
    ]

    ys_p, ys_c = fpair("유튜브_구독자")
    yv_p, yv_c = pair("유튜브_조회수")
    yr_p, yr_c = content_metric("유튜브", "reaction", YT_COLS)
    yc_p, yc_c = content_metric("유튜브", "count", YT_COLS)
    metrics["유튜브"] = [
        ("구독자", ys_p, ys_c, "pct", None),
        ("총 조회수", yv_p, yv_c, "pct", None),
        ("평균 조회율", _youtube_avg_view_rate(raw_wb, -2), _youtube_avg_view_rate(raw_wb, -1), "pp", None),
        ("총 반응", yr_p, yr_c, "pct", None),
        ("콘텐츠 수량", yc_p, yc_c, "pct", None),
    ]

    if_p, if_c = fpair("인스타_팔로워")
    id_p, id_c = pair("인스타_도달수")
    ir_p, ir_c = content_metric("인스타그램", "reaction", IG_COLS)
    ic_p, ic_c = content_metric("인스타그램", "count", IG_COLS)
    metrics["인스타그램"] = [
        ("팔로워", if_p, if_c, "pct", None),
        ("도달", id_p, id_c, "pct", None),
        ("반응", ir_p, ir_c, "pct", None),
        ("콘텐츠 수량", ic_p, ic_c, "pct", None),
    ]

    ff_p, ff_c = fpair("페북_팔로워")
    fd_p, fd_c = pair("페북_도달수")
    fr_p, fr_c = content_metric("페이스북", "reaction", FB_COLS)
    fc_p, fc_c = content_metric("페이스북", "count", FB_COLS)
    metrics["페이스북"] = [
        ("팔로워", ff_p, ff_c, "pct", None),
        ("도달", fd_p, fd_c, "pct", None),
        ("반응", fr_p, fr_c, "pct", None),
        ("콘텐츠 수량", fc_p, fc_c, "pct", None),
    ]

    bn_p, bn_c = fpair("블로그_이웃수")
    bvis_p, bvis_c = pair("블로그_방문횟수")
    bview_p, bview_c = pair("블로그_조회수")
    bc_p, bc_c = content_metric("블로그", "views", BLOG_COLS)
    bq_p, bq_c = content_metric("블로그", "count", BLOG_COLS)
    metrics["블로그"] = [
        ("이웃", bn_p, bn_c, "pct", None),
        ("방문자", None, None, "pct", None),
        ("방문횟수", bvis_p, bvis_c, "pct", "블로그_방문횟수"),
        ("총 조회수", bview_p, bview_c, "pct", "블로그_조회수"),
        ("콘텐츠 조회수", bc_p, bc_c, "pct", None),
        ("콘텐츠 수량", bq_p, bq_c, "pct", None),
    ]
    return metrics


def _fill_inflow_tables(slides, raw_wb, month: int, rec: _Recorder) -> None:
    metrics = _inflow_metrics(raw_wb, month)
    prev_m = month - 1 if month > 1 else 12

    for channel in CHANNELS:
        tbl, slide_no = find_table(
            slides, ["전월 대비 증감률"], slide_must_include=[SLIDE_KEYWORD[channel], "유입데이터 분석"]
        )
        if tbl is None:
            rec.todo(None, channel, "유입데이터 분석 표", "PPT에서 표를 찾지 못함")
            continue

        has_cumulative = len(tbl.columns) >= 6

        # 템플릿은 '전월 PPT'라서 그 표의 당월 열(=이번 달 기준으로는 전월)에 이미 전월 값이
        # 들어있다. 엑셀에서 전월 값을 못 구하는 지표는 이 값을 그대로 이어받는다.
        carried = {}
        for label, *_ in metrics[channel]:
            ri = find_row_by_label(tbl, label)
            if ri is not None and ri != 0:
                carried[label] = tbl.cell(ri, 2).text.strip()

        set_cell_text(tbl.cell(0, 1), f"{prev_m}월")
        set_cell_text(tbl.cell(0, 2), f"{month}월")

        done, todo = [], []
        for label, prev, cur, mode, monthly_key in metrics[channel]:
            ri = find_row_by_label(tbl, label)
            if ri is None or ri == 0:
                continue
            prev_text = carried.get(label, "")
            if cur is None:
                # 전월 값은 템플릿에서 이어받아 살려두고, 당월/증감률만 확인 필요로 남긴다.
                set_cell_text(tbl.cell(ri, 1), prev_text or CONFIRM_LABEL)
                set_cell_text(tbl.cell(ri, 2), CONFIRM_LABEL)
                set_cell_text(tbl.cell(ri, 3), CONFIRM_LABEL, color=CHANGE_NONE)
                if has_cumulative:
                    set_cell_text(tbl.cell(ri, 4), CONFIRM_LABEL)
                    set_cell_text(tbl.cell(ri, 5), CONFIRM_LABEL, color=CHANGE_NONE)
                todo.append(label)
                continue

            is_rate = mode == "pp"
            if prev is None:
                prev = _num(prev_text)  # 엑셀에서 못 구하면 템플릿 값으로 대체
            set_cell_text(tbl.cell(ri, 1), f"{prev:.2f}%" if (is_rate and prev is not None) else (_fmt_num(prev) if prev is not None else (prev_text or CONFIRM_LABEL)))
            set_cell_text(tbl.cell(ri, 2), f"{cur:.2f}%" if is_rate else _fmt_num(cur))
            change = _pp_change(prev, cur) if is_rate else _pct_change(prev, cur)
            set_cell_text(tbl.cell(ri, 3), change, color=_change_color(change))

            if has_cumulative:
                series = _monthly_series(raw_wb, monthly_key, month) if monthly_key else _content_series(raw_wb, channel, label, month)
                if series:
                    avg = sum(series) / len(series)
                    cum_change = _pct_change(avg, cur)
                    set_cell_text(tbl.cell(ri, 4), _fmt_num(_round_half_up(avg)))
                    set_cell_text(tbl.cell(ri, 5), cum_change, color=_change_color(cum_change))
                else:
                    set_cell_text(tbl.cell(ri, 4), CONFIRM_LABEL)
                    set_cell_text(tbl.cell(ri, 5), CONFIRM_LABEL, color=CHANGE_NONE)
            done.append(label)

        if done:
            rec.ok(slide_no, channel, "유입데이터 분석 표", f"{prev_m}월/{month}월 + 증감률 — {len(done)}개 지표")
        if todo:
            rec.todo(
                slide_no, channel, "유입데이터 분석 표", f"{', '.join(todo)} — 원본에 값이 없어 '확인 필요'",
                f"{', '.join(todo)} 지표는 취합 엑셀에서 값을 찾지 못해 '확인 필요'로 남겼습니다.",
            )


def _content_series(raw_wb, channel: str, label: str, month: int) -> list[float]:
    """콘텐츠 시트에서 뽑는 지표(조회수/수량)의 1월~당월 값들 — 누적 평균용."""
    ws, blocks = _content_blocks(raw_wb, channel)
    if ws is None or not blocks:
        return []
    cols = {"views": "H"}
    kind = "count" if "수량" in label else "views"
    out = []
    for b in blocks[: month]:
        v = _block_metric(ws, b, cols, kind)
        if v is not None:
            out.append(v)
    return out


# ── ② 콘텐츠 운영 내역 ───────────────────────────────────────────────────────
def _last_block_rows(raw_wb, sheet_name, cols):
    if raw_wb is None or sheet_name not in raw_wb.sheetnames:
        return None
    ws = raw_wb[sheet_name]
    blocks = find_month_blocks(ws)
    if not blocks:
        return None
    return _read_block_values(ws, blocks[-1], cols)


def _links_from(rows: list[dict]) -> dict[int, str]:
    """행 번호 -> 콘텐츠 URL. 표의 제목 셀에 하이퍼링크로 건다."""
    return {
        i: str(r["url"]).strip()
        for i, r in enumerate(rows)
        if isinstance(r.get("url"), str) and r["url"].strip().startswith("http")
    }


def _log_totals(rec: _Recorder, slide_no, channel: str, columns: dict[str, str]) -> None:
    unresolved = [name for name, value in columns.items() if value == CONFIRM_LABEL]
    if unresolved:
        rec.todo(
            slide_no, channel, "콘텐츠 운영 내역 표 TOTAL·평균",
            f"{', '.join(unresolved)} 합계 미계산",
            f"{', '.join(unresolved)} 열에 숫자로 읽을 수 없는 값(확인 필요·빈칸 등)이 있어 "
            "합계·평균을 계산하지 않았습니다. 로우데이터에서 값을 채운 뒤 다시 만들어주세요.",
        )


def _fill_newsroom(slides, raw_wb, rec: _Recorder) -> None:
    ch = "뉴스룸"
    rows = _last_block_rows(raw_wb, RAW_SHEET[ch], {"date": "C", "cat": "D", "title": "E", "url": "F", "type": "G", "views": "H"})
    if rows is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "취합 엑셀에 시트가 없음")
        return
    tbl, slide_no = find_table(slides, ["발행일", "콘텐츠 제목", "조회수"], ncols=6)
    if tbl is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "PPT에서 표를 찾지 못함")
        return
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(
        tbl,
        [[str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""), str(r["type"] or ""), _fmt_num(r["views"])] for i, r in enumerate(rows)],
        links=_links_from(rows), link_col=3,
    )
    total, avg = _agg([r["views"] for r in rows])
    fill_footer_row(tbl, 1 + len(rows), {0: "TOTAL", 5: total})
    fill_footer_row(tbl, 1 + len(rows) + 1, {0: "평균", 5: avg})
    rec.ok(slide_no, ch, "콘텐츠 운영 내역 표", f"{len(rows)}건 (제목에 URL 링크)")
    _log_totals(rec, slide_no, ch, {"조회수": total})


def _fill_youtube(slides, raw_wb, ad_wb, rec: _Recorder) -> None:
    ch = "유튜브"
    rows = _last_block_rows(
        raw_wb, RAW_SHEET[ch],
        {"date": "C", "cat": "D", "title": "E", "url": "F", "views": "I", "length": "J", "avg_dur": "K", "avg_pct": "L", "likes": "M", "comments": "N", "shares": "O", "fee": "H"},
    )
    if rows is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "취합 엑셀에 시트가 없음")
        return
    tbl, slide_no = find_table(slides, ["발행일", "콘텐츠 제목", "영상길이"])
    if tbl is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "PPT에서 표를 찾지 못함")
        return

    table_rows, views, cvs, reactions, fees, lengths, durs = [], [], [], [], [], [], []
    for i, r in enumerate(rows):
        campaign_views, ad_fee = _find_ad_fee_and_views(ad_wb, "유튜브_캠페인", r["title"] or "")
        fee = r["fee"] if _num(r["fee"]) is not None else ad_fee
        reaction = _row_total(r["likes"], r["comments"], r["shares"])
        views.append(r["views"]); cvs.append(campaign_views); reactions.append(reaction); fees.append(fee)
        lengths.append(r["length"]); durs.append(r["avg_dur"])
        table_rows.append([
            str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""),
            _fmt_num(r["views"]), _fmt_num(campaign_views), _fmt_time_slot(r["length"]),
            _fmt_time_slot(r["avg_dur"]), _fmt_pct(r["avg_pct"]), _fmt_num(reaction), _fmt_num(fee),
        ])
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(tbl, table_rows, links=_links_from(rows), link_col=3)

    t_v, a_v = _agg(views); t_c, a_c = _agg(cvs); t_r, a_r = _agg(reactions); t_f, _ = _agg(fees)
    fill_footer_row(tbl, 1 + len(rows), {0: "TOTAL", 4: t_v, 5: t_c, 8: "", 9: t_r, 10: t_f})
    fill_footer_row(tbl, 1 + len(rows) + 1, {0: "평균", 4: a_v, 5: a_c, 6: _fmt_time_slot(_avg_time_slot(lengths)), 7: _fmt_time_slot(_avg_time_slot(durs)), 9: a_r})
    rec.ok(slide_no, ch, "콘텐츠 운영 내역 표", f"{len(rows)}건 (제목에 URL 링크)")
    _log_totals(rec, slide_no, ch, {"조회수": t_v, "캠페인 조회수": t_c, "총반응": t_r, "캠페인비": t_f})
    rec.todo(
        slide_no, ch, "TOTAL행 평균조회율", "쇼츠·시흥캠 제외 가중평균이라 비워둠",
        "쇼츠·시흥캠 영상 제외 가중평균 수기 집계값이라 자동 계산하지 않고 비웠습니다. 직접 채워주세요.",
    )


def _fill_instagram(slides, raw_wb, ad_wb, rec: _Recorder) -> None:
    ch = "인스타그램"
    rows = _last_block_rows(
        raw_wb, RAW_SHEET[ch],
        {"date": "C", "cat": "D", "title": "E", "url": "F", "type": "G", "reach": "I", "likes": "M", "comments": "N", "saves": "O", "shares": "P", "fee": "H"},
    )
    if rows is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "취합 엑셀에 시트가 없음")
        return
    tbl, slide_no = find_table(slides, ["발행일", "콘텐츠 제목", "저장", "공유"])
    if tbl is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "PPT에서 표를 찾지 못함")
        return

    keys = ["reach", "likes", "comments", "saves", "shares"]
    cols = {k: [] for k in keys}
    table_rows = []
    for i, r in enumerate(rows):
        for k in keys:
            cols[k].append(r[k])
        table_rows.append([
            str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""), str(r["type"] or ""),
            _fmt_num(r["reach"]), _fmt_num(r["likes"]), _fmt_num(r["comments"]), _fmt_num(r["saves"]), _fmt_num(r["shares"]), _fmt_num(r["fee"]),
        ])
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(tbl, table_rows, links=_links_from(rows), link_col=3)

    agg = {k: _agg(cols[k]) for k in keys}
    fill_footer_row(tbl, 1 + len(rows), {0: "TOTAL", **{5 + i: agg[k][0] for i, k in enumerate(keys)}})
    fill_footer_row(tbl, 1 + len(rows) + 1, {0: "평균", **{5 + i: agg[k][1] for i, k in enumerate(keys)}})
    rec.ok(slide_no, ch, "콘텐츠 운영 내역 표", f"{len(rows)}건 (제목에 URL 링크)")
    _log_totals(rec, slide_no, ch, {"총 도달": agg["reach"][0], "좋아요": agg["likes"][0], "댓글": agg["comments"][0], "저장": agg["saves"][0], "공유": agg["shares"][0]})


def _fill_facebook(slides, raw_wb, rec: _Recorder) -> None:
    ch = "페이스북"
    rows = _last_block_rows(
        raw_wb, RAW_SHEET[ch],
        {"date": "C", "cat": "D", "title": "E", "url": "F", "type": "G", "reach": "I", "likes": "K", "comments": "L", "shares": "M", "fee": "H"},
    )
    if rows is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "취합 엑셀에 시트가 없음")
        return
    tbl, slide_no = find_table(slides, ["발행일", "콘텐츠 제목", "공유수", "좋아요수"])
    if tbl is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "PPT에서 표를 찾지 못함")
        return

    keys = ["reach", "shares", "likes", "comments"]
    cols = {k: [] for k in keys}
    table_rows = []
    for i, r in enumerate(rows):
        for k in keys:
            cols[k].append(r[k])
        table_rows.append([
            str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""), str(r["type"] or ""),
            _fmt_num(r["reach"]), _fmt_num(r["shares"]), _fmt_num(r["likes"]), _fmt_num(r["comments"]), str(r["fee"] or "-"),
        ])
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(tbl, table_rows, links=_links_from(rows), link_col=3)

    agg = {k: _agg(cols[k]) for k in keys}
    fill_footer_row(tbl, 1 + len(rows), {0: "TOTAL", **{5 + i: agg[k][0] for i, k in enumerate(keys)}})
    fill_footer_row(tbl, 1 + len(rows) + 1, {0: "평균", **{5 + i: agg[k][1] for i, k in enumerate(keys)}})
    rec.ok(slide_no, ch, "콘텐츠 운영 내역 표", f"{len(rows)}건 (제목에 URL 링크)")
    _log_totals(rec, slide_no, ch, {"콘텐츠 도달수": agg["reach"][0], "공유수": agg["shares"][0], "좋아요수": agg["likes"][0], "댓글수": agg["comments"][0]})


def _fill_blog(slides, raw_wb, rec: _Recorder) -> None:
    ch = "블로그"
    rows = _last_block_rows(raw_wb, RAW_SHEET[ch], {"date": "C", "cat": "D", "title": "E", "url": "F", "type": "G", "views": "H", "likes": "I", "comments": "J"})
    if rows is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "취합 엑셀에 시트가 없음")
        return
    tbl, slide_no = find_table(slides, ["발행일", "콘텐츠 제목", "총 반응"], ncols=7)
    if tbl is None:
        rec.todo(None, ch, "콘텐츠 운영 내역 표", "PPT에서 표를 찾지 못함")
        return

    table_rows, views, reactions = [], [], []
    for i, r in enumerate(rows):
        reaction = _row_total(r["likes"], r["comments"])
        views.append(r["views"]); reactions.append(reaction)
        table_rows.append([str(i + 1), str(r["date"] or ""), str(r["cat"] or ""), str(r["title"] or ""), str(r["type"] or ""), _fmt_num(r["views"]), _fmt_num(reaction)])
    sync_table_rows(tbl, len(rows), header_rows=1, footer_rows=2)
    fill_data_rows(tbl, table_rows, links=_links_from(rows), link_col=3)

    t_v, a_v = _agg(views); t_r, a_r = _agg(reactions)
    fill_footer_row(tbl, 1 + len(rows), {0: "TOTAL", 5: t_v, 6: t_r})
    fill_footer_row(tbl, 1 + len(rows) + 1, {0: "평균", 5: a_v, 6: a_r})
    rec.ok(slide_no, ch, "콘텐츠 운영 내역 표", f"{len(rows)}건 (제목에 URL 링크)")
    _log_totals(rec, slide_no, ch, {"조회수": t_v, "총 반응": t_r})


# ── ③ 캠페인 운영 개요 > 캠페인 운영 요약 ────────────────────────────────────
DATE_PREFIX_RE = re.compile(r"^\d{6}_")


@dataclass
class CampaignSpec:
    channel: str
    label: str                 # 결과 화면에 쓸 이름
    ad_sheet: str
    header: list[str]          # PPT 표 헤더 식별용
    ncols: int
    strip_prefix: bool         # 주제에서 '260611_' 접두 제거 여부
    ad_cols: dict              # PPT 열 인덱스 -> AD 시트 열문자
    rate_cols: dict            # 열 인덱스 -> ('pct'|'int') 표기 방식
    total_cols: list[int]      # TOTAL 행에서 단순 합계로 채울 열
    weighted: dict             # 평균 행 가중평균: 열 -> (분자 AD열, 분모 AD열, 배수)
    pct_decimals: int = 2      # 비율(VTR 등) 소수 자릿수 — 원본 표기가 채널마다 다르다


CAMPAIGN_SPECS = [
    CampaignSpec(
        channel="유튜브", label="캠페인 운영 요약(동영상 조회)", ad_sheet="유튜브_캠페인",
        header=["주제", "영상형태", "CPV"], ncols=10, strip_prefix=False,
        ad_cols={1: "D", 2: "E", 3: "F", 4: "G", 5: "H", 6: "I", 7: "J", 8: "K", 9: "L"},
        rate_cols={6: "int", 7: "pct"},
        total_cols=[4, 5, 8, 9],
        weighted={6: ("K", "H", 1), 7: ("H", "G", 100)},
    ),
    CampaignSpec(
        channel="인스타그램", label="캠페인 운영 요약(참여)", ad_sheet="인스타그램_참여",
        header=["주제", "참여당 비용"], ncols=8, strip_prefix=True,
        ad_cols={1: "D", 2: "E", 3: "F", 4: "G", 5: "H", 6: "O", 7: "P"},
        rate_cols={5: "int"},
        total_cols=[4, 6, 7],
        weighted={5: ("O", "G", 1)},
        pct_decimals=1,
    ),
    CampaignSpec(
        channel="인스타그램", label="캠페인 운영 요약(도달)", ad_sheet="인스타그램_도달",
        header=["주제", "도달당 비용"], ncols=9, strip_prefix=True,
        ad_cols={1: "D", 2: "E", 3: "F", 4: "J", 5: "H", 6: "N", 7: "O", 8: "P"},
        rate_cols={5: "int", 6: "int"},
        total_cols=[4, 7, 8],
        weighted={5: ("O", "J", 1000), 6: ("O", "I", 1000)},
        pct_decimals=1,
    ),
    CampaignSpec(
        channel="인스타그램", label="캠페인 운영 요약(동영상 조회)", ad_sheet="인스타그램_동영상 조회",
        header=["주제", "동영상 조회", "CPV"], ncols=9, strip_prefix=True,
        ad_cols={1: "D", 2: "E", 3: "F", 4: "G", 5: "O", 6: "P", 7: "Q", 8: "R"},
        rate_cols={5: "int", 6: "pct"},
        total_cols=[4, 7, 8],
        weighted={5: ("Q", "G", 1), 6: ("G", "I", 100)},
        pct_decimals=1,
    ),
]


def _fill_campaign_tables(slides, ad_wb, rec: _Recorder) -> None:
    if ad_wb is None:
        rec.todo(None, "광고", "캠페인 운영 요약 표", "AD 최종본이 없어 채우지 못함")
        return

    for spec in CAMPAIGN_SPECS:
        if spec.ad_sheet not in ad_wb.sheetnames:
            rec.todo(None, spec.channel, spec.label, f"AD 파일에 '{spec.ad_sheet}' 시트가 없음")
            continue
        ws = ad_wb[spec.ad_sheet]
        blocks = find_month_blocks(ws, num_col="C")
        if not blocks:
            rec.todo(None, spec.channel, spec.label, "AD 시트에서 이번 달 블록을 찾지 못함")
            continue
        block = blocks[-1]
        tbl, slide_no = find_table(slides, spec.header, ncols=spec.ncols)
        if tbl is None:
            rec.todo(None, spec.channel, spec.label, "PPT에서 표를 찾지 못함")
            continue

        raw = []
        for r in range(block.data_start_row, block.data_end_row + 1):
            raw.append({letter: ws.cell(row=r, column=_col_idx(letter)).value for letter in set(spec.ad_cols.values()) | {"G", "H", "I", "J", "K", "O", "Q"}})

        table_rows = []
        for i, row in enumerate(raw):
            cells = [""] * spec.ncols
            cells[0] = str(i + 1)
            for col, letter in spec.ad_cols.items():
                v = row.get(letter)
                if col == 1:
                    text = str(v or "")
                    cells[col] = DATE_PREFIX_RE.sub("", text).strip() if spec.strip_prefix else text
                elif col in spec.rate_cols:
                    n = _num(v)
                    if n is None:
                        cells[col] = CONFIRM_LABEL
                    elif spec.rate_cols[col] == "pct":
                        pct = n * 100 if n <= 1 else n
                        cells[col] = f"{pct:.{spec.pct_decimals}f}%"
                    else:
                        cells[col] = _fmt_int(n)
                elif col in (2, 3):
                    cells[col] = str(v or CONFIRM_LABEL)
                else:
                    cells[col] = _fmt_num(_num(v)) if _num(v) is not None else CONFIRM_LABEL
            table_rows.append(cells)

        n = len(table_rows)
        has_avg = len(tbl.rows) >= 3 and "평균" in tbl.cell(len(tbl.rows) - 1, 0).text
        footer_rows = 2 if has_avg else 1
        sync_table_rows(tbl, n, header_rows=1, footer_rows=footer_rows)
        fill_data_rows(tbl, table_rows)

        totals, avgs = {0: "TOTAL"}, {0: "평균"}
        for col in spec.total_cols:
            letter = spec.ad_cols[col]
            total, avg = _agg([r.get(letter) for r in raw])
            totals[col] = total
            avgs[col] = avg
        for col, (num_l, den_l, mult) in spec.weighted.items():
            num, den = _sum([r.get(num_l) for r in raw]), _sum([r.get(den_l) for r in raw])
            if num is None or not den:
                avgs[col] = CONFIRM_LABEL
            elif spec.rate_cols.get(col) == "pct":
                avgs[col] = f"{num / den * mult:.{spec.pct_decimals}f}%"
            else:
                avgs[col] = _fmt_int(num / den * mult)
        fill_footer_row(tbl, 1 + n, totals)
        if has_avg:
            fill_footer_row(tbl, 1 + n + 1, avgs)

        rec.ok(slide_no, spec.channel, spec.label, f"{n}건")
        unresolved = [c for c in spec.total_cols if totals.get(c) == CONFIRM_LABEL]
        if unresolved:
            rec.todo(slide_no, spec.channel, spec.label + " TOTAL", "일부 열 합계 미계산",
                     "AD 시트에 숫자로 읽을 수 없는 값이 있어 합계를 계산하지 않았습니다.")


def _update_title_slide(slides, target_year: int, target_month: int, rec: _Recorder):
    """표지 슬라이드의 '20XX년 X월' 텍스트만 갱신한다. 실제 템플릿은 '2026'/'년 '/'5'/'월'처럼
    연도·월이 서로 다른 run에 나뉘어 있어, run 하나짜리 정규식으로는 못 찾는다 — 문단 전체
    텍스트로 패턴을 확인한 뒤, 순수 숫자로만 된 run을 연도/월로 보고 그 run만 바꾼다."""
    pattern = re.compile(r"\d{4}\s*년\s*\d{1,2}\s*월")
    replaced = False
    for shape in slides[0].shapes:
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
    if replaced:
        rec.ok(1, "공통", "표지", f"{target_year}년 {target_month}월")
    else:
        rec.todo(1, "공통", "표지", "연도/월 텍스트를 찾지 못함", "표지의 '연도/월' 텍스트를 자동으로 찾지 못해 직접 수정이 필요합니다.")
