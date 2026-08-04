"""LX세미콘 월간 리포트 자동화 웹앱 (Streamlit)

담당자가 매달 원본 자료를 업로드하면
① 로우데이터(취합)·AD 엑셀 -> ② 월간보고서 PPT
순서로 만들어주는 도구. 판단이 애매하거나 원본에 없는 값은 절대 임의로 채우지 않고
"⚠ 확인 필요"로 남긴다.

실행:
    streamlit run app.py
"""

from io import BytesIO

import openpyxl
import streamlit as st
from pptx import Presentation

from core.ad_data import build_ad_data
from core.ad_lookup import extract_campaign_fee_rows
from core.follower_data import fill_daily_followers
from core.parsers.base import ParseError
from core.parsers.blog import parse_blog
from core.parsers.campaign_plan import parse_campaign_plan
from core.parsers.facebook import parse_facebook
from core.parsers.google_ads import parse_google_ads
from core.parsers.instagram import parse_instagram
from core.parsers.meta_ads import parse_meta_ads
from core.parsers.newsroom_ga import parse_newsroom_ga
from core.parsers.publish_list import parse_publish_list
from core.parsers.youtube_studio import parse_youtube_studio
from core.ppt_report import build_ppt_report
from core.raw_data import SHEET_NAMES, build_raw_data

st.set_page_config(page_title="LX세미콘 월간 리포트 자동화", page_icon="📊", layout="wide")

PAGE_BG = "#F8FAFC"
CARD_BG = "#FFFFFF"
CARD_BORDER = "#E5E7EB"
TEXT_TITLE = "#111827"
TEXT_MUTED = "#6B7280"

# 구역 구분은 카드 배경색이 아니라 '제목 좌측 컬러바'로만 한다.
BAR_PREV_RAW = "#2563EB"  # 전월 로우데이터(취합)
BAR_PREV_AD = "#7C3AED"   # 전월 로우데이터 AD
BAR_INPUT = "#F59E0B"     # 이번 달 입력 데이터
BAR_CONTENT = "#10B981"   # 콘텐츠 성과
BAR_ADS = "#EF4444"       # 광고 데이터

CARD_KEYS = ["card_prev", "card_input", "card_content", "card_ads"]
_card_selectors = ", ".join(f".st-key-{k}" for k in CARD_KEYS)

st.markdown(
    f"""
    <style>
    .stApp {{ background: {PAGE_BG}; }}

    /* 업로드 카드 — 흰 배경 + 연한 테두리 (배경 전체를 색칠하지 않는다) */
    {_card_selectors} {{
        background: {CARD_BG};
        border: 1px solid {CARD_BORDER};
        border-radius: 12px;
        padding: 1.25rem 1.4rem 0.85rem 1.4rem;
        height: 100%;
    }}

    /* 카드 제목 — 좌측 컬러바만으로 구역을 구분 */
    .card-title {{
        font-size: 1.02rem; font-weight: 700; color: {TEXT_TITLE};
        letter-spacing: -0.01em; line-height: 1.3;
        margin: 0.1rem 0 0.9rem 0; padding-left: 0.7rem;
        border-left: 4px solid;
    }}
    .card-title + .card-note {{ margin-top: -0.6rem; }}
    .card-note {{ font-size: 0.82rem; color: {TEXT_MUTED}; margin: 0 0 0.9rem 0.7rem; }}

    /* 전월 최종본 카드 안의 두 항목 — 각각 다른 컬러바 */
    .st-key-sub_prev_raw, .st-key-sub_prev_ad {{ padding-left: 0.7rem; border-left: 4px solid; }}
    .st-key-sub_prev_raw {{ border-left-color: {BAR_PREV_RAW}; }}
    .st-key-sub_prev_ad  {{ border-left-color: {BAR_PREV_AD}; margin-top: 0.9rem; }}

    /* 섹션 제목 (📊 채널 원본 데이터) */
    .section-title {{
        font-size: 1.5rem; font-weight: 800; color: {TEXT_TITLE};
        letter-spacing: -0.02em; margin: 0.4rem 0 0.15rem 0;
    }}
    .section-sub {{ font-size: 0.86rem; color: {TEXT_MUTED}; margin: 0 0 0.9rem 0; }}

    .panel-gap {{ height: 0.9rem; }}

    /* 실행 버튼 — 화면에서 가장 눈에 띄도록 크게, 알약 모양, 그림자 */
    .st-key-btn_raw button, .st-key-btn_ad button {{
        width: 100%;
        min-height: 5rem;
        border: none;
        border-radius: 999px;
        color: #fff;
        transition: transform .15s ease, box-shadow .15s ease, filter .15s ease;
    }}
    .st-key-btn_raw button p, .st-key-btn_ad button p {{
        font-size: 1.45rem !important;
        font-weight: 800 !important;
        letter-spacing: -0.02em;
        line-height: 1.2;
    }}
    .st-key-btn_raw button {{
        background: linear-gradient(135deg, #3B82F6 0%, {BAR_PREV_RAW} 55%, #1D4ED8 100%);
        box-shadow: 0 12px 26px {BAR_PREV_RAW}59, 0 2px 6px {BAR_PREV_RAW}40;
    }}
    .st-key-btn_ad button {{
        background: linear-gradient(135deg, #A855F7 0%, {BAR_PREV_AD} 55%, #6D28D9 100%);
        box-shadow: 0 12px 26px {BAR_PREV_AD}59, 0 2px 6px {BAR_PREV_AD}40;
    }}
    .st-key-btn_raw button:hover:enabled, .st-key-btn_ad button:hover:enabled {{
        transform: translateY(-3px);
        filter: brightness(1.06);
        color: #fff;
    }}
    .st-key-btn_raw button:hover:enabled {{ box-shadow: 0 18px 34px {BAR_PREV_RAW}66, 0 3px 8px {BAR_PREV_RAW}4D; }}
    .st-key-btn_ad button:hover:enabled  {{ box-shadow: 0 18px 34px {BAR_PREV_AD}66, 0 3px 8px {BAR_PREV_AD}4D; }}
    .st-key-btn_raw button:active:enabled, .st-key-btn_ad button:active:enabled {{
        transform: translateY(-1px);
    }}
    .st-key-btn_raw button:disabled, .st-key-btn_ad button:disabled {{
        background: #E5E7EB; color: {TEXT_MUTED}; box-shadow: none;
    }}
    .st-key-btn_raw button:disabled p, .st-key-btn_ad button:disabled p {{ color: {TEXT_MUTED}; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def check_password() -> bool:
    """st.secrets에 APP_PASSWORD가 설정된 경우에만 접속 비밀번호를 요구한다.
    secrets.toml 자체가 없는 로컬 환경에서도 에러 없이 통과하도록 예외를 방어한다."""
    try:
        required = st.secrets.get("APP_PASSWORD")
    except Exception:
        required = None
    if not required:
        return True
    if st.session_state.get("_authed"):
        return True
    st.title("🔒 접속 비밀번호")
    pw = st.text_input("비밀번호를 입력하세요", type="password")
    if st.button("확인"):
        if pw == required:
            st.session_state["_authed"] = True
            st.rerun()
        else:
            st.error("비밀번호가 올바르지 않습니다.")
    return False


if not check_password():
    st.stop()


def card_title(text: str, bar: str, note: str | None = None) -> None:
    """카드 제목. 구역 구분은 좌측 컬러바로만 한다(카드 배경은 항상 흰색)."""
    st.markdown(
        f'<div class="card-title" style="border-left-color:{bar}">{text}</div>'
        + (f'<div class="card-note">{note}</div>' if note else ""),
        unsafe_allow_html=True,
    )


def render_result(out: dict) -> None:
    """세션에 저장해둔 처리 결과(요약·확인 필요·다운로드)를 그린다.
    다운로드 버튼을 눌러 rerun이 일어나도 결과가 사라지지 않도록 세션에 담아 렌더한다."""
    st.markdown(f"#### {out['title']}")
    for note in out.get("notes", []):
        st.warning(note)
    for cap in out.get("captions", []):
        st.caption(cap)

    cols = st.columns(max(1, len(out["stats"])))
    for col, (name, matched, total) in zip(cols, out["stats"]):
        with col:
            st.metric(name, f"{matched}/{total}", help="매칭 성공 / 이번 달 건수")

    confirm = out["confirm"]
    if confirm:
        st.markdown(f"**⚠ 확인 필요 항목 ({len(confirm)}건)**")
        grouped: dict[str, list] = {}
        for area, location, reason in confirm:
            grouped.setdefault(area, []).append((location, reason))
        for area, items in grouped.items():
            with st.expander(f"{area} ({len(items)}건)"):
                for location, reason in items:
                    st.write(f"- **{location}** — {reason}")
    else:
        st.info("확인이 필요한 항목이 없습니다.")

    st.download_button(
        out["download_label"],
        data=out["bytes"],
        file_name=out["filename"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=out["download_key"],
    )


st.title("LX세미콘 월간 리포트 자동화")
st.caption("매달 원본 자료를 업로드하면 로우데이터(취합)·AD 엑셀과 월간보고서 PPT를 만들어줍니다.")
st.caption("🔒 업로드한 파일은 서버에 저장되지 않고 메모리에서만 처리되며, 세션이 끝나면 함께 사라집니다.")
st.caption("⚠ 판단이 애매하거나 원본에 없는 값은 임의로 채우지 않고 '확인 필요'로 표시합니다. 직접 확인 후 채워주세요.")

tab_data, tab_ppt = st.tabs(["① 로우데이터 (취합 · AD)", "② 월간보고서 PPT"])

MONTH_KR = {m: f"{m}월" for m in range(1, 13)}

with tab_data:
    col_y, col_m = st.columns(2)
    with col_y:
        target_year = st.number_input("이번 리포트 연도", min_value=2020, max_value=2100, value=2026, step=1)
    with col_m:
        target_month = st.number_input("이번 리포트 월", min_value=1, max_value=12, value=7, step=1)

    st.subheader("입력 파일")
    st.caption("★ 표시는 필수 파일입니다. 선택 파일이 없으면 해당 항목은 '확인 필요'로 비워둡니다.")

    top_l, top_r = st.columns(2)
    with top_l:
        with st.container(key="card_prev"):
            card_title("전월 최종본", BAR_PREV_RAW, "이번 달 블록을 이어붙일 기준 파일입니다.")
            with st.container(key="sub_prev_raw"):
                prev_fn_file = st.file_uploader("★ 전월 로우데이터(취합) 최종본 (_fn.xlsx)", type=["xlsx"], key="prev_raw_fn")
            with st.container(key="sub_prev_ad"):
                ad_prev_fn_file = st.file_uploader("★ 전월 로우데이터 AD 최종본 (_fn.xlsx)", type=["xlsx"], key="prev_ad_fn")
            st.markdown('<div class="panel-gap"></div>', unsafe_allow_html=True)

    with top_r:
        with st.container(key="card_input"):
            card_title("이번 달 입력 데이터", BAR_INPUT, "발행리스트·캠페인 플랜은 한 번만 올리면 양쪽에서 함께 씁니다.")
            publish_list_file = st.file_uploader("★ 이번 달 콘텐츠 발행리스트", type=["xlsx", "xls"], key="publish_list")
            plan_file = st.file_uploader("이번 달 캠페인 플랜 (캠페인비·집행 기간 매칭용)", type=["xlsx"], key="campaign_plan")
            daily_follower_file = st.file_uploader("(선택) 채널별 팔로워·구독자 Daily data", type=["xlsx"], key="daily_follower")
            st.markdown('<div class="panel-gap"></div>', unsafe_allow_html=True)

    st.markdown(
        '<div class="section-title">📊 채널 원본 데이터</div>'
        '<div class="section-sub">채널별 성과 및 광고 데이터를 업로드하세요</div>',
        unsafe_allow_html=True,
    )

    bot_l, bot_r = st.columns(2)
    with bot_l:
        with st.container(key="card_content"):
            card_title("콘텐츠 성과", BAR_CONTENT)
            newsroom_file = st.file_uploader("뉴스룸 콘텐츠별 조회수 원본 (GA)", type=["xlsx", "csv"], key="newsroom")
            yt_file = st.file_uploader("유튜브 스튜디오 내보내기", type=["xlsx"], key="youtube")
            insta_file = st.file_uploader("인스타그램 성과 원본 (메타 콘텐츠 인사이트)", type=["xlsx"], key="instagram")
            fb_file = st.file_uploader("페이스북 성과 원본 (페이스북 인사이트 CSV)", type=["csv"], key="facebook")
            blog_file = st.file_uploader("블로그 실적 원본 (네이버 블로그 통계)", type=["xlsx", "csv"], key="blog")
            st.markdown('<div class="panel-gap"></div>', unsafe_allow_html=True)

    with bot_r:
        with st.container(key="card_ads"):
            card_title("광고 데이터", BAR_ADS)
            google_file = st.file_uploader("구글애즈 광고 리포트 (CSV, UTF-16)", type=["csv"], key="google_ads")
            meta_file = st.file_uploader("메타 광고 리포트 (참여/도달/동영상조회/트래픽)", type=["xlsx"], key="meta_ads")
            st.markdown('<div class="panel-gap"></div>', unsafe_allow_html=True)

    raw_missing = []
    if prev_fn_file is None:
        raw_missing.append("전월 로우데이터(취합) 최종본")
    if publish_list_file is None:
        raw_missing.append("콘텐츠 발행리스트")

    ad_missing = []
    if ad_prev_fn_file is None:
        ad_missing.append("전월 로우데이터 AD 최종본")
    if publish_list_file is None:
        ad_missing.append("콘텐츠 발행리스트")

    optional_missing = [
        label
        for label, f in [
            ("뉴스룸 콘텐츠별 조회수 원본", newsroom_file),
            ("유튜브 스튜디오 내보내기", yt_file),
            ("인스타그램 성과 원본", insta_file),
            ("페이스북 성과 원본", fb_file),
            ("블로그 실적 원본", blog_file),
            ("구글애즈 광고 리포트", google_file),
            ("메타 광고 리포트", meta_file),
            ("이번 달 캠페인 플랜", plan_file),
            ("채널별 팔로워·구독자 Daily data", daily_follower_file),
        ]
        if f is None
    ]
    if optional_missing:
        st.info(
            "다음 선택 파일이 없습니다 — 해당 채널/항목은 '확인 필요'로 비워둔 채 진행됩니다: "
            + ", ".join(optional_missing)
        )

    st.markdown("")
    # 두 처리는 순서와 무관하게 독립 실행되므로 사이에 화살표를 두지 않는다.
    btn_l, btn_r = st.columns(2)
    with btn_l:
        with st.container(key="btn_raw"):
            run_raw = st.button(
                "📊 로우데이터 취합", disabled=bool(raw_missing), key="do_raw", use_container_width=True
            )
    with btn_r:
        with st.container(key="btn_ad"):
            run_ad = st.button(
                "📣 AD 데이터 취합", disabled=bool(ad_missing), key="do_ad", use_container_width=True
            )

    warn_l, warn_r = st.columns(2)
    with warn_l:
        if raw_missing:
            st.warning("⚠ 필수 파일 없음: " + ", ".join(raw_missing))
    with warn_r:
        if ad_missing:
            st.warning("⚠ 필수 파일 없음: " + ", ".join(ad_missing))
    st.caption("⚠ 버즈빌_캠페인 시트는 12개월 고정 캘린더 구조라 자동화 대상이 아니며 항상 수기 확인이 필요합니다.")

    def safe_parse(parser, file, label, notes: list, **kwargs):
        if file is None:
            return None
        try:
            return parser(file, **kwargs)
        except ParseError as e:
            notes.append(f"{label} 파싱 실패 — 이 파일은 무시하고 '확인 필요'로 남깁니다: {e}")
            return None

    def parse_shared_publish_list():
        try:
            return parse_publish_list(publish_list_file), None
        except ParseError as e:
            return None, f"콘텐츠 발행리스트 형식 오류: {e}"

    # ── 로우데이터(취합) 실행 ────────────────────────────────────────────────
    if run_raw:
        notes: list[str] = []
        captions: list[str] = []
        try:
            prev_wb = openpyxl.load_workbook(prev_fn_file)
        except Exception as e:
            st.error(f"전월 로우데이터(취합) 최종본을 여는 중 오류: {e}")
            st.stop()

        missing_sheets = [s for s in SHEET_NAMES.values() if s not in prev_wb.sheetnames]
        if missing_sheets:
            st.error(
                "업로드하신 파일에 다음 시트가 없습니다: " + ", ".join(missing_sheets) + "\n\n"
                "'전월 로우데이터(취합) 최종본' 파일이 맞는지 확인해주세요 "
                "(AD 파일이나 다른 파일을 잘못 올렸을 수 있습니다). "
                f"이 파일에 실제로 있는 시트: {', '.join(prev_wb.sheetnames)}"
            )
            st.stop()

        publish_rows, err = parse_shared_publish_list()
        if err:
            st.error(err)
            st.stop()

        insta_rows = safe_parse(parse_instagram, insta_file, "인스타그램 성과 원본", notes)
        fb_rows = safe_parse(parse_facebook, fb_file, "페이스북 성과 원본", notes)
        yt_rows = safe_parse(parse_youtube_studio, yt_file, "유튜브 스튜디오 내보내기", notes)
        newsroom_rows = safe_parse(parse_newsroom_ga, newsroom_file, "뉴스룸 콘텐츠별 조회수 원본", notes)
        blog_rows = safe_parse(parse_blog, blog_file, "블로그 실적 원본", notes)
        plan_rows = safe_parse(
            parse_campaign_plan, plan_file, "캠페인 플랜", notes, sheet_name=f"{int(target_month)}월"
        )

        # 이미 이 세션에서 AD 파일을 만들었다면, 캠페인비는 계획 예산보다 실지출 기준이
        # 정확하므로 그 결과물을 우선 참고한다(추가 업로드 없이 자동으로).
        yt_campaign_rows, insta_campaign_rows = None, None
        ad_out = st.session_state.get("ad_out")
        if ad_out is not None:
            try:
                gen_ad_wb = openpyxl.load_workbook(BytesIO(ad_out["bytes"]))
                month_label = MONTH_KR[int(target_month)]
                yt_campaign_rows = extract_campaign_fee_rows(gen_ad_wb, "유튜브_캠페인", month_label)
                insta_campaign_rows = []
                for sheet in ("인스타그램_참여", "인스타그램_도달", "인스타그램_동영상 조회"):
                    insta_campaign_rows.extend(extract_campaign_fee_rows(gen_ad_wb, sheet, month_label))
                captions.append("방금 만든 AD 데이터의 실지출 기준으로 캠페인비를 먼저 매칭했습니다.")
            except Exception as e:
                notes.append(f"방금 만든 AD 데이터를 캠페인비 매칭에 쓰지 못했습니다(캠페인 플랜으로 대체): {e}")

        try:
            result = build_raw_data(
                prev_wb=prev_wb,
                target_year=int(target_year),
                target_month=int(target_month),
                publish_rows=publish_rows,
                instagram_rows=insta_rows,
                facebook_rows=fb_rows,
                youtube_studio_rows=yt_rows,
                newsroom_ga_rows=newsroom_rows,
                blog_rows=blog_rows,
                youtube_campaign_rows=yt_campaign_rows,
                instagram_campaign_rows=insta_campaign_rows,
                campaign_plan_rows=plan_rows,
            )
        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
            st.stop()

        if daily_follower_file is not None:
            try:
                daily_wb = openpyxl.load_workbook(daily_follower_file, data_only=True)
                filled = fill_daily_followers(
                    result.workbook, daily_wb, int(target_year), int(target_month), result.confirm
                )
                if filled:
                    captions.append(
                        f"채널 팔로워 일별 로그 {filled}일치를 자동으로 채웠습니다 (월간 요약표는 수기 확인 필요)."
                    )
            except Exception as e:
                notes.append(f"채널별 팔로워 Daily data를 처리하는 중 오류가 발생해 이 부분은 건너뜁니다: {e}")

        buffer = BytesIO()
        result.workbook.save(buffer)
        st.session_state["raw_out"] = {
            "title": "로우데이터(취합) 결과",
            "bytes": buffer.getvalue(),
            "filename": f"_피알원__LX세미콘_월간_리포트_Raw_Data_{int(target_month)}월_fn.xlsx",
            "download_label": "취합 엑셀 다운로드",
            "download_key": "raw_download",
            "stats": [(ch, s.matched, s.total_published) for ch, s in result.stats.items()],
            "confirm": [(i.area, i.location, i.reason) for i in result.confirm.items],
            "notes": notes,
            "captions": captions,
        }

    # ── 로우데이터 AD 실행 ──────────────────────────────────────────────────
    if run_ad:
        notes = []
        try:
            ad_prev_wb = openpyxl.load_workbook(ad_prev_fn_file)
        except Exception as e:
            st.error(f"전월 로우데이터 AD 최종본을 여는 중 오류: {e}")
            st.stop()

        ad_required_sheets = ["유튜브_캠페인", "인스타그램_참여", "인스타그램_도달"]
        ad_missing_sheets = [s for s in ad_required_sheets if s not in ad_prev_wb.sheetnames]
        if ad_missing_sheets:
            st.error(
                "업로드하신 파일에 다음 시트가 없습니다: " + ", ".join(ad_missing_sheets) + "\n\n"
                "'전월 로우데이터 AD 최종본' 파일이 맞는지 확인해주세요. "
                f"이 파일에 실제로 있는 시트: {', '.join(ad_prev_wb.sheetnames)}"
            )
            st.stop()

        ad_publish_rows, err = parse_shared_publish_list()
        if err:
            st.error(err)
            st.stop()

        google_rows = safe_parse(parse_google_ads, google_file, "구글애즈 광고 리포트", notes)
        meta_rows = safe_parse(parse_meta_ads, meta_file, "메타 광고 리포트", notes)
        ad_plan_rows = safe_parse(
            parse_campaign_plan, plan_file, "캠페인 플랜", notes, sheet_name=f"{int(target_month)}월"
        )

        try:
            ad_result = build_ad_data(
                prev_wb=ad_prev_wb,
                target_year=int(target_year),
                target_month=int(target_month),
                publish_rows=ad_publish_rows,
                google_ads_rows=google_rows,
                meta_ads_rows=meta_rows,
                campaign_plan_rows=ad_plan_rows,
            )
        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
            st.stop()

        ad_buffer = BytesIO()
        ad_result.workbook.save(ad_buffer)
        st.session_state["ad_out"] = {
            "title": "로우데이터 AD 결과",
            "bytes": ad_buffer.getvalue(),
            "filename": f"_피알원_LX세미콘_월간_리포트_Raw_Data_AD_{int(target_month)}월_fn.xlsx",
            "download_label": "AD 엑셀 다운로드",
            "download_key": "ad_download",
            "stats": [(sheet, s.matched, s.total_published) for sheet, s in ad_result.stats.items()],
            "confirm": [(i.area, i.location, i.reason) for i in ad_result.confirm.items],
            "notes": notes,
            "captions": [],
        }

    if st.session_state.get("raw_out") or st.session_state.get("ad_out"):
        st.divider()
        res_l, res_r = st.columns(2)
        with res_l:
            if st.session_state.get("raw_out"):
                render_result(st.session_state["raw_out"])
        with res_r:
            if st.session_state.get("ad_out"):
                render_result(st.session_state["ad_out"])

with tab_ppt:
    col_y3, col_m3 = st.columns(2)
    with col_y3:
        ppt_target_year = st.number_input("이번 리포트 연도  ", min_value=2020, max_value=2100, value=2026, step=1, key="ppt_year")
    with col_m3:
        ppt_target_month = st.number_input("이번 리포트 월  ", min_value=1, max_value=12, value=7, step=1, key="ppt_month")

    st.subheader("입력 파일")
    st.caption("★ 표시는 필수 파일입니다. 취합·AD 엑셀은 '확인 필요' 항목까지 다 채운 최종본을 올려주세요.")
    c1, c2 = st.columns(2)
    with c1:
        ppt_prev_file = st.file_uploader("★ 전월 SNS 운영 리포트 PPT (_FN.pptx)", type=["pptx"], key="ppt_prev")
        ppt_raw_file = st.file_uploader("★ 이번 달 로우데이터(취합) 최종본", type=["xlsx"], key="ppt_raw")
    with c2:
        ppt_ad_file = st.file_uploader("이번 달 로우데이터 AD 최종본 (유튜브 캠페인 조회수용)", type=["xlsx"], key="ppt_ad")

    ppt_required_missing = []
    if ppt_prev_file is None:
        ppt_required_missing.append("전월 SNS 운영 리포트 PPT")
    if ppt_raw_file is None:
        ppt_required_missing.append("이번 달 로우데이터(취합) 최종본")
    if ppt_ad_file is None:
        st.info("AD 최종본이 없으면 유튜브 표의 '캠페인/조회수'·'캠페인비' 일부가 비워집니다.")

    st.caption(
        "⚠ 표 데이터(콘텐츠 운영 내역)만 자동으로 채웁니다. KPI 3개월 추이 표, 네이티브 차트, "
        "스크린샷 기반 수기 데이터(방문유형·체류시간·성별연령·키워드 등)는 전월 내용이 그대로 "
        "남아있으니 직접 교체해야 합니다."
    )

    ppt_process = st.button("처리 시작", type="primary", disabled=bool(ppt_required_missing), key="ppt_process")
    if ppt_required_missing:
        st.warning("⚠ 다음 필수 파일이 없어 처리할 수 없습니다: " + ", ".join(ppt_required_missing))

    if ppt_process:
        try:
            prev_ppt = Presentation(ppt_prev_file)
        except Exception as e:
            st.error(f"전월 PPT를 여는 중 오류: {e}")
            st.stop()

        try:
            ppt_raw_wb = openpyxl.load_workbook(ppt_raw_file, data_only=True)
        except Exception as e:
            st.error(f"로우데이터(취합) 최종본을 여는 중 오류: {e}")
            st.stop()

        ppt_missing_sheets = [s for s in SHEET_NAMES.values() if s not in ppt_raw_wb.sheetnames]
        if ppt_missing_sheets:
            st.error(
                "업로드하신 파일에 다음 시트가 없습니다: " + ", ".join(ppt_missing_sheets) + "\n\n"
                "'이번 달 로우데이터(취합) 최종본' 파일이 맞는지 확인해주세요. "
                f"이 파일에 실제로 있는 시트: {', '.join(ppt_raw_wb.sheetnames)}"
            )
            st.stop()

        ppt_ad_wb = None
        if ppt_ad_file is not None:
            try:
                ppt_ad_wb = openpyxl.load_workbook(ppt_ad_file, data_only=True)
            except Exception as e:
                st.warning(f"AD 최종본을 읽는 중 오류가 발생해 이 부분은 건너뜁니다: {e}")

        try:
            ppt_result = build_ppt_report(
                prev_ppt=prev_ppt, raw_wb=ppt_raw_wb, ad_wb=ppt_ad_wb,
                target_year=int(ppt_target_year), target_month=int(ppt_target_month),
            )
        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
            st.stop()

        st.success("처리를 완료했습니다.")

        st.subheader("체크리스트")
        n_filled = sum(1 for r in ppt_result.filled_tables if r.filled)
        n_total = len(ppt_result.filled_tables)
        st.metric("자동 채움 표", f"{n_filled}/{n_total}")
        for r in ppt_result.filled_tables:
            if r.filled:
                st.write(f"✅ {r.table_label} — {r.n_rows}건 자동 채움")
            else:
                st.write(f"⬜ {r.table_label} — {r.reason}")

        if ppt_result.confirm:
            st.subheader(f"⚠ 확인 필요 항목 ({len(ppt_result.confirm)}건)")
            for area, items in ppt_result.confirm.by_area().items():
                with st.expander(f"{area} ({len(items)}건)"):
                    for item in items:
                        st.write(f"- **{item.location}** — {item.reason}")

        ppt_buffer = BytesIO()
        ppt_result.presentation.save(ppt_buffer)
        ppt_buffer.seek(0)
        st.download_button(
            "PPT 다운로드",
            data=ppt_buffer,
            file_name=f"_피알원__LX세미콘_{int(ppt_target_month)}월_SNS_운영_리포트_FN.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key="ppt_download",
        )
