"""LX세미콘 월간 리포트 자동화 웹앱 (Streamlit)

담당자가 매달 원본 자료를 업로드하면
① 로우데이터(취합) 엑셀 -> ② 로우데이터 AD 엑셀 -> ③ 월간보고서 PPT
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
from core.raw_data import build_raw_data

st.set_page_config(page_title="LX세미콘 월간 리포트 자동화", page_icon="📊", layout="wide")


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

st.title("LX세미콘 월간 리포트 자동화")
st.caption(
    "매달 원본 자료를 업로드하면 ① 로우데이터(취합) → ② 로우데이터 AD → ③ 월간보고서 PPT 순으로 만들어줍니다."
)
st.caption("🔒 업로드한 파일은 서버에 저장되지 않고 메모리에서만 처리되며, 세션이 끝나면 함께 사라집니다.")
st.caption("⚠ 판단이 애매하거나 원본에 없는 값은 임의로 채우지 않고 '확인 필요'로 표시합니다. 직접 확인 후 채워주세요.")

tab1, tab2, tab3 = st.tabs(["① 로우데이터(취합)", "② 로우데이터 AD", "③ 월간보고서 PPT"])

MONTH_KR = {m: f"{m}월" for m in range(1, 13)}

with tab1:
    col_y, col_m = st.columns(2)
    with col_y:
        target_year = st.number_input("이번 리포트 연도", min_value=2020, max_value=2100, value=2026, step=1)
    with col_m:
        target_month = st.number_input("이번 리포트 월", min_value=1, max_value=12, value=7, step=1)

    st.subheader("입력 파일")
    st.caption("★ 표시는 필수 파일입니다. 선택 파일이 없으면 해당 항목은 '확인 필요'로 비워둡니다.")

    c1, c2 = st.columns(2)
    with c1:
        prev_fn_file = st.file_uploader("★ 전월 로우데이터(취합) 최종본 (_fn.xlsx)", type=["xlsx"])
        publish_list_file = st.file_uploader("★ 이번 달 콘텐츠 발행리스트", type=["xlsx", "xls"])
        insta_file = st.file_uploader("인스타그램 성과 원본 (메타 콘텐츠 인사이트)", type=["xlsx"])
        fb_file = st.file_uploader("페이스북 성과 원본 (페이스북 인사이트 CSV)", type=["csv"])
    with c2:
        yt_file = st.file_uploader("유튜브 스튜디오 내보내기", type=["xlsx"])
        newsroom_file = st.file_uploader("뉴스룸 콘텐츠별 조회수 원본 (GA)", type=["xlsx", "csv"])
        blog_file = st.file_uploader("블로그 실적 원본 (네이버 블로그 통계)", type=["xlsx", "csv"])
        ad_fn_file = st.file_uploader("(선택) 이번 달 로우데이터 AD 최종본 — 캠페인비 매칭용", type=["xlsx"])
        daily_follower_file = st.file_uploader("(선택) 채널별 팔로워/구독자 Daily data", type=["xlsx"])

    required_missing = []
    if prev_fn_file is None:
        required_missing.append("전월 로우데이터(취합) 최종본")
    if publish_list_file is None:
        required_missing.append("콘텐츠 발행리스트")

    optional_missing = []
    for label, f in [
        ("인스타그램 성과 원본", insta_file),
        ("페이스북 성과 원본", fb_file),
        ("유튜브 스튜디오 내보내기", yt_file),
        ("뉴스룸 콘텐츠별 조회수 원본", newsroom_file),
        ("블로그 실적 원본", blog_file),
        ("이번 달 AD 최종본", ad_fn_file),
        ("채널별 팔로워/구독자 Daily data", daily_follower_file),
    ]:
        if f is None:
            optional_missing.append(label)

    if optional_missing:
        st.info(
            "다음 선택 파일이 없습니다 — 해당 채널/항목은 '확인 필요'로 비워둔 채 진행됩니다: "
            + ", ".join(optional_missing)
        )

    process = st.button("처리 시작", type="primary", disabled=bool(required_missing))
    if required_missing:
        st.warning("⚠ 다음 필수 파일이 없어 처리할 수 없습니다: " + ", ".join(required_missing))

    if process:
        try:
            prev_wb = openpyxl.load_workbook(prev_fn_file)
        except Exception as e:
            st.error(f"전월 로우데이터(취합) 최종본을 여는 중 오류: {e}")
            st.stop()

        try:
            publish_rows = parse_publish_list(publish_list_file)
        except ParseError as e:
            st.error(f"콘텐츠 발행리스트 형식 오류: {e}")
            st.stop()

        def safe_parse(parser, file, label):
            if file is None:
                return None
            try:
                return parser(file)
            except ParseError as e:
                st.warning(f"{label} 파싱 실패 — 이 파일은 무시하고 '확인 필요'로 남깁니다: {e}")
                return None

        insta_rows = safe_parse(parse_instagram, insta_file, "인스타그램 성과 원본")
        fb_rows = safe_parse(parse_facebook, fb_file, "페이스북 성과 원본")
        yt_rows = safe_parse(parse_youtube_studio, yt_file, "유튜브 스튜디오 내보내기")
        newsroom_rows = safe_parse(parse_newsroom_ga, newsroom_file, "뉴스룸 콘텐츠별 조회수 원본")
        blog_rows = safe_parse(parse_blog, blog_file, "블로그 실적 원본")

        yt_campaign_rows, insta_campaign_rows = None, None
        if ad_fn_file is not None:
            try:
                ad_wb = openpyxl.load_workbook(ad_fn_file)
                month_label = MONTH_KR[int(target_month)]
                yt_campaign_rows = extract_campaign_fee_rows(ad_wb, "유튜브_캠페인", month_label)
                insta_campaign_rows = []
                for sheet in ("인스타그램_참여", "인스타그램_도달", "인스타그램_동영상 조회"):
                    insta_campaign_rows.extend(extract_campaign_fee_rows(ad_wb, sheet, month_label))
            except Exception as e:
                st.warning(f"AD 최종본을 읽는 중 오류가 발생해 캠페인비 매칭을 건너뜁니다: {e}")

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
                    st.caption(f"채널 팔로워 일별 로그 {filled}일치를 자동으로 채웠습니다 (월간 요약표는 수기 확인 필요).")
            except Exception as e:
                st.warning(f"채널별 팔로워 Daily data를 처리하는 중 오류가 발생해 이 부분은 건너뜁니다: {e}")

        st.success("처리를 완료했습니다.")

        st.subheader("매칭 결과 요약")
        cols = st.columns(5)
        for col, (channel, s) in zip(cols, result.stats.items()):
            with col:
                st.metric(channel, f"{s.matched}/{s.total_published}", help="매칭 성공 / 이번 달 발행 건수")

        if result.confirm:
            st.subheader(f"⚠ 확인 필요 항목 ({len(result.confirm)}건)")
            for area, items in result.confirm.by_area().items():
                with st.expander(f"{area} ({len(items)}건)"):
                    for item in items:
                        st.write(f"- **{item.location}** — {item.reason}")
        else:
            st.info("확인이 필요한 항목이 없습니다.")

        buffer = BytesIO()
        result.workbook.save(buffer)
        buffer.seek(0)
        st.download_button(
            "취합 엑셀 다운로드",
            data=buffer,
            file_name=f"_피알원__LX세미콘_월간_리포트_Raw_Data_{int(target_month)}월_fn.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab2:
    col_y2, col_m2 = st.columns(2)
    with col_y2:
        ad_target_year = st.number_input("이번 리포트 연도 ", min_value=2020, max_value=2100, value=2026, step=1, key="ad_year")
    with col_m2:
        ad_target_month = st.number_input("이번 리포트 월 ", min_value=1, max_value=12, value=7, step=1, key="ad_month")

    st.subheader("입력 파일")
    st.caption("★ 표시는 필수 파일입니다.")
    c1, c2 = st.columns(2)
    with c1:
        ad_prev_fn_file = st.file_uploader("★ 전월 로우데이터 AD 최종본 (_fn.xlsx)", type=["xlsx"], key="ad_prev_fn")
        ad_publish_list_file = st.file_uploader("★ 이번 달 콘텐츠 발행리스트", type=["xlsx", "xls"], key="ad_publish_list")
        ad_google_file = st.file_uploader("구글애즈 광고 리포트 (CSV, UTF-16)", type=["csv"], key="ad_google")
    with c2:
        ad_meta_file = st.file_uploader("메타 광고 리포트 (참여/도달/동영상조회/트래픽)", type=["xlsx"], key="ad_meta")
        ad_plan_file = st.file_uploader("이번 달 캠페인 플랜 — 집행 기간 매칭용", type=["xlsx"], key="ad_plan")

    ad_required_missing = []
    if ad_prev_fn_file is None:
        ad_required_missing.append("전월 로우데이터 AD 최종본")
    if ad_publish_list_file is None:
        ad_required_missing.append("콘텐츠 발행리스트")

    ad_optional_missing = []
    for label, f in [("구글애즈 광고 리포트", ad_google_file), ("메타 광고 리포트", ad_meta_file), ("캠페인 플랜", ad_plan_file)]:
        if f is None:
            ad_optional_missing.append(label)
    if ad_optional_missing:
        st.info(
            "다음 선택 파일이 없습니다 — 해당 시트/항목은 '확인 필요'로 비워둔 채 진행됩니다: "
            + ", ".join(ad_optional_missing)
        )
    st.caption("⚠ 버즈빌_캠페인 시트는 12개월 고정 캘린더 구조라 자동화 대상이 아니며 항상 수기 확인이 필요합니다.")

    ad_process = st.button("처리 시작", type="primary", disabled=bool(ad_required_missing), key="ad_process")
    if ad_required_missing:
        st.warning("⚠ 다음 필수 파일이 없어 처리할 수 없습니다: " + ", ".join(ad_required_missing))

    if ad_process:
        try:
            ad_prev_wb = openpyxl.load_workbook(ad_prev_fn_file)
        except Exception as e:
            st.error(f"전월 로우데이터 AD 최종본을 여는 중 오류: {e}")
            st.stop()

        try:
            ad_publish_rows = parse_publish_list(ad_publish_list_file)
        except ParseError as e:
            st.error(f"콘텐츠 발행리스트 형식 오류: {e}")
            st.stop()

        def ad_safe_parse(parser, file, label, **kwargs):
            if file is None:
                return None
            try:
                return parser(file, **kwargs)
            except ParseError as e:
                st.warning(f"{label} 파싱 실패 — 이 파일은 무시하고 '확인 필요'로 남깁니다: {e}")
                return None

        google_rows = ad_safe_parse(parse_google_ads, ad_google_file, "구글애즈 광고 리포트")
        meta_rows = ad_safe_parse(parse_meta_ads, ad_meta_file, "메타 광고 리포트")
        plan_rows = ad_safe_parse(
            parse_campaign_plan, ad_plan_file, "캠페인 플랜", sheet_name=f"{int(ad_target_month)}월"
        )

        try:
            ad_result = build_ad_data(
                prev_wb=ad_prev_wb,
                target_year=int(ad_target_year),
                target_month=int(ad_target_month),
                publish_rows=ad_publish_rows,
                google_ads_rows=google_rows,
                meta_ads_rows=meta_rows,
                campaign_plan_rows=plan_rows,
            )
        except Exception as e:
            st.error(f"처리 중 오류가 발생했습니다: {e}")
            st.stop()

        st.success("처리를 완료했습니다.")

        st.subheader("매칭 결과 요약")
        ad_cols = st.columns(len(ad_result.stats))
        for col, (sheet, s) in zip(ad_cols, ad_result.stats.items()):
            with col:
                st.metric(sheet, f"{s.matched}/{s.total_published}", help="매칭 성공 / 이번 달 캠페인 건수")

        if ad_result.confirm:
            st.subheader(f"⚠ 확인 필요 항목 ({len(ad_result.confirm)}건)")
            for area, items in ad_result.confirm.by_area().items():
                with st.expander(f"{area} ({len(items)}건)"):
                    for item in items:
                        st.write(f"- **{item.location}** — {item.reason}")
        else:
            st.info("확인이 필요한 항목이 없습니다.")

        ad_buffer = BytesIO()
        ad_result.workbook.save(ad_buffer)
        ad_buffer.seek(0)
        st.download_button(
            "AD 엑셀 다운로드",
            data=ad_buffer,
            file_name=f"_피알원_LX세미콘_월간_리포트_Raw_Data_AD_{int(ad_target_month)}월_fn.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="ad_download",
        )

with tab3:
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
