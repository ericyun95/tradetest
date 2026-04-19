import io
import os
import time
import streamlit as st

from main import (
    find_hs_codes,
    fetch_trade_data,
    score_countries,
    fetch_export_data,
    analyze_competitors,
    search_competitor_companies,
    get_buyer_channels,
    search_real_buyers,
    generate_word_report,
    fmt_usd,
    build_reason,
)

# ─────────────────────────────────────────────
# 페이지 설정
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Global Trade Mentor AI",
    page_icon="🌐",
    layout="wide",
)

st.title("🌐 Global Trade Mentor AI")
st.caption("수출 타겟 시장 · 경쟁사 · 바이어 자동 분석 도구  |  UN Comtrade 데이터 기반")
st.divider()

# ─────────────────────────────────────────────
# 사이드바 — API 키 설정
# ─────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ 설정")
    api_key = st.text_input(
        "UN Comtrade API Key",
        value=os.environ.get("COMTRADE_API_KEY", ""),
        type="password",
        placeholder="발급받은 API 키를 입력하세요",
    )
    st.caption("무료 발급: [comtradedeveloper.un.org](https://comtradedeveloper.un.org)")
    st.divider()
    st.markdown("""
**사용 가능 품목 예시**
- 식품: 라면, 김치, 녹차, 커피, 초콜릿
- 뷰티: 화장품, 립스틱, 샴푸, 향수
- 전자: 스마트폰, 배터리, 반도체
- 기타: 자동차, 전기차, 철강, 태양광패널
""")

# ─────────────────────────────────────────────
# 메인 — 입력
# ─────────────────────────────────────────────
col1, col2 = st.columns([3, 1])
with col1:
    query = st.text_input(
        "품목명 또는 HS Code를 입력하세요",
        placeholder="예: 라면, 포장재, 축중기 ...  또는  HS코드 직접입력: 190230"
    )
with col2:
    search_btn = st.button("🔍 HS Code 검색", use_container_width=True)

# ─────────────────────────────────────────────
# 세션 상태 초기화
# ─────────────────────────────────────────────
for key in ["candidates", "selected", "results"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ─────────────────────────────────────────────
# STEP 1 — HS Code 탐색
# ─────────────────────────────────────────────
if search_btn and query:
    candidates = find_hs_codes(query)
    if not candidates:
        st.warning("관련 HS Code를 찾지 못했습니다. 다른 키워드를 입력해보세요.")
    else:
        st.session_state.candidates = candidates
        st.session_state.selected   = None
        st.session_state.results    = None

if st.session_state.candidates:
    st.subheader("📦 HS Code 선택")
    options = {
        f"HS {c['code']}  —  {c['desc']}  (매칭 {c['match_score']}점)": c
        for c in st.session_state.candidates
    }
    choice = st.radio("분석에 사용할 HS Code를 선택하세요", list(options.keys()), index=0)
    st.session_state.selected = options[choice]

    run_btn = st.button("🚀 분석 시작", type="primary", use_container_width=False)

    # ─────────────────────────────────────────
    # STEP 2 — 전체 분석 실행
    # ─────────────────────────────────────────
    if run_btn:
        if not api_key:
            st.error("사이드바에서 UN Comtrade API 키를 입력해주세요.")
            st.stop()

        selected = st.session_state.selected
        results  = {}

        with st.status("분석 중...", expanded=True) as status:

            st.write("📥 수입 데이터 수집 중...")
            datasets = fetch_trade_data(selected["code"], api_key)
            cur_records = datasets.get(max(datasets.keys()), [])
            if not cur_records:
                status.update(label="데이터 수집 실패", state="error")
                st.error("데이터를 가져오지 못했습니다. API 키를 확인하거나 잠시 후 다시 시도해주세요.")
                st.stop()

            st.write("📊 타겟 시장 스코어링 중...")
            ranked = score_countries(datasets)
            if len(ranked) < 3:
                status.update(label="데이터 부족", state="error")
                st.stop()
            results["ranked"]   = ranked
            results["datasets"] = datasets

            st.write("🥊 수출 경쟁사 분석 중...")
            export_records, export_year = fetch_export_data(selected["code"], api_key)
            competitor = analyze_competitors(export_records, export_year)
            competitor["hs_desc"] = selected["desc"]
            for c in competitor.get("top3", []):
                c["companies"] = search_competitor_companies(c["name"], selected["desc"])
            results["competitor"] = competitor

            st.write(f"🔍 {ranked[0]['name']} 바이어 검색 중...")
            target   = ranked[0]["name"]
            channels = get_buyer_channels(selected["code"], target)
            buyers   = search_real_buyers(target, selected["desc"])
            results["channels"] = channels
            results["buyers"]   = buyers
            results["target"]   = target

            st.write("📄 Word 보고서 생성 중...")
            report_buf = io.BytesIO()
            from docx import Document
            tmp_path = f"/tmp/report_{selected['code']}.docx"
            generate_word_report(
                ranked[:3], selected, competitor, channels, buyers,
                output_path=tmp_path
            )
            with open(tmp_path, "rb") as f:
                report_buf = io.BytesIO(f.read())
            results["report_buf"]  = report_buf
            results["report_name"] = f"report_{selected['desc'].split('/')[0].strip().replace(' ','_')}.docx"

            status.update(label="분석 완료 ✅", state="complete", expanded=False)

        st.session_state.results = results

# ─────────────────────────────────────────────
# STEP 3 — 결과 출력
# ─────────────────────────────────────────────
if st.session_state.results:
    res      = st.session_state.results
    selected = st.session_state.selected
    ranked   = res["ranked"]
    datasets = res["datasets"]
    cur_yr   = max(datasets.keys())
    prv_yr   = min(datasets.keys())

    st.divider()

    # ── 다운로드 버튼 (상단 고정)
    st.download_button(
        label="📥 Word 보고서 다운로드",
        data=res["report_buf"],
        file_name=res["report_name"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
        use_container_width=True,
    )

    st.divider()

    # ── 섹션 1: 수출 타겟 국가 TOP 3
    st.subheader(f"🎯 수출 타겟 국가 TOP 3  |  {selected['desc']}")
    st.caption(f"HS Code {selected['code']}  |  분석 기간 {prv_yr}~{cur_yr}  |  수입규모 70% + 성장률 30% 스코어링")

    cols = st.columns(3)
    medals = ["🥇", "🥈", "🥉"]
    for i, c in enumerate(ranked[:3]):
        with cols[i]:
            g_sign = "+" if c["growth"] >= 0 else ""
            g_str  = f"{g_sign}{c['growth']:.1f}%" if c["prv_val"] > 0 else "N/A"
            st.metric(
                label=f"{medals[i]} {c['name']}",
                value=fmt_usd(c["cur_val"]),
                delta=g_str,
            )
            st.progress(int(c["total"]))
            st.caption(f"종합 점수: {c['total']:.1f}점")
            with st.expander("선정 근거"):
                st.write(build_reason(c, i + 1).split("] ", 1)[-1])

    st.divider()

    # ── 섹션 2: 경쟁사 분석
    competitor = res["competitor"]
    if competitor and competitor.get("top3"):
        st.subheader(f"🥊 수출 경쟁사 분석  |  {competitor['year']}년 기준")
        kr = competitor["korea"]
        if kr["rank"] and kr["val"] > 0:
            st.info(f"한국 현황: 세계 **{kr['rank']}위**  |  수출 **{fmt_usd(kr['val'])}**  |  점유율 **{kr['share']:.1f}%**")

        intensity_color = {"높음": "🔴", "중간": "🟡", "낮음": "🟢"}
        for i, c in enumerate(competitor["top3"]):
            icon = intensity_color.get(c["intensity"], "⚪")
            with st.expander(f"{['1위','2위','3위'][i]}  {c['name']}  —  수출 {fmt_usd(c['export_val'])}  |  점유율 {c['share']:.1f}%  {icon}", expanded=(i==0)):
                if c.get("companies"):
                    for comp in c["companies"]:
                        st.markdown(f"- [{comp['name']}]({comp['url']})")
                else:
                    st.caption("검색된 기업 없음")

    st.divider()

    # ── 섹션 3: 바이어 추천
    st.subheader(f"🤝 바이어 추천  |  {res['target']}")

    tab1, tab2 = st.tabs(["실제 바이어 기업", "유통 채널 전략"])

    with tab1:
        buyers = res["buyers"]
        if buyers:
            for i, b in enumerate(buyers[:3]):
                st.markdown(f"**{['①','②','③'][i]} {b['name']}**")
                st.markdown(f"🔗 {b['url']}")
                if b.get("reason"):
                    st.caption(b["reason"][:150])
                st.divider()
        else:
            st.caption("검색된 바이어 기업이 없습니다.")
        kotra_url = "https://www.kotra.or.kr/foreign/buyer/KTMITR060M.do"
        st.info(f"추가 바이어는 [KOTRA 바이어 DB]({kotra_url}) 에서 검색하세요.")

    with tab2:
        for i, ch in enumerate(res["channels"]):
            st.markdown(f"**{'①②③'[i]}  {ch['type']}**")
            st.caption(ch["desc"])
            st.markdown(f"→ **접근 전략:** {ch['strategy']}")
            if i < 2:
                st.divider()
