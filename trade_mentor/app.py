import sys
import os

sys.path.insert(0, os.path.dirname(__file__))

import time
import streamlit as st
import anthropic
from agents.parser import parse_input
from agents.top_markets import fetch_top_markets
from agents.buyer_research import research_buyers
from agents.report_gen import generate_report
from utils.docx_export import markdown_to_docx

st.set_page_config(
    page_title="Global Trade Mentor AI",
    page_icon="🌐",
    layout="wide",
)


@st.cache_resource
def get_anthropic_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])


with st.sidebar:
    st.title("🌐 Global Trade Mentor AI")
    st.caption("품목명 또는 HS Code를 입력하면 수입 현황과 바이어 후보를 분석합니다.")

    user_input = st.text_input(
        "품목명 또는 HS Code",
        placeholder="예: 에어프라이어 / 8516.60",
    )
    run_button = st.button("보고서 생성", type="primary", use_container_width=True)

if run_button:
    if not user_input.strip():
        st.error("품목명 또는 HS Code를 입력해주세요.")
        st.stop()

    client = get_anthropic_client()

    with st.spinner("HS Code 분석 중..."):
        try:
            parsed = parse_input(user_input.strip(), client)
        except ValueError as e:
            st.error(f"입력 분석 실패: {e}")
            st.stop()

    if parsed.get("confidence") == "low":
        st.warning(f"⚠️ {parsed.get('note', 'HS Code를 특정하기 어렵습니다. 결과를 확인해주세요.')}")

    st.info(
        f"**품목**: {parsed['product_name_ko']} ({parsed['product_name_en']})  |  "
        f"**HS Code**: {parsed['hs_code_display']}"
    )

    with st.spinner("UN Comtrade 수입 데이터 조회 중..."):
        try:
            top_markets = fetch_top_markets(
                hs_code=parsed["hs_code"],
                api_key=st.secrets["COMTRADE_API_KEY"],
            )
        except Exception as e:
            st.error(f"UN Comtrade 조회 오류: {e}")
            st.stop()

    if not top_markets:
        st.error("UN Comtrade에서 해당 품목의 수입 데이터를 찾을 수 없습니다.")
        st.stop()

    time.sleep(5)
    with st.spinner("바이어 후보 기업 조사 중..."):
        buyers = research_buyers(
            product_name_en=parsed["product_name_en"],
            hs_code_display=parsed["hs_code_display"],
            top_markets=top_markets,
            anthropic_client=client,
        )

    time.sleep(5)
    with st.spinner("종합 보고서 생성 중..."):
        report = generate_report(
            parsed=parsed,
            top_markets=top_markets,
            buyers=buyers,
            anthropic_client=client,
        )

    st.markdown(report)
    st.caption("출처: UN Comtrade  |  기업 정보: 웹 검색 기반")

    docx_bytes = markdown_to_docx(
        report,
        product_name_ko=parsed["product_name_ko"],
        hs_code_display=parsed["hs_code_display"],
    )
    filename = f"report_{parsed['hs_code']}_{parsed['product_name_ko']}.docx"
    st.download_button(
        label="📥 보고서 Word 다운로드",
        data=docx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
