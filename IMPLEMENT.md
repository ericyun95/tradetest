# IMPLEMENT.md — 구현 에이전트 하네스 (v2)

## ⚠️ 이 파일을 읽는 Claude Code에게

이 문서는 **구현 에이전트의 하네스**입니다.
**`CLAUDE.md` → `spec.md` → `DESIGN.md`를 먼저 읽었다고 가정합니다.** 스펙은 spec.md가 정답. 재정의 금지.
아래 순서대로만 구현하세요. 순서를 바꾸거나 추가 기능을 구현하지 마세요.
각 단계 완료 후 **체크리스트를 직접 업데이트**하세요.

---

## 구현 순서 (이 순서 엄수)

- [x] Step 1: 프로젝트 초기화
- [x] Step 2: `utils/hs_validator.py`
- [x] Step 3: `agents/parser.py`
- [x] Step 4: `agents/top_markets.py`
- [x] Step 5: `agents/buyer_research.py`
- [x] Step 6: `agents/report_gen.py`
- [x] Step 7: `app.py`
- [x] Step 8: 배포 설정

---

## Step 1: 프로젝트 초기화

### 1-1. 디렉토리 및 파일 생성

```bash
mkdir -p trade_mentor/agents trade_mentor/utils trade_mentor/.streamlit
touch trade_mentor/agents/__init__.py
touch trade_mentor/utils/__init__.py
```

### 1-2. requirements.txt

```txt
streamlit==1.45.0
anthropic==0.52.0
requests==2.32.3
```

### 1-3. .streamlit/secrets.toml (예시)

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
COMTRADE_API_KEY  = "your-comtrade-key-here"
```

### 1-4. .gitignore

```
.streamlit/secrets.toml
__pycache__/
*.pyc
.env
```

---

## Step 2: utils/hs_validator.py

역할: HS Code 유효성 검증 + 국가코드 매핑 (양방향)

```python
def validate_hs_code(code: str) -> tuple[bool, str]:
    """
    반환: (유효여부, 정제된 6자리 코드)
    - 점/공백 제거 후 숫자만 추출
    - 4자리면 뒤에 00 추가
    - 6자리가 아니면 False 반환
    """

COUNTRY_CODE_MAP = {
    "USA": 842, "VNM": 704, "DEU": 276, "JPN": 392,
    "CHN": 156, "IDN": 360, "THA": 764, "IND": 356,
    "AUS": 36,  "GBR": 826, "FRA": 250, "CAN": 124,
    "BRA": 76,  "MYS": 458, "PHL": 608, "SGP": 702,
    "SAU": 682, "ARE": 784, "MEX": 484, "NLD": 528,
}

NUMERIC_TO_NAME = {
    842: "United States", 704: "Viet Nam",    276: "Germany",
    392: "Japan",         156: "China",        360: "Indonesia",
    764: "Thailand",      356: "India",        36:  "Australia",
    826: "United Kingdom",250: "France",       124: "Canada",
    76:  "Brazil",        458: "Malaysia",     608: "Philippines",
    702: "Singapore",     682: "Saudi Arabia", 784: "UAE",
    484: "Mexico",        528: "Netherlands",
}

def get_all_reporter_codes() -> str:
    """20개국 코드를 콤마 구분 문자열로 반환 (Comtrade API 파라미터용)"""
    return ",".join(str(v) for v in COUNTRY_CODE_MAP.values())

def get_country_name(numeric_code: int) -> str:
    """숫자 코드 → 국가 영문명. 없으면 str(numeric_code) 반환."""
    return NUMERIC_TO_NAME.get(numeric_code, str(numeric_code))
```

---

## Step 3: agents/parser.py

역할: 사용자 입력 → Claude API 호출 → 표준화된 dict 반환

### 함수 시그니처

```python
def parse_input(user_input: str, anthropic_client) -> dict:
    """
    반환 형식:
    {
        "input_type": str,        # "hs_code" | "product_name"
        "hs_code": str,           # 6자리 숫자
        "hs_code_display": str,   # "xxxx.xx"
        "product_name_en": str,
        "product_name_ko": str,
        "confidence": str,        # "high" | "medium" | "low"
        "note": str | None
    }
    예외 시: ValueError 발생
    """
```

### 구현 규칙

- spec.md의 **프롬프트 #1** 그대로 사용 (수정 금지)
- `model="claude-sonnet-4-6"`, `max_tokens=300`
- 응답에서 JSON만 추출: `response.content[0].text`를 `json.loads()`
- JSON 파싱 실패 시 `ValueError("LLM 응답 파싱 실패")` 발생
- `confidence == "low"` 처리는 parser가 아닌 app.py에서 담당 (그냥 반환)

---

## Step 4: agents/top_markets.py

역할: UN Comtrade API → 20개국 일괄 조회 → Top 5 수입국 추출

### 4-1. 주요 함수

```python
def fetch_top_markets(
    hs_code: str,        # 6자리
    api_key: str,
    years: list[int] = None   # 기본: 최근 3개년
) -> list[dict]:
    """
    반환: Top 5 수입국 리스트 (금액 기준 내림차순)
    [
        {
            "rank": 1,
            "country_name": "United States",
            "numeric_code": 842,
            "total_value_usd": 1234567890,
            "by_year": {"2022": 400000000, "2023": 450000000, "2024": 384567890},
            "yoy_change_pct": -14.5    # 최신 2개년 대비
        },
        ...
    ]
    빈 리스트 반환 가능 (데이터 없을 시)
    """
```

### 4-2. 요청 구성

```python
BASE_URL = "https://comtradeapi.un.org/data/v1/get/C/A/HS"
params = {
    "reporterCode": get_all_reporter_codes(),   # 20개국 콤마 구분
    "period": ",".join(str(y) for y in years),
    "cmdCode": hs_code,
    "flowCode": "M",
    "maxRecords": 500,
    "format": "JSON",
}
headers = {"Ocp-Apim-Subscription-Key": api_key}
```

### 4-3. 집계 로직

```python
# 1. 응답 데이터 순회 → reporterCode 기준으로 primaryValue 연도별 합산
# 2. total = sum(by_year.values())
# 3. total 기준 내림차순 정렬
# 4. 상위 5개 슬라이싱
# 5. rank 필드 추가 (1~5)
# 6. yoy_change_pct 계산: (최신연도 - 전년도) / 전년도 * 100
#    전년도 데이터 없을 시 None
```

### 4-4. 예외 처리

```python
# HTTP 오류: response.raise_for_status()
# 빈 응답: data.get("data", []) == [] → 빈 리스트 반환
# primaryValue가 None인 레코드: 0으로 처리
```

---

## Step 5: agents/buyer_research.py

역할: Claude web_search 툴 → Top 5 국가별 바이어 후보 기업 조사

### 함수 시그니처

```python
def research_buyers(
    product_name_en: str,
    hs_code_display: str,
    top_markets: list[dict],    # fetch_top_markets() 반환값
    anthropic_client
) -> list[dict]:
    """
    반환:
    [
        {
            "country": "United States",
            "company_name": "...",
            "description": "...",
            "source": "..."
        },
        ...
    ]
    웹 검색 실패 시 빈 리스트 반환 (예외 발생 금지)
    """
```

### 구현 규칙

- spec.md의 **프롬프트 #2** 그대로 사용
- `model="claude-sonnet-4-6"`, `max_tokens=1500`
- `tools=[{"type": "web_search_20250305", "name": "web_search"}]`
- 응답에서 JSON 추출: tool_use 블록이 아닌 최종 text 블록에서 파싱
- JSON 파싱 실패 시 빈 리스트 반환 (앱 중단 금지)

### tool_use 응답 처리

```python
# response.content를 순회하여 type == "text"인 마지막 블록에서 JSON 추출
# json.loads() 실패 시 [] 반환
buyers = []
for block in response.content:
    if block.type == "text":
        try:
            result = json.loads(block.text)
            buyers = result.get("buyers", [])
        except json.JSONDecodeError:
            pass
return buyers
```

---

## Step 6: agents/report_gen.py

역할: 파싱 결과 + Top 5 데이터 + 바이어 조사 결과 → 종합 보고서 마크다운

### 함수 시그니처

```python
def generate_report(
    parsed: dict,               # parse_input() 반환값
    top_markets: list[dict],    # fetch_top_markets() 반환값
    buyers: list[dict],         # research_buyers() 반환값
    anthropic_client
) -> str:
    """
    반환: 마크다운 형식 보고서 문자열
    top_markets가 빈 리스트면 → "수입 데이터를 조회할 수 없어 보고서를 생성할 수 없습니다." 반환
    """
```

### 구현 규칙

- spec.md의 **프롬프트 #3** 그대로 사용
- `model="claude-sonnet-4-6"`, `max_tokens=2000`
- top_markets와 buyers를 `json.dumps(, ensure_ascii=False, indent=2)`로 직렬화하여 프롬프트에 삽입
- 반환값: `response.content[0].text`

---

## Step 7: app.py

### 7-1. 초기화 블록

```python
import streamlit as st
import anthropic
from agents.parser import parse_input
from agents.top_markets import fetch_top_markets
from agents.buyer_research import research_buyers
from agents.report_gen import generate_report

st.set_page_config(
    page_title="Global Trade Mentor AI",
    page_icon="🌐",
    layout="wide"
)

@st.cache_resource
def get_anthropic_client():
    return anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])
```

### 7-2. 사이드바

```python
with st.sidebar:
    st.title("🌐 Global Trade Mentor AI")
    st.caption("품목명 또는 HS Code를 입력하면 수입 현황과 바이어 후보를 분석합니다.")

    user_input = st.text_input(
        "품목명 또는 HS Code",
        placeholder="예: 에어프라이어 / 8516.60"
    )
    run_button = st.button("보고서 생성", type="primary", use_container_width=True)
```

### 7-3. 메인 실행 흐름

```python
if run_button:
    if not user_input:
        st.error("품목명 또는 HS Code를 입력해주세요.")
        st.stop()

    client = get_anthropic_client()

    # Step A: 입력 파싱
    with st.spinner("HS Code 분석 중..."):
        parsed = parse_input(user_input, client)

    if parsed.get("confidence") == "low":
        st.warning(f"⚠️ {parsed.get('note', 'HS Code 확인 필요')}")

    st.info(
        f"**품목**: {parsed['product_name_ko']} ({parsed['product_name_en']}) | "
        f"**HS Code**: {parsed['hs_code_display']}"
    )

    # Step B: Top 5 수입국 조회
    with st.spinner("UN Comtrade 수입 데이터 조회 중..."):
        top_markets = fetch_top_markets(
            hs_code=parsed["hs_code"],
            api_key=st.secrets["COMTRADE_API_KEY"]
        )

    if not top_markets:
        st.error("UN Comtrade에서 해당 품목의 수입 데이터를 찾을 수 없습니다.")
        st.stop()

    # Step C: 바이어 후보 조사
    with st.spinner("바이어 후보 기업 조사 중..."):
        buyers = research_buyers(
            product_name_en=parsed["product_name_en"],
            hs_code_display=parsed["hs_code_display"],
            top_markets=top_markets,
            anthropic_client=client
        )

    # Step D: 보고서 생성
    with st.spinner("종합 보고서 생성 중..."):
        report = generate_report(
            parsed=parsed,
            top_markets=top_markets,
            buyers=buyers,
            anthropic_client=client
        )

    st.markdown(report)
    st.caption("출처: UN Comtrade | 기업 정보: 웹 검색 기반")
```

---

## Step 8: 배포 설정

### Streamlit Cloud secrets 등록

Streamlit Cloud 대시보드 → App settings → Secrets:

```toml
ANTHROPIC_API_KEY = "sk-ant-..."
COMTRADE_API_KEY  = "..."
```

### requirements.txt 최종 확인

```txt
streamlit==1.45.0
anthropic==0.52.0
requests==2.32.3
```

---

## 구현 금지 사항

- 위 Step 외 추가 기능 구현 금지
- 사용자 국가 지정 UI 추가 금지
- 이메일 생성 기능 추가 금지
- 데이터베이스 연동 금지
- spec.md 프롬프트 임의 수정 금지
- 스트리밍 응답(`stream=True`) 사용 금지

## 구현 완료 기준

- [ ] `streamlit run app.py` 로컬 실행 성공 (API 키 입력 후 확인)
- [ ] 한글 품목명 입력 시 HS Code 추론 동작
- [ ] HS Code 입력 시 품목명 역추론 동작
- [ ] Top 5 수입국 정상 추출 및 순위 정렬 확인
- [ ] 바이어 후보 기업 1건 이상 반환 확인
- [ ] 보고서 5개 섹션 모두 포함 확인

**완료 후 → VERIFY.md로 이동**
