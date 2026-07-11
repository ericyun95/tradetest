# spec.md — 기술 스펙 전문 (v2)

> 이 파일은 모든 에이전트가 참조하는 단일 진실 공급원(Single Source of Truth)입니다.
> 에이전트는 이 파일을 수정하지 않습니다. 읽기 전용.

---

## 1. 구현 범위

| 포함 ✅ | 제외 ❌ |
|---|---|
| HS Code ↔ 품목명 정확 매핑 | 사용자 국가 지정 기능 |
| UN Comtrade 최대 수입국 Top 5 자동 발굴 | 영문 이메일 생성 |
| 바이어 후보 기업 조사 (Claude 웹 검색) | 바이어 스코어링 |
| 종합 보고서 생성 (마크다운) | 로그인/회원가입 |
| Streamlit 웹 UI | 데이터 저장/이력 관리 |

---

## 2. 파일 구조

```
trade_mentor/
├── app.py                      # Streamlit 진입점
├── agents/
│   ├── __init__.py
│   ├── parser.py               # 품목명/HS Code → 표준화 (Claude API)
│   ├── top_markets.py          # UN Comtrade → Top 5 수입국 추출
│   ├── buyer_research.py       # 웹 검색 → 바이어 후보 기업
│   └── report_gen.py           # 종합 보고서 생성 (Claude API)
├── utils/
│   ├── __init__.py
│   └── hs_validator.py         # HS Code 검증 + 국가코드 매핑
├── .streamlit/
│   └── secrets.toml
├── .gitignore
└── requirements.txt
```

---

## 3. 데이터 흐름

```
[사용자 입력]
  품목명(한글/영문) 또는 HS Code
        ↓
[agents/parser.py]
  Claude API #1 (max_tokens=300)
  → HS Code 6자리 확정
  → 품목명(한글/영문) 확정
        ↓
[agents/top_markets.py]
  UN Comtrade API — 20개국 일괄 조회
  → 국가별 수입 총액 집계 (최근 3개년)
  → Top 5 수입국 추출 (금액 기준 내림차순)
        ↓
[agents/buyer_research.py]
  Claude API #2 (web_search 툴, max_tokens=1500)
  → Top 5 국가별 바이어 후보 기업 2~3사 조사
        ↓
[agents/report_gen.py]
  Claude API #3 (max_tokens=2000)
  → 5개 섹션 종합 보고서 마크다운 생성
        ↓
[app.py — Streamlit UI]
  단일 스크롤 보고서 표시
```

---

## 4. 입력 처리 스펙

### 4-1. 입력 타입 판별

```
숫자 + 점(.) 조합으로만 구성 → HS Code
예: "8516.60", "851660", "85.16.60" → HS Code 처리
그 외 → 품목명 (한글/영문 무관)
```

### 4-2. Claude 프롬프트 #1 — 입력 파싱

```
SYSTEM:
당신은 무역 HS Code 전문가입니다.
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 출력하지 마세요.

USER:
다음 입력을 분석하세요: "{user_input}"

다음 JSON을 반환하세요:
{
  "input_type": "hs_code" | "product_name",
  "hs_code": "6자리 숫자 문자열 (점 없이)",
  "hs_code_display": "xxxx.xx 형식",
  "product_name_en": "영문 품목명 (간결하게)",
  "product_name_ko": "한글 품목명",
  "confidence": "high" | "medium" | "low",
  "note": null | "confidence가 low일 때 이유"
}

규칙:
- HS Code는 반드시 6자리. 4자리 입력 시 뒤에 00 추가.
- confidence가 low인 경우 가장 유력한 코드를 넣고 이유를 note에 추가.
- 절대 JSON 외 텍스트 출력 금지.

model: claude-sonnet-4-6
max_tokens: 300
```

---

## 5. UN Comtrade API 스펙

### 5-1. Top Markets 조회 방식

```
GET https://comtradeapi.un.org/data/v1/get/C/A/HS

params:
  reporterCode : "842,704,276,392,156,360,764,356,36,826,250,124,76,458,608,702,682,784,484,528"
                 (20개국 콤마 구분, 일괄 조회)
  period       : 최근 3개년 (예: "2022,2023,2024") — 현재연도-3 ~ 현재연도-1
  cmdCode      : HS 6자리
  flowCode     : "M" (수입)
  maxRecords   : 500
  format       : "JSON"

headers:
  Ocp-Apim-Subscription-Key: {COMTRADE_API_KEY}
```

### 5-2. Top 5 추출 로직

```python
# 응답 데이터에서 국가별 primaryValue 합산 (3개년 합계)
# 구조: {reporterCode: {"name": reporterDesc, "total": sum, "by_year": {year: value}}}
# total 기준 내림차순 정렬 → 상위 5개 선택
```

### 5-3. 추출 필드

```python
FIELDS = ["period", "reporterCode", "reporterDesc", "primaryValue", "netWgt"]
```

### 5-4. 예외 처리

| 상황 | 처리 |
|---|---|
| 빈 배열 반환 | "조회된 수입 데이터가 없습니다" 반환, Claude 호출 스킵 |
| HTTP 4xx/5xx | requests.HTTPError 발생 → st.error() |
| API Key 없음 | st.error() 후 st.stop() |

---

## 6. 국가코드 매핑 테이블

```python
COUNTRY_CODE_MAP = {
    "USA": 842, "VNM": 704, "DEU": 276, "JPN": 392,
    "CHN": 156, "IDN": 360, "THA": 764, "IND": 356,
    "AUS": 36,  "GBR": 826, "FRA": 250, "CAN": 124,
    "BRA": 76,  "MYS": 458, "PHL": 608, "SGP": 702,
    "SAU": 682, "ARE": 784, "MEX": 484, "NLD": 528,
}

# 역방향 매핑 (Comtrade 응답 코드 → 국가명 조회용)
NUMERIC_TO_NAME = {
    842: "United States", 704: "Viet Nam", 276: "Germany",
    392: "Japan",         156: "China",    360: "Indonesia",
    764: "Thailand",      356: "India",    36:  "Australia",
    826: "United Kingdom",250: "France",   124: "Canada",
    76:  "Brazil",        458: "Malaysia", 608: "Philippines",
    702: "Singapore",     682: "Saudi Arabia", 784: "UAE",
    484: "Mexico",        528: "Netherlands",
}
```

---

## 7. Claude 프롬프트 #2 — 바이어 후보 조사

```
SYSTEM:
You are a trade research assistant.
Use web search to find B2B importers, distributors, and buyers.
Return JSON only. No other text.

USER:
Product: {product_name_en} (HS Code: {hs_code_display})
Top importing countries: {top_5_country_names}

For each country, search for 2-3 major importing companies or distributors of this product.
Return:
{
  "buyers": [
    {
      "country": "country name",
      "company_name": "...",
      "description": "industry/specialty, one line",
      "source": "URL or publication"
    }
  ]
}

Rules:
- Only include companies verifiable via web search
- Do not invent or guess company names
- If no result found for a country, skip it

model: claude-sonnet-4-6
max_tokens: 1500
tools: [{"type": "web_search_20250305", "name": "web_search"}]
```

---

## 8. Claude 프롬프트 #3 — 종합 보고서 생성

```
SYSTEM:
당신은 무역 인텔리전스 분석가입니다.
주어진 데이터를 바탕으로 한국어 종합 보고서를 마크다운 형식으로 작성하세요.
섹션 번호와 제목을 반드시 포함하세요.

USER:
## 분석 대상
품목: {product_name_ko} ({product_name_en})
HS Code: {hs_code_display}

## UN Comtrade 수입 데이터
{top_markets_json}

## 바이어 후보 조사 결과
{buyer_research_json}

아래 5개 섹션으로 보고서를 작성하세요:

### 1. HS Code 확인
품목명(한/영), 6자리 코드, 코드 의미 1줄 설명

### 2. 수입 규모 추이
20개국 기준 최근 3개년 연도별 총 수입액 (USD), 전년 대비 성장률

### 3. 상위 수입국 Top 5
표 형식: 순위 | 국가명 | 수입액(USD) | 점유율 | 전년 대비 증감

### 4. 바이어 후보 기업
국가별 2~3개사: 기업명 | 설명 | 출처

### 5. 시장 진입 분석
상위 2개국 시장 특징 + 한국 수출 시 고려할 점 2~3가지

model: claude-sonnet-4-6
max_tokens: 2000
```

---

## 9. Streamlit UI 스펙

### 사이드바
- 앱 제목: "🌐 Global Trade Mentor AI"
- 입력: 품목명 또는 HS Code (placeholder: "예: 에어프라이어 / 8516.60")
- 버튼: "보고서 생성" (type="primary")

### 메인 영역
- 파싱 결과: `st.info()` 1줄 (HS Code + 품목명 확인)
- confidence == "low": `st.warning()` 추가 표시
- 보고서: `st.markdown()` 전체 출력 (탭 없이 단일 스크롤)
- 하단 caption: "출처: UN Comtrade | 기업 정보: 웹 검색 기반"

### spinner 텍스트
- 파싱 중: "HS Code 분석 중..."
- Comtrade 조회 중: "UN Comtrade 수입 데이터 조회 중..."
- 바이어 조사 중: "바이어 후보 기업 조사 중..."
- 보고서 생성 중: "종합 보고서 생성 중..."

---

## 10. 환경변수

```toml
# .streamlit/secrets.toml
ANTHROPIC_API_KEY = "sk-ant-..."
COMTRADE_API_KEY  = "..."
```

```
# .gitignore 필수 포함
.streamlit/secrets.toml
__pycache__/
*.pyc
.env
```
