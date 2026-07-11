# VERIFY.md — 검증 에이전트 하네스 (v2)

## ⚠️ 이 파일을 읽는 Claude Code에게

이 문서는 **검증 에이전트의 하네스**입니다.
DESIGN.md + IMPLEMENT.md 기반 구현이 완료된 후 실행합니다.
실패한 테스트에 대해서만 수정 제안을 작성하세요. 통과 항목은 설명 없이 ✅만 표기.

---

## 검증 범위

| 범위 | 포함 | 제외 |
|---|---|---|
| HS Code 파싱 정확도 | ✅ | UI 디자인 검토 |
| Top 5 수입국 추출 | ✅ | 성능 벤치마크 |
| 바이어 후보 조사 | ✅ | 보안 취약점 점검 |
| 보고서 완전성 (5개 섹션) | ✅ | 다국어 보고서 검증 |
| 예외 처리 | ✅ | 부하 테스트 |

---

## 테스트 케이스

### TC-01: 한글 품목명 입력 → HS Code 추론

```
입력:
  user_input = "에어프라이어"

기대 출력 (parser):
  hs_code: "851660" (또는 "851640" — 둘 다 허용)
  product_name_en: "Air Fryer" (대소문자 무관)
  confidence: "high" 또는 "medium"

판정 기준:
  ✅ PASS: hs_code가 6자리 숫자
  ✅ PASS: product_name_en에 "air fryer" 포함 (대소문자 무관)
  ❌ FAIL: hs_code가 6자리 미만이거나 비숫자 포함
```

---

### TC-02: HS Code 직접 입력 → 품목명 역추론

```
입력:
  user_input = "8516.60"

기대 출력 (parser):
  input_type: "hs_code"
  hs_code: "851660"
  hs_code_display: "8516.60"
  product_name_ko: 조리기기 관련 단어 포함

판정 기준:
  ✅ PASS: input_type == "hs_code"
  ✅ PASS: hs_code == "851660"
  ✅ PASS: product_name_ko 비어있지 않음
  ✅ PASS: hs_code_display == "8516.60"
```

---

### TC-03: 영문 품목명 입력

```
입력:
  user_input = "cosmetics"

기대 출력:
  hs_code 앞 2자리: "33" (화장품류)
  product_name_ko: 비어있지 않음

판정 기준:
  ✅ PASS: hs_code 앞 2자리 == "33"
  ✅ PASS: hs_code가 6자리 숫자
  ❌ FAIL: hs_code가 화장품과 무관한 코드
```

---

### TC-04: 점 없는 HS Code 입력

```
입력:
  user_input = "070200"

기대 출력:
  input_type: "hs_code"
  hs_code: "070200"
  hs_code_display: "0702.00"

판정 기준:
  ✅ PASS: hs_code == "070200"
  ✅ PASS: hs_code_display == "0702.00"
```

---

### TC-05: UN Comtrade Top Markets 추출

```
입력 (top_markets.py 직접 호출):
  hs_code = "851660"

판정 기준:
  ✅ PASS: 반환 리스트 길이 1 이상 (최대 5)
  ✅ PASS: 각 항목에 "rank", "country_name", "total_value_usd", "by_year" 키 존재
  ✅ PASS: rank 기준 total_value_usd 내림차순 정렬
  ✅ PASS: rank 값이 1부터 시작하는 연속 정수
  ⚠️ WARN: 빈 리스트 반환 시 → app.py에서 st.error() 출력 여부 확인
  ❌ FAIL: 예외 미처리로 앱 크래시
```

---

### TC-06: 바이어 후보 기업 조사

```
입력 (buyer_research.py 직접 호출):
  product_name_en = "Air Fryer"
  hs_code_display = "8516.60"
  top_markets = [{"country_name": "United States", ...}, ...]  # TC-05 결과 활용

판정 기준:
  ✅ PASS: 반환 리스트 길이 1 이상
  ✅ PASS: 각 항목에 "country", "company_name", "description" 키 존재
  ✅ PASS: company_name이 비어있지 않음
  ⚠️ WARN: 웹 검색 실패 시 빈 리스트 반환 (앱 크래시 금지)
  ❌ FAIL: 회사명이 명백히 허위이거나 "[Company]" 같은 플레이스홀더 포함
```

---

### TC-07: 보고서 완전성 검증

```
입력 (generate_report() 직접 호출):
  parsed = TC-01 결과
  top_markets = TC-05 결과
  buyers = TC-06 결과

판정 기준:
  ✅ PASS: 보고서에 "### 1." 또는 "## 1." 포함 (섹션 1 존재)
  ✅ PASS: "### 2." 또는 "## 2." 포함 (섹션 2 존재)
  ✅ PASS: "### 3." 또는 "## 3." 포함 (섹션 3 존재)
  ✅ PASS: "### 4." 또는 "## 4." 포함 (섹션 4 존재)
  ✅ PASS: "### 5." 또는 "## 5." 포함 (섹션 5 존재)
  ✅ PASS: 보고서에 표(|) 형식 1개 이상 포함
  ❌ FAIL: 5개 섹션 중 하나라도 누락
  ❌ FAIL: 영어 본문 포함 (섹션 제목 제외)
```

---

### TC-08: 예외 처리 — 빈 입력

```
입력:
  user_input = ""

기대 동작:
  st.error() 메시지 출력 후 st.stop()
  Claude API 호출 없음

판정 기준:
  ✅ PASS: 에러 메시지 표시
  ✅ PASS: API 호출 발생하지 않음
  ❌ FAIL: 앱 크래시 또는 빈 결과 표시
```

---

## 자동 검증 스크립트 (TC-01 ~ TC-04)

```python
# verify_run.py
import os
import anthropic
from agents.parser import parse_input

client = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])

TEST_CASES = [
    ("에어프라이어", {"hs_code_prefix": "851", "name_contains": "air fryer"}),
    ("8516.60",      {"input_type": "hs_code", "hs_code": "851660", "hs_code_display": "8516.60"}),
    ("cosmetics",    {"hs_code_prefix": "33"}),
    ("070200",       {"hs_code": "070200", "hs_code_display": "0702.00"}),
]

for user_input, expected in TEST_CASES:
    result = parse_input(user_input, client)
    errors = []

    if "hs_code_prefix" in expected:
        if not result["hs_code"].startswith(expected["hs_code_prefix"]):
            errors.append(f"hs_code prefix: got {result['hs_code']}")
    if "hs_code" in expected:
        if result["hs_code"] != expected["hs_code"]:
            errors.append(f"hs_code: expected {expected['hs_code']}, got {result['hs_code']}")
    if "hs_code_display" in expected:
        if result["hs_code_display"] != expected["hs_code_display"]:
            errors.append(f"hs_code_display: expected {expected['hs_code_display']}, got {result['hs_code_display']}")
    if "input_type" in expected:
        if result["input_type"] != expected["input_type"]:
            errors.append(f"input_type: expected {expected['input_type']}, got {result['input_type']}")
    if "name_contains" in expected:
        if expected["name_contains"] not in result["product_name_en"].lower():
            errors.append(f"product_name_en: '{expected['name_contains']}' not in '{result['product_name_en']}'")

    if errors:
        print(f"❌ FAIL: '{user_input}' → {errors}")
    else:
        print(f"✅ PASS: '{user_input}' → {result['hs_code_display']}")
```

실행:
```bash
cd trade_mentor
ANTHROPIC_API_KEY=sk-ant-... python verify_run.py
```

---

## 검증 판정 기준

| 결과 | 조건 |
|---|---|
| **PASS** | TC-01 ~ TC-08 모두 통과 |
| **CONDITIONAL PASS** | TC-06에서 빈 리스트 반환 (웹 검색 일시 불가) — 앱 크래시 없으면 허용 |
| **FAIL** | TC-01, TC-02, TC-05, TC-07, TC-08 중 하나라도 실패 |

## 검증 완료 후

- PASS → Streamlit Cloud 배포 진행
- FAIL → 실패 케이스만 IMPLEMENT.md로 돌아가 해당 Step 재구현
