# DESIGN.md — 설계 에이전트 하네스 (v2)

## ⚠️ 이 파일을 읽는 Claude Code에게

**Phase 1 전용 하네스입니다.**
이 파일을 읽기 전 `spec.md`를 반드시 먼저 읽으세요.
설계 에이전트의 역할은 **스펙의 공백을 발견하고 확인하는 것**입니다.
스펙을 재작성하거나 대안을 제안하지 마세요.

---

## 설계 에이전트 체크리스트

아래 항목을 spec.md 기준으로 검토하고 ✅/❌ 표기하세요.

### A. 데이터 흐름 완결성
- ✅ 사용자 입력 → parser → top_markets → buyer_research → report_gen → UI 흐름이 끊김 없이 연결되는가
- ✅ 각 단계의 입력/출력 타입이 명확히 정의되어 있는가
- ✅ 빈 데이터, API 오류 등 예외 경로가 모두 정의되어 있는가

### B. 프롬프트 완결성
- ✅ 프롬프트 #1 (파싱): JSON 필드가 구현에 필요한 모든 정보를 포함하는가
- ✅ 프롬프트 #2 (바이어 조사): web_search 툴 사용 방식 및 반환 JSON이 명확한가
- ✅ 프롬프트 #3 (보고서): 5개 섹션 구조와 표 형식이 명확히 정의되어 있는가
- ✅ 3개 프롬프트 모두 model과 max_tokens가 지정되어 있는가

### C. API 연동 완결성
- ✅ UN Comtrade 엔드포인트 URL이 명시되어 있는가
- ✅ 20개국 일괄 조회 파라미터(reporterCode 콤마 구분)가 정의되어 있는가
- ✅ Top 5 추출 로직(3개년 합산 → 내림차순 정렬)이 명시되어 있는가
- ✅ NUMERIC_TO_NAME 역방향 매핑이 정의되어 있는가 (Comtrade 응답 파싱용)

### D. 환경 설정 완결성
- ✅ Streamlit secrets 키 이름이 코드와 일치하는가 (`ANTHROPIC_API_KEY`, `COMTRADE_API_KEY`)
- ✅ `.gitignore`에 secrets.toml이 포함되어 있는가
- ✅ requirements.txt 패키지가 확정되어 있는가 (`streamlit`, `anthropic`, `requests`)

---

## 공백 발견 시 처리 방법

1. ❌ 항목을 목록으로 정리
2. 각 공백에 대해 **질문 1개**만 사용자에게 확인 요청
3. 답변 수렴 후 spec.md가 아닌 **이 파일 하단 "보완 사항"** 섹션에 기록
4. 모든 항목 ✅ 확인 후 "Phase 1 완료"를 선언하고 IMPLEMENT.md로 이동

---

## 보완 사항

### v2 재설계 변경 내역 (확정)

1. **국가 입력 제거**: 사용자가 국가를 지정하지 않음. Top 5 수입국 자동 발굴.
2. **이메일 생성 제거**: `email_gen.py` 삭제. `report_gen.py`로 대체.
3. **신규 파일 추가**:
   - `agents/top_markets.py` — UN Comtrade 20개국 일괄 조회 + Top 5 추출
   - `agents/buyer_research.py` — Claude web_search 툴로 바이어 후보 조사
4. **바이어 조사 소스**: UN Comtrade (국가 수준) + Claude 웹 검색 (기업 수준)
5. **보고서 구성**: 5개 섹션 (HS Code 확인 / 수입 추이 / Top 5 국가 / 바이어 후보 / 시장 분석)
6. **모델 ID**: `claude-sonnet-4-6` (고정)
7. **requirements.txt**: `streamlit`, `anthropic`, `requests` 3종 확정

---

## Phase 1 완료 선언

```
✅ Phase 1 완료 (v2 재설계 반영)
- 체크리스트 전 항목 통과 (A/B/C/D 전 항목 ✅)
- 주요 변경: 국가 입력 제거 / 이메일 제거 / Top 5 자동 발굴 / 바이어 웹 검색
- 다음 단계: IMPLEMENT.md Phase 2 시작
```
