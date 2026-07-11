# Claude Code 사용 가이드

## 사전 준비

### 1. API 키 발급
- **Anthropic API Key**: https://console.anthropic.com
- **UN Comtrade API Key**: https://comtradeapi.un.org → 무료 등록 후 발급 (하루 500콜 무료)

### 2. GitHub 저장소 생성 후 파일 업로드
```
trade_mentor/
├── CLAUDE.md        ← 필수
├── spec.md          ← 필수
├── DESIGN.md        ← 필수
├── IMPLEMENT.md     ← 필수
├── VERIFY.md        ← 필수
└── README.md
```
`.streamlit/secrets.toml`은 **절대 커밋하지 말 것** (`.gitignore`에 포함)

### 3. Claude Code 설치 및 프로젝트 열기
```bash
cd trade_mentor
claude
```

---

## Phase 1 — 설계 에이전트 프롬프트

```
CLAUDE.md와 spec.md를 읽고, DESIGN.md의 설계 에이전트로서 체크리스트를 실행해줘.
스펙에 공백이 있으면 알려주고, 없으면 Phase 1 완료를 선언해.
```

**예상 동작**: Claude Code가 CLAUDE.md → spec.md → DESIGN.md 순으로 읽고 체크리스트 항목을 검토. 공백 없으면 "Phase 1 완료" 선언.

**공백이 발견된 경우 예시 응답**:
```
❌ B-3: 프롬프트 #3에 플레이스홀더 금지 조건이 명시되어 있지 않습니다.
→ 질문: 이메일 초안에서 [Company Name] 같은 플레이스홀더를 허용하시겠습니까?
```
→ 답변하면 DESIGN.md 보완 사항 섹션에 기록 후 진행

---

## Phase 2 — 구현 에이전트 프롬프트

```
Phase 1 완료됐어. 이제 IMPLEMENT.md를 읽고 구현 에이전트로서 Step 1부터 순서대로 구현해줘.
각 Step 완료 시 체크리스트 업데이트하고 다음 Step으로 넘어가.
```

**예상 동작**: Step 1(초기화) → Step 2(hs_validator) → Step 3(parser) → Step 4(trade_data) → Step 5(email_gen) → Step 6(app.py) → Step 7(배포 설정) 순으로 파일 생성.

**중간에 막히면 쓸 프롬프트**:
```
Step 4까지 완료됐는데 UN Comtrade API가 429 에러를 반환해. 어떻게 처리할까?
```
→ Claude Code가 재시도 로직 또는 에러 메시지 처리 방법 제안

---

## Phase 3 — 검증 에이전트 프롬프트

```
구현 완료됐어. VERIFY.md를 읽고 검증 에이전트로서 TC-01부터 순서대로 실행해줘.
각 TC 결과를 ✅/❌로 표기하고, 실패한 케이스만 수정 방법을 알려줘.
```

**예상 동작**: TC-01 ~ TC-08 순서대로 실행. 통과 항목은 ✅만 표기, 실패 항목은 원인 + 수정 위치 안내.

**특정 TC만 재실행할 때**:
```
TC-06 이메일 품질 검증만 다시 실행해줘. 수정 후 재테스트야.
```

---

## 배포 — Streamlit Cloud

### 1. Streamlit Cloud 연결
- https://share.streamlit.io → GitHub 저장소 연결
- Main file: `app.py`

### 2. Secrets 등록
Streamlit Cloud 대시보드 → 앱 설정 → Secrets:
```toml
ANTHROPIC_API_KEY = "sk-ant-..."
COMTRADE_API_KEY = "..."
```

### 3. 배포 완료 후 확인 프롬프트
```
배포 링크가 나왔어. 마지막으로 VERIFY.md의 TC-01과 TC-06을 실제 배포 환경에서 수동으로 테스트하는 방법을 알려줘.
```

---

## 자주 쓸 추가 프롬프트

**특정 파일만 수정할 때**:
```
agents/parser.py에서 confidence가 low일 때 처리 방식만 수정해줘. 나머지는 건드리지 마.
```

**스펙 변경이 필요할 때**:
```
spec.md의 국가코드 매핑에 "KOR": 410 추가하고, 관련 코드도 업데이트해줘.
```

**에러 디버깅할 때**:
```
streamlit run app.py 했더니 아래 에러가 나. 원인과 수정 위치 알려줘.
[에러 메시지 붙여넣기]
```
