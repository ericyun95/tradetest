# CLAUDE.md — Global Trade Mentor AI

## 이 파일을 읽는 Claude Code에게

이 파일은 프로젝트 진입점입니다. **모든 작업 전 반드시 먼저 읽으세요.**  
지시 없이 코드를 작성하거나 파일을 생성하지 마세요.

---

## 프로젝트 한 줄 요약

> 품목명(한글/영문) 또는 HS Code + 국가명을 입력하면  
> UN Comtrade 무역 데이터 요약과 맞춤 영문 이메일 초안을 생성하는 Streamlit 웹 챗봇

---

## 파일 맵 (읽는 순서)

```
CLAUDE.md       ← 지금 여기 (진입점)
spec.md         ← 전체 기술 스펙 (아키텍처, API, 프롬프트, 파일 구조)
DESIGN.md       ← [설계 에이전트] 하네스
IMPLEMENT.md    ← [구현 에이전트] 하네스
VERIFY.md       ← [검증 에이전트] 하네스
```

---

## 3-Agent 실행 순서

### Phase 1 — 설계 에이전트
```
읽을 파일: spec.md → DESIGN.md
할 일: 스펙 공백 확인, 프롬프트 완성도 검토
출력: DESIGN.md 체크리스트 업데이트
다음: "설계 완료"를 명시적으로 선언 후 Phase 2 진행
```

### Phase 2 — 구현 에이전트
```
읽을 파일: spec.md → IMPLEMENT.md
할 일: IMPLEMENT.md의 Step 순서대로 코드 작성
출력: 실제 .py 파일들
다음: IMPLEMENT.md 체크리스트 모두 체크 후 Phase 3 진행
```

### Phase 3 — 검증 에이전트
```
읽을 파일: VERIFY.md
할 일: 테스트 케이스 TC-01 ~ TC-08 순서대로 실행
출력: 각 TC 결과 (✅/❌) + 실패 시 수정 제안
다음: 전체 PASS 시 배포 안내
```

---

## 전역 규칙 (모든 에이전트 공통)

1. **spec.md가 정답이다** — 스펙과 충돌하는 판단은 스펙을 따른다
2. **스코프 이탈 금지** — spec.md의 "제외 항목"은 구현하지 않는다
3. **순서 준수** — Phase를 건너뛰거나 병렬 실행하지 않는다
4. **완료 선언** — 각 Phase 완료 시 "Phase N 완료"를 명시적으로 출력한다
5. **모르면 멈춘다** — 불확실한 결정은 구현 전 사용자에게 질문한다

---

## 기술 스택 (확정, 변경 불가)

| 항목 | 결정값 |
|---|---|
| UI | Streamlit |
| LLM | Claude API `claude-sonnet-4-20250514` |
| 무역 데이터 | UN Comtrade API v2 |
| 언어 | Python 3.11+ |
| 배포 | Streamlit Cloud |
| 별도 DB | 없음 |
| 별도 번역 API | 없음 |
