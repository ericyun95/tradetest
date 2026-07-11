# HANDOFF.md — Global Trade Mentor AI 인수인계

> **다음 세션 Claude에게.** 이 파일 하나로 프로젝트의 현재 상태·발견된 버그·수정 내역·검증 결과·남은 일을 전부 파악할 수 있게 정리했다. `CLAUDE.md`(진입점)와 함께 이 파일을 먼저 읽어라.
> 작성 시점: **2026-07-11**. 작성 주체: 이전 세션 Claude (Opus 4.8).

---

## 0. 30초 요약

- **무엇**: 품목명(또는 HS Code)을 넣으면 UN Comtrade 수입 데이터 + 웹검색 바이어 조사로 **한국어 종합보고서(화면+Word)** 를 만드는 Streamlit 앱. (`CLAUDE.md`의 "이메일 생성"은 폐기된 v1 설명 — 실제는 v2 방향, [메모리 `project_direction_v2`] 참조)
- **이번 세션에 한 일**: "HS는 얼추 맞는데 그 뒤가 엉망"이라는 사용자 관찰을 출발점으로 → 원인 3종 규명 → 전부 수정 → **데이터 계층은 실측으로 검증 완료** → **파서(HS 파싱) 정확도를 대규모로 측정하는 하네스 구축 + 기준선 확보**.
- **현재 상태**: 데이터 계층(Comtrade) 수정·검증 끝. LLM 계층(바이어/보고서)은 모델 승격 등 수정은 했으나 **엔드투엔드 실행 검증은 미완**. 파서는 **정확도가 낮다는 것이 측정으로 드러남**(개선은 미착수 — 다음 단계).
- **작업 트리 커밋 안 됨**: 이번 변경사항은 모두 uncommitted. 마지막 커밋은 `bf6e6ef` (세션 시작 시점).

---

## 1. 아키텍처 / 파이프라인 (현재 코드 기준)

앱 진입점 `trade_mentor/app.py`가 4단계를 순차 실행:

```
사용자 입력(품목명/HS Code)
  │
  ▼ [1] agents/parser.py        HS Code 6자리 정규화 (LLM→JSON)         model: claude-haiku-4-5
  ▼ [2] agents/top_markets.py   UN Comtrade v2 수입데이터 → Top5 집계    (LLM 아님, requests)
  │      (time.sleep(5) — rate limit 완충)
  ▼ [3] agents/buyer_research.py 국가별 바이어 웹검색                     model: claude-sonnet-5 + web_search
  │      (time.sleep(5))
  ▼ [4] agents/report_gen.py    5섹션 한국어 보고서 생성                  model: claude-sonnet-5
  ▼ utils/docx_export.py         마크다운→Word 변환 (한글폰트 처리)
```

- **시크릿**: `trade_mentor/.streamlit/secrets.toml` → `ANTHROPIC_API_KEY`, `COMTRADE_API_KEY`.
- **국가 범위 하드코딩**: `utils/hs_validator.py`에 **조사 대상 20개국**의 reporterCode·국가명이 하드코딩돼 있음. 전 세계가 아니라 이 20개국이 시장의 프록시. (한계로 인지할 것 — 실제 최대 수입국이 이 밖이면 누락)

### ⚠️ 모델/스택 관련 주의
- `CLAUDE.md`의 "기술 스택(확정, 변경불가): LLM = claude-sonnet-4-20250514"는 **실제 코드와 불일치**. 실코드는 파서=Haiku 4.5, 바이어/보고서=Sonnet 5. (이번 세션에 품질 위해 승격, 사용자 승인함.) `claude-sonnet-4-20250514`는 2026-06-15 퇴역 예정이므로 되돌리지 말 것.
- `requirements.txt`의 `anthropic==0.52.0`은 구버전이지만, 모델 ID·툴은 문자열/딕셔너리로 API에 그대로 전달되므로 최신 모델(`claude-sonnet-5`)·최신 웹검색툴(`web_search_20260209`)이 정상 동작한다. SDK를 굳이 올릴 필요 없음.

---

## 2. 발견된 버그 3종과 근거 (이번 세션의 진단)

사용자 증상: "HS는 얼추 맞는데 그 이후(과거 데이터 수집·조사결과)가 엉망."

### 🔴 1차(치명적) — `top_markets.py` Comtrade 쿼리에 집계 필터 누락
- UN Comtrade는 `reporter×year×World`를 요청해도 내부적으로 **운송수단(motCode)·2차파트너(partner2Code)** 로 쪼개 반환. 실측: 영국/2023/녹차가 **필터 없으면 129행 vs `motCode=0,partner2Code=0`면 1행($15.67M)**.
- **결과 A(과거 데이터 누락)**: `maxRecords=500` 상한을 reporter 4~5개국×1개연도(2023)만으로 소진 → 2024·2025 행이 응답에 도착조차 못 함 → `by_year` 연도 1개 → YoY=None → "추이" 텅 빔. (수정 전 실측: 녹차 조회 시 **캐나다 1개국·2023만** 반환)
- **결과 B(수치 뻥튀기)**: 129행을 전부 합산 → 실제값의 8~10배.

### 🟠 2차 — 연도창이 미래를 봄
- 기존 `_default_years()` = `[현재-3,-2,-1]`. Comtrade 연간(A)은 1~1.5년 지연이라 최신 연도가 비기 쉬움.

### 🟡 3차 — LLM 계층 품질
- 3개 에이전트 전부 Haiku였음(관련성 판별 약함 → volza·블로그·무관품목 긁힘).
- `buyer_research.py max_tokens=1500`(웹검색 인용 붙으면 JSON 잘림→파싱실패→조용히 `[]`).
- `report_gen.py max_tokens=2000`(한국어 5섹션+표 잘림). 프롬프트가 "20개국 기준·점유율"을 요구하는데 실제로는 top5만 총합 없이 넘겨 **점유율을 날조**.

### 근본 원인(메타): 검증이 "모양"만 봤다
- `trade_mentor/verify_run.py`의 TC들이 구조(shape)만 확인(`len>=1`, 키 존재, 정렬)하고 **값의 정합성을 하나도 검사 안 함**. 그래서 캐나다 1개국·뻥튀기 값도 전부 PASS. → 이번에 "골든넘버(공개값 대조)" 방식으로 교체하는 게 교훈.

> 참고: 루트의 `report_녹차_20260419.docx`, `report_라면_20260419.docx`는 **이전(구버전) 코드**가 만든 산출물. 현재 5섹션 보고서 구조와 다르다. 진단 근거로만 참고.

---

## 3. 수정 내역 (전부 적용 완료)

### `agents/top_markets.py`
- 쿼리에 **`partnerCode=0, partner2Code=0, motCode=0, customsCode="C00"`** 추가 → reporter-year당 1행.
- `_default_years()` = **`[현재-4,-3,-2,-1]`** (4개년 버퍼).
- Top5 각 항목에 **`share_pct`**(조사 대상 20개국 합계 대비 점유율) 추가. YoY는 보유한 최신 2개 연도 기준.
- **✅ 검증(골든넘버)**: 녹차(090210) 조회 시 미국 $491M 1위·4개년 전부 반환, **영국 2023 = $15,666,155 → comtradeplus.un.org 공개값 $15.67M과 일치**.

### `agents/buyer_research.py`
- 모델 `claude-haiku-4-5` → **`claude-sonnet-5`**, `max_tokens` 1500 → **4000**.
- 웹검색툴 `web_search_20250305` → **`web_search_20260209`**(동적필터링) + **`blocked_domains`**(volza/panjiva/importgenius/tradeimex/imarcgroup/ensun 등 통계·리드 사이트 차단).
- **`pause_turn` 처리 루프** 추가(서버툴 다중 라운드 대응). 예외 시 여전히 `[]` 반환(안전).
- ⚠️ **엔드투엔드 실행 미검증** (Sonnet+웹검색은 토큰 비용 발생). import·문법만 확인함.

### `agents/report_gen.py`
- 모델 → **`claude-sonnet-5`**, `max_tokens` 2000 → **4000**.
- 프롬프트를 **실제 넘기는 데이터(share_pct, by_year)와 정합**하게 재작성. "데이터에 없는 수치는 지어내지 말라" 명시. 점유율은 `share_pct` 그대로 사용, YoY null이면 "데이터 부족" 표기, "※ 조사 대상 20개국 합계 기준" 각주.
- ⚠️ **엔드투엔드 실행 미검증** (동일 이유).

---

## 4. 검증 하네스 & 기준선 (이번 세션의 핵심 산출물)

파서(HS 파싱)가 downstream 전체의 입력인데, 이게 얼마나 정확한지 **대규모로 측정하는 도구**를 만들었다.

### 파일
- `tests/hs_samples.json` — **227개** 교과서형 품목(깨끗한 단어) + 정답 HS4(+일부 HS6).
- `tests/realistic_samples.json` — **57개** 현실 입력. 전라도 특산품(전남31·전북12) + 산업재(14). 지역명·브랜드·가공형태·규격 포함. 애매품목은 `alt4`(복수정답) 허용.
- `tests/verify_hs.py` — 러너.

### 실행법
```bash
cd trade_mentor
python tests/verify_hs.py                                   # 교과서 227개, 한글
python tests/verify_hs.py --samples realistic_samples.json  # 현실 57개
python tests/verify_hs.py --field en                        # 영문 입력으로
python tests/verify_hs.py --comtrade                        # 파서가 뱉은 HS로 Comtrade 데이터 잡히는지까지
python tests/verify_hs.py --limit 20                        # 앞 N개만
# 결과 상세는 $TMPDIR/hs_verify_<파일>_<필드>.csv 로 저장됨
```
- 채점 기준: **HS4 heading 일치**(주 지표. HS6은 5,612지선다라 전문가도 갈림), HS6 exact는 부가.
- `alt4`: 정당하게 두 heading으로 분류되는 품목(예 홍삼정 1302/2106)은 둘 다 정답 처리.
- 샘플 추가는 JSON에 한 줄씩 넣으면 됨(확장 용이).

### 📉 확보한 기준선 (개선 전 = baseline)
| 세트 | HS4 정확도 | 비고 |
|---|---|---|
| 교과서 227개 | **57.7%** (131/227) | HS6은 46.7% |
| 현실 57개 | **52.6%** (30/57) | 산업재 43%, 전북 50%, 전남 58% |

### 측정으로 드러난 문제(다음 개선의 타깃)
1. **가공형태 무시**: 굴비(염장 0305)→냉동 0303, 유자청(2008)→주스 2009, 건표고(0712)→생 0709. 지역·브랜드는 잘 버리는데 "간/건조/훈제" 같은 가공상태를 못 읽어 코드가 틀림.
2. **유사품목 혼동**: 다시마·미역→유채씨(1205), 무화과→바나나(0803), 석류→포도(0806), 참기름→대두유(1507), 굴착기·전동드릴→배터리, 프로젝터→카메라.
3. **산업재 지식 부족**: 부직포·부표·골판지박스·목재파레트 전멸.
4. **confidence 무용지물**: 두 세트 모두 오답의 ~39%가 `high` 라벨, `low` 라벨은 **0개**. → `app.py`는 confidence=='low'일 때만 경고를 띄우므로 **사용자는 오답에 대해 경고를 절대 못 받음**.
5. **비결정적**: 같은 입력도 run마다 코드가 달라짐(살균제 3808↔3809, 파레트 4415↔4419).

---

## 5. 남은 일 / 미완 영역

| # | 항목 | 상태 | 메모 |
|---|---|---|---|
| 1 | **파서 개선** | 미착수(권장 다음 단계) | Haiku→Sonnet 승격 + few-shot 예시 + 가공형태/유사품목 대응 + (선택)HS 유효성 검증. 개선 후 두 세트로 before/after 비교. 목표는 "HS6 정확"이 아니라 **HS4 정확도**(6자리는 4자리를 포함하므로 4자리가 천장). |
| 2 | **confidence 신뢰도** | 미해결 | 파서가 틀릴 때도 high. app의 경고 로직이 무력화됨. 개선 시 함께 손볼 것. |
| 3 | **바이어/보고서 엔드투엔드 검증** | 미완 | 코드 수정은 했으나 Sonnet+웹검색 실행 검증 안 함(토큰 비용). 실제 1건 돌려 산출물 눈으로 확인 필요. |
| 4 | **`verify_run.py` 갱신** | 미완 | 기존 shape-only TC를 골든넘버(공개값 대조) 방식으로 교체하면 값 버그 재발을 막음. |
| 5 | **20개국 하드코딩** | 설계 한계 | `hs_validator.py`의 20개국이 시장 프록시. 필요시 reporterCode를 "all"로 넓히는 것 검토(단 응답량·rate limit 주의). |

---

## 6. 파일 맵

```
trade_mentor/
├── app.py                      # Streamlit 진입점, 4단계 오케스트레이션
├── agents/
│   ├── parser.py               # HS 파싱 (Haiku). ★정확도 낮음 — 개선 타깃
│   ├── top_markets.py          # Comtrade 조회 ★수정·검증 완료
│   ├── buyer_research.py       # 바이어 웹검색 (Sonnet5) ★수정, 실행검증 미완
│   └── report_gen.py           # 보고서 생성 (Sonnet5) ★수정, 실행검증 미완
├── utils/
│   ├── hs_validator.py         # 20개국 코드 하드코딩
│   └── docx_export.py          # 마크다운→Word
├── tests/                      # ★이번 세션 신규
│   ├── hs_samples.json         # 227 교과서 샘플
│   ├── realistic_samples.json  # 57 현실 샘플(전라도+산업재)
│   └── verify_hs.py            # 파서 정확도 러너
├── verify_run.py               # 구 검증 스크립트(shape-only, 갱신 필요)
├── requirements.txt            # anthropic==0.52.0 (모델은 문자열이라 OK)
└── .streamlit/secrets.toml     # API 키 2종

루트/
├── CLAUDE.md                   # 진입점(단 v1 설명·모델 스택은 낡음 — §1 주의 참조)
├── HANDOFF.md                  # ← 지금 이 파일
├── spec.md / DESIGN.md / IMPLEMENT.md / VERIFY.md  # 3-에이전트 하네스 문서(초기 설계)
└── report_*.docx               # 구버전 코드 산출물(참고용)
```

## 7. 관련 메모리(자동 로드됨)
- `comtrade_pipeline_bugs` — 버그 3종 확정 원인·수정·골든넘버.
- `parser_accuracy_baseline` — 파서 정확도 기준선(교과서 57.7% / 현실 52.6%)·실패패턴·하네스 사용법.
- `project_direction_v2` — v2 방향 전환(이메일 폐기, Top5+바이어+보고서).

---

## 8. 다음 세션 첫 수 추천
1. 이 파일 + `CLAUDE.md` 읽기.
2. `python tests/verify_hs.py --samples realistic_samples.json` 한 번 돌려 현재 기준선(52.6% 부근) 재확인.
3. 파서 개선 착수(§5-1) → 두 세트로 before/after. 이게 사용자가 다음에 하려던 작업.
