import json
import time

SYSTEM_PROMPT = """당신은 무역 인텔리전스 분석가입니다.
아래에 주어진 데이터만 사용해 한국어 종합 보고서를 마크다운으로 작성하세요.
데이터에 없는 수치(수입액·점유율·성장률)는 절대 지어내지 마세요.
섹션 번호와 제목을 반드시 포함하세요."""

USER_TEMPLATE = """## 분석 대상
품목: {product_name_ko} ({product_name_en})
HS Code: {hs_code_display}

## 상위 수입국 데이터 (UN Comtrade, 조사 대상 20개국 중 상위 {n}개국)
각 국가 객체 필드 설명:
- country_name: 국가명
- total_value_usd: 조사 기간 총 수입액(USD)
- by_year: 연도별 수입액(USD) — 해당 국가가 실제 보유한 연도만 존재
- share_pct: 조사 대상 20개국 합계 대비 점유율(%)
- yoy_change_pct: 최신 2개 연도 기준 전년 대비 증감률(%), null이면 데이터 부족

```json
{top_markets_json}
```

## 바이어 후보 조사 결과 (웹 검색 기반)
```json
{buyer_research_json}
```

아래 5개 섹션으로 보고서를 작성하세요:

### 1. HS Code 확인
품목명(한/영), 6자리 코드, 코드 의미 1줄 설명

### 2. 수입 규모 추이
각 상위국의 by_year를 근거로 연도별 흐름과 성장세를 서술. 데이터가 있는 연도만 언급.

### 3. 상위 수입국 Top {n}
표 형식으로 작성. 열: 순위 | 국가명 | 총 수입액(USD) | 점유율 | 전년 대비 증감
- 총 수입액은 total_value_usd, 점유율은 share_pct, 증감은 yoy_change_pct 값을 그대로 사용.
- yoy_change_pct가 null이면 "데이터 부족"으로 표기.
- 점유율 아래에 "※ 조사 대상 20개국 합계 기준" 각주를 표 밑에 한 줄 추가.

### 4. 바이어 후보 기업
국가별로 company_name | description | source(URL)를 표 또는 목록으로 정리. 바이어 데이터가 비어 있으면 "웹 검색에서 검증된 바이어를 찾지 못함"이라고 명시.

### 5. 시장 진입 분석
상위 2개국 시장 특징 + 한국 수출 시 고려할 점 2~3가지 (데이터에 근거)."""


def generate_report(
    parsed: dict,
    top_markets: list[dict],
    buyers: list[dict],
    anthropic_client,
) -> str:
    if not top_markets:
        return "수입 데이터를 조회할 수 없어 보고서를 생성할 수 없습니다."

    prompt = USER_TEMPLATE.format(
        product_name_ko=parsed["product_name_ko"],
        product_name_en=parsed["product_name_en"],
        hs_code_display=parsed["hs_code_display"],
        n=len(top_markets),
        top_markets_json=json.dumps(top_markets, ensure_ascii=False, indent=2),
        buyer_research_json=json.dumps(buyers, ensure_ascii=False, indent=2),
    )

    for attempt in range(3):
        try:
            response = anthropic_client.messages.create(
                model="claude-sonnet-5",
                max_tokens=4000,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": prompt}],
            )
            return response.content[0].text
        except Exception as e:
            if "rate_limit" in str(e) and attempt < 2:
                time.sleep(60)
            else:
                raise
