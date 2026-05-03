import json
import time

SYSTEM_PROMPT = """당신은 무역 인텔리전스 분석가입니다.
주어진 데이터를 바탕으로 한국어 종합 보고서를 마크다운 형식으로 작성하세요.
섹션 번호와 제목을 반드시 포함하세요."""

USER_TEMPLATE = """## 분석 대상
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
상위 2개국 시장 특징 + 한국 수출 시 고려할 점 2~3가지"""


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
        top_markets_json=json.dumps(top_markets, ensure_ascii=False, indent=2),
        buyer_research_json=json.dumps(buyers, ensure_ascii=False, indent=2),
    )

    for attempt in range(3):
        try:
            response = anthropic_client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=2000,
                system=SYSTEM_PROMPT,
                messages=[{"role": "user", "content": prompt}],
            )
            return response.content[0].text
        except Exception as e:
            if "rate_limit" in str(e) and attempt < 2:
                time.sleep(60)
            else:
                raise
