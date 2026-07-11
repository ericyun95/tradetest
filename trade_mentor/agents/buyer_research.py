import json
import re
import time

# 무역 통계 집계/광고 디렉터리 사이트 — 실제 바이어가 아니라 데이터 판매·리드 사이트.
# 이전 버전에서 이런 도메인이 "바이어"로 잡혀 결과 품질을 망쳤다.
BLOCKED_DOMAINS = [
    "volza.com", "panjiva.com", "importgenius.com", "tradeimex.in",
    "imarcgroup.com", "ensun.io", "exportgenius.net", "seair.co.in",
    "trademo.com", "zauba.com", "connect2india.com", "eworldtrade.com",
    "tradewheel.com", "go4worldbusiness.com", "santoshexport.com",
]

SYSTEM_PROMPT = """You are a B2B trade research analyst.
Use web search to find real importing companies, distributors, and wholesalers for a product in specific countries.
Verify each company on its own official website. Return JSON only — no other text."""

USER_TEMPLATE = """Product: {product_name_en} (HS Code: {hs_code_display})
Top importing countries: {country_list}

For each country, find 2-3 REAL companies that import, distribute, or wholesale this specific product.
Return exactly this JSON:
{{
  "buyers": [
    {{
      "country": "country name",
      "company_name": "official company name",
      "description": "what they do, one line — must relate to this product category",
      "source": "the company's own official website URL"
    }}
  ]
}}

Hard rules:
- The company must actually deal in THIS product ({product_name_en}) — not an unrelated item.
- "source" MUST be the company's own official website (e.g. company.com). NEVER a trade-statistics or lead-generation site.
- Do NOT use blogs, marketplaces listings, news articles, or directory pages as a company.
- Do not invent or guess company names or URLs. If you cannot verify a company for a country, skip that country.
- Prefer distributors/wholesalers/importers over manufacturers."""


def _web_search_tool():
    return {
        "type": "web_search_20260209",
        "name": "web_search",
        "max_uses": 8,
        "blocked_domains": BLOCKED_DOMAINS,
    }


def _extract_buyers(content) -> list[dict]:
    """응답 content 블록들에서 buyers JSON을 추출."""
    for block in content:
        if getattr(block, "type", None) != "text":
            continue
        text = block.text.strip()
        if "```" in text:
            m = re.search(r"```(?:json)?\s*([\s\S]+?)```", text)
            if m:
                text = m.group(1).strip()
        m = re.search(r'\{[\s\S]*"buyers"[\s\S]*\}', text)
        if m:
            try:
                return json.loads(m.group()).get("buyers", [])
            except json.JSONDecodeError:
                continue
    return []


def research_buyers(
    product_name_en: str,
    hs_code_display: str,
    top_markets: list[dict],
    anthropic_client,
) -> list[dict]:
    country_list = ", ".join(m["country_name"] for m in top_markets)
    prompt = USER_TEMPLATE.format(
        product_name_en=product_name_en,
        hs_code_display=hs_code_display,
        country_list=country_list,
    )

    messages = [{"role": "user", "content": prompt}]

    try:
        # 웹 검색은 서버 툴이라 여러 라운드가 필요할 수 있음(pause_turn).
        for _ in range(5):
            for attempt in range(3):
                try:
                    response = anthropic_client.messages.create(
                        model="claude-sonnet-5",
                        max_tokens=4000,
                        system=SYSTEM_PROMPT,
                        tools=[_web_search_tool()],
                        messages=messages,
                    )
                    break
                except Exception as e:
                    if "rate_limit" in str(e) and attempt < 2:
                        time.sleep(30)
                    else:
                        raise
            if response.stop_reason == "pause_turn":
                # 서버 툴 루프가 중단됨 — 대화를 이어서 재개
                messages.append({"role": "assistant", "content": response.content})
                continue
            break
    except Exception:
        return []

    return _extract_buyers(response.content)
