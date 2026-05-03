import json
import re
import time

SYSTEM_PROMPT = """You are a trade research assistant.
Use web search to find B2B importers, distributors, and buyers.
Return JSON only. No other text."""

USER_TEMPLATE = """Product: {product_name_en} (HS Code: {hs_code_display})
Top importing countries: {country_list}

For each country, search for 2-3 major importing companies or distributors of this product.
Return:
{{
  "buyers": [
    {{
      "country": "country name",
      "company_name": "...",
      "description": "industry/specialty, one line",
      "source": "URL or publication"
    }}
  ]
}}

Rules:
- Only include companies verifiable via web search
- Do not invent or guess company names
- "source" must be the company's own official website URL (e.g. company.com), NOT trade data sites like volza.com, panjiva.com, or importgenius.com
- If you cannot find the official website, omit the source field
- If no verifiable company found for a country, skip that country"""


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

    try:
        for attempt in range(3):
            try:
                response = anthropic_client.messages.create(
                    model="claude-haiku-4-5-20251001",
                    max_tokens=1500,
                    system=SYSTEM_PROMPT,
                    tools=[{"type": "web_search_20250305", "name": "web_search"}],
                    messages=[{"role": "user", "content": prompt}],
                )
                break
            except Exception as e:
                if "rate_limit" in str(e) and attempt < 2:
                    time.sleep(30)
                else:
                    raise
    except Exception:
        return []

    buyers = []
    for block in response.content:
        if block.type == "text":
            text = block.text.strip()
            # 마크다운 코드 블록 제거
            if "```" in text:
                match = re.search(r"```(?:json)?\s*([\s\S]+?)```", text)
                if match:
                    text = match.group(1).strip()
            # JSON 객체만 추출
            match = re.search(r'\{[\s\S]*"buyers"[\s\S]*\}', text)
            if match:
                try:
                    result = json.loads(match.group())
                    buyers = result.get("buyers", [])
                except json.JSONDecodeError:
                    pass

    return buyers
