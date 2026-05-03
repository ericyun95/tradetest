import json
import time


SYSTEM_PROMPT = """당신은 무역 HS Code 전문가입니다.
반드시 아래 JSON 형식으로만 응답하세요. 다른 텍스트는 출력하지 마세요."""

USER_TEMPLATE = """다음 입력을 분석하세요: "{user_input}"

다음 JSON을 반환하세요:
{{
  "input_type": "hs_code" | "product_name",
  "hs_code": "6자리 숫자 문자열 (점 없이)",
  "hs_code_display": "xxxx.xx 형식",
  "product_name_en": "영문 품목명 (간결하게)",
  "product_name_ko": "한글 품목명",
  "confidence": "high" | "medium" | "low",
  "note": null | "confidence가 low일 때 이유"
}}

규칙:
- HS Code는 반드시 6자리. 4자리 입력 시 뒤에 00 추가.
- confidence가 low인 경우 가장 유력한 코드를 넣고 이유를 note에 추가.
- 절대 JSON 외 텍스트 출력 금지."""


def parse_input(user_input: str, anthropic_client) -> dict:
    for attempt in range(3):
        try:
            response = anthropic_client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=300,
                system=SYSTEM_PROMPT,
                messages=[
                    {"role": "user", "content": USER_TEMPLATE.format(user_input=user_input)}
                ],
            )
            break
        except Exception as e:
            if "rate_limit" in str(e) and attempt < 2:
                time.sleep(30)
            else:
                raise
    text = response.content[0].text.strip()
    # 마크다운 코드 블록 제거
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        raise ValueError(f"LLM 응답 파싱 실패: {text}")
