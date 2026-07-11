import datetime
import requests
from utils.hs_validator import get_all_reporter_codes, get_country_name

BASE_URL = "https://comtradeapi.un.org/data/v1/get/C/A/HS"


def _default_years() -> list[int]:
    # UN Comtrade 연간(A) 데이터는 1~1.5년 지연됨.
    # 최신 연도가 일부 국가에 대해 비어 있어도 3개년 추이를 확보하도록
    # 4개년 창을 조회한 뒤, 국가별로 실제 보유한 연도만 집계한다.
    current = datetime.date.today().year
    return [current - 4, current - 3, current - 2, current - 1]


def fetch_top_markets(
    hs_code: str,
    api_key: str,
    years: list[int] = None,
    top_n: int = 5,
) -> list[dict]:
    if years is None:
        years = _default_years()

    params = {
        "reporterCode": get_all_reporter_codes(),
        "period": ",".join(str(y) for y in years),
        "cmdCode": hs_code,
        "flowCode": "M",
        # 집계 필터 — 이게 없으면 reporter×year당 100+개 행(운송수단·2차파트너
        # 분할)이 반환되어 maxRecords 상한에 걸려 과거 연도가 누락되고,
        # 중복 합산으로 금액이 뻥튀기된다.
        "partnerCode": 0,       # World 총계만
        "partner2Code": 0,      # 2차 파트너 분할 제거
        "motCode": 0,           # 운송수단 분할 제거
        "customsCode": "C00",   # 세관절차 집계
        "maxRecords": 500,
        "format": "JSON",
    }
    headers = {"Ocp-Apim-Subscription-Key": api_key}

    response = requests.get(BASE_URL, params=params, headers=headers, timeout=30)
    response.raise_for_status()

    records = response.json().get("data", [])
    if not records:
        return []

    # 국가별 연도별 수입액 집계 (필터 덕분에 reporter-year당 1행)
    aggregated: dict[int, dict] = {}
    for row in records:
        code = row.get("reporterCode")
        period = str(row.get("period", ""))
        value = row.get("primaryValue") or 0
        if code is None:
            continue
        if code not in aggregated:
            aggregated[code] = {"name": get_country_name(code), "by_year": {}}
        aggregated[code]["by_year"][period] = (
            aggregated[code]["by_year"].get(period, 0) + value
        )

    # total 계산 및 정렬
    results = []
    for numeric_code, data in aggregated.items():
        total = sum(data["by_year"].values())
        results.append(
            {
                "numeric_code": numeric_code,
                "country_name": data["name"],
                "total_value_usd": total,
                "by_year": data["by_year"],
            }
        )

    results.sort(key=lambda x: x["total_value_usd"], reverse=True)

    # 점유율은 조사 대상 20개국 합계 기준 (전 세계가 아님 — 프록시임을 명시)
    universe_total = sum(r["total_value_usd"] for r in results) or 1
    top = results[:top_n]

    for i, item in enumerate(top, start=1):
        item["rank"] = i
        item["share_pct"] = round(item["total_value_usd"] / universe_total * 100, 1)

        # 전년 대비 증감률 — 실제 보유한 최신 2개 연도 기준
        by_year = item["by_year"]
        available = sorted(by_year.keys())
        if len(available) >= 2:
            latest = by_year[available[-1]]
            prev = by_year[available[-2]]
            item["yoy_change_pct"] = (
                round((latest - prev) / prev * 100, 1) if prev else None
            )
        else:
            item["yoy_change_pct"] = None

    return top
