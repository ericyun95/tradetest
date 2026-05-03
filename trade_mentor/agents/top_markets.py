import datetime
import requests
from utils.hs_validator import get_all_reporter_codes, get_country_name

BASE_URL = "https://comtradeapi.un.org/data/v1/get/C/A/HS"


def _default_years() -> list[int]:
    current = datetime.date.today().year
    return [current - 3, current - 2, current - 1]


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
        "maxRecords": 500,
        "format": "JSON",
    }
    headers = {"Ocp-Apim-Subscription-Key": api_key}

    response = requests.get(BASE_URL, params=params, headers=headers, timeout=30)
    response.raise_for_status()

    records = response.json().get("data", [])
    if not records:
        return []

    # 국가별 연도별 수입액 집계
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
    top = results[:top_n]

    # rank + yoy_change_pct 계산 (실제 보유 연도 기준)
    for i, item in enumerate(top, start=1):
        item["rank"] = i
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
