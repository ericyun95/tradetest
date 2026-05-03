import re

COUNTRY_CODE_MAP = {
    "USA": 842, "VNM": 704, "DEU": 276, "JPN": 392,
    "CHN": 156, "IDN": 360, "THA": 764, "IND": 356,
    "AUS": 36,  "GBR": 826, "FRA": 250, "CAN": 124,
    "BRA": 76,  "MYS": 458, "PHL": 608, "SGP": 702,
    "SAU": 682, "ARE": 784, "MEX": 484, "NLD": 528,
}

NUMERIC_TO_NAME = {
    842: "United States",  704: "Viet Nam",      276: "Germany",
    392: "Japan",          156: "China",          360: "Indonesia",
    764: "Thailand",       356: "India",          36:  "Australia",
    826: "United Kingdom", 250: "France",         124: "Canada",
    76:  "Brazil",         458: "Malaysia",       608: "Philippines",
    702: "Singapore",      682: "Saudi Arabia",   784: "UAE",
    484: "Mexico",         528: "Netherlands",
}


def validate_hs_code(code: str) -> tuple[bool, str]:
    digits = re.sub(r"[^\d]", "", code)
    if len(digits) == 4:
        digits = digits + "00"
    if len(digits) != 6:
        return False, ""
    return True, digits


def get_all_reporter_codes() -> str:
    return ",".join(str(v) for v in COUNTRY_CODE_MAP.values())


def get_country_name(numeric_code: int) -> str:
    return NUMERIC_TO_NAME.get(numeric_code, str(numeric_code))
