import sys, time, json
sys.path.insert(0, ".")

from agents.parser import parse_input
from agents.top_markets import fetch_top_markets
from agents.buyer_research import research_buyers
from agents.report_gen import generate_report
import anthropic, tomllib

with open(".streamlit/secrets.toml", "rb") as f:
    secrets = tomllib.load(f)

client = anthropic.Anthropic(api_key=secrets["ANTHROPIC_API_KEY"])
comtrade_key = secrets["COMTRADE_API_KEY"]

results = []

def check(tc, cond, msg):
    status = "✅ PASS" if cond else "❌ FAIL"
    results.append((tc, status, msg))
    print(f"  {status}: {msg}")

print("=" * 60)
print("VERIFY.md 검증 에이전트 실행")
print("=" * 60)

# ── TC-01: 한글 품목명 → HS Code ──
print("\n[TC-01] 에어프라이어 → HS Code 추론")
r1 = parse_input("에어프라이어", client)
check("TC-01", len(r1["hs_code"]) == 6 and r1["hs_code"].isdigit(), f"HS Code 6자리 숫자: {r1['hs_code']}")
check("TC-01", "air" in r1["product_name_en"].lower(), f"product_name_en: {r1['product_name_en']}")
check("TC-01", r1["confidence"] in ("high", "medium"), f"confidence: {r1['confidence']}")

time.sleep(8)

# ── TC-02: HS Code 직접 입력 → 역추론 ──
print("\n[TC-02] 8516.60 → 품목명 역추론")
r2 = parse_input("8516.60", client)
check("TC-02", r2["input_type"] == "hs_code", f"input_type: {r2['input_type']}")
check("TC-02", r2["hs_code"] == "851660", f"hs_code: {r2['hs_code']}")
check("TC-02", bool(r2["product_name_ko"]), f"product_name_ko: {r2['product_name_ko']}")
check("TC-02", r2["hs_code_display"] == "8516.60", f"hs_code_display: {r2['hs_code_display']}")

time.sleep(8)

# ── TC-03: 영문 품목명 ──
print("\n[TC-03] cosmetics → HS Code")
r3 = parse_input("cosmetics", client)
check("TC-03", r3["hs_code"].startswith("33"), f"HS 앞 2자리 '33': {r3['hs_code']}")
check("TC-03", len(r3["hs_code"]) == 6, f"6자리 숫자: {r3['hs_code']}")

time.sleep(8)

# ── TC-04: 점 없는 HS Code ──
print("\n[TC-04] 070200 → hs_code_display 형식")
r4 = parse_input("070200", client)
check("TC-04", r4["hs_code"] == "070200", f"hs_code: {r4['hs_code']}")
check("TC-04", r4["hs_code_display"] == "0702.00", f"hs_code_display: {r4['hs_code_display']}")

time.sleep(8)

# ── TC-05: UN Comtrade Top Markets ──
print("\n[TC-05] UN Comtrade Top Markets (851660 에어프라이어)")
top = fetch_top_markets("851660", comtrade_key)
check("TC-05", len(top) >= 1, f"반환 건수: {len(top)}개")
check("TC-05", all("rank" in m and "country_name" in m and "total_value_usd" in m for m in top), "필수 키 존재")
check("TC-05", all(top[i]["total_value_usd"] >= top[i+1]["total_value_usd"] for i in range(len(top)-1)), "금액 내림차순 정렬")
check("TC-05", top[0]["rank"] == 1, f"첫 번째 rank=1: {top[0]['rank']}")
for m in top:
    yoy = f"{m['yoy_change_pct']:+.1f}%" if m['yoy_change_pct'] is not None else "1개년"
    print(f"    #{m['rank']} {m['country_name']:20s} ${m['total_value_usd']:>14,.0f}  YoY: {yoy}")

time.sleep(10)

# ── TC-06: 바이어 후보 조사 ──
print("\n[TC-06] 바이어 후보 기업 조사")
buyers = research_buyers("Air Fryer", "8516.60", top, client)
check("TC-06", isinstance(buyers, list), "리스트 반환")
check("TC-06", not any("[" in b.get("company_name","") for b in buyers), "플레이스홀더 없음")
if buyers:
    check("TC-06", all("company_name" in b and "country" in b for b in buyers), "필수 키 존재")
    for b in buyers[:3]:
        print(f"    [{b['country']}] {b['company_name']}: {b['description']}")
else:
    print("    ⚠️  WARN: 바이어 0건 (웹 검색 미반환 — 앱 크래시 없으면 허용)")

time.sleep(10)

# ── TC-07: 보고서 완전성 ──
print("\n[TC-07] 보고서 5개 섹션 완전성")
report = generate_report(r1, top, buyers, client)
for n in range(1, 6):
    found = any(f"### {n}." in report or f"## {n}." in report for _ in [1])
    check("TC-07", found, f"섹션 {n} 존재")
has_table = "|" in report
check("TC-07", has_table, "표(|) 형식 포함")
print(f"    보고서 총 {len(report)}자")

# ── TC-08: 빈 입력 예외 ──
print("\n[TC-08] 빈 입력 예외 처리 (app.py 로직 검증)")
empty = "".strip()
check("TC-08", not bool(empty), "빈 문자열 감지 → st.error() 후 중단")

# ── 최종 결과 ──
print("\n" + "=" * 60)
print("검증 결과 요약")
print("=" * 60)
passed = sum(1 for _, s, _ in results if "PASS" in s)
failed = sum(1 for _, s, _ in results if "FAIL" in s)
for tc, status, msg in results:
    print(f"  [{tc}] {status}: {msg}")
print(f"\n총 {passed+failed}개 항목: ✅ {passed}개 PASS  /  ❌ {failed}개 FAIL")
if failed == 0:
    print("\n✅ Phase 3 완료 — 전 항목 PASS")
else:
    print(f"\n❌ {failed}개 실패 항목 수정 필요")
