"""
품목명 -> HS Code 파싱 정확도 대규모 검증.
보고서/바이어(토큰 많이 쓰는) 단계 이전, parser(Haiku)만 검증한다.

사용:
  python tests/verify_hs.py                 # 한글 품목명, 파서 정확도만
  python tests/verify_hs.py --field en      # 영문 품목명으로
  python tests/verify_hs.py --comtrade      # + 해당 HS로 Comtrade 데이터가 실제로 잡히는지
  python tests/verify_hs.py --limit 20      # 앞 N개만
"""
import argparse
import csv
import json
import os
import sys
import time
from collections import defaultdict

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tomllib
import anthropic
from agents.parser import parse_input
from agents.top_markets import fetch_top_markets

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)


def load_secrets():
    with open(os.path.join(ROOT, ".streamlit", "secrets.toml"), "rb") as f:
        return tomllib.load(f)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--field", choices=["ko", "en"], default="ko")
    ap.add_argument("--samples", default="hs_samples.json", help="샘플 JSON 파일명 (tests/ 기준)")
    ap.add_argument("--comtrade", action="store_true")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--sleep", type=float, default=0.0, help="샘플 간 대기(초, rate limit 완충)")
    args = ap.parse_args()

    secrets = load_secrets()
    client = anthropic.Anthropic(api_key=secrets["ANTHROPIC_API_KEY"])
    comtrade_key = secrets.get("COMTRADE_API_KEY")

    with open(os.path.join(HERE, args.samples), encoding="utf-8") as f:
        samples = json.load(f)
    if args.limit:
        samples = samples[: args.limit]

    rows = []
    n_hs4_ok = n_hs6_total = n_hs6_ok = n_err = n_ct_ok = n_ct_total = 0
    by_cat = defaultdict(lambda: [0, 0])  # cat -> [correct4, total]

    print(f"검증 시작: {len(samples)}개 샘플  |  입력필드={args.field}  |  comtrade={args.comtrade}\n")
    print(f"{'품목':16s} {'정답HS4':7s} {'파서HS':8s} {'HS4':4s} {'HS6':4s} {'conf':7s} {'CT':4s}")
    print("-" * 70)

    for s in samples:
        q = s[args.field]
        exp4, exp6 = s["hs4"], s.get("hs6")
        accept4 = [exp4] + s.get("alt4", [])  # 복수 정답 허용(정당하게 두 heading으로 분류되는 품목)
        got = got4 = conf = ""
        ok4 = ok6 = None
        ct = ""
        try:
            r = parse_input(q, client)
            got = r["hs_code"]
            got4 = got[:4]
            conf = r.get("confidence", "")
            ok4 = (got4 in accept4)
            if exp6:
                n_hs6_total += 1
                ok6 = (got == exp6)
                if ok6:
                    n_hs6_ok += 1
            if ok4:
                n_hs4_ok += 1
            by_cat[s["cat"]][1] += 1
            if ok4:
                by_cat[s["cat"]][0] += 1

            if args.comtrade and comtrade_key:
                try:
                    top = fetch_top_markets(got, comtrade_key)
                    n_ct_total += 1
                    if top:
                        n_ct_ok += 1
                        ct = "OK"
                    else:
                        ct = "empty"
                except Exception:
                    ct = "err"
                time.sleep(1.0)
        except Exception as e:
            n_err += 1
            got = f"ERR:{str(e)[:20]}"

        mark4 = "✓" if ok4 else ("✗" if ok4 is False else "-")
        mark6 = ("✓" if ok6 else "✗") if ok6 is not None else "-"
        print(f"{q[:16]:16s} {exp4:7s} {got:8s} {mark4:4s} {mark6:4s} {conf:7s} {ct:4s}")

        rows.append({
            "query": q, "cat": s["cat"], "expect_hs4": exp4, "expect_hs6": exp6 or "",
            "got_hs": got, "hs4_ok": ok4, "hs6_ok": ok6, "confidence": conf, "comtrade": ct,
        })
        if args.sleep:
            time.sleep(args.sleep)

    total = len(samples)
    print("\n" + "=" * 70)
    print("결과 요약")
    print("=" * 70)
    print(f"  HS4 heading 정확도 : {n_hs4_ok}/{total}  ({n_hs4_ok/total*100:.1f}%)")
    if n_hs6_total:
        print(f"  HS6 exact 정확도   : {n_hs6_ok}/{n_hs6_total}  ({n_hs6_ok/n_hs6_total*100:.1f}%)  (hs6 정답 보유 샘플만)")
    if n_err:
        print(f"  파싱 에러          : {n_err}건")
    if n_ct_total:
        print(f"  Comtrade 데이터 존재: {n_ct_ok}/{n_ct_total}  ({n_ct_ok/n_ct_total*100:.1f}%)")

    print("\n  카테고리별 HS4 정확도:")
    for cat, (c, t) in sorted(by_cat.items()):
        print(f"    {cat:12s} {c}/{t}  ({c/t*100:.0f}%)")

    print("\n  ✗ HS4 불일치 목록 (눈으로 확인):")
    for r in rows:
        if r["hs4_ok"] is False:
            print(f"    {r['query']:16s} 정답 {r['expect_hs4']}xx  ->  파서 {r['got_hs']}  (conf={r['confidence']})")

    tag = os.path.splitext(args.samples)[0]
    out = os.path.join(
        os.environ.get("TMPDIR", "/tmp"), f"hs_verify_{tag}_{args.field}.csv"
    )
    with open(out, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        w.writeheader()
        w.writerows(rows)
    print(f"\n  상세 결과 CSV: {out}")


if __name__ == "__main__":
    main()
