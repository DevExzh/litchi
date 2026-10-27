#!/usr/bin/env python3
"""Record 0377: the per-wave Markdown tables of the record from the
committed evidence (callgrind.tsv and abba.json next to this script)."""
import csv
import json
import statistics
from pathlib import Path

HERE = Path(__file__).resolve().parent
WAVES = [("W1 DOCX", ("docx",)), ("W2 PPTX", ("pptx",)), ("W3 PPT", ("ppt",)),
         ("W4 XLSX/XLSB", ("xlsx", "xlsb")), ("W5 DOC/XLS/RTF", ("doc", "xls", "rtf")),
         ("W6 ODF", ("odt", "oth", "odm", "ods", "odp", "odg", "odc", "odi", "odb", "odf")),
         ("W7 iWork", ("pages", "keynote", "numbers"))]


def budget(route):
    if "batch1000" in route:
        return 2.0, "2.0"
    if "batch" in route:
        return None, "scaling"
    if "open" in route:
        return 1.05, "1.05 or +20 µs"
    if route.endswith("_get"):
        return 1.10, "1.10 (read)"
    return 1.10, "1.10"


def main():
    ir = {r["route"]: r for r in csv.DictReader(open(HERE / "callgrind.tsv"), delimiter="\t")}
    wall = {}
    for run in json.loads((HERE / "abba.json").read_text())["runs"]:
        for c in run["cases"]:
            wall.setdefault(c["route"], {"base": [], "cand": []})[run["build"]].append(c)
    for wave, formats in WAVES:
        print(f"\n### {wave}\n")
        print("| Route | Unified Ir base → cand (ratio) | Leaf Ir ratio | Wall leaf / unified µs (cand) | Unified/leaf wall [spread] | Budget | Verdict |")
        print("|---|---|---:|---|---|---|---|")
        for route, r in ir.items():
            if route.split("_", 1)[0] not in formats or route == "docx_eager_open":
                continue
            ub, uc = float(r["unified_base_ir"]), float(r["unified_cand_ir"])
            lb, lc = float(r["leaf_base_ir"]), float(r["leaf_cand_ir"])
            w = wall.get(route)
            limit, text = budget(route)
            if w and w["cand"]:
                l = statistics.median(x for c in w["cand"] for x in c["leaf_block_median_ns"]) / 1e3
                u = statistics.median(x for c in w["cand"] for x in c["unified_block_median_ns"]) / 1e3
                rs = [a / b for c in w["cand"] for a, b in zip(c["unified_block_median_ns"], c["leaf_block_median_ns"])]
                ratio = statistics.median(rs)
                wtxt = f"{l:,.1f} / {u:,.1f}"
                rtxt = f"{ratio:.3f} [{min(rs):.3f}–{max(rs):.3f}]"
            else:
                ratio, wtxt, rtxt = uc / lc, "—", f"{uc / lc:.3f} (Ir)"
                u = l = None
            reg = max(uc / ub, lc / lb)
            ok_budget = limit is None or ratio <= limit or ("open" in route and u is not None and u - l <= 20)
            verdict = ("meets" if ok_budget else "miss") + ("" if reg <= 1.05 else ", >5 % Ir")
            print(f"| `{route}` | {ub:,.0f} → {uc:,.0f} ({uc / ub:.3f}) | {lc / lb:.3f} | {wtxt} | {rtxt} | {text} | {verdict} |")


if __name__ == "__main__":
    main()
