#!/usr/bin/env python3
"""Per-lifecycle inclusive Ir and call count for each (caller -> fingerprint)
call site, as an isolation pair over the --samples 1 and --samples 3 profiles."""
import subprocess, sys, pathlib, collections

S = pathlib.Path(__file__).resolve().parent

def sites(leg, case, samples):
    out = subprocess.run(
        [sys.executable, str(S / "callsites.py"),
         str(S / f"cg-{leg}-{case}-s{samples}.out"), "package_fingerprint"],
        capture_output=True, text=True).stdout
    table = {}
    for line in out.splitlines():
        parts = line.split("\t")
        if len(parts) < 4:
            continue
        incl = int(parts[0].replace(",", "").strip())
        calls = int(parts[1].replace(",", "").strip())
        caller = parts[3].split("\t")[0].strip()
        key = caller.split("::")[-1]
        entry = table.setdefault(key, [0, 0])
        entry[0] += incl
        entry[1] += calls
    return table

def main(case):
    for leg in ("before", "after"):
        t1, t3 = sites(leg, case, 1), sites(leg, case, 3)
        print(f"## {leg} {case}")
        for key in sorted(set(t1) | set(t3)):
            a, b = t1.get(key, [0, 0]), t3.get(key, [0, 0])
            incl = (b[0] - a[0]) // 2
            calls = (b[1] - a[1]) / 2
            if calls:
                print(f"   {key:28s} calls/life {calls:5.1f}  Ir/life {incl:15,d}  Ir/call {int(incl/calls):13,d}")

if __name__ == "__main__":
    for case in sys.argv[1:]:
        main(case)
