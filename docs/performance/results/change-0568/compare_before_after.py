#!/usr/bin/env python3
"""Field-by-field BEFORE/AFTER comparison for change 0568 (+0570).

The headline control: file-source and owned-readat `open` and `list` must hold
at 53 logical reads. `one-cell` must fall from 319. Anything else is reported
as a divergence, not smoothed over.
"""
from __future__ import annotations
import argparse, json, pathlib, statistics


def pct(before, after):
    if before in (None, 0):
        return None
    return (after - before) / before * 100.0


def p50(xs):
    return statistics.median(xs) if xs else None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--before", required=True)
    ap.add_argument("--after", required=True)
    ap.add_argument("--output", required=True)
    a = ap.parse_args()
    b = json.loads(pathlib.Path(a.before).read_text())["operations"]
    c = json.loads(pathlib.Path(a.after).read_text())["operations"]

    rows, divergences, control = [], [], {}
    for key in sorted(set(b) | set(c)):
        bo, co = b.get(key), c.get(key)
        if bo is None or co is None:
            divergences.append({"key": key, "why": "present in only one capture"})
            continue
        bl, cl = bo["logical"], co["logical"]
        bi, ci = bo.get("isolation") or {}, co.get("isolation") or {}
        row = {
            "mode": bo["mode"], "operation": bo["operation"],
            "read_calls": {"before": bl["read_calls"], "after": cl["read_calls"],
                           "delta": cl["read_calls"] - bl["read_calls"],
                           "percent": pct(bl["read_calls"], cl["read_calls"])},
            "read_bytes": {"before": bl["read_bytes"], "after": cl["read_bytes"],
                           "delta": cl["read_bytes"] - bl["read_bytes"],
                           "percent": pct(bl["read_bytes"], cl["read_bytes"])},
            "version_calls": {"before": bl["version_calls"], "after": cl["version_calls"],
                              "delta": cl["version_calls"] - bl["version_calls"]},
            "len_calls": {"before": bl["len_calls"], "after": cl["len_calls"],
                          "delta": cl["len_calls"] - bl["len_calls"]},
            "seek_calls": {"before": bl["seek_calls"], "after": cl["seek_calls"],
                           "delta": cl["seek_calls"] - bl["seek_calls"]},
            "pread64_per_op": {"before": bi.get("pread64_per_op"), "after": ci.get("pread64_per_op"),
                               "delta": (None if bi.get("pread64_per_op") is None or ci.get("pread64_per_op") is None
                                         else ci["pread64_per_op"] - bi["pread64_per_op"])},
            "statx_per_op": {"before": bi.get("statx_per_op"), "after": ci.get("statx_per_op"),
                             "delta": (None if bi.get("statx_per_op") is None or ci.get("statx_per_op") is None
                                       else ci["statx_per_op"] - bi["statx_per_op"])},
            "isolation_raw": {
                "before": {"n1": bi.get("whole_child_samples_1"), "n11": bi.get("whole_child_samples_11")},
                "after": {"n1": ci.get("whole_child_samples_1"), "n11": ci.get("whole_child_samples_11")},
            },
            "attribution_p50_ns": {"before": p50(bl.get("elapsed_ns_samples") or []),
                                   "after": p50(cl.get("elapsed_ns_samples") or [])},
            "source_version_stable": {"before": bl.get("source_version_stable"),
                                      "after": cl.get("source_version_stable")},
            "warmups_samples": {"before": [bl.get("warmups"), bl.get("samples")],
                                "after": [cl.get("warmups"), cl.get("samples")]},
        }
        rows.append(row)

        op = bo["operation"]
        if op in ("open", "list"):
            held = (bl["read_calls"] == cl["read_calls"] == 53
                    and bl["read_bytes"] == cl["read_bytes"]
                    and bi.get("pread64_per_op") == ci.get("pread64_per_op")
                    and bi.get("statx_per_op") == ci.get("statx_per_op"))
            control[f"{bo['mode']}/{op}"] = {
                "held": held,
                "read_calls_before": bl["read_calls"], "read_calls_after": cl["read_calls"],
                "read_bytes_before": bl["read_bytes"], "read_bytes_after": cl["read_bytes"],
                "pread64_before": bi.get("pread64_per_op"), "pread64_after": ci.get("pread64_per_op"),
                "statx_before": bi.get("statx_per_op"), "statx_after": ci.get("statx_per_op"),
            }
            if not held:
                divergences.append({
                    "key": key, "why": "CONTROL MOVED: open/list must be unchanged at 53 logical reads",
                    "read_calls": [bl["read_calls"], cl["read_calls"]],
                    "read_bytes": [bl["read_bytes"], cl["read_bytes"]],
                    "pread64_per_op": [bi.get("pread64_per_op"), ci.get("pread64_per_op")],
                    "statx_per_op": [ci and bi.get("statx_per_op"), ci.get("statx_per_op")],
                })
        if op == "one-cell" and cl["read_calls"] >= bl["read_calls"]:
            divergences.append({"key": key, "why": "one-cell did not fall",
                                "read_calls": [bl["read_calls"], cl["read_calls"]]})

    control_held = all(v["held"] for v in control.values())
    out = {
        "schema_version": 1,
        "comparison_kind": "litchi-perf-change-0568-before-after-counters",
        "before_summary": a.before, "after_summary": a.after,
        "candidate_carries": ["change 0568 windowed worksheet scan (crates/litchi-xls/)",
                              "change 0570 contiguous CFB FAT run batching (crates/litchi-cfb/src/file.rs)"],
        "control": control,
        "control_verdict": "HELD - open and list unchanged at 53 logical reads"
                           if control_held else "MOVED - see divergences",
        "rows": rows,
        "divergences": divergences,
    }
    pathlib.Path(a.output).write_text(json.dumps(out, indent=2) + "\n")

    w = f"{'mode/op':<26}{'reads b':>9}{'reads a':>9}{'d':>7}{'bytes b':>10}{'bytes a':>10}{'d bytes':>9}{'pread b':>9}{'pread a':>9}{'statx b':>9}{'statx a':>9}"
    print(w)
    for r in rows:
        k = f"{r['mode']}/{r['operation']}"
        rc, rb = r["read_calls"], r["read_bytes"]
        pr, st = r["pread64_per_op"], r["statx_per_op"]
        print(f"{k:<26}{rc['before']:>9}{rc['after']:>9}{rc['delta']:>+7}"
              f"{rb['before']:>10}{rb['after']:>10}{rb['delta']:>+9}"
              f"{str(pr['before']):>9}{str(pr['after']):>9}{str(st['before']):>9}{str(st['after']):>9}")
    print()
    print("control:", out["control_verdict"])
    for k, v in control.items():
        print(f"  {k:<26} held={v['held']}  reads {v['read_calls_before']} -> {v['read_calls_after']}"
              f"  bytes {v['read_bytes_before']} -> {v['read_bytes_after']}")
    if divergences:
        print("\nDIVERGENCES:")
        for d in divergences:
            print(" ", json.dumps(d))
    else:
        print("\nno divergences")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
