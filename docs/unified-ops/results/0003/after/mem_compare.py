#!/usr/bin/env python3
"""Compare the S10 allocation probe's before and after JSONL outputs.

usage: mem_compare.py BEFORE.jsonl AFTER.jsonl OUT.json [RELEASED.jsonl]

RELEASED is the after probe run with S10_RETAIN_SOURCE=0, that is
OpenOptions::retain_source(false).

Every probe metric is deterministic across measured rounds (the probe checks
it); the median round is reported anyway.
"""

import json
import statistics
import sys


def load(path):
    rows = {}
    for line in open(path):
        row = json.loads(line)
        rows[(row["format"], row["fixture"], row["route"])] = row
    return rows


def med(values):
    values = [v for v in values if v is not None]
    return statistics.median(values) if values else None


def main():
    before, after = load(sys.argv[1]), load(sys.argv[2])
    released = load(sys.argv[4]) if len(sys.argv) > 4 else {}
    out = []
    for key in sorted(set(before) | set(after)):
        b, a = before.get(key), after.get(key)
        entry = {"format": key[0], "fixture": key[1], "route": key[2]}
        for name, row in (("before", b), ("after", a), ("after_retain_source_false", released.get(key))):
            if row is None:
                if name != "after_retain_source_false":
                    entry[name] = None
                continue
            entry[name] = {
                "deterministic": row["deterministic"],
                "open_error": row["open_error"],
                "text_error": row["text_error"],
                "input_bytes": row["input_bytes"],
                "held_after_open": med(row["held_after_open"]),
                "peak_open": med(row["peak_open"]),
                "open_alloc_calls": med(row["open_alloc_calls"]),
                "text_bytes": row["text_bytes"],
                "peak_text": med(row["peak_text"]),
                "held_after_text": med(row["held_after_text"]),
                "live_after_drop": med(row["live_after_drop"]),
            }
        if b and a:
            for metric in ("held_after_open", "peak_open", "held_after_text", "peak_text", "open_alloc_calls"):
                x, y = entry["before"][metric], entry["after"][metric]
                if x is not None and y is not None:
                    entry[f"delta_{metric}"] = y - x
                    entry[f"ratio_{metric}"] = (y / x) if x else None
            entry["delta_held_after_open_over_input"] = (
                entry["delta_held_after_open"] / b["input_bytes"]
                if "delta_held_after_open" in entry else None)
        out.append(entry)
    json.dump({"schema": "litchi.unified-ops.0003.s10-allocation-probe.v1", "rows": out},
              open(sys.argv[3], "w"), indent=1)
    print(f"{'format':6} {'fixture':28} {'route':5} {'input':>9} {'held b':>10} {'held a':>10} "
          f"{'d held':>9} {'d/input':>7} {'peak b':>10} {'peak a':>10} {'calls b':>7} {'calls a':>7} {'text b/a':>12}")
    for e in out:
        b, a = e["before"], e["after"]
        if not b or not a:
            print(e["format"], e["fixture"], e["route"], "missing leg")
            continue
        name = e["fixture"].rsplit("/", 1)[-1]
        if b["open_error"] or a["open_error"]:
            print(f"{e['format']:6} {name:28} {e['route']:5} refused: before={bool(b['open_error'])} after={bool(a['open_error'])}")
            continue
        print(f"{e['format']:6} {name:28} {e['route']:5} {b['input_bytes']:9d} {b['held_after_open']:10.0f} "
              f"{a['held_after_open']:10.0f} {e['delta_held_after_open']:9.0f} "
              f"{e['delta_held_after_open_over_input']:7.2f} {b['peak_open']:10.0f} {a['peak_open']:10.0f} "
              f"{b['open_alloc_calls']:7.0f} {a['open_alloc_calls']:7.0f} {str(b['text_bytes'])+'/'+str(a['text_bytes']):>12}")


if __name__ == "__main__":
    main()
