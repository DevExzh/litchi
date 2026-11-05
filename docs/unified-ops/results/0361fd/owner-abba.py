#!/usr/bin/env python3
"""Owner ABBA: 5 blocks per route, one CPU, host load recorded per leg (not aborted).

Usage: abba.py OUT_JSON [--cpu N] [--blocks 5]
A block is "clean" when no cargo/rustc/valgrind/clang/ld process ran before or
after any of its legs. Every block, clean or not, is retained.
"""
import argparse, json, math, os, re, statistics, subprocess, time
from pathlib import Path

HERE = Path(__file__).resolve().parent.parent
ROUTES = ("numbers_headers", "keynote_media", "pages_footnote_insert", "pages_section_name")
BLOCKERS = re.compile(r"^(cargo|rustc|rustdoc|clippy-driver|valgrind|callgrind.*|memcheck.*|dhat.*|cc|c\+\+|clang|gcc|ld|lld|mold|cc1|cc1plus)$")
TARGET_NS = 500_000_000


def host():
    names = []
    for entry in Path("/proc").iterdir():
        if entry.name.isdigit():
            try:
                names.append((entry / "comm").read_text().strip())
            except OSError:
                pass
    blockers = sorted(n for n in names if BLOCKERS.match(n))
    return dict(time_ns=time.time_ns(), load=open("/proc/loadavg").read().split()[:3], blockers=blockers)


def run(side, route, count, cpu):
    before = host()
    out = subprocess.run(["taskset", "-c", str(cpu), str(HERE / f"build-{side}" / "owner-probe"),
                          route, str(count), str(HERE / "fixtures")], capture_output=True, text=True, check=True)
    value = json.loads(out.stdout)
    assert value["route"] == route and value["iterations"] == count
    return dict(side=side, value=value, before=before, after=host())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("out", type=Path)
    parser.add_argument("--cpu", type=int, default=29)
    parser.add_argument("--blocks", type=int, default=5)
    args = parser.parse_args()
    result = dict(schema="owner-abba.v1", cpu=args.cpu, target_ns=TARGET_NS, routes=[])
    for route in ROUTES:
        count = 10
        while True:
            legs = [run(side, route, count, args.cpu) for side in ("base", "candidate")]
            fastest = min(leg["value"]["ns"] for leg in legs)
            if fastest >= TARGET_NS:
                break
            count = max(count + 1, math.ceil(count * TARGET_NS / fastest * 1.1))
        for _ in range(2):
            for side in ("base", "candidate"):
                run(side, route, count, args.cpu)
        case = dict(route=route, count=count, blocks=[])
        for block in range(args.blocks):
            legs = [run(side, route, count, args.cpu) for side in ("base", "candidate", "candidate", "base")]
            assert len({leg["value"]["output_digest"] for leg in legs}) == 1
            assert len({leg["value"]["checksum"] for leg in legs}) == 1
            ns = [leg["value"]["ns"] for leg in legs]
            ratio = (ns[1] + ns[2]) / (ns[0] + ns[3])
            clean = not any(leg[w]["blockers"] for leg in legs for w in ("before", "after"))
            case["blocks"].append(dict(ratio=ratio, clean=clean, legs=legs))
            args.out.write_text(json.dumps(result | dict(routes=result["routes"] + [case]), indent=1))
        ratios = [b["ratio"] for b in case["blocks"]]
        clean = [b["ratio"] for b in case["blocks"] if b["clean"]]
        case["median"] = statistics.median(ratios)
        case["spread"] = [min(ratios), max(ratios)]
        case["clean_blocks"] = len(clean)
        case["clean_median"] = statistics.median(clean) if clean else None
        result["routes"].append(case)
        args.out.write_text(json.dumps(result, indent=1))
        print(route, count, f"median={case['median']:.4f}", f"spread={case['spread'][0]:.4f}-{case['spread'][1]:.4f}",
              f"clean={len(clean)}/{len(ratios)}", f"clean_median={case['clean_median']}", flush=True)


if __name__ == "__main__":
    main()
