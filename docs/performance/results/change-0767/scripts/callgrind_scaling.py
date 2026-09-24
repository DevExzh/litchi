#!/usr/bin/env python3
"""Change 0767: Callgrind instruction counts of one timed owner, by scale.

For each arm, mode and generated input (v3/v4, mini/regular, 1,000 / 3,000 /
10,000 streams), one pinned process runs one owner (no warmup) under
Callgrind with collection toggled on the probe's `timed_owner`, so only the
timed region is counted. The inclusive instruction counts of the owner and
of the functions this change concerns are extracted with
`callgrind_annotate --inclusive=yes`; the heads of those listings are kept.
Callgrind counts its own byte loops for `memset` and `memcpy`, so its totals
exceed native counts; the ratios between scales are what this lane reports.

Usage: callgrind_scaling.py OUT_DIR > callgrind-scaling.json
"""

import json
import os
import re
import subprocess
import sys

ROOT = "/home/zhuhe/code/litchi-worktrees/scratch/0767"
CORE = "28"
FUNCTIONS = {
    "validate_stream_allocations": r"OleFile<R>::validate_stream_allocations",
    "collect_exact": r"SectorChainScratch::collect_exact",
    "load_directory": r"OleFile<R>::load_directory",
    "load_minifat": r"OleFile<R>::load_minifat",
    "open_stream": r"OleFile<R>::open_stream",
    "find_entry": r"OleFile<R>::find_entry",
    "collect_sector_chain": r"litchi_cfb::file::collect_sector_chain\b",
    "end_chain_collect": r"EndChainScratch::collect",
    "memset": r"__memset",
}
MODES = ("cfb-open", "cfb-read-all")


def inclusive(listing, pattern):
    total = 0
    for line in listing.splitlines():
        match = re.match(r"\s*([\d,]+)\s+\(", line)
        if match and re.search(pattern, line):
            total = max(total, int(match.group(1).replace(",", "")))
    return total


def main():
    out_dir = sys.argv[1]
    os.makedirs(out_dir, exist_ok=True)
    rows = []
    for arm in ("base", "cand"):
        for mode in MODES:
            for version in ("v3", "v4"):
                for kind in ("mini", "regular"):
                    for count in (1000, 3000, 10000):
                        name = f"{version}-{kind}-{count}"
                        stem = f"{out_dir}/{arm}-{mode}-{name}"
                        command = [
                            "taskset", "-c", CORE, "valgrind", "--tool=callgrind",
                            f"--callgrind-out-file={stem}.out", "--toggle-collect=*timed_owner*",
                            f"{ROOT}/bin/{arm}/probe", "--mode", mode,
                            "--input", f"{ROOT}/gen/{name}.cfb",
                            "--warmups", "0", "--samples", "1",
                        ]
                        completed = subprocess.run(command, capture_output=True, text=True)
                        if completed.returncode != 0:
                            raise SystemExit(f"{command} failed: {completed.stderr[-2000:]}")
                        listing = subprocess.run(
                            ["callgrind_annotate", "--inclusive=yes", "--threshold=100", f"{stem}.out"],
                            capture_output=True, text=True, check=True).stdout
                        total = inclusive(listing, r"PROGRAM TOTALS")
                        row = {"arm": arm, "mode": mode, "input": name, "streams": count,
                               "owner_ir": total, "owner_ir_per_stream": round(total / count, 1)}
                        for key, pattern in FUNCTIONS.items():
                            row[f"{key}_ir"] = inclusive(listing, pattern)
                        rows.append(row)
                        head = "\n".join(listing.splitlines()[:60])
                        with open(f"{stem}.inclusive.txt", "w") as handle:
                            handle.write(head + "\n")
                        os.remove(f"{stem}.out")
    ratios = []
    for row in rows:
        if row["streams"] == 1000:
            continue
        smaller = 1000 if row["streams"] == 3000 else 3000
        previous = next(r for r in rows if r["arm"] == row["arm"] and r["mode"] == row["mode"]
                        and r["input"] == row["input"].rsplit("-", 1)[0] + f"-{smaller}")
        ratios.append({
            "arm": row["arm"], "mode": row["mode"], "input": row["input"],
            "streams_ratio": row["streams"] / smaller,
            "owner_ir_ratio": round(row["owner_ir"] / previous["owner_ir"], 3),
            "validate_stream_allocations_ir_ratio": round(
                row["validate_stream_allocations_ir"] / previous["validate_stream_allocations_ir"], 3)
            if previous["validate_stream_allocations_ir"] else None,
        })
    json.dump({"rows": rows, "ratios": ratios}, sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
