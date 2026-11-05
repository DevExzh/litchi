#!/usr/bin/env python3
"""S10 follow-up: does the RTF open gap depend on the allocator's state?

Runs the S10 `rtf` group (A: pre-M0a binary on S1's selectors, U and L: the
post-S9 binary on the unified and leaf selectors) as S1/S10 children, in
A U L L U A order, once with the default glibc malloc and once with
GLIBC_TUNABLES that stop heap trimming and pin the mmap threshold at its
maximum, so freed memory stays mapped and the next sample's allocations do
not fault fresh pages. Children run pinned, with S1's warm-up and sample
counts; each report lands in OUT/raw/<mode>-<pos><leg>-rtf.json.

usage: rtf_allocator_check.py --before-binary B --after-binary A --out DIR --cpu C [--blocks N]
"""

import argparse
import json
import os
import subprocess
import sys
import time
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import capture_s10  # noqa: E402

TUNABLES = "glibc.malloc.trim_threshold=4294967295:glibc.malloc.mmap_threshold=33554432"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--before-binary", required=True)
    parser.add_argument("--after-binary", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--cpu", required=True)
    parser.add_argument("--blocks", type=int, default=2)
    args = parser.parse_args()
    binaries = {"A": args.before_binary, "U": args.after_binary, "L": args.after_binary}
    os.makedirs(os.path.join(args.out, "raw"), exist_ok=True)
    log = open(os.path.join(args.out, "children.jsonl"), "a")
    for block in range(1, args.blocks + 1):
        for mode in ("default", "tunables"):
            for position, leg in enumerate("AULLUA", start=1):
                stem = f"B{block}-{mode}-{position}{leg}-rtf"
                out = os.path.join(args.out, "raw", stem + ".json")
                argv = ["taskset", "-c", args.cpu, binaries[leg],
                        "--case", ",".join(capture_s10.cases_for(leg, "rtf")),
                        "--warmup", "10", "--samples", "100", "--workers", "1", "--json", out]
                env = {**os.environ, "LC_ALL": "C", "TZ": "UTC"}
                if mode == "tunables":
                    env["GLIBC_TUNABLES"] = TUNABLES
                record = {"block": block, "mode": mode, "position": position, "leg": leg,
                          "argv": argv, "glibc_tunables": env.get("GLIBC_TUNABLES"),
                          "started_utc": datetime.now(timezone.utc).isoformat(),
                          "loadavg_before": open("/proc/loadavg").read().split()[:3]}
                started = time.monotonic()
                proc = subprocess.run(argv, cwd=capture_s10.REPO, env=env,
                                      stdout=subprocess.DEVNULL, stderr=subprocess.PIPE)
                record.update({"seconds": round(time.monotonic() - started, 3),
                               "exit_code": proc.returncode,
                               "stderr": proc.stderr.decode(errors="replace")[-500:]})
                log.write(json.dumps(record) + "\n")
                log.flush()
                print(stem, proc.returncode, flush=True)


if __name__ == "__main__":
    main()
