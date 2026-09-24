#!/usr/bin/env python3
"""ABBA-interleaved before/after process runner for change 0760.

Adapted from change 0743's driver. Every measured process is pinned to one
core with `taskset` and runs under `perf stat` for user-mode instructions and
cycles of the whole process. Each block runs before, after, after, before for
one case group, so each arm sees the same positions; blocks repeat for
`--rounds`. Raw harness JSON is kept per process together with the load
average observed just before it started.
"""

import argparse
import json
import os
import subprocess
import time

CASES = (
    "pptx_semantic_full_text,pptx_semantic_noop_edit_save,"
    "pptx_semantic_one_edit_save,pptx_semantic_one_percent_edit_save"
)

GROUPS = {
    "large": {"cases": CASES, "shape": "large", "samples": 15, "warmup": 3},
    "medium": {"cases": CASES, "shape": "medium", "samples": 60, "warmup": 10},
    "phases": {
        "cases": "pptx_semantic_opened_transaction_phases",
        "shape": "large,medium",
        "samples": 15,
        "warmup": 3,
    },
}


def load_average():
    with open("/proc/loadavg", encoding="ascii") as handle:
        return handle.read().split()[:3]


def run(binary, group, out_path, core):
    spec = GROUPS[group]
    command = [
        "perf", "stat", "-x,", "-e", "instructions:u,cycles:u", "-o", out_path + ".perf.csv",
        "taskset", "-c", core, binary,
        "--case", spec["cases"],
        "--semantic-shape", spec["shape"],
        "--samples", str(spec["samples"]),
        "--warmup", str(spec["warmup"]),
        "--json", out_path,
    ]
    started = time.time()
    load = load_average()
    completed = subprocess.run(command, capture_output=True, text=True, check=False)
    return {
        "command": command,
        "exit_code": completed.returncode,
        "stderr_tail": completed.stderr[-2000:],
        "load_average_before": load,
        "wall_seconds": round(time.time() - started, 3),
        "json": os.path.basename(out_path),
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--before", required=True)
    parser.add_argument("--after", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--rounds", type=int, default=4)
    parser.add_argument("--core", default="4")
    parser.add_argument("--groups", default="large,medium,phases")
    parser.add_argument("--round-offset", type=int, default=0)
    args = parser.parse_args()
    os.makedirs(args.out, exist_ok=True)
    binaries = {"before": args.before, "after": args.after}
    log = []
    for round_index in range(args.round_offset, args.round_offset + args.rounds):
        for group in args.groups.split(","):
            for position, arm in enumerate(["before", "after", "after", "before"]):
                name = f"r{round_index}-{group}-p{position}-{arm}.json"
                record = run(binaries[arm], group, os.path.join(args.out, name), args.core)
                record.update(
                    {"round": round_index, "group": group, "position": position, "arm": arm}
                )
                log.append(record)
                print(
                    f"round {round_index} {group} {arm} p{position} exit {record['exit_code']} "
                    f"{record['wall_seconds']}s load {' '.join(record['load_average_before'])}",
                    flush=True,
                )
                if record["exit_code"] != 0:
                    raise SystemExit(f"measured process failed: {record['stderr_tail']}")
    log_name = f"run-log-{args.round_offset}.json"
    with open(os.path.join(args.out, log_name), "w", encoding="utf-8") as handle:
        json.dump(log, handle, indent=1)


if __name__ == "__main__":
    main()
