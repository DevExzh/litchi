#!/usr/bin/env python3
"""0768 ABBA control runner.

Runs the before (A) and after (B) binaries in alternating ABBA order, pinned to
one core, each process wrapped in `perf stat -e instructions,cycles`. Binaries
are staged at equal-length paths (bin/A/..., bin/B/...) and every argument is
identical across legs, so argv lengths match.

Usage: abba.py OUT_DIR ROUNDS CORE
"""

import json
import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))


def cases(out, leg, rnd):
    fixtures = os.path.join(HERE, "fixtures")
    return {
        "harness_doc_semantic_one_edit_save": (
            [
                f"bin/{leg}/harness",
                "--case",
                "doc_semantic_one_edit_save",
                "--writer-shape",
                "tiny,large",
                "--samples",
                "30",
                "--warmup",
                "5",
                "--json",
                os.path.join(out, f"harness-r{rnd}-{leg}.json"),
            ],
            None,
        ),
        "probe_docnohf_format": (
            [
                f"bin/{leg}/probe",
                "--case",
                "docnohf",
                "--input",
                os.path.join(fixtures, "NoHeadFoot.doc"),
                "--operation",
                "format",
                "--warmups",
                "5",
                "--samples",
                "40",
            ],
            os.path.join(out, f"probe-docnohf-r{rnd}-{leg}.json"),
        ),
        "probe_docfloat_format": (
            [
                f"bin/{leg}/probe",
                "--case",
                "docfloat",
                "--input",
                os.path.join(fixtures, "FloatingPictures.doc"),
                "--operation",
                "format",
                "--warmups",
                "5",
                "--samples",
                "30",
            ],
            os.path.join(out, f"probe-docfloat-r{rnd}-{leg}.json"),
        ),
    }


def main():
    out, rounds, core = sys.argv[1], int(sys.argv[2]), sys.argv[3]
    os.makedirs(out, exist_ok=True)
    schedule = []
    for rnd in range(rounds):
        order = ["A", "B"] if rnd % 2 == 0 else ["B", "A"]
        for case in ["harness_doc_semantic_one_edit_save", "probe_docnohf_format", "probe_docfloat_format"]:
            for leg in order:
                argv, stdout_path = cases(out, leg, rnd)[case]
                perf_path = os.path.join(out, f"perf-{case}-r{rnd}-{leg}.csv")
                command = [
                    "perf", "stat", "-e", "instructions,cycles", "-x,", "-o", perf_path,
                    "--", "taskset", "-c", core, *argv,
                ]
                started = time.time()
                with open(stdout_path, "w") if stdout_path else open(os.devnull, "w") as sink:
                    status = subprocess.run(command, cwd=HERE, stdout=sink, stderr=subprocess.PIPE, text=True)
                with open("/proc/loadavg") as handle:
                    load = handle.read().split()[0]
                schedule.append({
                    "round": rnd, "case": case, "leg": leg, "argv": argv,
                    "exit": status.returncode, "wall_s": round(time.time() - started, 3),
                    "loadavg_1m_after": float(load),
                    "stderr_tail": status.stderr[-400:],
                })
                if status.returncode != 0:
                    print(f"FAILED {case} r{rnd} {leg}: {status.stderr[-400:]}", file=sys.stderr)
                    sys.exit(1)
    with open(os.path.join(out, "schedule.json"), "w") as handle:
        json.dump(schedule, handle, indent=1)


if __name__ == "__main__":
    main()
