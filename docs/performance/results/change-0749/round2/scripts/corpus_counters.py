#!/usr/bin/env python3
"""Change 0749 round two: per-owner instructions over the OLE2 corpus.

Every OLE2 fixture under test-data (512 bytes to 48 MiB, the magic, sorted,
as the repository's `sector_layout_corpus` test enumerates them) is
republished by the probe's `cfb-write-reuse` with both of that test's edits
(`same`: flip the largest stream's last byte; `grow`: append
`3 * sector_size + 7` bytes to it), in arms A, B and C. For each fixture,
edit and arm, two pinned processes with 20 and 120 owners run under
`perf stat -e instructions:u`; the difference divided by 100 is the per-owner
user-instruction count, which does not depend on code or heap layout.
Instruction counts are the evidence here because the corpus holds hundreds
of sub-10-microsecond writes.

A `describe` of each fixture records its stream and mini-stream counts; a
fixture the probe cannot republish (no stream, an empty largest stream for
`same`, a source the writer does not adopt) is recorded and skipped.
Outputs already present are reused, so the lane resumes in chunks.

Usage: corpus_counters.py OUT_DIR [MAX_SECONDS] > corpus.json
"""

import json
import os
import statistics
import subprocess
import sys
import time

ROOT = "/home/zhuhe/code/litchi-worktrees/scratch/0749"
TREE = "/home/zhuhe/code/litchi-worktrees/0749-cfb-reuse-plan-validation"
MAGIC = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
MAX_BYTES = 48 * 1024 * 1024
SIZES = (20, 120)
ARMS = ("A", "B", "C")
EDITS = ("same", "grow")


def fixtures():
    found = []
    for directory, subdirectories, files in os.walk(f"{TREE}/test-data"):
        subdirectories.sort()
        for name in sorted(files):
            path = os.path.join(directory, name)
            try:
                size = os.path.getsize(path)
                if size < 512 or size > MAX_BYTES:
                    continue
                with open(path, "rb") as handle:
                    if handle.read(8) != MAGIC:
                        continue
            except OSError:
                continue
            found.append(path)
    return sorted(found)


def run_json(command, path):
    if os.path.exists(path):
        with open(path) as handle:
            text = handle.read()
        return json.loads(text) if text.strip().startswith("{") else None
    completed = subprocess.run(command, capture_output=True, text=True, timeout=300)
    text = completed.stdout if completed.returncode == 0 else ""
    with open(path, "w") as handle:
        handle.write(text or f"error: {completed.stderr[-500:]}")
    return json.loads(text) if text.strip().startswith("{") else None


def instructions(command, path):
    if not os.path.exists(path):
        completed = subprocess.run(
            ["perf", "stat", "-x,", "-e", "instructions:u", "-o", path, "--"] + command,
            capture_output=True, text=True, timeout=300,
        )
        if completed.returncode != 0:
            with open(path, "w") as handle:
                handle.write("error\n")
            return None
    with open(path) as handle:
        for line in handle:
            parts = line.strip().split(",")
            if len(parts) > 3 and parts[0] and not line.startswith("#") and parts[2] == "instructions:u":
                return float(parts[0])
    return None


def main():
    out_dir = sys.argv[1]
    budget = float(sys.argv[2]) if len(sys.argv) > 2 else 1e9
    started = time.time()
    os.makedirs(out_dir, exist_ok=True)
    rows = []
    complete = True
    for number, path in enumerate(fixtures()):
        if time.time() - started > budget:
            complete = False
            break
        label = os.path.relpath(path, TREE)
        stem = f"{out_dir}/{number:03d}"
        describe = run_json(
            ["taskset", "-c", "20", f"{ROOT}/bin/C/probe", "--mode", "describe", "--edit", "grow", "--input", path],
            f"{stem}-describe.json",
        )
        row = {"fixture": label, "streams": None, "mini_streams": None, "edits": {}}
        if describe:
            row["streams"] = describe["stream_count"]
            row["mini_streams"] = describe["source_mini_streams"]
        for edit in EDITS:
            result = {}
            for arm in ARMS:
                base = ["taskset", "-c", "20", f"{ROOT}/bin/{arm}/probe", "--mode", "cfb-write-reuse",
                        "--edit", edit, "--input", path, "--warmups", "5", "--oracle", "first"]
                small = instructions(base + ["--samples", str(SIZES[0])], f"{stem}-{edit}-{arm}-n{SIZES[0]}.perf")
                large = instructions(base + ["--samples", str(SIZES[1])], f"{stem}-{edit}-{arm}-n{SIZES[1]}.perf")
                report = run_json(base + ["--samples", "1"], f"{stem}-{edit}-{arm}-report.json")
                if small is None or large is None or report is None:
                    result[arm] = None
                    continue
                result[arm] = {
                    "instructions_u_per_owner": round((large - small) / (SIZES[1] - SIZES[0]), 1),
                    "reused": (report.get("sector_layout") or {}).get("reused"),
                    "output_sha256": report["output_sha256"],
                }
            row["edits"][edit] = result
        rows.append(row)
    pairs = []
    for row in rows:
        for edit, result in row["edits"].items():
            if not all(result.get(arm) for arm in ARMS):
                continue
            base = result["A"]["instructions_u_per_owner"]
            pairs.append({
                "fixture": row["fixture"],
                "edit": edit,
                "all_mini": row["streams"] is not None and row["streams"] == row["mini_streams"],
                "reused": result["C"]["reused"],
                "same_output": len({result[arm]["output_sha256"] for arm in ARMS}) == 1,
                "B_vs_A_pct": round(100.0 * (result["B"]["instructions_u_per_owner"] / base - 1.0), 3),
                "C_vs_A_pct": round(100.0 * (result["C"]["instructions_u_per_owner"] / base - 1.0), 3),
                "C_vs_B_pct": round(100.0 * (result["C"]["instructions_u_per_owner"] / result["B"]["instructions_u_per_owner"] - 1.0), 3),
            })

    def describe_group(selected):
        if not selected:
            return None
        return {
            "pairs": len(selected),
            "median_B_vs_A_pct": round(statistics.median(p["B_vs_A_pct"] for p in selected), 3),
            "max_B_vs_A_pct": max(p["B_vs_A_pct"] for p in selected),
            "B_above_plus_5_pct": sum(p["B_vs_A_pct"] > 5 for p in selected),
            "median_C_vs_A_pct": round(statistics.median(p["C_vs_A_pct"] for p in selected), 3),
            "max_C_vs_A_pct": max(p["C_vs_A_pct"] for p in selected),
            "C_above_plus_1_pct": sum(p["C_vs_A_pct"] > 1 for p in selected),
            "C_above_plus_5_pct": sum(p["C_vs_A_pct"] > 5 for p in selected),
        }

    reused = [p for p in pairs if p["reused"]]
    summary = {
        "complete": complete,
        "fixtures_enumerated": len(fixtures()),
        "fixtures_examined": len(rows),
        "pairs_measured": len(pairs),
        "all_outputs_identical_across_arms": all(p["same_output"] for p in pairs),
        "reused_pairs": describe_group(reused),
        "reused_all_mini_stream_pairs": describe_group([p for p in reused if p["all_mini"]]),
        "declined_pairs": describe_group([p for p in pairs if not p["reused"]]),
        "c_above_plus_1_pct": [p for p in pairs if p["C_vs_A_pct"] > 1],
    }
    json.dump({"schema": "0749-corpus-instructions-v1", "summary": summary, "pairs": pairs, "rows": rows},
              sys.stdout, indent=1)
    print()


if __name__ == "__main__":
    main()
