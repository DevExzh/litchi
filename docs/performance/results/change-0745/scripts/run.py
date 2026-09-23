#!/usr/bin/env python3
"""Change 0745 three-arm process matrix with heap-layout randomization.

Arms: A = base 009d515bef, B = f5f2922750 (deferred digests),
C = e94bd56f58 (B plus single-read editor streams and a shared commit editor).

Every process is pinned with `taskset -c CORE`. Each round runs every case
once per arm. Round r launches each binary through a symlink whose name adds
8*r bytes to argv[0]; Rust's runtime copies argv into heap allocations, and
change 0745 found that this startup heap layout alone moves some PPT
lifecycle timings by up to +/-15% (glibc page-fault behaviour). Using the same
argv[0] length for every arm within a round keeps the comparison paired while
the rounds sample many heap layouts. Arm order cycles through all six
permutations; case order rotates each round. Command lines are otherwise
identical between arms (equal-length arm directory names). Raw reports,
stderr, commands, exit codes and start times are retained.

Usage: run.py OUT_DIR ROUNDS [CASE_PREFIX ...]
"""

import itertools
import json
import os
import subprocess
import sys
import time

ROOT = "/home/zhuhe/code/litchi-worktrees/scratch/0745"
TREE = "/home/zhuhe/code/litchi-worktrees/0745-ppt-lazy-artifact-digests"
SLIDES = f"{TREE}/test-data/poi/test-data/slideshow"
DOCS = f"{TREE}/test-data/ole/doc"
CORE = "16"
ARMS = ("A", "B", "C")
ORDERS = list(itertools.permutations(ARMS))


def binary(arm, name, round_index):
    """Symlink to bin/ARM/NAME whose path length grows by 8 bytes per round."""
    link_dir = f"{ROOT}/argv0/{arm}"
    os.makedirs(link_dir, exist_ok=True)
    link = f"{link_dir}/{name}-p{'q' * (8 * round_index)}"
    if not os.path.islink(link):
        os.symlink(f"{ROOT}/bin/{arm}/{name}", link)
    return link


def probe(mode, fixture, warmups=5, samples=50):
    return lambda arm, r, out: [
        binary(arm, "probe", r),
        "--mode", mode,
        "--input", fixture,
        "--warmups", str(warmups),
        "--samples", str(samples),
    ]


def p0734(case, fixture):
    return lambda arm, r, out: [
        binary(arm, "p0734", r),
        "--case", case,
        "--input", fixture,
        "--operation", "format",
        "--samples", "50",
        "--warmups", "3",
    ]


def harness(arm, r, out):
    return [
        binary(arm, "litchi-perf-baseline", r),
        "--case", "ppt_semantic_one_edit_save,ppt_semantic_noop_edit_save",
        "--writer-shape", "tiny,large",
        "--samples", "40",
        "--warmup", "5",
        "--json", out,
    ]


CASES = {
    "remove-45543": probe("remove", f"{SLIDES}/45543.ppt"),
    "remove-41246": probe("remove", f"{SLIDES}/41246-1.ppt"),
    "remove-durable-45543": probe("remove-durable", f"{SLIDES}/45543.ppt"),
    "chain-durable-45543": probe("chain-durable", f"{SLIDES}/45543.ppt"),
    "apply-durable-45543": probe("apply-durable", f"{SLIDES}/45543.ppt"),
    "noop-45543": probe("noop", f"{SLIDES}/45543.ppt"),
    "hide-45543": probe("hide", f"{SLIDES}/45543.ppt"),
    "doc-replace-floating": probe("doc-replace", f"{DOCS}/FloatingPictures.doc"),
    "doc-replace-durable-floating": probe(
        "doc-replace-durable", f"{DOCS}/FloatingPictures.doc"
    ),
    "p0734-ppt45543": p0734("ppt45543", f"{SLIDES}/45543.ppt"),
    "p0734-docfloat": p0734("docfloat", f"{DOCS}/FloatingPictures.doc"),
    "harness-ppt-semantic": harness,
}


def main():
    out_dir = sys.argv[1]
    rounds = int(sys.argv[2])
    prefixes = sys.argv[3:]
    names = [name for name in CASES if not prefixes or name.startswith(tuple(prefixes))]
    os.makedirs(out_dir, exist_ok=True)
    log = []
    for round_index in range(rounds):
        order = ORDERS[round_index % len(ORDERS)]
        shift = round_index % len(names)
        rotated = names[shift:] + names[:shift]
        for name in rotated:
            for arm in order:
                stem = f"{out_dir}/{name}/r{round_index:02d}-{arm}"
                os.makedirs(os.path.dirname(stem), exist_ok=True)
                json_path = f"{stem}.json"
                command = ["taskset", "-c", CORE] + CASES[name](arm, round_index, json_path)
                started = time.time()
                completed = subprocess.run(command, capture_output=True, text=True)
                if name != "harness-ppt-semantic":
                    with open(json_path, "w") as handle:
                        handle.write(completed.stdout)
                with open(f"{stem}.stderr", "w") as handle:
                    handle.write(completed.stderr)
                entry = {
                    "case": name,
                    "round": round_index,
                    "arm": arm,
                    "argv0_extra_bytes": 8 * round_index,
                    "command": command,
                    "exit_code": completed.returncode,
                    "started_unix": round(started, 3),
                    "wall_s": round(time.time() - started, 3),
                }
                log.append(entry)
                if completed.returncode != 0:
                    print(json.dumps(entry), file=sys.stderr)
    with open(f"{out_dir}/commands.json", "w") as handle:
        json.dump(log, handle, indent=1)
    failures = [entry for entry in log if entry["exit_code"] != 0]
    print(f"processes={len(log)} failures={len(failures)}")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
