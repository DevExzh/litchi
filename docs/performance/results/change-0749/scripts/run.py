#!/usr/bin/env python3
"""Change 0749 two-arm ABBA process matrix with heap-layout randomization.

Arms: base = ab29ac6291 (branch tip with record 0745), cand = the 0749
in-place Reuse-plan validation. Both arms' binaries are built with identical
cargo commands, flags and features (see binaries.json).

Every process is pinned with `taskset -c 20`. Round r runs every case once per
arm, in ABBA order across rounds (base,cand / cand,base / cand,base /
base,cand, repeating), with the case order rotated each round. As in change
0745, each binary is started through a symlink whose path grows by 8 bytes per
round: the probe's and the harness's own argument handling copy argv into the
heap, and 0745 found that this early heap layout alone moves some OLE2
lifecycle timings by up to +/-15%. Both arms use the same argv[0] length in a
round, so the comparison stays paired while the rounds sample many layouts.
Arm directory names have equal length. Raw reports, stderr, exact commands,
exit codes and start times are retained.

Usage: run.py OUT_DIR ROUNDS [CASE_PREFIX ...]

`ROUND_START` (default 0) continues a matrix: rounds ROUND_START ..
ROUND_START + ROUNDS - 1 keep the same order cycle and argv[0] progression,
and their command log is written to `commands-from-rROUND_START.json`.
"""

import json
import os
import subprocess
import sys
import time

ROOT = "/home/zhuhe/code/litchi-worktrees/scratch/0749"
TREE = "/home/zhuhe/code/litchi-worktrees/0749-cfb-reuse-plan-validation"
SLIDES = f"{TREE}/test-data/poi/test-data/slideshow"
DOCS = f"{TREE}/test-data/ole/doc"
CORE = "20"
ARMS = ("base", "cand")
ORDERS = (("base", "cand"), ("cand", "base"), ("cand", "base"), ("base", "cand"))

FIXTURES = {
    "45543": f"{SLIDES}/45543.ppt",
    "41246": f"{SLIDES}/41246-1.ppt",
    "floating": f"{DOCS}/FloatingPictures.doc",
    "nohf": f"{DOCS}/NoHeadFoot.doc",
}


def binary(arm, name, round_index):
    """Symlink to bin/ARM/NAME whose path length grows by 8 bytes per round."""
    link_dir = f"{ROOT}/argv0/{arm}"
    os.makedirs(link_dir, exist_ok=True)
    link = f"{link_dir}/{name}-p{'q' * (8 * round_index)}"
    if not os.path.islink(link):
        os.symlink(f"{ROOT}/bin/{arm}/{name}", link)
    return link


def probe(mode, fixture, warmups, samples):
    return lambda arm, r, out: [
        binary(arm, "probe", r),
        "--mode", mode,
        "--input", FIXTURES[fixture],
        "--warmups", str(warmups),
        "--samples", str(samples),
    ]


def harness(cases, shapes_flag, shapes, warmups=10, samples=100):
    return lambda arm, r, out: [
        binary(arm, "litchi-perf-baseline", r),
        "--case", cases,
        shapes_flag, shapes,
        "--samples", str(samples),
        "--warmup", str(warmups),
        "--json", out,
    ]


CASES = {
    # Public lifecycles (0728/0733/0734/0745 selectors).
    "ppt-remove-45543": probe("ppt-remove", "45543", 5, 60),
    "ppt-remove-41246": probe("ppt-remove", "41246", 5, 60),
    "doc-replace-floating": probe("doc-replace", "floating", 5, 60),
    "doc-replace-nohf": probe("doc-replace", "nohf", 10, 300),
    # 0728 common-container control, both policies.
    "container-reuse-45543": probe("container-reuse", "45543", 10, 200),
    "container-rewrite-45543": probe("container-rewrite", "45543", 10, 200),
    "container-reuse-floating": probe("container-reuse", "floating", 10, 200),
    "container-rewrite-floating": probe("container-rewrite", "floating", 10, 200),
    "container-reuse-nohf": probe("container-reuse", "nohf", 20, 1000),
    "container-rewrite-nohf": probe("container-rewrite", "nohf", 20, 1000),
    # CFB-only: one OleWriter::write_to of the edited model.
    "cfb-write-reuse-45543": probe("cfb-write-reuse", "45543", 50, 2000),
    "cfb-write-rewrite-45543": probe("cfb-write-rewrite", "45543", 50, 2000),
    "cfb-write-reuse-41246": probe("cfb-write-reuse", "41246", 50, 2000),
    "cfb-write-reuse-floating": probe("cfb-write-reuse", "floating", 50, 2000),
    "cfb-write-rewrite-floating": probe("cfb-write-rewrite", "floating", 50, 2000),
    "cfb-write-reuse-nohf": probe("cfb-write-reuse", "nohf", 100, 5000),
    "cfb-write-rewrite-nohf": probe("cfb-write-rewrite", "nohf", 100, 5000),
    # Controls that should not change.
    "ppt-noop-45543": probe("ppt-noop", "45543", 10, 200),
    "cfb-open-read-45543": probe("cfb-open-read", "45543", 50, 2000),
    # Harness selectors.
    "harness-ole-edit-save": harness(
        "doc_semantic_one_edit_save,ppt_semantic_one_edit_save,xls_semantic_one_edit_save",
        "--writer-shape", "tiny,large"),
    "harness-ole-noop-save": harness(
        "doc_semantic_noop_edit_save,ppt_semantic_noop_edit_save,xls_semantic_noop_edit_save",
        "--writer-shape", "tiny,large"),
    "harness-cfb-read": harness("cfb_open,cfb_read_one", "--shape", "many-small,few-large",
                                warmups=20, samples=200),
}


def main():
    out_dir = sys.argv[1]
    rounds = int(sys.argv[2])
    prefixes = sys.argv[3:]
    start = int(os.environ.get("ROUND_START", "0"))
    names = [name for name in CASES if not prefixes or name.startswith(tuple(prefixes))]
    os.makedirs(out_dir, exist_ok=True)
    log = []
    for round_index in range(start, start + rounds):
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
                if not name.startswith("harness-"):
                    with open(json_path, "w") as handle:
                        handle.write(completed.stdout)
                if completed.stderr:
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
    log_name = "commands.json" if start == 0 else f"commands-from-r{start:02d}.json"
    with open(f"{out_dir}/{log_name}", "w") as handle:
        json.dump(log, handle, indent=1)
    failures = [entry for entry in log if entry["exit_code"] != 0]
    print(f"processes={len(log)} failures={len(failures)}")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
