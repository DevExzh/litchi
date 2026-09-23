#!/usr/bin/env python3
"""Change 0749 round two: the mini-stream fix, three probe arms.

Arms (see binaries-r2.json):
  A = ab29ac6291, the base;
  B = 776ad05175, the first 0749 version, whose mini-stream readback is
      quadratic (reviewer's blocker);
  C = 59d853618f, the fix (root mini stream loaded once per validation,
      contiguous mini sectors compared as one range, chain buffers cleared
      in proportion to the chains).
Harness selectors run arms A and C only (B's reader is A's).

Every process is pinned with `taskset -c 20`. Round r runs every case once
per arm. Probe cases cycle through all six orders of (A, B, C); harness cases
alternate A,C / C,A / C,A / A,C. Case order rotates each round. As in 0745
and round one, each binary starts through a symlink whose path grows 8 bytes
per round, so all arms of a round share one argv[0] heap layout; arm
directory names have equal length. Raw reports, stderr, commands, exit codes
and start times are kept.

Usage: run2.py OUT_DIR ROUNDS [CASE_PREFIX ...]
"""

import itertools
import json
import os
import subprocess
import sys
import time

ROOT = "/home/zhuhe/code/litchi-worktrees/scratch/0749"
TREE = "/home/zhuhe/code/litchi-worktrees/0749-cfb-reuse-plan-validation"
T = f"{TREE}/test-data"
CORE = "20"
PROBE_ORDERS = list(itertools.permutations(("A", "B", "C")))
HARNESS_ORDERS = (("A", "C"), ("C", "A"), ("C", "A"), ("A", "C"))

INPUTS = {
    "45543": f"{T}/poi/test-data/slideshow/45543.ppt",
    "41246": f"{T}/poi/test-data/slideshow/41246-1.ppt",
    "floating": f"{T}/ole/doc/FloatingPictures.doc",
    "nohf": f"{T}/ole/doc/NoHeadFoot.doc",
    "hyperlink": f"{T}/ole/doc/hyperlink.doc",
    "empty": f"{T}/ole/ppt/empty.ppt",
    "checkboxes": f"{T}/ole/xls/WithCheckBoxes.xls",
    "v3-1000x2000": f"{ROOT}/gen/v3-1000x2000.cfb",
    "v3-3000x2000": f"{ROOT}/gen/v3-3000x2000.cfb",
    "v4-3000x2000": f"{ROOT}/gen/v4-3000x2000.cfb",
    "v4-10000x4000": f"{ROOT}/gen/v4-10000x4000.cfb",
}


def binary(arm, name, round_index):
    """Symlink to bin/ARM/NAME whose path length grows by 8 bytes per round."""
    link_dir = f"{ROOT}/argv0/{arm}"
    os.makedirs(link_dir, exist_ok=True)
    link = f"{link_dir}/{name}-p{'q' * (8 * round_index)}"
    if not os.path.islink(link):
        os.symlink(f"{ROOT}/bin/{arm}/{name}", link)
    return link


def probe(mode, source, edit, warmups, samples):
    return {
        "arms": "probe",
        "build": lambda arm, r, out: [
            binary(arm, "probe", r),
            "--mode", mode,
            "--edit", edit,
            "--input", INPUTS[source],
            "--warmups", str(warmups),
            "--samples", str(samples),
        ],
    }


def harness(cases, shapes_flag, shapes, warmups=10, samples=100):
    return {
        "arms": "harness",
        "build": lambda arm, r, out: [
            binary(arm, "litchi-perf-baseline", r),
            "--case", cases,
            shapes_flag, shapes,
            "--samples", str(samples),
            "--warmup", str(warmups),
            "--json", out,
        ],
    }


CASES = {
    # The four original fixtures, public edits (round one's CFB-only lane).
    "reuse-45543-public": probe("cfb-write-reuse", "45543", "public", 50, 2000),
    "reuse-41246-public": probe("cfb-write-reuse", "41246", "public", 50, 2000),
    "reuse-floating-public": probe("cfb-write-reuse", "floating", "public", 50, 2000),
    "reuse-nohf-public": probe("cfb-write-reuse", "nohf", "public", 100, 5000),
    # All-mini-stream fixtures the reviewer flagged, both corpus edits.
    "reuse-hyperlink-same": probe("cfb-write-reuse", "hyperlink", "same", 100, 5000),
    "reuse-hyperlink-grow": probe("cfb-write-reuse", "hyperlink", "grow", 100, 5000),
    "reuse-empty-same": probe("cfb-write-reuse", "empty", "same", 100, 5000),
    "reuse-empty-grow": probe("cfb-write-reuse", "empty", "grow", 100, 5000),
    "reuse-checkboxes-same": probe("cfb-write-reuse", "checkboxes", "same", 100, 5000),
    "reuse-checkboxes-grow": probe("cfb-write-reuse", "checkboxes", "grow", 100, 5000),
    # Generated many-mini-stream files, corpus length-changing edit.
    "reuse-v3-1000x2000-grow": probe("cfb-write-reuse", "v3-1000x2000", "grow", 3, 40),
    "reuse-v3-3000x2000-grow": probe("cfb-write-reuse", "v3-3000x2000", "grow", 2, 20),
    "reuse-v4-3000x2000-grow": probe("cfb-write-reuse", "v4-3000x2000", "grow", 2, 20),
    "reuse-v4-10000x4000-grow": probe("cfb-write-reuse", "v4-10000x4000", "grow", 1, 8),
    # Controls: the Rewrite policy and the reader path.
    "rewrite-45543-public": probe("cfb-write-rewrite", "45543", "public", 50, 2000),
    "rewrite-checkboxes-grow": probe("cfb-write-rewrite", "checkboxes", "grow", 100, 5000),
    "open-read-45543": probe("cfb-open-read", "45543", "public", 50, 2000),
    # Harness selectors (A and C).
    "harness-ole-edit-save": harness(
        "doc_semantic_one_edit_save,ppt_semantic_one_edit_save,xls_semantic_one_edit_save",
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
        shift = round_index % len(names)
        rotated = names[shift:] + names[:shift]
        for name in rotated:
            case = CASES[name]
            if case["arms"] == "probe":
                order = PROBE_ORDERS[round_index % len(PROBE_ORDERS)]
            else:
                order = HARNESS_ORDERS[round_index % len(HARNESS_ORDERS)]
            for arm in order:
                stem = f"{out_dir}/{name}/r{round_index:02d}-{arm}"
                os.makedirs(os.path.dirname(stem), exist_ok=True)
                json_path = f"{stem}.json"
                command = ["taskset", "-c", CORE] + case["build"](arm, round_index, json_path)
                started = time.time()
                completed = subprocess.run(command, capture_output=True, text=True, timeout=600)
                if case["arms"] == "probe":
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
