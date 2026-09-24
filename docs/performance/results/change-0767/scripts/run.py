#!/usr/bin/env python3
"""Change 0767 two-arm ABBA process matrix with heap-layout randomization.

Arms: base = 1d1044e3ac, cand = the 0767 change (af72192785). Both arms'
binaries are built with identical cargo commands, flags and features (see
binaries.json), and copied to bin/ARM/ so their paths have equal length.

Every process is pinned with `taskset -c 28`. Round r runs every case once per
arm, in ABBA order across rounds (base,cand / cand,base / cand,base /
base,cand, repeating), with the case order rotated each round. As in changes
0745 and 0749, each binary is started through a symlink whose path grows by 8
bytes per round, so both arms of a round share one argv[0] heap layout and
the rounds sample many layouts. Raw reports, stderr, exact commands, exit
codes and start times are retained.

Usage: run.py OUT_DIR ROUNDS [CASE_PREFIX ...]

`ROUND_START` (default 0) continues a matrix with the same order cycle and
argv[0] progression; its command log is `commands-from-rROUND_START.json`.
"""

import json
import os
import subprocess
import sys
import time

ROOT = "/home/zhuhe/code/litchi-worktrees/scratch/0767"
TREE = "/home/zhuhe/code/litchi-worktrees/0767-cfb-reparse-linear"
SLIDES = f"{TREE}/test-data/poi/test-data/slideshow"
OLE = f"{TREE}/test-data/ole"
CORE = "28"
ORDERS = (("base", "cand"), ("cand", "base"), ("cand", "base"), ("base", "cand"))

FIXTURES = {
    "45543": f"{SLIDES}/45543.ppt",
    "floating": f"{OLE}/doc/FloatingPictures.doc",
    "checkboxes": f"{OLE}/xls/WithCheckBoxes.xls",
}
for version in ("v3", "v4"):
    for kind in ("mini", "regular"):
        for count in (1000, 3000, 10000):
            FIXTURES[f"{version}-{kind}-{count}"] = f"{ROOT}/gen/{version}-{kind}-{count}.cfb"


def binary(arm, name, round_index):
    """Symlink to bin/ARM/NAME whose path length grows by 8 bytes per round."""
    link_dir = f"{ROOT}/argv0/{arm}"
    os.makedirs(link_dir, exist_ok=True)
    link = f"{link_dir}/{name}-p{'q' * (8 * round_index)}"
    if not os.path.islink(link):
        os.symlink(f"{ROOT}/bin/{arm}/{name}", link)
    return link


def probe(mode, fixture, warmups, samples, edit=None):
    def command(arm, r, out):
        argv = [
            binary(arm, "probe", r),
            "--mode", mode,
            "--input", FIXTURES[fixture],
            "--warmups", str(warmups),
            "--samples", str(samples),
        ]
        if edit:
            argv += ["--edit", edit]
        return argv
    return command


def harness(cases, shapes_flag, shapes, warmups=10, samples=100, extra=()):
    def command(arm, r, out):
        argv = [binary(arm, "litchi-perf-baseline", r), "--case", cases, shapes_flag, shapes]
        return argv + list(extra) + [
            "--samples", str(samples), "--warmup", str(warmups), "--json", out]
    return command


def many_stream_samples(count, per_stream_us):
    """Samples and warmups for about 0.25 s of timed work per process."""
    owner_us = max(count * per_stream_us, 1.0)
    samples = int(min(2000, max(20, 250_000 / owner_us)))
    return max(3, samples // 10), samples


CASES = {}
for version in ("v3", "v4"):
    for kind in ("mini", "regular"):
        for count in (1000, 3000, 10000):
            name = f"{version}-{kind}-{count}"
            warmups, samples = many_stream_samples(count, 0.45)
            CASES[f"open-{name}"] = probe("cfb-open", name, warmups, samples)
            warmups, samples = many_stream_samples(count, 0.8)
            CASES[f"readall-{name}"] = probe("cfb-read-all", name, warmups, samples)
# The shared reader's open and the common editor's open-and-capture.
for name in ("v3-mini-10000", "v3-regular-10000"):
    warmups, samples = many_stream_samples(10000, 0.45)
    CASES[f"shared-open-{name}"] = probe("shared-open", name, warmups, samples)
for name in ("v3-mini-1000", "v3-mini-3000", "v3-mini-10000"):
    count = int(name.rsplit("-", 1)[1])
    warmups, samples = many_stream_samples(count, 2.0)
    CASES[f"editor-open-{name}"] = probe("editor-open", name, warmups, samples)
# The Reuse write whose validation reparses the plan (0749's lane).
CASES["write-reuse-45543"] = probe("cfb-write-reuse", "45543", 50, 2000)
CASES["write-reuse-v3-mini-3000-grow"] = probe("cfb-write-reuse", "v3-mini-3000", 5, 40, "grow")
# Real small files: the common benign case.
CASES["open-45543"] = probe("cfb-open", "45543", 100, 5000)
CASES["open-floating"] = probe("cfb-open", "floating", 100, 5000)
CASES["open-checkboxes"] = probe("cfb-open", "checkboxes", 100, 5000)
CASES["openread-45543"] = probe("cfb-open-read", "45543", 50, 2000)
CASES["editor-open-45543"] = probe("editor-open", "45543", 50, 2000)
# Controls whose timed work does not run the changed code.
CASES["ctl-write-rewrite-45543"] = probe("cfb-write-rewrite", "45543", 50, 2000)
# Harness selectors.
CASES["harness-cfb"] = harness(
    "cfb_open,cfb_list_streams", "--shape", "tiny,many-small,few-large,wide-root",
    warmups=20, samples=200, extra=["--payload", "incompressible"])
CASES["harness-ole-open"] = harness(
    "doc_semantic_open,xls_semantic_open,ppt_semantic_open", "--semantic-shape",
    "tiny,medium,large", warmups=10, samples=100)


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
