#!/usr/bin/env python3
"""M0a S10 ABBA capture driver for the plan 0003 section 5.3 routes.

S1's capture.py runs one binary over one group map. S10 compares two
binaries and, inside the post-S9 binary, two legs, so this driver alternates
three kinds of child per S1 group:

- A: the pre-M0a binary running S1's selectors (capture.py's GROUPS, imported
  unchanged from ../before/capture.py);
- U: the post-S9 binary running the unified successors of those selectors
  (results/0003/s8b/route-map.json), plus the predecessors S9 kept (they run
  in both binaries under the same name);
- L: the post-S9 binary running the leaf twins.

Each block runs, per group, the children in --pattern order (default
A U L L U A): A-U-U-A is the ABBA order for the facade comparison and
U-L-L-U the ABBA order for the overhead check. Blocks run group by group, so
each group's ABBA sequence is contiguous in time.

Every child is S1's child: pinned with taskset, wrapped in GNU time, S1's
warm-up and sample counts, --workers 1, --filesystem-cache warm for the
filesystem groups, a schema-2 corpus manifest except for detection.

Load gate: before a child starts, while the 1-minute load average is above
--load-max, the driver sleeps 30 s and re-checks. The total wait over all
invocations that share --out is capped at --wait-cap-s (state in
wait-state.json); once the cap is spent, children start without waiting and
record load_gate "cap-exhausted".

A child that fails because the shared tmpfs quota is full (the harness's
filesystem samples live in the temporary directory) is re-run in place after
a 60 s pause, up to --quota-retries times; the failed attempt's files keep a
".quota-failed<k>" infix and its children.jsonl record says
"quota_failure": true.

Usage:
  capture_s10.py --before-binary PATH --after-binary PATH --out DIR --cpu C
                 [--groups g1,g2] [--blocks N] [--first-block N]
                 [--pattern AULLUA] [--load-max 8] [--wait-cap-s 7200]
                 [--quota-retries 5] [--filesystem-root DIR] [--gzip-raw]
"""

import argparse
import hashlib
import importlib.util
import json
import os
import subprocess
import sys
import time
from datetime import datetime, timezone

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.abspath(os.path.join(HERE, "..", "..", "..", "..", ".."))
ROUTE_MAP = os.path.join(HERE, "..", "s8b", "route-map.json")


def load_s1_capture():
    spec = importlib.util.spec_from_file_location(
        "s1_capture", os.path.join(HERE, "..", "before", "capture.py"))
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


S1 = load_s1_capture()

# Predecessors that S9 kept registered: they time the detection functions,
# the DOCX leaf owner or the ODS leaf owners, so the post-S9 binary runs them
# under the same name (the U children carry them).
KEPT = {
    "docx_read_one": [
        "docx_semantic_one_paragraph", "docx_semantic_one_paragraph_text",
        "docx_producer_source_selected_paragraph",
    ],
    "detection": [
        f"detect_{fmt}_{kind}"
        for fmt in ("odt", "ods", "odp")
        for kind in ("bytes", "reader", "polyglot", "catalog_alias")
    ],
    "ods_file": [
        "ods_file_eager_selected_cell", "ods_file_source_selected_cell",
        "ods_file_eager_selected_media", "ods_file_source_selected_media",
    ],
}


def successor_groups():
    """S1 group -> (unified cases, leaf cases), in S1's case order."""
    routes = json.load(open(ROUTE_MAP))["routes"]
    by_predecessor = {route["predecessor"]: route for route in routes}
    unified, leaf = {}, {}
    for group, (cases, _extra) in S1.GROUPS.items():
        unified[group], leaf[group] = [], []
        for case in cases:
            route = by_predecessor.get(case)
            if route is None:
                continue
            unified[group].append(route["unified"])
            if route["leaf"]:
                leaf[group].append(route["leaf"])
        unified[group] += KEPT.get(group, [])
    return unified, leaf


UNIFIED, LEAF = successor_groups()


def cases_for(leg, group):
    if leg == "A":
        return S1.GROUPS[group][0]
    if leg == "U":
        return UNIFIED[group]
    if leg == "L":
        return LEAF[group]
    raise ValueError(leg)


def loadavg():
    with open("/proc/loadavg") as handle:
        return handle.read().split()[:3]


def now():
    return datetime.now(timezone.utc).isoformat()


def append_record(path, line, attempts=40):
    """Append one children.jsonl line, waiting out a full tmpfs quota."""
    for _ in range(attempts):
        try:
            with open(path, "a") as handle:
                handle.write(line)
            return
        except OSError as error:
            if error.errno != 122:  # EDQUOT
                raise
            time.sleep(30)
    raise OSError(122, f"could not append to {path}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--before-binary", required=True)
    parser.add_argument("--after-binary", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--cpu", required=True)
    parser.add_argument("--groups", default=",".join(S1.GROUPS))
    parser.add_argument("--blocks", type=int, default=1)
    parser.add_argument("--first-block", type=int, default=1)
    parser.add_argument("--pattern", default="AULLUA")
    parser.add_argument("--samples", type=int, default=100)
    parser.add_argument("--warmup", type=int, default=10)
    parser.add_argument("--slow-groups", default="pptx_file,docx_file")
    parser.add_argument("--slow-samples", type=int, default=30)
    parser.add_argument("--slow-warmup", type=int, default=5)
    parser.add_argument("--load-max", type=float, default=8.0)
    parser.add_argument("--wait-cap-s", type=int, default=7200)
    parser.add_argument("--label", default="s10")
    parser.add_argument("--quota-retries", type=int, default=5)
    # The harness writes one corpus copy per filesystem case (12 x 17 MB for
    # PPTX) under the temporary directory unless told otherwise; on a host
    # whose shared tmpfs quota is full, point it at a disk directory instead.
    parser.add_argument("--filesystem-root", default=None)
    parser.add_argument("--gzip-raw", action="store_true")
    args = parser.parse_args()

    groups = args.groups.split(",")
    for group in groups:
        if group not in S1.GROUPS:
            sys.exit(f"unknown group {group}")
    slow = set(args.slow_groups.split(","))
    binaries = {
        "A": args.before_binary,
        "U": args.after_binary,
        "L": args.after_binary,
    }
    shas = {leg: S1.sha256_file(path) for leg, path in binaries.items()}
    os.makedirs(os.path.join(args.out, "raw"), exist_ok=True)
    log_path = os.path.join(args.out, "children.jsonl")
    wait_path = os.path.join(args.out, "wait-state.json")
    waited_total = json.load(open(wait_path))["waited_s"] if os.path.exists(wait_path) else 0

    for block in range(args.first_block, args.first_block + args.blocks):
        for group in groups:
            _cases, extra = S1.GROUPS[group]
            for position, leg in enumerate(args.pattern, start=1):
                cases = cases_for(leg, group)
                if not cases:
                    continue
                stem = f"B{block}-{position}{leg}-{group}"
                json_out = os.path.join(args.out, "raw", stem + ".json")
                manifest_out = os.path.join(args.out, "raw", stem + ".corpus-manifest.json")
                time_out = os.path.join(args.out, "raw", stem + ".time")
                stderr_out = os.path.join(args.out, "raw", stem + ".stderr")
                for path in (json_out, manifest_out, json_out + ".gz"):
                    if os.path.exists(path):
                        sys.exit(f"refusing to overwrite {path}")
                waited = 0
                gate = "ok"
                while float(loadavg()[0]) > args.load_max:
                    if waited_total >= args.wait_cap_s:
                        gate = "cap-exhausted"
                        break
                    time.sleep(30)
                    waited += 30
                    waited_total += 30
                    gate = "waited"
                    with open(wait_path, "w") as handle:
                        json.dump({"waited_s": waited_total, "cap_s": args.wait_cap_s}, handle)
                argv = [
                    "taskset", "-c", args.cpu,
                    "/usr/bin/time", "-v", "-o", time_out,
                    binaries[leg],
                    "--case", ",".join(cases),
                    "--warmup", str(args.slow_warmup if group in slow else args.warmup),
                    "--samples", str(args.slow_samples if group in slow else args.samples),
                    "--workers", "1",
                    *extra,
                    "--json", json_out,
                ]
                if group in S1.FILESYSTEM_GROUPS:
                    argv += ["--filesystem-cache", "warm"]
                    if args.filesystem_root:
                        argv += ["--filesystem-root", args.filesystem_root]
                if group not in S1.NO_MANIFEST:
                    argv += ["--corpus-manifest", manifest_out]
                for attempt in range(args.quota_retries + 1):
                    record = {
                        "label": args.label,
                        "block": block,
                        "position": position,
                        "leg": leg,
                        "group": group,
                        "argv": argv,
                        "cwd": REPO,
                        "binary_sha256": shas[leg],
                        "load_gate": gate,
                        "waited_s": waited,
                        "waited_total_s": waited_total,
                        "attempt": attempt,
                        "started_utc": now(),
                        "loadavg_before": loadavg(),
                    }
                    started = time.monotonic()
                    with open(stderr_out, "wb") as stderr:
                        proc = subprocess.run(argv, cwd=REPO, stdout=subprocess.DEVNULL, stderr=stderr,
                                              env={**os.environ, "LC_ALL": "C", "TZ": "UTC"})
                    record.update({
                        "finished_utc": now(),
                        "seconds": round(time.monotonic() - started, 3),
                        "loadavg_after": loadavg(),
                        "exit_code": proc.returncode,
                    })
                    if os.path.exists(time_out):
                        record["gnu_time"] = S1.parse_time_file(time_out)
                    # The harness writes its filesystem samples to the
                    # temporary directory, a tmpfs whose per-user quota every
                    # agent on this host shares. A child that fails on that
                    # quota is set aside (its files keep a .quota-failed<k>
                    # infix) and re-run in place after a pause.
                    quota = proc.returncode != 0 and "Disk quota exceeded" in open(
                        stderr_out, errors="replace").read()
                    record["quota_failure"] = quota
                    append_record(log_path, json.dumps(record) + "\n")
                    if args.gzip_raw and proc.returncode == 0:
                        for path in (json_out, manifest_out):
                            if os.path.exists(path):
                                subprocess.run(["gzip", "-9", "-n", path], check=True)
                    print(f"{stem}: exit={proc.returncode} {record['seconds']}s "
                          f"load={record['loadavg_before'][0]}->{record['loadavg_after'][0]} "
                          f"gate={gate} waited={waited}s attempt={attempt}", flush=True)
                    if not quota or attempt == args.quota_retries:
                        break
                    for path in (json_out, manifest_out, time_out, stderr_out):
                        if os.path.exists(path):
                            suffix = path[len(os.path.join(args.out, "raw", stem)):]
                            os.rename(path, os.path.join(args.out, "raw",
                                                         f"{stem}.quota-failed{attempt}{suffix}"))
                    time.sleep(60)


if __name__ == "__main__":
    main()
