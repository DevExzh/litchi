#!/usr/bin/env python3
"""Measure protected PPT probes with explicit phase-span qualification."""
import argparse
import hashlib
import json
import math
import os
from pathlib import Path
import re
import shlex
import statistics
import subprocess
import time

MINIMUM_NS = 150_000_000
TARGET_NS = 500_000_000
PHASES = ("first", "steady", "apply", "inverse")


def digest(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def artifact(path):
    path = Path(path).resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": digest(path)}


def save(path, value):
    Path(path).write_text(json.dumps(value, indent=2) + "\n")


def spans(row):
    # The legacy probe's commit_total_ns includes publication outside its
    # commit timer. At least half of the sorted commit samples are at least
    # their median. Allow for the printed median's 0.001-us rounding and use
    # that rigorous lower bound solely to qualify the steady timing span.
    steady_lower = math.floor(row["patches"] / 2) * max(
        0, row["commit_median_us"] * 1000 - 0.5
    )
    return dict(zip(PHASES, (
        row["first_total_ns"], steady_lower,
        row["apply_total_ns"], row["inverse_apply_total_ns"],
    )))


def metrics(row, rss):
    return {
        "first_mean_ns": row["first_total_ns"] / row["first_count"],
        "steady_median_ns": row["commit_median_us"] * 1000,
        "commit_and_publication_mean_ns": row["commit_total_ns"] / row["patches"],
        "apply_mean_ns": row["apply_total_ns"] / row["replay_count"],
        "inverse_mean_ns": row["inverse_apply_total_ns"] / row["replay_count"],
        "peak_rss_kib": rss,
        "retained_bytes_per_patch": row["retained_patch_bytes"] / row["patches"],
    }


def probe_command(arguments, pattern):
    # Match the program or interpreter's script, never strings embedded in a
    # source-review/hash-bookkeeping shell command. Actual child processes
    # remain independently visible to the guard.
    try:
        words = shlex.split(arguments)
    except ValueError:
        words = arguments.split()
    if not words:
        return False
    program = os.path.basename(words[0])
    if pattern.search(program):
        return True
    if re.fullmatch(r"(?:python(?:3(?:\.\d+)?)?|bash|sh|dash)", program):
        for word in words[1:]:
            if word in ("-c", "-lc", "-ic", "-m"):
                return False
            if word.startswith("-"):
                continue
            return bool(pattern.search(os.path.basename(word)))
    return False


def quiet(out, label, allowed_child=None):
    table = subprocess.check_output(
        ["ps", "-eo", "pid=,ppid=,comm=,args="], text=True
    )
    rows = []
    for line in table.splitlines():
        fields = line.split(None, 3)
        if len(fields) == 4:
            rows.append((int(fields[0]), int(fields[1]), fields[2], fields[3]))
    parents = {pid: parent for pid, parent, _, _ in rows}
    excluded = set()
    current = os.getpid()
    while current > 0 and current not in excluded:
        excluded.add(current)
        current = parents.get(current, 0)
    if allowed_child is not None:
        descendants = {allowed_child}
        while True:
            expanded = descendants | {pid for pid, parent in parents.items() if parent in descendants}
            if expanded == descendants:
                break
            descendants = expanded
        excluded.update(descendants)
    names = re.compile(
        r"^(cargo|rustc|rustdoc|rustfmt|clippy-driver|heaptrack|perf|"
        r"valgrind|callgrind|clang\+\+|ld|lld|ld\.lld|script_program|peculiar[_-].*)$"
    )
    probes = re.compile(
        r"(scoped_patch_probe|snappy-(timed|alloc)|ole_w5_probe|ole_package_clone_probe|"
        r"xlsx_(source_family|merge|tab_state)|litchi-perf-baseline|"
        r"run-abba-w5|run-counting-allocator-w5|ppt-phase-abba[^/ ]*\.py|run-matched-w5[^/ ]*\.py|test-adblock[^/ ]*\.py)"
    )
    # Installed cluster services predate this run and are not benchmark jobs.
    # Keep their full process rows in every snapshot; block source/build-tree
    # peculiar_shell probes and test drivers normally.
    blockers = [list(row) for row in rows if row[0] not in excluded
                and row[3].split(maxsplit=1)[0] != "/usr/local/bin/peculiar_shell" and (
        names.search(row[2]) or probe_command(row[3], probes)
    )]
    save(out / f"{label}.host.json", {
        "time_ns": time.time_ns(), "excluded_processes": sorted(excluded),
        "blockers": blockers, "processes": table,
    })
    if blockers:
        raise RuntimeError(f"host overlap at {label}: {blockers}")


def run_child(out, name, binary, fixture, counts, cpu):
    quiet(out, name + "-before")
    stdout = out / f"{name}.stdout"
    stderr = out / f"{name}.stderr"
    resource = out / f"{name}.time"
    command = ["taskset", "-c", str(cpu), "/usr/bin/time", "-v", "-o",
               str(resource), str(binary), str(fixture), *map(str, counts)]
    start = time.time_ns()
    overlap = None
    with stdout.open("w") as output, stderr.open("w") as errors:
        child = subprocess.Popen(command, stdout=output, stderr=errors)
        check = 0
        while True:
            try:
                returncode = child.wait(timeout=0.1)
                break
            except subprocess.TimeoutExpired:
                if overlap is None:
                    check += 1
                    try:
                        quiet(out, f"{name}-during-{check:04}", child.pid)
                    except RuntimeError as error:
                        overlap = str(error)
    receipt = {
        "command": command, "started_ns": start, "finished_ns": time.time_ns(),
        "exit_code": returncode, "overlap": overlap,
        "artifacts": [artifact(p) for p in (stdout, stderr, resource)],
    }
    save(out / f"{name}.receipt.json", receipt)
    # Preserve the child and receipt before overlap, parsing or qualification
    # checks can reject an observation.
    quiet(out, name + "-after")
    if overlap:
        raise RuntimeError(overlap)
    if returncode or stderr.stat().st_size:
        raise RuntimeError(f"failed probe {name}; inspect retained outputs")
    row = json.loads(stdout.read_text())
    if [row["patches"], row["first_count"], row["replay_count"]] != list(counts):
        raise RuntimeError(f"count mismatch: {name}")
    if row["source_len"] != fixture.stat().st_size or row["target_len"] <= 0:
        raise RuntimeError(f"artifact extent mismatch: {name}")
    for key in ("first_total_ns", "commit_total_ns", "apply_total_ns",
                "inverse_apply_total_ns", "retained_patch_bytes"):
        if not isinstance(row[key], int) or row[key] <= 0:
            raise RuntimeError(f"invalid field {key}: {name}")
    for key in ("commit_median_us", "apply_median_us", "inverse_apply_median_us",
                "first_commit_median_us"):
        if not math.isfinite(row[key]) or row[key] <= 0:
            raise RuntimeError(f"invalid median {key}: {name}")
    match = re.search(r"Maximum resident set size \(kbytes\): (\d+)", resource.read_text())
    if not match:
        raise RuntimeError(f"missing RSS: {name}")
    item = {"name": name, "counts": list(counts), "output": row,
            "spans_ns": spans(row), "metrics": metrics(row, int(match[1])),
            "receipt": artifact(out / f"{name}.receipt.json")}
    save(out / f"{name}.parsed.json", item)
    print(json.dumps({"name": name, "spans_ms": {
        key: round(value / 1e6, 3) for key, value in item["spans_ns"].items()
    }}), flush=True)
    return item


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--before", type=Path, required=True)
    parser.add_argument("--after", type=Path, required=True)
    parser.add_argument("--fixture", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--before-commit", required=True)
    parser.add_argument("--after-commit", required=True)
    parser.add_argument("--harness", type=Path, required=True)
    parser.add_argument("--cpu", type=int, default=29)
    parser.add_argument("--blocks", type=int, default=5)
    args = parser.parse_args()
    if args.blocks < 1 or args.cpu not in os.sched_getaffinity(0):
        parser.error("invalid blocks or CPU")
    args.out = args.out.resolve()
    args.out.mkdir()
    binaries = {"before": args.before.resolve(), "after": args.after.resolve()}
    args.fixture = args.fixture.resolve()
    manifest = {
        "before_commit": args.before_commit, "after_commit": args.after_commit,
        "runner": artifact(__file__), "harness": artifact(args.harness),
        "binaries": {name: artifact(path) for name, path in binaries.items()},
        "fixture": artifact(args.fixture), "cpu": args.cpu, "blocks": args.blocks,
        "minimum_phase_ns": MINIMUM_NS, "calibration_target_ns": TARGET_NS,
        "host_checks": "before, every 100ms while the child runs, and after; own timer/probe descendants excluded",
        "command": ["python3", str(Path(__file__).resolve()), *os.sys.argv[1:]],
        "steady_span": "lower bound: floor(N/2)*(printed_median_ns-0.5); not commit_total_ns",
        "commit_total_ns_scope": "complete commit loop including output serialization and patch retention",
        "fidelity": "probe checks every inverse artifact byte-for-byte; compare source/target extents across endpoints",
        "status": "running", "calibration": [], "paired": [],
    }
    save(args.out / "measurements.json", manifest)
    initial = (1000, 1000, 10000)
    counts = list(initial)
    for side, binary in binaries.items():
        item = run_child(args.out, f"calibration-{side}", binary, args.fixture, initial, args.cpu)
        manifest["calibration"].append(item)
        spans_ = item["spans_ns"]
        factors = (TARGET_NS / max(1, spans_["steady"]),
                   TARGET_NS / max(1, spans_["first"]),
                   max(TARGET_NS / max(1, spans_["apply"]),
                       TARGET_NS / max(1, spans_["inverse"])))
        counts = [max(previous, math.ceil(n * factor * 1.2 / 100) * 100)
                  for previous, n, factor in zip(counts, initial, factors)]
        save(args.out / "measurements.json", manifest)
    if any(value > ceiling for value, ceiling in zip(counts, (20000, 20000, 200000))):
        raise RuntimeError(f"calibration exceeds protected workload ceilings: {counts}")
    manifest["counts"] = counts
    save(args.out / "measurements.json", manifest)
    for block in range(args.blocks):
        legs = []
        for index, side in enumerate(("before", "after", "after", "before")):
            item = run_child(args.out, f"block-{block:02}-{index}-{side}", binaries[side],
                             args.fixture, counts, args.cpu)
            if min(item["spans_ns"].values()) < MINIMUM_NS:
                raise RuntimeError(f"short measured phase in {item['name']}")
            legs.append(item)
        lengths = {(item["output"]["source_len"], item["output"]["target_len"]) for item in legs}
        if len(lengths) != 1:
            raise RuntimeError(f"endpoint artifact lengths differ in block {block}")
        ratios = {
            key: (legs[1]["metrics"][key] + legs[2]["metrics"][key]) /
                 (legs[0]["metrics"][key] + legs[3]["metrics"][key])
            for key in legs[0]["metrics"]
        }
        manifest["paired"].append({"block": block, "after_before": ratios, "legs": legs})
        save(args.out / "measurements.json", manifest)
        print(json.dumps({"block": block, "ratios": ratios}), flush=True)
    manifest["medians"] = {
        key: statistics.median(row["after_before"][key] for row in manifest["paired"])
        for key in manifest["paired"][0]["after_before"]
    }
    manifest["status"] = "complete; coordinator review pending"
    save(args.out / "measurements.json", manifest)
    print(json.dumps(manifest["medians"]), flush=True)


if __name__ == "__main__":
    main()
