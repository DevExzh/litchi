#!/usr/bin/env python3
"""Matched constructor ABBA; all attempts and host snapshots are retained."""
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

TARGET_NS = 500_000_000
MINIMUM_NS = 150_000_000
CASES = ("field_path", "field_info", "message_info", "archive_info", "archive_object",
         "archive_object_reject", "parse", "validate", "encode",
          "retained_parse", "retained_validate", "retained_encoded_len", "retained_encode",
          "retained_validate_tight", "retained_encode_tight")


def artifact(path):
    path = Path(path).resolve()
    return dict(path=str(path), bytes=path.stat().st_size,
                sha256=hashlib.sha256(path.read_bytes()).hexdigest())


def save(path, value):
    Path(path).write_text(json.dumps(value, indent=2) + "\n")


def program(arguments):
    try:
        words = shlex.split(arguments)
    except ValueError:
        words = arguments.split()
    if not words:
        return ""
    name = os.path.basename(words[0])
    if re.fullmatch(r"(?:python(?:3(?:\.\d+)?)?|bash|sh|dash)", name):
        for word in words[1:]:
            if word in ("-c", "-lc", "-ic"):
                return name
            if word.startswith("-"):
                continue
            return os.path.basename(word)
    return name


def quiet(output, label, child=None):
    rows = []
    for entry in Path("/proc").iterdir():
        if not entry.name.isdigit():
            continue
        try:
            stat = (entry / "stat").read_text()
            parent = int(stat[stat.rfind(")") + 2:].split()[1])
            arguments = (entry / "cmdline").read_bytes().replace(b"\0", b" ").decode(errors="replace")
            executable = os.readlink(entry / "exe")
            rows.append(dict(pid=int(entry.name), parent=parent, exe=executable, args=arguments))
        except (FileNotFoundError, ProcessLookupError, PermissionError):
            continue
    parents = {row["pid"]: row["parent"] for row in rows}
    exempt = set()
    current = os.getpid()
    while current and current not in exempt:
        exempt.add(current)
        current = parents.get(current, 0)
    if child is not None:
        descendants = {child}
        while True:
            expanded = descendants | {pid for pid, parent in parents.items() if parent in descendants}
            if expanded == descendants:
                break
            descendants = expanded
        exempt |= descendants
    forbidden = re.compile(
        r"(?:^(?:cargo|rustc|rustdoc|rustfmt|clippy-driver|heaptrack|perf|valgrind|"
        r"callgrind|clang\+\+|clang|gcc|g\+\+|cc|cmake|ninja|make|ld|lld|ld\.lld|"
        r"mold|script_program|pytest|nextest)$|"
        r"peculiar|test-adblock|scoped_patch_probe|char-style|char_style|constructor|odraw.*probe|"
        r"file_w8_.*probe|ole_.*probe|xlsx_.*probe|litchi-perf-baseline|"
        r"run-abba|run-matched-w5|run-counting-allocator)")
    blockers = [row for row in rows if row["pid"] not in exempt
                and row["exe"] != "/usr/local/bin/peculiar_shell"
                and (forbidden.search(os.path.basename(row["exe"]))
                     or forbidden.search(program(row["args"])))]
    save(output / f"{label}.host.json", dict(time_ns=time.time_ns(),
         exempt=sorted(exempt), blockers=blockers, processes=rows))
    if blockers:
        raise RuntimeError(f"host overlap at {label}: {blockers}")


def invoke(output, label, binary, route, count, cpu):
    quiet(output, label + "-before")
    stdout, stderr = output / f"{label}.stdout", output / f"{label}.stderr"
    command = ["taskset", "-c", str(cpu), str(binary), route, str(count)]
    overlap = None
    started = time.time_ns()
    with stdout.open("w") as out, stderr.open("w") as err:
        child = subprocess.Popen(command, stdout=out, stderr=err)
        checks = 0
        while True:
            try:
                status = child.wait(timeout=0.1)
                break
            except subprocess.TimeoutExpired:
                checks += 1
                try:
                    quiet(output, f"{label}-during-{checks:04}", child.pid)
                except RuntimeError as error:
                    overlap = str(error)
                    child.terminate()
                    status = child.wait()
                    break
    receipt = dict(command=command, started_ns=started, finished_ns=time.time_ns(),
                   exit_code=status, overlap=overlap,
                   stdout=artifact(stdout), stderr=artifact(stderr))
    save(output / f"{label}.receipt.json", receipt)
    quiet(output, label + "-after")
    if status or overlap:
        raise RuntimeError(f"invalid attempt {label}: {receipt}")
    value = json.loads(stdout.read_text())
    if (value.get("route"), value.get("iterations")) != (route, count):
        raise RuntimeError(f"probe identity mismatch {label}")
    if not isinstance(value.get("ns"), int) or value["ns"] <= 0:
        raise RuntimeError(f"missing timing {label}")
    return value


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("base", type=Path)
    parser.add_argument("candidate", type=Path)
    parser.add_argument("out", type=Path)
    parser.add_argument("--cpu", type=int, default=29)
    args = parser.parse_args()
    assert args.cpu in os.sched_getaffinity(0)
    out = args.out.resolve()
    out.mkdir(parents=True, exist_ok=False)
    binaries = {"A": args.base.resolve(), "B": args.candidate.resolve()}
    manifest = dict(schema="constructor-abba.v1", binaries={key: artifact(value) for key, value in binaries.items()},
                    runner=artifact(__file__), cpu=args.cpu, target_ns=TARGET_NS,
                    minimum_ns=MINIMUM_NS, cases=[], status="running")
    save(out / "manifest.json", manifest)
    for route in CASES:
        prefix = route
        count = 1000
        calibrations = []
        while True:
            results = {side: invoke(out, f"{prefix}-cal-{count}-{side}", binary, route, count, args.cpu)
                       for side, binary in binaries.items()}
            calibrations.append(results)
            fastest = min(row["ns"] for row in results.values())
            if fastest >= TARGET_NS:
                break
            count = max(count + 1, math.ceil(count * TARGET_NS / fastest * 1.2))
            if count > 1_000_000_000:
                raise RuntimeError(f"calibration exceeds ceiling: {prefix}")
        for warmup in range(2):
            for side, binary in binaries.items():
                invoke(out, f"{prefix}-warm-{warmup}-{side}", binary, route, count, args.cpu)
        case = dict(route=route, count=count, calibrations=calibrations, blocks=[])
        manifest["cases"].append(case)
        save(out / "manifest.json", manifest)
        for block in range(5):
            legs = []
            for index, side in enumerate("ABBA"):
                label = f"{prefix}-b{block}-{index}-{side}"
                value = invoke(out, label, binaries[side], route, count, args.cpu)
                if value["ns"] < MINIMUM_NS:
                    raise RuntimeError(f"sample below minimum: {label}")
                legs.append(dict(side=side, label=label, value=value))
            assert len({row["value"]["checksum"] for row in legs}) == 1
            assert len({row["value"]["output_sha256"] for row in legs}) == 1
            ratio = sum(row["value"]["ns"] for row in legs[1:3]) / sum(legs[i]["value"]["ns"] for i in (0, 3))
            case["blocks"].append(dict(legs=legs, ratio=ratio))
            save(out / "manifest.json", manifest)
        case["median_ratio"] = statistics.median(row["ratio"] for row in case["blocks"])
        save(out / "manifest.json", manifest)
        print(prefix, case["median_ratio"], flush=True)
    quiet(out, "final")
    if {key: artifact(value) for key, value in binaries.items()} != manifest["binaries"]:
        raise RuntimeError("protected binaries changed during measurement")
    manifest["status"] = "complete"
    save(out / "manifest.json", manifest)


if __name__ == "__main__":
    main()
