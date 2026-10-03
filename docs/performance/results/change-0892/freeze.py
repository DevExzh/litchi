#!/usr/bin/env python3
"""Freeze both diagnostic populations before launching either profiler."""
import datetime
import json
import os
from pathlib import Path
import platform
import subprocess

import source_identity

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
HELPERS = ("source_identity.py", "freeze.py", "capture.py", "export.py", "reduce.py",
           "verify.py", "capture_callgrind.py", "analyze_callgrind.py")


def read(path):
    return json.loads(path.read_text())


def descriptor(path):
    return source_identity.identity(path)


def write(path, value):
    with path.open("x") as f:
        json.dump(value, f, indent=2)
        f.write("\n")


def main():
    assert not (P / "plan.json").exists() and not (P / "callgrind-plan.json").exists()
    source = source_identity.audit()
    assert source == read(P / "source-identity.json"), "source identity changed before freeze"
    review = read(P / "methodology-review.json")
    assert review["status"] == "pass", "methodology review must pass before capture"
    binary = read(P.parent / "change-0891/build.json")["binaries"]["native"]
    helpers = {name: descriptor(P / name) for name in HELPERS}
    assert review["helper_hashes"] == {name: item["sha256"] for name, item in helpers.items()}, "reviewed helpers changed"
    upstream = [P.parent / "change-0885" / name for name in ("capture.py", "verify.py")]
    upstream += [P.parent / "change-0886" / name for name in
                 ("capture.py", "export.py", "reduce.py", "verify_packet.py", "callgrind.py", "analyze_callgrind.py")]
    tools = [Path(x) for x in ("/usr/bin/perf", "/usr/bin/taskset", "/usr/bin/valgrind",
                              "/usr/bin/callgrind_annotate", "/usr/bin/readelf")]
    engine = Path("/usr/libexec/valgrind/callgrind-amd64-linux")
    assert engine.is_file(), "record the actual Callgrind engine identity"
    tools.append(engine)
    build_notes = subprocess.check_output(["/usr/bin/readelf", "-n", binary["path"]], text=True)
    build_ids = [x.split("Build ID:", 1)[1].strip() for x in build_notes.splitlines() if "Build ID:" in x]
    assert len(build_ids) == 1
    common = {
        "created_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
        "source_revision": source_identity.REVISION,
        "source_identity": descriptor(P / "source-identity.json"),
        "baseline_commit": source_identity.REVISION,
        "profile": "release", "fresh_child": True, "fail_fast": True,
        "binary": binary, "elf_build_id": build_ids[0], "cpu": 12,
        "case": "xlsx_producer_medium_source_full_cell_scan", "filesystem_cache": "warm",
        "schedule": "sequential; no retry or slow-row replacement",
        "failure_policy": "stop on command, semantic, custody, or zero-collection failure; preserve every attempt",
        "environment": {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"},
        "host": {"platform": platform.platform(), "affinity_allowed": sorted(os.sched_getaffinity(0)),
                 "loadavg": Path("/proc/loadavg").read_text().strip(), "exclusive_host": False,
                 "physical_cold_cache": False},
        "bindings": {"source_identity": descriptor(P / "source-identity.json"),
                     "methodology_review": descriptor(P / "methodology-review.json"),
                     "helpers": helpers, "upstream": [descriptor(x) for x in upstream],
                     "tools": [descriptor(x) for x in tools],
                     "prior_self_analysis": descriptor(P.parent / "change-0886/self-analysis.json"),
                     "prior_callgrind_analysis": descriptor(P.parent / "change-0886/callgrind-analysis.json")},
        "performance_claim": None,
    }
    assert 12 in common["host"]["affinity_allowed"]
    native = {**common, "schema": "litchi.xlsx-scan-cycle-profile-0892.v1",
              "children": 3, "warmup": 10, "samples": 2000,
              "event": "cycles:u", "frequency_hz": 499, "call_graph": "dwarf,16384",
              "perf_version": subprocess.check_output(["/usr/bin/perf", "--version"], text=True).strip(),
              "scope": "Whole-process sampled-IP self periods including setup, timer-excluded checks, destruction and reporting; no native inclusive attribution or latency claim."}
    guest = {**common, "schema": "litchi.xlsx-scan-callgrind-diagnostic-0892.v1",
             "children": 2, "samples": 10, "warmup": 0,
             "toggle_function": "litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan",
             "valgrind_version": subprocess.check_output(["/usr/bin/valgrind", "--version"], text=True).strip(),
             "scope": "Entire producer full-cell runner, including oracle setup and timer-excluded work; guest instructions separate from native cycles and latency."}
    write(P / "plan.json", native)
    write(P / "callgrind-plan.json", guest)
    print("Both diagnostic plans frozen; no workload launched")


if __name__ == "__main__":
    main()
