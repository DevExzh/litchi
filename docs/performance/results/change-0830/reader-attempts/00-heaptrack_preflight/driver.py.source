"""Root-only, serial execution for the 0830 XLSX allocation investigation.

Run stages in order: prepare, quality, build, capture. Analysis is separate.
Every child has a durable command receipt, stdout/stderr log, and exit status.
Inputs are hashed before and after each stage; existing evidence is never replaced.
"""
from __future__ import annotations

import hashlib
import itertools
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BASE = "87eb57182be1622385dc3b28dfc3c7be868dca32"
TARGET = ROOT.parent / "litchi-target-0830"
MANIFEST = P / "probe-src/Cargo.toml"
INPUT = ROOT / "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx"
REFERENCE = ROOT / "docs/performance/results/change-0821/artifacts/real-001-xlsx/default.xlsx"
OWNER = "xlsx_edit_profile_0830::edit_region_0830"
UNRELATED = {
    "docs/FORMAT_IMPLEMENTATION_REVIEW.md": "bffd00f144c4c1bbb3b0805d21352b40e9ae46f7366c03d6e581ca61b1f27ce5",
    "docs/UNIFIED_OPS_API_DESIGN.md": "f5672c38393a2a6c52f028b2a501ddad93766ef974dbdb46cdc3c45e1db3ef6d",
    "matrix-analysis.json": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
}


def sha(path):
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def read(path):
    return json.loads(Path(path).read_text())


def write(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x") as f:
        f.write(json.dumps(value, indent=2, sort_keys=True) + "\n")


def output(argv):
    return subprocess.check_output(argv, cwd=ROOT, text=True).strip()


def inventory():
    tracked = output(["git", "ls-files", "crates", "Cargo.toml", "Cargo.lock",
                      "rustfmt.toml", "clippy.toml", ".cargo", "rust-toolchain.toml",
                      "docs/adr", "docs/GOAL.md", "docs/CRUD_Scenario_Checklist.md",
                      "tools/perf-baseline"]).splitlines()
    result = {name: sha(ROOT / name) for name in tracked}
    for path in (INPUT, REFERENCE):
        result[str(path.relative_to(ROOT))] = sha(path)
    for path in sorted(P.rglob("*")):
        if path.is_file() and (path.name == "driver.py" or "probe-src" in path.parts):
            result[str(path.relative_to(ROOT))] = sha(path)
    return result


def check_inputs():
    frozen = read(P / "inputs.json")
    current = inventory()
    assert current == frozen, "source, normative, corpus, driver or probe drift"
    assert all(sha(ROOT / n) == h for n, h in UNRELATED.items()), "unrelated work drift"
    assert output(["git", "rev-parse", "HEAD"]) == BASE, "base changed"
    prepared = read(P / "prepare.json")
    for name in ("inputs", "plan", "host"):
        assert sha(P / (name + ".json")) == prepared[name + "_sha256"], name


def environment(arm="ordinary"):
    forbidden = [k for k in os.environ if k.startswith(("CARGO_PROFILE_", "CARGO_TARGET_"))
                 or k in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_WRAPPER",
                          "RUSTC_WORKSPACE_WRAPPER", "LD_PRELOAD", "LD_LIBRARY_PATH")]
    assert not forbidden, f"ambient build/profiler overrides: {forbidden}"
    env = dict(os.environ)
    env.update(CARGO_TARGET_DIR=str(TARGET / arm), CARGO_BUILD_JOBS="2",
               CARGO_INCREMENTAL="0", RUSTFLAGS="-C force-frame-pointers=yes" if arm == "fp" else "",
               RUSTDOCFLAGS="-D warnings", LC_ALL="C", TZ="UTC")
    return env


def run(name, argv, arm="ordinary"):
    receipt = P / "commands" / (name + ".json")
    start = P / "commands" / (name + ".started.json")
    env = environment(arm)
    row = {"argv": list(map(str, argv)), "cwd": str(ROOT), "started_unix": time.time(),
           "input_inventory_sha256": sha(P / "inputs.json"),
           "environment": {k: env[k] for k in ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS",
               "CARGO_INCREMENTAL", "RUSTFLAGS", "RUSTDOCFLAGS", "LC_ALL", "TZ")}}
    write(start, row)
    with (P / "commands" / (name + ".log")).open("xb") as log:
        result = subprocess.run(row["argv"], cwd=ROOT, env=env, stdout=log, stderr=subprocess.STDOUT)
    row.update(finished_unix=time.time(), exit_code=result.returncode,
               log_sha256=sha(P / "commands" / (name + ".log")))
    write(receipt, row)
    print(name, result.returncode, flush=True)
    assert result.returncode == 0, name


def prepare():
    assert output(["git", "rev-parse", "HEAD"]) == BASE
    assert not TARGET.exists(), TARGET
    assert all(sha(ROOT / n) == h for n, h in UNRELATED.items())
    assert INPUT.stat().st_size == 8435
    assert sha(INPUT) == "d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4"
    assert REFERENCE.stat().st_size == 8521
    assert sha(REFERENCE) == "0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68"
    lock_parity = read(P / "lock-parity.json")
    assert lock_parity["status"] == "pass" and lock_parity["mismatches"] == []
    assert sha(ROOT / "Cargo.lock") == lock_parity["root_lock_sha256"]
    assert sha(P / "probe-src/Cargo.lock") == lock_parity["probe_lock_sha256"]
    architecture = read(P.parent / "change-0829/architecture-inputs.json")
    assert all(sha(ROOT / n) == h for n, h in architecture.items())
    assert 12 in os.sched_getaffinity(0)
    environment()
    plan = {"schema": "litchi.performance.0830.plan.v1", "base": BASE,
            "owner": OWNER, "cpu": 12, "input": str(INPUT), "reference": str(REFERENCE),
            "scope": "Public one-cell edit including old Workbook replacement/drop; fresh open, save and readback outside timer.",
            "qualification": {"arms": ["direct", "wrapped", "fp"], "samples": 3, "warmup": 0},
            "native": {"blocks": 6, "samples": 30, "warmup": 3,
                       "order": list(itertools.permutations(["direct", "wrapped", "fp"]))},
            "heaptrack": {"repeats": 2, "samples": 5, "warmup": 0, "arm": "fp",
                          "requested_bytes_only": True, "profiled_latency_claim": False},
            "expected_reports": 23, "expected_measured_samples": 559,
            "statistics": {"within_process": "nearest-rank", "across_process": "midpoint median",
                           "bootstrap_resamples": 10000, "seed": 830830, "endpoints": [250, 9749]},
            "quality_scope": "Fresh isolated probe fmt/check/test/clippy/doc; no claim of fresh full production test suite.",
            "limitations": "Single real workbook; heaptrack malloc interposition differs from Rust allocator counters. No per-edit peak/live or physical-copy claim."}
    write(P / "plan.json", plan)
    host = {"platform": platform.platform(), "affinity": sorted(os.sched_getaffinity(0)),
            "cpuinfo": Path("/proc/cpuinfo").read_text(), "meminfo": Path("/proc/meminfo").read_text(),
            "loadavg": Path("/proc/loadavg").read_text(), "cgroup": Path("/proc/self/cgroup").read_text(),
            "filesystem": output(["findmnt", "-n", "-o", "SOURCE,FSTYPE,OPTIONS", "-T", str(ROOT)]),
            "tools": {" ".join(a): output(a) for a in (["rustc", "-Vv"], ["cargo", "-V"],
                ["heaptrack", "--version"], ["heaptrack_print", "--version"], ["python3", "--version"],
                ["zstd", "--version"], ["nm", "--version"] )}}
    current = Path("/sys/fs/cgroup") / host["cgroup"].split("0::", 1)[1].strip().lstrip("/")
    limits = []
    while current.is_relative_to(Path("/sys/fs/cgroup")):
        limits.append({"path": str(current), **{name: (current / name).read_text().strip()
                       if (current / name).is_file() else None
                       for name in ("cpu.max", "memory.max", "cpuset.cpus.effective")}})
        if current == Path("/sys/fs/cgroup"):
            break
        current = current.parent
    host["cgroup_limits"] = limits
    write(P / "host.json", host)
    write(P / "unrelated.json", UNRELATED)
    write(P / "inputs.json", inventory())
    write(P / "prepare.json", {"status": "pass", "base": BASE, "normative_files": len(architecture),
          "inputs_sha256": sha(P / "inputs.json"), "plan_sha256": sha(P / "plan.json"),
          "host_sha256": sha(P / "host.json")})


def quality():
    commands = [
        ("fmt", ["cargo", "fmt", "--manifest-path", MANIFEST, "--", "--check"]),
        ("check", ["cargo", "check", "--offline", "--locked", "--manifest-path", MANIFEST, "--all-targets"]),
        ("test", ["cargo", "test", "--offline", "--locked", "--manifest-path", MANIFEST]),
        ("clippy", ["cargo", "clippy", "--offline", "--locked", "--manifest-path", MANIFEST, "--all-targets", "--", "-D", "warnings"]),
        ("doc", ["cargo", "doc", "--offline", "--locked", "--manifest-path", MANIFEST, "--no-deps"]),
    ]
    for name, argv in commands:
        run("quality-" + name, argv)


def binary(arm):
    return TARGET / ("fp" if arm == "fp" else "ordinary") / "release/xlsx-edit-profile-0830"


def build():
    for arm in ("ordinary", "fp"):
        run("build-" + arm, ["cargo", "build", "--release", "--offline", "--locked",
                             "--manifest-path", MANIFEST], arm)
        write(P / ("binary-" + arm + ".json"), {"path": str(binary(arm)),
              "bytes": binary(arm).stat().st_size, "sha256": sha(binary(arm))})
        run("nm-" + arm, ["nm", "-S", "--defined-only", "--demangle=rust", binary(arm)], arm)
        matches = [line for line in (P / "commands" / ("nm-" + arm + ".log")).read_text().splitlines()
                   if line.split(maxsplit=3)[-1] == OWNER]
        assert len(matches) == 1, matches
        addr, size, kind, name = matches[0].split(maxsplit=3)
        assert kind in ("t", "T") and int(size, 16) > 0
        run("objdump-" + arm, ["objdump", "-d", "--demangle=rust", "--start-address=" + str(int(addr, 16)),
                               "--stop-address=" + str(int(addr, 16) + int(size, 16)), binary(arm)], arm)
        assembly = (P / "commands" / ("objdump-" + arm + ".log")).read_text()
        assert OWNER in assembly and "call" in assembly
        if arm == "fp":
            assert "%rbp" in assembly and "%rsp,%rbp" in assembly
        write(P / ("symbol-" + arm + ".json"), {"row": matches[0], "address": int(addr, 16),
              "size": int(size, 16), "binary_sha256": sha(binary(arm)), "status": "pass"})


def probe_args(arm, destination, samples, warmup):
    return [str(binary(arm)), "--mode", "direct" if arm == "direct" else "wrapped",
            "--input", str(INPUT), "--reference", str(REFERENCE), "--output", str(destination),
            "--samples", str(samples), "--warmup", str(warmup)]


def verify_report(path, count, warmup, arm):
    report = read(path)
    assert report["base_revision"] == BASE
    assert report["all_verified"] is True and report["warmup_verified"] is True
    assert report["warmup"] == warmup and report["samples_requested"] == count
    assert report["mode"] == ("direct" if arm == "direct" else "wrapped")
    samples = report["samples"]
    assert len(samples) == count
    assert report["elapsed_ns"]["sample_order"] == list(range(count))
    assert report["elapsed_ns"]["samples"] == [s["elapsed_ns"] for s in samples]
    assert all(s["index"] == i and isinstance(s["elapsed_ns"], int) and s["elapsed_ns"] > 0
               and all(v is True for v in s["verification"].values()) for i, s in enumerate(samples))
    assert report["output"] == {"bytes": 8521, "sha256": sha(REFERENCE)}


def capture():
    plan = read(P / "plan.json")
    for arm in ("ordinary", "fp"):
        assert sha(binary(arm)) == read(P / ("binary-" + arm + ".json"))["sha256"]
    for arm in ("direct", "wrapped", "fp"):
        dest = P / "qualification" / (arm + ".json")
        dest.parent.mkdir(exist_ok=True)
        run("qualification-" + arm, ["taskset", "-c", "12", *probe_args(arm, dest, 3, 0)],
            "fp" if arm == "fp" else "ordinary")
        verify_report(dest, 3, 0, arm)
    for block, arms in enumerate(plan["native"]["order"]):
        for arm in arms:
            name = f"native-{block:02}-{arm}"
            dest = P / "native" / (name + ".json")
            dest.parent.mkdir(exist_ok=True)
            rss = dest.with_suffix(".rss")
            run(name, ["taskset", "-c", "12", "/usr/bin/time", "-f", "%M", "-o", rss,
                       *probe_args(arm, dest, 30, 3)], "fp" if arm == "fp" else "ordinary")
            verify_report(dest, 30, 3, arm)
    for repeat in range(2):
        name = f"heaptrack-{repeat}"
        folder = P / name
        folder.mkdir()
        dest = folder / "report.json"
        run(name, ["taskset", "-c", "12", "heaptrack", "--record-only", "-o", folder / "trace",
                   *probe_args("fp", dest, 5, 0)], "fp")
        verify_report(dest, 5, 0, "fp")
        assert (folder / "trace.zst").is_file()
        # A separate retained decoded stream permits pure offline readers.
        run(name + "-decode", ["zstd", "-d", "--keep", "-o", folder / "trace.txt", folder / "trace.zst"], "fp")
        for kind, extra in (("whole", []), ("owner", ["--filter-bt-function", OWNER])):
            run(name + "-print-" + kind, ["heaptrack_print", "-f", folder / "trace.zst",
                "--disable-embedded-suppressions", "--disable-builtin-suppressions", "-n", "20", *extra], "fp")


def main():
    stage = sys.argv[1]
    assert stage in ("prepare", "quality", "build", "capture")
    if stage == "prepare":
        prepare()
    else:
        check_inputs()
        if stage == "build":
            assert read(P / "quality.json")["status"] == "pass"
        if stage == "capture":
            assert read(P / "build.json")["status"] == "pass"
        globals()[stage]()
        check_inputs()
        write(P / (stage + ".json"), {"status": "pass", "inputs_sha256": sha(P / "inputs.json"),
                                     "finished_unix": time.time()})
    print(stage + " PASS", flush=True)


if __name__ == "__main__":
    main()
