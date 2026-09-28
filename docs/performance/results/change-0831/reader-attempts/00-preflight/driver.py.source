"""Root-only matched XLSX empty-column-action experiment, with durable receipts."""
from __future__ import annotations

import hashlib
import json
import os
from pathlib import Path
import platform
import shutil
import subprocess
import sys
import time
import tomllib

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BASE = "3bcaee6f418a78ee62bab76aed0d6d9f1f93d024"
TARGET = ROOT.parent / "litchi-target-0831"
SCRATCH = ROOT.parent / "litchi-fs-0831"
SOURCE = "crates/litchi-xlsx/src/raw/worksheet/edit/validation.rs"
TOOL = ROOT / "tools/perf-baseline"
INPUT = ROOT / "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx"
REFERENCE = ROOT / "docs/performance/results/change-0821/artifacts/real-001-xlsx/default.xlsx"
ORACLE_MANIFEST = ROOT / "docs/performance/results/change-0830/probe-src/Cargo.toml"
UNRELATED = json.loads((P.parent / "change-0830/unrelated.json").read_text())


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
    names = output(["git", "ls-files", "crates", "Cargo.toml", "Cargo.lock", "rustfmt.toml",
                    "clippy.toml", ".cargo", "rust-toolchain.toml", "docs/adr", "docs/GOAL.md",
                    "docs/CRUD_Scenario_Checklist.md", "tools"]).splitlines()
    rows = {n: sha(ROOT / n) for n in names}
    for path in (INPUT, REFERENCE, P / "driver.py", *sorted((P / "candidate").glob("*.rs")),
                 *sorted(ORACLE_MANIFEST.parent.rglob("*"))):
        if path.is_file():
            rows[str(path.relative_to(ROOT))] = sha(path)
    return rows


def check(leg):
    assert leg in ("before", "after")
    assert output(["git", "rev-parse", "HEAD"]) == BASE
    frozen = read(P / "inputs.json")
    expected = dict(frozen)
    expected[SOURCE] = sha(P / "candidate" / (leg + "-validation.rs"))
    assert inventory() == expected, "source, driver, normative, tool, oracle or corpus drift"
    assert all(sha(ROOT / n) == h for n, h in UNRELATED.items()), "unrelated work drift"
    freeze = read(P / "prepare.json")
    assert sha(P / "plan.json") == freeze["plan_sha256"]
    assert sha(P / "inputs.json") == freeze["inputs_sha256"]
    assert sha(P / "host.json") == freeze["host_sha256"]


def env(kind):
    assert kind in ("quality", "release", "capture")
    forbidden = [k for k in os.environ if k.startswith(("CARGO_PROFILE_", "CARGO_TARGET_")) or
                 k in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_WRAPPER", "RUSTC_WORKSPACE_WRAPPER", "LD_PRELOAD")]
    assert not forbidden, forbidden
    result = dict(os.environ)
    result.update(CARGO_TARGET_DIR=str(TARGET / ("quality" if kind == "quality" else "release")),
                  CARGO_BUILD_JOBS="2", CARGO_INCREMENTAL="0", CARGO_PROFILE_DEV_DEBUG="0",
                  RUSTDOCFLAGS="-D warnings", PYTHONDONTWRITEBYTECODE="1", LC_ALL="C", TZ="UTC")
    if kind == "release":
        result.update(CARGO_PROFILE_RELEASE_OPT_LEVEL="3", CARGO_PROFILE_RELEASE_DEBUG="1",
                      CARGO_PROFILE_RELEASE_LTO="thin", CARGO_PROFILE_RELEASE_CODEGEN_UNITS="1",
                      CARGO_PROFILE_RELEASE_PANIC="unwind", CARGO_PROFILE_RELEASE_INCREMENTAL="false")
    return result


def run(name, argv, kind, leg):
    check(leg)
    environment = env(kind)
    row = {"argv": list(map(str, argv)), "cwd": str(ROOT), "source_leg": leg,
           "source_sha256": sha(ROOT / SOURCE), "started_unix": time.time(),
           "input_inventory_sha256": sha(P / "inputs.json"),
           "environment": {k: v for k, v in environment.items()
                           if k.startswith(("CARGO_PROFILE_", "CARGO_TARGET_", "RUSTDOC")) or
                           k in ("CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "LC_ALL", "TZ", "PYTHONDONTWRITEBYTECODE")}}
    write(P / "commands" / (name + ".started.json"), row)
    log = P / "commands" / (name + ".log")
    with log.open("xb") as f:
        child = subprocess.run(row["argv"], cwd=ROOT, env=environment, stdout=f, stderr=subprocess.STDOUT)
    row.update(exit_code=child.returncode, finished_unix=time.time(), log_sha256=sha(log))
    write(P / "commands" / (name + ".json"), row)
    print(name, child.returncode, flush=True)
    assert child.returncode == 0, name
    check(leg)


def prepare():
    assert output(["git", "rev-parse", "HEAD"]) == BASE
    assert not TARGET.exists() and not SCRATCH.exists()
    assert (ROOT / SOURCE).read_bytes() == (P / "candidate/before-validation.rs").read_bytes()
    before = (P / "candidate/before-validation.rs").read_text()
    after = (P / "candidate/after-validation.rs").read_text()
    needle = "    let mut owners = Assignments::new()?;"
    assert before.count(needle) == 1
    assert after == before.replace(needle, "    if actions.is_empty() {\n        return Ok(());\n    }\n" + needle)
    assert all(sha(ROOT / n) == h for n, h in read(P / "architecture-inputs.json").items())
    assert all(sha(ROOT / n) == h for n, h in UNRELATED.items())
    assert sha(INPUT) == "d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4"
    assert sha(REFERENCE) == "0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68"
    cases = [{"id": "real-" + phase, "case": "xlsx_real_file_ordinary_save_" + phase,
              "shape": None, "native_samples": 500} for phase in ("edit", "lifecycle")]
    for operation in ("one-cell", "one-percent"):
        for shape in ("tiny", "medium", "dense-wide"):
            cases.append({"id": operation + "-" + shape, "case": "xlsx_" + operation.replace("-", "_") + "_commit_save",
                          "shape": shape, "native_samples": 30})
    cases.append({"id": "noop-medium", "case": "xlsx_noop_commit_save", "shape": "medium", "native_samples": 500})
    plan = {"schema": "litchi.performance.0831.plan.v1", "base": BASE, "source": SOURCE,
            "cpu": 12, "cases": cases, "qualification": {"samples": 1, "warmup": 0, "reports": 36},
            "native": {"blocks": 6, "warmup": 3, "orders": [["before", "after"], ["after", "before"]] * 3},
            "observer": {"blocks": 2, "samples": 3, "warmup": 0, "orders": [["before", "after"], ["after", "before"]]},
            "expected_reports": 180, "expected_samples": 20304,
            "statistics": {"within_process": "nearest-rank", "across_process": "midpoint median",
                           "bootstrap_resamples": 10000, "seed": 831831, "endpoints": [250, 9749]},
            "adoption": {"primary": "at least 15% reduction in real-edit requested allocated bytes with identical output",
                         "regression_flags": "paired native p50 lower CI above 1.05; paired RSS above 1.05; any allocation/region-peak increase",
                         "review": "Review every flag; no pooling observer latency or hiding individual cases. Synthetic 30-sample cases are diagnostic guards."},
            "quality": "Fresh XLSX fmt/check/all-feature tests/clippy/doc/boundaries, pinned 0830 oracle tests and harness library tests on both source legs. Root and standalone harness locks differ but each is fixed across legs."}
    assert 12 in os.sched_getaffinity(0)
    env("quality")
    def package_key(row):
        return tuple(row.get(k) for k in ("name", "version", "source", "checksum"))
    root_packages = {package_key(v) for v in tomllib.loads((ROOT / "Cargo.lock").read_text())["package"]}
    tool_packages = tomllib.loads((TOOL / "Cargo.lock").read_text())["package"]
    missing = [package_key(v) for v in tool_packages if "source" in v and package_key(v) not in root_packages]
    write(P / "lock-parity.json", {"status": "comparison-complete", "mismatches": missing,
          "all_external_tool_entries_match_root": not missing,
          "root_sha256": sha(ROOT / "Cargo.lock"), "tool_sha256": sha(TOOL / "Cargo.lock")})
    write(P / "plan.json", plan)
    write(P / "inputs.json", inventory())
    host = {"platform": platform.platform(), "affinity": sorted(os.sched_getaffinity(0)),
            **{n: Path("/proc", n).read_text() for n in ("cpuinfo", "meminfo", "loadavg")},
            "tools": {" ".join(a): output(a) for a in (["rustc", "-Vv"], ["cargo", "-V"], ["python3", "--version"], ["git", "--version"])},
            "filesystem": output(["findmnt", "-n", "-o", "SOURCE,FSTYPE,OPTIONS", "-T", str(ROOT)])}
    host["cgroup"] = Path("/proc/self/cgroup").read_text()
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
    write(P / "prepare.json", {"status": "pass", "base": BASE, "inputs_sha256": sha(P / "inputs.json"),
                                "plan_sha256": sha(P / "plan.json"), "host_sha256": sha(P / "host.json")})
    SCRATCH.mkdir()
    (SCRATCH / ".owner").write_text("litchi-performance-0831-owned-scratch\n")


def install():
    check("before")
    assert read(P / "build-before.json")["status"] == "pass"
    (ROOT / SOURCE).write_bytes((P / "candidate/after-validation.rs").read_bytes())
    check("after")
    write(P / "install.json", {"status": "pass", "source": SOURCE, "before": sha(P / "candidate/before-validation.rs"),
                               "after": sha(ROOT / SOURCE)})


def quality(leg):
    prefix = ["--offline", "--locked", "-p", "litchi-xlsx"]
    commands = [
        ("fmt", ["cargo", "fmt", "-p", "litchi-xlsx", "--", "--check"]),
        ("check", ["cargo", "check", *prefix, "--all-features", "--all-targets"]),
        ("test", ["cargo", "test", *prefix, "--all-features", "--", "--test-threads=2"]),
        ("clippy", ["cargo", "clippy", *prefix, "--all-features", "--all-targets", "--", "-D", "warnings"]),
        ("doc", ["cargo", "doc", *prefix, "--all-features", "--no-deps"]),
        ("boundaries", ["python3", "-B", ROOT / "tools/check_crate_boundaries.py"]),
        ("oracle", ["cargo", "test", "--offline", "--locked", "--manifest-path", ORACLE_MANIFEST]),
        ("harness-test", ["cargo", "test", "--offline", "--locked", "--manifest-path", TOOL / "Cargo.toml",
                          "--lib", "--features", "allocator-metrics,ordinary-save-process-metrics", "--", "--test-threads=2"]),
    ]
    for name, argv in commands:
        run(f"quality-{leg}-{name}", argv, "quality", leg)
    write(P / ("quality-" + leg + ".json"), {"status": "pass", "gates": len(commands),
          "source_sha256": sha(ROOT / SOURCE), "inputs_sha256": sha(P / "inputs.json")})


def build(leg):
    assert read(P / ("quality-" + leg + ".json"))["status"] == "pass"
    binaries = {}
    for logical, name, features in (("native", "litchi-perf-baseline", []),
            ("observer", "litchi-perf-baseline-alloc", ["--features", "allocator-metrics,ordinary-save-process-metrics"])):
        run(f"build-{leg}-{logical}", ["cargo", "build", "--offline", "--locked", "--release",
            "--manifest-path", TOOL / "Cargo.toml", "--bin", name, *features], "release", leg)
        source = TARGET / "release/release" / name
        dest = TARGET / (leg + "-" + logical)
        assert not dest.exists()
        shutil.copy2(source, dest)
        binaries[logical] = {"path": str(dest), "bytes": dest.stat().st_size, "sha256": sha(dest)}
    write(P / ("build-" + leg + ".json"), {"status": "pass", "binaries": binaries,
            "source_sha256": sha(ROOT / SOURCE), "inputs_sha256": sha(P / "inputs.json")})


def run_case(lane, block, leg, logical, case, samples, warmup, source_leg):
    binding = read(P / ("build-" + leg + ".json"))["binaries"][logical]
    assert sha(binding["path"]) == binding["sha256"]
    stem = (leg + "-" + logical + "-" + case["id"]) if lane == "qualification" else f"{block:02}-{leg}-{case['id']}"
    dest = P / lane / (stem + ".json")
    dest.parent.mkdir(exist_ok=True)
    argv = ["taskset", "-c", "12", "/usr/bin/time", "-f", "%M", "-o", dest.with_suffix(".rss"),
            binding["path"], "--case", case["case"], "--warmup", str(warmup), "--samples", str(samples),
            "--json", dest, "--filesystem-root", SCRATCH]
    if case["shape"] is None:
        argv += ["--ooxml-file", INPUT]
    else:
        argv += ["--xlsx-shape", case["shape"]]
    run(lane + "-" + stem, argv, "capture", source_leg)
    report = read(dest)
    assert report["configuration"]["samples_per_case"] == samples
    assert len(report["results"]) == 1
    result = report["results"][0]
    assert result["case"] == case["case"] and len(result["elapsed_ns"]["samples"]) == samples
    assert report["binary_identity"]["binary_sha256"] == binding["sha256"]
    if case["shape"] is None:
        assert result["corpus"]["archive_sha256"] == sha(INPUT)
    assert report["tool"]["instrumentation"] == ("none" if logical == "native" else "system_allocator_operation_scoped")


def qualification(leg):
    plan = read(P / "plan.json")
    for logical in ("native", "observer"):
        for case in plan["cases"]:
            run_case("qualification", 0, leg, logical, case, 1, 0, leg)
    write(P / ("qualification-" + leg + ".json"), {"status": "pass", "reports": 18, "samples": 18})


def capture():
    assert all(read(P / ("qualification-" + leg + ".json"))["status"] == "pass" for leg in ("before", "after"))
    plan = read(P / "plan.json")
    for lane in ("native", "observer"):
        for block, order in enumerate(plan[lane]["orders"]):
            for case in plan["cases"]:
                for leg in order:
                    count = case["native_samples"] if lane == "native" else 3
                    run_case(lane, block, leg, lane, case, count, plan[lane]["warmup"], "after")
    write(P / "capture.json", {"status": "pass", "reports": 180, "samples": 20304,
                              "inputs_sha256": sha(P / "inputs.json")})


def main():
    stage = sys.argv[1]
    if stage == "prepare":
        prepare()
    elif stage in ("before", "after"):
        check(stage)
        quality(stage)
        build(stage)
        qualification(stage)
    elif stage == "install":
        install()
    elif stage == "capture":
        check("after")
        capture()
    else:
        raise ValueError(stage)
    print(stage + " PASS", flush=True)


if __name__ == "__main__":
    main()
