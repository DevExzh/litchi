"""Root-only matched XLSX column-assignment experiment.

This driver owns compilation and workload capture in this packet.  It freezes
the complete source census before any build, keeps the candidate sources under
their repository paths, and refuses to overwrite any receipt.
The current revision is the already-tested 0831 after leg, so that leg's
quality evidence is reused only after its command receipts, logs, and
normalized input census have been verified.  The candidate after leg receives
fresh quality, build, qualification, and comparative capture work.
"""
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
BASE = "7eeaab48c0281b06f53527d4c4f4ea79050d27e7"
TARGET = ROOT.parent / "litchi-target-0832"
SCRATCH = ROOT.parent / "litchi-fs-0832"
TOOL = ROOT / "tools/perf-baseline"
INPUT = ROOT / "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx"
REFERENCE = ROOT / "docs/performance/results/change-0821/artifacts/real-001-xlsx/default.xlsx"
ORACLE_MANIFEST = ROOT / "docs/performance/results/change-0830/probe-src/Cargo.toml"
PREVIOUS = P.parent / "change-0831"
UNRELATED = json.loads((P / "unrelated.json").read_text())
ARCHITECTURE_INPUTS = P / "architecture-inputs.json"
CENSUS_SCRIPT = P / "census.py"
CENSUS_REPORT = P / "census.json"
DESIGN = P / "design.md"
SOURCE_REVIEW = P / "source-review.md"

PRIMARY_SOURCE = "crates/litchi-xlsx/src/raw/worksheet/edit/validation.rs"
SOURCE = PRIMARY_SOURCE
SOURCES = (
    "crates/litchi-xlsx/src/column.rs",
    "crates/litchi-xlsx/src/raw/worksheet/codec.rs",
    "crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/columns.rs",
    PRIMARY_SOURCE,
)
OBSERVER_IDENTITY = "ordinary_save_procfs_and_system_allocator_operation_scoped"
NATIVE_IDENTITY = "none"


def sha(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def read(path):
    return json.loads(Path(path).read_text())


def write(path, value):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x") as stream:
        stream.write(json.dumps(value, indent=2, sort_keys=True) + "\n")


def output(argv):
    return subprocess.check_output([str(part) for part in argv], cwd=ROOT, text=True).strip()


def candidate_path(leg, source):
    assert leg in ("before", "after")
    path = P / "candidate" / leg / source
    assert path.is_file() and not path.is_symlink(), path
    return path


def archive_source_map(leg):
    return {source: sha(candidate_path(leg, source)) for source in SOURCES}


def packet_source_map(leg):
    # Readers bind receipts by the full repository source paths.  The archived
    # candidate paths remain independently hashed in ``inputs.json`` and in
    # the candidate manifest.
    return archive_source_map(leg)


def root_source_map():
    return {source: sha(ROOT / source) for source in SOURCES}


def tracked_names():
    return output([
        "git", "ls-files", "crates", "Cargo.toml", "Cargo.lock", "rustfmt.toml",
        "clippy.toml", ".cargo", "rust-toolchain.toml", "docs/adr", "docs/GOAL.md",
        "docs/CRUD_Scenario_Checklist.md", "tools",
    ]).splitlines()


def root_inventory():
    """Hash all repository inputs used by the build, test, and harness commands."""
    rows = {name: sha(ROOT / name) for name in tracked_names()}
    for path in (INPUT, REFERENCE, *sorted(ORACLE_MANIFEST.parent.rglob("*"))):
        if path.is_file():
            rows[str(path.relative_to(ROOT))] = sha(path)
    return rows


def packet_input_paths():
    paths = [ARCHITECTURE_INPUTS, P / "unrelated.json", CENSUS_SCRIPT, CENSUS_REPORT]
    for path in (DESIGN, SOURCE_REVIEW):
        if path.is_file():
            paths.append(path)
    return paths


def inventory():
    rows = root_inventory()
    rows[str(P.joinpath("driver.py").relative_to(ROOT))] = sha(P / "driver.py")
    for path in packet_input_paths():
        assert path.is_file() and not path.is_symlink(), path
        rows[str(path.relative_to(ROOT))] = sha(path)
    for path in sorted((P / "candidate").rglob("*")):
        if path.is_file():
            assert not path.is_symlink(), path
            rows[str(path.relative_to(ROOT))] = sha(path)
    return rows


def assert_candidate_archive():
    expected = {Path(leg) / source for leg in ("before", "after") for source in SOURCES}
    expected |= {Path(name) for name in ("manifest.json", "candidate.patch", "pre-freeze-correction.json")}
    actual = {
        path.relative_to(P / "candidate")
        for path in (P / "candidate").rglob("*")
        if path.is_file()
    }
    assert actual == expected, {"missing": sorted(expected - actual), "extra": sorted(actual - expected)}
    for leg in ("before", "after"):
        for source in SOURCES:
            candidate_path(leg, source)


def expected_inventory(leg):
    frozen = read(P / "inputs.json")
    expected = dict(frozen)
    for source in SOURCES:
        expected[source] = sha(candidate_path(leg, source))
    return expected


def check(leg):
    assert leg in ("before", "after")
    assert output(["git", "rev-parse", "HEAD"]) == BASE
    assert inventory() == expected_inventory(leg), "source, driver, normative, tool, oracle or corpus drift"
    assert root_source_map() == archive_source_map(leg), "production source does not match the archived leg"
    assert all(sha(ROOT / name) == digest for name, digest in UNRELATED.items()), "unrelated work drift"
    freeze = read(P / "prepare.json")
    assert sha(P / "plan.json") == freeze["plan_sha256"]
    assert sha(P / "inputs.json") == freeze["inputs_sha256"]
    assert sha(P / "host.json") == freeze["host_sha256"]
    assert sha(P / "baseline-witness.json") == freeze["baseline_witness_sha256"]


def env(kind):
    assert kind in ("quality", "release", "capture")
    forbidden = [
        key for key in os.environ
        if key.startswith(("CARGO_PROFILE_", "CARGO_TARGET_"))
        or key in (
            "RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_WRAPPER",
            "RUSTC_WORKSPACE_WRAPPER", "LD_PRELOAD",
        )
    ]
    assert not forbidden, forbidden
    result = dict(os.environ)
    result.update(
        CARGO_TARGET_DIR=str(TARGET / ("quality" if kind == "quality" else "release")),
        CARGO_BUILD_JOBS="2",
        CARGO_INCREMENTAL="0",
        CARGO_PROFILE_DEV_DEBUG="0",
        RUSTDOCFLAGS="-D warnings",
        PYTHONDONTWRITEBYTECODE="1",
        LC_ALL="C",
        TZ="UTC",
    )
    if kind == "release":
        result.update(
            CARGO_PROFILE_RELEASE_OPT_LEVEL="3",
            CARGO_PROFILE_RELEASE_DEBUG="1",
            CARGO_PROFILE_RELEASE_LTO="thin",
            CARGO_PROFILE_RELEASE_CODEGEN_UNITS="1",
            CARGO_PROFILE_RELEASE_PANIC="unwind",
            CARGO_PROFILE_RELEASE_INCREMENTAL="false",
        )
    return result


def run(name, argv, kind, source_leg, binary_leg=None):
    """Run one immutable child command and retain a terminal receipt on failure."""
    check(source_leg)
    environment = env(kind)
    normalized = list(map(str, argv))
    row = {
        "argv": normalized,
        "cwd": str(ROOT),
        "source_leg": source_leg,
        "source_sha256": sha(ROOT / PRIMARY_SOURCE),
        "source_map": root_source_map(),
        "sources_sha256": packet_source_map(source_leg),
        "binary_leg": binary_leg,
        "started_unix": time.time(),
        "input_inventory_sha256": sha(P / "inputs.json"),
        "environment": {
            key: value for key, value in environment.items()
            if key.startswith(("CARGO_PROFILE_", "CARGO_TARGET_", "RUSTDOC"))
            or key in (
                "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "LC_ALL", "TZ",
                "PYTHONDONTWRITEBYTECODE",
            )
        },
    }
    started = P / "commands" / (name + ".started.json")
    terminal = P / "commands" / (name + ".json")
    log = P / "commands" / (name + ".log")
    assert all(not path.exists() and not path.is_symlink() for path in (started, terminal, log))
    write(started, row)
    try:
        with log.open("xb") as stream:
            child = subprocess.run(normalized, cwd=ROOT, env=environment,
                                   stdout=stream, stderr=subprocess.STDOUT)
        row.update(exit_code=child.returncode, finished_unix=time.time(), log_sha256=sha(log))
    except BaseException as error:
        row.update(
            exit_code=None,
            finished_unix=time.time(),
            log_sha256=sha(log) if log.exists() else None,
            exception_type=type(error).__name__,
            exception=str(error),
        )
        write(terminal, row)
        raise
    write(terminal, row)
    print(name, row["exit_code"], flush=True)
    assert row["exit_code"] == 0, name
    check(source_leg)


def verify_previous_quality():
    """Verify the completed 0831 after quality as a reusable baseline witness."""
    assert PREVIOUS.is_dir()
    previous_inputs = read(PREVIOUS / "inputs.json")
    normalized = {
        name: digest for name, digest in previous_inputs.items()
        if not name.startswith("docs/performance/results/change-0831/")
    }
    normalized[PRIMARY_SOURCE] = sha(PREVIOUS / "candidate/after-validation.rs")
    assert root_inventory() == normalized, "current root is not the normalized 0831 after census"
    previous_quality = read(PREVIOUS / "quality-after.json")
    assert previous_quality["status"] == "pass" and previous_quality["gates"] == 8
    assert previous_quality["source_sha256"] == sha(PREVIOUS / "candidate/after-validation.rs")
    assert previous_quality["inputs_sha256"] == sha(PREVIOUS / "inputs.json")
    command_rows = {}
    expected_names = ("fmt", "check", "test", "clippy", "doc", "boundaries", "oracle", "harness-test")
    for name in expected_names:
        terminal = PREVIOUS / "commands" / ("quality-after-" + name + ".json")
        started = PREVIOUS / "commands" / ("quality-after-" + name + ".started.json")
        log = PREVIOUS / "commands" / ("quality-after-" + name + ".log")
        receipt = read(terminal)
        start = read(started)
        assert receipt["exit_code"] == 0
        assert all(receipt[key] == value for key, value in start.items())
        assert receipt["source_leg"] == "after"
        assert receipt["source_sha256"] == previous_quality["source_sha256"]
        assert receipt["input_inventory_sha256"] == previous_quality["inputs_sha256"]
        assert receipt["log_sha256"] == sha(log)
        command_rows[name] = {
            "receipt": str(terminal.relative_to(ROOT)),
            "receipt_sha256": sha(terminal),
            "started": str(started.relative_to(ROOT)),
            "log": str(log.relative_to(ROOT)),
            "log_sha256": sha(log),
        }
    return {
        "status": "verified-reusable",
        "packet": "docs/performance/results/change-0831",
        "base": BASE,
        "quality": str((PREVIOUS / "quality-after.json").relative_to(ROOT)),
        "quality_sha256": sha(PREVIOUS / "quality-after.json"),
        "previous_inputs_sha256": sha(PREVIOUS / "inputs.json"),
        "normalized_input_rule": "0831 packet paths omitted; sole validation source replaced with 0831 after archive",
        "commands": command_rows,
    }


def prepare():
    assert output(["git", "rev-parse", "HEAD"]) == BASE
    assert not TARGET.exists() and not SCRATCH.exists()
    assert_candidate_archive()
    assert root_source_map() == archive_source_map("before")
    assert all(sha(ROOT / name) == digest for name, digest in read(ARCHITECTURE_INPUTS).items())
    assert all(sha(ROOT / name) == digest for name, digest in UNRELATED.items())
    assert sha(INPUT) == "d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4"
    assert sha(REFERENCE) == "0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68"
    baseline = verify_previous_quality()
    cases = [
        {"id": "real-" + phase, "case": "xlsx_real_file_ordinary_save_" + phase,
         "shape": None, "native_samples": 500}
        for phase in ("edit", "lifecycle")
    ]
    for operation in ("one-cell", "one-percent"):
        for shape in ("tiny", "medium", "dense-wide"):
            cases.append({
                "id": operation + "-" + shape,
                "case": "xlsx_" + operation.replace("-", "_") + "_commit_save",
                "shape": shape,
                "native_samples": 30,
            })
    cases.append({
        "id": "noop-medium", "case": "xlsx_noop_commit_save", "shape": "medium",
        "native_samples": 500,
    })
    plan = {
        "schema": "litchi.performance.0832.plan.v1",
        "base": BASE,
        "sources": list(SOURCES),
        "primary_source": PRIMARY_SOURCE,
        "source": PRIMARY_SOURCE,
        "cpu": 12,
        "cases": cases,
        "qualification": {"samples": 1, "warmup": 0, "reports": 36},
        "qualification_admission_required_before_capture": True,
        "native": {
            "blocks": 6, "warmup": 3,
            "orders": [["before", "after"], ["after", "before"]] * 3,
        },
        "observer": {
            "blocks": 2, "samples": 3, "warmup": 0,
            "orders": [["before", "after"], ["after", "before"]],
            "identity": OBSERVER_IDENTITY,
        },
        "expected_reports": 180,
        "expected_samples": 20304,
        "expected_fresh_child_commands": 192,
        "statistics": {
            "within_process": "nearest-rank",
            "across_process": "midpoint median",
            "bootstrap_resamples": 10000,
            "seed": 832832,
            "endpoints": [250, 9749],
        },
        "adoption": {
            "primary": "at least 50% reduction in real-edit requested allocated bytes with identical output",
            "regression_flags": (
                "paired native p50 lower CI above 1.05; paired RSS above 1.05; "
                "any allocation call/bytes/region-peak increase"
            ),
            "peak_review": (
                "Review region peak after subtracting each leg's entry baseline; retain raw peak, "
                "entry, peak-above-entry, and net-live diagnostics."
            ),
            "review": "Review every flag; keep every case visible; observer latency is descriptive only.",
        },
        "quality": {
            "before": "verified reuse of 0831 after quality receipts and normalized full input census",
            "after": "fresh eight-gate XLSX quality run",
            "gates": ["fmt", "check", "test", "clippy", "doc", "boundaries", "oracle", "harness-test"],
            "locks": "Root and standalone harness locks differ and remain fixed independently.",
        },
        "custody": "Every child receipt contains source_leg, binary_leg, primary source hash, and all four source hashes.",
    }
    assert 12 in os.sched_getaffinity(0)
    env("quality")

    def package_key(row):
        return tuple(row.get(key) for key in ("name", "version", "source", "checksum"))

    root_packages = {
        package_key(value)
        for value in tomllib.loads((ROOT / "Cargo.lock").read_text())["package"]
    }
    tool_packages = tomllib.loads((TOOL / "Cargo.lock").read_text())["package"]
    missing = [
        package_key(value) for value in tool_packages
        if "source" in value and package_key(value) not in root_packages
    ]
    write(P / "lock-parity.json", {
        "status": "comparison-complete",
        "mismatches": missing,
        "all_external_tool_entries_match_root": not missing,
        "root_sha256": sha(ROOT / "Cargo.lock"),
        "tool_sha256": sha(TOOL / "Cargo.lock"),
        "policy": "retain root/tool lock differences; no cross-lock equality assertion",
    })
    write(P / "plan.json", plan)
    write(P / "inputs.json", inventory())
    host = {
        "platform": platform.platform(),
        "affinity": sorted(os.sched_getaffinity(0)),
        **{name: Path("/proc", name).read_text() for name in ("cpuinfo", "meminfo", "loadavg")},
        "tools": {
            " ".join(argv): output(argv)
            for argv in (("rustc", "-Vv"), ("cargo", "-V"), ("python3", "--version"), ("git", "--version"))
        },
        "filesystem": output(["findmnt", "-n", "-o", "SOURCE,FSTYPE,OPTIONS", "-T", str(ROOT)]),
    }
    host["cgroup"] = Path("/proc/self/cgroup").read_text()
    cgroup_suffix = host["cgroup"].split("0::", 1)[1].strip().lstrip("/")
    current = Path("/sys/fs/cgroup") / cgroup_suffix
    limits = []
    while current.is_relative_to(Path("/sys/fs/cgroup")):
        limits.append({
            "path": str(current),
            **{
                name: (current / name).read_text().strip() if (current / name).is_file() else None
                for name in ("cpu.max", "memory.max", "cpuset.cpus.effective")
            },
        })
        if current == Path("/sys/fs/cgroup"):
            break
        current = current.parent
    host["cgroup_limits"] = limits
    write(P / "host.json", host)
    write(P / "baseline-witness.json", baseline)
    write(P / "prepare.json", {
        "status": "pass",
        "base": BASE,
        "inputs_sha256": sha(P / "inputs.json"),
        "plan_sha256": sha(P / "plan.json"),
        "host_sha256": sha(P / "host.json"),
        "baseline_witness_sha256": sha(P / "baseline-witness.json"),
    })
    SCRATCH.mkdir()
    (SCRATCH / ".owner").write_text("litchi-performance-0832-owned-scratch\n")


def install():
    check("before")
    assert read(P / "build-before.json")["status"] == "pass"
    replacement = {source: candidate_path("after", source).read_bytes() for source in SOURCES}
    for source, contents in replacement.items():
        destination = ROOT / source
        assert destination.is_file() and not destination.is_symlink()
        destination.write_bytes(contents)
    check("after")
    write(P / "install.json", {
        "status": "pass",
        "source": PRIMARY_SOURCE,
        "sources_sha256": packet_source_map("after"),
        "sources": {
            source: {
                "path": str(candidate_path("after", source).relative_to(P)),
                "bytes": candidate_path("after", source).stat().st_size,
                "sha256": sha(candidate_path("after", source)),
            }
            for source in SOURCES
        },
        "transitions": {
            source: {"before": archive_source_map("before")[source], "after": sha(ROOT / source)}
            for source in SOURCES
        },
    })


def quality(leg):
    if leg == "before":
        witness = read(P / "baseline-witness.json")
        assert witness["status"] == "verified-reusable"
        assert witness["base"] == BASE
        write(P / "quality-before.json", {
            "status": "pass",
            "gates": 8,
            "reused": True,
            "reused_from": witness["packet"],
            "witness_sha256": sha(P / "baseline-witness.json"),
            "source_sha256": sha(ROOT / PRIMARY_SOURCE),
            "source_map": root_source_map(),
            "sources_sha256": packet_source_map("before"),
            "inputs_sha256": sha(P / "inputs.json"),
        })
        return
    assert leg == "after"
    prefix = ["--offline", "--locked", "-p", "litchi-xlsx"]
    commands = [
        ("fmt", ["cargo", "fmt", "-p", "litchi-xlsx", "--", "--check"]),
        ("check", ["cargo", "check", *prefix, "--all-features", "--all-targets"]),
        ("test", ["cargo", "test", *prefix, "--all-features", "--", "--test-threads=2"]),
        ("clippy", ["cargo", "clippy", *prefix, "--all-features", "--all-targets", "--", "-D", "warnings"]),
        ("doc", ["cargo", "doc", *prefix, "--all-features", "--no-deps"]),
        ("boundaries", ["python3", "-B", ROOT / "tools/check_crate_boundaries.py"]),
        ("oracle", ["cargo", "test", "--offline", "--locked", "--manifest-path", ORACLE_MANIFEST]),
        ("harness-test", [
            "cargo", "test", "--offline", "--locked", "--manifest-path", TOOL / "Cargo.toml",
            "--lib", "--features", "allocator-metrics,ordinary-save-process-metrics", "--", "--test-threads=2",
        ]),
    ]
    for name, argv in commands:
        run("quality-after-" + name, argv, "quality", "after")
    write(P / "quality-after.json", {
        "status": "pass",
        "gates": len(commands),
        "reused": False,
        "source_sha256": sha(ROOT / PRIMARY_SOURCE),
        "source_map": root_source_map(),
        "sources_sha256": packet_source_map("after"),
        "inputs_sha256": sha(P / "inputs.json"),
    })


def build(leg):
    assert read(P / ("quality-" + leg + ".json"))["status"] == "pass"
    binaries = {}
    for logical, name, features in (
        ("native", "litchi-perf-baseline", []),
        ("observer", "litchi-perf-baseline-alloc", ["--features", "allocator-metrics,ordinary-save-process-metrics"]),
    ):
        run(
            "build-" + leg + "-" + logical,
            ["cargo", "build", "--offline", "--locked", "--release", "--manifest-path", TOOL / "Cargo.toml",
             "--bin", name, *features],
            "release", leg, leg,
        )
        source = TARGET / "release/release" / name
        destination = TARGET / (leg + "-" + logical)
        assert source.is_file() and not source.is_symlink()
        assert not destination.exists() and not destination.is_symlink()
        shutil.copy2(source, destination)
        binaries[logical] = {
            "path": str(destination),
            "bytes": destination.stat().st_size,
            "sha256": sha(destination),
        }
    write(P / ("build-" + leg + ".json"), {
        "status": "pass",
        "binaries": binaries,
        "source_sha256": sha(ROOT / PRIMARY_SOURCE),
        "source_map": root_source_map(),
        "sources_sha256": packet_source_map(leg),
        "inputs_sha256": sha(P / "inputs.json"),
    })


def run_case(lane, block, leg, logical, case, samples, warmup, source_leg):
    assert source_leg in ("before", "after")
    build_receipt = read(P / ("build-" + leg + ".json"))
    assert build_receipt["status"] == "pass"
    assert build_receipt["source_map"] == archive_source_map(leg)
    binding = build_receipt["binaries"][logical]
    binary = Path(binding["path"])
    assert binary.is_file() and not binary.is_symlink()
    assert sha(binary) == binding["sha256"] and binary.stat().st_size == binding["bytes"]
    stem = (leg + "-" + logical + "-" + case["id"] if lane == "qualification"
            else f"{block:02}-{leg}-{case['id']}")
    destination = P / lane / (stem + ".json")
    rss = destination.with_suffix(".rss")
    lane_dir = destination.parent
    assert not lane_dir.is_symlink()
    lane_dir.mkdir(exist_ok=True)
    command_name = lane + "-" + stem
    terminal = P / "commands" / (command_name + ".json")
    started = P / "commands" / (command_name + ".started.json")
    log = P / "commands" / (command_name + ".log")
    assert all(not path.exists() and not path.is_symlink() for path in (destination, rss, terminal, started, log))
    argv = [
        "taskset", "-c", "12", "/usr/bin/time", "-f", "%M", "-o", rss,
        binary, "--case", case["case"], "--warmup", str(warmup), "--samples", str(samples),
        "--json", destination, "--filesystem-root", SCRATCH,
    ]
    argv += ["--ooxml-file", INPUT] if case["shape"] is None else ["--xlsx-shape", case["shape"]]
    run(command_name, argv, "capture", source_leg, leg)
    receipt = read(terminal)
    start = read(started)
    assert all(receipt[key] == value for key, value in start.items())
    assert receipt["argv"] == list(map(str, argv)) and receipt["cwd"] == str(ROOT)
    assert receipt["exit_code"] == 0 and receipt["finished_unix"] >= receipt["started_unix"]
    assert receipt["source_leg"] == source_leg
    assert receipt["source_map"] == archive_source_map(source_leg)
    assert receipt["sources_sha256"] == packet_source_map(source_leg)
    assert receipt["binary_leg"] == leg
    assert receipt["input_inventory_sha256"] == sha(P / "inputs.json")
    assert receipt["log_sha256"] == sha(log)
    report = read(destination)
    assert report["configuration"]["samples_per_case"] == samples
    assert report["configuration"]["warmup_iterations_per_case"] == warmup
    assert report["configuration"]["cases"] == [case["case"]]
    assert len(report["results"]) == 1
    result = report["results"][0]
    assert result["case"] == case["case"] and len(result["elapsed_ns"]["samples"]) == samples
    assert report["binary_identity"]["binary_sha256"] == binding["sha256"]
    assert report["binary_identity"]["binary_bytes"] == binding["bytes"]
    assert result["corpus"]["shape"] == ("real-file" if case["shape"] is None else case["shape"])
    if case["shape"] is None:
        assert result["corpus"]["archive_sha256"] == sha(INPUT)
    else:
        assert case["shape"] in report["configuration"]["xlsx_shapes"]
    expected_identity = NATIVE_IDENTITY if logical == "native" else OBSERVER_IDENTITY
    assert report["tool"]["instrumentation"] == expected_identity
    rss_text = rss.read_text().strip()
    assert rss_text.isdigit() and int(rss_text) > 0


def qualification(leg):
    plan = read(P / "plan.json")
    for logical in ("native", "observer"):
        for case in plan["cases"]:
            run_case("qualification", 0, leg, logical, case, 1, 0, leg)
    write(P / ("qualification-" + leg + ".json"), {
        "status": "pass", "reports": 18, "samples": 18,
        "source_map": archive_source_map(leg), "sources_sha256": packet_source_map(leg),
        "inputs_sha256": sha(P / "inputs.json"),
    })


def capture():
    assert all(read(P / ("qualification-" + leg + ".json"))["status"] == "pass"
               for leg in ("before", "after"))
    admission_path = P / "qualification-admission.json"
    assert admission_path.is_file() and not admission_path.is_symlink()
    admission = read(admission_path)
    assert admission["status"] == "pass"
    assert admission["reports"] == 36 and admission["samples"] == 36
    assert admission["analysis_reader_sha256"] == sha(P / "analyze.py")
    assert admission["independent_reader_sha256"] == sha(P / "audit.py")
    qualified = admission["qualified_report_hashes"]
    assert isinstance(qualified, dict) and len(qualified) == 36
    for relative_path, digest in qualified.items():
        report = P / relative_path
        assert report.is_file() and not report.is_symlink() and sha(report) == digest
    plan = read(P / "plan.json")
    for lane in ("native", "observer"):
        for block, order in enumerate(plan[lane]["orders"]):
            for case in plan["cases"]:
                for leg in order:
                    samples = case["native_samples"] if lane == "native" else 3
                    run_case(lane, block, leg, lane, case, samples, plan[lane]["warmup"], "after")
    write(P / "capture.json", {
        "status": "pass", "reports": 180, "samples": 20304,
        "measured_reports": 144, "measured_samples": 20268,
        "inputs_sha256": sha(P / "inputs.json"),
        "qualification_admission_sha256": sha(admission_path),
    })


def main():
    assert len(sys.argv) == 2
    stage = sys.argv[1]
    if stage == "prepare":
        prepare()
    elif stage == "before":
        check("before")
        quality("before")
        build("before")
        qualification("before")
    elif stage == "install":
        install()
    elif stage == "after":
        check("after")
        quality("after")
        build("after")
        qualification("after")
    elif stage == "capture":
        check("after")
        capture()
    else:
        raise ValueError(stage)
    print(stage + " PASS", flush=True)


if __name__ == "__main__":
    main()
