#!/usr/bin/env python3
"""Freeze, run, and reduce the 0914 DOCX section-layout baseline.

The benchmark is built from a selective ``git archive`` of the committed
source tree in an external custody directory.  This helper never builds from
the shared dirty checkout and never writes binaries or raw child reports into
the repository.  It uses only the Python standard library.

Actions:

``prepare``
    Create the external source snapshot from ``HEAD``.  It refuses to replace
    an existing custody directory.
``freeze``
    Record the source revision/tree/archive and the run protocol.
``build``
    Build the release harness with ``--locked`` in an external target.
``run``
    Run one or all of four fresh pinned child processes.
``reduce``
    Validate every report and process receipt, then write the small retained
    reduction under this change directory.
``verify``
    Revalidate the retained reduction and every still-present external input.

The timed benchmark operation is the harness's six phase sum.  ``time -v``
process metrics are explicitly setup-inclusive and are never substituted for
the operation timer.  This packet is descriptive baseline evidence and never
calculates a speedup.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import os
from pathlib import Path
import platform
import re
import shutil
import statistics
import subprocess
import sys
import tarfile
import time


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
EXTERNAL = Path(os.environ.get("LITCHI_PERF_0914_ROOT", "/home/zhuhe/litchi-performance-0914")).resolve()
SOURCE = EXTERNAL / "source" / "litchi"
ARCHIVE = EXTERNAL / "source-HEAD-code.tar"
TARGET = EXTERNAL / "target"
BINARY = TARGET / "release" / "litchi-perf-baseline"
BUILD_DIR = EXTERNAL / "build"
RUN_DIR = EXTERNAL / "runs"

CASE = "docx_source_backed_existing_section_layout_edit_save"
WARMUP = 30
SAMPLES = 200
CPU = int(os.environ.get("LITCHI_PERF_0914_CPU", "12"))
DEFAULT_REVISION = "a4ca85e28cdea151176534b7f8dc912253a1ca1e"
ARCHIVE_PATHS = (
    "Cargo.toml",
    "clippy.toml",
    "deny.toml",
    "rust-toolchain.toml",
    "rustfmt.toml",
    ".cargo",
    "crates",
    "test-data/poi/test-data/spreadsheet/54016.xls",
    "test-data/rtf/watermark.rtf",
    "tools/perf-baseline",
)
SOURCE_SCOPE = HERE / "source-scope.md"
EXPECTED = {
    "generator": "litchi-docx-section-layout-source-v1",
    "name": "docx-source-backed-existing-section-layout",
    "package_format": "DOCX/OPC/ZIP",
    "shape": "existing-section-layout",
    "payload_kind": "deterministic-incompressible-media",
    "compression": "deflate",
    "entry_count": 7,
    "archive_member_count": 10,
    "entry_bytes": 524288,
    "uncompressed_payload_bytes": 2116441,
    "archive_bytes": 2101427,
    "archive_sha256": "183204a832a3518bf1d4474e9b7ff347ad7d0dc95ad2e9b1dbdde405837a29e8",
    "target_entry": "word/document.xml",
    "target_payload_bytes": 18855,
    "target_payload_sha256": "67cb9066f112bfee6905a23f3eabdc76b65c86cd0c6613dd60a5a18895ceac81",
    "output_bytes": 2101446,
    "output_sha256": "5e87e9ca9fd6b9a98933c36d9aee1e848bc97eb7693229f7949514830121c4ec",
}
CORPUS_EXPECTED = {key: value for key, value in EXPECTED.items() if not key.startswith("output_")}


class EvidenceError(RuntimeError):
    pass


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def fail(message: str) -> None:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, object]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    require(not path.exists(), f"refusing to replace retained file: {path}")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    temporary.replace(path)


def replace_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(path.name + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    temporary.replace(path)


def read_json(path: Path) -> object:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def validate_source_record(source_record: dict[str, object]) -> None:
    """Verify that the external source custody still matches its freeze receipt."""
    revision = source_record.get("source_revision")
    require(isinstance(revision, str), "source revision missing")
    require(revision == configured_revision(), "source revision differs from configured freeze")
    require(git("rev-parse", "HEAD") == revision, "checkout HEAD differs from source receipt")
    require(source_record.get("archive_paths") == list(ARCHIVE_PATHS), "source archive path manifest changed")
    require(descriptor(ARCHIVE) == source_record.get("archive"), "source archive changed since freeze")
    require(source_tree_identity(revision) == source_record.get("source_tree"), "source git tree identity changed")
    critical = source_record.get("critical_files")
    require(isinstance(critical, dict), "source critical-file manifest missing")
    if SOURCE.is_dir():
        source_files = [path for path in SOURCE.rglob("*") if path.is_file() and not path.is_symlink()]
        require(len(source_files) == source_record.get("source_file_count"), "frozen source file count changed")
        for name, expected in critical.items():
            require(descriptor(SOURCE / name) == expected, f"frozen source file changed: {name}")
    else:
        # Verification remains possible after deleting the extracted checkout:
        # the archive is the authoritative source custody artifact.
        with tarfile.open(ARCHIVE, mode="r") as archive:
            members = {member.name: member for member in archive.getmembers() if member.isfile()}
            require(len(members) == source_record.get("source_file_count"), "frozen archive file count changed")
            for name, expected in critical.items():
                member = members.get(f"litchi/{name}")
                require(member is not None, f"frozen archive file is missing: {name}")
                stream = archive.extractfile(member)
                require(stream is not None, f"frozen archive file cannot be read: {name}")
                digest = hashlib.sha256()
                for chunk in iter(lambda: stream.read(1 << 20), b""):
                    digest.update(chunk)
                actual = {"path": expected["path"], "bytes": member.size, "sha256": digest.hexdigest()}
                require(actual == expected, f"frozen archive file changed: {name}")


def run_command(argv: list[str], *, cwd: Path | None = None, env: dict[str, str] | None = None,
                stdout: Path | None = None, stderr: Path | None = None) -> subprocess.CompletedProcess[str]:
    out = stdout.open("w", encoding="utf-8") if stdout else subprocess.PIPE
    err = stderr.open("w", encoding="utf-8") if stderr else subprocess.PIPE
    try:
        return subprocess.run(
            argv,
            cwd=cwd,
            env=env,
            stdin=subprocess.DEVNULL,
            stdout=out,
            stderr=err,
            text=True,
            check=False,
        )
    finally:
        if stdout:
            out.close()
        if stderr:
            err.close()


def git(*args: str, check: bool = True) -> str:
    result = subprocess.run(
        ["git", *args], cwd=ROOT, stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, check=False,
    )
    if check and result.returncode:
        fail(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout.strip()


def configured_revision() -> str:
    revision = os.environ.get("LITCHI_PERF_0914_REVISION", DEFAULT_REVISION)
    require(re.fullmatch(r"[0-9a-f]{40}", revision) is not None, f"invalid source revision: {revision}")
    return revision


def source_tree_identity(revision: str) -> dict[str, object]:
    listing = git("ls-tree", "-r", "--full-tree", "--format=%(objectname) %(path)", revision, "--", *ARCHIVE_PATHS)
    names = git("ls-tree", "-r", "--full-tree", "--name-only", revision, "--", *ARCHIVE_PATHS)
    listing_bytes = (listing + "\n").encode()
    name_lines = [line for line in names.splitlines() if line]
    return {
        "file_count": len(name_lines),
        "listing_sha256": hashlib.sha256(listing_bytes).hexdigest(),
        "listing_bytes": len(listing_bytes),
    }


def host_snapshot() -> dict[str, object]:
    affinity = sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else []
    values: dict[str, object] = {
        "utc": now(),
        "hostname": platform.node(),
        "os": platform.platform(),
        "kernel": platform.release(),
        "machine": platform.machine(),
        "python": platform.python_version(),
        "logical_cpus": os.cpu_count(),
        "process_affinity": affinity,
        "loadavg": list(os.getloadavg()),
    }
    for path, key in ((Path("/proc/cpuinfo"), "cpuinfo"), (Path("/proc/meminfo"), "meminfo"),
                      (Path("/proc/sys/kernel/perf_event_paranoid"), "perf_event_paranoid")):
        try:
            values[key] = path.read_text(encoding="utf-8").strip()
        except OSError:
            values[key] = None
    try:
        values["rustc_version"] = subprocess.check_output(["rustc", "--version"], text=True).strip()
        values["rustc_verbose"] = subprocess.check_output(["rustc", "-vV"], text=True).strip()
        values["cargo_version"] = subprocess.check_output(["cargo", "--version"], text=True).strip()
    except (OSError, subprocess.CalledProcessError) as error:
        values["toolchain_error"] = str(error)
    return values


def prepare() -> None:
    require(not EXTERNAL.exists(), f"external custody directory already exists: {EXTERNAL}")
    EXTERNAL.mkdir(parents=True)
    source_parent = EXTERNAL / "source"
    source_parent.mkdir()
    revision = configured_revision()
    archive_command = ["git", "archive", "--format=tar", "--prefix=litchi/", revision, "--", *ARCHIVE_PATHS]
    with ARCHIVE.open("wb") as output:
        result = subprocess.run(archive_command, cwd=ROOT, stdin=subprocess.DEVNULL, stdout=output, stderr=subprocess.PIPE)
    require(result.returncode == 0, f"git archive failed: {result.stderr.decode(errors='replace')}")
    subprocess.run(["tar", "-xf", str(ARCHIVE), "-C", str(source_parent)], check=True)
    require(SOURCE.is_dir(), f"archive did not extract expected source root: {SOURCE}")
    print(json.dumps({"status": "prepared", "source_revision": revision, "archive": descriptor(ARCHIVE), "source": str(SOURCE)}))


def freeze() -> None:
    require(ROOT.is_dir(), "repository root is missing")
    require(SOURCE.is_dir() and ARCHIVE.is_file(), "prepare must run before freeze")
    require(not (EXTERNAL / "source-v1.json").exists(), "refusing to replace an existing source receipt; use fresh custody")
    require(not (EXTERNAL / "protocol-v1.json").exists(), "refusing to replace an existing protocol receipt; use fresh custody")
    require(CPU in (sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else range(128)),
            f"requested CPU {CPU} is outside this process affinity")
    revision = configured_revision()
    require(git("rev-parse", "HEAD") == revision, "checkout HEAD differs from the explicitly frozen source revision")
    status = git("status", "--porcelain")
    source_tree = source_tree_identity(revision)
    archive = descriptor(ARCHIVE)
    source_files = [path for path in SOURCE.rglob("*") if path.is_file() and not path.is_symlink()]
    critical = [
        "Cargo.toml",
        "tools/perf-baseline/Cargo.toml",
        "tools/perf-baseline/Cargo.lock",
        "tools/perf-baseline/src/lib.rs",
        "tools/perf-baseline/src/main.rs",
        "test-data/poi/test-data/spreadsheet/54016.xls",
        "test-data/rtf/watermark.rtf",
    ]
    critical_descriptors = {name: descriptor(SOURCE / name) for name in critical}
    source_record = {
        "schema": "litchi.performance.0914.source.v1",
        "created_utc": now(),
        "repository": str(ROOT),
        "source_revision": revision,
        "worktree_status_at_freeze": status,
        "worktree_dirty_at_freeze": bool(status),
        "archive_paths": list(ARCHIVE_PATHS),
        "archive_command": f"git archive --format=tar --prefix=litchi/ {revision} -- Cargo.toml clippy.toml deny.toml rust-toolchain.toml rustfmt.toml .cargo crates test-data/poi/test-data/spreadsheet/54016.xls test-data/rtf/watermark.rtf tools/perf-baseline",
        "archive": archive,
        "source_root": str(SOURCE),
        "source_file_count": len(source_files),
        "source_tree": source_tree,
        "critical_files": critical_descriptors,
        "supplementary_compile_fixtures": {
            name: critical_descriptors[name]
            for name in (
                "test-data/poi/test-data/spreadsheet/54016.xls",
                "test-data/rtf/watermark.rtf",
            )
        },
        "cargo_lock_sha256": critical_descriptors["tools/perf-baseline/Cargo.lock"]["sha256"],
        "host": host_snapshot(),
    }
    replace_json(EXTERNAL / "source-v1.json", source_record)
    protocol = {
        "schema": "litchi.performance.0914.protocol.v1",
        "created_utc": now(),
        "case": CASE,
        "warmup_iterations": WARMUP,
        "samples": SAMPLES,
        "children": ["child-a", "child-b", "child-c", "child-d"],
        "schedule": "four independent same-route fresh child processes, sequential, no eager control and no speedup comparison",
        "cpu": CPU,
        "source_revision": revision,
        "binary_profile": "release",
        "build_flags": ["--release", "--locked", "-j2"],
        "build_command": "CARGO_TARGET_DIR=/home/zhuhe/litchi-performance-0914/target cargo build --release --locked -j2 --manifest-path /home/zhuhe/litchi-performance-0914/source/litchi/tools/perf-baseline/Cargo.toml",
        "run_command_template": "taskset -c 12 /usr/bin/time -v -o <process.txt> <binary> --warmup 30 --samples 200 --case docx_source_backed_existing_section_layout_edit_save --json <report.json>",
        "timing_scope": "harness elapsed_ns: preparation, source-backed open, section query, edit, commit, sequential publication; exact six-phase sum",
        "process_scope": "GNU time -v: setup-inclusive child process including corpus generation, benchmark, verification, JSON serialization, and process teardown",
        "performance_claim": "none",
        "allocation_metrics": "unavailable in normal release binary; no allocator-instrumented timing claim",
        "physical_io_metrics": "unavailable; source is in-memory InstrumentedSource and counters are logical harness observations",
        "host_at_freeze": host_snapshot(),
        "source_scope": descriptor(SOURCE_SCOPE) if SOURCE_SCOPE.is_file() else None,
    }
    replace_json(EXTERNAL / "protocol-v1.json", protocol)
    print(json.dumps({"status": "frozen", "source_revision": revision, "archive": archive, "cpu": CPU}))


def build() -> None:
    source_record = read_json(EXTERNAL / "source-v1.json")
    require(isinstance(source_record, dict), "source receipt malformed")
    validate_source_record(source_record)
    BUILD_DIR.mkdir(parents=True, exist_ok=True)
    stdout_path = BUILD_DIR / "build.stdout.log"
    stderr_path = BUILD_DIR / "build.stderr.log"
    exit_path = BUILD_DIR / "build.exit"
    require(not (EXTERNAL / "build-v1.json").exists(), "refusing to replace an existing build receipt; use fresh custody")
    require(not exit_path.exists(), "refusing to adopt or replace an existing build exit receipt; use fresh custody")
    env = os.environ.copy()
    env.update({"CARGO_TARGET_DIR": str(TARGET)})
    argv = ["cargo", "build", "--release", "--locked", "-j2", "--manifest-path", str(SOURCE / "tools/perf-baseline/Cargo.toml")]
    started = now()
    result = run_command(argv, cwd=SOURCE, env=env, stdout=stdout_path, stderr=stderr_path)
    exit_path.write_text(f"exit_code={result.returncode}\n", encoding="utf-8")
    receipt: dict[str, object] = {
        "schema": "litchi.performance.0914.build.v1",
        "started_utc": started,
        "finished_utc": now(),
        "argv": argv,
        "cwd": str(SOURCE),
        "environment": {
            key: env.get(key)
            for key in (
                "CARGO_TARGET_DIR", "RUSTFLAGS", "CARGO_BUILD_TARGET", "CARGO_ENCODED_RUSTFLAGS",
                "RUSTC_WRAPPER", "CARGO_PROFILE_RELEASE_OPT_LEVEL", "CARGO_PROFILE_RELEASE_LTO",
                "CARGO_PROFILE_RELEASE_CODEGEN_UNITS", "CARGO_PROFILE_RELEASE_DEBUG",
                "CARGO_PROFILE_RELEASE_STRIP",
            )
        },
        "source": source_record,
        "source_revision": source_record["source_revision"],
        "cargo_lock": descriptor(SOURCE / "tools/perf-baseline/Cargo.lock"),
        "rustc_version": subprocess.check_output(["rustc", "--version"], text=True).strip(),
        "cargo_version": subprocess.check_output(["cargo", "--version"], text=True).strip(),
        "stdout": descriptor(stdout_path),
        "stderr": descriptor(stderr_path),
        "exit": descriptor(exit_path),
        "exit_code": result.returncode,
        "receipt_mode": "built-by-script",
    }
    if result.returncode == 0:
        receipt["binary"] = descriptor(BINARY)
    replace_json(EXTERNAL / "build-v1.json", receipt)
    require(result.returncode == 0, f"release build failed; see {stderr_path}")
    print(json.dumps({"status": "built", "binary": receipt["binary"], "build": str(EXTERNAL / "build-v1.json")}))


def parse_time_v(path: Path) -> dict[str, object]:
    text = path.read_text(encoding="utf-8", errors="replace")
    values: dict[str, object] = {"raw_sha256": hashlib.sha256(text.encode()).hexdigest()}
    patterns = {
        "user_seconds": r"User time \(seconds\): ([0-9.]+)",
        "system_seconds": r"System time \(seconds\): ([0-9.]+)",
        "elapsed": r"Elapsed \(wall clock\) time \(h\?:m\?:s\.?\): ([0-9:.]+)",
        "max_rss_kib": r"Maximum resident set size \(kbytes\): (\d+)",
        "major_page_faults": r"Major \(requiring I/O\) page faults: (\d+)",
        "minor_page_faults": r"Minor \(reclaiming a frame\) page faults: (\d+)",
        "voluntary_context_switches": r"Voluntary context switches: (\d+)",
        "involuntary_context_switches": r"Involuntary context switches: (\d+)",
        "file_system_inputs": r"File system inputs: (\d+)",
        "file_system_outputs": r"File system outputs: (\d+)",
    }
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            value = match.group(1)
            values[key] = float(value) if key.endswith("seconds") else value
    require("max_rss_kib" in values, f"time -v output lacks max RSS: {path}")
    values["bytes"] = path.stat().st_size
    values["path"] = str(path)
    return values


def run_leg(label: str) -> None:
    protocol = read_json(EXTERNAL / "protocol-v1.json")
    build_record = read_json(EXTERNAL / "build-v1.json")
    require(isinstance(protocol, dict) and isinstance(build_record, dict), "protocol/build receipts malformed")
    require(build_record["exit_code"] == 0, "cannot run after failed build")
    require(descriptor(BINARY) == build_record["binary"], "binary changed since build receipt")
    require(label in {"child-a", "child-b", "child-c", "child-d"}, f"unknown child label: {label}")
    leg = RUN_DIR / label
    require(not leg.exists(), f"refusing to replace existing child directory: {leg}")
    leg.mkdir(parents=True)
    report = leg / "report.json"
    process = leg / "process.txt"
    stdout = leg / "stdout.txt"
    stderr = leg / "stderr.txt"
    argv = [
        "/usr/bin/taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o", str(process),
        str(BINARY), "--warmup", str(WARMUP), "--samples", str(SAMPLES), "--case", CASE,
        "--json", str(report),
    ]
    env = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC", "RUST_BACKTRACE": "0"}
    receipt = {
        "schema": "litchi.performance.0914.child.v1",
        "label": label,
        "started_utc": now(),
        "argv": argv,
        "cwd": str(SOURCE),
        "environment": env,
        "protocol": descriptor(EXTERNAL / "protocol-v1.json"),
        "build": descriptor(EXTERNAL / "build-v1.json"),
        "binary": descriptor(BINARY),
        "cpu": CPU,
        "host_load_before": list(os.getloadavg()),
    }
    write_json(leg / "start.json", receipt)
    result = run_command(argv, cwd=SOURCE, env=env, stdout=stdout, stderr=stderr)
    receipt.update({
        "finished_utc": now(),
        "exit_code": result.returncode,
        "host_load_after": list(os.getloadavg()),
        "report": descriptor(report) if report.exists() else None,
        "process": descriptor(process) if process.exists() else None,
        "stdout": descriptor(stdout),
        "stderr": descriptor(stderr),
        "process_metrics": parse_time_v(process) if process.exists() else None,
    })
    replace_json(leg / "receipt.json", receipt)
    require(result.returncode == 0, f"child {label} failed; see {stderr}")
    require(report.is_file(), f"child {label} produced no report")
    print(json.dumps({"status": "child-complete", "label": label, "report": descriptor(report), "process_metrics": receipt["process_metrics"]}))


def integer(value: object, label: str, positive: bool = False) -> int:
    require(type(value) is int, f"{label} must be integer")
    require(value > 0 if positive else value >= 0, f"{label} out of range")
    return value


def percentile(values: list[int], fraction: float) -> int:
    ordered = sorted(values)
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def stats(values: list[int]) -> dict[str, object]:
    require(values, "empty statistic vector")
    return {
        "samples": len(values),
        "min": min(values),
        "p50": (sorted(values)[(len(values) - 1) // 2] + sorted(values)[len(values) // 2]) // 2,
        "p95": percentile(values, 0.95),
        "p99": percentile(values, 0.99),
        "max": max(values),
        "mean": statistics.fmean(values),
    }


def regeneration_command(source_revision: str) -> str:
    replay_path = EXTERNAL / "reduced-replay.json"
    return (
        f"LITCHI_PERF_0914_ROOT={EXTERNAL} LITCHI_PERF_0914_REVISION={source_revision} "
        f"python3 docs/performance/results/change-0914/capture.py reduce --output {replay_path}"
    )


def validate_report(report: dict[str, object], label: str, binary: dict[str, object]) -> dict[str, object]:
    require(report.get("schema_version") == 1, f"{label}: unexpected report schema")
    require(report.get("configuration", {}).get("samples_per_case") == SAMPLES, f"{label}: sample count mismatch")
    require(report.get("configuration", {}).get("warmup_iterations_per_case") == WARMUP, f"{label}: warmup mismatch")
    require(report.get("configuration", {}).get("cases") == [CASE], f"{label}: selected case mismatch")
    require(report.get("configuration", {}).get("corpus_shapes") == ["tiny", "many-small", "few-large", "wide-root"], f"{label}: shape axis was changed")
    require(report.get("configuration", {}).get("payload_kinds") == ["compressible", "incompressible"], f"{label}: payload axis was changed")
    require(report.get("binary_identity", {}).get("binary_sha256") == binary["sha256"], f"{label}: binary identity mismatch")
    require(report.get("binary_identity", {}).get("binary_bytes") == binary["bytes"], f"{label}: binary size mismatch")
    require(report.get("tool", {}).get("profile") == "release", f"{label}: report is not release profile")
    require(report.get("tool", {}).get("instrumentation") == "none", f"{label}: unexpected instrumentation")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1, f"{label}: expected one result")
    result = results[0]
    require(result.get("case") == CASE, f"{label}: result case mismatch")
    corpus = result.get("corpus")
    require(isinstance(corpus, dict), f"{label}: missing corpus manifest")
    for key, expected in CORPUS_EXPECTED.items():
        require(key in corpus, f"{label}: corpus field {key} is missing")
        require(corpus[key] == expected, f"{label}: corpus {key} mismatch")
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict), f"{label}: elapsed statistics missing")
    samples = elapsed.get("samples")
    require(isinstance(samples, list) and len(samples) == SAMPLES, f"{label}: elapsed samples mismatch")
    sample_order = elapsed.get("sample_order")
    require(isinstance(sample_order, list) and sorted(sample_order) == list(range(SAMPLES)), f"{label}: sample order mismatch")
    totals = [integer(value, f"{label}.elapsed_ns.samples[{index}]", True) for index, value in enumerate(samples)]
    source = result.get("source")
    require(isinstance(source, dict), f"{label}: source summary missing")
    section = source.get("docx_section_layout")
    require(isinstance(section, dict), f"{label}: section-layout summary missing")
    require(section.get("performance_claim") == "none", f"{label}: unsupported performance claim")
    require(section.get("sample_order") == sample_order, f"{label}: section sample order mismatch")
    section_identity = {
        "implementation": "litchi-docx-source-backed-section-layout",
        "source_bytes": EXPECTED["archive_bytes"],
        "source_sha256": EXPECTED["archive_sha256"],
        "expected_output_sha256": EXPECTED["output_sha256"],
        "total_main_story_paragraph_count": 257,
        "section_count": 3,
        "paragraph_owned_section_positions": [64, 129],
        "selected_paragraph": 64,
        "columns_before": 2,
        "columns_after": 3,
        "media_count": 4,
        "media_bytes": EXPECTED["entry_bytes"],
    }
    for key, expected in section_identity.items():
        require(section.get(key) == expected, f"{label}: section field {key} mismatch")
    for key in ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns"):
        values = section.get(key)
        require(isinstance(values, list) and len(values) == SAMPLES, f"{label}: phase {key} cardinality")
        integer_values = [integer(value, f"{label}.{key}[{index}]", True) for index, value in enumerate(values)]
        section[key] = integer_values
    for key in (
        "phase_sum_verified", "semantic_reopen_all_sections_verified", "table_cell_sectpr_excluded_verified",
        "raw_untouched_members_verified", "raw_member_order_and_comment_verified", "header_footer_relationships_verified",
        "source_immutability_verified", "patch_forward_inverse_verified", "publication_inverse_exact_source_verified",
        "stale_foreign_signed_noop_limits_partial_sink_verified", "cache_counters_verified", "sink_counters_verified",
    ):
        require(section.get(key) is True, f"{label}: gate {key} failed")
    require(sum(section[key][0] for key in ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")) > 0, f"{label}: empty phase values")
    for index in range(SAMPLES):
        require(sum(section[key][index] for key in ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")) == totals[index], f"{label}: phase sum mismatch at {index}")
    source_reads = source.get("read_calls")
    source_bytes = source.get("read_bytes")
    ordinary_calls = source.get("ordinary_payload_read_calls")
    ordinary_bytes = source.get("ordinary_payload_read_bytes")
    for key, values in (("read_calls", source_reads), ("read_bytes", source_bytes), ("ordinary_payload_read_calls", ordinary_calls), ("ordinary_payload_read_bytes", ordinary_bytes)):
        require(isinstance(values, list) and len(values) == SAMPLES, f"{label}: source {key} cardinality")
        require(all(integer(value, f"{label}.source.{key}") > 0 for value in values), f"{label}: source {key} has zero")
    max_in_flight = source.get("max_in_flight_reads")
    require(isinstance(max_in_flight, list) and len(max_in_flight) == SAMPLES, f"{label}: max in-flight cardinality")
    require(all(integer(value, f"{label}.source.max_in_flight_reads") > 0 for value in max_in_flight), f"{label}: max in-flight source reads has zero")
    materializations = source.get("ordinary_payload_materializations")
    require(isinstance(materializations, list) and materializations == [1] * SAMPLES, f"{label}: materialization contract")
    for key, values in (
        ("source_read_calls", section.get("source_read_calls")),
        ("source_read_bytes", section.get("source_read_bytes")),
        ("ordinary_payload_materializations", section.get("ordinary_payload_materializations")),
    ):
        require(isinstance(values, list) and len(values) == SAMPLES, f"{label}: section {key} cardinality")
    require(section["source_read_calls"] == source_reads, f"{label}: source read-call vectors are misaligned")
    require(section["source_read_bytes"] == source_bytes, f"{label}: source read-byte vectors are misaligned")
    require(section["ordinary_payload_materializations"] == materializations, f"{label}: materialization vectors are misaligned")
    cache_vectors = {
        "hits": section.get("source_cache_before_publication_after_commit_hits"),
        "cold_loads": section.get("source_cache_before_publication_after_commit_cold_loads"),
        "successful_loads": section.get("source_cache_before_publication_after_commit_successful_loads"),
        "retained_entries": section.get("source_cache_before_publication_after_commit_retained_entries"),
        "retained_bytes": section.get("source_cache_before_publication_after_commit_retained_bytes"),
    }
    for key, values in cache_vectors.items():
        require(isinstance(values, list) and len(values) == SAMPLES, f"{label}: cache {key} cardinality")
        require(all(integer(value, f"{label}.cache.{key}") >= 0 for value in values), f"{label}: cache {key} out of range")
    require(cache_vectors["cold_loads"] == [1] * SAMPLES, f"{label}: cache cold-load contract")
    require(cache_vectors["successful_loads"] == [1] * SAMPLES, f"{label}: cache successful-load contract")
    require(cache_vectors["retained_entries"] == [1] * SAMPLES, f"{label}: cache retained-entry contract")
    require(all(value > 0 for value in cache_vectors["retained_bytes"]), f"{label}: cache retained bytes are zero")
    require(len(set(cache_vectors["hits"])) == 1, f"{label}: cache hit counter is nondeterministic")
    sink = result.get("sink")
    require(isinstance(sink, dict), f"{label}: sink summary missing")
    require(sink.get("accepted_bytes") == EXPECTED["output_bytes"], f"{label}: sink output bytes mismatch")
    require(integer(sink.get("write_calls"), f"{label}.sink.write_calls", True) > 0, f"{label}: sink has no writes")
    require(integer(sink.get("largest_write"), f"{label}.sink.largest_write", True) <= 65536, f"{label}: sink write bound")
    buckets = sink.get("write_size_buckets")
    require(isinstance(buckets, dict), f"{label}: sink write buckets missing")
    bucket_keys = ("bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536")
    for key in bucket_keys:
        integer(buckets.get(key), f"{label}.sink.write_size_buckets.{key}")
    require(sum(buckets[key] for key in bucket_keys) == sink["write_calls"], f"{label}: sink bucket count mismatch")
    require(buckets["bytes_over_65536"] == 0, f"{label}: sink has oversized write")
    output_hashes = section.get("output_sha256")
    require(output_hashes == [EXPECTED["output_sha256"]] * SAMPLES, f"{label}: output hash stability")
    require(result.get("output_sha256") == EXPECTED["output_sha256"], f"{label}: top-level output hash")
    phase_stats = {key: stats(section[key]) for key in ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")}
    return {
        "label": label,
        "elapsed_ns": stats(totals),
        "phase_ns": phase_stats,
        "source_read_calls": stats([int(value) for value in source_reads]),
        "source_read_bytes": stats([int(value) for value in source_bytes]),
        "ordinary_payload_read_calls": stats([int(value) for value in ordinary_calls]),
        "ordinary_payload_read_bytes": stats([int(value) for value in ordinary_bytes]),
        "ordinary_payload_materializations": 1,
        "max_in_flight_reads": stats([int(value) for value in max_in_flight]),
        "cache_before_publication_after_commit": {
            key: stats([int(value) for value in values])
            for key, values in cache_vectors.items()
        },
        "output_bytes": sink["accepted_bytes"],
        "output_sha256": result["output_sha256"],
        "sink": {
            "accepted_bytes": sink["accepted_bytes"],
            "write_calls": sink["write_calls"],
            "largest_write": sink["largest_write"],
            "write_size_buckets": {key: buckets[key] for key in bucket_keys},
        },
        "sink_largest_write": sink["largest_write"],
        "gates": {key: section[key] for key in (
            "phase_sum_verified", "semantic_reopen_all_sections_verified", "table_cell_sectpr_excluded_verified",
            "raw_untouched_members_verified", "raw_member_order_and_comment_verified", "header_footer_relationships_verified",
            "source_immutability_verified", "patch_forward_inverse_verified", "publication_inverse_exact_source_verified",
            "stale_foreign_signed_noop_limits_partial_sink_verified", "cache_counters_verified", "sink_counters_verified",
        )},
    }


def compute_reduction() -> dict[str, object]:
    protocol_path = EXTERNAL / "protocol-v1.json"
    build_path = EXTERNAL / "build-v1.json"
    source_path = EXTERNAL / "source-v1.json"
    source_record = read_json(source_path)
    protocol = read_json(protocol_path)
    build_record = read_json(build_path)
    require(isinstance(source_record, dict), "source receipt malformed")
    require(isinstance(protocol, dict) and isinstance(build_record, dict), "protocol/build malformed")
    validate_source_record(source_record)
    require(protocol.get("source_revision") == source_record["source_revision"], "protocol source revision mismatch")
    require(protocol.get("case") == CASE and protocol.get("warmup_iterations") == WARMUP and protocol.get("samples") == SAMPLES, "protocol sampling contract changed")
    require(protocol.get("children") == ["child-a", "child-b", "child-c", "child-d"], "protocol child contract changed")
    require(protocol.get("cpu") == CPU and protocol.get("performance_claim") == "none", "protocol performance contract changed")
    require(build_record.get("source_revision") == source_record["source_revision"], "build source revision mismatch")
    require(build_record.get("source") == source_record, "build source receipt differs from frozen source")
    require(protocol.get("build_flags") == ["--release", "--locked", "-j2"], "build flags changed")
    require(build_record.get("argv") == [
        "cargo", "build", "--release", "--locked", "-j2",
        "--manifest-path", str(SOURCE / "tools/perf-baseline/Cargo.toml"),
    ], "build argv changed")
    binary = build_record.get("binary")
    require(isinstance(binary, dict), "build binary descriptor missing")
    require(descriptor(BINARY) == binary, "binary descriptor changed")
    legs: list[dict[str, object]] = []
    for label in ("child-a", "child-b", "child-c", "child-d"):
        folder = RUN_DIR / label
        receipt = read_json(folder / "receipt.json")
        report_path = folder / "report.json"
        report = read_json(report_path)
        require(isinstance(receipt, dict) and isinstance(report, dict), f"{label}: malformed receipt/report")
        require(receipt.get("exit_code") == 0, f"{label}: nonzero child exit")
        require(receipt.get("cpu") == CPU, f"{label}: CPU mismatch")
        require(receipt.get("binary") == binary, f"{label}: binary receipt mismatch")
        require(receipt.get("protocol") == descriptor(protocol_path), f"{label}: protocol descriptor mismatch")
        require(receipt.get("build") == descriptor(build_path), f"{label}: build descriptor mismatch")
        require(receipt.get("report") == descriptor(report_path), f"{label}: report descriptor mismatch")
        process_path = folder / "process.txt"
        require(receipt.get("process") == descriptor(process_path), f"{label}: process descriptor mismatch")
        reduced = validate_report(report, label, binary)
        process = receipt.get("process_metrics")
        require(isinstance(process, dict), f"{label}: process metrics missing")
        reduced["report"] = descriptor(report_path)
        reduced["receipt"] = descriptor(folder / "receipt.json")
        reduced["process"] = process
        legs.append(reduced)
    medians = [leg["elapsed_ns"]["p50"] for leg in legs]
    pooled = []
    pooled_phases: dict[str, list[int]] = {key: [] for key in ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")}
    for leg in legs:
        report = read_json(Path(leg["report"]["path"]))
        pooled.extend(report["results"][0]["elapsed_ns"]["samples"])
        section = report["results"][0]["source"]["docx_section_layout"]
        for key in pooled_phases:
            pooled_phases[key].extend(section[key])
    pooled_elapsed_sum = sum(int(value) for value in pooled)
    pooled_phase_sum = {key: sum(values) for key, values in pooled_phases.items()}
    require(sum(pooled_phase_sum.values()) == pooled_elapsed_sum, "pooled phase sums do not equal pooled elapsed")
    result = {
        "schema": "litchi.performance.0914.docx-section-layout-reduction.v1",
        "source_revision": source_record["source_revision"],
        "source_receipt": descriptor(source_path),
        "case": CASE,
        "corpus": EXPECTED,
        "protocol": descriptor(protocol_path),
        "build": descriptor(build_path),
        "binary": binary,
        "children": 4,
        "warmup_iterations": WARMUP,
        "samples_per_child": SAMPLES,
        "cpu": CPU,
        "timing_scope": protocol["timing_scope"],
        "process_scope": protocol["process_scope"],
        "performance_claim": "none",
        "allocation_metrics": protocol["allocation_metrics"],
        "physical_io_metrics": protocol["physical_io_metrics"],
        "child_elapsed_p50_ns": medians,
        "child_elapsed_p50_range_ns": {"min": min(medians), "max": max(medians)},
        "pooled_elapsed_descriptive_only": stats([int(value) for value in pooled]),
        "pooled_phase_descriptive_only": {
            key: stats(values) for key, values in pooled_phases.items()
        },
        "pooled_phase_fraction_of_elapsed": {
            key: (value / pooled_elapsed_sum) for key, value in pooled_phase_sum.items()
        },
        "children_detail": legs,
        "regeneration_command": regeneration_command(source_record["source_revision"]),
        "guards": {
            "all_children_exit_zero": True,
            "all_reports_match_binary": True,
            "all_reports_match_source_corpus": True,
            "all_reports_match_warmup_and_samples": True,
            "all_semantic_and_preservation_gates": True,
            "all_source_and_sink_gates": True,
            "all_output_hashes_stable": True,
            "no_speedup_claim": True,
        },
    }
    return result


def reduce(output: Path | None = None) -> None:
    result = compute_reduction()
    output_path = (output or (HERE / "reduced.json")).resolve()
    replace_json(output_path, result)
    print(json.dumps({"status": "reduced", "output": descriptor(output_path), "child_elapsed_p50_ns": result["child_elapsed_p50_ns"]}))


def verify() -> None:
    reduced_path = HERE / "reduced.json"
    reduced = read_json(reduced_path)
    require(isinstance(reduced, dict), "reduced report malformed")
    require(reduced.get("schema") == "litchi.performance.0914.docx-section-layout-reduction.v1", "reduced schema mismatch")
    require(reduced.get("performance_claim") == "none", "reduced report claims performance")
    require(reduced.get("regeneration_command") == regeneration_command(reduced["source_revision"]), "regeneration command changed")
    source_record = read_json(EXTERNAL / "source-v1.json")
    require(source_record["source_revision"] == reduced["source_revision"], "reduced source revision mismatch")
    require(descriptor(EXTERNAL / "source-v1.json") == reduced["source_receipt"], "source descriptor mismatch")
    protocol = read_json(EXTERNAL / "protocol-v1.json")
    require(protocol["case"] == CASE and protocol["warmup_iterations"] == WARMUP and protocol["samples"] == SAMPLES, "protocol mismatch")
    recomputed = compute_reduction()
    require(recomputed == reduced, "retained reduction differs from deterministic recomputation")
    print(json.dumps({"status": "verified", "reduced": descriptor(reduced_path)}))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("prepare", "freeze", "build", "run", "reduce", "verify"))
    parser.add_argument("--label", choices=("child-a", "child-b", "child-c", "child-d"))
    parser.add_argument("--output", type=Path, help="reduction output path for reduce; defaults to the retained reduced.json")
    args = parser.parse_args()
    try:
        if args.action == "prepare":
            prepare()
        elif args.action == "freeze":
            freeze()
        elif args.action == "build":
            build()
        elif args.action == "run":
            if args.label:
                run_leg(args.label)
            else:
                for label in ("child-a", "child-b", "child-c", "child-d"):
                    run_leg(label)
        elif args.action == "reduce":
            reduce(args.output)
        elif args.action == "verify":
            verify()
    except (EvidenceError, OSError, subprocess.SubprocessError) as error:
        print(f"capture failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
