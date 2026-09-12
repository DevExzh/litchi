#!/usr/bin/env python3
"""Fail-closed verifier for the source-bound 0517 DOCX evidence bundle.

This verifier only reads retained evidence.  Row-level CSV invariants are
delegated to ``capture.validate`` and the raw Callgrind grammar to the 0515
parser.  A missing candidate epoch is reported as pending; ``--require-complete``
turns that state into a failure for the final seal.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import hashlib
import importlib.util
import json
import math
import re
import subprocess
import sys
import tarfile
from collections import Counter
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
SCRATCH = Path("/tmp/litchi-goal-0517")
BASE_REVISION = "201f65094ff0ca6186fbf1e2ffa6b3641290253b"
SOURCE_PRODUCTION = "crates/litchi-opc/src/source_backed.rs"
SOURCE_TEST = "crates/litchi-opc/tests/source_xml_publication.rs"
DEFAULT_CANDIDATE_PATHS = {SOURCE_PRODUCTION, SOURCE_TEST}
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
RSS_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")
NAME_RE = re.compile(r"^p(?:128|512)-k(?:1|8|32)-(?:owned|file)-(?:repeated|batch)$")
PROFILE_FUNCTION = (
    "litchi_docx::source_backed::Package::publish_document_commit_to_stream"
)
PROFILE_CALLER = "managed_paragraph_batch_perf::run_sample"
EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "cache-misses",
    "context-switches",
    "cpu-migrations",
    "page-faults",
)
NATIVE_NAMES = {
    f"p{paragraphs}-k{count}-{source}-{mode}"
    for paragraphs in (128, 512)
    for count in (1, 8, 32)
    for source in ("owned", "file")
    for mode in ("repeated", "batch")
}
PROFILE_NAMES = {
    "p128-k1-owned-repeated",
    "p128-k1-owned-batch",
    "p512-k1-owned-repeated",
    "p512-k1-owned-batch",
    "p512-k32-owned-repeated",
    "p512-k32-owned-batch",
}
LANE_SPECS = {
    "preflight": (NATIVE_NAMES, 1, 0, 1, "native"),
    "r1": (NATIVE_NAMES, 30, 3, 2, "native"),
    "r2": (NATIVE_NAMES, 30, 3, 2, "native"),
    "profile-preflight": ({"p128-k1-owned-batch"}, 1, 0, 1, "profile"),
    "profile-r1": (PROFILE_NAMES, 1, 0, 1, "profile"),
    "profile-r2": (PROFILE_NAMES, 1, 0, 1, "profile"),
    "hardware": (PROFILE_NAMES, 30, 3, 2, "hardware"),
    "after-r1": (NATIVE_NAMES, 30, 3, 2, "native"),
    "after-r2": (NATIVE_NAMES, 30, 3, 2, "native"),
    "profile-after-r1": (PROFILE_NAMES, 1, 0, 1, "profile"),
    "profile-after-r2": (PROFILE_NAMES, 1, 0, 1, "profile"),
}
BASELINE_LANES = (
    "preflight", "r1", "r2", "profile-preflight", "profile-r1",
    "profile-r2", "hardware",
)
CANDIDATE_LANES = (
    "after-r1", "after-r2", "profile-after-r1", "profile-after-r2",
)
QUALITY_COMMANDS = {
    "fmt": ["cargo", "fmt", "--all", "--check"],
    "ooxml-tests": [
        "cargo", "test", "--locked", "-p", "litchi-opc", "-p", "litchi-docx",
        "-p", "litchi-xlsx", "-p", "litchi-pptx", "-p", "litchi-xlsb",
        "--all-features", "--", "--test-threads=2",
    ],
    "workspace-check": ["cargo", "check", "--locked", "--workspace", "--all-features"],
    "clippy": [
        "cargo", "clippy", "--locked", "-p", "litchi-opc", "-p", "litchi-docx",
        "--all-features", "--lib", "--", "-D", "warnings",
    ],
    "rustdoc": [
        "cargo", "doc", "--locked", "-p", "litchi-opc", "-p", "litchi-docx",
        "--all-features", "--no-deps",
    ],
    "boundaries": ["python3", "-B", "tools/check_crate_boundaries.py"],
    "zip64-generate": [
        "python3", "-B", "docs/performance/results/change-0415/interop.py", "generate",
        str(SCRATCH / "independent-zip64.zip"),
    ],
    "zip64-test": [
        "cargo", "test", "--locked", "-p", "litchi-opc", "--all-features",
        "--test", "external_zip64_source", "--", "--ignored", "--nocapture",
    ],
    "claims": [
        "python3", "-B", "tools/check_perf_claims.py", "--registry",
        "docs/performance/claim-registry-v1.json", "--repo-root", ".",
        "--evidence-root", ".", "--mode", "strict",
    ],
}
QUALITY_ENVIRONMENT = {
    "CARGO_TARGET_DIR": "/home/zhuhe/litchi-goal-0517-target",
    "CARGO_BUILD_JOBS": "2",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_DEBUG": "0",
    "CARGO_PROFILE_DEV_DEBUG": "0",
    "CARGO_PROFILE_TEST_DEBUG": "0",
    "TMPDIR": str(SCRATCH),
    "RUSTDOCFLAGS": "-D warnings",
}
ZIP64_DECLARED_BYTES = 2**32
ZIP64_BUFFER_BYTES = 65536
ZIP64_PHYSICAL_LIMIT = 64 * 1024 * 1024
ZIP64_ENTRIES = ("[Content_Types].xml", "_rels/.rels", "document.xml", "large.bin")
HELPER_PATHS = {
    "docs/performance/results/change-0500/verify-evidence.py",
    "docs/performance/results/change-0515/analyze.py",
    "docs/performance/results/change-0415/interop.py",
    "tools/check_crate_boundaries.py",
    "tools/crate_boundaries.json",
    "tools/check_perf_claims.py",
    "docs/performance/claim-registry-v1.json",
}


class VerificationError(ValueError):
    """A missing, malformed, or inconsistent retained artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def _reject_constant(value: str) -> None:
    raise VerificationError(f"non-finite JSON number {value!r}")


def _no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def load_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_pairs,
            parse_constant=_reject_constant,
        )
    except VerificationError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise VerificationError(f"cannot load {path}: {error}") from error


def regular(path: Path, context: str, *, nonempty: bool = True) -> Path:
    require(path.is_file() and not path.is_symlink(),
            f"{context} is missing, symlinked, or not a regular file")
    if nonempty:
        require(path.stat().st_size > 0, f"{context} is empty")
    return path


def digest(path: Path) -> str:
    regular(path, f"hash target {path}", nonempty=False)
    result = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            result.update(block)
    return result.hexdigest()


def check_digest(value: Any, context: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{context} is not a lowercase SHA-256 digest")
    return value


def safe_relative(value: Any, context: str) -> Path:
    require(isinstance(value, str) and value, f"{context} is not a path")
    path = Path(value)
    require(not path.is_absolute() and ".." not in path.parts
            and path.as_posix() == value, f"{context} is unsafe")
    return path


def load_capture() -> Any:
    """Load capture.py without entering its CLI or launching a child."""

    if str(HERE) not in sys.path:
        sys.path.insert(0, str(HERE))
    try:
        spec = importlib.util.spec_from_file_location("change0517_capture_verify", HERE / "capture.py")
        require(spec is not None and spec.loader is not None, "cannot load capture.py")
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)
        return module
    except (OSError, ImportError, AttributeError) as error:
        raise VerificationError(f"cannot load capture.py: {error}") from error


capture = load_capture()


def load_profile_parser() -> Any:
    path = HERE.parent / "change-0515" / "analyze.py"
    regular(path, "0515 Callgrind parser")
    spec = importlib.util.spec_from_file_location("change0515_parser_for_0517", path)
    require(spec is not None and spec.loader is not None,
            "cannot load 0515 Callgrind parser")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as error:  # parser turns malformed profiles into errors later
        raise VerificationError(f"cannot initialize Callgrind parser: {error}") from error
    return module


profile_parser = load_profile_parser()


def load_analyzer(path: Path, name: str) -> Any:
    regular(path, name)
    spec = importlib.util.spec_from_file_location(
        "change0517_" + name.replace(".", "_"), path
    )
    require(spec is not None and spec.loader is not None,
            f"cannot load {name}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as error:
        raise VerificationError(f"cannot initialize {name}: {error}") from error
    return module


profile_analyzer = load_analyzer(HERE / "analyze_profiles.py", "analyze_profiles.py")
native_analyzer = load_analyzer(HERE / "analyze_native.py", "analyze_native.py")
hardware_analyzer = load_analyzer(HERE / "analyze_hardware.py", "analyze_hardware.py")
counter_analyzer = load_analyzer(HERE / "analyze_counters.py", "analyze_counters.py")
candidate_comparator = load_analyzer(HERE / "compare_candidate.py", "compare_candidate.py")
profile_comparator = load_analyzer(HERE / "compare_profile_lanes.py", "compare_profile_lanes.py")


def manifest(path: Path, context: str) -> dict[str, str]:
    value = load_json(path)
    require(isinstance(value, dict) and value, f"{context} is not a non-empty object")
    result: dict[str, str] = {}
    for name, expected in value.items():
        safe_relative(name, f"{context} path")
        result[name] = check_digest(expected, f"{context}.{name}")
    return result


def live_source_manifest() -> dict[str, str]:
    try:
        # capture.py imports the authoritative run module; invoking it does no
        # mutation and keeps the inventory definition in one place.
        result = capture.sources()
    except (OSError, subprocess.SubprocessError) as error:
        raise VerificationError(f"cannot enumerate current sources: {error}") from error
    require(isinstance(result, dict) and result, "live source manifest is empty")
    return result


def base_source_manifest() -> dict[str, str]:
    """Hash source files from the declared Git base without touching the tree."""

    command = [
        "git", "archive", "--format=tar", BASE_REVISION,
        "crates", "tools/perf-baseline", "Cargo.toml",
        "rust-toolchain.toml", ".cargo",
    ]
    try:
        process = subprocess.Popen(command, cwd=REPO, stdout=subprocess.PIPE,
                                   stderr=subprocess.PIPE)
    except OSError as error:
        raise VerificationError(f"cannot archive base sources: {error}") from error
    result: dict[str, str] = {}
    try:
        assert process.stdout is not None
        with tarfile.open(fileobj=process.stdout, mode="r|") as archive:
            for member in archive:
                if not member.isfile() or Path(member.name).suffix not in {".rs", ".toml", ".lock"}:
                    continue
                stream = archive.extractfile(member)
                require(stream is not None, f"cannot read base source {member.name}")
                value = hashlib.sha256()
                for block in iter(lambda: stream.read(1024 * 1024), b""):
                    value.update(block)
                result[member.name] = value.hexdigest()
        stderr = process.stderr.read().decode("utf-8", "replace")
        code = process.wait()
    except (OSError, tarfile.TarError) as error:
        process.kill()
        process.wait()
        raise VerificationError(f"cannot read base archive: {error}") from error
    require(code == 0, f"git archive failed: {stderr.strip()}")
    require(result, "base source archive is empty")
    return dict(sorted(result.items()))


def parse_time(value: Any, context: str) -> dt.datetime:
    require(isinstance(value, str), f"{context} is not an ISO timestamp")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise VerificationError(f"{context} is not an ISO timestamp: {error}") from error
    require(parsed.tzinfo is not None, f"{context} has no timezone")
    return parsed.astimezone(dt.timezone.utc)


def receipt_interval(receipt: dict[str, Any], context: str) -> tuple[float, float]:
    started = parse_time(receipt.get("started_utc"), f"{context}.started_utc")
    elapsed = receipt.get("elapsed_seconds")
    require(isinstance(elapsed, (int, float)) and not isinstance(elapsed, bool)
            and math.isfinite(float(elapsed)) and float(elapsed) > 0,
            f"{context}.elapsed_seconds is not positive and finite")
    start = started.timestamp()
    return start, start + float(elapsed)


def build_command() -> list[str]:
    return [
        "cargo", "build", "--release", "--locked", "-p", "litchi-docx",
        "--example", "managed_paragraph_batch_perf",
    ]


def verify_build(variant: str, intervals: list[tuple[float, float, str]]) -> dict[str, Any]:
    directory = HERE / variant
    receipt_path = directory / "build-receipt.json"
    log_path = directory / "build.log"
    source_path = directory / "source-manifest.json"
    source = manifest(source_path, f"{variant} source manifest")
    source_sha = digest(source_path)
    receipt = load_json(receipt_path)
    require(receipt.get("command") == build_command(), f"{variant} build command differs")
    require(receipt.get("exit_code") == 0, f"{variant} build failed")
    require(receipt.get("source_unchanged") is True, f"{variant} build changed sources")
    require(receipt.get("source_manifest_sha256") == source_sha,
            f"{variant} build source manifest binding differs")
    require(receipt.get("log_sha256") == digest(log_path),
            f"{variant} build log hash differs")
    binary_name = f"managed-paragraph-{variant}"
    expected_binary = SCRATCH / binary_name
    require(receipt.get("binary") == str(expected_binary),
            f"{variant} build binary path differs")
    binary_sha = check_digest(receipt.get("binary_sha256"), f"{variant} binary_sha256")
    binary = Path(receipt["binary"])
    if binary.exists():
        regular(binary, f"{variant} binary")
        require(digest(binary) == binary_sha, f"{variant} binary hash differs")
    symbols_path = directory / "symbols.txt"
    symbols_sha = None
    if symbols_path.exists():
        regular(symbols_path, f"{variant} symbol inventory")
        symbols = symbols_path.read_text(encoding="utf-8")
        require(sum(PROFILE_FUNCTION in line for line in symbols.splitlines()) == 1
                and sum(PROFILE_CALLER in line for line in symbols.splitlines()) == 1,
                f"{variant} symbol inventory omits or duplicates the profile boundary")
        symbols_sha = digest(symbols_path)
    # A final cleaned evidence bundle may deliberately remove the external
    # binary; all captures still retain its hash binding in their receipts.
    settings = receipt.get("environment")
    expected_settings = {
        "CARGO_TARGET_DIR": "/home/zhuhe/litchi-goal-0517-target",
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_DEBUG": "0",
        "CARGO_PROFILE_DEV_DEBUG": "0",
        "CARGO_PROFILE_TEST_DEBUG": "0",
        "TMPDIR": str(SCRATCH),
    }
    require(settings == expected_settings, f"{variant} build environment differs")
    intervals.append((*receipt_interval(receipt, f"{variant} build"), f"{variant}/build"))
    return {
        "source_manifest_sha256": source_sha,
        "source_files": len(source),
        "binary_sha256": binary_sha,
        "binary_present": binary.exists(),
        "symbols_sha256": symbols_sha,
        "log_sha256": digest(log_path),
        "receipt_sha256": digest(receipt_path),
    }


def expected_command(lane: str, name: str, variant: str, plan: dict[str, Any],
                     binary: str) -> list[str]:
    paragraphs, replacements, provider, mode = capture.prior.parse_name(name)
    samples, warmups, repeats, kind = LANE_SPECS[lane][1:]
    output = HERE / lane / f"{name}.csv"
    scratch = SCRATCH / "corpora" / f"{lane}-{name}"
    command = [
        binary, "--paragraphs", str(paragraphs), "--replacements", str(replacements),
        "--source", provider, "--mode", mode, "--samples", str(samples),
        "--warmups", str(warmups), "--repeats", str(repeats),
        "--artifact-dir", str(scratch), "--output", str(output),
    ]
    if kind == "profile":
        raw = HERE / lane / f"{name}.callgrind"
        command = [
            "valgrind", "--tool=callgrind", "--collect-atstart=no",
            "--zero-before=" + plan["profile"]["zero_before"],
            "--toggle-collect=" + plan["profile"]["toggle_collect"],
            "--callgrind-out-file=" + str(raw), *command,
        ]
    elif kind == "hardware":
        counters = HERE / lane / f"{name}.perf.csv"
        command = [
            "perf", "stat", "-x,", "-o", str(counters), "-e",
            ",".join(plan["hardware"]["events"]), "--", *command,
        ]
    return ["/usr/bin/time", "-v", "taskset", "-c", "2", *command]


def read_rows(path: Path) -> list[dict[str, str]]:
    regular(path, f"CSV {path}")
    try:
        with path.open(newline="", encoding="utf-8") as stream:
            rows = list(csv.DictReader(stream))
    except (OSError, UnicodeError, csv.Error) as error:
        raise VerificationError(f"cannot read CSV {path}: {error}") from error
    require(rows, f"empty CSV {path}")
    return rows


def verify_artifacts(directory: Path, name: str, receipt: dict[str, Any],
                     kind: str) -> dict[str, str]:
    suffixes = {"csv", "stdout", "stderr"}
    if kind == "profile":
        suffixes.add("callgrind")
    elif kind == "hardware":
        suffixes.add("perf.csv")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{directory.name}/{name}: artifact map missing")
    expected = {f"{name}.{suffix}" for suffix in suffixes}
    require(set(artifacts) == expected,
            f"{directory.name}/{name}: artifact inventory differs")
    checked: dict[str, str] = {}
    for relative, expected_sha in artifacts.items():
        path = safe_relative(relative, f"{directory.name}/{name} artifact")
        require(path.name == relative and len(path.parts) == 1,
                f"{directory.name}/{name}: artifact is not a basename")
        target = regular(
            directory / path,
            f"{directory.name}/{name} artifact {relative}",
            nonempty=not relative.endswith(".stdout"),
        )
        checked[relative] = check_digest(expected_sha,
                                         f"{directory.name}/{name}.artifacts.{relative}")
        require(digest(target) == checked[relative],
                f"{directory.name}/{name} artifact {relative} hash differs")
    return checked


def verify_profile_scope(path: Path) -> dict[str, Any]:
    try:
        summary, edges = profile_parser.parse_raw(path)
    except Exception as error:
        raise VerificationError(f"cannot parse {path.name}: {error}") from error
    matches = [
        edge for edge in edges
        if profile_parser.base_name(edge.parent) == PROFILE_CALLER
        and profile_parser.base_name(edge.child) == PROFILE_FUNCTION
    ]
    positive = [edge for edge in matches if edge.calls > 0]
    require(len(positive) == 1 and positive[0].calls == 1,
            f"{path.name}: publication edge is not exactly one positive call")
    require(positive[0].cost > 0, f"{path.name}: publication Ir cost is not positive")
    selected = [edge for edge in edges
                if profile_parser.base_name(edge.child) == PROFILE_FUNCTION
                and edge.calls > 0]
    require(len(selected) == 1 and selected[0] is positive[0],
            f"{path.name}: publication has an unexpected positive caller")
    require(summary == positive[0].cost,
            f"{path.name}: process Ir total is not scoped to publication")
    try:
        analyzed = profile_analyzer.parse_profile(path, PROFILE_FUNCTION)
    except Exception as error:
        raise VerificationError(f"{path.name}: profile analyzer replay failed: {error}") from error
    require(analyzed["validation"]["exactly_one_positive_call"] is True
            and analyzed["validation"]["selected_self_plus_direct_equals_inclusive"] is True
            and analyzed["validation"]["publication_total_scoped_to_summary"] is True,
            f"{path.name}: profile analyzer scope replay failed")
    return {
        "summary_ir": summary,
        "publication_calls": positive[0].calls,
        "publication_ir": positive[0].cost,
        "publication_caller": PROFILE_CALLER,
        "publication_function": PROFILE_FUNCTION,
    }


def verify_perf(path: Path, events: tuple[str, ...]) -> dict[str, dict[str, float | int | None]]:
    regular(path, f"perf counters {path}")
    values: dict[str, dict[str, float | int | None]] = {}
    try:
        with path.open(newline="", encoding="utf-8") as stream:
            for row in csv.reader(stream):
                if not row or row[0].startswith("#"):
                    continue
                require(len(row) >= 5, f"{path.name}: malformed perf row")
                event = row[2]
                require(event not in values, f"{path.name}: duplicate event {event}")
                raw = row[0].strip()
                if raw in {"<not counted>", "<not supported>", "-"}:
                    value: int | None = None
                else:
                    try:
                        value = int(raw.replace(",", ""))
                    except ValueError as error:
                        raise VerificationError(f"{path.name}: nonnumeric count {raw!r}") from error
                    require(value >= 0, f"{path.name}: negative count")
                try:
                    runtime = int(row[3].replace(",", ""))
                    coverage = float(row[4].rstrip("%"))
                except ValueError as error:
                    raise VerificationError(f"{path.name}: malformed runtime/coverage") from error
                require(runtime >= 0 and math.isfinite(coverage),
                        f"{path.name}: invalid runtime/coverage")
                values[event] = {"value": value, "runtime_ns": runtime,
                                 "coverage_percent": coverage}
    except (OSError, UnicodeError, csv.Error) as error:
        raise VerificationError(f"cannot read {path}: {error}") from error
    require(set(values) == set(events), f"{path.name}: event inventory differs")
    for event, value in values.items():
        if value["value"] is not None:
            require(float(value["coverage_percent"]) >= 99.9,
                    f"{path.name}: incomplete {event} coverage")
    return values


def verify_capture(lane: str, name: str, variant: str, plan: dict[str, Any],
                   build: dict[str, Any], intervals: list[tuple[float, float, str]]) -> dict[str, Any]:
    directory = HERE / lane
    path = directory / f"{name}.json"
    receipt = load_json(path)
    samples, warmups, repeats, kind = LANE_SPECS[lane][1:]
    require(receipt.get("exit_code") == 0, f"{lane}/{name}: child failed")
    require(receipt.get("source_unchanged") is True, f"{lane}/{name}: source changed")
    require(receipt.get("cleanup_verified") is True, f"{lane}/{name}: scratch cleanup failed")
    require(receipt.get("samples") == samples and receipt.get("warmups") == warmups
            and receipt.get("repeats") == repeats,
            f"{lane}/{name}: sample settings differ")
    require(receipt.get("source_manifest_sha256") == build["source_manifest_sha256"],
            f"{lane}/{name}: source binding differs from build")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{lane}/{name}: binary binding differs from build")
    require(receipt.get("plan_sha256") == digest(HERE / "plan.json"),
            f"{lane}/{name}: plan binding differs")
    if variant == "candidate":
        require(receipt.get("candidate_plan_sha256") == digest(HERE / "candidate-plan.json"),
                f"{lane}/{name}: candidate plan binding differs")
    require(receipt.get("scope") == plan["profile"]["scope"] if kind == "profile" else
            receipt.get("scope") == plan["hardware"]["scope"] if kind == "hardware" else
            receipt.get("scope") == plan["native"]["scope"],
            f"{lane}/{name}: scope differs from plan")
    binary = str(SCRATCH / f"managed-paragraph-{variant}")
    require(receipt.get("command") == expected_command(lane, name, variant, plan, binary),
            f"{lane}/{name}: command differs from frozen capture protocol")
    scratch = SCRATCH / "corpora" / f"{lane}-{name}"
    require(not scratch.exists(), f"{lane}/{name}: owned artifact directory remains")
    artifacts = verify_artifacts(directory, name, receipt, kind)
    rows = read_rows(directory / f"{name}.csv")
    try:
        capture.validate(name, rows, samples, warmups, repeats)
    except (AssertionError, KeyError, TypeError, ValueError) as error:
        raise VerificationError(f"{lane}/{name}: CSV oracle validation failed: {error}") from error
    identity = tuple(rows[0][key] for key in capture.prior.IDENTITY_KEYS)
    require(len({tuple(row[key] for key in capture.prior.IDENTITY_KEYS) for row in rows}) == 1,
            f"{lane}/{name}: output identity changes")
    profile = None
    counters = None
    if kind == "profile":
        profile = verify_profile_scope(directory / f"{name}.callgrind")
    elif kind == "hardware":
        counters = verify_perf(directory / f"{name}.perf.csv", tuple(plan["hardware"]["events"]))
    intervals.append((*receipt_interval(receipt, f"{lane}/{name}"), f"{lane}/{name}"))
    return {
        "identity": identity,
        "artifact_sha256": artifacts,
        "receipt_sha256": digest(path),
        "profile": profile,
        "hardware": counters,
    }


def verify_lane(lane: str, variant: str, plan: dict[str, Any], build: dict[str, Any],
                intervals: list[tuple[float, float, str]]) -> dict[str, Any]:
    directory = HERE / lane
    require(directory.is_dir(), f"missing lane directory: {lane}")
    names, _samples, _warmups, _repeats, kind = LANE_SPECS[lane]
    csv_names = {
        path.name.removesuffix(".csv")
        for path in directory.glob("*.csv")
        if not path.name.endswith(".perf.csv")
    }
    require(csv_names == names, f"{lane}: CSV inventory differs")
    for name in names:
        regular(directory / f"{name}.json", f"{lane}/{name} receipt")
        regular(directory / f"{name}.stdout", f"{lane}/{name} stdout", nonempty=False)
        regular(directory / f"{name}.stderr", f"{lane}/{name} stderr")
        if kind == "profile":
            regular(directory / f"{name}.callgrind", f"{lane}/{name} Callgrind")
        elif kind == "hardware":
            regular(directory / f"{name}.perf.csv", f"{lane}/{name} perf counters")
    result = {name: verify_capture(lane, name, variant, plan, build, intervals)
              for name in sorted(names)}
    # Repeated and batch routes for the same deterministic fixture must emit
    # the same output identity.
    for stem in {name.rsplit("-", 1)[0] for name in names}:
        left = result.get(f"{stem}-repeated")
        right = result.get(f"{stem}-batch")
        if left is not None and right is not None:
            require(left["identity"] == right["identity"],
                    f"{lane}/{stem}: route output identities differ")
    return result


def verify_timeline(intervals: list[tuple[float, float, str]]) -> list[dict[str, Any]]:
    ordered = sorted(intervals, key=lambda item: (item[0], item[1], item[2]))
    for previous, current in zip(ordered, ordered[1:]):
        require(current[0] >= previous[1],
                f"capture/build intervals overlap: {previous[2]} and {current[2]}")
    return [{"label": label, "start": start, "end": end}
            for start, end, label in ordered]


def verify_order(lane: str, records: dict[str, Any], plan: dict[str, Any]) -> None:
    if lane not in {"r1", "r2", "profile-r1", "profile-r2",
                    "after-r1", "after-r2", "profile-after-r1", "profile-after-r2"}:
        return
    names = (list(plan["profile"]["cases"])
             if lane.startswith("profile") else sorted(records))
    if lane.endswith("r2"):
        names.reverse()
    # Receipt start times are stored in the record only through the timeline,
    # so use file mtime as a tie-resistant fallback only when starts are equal.
    starts = []
    for name in records:
        receipt = load_json(HERE / lane / f"{name}.json")
        starts.append((parse_time(receipt["started_utc"], f"{lane}/{name}"), name))
    observed = [name for _time, name in sorted(starts)]
    require(observed == names, f"{lane}: capture order differs from frozen plan")


def verify_manifest_epoch(variant: str) -> tuple[dict[str, str], str]:
    path = HERE / variant / "source-manifest.json"
    value = manifest(path, f"{variant} source manifest")
    if variant == "baseline":
        require(value == base_source_manifest(),
                "baseline source manifest differs from declared Git base")
    else:
        require(value == live_source_manifest(),
                "candidate source manifest differs from current source tree")
    return value, digest(path)


def verify_candidate_diff(baseline: dict[str, str], candidate: dict[str, str],
                          candidate_plan: dict[str, Any]) -> dict[str, Any]:
    changed = set(baseline) ^ set(candidate)
    changed |= {name for name in baseline.keys() & candidate.keys()
                if baseline[name] != candidate[name]}
    explicit = None
    for key in ("intended_source_paths", "allowed_source_paths", "changed_paths"):
        value = candidate_plan.get(key)
        if value is not None:
            require(isinstance(value, list) and all(isinstance(item, str) for item in value),
                    f"candidate plan {key} is not a path list")
            explicit = {safe_relative(item, f"candidate plan {key}").as_posix() for item in value}
            break
    if explicit is None:
        require(changed == DEFAULT_CANDIDATE_PATHS,
                f"candidate source diff differs from the frozen two-file policy: {sorted(changed)}")
        expected = sorted(DEFAULT_CANDIDATE_PATHS)
        policy = "frozen candidate-plan two-file OPC production/test policy"
    else:
        require(changed == explicit,
                f"candidate source diff differs from declared paths: changed={sorted(changed)} declared={sorted(explicit)}")
        expected = sorted(explicit)
        policy = "candidate-plan.json intended_source_paths"
    return {"changed_paths": sorted(changed), "allowed_paths": expected, "policy": policy}


def verify_source_patch(required: bool = False) -> dict[str, Any] | None:
    """Bind the retained patch to the current tree relative to the base."""

    path = HERE / "candidate" / "source.patch"
    if not path.exists():
        require(not required, "candidate source patch is missing")
        return None
    regular(path, "candidate source patch")
    try:
        process = subprocess.run(
            [
                "git", "diff", "--no-ext-diff", "--no-color", BASE_REVISION,
                "--", SOURCE_PRODUCTION, SOURCE_TEST,
            ],
            cwd=REPO,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
    except OSError as error:
        raise VerificationError(f"cannot inspect candidate source patch: {error}") from error
    require(process.returncode == 0,
            "git diff failed while binding candidate source patch")
    retained = path.read_bytes()
    require(retained == process.stdout,
            "candidate source patch differs from the current base-to-candidate diff")
    return {
        "sha256": digest(path),
        "bytes": len(retained),
        "matches_base_diff": True,
        "paths": [SOURCE_PRODUCTION, SOURCE_TEST],
    }


def verify_helper_manifest(required: bool = False) -> dict[str, Any] | None:
    path = HERE / "helper-manifest.json"
    if not path.exists():
        require(not required, "helper manifest is missing")
        return None
    value = manifest(path, "helper manifest")
    require(set(value) == HELPER_PATHS,
            "helper manifest dependency inventory differs")
    for relative, expected in value.items():
        target = REPO / safe_relative(relative, "helper manifest path")
        require(digest(target) == expected,
                f"helper manifest hash differs for {relative}")
    return {"sha256": digest(path), "entries": len(value)}


def verify_hardware_summary(plan: dict[str, Any]) -> dict[str, Any] | None:
    path = HERE / "hardware-summary.json"
    try:
        replay = hardware_analyzer.analyze()
    except Exception as error:
        raise VerificationError(f"hardware analyzer replay failed: {error}") from error
    if not path.exists():
        return {"replayed": True, "cases": len(replay["cases"]),
                "summary_retained": False}
    summary = load_json(path)
    require(summary == replay, "hardware summary differs from analyzer replay")
    require(summary.get("scope") == "Whole child baseline only; includes fixture/preflight and output oracles; not publication attribution or candidate comparison",
            "hardware summary scope differs")
    cases = summary.get("cases")
    require(isinstance(cases, list), "hardware summary cases missing")
    observed = {}
    for item in cases:
        require(isinstance(item, dict) and isinstance(item.get("case"), str),
                "hardware summary case malformed")
        require(item["case"] not in observed, "hardware summary has duplicate case")
        observed[item["case"]] = item
    require(set(observed) == PROFILE_NAMES, "hardware summary case inventory differs")
    for name, item in observed.items():
        counters = verify_perf(HERE / "hardware" / f"{name}.perf.csv", tuple(plan["hardware"]["events"]))
        expected = item.get("events")
        require(isinstance(expected, dict) and set(expected) == set(EVENTS),
                f"hardware summary {name}: event inventory differs")
        for event in EVENTS:
            for field in ("value", "runtime_ns", "coverage_percent"):
                actual = counters[event][field]
                wanted = expected[event].get(field)
                if isinstance(actual, float):
                    require(math.isclose(float(actual), float(wanted), rel_tol=1e-12, abs_tol=1e-9),
                            f"hardware summary {name}/{event}/{field} differs")
                else:
                    require(actual == wanted, f"hardware summary {name}/{event}/{field} differs")
        cycles = counters["cycles"]["value"]
        instructions = counters["instructions"]["value"]
        require(cycles and instructions, f"hardware summary {name}: IPC inputs unavailable")
        require(math.isclose(float(item.get("IPC")), instructions / cycles,
                             rel_tol=1e-12, abs_tol=1e-12),
                f"hardware summary {name}: IPC differs")
    return {"sha256": digest(path), "cases": len(cases)}


def recompute_native_report(lanes: list[str]) -> dict[str, Any]:
    campaigns = {lane: native_analyzer.load_campaign(lane) for lane in lanes}
    bindings = {tuple(campaign["binding"].values()) for campaign in campaigns.values()}
    require(len(bindings) == 1, "native report campaigns do not share one binding")
    first = campaigns[lanes[0]]
    seed = native_analyzer.DEFAULT_SEED
    bootstraps = native_analyzer.DEFAULT_BOOTSTRAPS
    return {
        "schema": "managed_paragraph_native_analysis_v1",
        "campaigns": lanes,
        "seed": seed,
        "bootstrap_iterations": bootstraps,
        "thresholds_percent": {"p50": 5, "mean": 5, "p95": 10, "p99": 15, "rss": 5},
        "scope": "native r1/r2 (or explicitly supplied equivalent lanes), measured warm=false rows only; phase clocks include returned Snapshot drop in publish_ns; RSS is whole-child",
        "binding": first["binding"],
        "campaign_data": {
            lane: {
                "binding": campaigns[lane]["binding"],
                "cases": {
                    name: native_analyzer.compact_case(case)
                    for name, case in campaigns[lane]["cases"].items()
                },
            }
            for lane in lanes
        },
        "api_choice": {
            lane: native_analyzer.api_choice(campaigns[lane], seed, bootstraps)
            for lane in lanes
        },
        "cross_campaign": native_analyzer.cross_campaign(
            campaigns[lanes[0]], campaigns[lanes[1]], seed, bootstraps
        ),
    }


def recompute_candidate_comparison() -> dict[str, Any]:
    baseline_lanes = ("r1", "r2")
    candidate_lanes = ("after-r1", "after-r2")
    native = candidate_comparator.native
    baseline = {lane: native.load_campaign(lane) for lane in baseline_lanes}
    candidate = {lane: native.load_campaign(lane) for lane in candidate_lanes}
    for lane, campaign in (*baseline.items(), *candidate.items()):
        candidate_comparator.verify_anchor(lane, campaign)
    require(baseline["r1"]["binding"] == baseline["r2"]["binding"],
            "candidate comparison baseline bindings differ")
    require(candidate["after-r1"]["binding"] == candidate["after-r2"]["binding"],
            "candidate comparison candidate bindings differ")
    all_campaigns = [baseline["r1"], baseline["r2"], candidate["after-r1"], candidate["after-r2"]]
    identities = candidate_comparator.check_output_identity(all_campaigns)
    pairs = [
        {"id": "r1", "baseline": "r1", "candidate": "after-r1"},
        {"id": "r2", "baseline": "r2", "candidate": "after-r2"},
    ]
    seed = native.DEFAULT_SEED
    bootstraps = native.DEFAULT_BOOTSTRAPS
    comparisons = {
        pair["id"]: candidate_comparator.compare_pair(
            baseline[pair["baseline"]], candidate[pair["candidate"]],
            pair["baseline"], pair["candidate"], seed, bootstraps
        )
        for pair in pairs
    }
    flag_counts = Counter(
        f"{flag['metric']}.{flag['stat']}"
        for records in comparisons.values()
        for record in records
        for flag in record["adverse_flags"]
    )
    return {
        "schema": "managed_paragraph_native_candidate_comparison_v1",
        "pairs": pairs,
        "seed": seed,
        "bootstrap_iterations": bootstraps,
        "thresholds_percent": {"p50": 5, "mean": 5, "p95": 10, "p99": 15, "rss": 5},
        "scope": "same-API native baseline/candidate comparison; 24 cases per pair; warm=false rows only; publish_ns includes returned Snapshot drop; RSS is whole-child",
        "bindings": {
            "baseline": {
                lane: candidate_comparator.public_binding(baseline[lane])
                for lane in baseline_lanes
            },
            "candidate": {
                lane: candidate_comparator.public_binding(candidate[lane])
                for lane in candidate_lanes
            },
        },
        "output_identities": identities,
        "comparisons": comparisons,
        "adverse_flag_counts": dict(sorted(flag_counts.items())),
    }


def recompute_profile_comparison() -> dict[str, Any]:
    baseline = profile_comparator.load_variant("baseline")
    candidate = profile_comparator.load_variant("candidate")
    comparisons = profile_comparator.compare(baseline, candidate)
    validation = {
        "baseline_profile_count": len(baseline),
        "candidate_profile_count": len(candidate),
        "same_api_for_all_pairs": all(record["same_api"] for record in comparisons),
        "baseline_exact_one_positive_call": all(
            profile["annotation_checks"]["raw_exactly_one_positive_call"]
            and profile["publication_calls"] == 1
            for profile in baseline.values()
        ),
        "candidate_exact_one_positive_call": all(
            profile["annotation_checks"]["raw_exactly_one_positive_call"]
            and profile["publication_calls"] == 1
            for profile in candidate.values()
        ),
        "all_annotation_checks": all(
            all(profile["annotation_checks"].values())
            for profile in (*baseline.values(), *candidate.values())
        ),
    }
    return {
        "schema": "managed_docx_publication_profile_comparison_v1",
        "matrix": {
            "arms": list(profile_comparator.ARMS),
            "repeats": ["r1", "r2"],
            "profiles_per_variant": len(baseline),
            "matched_pairs": len(comparisons),
        },
        "scope": "Callgrind Ir publication-method diagnostic; owned source; one measured sample; zero warmups; one harness repeat; publication excludes caller Snapshot drop; no RSS claim",
        "functions": {
            "publication": profile_comparator.PUBLICATION,
            "direct_owner": profile_comparator.TOPOLOGY,
            "topology": profile_comparator.TOPOLOGY,
            "xml_validator": profile_comparator.XML_VALIDATOR,
        },
        "bindings": {
            "baseline": profile_comparator.binding(baseline),
            "candidate": profile_comparator.binding(candidate),
        },
        "validation": validation,
        "comparisons": comparisons,
    }


def verify_analyses(plan: dict[str, Any], captures: dict[str, Any],
                    required: bool = False) -> dict[str, Any]:
    checked: dict[str, Any] = {}
    paths = [HERE / "native-analysis.json", HERE / "profile-analysis.json"]
    paths.extend(sorted(HERE.glob("profile-*/profile-analysis.json")))
    paths.extend(sorted(HERE.glob("*native-analysis*.json")))
    required_reports = {
        HERE / "native-analysis.json",
        HERE / "candidate-native-analysis.json",
        HERE / "profile-analysis.json",
        HERE / "profile-after-analysis.json",
        HERE / "candidate-comparison.json",
        HERE / "profile-comparison.json",
        HERE / "counter-comparison.json",
    }
    if required:
        for report_path in required_reports:
            require(report_path.is_file(), f"required analysis report is missing: {report_path.name}")
    for path in dict.fromkeys(paths):
        if not path.exists():
            continue
        report = load_json(path)
        require(isinstance(report, dict), f"analysis report {path.name} is not an object")
        if "native-analysis" in path.name:
            require(report.get("schema") == "managed_paragraph_native_analysis_v1",
                    "native analysis schema differs")
            campaigns = report.get("campaigns")
            require(isinstance(campaigns, list) and len(campaigns) == 2,
                    "native analysis campaign binding differs")
            require(all(isinstance(lane, str) and (HERE / lane).is_dir()
                        for lane in campaigns),
                    "native analysis references an absent campaign")
            require(isinstance(report.get("campaign_data"), dict),
                    "native analysis campaign data missing")
            for lane in campaigns:
                cases = report["campaign_data"].get(lane, {}).get("cases")
                require(isinstance(cases, dict) and set(cases) == NATIVE_NAMES,
                        f"native analysis {lane} case inventory differs")
            expected_lanes = ["after-r1", "after-r2"] if path.name.startswith("candidate-") else ["r1", "r2"]
            try:
                replay = recompute_native_report(expected_lanes)
            except Exception as error:
                raise VerificationError(f"native analysis replay failed for {path.name}: {error}") from error
            require(report == replay,
                    f"native analysis {path.name} differs from analyzer replay")
        else:
            require(report.get("schema") == "docx_callgrind_publication_analysis_v1",
                    "profile analysis schema differs")
            require(report.get("selected_function") == PROFILE_FUNCTION,
                    "profile analysis function differs")
            profiles = report.get("profiles")
            require(isinstance(profiles, list) and report.get("profile_count") == len(profiles),
                    "profile analysis count differs")
            require(profiles, "profile analysis is empty")
            for item in profiles:
                require(item.get("selected_function") == PROFILE_FUNCTION,
                        "profile analysis selected function differs")
                raw = Path(item.get("path", ""))
                if not raw.is_absolute():
                    raw = (REPO / raw).resolve()
                regular(raw, "profile analysis raw profile")
                require(item.get("sha256") == digest(raw),
                        f"profile analysis hash differs for {raw.name}")
                require(item.get("validation", {}).get("exactly_one_positive_call") is True
                        and item.get("validation", {}).get("publication_total_scoped_to_summary") is True,
                        f"profile analysis validation failed for {raw.name}")
        checked[path.relative_to(HERE).as_posix()] = {"sha256": digest(path)}
    for name, replay_function in (
        ("candidate-comparison.json", recompute_candidate_comparison),
        ("profile-comparison.json", recompute_profile_comparison),
    ):
        path = HERE / name
        if not path.exists():
            continue
        report = load_json(path)
        try:
            replay = replay_function()
        except Exception as error:
            raise VerificationError(f"comparison replay failed for {name}: {error}") from error
        require(report == replay, f"comparison report {name} differs from analyzer replay")
        checked[name] = {"sha256": digest(path), "replayed": True}
    counter_path = HERE / "counter-comparison.json"
    if counter_path.exists():
        report = load_json(counter_path)
        require(isinstance(report, dict), "counter comparison is not an object")
        try:
            replay = counter_analyzer.analyze()
        except Exception as error:
            raise VerificationError(f"counter analyzer replay failed: {error}") from error
        require(report == replay, "counter comparison differs from analyzer replay")
        require(report.get("rows_checked") == 3168
                and isinstance(report.get("cases"), list)
                and len(report["cases"]) == 48,
                "counter comparison inventory differs")
        checked["counter-comparison.json"] = {
            "sha256": digest(counter_path),
            "replayed": True,
            "rows_checked": report["rows_checked"],
        }
    return checked


def verify_sha256sums() -> dict[str, Any] | None:
    path = HERE / "SHA256SUMS"
    if not path.exists():
        return None
    regular(path, "SHA256SUMS")
    entries: dict[str, str] = {}
    evidence_root = Path("docs/performance/results/change-0517")
    for number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        parts = line.split(maxsplit=1)
        require(len(parts) == 2, f"SHA256SUMS line {number} is malformed")
        expected = check_digest(parts[0], f"SHA256SUMS line {number}")
        name = parts[1].lstrip(" *")
        relative = safe_relative(name, f"SHA256SUMS line {number} path")
        if relative == evidence_root or relative.as_posix().startswith(
                evidence_root.as_posix() + "/"):
            target = REPO / relative
            key = relative.relative_to(evidence_root).as_posix()
        else:
            target = HERE / relative
            key = relative.as_posix()
        require(key not in entries, f"SHA256SUMS duplicates {key}")
        require(digest(target) == expected, f"SHA256SUMS hash differs for {key}")
        entries[key] = expected
    require(entries, "SHA256SUMS is empty")
    actual: set[str] = set()
    for leaf in HERE.rglob("*"):
        if leaf.is_symlink():
            raise VerificationError(f"SHA256SUMS evidence tree contains symlink: {leaf}")
        if leaf.is_file():
            relative = leaf.relative_to(HERE).as_posix()
            if relative != "SHA256SUMS":
                actual.add(relative)
    require(set(entries) == actual,
            "SHA256SUMS inventory differs from retained evidence leaves")
    return {"entries": len(entries)}


def verify_zip64_generator(log_path: Path) -> dict[str, Any]:
    """Validate the retained independent ZIP64 generator report.

    The report is retained after the temporary corpus is cleaned.  When the
    corpus is still available, bind both its byte count and digest to the
    report as an additional tamper check.
    """

    regular(log_path, "ZIP64 generator log")
    report = load_json(log_path)
    require(isinstance(report, dict), "ZIP64 generator log is not an object")
    require(set(report) == {
        "archive_bytes", "archive_sha256", "buffer_bytes", "entries", "python", "zlib",
    }, "ZIP64 generator report fields differ")
    archive_bytes = report.get("archive_bytes")
    require(isinstance(archive_bytes, int) and not isinstance(archive_bytes, bool)
            and 0 < archive_bytes <= ZIP64_PHYSICAL_LIMIT,
            "ZIP64 generator physical archive size is out of bounds")
    require(archive_bytes < ZIP64_DECLARED_BYTES,
            "ZIP64 generator archive is not physically bounded below the declared size")
    archive_sha = check_digest(report.get("archive_sha256"),
                               "ZIP64 generator archive_sha256")
    require(report.get("buffer_bytes") == ZIP64_BUFFER_BYTES,
            "ZIP64 generator buffer size differs")
    require(isinstance(report.get("python"), str) and report["python"],
            "ZIP64 generator Python version is missing")
    require(isinstance(report.get("zlib"), str) and report["zlib"],
            "ZIP64 generator zlib version is missing")

    entries = report.get("entries")
    require(isinstance(entries, list) and len(entries) == len(ZIP64_ENTRIES),
            "ZIP64 generator entry count differs")
    require(tuple(item.get("name") for item in entries if isinstance(item, dict))
            == ZIP64_ENTRIES, "ZIP64 generator entry inventory differs")
    for item in entries:
        require(isinstance(item, dict) and set(item) == {
            "compressed_bytes", "crc32", "flags", "name", "sha256",
            "uncompressed_bytes", "version_needed",
        }, "ZIP64 generator entry fields differ")
        for field in ("compressed_bytes", "crc32", "flags", "uncompressed_bytes", "version_needed"):
            value = item[field]
            require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
                    f"ZIP64 generator {item['name']}/{field} is invalid")
        require(item["compressed_bytes"] <= archive_bytes,
                f"ZIP64 generator {item['name']} compressed size exceeds archive")
        check_digest(item["sha256"], f"ZIP64 generator {item['name']} sha256")
    large = entries[-1]
    require(large["uncompressed_bytes"] == ZIP64_DECLARED_BYTES,
            "ZIP64 generator large entry does not declare 4 GiB")
    require(large["compressed_bytes"] > 0 and large["compressed_bytes"] < archive_bytes,
            "ZIP64 generator large entry is not physically compressed")
    require(large["version_needed"] >= 45,
            "ZIP64 generator large entry is not ZIP64")

    corpus = SCRATCH / "independent-zip64.zip"
    present = corpus.exists()
    if present:
        regular(corpus, "ZIP64 generator corpus")
        require(corpus.stat().st_size == archive_bytes,
                "ZIP64 generator corpus size differs from report")
        require(digest(corpus) == archive_sha,
                "ZIP64 generator corpus hash differs from report")
    return {
        "archive_bytes": archive_bytes,
        "archive_sha256": archive_sha,
        "declared_large_bytes": large["uncompressed_bytes"],
        "physical_limit_bytes": ZIP64_PHYSICAL_LIMIT,
        "corpus_present": present,
    }


def verify_quality_receipts(candidate_source_sha: str,
                            intervals: list[tuple[float, float, str]],
                            required: bool = False) -> dict[str, Any] | None:
    directory = HERE / "checks"
    if not directory.exists():
        require(not required, "candidate quality-gate receipts are missing")
        return None
    expected = set(QUALITY_COMMANDS)
    receipts = {path.stem for path in directory.glob("*.json")}
    if not receipts:
        require(not required, "candidate quality-gate receipts are missing")
        return {"status": "pending", "missing": sorted(expected), "checked": {}}
    extra = receipts - expected
    require(not extra, f"quality receipt inventory has unexpected files: {sorted(extra)}")
    missing = expected - receipts
    require(not (missing and required),
            "candidate quality-gate receipts are incomplete: " + ", ".join(sorted(missing)))
    result = {}
    for name in sorted(receipts):
        receipt_path = directory / f"{name}.json"
        receipt = load_json(receipt_path)
        require(isinstance(receipt, dict), f"quality check {name} receipt is not an object")
        require(receipt.get("command") == QUALITY_COMMANDS[name],
                f"quality check {name} command differs")
        log_path = directory / f"{name}.log"
        require(receipt.get("exit_code") == 0 and receipt.get("source_unchanged") is True,
                f"quality check {name} failed or changed source")
        require(receipt.get("source_manifest_sha256") == candidate_source_sha,
                f"quality check {name} source binding differs")
        require(receipt.get("log_sha256") == digest(log_path),
                f"quality check {name} log hash differs")
        environment = receipt.get("environment")
        expected_environment = dict(QUALITY_ENVIRONMENT)
        if name in {"zip64-generate", "zip64-test", "claims"}:
            expected_environment["LITCHI_0415_PYTHON_ZIP"] = str(
                SCRATCH / "independent-zip64.zip"
            )
        require(environment == expected_environment,
                f"quality check {name} environment differs")
        intervals.append((*receipt_interval(receipt, f"checks/{name}"), f"checks/{name}"))
        checked = {"receipt_sha256": digest(receipt_path), "log_sha256": digest(log_path)}
        if name == "zip64-generate":
            checked["generator"] = verify_zip64_generator(log_path)
        result[name] = checked
    if missing:
        return {"status": "pending", "missing": sorted(missing), "checked": result}
    return {"status": "pass", "checked": result}


def capture_proof(lanes: dict[str, dict[str, Any]]) -> dict[str, Any]:
    return {
        lane: {
            name: {
                "receipt_sha256": item["receipt_sha256"],
                "artifact_sha256": item["artifact_sha256"],
            }
            for name, item in sorted(records.items())
        }
        for lane, records in sorted(lanes.items())
    }


def replay_native_analyzer(lanes: list[str]) -> list[str]:
    for lane in lanes:
        try:
            native_analyzer.load_campaign(lane)
        except Exception as error:
            raise VerificationError(f"native analyzer replay failed for {lane}: {error}") from error
    return lanes


def verify_bundle(require_complete: bool = False) -> dict[str, Any]:
    plan = load_json(HERE / "plan.json")
    require(plan.get("base_revision") == BASE_REVISION, "plan base revision differs")
    candidate_plan = load_json(HERE / "candidate-plan.json")
    intervals: list[tuple[float, float, str]] = []
    baseline_manifest, baseline_manifest_sha = verify_manifest_epoch("baseline")
    baseline_build = verify_build("baseline", intervals)
    baseline = {}
    for lane in BASELINE_LANES:
        baseline[lane] = verify_lane(lane, "baseline", plan, baseline_build, intervals)
        verify_order(lane, baseline[lane], plan)
    replayed_native = replay_native_analyzer(["r1", "r2"])

    candidate_dirs = [HERE / lane for lane in CANDIDATE_LANES]
    candidate_complete = all(path.is_dir() for path in candidate_dirs)
    candidate = None
    candidate_build = None
    candidate_manifest_sha = None
    diff = None
    source_patch = None
    quality = None
    if any(path.exists() for path in candidate_dirs) or (HERE / "candidate").exists():
        candidate = {}
        candidate_manifest, candidate_manifest_sha = verify_manifest_epoch("candidate")
        diff = verify_candidate_diff(baseline_manifest, candidate_manifest, candidate_plan)
        source_patch = verify_source_patch(required=require_complete)
        candidate_build = verify_build("candidate", intervals)
        for lane in CANDIDATE_LANES:
            if not (HERE / lane).is_dir():
                continue
            candidate[lane] = verify_lane(lane, "candidate", plan, candidate_build, intervals)
            verify_order(lane, candidate[lane], plan)
        candidate_complete = candidate_complete and all(lane in candidate for lane in CANDIDATE_LANES)
        if candidate_complete:
            for name in NATIVE_NAMES:
                before = baseline["r1"][name]["identity"]
                require(candidate["after-r1"][name]["identity"] == before
                        and candidate["after-r2"][name]["identity"] == before,
                        f"{name}: candidate output identity differs from baseline")
            quality = verify_quality_receipts(
                candidate_manifest_sha, intervals, required=require_complete
            )
            replayed_native.extend(replay_native_analyzer(["after-r1", "after-r2"]))
    if not candidate_complete and require_complete:
        missing = [lane for lane in CANDIDATE_LANES if not (HERE / lane).is_dir()]
        raise VerificationError("candidate evidence is incomplete: " + ", ".join(missing))

    timeline = verify_timeline(intervals)
    hardware_summary = verify_hardware_summary(plan)
    analyses = verify_analyses(plan, baseline,
                               required=require_complete and candidate_complete)
    helper_manifest = verify_helper_manifest(required=require_complete and candidate_complete)
    sums = verify_sha256sums()
    status = "pass" if candidate_complete else "pending-candidate"
    if candidate_complete and quality is None:
        status = "pending-quality"
    elif candidate_complete and quality.get("status") != "pass":
        status = "pending-quality"
    elif candidate_complete and (source_patch is None or helper_manifest is None):
        status = "pending-quality"
    return {
        "schema": "litchi-0517-verifier-v1",
        "status": status,
        "base_revision": BASE_REVISION,
        "priority": "OLE2/OOXML; ODF deferred; iWork excluded",
        "baseline": {"build": baseline_build,
                      "lanes": {lane: len(value) for lane, value in baseline.items()},
                      "source_manifest_sha256": baseline_manifest_sha,
                      "capture_receipts": capture_proof(baseline)},
        "candidate": None if candidate is None else {
            "build": candidate_build,
            "lanes": {lane: len(value) for lane, value in candidate.items()},
            "source_manifest_sha256": candidate_manifest_sha,
            "diff": diff,
            "source_patch": source_patch,
            "capture_receipts": capture_proof(candidate),
        },
        "candidate_required_lanes": list(CANDIDATE_LANES),
        "candidate_complete": candidate_complete,
        "quality_checks": quality,
        "analyzer_replays": replayed_native,
        "hardware_summary": hardware_summary,
        "analyses": analyses,
        "helper_manifest": helper_manifest,
        "timeline": timeline,
        "sha256sums": sums,
        "scope": plan["native"]["scope"],
        "claims": plan["claims"],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--require-complete", action="store_true",
                        help="fail when candidate lanes have not all been retained")
    parser.add_argument("--output", type=Path, default=HERE / "verification.json")
    args = parser.parse_args()
    try:
        result = verify_bundle(args.require_complete)
    except (VerificationError, OSError, KeyError, ValueError, subprocess.SubprocessError) as error:
        print(f"verification failed: {error}", file=sys.stderr)
        return 2
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n",
                           encoding="utf-8")
    print(json.dumps({"status": result["status"],
                      "candidate_complete": result["candidate_complete"],
                      "timeline_events": len(result["timeline"])}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
