#!/usr/bin/env python3
"""Fail-closed verifier for the source-bound 0519 DOCX evidence bundle.

This verifier only reads retained evidence.  Row-level CSV invariants are
delegated to ``capture.validate`` and the raw Callgrind grammar to the 0515
parser.  A missing candidate epoch is reported as pending; ``--require-complete``
turns that state into a failure for the final seal.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import difflib
import hashlib
import importlib.util
import json
import math
import re
import subprocess
import sys
import tarfile
import tempfile
from collections import Counter
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
SCRATCH = Path("/tmp/litchi-goal-0519")
BASE_REVISION = "45d71cb6f0cd5d003c544b61c748552a513b0245"
SOURCE_PRODUCTION = "crates/litchi-opc/src/xml_splice.rs"
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
    "alloc-baseline-r1": (NATIVE_NAMES, 1, 0, 1, "allocation"),
    "alloc-baseline-r2": (NATIVE_NAMES, 1, 0, 1, "allocation"),
    "after-r1": (NATIVE_NAMES, 30, 3, 2, "native"),
    "after-r2": (NATIVE_NAMES, 30, 3, 2, "native"),
    "profile-after-r1": (PROFILE_NAMES, 1, 0, 1, "profile"),
    "profile-after-r2": (PROFILE_NAMES, 1, 0, 1, "profile"),
    "hardware-after": (PROFILE_NAMES, 30, 3, 2, "hardware"),
    "alloc-candidate-r1": (NATIVE_NAMES, 1, 0, 1, "allocation"),
    "alloc-candidate-r2": (NATIVE_NAMES, 1, 0, 1, "allocation"),
}
BASELINE_LANES = (
    "preflight", "r1", "r2", "profile-preflight", "profile-r1",
    "profile-r2", "hardware", "alloc-baseline-r1", "alloc-baseline-r2",
)
CANDIDATE_LANES = (
    "after-r1", "after-r2", "profile-after-r1", "profile-after-r2",
    "hardware-after", "alloc-candidate-r1", "alloc-candidate-r2",
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
    "CARGO_TARGET_DIR": "/home/zhuhe/litchi-goal-0519-target",
    "CARGO_BUILD_JOBS": "2",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_DEBUG": "0",
    "CARGO_PROFILE_DEV_DEBUG": "0",
    "CARGO_PROFILE_TEST_DEBUG": "0",
    "TMPDIR": str(SCRATCH),
    "RUSTDOCFLAGS": "-D warnings",
    "LITCHI_0415_PYTHON_ZIP": str(SCRATCH / "independent-zip64.zip"),
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
TAIL_GUARD_PLAN = HERE / "tail-guard-plan.json"
TAIL_GUARD_DIR = HERE / "tail-guard"
TAIL_GUARD_REPORT = HERE / "tail-guard.json"
TAIL_GUARD_MARKDOWN = HERE / "tail-guard.md"
TAIL_GUARD_SCRIPT = HERE / "tail_guard.py"
TAIL_GUARD_PLAN_SCHEMA = "managed_paragraph_tail_guard_plan_0519_v2"
TAIL_GUARD_CAPTURE_SCHEMA = "managed_paragraph_tail_guard_capture_0519_v2"
TAIL_GUARD_REPORT_SCHEMA = "managed_paragraph_tail_guard_0519_v2"
TAIL_GUARD_CASES = {
    "1": "p128-k1-file-batch",
    "2": "p128-k1-owned-batch",
}
TAIL_GUARD_ORDER = (
    ("A1", "baseline", "1"),
    ("A1", "baseline", "2"),
    ("B1", "candidate", "1"),
    ("B1", "candidate", "2"),
    ("B2", "candidate", "2"),
    ("B2", "candidate", "1"),
    ("A2", "baseline", "2"),
    ("A2", "baseline", "1"),
)
TAIL_GUARD_PAIR_ORDER = (("A1", "B1"), ("A2", "B2"))
TAIL_GUARD_PHASES = (
    "elapsed_ns", "open_ns", "edit_ns", "commit_ns", "publish_ns", "drop_ns",
)


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
        spec = importlib.util.spec_from_file_location("change0519_capture_verify", HERE / "capture.py")
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
    spec = importlib.util.spec_from_file_location("change0515_parser_for_0519", path)
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
        "change0519_" + name.replace(".", "_"), path
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


def load_optional_analyzer(path: Path, name: str) -> Any | None:
    """Defer optional replay dependencies until their evidence is present."""

    return load_analyzer(path, name) if path.exists() else None


hardware_analyzer = load_optional_analyzer(HERE / "analyze_hardware.py", "analyze_hardware.py")
counter_analyzer = load_optional_analyzer(HERE / "analyze_counters.py", "analyze_counters.py")
allocation_analyzer = load_optional_analyzer(HERE / "analyze_allocations.py", "analyze_allocations.py")
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


def verify_adr_manifest() -> dict[str, Any]:
    """Bind the retained ADR inventory to both the declared base and files."""

    path = HERE / "adr-manifest.json"
    value = load_json(path)
    require(isinstance(value, dict) and value.get("schema_version") == 1
            and value.get("algorithm") == "sha256"
            and value.get("revision") == BASE_REVISION,
            "ADR manifest header differs")
    files = value.get("files")
    require(isinstance(files, dict) and files, "ADR manifest files are missing")
    checked = 0
    for relative, expected in files.items():
        target = REPO / safe_relative(relative, "ADR manifest path")
        check_digest(expected, f"ADR manifest.{relative}")
        regular(target, f"ADR file {relative}")
        require(digest(target) == expected, f"ADR file hash differs for {relative}")
        try:
            process = subprocess.run(["git", "show", f"{BASE_REVISION}:{relative}"],
                                     cwd=REPO, stdout=subprocess.PIPE,
                                     stderr=subprocess.PIPE, check=False)
        except OSError as error:
            raise VerificationError(f"cannot read base ADR {relative}: {error}") from error
        require(process.returncode == 0 and hashlib.sha256(process.stdout).hexdigest() == expected,
                f"ADR base binding differs for {relative}")
        checked += 1
    return {"sha256": digest(path), "revision": BASE_REVISION, "entries": checked}


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
        "CARGO_TARGET_DIR": "/home/zhuhe/litchi-goal-0519-target",
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
        "binary_path": str(expected_binary),
        "binary_present": binary.exists(),
        "symbols_sha256": symbols_sha,
        "log_sha256": digest(log_path),
        "receipt_sha256": digest(receipt_path),
    }


PROBE = HERE / "allocator-probe"
ALLOCATOR_BUILD_COMMAND = [
    "cargo", "build", "--release", "--locked", "--manifest-path",
    str(PROBE / "Cargo.toml"), "--features", "allocator-metrics",
]
ALLOCATOR_BUILD_ENVIRONMENT = {
    "CARGO_TARGET_DIR": "/home/zhuhe/litchi-goal-0519-target/allocator",
    "CARGO_BUILD_JOBS": "2",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_DEBUG": "0",
    "TMPDIR": str(SCRATCH),
}


def probe_manifest() -> dict[str, str]:
    require(PROBE.is_dir(), "allocator probe directory is missing")
    result: dict[str, str] = {}
    for path in sorted(PROBE.rglob("*")):
        if path.is_symlink():
            raise VerificationError(f"allocator probe contains symlink: {path}")
        if path.is_file():
            result[path.relative_to(REPO).as_posix()] = digest(path)
    require(result, "allocator probe manifest is empty")
    return result


def verify_probe_binding() -> dict[str, Any]:
    binding_path = PROBE / "source-binding.json"
    binding = load_json(binding_path)
    require(isinstance(binding, dict), "allocator source binding is not an object")
    for field in ("source", "generated", "diff"):
        safe_relative(binding.get(field), f"allocator source binding {field}")
    source_path = REPO / safe_relative(binding["source"], "allocator source path")
    generated_path = REPO / safe_relative(binding["generated"], "allocator generated path")
    diff_path = REPO / safe_relative(binding["diff"], "allocator source diff path")
    regular(source_path, "allocator source example")
    regular(generated_path, "allocator generated source")
    regular(diff_path, "allocator source diff")
    source_sha = check_digest(binding.get("source_sha256"), "allocator source_sha256")
    generated_sha = check_digest(binding.get("generated_sha256"), "allocator generated_sha256")
    require(source_sha == digest(source_path),
            "allocator source binding source hash differs")
    require(generated_sha == digest(generated_path),
            "allocator source binding generated hash differs")
    require(binding.get("source_lines") == len(source_path.read_text(encoding="utf-8").splitlines()),
            "allocator source binding source line count differs")
    require(binding.get("generated_lines") == len(generated_path.read_text(encoding="utf-8").splitlines()),
            "allocator source binding generated line count differs")
    expected_diff = "".join(difflib.unified_diff(
        source_path.read_text(encoding="utf-8").splitlines(keepends=True),
        generated_path.read_text(encoding="utf-8").splitlines(keepends=True),
        fromfile=binding["source"], tofile=binding["generated"],
    ))
    require(diff_path.read_text(encoding="utf-8") == expected_diff,
            "allocator source diff is not the exact generated-source patch")
    canonical = binding.get("canonical_modules")
    require(isinstance(canonical, dict) and canonical, "allocator canonical module binding missing")
    for name in canonical.values():
        path = REPO / safe_relative(name, "allocator canonical module path")
        regular(path, "allocator canonical module")
    anchors = binding.get("anchors")
    require(isinstance(anchors, dict) and all(isinstance(v, str) and v for v in anchors.values()),
            "allocator source binding anchors missing")
    return {"sha256": digest(binding_path), "source": binding["source"],
            "source_sha256": source_sha,
            "generated": binding["generated"], "diff": binding["diff"]}


def verify_allocator_build(variant: str, expected_source: dict[str, str],
                           intervals: list[tuple[float, float, str]]) -> dict[str, Any]:
    directory = HERE / ("allocator-" + variant)
    receipt_path = directory / "build-receipt.json"
    log_path = directory / "build.log"
    source_path = directory / "source-manifest.json"
    probe_path = directory / "probe-manifest.json"
    source = manifest(source_path, f"allocator-{variant} source manifest")
    require(source == expected_source,
            f"allocator-{variant} source manifest differs from document build")
    probe = manifest(probe_path, f"allocator-{variant} probe manifest")
    require(probe == probe_manifest(), f"allocator-{variant} probe manifest differs")
    receipt = load_json(receipt_path)
    require(receipt.get("command") == ALLOCATOR_BUILD_COMMAND,
            f"allocator-{variant} build command differs")
    require(receipt.get("environment") == ALLOCATOR_BUILD_ENVIRONMENT,
            f"allocator-{variant} build environment differs")
    require(receipt.get("exit_code") == 0 and receipt.get("source_unchanged") is True
            and receipt.get("probe_unchanged") is True,
            f"allocator-{variant} build failed or changed sources")
    source_sha = digest(source_path)
    probe_sha = digest(probe_path)
    require(receipt.get("source_manifest_sha256") == source_sha,
            f"allocator-{variant} source manifest binding differs")
    require(receipt.get("probe_manifest_sha256") == probe_sha,
            f"allocator-{variant} probe manifest binding differs")
    require(receipt.get("log_sha256") == digest(log_path),
            f"allocator-{variant} build log hash differs")
    expected_binary = SCRATCH / ("allocation-" + variant)
    require(receipt.get("binary") == str(expected_binary),
            f"allocator-{variant} binary path differs")
    binary_sha = check_digest(receipt.get("binary_sha256"),
                               f"allocator-{variant} binary_sha256")
    if expected_binary.exists():
        regular(expected_binary, f"allocator-{variant} binary")
        require(digest(expected_binary) == binary_sha,
                f"allocator-{variant} binary hash differs")
    intervals.append((*receipt_interval(receipt, f"allocator-{variant} build"),
                      f"allocator-{variant}/build"))
    return {
        "source_manifest_sha256": source_sha,
        "probe_manifest_sha256": probe_sha,
        "source_files": len(source),
        "probe_files": len(probe),
        "binary_sha256": binary_sha,
        "binary_path": str(expected_binary),
        "binary_present": expected_binary.exists(),
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


ALLOCATION_SCOPE = "publish_document_commit_to_stream_method_only_before_returned_snapshot_drop"
ALLOCATION_SAMPLE_SCOPE = "operation_global_system_allocator"
ALLOCATION_FIELDS = (
    "status", "scope", "allocation_calls", "deallocation_calls",
    "reallocation_calls", "failed_allocation_calls", "allocated_bytes",
    "deallocated_bytes", "live_bytes_before", "live_bytes_after",
    "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes",
)


def verify_allocation_sample(path: Path, name: str) -> dict[str, Any]:
    """Validate the probe's one JSON sample independently of its CSV oracle."""

    regular(path, f"allocation sample {path}")
    lines = [line for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    require(len(lines) == 1, f"{path.name}: allocation stdout must contain one JSON sample")
    try:
        value = json.loads(lines[0], object_pairs_hook=_no_duplicate_pairs,
                           parse_constant=_reject_constant)
    except (OSError, UnicodeError, json.JSONDecodeError, VerificationError) as error:
        raise VerificationError(f"{path.name}: malformed allocation sample: {error}") from error
    require(isinstance(value, dict), f"{path.name}: allocation sample is not an object")
    require(set(value) == {"tag", "scope", "case", "repeat", "ordinal", "warmup",
                           "allocationSample"},
            f"{path.name}: allocation sample fields differ")
    require(value["tag"] == "allocationSample" and value["scope"] == ALLOCATION_SCOPE
            and value["case"] == name,
            f"{path.name}: allocation sample identity differs")
    require(value["repeat"] == 0 and value["ordinal"] == 0 and value["warmup"] is False,
            f"{path.name}: allocation sample ordinal differs")
    sample = value["allocationSample"]
    require(isinstance(sample, dict) and set(sample) == set(ALLOCATION_FIELDS),
            f"{path.name}: allocation counters differ")
    require(sample["status"] == "measured" and sample["scope"] == ALLOCATION_SAMPLE_SCOPE,
            f"{path.name}: allocation status/scope differs")
    for field in ALLOCATION_FIELDS[2:]:
        item = sample[field]
        require(isinstance(item, int) and not isinstance(item, bool) and item >= 0,
                f"{path.name}: allocation counter {field} is invalid")
    require(sample["failed_allocation_calls"] == 0,
            f"{path.name}: failed allocation calls are nonzero")
    require(sample["live_bytes_before"] + sample["allocated_bytes"]
            - sample["deallocated_bytes"] == sample["live_bytes_after"],
            f"{path.name}: live-byte balance does not reconcile")
    require(sample["peak_live_bytes_before"] <= sample["peak_live_bytes_after"],
            f"{path.name}: absolute peak moved backwards")
    require(sample["region_peak_live_bytes"] >= max(
        sample["live_bytes_before"], sample["live_bytes_after"]),
            f"{path.name}: region peak is below live bytes")
    require(sample["peak_live_bytes_before"] >= sample["live_bytes_before"]
            and sample["peak_live_bytes_after"] >= sample["live_bytes_after"],
            f"{path.name}: peak live bytes are below live bytes")
    require(sample["region_peak_live_bytes"] <= sample["peak_live_bytes_after"],
            f"{path.name}: region peak exceeds absolute peak")
    return sample


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


def verify_profile_function(path: Path, selected_name: str) -> dict[str, Any]:
    """Validate a debug profile's selected function without assuming its caller."""

    try:
        summary, edges = profile_parser.parse_raw(path)
    except Exception as error:
        raise VerificationError(f"cannot parse {path.name}: {error}") from error
    selected = [edge for edge in edges
                if profile_parser.base_name(edge.child) == selected_name
                and edge.calls > 0]
    require(len(selected) == 1 and selected[0].calls == 1,
            f"{path.name}: selected function does not have one positive call")
    require(selected[0].cost > 0 and summary == selected[0].cost,
            f"{path.name}: selected function is not scoped to the profile summary")
    return {"summary_ir": summary, "selected_function": selected_name,
            "caller": profile_parser.base_name(selected[0].parent),
            "calls": selected[0].calls, "inclusive_ir": selected[0].cost}


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
    if kind == "allocation":
        # allocations.py records the fixed one-sample probe without repeating
        # the CLI settings in every receipt; the command itself is checked
        # below and the stdout sample is checked independently.
        require("samples" not in receipt and "warmups" not in receipt
                and "repeats" not in receipt,
                f"{lane}/{name}: allocation receipt unexpectedly has sample settings")
    else:
        require(receipt.get("samples") == samples and receipt.get("warmups") == warmups
                and receipt.get("repeats") == repeats,
                f"{lane}/{name}: sample settings differ")
    require(receipt.get("source_manifest_sha256") == build["source_manifest_sha256"],
            f"{lane}/{name}: source binding differs from build")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{lane}/{name}: binary binding differs from build")
    if kind == "allocation":
        require(receipt.get("candidate_plan_sha256") == digest(HERE / "candidate-plan.json"),
                f"{lane}/{name}: candidate plan binding differs")
    else:
        require(receipt.get("plan_sha256") == digest(HERE / "plan.json"),
                f"{lane}/{name}: plan binding differs")
    if variant == "candidate":
        require(receipt.get("candidate_plan_sha256") == digest(HERE / "candidate-plan.json"),
                f"{lane}/{name}: candidate plan binding differs")
    if kind == "allocation":
        require("scope" not in receipt, f"{lane}/{name}: allocation receipt has an unexpected scope")
    else:
        expected_scope = (plan["profile"]["scope"] if kind == "profile" else
                          plan["hardware"]["scope"] if kind == "hardware" else
                          plan["native"]["scope"])
        require(receipt.get("scope") == expected_scope,
                f"{lane}/{name}: scope differs from plan")
    binary = build["binary_path"]
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
    elif kind == "allocation":
        profile = {"allocation": verify_allocation_sample(directory / f"{name}.stdout", name)}
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
    if lane not in {"r1", "r2", "profile-r1", "profile-r2", "hardware",
                    "after-r1", "after-r2", "profile-after-r1", "profile-after-r2",
                    "hardware-after", "alloc-baseline-r1", "alloc-baseline-r2",
                    "alloc-candidate-r1", "alloc-candidate-r2"}:
        return
    names = (list(plan["profile"]["cases"])
             if lane.startswith("profile") or lane.startswith("hardware") else sorted(records))
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
                          candidate_plan: dict[str, Any],
                          require_complete: bool = False) -> dict[str, Any]:
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
    replay_path = HERE / "candidate" / "patch-replay.json"
    if explicit is None and replay_path.exists():
        replay = load_json(replay_path)
        files = replay.get("files") if isinstance(replay, dict) else None
        require(isinstance(files, dict) and files,
                "candidate patch replay has no file inventory")
        explicit = {safe_relative(item, "candidate patch replay path").as_posix()
                    for item in files}
    if explicit is None:
        if changed != DEFAULT_CANDIDATE_PATHS and not require_complete:
            # A candidate can be observed before the root has retained its
            # exact patch replay.  Preserve the actual diff for the pending
            # result; the final complete mode still requires that replay to
            # declare the authoritative path inventory.
            return {"changed_paths": sorted(changed), "allowed_paths": None,
                    "policy": "awaiting candidate/patch-replay.json path inventory"}
        require(changed == DEFAULT_CANDIDATE_PATHS,
                f"candidate source diff differs from the frozen candidate path policy: {sorted(changed)}")
        expected = sorted(DEFAULT_CANDIDATE_PATHS)
        policy = "frozen candidate-plan source path policy"
    else:
        require(changed == explicit,
                f"candidate source diff differs from declared paths: changed={sorted(changed)} declared={sorted(explicit)}")
        expected = sorted(explicit)
        policy = ("candidate-plan.json intended_source_paths"
                  if any(key in candidate_plan for key in
                         ("intended_source_paths", "allowed_source_paths", "changed_paths"))
                  else "candidate/patch-replay.json files")
    return {"changed_paths": sorted(changed), "allowed_paths": expected, "policy": policy}


def _base_has_path(relative: str) -> bool:
    process = subprocess.run(["git", "cat-file", "-e", f"{BASE_REVISION}:{relative}"],
                             cwd=REPO, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                             check=False)
    return process.returncode == 0


def exact_source_diff(paths: list[str]) -> bytes:
    chunks: list[bytes] = []
    for relative in paths:
        if _base_has_path(relative):
            command = ["git", "diff", "--no-ext-diff", "--no-color", BASE_REVISION,
                       "--", relative]
        else:
            command = ["git", "diff", "--no-index", "--no-ext-diff", "--no-color",
                       "/dev/null", relative]
        process = subprocess.run(command, cwd=REPO, stdout=subprocess.PIPE,
                                 stderr=subprocess.PIPE, check=False)
        require(process.returncode in {0, 1},
                f"git diff failed while binding {relative}")
        chunks.append(process.stdout)
    return b"".join(chunks)


def patch_paths(patch: bytes) -> list[str]:
    """Recover the retained diff section order without trusting its contents."""

    names = []
    for line in patch.splitlines():
        match = re.fullmatch(rb"diff --git a/(.+) b/(.+)", line)
        if match:
            left, right = (item.decode("utf-8", "strict") for item in match.groups())
            require(left == right, "candidate source patch renames a source path")
            names.append(left)
    require(len(names) == len(set(names)),
            "candidate source patch repeats a source path")
    return names


def verify_patch_replay(paths: list[str], baseline: dict[str, str],
                        candidate: dict[str, str]) -> dict[str, Any] | None:
    replay_path = HERE / "candidate" / "patch-replay.json"
    patch_path = HERE / "candidate" / "source.patch"
    if not replay_path.exists():
        require(not patch_path.exists(), "candidate source patch has no patch replay")
        return None
    replay = load_json(replay_path)
    require(isinstance(replay, dict)
            and replay.get("base_revision") == BASE_REVISION
            and replay.get("exact_replay") is True
            and replay.get("temporary_removed") is True,
            "candidate patch replay header differs")
    files = replay.get("files")
    require(isinstance(files, dict) and set(files) == set(paths),
            "candidate patch replay file inventory differs")
    regular(patch_path, "candidate source patch")
    expected_patch_sha = check_digest(replay.get("source_patch_sha256"),
                                      "candidate patch replay source_patch_sha256")
    require(digest(patch_path) == expected_patch_sha,
            "candidate patch replay source patch hash differs")
    for relative in sorted(paths):
        item = files[relative]
        require(isinstance(item, dict) and set(item) == {"before_sha256", "after_sha256"},
                f"candidate patch replay {relative} entry differs")
        after_sha = check_digest(item.get("after_sha256"),
                                 f"candidate patch replay {relative}.after_sha256")
        require(candidate.get(relative) == after_sha,
                f"candidate patch replay after hash differs for {relative}")
        before_sha = item.get("before_sha256")
        if before_sha is None:
            require(not _base_has_path(relative),
                    f"candidate patch replay marks existing base file as new: {relative}")
        else:
            before_sha = check_digest(before_sha,
                                      f"candidate patch replay {relative}.before_sha256")
            require(_base_has_path(relative),
                    f"candidate patch replay base file is missing: {relative}")
            base_bytes = subprocess.check_output(["git", "show", f"{BASE_REVISION}:{relative}"],
                                                 cwd=REPO)
            require(hashlib.sha256(base_bytes).hexdigest() == before_sha,
                    f"candidate patch replay before hash differs for {relative}")
            require(baseline.get(relative) == before_sha,
                    f"candidate patch replay is not bound to retained baseline for {relative}")
        target = REPO / safe_relative(relative, "candidate patch replay path")
        regular(target, f"candidate patch replay current file {relative}")
        require(digest(target) == after_sha,
                f"candidate patch replay current hash differs for {relative}")
    require(patch_path.read_bytes() == exact_source_diff(list(files)),
            "candidate source patch differs from exact base-to-current diff")
    return {"sha256": digest(replay_path), "source_patch_sha256": expected_patch_sha,
            "files": sorted(paths), "exact_replay": True}


def verify_source_patch(paths: list[str] | None = None,
                        required: bool = False) -> dict[str, Any] | None:
    """Bind the retained patch to the current tree relative to the base."""

    path = HERE / "candidate" / "source.patch"
    if not path.exists():
        require(not required, "candidate source patch is missing")
        return None
    regular(path, "candidate source patch")
    selected = sorted(paths or DEFAULT_CANDIDATE_PATHS)
    require(selected and all(isinstance(item, str) for item in selected),
            "candidate source patch path list is empty")
    retained = path.read_bytes()
    order = patch_paths(retained)
    require(set(order) == set(selected),
            "candidate source patch path inventory differs from declared paths")
    require(retained == exact_source_diff(order),
            "candidate source patch differs from the current base-to-candidate diff")
    return {
        "sha256": digest(path),
        "bytes": len(retained),
        "matches_base_diff": True,
        "paths": selected,
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


def verify_hardware_summary(plan: dict[str, Any], candidate_complete: bool = False) -> dict[str, Any]:
    """Replay the retained six-case hardware diagnostic when both arms exist."""

    path = HERE / "hardware-summary.json"
    have_after = (HERE / "hardware-after").is_dir()
    if not have_after:
        require(not path.exists(), "hardware summary exists without hardware-after lane")
        return {"replayed": False, "cases": len(PROFILE_NAMES), "summary_retained": False}
    require(hardware_analyzer is not None,
            "hardware-after evidence requires analyze_hardware.py")
    try:
        replay = hardware_analyzer.analyze()
    except Exception as error:
        raise VerificationError(f"hardware analyzer replay failed: {error}") from error
    if not path.exists():
        require(not candidate_complete, "complete candidate is missing hardware summary")
        return {"replayed": True, "cases": len(replay["baseline"]),
                "summary_retained": False}
    summary = load_json(path)
    require(summary == replay, "hardware summary differs from analyzer replay")
    require(summary.get("scope") == replay.get("scope"), "hardware summary scope differs")
    for arm in ("baseline", "candidate"):
        cases = summary.get(arm)
        require(isinstance(cases, list) and len(cases) == len(PROFILE_NAMES),
                f"hardware summary {arm} case inventory differs")
        require({item.get("case") for item in cases} == PROFILE_NAMES,
                f"hardware summary {arm} case names differ")
    return {"sha256": digest(path), "cases": len(summary["baseline"]), "replayed": True}


def recompute_native_report(lanes: list[str]) -> dict[str, Any]:
    campaigns = {lane: native_analyzer.load_campaign(lane) for lane in lanes}
    bindings = {tuple(campaign["binding"].values()) for campaign in campaigns.values()}
    require(len(bindings) == 1, "native report campaigns do not share one binding")
    first = campaigns[lanes[0]]
    seed = native_analyzer.DEFAULT_SEED
    bootstraps = native_analyzer.DEFAULT_BOOTSTRAPS
    return {
        "schema": "managed_paragraph_native_analysis_0519_v1",
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
        "schema": "managed_paragraph_native_candidate_comparison_0519_v1",
        "pairs": pairs,
        "seed": seed,
        "bootstrap_iterations": bootstraps,
        "thresholds_percent": {"p50": 5, "mean": 5, "p95": 10, "p99": 15, "rss": 5},
        "scope": "same-API native baseline/candidate comparison; 24 cases per pair; warm=false rows only; publish_ns includes returned Snapshot drop; RSS is whole-child",
        "counter_policy": {
            "allowed_difference_fields": sorted(candidate_comparator.ALLOWED_COUNTER_DIFFERENCES),
            "guard_fields": list(candidate_comparator.GUARD_COUNTERS),
            "allowed_difference_note": "retained live/cache gauges are descriptive comparison data; Work is expected unchanged; release, output, input, and source-read fields remain guards",
        },
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
            "snapshot_owner": sorted(profile_comparator.SNAPSHOT_OWNER_NAMES),
            "fresh_docx_scan": profile_comparator.FRESH_DOCX_SCAN_PREFIX,
            "fresh_snapshot_build": profile_comparator.FRESH_SNAPSHOT_BUILD_PREFIX,
            "topology": profile_comparator.TOPOLOGY,
            "xml_validator": profile_comparator.XML_VALIDATOR,
            "xml_validator_policy": profile_comparator.XML_VALIDATOR_POLICY,
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
    paths = [HERE / "baseline-native-analysis.json", HERE / "profile-analysis.json"]
    paths.extend(sorted(HERE.glob("profile-*/profile-analysis.json")))
    paths.extend(sorted(HERE.glob("*native-analysis*.json")))
    required_reports = {
        HERE / "baseline-native-analysis.json",
        HERE / "candidate-native-analysis.json",
        HERE / "profile-analysis.json",
        HERE / "profile-after-analysis.json",
        HERE / "candidate-comparison.json",
        HERE / "profile-comparison.json",
        HERE / "publication-allocation-comparison.json",
        HERE / "counter-comparison.json",
        HERE / "hardware-summary.json",
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
            require(report.get("schema") == "managed_paragraph_native_analysis_0519_v1",
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
        require(counter_analyzer is not None,
                "counter comparison requires analyze_counters.py")
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
    evidence_root = Path("docs/performance/results/change-0519")
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


def verify_focused_receipts(intervals: list[tuple[float, float, str]]) -> dict[str, Any]:
    """Check retained focused correctness runs without treating them as benchmarks."""

    checked: dict[str, Any] = {}
    focused_directories = []
    if (HERE / "focused-check").is_dir():
        focused_directories.append(HERE / "focused-check")
    focused_directories.extend(HERE.glob("focused-*-check"))
    for directory in sorted(set(focused_directories)):
        if not directory.is_dir():
            continue
        source_path = directory / "source-manifest.json"
        if not source_path.exists():
            raise VerificationError(f"{directory.name}: source manifest is missing")
        retained_manifest = manifest(source_path, f"{directory.name} source manifest")
        retained_sha = digest(source_path)
        # These manifests deliberately snapshot the source at each focused
        # run; they need not all equal the final candidate epoch.
        receipts = sorted(directory.glob("receipt*.json"))
        require(receipts, f"{directory.name}: focused receipt is missing")
        for receipt_path in receipts:
            receipt = load_json(receipt_path)
            require(isinstance(receipt, dict) and isinstance(receipt.get("command"), list)
                    and receipt["command"], f"{receipt_path.name}: focused command is missing")
            require(receipt.get("exit_code") == 0 and receipt.get("source_unchanged") is True,
                    f"{receipt_path.name}: focused check failed or changed source")
            require(receipt.get("source_manifest_sha256") == retained_sha,
                    f"{receipt_path.name}: focused source binding differs")
            suffix = receipt_path.stem.removeprefix("receipt")
            log_path = directory / (("check" + suffix) if suffix else "check")
            log_path = log_path.with_suffix(".log")
            require(receipt.get("log_sha256") == digest(log_path),
                    f"{receipt_path.name}: focused log hash differs")
            intervals.append((*receipt_interval(receipt, f"{directory.name}/{receipt_path.name}"),
                              f"{directory.name}/{receipt_path.name}"))
            checked[f"{directory.name}/{receipt_path.name}"] = {
                "receipt_sha256": digest(receipt_path),
                "log_sha256": digest(log_path),
                "source_manifest_sha256": retained_sha,
                "source_files": len(retained_manifest),
            }
    return checked


def verify_report_replay(required: bool = False) -> dict[str, Any] | None:
    """Validate the retained exact-byte replay ledger without rerunning tools."""

    path = HERE / "report-replay.json"
    if not path.exists():
        require(not required, "report replay ledger is missing")
        return None
    report = load_json(path)
    require(isinstance(report, dict) and report.get("temporary_replay_removed") is True,
            "report replay header differs")
    records = report.get("reports")
    require(isinstance(records, list), "report replay records are missing")
    expected = {
        "baseline-native-analysis": ("analyze_native.py", [], "--json"),
        "candidate-native-analysis": ("analyze_native.py",
                                      ["--lanes", "after-r1", "after-r2"], "--json"),
        "candidate-comparison": ("compare_candidate.py", [], "--json"),
        "profile-comparison": ("compare_profile_lanes.py", [], "--json"),
        "publication-allocation-comparison": ("analyze_allocations.py", [], "--output"),
    }
    require(len(records) == len(expected), "report replay count differs")
    checked: dict[str, Any] = {}
    for item in records:
        require(isinstance(item, dict) and set(item) == {
            "command", "exit_code", "helper_sha256", "output_sha256",
        }, "report replay record fields differ")
        command = item["command"]
        require(isinstance(command, list) and len(command) >= 4,
                "report replay command is malformed")
        output = item["output_sha256"]
        require(isinstance(output, dict) and all(isinstance(key, str) for key in output)
                and all(isinstance(value, str) for value in output.values()),
                "report replay output inventory differs")
        json_names = [key for key in output if key.endswith(".json")]
        require(len(json_names) == 1, "report replay JSON output is missing or duplicated")
        name = Path(json_names[0]).name.removesuffix(".json")
        require(name in expected and name not in checked,
                "report replay report name is unexpected or duplicated")
        script, extra, option = expected[name]
        expected_command = ["python3", "-B", str(HERE / script), *extra, option,
                            str(SCRATCH / "report-replay" / (name + ".json")),
                            "--markdown", str(SCRATCH / "report-replay" / (name + ".md"))]
        require(command == expected_command, f"report replay command differs for {name}")
        require(item["exit_code"] == 0, f"report replay failed for {name}")
        check_digest(item["helper_sha256"], f"report replay {name} helper_sha256")
        require(item["helper_sha256"] == digest(HERE / script),
                f"report replay helper hash differs for {name}")
        require(set(output) == {name + ".json", name + ".md"},
                f"report replay output names differ for {name}")
        for suffix in (".json", ".md"):
            retained = HERE / (name + suffix)
            regular(retained, f"report replay retained {name}{suffix}")
            expected_sha = check_digest(output[name + suffix],
                                        f"report replay {name}{suffix} hash")
            require(expected_sha == digest(retained),
                    f"report replay output hash differs for {name}{suffix}")
        checked[name] = {"helper_sha256": item["helper_sha256"],
                         "outputs": dict(output)}
    return {"sha256": digest(path), "reports": checked,
            "temporary_replay_removed": True}


def verify_allocation_report(builds: dict[str, Any], required: bool = False) -> dict[str, Any] | None:
    path = HERE / "publication-allocation-comparison.json"
    if not path.exists():
        require(not required, "publication allocation comparison is missing")
        return None
    require(allocation_analyzer is not None,
            "publication allocation comparison requires analyze_allocations.py")
    report = load_json(path)
    require(isinstance(report, dict) and report.get("schema_version") == 1
            and report.get("status") == "complete"
            and report.get("claim_scope") == allocation_analyzer.CLAIM_SCOPE,
            "publication allocation comparison header differs")
    build_scope = report.get("allocator_build_scope")
    require(isinstance(build_scope, dict)
            and build_scope.get("kind") == "standalone_allocator_probe_release_binary"
            and isinstance(build_scope.get("comparison_validity"), str)
            and build_scope["comparison_validity"]
            and isinstance(build_scope.get("absolute_count_limit"), str)
            and build_scope["absolute_count_limit"],
            "publication allocation build scope differs")
    standalone = build_scope.get("standalone_probe")
    workspace = build_scope.get("workspace_native")
    require(isinstance(standalone, dict) and isinstance(workspace, dict),
            "publication allocation build scope entries are missing")
    for item, context in ((standalone, "standalone allocator build scope"),
                          (workspace, "workspace native build scope")):
        require(isinstance(item.get("manifest"), str)
                and isinstance(item.get("lockfile"), str),
                f"{context} paths are missing")
        # The analyzer emits probe paths relative to the evidence directory,
        # while workspace paths are absolute so the build context is explicit.
        root = HERE if item is standalone else REPO
        manifest_path = root / item["manifest"] if not Path(item["manifest"]).is_absolute() \
            else Path(item["manifest"])
        lockfile_path = root / item["lockfile"] if not Path(item["lockfile"]).is_absolute() \
            else Path(item["lockfile"])
        require(item.get("manifest_sha256") == digest(manifest_path)
                and item.get("lockfile_sha256") == digest(lockfile_path),
                f"{context} hashes differ")
    require(standalone.get("manifest") == "allocator-probe/Cargo.toml"
            and standalone.get("lockfile") == "allocator-probe/Cargo.lock"
            and standalone.get("dependency_resolution") == "independent standalone lockfile"
            and standalone.get("release_profile") == {
                "profile_section": "no [profile.release]; Cargo defaults",
                "lto": "off (Cargo default)",
                "panic": "unwind (Cargo default)",
            }, "standalone allocator build scope details differ")
    require(workspace.get("manifest") == str(REPO / "Cargo.toml")
            and workspace.get("lockfile") == str(REPO / "Cargo.lock")
            and workspace.get("release_profile") == {"lto": "true", "panic": "abort"},
            "workspace native build scope details differ")
    require(report.get("candidate_plan_sha256") == digest(HERE / "candidate-plan.json"),
            "publication allocation comparison candidate plan binding differs")
    require(report.get("allocation_scope") == allocation_analyzer.ALLOCATION_SCOPE
            and report.get("timing_scope") == allocation_analyzer.TIMING_SCOPE,
            "publication allocation comparison scope differs")
    observer = report.get("observer")
    require(isinstance(observer, dict), "publication allocation observer report is missing")
    observer_source = REPO / "tools/perf-baseline/src/allocation_metrics.rs"
    require(observer.get("source_sha256") == digest(observer_source),
            "publication allocation observer source hash differs")
    require(observer.get("revision") == allocation_analyzer.OBSERVER_REVISION
            and observer.get("revision_sha256") == hashlib.sha256(
                allocation_analyzer.OBSERVER_REVISION.encode()).hexdigest(),
            "publication allocation observer revision differs")
    probe = report.get("probe")
    require(isinstance(probe, dict)
            and probe.get("source_sha256") == probe_binding_source_sha()
            and probe.get("generated_sha256") == digest(PROBE / "src/main.rs"),
            "publication allocation probe binding differs")
    report_builds = report.get("builds")
    require(isinstance(report_builds, dict) and set(report_builds) == {"baseline", "candidate"},
            "publication allocation build inventory differs")
    for stage in ("baseline", "candidate"):
        item = report_builds[stage]
        require(isinstance(item, dict) and item.get("valid") is True,
                f"publication allocation {stage} build is not valid")
        expected = builds[stage]
        require(item.get("receipt_sha256") == expected["receipt_sha256"]
                and item.get("source_manifest_sha256") == expected["source_manifest_sha256"]
                and item.get("probe_manifest_sha256") == expected["probe_manifest_sha256"]
                and item.get("binary_sha256") == expected["binary_sha256"],
                f"publication allocation {stage} build binding differs")
    captures = report.get("captures")
    require(isinstance(captures, dict)
            and set(captures) == {"baseline-r1", "baseline-r2", "candidate-r1", "candidate-r2"},
            "publication allocation capture inventory differs")
    for key, item in captures.items():
        require(isinstance(item, dict) and item.get("valid") is True
                and item.get("complete") is True
                and item.get("missing_cases") == []
                and len(item.get("rows", [])) == len(NATIVE_NAMES),
                f"publication allocation {key} is incomplete")
    comparisons = report.get("comparisons")
    require(isinstance(comparisons, list) and len(comparisons) == len(NATIVE_NAMES),
            "publication allocation comparison case count differs")
    require(report.get("warnings") == [], "publication allocation report has warnings")
    return {"sha256": digest(path), "comparisons": len(comparisons), "status": "complete"}


def probe_binding_source_sha() -> str:
    binding = load_json(PROBE / "source-binding.json")
    return check_digest(binding.get("source_sha256"), "allocator source binding source_sha256")


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


def count_adverse_flags(report: dict[str, Any]) -> int:
    comparisons = report.get("comparisons")
    require(isinstance(comparisons, dict), "candidate comparison has no comparisons")
    total = 0
    for records in comparisons.values():
        require(isinstance(records, list), "candidate comparison records are malformed")
        for record in records:
            require(isinstance(record, dict), "candidate comparison record is malformed")
            flags = record.get("adverse_flags", [])
            require(isinstance(flags, list), "candidate comparison flags are malformed")
            total += len(flags)
    return total


def verify_tail_guard(
    intervals: list[tuple[float, float, str]],
    builds: dict[str, dict[str, Any]],
    candidate_source_sha: str | None,
    *,
    required: bool = False,
) -> dict[str, Any] | None:
    """Validate the bounded p99 follow-up and replay its retained report.

    Tail-guard captures use retained binaries with the current candidate source
    tree left checked out.  Consequently a baseline receipt is bound to the
    baseline binary but still records the unchanged candidate source snapshot;
    this is deliberate and keeps the binary/source custody distinction visible.
    """

    if not TAIL_GUARD_PLAN.exists():
        require(not required, "tail-guard plan is missing")
        return None
    plan = load_json(TAIL_GUARD_PLAN)
    if plan.get("schema") != TAIL_GUARD_PLAN_SCHEMA:
        require(not required, "tail-guard plan is not the corrected v2 protocol")
        return {"status": "pending", "reason": "tail-guard plan is not corrected v2"}
    require(plan.get("base_revision") == BASE_REVISION,
            "tail-guard base revision differs")
    protocol = plan.get("protocol")
    require(isinstance(protocol, dict), "tail-guard protocol is missing")
    require(protocol.get("cases") == TAIL_GUARD_CASES,
            "tail-guard case inventory differs")
    expected_order = [
        {"slot": slot, "variant": variant, "case_id": case_id,
         "case": TAIL_GUARD_CASES[case_id]}
        for slot, variant, case_id in TAIL_GUARD_ORDER
    ]
    require(protocol.get("order") == expected_order,
            "tail-guard process order differs")
    require(protocol.get("samples") == 200
            and protocol.get("warmups") == 10
            and protocol.get("repeats") == 1
            and protocol.get("cpu") == 2
            and protocol.get("fresh_process_per_case") is True,
            "tail-guard sample protocol differs")
    require(protocol.get("phase_metrics") == list(TAIL_GUARD_PHASES),
            "tail-guard phase inventory differs")

    plan_builds = plan.get("builds")
    require(isinstance(plan_builds, dict) and set(plan_builds) == {"baseline", "candidate"},
            "tail-guard build inventory differs")
    for variant in ("baseline", "candidate"):
        retained = builds.get(variant)
        if retained is None:
            require(not required, f"tail-guard {variant} build binding is unavailable")
            continue
        item = plan_builds[variant]
        require(isinstance(item, dict), f"tail-guard {variant} build entry is malformed")
        require(item.get("variant") == variant
                and item.get("binary_path") == retained["binary_path"]
                and item.get("binary_sha256") == retained["binary_sha256"]
                and item.get("source_manifest_sha256") == retained["source_manifest_sha256"]
                and item.get("build_receipt_sha256") == retained["receipt_sha256"],
                f"tail-guard {variant} build binding differs")
        require(item.get("build_receipt_path") == str(HERE / variant / "build-receipt.json")
                and item.get("build_manifest_path") == str(HERE / variant / "source-manifest.json"),
                f"tail-guard {variant} build paths differ")
    if candidate_source_sha is not None:
        require(plan.get("candidate_source_manifest_current_sha256") == candidate_source_sha,
                "tail-guard candidate source binding differs")

    original = plan.get("original_comparison")
    require(isinstance(original, dict)
            and original.get("path") == str(HERE / "candidate-comparison.json"),
            "tail-guard original comparison path differs")
    original_path = HERE / "candidate-comparison.json"
    original_sha = digest(original_path)
    require(original.get("sha256") == original_sha,
            "tail-guard original comparison hash differs")
    require(original.get("adverse_flag_count") == count_adverse_flags(load_json(original_path)),
            "tail-guard original comparison flag count differs")

    if not TAIL_GUARD_REPORT.exists():
        require(not required, "tail-guard report is missing")
        return {"status": "pending", "plan_sha256": digest(TAIL_GUARD_PLAN)}
    require(TAIL_GUARD_DIR.is_dir(), "tail-guard capture directory is missing")
    regular(TAIL_GUARD_REPORT, "tail-guard report")
    regular(TAIL_GUARD_MARKDOWN, "tail-guard markdown")

    entries = [(slot, variant, case_id, TAIL_GUARD_CASES[case_id])
               for slot, variant, case_id in TAIL_GUARD_ORDER]
    expected_names = {
        f"{slot}-{case}.{suffix}"
        for slot, _variant, _case_id, case in entries
        for suffix in ("csv", "json", "stdout", "stderr")
    }
    actual_names: set[str] = set()
    for path in TAIL_GUARD_DIR.rglob("*"):
        require(not path.is_symlink(), f"tail-guard tree contains symlink: {path}")
        require(path.is_file(), f"tail-guard tree contains non-file: {path}")
        actual_names.add(path.relative_to(TAIL_GUARD_DIR).as_posix())
    require(actual_names == expected_names,
            "tail-guard capture inventory differs")

    tail_intervals: list[tuple[float, float, str]] = []
    for slot, variant, case_id, case in entries:
        name = f"{slot}-{case}"
        receipt_path = TAIL_GUARD_DIR / f"{name}.json"
        receipt = load_json(receipt_path)
        require(receipt.get("schema") == TAIL_GUARD_CAPTURE_SCHEMA,
                f"tail-guard {name}: receipt schema differs")
        require(receipt.get("slot") == slot and receipt.get("variant") == variant
                and receipt.get("case_id") == case_id and receipt.get("case") == case,
                f"tail-guard {name}: receipt identity differs")
        require(receipt.get("exit_code") == 0
                and receipt.get("source_unchanged") is True
                and receipt.get("cleanup_verified") is True
                and receipt.get("fresh_process_per_case") is True,
                f"tail-guard {name}: capture failed, mutated source, or leaked scratch")
        require(receipt.get("cpu") == 2 and receipt.get("samples") == 200
                and receipt.get("warmups") == 10 and receipt.get("repeats") == 1,
                f"tail-guard {name}: capture settings differ")
        require(receipt.get("scope") == plan["scope"],
                f"tail-guard {name}: scope differs")
        require(receipt.get("plan_sha256") == digest(TAIL_GUARD_PLAN),
                f"tail-guard {name}: plan binding differs")

        build = builds.get(variant)
        require(build is not None, f"tail-guard {name}: {variant} build binding is missing")
        require(receipt.get("build_receipt_path") == str(HERE / variant / "build-receipt.json")
                and receipt.get("build_manifest_path") == str(HERE / variant / "source-manifest.json")
                and receipt.get("build_receipt_sha256") == build["receipt_sha256"]
                and receipt.get("build_source_manifest_sha256") == build["source_manifest_sha256"]
                and receipt.get("binary_path") == build["binary_path"]
                and receipt.get("binary_sha256") == build["binary_sha256"]
                and receipt.get("binary_sha256_before") == build["binary_sha256"]
                and receipt.get("binary_sha256_after") == build["binary_sha256"],
                f"tail-guard {name}: build or binary binding differs")
        require(candidate_source_sha is not None
                and receipt.get("source_manifest_before_sha256") == candidate_source_sha
                and receipt.get("source_manifest_after_sha256") == candidate_source_sha
                and receipt.get("candidate_source_manifest_sha256") == candidate_source_sha
                and receipt.get("candidate_source_before_matches_build") is True
                and receipt.get("candidate_source_after_matches_build") is True,
                f"tail-guard {name}: source snapshot binding differs")
        expected_baseline_binding = (
            "not_required_for_retained_baseline_binary" if variant == "baseline"
            else "candidate_manifest_required_before_and_after"
        )
        require(receipt.get("working_tree_baseline_binding") == expected_baseline_binding,
                f"tail-guard {name}: working-tree binding policy differs")

        paragraphs, replacements, source, mode = capture.prior.parse_name(case)
        csv_path = TAIL_GUARD_DIR / f"{name}.csv"
        stdout_path = TAIL_GUARD_DIR / f"{name}.stdout"
        stderr_path = TAIL_GUARD_DIR / f"{name}.stderr"
        expected_command = [
            "/usr/bin/time", "-v", "taskset", "-c", "2", build["binary_path"],
            "--paragraphs", str(paragraphs), "--replacements", str(replacements),
            "--source", source, "--mode", mode,
            "--samples", "200", "--warmups", "10", "--repeats", "1",
            "--artifact-dir", str(SCRATCH / "tail-guard-corpora" / name),
            "--output", str(csv_path),
        ]
        require(receipt.get("command") == expected_command,
                f"tail-guard {name}: command differs")
        artifacts = receipt.get("artifacts")
        require(isinstance(artifacts, dict)
                and set(artifacts) == {f"{name}.csv", f"{name}.stdout", f"{name}.stderr"},
                f"tail-guard {name}: artifact inventory differs")
        for path in (csv_path, stdout_path, stderr_path):
            expected_sha = check_digest(artifacts.get(path.name),
                                        f"tail-guard {name} {path.name} hash")
            regular(path, f"tail-guard {name} {path.name}",
                    nonempty=path is not stdout_path)
            require(digest(path) == expected_sha,
                    f"tail-guard {name}: artifact hash differs for {path.name}")
        rows = read_rows(csv_path)
        try:
            capture.validate(case, rows, 200, 10, 1)
        except (AssertionError, KeyError, TypeError, ValueError) as error:
            raise VerificationError(f"tail-guard {name}: CSV oracle validation failed: {error}") from error
        require(len(rows) == 210, f"tail-guard {name}: row count differs")
        require(all(row.get("budget_managed") == "true"
                    and all(row.get(field) == "true" for field in capture.prior.BOOLEAN_KEYS)
                    for row in rows),
                f"tail-guard {name}: CSV release/oracle guard is false")
        interval = (*receipt_interval(receipt, f"tail-guard/{name}"), f"tail-guard/{name}")
        tail_intervals.append(interval)
        intervals.append(interval)

    ordered_labels = [item[2] for item in sorted(tail_intervals, key=lambda item: (item[0], item[1], item[2]))]
    expected_labels = [f"tail-guard/{slot}-{case}" for slot, _variant, _case_id, case in entries]
    require(ordered_labels == expected_labels,
            "tail-guard timestamps do not follow the frozen serial order")

    report = load_json(TAIL_GUARD_REPORT)
    require(report.get("schema") == TAIL_GUARD_REPORT_SCHEMA
            and report.get("plan_sha256") == digest(TAIL_GUARD_PLAN)
            and report.get("protocol") == plan["protocol"]
            and report.get("bindings") == plan["builds"],
            "tail-guard report header differs")
    require(report.get("summary", {}).get("runs") == len(entries)
            and report.get("summary", {}).get("cases") == len(TAIL_GUARD_CASES)
            and report.get("summary", {}).get("paired_comparisons") == 4
            and report.get("summary", {}).get("original_adverse_flags_retained")
            == original["adverse_flag_count"],
            "tail-guard report summary differs")
    report_original = report.get("original_comparison")
    require(isinstance(report_original, dict)
            and report_original.get("report_sha256") == original_sha
            and report_original.get("total_adverse_flags") == original["adverse_flag_count"],
            "tail-guard report does not retain the original flags")

    # The tail helper is itself the bounded statistical analyzer.  Replay it
    # into an auto-cleaned temporary destination, then compare both retained
    # report forms byte-for-byte without mutating the evidence tree.
    tail_module = load_analyzer(TAIL_GUARD_SCRIPT, "tail_guard.py")
    with tempfile.TemporaryDirectory(prefix="litchi-verify-0519-") as directory:
        replay_json = Path(directory) / "tail-guard.json"
        replay_markdown = Path(directory) / "tail-guard.md"
        tail_module.REPORT_JSON = replay_json
        tail_module.REPORT_MD = replay_markdown
        try:
            replay = tail_module.analyze(plan)
        except Exception as error:
            raise VerificationError(f"tail-guard analyzer replay failed: {error}") from error
        require(report == replay, "tail-guard report differs from analyzer replay")
        require(TAIL_GUARD_MARKDOWN.read_text(encoding="utf-8")
                == replay_markdown.read_text(encoding="utf-8"),
                "tail-guard markdown differs from analyzer replay")
    return {
        "status": "pass",
        "plan_sha256": digest(TAIL_GUARD_PLAN),
        "report_sha256": digest(TAIL_GUARD_REPORT),
        "markdown_sha256": digest(TAIL_GUARD_MARKDOWN),
        "runs": len(entries),
        "paired_comparisons": 4,
        "original_adverse_flags_retained": original["adverse_flag_count"],
        "analyzer_replayed": True,
    }


def verify_bundle(require_complete: bool = False) -> dict[str, Any]:
    plan = load_json(HERE / "plan.json")
    require(plan.get("base_revision") == BASE_REVISION, "plan base revision differs")
    candidate_plan = load_json(HERE / "candidate-plan.json")
    require(candidate_plan.get("base") == BASE_REVISION,
            "candidate plan base revision differs")
    intervals: list[tuple[float, float, str]] = []
    adr = verify_adr_manifest()
    probe_binding = verify_probe_binding()
    baseline_manifest, baseline_manifest_sha = verify_manifest_epoch("baseline")
    require(baseline_manifest.get(probe_binding["source"]) == probe_binding["source_sha256"],
            "allocator probe source is not bound to the baseline source manifest")
    baseline_build = verify_build("baseline", intervals)
    baseline_allocator_build = verify_allocator_build("baseline", baseline_manifest, intervals)
    baseline = {}
    for lane in BASELINE_LANES:
        build = baseline_allocator_build if lane.startswith("alloc-") else baseline_build
        baseline[lane] = verify_lane(lane, "baseline", plan, build, intervals)
        verify_order(lane, baseline[lane], plan)
    replayed_native = replay_native_analyzer(["r1", "r2"])

    candidate_dirs = [HERE / lane for lane in CANDIDATE_LANES]
    candidate_signal = ((HERE / "candidate").exists()
                        or (HERE / "allocator-candidate").exists()
                        or any(path.exists() for path in candidate_dirs))
    candidate_complete = candidate_signal and (HERE / "candidate").is_dir() \
        and (HERE / "allocator-candidate").is_dir() \
        and all(path.is_dir() for path in candidate_dirs)
    candidate = None
    candidate_build = None
    candidate_allocator_build = None
    candidate_manifest = None
    candidate_manifest_sha = None
    diff = None
    source_patch = None
    patch_replay = None
    quality = None
    if any(path.exists() for path in candidate_dirs) or (HERE / "candidate").exists():
        candidate = {}
        candidate_manifest, candidate_manifest_sha = verify_manifest_epoch("candidate")
        diff = verify_candidate_diff(baseline_manifest, candidate_manifest, candidate_plan,
                                     require_complete=require_complete)
        source_patch = verify_source_patch(diff["changed_paths"], required=require_complete)
        patch_replay = verify_patch_replay(diff["changed_paths"], baseline_manifest,
                                           candidate_manifest)
        candidate_build_path = HERE / "candidate" / "build-receipt.json"
        if candidate_build_path.exists():
            candidate_build = verify_build("candidate", intervals)
        else:
            require(not require_complete, "candidate build receipt is missing")
        allocator_build_path = HERE / "allocator-candidate" / "build-receipt.json"
        if allocator_build_path.exists():
            candidate_allocator_build = verify_allocator_build(
                "candidate", candidate_manifest, intervals
            )
        else:
            require(not require_complete, "allocator-candidate build receipt is missing")
        for lane in CANDIDATE_LANES:
            if not (HERE / lane).is_dir():
                continue
            build = (candidate_allocator_build if lane.startswith("alloc-")
                     else candidate_build)
            if build is None:
                require(not require_complete, f"{lane}: candidate build receipt is missing")
                continue
            candidate[lane] = verify_lane(lane, "candidate", plan, build, intervals)
            verify_order(lane, candidate[lane], plan)
        candidate_complete = (candidate_complete and candidate_build is not None
                              and candidate_allocator_build is not None
                              and all(lane in candidate for lane in CANDIDATE_LANES))
        if candidate_complete:
            for name in NATIVE_NAMES:
                before = baseline["r1"][name]["identity"]
                require(candidate["after-r1"][name]["identity"] == before
                        and candidate["after-r2"][name]["identity"] == before,
                        f"{name}: candidate output identity differs from baseline")
            for before_lane, after_lane, names in (
                ("alloc-baseline-r1", "alloc-candidate-r1", NATIVE_NAMES),
                ("alloc-baseline-r2", "alloc-candidate-r2", NATIVE_NAMES),
                ("profile-r1", "profile-after-r1", PROFILE_NAMES),
                ("profile-r2", "profile-after-r2", PROFILE_NAMES),
                ("hardware", "hardware-after", PROFILE_NAMES),
            ):
                for name in names:
                    require(candidate[after_lane][name]["identity"]
                            == baseline[before_lane][name]["identity"],
                            f"{name}: candidate {after_lane} identity differs from baseline")
            quality = verify_quality_receipts(
                candidate_manifest_sha, intervals, required=require_complete
            )
            replayed_native.extend(replay_native_analyzer(["after-r1", "after-r2"]))
    if not candidate_complete and require_complete:
        missing = [lane for lane in CANDIDATE_LANES if not (HERE / lane).is_dir()]
        raise VerificationError("candidate evidence is incomplete: " + ", ".join(missing))

    hardware_summary = verify_hardware_summary(plan, candidate_complete)
    focused = verify_focused_receipts(intervals)
    tail_guard = verify_tail_guard(
        intervals,
        {"baseline": baseline_build, "candidate": candidate_build},
        candidate_manifest_sha,
        required=require_complete and candidate_complete,
    )
    report_replay = verify_report_replay(required=require_complete and candidate_complete)
    allocation_report = None
    allocation_report_path = HERE / "publication-allocation-comparison.json"
    if allocation_report_path.exists():
        require(candidate_allocator_build is not None,
                "publication allocation comparison has no candidate allocator build")
        allocation_report = verify_allocation_report(
            {"baseline": baseline_allocator_build, "candidate": candidate_allocator_build},
            required=require_complete and candidate_complete,
        )
    elif require_complete and candidate_complete:
        raise VerificationError("publication allocation comparison is missing")
    analyses = verify_analyses(plan, baseline,
                               required=require_complete and candidate_complete)
    helper_manifest = verify_helper_manifest(required=False)
    sums = verify_sha256sums()
    # Focused receipts contribute intervals too.  Audit the complete interval
    # set only after every retained receipt has been read.
    timeline = verify_timeline(intervals)
    required_reports = {
        "baseline-native-analysis.json", "candidate-native-analysis.json",
        "profile-analysis.json", "profile-after-analysis.json",
        "candidate-comparison.json", "profile-comparison.json",
        "counter-comparison.json", "hardware-summary.json",
        "publication-allocation-comparison.json", "report-replay.json",
    }
    reports_complete = all((HERE / name).is_file() for name in required_reports)
    tail_complete = tail_guard is not None and tail_guard.get("status") == "pass"
    status = "pending-candidate" if not candidate_complete else "pass"
    if candidate_complete and not tail_complete:
        status = "pending-tail-guard"
    elif candidate_complete and (quality is None or quality.get("status") != "pass"
                                 or source_patch is None or not reports_complete):
        status = "pending-quality" if (quality is None or quality.get("status") != "pass"
                                        or source_patch is None) else "pending-reports"
    return {
        "schema": "litchi-0519-verifier-v1",
        "status": status,
        "base_revision": BASE_REVISION,
        "priority": "OLE2/OOXML; ODF deferred; iWork excluded",
        "baseline": {"build": baseline_build,
                      "allocator_build": baseline_allocator_build,
                      "lanes": {lane: len(value) for lane, value in baseline.items()},
                      "source_manifest_sha256": baseline_manifest_sha,
                      "capture_receipts": capture_proof(baseline)},
        "candidate": None if candidate is None else {
            "build": candidate_build,
            "allocator_build": candidate_allocator_build,
            "lanes": {lane: len(value) for lane, value in candidate.items()},
            "source_manifest_sha256": candidate_manifest_sha,
            "diff": diff,
            "source_patch": source_patch,
            "patch_replay": patch_replay,
            "capture_receipts": capture_proof(candidate),
        },
        "candidate_required_lanes": list(CANDIDATE_LANES),
        "candidate_complete": candidate_complete,
        "quality_checks": quality,
        "analyzer_replays": replayed_native,
        "hardware_summary": hardware_summary,
        "analyses": analyses,
        "helper_manifest": helper_manifest,
        "adr_manifest": adr,
        "probe_binding": probe_binding,
        "focused_checks": focused,
        "tail_guard": tail_guard,
        "report_replay": report_replay,
        "allocation_report": allocation_report,
        "reports_complete": reports_complete,
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
