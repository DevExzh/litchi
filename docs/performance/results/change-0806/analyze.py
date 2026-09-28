"""Fail-closed offline replay for the 0806 PPTX workflow packet.

This module only consumes frozen plans, retained receipts, reports, and sealed
historical oracles.  It never builds, runs a probe, invokes a profiler, or
uses a timing command.  Native elapsed values and scoped allocation counters
are paired within each fresh alternating block; no historical timing is
pooled with the 0806 measurements.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import re
import statistics
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

try:
    import custody as custody_api
except ModuleNotFoundError:
    # Keep direct import/replay from a repository-root harness deterministic;
    # normal script execution already places the packet directory on sys.path.
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import custody as custody_api


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
CUSTODY_ROOT = custody_api.ROOT.resolve()
CUSTODY_TARGET = custody_api.TARGET.resolve()
ORIGIN_PATH = PACKET / "origin.json"
SOURCE_ALLOWLIST = (
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes/tests.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
)
AMENDMENT_HELPER_FILES = (
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
)
AMENDMENT_SHARED_TEST = "crates/litchi-opc/src/xml_attributes/tests.rs"
AMENDMENT_ARCHIVES = {
    "litchi-ole-common-xml_attributes.rs": AMENDMENT_HELPER_FILES[0],
    "litchi-opc-xml_attributes.rs": AMENDMENT_HELPER_FILES[1],
    "litchi-sign-xml_attributes.rs": AMENDMENT_HELPER_FILES[2],
    "litchi-xldm-xml_attributes.rs": AMENDMENT_HELPER_FILES[3],
    "xml-minifier-xml_attributes.rs": AMENDMENT_HELPER_FILES[4],
}
PREFLIGHT_ARCHIVES = {
    **AMENDMENT_ARCHIVES,
    "litchi-opc-xml_attributes-tests.rs": AMENDMENT_SHARED_TEST,
}
VISIBILITY_SCHEMA = "litchi.performance.0806.visibility-amendment.v1"
VISIBILITY_APPLICATION_SCHEMA = (
    "litchi.performance.0806.visibility-amendment-application.v1"
)
VISIBILITY_ARCHIVE = "litchi-ole-common-xml_attributes.rs"
VISIBILITY_PRODUCTION = "crates/litchi-ole-common/src/xml_attributes.rs"
VISIBILITY_DECLARATIONS = {
    "litchi-ole-common-xml_attributes.rs": {
        "baseline": {"BytesStartExt": "pub", "CheckedAttributes": "pub"},
        "current": {"BytesStartExt": "pub(crate)", "CheckedAttributes": "pub(crate)"},
        "status": "two narrowed declarations; amendment restores both",
    },
    "litchi-opc-xml_attributes.rs": {
        "baseline": {"BytesStartExt": "pub", "CheckedAttributes": "pub"},
        "current": {"BytesStartExt": "pub", "CheckedAttributes": "pub"},
        "status": "unchanged",
    },
    "litchi-sign-xml_attributes.rs": {
        "baseline": {"BytesStartExt": "pub(crate)", "CheckedAttributes": "pub(crate)"},
        "current": {"BytesStartExt": "pub(crate)", "CheckedAttributes": "pub(crate)"},
        "status": "unchanged",
    },
    "litchi-xldm-xml_attributes.rs": {
        "baseline": {"BytesStartExt": "pub(crate)", "CheckedAttributes": "pub(crate)"},
        "current": {"BytesStartExt": "pub(crate)", "CheckedAttributes": "pub(crate)"},
        "status": "unchanged",
    },
    "xml-minifier-xml_attributes.rs": {
        "baseline": {"BytesStartExt": "pub(crate)", "CheckedAttributes": "pub(crate)"},
        "current": {"BytesStartExt": "pub(crate)", "CheckedAttributes": "pub(crate)"},
        "status": "unchanged",
    },
}
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
MODES = ("capture", "commit", "lifecycle")
CASES = tuple({"shape": shape, "mode": mode} for shape in SHAPES for mode in MODES)
ORIGINAL_SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor")
ORIGINAL_CASES = tuple({"shape": shape, "mode": mode}
                       for shape in ORIGINAL_SHAPES for mode in MODES)
LEGS = ("before", "after")
DIMENSIONS = {
    "tiny": (3, 4),
    "medium": (12, 8),
    "large": (100, 100),
    "vendor": (12, 8),
    "unicode-vendor": (12, 8),
    "valid-4attr": (12, 8),
}
VALID_FOUR_ATTRIBUTE_URIS = [
    "urn:litchi:perf:0806:extension:one",
    "urn:litchi:perf:0806:extension:two",
    "urn:litchi:perf:0806:extension:three",
    "urn:litchi:perf:0806:extension:four",
]
VALID_FOUR_ATTRIBUTE_NAMES = [
    "lx1:probeOne", "lx2:probeTwo", "lx3:probeThree", "lx4:probeFour",
]
VALID_FOUR_ATTRIBUTE_VALUES = [
    "litchi-perf-0806-valid-4attr-one",
    "litchi-perf-0806-valid-4attr-two",
    "litchi-perf-0806-valid-4attr-three",
    "litchi-perf-0806-valid-4attr-four",
]
PROBE_FILES = {
    "schema": "litchi.pptx.public-workflow-probe-0806.v1",
    "tool": "public-pptx-probe-0806",
    "marker": "litchi-perf-0780-static-mce-capabilities",
}
TIMING_SCOPES = {
    "capture": "Package::opened_presentation only",
    "commit": (
        "Transaction::commit only; package capture and one set_shape_text staging "
        "are outside the clock"
    ),
    "lifecycle": (
        "Package::opened_presentation, edit, set_shape_text, commit, "
        "apply_opened_presentation_commit, and Package::to_bytes"
    ),
}
ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
    "net_live",
    "peak_above_entry",
)
RAW_ALLOCATION_FIELDS = ALLOCATION_FIELDS[:11]
NATIVE_METRICS = ("p50", "mean", "p95", "p99")
BOOTSTRAP_SEED = 806080
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9749
BOOTSTRAP_CONFIDENCE = 0.95
HISTORICAL_PACKET = "docs/performance/results/change-0792"
PROFILE_PACKET = "docs/performance/results/change-0794"
QUALITY_COMMANDS = (
    ["cargo", "fmt", "-p", "litchi-opc", "-p", "litchi-ole-common",
     "-p", "litchi-sign", "-p", "litchi-xldm", "-p", "litchi-formula",
     "-p", "xml-minifier", "-p", "litchi-ooxml-common", "-p", "litchi-docx",
     "-p", "litchi-xlsx", "-p", "litchi-pptx", "-p", "litchi-doc",
     "-p", "litchi-xls", "-p", "litchi-ppt", "-p", "litchi-xlsb", "--", "--check"],
    ["cargo", "check", "--offline", "--locked", "-p", "litchi-opc",
     "-p", "litchi-ole-common", "-p", "litchi-sign", "-p", "litchi-xldm",
     "-p", "litchi-formula", "-p", "xml-minifier", "-p", "litchi-ooxml-common",
     "-p", "litchi-docx", "-p", "litchi-xlsx", "-p", "litchi-pptx", "-p",
     "litchi-doc", "-p", "litchi-xls", "-p", "litchi-ppt", "-p", "litchi-xlsb",
     "--all-features", "--all-targets"],
    ["cargo", "test", "--offline", "--locked", "-p", "litchi-opc", "-p",
     "litchi-ole-common", "-p", "litchi-sign", "-p", "litchi-xldm", "-p",
     "litchi-formula", "-p", "xml-minifier", "-p", "litchi-ooxml-common", "-p",
     "litchi-docx", "-p", "litchi-xlsx", "-p", "litchi-pptx", "-p", "litchi-doc",
     "-p", "litchi-xls", "-p", "litchi-ppt", "-p", "litchi-xlsb", "--all-features",
     "--", "--test-threads=2"],
    ["cargo", "clippy", "--offline", "--locked", "-p", "litchi-opc", "-p",
     "litchi-ole-common", "-p", "litchi-sign", "-p", "litchi-xldm", "-p",
     "litchi-formula", "-p", "xml-minifier", "-p", "litchi-ooxml-common", "-p",
     "litchi-docx", "-p", "litchi-xlsx", "-p", "litchi-pptx", "-p", "litchi-doc",
     "-p", "litchi-xls", "-p", "litchi-ppt", "-p", "litchi-xlsb", "--all-features",
     "--lib", "--", "-D", "warnings"],
    ["cargo", "doc", "--offline", "--locked", "-p", "litchi-opc", "-p",
     "litchi-ole-common", "-p", "litchi-sign", "-p", "litchi-xldm", "-p",
     "litchi-formula", "-p", "xml-minifier", "-p", "litchi-ooxml-common", "-p",
     "litchi-docx", "-p", "litchi-xlsx", "-p", "litchi-pptx", "-p", "litchi-doc",
     "-p", "litchi-xls", "-p", "litchi-ppt", "-p", "litchi-xlsb", "--all-features",
     "--no-deps"],
    ["python3", "-B", "tools/check_crate_boundaries.py"],
)
PROBE_WARNING_HEADLINES = (
    "warning: function `enable` is never used",
    "warning: function `record_allocation` is never used",
    "warning: function `record_deallocation` is never used",
    "warning: function `record_reallocation` is never used",
    "warning: function `record_failed_allocation` is never used",
    "warning: function `unavailable_sample` is never used",
    "warning: struct `CallbackEntryGuard` is never constructed",
    "warning: associated function `enter` is never used",
    "warning: multiple methods are never used",
    "warning: function `checked_add` is never used",
    "warning: function `checked_sub` is never used",
)
PROBE_WARNING_SUMMARY = (
    "warning: `mce-capabilities-probe` (bin \"namespace-uri-probe\") generated 11 warnings"
)
HEX = frozenset("0123456789abcdefABCDEF")


class ReplayError(RuntimeError):
    """Evidence is missing, stale, or contradictory."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def digest_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def is_revision(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 40 and all(c in HEX for c in value)


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def timestamp_utc(value: Any, label: str) -> float:
    require(isinstance(value, str)
            and re.fullmatch(r"\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?Z", value)
            is not None, f"{label} is not a UTC timestamp")
    try:
        parsed = datetime.fromisoformat(value[:-1] + "+00:00")
    except ValueError as error:
        fail(f"{label} is not parseable: {error}")
    require(parsed.tzinfo is not None and parsed.utcoffset() == timezone.utc.utcoffset(parsed),
            f"{label} is not UTC")
    return parsed.timestamp()


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def file_identity(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def origin() -> dict[str, Any]:
    value = read_json(ORIGIN_PATH)
    require(isinstance(value, dict), "origin.json is malformed")
    require(is_revision(value.get("base")), "origin base revision is invalid")
    require(set(value) == {"base", "unrelated", "worktrees"},
            "origin schema changed")
    require(ROOT.resolve() == CUSTODY_ROOT,
            "packet root differs from custody root")
    require(CUSTODY_TARGET.is_absolute()
            and CUSTODY_TARGET.name == "litchi-target-0806",
            "custody target convention changed")
    # The frozen origin deliberately records only historical workspace state.
    # Derive these operational paths from the packet-owned custody module
    # rather than mutating or extending origin.json.
    derived = dict(value)
    derived["main"] = str(CUSTODY_ROOT)
    derived["target"] = str(CUSTODY_TARGET)
    return derived


def _candidate_paths(raw: Path, *, packet_bound: bool) -> list[Path]:
    candidates: list[Path] = []
    owned = Path(origin().get("worktree", origin().get("main", ROOT))).resolve()
    if raw.is_absolute():
        try:
            candidates.append(ROOT / raw.resolve().relative_to(owned))
        except ValueError:
            pass
        parts = raw.parts
        for marker in ("change-0806", "change-0805", "change-0804", "change-0794", "change-0793", "change-0792", "build-before", "build-after", "native",
                       "allocation", "qualification", "probe-tests"):
            if marker in parts:
                index = parts.index(marker)
                if marker in {"change-0806", "change-0805", "change-0804", "change-0792", "change-0793", "change-0794"}:
                    candidates.append(PACKET.joinpath(*parts[index + 1:]))
                else:
                    candidates.append(PACKET / marker / Path(*parts[index + 1:]))
        candidates.append(raw)
    else:
        text = str(raw).replace("\\", "/")
        for packet_name in ("change-0806", "change-0805", "change-0804", "change-0794", "change-0793", "change-0792"):
            prefix = f"docs/performance/results/{packet_name}/"
            if text.startswith(prefix):
                candidates.append(PACKET / text[len(prefix):] if packet_name == "change-0806"
                                 else ROOT / "docs/performance/results" / packet_name / text[len(prefix):])
        candidates.extend((PACKET / raw, ROOT / raw))
    if not packet_bound:
        candidates.append(raw)
    return candidates


def resolve_path(value: Any, *, packet_bound: bool = True) -> Path:
    require(isinstance(value, str) and value, f"invalid artifact path: {value!r}")
    raw = Path(value)
    candidates = _candidate_paths(raw, packet_bound=packet_bound)
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            path = candidate.resolve()
            if packet_bound:
                try:
                    path.relative_to(PACKET.resolve())
                except ValueError:
                    continue
            return path
    path = (candidates[0] if candidates else raw).resolve(strict=False)
    if packet_bound:
        try:
            path.relative_to(PACKET.resolve())
        except ValueError:
            fail(f"artifact path escaped packet: {value}")
    return path


def artifact(value: Any, label: str, *, packet_bound: bool = True,
             allow_missing: bool = False) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact receipt")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = resolve_path(raw, packet_bound=packet_bound)
    if not path.is_file():
        if allow_missing:
            return None
        fail(f"missing {label}: {raw}")
    require(not path.is_symlink(), f"{label} is a symlink: {raw}")
    require(path.stat().st_size == size and sha256(path) == digest,
            f"{label} bytes or sha256 changed")
    return path


def artifact_path(value: Any, label: str, *, packet_bound: bool = True) -> Path:
    path = artifact(value, label, packet_bound=packet_bound)
    assert path is not None
    return path


def _git_output(arguments: list[str], label: str) -> str:
    try:
        return subprocess.check_output(["git", *arguments], cwd=ROOT, text=True)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {label}: {error}")


def git_blobs(revision: str, names: Iterable[str], label: str) -> dict[str, bytes]:
    ordered = list(names)
    require(is_revision(revision), f"{label} revision is invalid")
    require(len(set(ordered)) == len(ordered), f"{label} contains duplicate paths")
    require(all(name and "\n" not in name and "\0" not in name for name in ordered),
            f"{label} contains an invalid path")
    request = "".join(f"{revision}:{name}\n" for name in ordered).encode()
    try:
        result = subprocess.run(["git", "cat-file", "--batch"], cwd=ROOT,
                                input=request, stdout=subprocess.PIPE, check=True)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {label}: {error}")
    data, offset, values = result.stdout, 0, {}
    for name in ordered:
        end = data.find(b"\n", offset)
        require(end >= 0, f"{label} header is truncated: {name}")
        header = data[offset:end].split()
        require(len(header) == 3 and header[1] == b"blob",
                f"{label} entry is not a blob: {name}")
        try:
            size = int(header[2])
        except ValueError:
            fail(f"{label} size is invalid: {name}")
        offset = end + 1
        require(size >= 0 and offset + size <= len(data), f"{label} is truncated: {name}")
        values[name] = data[offset:offset + size]
        offset += size
        require(data[offset:offset + 1] == b"\n", f"{label} terminator is missing: {name}")
        offset += 1
    require(offset == len(data), f"{label} has trailing data")
    return values


def source_manifest(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is malformed")
    require(is_revision(value.get("revision")), f"{label}.revision is invalid")
    files = value.get("files")
    require(isinstance(files, dict) and files, f"{label}.files is missing")
    for name, digest in files.items():
        require(isinstance(name, str) and name and is_sha(digest),
                f"{label} contains an invalid file digest")
    return {"revision": value["revision"], "files": dict(files)}


def source_files_equal(left: dict[str, Any], right: dict[str, Any]) -> bool:
    return left.get("files") == right.get("files")


def current_source_files() -> dict[str, str]:
    try:
        raw = subprocess.check_output(
            ["git", "ls-files", "-z", "--", "crates", "Cargo.toml", "clippy.toml",
             ".cargo/config.toml", "rust-toolchain.toml"], cwd=ROOT)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read current source census: {error}")
    result: dict[str, str] = {}
    for name in (item for item in raw.decode().split("\0") if item):
        path = ROOT / name
        require(path.is_file() and not path.is_symlink(), f"live source missing: {name}")
        result[name] = sha256(path)
    return result


def load_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(isinstance(plan, dict) and plan.get("schema") == "litchi.performance.0806.v1",
            "plan schema changed")
    require(plan.get("cpu") == 12 and plan.get("cases") == list(CASES)
            and len(CASES) == 18,
            "plan cases or CPU changed")
    require(plan.get("source_allowlist") == list(SOURCE_ALLOWLIST),
            "source allowlist changed")
    require(isinstance(plan.get("scope"), str) and plan["scope"], "plan scope missing")
    require({case["shape"] for case in CASES} == set(SHAPES)
            and {case["mode"] for case in CASES} == set(MODES),
            "PPTX case matrix changed")
    native, allocation = plan.get("native"), plan.get("allocation")
    require(isinstance(native, dict) and isinstance(allocation, dict),
            "plan lanes are missing")
    require(native.get("blocks") == 6 and native.get("samples") == 30
            and native.get("warmup") == 3, "native lane changed")
    require(allocation.get("blocks") == 2 and allocation.get("samples") == 3
            and allocation.get("warmup") == 0, "allocation lane changed")
    require(native.get("orders") == [
        ["before", "after"], ["after", "before"], ["before", "after"],
        ["after", "before"], ["after", "before"], ["before", "after"],
    ], "alternating order changed")
    profile = plan.get("profile")
    require(isinstance(profile, dict)
            and profile.get("mode") == "capture"
            and profile.get("shape") == "large"
            and profile.get("samples") == 1
            and profile.get("warmup") == 0
            and profile.get("orders") == [["before", "after"], ["after", "before"]]
            and profile.get("owner") == "namespace_uri_probe::capture_region_0793"
            and profile.get("counter_match") == (
                "owner allocation flamegraph cost equals allocation_calls alone; reallocation_calls is a subset")
            and profile.get("filtered_conservation") == (
                "filtered stacks equal exact-owner subset of whole stacks")
            and profile.get("whole_conservation") == (
                "whole flamegraph sum equals print summary equals histogram count"),
            "profile lane changed")
    return plan


def check_premeasurement_inputs() -> dict[str, Any]:
    """Recheck the analysis contract that was frozen before native capture."""

    path = PACKET / "pre-measurement-inputs.json"
    if path.is_file():
        value = read_json(path)
        expected_names = {"adoption-policy.json", "analysis-plan.json", "plan.json"}
        require(isinstance(value, dict) and expected_names.issubset(value),
                "pre-measurement input set changed")
        for name, digest in value.items():
            require(is_sha(digest), f"pre-measurement digest is invalid: {name}")
            current = PACKET / name
            require(current.is_file() and sha256(current) == digest,
                    f"pre-measurement input changed: {name}")
    else:
        value = None
    contract = read_json(PACKET / "analysis-plan.json")
    require(isinstance(contract, dict)
            and contract.get("bootstrap", {}).get("seed") == BOOTSTRAP_SEED
            and contract.get("bootstrap", {}).get("resamples") == BOOTSTRAP_RESAMPLES
            and contract.get("bootstrap", {}).get("sorted_zero_based_endpoints")
            == [BOOTSTRAP_LOW_RANK, BOOTSTRAP_HIGH_RANK]
            and contract.get("process_quantile") == "nearest rank ceil(n*p)-1"
            and "all samples" in contract.get("resource_guard", ""),
            "analysis plan changed")
    qualification = PACKET / "qualification-contract.json"
    qualification_identity = (file_identity(qualification)
                              if qualification.is_file() else None)
    return {"receipt": None if value is None else file_identity(path),
            "contract": file_identity(PACKET / "analysis-plan.json"),
            "bootstrap_endpoints": [BOOTSTRAP_LOW_RANK, BOOTSTRAP_HIGH_RANK],
            "qualification_contract": qualification_identity}


def load_policy() -> dict[str, Any]:
    path = PACKET / "adoption-policy.json"
    value = read_json(path)
    require(isinstance(value, dict), "adoption policy is malformed")
    require(value.get("frozen_before_build") is True
            and value.get("allocation_count_alone_sufficient") is False
            and value.get("useful_public_workflow_benefit_required") is True,
            "adoption policy freeze changed")
    latency, benefit, memory = value.get("latency"), value.get("benefit"), value.get("memory")
    require(isinstance(latency, dict) and latency.get("metric") == "paired process p50"
            and latency.get("maximum_ratio") == 1.05
            and latency.get("bootstrap95_low_must_exceed") == 1.0
            and latency.get("resamples") == BOOTSTRAP_RESAMPLES
            and latency.get("seed") == BOOTSTRAP_SEED
            and latency.get("any_case_violation_rejects") is True,
            "latency policy changed")
    require(isinstance(benefit, dict) and benefit.get("eligible_modes") == ["capture", "lifecycle"]
            and benefit.get("minimum_improvement_percent") == 3.0
            and benefit.get("bootstrap95_high_below") == 1.0
            and benefit.get("at_least_one_case_required") is True,
            "benefit policy changed")
    require(isinstance(memory, dict)
            and memory.get("net_live_increase_allowed") == 0
            and memory.get("peak_above_entry_increase_allowed") == 0
            and memory.get("allocation_calls_increase_allowed") == 0
            and memory.get("allocated_bytes_increase_allowed") == 0,
            "memory policy changed")
    require(value.get("scope") == (
        "All eighteen cases, including unchanged ordinary/vendor controls and valid-4attr; helper gains alone insufficient."),
            "policy scope changed")
    return {"receipt": file_identity(path), "policy": value}


def check_candidate_application(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    """Bind the measured source census to the original candidate archive.

    The candidate archive is optional while the packet is being measured.  If
    it exists, every changed source file must have one byte-for-byte baseline
    and candidate copy, plus a patch and review artifact.  This keeps the
    analyzer independent of a particular candidate implementation while
    preventing an archive from silently describing a different source tree.
    """

    changed = sorted(name for name in set(before["files"]) | set(after["files"])
                     if before["files"].get(name) != after["files"].get(name))
    require(changed and set(changed).issubset(SOURCE_ALLOWLIST),
            f"source changed outside explicit allowlist: {changed}")
    application_path = PACKET / "application.json"
    require(application_path.is_file(), "candidate application witness missing")
    application = read_json(application_path)
    require(application.get("source", {}).get("files") == after["files"],
            "applied candidate source differs from measured source")
    manifest_path = artifact_path(application.get("manifest"), "candidate manifest")
    patch_witness = artifact_path(application.get("patch"), "applied candidate patch")
    manifest = read_json(manifest_path)
    require(manifest.get("base_commit") == origin()["base"], "candidate manifest base changed")
    declared_patch = manifest.get("patch")
    if isinstance(declared_patch, dict) and declared_patch.get("sha256"):
        require(sha256(patch_witness) == declared_patch["sha256"],
                "candidate manifest patch changed")
    manifest_files = manifest.get("files")
    if isinstance(manifest_files, dict):
        manifest_paths = sorted(
            value.get("production_path") for value in manifest_files.values()
            if isinstance(value, dict) and isinstance(value.get("production_path"), str)
        )
        require(manifest_paths == changed, "candidate manifest source set changed")
    else:
        require(sorted(row["production"] for row in manifest_files or []) == changed,
                "candidate manifest source set changed")
    candidate = PACKET / "candidate"
    archive = {"changed_files": changed}
    if candidate.exists():
        require(candidate.is_dir(), "candidate archive is not a directory")
        patch_candidates = [candidate / "candidate.patch", PACKET / "candidate.patch"]
        review_candidates = [candidate / "candidate-review.md", candidate / "source-review.md",
                             candidate / "design.md",
                             PACKET / "candidate-review.md", PACKET / "source-review.md"]
        patch_path = next((path for path in patch_candidates if path.is_file()), None)
        review_path = next((path for path in review_candidates if path.is_file()), None)
        require(patch_path is not None and review_path is not None,
                "candidate archive patch or review is missing")
        before_root, after_root = candidate / "before", candidate / "after"
        require(before_root.is_dir() and after_root.is_dir(),
                "candidate archive before/after directories are missing")

        def archived(root: Path, name: str, label: str) -> Path:
            exact = root / name
            if exact.is_file():
                return exact
            archive_names = {
                **{production: archive for archive, production in AMENDMENT_ARCHIVES.items()},
                AMENDMENT_SHARED_TEST: "litchi-opc-xml_attributes-tests.rs",
            }
            flat = root / archive_names.get(
                name, Path(name).parts[1] + "-" + Path(name).name
            )
            if flat.is_file() and not flat.is_symlink():
                return flat
            matches = [path for path in root.rglob(Path(name).name)
                       if path.is_file() and not path.is_symlink()]
            require(len(matches) == 1, f"{label} archive copy is missing or ambiguous")
            return matches[0]

        copies = []
        for name in changed:
            before_path = archived(before_root, name, f"candidate {name} before")
            after_path = archived(after_root, name, f"candidate {name} after")
            require(sha256(before_path) == before["files"].get(name),
                    f"candidate {name} before archive differs from baseline")
            require(sha256(after_path) == after["files"].get(name),
                    f"candidate {name} after archive differs from build source")
            copies.append({"path": name,
                           "before": file_identity(before_path),
                           "after": file_identity(after_path)})
        archive.update({"patch": file_identity(patch_path),
                        "review": file_identity(review_path), "copies": copies})
    return archive


def strict_packet_artifact(value: Any, label: str, expected: Path) -> dict[str, Any]:
    """Require a three-field artifact receipt bound to one packet path."""

    require(isinstance(value, dict) and set(value) == {"path", "bytes", "sha256"},
            f"{label} descriptor changed")
    actual = artifact_path(value, label)
    require(actual.resolve() == expected.resolve(), f"{label} path changed")
    require(value["bytes"] == expected.stat().st_size
            and value["sha256"] == sha256(expected),
            f"{label} identity changed")
    return file_identity(expected)


def preflight_artifact(value: Any, label: str, expected: Path) -> dict[str, Any]:
    """Check an amendment-preflight receipt without following aliases."""

    require(isinstance(value, dict) and set(value) == {"path", "bytes", "sha256"},
            f"{label} descriptor changed")
    raw = value["path"]
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    actual = Path(raw) if Path(raw).is_absolute() else (PACKET / "amendment-preflight" / raw)
    actual = actual.resolve()
    expected = expected.resolve()
    require(actual == expected, f"{label} path changed")
    require(actual.is_file() and not actual.is_symlink(), f"missing {label}: {raw}")
    require(value["bytes"] == actual.stat().st_size
            and value["sha256"] == sha256(actual), f"{label} identity changed")
    return {"path": str(actual.relative_to(PACKET)),
            "bytes": actual.stat().st_size, "sha256": sha256(actual)}


def preflight_source_archives() -> dict[str, dict[str, dict[str, Any]]]:
    """Return and bind the six-file before/after archive identities."""

    root = PACKET / "amendment-preflight" / "source"
    packet_root = PACKET / "amendment-preflight"
    result: dict[str, dict[str, dict[str, Any]]] = {}
    expected_names = set(PREFLIGHT_ARCHIVES)
    for leg in LEGS:
        directory = root / leg
        require(directory.is_dir() and not directory.is_symlink(),
                f"amendment preflight {leg} source archive is missing")
        actual_names = {path.name for path in directory.iterdir()
                        if path.is_file() and not path.is_symlink()}
        require(actual_names == expected_names,
                f"amendment preflight {leg} source archive set changed")
        result[leg] = {
            name: {"path": str((directory / name).relative_to(packet_root)),
                   "bytes": (directory / name).stat().st_size,
                   "sha256": sha256(directory / name)}
            for name in sorted(expected_names)
        }
    return result


def check_amendment_preflight(preflight_before: dict[str, Any],
                              amended_source: dict[str, Any]) -> dict[str, Any]:
    """Validate the fresh protected preflight bound to the quality amendment.

    This is a receipt reader only.  It checks the immutable source handoff,
    the mirror quality gates, and the complete native receipt set before
    accepting the preflight decision that authorizes the amended source to
    re-enter the frozen 0806 workflow.
    """

    root = PACKET / "amendment-preflight"
    manifest = read_json(root / "manifest.json")
    require(isinstance(manifest, dict)
            and set(manifest) == {"schema", "packet", "source", "probe",
                                  "amendment", "execution", "decision_contract"}
            and manifest["schema"] == "litchi.performance.0806.amendment-preflight-manifest.v1"
            and manifest["packet"] == "change-0806/amendment-preflight",
            "amendment preflight manifest schema changed")
    require(manifest["source"] == {
        "before": "source/before",
        "after": "source/after",
        "before_lineage": "../candidate/before",
        "after_lineage": "../candidate-quality-amendment/after",
        "handoff_before_is_not_used": (
            "../candidate-quality-amendment/before contains the already-applied 0806 candidate and is excluded from this comparison"
        ),
    }, "amendment preflight source lineage changed")
    require(manifest["probe"] == {
        "lineage": "../change-0805/probe-src",
        "case_archive": "../change-0805/cases.json",
        "fixture_archive": "../change-0805/fixtures.json",
        "schema": "litchi.attribute-boundary-probe.v1",
        "tool": "attribute-boundary-probe-0805",
        "clone_advances": [0, 1, 2, 3, 4, 5, 32, 33],
    }, "amendment preflight probe lineage changed")
    require(manifest["amendment"] == {
        "source_paths": list(AMENDMENT_HELPER_FILES),
        "shared_test_path": AMENDMENT_SHARED_TEST,
        "constructor_rewrite": {
            "from": "let mut attributes = tag.attributes(); followed by attributes.with_checks(false);",
            "to": "let attributes = tag.unchecked_attributes();",
            "algorithm_change": False,
        },
    }, "amendment preflight source operation changed")
    require(manifest["execution"] == {
        "target": "/home/zhuhe/code/litchi-target-0806/amendment-preflight",
        "cpu": 12,
        "native_blocks": 6,
        "native_samples": 30,
        "native_warmup": 3,
        "native_iterations": 4096,
        "bootstrap_seed": 806082,
        "profiles": False,
        "callgrind": False,
        "historical_timing_pooling": False,
    }, "amendment preflight execution contract changed")
    require(manifest["decision_contract"] == {
        "required_fields": [
            "advance_to_workflow_trials",
            "production_adoption",
            "protected_consume_regressions",
            "dominant_class_benefits",
        ],
        "production_adoption": False,
    }, "amendment preflight decision contract changed")

    plan = read_json(root / "plan.json")
    require(isinstance(plan, dict)
            and set(plan) == {"schema", "purpose", "lineage", "scope", "cpu", "legs",
                              "modes", "case_count", "clone_advances", "native", "analysis",
                              "policy", "quality", "amendment", "source_allowlist"}
            and plan["schema"] == "litchi.performance.0806.amendment-preflight.v1"
            and plan["cpu"] == 12 and plan["legs"] == ["before", "after"]
            and plan["modes"] == ["construct", "consume"]
            and plan["case_count"] == 39
            and plan["clone_advances"] == [0, 1, 2, 3, 4, 5, 32, 33]
            and plan["source_allowlist"] == list(SOURCE_ALLOWLIST),
            "amendment preflight plan changed")
    require(plan["lineage"] == {
        "prior_packet": "../change-0805",
        "prior_seal": "../change-0805/seal.json",
        "original_candidate_packet": "../candidate",
        "amendment_handoff": "../candidate-quality-amendment",
        "original_production_before": "../candidate/before",
        "amended_candidate_after": "../candidate-quality-amendment/after",
        "comparison": "source/before versus source/after",
        "historical_timing_pooling": False,
    }, "amendment preflight plan lineage changed")
    require(plan["scope"] == {
        "claim": "protected native micro-input timing only",
        "public_workflow_speedup": False,
        "resource_claim": False,
        "production_adoption": False,
        "profiles": False,
        "callgrind": False,
        "allocator_measurement": False,
        "native_process_elapsed_is_diagnostic": True,
    }, "amendment preflight scope changed")
    require(plan["native"] == {
        "blocks": 6,
        "orders": [
            ["before", "after"], ["after", "before"], ["before", "after"],
            ["after", "before"], ["after", "before"], ["before", "after"],
        ],
        "samples": 30,
        "warmup": 3,
        "iterations": 4096,
    }, "amendment preflight native lane changed")
    require(plan["analysis"] == {
        "bootstrap_seed": 806082,
        "bootstrap_resamples": 10000,
        "bootstrap_statistic": "median of six paired process-p50 ratios",
        "process_p50": "nearest rank ceil(n/2)-1",
        "zero_based_endpoints": [250, 9749],
        "diagnostic_ratio_above": 1.05,
        "diagnostic_ci_low_above": 1.0,
    }, "amendment preflight analysis contract changed")
    require(plan["policy"] == {
        "benefit_mode": "consume",
        "benefit_cases": ["distinct-1", "distinct-2"],
        "benefit_ratio_at_most": 0.97,
        "benefit_ci_high_below": 1.0,
        "protected_consume_cases": [
            "distinct-0", "distinct-1", "distinct-2", "duplicate-valid-after-1",
            "duplicate-valid-after-2", "duplicate-long-quoted-after-1",
            "duplicate-long-quoted-after-2", "duplicate-long-quoted-after-33",
            "duplicate-long-unterminated-after-1", "duplicate-long-unterminated-after-2",
            "duplicate-long-unterminated-after-33", "duplicate-unquoted-after-1",
            "syntax-flag-after-0", "syntax-flag-after-2", "syntax-unique-tail-after-0",
            "syntax-unique-tail-after-2", "syntax-equals-value-after-0",
            "syntax-equals-value-after-2",
        ],
        "protected_ratio_above": 1.05,
        "failure_action": "retain both archives and reject amendment; do not run public workflow adoption captures",
        "success_action": "amendment eligible only for root review; this packet never adopts production",
    }, "amendment preflight policy changed")
    require(plan["quality"] == {
        "helper_crates": ["litchi-opc", "litchi-ole-common", "litchi-sign",
                           "litchi-xldm", "xml-minifier"],
        "before_test_count": 70,
        "after_test_count": 100,
        "clippy": "offline locked workspace mirror all targets with -D warnings",
        "format": "cargo fmt --check",
        "source_paths": list(SOURCE_ALLOWLIST),
    }, "amendment preflight quality contract changed")
    require(plan["amendment"] == {
        "kind": "mechanical constructor reuse",
        "changed_helper_count": 5,
        "required_constructor_expression": "let attributes = tag.unchecked_attributes();",
        "forbidden_constructor_expression": "attributes.with_checks(false);",
        "runtime_algorithm_claim": "none; only the existing helper call route is repaired",
        "source_tests_unchanged_from_0805": True,
    }, "amendment preflight amendment contract changed")

    archive_ids = preflight_source_archives()
    for name, production in AMENDMENT_ARCHIVES.items():
        before_path = root / "source/before" / name
        after_path = root / "source/after" / name
        require(sha256(before_path) == preflight_before["files"][production]
                and sha256(after_path) == amended_source["files"][production],
                f"amendment preflight source archive differs: {production}")
    require(sha256(root / "source/before" / "litchi-opc-xml_attributes-tests.rs")
            == preflight_before["files"][AMENDMENT_SHARED_TEST]
            and sha256(root / "source/after" / "litchi-opc-xml_attributes-tests.rs")
            == amended_source["files"][AMENDMENT_SHARED_TEST],
            "amendment preflight shared test archive changed")

    quality = read_json(root / "quality/complete.json")
    require(isinstance(quality, dict)
            and set(quality) == {"schema", "rows", "test_counts", "expected_test_counts",
                                 "source_archives", "scope"}
            and quality["schema"] == "litchi.performance.0806.amendment-quality.v1"
            and quality["test_counts"] == {"before": 70, "after": 100}
            and quality["expected_test_counts"] == {"before": 70, "after": 100}
            and quality["scope"] == (
                "five helper mirror crates plus shared OPC tests; no full production-crate claim"),
            "amendment preflight quality receipt changed")
    expected_archive_digests = {
        leg: {name: sha256(root / "source" / leg / name)
              for name in sorted(PREFLIGHT_ARCHIVES)} for leg in LEGS
    }
    require(quality["source_archives"] == expected_archive_digests,
            "amendment preflight quality source archive changed")
    quality_rows = quality["rows"]
    require(isinstance(quality_rows, list) and len(quality_rows) == 6,
            "amendment preflight quality gate count changed")
    previous = None
    for index, row in enumerate(quality_rows):
        require(isinstance(row, dict) and set(row) == {
            "leg", "command", "started", "ended", "exit_code", "log"
        } and row["leg"] in LEGS and row["exit_code"] == 0
                and isinstance(row["started"], (int, float))
                and isinstance(row["ended"], (int, float))
                and row["started"] <= row["ended"]
                and (previous is None or previous <= row["started"]),
                f"amendment preflight quality row {index} changed")
        previous = row["ended"]
        expected_log = root / "quality" / f"{row['leg']}-{index % 3}.log"
        preflight_artifact(row["log"], f"amendment preflight quality log {index}", expected_log)
    require([row["leg"] for row in quality_rows] == ["before"] * 3 + ["after"] * 3,
            "amendment preflight quality leg order changed")

    target = Path(manifest["execution"]["target"])
    require(target.is_absolute() and target.name == "amendment-preflight",
            "amendment preflight target changed")
    cleanup_path = root / "cleanup.json"
    cleanup = None
    if cleanup_path.is_file():
        cleanup = read_json(cleanup_path)
        require(isinstance(cleanup, dict)
                and set(cleanup) == {
                    "schema", "target", "target_removed", "removed_target_bytes",
                    "removed_binaries", "removed_failed_binaries", "native_complete",
                }
                and cleanup["schema"] == "litchi.performance.0806.amendment-cleanup.v1"
                and cleanup["target"] == str(target)
                and cleanup["target_removed"] is True
                and not target.exists() and not target.is_symlink(),
                "amendment preflight cleanup changed")
        nonnegative_int(cleanup["removed_target_bytes"],
                        "amendment preflight cleanup removed_target_bytes")
        removed = cleanup["removed_binaries"]
        require(isinstance(removed, list) and len(removed) == 2
                and all(isinstance(item, dict)
                        and set(item) == {"path", "bytes", "sha256"}
                        and isinstance(item["path"], str)
                        and isinstance(item["bytes"], int)
                        and not isinstance(item["bytes"], bool)
                        and item["bytes"] >= 0
                        and is_sha(item["sha256"])
                        for item in removed)
                and cleanup["removed_failed_binaries"] == [],
                "amendment preflight cleanup binary count changed")
        preflight_artifact(cleanup["native_complete"],
                           "amendment preflight cleanup native witness",
                           root / "native/complete.json")

    builds: dict[str, dict[str, Any]] = {}
    expected_probe = {
        str(path.relative_to(root)): sha256(path)
        for path in (root / "probe-src").rglob("*")
        if path.is_file() and path.name != "Cargo.toml"
    }
    for leg in LEGS:
        build = read_json(root / f"build-{leg}/build.json")
        require(isinstance(build, dict)
                and set(build) == {"schema", "leg", "source", "archive", "probe",
                                   "binary", "lock", "command", "environment",
                                   "profiles", "callgrind"}
                and build["schema"] == "litchi.performance.0806.amendment-build.v1"
                and build["leg"] == leg and build["archive"] == archive_ids[leg]
                and build["probe"] == expected_probe
                and build["profiles"] is False and build["callgrind"] is False,
                f"amendment preflight {leg} build manifest changed")
        source = build["source"]
        require(isinstance(source, dict) and is_revision(source.get("revision"))
                and isinstance(source.get("files"), dict),
                f"amendment preflight {leg} build source changed")
        preflight_artifact(build["lock"], f"amendment preflight {leg} lock",
                           root / "probe-src/Cargo.lock")
        binary = build["binary"]
        require(isinstance(binary, dict) and set(binary) == {"path", "bytes", "sha256"},
                f"amendment preflight {leg} binary descriptor changed")
        binary_path = Path(binary["path"])
        if binary_path.is_file() and not binary_path.is_symlink():
            require(binary_path.stat().st_size == binary["bytes"]
                    and sha256(binary_path) == binary["sha256"],
                    f"amendment preflight {leg} binary changed")
        else:
            require(cleanup is not None and binary in cleanup["removed_binaries"],
                    f"amendment preflight {leg} binary disappeared without cleanup")
        command = build["command"]
        require(isinstance(command, dict) and command.get("exit_code") == 0
                and isinstance(command.get("command"), list)
                and command["command"][:3] == ["cargo", "build", "--offline"],
                f"amendment preflight {leg} build command changed")
        preflight_artifact(command["log"], f"amendment preflight {leg} build log",
                           root / f"build-{leg}/native.log")
        require(build["environment"].get("CARGO_BUILD_JOBS") == "2"
                and build["environment"].get("CARGO_INCREMENTAL") == "0",
                f"amendment preflight {leg} build environment changed")
        builds[leg] = build

    if cleanup is not None:
        require({(item["path"], item["bytes"], item["sha256"])
                 for item in cleanup["removed_binaries"]}
                == {(builds[leg]["binary"]["path"], builds[leg]["binary"]["bytes"],
                     builds[leg]["binary"]["sha256"])
                    for leg in LEGS},
                "amendment preflight cleanup binary identities changed")

    native_complete = read_json(root / "native/complete.json")
    expected_children = 6 * 39 * 2 * 2
    expected_samples = expected_children * 30
    require(isinstance(native_complete, dict)
            and set(native_complete) == {"schema", "children", "expected_children", "samples",
                                        "receipts", "source", "profiles", "callgrind",
                                        "historical_timing_pooling"}
            and native_complete["schema"] == "litchi.performance.0806.amendment-native.v1"
            and native_complete["children"] == expected_children
            and native_complete["expected_children"] == expected_children
            and native_complete["samples"] == expected_samples
            and native_complete["profiles"] is False
            and native_complete["callgrind"] is False
            and native_complete["historical_timing_pooling"] is False,
            "amendment preflight native completion changed")
    preflight_artifact(native_complete["source"], "amendment preflight native source",
                       root / "native/source.json")
    receipts_path = root / "native/receipts.json"
    preflight_artifact(native_complete["receipts"], "amendment preflight native receipts",
                       receipts_path)
    receipts = read_json(receipts_path)
    require(isinstance(receipts, list) and len(receipts) == expected_children,
            "amendment preflight native receipt count changed")
    case_ids = [case["id"] for case in read_json(root / "cases.json")]
    require(len(case_ids) == 39 and len(set(case_ids)) == 39,
            "amendment preflight case archive changed")
    expected_orders = plan["native"]["orders"]
    seen: set[tuple[int, str, str, str]] = set()
    for index, receipt in enumerate(receipts):
        require(isinstance(receipt, dict)
                and set(receipt) == {"schema", "block", "case", "mode", "leg", "command",
                                     "started", "ended", "exit_code", "binary", "log", "rss", "report"}
                and receipt["schema"] == "litchi.performance.0806.amendment-native-receipt.v1"
                and receipt["exit_code"] == 0 and receipt["leg"] in LEGS
                and receipt["case"] in case_ids and receipt["mode"] in plan["modes"]
                and isinstance(receipt["block"], int)
                and 0 <= receipt["block"] < 6
                and isinstance(receipt["started"], (int, float))
                and isinstance(receipt["ended"], (int, float))
                and receipt["started"] <= receipt["ended"],
                f"amendment preflight native receipt {index} changed")
        identity = (receipt["block"], receipt["case"], receipt["mode"], receipt["leg"])
        require(identity not in seen, f"duplicate amendment preflight native receipt: {identity}")
        seen.add(identity)
        block_order = expected_orders[receipt["block"]]
        case_index = case_ids.index(receipt["case"])
        mode_index = plan["modes"].index(receipt["mode"])
        require(receipt["leg"] in block_order,
                f"amendment preflight native order changed at receipt {index}")
        expected_index = (receipt["block"] * len(case_ids) * len(plan["modes"]) * 2
                          + (case_index * len(plan["modes"]) + mode_index) * 2
                          + block_order.index(receipt["leg"]))
        require(index == expected_index,
                f"amendment preflight native schedule position changed at receipt {index}")
        require(receipt["binary"] == builds[receipt["leg"]]["binary"],
                f"amendment preflight native binary custody changed at receipt {index}")
        stem = f"{receipt['block']}-{receipt['case']}-{receipt['mode']}-{receipt['leg']}"
        preflight_artifact(receipt["log"], f"amendment preflight native log {index}",
                           root / "native" / f"{stem}.log")
        preflight_artifact(receipt["rss"], f"amendment preflight native RSS {index}",
                           root / "native" / f"{stem}.rss")
        report_path = root / "native" / f"{stem}.json"
        preflight_artifact(receipt["report"], f"amendment preflight native report {index}",
                           report_path)
        report = read_json(report_path)
        require(report.get("schema") == "litchi.attribute-boundary-probe.v1"
                and report.get("tool") == "attribute-boundary-probe-0805"
                and report.get("leg") == receipt["leg"]
                and report.get("mode") == receipt["mode"]
                and report.get("case") == receipt["case"]
                and report.get("iterations") == 4096
                and report.get("warmup") == 3
                and report.get("samples_requested") == 30
                and isinstance(report.get("samples"), list)
                and len(report["samples"]) == 30,
                f"amendment preflight native report {index} changed")
    require(len(seen) == expected_children,
            "amendment preflight native schedule is incomplete")

    analysis_path = root / "analysis.json"
    decision_path = root / "decision.json"
    analysis = read_json(analysis_path)
    decision = read_json(decision_path)
    require(isinstance(analysis, dict)
            and set(analysis) == {"schema", "rows", "policy", "claims"}
            and analysis["schema"] == "litchi.performance.0806.amendment-analysis.v1"
            and analysis["policy"] == plan["policy"]
            and analysis["claims"] == plan["scope"],
            "amendment preflight analysis schema changed")
    rows = analysis["rows"]
    require(isinstance(rows, list) and len(rows) == 39 * 2,
            "amendment preflight analysis row count changed")
    row_keys = set()
    for row in rows:
        require(isinstance(row, dict)
                and set(row) == {"case", "mode", "process_p50_before", "process_p50_after",
                                 "paired_ratios", "ratio_median", "bootstrap_ci_low",
                                 "bootstrap_ci_high", "change_percent_median",
                                 "diagnostic_regression"}
                and row["case"] in case_ids and row["mode"] in plan["modes"]
                and (row["case"], row["mode"]) not in row_keys
                and isinstance(row["process_p50_before"], list)
                and isinstance(row["process_p50_after"], list)
                and isinstance(row["paired_ratios"], list)
                and len(row["process_p50_before"]) == 6
                and len(row["process_p50_after"]) == 6
                and len(row["paired_ratios"]) == 6
                and isinstance(row["diagnostic_regression"], bool),
                "amendment preflight analysis row changed")
        row_keys.add((row["case"], row["mode"]))
        for number in (*row["process_p50_before"], *row["process_p50_after"],
                       *row["paired_ratios"], row["ratio_median"],
                       row["bootstrap_ci_low"], row["bootstrap_ci_high"],
                       row["change_percent_median"]):
            finite_number(number, "amendment preflight analysis number")
    require(row_keys == {(case, mode) for case in case_ids for mode in plan["modes"]},
            "amendment preflight analysis coverage changed")
    require(isinstance(decision, dict)
            and set(decision) == {"schema", "legacy_schema", "packet", "seed",
                                  "advance_to_workflow_trials", "production_adoption",
                                  "protected_consume_regressions", "dominant_class_benefits",
                                  "benefit_policy_passed", "all_consume_regressions", "counts",
                                  "analysis", "independent_audit", "custody"}
            and decision["schema"] == "litchi.performance.0806.amendment-decision.v1"
            and decision["legacy_schema"] == "litchi.performance.0805.preflight-decision.v1"
            and decision["packet"] == "change-0806/amendment-preflight"
            and decision["seed"] == 806082
            and decision["advance_to_workflow_trials"] is True
            and decision["production_adoption"] is False
            and decision["protected_consume_regressions"] == []
            and decision["dominant_class_benefits"] == {"distinct-1": True, "distinct-2": True}
            and decision["benefit_policy_passed"] is True,
            "amendment preflight decision changed")
    strict_preflight_artifact = preflight_artifact(
        decision["analysis"], "amendment preflight decision analysis", analysis_path)
    strict_independent_audit = preflight_artifact(
        decision["independent_audit"], "amendment preflight independent audit",
        root / "root-native-audit.json")
    independent_audit = read_json(root / "root-native-audit.json")
    require(isinstance(independent_audit, dict)
            and independent_audit.get("schema")
            == "litchi.performance.0806.amendment-root-native-audit.v1"
            and independent_audit.get("passed") is True
            and independent_audit.get("matches_primary_analysis") is True
            and independent_audit.get("native_reports") == 936
            and independent_audit.get("native_samples") == 28_080
            and independent_audit.get("advance_to_workflow_trials") is True
            and independent_audit.get("production_adoption") is False
            and independent_audit.get("protected_consume_regressions") == []
            and independent_audit.get("dominant_class_benefits") == {
                "distinct-1": True, "distinct-2": True,
            }, "amendment preflight independent audit changed")
    require(decision["all_consume_regressions"] == [
        {
            "case": row["case"], "mode": row["mode"],
            "ratio_median": row["ratio_median"],
            "bootstrap_ci_low": row["bootstrap_ci_low"],
            "bootstrap_ci_high": row["bootstrap_ci_high"],
            "change_percent_median": row["change_percent_median"],
        }
        for row in rows
        if row["mode"] == "consume"
        and row["ratio_median"] > plan["analysis"]["diagnostic_ratio_above"]
    ], "amendment preflight consume regression list changed")
    require(decision["counts"] == {
        "case_count": 39,
        "native_reports": expected_children,
        "native_children": expected_children,
        "native_samples": expected_samples,
    }, "amendment preflight decision counts changed")
    custody = decision["custody"]
    require(isinstance(custody, dict)
            and set(custody) == {"archives", "builds", "native", "probe", "source",
                                  "quality", "profiles", "callgrind",
                                  "historical_timing_pooling"}
            and custody["archives"] == archive_ids
            and custody["profiles"] is False and custody["callgrind"] is False
            and custody["historical_timing_pooling"] is False,
            "amendment preflight decision custody changed")
    for leg in LEGS:
        preflight_artifact(custody["builds"][leg],
                           f"amendment preflight decision {leg} build",
                           root / f"build-{leg}/build.json")
    preflight_artifact(custody["native"], "amendment preflight decision native",
                       root / "native/complete.json")
    preflight_artifact(custody["source"], "amendment preflight decision source",
                       root / "source.json")
    preflight_artifact(custody["quality"], "amendment preflight decision quality",
                       root / "quality/complete.json")
    require(custody["probe"] == {
        str(path.relative_to(root)): sha256(path)
        for path in (root / "probe-src").rglob("*")
        if path.is_file() and path.name != "Cargo.toml"
    }, "amendment preflight decision probe custody changed")

    return {
        "schema": decision["schema"],
        "decision": file_identity(decision_path),
        "analysis": strict_preflight_artifact,
        "independent_audit": strict_independent_audit,
        "plan": file_identity(root / "plan.json"),
        "source_archives": archive_ids,
        "quality": {"path": rel(root / "quality/complete.json"),
                    "test_counts": quality["test_counts"], "gates": len(quality_rows)},
        "native": {"children": expected_children, "samples": expected_samples},
        "advance_to_workflow_trials": True,
        "production_adoption": False,
        "protected_consume_regressions": [],
        "dominant_class_benefits": dict(decision["dominant_class_benefits"]),
    }


def check_quality_amendment(before: dict[str, Any],
                            candidate_source: dict[str, Any],
                            quality_source: dict[str, Any]) -> dict[str, Any]:
    """Bind the intermediate source to the reviewed five-helper amendment."""

    root = PACKET / "candidate-quality-amendment"
    manifest_path = root / "manifest.json"
    patch_path = root / "candidate-quality-amendment.patch"
    review_path = root / "source-review.md"
    manifest = read_json(manifest_path)
    require(isinstance(manifest, dict)
            and set(manifest) == {"schema", "change", "base_commit", "parent_candidate",
                                  "files", "shared_files", "patch", "amendment", "review"}
            and manifest["schema"] == "litchi.performance.0806.quality-amendment.v1"
            and manifest["change"] == 806
            and manifest["base_commit"] == before["revision"],
            "quality amendment manifest identity changed")

    expected_parent = {
        "manifest": PACKET / "candidate/manifest.json",
        "patch": PACKET / "candidate/candidate.patch",
        "application": PACKET / "application.json",
        "source": PACKET / "source.json",
    }
    parent = manifest["parent_candidate"]
    require(isinstance(parent, dict) and set(parent) == set(expected_parent),
            "quality amendment parent witness changed")
    for name, expected in expected_parent.items():
        strict_packet_artifact(parent[name], f"quality amendment parent {name}", expected)
    original_application = read_json(expected_parent["application"])
    require(isinstance(original_application, dict)
            and set(original_application) == {"manifest", "patch", "source"}
            and original_application["source"] == candidate_source,
            "original candidate application was altered")
    require(candidate_source["revision"] == before["revision"],
            "original candidate source revision changed")

    amendment_files = manifest["files"]
    require(isinstance(amendment_files, dict)
            and set(amendment_files) == set(AMENDMENT_ARCHIVES),
            "quality amendment helper archive set changed")
    amended_file_receipts: dict[str, dict[str, Any]] = {}
    for archive_name, production in AMENDMENT_ARCHIVES.items():
        row = amendment_files[archive_name]
        require(isinstance(row, dict)
                and set(row) == {"production_path", "before", "after"}
                and row["production_path"] == production,
                f"quality amendment file record changed: {production}")
        before_path = root / "before" / archive_name
        after_path = root / "after" / archive_name
        strict_packet_artifact(row["before"], f"quality amendment before {production}", before_path)
        strict_packet_artifact(row["after"], f"quality amendment after {production}", after_path)
        require(sha256(before_path) == candidate_source["files"][production]
                and sha256(after_path) == quality_source["files"][production]
                and sha256(before_path) != sha256(after_path),
                f"quality amendment source archive differs: {production}")
        old = before_path.read_bytes()
        new = after_path.read_bytes()
        constructor = (
            b"    #[allow(clippy::disallowed_methods)]\n"
            b"    #[inline]\n"
            b"    fn new(tag: &'a BytesStart<'a>) -> Self {\n"
            b"        let mut attributes = tag.attributes();\n"
            b"        attributes.with_checks(false);"
        )
        replacement = (
            b"    #[inline]\n"
            b"    fn new(tag: &'a BytesStart<'a>) -> Self {\n"
            b"        let attributes = tag.unchecked_attributes();"
        )
        require(old.count(constructor) == 1
                and new == old.replace(constructor, replacement, 1),
                f"quality amendment constructor transformation changed: {production}")
        amended_file_receipts[production] = file_identity(after_path)

    shared = manifest["shared_files"]
    require(isinstance(shared, dict)
            and set(shared) == {"litchi-opc-xml_attributes-tests.rs"},
            "quality amendment shared-file witness changed")
    shared_row = shared["litchi-opc-xml_attributes-tests.rs"]
    require(isinstance(shared_row, dict)
            and set(shared_row) == {"production_path", "original_candidate_after",
                                    "amendment_action"}
            and shared_row["production_path"] == AMENDMENT_SHARED_TEST
            and shared_row["amendment_action"]
            == "byte-identical; omitted from this five-helper amendment",
            "quality amendment shared-file record changed")
    shared_path = PACKET / "candidate/after/litchi-opc-xml_attributes-tests.rs"
    strict_packet_artifact(shared_row["original_candidate_after"],
                           "quality amendment shared test", shared_path)
    require(sha256(shared_path) == candidate_source["files"][AMENDMENT_SHARED_TEST]
            and quality_source["files"][AMENDMENT_SHARED_TEST]
            == candidate_source["files"][AMENDMENT_SHARED_TEST],
            "quality amendment shared test changed")

    strict_packet_artifact(manifest["patch"], "quality amendment patch", patch_path)
    require(manifest["review"] == str(review_path)
            and review_path.is_file() and not review_path.is_symlink(),
            "quality amendment review witness changed")
    patch_text = patch_path.read_text(errors="replace")
    require(patch_text.count("diff --git a/") == 5
            and "tests.rs" not in patch_text
            and "litchi-formula" not in patch_text
            and "let attributes = tag.unchecked_attributes();" in patch_text
            and "attributes.with_checks(false);" in patch_text,
            "quality amendment patch scope changed")
    require(manifest["amendment"] == {
        "reason": "The exact 0805 candidate left CheckedAttributes::unchecked_attributes unused in litchi-sign under warnings-denied probe quality.",
        "operation": "In each of the five helper files, CheckedAttributes::new now obtains its unchecked Attributes iterator through the existing BytesStartExt::unchecked_attributes helper.",
        "removed_allow": "The constructor-local clippy::disallowed_methods allow is removed because the constructor no longer calls quick-xml directly; the helper's existing allow remains in place at its only direct with_checks(false) call.",
        "algorithm_state_and_public_api": "Unchanged. The same unchecked iterator, tag reference, phase initialization, and iterator state are retained; the OPC test source is unchanged.",
        "runtime_requalification_required": True,
        "production_apply_required": True,
    }, "quality amendment rationale changed")

    application_path = PACKET / "quality-amendment-application.json"
    application = read_json(application_path)
    require(isinstance(application, dict)
            and set(application) == {"schema", "original_application", "manifest", "patch",
                                     "preflight", "source"}
            and application["schema"] == "litchi.performance.0806.quality-amendment-application.v1",
            "quality amendment application schema changed")
    strict_packet_artifact(application["original_application"],
                           "quality amendment original application",
                           PACKET / "application.json")
    strict_packet_artifact(application["manifest"], "quality amendment application manifest",
                           manifest_path)
    strict_packet_artifact(application["patch"], "quality amendment application patch",
                           patch_path)
    decision = check_amendment_preflight(before, quality_source)
    strict_packet_artifact(application["preflight"], "quality amendment application preflight",
                           PACKET / "amendment-preflight/decision.json")
    require(application["source"] == quality_source,
            "quality amendment application source differs from intermediate build")
    changed = sorted(name for name in set(candidate_source["files"]) | set(quality_source["files"])
                     if candidate_source["files"].get(name) != quality_source["files"].get(name))
    require(changed == list(AMENDMENT_HELPER_FILES),
            "quality amendment changed source outside five helpers")
    final_changed = sorted(name for name in set(before["files"]) | set(quality_source["files"])
                           if before["files"].get(name) != quality_source["files"].get(name))
    require(final_changed == list(SOURCE_ALLOWLIST),
            "amended source change set differs from frozen allowlist")
    require(quality_source["revision"] == before["revision"],
            "quality amendment source revision differs from origin")
    for production, receipt in amended_file_receipts.items():
        require(quality_source["files"][production] == receipt["sha256"],
                f"quality amendment source witness differs: {production}")
    return {
        "changed_files": final_changed,
        "original_candidate": check_candidate_application(before, candidate_source),
        "application": file_identity(application_path),
        "manifest": file_identity(manifest_path),
        "patch": file_identity(patch_path),
        "preflight": decision,
        "source": quality_source,
    }


def check_visibility_amendment(quality_source: dict[str, Any],
                               final_source: dict[str, Any]) -> dict[str, Any]:
    """Bind the final source to the reviewed two-token OLE visibility repair."""

    root = PACKET / "candidate-visibility-amendment"
    manifest_path = root / "manifest.json"
    patch_path = root / "candidate-visibility-amendment.patch"
    review_path = root / "source-review.md"
    expected_inventory = {
        "manifest.json", "candidate-visibility-amendment.patch", "source-review.md",
        f"before/{VISIBILITY_ARCHIVE}", f"after/{VISIBILITY_ARCHIVE}",
    }
    inventory = {str(path.relative_to(root)) for path in root.rglob("*")
                 if path.is_file() and not path.is_symlink()}
    require(inventory == expected_inventory,
            "visibility amendment archive inventory changed")

    manifest = read_json(manifest_path)
    require(isinstance(manifest, dict)
            and set(manifest) == {
                "schema", "change", "base_commit", "parent_application", "files",
                "patch", "visibility_audit", "scope", "review",
            }
            and manifest["schema"] == VISIBILITY_SCHEMA
            and manifest["change"] == 806
            and manifest["base_commit"] == quality_source["revision"],
            "visibility amendment manifest identity changed")
    parent_path = PACKET / "quality-amendment-application.json"
    strict_packet_artifact(manifest["parent_application"],
                           "visibility amendment parent application", parent_path)
    parent_application = read_json(parent_path)
    require(isinstance(parent_application, dict)
            and set(parent_application) == {
                "schema", "original_application", "manifest", "patch", "preflight", "source",
            }
            and parent_application["schema"]
            == "litchi.performance.0806.quality-amendment-application.v1"
            and parent_application["source"] == quality_source,
            "visibility amendment parent source changed")
    strict_packet_artifact(parent_application["original_application"],
                           "visibility parent original application", PACKET / "application.json")
    strict_packet_artifact(parent_application["preflight"],
                           "visibility parent preflight", PACKET / "amendment-preflight/decision.json")

    amendment_files = manifest["files"]
    require(isinstance(amendment_files, dict)
            and set(amendment_files) == {VISIBILITY_ARCHIVE},
            "visibility amendment file set changed")
    row = amendment_files[VISIBILITY_ARCHIVE]
    before_path = root / "before" / VISIBILITY_ARCHIVE
    after_path = root / "after" / VISIBILITY_ARCHIVE
    require(isinstance(row, dict)
            and set(row) == {"production_path", "before", "after"}
            and row["production_path"] == VISIBILITY_PRODUCTION,
            "visibility amendment file record changed")
    strict_packet_artifact(row["before"], "visibility amendment before source", before_path)
    strict_packet_artifact(row["after"], "visibility amendment after source", after_path)
    require(sha256(before_path) == quality_source["files"][VISIBILITY_PRODUCTION]
            and sha256(after_path) == final_source["files"][VISIBILITY_PRODUCTION],
            "visibility amendment source boundary changed")
    old = before_path.read_bytes()
    new = after_path.read_bytes()
    require(old.count(b"pub(crate) trait BytesStartExt") == 1
            and old.count(b"pub(crate) struct CheckedAttributes") == 1
            and new == old.replace(b"pub(crate) trait BytesStartExt",
                                   b"pub trait BytesStartExt", 1).replace(
                                       b"pub(crate) struct CheckedAttributes",
                                       b"pub struct CheckedAttributes", 1),
            "visibility amendment transformation changed")

    baseline_hashes = {
        name: sha256(PACKET / "candidate/before" / name)
        for name in sorted(VISIBILITY_DECLARATIONS)
    }
    current_hashes = {
        name: quality_source["files"][AMENDMENT_ARCHIVES[name]]
        for name in sorted(VISIBILITY_DECLARATIONS)
    }
    require(manifest["visibility_audit"] == {
        "baseline_source": "docs/performance/results/change-0806/candidate/before/*.rs",
        "current_source": (
            "the five helper files at the parent quality-amendment application boundary"
        ),
        "baseline_hashes": baseline_hashes,
        "current_hashes": current_hashes,
        "declarations": VISIBILITY_DECLARATIONS,
        "result": (
            "The OLE helper's two public declarations are the only visibility narrowings "
            "across the five helper files; all other item declarations retain their baseline visibility."
        ),
    }, "visibility audit changed")
    require(manifest["scope"] == {
        "changed_tokens": 2,
        "changed_files": 1,
        "public_api_restored": [
            "litchi_ole_common::xml_attributes::BytesStartExt",
            "litchi_ole_common::xml_attributes::CheckedAttributes",
        ],
        "algorithm_or_behavior_change": False,
        "runtime_requalification_required": True,
        "production_apply_required": True,
    }, "visibility amendment scope changed")
    require(manifest["review"] == str(review_path),
            "visibility amendment review path changed")
    require(review_path.is_file() and not review_path.is_symlink(),
            "visibility amendment review is missing")

    strict_packet_artifact(manifest["patch"], "visibility amendment patch", patch_path)
    patch_text = patch_path.read_text(errors="replace")
    require(patch_text.count("diff --git a/") == 1
            and "diff --git a/crates/litchi-ole-common/src/xml_attributes.rs "
            in patch_text
            and "pub(crate) trait BytesStartExt" in patch_text
            and "pub trait BytesStartExt" in patch_text
            and "pub(crate) struct CheckedAttributes" in patch_text
            and "pub struct CheckedAttributes" in patch_text
            and "crates/litchi-opc" not in patch_text
            and "tests.rs" not in patch_text,
            "visibility amendment patch scope changed")

    changed = sorted(name for name in set(quality_source["files"]) | set(final_source["files"])
                     if quality_source["files"].get(name) != final_source["files"].get(name))
    require(changed == [VISIBILITY_PRODUCTION],
            "visibility amendment changed source outside OLE helper")
    require(final_source["revision"] == quality_source["revision"],
            "visibility amendment source revision changed")

    application_path = PACKET / "visibility-amendment-application.json"
    application = read_json(application_path)
    require(isinstance(application, dict)
            and set(application) == {
                "schema", "original_application", "manifest", "patch", "source",
            }
            and application["schema"] == VISIBILITY_APPLICATION_SCHEMA
            and application["source"] == final_source,
            "visibility amendment application schema changed")
    strict_packet_artifact(application["original_application"],
                           "visibility application parent", parent_path)
    strict_packet_artifact(application["manifest"],
                           "visibility application manifest", manifest_path)
    strict_packet_artifact(application["patch"],
                           "visibility application patch", patch_path)
    return {
        "changed_files": list(SOURCE_ALLOWLIST),
        "visibility_changed_files": changed,
        "application": file_identity(application_path),
        "manifest": file_identity(manifest_path),
        "patch": file_identity(patch_path),
        "parent_application": file_identity(parent_path),
        "source": final_source,
        "quality_source": quality_source,
        "algorithm_or_behavior_change": False,
    }


def probe_files() -> dict[str, str]:
    root = PACKET / "probe-src"
    require(root.is_dir(), "probe source directory is missing")
    return {str(path.relative_to(PACKET)): sha256(path) for path in root.rglob("*")
            if path.is_file() and path.name not in {"Cargo.lock", "Cargo.toml"}}


def fixed_release_profile() -> None:
    path = PACKET / "probe-src/Cargo.toml.template"
    text = path.read_text() if path.is_file() else ""
    require(all(line in text for line in (
        "[profile.release]", "opt-level = 3", "debug = 1", "lto = \"thin\"",
        "codegen-units = 1", "panic = \"unwind\"",
    )), "probe release profile changed")


def normalize(value: str) -> str:
    owned = str(Path(origin().get("main", ROOT)).resolve())
    return value.replace(owned, str(ROOT.resolve()))


def expected_build_command(feature: str | None) -> list[str]:
    command = ["cargo", "build", "--offline", "--release", "--manifest-path",
               str(PACKET / "probe-src/Cargo.toml")]
    if feature:
        command += ["--features", feature]
    command.append("--locked")
    return command


def build_kind(command: list[str]) -> str:
    if "allocator-metrics" in command:
        return "allocation"
    if "capture-profile" in command:
        return "profile"
    return "native"


def load_frozen_inputs(directory: Path, label: str) -> dict[str, str]:
    value = read_json(directory / "frozen-inputs.json")
    expected_names = {"plan.json", "adoption-policy.json", "build.py", "capture.py",
                      "architecture-inputs.json", "analysis-plan.json", "quality.py",
                      "probe_quality.py",
                      "profile.py", "origin.json", "inheritance.json", "host.json"}
    require(isinstance(value, dict) and set(value) == expected_names,
            f"{label} frozen input set changed")
    result = {}
    for name, digest in value.items():
        require(is_sha(digest), f"{label} frozen input digest invalid: {name}")
        path = PACKET / name
        require(path.is_file() and sha256(path) == digest,
                f"{label} frozen input changed: {name}")
        result[name] = digest
    return result


def cleanup_witness() -> tuple[dict[str, Any] | None, bool]:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None, False
    value = read_json(path)
    require(isinstance(value, dict)
            and set(value) == {"removed_binaries", "removed_target_bytes", "target",
                               "target_removed"}, "cleanup schema changed")
    require(value.get("target") == origin()["target"] and value.get("target_removed") is True,
            "cleanup target witness changed")
    nonnegative_int(value.get("removed_target_bytes"), "cleanup removed_target_bytes")
    binaries = value.get("removed_binaries")
    require(isinstance(binaries, list) and len(binaries) == 6,
            "cleanup binary witness cardinality changed")
    return value, True


def cleanup_matches(cleanup: Any, receipt: dict[str, Any]) -> bool:
    if not isinstance(cleanup, dict):
        return False
    expected = (receipt.get("path"), receipt.get("bytes"), receipt.get("sha256"))
    for item in cleanup.get("removed_binaries", []):
        if isinstance(item, dict) and (item.get("path"), item.get("bytes"), item.get("sha256")) == expected:
            return True
    return False


def validate_binary(receipt: Any, label: str, cleanup: Any, cleanup_ok: bool) -> None:
    require(isinstance(receipt, dict), f"{label} receipt missing")
    path = artifact(receipt, label, packet_bound=False, allow_missing=True)
    if path is not None:
        return
    require(cleanup_ok and cleanup_matches(cleanup, receipt),
            f"{label} missing without exact cleanup witness")


def load_builds(plan: dict[str, Any]) -> tuple[dict[str, Any], Any, bool]:
    fixed_release_profile()
    cleanup, cleanup_ok = cleanup_witness()
    builds: dict[str, Any] = {}
    for leg in LEGS:
        directory = PACKET / f"build-{leg}"
        manifest = read_json(directory / "build.json")
        require(isinstance(manifest, dict), f"{leg} build manifest malformed")
        frozen = load_frozen_inputs(directory, leg)
        source_path = artifact_path(manifest.get("source"), f"{leg} build source")
        source = source_manifest(read_json(source_path), f"{leg} source")
        inventory = manifest.get("probe")
        require(inventory == probe_files(), f"{leg} probe inventory changed")
        lock = manifest.get("lock")
        lock_path = artifact_path(lock, f"{leg} probe lock")
        require(lock_path == (PACKET / "probe-src/Cargo.lock").resolve(),
                f"{leg} lock path did not relocate")
        binaries = manifest.get("binaries")
        require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation", "profile"},
                f"{leg} binary map changed")
        for name, receipt in binaries.items():
            validate_binary(receipt, f"{leg} {name} binary", cleanup, cleanup_ok)
        rows = manifest.get("rows")
        require(isinstance(rows, list) and len(rows) == 3, f"{leg} build rows incomplete")
        expected = {"native": expected_build_command(None),
                    "allocation": expected_build_command("allocator-metrics"),
                    "profile": expected_build_command("capture-profile")}
        seen: set[str] = set()
        for row in rows:
            require(isinstance(row, dict) and row.get("exit_code") == 0,
                    f"{leg} build command failed")
            command = row.get("command")
            require(isinstance(command, list), f"{leg} build command malformed")
            command = [normalize(item) for item in command]
            kind = build_kind(command)
            require(kind not in seen and command == expected[kind],
                    f"{leg} {kind} build command changed")
            seen.add(kind)
            artifact_path(row.get("log"), f"{leg} {kind} build log")
        require(seen == {"native", "allocation", "profile"}, f"{leg} build rows incomplete")
        env = manifest.get("environment")
        require(isinstance(env, dict) and env.get("CARGO_BUILD_JOBS") == "2"
                and env.get("CARGO_INCREMENTAL") == "0",
                f"{leg} build environment changed")
        builds[leg] = {"manifest": manifest, "source": source, "source_path": source_path,
                       "probe": inventory, "lock": lock, "binaries": binaries,
                       "frozen": frozen}
    before, after = builds["before"]["source"], builds["after"]["source"]
    changed = sorted(name for name in set(before["files"]) | set(after["files"])
                     if before["files"].get(name) != after["files"].get(name))
    allowed = set(SOURCE_ALLOWLIST)
    require(all(name in before["files"] and name in after["files"] for name in SOURCE_ALLOWLIST),
            "candidate source allowlist file is missing from source census")
    require(changed and set(changed).issubset(allowed),
            f"source changed outside explicit allowlist: {changed}")
    require(before["revision"] == origin()["base"], "before source revision differs from origin")
    require(builds["before"]["lock"]["sha256"] == builds["after"]["lock"]["sha256"],
            "probe lock changed between builds")
    require(builds["before"]["probe"] == builds["after"]["probe"],
            "probe source changed between builds")
    require(builds["before"]["frozen"] == builds["after"]["frozen"],
            "frozen inputs changed between builds")
    for lane in ("native", "allocation", "qualification"):
        path = PACKET / lane / "source.json"
        if path.is_file():
            expected_source = before if lane == "qualification" else after
            require(source_files_equal(source_manifest(read_json(path), f"{lane} source"),
                                       expected_source), f"{lane} source differs from build")
    return builds, cleanup, cleanup_ok


def load_architecture_inputs() -> dict[str, Any]:
    path = PACKET / "architecture-inputs.json"
    value = read_json(path)
    require(isinstance(value, dict) and len(value) == 35, "architecture input count changed")
    for name, digest in value.items():
        require(isinstance(name, str) and not name.startswith("/") and is_sha(digest),
                f"architecture input invalid: {name}")
        live = ROOT / name
        require(live.is_file() and sha256(live) == digest,
                f"architecture input changed: {name}")
    base = origin()["base"]
    blobs = git_blobs(base, value, "architecture inputs")
    require({name: digest_bytes(data) for name, data in blobs.items()} == value,
            "architecture origin blobs changed")
    return {"receipt": file_identity(path), "revision": base, "count": 35,
            "files": dict(value), "live_files_match": True, "origin_blob_hashes_match": True}


def check_quality(after_source: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "quality.json"
    value = read_json(path)
    require(isinstance(value, dict), "quality.json malformed")
    source_path = artifact_path(value.get("source"), "quality source")
    require(source_files_equal(source_manifest(read_json(source_path), "quality source"),
                               after_source), "quality source differs from after build")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 6, "quality gate count changed")
    logs = []
    test_summary = None
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0
                and row.get("command") == QUALITY_COMMANDS[index],
                f"quality gate {index} changed or failed")
        log_path = artifact_path(row.get("log"), f"quality gate {index} log")
        logs.append(rel(log_path))
        if index == 2:
            matches = re.findall(
                r"^test result: (?:ok|FAILED)\.\s+(\d+) passed; (\d+) failed; "
                r"(\d+) ignored; (\d+) measured; (\d+) filtered out;",
                log_path.read_text(), re.MULTILINE)
            require(matches, "quality test gate has no parseable result")
            test_summary = {"suites": len(matches),
                            "passed": sum(int(row[0]) for row in matches),
                            "failed": sum(int(row[1]) for row in matches),
                            "ignored": sum(int(row[2]) for row in matches)}
            require(test_summary["failed"] == 0, "quality test gate contains failures")
    env = value.get("environment")
    require(isinstance(env, dict) and env.get("CARGO_BUILD_JOBS") == "2"
            and env.get("CARGO_INCREMENTAL") == "0"
            and env.get("CARGO_PROFILE_DEV_DEBUG") == "0"
            and env.get("RUSTDOCFLAGS") == "-D warnings",
            "quality environment changed")
    require(test_summary is not None, "quality test summary is missing")
    return {"gates": 6, "commands": [list(row["command"]) for row in rows],
            "logs": logs, "source": rel(source_path), "test_summary": test_summary}


def check_build_warning_scope(builds: dict[str, Any]) -> dict[str, Any]:
    """Retain the known standalone-probe warnings without treating them as gates."""

    rows = []
    for leg in LEGS:
        manifest = builds[leg]["manifest"]
        for row in manifest["rows"]:
            kind = build_kind(row["command"])
            log_path = artifact_path(row["log"], f"{leg} {kind} build log")
            text = log_path.read_text()
            require("error:" not in text.lower(), f"{leg} {kind} build log has an error")
            warning_lines = [line for line in text.splitlines() if line.startswith("warning:")]
            generated = [line for line in text.splitlines()
                         if "mce-capabilities-probe" in line and "generated" in line]
            if kind == "allocation":
                require(not warning_lines and not generated,
                        f"{leg} allocation build warning scope changed")
                summary = None
            else:
                require(warning_lines == [*PROBE_WARNING_HEADLINES, PROBE_WARNING_SUMMARY]
                        and generated == [PROBE_WARNING_SUMMARY],
                        f"{leg} {kind} retained warning scope changed")
                summary = PROBE_WARNING_SUMMARY
            rows.append({"leg": leg, "kind": kind, "warning_lines": len(warning_lines),
                         "summary": summary})
    return {"rows": rows, "scope": "standalone probe warnings retained; no build errors"}


def check_probe_tests(builds: dict[str, Any]) -> dict[str, Any]:
    """Verify the probe-only format/test/Clippy gates captured by root.

    These gates are deliberately separate from the production-crate quality
    lane.  Their source and generated probe inventory must remain bound to the
    corresponding before or after build, and every command must have passed.
    """

    expected = [
        ["cargo", "fmt", "--manifest-path", str(PACKET / "probe-src/Cargo.toml"),
         "--", "--check"],
        ["cargo", "test", "--offline", "--locked", "--release", "--manifest-path",
         str(PACKET / "probe-src/Cargo.toml"), "--all-features", "--",
         "--test-threads=1"],
        ["cargo", "clippy", "--offline", "--locked", "--release", "--manifest-path",
         str(PACKET / "probe-src/Cargo.toml"), "--all-features", "--all-targets",
         "--", "-D", "warnings"],
    ]
    output: dict[str, Any] = {"commands": expected, "legs": {}, "gates": 0,
                               "rerun": True}
    for leg in LEGS:
        directory = PACKET / f"probe-quality-{leg}"
        complete = read_json(directory / "complete.json")
        inputs_path = artifact_path(complete.get("inputs"), f"{leg} probe quality inputs")
        inputs = read_json(inputs_path)
        require(inputs.get("source") == builds[leg]["source"],
                f"{leg} probe quality source differs from build")
        probe = inputs.get("probe")
        expected_probe = {
            str(path.relative_to(PACKET / "probe-src")): sha256(path)
            for path in (PACKET / "probe-src").rglob("*")
            if path.is_file()
        }
        require(isinstance(probe, dict) and probe == expected_probe,
                f"{leg} probe quality inventory differs from build")
        require(complete.get("receipts") is not None,
                f"{leg} probe quality receipts are missing")
        receipt_path = artifact_path(complete["receipts"], f"{leg} probe quality receipts")
        rows = read_json(receipt_path)
        require(isinstance(rows, list) and len(rows) == 3,
                f"{leg} probe quality gate count changed")
        checked = []
        previous = float("-inf")
        for index, row in enumerate(rows):
            require(row.get("command") == expected[index]
                    and row.get("exit_code") == 0,
                    f"{leg} probe quality command {index} changed or failed")
            started, ended = float(row["started"]), float(row["ended"])
            require(started <= ended and previous <= started,
                    f"{leg} probe quality receipt order changed")
            previous = ended
            log = artifact_path(row.get("log"), f"{leg} probe quality log {index}")
            checked.append({"command": list(row["command"]), "log": rel(log),
                            "exit_code": row["exit_code"]})
        output["legs"][leg] = {"inputs": rel(inputs_path), "receipts": rel(receipt_path),
                                "gates": checked}
        output["gates"] += len(rows)
    require(output["gates"] == 6, "probe quality aggregate gate count changed")
    return output


def prior_seal(packet: Path, label: str) -> dict[str, Any]:
    seal_path = packet / "seal.json"
    seal = read_json(seal_path)
    require(isinstance(seal, dict) and isinstance(seal.get("files"), dict),
            f"{label} seal malformed")
    files = seal["files"]
    for name, digest in files.items():
        path = packet / name
        require(path.is_file() and not path.is_symlink() and sha256(path) == digest,
                f"{label} sealed file changed: {name}")
    actual = {str(path.relative_to(packet)): sha256(path) for path in packet.rglob("*")
              if path.is_file() and path.name != "seal.json"}
    require(actual == files, f"{label} seal inventory changed")
    return seal


def load_inheritance() -> dict[str, Any]:
    value = read_json(PACKET / "inheritance.json")
    require(isinstance(value, dict), "inheritance receipt malformed")
    fixture = ROOT / "docs/performance/results/change-0792"
    harness = ROOT / "docs/performance/results/change-0794"
    prior = ROOT / "docs/performance/results/change-0805"
    fixture_seal_path = fixture / "seal.json"
    harness_seal_path = harness / "seal.json"
    prior_seal_path = prior / "seal.json"
    require(value.get("fixture_packet") == "../change-0792"
            and is_sha(value.get("fixture_seal"))
            and fixture_seal_path.is_file()
            and sha256(fixture_seal_path) == value["fixture_seal"],
            "fixture inheritance reference changed")
    require(value.get("harness_packet") == "../change-0794"
            and is_sha(value.get("harness_seal"))
            and harness_seal_path.is_file()
            and sha256(harness_seal_path) == value["harness_seal"],
            "harness inheritance reference changed")
    require(value.get("prior_packet") == "../change-0805"
            and is_sha(value.get("prior_seal"))
            and prior_seal_path.is_file()
            and sha256(prior_seal_path) == value["prior_seal"],
            "prior packet inheritance reference changed")
    fixture_value = prior_seal(fixture, "0792 fixture")
    harness_value = prior_seal(harness, "0794 harness")
    prior_value = prior_seal(prior, "0805 prior packet")

    production = value.get("production_source")
    production_path = artifact_path(production, "inherited production source", packet_bound=False)
    production_manifest = source_manifest(read_json(production_path),
                                          "inherited production source")
    probe = value.get("probe_reference")
    require(isinstance(probe, dict) and set(probe) == {
        "Cargo.lock", "Cargo.toml.template", "src/allocation_metrics.rs",
        "src/counting_allocator.rs", "src/main.rs"},
            "probe inheritance references changed")
    lineage = probe_lineage(probe, harness)
    return {"fixture_packet": "docs/performance/results/change-0792",
            "fixture_seal": value["fixture_seal"],
            "fixture_sealed_files": len(fixture_value["files"]),
            "harness_packet": "docs/performance/results/change-0794",
            "harness_seal": value["harness_seal"],
            "harness_sealed_files": len(harness_value["files"]),
            "prior_packet": "docs/performance/results/change-0805",
            "prior_seal": value["prior_seal"],
            "prior_sealed_files": len(prior_value["files"]),
            "production_source": file_identity(production_path),
            "production_source_files": production_manifest["files"],
            "probe_references": sorted(probe), "probe_lineage": lineage,
            "timings_imported": False}


def probe_lineage(references: dict[str, Any], harness: Path) -> dict[str, Any]:
    """Bind the 0806 probe to the sealed 0794 copy and reviewed setup fixes.

    The 0806 public shape changes ``main.rs``.  The allocator support files
    retain their 0794 implementation and identity-only renames in the
    immutable setup archive, followed by two exact test-isolation cfg fixes:
    the global allocator is disabled under ``cfg(test)`` and the unavailable
    sample is compiled only without the allocator feature.  No broad source
    exemption is accepted.
    """

    historical_root = harness / "probe-src"
    current_root = PACKET / "probe-src"
    setup_roots = sorted(
        (path for path in PACKET.glob("probe-src-setup-*") if path.is_dir()),
        key=lambda path: int(path.name.rsplit("-", 1)[1])
        if path.name.rsplit("-", 1)[1].isdigit() else -1,
    )
    require(setup_roots and setup_roots[0].name == "probe-src-setup-0",
            "immutable probe setup archive is missing")
    require([int(path.name.rsplit("-", 1)[1]) for path in setup_roots]
            == list(range(len(setup_roots))),
            "probe setup archive sequence changed")
    setup_root = setup_roots[0]
    expected_names = {
        "Cargo.lock", "Cargo.toml.template", "src/allocation_metrics.rs",
        "src/counting_allocator.rs", "src/main.rs",
    }
    require(set(references) == expected_names, "probe inheritance files changed")
    current_hashes: dict[str, str] = {}
    fixes: dict[str, str] = {}

    def bytes_at(root: Path, name: str, label: str) -> bytes:
        path = root / name
        require(path.is_file() and not path.is_symlink(), f"{label} is missing")
        return path.read_bytes()

    for name, digest in references.items():
        historical = bytes_at(historical_root, name, f"0794 {name}")
        setup = bytes_at(setup_root, name, f"0806 setup {name}")
        current = bytes_at(current_root, name, f"0806 current {name}")
        require(is_sha(digest) and digest_bytes(historical) == digest,
                f"probe historical reference changed: {name}")
        current_hashes[name] = digest_bytes(current)
        if name in {"Cargo.lock", "Cargo.toml.template"}:
            require(setup == historical and current == setup,
                    f"probe immutable dependency input changed: {name}")
            continue
        if name == "src/main.rs":
            require(b"valid-4attr" in setup and b"valid-4attr" in current,
                    "0806 main probe shape is missing")
            continue
        if name == "src/allocation_metrics.rs":
            identity_normalized = historical.replace(
                b"0785 PPTX probe", b"0806 PPTX probe"
            ).replace(b"namespace-uri-probe-0785", b"namespace-uri-probe-0806")
        elif name == "src/counting_allocator.rs":
            identity_normalized = historical.replace(
                b"0785 PPTX namespace URI probe",
                b"0806 PPTX public-workflow probe",
            )
        else:
            identity_normalized = historical
        require(identity_normalized == setup,
                f"probe support lineage differs beyond identity strings: {name}")
        if name == "src/counting_allocator.rs":
            old = bytes((10,)) + b"#[global_allocator]" + bytes((10,))
            new = (b"\n// Unit tests invoke this wrapper directly while the counter tests exercise\n"
                   b"// synthetic callbacks. Keeping the process allocator unwrapped in the test\n"
                   b"// harness prevents unrelated test runtime allocations from entering those\n"
                   b"// synthetic regions; release binaries retain the production probe wrapper.\n"
                   b"#[cfg_attr(not(test), global_allocator)]\n")
            require(setup.count(old) == 1 and current == setup.replace(old, new, 1),
                    "counting allocator test-isolation amendment changed")
            fixes[name] = "exact cfg(test) global-allocator amendment"
        elif name == "src/allocation_metrics.rs":
            old = b"pub(crate) fn unavailable_sample()"
            new = b"#[cfg(not(feature = \"allocator-metrics\"))]\n" + old
            require(setup.count(old) == 1 and current.count(new) == 1,
                    "allocation unavailable-sample cfg amendment changed")
            fixes[name] = "exact non-allocator unavailable-sample amendment"
        else:
            fail(f"unrecognized probe lineage file: {name}")
    require(len(setup_roots) >= 2, "immutable probe setup-1 archive is missing")
    setup_one = setup_roots[1]
    setup_zero_counting = bytes_at(setup_root, "src/counting_allocator.rs",
                                   "0806 setup-0 counting allocator")
    setup_one_counting = bytes_at(setup_one, "src/counting_allocator.rs",
                                  "0806 setup-1 counting allocator")
    old = bytes((10,)) + b"#[global_allocator]" + bytes((10,))
    new = (bytes((10,))
           + b"// Unit tests invoke this wrapper directly while the counter tests exercise"
           + bytes((10,))
           + b"// synthetic callbacks. Keeping the process allocator unwrapped in the test"
           + bytes((10,))
           + b"// harness prevents unrelated test runtime allocations from entering those"
           + bytes((10,))
           + b"// synthetic regions; release binaries retain the production probe wrapper."
           + bytes((10,))
           + b"#[cfg_attr(not(test), global_allocator)]"
           + bytes((10,)))
    require(setup_zero_counting.count(old) == 1
            and setup_one_counting == setup_zero_counting.replace(old, new, 1)
            and current_hashes["src/counting_allocator.rs"]
            == digest_bytes(setup_one_counting),
            "counting allocator setup repair changed")
    setup_zero_metrics = bytes_at(setup_root, "src/allocation_metrics.rs",
                                  "0806 setup-0 allocation metrics")
    setup_one_metrics = bytes_at(setup_one, "src/allocation_metrics.rs",
                                 "0806 setup-1 allocation metrics")
    old = b"pub(crate) fn unavailable_sample()"
    new = b'#[cfg(not(feature = "allocator-metrics"))]' + bytes((10,)) + old
    require(setup_zero_metrics.count(old) == 1
            and setup_one_metrics == setup_zero_metrics.replace(old, new, 1),
            "allocation unavailable-sample setup repair changed")
    fixes["setup-1"] = (
        "exact cfg(test) global-allocator and non-allocator unavailable-sample repairs"
    )
    require(len(setup_roots) == 3,
            "probe setup archive set must contain exactly setup-0 through setup-2")
    setup_two = setup_roots[2]
    for name in expected_names - {"src/main.rs"}:
        require(bytes_at(setup_two, name, f"0806 setup-2 {name}")
                == bytes_at(current_root, name, f"0806 current {name}"),
                f"probe setup-2 support input differs from reviewed current: {name}")
    setup_one_main = bytes_at(setup_one, "src/main.rs", "0806 setup-1 main")
    setup_two_main = bytes_at(setup_two, "src/main.rs", "0806 setup-2 main")
    newline = bytes((10,))
    test_marker = newline + b"#[cfg(test)]" + newline + b"mod tests {"
    test_start = setup_one_main.find(test_marker)
    require(test_start >= 0, "setup-1 probe test module is missing")
    test_end = setup_one_main.find(newline + b"fn identity(", test_start)
    require(test_end > test_start, "setup-1 probe test module boundary changed")
    test_block = setup_one_main[test_start:test_end]
    setup_two_expected = setup_one_main[:test_start] + setup_one_main[test_end:]
    old_doc = (b"//!   serialization as one public operation." + newline
               + b"//! Package")
    new_doc = (b"//!   serialization as one public operation." + newline
               + b"//!" + newline + b"//! Package")
    require(setup_two_expected.count(old_doc) == 1,
            "setup-2 documentation continuation boundary changed")
    setup_two_expected = setup_two_expected.replace(old_doc, new_doc, 1)
    setup_two_expected += test_block
    require(setup_two_main == setup_two_expected,
            "setup-2 test-module placement or documentation repair changed")
    fixes["setup-2"] = "exact test-module relocation and documentation continuation repair"

    audit_path = PACKET / "probe-amendment-audit.json"
    audit = read_json(audit_path)
    require(isinstance(audit, dict)
            and set(audit) == {
                "schema", "passed", "reader", "files", "changes",
                "new_round_trip_test_sha256",
                "counter_logic_and_layout_preserved",
                "fixture_logic_and_tests_preserved",
            }
            and audit.get("schema") == "litchi.performance.0806.probe-amendment-audit.v1"
            and audit.get("passed") is True
            and audit.get("counter_logic_and_layout_preserved") is True
            and audit.get("fixture_logic_and_tests_preserved") is True,
            "probe amendment audit schema or result changed")
    expected_changes = [
        "test-only allocator registration isolation",
        "feature gate on unused fallback sample helper",
        "seven non-test dead-code allowances on retained support items",
        "unchanged test module relocated to file end",
        "explicit valid-4attr serialization name",
        "six-shape round-trip test with test-only derives",
        "comments and module documentation whitespace",
    ]
    require(audit.get("changes") == expected_changes,
            "probe amendment audit change inventory changed")
    require(audit.get("new_round_trip_test_sha256")
            == "942595b375a60d421c7638d8c8bbd6e9b8283350d1b7c0727471ffb211f81f98",
            "probe shape round-trip amendment changed")

    def audited_file(raw: Any, expected: Path, label: str) -> dict[str, Any]:
        require(isinstance(raw, dict) and set(raw) == {"path", "bytes", "sha256"},
                f"{label} descriptor changed")
        actual = artifact_path(raw, label)
        require(actual.resolve() == expected.resolve()
                and raw["bytes"] == expected.stat().st_size
                and raw["sha256"] == sha256(expected),
                f"{label} identity changed")
        return file_identity(expected)

    reader = audited_file(audit["reader"], PACKET / "probe_amendment_audit.py",
                          "probe amendment audit reader")
    audited = {}
    for name in ("main.rs", "allocation_metrics.rs", "counting_allocator.rs"):
        row = audit["files"].get(name)
        require(isinstance(row, dict) and set(row) == {"initial", "current"},
                f"probe amendment audit files changed: {name}")
        audited[name] = {
            "initial": audited_file(row["initial"], setup_root / "src" / name,
                                     f"probe amendment initial {name}"),
            "current": audited_file(row["current"], current_root / "src" / name,
                                     f"probe amendment current {name}"),
        }
    fixes["probe-amendment-audit"] = (
        "exact seven dead-code allowances, comments/doc whitespace, and unchanged test relocation"
    )
    latest = setup_roots[-1]
    latest_hashes = {
        name: digest_bytes(bytes_at(latest, name, f"0806 {latest.name} {name}"))
        for name in expected_names
    }
    return {"historical_packet": "docs/performance/results/change-0794/probe-src",
            "setup_packet": latest.name,
            "setup_archives": [path.name for path in setup_roots],
            "current_packet": "probe-src", "historical_references": dict(references),
            "setup_hashes": latest_hashes, "current_hashes": current_hashes,
            "exact_fixes": fixes, "amendment_audit": file_identity(audit_path),
            "amendment_audit_reader": reader, "amendment_audit_files": audited}


def load_prior_qualification() -> dict[str, Any]:
    packet = ROOT / HISTORICAL_PACKET
    seal = prior_seal(packet, "0792 fixture qualification")
    source_path = packet / "build-after/source.json"
    require(seal["files"].get("build-after/source.json") == sha256(source_path),
            "0792 after source is not sealed")
    source = source_manifest(read_json(source_path), "0792 after source")
    rows = []
    for case in ORIGINAL_CASES:
        stem = f"0-{case['shape']}-{case['mode']}-before"
        relative = f"qualification/{stem}.json"
        path = packet / relative
        require(seal["files"].get(relative) == sha256(path),
                f"0792 fixture qualification seal missing: {stem}")
        report = read_json(path)
        require(isinstance(report, dict) and is_sha(report.get("source", {}).get("sha256")),
                f"0792 fixture qualification report malformed: {stem}")
        sample = report.get("samples")
        require(isinstance(sample, list) and len(sample) == 1, f"0792 fixture qualification samples changed: {stem}")
        output = sample[0].get("output")
        require(isinstance(output, dict) and is_sha(output.get("sha256")),
                f"0792 fixture qualification output missing: {stem}")
        rows.append({"case": f"{case['shape']}/{case['mode']}",
                     "report_source": {"bytes": report["source"]["bytes"],
                                       "sha256": report["source"]["sha256"]},
                     "output": {"bytes": output["bytes"], "sha256": output["sha256"]},
                     "report_sha256": sha256(path)})
    return {"packet": HISTORICAL_PACKET, "seal_schema": seal.get("schema"),
            "source": source, "reports": rows,
            "timings_imported": False}


def before_qualification_contract(entries: list[dict[str, Any]],
                                  builds: dict[str, Any]) -> dict[str, Any]:
    """Check the strict, pre-after-build four-attribute qualification oracle.

    Root writes ``qualification-four-attr.json`` immediately after the single
    before-only qualification lane and before compiling the after leg.  The
    contract has one mode-specific oracle per valid-4attr row, plus immutable
    baseline build/probe descriptors.  It records package identities and the
    complete semantic/preservation verification object, never elapsed values;
    this reader deliberately has no compatibility aliases or optional fields.
    """

    path = PACKET / "qualification-four-attr.json"
    require(path.is_file(), "qualification-four-attr.json is missing")
    contract = read_json(path)
    require(isinstance(contract, dict)
            and set(contract) == {"schema", "freeze", "baseline", "fixture", "oracles"}
            and contract.get("schema") == (
                "litchi.performance.0806.qualification-four-attr.v1"),
            "four-attribute qualification oracle schema changed")

    freeze = contract["freeze"]
    require(isinstance(freeze, dict)
            and set(freeze) == {"stage", "after_before_qualification",
                                "before_after_build", "timings_imported",
                                "frozen_at_utc"}
            and freeze.get("stage") == "after-before-qualification"
            and freeze.get("after_before_qualification") is True
            and freeze.get("before_after_build") is True
            and freeze.get("timings_imported") is False
            and isinstance(freeze.get("frozen_at_utc"), str),
            "four-attribute qualification freeze boundary changed")
    frozen_at = timestamp_utc(freeze["frozen_at_utc"],
                              "four-attribute qualification freeze")

    baseline = contract["baseline"]
    require(isinstance(baseline, dict) and set(baseline) == {"build", "probe"},
            "four-attribute baseline descriptor changed")
    build_descriptor = baseline["build"]
    require(isinstance(build_descriptor, dict)
            and set(build_descriptor) == {"source", "binaries"}
            and build_descriptor["source"] == builds["before"]["source"]
            and build_descriptor["binaries"] == builds["before"]["binaries"],
            "four-attribute baseline build descriptor differs from before build")
    probe_descriptor = baseline["probe"]
    require(isinstance(probe_descriptor, dict)
            and set(probe_descriptor) == {"schema", "tool", "marker", "files", "lock"}
            and probe_descriptor["schema"] == PROBE_FILES["schema"]
            and probe_descriptor["tool"] == PROBE_FILES["tool"]
            and probe_descriptor["marker"] == PROBE_FILES["marker"]
            and probe_descriptor["files"] == builds["before"]["probe"]
            and probe_descriptor["lock"] == builds["before"]["lock"],
            "four-attribute baseline probe descriptor differs from before build")

    fixture = contract["fixture"]
    require(isinstance(fixture, dict)
            and fixture == {
                "injection": "valid-four-distinct-namespaced-extension-attributes",
                "slide_parts": DIMENSIONS["valid-4attr"][0],
                "replaced_text_tags": (DIMENSIONS["valid-4attr"][0]
                                        * DIMENSIONS["valid-4attr"][1]),
                "namespace_declarations": 4,
                "namespaced_attributes": 4,
                "namespace_uris": VALID_FOUR_ATTRIBUTE_URIS,
                "attribute_names": VALID_FOUR_ATTRIBUTE_NAMES,
            }, "four-attribute qualification fixture descriptor changed")

    rows = {entry["identity"]["mode"]: entry for entry in entries
            if entry["identity"]["shape"] == "valid-4attr"
            and entry["identity"]["lane"] == "qualification"}
    require(set(rows) == set(MODES), "four-attribute qualification coverage changed")
    qualification_end = max(entry["ended"] for entry in entries
                            if entry["identity"]["lane"] == "qualification")
    require(frozen_at > qualification_end,
            "four-attribute qualification freeze predates receipt completion")
    oracles = contract["oracles"]
    require(isinstance(oracles, list) and len(oracles) == len(MODES),
            "four-attribute qualification oracle count changed")
    seen: set[str] = set()
    observations = []
    for oracle in oracles:
        require(isinstance(oracle, dict)
                and set(oracle) == {"shape", "mode", "report", "source", "output",
                                    "fixture", "verification"}
                and oracle.get("shape") == "valid-4attr"
                and oracle.get("mode") in MODES
                and oracle["mode"] not in seen,
                "four-attribute mode oracle descriptor changed")
        mode = oracle["mode"]
        seen.add(mode)
        entry = rows[mode]
        report_path = entry["report_path"]
        report = read_json(report_path)
        sample = report["samples"][0]
        require(oracle["report"] == file_identity(report_path)
                and oracle["source"] == report["source"]
                and oracle["output"] == sample["output"]
                and oracle["fixture"] == report["fixture"]
                and oracle["verification"] == sample["verification"],
                f"four-attribute {mode} qualification oracle differs from report")
        require("elapsed_ns" not in oracle["verification"],
                f"four-attribute {mode} qualification oracle imports timing")
        observations.append({"mode": mode, "report": file_identity(report_path)})
    require(seen == set(MODES), "four-attribute qualification mode coverage changed")
    return {"path": rel(path), "sha256": sha256(path), "timings_imported": False,
            "frozen_at_utc": freeze["frozen_at_utc"],
            "qualification_ended_at": qualification_end,
            "rows": observations, "schema": contract["schema"]}


def check_after_build_inputs(qualification: dict[str, Any],
                             builds: dict[str, Any],
                             qualification_entries: list[dict[str, Any]]) -> dict[str, Any]:
    """Verify the witness frozen immediately before the after compilation.

    The witness is intentionally separate from ``build.py`` frozen inputs:
    that driver is shared by both legs, while this handoff only exists after
    baseline qualification and the final visibility amendment application. It
    binds the oracle, application witness, and applied source manifest with one
    UTC timestamp.
    """

    path = PACKET / "after-build-inputs.json"
    require(path.is_file(), "after-build-inputs.json is missing")
    value = read_json(path)
    require(isinstance(value, dict)
            and set(value) == {"schema", "qualification", "application", "source",
                               "frozen_at_utc"}
            and value.get("schema") == "litchi.performance.0806.after-build-inputs.v1",
            "after-build-inputs schema changed")

    def bound_artifact(raw: Any, label: str, expected: Path) -> dict[str, Any]:
        require(isinstance(raw, dict) and set(raw) == {"path", "bytes", "sha256"},
                f"{label} descriptor changed")
        actual = artifact_path(raw, label)
        require(actual.resolve() == expected.resolve(),
                f"{label} path changed")
        return file_identity(actual)

    qualification_path = PACKET / "qualification-four-attr.json"
    application_path = PACKET / "visibility-amendment-application.json"
    qualification_identity = bound_artifact(value["qualification"],
                                            "after-build qualification",
                                            qualification_path)
    application_identity = bound_artifact(value["application"],
                                          "after-build application",
                                          application_path)
    contract_identity = {"path": qualification_identity["path"],
                         "bytes": qualification_identity["bytes"],
                         "sha256": qualification_identity["sha256"]}
    require(qualification_identity == contract_identity,
            "after-build qualification identity is malformed")
    require(qualification_identity["sha256"] == qualification["sha256"],
            "after-build qualification differs from inspected contract")

    application = read_json(application_path)
    require(isinstance(application, dict)
            and set(application) == {"schema", "original_application", "manifest",
                                     "patch", "source"}
            and application.get("schema") == VISIBILITY_APPLICATION_SCHEMA
            and isinstance(application.get("source"), dict),
            "after-build application witness changed")
    strict_packet_artifact(application["original_application"],
                           "after-build original application",
                           PACKET / "quality-amendment-application.json")
    strict_packet_artifact(application["manifest"], "after-build visibility manifest",
                           PACKET / "candidate-visibility-amendment/manifest.json")
    strict_packet_artifact(application["patch"], "after-build visibility patch",
                           PACKET / "candidate-visibility-amendment/candidate-visibility-amendment.patch")
    parent_application = read_json(PACKET / "quality-amendment-application.json")
    require(isinstance(parent_application, dict)
            and set(parent_application) == {
                "schema", "original_application", "manifest", "patch", "preflight", "source",
            }
            and parent_application.get("schema")
            == "litchi.performance.0806.quality-amendment-application.v1",
            "after-build parent application witness changed")
    strict_packet_artifact(parent_application["original_application"],
                           "after-build parent original application",
                           PACKET / "application.json")
    strict_packet_artifact(parent_application["manifest"],
                           "after-build parent amendment manifest",
                           PACKET / "candidate-quality-amendment/manifest.json")
    strict_packet_artifact(parent_application["patch"],
                           "after-build parent amendment patch",
                           PACKET / "candidate-quality-amendment/candidate-quality-amendment.patch")
    strict_packet_artifact(parent_application["preflight"],
                           "after-build parent preflight",
                           PACKET / "amendment-preflight/decision.json")
    source = source_manifest(value["source"], "after-build source manifest")
    require(value["source"] == application["source"],
            "after-build source differs from application witness")
    require(source == builds["after"]["source"],
            "after-build source differs from after build source")

    frozen_at = timestamp_utc(value["frozen_at_utc"], "after-build inputs freeze")
    qualification_end = max(entry["ended"] for entry in qualification_entries)
    qualification_freeze = timestamp_utc(qualification["frozen_at_utc"],
                                         "four-attribute qualification freeze")
    require(frozen_at > qualification_end and frozen_at > qualification_freeze,
            "after-build inputs freeze is not after qualification")
    build_rows = builds["after"]["manifest"].get("rows")
    require(isinstance(build_rows, list) and build_rows,
            "after build command receipts are missing")
    starts = []
    for index, row in enumerate(build_rows):
        require(isinstance(row, dict), f"after build row {index} is malformed")
        started, ended = row.get("started"), row.get("ended")
        finite_number(started, f"after build row {index} started")
        finite_number(ended, f"after build row {index} ended")
        require(started <= ended, f"after build row {index} interval changed")
        starts.append(float(started))
    after_build_started = min(starts)
    require(frozen_at < after_build_started,
            "after-build inputs freeze does not precede after compilation")
    return {"path": rel(path), "sha256": sha256(path),
            "schema": value["schema"], "frozen_at_utc": value["frozen_at_utc"],
            "qualification": qualification_identity,
            "application": application_identity,
            "source": source,
            "qualification_ended_at": qualification_end,
            "qualification_frozen_at_utc": qualification["frozen_at_utc"],
            "after_build_started_at": after_build_started,
            "timings_imported": False}


def freeze_qualification_contract() -> dict[str, Any]:
    """Write the three mode-specific oracle after baseline qualification.

    Root invokes this command after ``capture.py qualification`` and after an
    independent review, while ``build-after/build.json`` is still absent.  It
    performs the same report checks used by the later full replay and writes
    the exact strict shape consumed by :func:`before_qualification_contract`.
    """

    output = PACKET / "qualification-four-attr.json"
    require(not output.exists(), "qualification-four-attr.json already exists")
    require(not (PACKET / "build-after").exists()
            and not (PACKET / "build-after").is_symlink(),
            "four-attribute oracle must be frozen before after build")
    plan = load_plan()
    build_manifest = read_json(PACKET / "build-before/build.json")
    source_path = artifact_path(build_manifest.get("source"),
                                "before qualification build source")
    source = source_manifest(read_json(source_path), "before qualification source")
    require(source["revision"] == origin()["base"],
            "before qualification source revision differs from origin")
    binaries = build_manifest.get("binaries")
    require(isinstance(binaries, dict)
            and set(binaries) == {"native", "allocation", "profile"},
            "before qualification binary descriptor changed")
    for kind, value in binaries.items():
        artifact(value, f"before qualification {kind} binary", packet_bound=False)
    probe = build_manifest.get("probe")
    lock = build_manifest.get("lock")
    require(isinstance(probe, dict) and probe == {
                str(path.relative_to(PACKET)): sha256(path)
                for path in (PACKET / "probe-src").rglob("*")
                if path.is_file() and path.name not in {"Cargo.lock", "Cargo.toml"}},
            "before qualification probe descriptor changed")
    artifact(lock, "before qualification probe lock")

    complete = read_json(PACKET / "qualification/complete.json")
    complete_source = artifact_path(complete.get("source"),
                                    "before qualification complete source")
    require(source_files_equal(source_manifest(read_json(complete_source),
                                               "before qualification complete source"),
                               source), "qualification source differs from before build")
    receipts_path = artifact_path(complete.get("receipts"),
                                  "before qualification receipts")
    rows = read_json(receipts_path)
    jobs = expected_jobs(plan, "qualification")
    require(isinstance(rows, list) and len(rows) == len(jobs),
            "before qualification receipt count changed")
    build = {"source": source, "binaries": binaries}
    valid_reports: dict[str, dict[str, Any]] = {}
    for index, (row, job) in enumerate(zip(rows, jobs)):
        require(isinstance(row, dict)
                and all(row.get(key) == job[key]
                        for key in ("lane", "block", "shape", "mode", "leg"))
                and row.get("exit_code") == 0,
                f"before qualification row {index} changed")
        started, ended = row.get("started"), row.get("ended")
        finite_number(started, f"before qualification row {index} started")
        finite_number(ended, f"before qualification row {index} ended")
        require(started <= ended,
                f"before qualification row {index} interval changed")
        report_path = artifact_path(row.get("report"),
                                    f"before qualification row {index} report")
        report = read_json(report_path)
        validate_report(report, job, "allocation", binaries["allocation"],
                        f"before qualification row {index}")
        if job["shape"] == "valid-4attr":
            valid_reports[job["mode"]] = {"path": report_path, "report": report}
    require(set(valid_reports) == set(MODES),
            "before qualification valid-4attr mode coverage changed")
    fixture = valid_reports["capture"]["report"]["fixture"]
    for mode in MODES:
        require(valid_reports[mode]["report"]["fixture"] == fixture,
                "before qualification fixture identity differs by mode")
    require(fixture == {
        "injection": "valid-four-distinct-namespaced-extension-attributes",
        "slide_parts": DIMENSIONS["valid-4attr"][0],
        "replaced_text_tags": DIMENSIONS["valid-4attr"][0] * DIMENSIONS["valid-4attr"][1],
        "namespace_declarations": 4,
        "namespaced_attributes": 4,
        "namespace_uris": VALID_FOUR_ATTRIBUTE_URIS,
        "attribute_names": VALID_FOUR_ATTRIBUTE_NAMES,
    }, "before qualification fixture identity changed")
    qualification_end = max(float(row["ended"]) for row in rows)
    frozen_at = datetime.now(timezone.utc)
    require(frozen_at.timestamp() > qualification_end,
            "qualification contract would predate receipt completion")
    contract = {
        "schema": "litchi.performance.0806.qualification-four-attr.v1",
        "freeze": {
            "stage": "after-before-qualification",
            "after_before_qualification": True,
            "before_after_build": True,
            "timings_imported": False,
            "frozen_at_utc": frozen_at.isoformat().replace("+00:00", "Z"),
        },
        "baseline": {
            "build": build,
            "probe": {"schema": PROBE_FILES["schema"],
                      "tool": PROBE_FILES["tool"],
                      "marker": PROBE_FILES["marker"],
                      "files": probe, "lock": lock},
        },
        "fixture": fixture,
        "oracles": [],
    }
    for mode in MODES:
        item = valid_reports[mode]
        report_path, report = item["path"], item["report"]
        sample = report["samples"][0]
        verification = sample["verification"]
        require("elapsed_ns" not in verification,
                f"before qualification {mode} verification contains timing")
        contract["oracles"].append({
            "shape": "valid-4attr", "mode": mode,
            "report": file_identity(report_path),
            "source": report["source"], "output": sample["output"],
            "fixture": report["fixture"], "verification": verification,
        })
    output.write_text(json.dumps(contract, indent=2, sort_keys=True) + "\n")
    return {"path": rel(output), "sha256": sha256(output),
            "timings_imported": False, "rows": [
                {"mode": mode, "report": file_identity(valid_reports[mode]["path"])}
                for mode in MODES], "schema": contract["schema"]}


def expected_jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    blocks = 1 if lane == "qualification" else plan[lane]["blocks"]
    samples = 1 if lane == "qualification" else plan[lane]["samples"]
    warmup = 0 if lane == "qualification" else plan[lane]["warmup"]
    jobs = []
    for block in range(blocks):
        order = ["before"] if lane == "qualification" else plan["native"]["orders"][block]
        for case in CASES:
            for leg in order:
                jobs.append({"lane": lane, "block": block, **case, "leg": leg,
                             "samples": samples, "warmup": warmup})
    return jobs


def expected_command(plan: dict[str, Any], job: dict[str, Any], binary: dict[str, Any],
                     report: Path, rss: Path) -> list[str]:
    return ["/usr/bin/time", "-f", "%M", "-o", str(rss), "taskset", "-c",
            str(plan["cpu"]), normalize(binary["path"]), "--mode", job["mode"],
            "--shape", job["shape"], "--samples", str(job["samples"]),
            "--warmup", str(job["warmup"]), "--output", str(report)]


def fixture_check(report: dict[str, Any], shape: str, label: str) -> None:
    slides, shapes = DIMENSIONS[shape]
    require(report.get("slides") == slides and report.get("shapes_per_slide") == shapes,
            f"{label} fixture dimensions changed")
    fixture = report.get("fixture")
    require(isinstance(fixture, dict), f"{label} fixture metadata missing")
    if shape == "valid-4attr":
        require(fixture.get("injection") ==
                "valid-four-distinct-namespaced-extension-attributes"
                and fixture.get("slide_parts") == slides
                and fixture.get("replaced_text_tags") == slides * shapes
                and fixture.get("namespace_declarations") == 4
                and fixture.get("namespaced_attributes") == 4,
                f"{label} four-attribute fixture metadata changed")
        uris = fixture.get("namespace_uris")
        names = fixture.get("attribute_names")
        require(uris == VALID_FOUR_ATTRIBUTE_URIS
                and names == VALID_FOUR_ATTRIBUTE_NAMES,
                f"{label} four-attribute namespace oracle changed")
        return
    vendor = shape in {"vendor", "unicode-vendor"}
    injection = {"vendor": "same-length-known-uri-near-misses",
                 "unicode-vendor": "same-length-valid-utf8-unknown-uris"}.get(shape, "none")
    require(fixture.get("injection") == injection
            and fixture.get("slide_parts") == (slides if vendor else 0)
            and fixture.get("replaced_text_tags") == (slides * shapes if vendor else 0)
            and fixture.get("namespace_declarations") == (6 if vendor else 0)
            and fixture.get("namespaced_attributes") == (6 if vendor else 0),
            f"{label} fixture injection changed")
    for key in ("namespace_uris", "attribute_names"):
        values = fixture.get(key)
        require(isinstance(values, list), f"{label} fixture {key} missing")
        if vendor:
            require(len(values) == 6 and all(isinstance(v, str) and v for v in values)
                    and len(set(values)) == 6, f"{label} fixture {key} changed")
        else:
            require(values == [], f"{label} ordinary fixture {key} changed")


def stats(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "empty metric vector")
    for index, value in enumerate(vector):
        finite_number(value, f"metric[{index}]")
    ordered = sorted(vector)

    def nearest(percentile: int) -> int | float:
        return ordered[max(1, math.ceil(percentile * len(ordered) / 100)) - 1]

    return {"count": len(vector), "values": vector, "min": min(vector),
            "p50": nearest(50), "mean": statistics.mean(vector),
            "p95": nearest(95), "p99": nearest(99), "max": max(vector)}


def spread(values: Iterable[int | float]) -> float:
    vector = [float(value) for value in values]
    require(vector and all(math.isfinite(v) for v in vector), "spread vector invalid")
    low, high = min(vector), max(vector)
    return 0.0 if low == high else (float("inf") if low == 0
                                    else (high - low) * 100.0 / abs(low))


def distribution(values: Iterable[int | float]) -> dict[str, Any]:
    result = stats(values)
    result["spread_percent"] = spread(result["values"])
    result["flag_over_5_percent"] = result["spread_percent"] > 5.0
    return result


def allocation_sample(sample: dict[str, Any], label: str) -> dict[str, int]:
    value = sample.get("allocation")
    require(isinstance(value, dict) and value.get("status") == "measured"
            and value.get("scope") == "operation_global_system_allocator",
            f"{label} allocation scope changed")
    result = {}
    for field in RAW_ALLOCATION_FIELDS:
        nonnegative_int(value.get(field), f"{label} allocation {field}")
        result[field] = value[field]
    require(result["live_bytes_after"] == result["live_bytes_before"]
            + result["allocated_bytes"] - result["deallocated_bytes"],
            f"{label} allocation accounting changed")
    require(result["region_peak_live_bytes"] >= result["live_bytes_before"]
            and result["region_peak_live_bytes"] >= result["live_bytes_after"]
            and result["peak_live_bytes_after"] >= result["peak_live_bytes_before"]
            and result["peak_live_bytes_after"] >= result["region_peak_live_bytes"],
            f"{label} allocation peak ordering changed")
    require(result["failed_allocation_calls"] == 0, f"{label} failed allocation observed")
    result["net_live"] = result["live_bytes_after"] - result["live_bytes_before"]
    result["peak_above_entry"] = result["region_peak_live_bytes"] - result["live_bytes_before"]
    return result


def extension_verification(verification: dict[str, Any], report: dict[str, Any],
                           label: str) -> None:
    """Require the complete valid-4attr preservation oracle on every sample.

    The probe checks declarations on each slide root and the exact name/value
    pairs on every text tag.  Reports expose the resulting cardinalities and
    identities; replay binds those fields to the frozen fixture dimensions so
    an aggregate boolean cannot hide a missing or duplicate attribute.
    """

    slides = DIMENSIONS["valid-4attr"][0]
    shapes = DIMENSIONS["valid-4attr"][1]
    text_tags = slides * shapes
    require(verification.get("extension_preservation_check") is True
            and verification.get("extension_text_tags") == text_tags
            and verification.get("extension_attributes_per_text_tag") == 4
            and verification.get("extension_attribute_occurrences") == text_tags * 4
            and verification.get("extension_value_occurrences") == text_tags * 4
            and verification.get("extension_namespace_declarations_per_slide") == 4
            and verification.get("extension_namespace_uris") == VALID_FOUR_ATTRIBUTE_URIS
            and verification.get("extension_attribute_names") == VALID_FOUR_ATTRIBUTE_NAMES
            and verification.get("extension_attribute_values") == VALID_FOUR_ATTRIBUTE_VALUES,
            f"{label} four-attribute preservation oracle changed")


def validate_report(report: dict[str, Any], job: dict[str, Any], kind: str,
                    binary: dict[str, Any], label: str) -> dict[str, Any]:
    require(report.get("schema") == PROBE_FILES["schema"]
            and report.get("tool") == PROBE_FILES["tool"]
            and report.get("marker") == PROBE_FILES["marker"],
            f"{label} probe identity changed")
    require(report.get("mode") == job["mode"] and report.get("shape") == job["shape"]
            and report.get("timing_scope") == TIMING_SCOPES[job["mode"]],
            f"{label} shape/mode/timing identity changed")
    fixture_check(report, job["shape"], label)
    source = report.get("source")
    require(isinstance(source, dict) and is_sha(source.get("sha256")),
            f"{label} source identity missing")
    positive_int(source.get("bytes"), f"{label} source bytes")
    require(report.get("warmup") == job["warmup"]
            and report.get("samples_requested") == job["samples"],
            f"{label} sample configuration changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == job["samples"],
            f"{label} sample count changed")
    allocator = report.get("allocator")
    require(isinstance(allocator, dict) and allocator.get("binary") == Path(binary["path"]).name,
            f"{label} allocator identity changed")
    if kind == "native":
        require(allocator.get("instrumentation") == "none"
                and allocator.get("allocator") == "Rust system allocator"
                and allocator.get("counter_revision") is None
                and all(sample.get("allocation") is None for sample in samples),
                f"{label} native allocation instrumentation changed")
    else:
        require(allocator.get("instrumentation") == "system_allocator_operation_scoped"
                and allocator.get("allocator") == "CountingSystemAllocator(std::alloc::System)"
                and allocator.get("counter_revision") == "serialized_region_peak_v3",
                f"{label} allocation instrumentation changed")
    elapsed, allocations, outputs = [], {field: [] for field in ALLOCATION_FIELDS}, []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{label} sample {index} identity changed")
        elapsed_ns = sample.get("elapsed_ns")
        nonnegative_int(elapsed_ns, f"{label} elapsed_ns")
        elapsed.append(elapsed_ns)
        require(sample.get("source_sha256") == source["sha256"],
                f"{label} sample source changed")
        metrics = sample.get("metrics")
        require(isinstance(metrics, dict) and metrics.get("elapsed_ns") == elapsed_ns
                and metrics.get("slides") == report["slides"]
                and metrics.get("shapes_per_slide") == report["shapes_per_slide"],
                f"{label} raw metric identity changed")
        verification = sample.get("verification")
        require(isinstance(verification, dict) and verification.get("semantic_check") is True
                and verification.get("reopened") is True
                and verification.get("expected_text") == verification.get("actual_text")
                and is_sha(verification.get("semantic_text_sha256"))
                and isinstance(verification.get("expected_text"), str)
                and isinstance(verification.get("actual_text"), str),
                f"{label} semantic verification failed")
        nonnegative_int(verification.get("semantic_text_bytes"), f"{label} semantic bytes")
        if job["shape"] == "valid-4attr":
            extension_verification(verification, report, f"{label} sample {index}")
        else:
            require(all(verification.get(field) is None for field in (
                "extension_preservation_check", "extension_text_tags",
                "extension_attributes_per_text_tag", "extension_attribute_occurrences",
                "extension_value_occurrences",
                "extension_namespace_declarations_per_slide", "extension_namespace_uris",
                "extension_attribute_names", "extension_attribute_values")),
                    f"{label} unexpected extension oracle on ordinary fixture")
        output = sample.get("output")
        require(isinstance(output, dict) and is_sha(output.get("sha256")),
                f"{label} output identity missing")
        nonnegative_int(output.get("bytes"), f"{label} output bytes")
        require(verification.get("readback_bytes") == output["bytes"]
                and verification.get("readback_sha256") == output["sha256"],
                f"{label} readback identity changed")
        expected_marker = job["mode"] in {"commit", "lifecycle"}
        require(verification.get("marker_matches") is (expected_marker or None),
                f"{label} marker verification changed")
        vendor = job["shape"] in {"vendor", "unicode-vendor"}
        require(verification.get("unknown_namespace_check") is (True if vendor else None)
                and verification.get("unknown_namespace_occurrences")
                == (DIMENSIONS[job["shape"]][0] * DIMENSIONS[job["shape"]][1] if vendor else None),
                f"{label} namespace oracle changed")
        outputs.append((output["bytes"], output["sha256"]))
        if kind == "allocation":
            values = allocation_sample(sample, f"{label} sample {index}")
            for field, number in values.items():
                allocations[field].append(number)
        else:
            require(sample.get("allocation") is None, f"{label} native allocation present")
    require(len(set(outputs)) == 1, f"{label} output identity is not deterministic")
    return {"stats": stats(elapsed), "elapsed": elapsed,
            "allocation": None if kind == "native" else allocations,
            "outputs": outputs}


def load_lane(plan: dict[str, Any], lane: str, builds: dict[str, Any],
              expected_source: dict[str, Any], cleanup: Any, cleanup_ok: bool) -> list[dict[str, Any]]:
    directory = PACKET / lane
    complete = read_json(directory / "complete.json")
    jobs = expected_jobs(plan, lane)
    require(complete.get("children") == len(jobs), f"{lane} child count changed")
    source_path = artifact_path(complete.get("source"), f"{lane} complete source")
    require(source_files_equal(source_manifest(read_json(source_path), f"{lane} source"),
                               expected_source), f"{lane} complete source changed")
    receipts_path = artifact_path(complete.get("receipts"), f"{lane} receipts")
    rows = read_json(receipts_path)
    require(isinstance(rows, list) and len(rows) == len(jobs), f"{lane} receipt count changed")
    entries, seen = [], set()
    source_by_shape: dict[str, str] = {}
    for index, (row, job) in enumerate(zip(rows, jobs)):
        label = f"{lane} child {index}"
        require(isinstance(row, dict) and all(row.get(key) == job[key]
                for key in ("lane", "block", "shape", "mode", "leg")),
                f"{label} identity changed")
        identity = (job["block"], job["shape"], job["mode"], job["leg"])
        require(identity not in seen, f"{label} duplicate identity")
        seen.add(identity)
        require(row.get("exit_code") == 0, f"{label} failed")
        started, ended = row.get("started"), row.get("ended")
        finite_number(started, f"{label} started")
        finite_number(ended, f"{label} ended")
        require(started <= ended, f"{label} receipt interval changed")
        kind = "allocation" if lane in {"allocation", "qualification"} else "native"
        expected_binary = builds[job["leg"]]["binaries"][kind]
        binary = row.get("binary")
        require(isinstance(binary, dict) and binary.get("bytes") == expected_binary.get("bytes")
                and binary.get("sha256") == expected_binary.get("sha256"),
                f"{label} binary identity changed")
        validate_binary(binary, f"{label} binary", cleanup, cleanup_ok)
        report_path = artifact_path(row.get("report"), f"{label} report")
        log_path = artifact_path(row.get("log"), f"{label} log")
        rss_path = artifact_path(row.get("rss"), f"{label} RSS")
        rss_text = rss_path.read_text().strip()
        require(rss_text.isdigit(), f"{label} RSS is not an integer")
        rss = int(rss_text)
        command = row.get("command")
        require(isinstance(command, list)
                and [normalize(item) for item in command]
                == expected_command(plan, job, expected_binary, report_path, rss_path),
                f"{label} command changed")
        report = read_json(report_path)
        outcome = validate_report(report, job, kind, expected_binary, label)
        digest = report["source"]["sha256"]
        require(source_by_shape.setdefault(job["shape"], digest) == digest,
                f"{label} fixture source changed")
        entries.append({"identity": job, "report_path": report_path,
                        "report_sha256": sha256(report_path), "log": rel(log_path),
                        "started": float(started), "ended": float(ended),
                        "rss_kib": rss, "source_sha256": digest,
                        "stats": outcome["stats"], "elapsed": outcome["elapsed"],
                        "allocation": outcome["allocation"], "outputs": outcome["outputs"]})
    return entries


def prior_fixture_parity(entries: list[dict[str, Any]], historical: dict[str, Any]) -> dict[str, Any]:
    prior = {row["case"]: row for row in historical["reports"]}
    current = {}
    for item in entries:
        identity = item["identity"]
        if identity["lane"] != "qualification":
            continue
        key = f"{identity['shape']}/{identity['mode']}"
        if identity["shape"] == "valid-4attr":
            continue
        report = read_json(item["report_path"])
        sample = report["samples"][0]
        source, output = report["source"], sample["output"]
        require({"bytes": source["bytes"], "sha256": source["sha256"]} == prior[key]["report_source"]
                and {"bytes": output["bytes"], "sha256": output["sha256"]} == prior[key]["output"],
                f"0792 fixture parity changed: {key}")
        current[key] = {"source": prior[key]["report_source"],
                        "output": prior[key]["output"]}
    require(set(current) == {f"{shape}/{mode}" for shape in ORIGINAL_SHAPES for mode in MODES},
            "0792 fixture parity coverage changed")
    return {"historical": HISTORICAL_PACKET, "rows": current, "timings_imported": False}


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(values, "cannot bootstrap empty ratio vector")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        estimates.append(statistics.median(values[rng.randrange(len(values))] for _ in values))
    estimates.sort()
    return {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
            "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median",
            "ci_low": estimates[BOOTSTRAP_LOW_RANK],
            "ci_high": estimates[BOOTSTRAP_HIGH_RANK],
            "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK}


def pair_ratio(before: float, after: float) -> dict[str, Any]:
    if before == 0:
        equal = after == 0
        return {"before": before, "after": after, "ratio": 1.0 if equal else None,
                "change_percent": 0.0 if equal else None, "relative_change_defined": False,
                "zero_baseline_equal": equal, "zero_to_nonzero": not equal,
                "over_5_percent": not equal}
    ratio = after / before
    return {"before": before, "after": after, "ratio": ratio,
            "change_percent": (ratio - 1.0) * 100.0, "relative_change_defined": True,
            "zero_baseline_equal": False, "zero_to_nonzero": False,
            "over_5_percent": ratio > 1.05}


def paired(entries: list[dict[str, Any]], metrics: Iterable[str]) -> dict[str, Any]:
    by_key = {(x["identity"]["shape"], x["identity"]["mode"],
               x["identity"]["leg"], x["identity"]["block"]): x for x in entries}
    output = {}
    for shape, mode in sorted({(x["identity"]["shape"], x["identity"]["mode"]) for x in entries}):
        blocks = sorted({x["identity"]["block"] for x in entries
                         if x["identity"]["shape"] == shape and x["identity"]["mode"] == mode})
        values = {}
        for metric in metrics:
            rows, ratios = [], []
            for block in blocks:
                before = by_key[(shape, mode, "before", block)]
                after = by_key[(shape, mode, "after", block)]
                if metric == "rss_kib":
                    left, right = before["rss_kib"], after["rss_kib"]
                elif before["allocation"] is not None:
                    left, right = stats(before["allocation"][metric])["p50"], stats(after["allocation"][metric])["p50"]
                else:
                    left, right = before["stats"][metric], after["stats"][metric]
                row = pair_ratio(float(left), float(right)); row["block"] = block
                rows.append(row)
                if row["ratio"] is not None:
                    ratios.append(row["ratio"])
            median = statistics.median(ratios) if ratios else None
            values[metric] = {"by_block": rows, "ratio_median": median,
                              "change_percent_median": None if median is None else (median - 1) * 100,
                              "bootstrap": bootstrap(ratios) if ratios else {"ci_low": None, "ci_high": None,
                                "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                                "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median",
                                "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK},
                              "defined_ratio_blocks": len(ratios),
                              "undefined_ratio_blocks": len(rows) - len(ratios),
                              "regression_over_5_percent": any(r["over_5_percent"] for r in rows)}
        output[f"{shape}/{mode}"] = {"shape": shape, "mode": mode, "blocks": len(blocks),
            "metrics": values, "comparison": "after/before paired by alternating capture block"}
    return output


def native_analysis(entries: list[dict[str, Any]]) -> dict[str, Any]:
    groups, spread_flags = {}, []
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for item in entries:
        identity = item["identity"]
        grouped.setdefault((identity["shape"], identity["mode"], identity["leg"]), []).append(item)
    for key, items in sorted(grouped.items()):
        elapsed = {metric: distribution(item["stats"][metric] for item in items)
                   for metric in NATIVE_METRICS}
        rss = distribution(item["rss_kib"] for item in items)
        for metric, result in (*elapsed.items(), ("rss_kib", rss)):
            if result["flag_over_5_percent"]:
                spread_flags.append({"group": list(key), "metric": metric,
                                     "spread_percent": result["spread_percent"]})
        groups["/".join(key)] = {"shape": key[0], "mode": key[1], "leg": key[2],
            "processes": [{"block": item["identity"]["block"], "stats": item["stats"],
                            "elapsed": item["elapsed"], "rss_kib": item["rss_kib"],
                            "report": rel(item["report_path"]),
                            "report_sha256": item["report_sha256"]}
                           for item in sorted(items, key=lambda x: x["identity"]["block"])],
            "elapsed_distribution_across_processes": elapsed,
            "per_process_rss_distribution": rss}
    paired_values = paired(entries, (*NATIVE_METRICS, "rss_kib"))
    regressions = []
    for key, group in paired_values.items():
        for metric, result in group["metrics"].items():
            if result["regression_over_5_percent"]:
                regressions.append({"group": key, "metric": metric,
                    "ratio_median": result["ratio_median"],
                    "change_percent_median": result["change_percent_median"]})
    return {"groups": groups, "spread_flags_over_5_percent": spread_flags,
            "regression_flags_over_5_percent": regressions,
            "paired_by_block_before_after": paired_values,
            "rss_is_separate_from_elapsed": True, "allocation_metrics_present": False}


def allocation_analysis(entries: list[dict[str, Any]]) -> dict[str, Any]:
    groups, spread_flags = {}, []
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for item in entries:
        identity = item["identity"]
        grouped.setdefault((identity["shape"], identity["mode"], identity["leg"]), []).append(item)
    for key, items in sorted(grouped.items()):
        metrics = {}
        for field in ALLOCATION_FIELDS:
            per_process = [{"block": item["identity"]["block"], "values": item["allocation"][field],
                            "stats": stats(item["allocation"][field])}
                           for item in sorted(items, key=lambda x: x["identity"]["block"])]
            repeats = [item["stats"]["p50"] for item in per_process]
            spread_value = spread(repeats)
            metrics[field] = {"per_process": per_process, "repeat_p50_values": repeats,
                              "spread_percent": spread_value,
                              "flag_over_5_percent": spread_value > 5}
            if spread_value > 5:
                spread_flags.append({"group": list(key), "metric": field,
                                     "spread_percent": spread_value})
        groups["/".join(key)] = {"shape": key[0], "mode": key[1], "leg": key[2],
                                  "blocks": len(items), "metrics": metrics,
                                  "elapsed_not_mixed": True}
    paired_values = paired(entries, ALLOCATION_FIELDS)
    regressions = []
    for key, group in paired_values.items():
        for metric, result in group["metrics"].items():
            if result["regression_over_5_percent"]:
                regressions.append({"group": key, "metric": metric,
                    "ratio_median": result["ratio_median"],
                    "change_percent_median": result["change_percent_median"]})
    return {"groups": groups, "fields": list(ALLOCATION_FIELDS),
            "spread_flags_over_5_percent": spread_flags,
            "regression_flags_over_5_percent": regressions,
            "paired_by_block_before_after": paired_values,
            "allocation_is_separate_from_elapsed": True}


def decision_guards(native: dict[str, Any], allocation: dict[str, Any],
                    policy: dict[str, Any]) -> dict[str, Any]:
    value = policy["policy"]
    latency, benefit = value["latency"], value["benefit"]
    violations, benefits, resources = [], [], []
    for key, group in sorted(native["paired_by_block_before_after"].items()):
        metric = group["metrics"]["p50"]
        if metric["ratio_median"] is not None and metric["ratio_median"] > latency["maximum_ratio"] \
                and metric["bootstrap"]["ci_low"] > latency["bootstrap95_low_must_exceed"]:
            violations.append({"case": key, "ratio_median": metric["ratio_median"],
                "bootstrap_ci_low": metric["bootstrap"]["ci_low"],
                "bootstrap_ci_high": metric["bootstrap"]["ci_high"],
                "change_percent_median": metric["change_percent_median"]})
        shape, mode = key.split("/", 1)
        if mode in benefit["eligible_modes"] and metric["ratio_median"] is not None \
                and metric["bootstrap"]["ci_high"] < benefit["bootstrap95_high_below"] \
                and (1 - metric["ratio_median"]) * 100 >= benefit["minimum_improvement_percent"]:
            benefits.append({"case": key, "improvement_percent": (1 - metric["ratio_median"]) * 100,
                "ratio_median": metric["ratio_median"], "bootstrap_ci_low": metric["bootstrap"]["ci_low"],
                "bootstrap_ci_high": metric["bootstrap"]["ci_high"]})
    for key, group in sorted(allocation["paired_by_block_before_after"].items()):
        for metric_name in ("allocation_calls", "allocated_bytes", "net_live", "peak_above_entry"):
            for row in group["metrics"][metric_name]["by_block"]:
                if row["after"] > row["before"]:
                    resources.append({"case": key, "metric": metric_name, "block": row["block"],
                                      "before": row["before"], "after": row["after"]})
    return {"latency_violations": violations, "resource_violations": resources,
            "eligible_benefits": benefits, "benefit_satisfied": bool(benefits),
            "latency_guard_passed": not violations, "resource_guard_passed": not resources,
            "adoption_eligible": not violations and not resources and bool(benefits),
            "policy_thresholds": {"maximum_ratio": latency["maximum_ratio"],
                "bootstrap95_low_must_exceed": latency["bootstrap95_low_must_exceed"],
                "minimum_improvement_percent": benefit["minimum_improvement_percent"]}}


def disposition(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "disposition.json"
    if not path.is_file():
        require(current_source_files() == after["files"],
                "undecided packet requires live candidate source")
        return {"status": "pending", "production_change_retained": None,
                "live_source_matches": "after"}
    value = read_json(path)
    require(isinstance(value, dict) and value.get("status") in {"retained", "rejected"},
            "candidate disposition is malformed")
    status = value["status"]
    require(value.get("production_change_retained") is (status == "retained"),
            "candidate disposition flag changed")
    if status == "retained":
        require(current_source_files() == after["files"], "retained source differs from after")
        return {"status": status, "production_change_retained": True,
                "live_source_matches": "after"}
    expected_restored = dict(before["files"])
    require(current_source_files() == expected_restored,
            "rejected source differs from approved restoration")
    restored = artifact_path(value.get("restored_source"), "restored source census")
    restored_manifest = source_manifest(read_json(restored), "restored source census")
    require(restored_manifest["files"] == expected_restored,
            "restored source differs from approved restoration")
    return {"status": status, "production_change_retained": False,
            "live_source_matches": "before",
            "restored_source": file_identity(restored)}


def analyze() -> dict[str, Any]:
    plan = load_plan()
    premeasurement = check_premeasurement_inputs()
    policy = load_policy()
    builds, cleanup, cleanup_ok = load_builds(plan)
    before, after = builds["before"]["source"], builds["after"]["source"]
    original_application = read_json(PACKET / "application.json")
    candidate_source = source_manifest(original_application.get("source"),
                                      "original candidate application source")
    candidate_application = check_candidate_application(before, candidate_source)
    quality_application = read_json(PACKET / "quality-amendment-application.json")
    quality_source = source_manifest(quality_application.get("source"),
                                     "quality amendment application source")
    quality_amendment = check_quality_amendment(before, candidate_source, quality_source)
    visibility_amendment = check_visibility_amendment(quality_source, after)
    architecture = load_architecture_inputs()
    inheritance = load_inheritance()
    require(inheritance["production_source_files"] == before["files"],
            "inherited production source differs from 0806 baseline")
    historical = load_prior_qualification()
    quality = check_quality(after)
    probe_tests = check_probe_tests(builds)
    build_warnings = check_build_warning_scope(builds)
    native_entries = load_lane(plan, "native", builds, after, cleanup, cleanup_ok)
    allocation_entries = load_lane(plan, "allocation", builds, after, cleanup, cleanup_ok)
    qualification_entries = load_lane(plan, "qualification", builds, before, cleanup, cleanup_ok)
    require(len(native_entries) == 216 and len(allocation_entries) == 72
            and len(qualification_entries) == 18, "lane cardinality changed")
    all_entries = (*native_entries, *allocation_entries, *qualification_entries)
    require(sum(item["stats"]["count"] for item in all_entries) == 6714,
            "sample cardinality changed")
    fixture_parity = prior_fixture_parity(qualification_entries, historical)
    new_qualification = before_qualification_contract(qualification_entries, builds)
    after_build_inputs = check_after_build_inputs(new_qualification, builds,
                                                  qualification_entries)
    native = native_analysis(native_entries)
    allocation = allocation_analysis(allocation_entries)
    guards = decision_guards(native, allocation, policy)
    binary_ids = {leg: {kind: {"bytes": builds[leg]["binaries"][kind]["bytes"],
                               "sha256": builds[leg]["binaries"][kind]["sha256"]}
                        for kind in ("native", "allocation", "profile")} for leg in LEGS}
    return {
        "schema": "litchi.performance.0806.workflow-analysis.v1",
        "plan_schema": plan["schema"],
        "counts": {"reports": 306, "samples": 6714, "native_reports": 216,
                    "allocation_reports": 72, "qualification_reports": 18},
        "source": {"before": before, "after": after,
                    "quality_amendment": quality_source,
                    "changed_files": list(visibility_amendment["changed_files"])},
        "premeasurement": premeasurement,
        "candidate_application": candidate_application,
        "quality_amendment_application": quality_amendment,
        "visibility_amendment_application": visibility_amendment,
        "policy": policy,
        "quality": quality,
        "probe_tests": probe_tests,
        "build_warnings": build_warnings,
        "architecture_inputs": architecture,
        "inheritance": inheritance,
        "historical_qualification": historical,
        "fixture_parity": fixture_parity,
        "new_qualification_oracle": new_qualification,
        "after_build_inputs": after_build_inputs,
        "decision_guards": guards,
        "disposition": disposition(before, after),
        "binary_identities": binary_ids,
        "cleanup_contract": {"schema": "litchi.performance.0806.cleanup.v1",
                              "target": origin()["target"], "binary_count": 6,
                              "exact_witness_required": True},
        "native": {"children": 216, "blocks": 6, "samples": 30, "warmup": 3,
                    "analysis": native,
                    "receipts": [{"shape": x["identity"]["shape"], "mode": x["identity"]["mode"],
                                  "block": x["identity"]["block"], "leg": x["identity"]["leg"],
                                  "report": rel(x["report_path"]), "report_sha256": x["report_sha256"]}
                                 for x in native_entries]},
        "allocation": {"children": 72, "blocks": 2, "samples": 3, "warmup": 0,
                        "analysis": allocation,
                        "receipts": [{"shape": x["identity"]["shape"], "mode": x["identity"]["mode"],
                                      "block": x["identity"]["block"], "leg": x["identity"]["leg"],
                                      "report": rel(x["report_path"]), "report_sha256": x["report_sha256"]}
                                     for x in allocation_entries]},
        "qualification": {"children": 18, "before_source_only": True,
                           "rows": [{"shape": x["identity"]["shape"], "mode": x["identity"]["mode"],
                                     "report": rel(x["report_path"]), "report_sha256": x["report_sha256"]}
                                    for x in qualification_entries]},
        "bootstrap": {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                       "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median",
                       "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK},
        "verification": {
            "all_raw_report_samples_checked": True,
            "semantic_verification_checked": True,
            "source_binary_probe_lock_receipts_checked": True,
            "frozen_inputs_checked": True,
            "architecture_inputs_checked": True,
            "historical_qualification_checked_without_timings": True,
            "new_four_attribute_oracle_checked_without_timings": True,
            "after_build_inputs_checked": True,
            "quality_amendment_application_checked": True,
            "visibility_amendment_application_checked": True,
            "amendment_preflight_checked": True,
            "inherited_probe_checked_without_timings": True,
            "quality_commands_checked_exactly": True,
            "probe_tests_checked": True,
            "native_has_no_allocation_metrics": True,
            "allocation_memory_guards_use_block_medians": True,
            "allocation_calls_and_bytes_guards_checked": True,
            "cleanup_binary_witness_required": True,
            "source_change_allowlist": list(visibility_amendment["changed_files"]),
        },
        "limits": [
            "Timing and allocation comparisons are descriptive paired evidence.",
            "RSS is the whole-process /usr/bin/time maximum resident-set gauge.",
            "Allocator counters are scoped operation counters, not physical-copy or causal proof.",
            "No cold-cache, device-floor, concurrency, historical timing, or profile timing claim is made.",
        ],
    }


def main() -> None:
    args = set(sys.argv[1:])
    if args == {"--freeze-qualification"}:
        print(json.dumps(freeze_qualification_contract(), sort_keys=True))
        return
    require(args <= {"--write", "--check"} and args,
            "usage: analyze.py --freeze-qualification, --write or analyze.py --check")
    result = analyze()
    output = PACKET / "analysis.json"
    encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
    if "--write" in args:
        output.write_text(encoded)
    if "--check" in args:
        require(output.is_file() and read_json(output) == result,
                "analysis.json does not replay byte-for-byte")
    print(json.dumps({"reports": result["counts"]["reports"],
                      "samples": result["counts"]["samples"]}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except ReplayError as error:
        print(f"analysis failed: {error}", file=sys.stderr)
        raise SystemExit(1)
