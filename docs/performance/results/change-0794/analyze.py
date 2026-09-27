"""Fail-closed offline replay for the 0794 shared XML-attribute packet.

This module only consumes retained receipts and report JSON.  It never runs a
probe, compiler, profiler, or timing command.  The six binary receipts may be
verified either while the owned target exists or after cleanup has recorded an
exact hash witness.
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
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
ORIGIN_PATH = PACKET / "origin.json"
SOURCE_ALLOWLIST = (
    "crates/litchi-formula/src/omml/xml_attributes.rs",
    "crates/litchi-ole-common/src/xml_attributes.rs",
    "crates/litchi-opc/src/xml_attributes.rs",
    "crates/litchi-sign/src/xml_attributes.rs",
    "crates/litchi-xldm/src/xml_attributes.rs",
    "crates/xml-minifier/src/xml_attributes.rs",
)
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor")
MODES = ("capture", "commit", "lifecycle")
CASES = tuple({"shape": shape, "mode": mode} for shape in SHAPES for mode in MODES)
LEGS = ("before", "after")
DIMENSIONS = {
    "tiny": (3, 4),
    "medium": (12, 8),
    "large": (100, 100),
    "vendor": (12, 8),
    "unicode-vendor": (12, 8),
}
PROBE_FILES = {
    "schema": "litchi.pptx.namespace-uri-probe.v1",
    "tool": "namespace-uri-probe-0785",
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
BOOTSTRAP_SEED = 794079
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9749
BOOTSTRAP_CONFIDENCE = 0.95
HISTORICAL_PACKET = "docs/performance/results/change-0792"
PROFILE_PACKET = "docs/performance/results/change-0793"
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
    require(isinstance(value.get("main"), str) and value["main"],
            "origin main path is missing")
    require(isinstance(value.get("target"), str) and value["target"],
            "origin target path is missing")
    return value


def _candidate_paths(raw: Path, *, packet_bound: bool) -> list[Path]:
    candidates: list[Path] = []
    owned = Path(origin().get("worktree", origin().get("main", ROOT))).resolve()
    if raw.is_absolute():
        try:
            candidates.append(ROOT / raw.resolve().relative_to(owned))
        except ValueError:
            pass
        parts = raw.parts
        for marker in ("change-0794", "change-0793", "change-0792", "build-before", "build-after", "native",
                       "allocation", "qualification", "probe-tests"):
            if marker in parts:
                index = parts.index(marker)
                if marker in {"change-0792", "change-0793", "change-0794"}:
                    candidates.append(PACKET.joinpath(*parts[index + 1:]))
                else:
                    candidates.append(PACKET / marker / Path(*parts[index + 1:]))
        candidates.append(raw)
    else:
        text = str(raw).replace("\\", "/")
        for packet_name in ("change-0794", "change-0793", "change-0792"):
            prefix = f"docs/performance/results/{packet_name}/"
            if text.startswith(prefix):
                candidates.append(PACKET / text[len(prefix):] if packet_name == "change-0794"
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
    require(isinstance(plan, dict) and plan.get("schema") == "litchi.performance.0794.v1",
            "plan schema changed")
    require(plan.get("cpu") == 12 and plan.get("cases") == list(CASES),
            "plan cases or CPU changed")
    require(plan.get("source_allowlist") == list(SOURCE_ALLOWLIST),
            "source allowlist changed")
    require(isinstance(plan.get("scope"), str) and plan["scope"], "plan scope missing")
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
    return {"receipt": None if value is None else file_identity(path),
            "contract": file_identity(PACKET / "analysis-plan.json"),
            "bootstrap_endpoints": [BOOTSTRAP_LOW_RANK, BOOTSTRAP_HIGH_RANK]}


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
        "All15cases includingvendorfallbackcontrols; diagnostichelpergains aloneinsufficient."),
            "policy scope changed")
    return {"receipt": file_identity(path), "policy": value}


def check_candidate_application(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    """Bind the measured source census to the five-file candidate archive.

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
    require(sha256(patch_witness) == manifest.get("patch", {}).get("sha256"),
            "candidate manifest patch changed")
    require(sorted(row["production"] for row in manifest.get("files", [])) == changed,
            "candidate manifest source set changed")
    candidate = PACKET / "candidate"
    archive = {"changed_files": changed}
    if candidate.exists():
        require(candidate.is_dir(), "candidate archive is not a directory")
        patch_candidates = [candidate / "candidate.patch", PACKET / "candidate.patch"]
        review_candidates = [candidate / "candidate-review.md", candidate / "source-review.md",
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
            flat = root / (Path(name).parts[1] + "-" + Path(name).name)
            if flat.is_file():
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
    require(before["files"].get(SOURCE_ALLOWLIST[0]) == after["files"].get(SOURCE_ALLOWLIST[0]),
            "lenient formula helper changed outside the candidate")
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
            require(warning_lines and all("probe-src" in line or line.startswith("warning: `mce-capabilities-probe`")
                                         or line.startswith("warning: function")
                                         or line.startswith("warning: struct")
                                         or line.startswith("warning: field")
                                         or line.startswith("warning: methods")
                                         or line.startswith("warning: associated")
                                         or line.startswith("warning: multiple")
                                         for line in warning_lines),
                    f"{leg} {kind} build warning scope changed")
            generated = [line for line in text.splitlines()
                         if "mce-capabilities-probe" in line and "generated" in line]
            require(len(generated) == 1, f"{leg} {kind} warning summary is missing")
            rows.append({"leg": leg, "kind": kind, "warning_lines": len(warning_lines),
                         "summary": generated[0].strip()})
    return {"rows": rows, "scope": "standalone probe warnings retained; no build errors"}


def check_probe_tests(before: dict[str, Any]) -> dict[str, Any]:
    """Use the sealed 0793 probe-test result; this packet does not rerun it."""

    packet = ROOT / PROFILE_PACKET
    seal = prior_seal(packet, "0793 inherited probe tests")
    build = read_json(packet / "build/receipt.json")
    prior = read_json(packet / "analysis.json")
    prior_probe = build.get("probe")
    require(isinstance(prior_probe, dict)
            and {name: prior_probe.get(name) for name in before["probe"]} == before["probe"],
            "inherited probe source inventory changed")
    require(prior.get("fresh_probe_tests") == 35,
            "inherited probe test count changed")
    require("analysis.json" in seal.get("files", {}),
            "inherited probe test analysis is not sealed")
    return {"inherited_packet": "docs/performance/results/change-0793",
            "inherited_seal_schema": seal.get("schema"),
            "commands": [], "rows": 0, "tests": [28, 7], "filtered": [0, 26],
            "fresh_tests": 35, "warnings_retained": True,
            "rerun": False}


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
    prior = ROOT / PROFILE_PACKET
    fixture_seal_path = fixture / "seal.json"
    prior_seal_path = prior / "seal.json"
    require(value.get("fixture_packet") == "../change-0792"
            and is_sha(value.get("fixture_seal"))
            and fixture_seal_path.is_file()
            and sha256(fixture_seal_path) == value["fixture_seal"],
            "fixture inheritance reference changed")
    require(value.get("prior_packet") == "../change-0793"
            and is_sha(value.get("prior_seal"))
            and prior_seal_path.is_file()
            and sha256(prior_seal_path) == value["prior_seal"],
            "prior packet inheritance reference changed")
    fixture_value = prior_seal(fixture, "0792 fixture")
    prior_value = prior_seal(prior, "0793 prior packet")

    production = value.get("production_source")
    production_path = artifact_path(production, "inherited production source", packet_bound=False)
    production_manifest = source_manifest(read_json(production_path),
                                          "inherited production source")
    probe = value.get("probe_reference")
    require(isinstance(probe, dict) and set(probe) == {
        "Cargo.lock", "Cargo.toml.template", "src/allocation_metrics.rs",
        "src/counting_allocator.rs", "src/main.rs"},
            "probe inheritance references changed")
    old_probe = prior / "probe-src"
    for name, digest in probe.items():
        expected = old_probe / name
        current = PACKET / "probe-src" / name
        require(is_sha(digest) and expected.is_file() and current.is_file()
                and sha256(expected) == digest and sha256(current) == digest,
                f"probe inheritance changed: {name}")
    return {"fixture_packet": "docs/performance/results/change-0792",
            "fixture_seal": value["fixture_seal"],
            "fixture_sealed_files": len(fixture_value["files"]),
            "prior_packet": PROFILE_PACKET,
            "prior_seal": value["prior_seal"],
            "prior_sealed_files": len(prior_value["files"]),
            "production_source": file_identity(production_path),
            "production_source_files": production_manifest["files"],
            "probe_references": sorted(probe), "timings_imported": False}


def load_prior_qualification() -> dict[str, Any]:
    packet = ROOT / HISTORICAL_PACKET
    seal = prior_seal(packet, "0792 fixture qualification")
    source_path = packet / "build-after/source.json"
    require(seal["files"].get("build-after/source.json") == sha256(source_path),
            "0792 after source is not sealed")
    source = source_manifest(read_json(source_path), "0792 after source")
    rows = []
    for case in CASES:
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
        report = read_json(item["report_path"])
        sample = report["samples"][0]
        source, output = report["source"], sample["output"]
        require({"bytes": source["bytes"], "sha256": source["sha256"]} == prior[key]["report_source"]
                and {"bytes": output["bytes"], "sha256": output["sha256"]} == prior[key]["output"],
                f"0792 fixture parity changed: {key}")
        current[key] = {"source": prior[key]["report_source"],
                        "output": prior[key]["output"]}
    require(set(current) == {f"{shape}/{mode}" for shape in SHAPES for mode in MODES},
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
    application = check_candidate_application(before, after)
    architecture = load_architecture_inputs()
    inheritance = load_inheritance()
    require(inheritance["production_source_files"] == before["files"],
            "inherited production source differs from 0794 baseline")
    historical = load_prior_qualification()
    quality = check_quality(after)
    probe_tests = check_probe_tests(builds["before"])
    build_warnings = check_build_warning_scope(builds)
    native_entries = load_lane(plan, "native", builds, after, cleanup, cleanup_ok)
    allocation_entries = load_lane(plan, "allocation", builds, after, cleanup, cleanup_ok)
    qualification_entries = load_lane(plan, "qualification", builds, before, cleanup, cleanup_ok)
    require(len(native_entries) == 180 and len(allocation_entries) == 60
            and len(qualification_entries) == 15, "lane cardinality changed")
    all_entries = (*native_entries, *allocation_entries, *qualification_entries)
    require(sum(item["stats"]["count"] for item in all_entries) == 5595,
            "sample cardinality changed")
    fixture_parity = prior_fixture_parity(qualification_entries, historical)
    native = native_analysis(native_entries)
    allocation = allocation_analysis(allocation_entries)
    guards = decision_guards(native, allocation, policy)
    binary_ids = {leg: {kind: {"bytes": builds[leg]["binaries"][kind]["bytes"],
                               "sha256": builds[leg]["binaries"][kind]["sha256"]}
                        for kind in ("native", "allocation", "profile")} for leg in LEGS}
    return {
        "schema": "litchi-0794-shared-xml-attribute-analysis-v1",
        "plan_schema": plan["schema"],
        "counts": {"reports": 255, "samples": 5595, "native_reports": 180,
                    "allocation_reports": 60, "qualification_reports": 15},
        "source": {"before": before, "after": after,
                    "changed_files": list(application["changed_files"])},
        "premeasurement": premeasurement,
        "candidate_application": application,
        "policy": policy,
        "quality": quality,
        "probe_tests": probe_tests,
        "build_warnings": build_warnings,
        "architecture_inputs": architecture,
        "inheritance": inheritance,
        "historical_qualification": historical,
        "fixture_parity": fixture_parity,
        "decision_guards": guards,
        "disposition": disposition(before, after),
        "binary_identities": binary_ids,
        "cleanup_contract": {"schema": "litchi.performance.0794.cleanup.v1",
                              "target": origin()["target"], "binary_count": 6,
                              "exact_witness_required": True},
        "native": {"children": 180, "blocks": 6, "samples": 30, "warmup": 3,
                    "analysis": native,
                    "receipts": [{"shape": x["identity"]["shape"], "mode": x["identity"]["mode"],
                                  "block": x["identity"]["block"], "leg": x["identity"]["leg"],
                                  "report": rel(x["report_path"]), "report_sha256": x["report_sha256"]}
                                 for x in native_entries]},
        "allocation": {"children": 60, "blocks": 2, "samples": 3, "warmup": 0,
                        "analysis": allocation,
                        "receipts": [{"shape": x["identity"]["shape"], "mode": x["identity"]["mode"],
                                      "block": x["identity"]["block"], "leg": x["identity"]["leg"],
                                      "report": rel(x["report_path"]), "report_sha256": x["report_sha256"]}
                                     for x in allocation_entries]},
        "qualification": {"children": 15, "before_source_only": True,
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
            "inherited_probe_checked_without_timings": True,
            "quality_commands_checked_exactly": True,
            "probe_tests_checked": True,
            "native_has_no_allocation_metrics": True,
            "allocation_memory_guards_use_block_medians": True,
            "allocation_calls_and_bytes_guards_checked": True,
            "cleanup_binary_witness_required": True,
            "source_change_allowlist": list(application["changed_files"]),
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
    require(args <= {"--write", "--check"} and args,
            "usage: analyze.py --write or analyze.py --check")
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
