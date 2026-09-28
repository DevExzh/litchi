"""Small offline-only custody and fixture helpers for change 0807.

The capture/build drivers own execution.  The analysis modules only replay
their retained receipts and raw outputs.  This module deliberately keeps
binary custody separate from packet-artifact custody so a final cleanup can
remove the owned target without making retained reports unverifiable.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path
from types import ModuleType
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
TARGET = Path("/home/zhuhe/code/litchi-target-0807")
HEX64 = re.compile(r"[0-9a-f]{64}")
SHAPES = ("tiny", "medium", "large")
DIMENSIONS = {"tiny": (3, 4), "medium": (12, 8), "large": (100, 100)}
MARKER = "litchi-perf-0780-static-mce-capabilities"
SCHEMA = "litchi.pptx.capture-profile-probe.v1"
TOOL = "pptx-capture-probe-0784"
OWNER = "pptx_capture_probe::capture_region_0784"
CHILD = "litchi_pptx::package::model::Package::opened_presentation_with_limits"
NESTED = "litchi_pptx::opened::model::capture_internal"
SCAN = "litchi_pptx::notes::codec::scan_processed_xml"
FINGERPRINT = "litchi_pptx::opened::model::package_fingerprint_with_memo"
RESOLVED = "litchi_pptx::notes::resolved"
SHA = "sha2::sha256::x86_sha::compress"
LOAD_INDEX = "litchi_pptx::notes::package::load_index_with_slide_root_proofs"
BOOTSTRAP_SEED = 807080
BOOTSTRAP_RESAMPLES = 10000


class EvidenceError(ValueError):
    """A missing, stale, or contradictory retained artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"invalid JSON {path}: {error}") from error


def write_or_check(path: Path, value: Any, check: bool) -> None:
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    if path.exists():
        require(path.is_file() and not path.is_symlink(), f"output is not regular: {path}")
        require(path.read_text(encoding="utf-8") == encoded,
                f"replayed output differs: {path.name}")
    else:
        require(not check, f"missing expected output: {path.name}")
        path.write_text(encoded, encoding="utf-8")


def _packet_candidate(raw: str) -> Path:
    marker = "/change-0807/"
    if marker in raw:
        return (PACKET / raw.split(marker, 1)[1]).resolve()
    path = Path(raw)
    if path.is_absolute():
        return path.resolve()
    return (PACKET / path).resolve()


def packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: artifact path is missing")
    path = _packet_candidate(raw)
    try:
        path.relative_to(PACKET.resolve())
    except ValueError as error:
        raise EvidenceError(f"{label}: path escapes packet: {raw}") from error
    return path


def external_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: path is missing")
    path = Path(raw)
    require(path.is_absolute(), f"{label}: executable path must be absolute")
    return path.resolve()


def artifact(value: Any, label: str, *, allow_missing: bool = False) -> Path:
    require(isinstance(value, dict), f"{label}: artifact is not an object")
    path = packet_path(value.get("path"), label)
    size = value.get("bytes")
    digest = value.get("sha256")
    require(type(size) is int and size >= 0, f"{label}: byte count is invalid")
    require(isinstance(digest, str) and HEX64.fullmatch(digest),
            f"{label}: SHA-256 is invalid")
    if not path.is_file() or path.is_symlink():
        require(allow_missing, f"{label}: artifact is missing: {path}")
        return path
    require(path.stat().st_size == size, f"{label}: byte count changed")
    require(sha256(path) == digest, f"{label}: SHA-256 changed")
    return path


def external_artifact(value: Any, label: str, cleanup: dict[str, Any] | None = None) -> dict[str, Any]:
    """Verify a binary while live or against the exact post-cleanup witness."""

    require(isinstance(value, dict), f"{label}: binary identity is missing")
    path = external_path(value.get("path"), label)
    size = value.get("bytes")
    digest = value.get("sha256")
    require(type(size) is int and size > 0, f"{label}: binary byte count is invalid")
    require(isinstance(digest, str) and HEX64.fullmatch(digest),
            f"{label}: binary SHA-256 is invalid")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == size, f"{label}: binary byte count changed")
        require(sha256(path) == digest, f"{label}: binary SHA-256 changed")
        return {"path": str(path), "bytes": size, "sha256": digest}
    require(cleanup is not None, f"{label}: binary is missing without cleanup witness")
    require(cleanup.get("target_removed") is True,
            "cleanup witness does not mark target_removed")
    require(cleanup.get("target") == str(TARGET), "cleanup target path changed")
    require(not TARGET.exists(), "owned target still exists after cleanup")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list), "cleanup removed_binaries is not a list")
    matches = [item for item in removed if isinstance(item, dict)
               and item.get("path") == str(path)]
    require(len(matches) == 1, f"{label}: exact removed binary witness is missing")
    witness = matches[0]
    require(witness.get("bytes") == size and witness.get("sha256") == digest,
            f"{label}: cleanup witness differs from binary identity")
    return {"path": str(path), "bytes": size, "sha256": digest}


def cleanup_witness() -> dict[str, Any] | None:
    path = PACKET / "cleanup.json"
    return read_json(path) if path.is_file() else None


def validate_cleanup(cleanup: dict[str, Any] | None,
                     binaries: dict[str, Any]) -> None:
    """When present, require the complete planned removed-binary set."""

    if cleanup is None:
        return
    require(cleanup.get("target_removed") is True,
            "cleanup witness does not mark target_removed")
    require(cleanup.get("target") == str(TARGET), "cleanup target path changed")
    require(not TARGET.exists(), "owned target still exists after cleanup")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list), "cleanup removed_binaries is not a list")
    expected = {str(external_path(value.get("path"), name))
                for name, value in binaries.items()}
    actual = {item.get("path") for item in removed if isinstance(item, dict)}
    require(len(removed) == len(expected) and actual == expected,
            "cleanup removed_binaries set differs from the planned binaries")
    for name, value in binaries.items():
        external_artifact(value, name, cleanup)


def source_identity() -> dict[str, Any]:
    """Bind current production files against the prior packet's source manifest."""

    source_path = PACKET / "build" / "source.json"
    require(source_path.is_file(), "build/source.json is missing")
    inheritance = read_json(PACKET / "inheritance.json")
    reference = inheritance.get("production_reference")
    require(isinstance(reference, dict), "production source inheritance is missing")
    raw_path = reference.get("path")
    require(isinstance(raw_path, str) and raw_path, "production source path is missing")
    copied_path = (PACKET / raw_path).resolve()
    require(copied_path.is_file() and not copied_path.is_symlink(),
            "inherited production source is missing")
    require(reference.get("sha256") == sha256(copied_path)
            and reference.get("bytes") == copied_path.stat().st_size,
            "inherited production source identity changed")
    source = read_json(source_path)
    copied = read_json(copied_path)
    origin = read_json(PACKET / "origin.json")
    require(source.get("revision") == origin.get("base"),
            "build source revision differs from origin base")
    require(source.get("files") == copied.get("files"),
            "build source files differ from prior frozen source")
    manifest = json.dumps(source.get("files"), sort_keys=True,
                          separators=(",", ":")).encode()
    return {"path": "build/source.json", "sha256": sha256(source_path),
            "revision": source.get("revision"),
            "file_count": len(source.get("files", {})),
            "file_manifest_sha256": hashlib.sha256(manifest).hexdigest()}


def normalize_command(command: Any) -> Any:
    """Map an archived absolute packet path to the current packet location."""

    if isinstance(command, list):
        return [normalize_command(value) for value in command]
    if isinstance(command, str):
        marker = "/change-0807/"
        if marker in command:
            prefix, suffix = command.split(marker, 1)
            if prefix.startswith("/"):
                return str(PACKET / suffix)
    return command


def _sealed_fixture(packet: str, relative: str, label: str) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    base = ROOT / "docs" / "performance" / "results" / packet
    seal_path = base / "seal.json"
    fixture_path = base / relative
    seal = read_json(seal_path)
    files = seal.get("files")
    require(isinstance(files, dict) and files.get(relative) == sha256(fixture_path),
            f"{packet}: fixture seal differs for {label}")
    value = read_json(fixture_path)
    samples = value.get("samples")
    require(isinstance(samples, list) and samples and isinstance(samples[0], dict),
            f"{packet}: fixture sample is missing for {label}")
    return value, samples[0], {
        "packet": packet, "path": str(fixture_path), "sha256": sha256(fixture_path),
        "seal_sha256": sha256(seal_path),
    }


def fixture(packet: str, shape: str) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    """Return a sealed capture fixture, its first sample, and seal identity."""

    require(shape in SHAPES, f"unknown shape: {shape}")
    return _sealed_fixture(packet, f"qualification/0-{shape}-capture-before.json", shape)


def oracle(shape: str) -> dict[str, Any]:
    """Require exact cross-packet source/output/semantic parity."""

    old, old_sample, old_meta = fixture("change-0780", shape)
    profile, profile_sample, profile_meta = _sealed_fixture(
        "change-0784", f"native/0-{shape}-profile.json", shape)
    current, current_sample, current_meta = fixture("change-0785", shape)
    require(old.get("source") == profile.get("source") == current.get("source"),
            f"{shape}: cross-packet source oracle differs")
    require(old_sample.get("output") == profile_sample.get("output")
            == current_sample.get("output"),
            f"{shape}: cross-packet output oracle differs")
    old_ver = old_sample.get("verification")
    profile_ver = profile_sample.get("verification")
    current_ver = current_sample.get("verification")
    require(isinstance(old_ver, dict) and isinstance(profile_ver, dict)
            and isinstance(current_ver, dict),
            f"{shape}: semantic oracle is malformed")
    semantic_fields = (
        "semantic_check", "reopened", "expected_text", "actual_text",
        "semantic_text_bytes", "semantic_text_sha256", "readback_bytes",
        "readback_sha256", "marker_matches",
    )
    expected_semantic = {key: old_ver.get(key) for key in semantic_fields}
    require({key: profile_ver.get(key) for key in semantic_fields}
            == expected_semantic
            == {key: current_ver.get(key) for key in semantic_fields},
            f"{shape}: cross-packet semantic oracle differs")
    return {
        "source": old["source"], "output": old_sample["output"],
        "verification": expected_semantic,
        "fixtures": [old_meta, profile_meta, current_meta],
    }


def check_report_identity(report: dict[str, Any], shape: str, *, samples: int,
                          warmup: int, binary: str) -> dict[str, Any]:
    expected = oracle(shape)
    require(report.get("schema") == SCHEMA, f"{shape}: report schema changed")
    require(report.get("tool") == TOOL, f"{shape}: report tool changed")
    require(report.get("mode") == "capture" and report.get("shape") == shape,
            f"{shape}: report mode/shape changed")
    require(report.get("timing_scope") == "Package::opened_presentation only",
            f"{shape}: timing scope changed")
    require(report.get("marker") == MARKER, f"{shape}: marker changed")
    require(report.get("source") == expected["source"], f"{shape}: source oracle changed")
    require((report.get("slides"), report.get("shapes_per_slide")) == DIMENSIONS[shape],
            f"{shape}: dimensions changed")
    require(report.get("warmup") == warmup and report.get("samples_requested") == samples,
            f"{shape}: sample policy changed")
    rows = report.get("samples")
    require(isinstance(rows, list) and len(rows) == samples,
            f"{shape}: sample count changed")
    require(report.get("allocator") == {
        "binary": binary, "allocator": "Rust system allocator",
        "instrumentation": "none", "counter_revision": None,
    }, f"{shape}: allocator identity changed")
    values: list[int] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("index") == index,
                f"{shape}: sample index changed")
        elapsed = row.get("elapsed_ns")
        require(type(elapsed) is int and elapsed > 0, f"{shape}: elapsed value is invalid")
        require(row.get("source_sha256") == expected["source"]["sha256"],
                f"{shape}: sample source hash changed")
        require(row.get("output") == expected["output"], f"{shape}: output oracle changed")
        verification = row.get("verification")
        require(isinstance(verification, dict), f"{shape}: verification is missing")
        require({key: verification.get(key) for key in expected["verification"]}
                == expected["verification"], f"{shape}: semantic oracle changed")
        metrics = row.get("metrics")
        require(metrics == {
            "elapsed_ns": elapsed, "slides": DIMENSIONS[shape][0],
            "shapes_per_slide": DIMENSIONS[shape][1],
            "captured_slides": DIMENSIONS[shape][0],
            "captured_shapes_per_slide": DIMENSIONS[shape][1],
        }, f"{shape}: metrics changed")
        require(row.get("allocation") is None or "allocation" not in row,
                f"{shape}: normal run contains allocation instrumentation")
        values.append(elapsed)
    return {"source": expected["source"], "output": expected["output"],
            "verification": expected["verification"], "values": values,
            "fixtures": expected["fixtures"]}


def load_legacy(module_name: str, filename: str) -> ModuleType:
    """Load one hash-bound 0784 parser without invoking its historical driver."""

    path = ROOT / "docs" / "performance" / "results" / "change-0784" / filename
    require(path.is_file(), f"missing historical parser: {path}")
    spec = importlib.util.spec_from_file_location(module_name, path)
    require(spec is not None and spec.loader is not None, f"cannot load parser: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_profile_parser() -> ModuleType:
    module = load_legacy("change0784_profile_parser_0807", "profile_analysis.py")
    module.HERE = PACKET
    return module


def load_perf_parser() -> ModuleType:
    old = ROOT / "docs" / "performance" / "results" / "change-0784"
    old_text = str(old)
    if old_text not in sys.path:
        sys.path.insert(0, old_text)
    return load_legacy("change0784_perf_parser_0807", "perf_analysis.py")
