"""Fail-closed custody helpers for the 0824 outer-whitespace trial.

The root agent owns Cargo, quality gates, workload execution, and cleanup.
This module only freezes identities, validates child receipts, and checks the
semantic fields emitted by the two packet-local probes.  It never starts a
workload or profiler itself.
"""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = Path("/home/zhuhe/code/litchi-target-0824")
BASE = "7b268927bfe3ed9c9000bfd567fc4ce656ec0cd9"
BASE_SHORT = "7b268927bf"
ALLOWLIST = (
    "crates/litchi-pptx/src/opened/transaction.rs",
    "crates/litchi-pptx/src/opened/xml.rs",
)
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
MODES = ("capture", "commit", "lifecycle")
PROBES = ("synthetic", "real")
DRIVER_FILES = ("plan.json", "custody.py", "freeze.py", "build.py", "capture.py", "quality.py")
CANDIDATE_FILES = (
    "candidate/before/transaction.rs",
    "candidate/before/xml.rs",
    "candidate/after/transaction.rs",
    "candidate/after/xml.rs",
    "candidate/combinedcandidate.patch",
    "candidate/design.md",
)
UNRELATED = {
    "docs/FORMAT_IMPLEMENTATION_REVIEW.md":
        "bffd00f144c4c1bbb3b0805d21352b40e9ae46f7366c03d6e581ca61b1f27ce5",
    "docs/UNIFIED_OPS_API_DESIGN.md":
        "f5672c38393a2a6c52f028b2a501ddad93766ef974dbdb46cdc3c45e1db3ef6d",
    "matrix-analysis.json":
        "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
}
REAL_INPUT = {
    "path": "test-data/ooxml/pptx/shapes.pptx",
    "bytes": 68822,
    "sha256": "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571",
}
REAL_REFERENCE = {
    "path": "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx",
    "bytes": 68284,
    "sha256": "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf",
}
REAL_OUTPUT = {"bytes": REAL_REFERENCE["bytes"], "sha256": REAL_REFERENCE["sha256"]}
REAL_FULL_TEXT_SHA256 = "5e9363235c4a1b158819a66be5aa46bf34d71b7da24592d855aad3ca90ce8b82"
REAL_MARKER = "litchi-perf-0638-ordinary-save"
REAL_SLIDE_COUNT = 6
RAW_ALLOC = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)


def sha(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def read(path: Path | str) -> Any:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write(path: Path | str, value: Any) -> None:
    Path(path).write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def artifact(path: Path | str) -> dict[str, Any]:
    path = Path(path)
    assert path.is_file() and not path.is_symlink(), f"missing artifact: {path}"
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def tracked(prefixes: tuple[str, ...] = (
    "crates", "Cargo.toml", "clippy.toml", ".cargo/config.toml", "rust-toolchain.toml",
)) -> list[str]:
    raw = subprocess.check_output(["git", "ls-files", "-z", "--", *prefixes], cwd=ROOT)
    return [name for name in raw.decode().split("\0") if name]


def source() -> dict[str, Any]:
    names = tracked()
    return {
        "revision": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip(),
        "files": {name: sha(ROOT / name) for name in names},
    }


def tool_source() -> dict[str, str]:
    return {name: sha(ROOT / name) for name in tracked(("tools/perf-baseline",))}


def probe_files(name: str) -> dict[str, str]:
    root = P / f"{name}-probe-src"
    assert root.is_dir() and not root.is_symlink(), f"missing probe source: {root}"
    result = {}
    for path in sorted(root.rglob("*")):
        if path.is_file():
            assert not path.is_symlink(), f"probe symlink: {path}"
            result[str(path.relative_to(root))] = sha(path)
    assert result, f"empty probe source: {name}"
    return result


def packet_files(names: tuple[str, ...]) -> dict[str, str]:
    result = {}
    for name in names:
        path = P / name
        assert path.is_file() and not path.is_symlink(), f"missing packet file: {name}"
        result[name] = sha(path)
    return result


def assert_candidate_archive() -> dict[str, str]:
    """Bind every archived production file before freezing the trial."""
    files = packet_files(CANDIDATE_FILES)
    patch_path = P / "candidate/combinedcandidate.patch"
    patch_text = patch_path.read_text(encoding="utf-8")
    assert patch_text.strip(), "combined candidate patch is empty"
    assert all(production in patch_text for production in ALLOWLIST)
    assert "crates/litchi-pptx/src/shape/reader.rs" not in patch_text
    assert (P / "candidate/design.md").stat().st_size > 0
    for production, before_name, after_name in (
        (ALLOWLIST[0], "candidate/before/transaction.rs", "candidate/after/transaction.rs"),
        (ALLOWLIST[1], "candidate/before/xml.rs", "candidate/after/xml.rs"),
    ):
        source = ROOT / production
        before = P / before_name
        after = P / after_name
        assert source.is_file() and before.is_file() and after.is_file()
        assert sha(source) == files[before_name], f"base archive mismatch: {production}"
        assert before.read_bytes() == source.read_bytes(), f"base archive bytes differ: {production}"
        assert before.read_bytes() != after.read_bytes(), f"candidate did not change: {production}"
    return files


def root_inputs() -> dict[str, str]:
    return {
        "Cargo.lock": sha(P / "inputs/root-Cargo.lock"),
        "rustfmt.toml": sha(P / "inputs/rustfmt.toml"),
        "tool-Cargo.lock": sha(P / "inputs/tool-Cargo.lock"),
    }


def assert_root_inputs() -> dict[str, str]:
    declared = read(P / "root-inputs.json")
    result = root_inputs()
    expected = {name: descriptor["sha256"] for name, descriptor in declared.items()
                if name in {"Cargo.lock", "rustfmt.toml", "tool-Cargo.lock"}}
    assert result == expected, "packet root-input hashes changed"
    for name, packet_name in (
        ("Cargo.lock", "root-Cargo.lock"),
        ("rustfmt.toml", "rustfmt.toml"),
        ("tools/perf-baseline/Cargo.lock", "tool-Cargo.lock"),
    ):
        source_path = ROOT / name
        packet_path = P / "inputs" / packet_name
        assert source_path.is_file() and packet_path.is_file()
        assert sha(source_path) == sha(packet_path), f"root input changed: {name}"
    return result


def architecture() -> dict[str, str]:
    declared = read(P / "architecture-inputs.json")
    assert isinstance(declared, dict) and len(declared) == 35
    actual = {name: sha(ROOT / name) for name in declared}
    assert actual == declared, "normative architecture input changed"
    return actual


def unrelated() -> dict[str, str]:
    actual = {name: sha(ROOT / name) for name in UNRELATED}
    assert actual == UNRELATED, "unrelated workspace file changed"
    return actual


def check_no_overrides() -> None:
    for name in (
        "RUSTUP_TOOLCHAIN", "RUSTFLAGS", "RUSTDOCFLAGS", "CARGO_ENCODED_RUSTFLAGS",
        "CARGO_BUILD_RUSTFLAGS", "CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS",
        "RUSTC_WRAPPER", "RUSTC_WORKSPACE_WRAPPER",
    ):
        assert not os.environ.get(name), f"inherited build override: {name}"


def assert_static() -> None:
    plan = read(P / "plan.json")
    assert plan["schema"] == "litchi.performance.0824.pptx-outer-whitespace.v1"
    assert plan["base"] == BASE and plan["cpu"] == 12
    assert plan["target"] == str(TARGET)
    assert plan["source_allowlist"] == list(ALLOWLIST)
    assert plan["candidate"]["production_files"] == list(ALLOWLIST)
    origin = read(P / "origin.json")
    assert origin["base"] == BASE
    assert origin["target"] == str(TARGET)
    assert origin["source_allowlist"] == list(ALLOWLIST)
    assert origin["production_changed_at_freeze"] is False
    assert origin["runtime_harness_changed"] is False and origin["tool_changed"] is False
    assert origin["unrelated"] == UNRELATED
    host = read(P / "host.json")
    assert host["affinity_selected"] == [12]
    assert host["target"] == str(TARGET)
    assert len(read(P / "architecture-inputs.json")) == 35
    assert_root_inputs()
    for name in PROBES:
        probe_files(name)
    for descriptor in (REAL_INPUT, REAL_REFERENCE):
        path = ROOT / descriptor["path"]
        assert path.is_file() and path.stat().st_size == descriptor["bytes"]
        assert sha(path) == descriptor["sha256"]
    check_no_overrides()


def changed_files(before: dict[str, Any], after: dict[str, Any]) -> set[str]:
    left = before["files"]
    right = after["files"]
    return {name for name in left.keys() | right.keys() if left.get(name) != right.get(name)}


def stable(frozen: dict[str, Any] | None = None) -> None:
    assert_static()
    if frozen is not None:
        assert source() == frozen["source"], "production source changed unexpectedly"
        stable_inputs(frozen)


def stable_inputs(frozen: dict[str, Any]) -> None:
    """Check packet inputs while allowing the one declared production edit."""
    assert assert_root_inputs() == frozen["root_inputs"], "root input custody changed"
    assert architecture() == frozen["architecture"], "architecture custody changed"
    assert unrelated() == frozen["unrelated"], "unrelated workspace file changed"
    assert tool_source() == frozen["tool"], "tool source changed unexpectedly"
    for name in PROBES:
        assert probe_files(name) == frozen["probes"][name], f"{name} probe changed"
    if "candidate" in frozen:
        assert packet_files(CANDIDATE_FILES) == frozen["candidate"], "candidate archive changed"
    if "drivers" in frozen:
        assert packet_files(DRIVER_FILES) == frozen["drivers"], "packet driver changed"
    for name in (
        "plan", "host", "toolchain", "root_input_manifest", "architecture_manifest",
        "corpus", "lock_parity",
    ):
        descriptor = frozen.get(name)
        assert isinstance(descriptor, dict), f"frozen metadata missing: {name}"
        assert artifact(descriptor["path"]) == descriptor, f"frozen metadata changed: {name}"


def assert_quality_before() -> dict[str, Any]:
    receipt = read(P / "quality-before.json")
    assert receipt.get("status") == "pass", "baseline production quality is not complete"
    source_path = ROOT / receipt["source"]
    quality_source = read(source_path)
    current = source()
    assert all(
        quality_source.get(name) == digest
        for name, digest in current["files"].items()
    ), "baseline quality source does not match the complete covered production source"
    assert ".cargo/config.toml" in quality_source, "quality source omitted .cargo/config.toml"
    return receipt


def historical_synthetic_report(shape: str, mode: str) -> dict[str, Any]:
    path = ROOT / "docs/performance/results/change-0806/qualification" / f"0-{shape}-{mode}-before.json"
    seal = read(ROOT / "docs/performance/results/change-0806/seal.json")
    relative = str(path.relative_to(ROOT / "docs/performance/results/change-0806"))
    assert seal.get("files", {}).get(relative) == sha(path), f"historical oracle changed: {relative}"
    return read(path)


def _is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in "0123456789abcdef" for c in value)


def _check_alloc(value: Any, label: str) -> None:
    assert isinstance(value, dict) and value.get("status") == "measured"
    assert value.get("scope") == "operation_global_system_allocator"
    for field in RAW_ALLOC:
        assert isinstance(value.get(field), int) and value[field] >= 0, f"{label}.{field}"
    assert value["failed_allocation_calls"] == 0
    assert value["live_bytes_after"] == value["live_bytes_before"] + value["allocated_bytes"] - value["deallocated_bytes"]
    assert value["region_peak_live_bytes"] >= value["live_bytes_before"]
    assert value["region_peak_live_bytes"] >= value["live_bytes_after"]


def check_synthetic_report(path: Path, *, shape: str, mode: str, samples: int,
                           warmup: int, allocation: bool, binary_name: str) -> dict[str, Any]:
    report = read(path)
    historical = historical_synthetic_report(shape, mode)
    assert set(report) == {"schema", "tool", "mode", "shape", "slides", "shapes_per_slide",
                           "timing_scope", "marker", "source", "fixture", "warmup",
                           "samples_requested", "samples", "allocator"}
    assert report["schema"] == "litchi.pptx.public-workflow-probe-0806.v1"
    assert report["tool"] == "public-pptx-probe-0806"
    assert report["mode"] == mode and report["shape"] == shape
    assert report["warmup"] == warmup and report["samples_requested"] == samples
    assert report["fixture"] == historical["fixture"]
    assert report["slides"] == historical["slides"]
    assert report["shapes_per_slide"] == historical["shapes_per_slide"]
    assert report["timing_scope"] == historical["timing_scope"]
    assert report["marker"] == historical["marker"]
    assert report["source"] == historical["source"]
    allocator = report["allocator"]
    assert allocator["binary"] == binary_name
    if allocation:
        assert allocator["instrumentation"] == "system_allocator_operation_scoped"
        assert allocator["counter_revision"] == "serialized_region_peak_v3"
    else:
        assert allocator["instrumentation"] == "none" and allocator["counter_revision"] is None
    rows = report["samples"]
    assert isinstance(rows, list) and len(rows) == samples
    historical_sample = historical["samples"][0]
    for index, sample in enumerate(rows):
        expected_keys = {"index", "elapsed_ns", "metrics", "source_sha256", "output", "verification"}
        if allocation:
            expected_keys.add("allocation")
        assert set(sample) == expected_keys and sample["index"] == index
        assert isinstance(sample["elapsed_ns"], int) and sample["elapsed_ns"] > 0
        assert sample["source_sha256"] == report["source"]["sha256"]
        expected_metrics = {"elapsed_ns", "slides", "shapes_per_slide"}
        if mode == "capture":
            expected_metrics |= {"captured_slides", "captured_shapes_per_slide"}
        assert isinstance(sample["metrics"], dict) and set(sample["metrics"]) == expected_metrics
        assert sample["metrics"]["elapsed_ns"] == sample["elapsed_ns"]
        assert sample["metrics"]["slides"] == report["slides"]
        assert sample["metrics"]["shapes_per_slide"] == report["shapes_per_slide"]
        if mode == "capture":
            assert sample["metrics"]["captured_slides"] == report["slides"]
            assert sample["metrics"]["captured_shapes_per_slide"] == report["shapes_per_slide"]
        verification = sample["verification"]
        common_verification = {
            "semantic_check", "reopened", "expected_text", "actual_text",
            "semantic_text_bytes", "semantic_text_sha256", "readback_bytes",
            "readback_sha256", "marker_matches", "unknown_namespace_check",
            "unknown_namespace_occurrences",
        }
        extension_verification = {
            "extension_preservation_check", "extension_text_tags",
            "extension_attributes_per_text_tag", "extension_attribute_occurrences",
            "extension_value_occurrences", "extension_namespace_declarations_per_slide",
            "extension_namespace_uris", "extension_attribute_names", "extension_attribute_values",
        }
        assert set(verification) == (common_verification | extension_verification
                                      if shape == "valid-4attr" else common_verification)
        assert verification["semantic_check"] is True and verification["reopened"] is True
        assert verification["expected_text"] == verification["actual_text"]
        assert isinstance(verification["semantic_text_bytes"], int)
        assert _is_sha(verification["semantic_text_sha256"])
        assert verification["readback_bytes"] == sample["output"]["bytes"]
        assert verification["readback_sha256"] == sample["output"]["sha256"]
        assert sample["output"] == historical_sample["output"]
        assert verification["semantic_text_bytes"] == historical_sample["verification"]["semantic_text_bytes"]
        assert verification["semantic_text_sha256"] == historical_sample["verification"]["semantic_text_sha256"]
        assert verification["marker_matches"] == historical_sample["verification"]["marker_matches"]
        assert verification["unknown_namespace_check"] == historical_sample["verification"]["unknown_namespace_check"]
        assert verification["unknown_namespace_occurrences"] == historical_sample["verification"]["unknown_namespace_occurrences"]
        if shape == "valid-4attr":
            for field in extension_verification:
                assert verification[field] == historical_sample["verification"][field]
        if allocation:
            _check_alloc(sample["allocation"], f"{shape}/{mode}/{index}")
    return report


def check_real_report(path: Path, *, samples: int, warmup: int, allocation: bool,
                      binary_name: str) -> dict[str, Any]:
    report = read(path)
    required = {"schema", "tool", "base_revision", "mode", "timing_scope", "input",
                "reference", "output", "marker", "target", "target_text",
                "full_text_sha256", "full_text_digest", "slide_count", "warmup",
                "samples_requested", "warmup_verified", "all_verified", "elapsed_ns",
                "samples", "allocator"}
    assert set(report) == required
    assert report["schema"] == "litchi.performance.0824.pptx-edit-trial.v1"
    assert report["base_revision"] == BASE_SHORT
    assert report["tool"] == "pptx-edit-profile-0822" and report["mode"] == "direct"
    assert report["input"] == REAL_INPUT and report["reference"] == REAL_REFERENCE
    assert report["output"] == REAL_OUTPUT and report["marker"] == REAL_MARKER
    assert report["target"] == {"slide": 0, "shape": 0}
    assert report["target_text"] == REAL_MARKER
    assert report["full_text_sha256"] == REAL_FULL_TEXT_SHA256
    assert report["full_text_digest"] == REAL_FULL_TEXT_SHA256
    assert report["slide_count"] == REAL_SLIDE_COUNT
    assert report["warmup"] == warmup and report["samples_requested"] == samples
    assert report["warmup_verified"] is True and report["all_verified"] is True
    allocator = report["allocator"]
    assert allocator["binary"] == binary_name
    if allocation:
        assert allocator["instrumentation"] == "system_allocator_operation_scoped"
        assert allocator["counter_revision"] == "serialized_region_peak_v3"
    else:
        assert allocator["instrumentation"] == "none" and allocator["counter_revision"] is None
    elapsed = report["elapsed_ns"]
    assert elapsed["unit"] == "ns" and elapsed["sample_order"] == list(range(samples))
    assert isinstance(elapsed["samples"], list) and len(elapsed["samples"]) == samples
    rows = report["samples"]
    assert isinstance(rows, list) and len(rows) == samples
    fields = {"all_verified", "input_hash_verified", "reference_hash_verified",
              "output_hash_verified", "output_size_verified", "output_bytes_verified",
              "reopened", "marker_verified", "target_verified",
              "full_text_digest_verified", "slide_count_verified"}
    for index, sample in enumerate(rows):
        expected = {"index", "elapsed_ns", "output", "verification"}
        if allocation:
            expected.add("allocation")
        assert set(sample) == expected and sample["index"] == index
        assert sample["elapsed_ns"] == elapsed["samples"][index] and sample["elapsed_ns"] > 0
        assert sample["output"] == REAL_OUTPUT
        verification = sample["verification"]
        assert set(verification) == fields and all(verification.values())
        if allocation:
            _check_alloc(sample["allocation"], f"real/{index}")
    return report
