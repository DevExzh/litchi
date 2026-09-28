"""Offline replay for the 0811 current-source native diagnostic.

The 0811 packet does not contain a candidate or an adoption gate.  It records
the cost of three deliberately different probe executables (ordinary,
non-inlined profiling wrapper, and frame-pointer wrapper) and a small native
``perf`` sample.  This reader never invokes Cargo, the probe, perf, or a
profiler.  It checks retained receipts, source/probe custody, the sealed 0810
public-workflow oracle, and the deterministic numerical summaries.
"""

from __future__ import annotations

import gzip
import hashlib
import importlib.util
import json
import math
import random
import statistics
import subprocess
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
SEALED_0810 = ROOT / "docs" / "performance" / "results" / "change-0810"
TARGET = Path("/home/zhuhe/code/litchi-target-0811")

SHAPES = ("tiny", "medium", "large")
VARIANTS = ("control", "profile", "fp")
CASES = tuple((shape, variant) for shape in SHAPES for variant in VARIANTS)
DIMENSIONS = {"tiny": (3, 4), "medium": (12, 8), "large": (100, 100)}
PROBE_SCHEMA = "litchi.pptx.public-workflow-probe-0806.v1"
PROBE_TOOL = "public-pptx-probe-0806"
MARKER = "litchi-perf-0780-static-mce-capabilities"
OWNER = "namespace_uri_probe::capture_region_0793"
BOOTSTRAP_SEED = 811811
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW = 250
BOOTSTRAP_HIGH = 9749
SOURCE_COUNT = 9196

SOURCE_KEYS = frozenset({"bytes", "sha256"})
FIXTURE_KEYS = frozenset({
    "injection", "slide_parts", "replaced_text_tags", "namespace_declarations",
    "namespaced_attributes", "namespace_uris", "attribute_names",
})
REPORT_KEYS = frozenset({
    "schema", "tool", "mode", "shape", "slides", "shapes_per_slide",
    "timing_scope", "marker", "source", "fixture", "warmup",
    "samples_requested", "samples", "allocator",
})
SAMPLE_KEYS = frozenset({
    "index", "elapsed_ns", "metrics", "source_sha256", "output", "verification",
})
VERIFICATION_KEYS = frozenset({
    "semantic_check", "reopened", "expected_text", "actual_text",
    "semantic_text_bytes", "semantic_text_sha256", "readback_bytes",
    "readback_sha256", "marker_matches", "unknown_namespace_check",
    "unknown_namespace_occurrences", "extension_preservation_check",
    "extension_text_tags", "extension_attributes_per_text_tag",
    "extension_attribute_occurrences", "extension_value_occurrences",
    "extension_namespace_declarations_per_slide", "extension_namespace_uris",
    "extension_attribute_names", "extension_attribute_values",
})


class ReplayError(RuntimeError):
    """A missing, stale, or contradictory retained artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReplayError(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ReplayError(f"invalid JSON {path}: {error}") from error


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(
        character in "0123456789abcdef" for character in value
    )


def integer(value: Any, label: str, *, positive: bool = False) -> None:
    require(
        isinstance(value, int)
        and not isinstance(value, bool)
        and (value > 0 if positive else value >= 0),
        f"{label}: invalid integer",
    )


def finite_number(value: Any, label: str, *, positive: bool = False) -> None:
    require(
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(float(value))
        and (value > 0 if positive else value >= 0),
        f"{label}: invalid number",
    )


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def packet_path(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}: path is missing")
    raw = Path(value)
    path = raw if raw.is_absolute() else PACKET / raw
    path = path.resolve()
    try:
        path.relative_to(PACKET.resolve())
    except ValueError as error:
        raise ReplayError(f"{label}: path escapes packet: {value}") from error
    return path


def artifact(value: Any, label: str, *, allow_missing: bool = False) -> Path:
    require(isinstance(value, dict), f"{label}: artifact is malformed")
    path = packet_path(value.get("path"), label)
    integer(value.get("bytes"), f"{label}.bytes")
    require(is_sha(value.get("sha256")), f"{label}.sha256 is malformed")
    if not path.is_file() or path.is_symlink():
        require(allow_missing, f"{label}: artifact is missing: {path}")
        return path
    require(path.stat().st_size == value["bytes"], f"{label}: byte count changed")
    require(sha(path) == value["sha256"], f"{label}: SHA-256 changed")
    return path


def retained_artifact(value: Any, label: str) -> Path:
    """Verify an artifact in a previously sealed packet outside this packet."""
    require(isinstance(value, dict), f"{label}: artifact is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and Path(raw).is_absolute(),
            f"{label}: retained path is not absolute")
    path = Path(raw).resolve()
    integer(value.get("bytes"), f"{label}.bytes")
    require(is_sha(value.get("sha256")), f"{label}.sha256 is malformed")
    require(path.is_file() and not path.is_symlink(), f"{label}: retained file is missing")
    require(path.stat().st_size == value["bytes"] and sha(path) == value["sha256"],
            f"{label}: retained identity changed")
    return path


def external_binary(value: Any, label: str, cleanup: dict[str, Any] | None) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: binary identity is malformed")
    path = Path(value.get("path", ""))
    require(path.is_absolute(), f"{label}: binary path is not absolute")
    integer(value.get("bytes"), f"{label}.bytes", positive=True)
    require(is_sha(value.get("sha256")), f"{label}.sha256 is malformed")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == value["bytes"] and sha(path) == value["sha256"],
                f"{label}: live binary identity changed")
        return {"path": str(path), "bytes": value["bytes"], "sha256": value["sha256"]}
    require(cleanup is not None and cleanup.get("target_removed") is True,
            f"{label}: binary is missing without cleanup witness")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list), "cleanup removed_binaries is malformed")
    matches = [row for row in removed if isinstance(row, dict) and row.get("path") == str(path)]
    require(
        len(matches) == 1
        and matches[0].get("bytes") == value["bytes"]
        and matches[0].get("sha256") == value["sha256"],
        f"{label}: exact removed binary witness is missing",
    )
    return {"path": str(path), "bytes": value["bytes"], "sha256": value["sha256"]}


def source_files(value: Any, label: str) -> dict[str, str]:
    require(isinstance(value, dict) and isinstance(value.get("files"), dict),
            f"{label}: source manifest is malformed")
    files = value["files"]
    require(files and all(isinstance(name, str) and is_sha(digest)
                          for name, digest in files.items()),
            f"{label}: source file census is malformed")
    require(len(files) == SOURCE_COUNT, f"{label}: source file count changed")
    return dict(files)


def current_source_files() -> dict[str, str]:
    raw = subprocess.check_output(
        ["git", "ls-files", "-z", "--", "crates", "Cargo.toml", "clippy.toml",
         ".cargo/config.toml", "rust-toolchain.toml"],
        cwd=ROOT,
    )
    names = [name for name in raw.decode().split("\0") if name]
    require(len(names) == SOURCE_COUNT, "live production source file count changed")
    return {name: sha(ROOT / name) for name in names}


def source_manifest(path: Path, label: str) -> dict[str, Any]:
    value = read(path)
    return {"files": source_files(value, label)}


def load_cleanup() -> dict[str, Any] | None:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None
    value = read(path)
    require(value.get("schema") == "litchi.performance.0811.cleanup.v1",
            "cleanup schema changed")
    require(value.get("target") == str(TARGET)
            and value.get("target_removed") is True
            and value.get("binaries_verified_before_removal") is True,
            "cleanup target contract changed")
    removed = value.get("removed_binaries")
    require(isinstance(removed, list) and len(removed) == 3,
            "cleanup binary count changed")
    for row in removed:
        require(isinstance(row, dict) and Path(row.get("path", "")).is_absolute()
                and isinstance(row.get("bytes"), int) and row["bytes"] > 0
                and is_sha(row.get("sha256")), "cleanup binary witness malformed")
    source_receipt = value.get("source_manifest")
    if isinstance(source_receipt, dict):
        artifact(source_receipt, "cleanup source manifest")
    return value


def sealed_source() -> dict[str, str]:
    path = SEALED_0810 / "build-after" / "source.json"
    seal = read(SEALED_0810 / "seal.json")
    relative_path = "docs/performance/results/change-0810/build-after/source.json"
    # The 0810 packet seal uses packet-relative names.
    if relative_path not in seal.get("files", {}):
        relative_path = "build-after/source.json"
    require(seal.get("files", {}).get(relative_path) == sha(path),
            "sealed 0810 after source manifest changed")
    return source_files(read(path), "sealed 0810 after source")


def root_inputs() -> dict[str, str]:
    origin = read(PACKET / "origin.json")
    root = origin.get("root_inputs")
    require(isinstance(root, dict), "origin root inputs are missing")
    expected = {
        "Cargo.lock": root.get("root-Cargo.lock"),
        "rustfmt.toml": root.get("rustfmt.toml"),
    }
    require(all(is_sha(value) for value in expected.values()),
            "origin root input hashes are malformed")
    copies = {
        "Cargo.lock": PACKET / "inputs" / "root-Cargo.lock",
        "rustfmt.toml": PACKET / "inputs" / "rustfmt.toml",
    }
    for name, path in copies.items():
        require(path.is_file() and sha(path) == expected[name],
                f"frozen root input changed: {name}")
        require((ROOT / name).is_file() and sha(ROOT / name) == expected[name],
                f"live root input changed: {name}")
    return expected


def probe_custody() -> dict[str, str]:
    old_probe = SEALED_0810 / "probe-src"
    old_inputs = read(SEALED_0810 / "inheritance.json").get("probe", {})
    expected = old_inputs.get("files") if isinstance(old_inputs, dict) else None
    if not isinstance(expected, dict):
        expected = {
            str(path.relative_to(old_probe)): sha(path)
            for path in sorted(old_probe.rglob("*")) if path.is_file()
        }
    require(set(expected) == {"Cargo.lock", "Cargo.toml", "Cargo.toml.template",
                              "src/main.rs", "src/counting_allocator.rs",
                              "src/allocation_metrics.rs"},
            "sealed probe file set changed")
    current = {
        str(path.relative_to(PACKET / "probe-src")): sha(path)
        for path in sorted((PACKET / "probe-src").rglob("*")) if path.is_file()
    }
    require(current == expected, "0811 probe is not the exact sealed 0810 copy")
    return expected


def plan_and_custody() -> tuple[dict[str, Any], dict[str, str], dict[str, str]]:
    plan = read(PACKET / "plan.json")
    require(plan.get("schema") == "litchi.performance.0811.current-production.v1",
            "plan schema changed")
    require(plan.get("source_revision") == "2894fbd628434bad4ef45af4f0f0b9a8d4d23468"
            and plan.get("source_file_count") == SOURCE_COUNT
            and plan.get("source_allowlist") == []
            and plan.get("cpu") == 12
            and plan.get("shapes") == list(SHAPES)
            and plan.get("modes") == ["capture"], "source or case scope changed")
    native = plan.get("native")
    require(isinstance(native, dict), "native plan missing")
    require(native.get("blocks") == 6 and native.get("orders") == [
        ["control", "profile", "fp"], ["profile", "fp", "control"],
        ["fp", "control", "profile"], ["fp", "profile", "control"],
        ["profile", "control", "fp"], ["control", "fp", "profile"],
    ] and native.get("shapes") == list(SHAPES)
            and native.get("samples") == 30 and native.get("warmup") == 3
            and native.get("reports") == 54 and native.get("samples_total") == 1620,
            "native plan changed")
    perf = plan.get("perf")
    require(isinstance(perf, dict)
            and perf.get("binary") == "fp" and perf.get("mode") == "capture"
            and perf.get("shape") == "large" and perf.get("orders") == [["fp"], ["fp"]]
            and perf.get("repeats") == 2 and perf.get("samples") == 100
            and perf.get("warmup") == 0 and perf.get("cpu") == 12
            and perf.get("event") == "cycles:u" and perf.get("frequency_hz") == 499
            and perf.get("call_graph") == "fp"
            and perf.get("owner") == OWNER and perf.get("reports") == 2
            and perf.get("samples_total") == 200, "perf plan changed")
    stats_plan = plan.get("statistics")
    require(isinstance(stats_plan, dict)
            and stats_plan.get("native") == ["p50", "p95", "p99", "mean", "spread", "rss"]
            and stats_plan.get("paired_ratios") == ["profile/control", "fp/profile"]
            and stats_plan.get("adoption_threshold") is None
            and stats_plan.get("speedup_claim") is False
            and stats_plan.get("allocation_lane") is False
            and stats_plan.get("callgrind_lane") is False,
            "statistics policy changed")
    bootstrap = stats_plan.get("bootstrap")
    require(bootstrap == {
        "resamples": BOOTSTRAP_RESAMPLES, "seed": BOOTSTRAP_SEED,
        "statistic": "median", "sorted_zero_based_endpoints": [BOOTSTRAP_LOW, BOOTSTRAP_HIGH],
    }, "bootstrap policy changed")
    root_inputs()
    probe_custody()
    expected_source = sealed_source()
    live = current_source_files()
    require(live == expected_source, "live production source differs from sealed 0810 after")
    return plan, expected_source, live


def fixture(shape: str) -> dict[str, Any]:
    require(shape in SHAPES, f"unknown shape: {shape}")
    relative_path = f"native/0-{shape}-capture-after.json"
    path = SEALED_0810 / relative_path
    seal = read(SEALED_0810 / "seal.json")
    sealed_key = f"docs/performance/results/change-0810/{relative_path}"
    if sealed_key not in seal.get("files", {}):
        sealed_key = relative_path
    require(seal.get("files", {}).get(sealed_key) == sha(path),
            f"sealed 0810 fixture changed: {shape}")
    report = read(path)
    require(set(report) == REPORT_KEYS, f"sealed fixture fields changed: {shape}")
    fixture_value = report.get("fixture")
    require(isinstance(fixture_value, dict) and set(fixture_value) == FIXTURE_KEYS,
            f"sealed fixture map malformed: {shape}")
    return report


def validate_report(report: dict[str, Any], shape: str, binary_name: str,
                    samples: int, warmup: int, label: str) -> dict[str, Any]:
    oracle = fixture(shape)
    require(set(report) == REPORT_KEYS, f"{label}: report fields changed")
    require(report.get("schema") == PROBE_SCHEMA and report.get("tool") == PROBE_TOOL
            and report.get("mode") == "capture" and report.get("shape") == shape
            and report.get("timing_scope") == "Package::opened_presentation only"
            and report.get("marker") == MARKER,
            f"{label}: probe identity changed")
    require(report.get("slides") == oracle.get("slides")
            and report.get("shapes_per_slide") == oracle.get("shapes_per_slide")
            and report.get("fixture") == oracle.get("fixture")
            and report.get("source") == oracle.get("source"),
            f"{label}: fixture/source oracle changed")
    require(report.get("warmup") == warmup and report.get("samples_requested") == samples,
            f"{label}: sample policy changed")
    source = report.get("source")
    require(isinstance(source, dict) and set(source) == SOURCE_KEYS
            and is_sha(source.get("sha256")), f"{label}: source identity malformed")
    integer(source.get("bytes"), f"{label}: source bytes", positive=True)
    allocator = report.get("allocator")
    require(isinstance(allocator, dict)
            and allocator.get("allocator") == "Rust system allocator"
            and allocator.get("instrumentation") == "none"
            and allocator.get("counter_revision") is None
            and allocator.get("binary") == binary_name,
            f"{label}: allocator identity changed")
    rows = report.get("samples")
    require(isinstance(rows, list) and len(rows) == samples, f"{label}: sample count changed")
    values: list[int] = []
    outputs: list[tuple[int, str]] = []
    expected_semantic = oracle["samples"][0].get("verification")
    require(isinstance(expected_semantic, dict),
            f"{label}: sealed semantic oracle malformed")
    expected_output = oracle["samples"][0].get("output")
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and set(row) == SAMPLE_KEYS
                and row.get("index") == index, f"{label}: sample fields changed")
        integer(row.get("elapsed_ns"), f"{label} sample {index}: elapsed", positive=True)
        values.append(row["elapsed_ns"])
        require(row.get("source_sha256") == source["sha256"]
                and row.get("output") == expected_output,
                f"{label} sample {index}: source/output oracle changed")
        verification = row.get("verification")
        require(verification == expected_semantic,
                f"{label} sample {index}: semantic verification changed")
        metrics = row.get("metrics")
        require(isinstance(metrics, dict)
                and metrics == {
                    "captured_shapes_per_slide": report["shapes_per_slide"],
                    "captured_slides": report["slides"],
                    "elapsed_ns": row["elapsed_ns"],
                    "shapes_per_slide": report["shapes_per_slide"],
                    "slides": report["slides"],
                }, f"{label} sample {index}: metrics changed")
        outputs.append((expected_output["bytes"], expected_output["sha256"]))
    require(len(set(outputs)) == 1, f"{label}: nondeterministic output")
    return {
        "source": source,
        "output": expected_output,
        "verification": expected_semantic,
        "elapsed": values,
        "stats": distribution(values),
        "output_identity": {"bytes": expected_output["bytes"],
                             "sha256": expected_output["sha256"]},
    }


def nearest(values: Iterable[float], percentile: float) -> float:
    ordered = sorted(values)
    require(ordered, "empty quantile vector")
    return ordered[max(1, math.ceil(len(ordered) * percentile)) - 1]


def distribution(values: Iterable[float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "empty timing vector")
    for index, value in enumerate(vector):
        finite_number(value, f"timing[{index}]", positive=True)
    return {
        "count": len(vector), "values": vector, "min": min(vector),
        "p50": nearest(vector, .50), "mean": statistics.mean(vector),
        "p95": nearest(vector, .95), "p99": nearest(vector, .99), "max": max(vector),
    }


def spread(values: Iterable[float]) -> float:
    vector = list(values)
    require(vector, "empty spread vector")
    low, high = min(vector), max(vector)
    return 0.0 if low == high else (float("inf") if low == 0 else (high - low) * 100.0 / abs(low))


def pair_ratio(left: float, right: float) -> dict[str, Any]:
    require(left > 0 and right > 0, "paired timing metric is not positive")
    ratio = right / left
    return {
        "before": left, "after": right, "ratio": ratio,
        "change_percent": (ratio - 1.0) * 100.0,
        "relative_change_defined": True,
        "zero_baseline_equal": False, "zero_to_nonzero": False,
        "over_5_percent": ratio > 1.05,
    }


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(len(values) == 6, "paired bootstrap must use six block ratios")
    rng = random.Random(BOOTSTRAP_SEED)
    draws = sorted(
        statistics.median(values[rng.randrange(len(values))] for _ in values)
        for _ in range(BOOTSTRAP_RESAMPLES)
    )
    return {
        "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
        "statistic": "median", "confidence": .95,
        "low_rank": BOOTSTRAP_LOW, "high_rank": BOOTSTRAP_HIGH,
        "ci_low": draws[BOOTSTRAP_LOW], "ci_high": draws[BOOTSTRAP_HIGH],
    }


def binary_name(value: Any, label: str) -> str:
    if isinstance(value, str):
        return value
    if isinstance(value, dict):
        path = value.get("path")
        require(isinstance(path, str) and path, f"{label}: binary name is missing")
        return Path(path).name
    raise ReplayError(f"{label}: binary name is malformed")


def load_build(cleanup: dict[str, Any] | None) -> tuple[dict[str, Any], dict[str, Any]]:
    candidates = [PACKET / "build" / "build.json", PACKET / "build.json",
                  PACKET / "build-current" / "build.json"]
    paths = [path for path in candidates if path.is_file()]
    require(len(paths) == 1, "expected one 0811 build manifest")
    path = paths[0]
    manifest = read(path)
    require(isinstance(manifest, dict)
            and str(manifest.get("schema", "")).startswith("litchi.performance.0811.build"),
            "build schema changed")
    source_receipt = manifest.get("source")
    if isinstance(source_receipt, dict) and "path" in source_receipt:
        source_path = artifact(source_receipt, "build source")
    else:
        source_path = path.parent / "source.json"
        require(source_path.is_file(), "build source receipt is missing")
    recorded_source = source_manifest(source_path, "build source")
    sealed = sealed_source()
    require(recorded_source["files"] == sealed, "build source differs from sealed 0810 after")
    binaries = manifest.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == set(VARIANTS),
            "build binary matrix changed")
    verified = {
        variant: external_binary(binaries[variant], f"build {variant}", cleanup)
        for variant in VARIANTS
    }
    if cleanup is not None:
        expected_paths = {identity["path"] for identity in verified.values()}
        actual_paths = {row["path"] for row in cleanup["removed_binaries"]}
        require(actual_paths == expected_paths, "cleanup binary witness set changed")
        source_witness = cleanup.get("source_manifest")
        if isinstance(source_witness, dict):
            source_witness_path = artifact(source_witness, "cleanup source manifest")
            require(source_files(read(source_witness_path), "cleanup source manifest")
                    == recorded_source["files"], "cleanup source witness changed")
    rows = manifest.get("rows")
    require(isinstance(rows, list) and len(rows) == 3, "build command cardinality changed")
    seen_rows = set()
    for row in rows:
        require(isinstance(row, dict) and row.get("variant") in VARIANTS
                and row["variant"] not in seen_rows and row.get("exit_code") == 0,
                "build variant receipt changed")
        seen_rows.add(row["variant"])
    require(seen_rows == set(VARIANTS), "build variant receipts incomplete")
    frozen_ref = manifest.get("frozen_inputs")
    require(isinstance(frozen_ref, dict), "build frozen inputs are missing")
    frozen_path = artifact(frozen_ref, "build frozen inputs")
    frozen = read(frozen_path)
    require(frozen.get("schema") == "litchi.performance.0811.frozen-inputs.v1"
            and frozen.get("source", {}).get("files") == recorded_source["files"]
            and frozen.get("probe") == probe_custody()
            and frozen.get("root_inputs") == root_inputs(),
            "build frozen custody changed")
    drivers = frozen.get("drivers")
    require(isinstance(drivers, dict) and drivers,
            "build frozen driver hashes are missing")
    for name, digest in drivers.items():
        driver_path = PACKET / name
        require(driver_path.is_file() and is_sha(digest) and sha(driver_path) == digest,
                f"frozen driver changed: {name}")
    probe = manifest.get("probe")
    if isinstance(probe, dict):
        expected = probe_custody()
        # Build manifests may retain absolute packet paths or packet-relative
        # source names; compare by the six relative names.
        normalized = {}
        for name, digest in probe.items():
            normalized[str(Path(name).as_posix()).removeprefix("probe-src/")] = digest
        require(normalized == expected, "build probe manifest changed")
    return manifest, {"source": recorded_source, "binaries": verified,
                       "manifest_path": relative(path)}


def normalize_command(value: Any) -> Any:
    if not isinstance(value, list):
        return value
    result = []
    for item in value:
        if isinstance(item, str):
            # Root drivers use absolute packet paths.  Keep command identity
            # strict while permitting the packet's absolute path to move.
            marker = "/change-0811/"
            if marker in item and item.startswith("/"):
                item = str(PACKET / item.split(marker, 1)[1])
        result.append(item)
    return result


def native_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        {"block": block, "shape": shape, "variant": variant,
         "samples": plan["native"]["samples"], "warmup": plan["native"]["warmup"]}
        for block, order in enumerate(plan["native"]["orders"])
        for shape in SHAPES for variant in order
    ]


def row_variant(row: dict[str, Any], label: str) -> str:
    for key in ("variant", "binary_name", "leg", "kind"):
        value = row.get(key)
        if isinstance(value, str) and value in VARIANTS:
            return value
    binary = row.get("binary")
    if isinstance(binary, dict):
        name = Path(str(binary.get("path", ""))).name
        for variant in VARIANTS:
            if name == variant or name.endswith("-" + variant):
                return variant
    raise ReplayError(f"{label}: binary variant is missing")


def receipt_source(path: Path, label: str) -> dict[str, str]:
    value = read(path)
    return source_files(value, label)


def load_native(plan: dict[str, Any], build: dict[str, Any],
                cleanup: dict[str, Any] | None) -> list[dict[str, Any]]:
    directory = PACKET / "native"
    complete = read(directory / "complete.json")
    require(complete.get("schema") == "litchi.performance.0811.native.complete.v1"
            and complete.get("blocks") == 6 and complete.get("reports") == 54
            and complete.get("samples") == 1620
            and complete.get("plan_sha256") == sha(PACKET / "plan.json")
            and complete.get("build_sha256") == sha(PACKET / "build" / "build.json"),
            "native completion receipt changed")
    source_ref = complete.get("source")
    if isinstance(source_ref, dict):
        source_path = artifact(source_ref, "native source")
        require(source_files(read(source_path), "native source") == build["source"]["files"],
                "native source census changed")
    receipts = read(artifact(complete.get("receipts"), "native receipts"))
    require(isinstance(receipts, list) and len(receipts) == 54,
            "native receipt cardinality changed")
    jobs = native_jobs(plan)
    entries: list[dict[str, Any]] = []
    previous_end = float("-inf")
    for row, job in zip(receipts, jobs):
        label = f"native/{job['block']}-{job['shape']}-{job['variant']}"
        require(isinstance(row, dict), f"{label}: receipt malformed")
        require(row.get("schema") == "litchi.performance.0811.native-receipt.v1"
                and row.get("lane") == "native" and row.get("mode") == "capture"
                and row.get("block") == job["block"] and row.get("shape") == job["shape"]
                and row_variant(row, label) == job["variant"], f"{label}: matrix changed")
        require(row.get("exit_code") == 0, f"{label}: process failed")
        started, ended = row.get("started"), row.get("ended")
        finite_number(started, f"{label}: started")
        finite_number(ended, f"{label}: ended")
        require(previous_end <= started <= ended, f"{label}: receipt order changed")
        previous_end = ended
        log_path = artifact(row.get("log"), f"{label} log")
        require(not log_path.read_bytes(), f"{label}: workload log is not empty")
        rss_path = artifact(row.get("rss"), f"{label} RSS")
        rss_text = rss_path.read_text(encoding="utf-8").strip()
        require(rss_text.isdigit() and int(rss_text) > 0, f"{label}: RSS malformed")
        report_path = artifact(row.get("report"), f"{label} report")
        binary = build["binaries"][job["variant"]]
        source_ref = row.get("source")
        require(isinstance(source_ref, dict), f"{label}: source receipt missing")
        source_path = artifact(source_ref, f"{label} source")
        require(source_files(read(source_path), f"{label} source") == build["source"]["files"],
                f"{label}: source census changed")
        require(row.get("probe") == probe_custody()
                and row.get("root_inputs") == root_inputs(), f"{label}: custody changed")
        command = [
            "/usr/bin/time", "-f", "%M", "-o", row["rss"]["path"],
            "taskset", "-c", str(plan["cpu"]), binary["path"],
            "--mode", "capture", "--shape", job["shape"],
            "--samples", str(job["samples"]), "--warmup", str(job["warmup"]),
            "--output", row["report"]["path"],
        ]
        require(normalize_command(row.get("command")) == normalize_command(command),
                f"{label}: command changed")
        report = read(report_path)
        outcome = validate_report(report, job["shape"], Path(binary["path"]).name,
                                  job["samples"], job["warmup"], label)
        stats = outcome["stats"]
        item = {
            "block": job["block"], "shape": job["shape"], "variant": job["variant"],
            "sample_count": len(outcome["elapsed"]), "stats": stats,
            "rss_kib": int(rss_text), "report": relative(report_path),
            "report_sha256": sha(report_path), "rss_artifact": relative(rss_path),
            "log_artifact": relative(log_path), "command": row["command"],
            "binary": binary,
            "source_output_semantic_identity_matches": True,
        }
        entries.append(item)
    require(sum(item["sample_count"] for item in entries) == 1620,
            "native sample cardinality changed")
    return entries


def paired_summary(entries: list[dict[str, Any]]) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    lookup = {(e["block"], e["shape"], e["variant"]): e for e in entries}
    metrics = ("p50", "mean", "p95", "p99", "rss_kib")
    summaries: dict[str, Any] = {}
    spread_flags: list[dict[str, Any]] = []
    for shape in SHAPES:
        variants: dict[str, Any] = {}
        for variant in VARIANTS:
            rows = [lookup[(block, shape, variant)] for block in range(6)]
            values: dict[str, Any] = {}
            for metric in metrics:
                vector = [row["stats"][metric] if metric != "rss_kib" else row[metric]
                          for row in rows]
                spread_value = spread(vector)
                values[metric] = {
                    "values": vector, "median": statistics.median(vector),
                    "spread_percent": spread_value,
                }
                if spread_value > 5.0:
                    spread_flags.append({"shape": shape, "variant": variant,
                                         "metric": metric, "spread_percent": spread_value,
                                         "descriptive_only": True})
            variants[variant] = values
        ratios: dict[str, Any] = {}
        for pair, numerator, denominator in (("profile/control", "profile", "control"),
                                             ("fp/profile", "fp", "profile")):
            metric_rows: dict[str, Any] = {}
            for metric in metrics:
                values = []
                by_block = []
                for block in range(6):
                    left = (lookup[(block, shape, denominator)]["stats"][metric]
                            if metric != "rss_kib" else lookup[(block, shape, denominator)][metric])
                    right = (lookup[(block, shape, numerator)]["stats"][metric]
                             if metric != "rss_kib" else lookup[(block, shape, numerator)][metric])
                    item = {"block": block, **pair_ratio(left, right)}
                    by_block.append(item)
                    values.append(item["ratio"])
                metric_rows[metric] = {
                    "by_block": by_block,
                    "ratio_values": values,
                    "ratio_median": statistics.median(values),
                    "change_percent_median": (statistics.median(values) - 1.0) * 100.0,
                    "bootstrap": bootstrap(values),
                    "spread_percent": spread(values),
                }
            ratios[pair] = metric_rows
        summaries[shape] = {"variants": variants, "paired_ratios": ratios}
    return summaries, spread_flags


def load_quality_reuse() -> dict[str, Any]:
    # The production six-gate output belongs to 0810 and is deliberately
    # reused.  A packet receipt, when present, must point to that exact file;
    # accepting the sealed source directly keeps this reader stable after the
    # 0811 packet is committed.
    old = SEALED_0810 / "quality-summary.json"
    require(old.is_file(), "sealed 0810 quality summary is missing")
    old_hash = sha(old)
    candidates = [PACKET / name for name in
                  ("quality.json", "quality-reuse.json", "quality-reused.json",
                   "quality-summary.json")]
    receipt = next((path for path in candidates if path.is_file()), None)
    if receipt is not None:
        value = read(receipt)
        require(isinstance(value, dict), "quality reuse receipt malformed")
        require(value.get("schema") == "litchi.performance.0811.quality-reuse.v1"
                and value.get("gate_count") == 6
                and value.get("cargo_executed") is False,
                "quality reuse receipt changed")
        source_ref = value.get("source")
        if isinstance(source_ref, dict):
            source_path = retained_artifact(source_ref, "quality reuse source")
            require(source_path == (SEALED_0810 / "build-after/source.json").resolve(),
                    "quality reuse source path changed")
        inputs_receipt = value.get("inputs")
        require(isinstance(inputs_receipt, dict), "quality reuse inputs are missing")
        inputs_path = artifact(inputs_receipt, "quality reuse inputs")
        inputs = read(inputs_path)
        require(inputs.get("sealed_0810_source_files_equal") is True,
                "quality reuse source equality changed")
        summary_ref = inputs.get("sealed_0810_summary")
        require(isinstance(summary_ref, dict) and summary_ref.get("sha256") == old_hash,
                "quality reuse summary changed")
        retained_artifact(summary_ref, "quality reuse summary")
        tests = value.get("tests")
        require(isinstance(tests, dict) and tests.get("passed") == 1241
                and tests.get("failed") == 0 and tests.get("ignored") == 3
                and tests.get("suites") == 85, "reused production test counts changed")
        gates = value.get("gates")
        require(isinstance(gates, list) and len(gates) == 6
                and all(isinstance(gate, dict) and gate.get("exit_code") == 0
                        and gate.get("status") == "pass" for gate in gates),
                "reused production gates changed")
        for index, gate in enumerate(gates):
            retained_artifact(gate.get("log"), f"reused production gate {index + 1} log")
    return {"packet": "change-0810", "path": "quality-summary.json",
            "sha256": old_hash, "reused": True, "gate_count": 6}


def load_probe_quality() -> dict[str, Any]:
    path = PACKET / "probe-quality.json"
    require(path.is_file(), "fresh probe quality completion is missing")
    value = read(path)
    require(value.get("schema") == "litchi.performance.0811.probe-quality.v1"
            and value.get("gate_count") == 3 and value.get("tests_passed") == 36,
            "fresh probe quality counts changed")
    complete = read(PACKET / "probe-quality/complete.json")
    require(complete == {"schema": "litchi.performance.0811.probe-quality.complete.v1",
                         **{key: value[key] for key in ("gate_count", "tests_passed", "inputs", "receipts")}},
            "fresh probe completion differs from its summary")
    inputs_path = artifact(value.get("inputs"), "probe quality inputs")
    inputs = read(inputs_path)
    require(source_files(inputs.get("source"), "probe quality source") == sealed_source()
            and inputs.get("probe") == probe_custody()
            and inputs.get("root_inputs") == root_inputs(),
            "probe quality custody changed")
    receipts_path = artifact(value.get("receipts"), "probe quality receipts")
    receipts = read(receipts_path)
    require(isinstance(receipts, list) and len(receipts) == 3,
            "probe quality receipt cardinality changed")
    for index, row in enumerate(receipts):
        require(isinstance(row, dict) and row.get("gate") == index + 1
                and row.get("exit_code") == 0, f"probe quality gate {index + 1} changed")
        artifact(row.get("log"), f"probe quality gate {index + 1} log")
    require(value.get("target") == str(TARGET / "probe-quality"),
            "probe quality target changed")
    return {"path": relative(path), "sha256": sha(path), "gate_count": 3,
            "tests_passed": 36}


def load_hash_bound_perf_parser() -> Any:
    # 0784's parser is used only as a syntax decoder.  Its bytes are sealed by
    # 0784 and checked before loading; none of its historical drivers execute.
    path = ROOT / "docs" / "performance" / "results" / "change-0784" / "perf_analysis.py"
    seal = read(ROOT / "docs" / "performance" / "results" / "change-0784" / "seal.json")
    expected = seal.get("files", {}).get("perf_analysis.py")
    if expected is None:
        expected = seal.get("files", {}).get("perf_analysis.py")
    require(is_sha(expected) and sha(path) == expected, "sealed 0784 perf parser changed")
    module_name = "sealed_change0784_perf_parser_0811"
    spec = importlib.util.spec_from_file_location(module_name, path)
    require(spec is not None and spec.loader is not None, "cannot load sealed perf parser")
    module = importlib.util.module_from_spec(spec)
    # The old parser imports its old custody/native modules.  They are only
    # imported to expose samples(); adding the historical directory is enough.
    old = str(path.parent)
    if old not in sys.path:
        sys.path.insert(0, old)
    spec.loader.exec_module(module)
    require(callable(getattr(module, "samples", None)), "sealed perf parser API changed")
    return module


def gzip_or_plain(value: Any, label: str) -> tuple[bytes, dict[str, Any] | None]:
    require(isinstance(value, dict), f"{label}: frame artifact is missing")
    path = artifact(value, label)
    data = path.read_bytes()
    original = None
    if path.name.endswith(".gz"):
        try:
            with gzip.open(path, "rb") as stream:
                data = stream.read()
        except (OSError, EOFError) as error:
            raise ReplayError(f"{label}: invalid gzip: {error}") from error
        original = value.get("original") if isinstance(value.get("original"), dict) else None
    return data, original


def parse_perf_stream(data: bytes, parser: Any, label: str) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    # Normal perf-script output has one parser block per sampled event.  Lost
    # event notices do not have a call stack; count them separately and retain
    # the parser's strict sample decoder for all actual stacks.
    text = data.decode("utf-8", errors="replace")
    lost_lines = [line for line in text.splitlines()
                  if "PERF_RECORD_LOST" in line or "lost" in line.lower()
                  and "sample" in line.lower()]
    try:
        samples = parser.samples(data)
    except (UnicodeError, ValueError, TypeError, IndexError, AssertionError) as error:
        # A status line can make the historical parser reject the whole stream;
        # remove only blocks whose header is explicitly a lost event.
        blocks = []
        for block in text.strip().split("\n\n"):
            first = block.splitlines()[0] if block.splitlines() else ""
            if "PERF_RECORD_LOST" not in first and not ("lost" in first.lower()
                                                         and "sample" in first.lower()):
                blocks.append(block)
        try:
            samples = parser.samples(("\n\n".join(blocks) + "\n").encode())
        except (UnicodeError, ValueError, TypeError, IndexError, AssertionError) as second:
            raise ReplayError(f"{label}: perf stream parse failed: {second}") from error
    require(samples, f"{label}: perf stream is empty")
    require(len({sample["timestamp"] for sample in samples}) == len(samples),
            f"{label}: perf timestamps are not unique")
    require(all(sample["period"] > 0 for sample in samples),
            f"{label}: non-positive sample period")
    return samples, {
        "lost_event_lines": len(lost_lines),
        "status_lines": len([line for line in text.splitlines()
                              if line.startswith("PERF_RECORD") or "lost" in line.lower()]),
        "parser_samples": len(samples),
        "raw_text_lines": len(text.splitlines()),
    }


def stack_diagnostics(data: bytes, parser: Any, label: str) -> dict[str, Any]:
    samples, status = parse_perf_stream(data, parser, label)
    qualified = []
    unresolved = 0
    unknown_interior = 0
    self_leaf_counts: Counter[str] = Counter()
    self_leaf_period: Counter[str] = Counter()
    inclusive_symbols: Counter[str] = Counter()
    inclusive_period: Counter[str] = Counter()
    inclusive_paths: Counter[tuple[str, ...]] = Counter()
    inclusive_path_period: Counter[tuple[str, ...]] = Counter()
    unresolved_frames = 0
    unknown_interior_frames = 0
    for sample in samples:
        symbols = sample["symbols"]
        unknown = [symbol for symbol in symbols
                   if any(token in symbol.lower() for token in ("[unknown]", "??", "<unknown>"))]
        if unknown:
            unresolved += 1
            unresolved_frames += len(unknown)
        if OWNER not in symbols:
            continue
        require(symbols.count(OWNER) == 1, f"{label}: exact owner appears more than once")
        owner_index = symbols.index(OWNER)
        interior = symbols[:owner_index]
        if any(any(token in symbol.lower() for token in ("[unknown]", "??", "<unknown>"))
               for symbol in interior):
            unknown_interior += 1
            unknown_interior_frames += sum(
                any(token in symbol.lower() for token in ("[unknown]", "??", "<unknown>"))
                for symbol in interior
            )
        qualified.append(sample)
        if interior:
            self_leaf_counts[interior[0]] += 1
            self_leaf_period[interior[0]] += sample["period"]
        else:
            # The owner itself is the sampled leaf in the unusual exact-owner
            # root case.  Keeping it makes the leaf partition reconstruct the
            # full exact-owner count instead of silently dropping that case.
            self_leaf_counts[OWNER] += 1
            self_leaf_period[OWNER] += sample["period"]
        # Every prefix ending at the owner is an inclusive call path.  Paths
        # overlap by design and are never summed as a cost estimate.
        prefix = tuple(symbols[:owner_index + 1])
        inclusive_paths[prefix] += 1
        inclusive_path_period[prefix] += sample["period"]
        for symbol in symbols[:owner_index + 1]:
            inclusive_symbols[symbol] += 1
            inclusive_period[symbol] += sample["period"]
    require(sum(1 for _ in qualified) <= len(samples), f"{label}: owner partition changed")
    owner_period = sum(sample["period"] for sample in qualified)
    total_period = sum(sample["period"] for sample in samples)
    top_paths = sorted(inclusive_paths.items(), key=lambda item: (-item[1], item[0]))[:50]
    top_path_period = sorted(inclusive_path_period.items(), key=lambda item: (-item[1], item[0]))[:50]
    return {
        "whole_process_samples": len(samples),
        "whole_process_period": total_period,
        "owner_qualified_samples": len(qualified),
        "owner_qualified_period": owner_period,
        "unqualified_samples": len(samples) - len(qualified),
        "unresolved_samples": unresolved,
        "unresolved_frames": unresolved_frames,
        "qualified_stacks_with_unknown_interior": unknown_interior,
        "unknown_interior_frames": unknown_interior_frames,
        "self_leaf_samples": [[name, count] for name, count in self_leaf_counts.most_common(50)],
        "self_leaf_period": [[name, value] for name, value in self_leaf_period.most_common(50)],
        "inclusive_symbol_samples": [[name, count] for name, count in inclusive_symbols.most_common(50)],
        "inclusive_symbol_period": [[name, value] for name, value in inclusive_period.most_common(50)],
        "inclusive_callpath_counts": [
            {"path": list(path), "samples": count, "period": inclusive_path_period[path]}
            for path, count in top_paths
        ],
        "inclusive_callpath_period": [
            {"path": list(path), "samples": inclusive_paths[path], "period": period}
            for path, period in top_path_period
        ],
        "count_semantics": {
            "self_leaf": "one top frame per exact-owner-qualified sample; owner is used when no interior exists",
            "inclusive_symbol": "per-symbol frame occurrence through the owner, so repeated frames are counted repeatedly",
            "inclusive_callpath": "one complete root-to-owner path per exact-owner-qualified sample; paths overlap",
        },
        "status_coverage": {
            **status,
            "parsed_stack_samples": len(samples),
            "owner_qualified_samples": len(qualified),
            "unqualified_samples": len(samples) - len(qualified),
            "coverage_percent": (100.0 * len(samples)
                                  / (len(samples) + status["lost_event_lines"])),
            "lost_events_explicit": True,
        },
    }


def find_perf_directory() -> Path:
    candidates = [PACKET / "perf", PACKET / "perf-fp", PACKET / "perf-capture"]
    path = next((candidate for candidate in candidates if candidate.is_dir()), None)
    require(path is not None, "perf capture directory is missing")
    return path


def perf_artifact_from(row: dict[str, Any], keys: tuple[str, ...], label: str) -> Any:
    for key in keys:
        if isinstance(row.get(key), dict):
            return row[key]
    artifacts = row.get("artifacts")
    if isinstance(artifacts, dict):
        for key in keys:
            value = artifacts.get(key)
            if isinstance(value, dict):
                return value
        # Match by suffix for drivers that key artifacts by filename.
        for value in artifacts.values():
            if isinstance(value, dict) and any(key in str(value.get("path", "")) for key in keys):
                return value
    raise ReplayError(f"{label}: artifact is missing")


def load_perf(plan: dict[str, Any], build: dict[str, Any],
              cleanup: dict[str, Any] | None, parser: Any) -> dict[str, Any]:
    directory = find_perf_directory()
    complete_path = directory / "complete.json"
    complete = read(complete_path)
    require(complete.get("schema") == "litchi.performance.0811.perf.complete.v1"
            and complete.get("repeats") == 2
            and complete.get("processes", complete.get("reports")) == 2
            and complete.get("reports", 2) == 2
            and complete.get("samples", 200) == 200
            and complete.get("plan_sha256") == sha(PACKET / "plan.json")
            and complete.get("build_sha256") == sha(PACKET / "build" / "build.json")
            and complete.get("event") == "cycles:u"
            and complete.get("frequency_hz") == 499
            and complete.get("call_graph") == "fp"
            and complete.get("owner") == OWNER,
            "perf completion cardinality changed")
    receipts_path = directory / "receipts.json"
    receipts = read(receipts_path)
    require(isinstance(receipts, list) and len(receipts) == 2,
            "perf receipt cardinality changed")
    compression_path = directory / "compression.json"
    compression = read(compression_path)
    require(isinstance(compression, list) and len(compression) == 4,
            "perf compression cardinality changed")
    compressed: dict[tuple[int, str], dict[str, Any]] = {}
    for item in compression:
        require(isinstance(item, dict) and item.get("compression") == "gzip"
                and item.get("gzip_mtime") == 0, "perf compression row changed")
        repeat = item.get("repeat")
        kind = item.get("kind")
        require(repeat in (0, 1) and kind in ("raw", "frames"),
                "perf compression identity changed")
        key = (repeat, kind)
        require(key not in compressed, "duplicate perf compression identity")
        original = item.get("original")
        compressed_descriptor = item.get("compressed")
        require(isinstance(original, dict) and isinstance(compressed_descriptor, dict),
                "perf compression descriptors are incomplete")
        original_path = packet_path(original.get("path"), f"perf {repeat} {kind} original")
        artifact(original, f"perf {repeat} {kind} original", allow_missing=True)
        compressed_path = artifact(compressed_descriptor,
                                   f"perf {repeat} {kind} compressed")
        try:
            with gzip.open(compressed_path, "rb") as stream:
                uncompressed = stream.read()
        except (OSError, EOFError) as error:
            raise ReplayError(f"perf {repeat} {kind}: invalid gzip") from error
        require(len(uncompressed) == original.get("bytes")
                and hashlib.sha256(uncompressed).hexdigest() == original.get("sha256")
                and item.get("decompressed_sha256") == original.get("sha256"),
                f"perf {repeat} {kind}: compressed identity changed")
        require(original_path.name == f"{repeat}.data" if kind == "raw"
                else original_path.name == f"{repeat}.frames",
                f"perf {repeat} {kind}: original filename changed")
        compressed[key] = {"original": original, "compressed": compressed_descriptor,
                           "data": uncompressed}
    require(set(compressed) == {(0, "raw"), (0, "frames"), (1, "raw"), (1, "frames")},
            "perf compression matrix incomplete")
    decode_complete = read(directory / "decode-complete.json")
    require(decode_complete.get("schema") == "litchi.performance.0811.decode.complete.v1"
            and decode_complete.get("reports") == 2
            and decode_complete.get("logical_artifacts") == 4
            and decode_complete.get("plan_sha256") == sha(PACKET / "plan.json")
            and decode_complete.get("build_sha256") == sha(PACKET / "build" / "build.json"),
            "perf decode completion changed")
    decode_receipts = read(directory / "decode-receipts.json")
    require(isinstance(decode_receipts, list) and len(decode_receipts) == 2,
            "perf decode receipt cardinality changed")
    for index, decoded in enumerate(decode_receipts):
        label = f"decode/{index}"
        require(isinstance(decoded, dict)
                and decoded.get("schema") == "litchi.performance.0811.decode-receipt.v1"
                and decoded.get("repeat") == index and decoded.get("exit_code") == 0
                and decoded.get("binary") == build["binaries"]["fp"],
                f"{label}: decode receipt changed")
        require(decoded.get("raw") == compressed[(index, "raw")]["original"],
                f"{label}: raw input changed")
        require(decoded.get("frames") == compressed[(index, "frames")]["original"],
                f"{label}: frame output changed")
        decoded_source = decoded.get("source")
        require(isinstance(decoded_source, dict), f"{label}: source receipt missing")
        decoded_source_path = artifact(decoded_source, f"{label} source")
        require(source_files(read(decoded_source_path), f"{label} source")
                == build["source"]["files"], f"{label}: source census changed")
        require(decoded.get("probe") == probe_custody()
                and decoded.get("root_inputs") == root_inputs(),
                f"{label}: custody changed")
        artifact(decoded.get("raw"), f"{label} raw", allow_missing=True)
        artifact(decoded.get("frames"), f"{label} frames", allow_missing=True)
        artifact(decoded.get("log"), f"{label} log")
        expected_decode_command = ["perf", "script", "--no-inline", "--ns", "-i",
                                   decoded["raw"]["path"]]
        require(normalize_command(decoded.get("command"))
                == normalize_command(expected_decode_command),
                f"{label}: decode command changed")
    rows: list[dict[str, Any]] = []
    for index, row in enumerate(receipts):
        label = f"perf/{index}"
        require(isinstance(row, dict)
                and row.get("schema") == "litchi.performance.0811.perf-receipt.v1"
                and row.get("lane") == "perf"
                and row.get("repeat", row.get("index")) == index
                and row.get("exit_code") == 0, f"{label}: receipt identity changed")
        binary_ref = row.get("binary")
        if isinstance(binary_ref, dict):
            require(binary_ref == build["binaries"]["fp"], f"{label}: binary identity changed")
        raw = perf_artifact_from(row, ("raw", "data"), f"{label} raw")
        report = perf_artifact_from(row, ("report",), f"{label} report")
        log = perf_artifact_from(row, ("log",), f"{label} log")
        artifact(raw, f"{label} raw", allow_missing=True)
        report_path = artifact(report, f"{label} report")
        artifact(log, f"{label} log")
        source_ref = row.get("source")
        require(isinstance(source_ref, dict), f"{label}: source receipt missing")
        source_path = artifact(source_ref, f"{label} source")
        require(source_files(read(source_path), f"{label} source") == build["source"]["files"],
                f"{label}: source census changed")
        require(row.get("probe") == probe_custody()
                and row.get("root_inputs") == root_inputs(), f"{label}: custody changed")
        require(raw == compressed[(index, "raw")]["original"],
                f"{label}: raw identity changed")
        expected_command = [
            "taskset", "-c", str(plan["perf"]["cpu"]), "perf", "record",
            "--no-buildid-cache", "-e", plan["perf"]["event"], "-F",
            str(plan["perf"]["frequency_hz"]), "--call-graph",
            plan["perf"]["call_graph"], "-o", raw["path"], "--",
            build["binaries"]["fp"]["path"], "--mode", "capture", "--shape", "large",
            "--samples", str(plan["perf"]["samples"]), "--warmup",
            str(plan["perf"]["warmup"]), "--output", report["path"],
        ]
        require(normalize_command(row.get("command")) == normalize_command(expected_command),
                f"{label}: command changed")
        report_outcome = validate_report(
            read(report_path), "large", Path(build["binaries"]["fp"]["path"]).name,
            plan["perf"]["samples"], plan["perf"]["warmup"], label,
        )
        frame_original = compressed[(index, "frames")]["original"]
        frame_descriptor = compressed[(index, "frames")]["compressed"]
        frame_path = artifact(frame_descriptor, f"{label} frames")
        require(frame_original.get("path", "").endswith(f"{index}.frames"),
                f"{label}: frame original identity changed")
        frame_data = compressed[(index, "frames")]["data"]
        stack = stack_diagnostics(frame_data, parser, label)
        rows.append({
            "repeat": index, "report": relative(report_path),
            "report_sha256": sha(report_path), "raw": raw,
            "frames": frame_descriptor, "frames_sha256": frame_descriptor.get("sha256"),
            "report_stats": report_outcome["stats"], "stack": stack,
            "command": row.get("command"),
        })
    return {
        "directory": relative(directory), "reports": rows,
        "summary": {
            "reports": len(rows),
            "samples": sum(item["report_stats"]["count"] for item in rows),
            "whole_process_samples": sum(item["stack"]["whole_process_samples"] for item in rows),
            "owner_qualified_samples": sum(item["stack"]["owner_qualified_samples"] for item in rows),
            "unresolved_samples": sum(item["stack"]["unresolved_samples"] for item in rows),
            "lost_event_lines": sum(item["stack"]["status_coverage"]["lost_event_lines"]
                                    for item in rows),
        },
    }


def analyze() -> dict[str, Any]:
    plan, sealed, live = plan_and_custody()
    cleanup = load_cleanup()
    manifest, build = load_build(cleanup)
    parser = load_hash_bound_perf_parser()
    native = load_native(plan, build, cleanup)
    summaries, spread_flags = paired_summary(native)
    perf = load_perf(plan, build, cleanup, parser)
    return {
        "schema": "litchi.performance.0811.current-production-analysis.v1",
        "plan": {"path": "plan.json", "sha256": sha(PACKET / "plan.json"),
                  "schema": plan["schema"]},
        "source": {
            "file_count": SOURCE_COUNT,
            "sealed_0810_after_files": sealed,
            "live_files": live,
            "live_matches_sealed_0810_after": True,
            "revision_equality_not_required_after_commit": True,
        },
        "root_inputs": root_inputs(),
        "probe": {
            "schema": PROBE_SCHEMA, "tool": PROBE_TOOL,
            "files": probe_custody(), "source": "sealed change-0810 probe-src",
            "full_semantic_oracle": True,
        },
        "quality": {
            "production": load_quality_reuse(),
            "probe": load_probe_quality(),
        },
        "build": {
            "manifest": build["manifest_path"],
            "binaries": build["binaries"],
            "three_binary_matrix": True,
        },
        "native": {
            "reports": len(native), "samples": sum(item["sample_count"] for item in native),
            "blocks": 6, "samples_per_report": 30, "warmup": 3,
            "processes": native, "summaries": summaries,
            "spread_flags_over_5_percent": spread_flags,
            "ratio_pairs": ["profile/control", "fp/profile"],
            "bootstrap": {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                          "statistic": "median of six paired block ratios",
                          "low_rank": BOOTSTRAP_LOW, "high_rank": BOOTSTRAP_HIGH},
        },
        "perf": perf,
        "diagnostic_boundary": {
            "phase_fraction_claim_authorized": False,
            "causal_cost_claim_authorized": False,
            "native_speedup_claim_authorized": False,
            "adoption_gate": None,
            "allocation_lane": False,
            "callgrind_lane": False,
            "exact_owner": OWNER,
            "owner_counts_are_descriptive_only": True,
            "nested_inclusive_counts_overlap": True,
            "unresolved_and_lost_status_retained": True,
        },
        "cleanup_contract": {
            "target": str(TARGET), "binary_count": 3,
            "exact_identity_witness_required": True,
            "output_stable_before_and_after_cleanup": True,
        },
        "verification": {
            "all_native_reports_checked": True,
            "all_perf_reports_checked": True,
            "all_output_and_semantic_oracles_checked": True,
            "current_source_files_checked_without_revision_equality": True,
            "hash_bound_0784_perf_parser_checked": True,
            "raw_perf_decoded_and_gzip_identity_checked": True,
            "no_historical_timing_pool": True,
            "no_cross_format_lane": True,
        },
        "limits": [
            "The native ratios characterize control, wrapper, and frame-pointer perturbations only.",
            "Whole-process and exact-owner sampled stacks are retained as attribution diagnostics; no phase fraction, causal cycle cost, latency speedup, or adoption claim is made.",
        ],
    }


def main() -> None:
    args = set(sys.argv[1:])
    require(args in ({"--write"}, {"--check"}), "use exactly --write or --check")
    value = analyze()
    path = PACKET / "analysis.json"
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    if "--write" in args:
        require(not path.exists(), "refusing to overwrite analysis.json")
        path.write_text(encoded, encoding="utf-8")
    else:
        require(path.is_file() and path.read_text(encoding="utf-8") == encoded,
                "analysis.json does not replay byte-for-byte")
    print("0811 current-production analysis PASS 56 reports/1820 samples")


if __name__ == "__main__":
    try:
        main()
    except ReplayError as error:
        print(f"analysis failed: {error}", file=sys.stderr)
        raise SystemExit(1)
