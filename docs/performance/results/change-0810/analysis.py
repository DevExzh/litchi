"""Fail-closed offline replay for the 0810 PPTX workflow packet.

The drivers own compilation and execution.  This reader only verifies retained
receipts, raw probe reports, the sealed 0806 qualification oracle, and the
frozen public adoption policy.  It deliberately has no subprocess, Cargo, or
profiler path.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import statistics
import sys
from functools import lru_cache
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
TARGET = Path("/home/zhuhe/code/litchi-target-0810")
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
MODES = ("capture", "commit", "lifecycle")
CASES = tuple((shape, mode) for shape in SHAPES for mode in MODES)
LEGS = ("before", "after")
DIMENSIONS = {
    "tiny": (3, 4), "medium": (12, 8), "large": (100, 100),
    "vendor": (12, 8), "unicode-vendor": (12, 8), "valid-4attr": (12, 8),
}
TIMING_SCOPES = {
    "capture": "Package::opened_presentation only",
    "commit": "Transaction::commit only; package capture and one set_shape_text staging are outside the clock",
    "lifecycle": "Package::opened_presentation, edit, set_shape_text, commit, apply_opened_presentation_commit, and Package::to_bytes",
}
PROBE_SCHEMA = "litchi.pptx.public-workflow-probe-0806.v1"
PROBE_TOOL = "public-pptx-probe-0806"
MARKER = "litchi-perf-0780-static-mce-capabilities"
VALID_URIS = [
    "urn:litchi:perf:0806:extension:one",
    "urn:litchi:perf:0806:extension:two",
    "urn:litchi:perf:0806:extension:three",
    "urn:litchi:perf:0806:extension:four",
]
VALID_NAMES = ["lx1:probeOne", "lx2:probeTwo", "lx3:probeThree", "lx4:probeFour"]
VALID_VALUES = [
    "litchi-perf-0806-valid-4attr-one",
    "litchi-perf-0806-valid-4attr-two",
    "litchi-perf-0806-valid-4attr-three",
    "litchi-perf-0806-valid-4attr-four",
]
RAW_ALLOC = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
ALLOC_METRICS = (*RAW_ALLOC, "net_live", "peak_above_entry")
BOOTSTRAP_SEED = 810810
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW = 250
BOOTSTRAP_HIGH = 9749
REPORT_KEYS = frozenset({
    "schema", "tool", "mode", "shape", "slides", "shapes_per_slide",
    "timing_scope", "marker", "source", "fixture", "warmup",
    "samples_requested", "samples", "allocator",
})
FIXTURE_KEYS = frozenset({
    "injection", "slide_parts", "replaced_text_tags", "namespace_declarations",
    "namespaced_attributes", "namespace_uris", "attribute_names",
})
SOURCE_KEYS = frozenset({"bytes", "sha256"})
ALLOC_KEYS = frozenset({
    "status", "scope", *RAW_ALLOC,
})
COMMON_VERIFICATION_KEYS = frozenset({
    "semantic_check", "reopened", "expected_text", "actual_text",
    "semantic_text_bytes", "semantic_text_sha256", "readback_bytes",
    "readback_sha256", "marker_matches", "unknown_namespace_check",
    "unknown_namespace_occurrences",
})
EXTENSION_VERIFICATION_KEYS = frozenset({
    "extension_preservation_check", "extension_text_tags",
    "extension_attributes_per_text_tag", "extension_attribute_occurrences",
    "extension_value_occurrences", "extension_namespace_declarations_per_slide",
    "extension_namespace_uris", "extension_attribute_names",
    "extension_attribute_values",
})


class ReplayError(RuntimeError):
    pass


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
    return isinstance(value, str) and len(value) == 64 and all(c in "0123456789abcdef" for c in value)


def number(value: Any, label: str, *, nonnegative: bool = True) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label}: not finite")
    if nonnegative:
        require(value >= 0, f"{label}: negative")


def integer(value: Any, label: str, *, positive: bool = False) -> None:
    require(isinstance(value, int) and not isinstance(value, bool)
            and (value > 0 if positive else value >= 0), f"{label}: invalid integer")


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def packet_path(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}: path missing")
    raw = Path(value)
    path = raw if raw.is_absolute() else PACKET / raw
    # Driver receipts use absolute packet paths.  Permit those, but never allow
    # a packet artifact reference to escape the packet.
    path = path.resolve()
    try:
        path.relative_to(PACKET.resolve())
    except ValueError as error:
        raise ReplayError(f"{label}: path escapes packet: {value}") from error
    return path


def artifact(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: artifact malformed")
    path = packet_path(value.get("path"), label)
    integer(value.get("bytes"), f"{label}.bytes")
    require(is_sha(value.get("sha256")), f"{label}.sha256 malformed")
    require(path.is_file() and not path.is_symlink(), f"{label}: file missing")
    require(path.stat().st_size == value["bytes"], f"{label}: byte count changed")
    require(sha(path) == value["sha256"], f"{label}: digest changed")
    return path


def external_binary(value: Any, label: str, cleanup: dict[str, Any] | None) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: binary identity malformed")
    path = Path(value.get("path", ""))
    require(path.is_absolute(), f"{label}: binary path is not absolute")
    integer(value.get("bytes"), f"{label}.bytes", positive=True)
    require(is_sha(value.get("sha256")), f"{label}.sha256 malformed")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == value["bytes"] and sha(path) == value["sha256"],
                f"{label}: live binary identity changed")
        return {"path": str(path), "bytes": value["bytes"], "sha256": value["sha256"]}
    require(cleanup is not None and cleanup.get("target_removed") is True,
            f"{label}: missing binary without cleanup witness")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list), "cleanup removed_binaries malformed")
    matches = [row for row in removed if isinstance(row, dict) and row.get("path") == str(path)]
    require(len(matches) == 1 and matches[0].get("bytes") == value["bytes"]
            and matches[0].get("sha256") == value["sha256"],
            f"{label}: exact removed binary witness missing")
    return {"path": str(path), "bytes": value["bytes"], "sha256": value["sha256"]}


def source_manifest(path: Path, label: str) -> dict[str, Any]:
    value = read(path)
    require(isinstance(value, dict) and isinstance(value.get("files"), dict),
            f"{label}: source manifest malformed")
    files = value["files"]
    require(files and all(isinstance(k, str) and is_sha(v) for k, v in files.items()),
            f"{label}: source file census malformed")
    return {"revision": value.get("revision"), "files": dict(files)}


def frozen_inputs(path: Path, label: str) -> dict[str, Any]:
    value = read(path)
    names = {
        "plan.json", "adoption-policy.json", "analysis-plan.json", "custody.py",
        "build.py", "capture.py", "profile.py", "quality.py", "probe_quality.py",
        "apply_candidate.py", "restore_candidate.py", "origin.json", "host.json",
        "inheritance.json", "architecture-inputs.json", "inputs/root-Cargo.lock",
        "inputs/rustfmt.toml",
    }
    require(isinstance(value, dict) and set(value) == {"packet", "root_inputs"}
            and isinstance(value["packet"], dict)
            and set(value["packet"]) == names
            and value["root_inputs"] == frozen_root_inputs(),
            f"{label}: frozen input schema changed")
    for name, digest in value["packet"].items():
        require(is_sha(digest), f"{label}: invalid frozen digest: {name}")
        path = PACKET / name
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{label}: frozen packet input changed: {name}")
    return value


def current_source() -> dict[str, str]:
    # The production census is deliberately derived without invoking a shell.
    import subprocess
    raw = subprocess.check_output(["git", "ls-files", "-z", "--", "crates", "Cargo.toml",
                                   "clippy.toml", ".cargo/config.toml", "rust-toolchain.toml"],
                                  cwd=ROOT)
    names = [name for name in raw.decode().split("\0") if name]
    return {name: sha(ROOT / name) for name in names}


@lru_cache(maxsize=1)
def frozen_root_inputs() -> dict[str, str]:
    """Bind the root Cargo/rustfmt inputs copied before the first build."""
    origin = read(PACKET / "origin.json")
    require(origin.get("schema") == "litchi.performance.0810.origin.v1",
            "origin schema changed")
    root_meta = origin.get("root_inputs")
    require(isinstance(root_meta, dict)
            and root_meta.get("copies_captured_before_first_build") is True
            and is_sha(root_meta.get("root-Cargo.lock"))
            and is_sha(root_meta.get("rustfmt.toml")),
            "root input custody changed")
    copies = {
        "Cargo.lock": PACKET / "inputs/root-Cargo.lock",
        "rustfmt.toml": PACKET / "inputs/rustfmt.toml",
    }
    result = {name: root_meta[f"root-{name}"] if name == "Cargo.lock"
              else root_meta[name] for name in copies}
    for name, copy in copies.items():
        require(copy.is_file() and not copy.is_symlink(),
                f"missing frozen root input: {copy}")
        require(sha(copy) == result[name], f"frozen root input changed: {name}")
        source = ROOT / name
        require(source.is_file() and sha(source) == result[name],
                f"live root input changed: {name}")
    return result


@lru_cache(maxsize=None)
def sealed_fixture(shape: str) -> dict[str, Any]:
    """Load one fixture map from the sealed 0806 qualification corpus.

    The probe fixture is independent of the production candidate.  Binding
    every replayed report to this archived map prevents a driver or probe
    change from being hidden behind the shape label alone.
    """
    old = ROOT / "docs/performance/results/change-0806"
    relative = f"qualification/0-{shape}-capture-before.json"
    path = old / relative
    seal = read(old / "seal.json")
    require(seal.get("files", {}).get(relative) == sha(path),
            f"0806 fixture seal changed: {relative}")
    report = read(path)
    fixture = report.get("fixture")
    require(isinstance(fixture, dict) and set(fixture) == FIXTURE_KEYS,
            f"0806 fixture malformed: {shape}")
    return fixture


def frozen_plan() -> tuple[dict[str, Any], dict[str, Any]]:
    frozen_root_inputs()
    origin = read(PACKET / "origin.json")
    require(origin.get("base") == "3677e31be5c9d5582a1f6d531ebb4d54db5a0acc"
            and origin.get("production_source", {}).get("revision") == origin.get("base")
            and origin.get("production_source", {}).get("tracked_file_count") == 9196
            and origin.get("production_source", {}).get("candidate_allowlist") ==
            ["crates/litchi-pptx/src/notes/codec.rs"],
            "origin source custody changed")
    plan = read(PACKET / "plan.json")
    require(plan.get("schema") == "litchi.performance.0810.v1", "plan schema changed")
    require(plan.get("cpu") == 12 and plan.get("source_allowlist") ==
            ["crates/litchi-pptx/src/notes/codec.rs"], "plan scope changed")
    require(plan.get("cases") == [{"mode": mode, "shape": shape}
            for shape in SHAPES for mode in MODES], "case order changed")
    for lane, counts in (("qualification", (1, 1, 0, 18, 18)),
                         ("native", (6, 30, 3, 216, 6480)),
                         ("allocation", (2, 3, 0, 72, 216))):
        row = plan.get(lane)
        require(isinstance(row, dict), f"{lane} plan missing")
        blocks, samples, warmup, reports, total = counts
        require((row.get("blocks"), row.get("samples"), row.get("warmup"),
                 row.get("reports"), row.get("samples_total")) == counts,
                f"{lane} plan changed")
    qualification = plan["qualification"]
    require(qualification.get("orders") == [["before"]]
            and qualification.get("binary") == "allocation",
            "qualification schedule changed")
    native = plan["native"]
    require(native["orders"] == [["before", "after"], ["after", "before"],
                                  ["before", "after"], ["after", "before"],
                                  ["after", "before"], ["before", "after"]],
            "native alternating order changed")
    allocation = plan["allocation"]
    require(allocation.get("orders") == [["before", "after"], ["after", "before"]],
            "allocation alternating order changed")
    profile = plan.get("profile")
    require(isinstance(profile, dict) and profile.get("owner") ==
            "namespace_uri_probe::capture_region_0793"
            and profile.get("binary") == "profile"
            and profile.get("mode") == "capture"
            and profile.get("shape") == "large"
            and profile.get("orders") == [["before", "after"], ["after", "before"]]
            and profile.get("repeats") == 2
            and profile.get("samples") == 1
            and profile.get("warmup") == 0
            and profile.get("collect_at_start") is False
            and profile.get("events") == ["Ir"]
            and profile.get("expected_numbered_parts") == 1
            and profile.get("reports") == 4 and profile.get("samples_total") == 4,
            "profile plan changed")
    boot = plan.get("bootstrap")
    require(boot == {"resamples": BOOTSTRAP_RESAMPLES, "seed": BOOTSTRAP_SEED,
                     "statistic": "median", "sorted_zero_based_endpoints": [BOOTSTRAP_LOW, BOOTSTRAP_HIGH]},
            "bootstrap plan changed")
    require(plan.get("totals") == {"reports": 310, "samples": 6718},
            "plan totals changed")
    policy = read(PACKET / "adoption-policy.json")
    require(policy.get("schema") == "litchi.performance.0810.adoption-policy.v1",
            "adoption policy schema changed")
    require(policy.get("latency", {}).get("seed") == BOOTSTRAP_SEED
            and policy["latency"].get("resamples") == BOOTSTRAP_RESAMPLES
            and policy["benefit"].get("minimum_improvement_percent") == 3.0
            and policy["benefit"].get("eligible_modes") == ["capture", "lifecycle"]
            and policy.get("scope") ==
            "All eighteen PPTX capture, commit, and lifecycle rows, including ordinary, vendor, unicode-vendor, and valid-4attr controls."
            and policy.get("frozen_before_build") is True
            and policy.get("useful_public_workflow_benefit_required") is True
            and policy.get("allocation_count_alone_sufficient") is False
            and policy.get("quantile") == "nearest rank ceil(n*p)-1"
            and policy.get("cross_format") ==
            "not included in this private PPTX codec experiment",
            "adoption thresholds changed")
    require(policy.get("memory") == {
        "allocation_calls_increase_allowed": 0,
        "allocated_bytes_increase_allowed": 0,
        "net_live_increase_allowed": 0,
        "peak_above_entry_increase_allowed": 0,
        "comparison": "each paired allocation block median",
    }, "memory policy changed")
    require(policy.get("benefit") == {
        "eligible_modes": ["capture", "lifecycle"],
        "at_least_one_case_required": True,
        "minimum_improvement_percent": 3.0,
        "bootstrap95_high_below": 1.0,
    } and policy.get("latency") == {
        "metric": "paired process p50",
        "maximum_ratio": 1.05,
        "bootstrap95_low_must_exceed": 1.0,
        "any_case_violation_rejects": True,
        "resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED,
    }, "latency/benefit policy changed")
    analysis_plan = read(PACKET / "analysis-plan.json")
    require(analysis_plan == {
        "schema": "litchi.performance.0810.analysis-plan.v1",
        "pairing": "same block, case, and sample index; six alternating native blocks and two alternating allocation blocks",
        "process_quantile": "nearest rank ceil(n*p)-1",
        "bootstrap": {
            "draw": "Python random.Random.randrange with replacement over six paired process ratios per draw",
            "resamples": BOOTSTRAP_RESAMPLES,
            "seed": BOOTSTRAP_SEED,
            "statistic": "median",
            "sorted_zero_based_endpoints": [BOOTSTRAP_LOW, BOOTSTRAP_HIGH],
        },
        "benefit": "At least one capture or lifecycle row must improve by at least 3 percent with bootstrap high endpoint below 1.",
        "veto": "Any of the eighteen rows with ratio above 1.05 and bootstrap low endpoint above 1 rejects adoption.",
        "resource_guard": "Compare each paired allocation block median for allocation calls, allocated bytes, net live bytes, and peak above entry; preserve all samples and report spread.",
        "profile": "Parse four owner-scoped Callgrind Ir publications for exact owner namespace_uri_probe::capture_region_0793; do not infer latency or RSS.",
    }, "analysis plan changed")
    return plan, policy


def load_build_leg(leg: str, cleanup: dict[str, Any] | None) -> dict[str, Any]:
    directory = PACKET / f"build-{leg}"
    manifest = read(directory / "build.json")
    frozen_inputs(directory / "frozen-inputs.json", f"{leg} build")
    require(manifest.get("schema") == f"litchi.performance.0810.build-{leg}.v1",
            f"{leg} build schema changed")
    source_receipt = manifest.get("source")
    source_path = packet_path(source_receipt.get("path") if isinstance(source_receipt, dict)
                              else None, f"{leg} source receipt")
    require(source_path == (directory / "source.json").resolve(),
            f"{leg} source path changed")
    artifact(source_receipt, f"{leg} source receipt")
    source = source_manifest(directory / "source.json", f"{leg} source")
    require(manifest.get("root_inputs") == frozen_root_inputs(),
            f"{leg} root input receipt changed")
    lock = manifest.get("lock")
    require(isinstance(lock, dict), f"{leg} probe lock receipt missing")
    require(artifact(lock, f"{leg} probe lock") == (PACKET / "probe-src/Cargo.lock").resolve(),
            f"{leg} probe lock path changed")
    binaries = manifest.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation", "profile"},
            f"{leg} binary matrix changed")
    verified = {name: external_binary(value, f"{leg} {name}", cleanup)
                for name, value in binaries.items()}
    probe = manifest.get("probe")
    require(isinstance(probe, dict), f"{leg} probe manifest missing")
    for name, digest in probe.items():
        path = packet_path(name, f"{leg} probe {name}")
        require(sha(path) == digest, f"{leg} probe source changed: {name}")
    rows = manifest.get("rows", manifest.get("commands"))
    require(isinstance(rows, list) and len(rows) == 3, f"{leg} build row count changed")
    seen_names = set()
    for row in rows:
        require(isinstance(row, dict) and row.get("name") in {"native", "allocation", "profile"}
                and row.get("name") not in seen_names
                and isinstance(row.get("command"), list)
                and row.get("exit_code") == 0,
                f"{leg} build failed")
        seen_names.add(row["name"])
        number(row.get("started"), f"{leg} {row['name']} started")
        number(row.get("ended"), f"{leg} {row['name']} ended")
        require(row["started"] <= row["ended"], f"{leg} {row['name']} interval changed")
        artifact(row.get("log"), f"{leg} build log")
    require(seen_names == {"native", "allocation", "profile"},
            f"{leg} build names changed")
    environment = manifest.get("environment")
    require(environment == {
        "CARGO_TARGET_DIR": str(TARGET),
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
    }, f"{leg} build environment changed")
    return {"manifest": manifest, "source": source, "binaries": verified}


def build_state(plan: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], dict[str, Any] | None]:
    cleanup_path = PACKET / ("early-stop-cleanup.json" if
                             (PACKET / "early-stop-cleanup.json").is_file()
                             else "cleanup.json")
    cleanup = read(cleanup_path) if cleanup_path.is_file() else None
    result: dict[str, Any] = {}
    for leg in LEGS:
        result[leg] = load_build_leg(leg, cleanup)
    before, after = result["before"]["source"], result["after"]["source"]
    require(before.get("revision") == "3677e31be5c9d5582a1f6d531ebb4d54db5a0acc",
            "before source revision changed")
    require(isinstance(after.get("revision"), str) and len(after["revision"]) == 40,
            "after source revision malformed")
    require(before["files"] != after["files"], "candidate build did not change source")
    changed = {name for name in before["files"].keys() | after["files"].keys()
               if before["files"].get(name) != after["files"].get(name)}
    require(changed == set(plan["source_allowlist"]), f"production source allowlist changed: {changed}")
    return result, before, after, cleanup


def fixture_check(report: dict[str, Any], shape: str, label: str) -> None:
    require((report.get("slides"), report.get("shapes_per_slide")) == DIMENSIONS[shape],
            f"{label}: dimensions changed")
    fixture = report.get("fixture")
    require(isinstance(fixture, dict) and set(fixture) == FIXTURE_KEYS,
            f"{label}: fixture fields changed")
    require(fixture == sealed_fixture(shape), f"{label}: fixture oracle changed")
    if shape == "valid-4attr":
        require(fixture.get("injection") == "valid-four-distinct-namespaced-extension-attributes"
                and fixture.get("slide_parts") == 12 and fixture.get("replaced_text_tags") == 96
                and fixture.get("namespace_declarations") == 4
                and fixture.get("namespaced_attributes") == 4
                and fixture.get("namespace_uris") == VALID_URIS
                and fixture.get("attribute_names") == VALID_NAMES,
                f"{label}: valid-4attr fixture changed")
    else:
        vendor = shape in {"vendor", "unicode-vendor"}
        expected = {"vendor": "same-length-known-uri-near-misses",
                    "unicode-vendor": "same-length-valid-utf8-unknown-uris"}.get(shape, "none")
        require(fixture.get("injection") == expected
                and fixture.get("slide_parts") == (12 if vendor else 0)
                and fixture.get("replaced_text_tags") == (96 if vendor else 0)
                and fixture.get("namespace_declarations") == (6 if vendor else 0)
                and fixture.get("namespaced_attributes") == (6 if vendor else 0),
                f"{label}: fixture injection changed")
        if vendor:
            for key in ("namespace_uris", "attribute_names"):
                values = fixture.get(key)
                require(isinstance(values, list) and len(values) == 6
                        and all(isinstance(value, str) and value for value in values)
                        and len(set(values)) == 6, f"{label}: vendor oracle changed")
        else:
            require(fixture.get("namespace_uris") == [] and fixture.get("attribute_names") == [],
                    f"{label}: ordinary fixture changed")


def check_extension(verification: dict[str, Any], label: str) -> None:
    require(verification.get("extension_preservation_check") is True
            and verification.get("extension_text_tags") == 96
            and verification.get("extension_attributes_per_text_tag") == 4
            and verification.get("extension_attribute_occurrences") == 384
            and verification.get("extension_value_occurrences") == 384
            and verification.get("extension_namespace_declarations_per_slide") == 4
            and verification.get("extension_namespace_uris") == VALID_URIS
            and verification.get("extension_attribute_names") == VALID_NAMES
            and verification.get("extension_attribute_values") == VALID_VALUES,
            f"{label}: extension preservation oracle changed")


def nearest(values: Iterable[float], percentile: float) -> float:
    ordered = sorted(values)
    require(ordered, "empty quantile vector")
    return ordered[max(1, math.ceil(len(ordered) * percentile)) - 1]


def stats(values: Iterable[float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "empty metric vector")
    for index, value in enumerate(vector):
        number(value, f"metric[{index}]")
    return {"count": len(vector), "values": vector, "min": min(vector),
            "p50": nearest(vector, .5), "mean": statistics.mean(vector),
            "p95": nearest(vector, .95), "p99": nearest(vector, .99), "max": max(vector)}


def spread(values: Iterable[float]) -> float:
    vector = [float(value) for value in values]
    require(vector, "empty spread vector")
    low, high = min(vector), max(vector)
    return 0.0 if low == high else (float("inf") if low == 0 else (high - low) * 100 / abs(low))


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(values, "empty bootstrap vector")
    rng = random.Random(BOOTSTRAP_SEED)
    draws = sorted(statistics.median(values[rng.randrange(len(values))] for _ in values)
                   for _ in range(BOOTSTRAP_RESAMPLES))
    return {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
            "statistic": "median", "confidence": .95,
            "low_rank": BOOTSTRAP_LOW, "high_rank": BOOTSTRAP_HIGH,
            "ci_low": draws[BOOTSTRAP_LOW], "ci_high": draws[BOOTSTRAP_HIGH]}


def pair_ratio(before: float, after: float) -> dict[str, Any]:
    """Describe a paired ratio without turning equal zeroes into a fake gain."""
    if before == 0:
        equal = after == 0
        return {
            "before": before, "after": after,
            "ratio": 1.0 if equal else None,
            "change_percent": 0.0 if equal else None,
            "relative_change_defined": False,
            "zero_baseline_equal": equal,
            "zero_to_nonzero": not equal,
            "over_5_percent": not equal,
        }
    ratio = after / before
    return {
        "before": before, "after": after, "ratio": ratio,
        "change_percent": (ratio - 1.0) * 100.0,
        "relative_change_defined": True,
        "zero_baseline_equal": False,
        "zero_to_nonzero": False,
        "over_5_percent": ratio > 1.05,
    }


def allocation_sample(sample: dict[str, Any], label: str) -> dict[str, int]:
    value = sample.get("allocation")
    require(isinstance(value, dict) and value.get("status") == "measured"
            and value.get("scope") == "operation_global_system_allocator",
            f"{label}: allocation scope changed")
    result: dict[str, int] = {}
    for field in RAW_ALLOC:
        integer(value.get(field), f"{label}.{field}")
        result[field] = value[field]
    require(result["live_bytes_after"] == result["live_bytes_before"]
            + result["allocated_bytes"] - result["deallocated_bytes"],
            f"{label}: allocation accounting changed")
    require(result["region_peak_live_bytes"] >= result["live_bytes_before"]
            and result["region_peak_live_bytes"] >= result["live_bytes_after"]
            and result["peak_live_bytes_after"] >= result["peak_live_bytes_before"]
            and result["peak_live_bytes_after"] >= result["region_peak_live_bytes"],
            f"{label}: allocation peak ordering changed")
    require(result["failed_allocation_calls"] == 0, f"{label}: failed allocation")
    result["net_live"] = result["live_bytes_after"] - result["live_bytes_before"]
    result["peak_above_entry"] = result["region_peak_live_bytes"] - result["live_bytes_before"]
    return result


def validate_report(report: dict[str, Any], shape: str, mode: str, leg: str,
                    samples: int, warmup: int, kind: str, binary: dict[str, Any],
                    label: str) -> dict[str, Any]:
    require(set(report) == REPORT_KEYS, f"{label}: report fields changed")
    require(report.get("schema") == PROBE_SCHEMA and report.get("tool") == PROBE_TOOL
            and report.get("marker") == MARKER and report.get("mode") == mode
            and report.get("shape") == shape and report.get("timing_scope") == TIMING_SCOPES[mode],
            f"{label}: probe identity changed")
    fixture_check(report, shape, label)
    require(report.get("warmup") == warmup and report.get("samples_requested") == samples,
            f"{label}: sample policy changed")
    source = report.get("source")
    require(isinstance(source, dict) and set(source) == SOURCE_KEYS
            and is_sha(source.get("sha256")), f"{label}: source fields changed")
    integer(source.get("bytes"), f"{label}: source bytes", positive=True)
    allocator = report.get("allocator")
    require(isinstance(allocator, dict) and set(allocator) == {
        "binary", "allocator", "instrumentation", "counter_revision"
    } and allocator.get("binary") == Path(binary["path"]).name,
            f"{label}: allocator identity changed")
    if kind == "native":
        require(allocator.get("instrumentation") == "none"
                and allocator.get("allocator") == "Rust system allocator"
                and allocator.get("counter_revision") is None, f"{label}: native allocator changed")
    else:
        require(allocator.get("instrumentation") == "system_allocator_operation_scoped"
                and allocator.get("allocator") == "CountingSystemAllocator(std::alloc::System)"
                and allocator.get("counter_revision") == "serialized_region_peak_v3",
                f"{label}: allocation allocator changed")
    rows = report.get("samples")
    require(isinstance(rows, list) and len(rows) == samples, f"{label}: sample count changed")
    elapsed: list[float] = []
    allocation: dict[str, list[float]] = {field: [] for field in ALLOC_METRICS}
    outputs: list[tuple[int, str]] = []
    expected_metrics = ({"captured_shapes_per_slide", "captured_slides", "elapsed_ns",
                         "shapes_per_slide", "slides"}
                        if mode == "capture"
                        else {"elapsed_ns", "shapes_per_slide", "slides"})
    expected_sample_keys = ({"index", "elapsed_ns", "metrics", "source_sha256",
                             "output", "verification"}
                            if kind == "native"
                            else {"index", "elapsed_ns", "metrics", "source_sha256",
                                  "output", "allocation", "verification"})
    for index, sample in enumerate(rows):
        require(isinstance(sample, dict) and set(sample) == expected_sample_keys
                and sample.get("index") == index, f"{label}: sample fields changed")
        integer(sample.get("elapsed_ns"), f"{label}: elapsed", positive=True)
        elapsed.append(sample["elapsed_ns"])
        require(sample.get("source_sha256") == source["sha256"], f"{label}: sample source changed")
        metrics = sample.get("metrics")
        require(isinstance(metrics, dict) and set(metrics) == expected_metrics
                and metrics.get("elapsed_ns") == sample["elapsed_ns"]
                and metrics.get("slides") == report["slides"]
                and metrics.get("shapes_per_slide") == report["shapes_per_slide"]
                and (mode != "capture" or (
                    metrics.get("captured_slides") == report["slides"]
                    and metrics.get("captured_shapes_per_slide") == report["shapes_per_slide"])),
                f"{label}: raw metrics changed")
        verification = sample.get("verification")
        require(isinstance(verification, dict) and verification.get("semantic_check") is True
                and verification.get("reopened") is True
                and isinstance(verification.get("expected_text"), str)
                and verification.get("expected_text") == verification.get("actual_text")
                and is_sha(verification.get("semantic_text_sha256")),
                f"{label}: semantic verification failed")
        expected_verification = COMMON_VERIFICATION_KEYS | (
            EXTENSION_VERIFICATION_KEYS if shape == "valid-4attr" else frozenset()
        )
        require(set(verification) == expected_verification,
                f"{label}: verification fields changed")
        integer(verification.get("semantic_text_bytes"), f"{label}: semantic bytes")
        if shape == "valid-4attr":
            check_extension(verification, f"{label} sample {index}")
        else:
            require(all(verification.get(field) is None for field in (
                "extension_preservation_check", "extension_text_tags",
                "extension_attributes_per_text_tag", "extension_attribute_occurrences",
                "extension_value_occurrences", "extension_namespace_declarations_per_slide",
                "extension_namespace_uris", "extension_attribute_names", "extension_attribute_values")),
                f"{label}: unexpected extension oracle")
        output = sample.get("output")
        require(isinstance(output, dict) and set(output) == {"bytes", "sha256"}
                and is_sha(output.get("sha256")), f"{label}: output fields changed")
        integer(output.get("bytes"), f"{label}: output bytes")
        require(verification.get("readback_bytes") == output["bytes"]
                and verification.get("readback_sha256") == output["sha256"], f"{label}: readback changed")
        expected_marker = mode in {"commit", "lifecycle"}
        require(verification.get("marker_matches") is (True if expected_marker else None),
                f"{label}: marker oracle changed")
        vendor = shape in {"vendor", "unicode-vendor"}
        require(verification.get("unknown_namespace_check") is (True if vendor else None)
                and verification.get("unknown_namespace_occurrences") ==
                (DIMENSIONS[shape][0] * DIMENSIONS[shape][1] if vendor else None),
                f"{label}: namespace oracle changed")
        outputs.append((output["bytes"], output["sha256"]))
        if kind == "allocation":
            require(isinstance(sample["allocation"], dict)
                    and set(sample["allocation"]) == ALLOC_KEYS,
                    f"{label}: allocation fields changed")
            values = allocation_sample(sample, f"{label} sample {index}")
            for field, value in values.items():
                allocation[field].append(value)
        else:
            require(sample.get("allocation") is None, f"{label}: native allocation present")
    require(len(set(outputs)) == 1, f"{label}: nondeterministic output")
    return {"source": source, "elapsed": elapsed, "stats": stats(elapsed),
            "allocation": None if kind == "native" else allocation, "outputs": outputs}


def jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    row = plan[lane]
    if lane == "qualification":
        # Keep qualification explicitly before-only; this avoids importing a
        # native timing order into the historical oracle lane.
        blocks, samples, warmup = 1, row["samples"], row["warmup"]
        orders = row["orders"]
    else:
        blocks, samples, warmup = row["blocks"], row["samples"], row["warmup"]
        orders = row["orders"]
    result = []
    for block in range(blocks):
        for shape, mode in CASES:
            for leg in orders[block]:
                result.append({"lane": lane, "block": block, "shape": shape, "mode": mode,
                               "leg": leg, "samples": samples, "warmup": warmup})
    return result


def expected_command(plan: dict[str, Any], job: dict[str, Any], binary: dict[str, Any],
                     report: Path, rss: Path) -> list[str]:
    return ["/usr/bin/time", "-f", "%M", "-o", str(rss), "taskset", "-c", str(plan["cpu"]),
            binary["path"], "--mode", job["mode"], "--shape", job["shape"],
            "--samples", str(job["samples"]), "--warmup", str(job["warmup"]),
            "--output", str(report)]


def normalize_command(value: Any) -> Any:
    if isinstance(value, list):
        return [normalize_command(item) for item in value]
    if isinstance(value, str):
        marker = "/change-0810/"
        if marker in value and value.startswith("/"):
            return str(PACKET / value.split(marker, 1)[1])
    return value


def load_lane(plan: dict[str, Any], lane: str, builds: dict[str, Any], cleanup: dict[str, Any] | None) -> list[dict[str, Any]]:
    directory = PACKET / lane
    complete = read(directory / "complete.json")
    expected = jobs(plan, lane)
    require(complete.get("schema") == f"litchi.performance.0810.{lane}.complete.v1",
            f"{lane}: complete schema changed")
    require(complete.get("children") == len(expected), f"{lane}: child count changed")
    require(complete.get("reports") == len(expected)
            and complete.get("samples") == len(expected) * plan[lane]["samples"]
            and complete.get("plan_sha256") == sha(PACKET / "plan.json"),
            f"{lane}: complete counts changed")
    source = artifact(complete.get("source"), f"{lane}: complete source")
    # Source manifests for a lane are required to be a byte-equivalent census,
    # even though the receipt itself may use an absolute path.
    expected_source = builds["before"]["source"] if lane == "qualification" else builds["after"]["source"]
    require(source_manifest(source, f"{lane}: complete source")["files"] == expected_source["files"],
            f"{lane}: complete source differs from build source")
    receipts = read(artifact(complete.get("receipts"), f"{lane}: receipts"))
    require(isinstance(receipts, list) and len(receipts) == len(expected), f"{lane}: receipt count changed")
    entries, seen = [], set()
    for index, (row, job) in enumerate(zip(receipts, expected)):
        label = f"{lane}/{index}"
        require(isinstance(row, dict)
                and row.get("schema") == "litchi.performance.0810.capture-receipt.v1"
                and row.get("root_inputs") == frozen_root_inputs()
                and all(row.get(key) == job[key]
                for key in ("lane", "block", "shape", "mode", "leg")), f"{label}: identity changed")
        identity = tuple(job[key] for key in ("block", "shape", "mode", "leg"))
        require(identity not in seen, f"{label}: duplicate identity")
        seen.add(identity)
        require(row.get("exit_code") == 0, f"{label}: process failed")
        number(row.get("started"), f"{label}: started")
        number(row.get("ended"), f"{label}: ended")
        require(row["started"] <= row["ended"], f"{label}: receipt interval changed")
        # Qualification is executed by the allocation-instrumented before
        # binary so its one sample carries the full semantic and allocator
        # oracle before candidate application.
        kind = "allocation" if lane in {"allocation", "qualification"} else "native"
        expected_binary = builds[job["leg"]]["binaries"][kind]
        binary = row.get("binary")
        require(isinstance(binary, dict) and binary.get("bytes") == expected_binary["bytes"]
                and binary.get("sha256") == expected_binary["sha256"], f"{label}: binary changed")
        external_binary(binary, f"{label}: binary", cleanup)
        report = artifact(row.get("report"), f"{label}: report")
        log = artifact(row.get("log"), f"{label}: log")
        rss = artifact(row.get("rss"), f"{label}: RSS")
        text = rss.read_text(encoding="utf-8").strip()
        require(text.isdigit() and int(text) > 0, f"{label}: RSS malformed")
        require(normalize_command(row.get("command")) ==
                expected_command(plan, job, expected_binary, report, rss), f"{label}: command changed")
        outcome = validate_report(read(report), job["shape"], job["mode"], job["leg"],
                                 job["samples"], job["warmup"], kind, expected_binary, label)
        entries.append({"identity": job, "report": report, "report_sha256": sha(report),
                        "log": rel(log), "rss_kib": int(text), "outcome": outcome})
    require(len(seen) == len(expected), f"{lane}: receipt identities incomplete")
    return entries


def paired(entries: list[dict[str, Any]], metrics: Iterable[str], *, allocation: bool = False) -> dict[str, Any]:
    lookup = {(e["identity"]["shape"], e["identity"]["mode"], e["identity"]["leg"], e["identity"]["block"]): e
              for e in entries}
    output = {}
    for shape, mode in CASES:
        rows = [e for e in entries if e["identity"]["shape"] == shape and e["identity"]["mode"] == mode]
        blocks = sorted({e["identity"]["block"] for e in rows})
        if not blocks:
            continue
        metric_rows = {}
        for metric in metrics:
            by_block = []
            ratios = []
            for block in blocks:
                before = lookup[(shape, mode, "before", block)]
                after = lookup[(shape, mode, "after", block)]
                if allocation:
                    left = stats(before["outcome"]["allocation"][metric])["p50"]
                    right = stats(after["outcome"]["allocation"][metric])["p50"]
                else:
                    left = before["outcome"]["stats"][metric]
                    right = after["outcome"]["stats"][metric]
                item = {"block": block, **pair_ratio(left, right)}
                by_block.append(item)
                if item["ratio"] is not None:
                    ratios.append(float(item["ratio"]))
            require(ratios, f"{shape}/{mode}/{metric}: no defined paired ratios")
            metric_rows[metric] = {"by_block": by_block, "ratio_median": statistics.median(ratios),
                                   "change_percent_median": (statistics.median(ratios) - 1) * 100,
                                   "bootstrap": bootstrap(ratios),
                                   "regression_over_5_percent": any(x["over_5_percent"] for x in by_block)}
        output[f"{shape}/{mode}"] = {"shape": shape, "mode": mode, "blocks": len(blocks),
                                      "metrics": metric_rows, "comparison": "after/before paired by block"}
    return output


def lane_analysis(entries: list[dict[str, Any]], *, allocation: bool) -> dict[str, Any]:
    metrics = ALLOC_METRICS if allocation else ("p50", "mean", "p95", "p99")
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    spread_flags = []
    for entry in entries:
        ident = entry["identity"]
        key = (ident["shape"], ident["mode"], ident["leg"])
        grouped.setdefault(key, []).append(entry)
    derived = {}
    for key, values in sorted(grouped.items()):
        if allocation:
            group = {metric: {"per_process": [
                {"block": e["identity"]["block"], "stats": stats(e["outcome"]["allocation"][metric])}
                for e in sorted(values, key=lambda x: x["identity"]["block"])],
                "repeat_p50_values": [stats(e["outcome"]["allocation"][metric])["p50"]
                                      for e in sorted(values, key=lambda x: x["identity"]["block"])],
                "spread_percent": spread(stats(e["outcome"]["allocation"][metric])["p50"] for e in values)}
                     for metric in metrics}
        else:
            group = {metric: {"processes": [e["outcome"]["stats"][metric] for e in values],
                              "spread_percent": spread(e["outcome"]["stats"][metric] for e in values)}
                     for metric in metrics}
        rss_values = [e["rss_kib"] for e in sorted(values,
                                                     key=lambda x: x["identity"]["block"])]
        group["rss_review"] = {
            "processes_kib": rss_values,
            "spread_percent": spread(rss_values),
            "review_only": True,
        }
        for metric, row in group.items():
            if row["spread_percent"] > 5 and metric != "rss_review":
                spread_flags.append({"group": list(key), "metric": metric,
                                     "spread_percent": row["spread_percent"]})
        if group["rss_review"]["spread_percent"] > 5:
            spread_flags.append({"group": list(key), "metric": "rss_kib",
                                 "spread_percent": group["rss_review"]["spread_percent"],
                                 "review_only": True})
        derived["/".join(key)] = group
    paired_values = paired(entries, metrics, allocation=allocation)
    regressions = [{"group": key, "metric": metric, "ratio_median": value["ratio_median"]}
                   for key, group in paired_values.items() for metric, value in group["metrics"].items()
                   if value["regression_over_5_percent"]]
    return {"groups": derived,
            "spread_flags_over_5_percent": spread_flags,
            "regression_flags_over_5_percent": regressions,
            "paired_by_block_before_after": paired_values,
            "allocation_is_separate_from_elapsed": allocation}


def guards(native: dict[str, Any], allocation: dict[str, Any], policy: dict[str, Any]) -> dict[str, Any]:
    latency = policy["latency"]
    benefit = policy["benefit"]
    violations, benefits, resources = [], [], []
    for key, group in sorted(native["paired_by_block_before_after"].items()):
        metric = group["metrics"]["p50"]
        if metric["ratio_median"] > latency["maximum_ratio"] and metric["bootstrap"]["ci_low"] > latency["bootstrap95_low_must_exceed"]:
            violations.append({"case": key, "ratio_median": metric["ratio_median"],
                               "bootstrap_ci_low": metric["bootstrap"]["ci_low"],
                               "bootstrap_ci_high": metric["bootstrap"]["ci_high"]})
        shape, mode = key.split("/", 1)
        if mode in benefit["eligible_modes"] and (1 - metric["ratio_median"]) * 100 >= benefit["minimum_improvement_percent"] \
                and metric["bootstrap"]["ci_high"] < benefit["bootstrap95_high_below"]:
            benefits.append({"case": key, "improvement_percent": (1 - metric["ratio_median"]) * 100,
                             "ratio_median": metric["ratio_median"],
                             "bootstrap_ci_low": metric["bootstrap"]["ci_low"],
                             "bootstrap_ci_high": metric["bootstrap"]["ci_high"]})
    for key, group in sorted(allocation["paired_by_block_before_after"].items()):
        for metric in ("allocation_calls", "allocated_bytes", "net_live", "peak_above_entry"):
            for row in group["metrics"][metric]["by_block"]:
                if row["after"] > row["before"]:
                    resources.append({"case": key, "metric": metric, "block": row["block"],
                                      "before": row["before"], "after": row["after"]})
    return {"latency_violations": violations, "resource_violations": resources,
            "eligible_benefits": benefits, "benefit_satisfied": bool(benefits),
            "latency_guard_passed": not violations, "resource_guard_passed": not resources,
            "adoption_eligible": not violations and not resources and bool(benefits),
            "policy_thresholds": {"maximum_ratio": latency["maximum_ratio"],
                                  "minimum_improvement_percent": benefit["minimum_improvement_percent"]}}


def disposition(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "disposition.json"
    if not path.is_file():
        require(current_source() == after["files"], "candidate source is not live before disposition")
        return {"status": "pending", "production_change_retained": None, "live_source_matches": "after"}
    value = read(path)
    require(value.get("status") in {"retained", "rejected"}, "disposition malformed")
    retained = value["status"] == "retained"
    require(value.get("production_change_retained") is retained, "disposition retention flag changed")
    expected = after["files"] if retained else before["files"]
    require(current_source() == expected, "live source does not match disposition")
    if not retained:
        restored = artifact(value.get("restored_source"), "restored source")
        require(source_manifest(restored, "restored source")["files"] == before["files"],
                "restored source manifest differs from before")
    return {"status": value["status"], "production_change_retained": retained,
            "live_source_matches": "after" if retained else "before"}


def historical_qualification(entries: list[dict[str, Any]]) -> dict[str, Any]:
    old = ROOT / "docs/performance/results/change-0806"
    seal = read(old / "seal.json")
    rows = []
    for entry in entries:
        ident = entry["identity"]
        path = entry["report"]
        old_path = old / "qualification" / f"0-{ident['shape']}-{ident['mode']}-before.json"
        relative = str(old_path.relative_to(old))
        require(seal["files"].get(relative) == sha(old_path), f"0806 qualification seal changed: {relative}")
        current = read(path)
        prior = read(old_path)
        require(current["source"] == prior["source"]
                and current["fixture"] == prior["fixture"]
                and current["samples"][0]["output"] == prior["samples"][0]["output"]
                and current["samples"][0]["verification"] == prior["samples"][0]["verification"],
                f"historical qualification oracle changed: {relative}")
        rows.append({"shape": ident["shape"], "mode": ident["mode"], "historical": relative,
                     "timings_imported": False})
    require(len(rows) == 18, "qualification cardinality changed")
    return {"packet": "docs/performance/results/change-0806", "rows": rows, "timings_imported": False}


def analyze() -> dict[str, Any]:
    plan, policy = frozen_plan()
    builds, before, after, cleanup = build_state(plan)
    application = read(PACKET / "application.json")
    require(isinstance(application, dict), "candidate application missing")
    candidate = application.get("source")
    require(isinstance(candidate, dict) and candidate.get("files") == after["files"],
            "candidate application source differs from after build")
    native = load_lane(plan, "native", builds, cleanup)
    allocation = load_lane(plan, "allocation", builds, cleanup)
    qualification = load_lane(plan, "qualification", builds, cleanup)
    require(len(native) == 216 and len(allocation) == 72 and len(qualification) == 18,
            "lane cardinality changed")
    require(sum(len(e["outcome"]["elapsed"]) for e in native) == 6480
            and sum(len(e["outcome"]["elapsed"]) for e in allocation) == 216
            and sum(len(e["outcome"]["elapsed"]) for e in qualification) == 18,
            "sample cardinality changed")
    native_result = lane_analysis(native, allocation=False)
    allocation_result = lane_analysis(allocation, allocation=True)
    decision = guards(native_result, allocation_result, policy)
    # Verify the live source census on every replay, but keep the numerical
    # analysis immutable when the root later records retained versus rejected
    # disposition.  The separate receipt owns that terminal decision.
    disposition(before, after)
    return {
        "schema": "litchi.performance.0810.workflow-analysis.v1",
        "plan_schema": plan["schema"],
        "counts": {"reports": 306, "samples": 6714, "native_reports": 216,
                    "allocation_reports": 72, "qualification_reports": 18},
        "source": {"before": before, "after": after,
                    "changed_files": sorted(name for name in before["files"]
                                             if before["files"].get(name) != after["files"].get(name))},
        "policy": policy,
        "historical_qualification": historical_qualification(qualification),
        "native": {"children": 216, "blocks": 6, "samples": 30, "warmup": 3,
                    "analysis": native_result,
                    "receipts": [{**e["identity"], "report": rel(e["report"]),
                                  "report_sha256": e["report_sha256"]} for e in native]},
        "allocation": {"children": 72, "blocks": 2, "samples": 3, "warmup": 0,
                        "analysis": allocation_result,
                        "receipts": [{**e["identity"], "report": rel(e["report"]),
                                      "report_sha256": e["report_sha256"]} for e in allocation]},
        "qualification": {"children": 18, "before_source_only": True,
                           "rows": [{**e["identity"], "report": rel(e["report"]),
                                     "report_sha256": e["report_sha256"]} for e in qualification]},
        "decision_guards": decision,
        "source_disposition_contract": {
            "path": "disposition.json",
            "schema": "litchi.performance.0810.disposition.v1",
            "retention_decision_external": True,
        },
        "bootstrap": {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                       "confidence": .95, "statistic": "median",
                       "low_rank": BOOTSTRAP_LOW, "high_rank": BOOTSTRAP_HIGH},
        "verification": {"all_raw_report_samples_checked": True,
                          "semantic_verification_checked": True,
                          "historical_qualification_checked_without_timings": True,
                          "allocation_guards_use_block_medians": True,
                          "source_allowlist_checked": True,
                          "no_cross_format_lane": True,
                          "profile_is_separate": True},
        "cleanup_contract": {"target": str(TARGET), "binary_count": 6,
                              "exact_witness_required": True},
        "limits": ["No historical timing pool, cross-format, heaptrack, perf, latency-profile, RSS, or causal cycle claim."],
    }


def qualification_audit(*, acceptance: bool) -> dict[str, Any]:
    """Replay the before-only qualification before candidate application.

    This path intentionally loads only ``build-before``.  It is the custody
    boundary consumed by ``apply_candidate.py`` and cannot accidentally accept
    a qualification after the candidate has been applied.
    """
    plan, _policy = frozen_plan()
    cleanup_path = PACKET / ("early-stop-cleanup.json" if
                             (PACKET / "early-stop-cleanup.json").is_file()
                             else "cleanup.json")
    cleanup = read(cleanup_path) if cleanup_path.is_file() else None
    build = load_build_leg("before", cleanup)
    if acceptance:
        require(not (PACKET / "application.json").exists()
                and not (PACKET / "build-after").exists(),
                "qualification acceptance must precede candidate application")
        require(current_source() == build["source"]["files"],
                "qualification acceptance source differs from before build")
    builds = {"before": build, "after": build}
    entries = load_lane(plan, "qualification", builds, cleanup)
    require(len(entries) == 18 and sum(len(e["outcome"]["elapsed"]) for e in entries) == 18,
            "qualification cardinality changed")
    historical = historical_qualification(entries)
    return {
        "schema": "litchi.performance.0810.qualification-audit.v1",
        "passed": True,
        "accepted_before_application": True,
        "application_absent_at_acceptance": True,
        "build_after_absent_at_acceptance": True,
        "accepted_source_revision": build["source"]["revision"],
        "accepted_source_file_count": len(build["source"]["files"]),
        "reports": 18,
        "samples": 18,
        "before_source": build["source"],
        "rows": [{"shape": e["identity"]["shape"], "mode": e["identity"]["mode"],
                  "report": rel(e["report"]), "report_sha256": e["report_sha256"]}
                 for e in entries],
        "historical_oracle": historical,
        "timings_imported": False,
        "all_semantic_oracles_checked": True,
    }


def main() -> None:
    args = set(sys.argv[1:])
    require(args in ({"--write"}, {"--check"}, {"--qualification", "--write"},
                     {"--qualification", "--check"}),
            "use --write/--check or --qualification with one of them")
    qualification = "--qualification" in args
    value = qualification_audit(acceptance="--write" in args) if qualification else analyze()
    path = PACKET / ("qualification-audit.json" if qualification else "analysis.json")
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    if "--write" in args:
        path.write_text(encoded, encoding="utf-8")
    else:
        require(path.is_file() and path.read_text(encoding="utf-8") == encoded,
                "analysis.json does not replay byte-for-byte")
    if qualification:
        print("0810 qualification audit PASS 18 reports/18 samples")
    else:
        print(f"0810 analysis PASS {value['counts']['reports']} reports/{value['counts']['samples']} samples")


if __name__ == "__main__":
    try:
        main()
    except ReplayError as error:
        print(f"analysis failed: {error}", file=sys.stderr)
        raise SystemExit(1)
