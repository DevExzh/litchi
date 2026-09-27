"""Offline replay and analysis for the 0783 PPTX phase diagnostic.

The capture is intentionally a very small, evidence-only program.  This
module is the other half of that contract: it reads the retained receipts,
checks every identity and semantic oracle, and derives the phase summaries.
It never starts Cargo, a probe, or a profiler.  In particular, it remains
usable after the owned target has been removed when ``cleanup.json`` contains
exact binary witnesses.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
HISTORICAL_PACKET = ROOT / "docs/performance/results/change-0780"

PLAN_SCHEMA = "litchi.performance.0783.v1"
REPORT_SCHEMA = "litchi.pptx.lifecycle-phase-probe.v1"
REPORT_TOOL = "pptx-phase-probe-0783"
MARKER = "litchi-perf-0780-static-mce-capabilities"
ANALYSIS_SCHEMA = "litchi-0783-lifecycle-phase-analysis-v1"
PHASE_KEYS = (
    "phase_capture_ns",
    "phase_stage_ns",
    "phase_commit_ns",
    "phase_apply_ns",
    "phase_serialize_ns",
)
BASE_METRIC_KEYS = frozenset({"elapsed_ns", "slides", "shapes_per_slide"})
PHASE_METRIC_KEYS = BASE_METRIC_KEYS | set(PHASE_KEYS)
SHAPES = ("tiny", "medium", "large")
LEGS = ("control", "phases")
SHAPE_DIMENSIONS = {"tiny": (3, 4), "medium": (12, 8), "large": (100, 100)}
TIMING_SCOPE = (
    "Package::opened_presentation, edit, set_shape_text, commit, "
    "apply_opened_presentation_commit, and Package::to_bytes; phase-timing "
    "is a diagnostic clock-boundary mode with clock overhead and no causal "
    "historical-regression claim"
)
BOOTSTRAP_SEED = 783078
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_CONFIDENCE = 0.95
HEX = frozenset("0123456789abcdefABCDEF")


class ReplayError(RuntimeError):
    """Raised when retained evidence is missing or contradictory."""


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


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


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


def origin() -> dict[str, Any]:
    value = read_json(PACKET / "origin.json")
    require(isinstance(value, dict), "origin.json is malformed")
    base = value.get("base")
    require(isinstance(base, str) and base, "origin base is missing")
    return value


def _path_candidates(raw: Path) -> list[Path]:
    """Resolve a receipt from its capture location into this packet.

    A capture may have been made from an owned worktree and copied into the
    main packet later.  Only the explicit ``change-0783`` path component is
    relocated; no broad basename search is allowed.
    """

    candidates: list[Path] = []
    if raw.is_absolute():
        candidates.append(raw)
        parts = raw.parts
        if "change-0783" in parts:
            index = parts.index("change-0783")
            candidates.append(PACKET.joinpath(*parts[index + 1:]))
        owned = origin().get("owned_worktree")
        if isinstance(owned, str) and owned:
            owned_path = Path(owned).resolve()
            try:
                candidates.append(ROOT / raw.resolve().relative_to(owned_path))
            except ValueError:
                pass
    else:
        text = str(raw).replace("\\", "/")
        prefix = "docs/performance/results/change-0783/"
        if text.startswith(prefix):
            candidates.append(PACKET / text[len(prefix):])
        candidates.extend((PACKET / raw, ROOT / raw))
    unique: list[Path] = []
    for candidate in candidates:
        if candidate not in unique:
            unique.append(candidate)
    return unique


def resolve_path(value: Any, *, packet_bound: bool = True) -> Path:
    require(isinstance(value, str) and value, f"invalid artifact path: {value!r}")
    raw = Path(value)
    candidates = _path_candidates(raw)
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
    require(path.stat().st_size == size, f"{label}.bytes changed")
    require(sha256(path) == digest, f"{label}.sha256 changed")
    return path


def artifact_path(value: Any, label: str, *, packet_bound: bool = True) -> Path:
    path = artifact(value, label, packet_bound=packet_bound)
    assert path is not None
    return path


def normalize_path_text(value: Any) -> str:
    require(isinstance(value, str), f"receipt path is not text: {value!r}")
    text = value.replace("\\", "/")
    marker = "/change-0783/"
    if marker in text:
        text = str(PACKET / text.split(marker, 1)[1])
    owned = origin().get("owned_worktree")
    if isinstance(owned, str) and owned:
        text = text.replace(str(Path(owned).resolve()), str(ROOT.resolve()))
    return text


def source_manifest(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    revision = value.get("revision")
    require(isinstance(revision, str) and revision and all(c in HEX for c in revision),
            f"{label}.revision is invalid")
    files = value.get("files")
    require(isinstance(files, dict) and files, f"{label}.files is missing")
    for name, digest in files.items():
        require(isinstance(name, str) and name and is_sha(digest),
                f"{label} contains an invalid file digest")
    return {"revision": revision, "files": dict(files)}


def probe_files() -> set[str]:
    root = PACKET / "probe-src"
    require(root.is_dir(), "probe source directory is missing")
    return {str(path.relative_to(PACKET)) for path in root.rglob("*")
            if path.is_file() and not path.is_symlink()}


def cleanup_has(cleanup: Any, receipt: dict[str, Any]) -> bool:
    expected = (receipt.get("path"), receipt.get("bytes", receipt.get("size")),
                receipt.get("sha256", receipt.get("digest")))
    if not (isinstance(expected[0], str) and isinstance(expected[1], int)
            and not isinstance(expected[1], bool) and expected[1] >= 0
            and is_sha(expected[2])):
        return False
    if isinstance(cleanup, dict):
        actual = (cleanup.get("path"), cleanup.get("bytes", cleanup.get("size")),
                  cleanup.get("sha256", cleanup.get("digest")))
        if actual == expected:
            return True
        return any(cleanup_has(value, receipt) for value in cleanup.values())
    if isinstance(cleanup, list):
        return any(cleanup_has(value, receipt) for value in cleanup)
    return False


def load_cleanup() -> tuple[Any, bool]:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None, False
    value = read_json(path)
    require(isinstance(value, dict), "cleanup.json is malformed")
    verified = value.get("verified_before_removal") is True or value.get("verified") is True
    return value, verified


def validate_binary(receipt: Any, label: str, cleanup: Any, cleanup_verified: bool) -> None:
    require(isinstance(receipt, dict), f"{label} receipt is missing")
    path = artifact(receipt, label, packet_bound=False, allow_missing=True)
    if path is not None:
        return
    require(cleanup_verified, f"{label} is missing without cleanup verification")
    require(cleanup_has(cleanup, receipt), f"{label} lacks an exact cleanup witness")


def load_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema changed")
    require(plan.get("cpu") == 12, "capture CPU changed")
    require(plan.get("shapes") == list(SHAPES), "shape order changed")
    require(plan.get("samples") == 30 and plan.get("warmup") == 3,
            "sample configuration changed")
    require(plan.get("phase_keys") == list(PHASE_KEYS), "phase key order changed")
    expected_orders = [
        ["control", "phases"], ["phases", "control"],
        ["control", "phases"], ["phases", "control"],
        ["phases", "control"], ["control", "phases"],
    ]
    require(plan.get("orders") == expected_orders, "alternating order changed")
    bootstrap = plan.get("bootstrap")
    require(isinstance(bootstrap, dict)
            and bootstrap.get("seed") == BOOTSTRAP_SEED
            and bootstrap.get("resamples") == BOOTSTRAP_RESAMPLES,
            "bootstrap configuration changed")
    return plan


def load_build(plan: dict[str, Any], cleanup: Any, cleanup_verified: bool) -> dict[str, Any]:
    directory = PACKET / "build"
    build_path = directory / "build.json"
    build = read_json(build_path)
    require(isinstance(build, dict), "build manifest is malformed")
    source_path = artifact_path(build.get("source"), "build source")
    source = source_manifest(read_json(source_path), "build source")
    require(source["revision"] == origin()["base"], "build source revision changed")
    inventory = build.get("probe")
    require(isinstance(inventory, dict) and set(inventory) == probe_files(),
            "probe inventory changed")
    for name, digest in inventory.items():
        require(is_sha(digest), f"probe digest is invalid: {name}")
        path = PACKET / name
        require(path.is_file() and not path.is_symlink(), f"missing probe file: {name}")
        require(sha256(path) == digest, f"probe file changed: {name}")
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == set(LEGS),
            "binary map changed")
    for leg in LEGS:
        validate_binary(binaries[leg], f"{leg} binary", cleanup, cleanup_verified)
    rows = build.get("commands")
    require(isinstance(rows, list) and len(rows) == 2, "build command cardinality changed")
    manifest = normalize_path_text(str(PACKET / "probe-src/Cargo.toml"))
    expected = [
        ["cargo", "build", "--offline", "--locked", "--release", "--manifest-path", manifest],
        ["cargo", "build", "--offline", "--locked", "--release", "--manifest-path", manifest,
         "--features", "phase-timing"],
    ]
    for row, expected_command in zip(rows, expected):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                "build command failed")
        command = row.get("command")
        require(isinstance(command, list), "build command is malformed")
        require([normalize_path_text(item) if isinstance(item, str) else item for item in command]
                == expected_command, "build command changed")
        artifact_path(row.get("log"), "build command log")
    environment = build.get("environment")
    require(isinstance(environment, dict)
            and environment.get("CARGO_BUILD_JOBS") == "2"
            and environment.get("CARGO_INCREMENTAL") == "0",
            "build environment changed")
    return {"manifest": build, "path": build_path, "source": source,
            "source_path": source_path, "probe": inventory, "binaries": binaries}


def source_text(slide: int, shape: int) -> str:
    return f"litchi-perf-baseline-pptx-semantic-v1-source-{slide:03}-{shape:03}"


def semantic_text(shape: str) -> str:
    slides, boxes = SHAPE_DIMENSIONS[shape]
    values = []
    for slide in range(slides):
        for index in range(boxes):
            values.append(MARKER if slide == 0 and index == 0
                          else source_text(slide, index))
    return "\n".join(values)


def expected_output(shape: str) -> dict[str, Any]:
    text = semantic_text(shape).encode("utf-8")
    return {"semantic_text_bytes": len(text),
            "semantic_text_sha256": hashlib.sha256(text).hexdigest()}


def _historical_file(relative: str, seal: dict[str, Any]) -> Path:
    path = HISTORICAL_PACKET / relative
    require(path.is_file() and not path.is_symlink(), f"missing historical evidence: {relative}")
    expected = seal.get("files", {}).get(relative)
    require(is_sha(expected), f"historical seal lacks {relative}")
    require(sha256(path) == expected, f"historical evidence changed: {relative}")
    return path


def validate_historical_report(path: Path, shape: str) -> dict[str, Any]:
    report = read_json(path)
    label = f"historical {rel(path)}"
    require(report.get("schema") == "litchi.pptx.static-mce-capabilities-probe.v1",
            f"{label} schema changed")
    require(report.get("tool") == "mce-capabilities-probe-0780", f"{label} tool changed")
    require(report.get("mode") == "lifecycle" and report.get("shape") == shape,
            f"{label} identity changed")
    require(report.get("marker") == MARKER, f"{label} marker changed")
    require((report.get("slides"), report.get("shapes_per_slide")) == SHAPE_DIMENSIONS[shape],
            f"{label} dimensions changed")
    require(report.get("warmup") == 3 and report.get("samples_requested") == 30,
            f"{label} sample configuration changed")
    source = report.get("source")
    require(isinstance(source, dict) and is_sha(source.get("sha256")),
            f"{label} source identity missing")
    positive_int(source.get("bytes"), f"{label} source bytes")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 30, f"{label} samples changed")
    outputs: list[tuple[int, str]] = []
    expected = expected_output(shape)
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{label} sample {index} index changed")
        require(sample.get("source_sha256") == source["sha256"],
                f"{label} sample {index} source changed")
        output = sample.get("output")
        require(isinstance(output, dict) and is_sha(output.get("sha256")),
                f"{label} sample {index} output missing")
        nonnegative_int(output.get("bytes"), f"{label} sample {index} output bytes")
        outputs.append((output["bytes"], output["sha256"]))
        verification = sample.get("verification")
        require(isinstance(verification, dict)
                and verification.get("semantic_check") is True
                and verification.get("reopened") is True
                and verification.get("expected_text") == MARKER
                and verification.get("actual_text") == MARKER
                and verification.get("marker_matches") is True,
                f"{label} sample {index} semantic verification changed")
        require(verification.get("semantic_text_bytes") == expected["semantic_text_bytes"]
                and verification.get("semantic_text_sha256") == expected["semantic_text_sha256"]
                and verification.get("readback_bytes") == output["bytes"]
                and verification.get("readback_sha256") == output["sha256"],
                f"{label} sample {index} readback identity changed")
    require(len(set(outputs)) == 1, f"{label} output is not deterministic")
    return {"source": {"bytes": source["bytes"], "sha256": source["sha256"]},
            "output": {"bytes": outputs[0][0], "sha256": outputs[0][1]},
            "semantic": expected}


def load_historical() -> dict[str, Any]:
    seal = read_json(HISTORICAL_PACKET / "seal.json")
    require(isinstance(seal, dict) and isinstance(seal.get("files"), dict),
            "historical 0780 seal is malformed")
    result: dict[str, Any] = {}
    receipts: list[dict[str, Any]] = []
    for shape in SHAPES:
        identities: list[dict[str, Any]] = []
        for block in range(6):
            for leg in ("before", "after"):
                relative = f"native/{block}-{shape}-lifecycle-{leg}.json"
                path = _historical_file(relative, seal)
                identity = validate_historical_report(path, shape)
                identities.append(identity)
                receipts.append({"shape": shape, "block": block, "leg": leg,
                                 "report": relative, "report_sha256": sha256(path)})
        require(len({(item["source"]["bytes"], item["source"]["sha256"])
                      for item in identities}) == 1,
                f"historical {shape} source is not stable")
        require(len({(item["output"]["bytes"], item["output"]["sha256"])
                      for item in identities}) == 1,
                f"historical {shape} output is not stable")
        result[shape] = identities[0]
    return {"source_packet": "../change-0780", "seal_sha256": sha256(HISTORICAL_PACKET / "seal.json"),
            "by_shape": result, "lifecycle_reports": receipts}


def nearest_rank(values: Iterable[int | float], percentile: int) -> int | float:
    vector = sorted(values)
    require(vector, "empty metric vector")
    rank = max(1, math.ceil(percentile * len(vector) / 100))
    return vector[rank - 1]


def stats(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "empty metric vector")
    for index, value in enumerate(vector):
        finite_number(value, f"metric[{index}]")
    return {"count": len(vector), "min": min(vector), "p50": nearest_rank(vector, 50),
            "mean": statistics.mean(vector), "p95": nearest_rank(vector, 95),
            "p99": nearest_rank(vector, 99), "max": max(vector)}


def spread(values: Iterable[int | float]) -> float:
    vector = [float(value) for value in values]
    require(vector and all(math.isfinite(value) for value in vector), "spread vector invalid")
    low, high = min(vector), max(vector)
    if low == 0:
        return 0.0 if high == 0 else float("inf")
    return (high - low) * 100.0 / abs(low)


def distribution(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    result = stats(vector)
    result["values"] = vector
    result["spread_percent"] = spread(vector)
    result["flag_over_5_percent"] = result["spread_percent"] > 5.0
    return result


def pair_ratio(control: float, phase: float) -> dict[str, Any]:
    if control == 0:
        equal_zero = phase == 0
        return {"control": control, "phases": phase,
                "ratio": 1.0 if equal_zero else None,
                "change_percent": 0.0 if equal_zero else None,
                "relative_change_defined": False,
                "zero_baseline_equal": equal_zero,
                "zero_to_nonzero": not equal_zero,
                "over_5_percent": not equal_zero}
    ratio = phase / control
    return {"control": control, "phases": phase, "ratio": ratio,
            "change_percent": (ratio - 1.0) * 100.0,
            "relative_change_defined": True, "zero_baseline_equal": False,
            "zero_to_nonzero": False, "over_5_percent": (ratio - 1.0) * 100.0 > 5.0}


def bootstrap_ci(values: list[float]) -> dict[str, Any]:
    require(values, "cannot bootstrap an empty ratio vector")
    rng = random.Random(BOOTSTRAP_SEED)
    medians: list[float] = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        sample = [values[rng.randrange(len(values))] for _ in values]
        medians.append(statistics.median(sample))
    medians.sort()
    low = max(0, math.floor((1.0 - BOOTSTRAP_CONFIDENCE) / 2.0 * len(medians)))
    high = min(len(medians) - 1,
               math.ceil((1.0 + BOOTSTRAP_CONFIDENCE) / 2.0 * len(medians)) - 1)
    return {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
            "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median",
            "ci_low": medians[low], "ci_high": medians[high]}


def validate_report(report: dict[str, Any], shape: str, leg: str,
                    historical: dict[str, Any], binary: dict[str, Any],
                    label: str) -> dict[str, Any]:
    require(report.get("schema") == REPORT_SCHEMA, f"{label} schema changed")
    require(report.get("tool") == REPORT_TOOL, f"{label} tool changed")
    require(report.get("mode") == "lifecycle" and report.get("shape") == shape,
            f"{label} shape/mode changed")
    require(report.get("timing_scope") == TIMING_SCOPE, f"{label} timing scope changed")
    require(report.get("marker") == MARKER, f"{label} marker changed")
    require((report.get("slides"), report.get("shapes_per_slide"))
            == SHAPE_DIMENSIONS[shape], f"{label} dimensions changed")
    require(report.get("warmup") == 3 and report.get("samples_requested") == 30,
            f"{label} sample configuration changed")
    source = report.get("source")
    require(isinstance(source, dict) and source.get("bytes") == historical["source"]["bytes"]
            and source.get("sha256") == historical["source"]["sha256"],
            f"{label} source identity differs from historical 0780 lifecycle")
    allocator = report.get("allocator")
    require(isinstance(allocator, dict), f"{label} allocator identity missing")
    require(allocator.get("binary") == Path(str(binary["path"])).name
            and allocator.get("allocator") == "Rust system allocator"
            and allocator.get("instrumentation") == "none"
            and allocator.get("counter_revision") is None,
            f"{label} allocator identity changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 30, f"{label} sample count changed")
    elapsed: list[int] = []
    phases: dict[str, list[int]] = {key: [] for key in PHASE_KEYS}
    fractions: dict[str, list[float]] = {key: [] for key in PHASE_KEYS}
    outputs: list[tuple[int, str]] = []
    expected_semantic = expected_output(shape)
    for index, sample in enumerate(samples):
        sample_label = f"{label} sample {index}"
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{sample_label} index changed")
        total = sample.get("elapsed_ns")
        positive_int(total, f"{sample_label} elapsed_ns")
        elapsed.append(total)
        metrics = sample.get("metrics")
        require(isinstance(metrics, dict) and metrics.get("elapsed_ns") == total
                and metrics.get("slides") == report["slides"]
                and metrics.get("shapes_per_slide") == report["shapes_per_slide"],
                f"{sample_label} raw metrics changed")
        keys = set(metrics)
        require(keys == (PHASE_METRIC_KEYS if leg == "phases" else BASE_METRIC_KEYS),
                f"{sample_label} phase metric presence changed")
        require(sample.get("source_sha256") == source["sha256"],
                f"{sample_label} source digest changed")
        verification = sample.get("verification")
        require(isinstance(verification, dict)
                and verification.get("semantic_check") is True
                and verification.get("reopened") is True
                and verification.get("expected_text") == MARKER
                and verification.get("actual_text") == MARKER
                and verification.get("marker_matches") is True,
                f"{sample_label} semantic verification changed")
        output = sample.get("output")
        require(isinstance(output, dict) and is_sha(output.get("sha256")),
                f"{sample_label} output identity missing")
        nonnegative_int(output.get("bytes"), f"{sample_label} output bytes")
        require((output["bytes"], output["sha256"]) ==
                (historical["output"]["bytes"], historical["output"]["sha256"]),
                f"{sample_label} output differs from historical 0780 lifecycle")
        outputs.append((output["bytes"], output["sha256"]))
        require(verification.get("semantic_text_bytes") == expected_semantic["semantic_text_bytes"]
                and verification.get("semantic_text_sha256") == expected_semantic["semantic_text_sha256"]
                and verification.get("readback_bytes") == output["bytes"]
                and verification.get("readback_sha256") == output["sha256"],
                f"{sample_label} readback identity changed")
        if leg == "phases":
            total_phase = 0
            for key in PHASE_KEYS:
                value = metrics.get(key)
                nonnegative_int(value, f"{sample_label} {key}")
                phases[key].append(value)
                total_phase += value
                fractions[key].append(value / total)
            require(total_phase == total, f"{sample_label} phase sum does not equal elapsed")
    require(len(set(outputs)) == 1, f"{label} output is not deterministic")
    result: dict[str, Any] = {"elapsed": elapsed, "stats": stats(elapsed),
                              "phase_stats": None, "phase_fraction_medians": None}
    if leg == "phases":
        result["phase_stats"] = {key: stats(values) for key, values in phases.items()}
        result["phase_fraction_medians"] = {
            key: statistics.median(values) for key, values in fractions.items()
        }
    return result


def expected_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    jobs: list[dict[str, Any]] = []
    for block, order in enumerate(plan["orders"]):
        for shape in plan["shapes"]:
            for leg in order:
                jobs.append({"block": block, "shape": shape, "leg": leg,
                             "samples": plan["samples"], "warmup": plan["warmup"]})
    return jobs


def load_native(plan: dict[str, Any], build: dict[str, Any],
                historical: dict[str, Any], cleanup: Any,
                cleanup_verified: bool) -> list[dict[str, Any]]:
    directory = PACKET / "native"
    require(directory.is_dir(), "native capture directory is missing")
    complete = read_json(directory / "complete.json")
    require(isinstance(complete, dict) and complete.get("processes") == 36,
            "native process cardinality changed")
    require(complete.get("plan_sha256") == sha256(PACKET / "plan.json"),
            "native plan identity changed")
    require(complete.get("build_sha256") == sha256(build["path"]),
            "native build identity changed")
    receipt_path = directory / "receipts.json"
    rows = read_json(receipt_path)
    jobs = expected_jobs(plan)
    require(isinstance(rows, list) and len(rows) == len(jobs) == 36,
            "native receipt cardinality changed")
    entries: list[dict[str, Any]] = []
    seen: set[tuple[int, str, str]] = set()
    for index, (row, job) in enumerate(zip(rows, jobs)):
        label = f"native child {index}"
        require(isinstance(row, dict), f"{label} is not an object")
        require(row.get("block") == job["block"] and row.get("shape") == job["shape"]
                and row.get("leg") == job["leg"], f"{label} order identity changed")
        identity = (job["block"], job["shape"], job["leg"])
        require(identity not in seen, f"duplicate native identity: {identity}")
        seen.add(identity)
        require(row.get("exit_code") == 0, f"{label} failed")
        binary = build["binaries"][job["leg"]]
        validate_binary(binary, f"{label} binary", cleanup, cleanup_verified)
        report_path = artifact_path(row.get("report"), f"{label} report")
        log_path = artifact_path(row.get("log"), f"{label} log")
        rss_path = artifact_path(row.get("rss"), f"{label} RSS")
        rss_text = rss_path.read_text().strip()
        require(rss_text.isdigit(), f"{label} RSS is not an integer")
        rss = int(rss_text)
        positive_int(rss, f"{label} RSS")
        expected_command = [
            "/usr/bin/time", "-f", "%M", "-o", normalize_path_text(str(rss_path)),
            "taskset", "-c", str(plan["cpu"]), normalize_path_text(binary["path"]),
            "--mode", "lifecycle", "--shape", job["shape"],
            "--samples", str(job["samples"]), "--warmup", str(job["warmup"]),
            "--output", normalize_path_text(str(report_path)),
        ]
        command = row.get("command")
        require(isinstance(command, list), f"{label} command is malformed")
        normalized = [normalize_path_text(item) if isinstance(item, str) else item
                      for item in command]
        require(normalized == expected_command, f"{label} command changed")
        report = read_json(report_path)
        identity = historical["by_shape"][job["shape"]]
        outcome = validate_report(report, job["shape"], job["leg"], identity,
                                  binary, label)
        entries.append({"identity": job, "report": report,
                        "report_path": report_path, "report_sha256": sha256(report_path),
                        "report_bytes": report_path.stat().st_size,
                        "log": rel(log_path), "rss_kib": rss,
                        "source_sha256": report["source"]["sha256"],
                        "stats": outcome["stats"],
                        "phase_stats": outcome["phase_stats"],
                        "phase_fraction_medians": outcome["phase_fraction_medians"]})
    require(len(seen) == 36, "native identities are incomplete")
    return entries


def phase_summary(entries: list[dict[str, Any]]) -> dict[str, Any]:
    grouped: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for entry in entries:
        identity = entry["identity"]
        grouped.setdefault((identity["shape"], identity["leg"]), []).append(entry)
    groups: dict[str, Any] = {}
    spread_flags: list[dict[str, Any]] = []
    for (shape, leg), items in sorted(grouped.items()):
        items = sorted(items, key=lambda item: item["identity"]["block"])
        total_p50 = [item["stats"]["p50"] for item in items]
        rss_values = [item["rss_kib"] for item in items]
        total_distribution = distribution(total_p50)
        rss_distribution = distribution(rss_values)
        if total_distribution["flag_over_5_percent"]:
            spread_flags.append({"group": f"{shape}/{leg}", "metric": "total_p50_ns",
                                 "spread_percent": total_distribution["spread_percent"]})
        if rss_distribution["flag_over_5_percent"]:
            spread_flags.append({"group": f"{shape}/{leg}", "metric": "rss_p50_kib",
                                 "spread_percent": rss_distribution["spread_percent"]})
        process_rows = []
        summary: dict[str, Any] = {
            "process_count": len(items),
            "total_p50_ns": total_distribution,
            "rss_p50_kib": rss_distribution,
            "summary_median": {
                "total_p50_ns": statistics.median(total_p50),
                "rss_p50_kib": statistics.median(rss_values),
            },
        }
        if leg == "phases":
            phase_p50: dict[str, list[int | float]] = {key: [] for key in PHASE_KEYS}
            fractions: dict[str, list[int | float]] = {key: [] for key in PHASE_KEYS}
            for item in items:
                for key in PHASE_KEYS:
                    phase_p50[key].append(item["phase_stats"][key]["p50"])
                    fractions[key].append(item["phase_fraction_medians"][key])
                process_rows.append({
                    "block": item["identity"]["block"],
                    "total_p50_ns": item["stats"]["p50"],
                    "phase_p50_ns": {key: item["phase_stats"][key]["p50"] for key in PHASE_KEYS},
                    "phase_fraction_medians": dict(item["phase_fraction_medians"]),
                    "rss_kib": item["rss_kib"],
                    "report": rel(item["report_path"]),
                    "report_sha256": item["report_sha256"],
                })
            summary["phase_p50_ns"] = {key: distribution(values) for key, values in phase_p50.items()}
            summary["phase_fraction_medians"] = {
                key: distribution(values) for key, values in fractions.items()
            }
            summary["summary_median"]["phase_p50_ns"] = {
                key: statistics.median(values) for key, values in phase_p50.items()
            }
            summary["summary_median"]["phase_fraction_medians"] = {
                key: statistics.median(values) for key, values in fractions.items()
            }
            for metric, values in phase_p50.items():
                result = summary["phase_p50_ns"][metric]
                if result["flag_over_5_percent"]:
                    spread_flags.append({"group": f"{shape}/{leg}",
                                         "metric": f"{metric}_p50",
                                         "spread_percent": result["spread_percent"]})
        else:
            require(all(item["phase_stats"] is None for item in items),
                    f"{shape}/{leg} unexpectedly has phase analysis")
            process_rows = [{
                "block": item["identity"]["block"],
                "total_p50_ns": item["stats"]["p50"],
                "rss_kib": item["rss_kib"],
                "report": rel(item["report_path"]),
                "report_sha256": item["report_sha256"],
            } for item in items]
        summary["rss_is_separate_whole_process_gauge"] = True
        groups[f"{shape}/{leg}"] = {
            "shape": shape, "leg": leg, "processes": process_rows,
            "summary": summary,
        }
    return {"groups": groups, "spread_flags_over_5_percent": spread_flags,
            "rss_is_separate_from_elapsed": True,
            "phase_shares_are_process_medians": True}


def paired_total(entries: list[dict[str, Any]]) -> dict[str, Any]:
    by_key = {(item["identity"]["shape"], item["identity"]["leg"],
               item["identity"]["block"]): item for item in entries}
    result: dict[str, Any] = {}
    for shape in SHAPES:
        rows: list[dict[str, Any]] = []
        ratios: list[float] = []
        for block in range(6):
            control = by_key[(shape, "control", block)]["stats"]["p50"]
            phases = by_key[(shape, "phases", block)]["stats"]["p50"]
            pair = pair_ratio(float(control), float(phases))
            pair["block"] = block
            rows.append(pair)
            if pair["ratio"] is not None:
                ratios.append(pair["ratio"])
        require(len(ratios) == 6, f"{shape} paired ratio is incomplete")
        ratio_median = statistics.median(ratios)
        result[shape] = {
            "shape": shape, "blocks": 6,
            "total_p50_ns": {
                "by_block": rows, "ratios": ratios,
                "ratio_median": ratio_median,
                "change_percent_median": (ratio_median - 1.0) * 100.0,
                "spread_percent": spread(ratios),
                "flag_over_5_percent": spread(ratios) > 5.0,
                "bootstrap": bootstrap_ci(ratios),
                "material_slowdown_guard": ratio_median > 1.05 and bootstrap_ci(ratios)["ci_low"] > 1.0,
            },
            "comparison": "phases/control paired by alternating capture block",
        }
    return result


def analysis_result(plan: dict[str, Any], build: dict[str, Any],
                    historical: dict[str, Any], entries: list[dict[str, Any]]) -> dict[str, Any]:
    source_digest = {entry["identity"]["shape"]: entry["source_sha256"] for entry in entries}
    require(len(source_digest) == len(SHAPES), "source digest shape coverage changed")
    for entry in entries:
        require(entry["source_sha256"] == source_digest[entry["identity"]["shape"]],
                "source digest changed across processes")
    return {
        "schema": ANALYSIS_SCHEMA,
        "plan_schema": plan["schema"],
        "build": {
            "source": build["source"],
            "source_file": rel(build["source_path"]),
            "probe_files": build["probe"],
            "binaries": {leg: {"path": normalize_path_text(build["binaries"][leg]["path"]),
                               "bytes": build["binaries"][leg]["bytes"],
                               "sha256": build["binaries"][leg]["sha256"]}
                         for leg in LEGS},
            "build_manifest": rel(build["path"]),
        },
        "historical_0780": historical,
        "native": {
            "children": len(entries), "blocks": 6, "samples": 30, "warmup": 3,
            "analysis": phase_summary(entries),
            "receipts": [{"shape": item["identity"]["shape"],
                          "block": item["identity"]["block"],
                          "leg": item["identity"]["leg"],
                          "report": rel(item["report_path"]),
                          "report_sha256": item["report_sha256"],
                          "report_bytes": item["report_bytes"]} for item in entries],
            "paired_phase_control": paired_total(entries),
        },
        "bootstrap": {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                       "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median"},
        "verification": {
            "exact_36_receipt_order_and_cardinality_checked": True,
            "build_commands_source_and_binary_custody_checked": True,
            "report_file_hashes_and_bytes_checked": True,
            "full_30_samples_and_3_warmups_checked": True,
            "schema_tool_timing_scope_checked": True,
            "historical_0780_source_output_parity_checked": True,
            "deterministic_output_checked": True,
            "phase_sum_equals_elapsed_for_every_sample": True,
            "control_has_no_phase_keys": True,
            "nearest_rank_process_p50_recomputed": True,
            "phase_fraction_median_per_process_recomputed": True,
            "six_process_summary_medians_recomputed": True,
            "paired_phase_control_bootstrap_recomputed": True,
            "rss_separate_whole_process_gauge": True,
            "cleanup_binary_witness_required_when_missing": True,
        },
        "limits": [
            "Phase durations are descriptive clock-boundary diagnostics; they are not CPU profiles or causal attribution of the historical 0780 regression.",
            "RSS is a whole-process /usr/bin/time peak including setup, readback and destruction outside the lifecycle clock.",
            "The paired ratio compares phase-instrumented and feature-off probe processes; it does not claim a production speedup or a cold-cache, range-scaling, concurrent, or native Office result.",
        ],
    }


def analyze() -> dict[str, Any]:
    plan = load_plan()
    cleanup, cleanup_verified = load_cleanup()
    build = load_build(plan, cleanup, cleanup_verified)
    historical = load_historical()
    entries = load_native(plan, build, historical, cleanup, cleanup_verified)
    require(len(entries) == 36, "native entry cardinality changed")
    return analysis_result(plan, build, historical, entries)


def render_table(result: dict[str, Any]) -> str:
    groups = result["native"]["analysis"]["groups"]
    lines = [
        "# 0783 lifecycle phase summary",
        "",
        "Process p50 values use nearest-rank p50 over 30 measured samples; phase shares are medians within each process and then medians across six alternating processes.",
        "",
        "| Shape | Leg | Total p50 (ns) | Capture | Stage | Commit | Apply | Serialize | RSS p50 (KiB) | Total spread >5% | RSS spread >5% |",
        "| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- | --- |",
    ]
    for shape in SHAPES:
        for leg in LEGS:
            group = groups[f"{shape}/{leg}"]["summary"]
            fractions = group.get("summary_median", {}).get("phase_fraction_medians", {})
            cells = [shape, leg,
                     f"{group['summary_median']['total_p50_ns']:.3f}"]
            cells.extend(f"{100.0 * fractions[key]:.3f}%" if key in fractions else "—" for key in PHASE_KEYS)
            cells.extend([
                f"{group['summary_median']['rss_p50_kib']:.3f}",
                str(group["total_p50_ns"]["flag_over_5_percent"]),
                str(group["rss_p50_kib"]["flag_over_5_percent"]),
            ])
            lines.append("| " + " | ".join(cells) + " |")
    lines.extend([
        "",
        "| Shape | Phases/control total p50 ratio | Change | 95% bootstrap CI | Ratio spread >5% |",
        "| --- | ---: | ---: | --- | --- |",
    ])
    paired = result["native"]["paired_phase_control"]
    for shape in SHAPES:
        value = paired[shape]["total_p50_ns"]
        boot = value["bootstrap"]
        lines.append("| " + " | ".join([
            shape, f"{value['ratio_median']:.6f}",
            f"{value['change_percent_median']:.3f}%",
            f"[{boot['ci_low']:.6f}, {boot['ci_high']:.6f}]",
            str(value["flag_over_5_percent"]),
        ]) + " |")
    lines.append("")
    return "\n".join(lines)


def json_bytes(value: Any) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()


def check_bytes(path: Path, expected: bytes, label: str) -> None:
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    actual = path.read_bytes()
    require(actual == expected, f"{label} bytes do not replay")


def main() -> None:
    args = set(sys.argv[1:])
    require(args <= {"--write", "--check"}, "usage: analyze.py --write | --check")
    require(len(args) == 1, "usage: analyze.py --write | --check")
    result = analyze()
    analysis_path = PACKET / "analysis.json"
    table_path = PACKET / "phase-summary.md"
    expected_analysis = json_bytes(result)
    expected_table = render_table(result).encode()
    if "--check" in args:
        check_bytes(analysis_path, expected_analysis, "analysis.json")
        check_bytes(table_path, expected_table, "phase-summary.md")
        print("0783 analysis and phase table replay PASS")
    else:
        analysis_path.write_bytes(expected_analysis)
        table_path.write_bytes(expected_table)
        print("Wrote analysis.json and phase-summary.md")


if __name__ == "__main__":
    try:
        main()
    except ReplayError as error:
        print(f"analyze.py: {error}", file=sys.stderr)
        raise SystemExit(1)
