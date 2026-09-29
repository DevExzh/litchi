"""Fail-closed offline readers for the 0841 PPT diagnosis packet.

The probe writes the JSON reports and Callgrind writes its raw profile.  This
module never starts a probe, imports the production crate, or treats a probe
reported digest as an independent digest.  The ten input identities are
anchored to the sealed 0782 qualification reports; output identities are
checked against those same retained reports.  CFB parsing is delegated to the
sealed 0840 parser after checking its file digest, which keeps the OLE2 walk
in one independently reviewed implementation.
"""

from __future__ import annotations

import functools
import hashlib
import importlib.util
import json
import math
import random
import re
import statistics
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


PACKET = Path(__file__).resolve().parent
REPORT_SCHEMA = "litchi.ppt.profile-probe-0841.v1"
REPORT_TOOL = "ppt-profile-probe-0841"
GENERATOR_SCHEMA = "litchi.ppt.borrowed-text-probe.v1"
GENERATOR_TOOL = "ppt-borrowed-text-probe-0781"
PROFILE_OWNER = "ppt_profile_probe::profile_write"
PROFILE_ANALYZER_SHA256 = "70b0bb3fc665860a9aa586725d72ced75e41de06f95e3b1c5408c70998e926b6"
CFB_READER_SHA256 = "7771347c1840f9be7997762a0e6d05657550dce5ceed358da0f3c396e3f302af"
SHAPES = ("tiny", "many", "payload", "unicode", "rich")
MODES = ("write", "lifecycle")
DIMENSIONS = {
    "tiny": (2, 3),
    "many": (100, 10),
    "payload": (16, 4),
    "unicode": (16, 4),
    "rich": (2, 2),
}
PROFILE_SHAPES = ("many", "unicode", "payload", "rich")
BOOTSTRAP_SEED = 841084
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_INDEXES = (249, 9749)
BOOTSTRAP_CONFIDENCE = 0.95
MAX_SAMPLES = 100_000
MAX_WARMUP = 100_000
HEX = frozenset("0123456789abcdefABCDEF")

TIMING_SCOPE = {
    "write": (
        "Writer::write_to only; Writer construction and slide/text-box authoring are outside the clock"
    ),
    "lifecycle": "Writer::new, add_slide, add_textbox/add_rich_textbox, and Writer::write_to",
}

REPORT_KEYS = {
    "schema",
    "tool",
    "profile_owner",
    "mode",
    "shape",
    "slides",
    "boxes_per_slide",
    "text_boxes",
    "timing_scope",
    "source",
    "expected_semantic_text_bytes",
    "expected_raw_text_bytes",
    "warmup",
    "samples_requested",
    "samples",
    "allocator",
}
SAMPLE_KEYS = {
    "index",
    "elapsed_ns",
    "metrics",
    "source_sha256",
    "output",
    "verification",
}
METRIC_KEYS = {"elapsed_ns", "slides", "boxes_per_slide"}
VERIFICATION_KEYS = {
    "reopened",
    "slide_count",
    "expected_slide_count",
    "slide_count_match",
    "semantic_check",
    "expected_semantic_text_bytes",
    "expected_semantic_text_sha256",
    "semantic_text_bytes",
    "semantic_text_sha256",
    "exact_text_match",
    "raw_text_check",
    "expected_raw_text_bytes",
    "expected_raw_text_sha256",
    "raw_text_bytes",
    "raw_text_sha256",
    "raw_text_box_count",
    "expected_raw_text_box_count",
    "raw_text_atom_count",
    "raw_text_boxes_match",
    "raw_text_match",
}


class ReaderError(ValueError):
    """Raised when retained evidence is absent, malformed, or contradictory."""


def fail(message: str) -> None:
    raise ReaderError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _json(path: Path) -> Any:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing JSON report: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON report {path}: {error}")


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> str:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def _sha(value: Any, label: str) -> str:
    require(
        isinstance(value, str) and len(value) == 64 and all(char in HEX for char in value),
        f"{label} is not a SHA-256 hexadecimal string",
    )
    return value.lower()


def _uint(value: Any, label: str, *, positive: bool = False) -> int:
    require(
        isinstance(value, int) and not isinstance(value, bool) and value >= 0,
        f"{label} is not a non-negative integer",
    )
    if positive:
        require(value > 0, f"{label} is not positive")
    return value


def _finite(value: Any, label: str) -> float:
    require(
        isinstance(value, (int, float)) and not isinstance(value, bool),
        f"{label} is not numeric",
    )
    result = float(value)
    require(math.isfinite(result), f"{label} is not finite")
    return result


def _identity(value: Any, label: str, *, positive: bool = True) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an identity object")
    require(set(value) == {"bytes", "sha256"}, f"{label} fields changed")
    result = {
        "bytes": _uint(value.get("bytes"), f"{label}.bytes", positive=positive),
        "sha256": _sha(value.get("sha256"), f"{label}.sha256"),
    }
    return result


def _read_sealed_file(packet: Path, relative: str, *, label: str) -> tuple[Path, Any]:
    seal_path = packet / "seal.json"
    seal = _json(seal_path)
    require(isinstance(seal, dict), f"{label} seal is not an object")
    files = seal.get("files")
    require(isinstance(files, dict), f"{label} seal file map is missing")
    # Historical seals use a packet-relative key in some changes and a
    # repository-relative key in others.  Resolve only the exact packet file;
    # a generic suffix match would allow a draft with the same basename.
    key = relative
    if key not in files:
        candidate = f"docs/performance/results/{packet.name}/{relative}"
        require(candidate in files, f"{label} seal has no entry for {relative}")
        key = candidate
    expected = _sha(files.get(key), f"{label} seal {key}")
    path = packet / relative
    require(path.is_file() and not path.is_symlink(), f"missing sealed {label}: {path}")
    require(_sha256_file(path) == expected, f"sealed {label} changed: {relative}")
    return path, expected


@functools.lru_cache(maxsize=None)
def baseline_identity(shape: str, mode: str) -> dict[str, Any]:
    """Return the immutable 0782 qualification identity for one case."""

    require(shape in SHAPES, f"unknown PPT shape {shape!r}")
    require(mode in MODES, f"unknown PPT mode {mode!r}")
    old_packet = PACKET.parent / "change-0782"
    relative = f"qualification/0-{shape}-{mode}-before.json"
    path, sealed_sha = _read_sealed_file(old_packet, relative, label="0782")
    report = _json(path)
    require(isinstance(report, dict), f"0782 {relative} is not an object")
    require(report.get("schema") == GENERATOR_SCHEMA, f"0782 {relative} schema changed")
    require(report.get("tool") == GENERATOR_TOOL, f"0782 {relative} tool changed")
    require(report.get("mode") == mode and report.get("shape") == shape,
            f"0782 {relative} mode/shape changed")
    dimensions = DIMENSIONS[shape]
    require((report.get("slides"), report.get("boxes_per_slide")) == dimensions,
            f"0782 {relative} dimensions changed")
    require(report.get("text_boxes") == dimensions[0] * dimensions[1],
            f"0782 {relative} text-box count changed")
    source = _identity(report.get("source"), f"0782 {relative}.source")
    semantic_bytes = _uint(report.get("expected_semantic_text_bytes"),
                           f"0782 {relative}.expected_semantic_text_bytes")
    raw_bytes = _uint(report.get("expected_raw_text_bytes"),
                      f"0782 {relative}.expected_raw_text_bytes")
    require(report.get("warmup") == 0 and report.get("samples_requested") == 1,
            f"0782 {relative} qualification schedule changed")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1,
            f"0782 {relative} qualification sample count changed")
    sample = samples[0]
    require(isinstance(sample, dict) and sample.get("index") == 0,
            f"0782 {relative} qualification sample is malformed")
    elapsed = _uint(sample.get("elapsed_ns"), f"0782 {relative}.elapsed_ns", positive=True)
    metrics = sample.get("metrics")
    require(isinstance(metrics, dict) and set(metrics) == METRIC_KEYS,
            f"0782 {relative} qualification metrics changed")
    require(metrics.get("elapsed_ns") == elapsed
            and metrics.get("slides") == dimensions[0]
            and metrics.get("boxes_per_slide") == dimensions[1],
            f"0782 {relative} qualification metrics disagree")
    require(sample.get("source_sha256") == source["sha256"],
            f"0782 {relative} qualification source digest changed")
    output = _identity(sample.get("output"), f"0782 {relative}.output")
    verification = sample.get("verification")
    require(isinstance(verification, dict) and set(verification) == VERIFICATION_KEYS,
            f"0782 {relative} verification fields changed")
    _validate_verification_fields(verification, shape, source, semantic_bytes, raw_bytes,
                                  f"0782 {relative}")
    require(verification["semantic_text_sha256"] == verification["expected_semantic_text_sha256"],
            f"0782 {relative} semantic digest is not self-consistent")
    require(verification["raw_text_sha256"] == verification["expected_raw_text_sha256"],
            f"0782 {relative} raw digest is not self-consistent")
    return {
        "shape": shape,
        "mode": mode,
        "sealed_report": str(path),
        "sealed_report_sha256": sealed_sha,
        "source": source,
        "slides": dimensions[0],
        "boxes_per_slide": dimensions[1],
        "text_boxes": dimensions[0] * dimensions[1],
        "expected_semantic_text_bytes": semantic_bytes,
        "expected_raw_text_bytes": raw_bytes,
        "semantic_text": {
            "bytes": verification["expected_semantic_text_bytes"],
            "sha256": verification["expected_semantic_text_sha256"],
        },
        "raw_text": {
            "bytes": verification["expected_raw_text_bytes"],
            "sha256": verification["expected_raw_text_sha256"],
        },
        "verification": dict(verification),
        "raw_text_atom_count": verification["raw_text_atom_count"],
        "output": output,
        "qualification_metrics": dict(metrics),
        "qualification_elapsed_ns": elapsed,
        "generator_identity": {
            "schema": report.get("schema"),
            "tool": report.get("tool"),
            "probe_schema": REPORT_SCHEMA,
            "probe_tool": REPORT_TOOL,
            "profile_owner": PROFILE_OWNER,
            "source": source,
            "source_report": str(path),
            "source_report_sha256": sealed_sha,
            "semantic_text": {
                "bytes": verification["expected_semantic_text_bytes"],
                "sha256": verification["expected_semantic_text_sha256"],
            },
            "raw_text": {
                "bytes": verification["expected_raw_text_bytes"],
                "sha256": verification["expected_raw_text_sha256"],
            },
        },
    }


def _validate_verification_fields(
    verification: Mapping[str, Any], shape: str, source: Mapping[str, Any],
    semantic_bytes: int, raw_bytes: int, label: str,
) -> dict[str, Any]:
    require(set(verification) == VERIFICATION_KEYS, f"{label}.verification fields changed")
    dimensions = DIMENSIONS[shape]
    for name in (
        "reopened",
        "slide_count_match",
        "semantic_check",
        "exact_text_match",
        "raw_text_check",
        "raw_text_boxes_match",
        "raw_text_match",
    ):
        require(verification.get(name) is True, f"{label}.verification.{name} is not true")
    require(_uint(verification.get("slide_count"), f"{label}.verification.slide_count")
            == dimensions[0], f"{label}.verification.slide_count changed")
    require(_uint(verification.get("expected_slide_count"),
                  f"{label}.verification.expected_slide_count") == dimensions[0],
            f"{label}.verification.expected_slide_count changed")
    require(verification.get("slide_count") == verification.get("expected_slide_count"),
            f"{label}.verification slide counts disagree")
    for prefix, expected, digest_name in (
        ("semantic", semantic_bytes, "semantic_text_sha256"),
        ("raw", raw_bytes, "raw_text_sha256"),
    ):
        expected_bytes_name = f"expected_{prefix}_text_bytes"
        actual_bytes_name = f"{prefix}_text_bytes"
        expected_digest_name = f"expected_{prefix}_text_sha256"
        require(_uint(verification.get(expected_bytes_name), f"{label}.{expected_bytes_name}") == expected,
                f"{label}.{expected_bytes_name} differs from generator identity")
        require(_uint(verification.get(actual_bytes_name), f"{label}.{actual_bytes_name}") == expected,
                f"{label}.{actual_bytes_name} differs from expected bytes")
        expected_digest = _sha(verification.get(expected_digest_name),
                               f"{label}.{expected_digest_name}")
        actual_digest = _sha(verification.get(digest_name), f"{label}.{digest_name}")
        require(actual_digest == expected_digest, f"{label}.{digest_name} differs from expected digest")
    expected_boxes = _uint(verification.get("expected_raw_text_box_count"),
                          f"{label}.expected_raw_text_box_count")
    actual_boxes = _uint(verification.get("raw_text_box_count"),
                        f"{label}.raw_text_box_count")
    atoms = _uint(verification.get("raw_text_atom_count"),
                  f"{label}.raw_text_atom_count", positive=True)
    require(expected_boxes == dimensions[0] * dimensions[1],
            f"{label}.expected_raw_text_box_count differs from dimensions")
    require(actual_boxes == expected_boxes, f"{label}.raw_text_box_count differs")
    require(atoms == actual_boxes, f"{label}.raw_text_atom_count differs from text-box count")
    _sha(verification.get("expected_semantic_text_sha256"),
         f"{label}.expected_semantic_text_sha256")
    _sha(verification.get("expected_raw_text_sha256"),
         f"{label}.expected_raw_text_sha256")
    require(isinstance(source, Mapping), f"{label}.source identity is invalid")
    _identity(dict(source), f"{label}.source")
    return dict(verification)


def _validate_allocator(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}.allocator is missing")
    require(set(value) == {"binary", "allocator", "instrumentation", "counter_revision"},
            f"{label}.allocator fields changed")
    require(value.get("binary") == "ppt-profile-probe",
            f"{label}.allocator.binary identity changed")
    require(value.get("allocator") == "Rust system allocator",
            f"{label}.allocator identity changed")
    require(value.get("instrumentation") == "none",
            f"{label}.allocator instrumentation changed")
    require(value.get("counter_revision") is None,
            f"{label}.allocator counter revision changed")
    return dict(value)


def _validate_artifact(path: Path, expected: Mapping[str, Any], label: str) -> dict[str, Any]:
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"{label} artifact is missing: {path}")
    try:
        data = path.read_bytes()
    except OSError as error:
        fail(f"cannot read {label} artifact {path}: {error}")
    identity = {"bytes": len(data), "sha256": _sha256(data)}
    expected_identity = _identity(dict(expected), f"{label}.artifact")
    require(identity == expected_identity, f"{label} artifact identity differs from report")
    parsed = parse_cfb_artifact(data)
    require(parsed.get("format") == "CFB/OLE2", f"{label} artifact is not CFB/OLE2")
    streams = parsed.get("streams")
    require(isinstance(streams, list) and streams, f"{label} CFB stream inventory is empty")
    document = next((row for row in streams if row.get("name") == "PowerPoint Document"), None)
    require(isinstance(document, dict), f"{label} PowerPoint Document stream is missing")
    require(_uint(document.get("bytes"), f"{label}.PowerPoint Document.bytes", positive=True)
            == document.get("bytes"), f"{label} PowerPoint Document length is invalid")
    _sha(document.get("sha256"), f"{label}.PowerPoint Document.sha256")
    return {
        "bytes": identity["bytes"],
        "sha256": identity["sha256"],
        "powerpoint_document": {
            "bytes": document["bytes"],
            "sha256": document["sha256"],
        },
        "streams": [
            {key: row[key] for key in ("name", "bytes", "sha256", "storage") if key in row}
            for row in streams
        ],
        "parser": {
            "header": parsed.get("header"),
            "directory_entries": len(parsed.get("directory", [])),
            "stream_count": len(streams),
        },
    }


def validate_report(
    path: Path,
    shape: str,
    mode: str,
    samples: int,
    warmup: int,
    artifact_path: Path | None = None,
    expected_output: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate one native, qualification, or profile JSON report.

    ``samples`` and ``warmup`` are supplied by the retained schedule rather
    than inferred from the report.  This prevents a report from redefining its
    own sample cardinality.  ``expected_output`` is optional for callers that
    already loaded the sealed qualification identity; the 0782 seal remains
    the mandatory anchor either way.
    """

    require(shape in SHAPES, f"unknown PPT shape {shape!r}")
    require(mode in MODES, f"unknown PPT mode {mode!r}")
    require(type(samples) is int and 0 < samples <= MAX_SAMPLES,
            "samples must be a positive bounded integer")
    require(type(warmup) is int and 0 <= warmup <= MAX_WARMUP,
            "warmup must be a bounded non-negative integer")
    baseline = baseline_identity(shape, mode)
    report = _json(Path(path))
    require(isinstance(report, dict), f"{path} report is not an object")
    label = str(path)
    report_keys = set(report)
    require(report_keys in (REPORT_KEYS, REPORT_KEYS | {"artifact"}),
            f"{label} report fields changed")
    require(report.get("schema") == REPORT_SCHEMA, f"{label} schema changed")
    require(report.get("tool") == REPORT_TOOL, f"{label} tool changed")
    require(report.get("profile_owner") == PROFILE_OWNER,
            f"{label} profile owner changed")
    require(report.get("mode") == mode and report.get("shape") == shape,
            f"{label} mode/shape changed")
    dimensions = DIMENSIONS[shape]
    require(_uint(report.get("slides"), f"{label}.slides") == dimensions[0],
            f"{label} slide dimensions changed")
    require(_uint(report.get("boxes_per_slide"), f"{label}.boxes_per_slide") == dimensions[1],
            f"{label} box dimensions changed")
    require(_uint(report.get("text_boxes"), f"{label}.text_boxes")
            == dimensions[0] * dimensions[1], f"{label} text-box count changed")
    require(report.get("timing_scope") == TIMING_SCOPE[mode],
            f"{label} timing scope changed")
    require(report.get("source") == baseline["source"],
            f"{label} source identity differs from sealed 0782 generator")
    require(_uint(report.get("expected_semantic_text_bytes"),
                  f"{label}.expected_semantic_text_bytes")
            == baseline["expected_semantic_text_bytes"],
            f"{label} expected semantic byte count differs")
    require(_uint(report.get("expected_raw_text_bytes"), f"{label}.expected_raw_text_bytes")
            == baseline["expected_raw_text_bytes"],
            f"{label} expected raw byte count differs")
    require(_uint(report.get("warmup"), f"{label}.warmup") == warmup
            and _uint(report.get("samples_requested"), f"{label}.samples_requested") == samples,
            f"{label} schedule differs from retained capture plan")
    allocator = _validate_allocator(report.get("allocator"), label)
    rows = report.get("samples")
    require(isinstance(rows, list) and len(rows) == samples,
            f"{label} sample count differs from retained capture plan")
    expected = baseline["output"]
    if expected_output is not None:
        supplied = _identity(expected_output, f"{label}.expected_output")
        require(supplied == expected, f"{label} supplied output anchor differs from 0782 seal")
    normalized: list[dict[str, Any]] = []
    for index, sample in enumerate(rows):
        sample_label = f"{label}.samples[{index}]"
        require(isinstance(sample, dict) and set(sample) == SAMPLE_KEYS,
                f"{sample_label} fields changed")
        require(sample.get("index") == index, f"{sample_label}.index is not contiguous")
        elapsed = _uint(sample.get("elapsed_ns"), f"{sample_label}.elapsed_ns", positive=True)
        metrics = sample.get("metrics")
        require(isinstance(metrics, dict) and set(metrics) == METRIC_KEYS,
                f"{sample_label}.metrics fields changed")
        require(_uint(metrics.get("elapsed_ns"), f"{sample_label}.metrics.elapsed_ns") == elapsed,
                f"{sample_label}.metrics.elapsed_ns disagrees")
        require(_uint(metrics.get("slides"), f"{sample_label}.metrics.slides") == dimensions[0]
                and _uint(metrics.get("boxes_per_slide"),
                          f"{sample_label}.metrics.boxes_per_slide") == dimensions[1],
                f"{sample_label}.metrics dimensions changed")
        require(sample.get("source_sha256") == baseline["source"]["sha256"],
                f"{sample_label}.source_sha256 differs from generator identity")
        output = _identity(sample.get("output"), f"{sample_label}.output")
        require(output == expected, f"{sample_label}.output differs from retained qualification")
        verification = sample.get("verification")
        require(isinstance(verification, dict), f"{sample_label}.verification is missing")
        _validate_verification_fields(
            verification,
            shape,
            baseline["source"],
            baseline["expected_semantic_text_bytes"],
            baseline["expected_raw_text_bytes"],
            sample_label,
        )
        # Every semantic and raw field is part of the sealed generator
        # identity.  This also makes the reader reject a report whose actual
        # oracle happened to remain self-consistent after a corpus mutation.
        require(verification == baseline["verification"],
                f"{sample_label}.verification differs from sealed qualification")
        normalized.append({
            "index": index,
            "elapsed_ns": elapsed,
            "wall_ns": elapsed,
            "metrics": dict(metrics),
            "source_sha256": sample["source_sha256"],
            "output": output,
            "verification": dict(verification),
        })
    artifact = None
    if artifact_path is not None:
        require("artifact" in report, f"{label} artifact identity is missing")
        report_artifact = _identity(report.get("artifact"), f"{label}.artifact")
        require(report_artifact == expected,
                f"{label}.artifact differs from sealed qualification output")
        artifact = _validate_artifact(Path(artifact_path), report_artifact, label)
    else:
        require("artifact" not in report,
                f"{label} retained artifact was not independently supplied")
    walls = [row["elapsed_ns"] for row in normalized]
    return {
        "report": {
            "path": str(Path(path)),
            "bytes": Path(path).stat().st_size,
            "sha256": _sha256_file(Path(path)),
        },
        "schema": REPORT_SCHEMA,
        "tool": REPORT_TOOL,
        "case": {"shape": shape, "mode": mode},
        "shape": shape,
        "mode": mode,
        "samples_requested": samples,
        "warmup": warmup,
        "generator_identity": baseline["generator_identity"],
        "corpus_identity": baseline["source"],
        "source": baseline["source"],
        "output": expected,
        "artifact_identity": artifact,
        "verification": baseline["verification"],
        "verification_ok": True,
        "wall_ns": walls,
        "cpu_ns": None,
        "p50_ns": nearest_rank(walls, 0.50),
        "p95_ns": nearest_rank(walls, 0.95),
        "p99_ns": nearest_rank(walls, 0.99),
        "mean_ns": statistics.fmean(walls),
        "samples": normalized,
        "allocator": allocator,
    }


def parse_report(*args: Any, **kwargs: Any) -> dict[str, Any]:
    """Compatibility alias for analysis callers using the older spelling."""

    return validate_report(*args, **kwargs)


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    require(isinstance(quantile, (int, float)) and not isinstance(quantile, bool)
            and 0.0 <= float(quantile) <= 1.0, "quantile is outside [0, 1]")
    ordered = sorted(_finite(value, "nearest-rank input") for value in values)
    require(ordered, "nearest-rank received no values")
    rank = max(1, math.ceil(float(quantile) * len(ordered)))
    return ordered[min(rank, len(ordered)) - 1]


def bootstrap(values: Sequence[float], *, seed: int = BOOTSTRAP_SEED,
              resamples: int = BOOTSTRAP_RESAMPLES,
              endpoint_indexes: tuple[int, int] = BOOTSTRAP_INDEXES) -> dict[str, Any]:
    require(values, "bootstrap received no values")
    require(type(resamples) is int and resamples > max(endpoint_indexes) >= 0,
            "bootstrap endpoint indexes are invalid")
    numbers = [_finite(value, "bootstrap input") for value in values]
    rng = random.Random(seed)
    estimates = sorted(
        statistics.median(rng.choice(numbers) for _ in numbers)
        for _ in range(resamples)
    )
    return {
        "estimate": statistics.median(numbers),
        "lower": estimates[endpoint_indexes[0]],
        "upper": estimates[endpoint_indexes[1]],
        "seed": seed,
        "resamples": resamples,
        "confidence": BOOTSTRAP_CONFIDENCE,
        "statistic": "median",
        "endpoint_indexes": list(endpoint_indexes),
    }


def bootstrap_ratio(before: Sequence[float], after: Sequence[float],
                    name: str = "ratio") -> dict[str, Any]:
    require(len(before) == len(after) and before,
            f"{name} pair cardinality is empty or differs")
    old = [_finite(value, f"{name}.before") for value in before]
    new = [_finite(value, f"{name}.after") for value in after]
    ratios: list[float | None] = []
    for index, (left, right) in enumerate(zip(old, new)):
        require(left >= 0 and right >= 0, f"{name} pair {index} is negative")
        ratios.append(right / left if left > 0 else None)
    usable = [value for value in ratios if value is not None]
    interval = bootstrap(usable) if usable else None
    result: dict[str, Any] = {
        "name": name,
        "before": old,
        "after": new,
        "ratios": ratios,
        "block_ratios": ratios,
        "raw_distributions_retained": True,
    }
    if interval is None:
        result.update({"estimate": None, "ci95_low": None, "ci95_high": None,
                       "bootstrap": None})
    else:
        result.update({"estimate": interval["estimate"],
                       "ci95_low": interval["lower"],
                       "ci95_high": interval["upper"],
                       "bootstrap": interval})
    return result


def paired_metric(before: Sequence[float], after: Sequence[float],
                  name: str = "metric") -> dict[str, Any]:
    require(len(before) == len(after) and before,
            f"{name} pair cardinality is empty or differs")
    old = [_finite(value, f"{name}.before") for value in before]
    new = [_finite(value, f"{name}.after") for value in after]
    difference = bootstrap([right - left for left, right in zip(old, new)])
    result = bootstrap_ratio(old, new, name)
    result.update({
        "before_block_values": old,
        "after_block_values": new,
        "block_deltas": [right - left for left, right in zip(old, new)],
        "difference_bootstrap": difference,
        "difference_estimate": difference["estimate"],
        "difference_ci95_low": difference["lower"],
        "difference_ci95_high": difference["upper"],
    })
    return result


@functools.lru_cache(maxsize=1)
def _cfb_reader_module() -> Any:
    """Load the 0840 parser only after replaying its sealed source digest."""

    packet = PACKET.parent / "change-0840"
    relative = "readers.py"
    path, expected = _read_sealed_file(packet, relative, label="0840 CFB parser")
    require(expected == CFB_READER_SHA256,
            "0840 CFB parser seal does not match the retained implementation")
    spec = importlib.util.spec_from_file_location("litchi_0840_cfb_reader", path)
    require(spec is not None and spec.loader is not None,
            "cannot load sealed 0840 CFB parser")
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as error:  # pragma: no cover - exercised by custody mutations
        fail(f"sealed 0840 CFB parser failed to load: {error}")
    require(_sha256_file(path) == expected, "sealed 0840 CFB parser changed after load")
    require(callable(getattr(module, "parse_cfb_artifact", None)),
            "sealed 0840 parser has no parse_cfb_artifact")
    return module


def parse_cfb_artifact(data: bytes | bytearray | memoryview) -> dict[str, Any]:
    """Parse a CFB artifact using the hash-bound 0840 independent parser."""

    try:
        return _cfb_reader_module().parse_cfb_artifact(data)
    except ReaderError:
        raise
    except Exception as error:
        fail(f"CFB parser rejected artifact: {error}")


def validate_cfb_artifact(data: bytes | bytearray | memoryview) -> dict[str, Any]:
    """Compatibility alias returning the independent CFB parse."""

    return parse_cfb_artifact(data)


# The following parser is a packet-local adapter of the 0784 raw Callgrind
# grammar.  It is intentionally not presented as byte-for-byte reuse: this
# copy adds 0841's strict duplicate/header and non-negative-cost checks and
# keeps the 0841 owner/report vocabulary local.  _check_profile_parser_source
# binds the adapted implementation to the reviewed 0784 source digest so a
# changed historical parser cannot silently redefine these semantics.
FUNCTION_RE = re.compile(r"^(fn|cfn)=\((\d+)\)(?:\s+(.*))?$")
CALLS_RE = re.compile(r"^calls=([+\-]?[0-9][0-9,]*)")
EVENTS_RE = re.compile(r"^events:\s*(.*)$")
SUMMARY_RE = re.compile(r"^summary:\s*(.*)$")
TOTALS_RE = re.compile(r"^totals:\s*(.*)$")
PART_RE = re.compile(r"^part:\s*(\d+)$")
TRIGGER_RE = re.compile(r"^desc:\s+Trigger:\s+(.*)$")
CMD_RE = re.compile(r"^cmd:\s?(.*)$")


def _check_profile_parser_source() -> tuple[Path, str]:
    """Recheck the sealed 0784 parser source before every raw parse.

    The local grammar below is deliberately adapted for the new owner and
    report shape, but its compressed-name, edge, and cost semantics are bound
    to the reviewed 0784 implementation.  Rechecking on every entry also
    makes a post-import mutation fail closed rather than relying on a cached
    module object.
    """

    packet = PACKET.parent / "change-0784"
    path, expected = _read_sealed_file(packet, "profile_analysis.py",
                                       label="0784 Callgrind parser")
    require(expected == PROFILE_ANALYZER_SHA256,
            "0784 Callgrind parser seal does not match the retained implementation")
    return path, expected


def _parse_ints(text: str) -> list[int]:
    values: list[int] = []
    for token in text.split():
        try:
            values.append(int(token.replace(",", "")))
        except ValueError:
            break
    return values


def _header_values(lines: Sequence[str], label: str) -> dict[str, Any]:
    events: list[str] | None = None
    summary: int | None = None
    totals: int | None = None
    part: int | None = None
    trigger: str | None = None
    command: str | None = None
    for line in lines:
        stripped = line.strip()
        if match := EVENTS_RE.match(stripped):
            require(events is None, f"{label}: duplicate events header")
            events = match.group(1).split()
        elif match := SUMMARY_RE.match(stripped):
            values = _parse_ints(match.group(1))
            require(values, f"{label}: summary is not numeric")
            require(summary is None, f"{label}: duplicate summaries")
            summary = values[0]
        elif match := TOTALS_RE.match(stripped):
            values = _parse_ints(match.group(1))
            require(values, f"{label}: totals is not numeric")
            require(totals is None, f"{label}: duplicate totals")
            totals = values[0]
        elif match := PART_RE.match(stripped):
            require(part is None, f"{label}: duplicate parts")
            part = int(match.group(1))
        elif match := TRIGGER_RE.match(stripped):
            require(trigger is None, f"{label}: duplicate triggers")
            trigger = match.group(1)
        elif command is None and (match := CMD_RE.match(stripped)):
            command = match.group(1)
    require(events == ["Ir"], f"{label}: Callgrind event set is not exactly Ir")
    require(summary is not None and summary >= 0, f"{label}: summary is missing or negative")
    require(totals is not None and totals >= 0, f"{label}: totals is missing or negative")
    require(part is not None, f"{label}: part is missing")
    require(trigger is not None and trigger, f"{label}: trigger is missing")
    return {
        "events": events,
        "summary_ir": summary,
        "totals_ir": totals,
        "part": part,
        "trigger": trigger,
        "command": command,
    }


def _add_name(names: dict[int, str], function_id: int, name: str, label: str) -> None:
    if not name:
        return
    prior = names.get(function_id)
    require(prior is None or prior == name,
            f"{label}: function id {function_id} has conflicting names")
    names[function_id] = name


def _cost_line(line: str, event_index: int) -> tuple[int, str] | None:
    fields = line.split()
    if len(fields) <= event_index + 1:
        return None
    position = fields[0]
    if position != "*" and not re.fullmatch(r"[+\-]?\d+", position):
        return None
    try:
        return int(fields[event_index + 1].replace(",", "")), position
    except ValueError:
        return None


def _position_kind(position: str) -> str:
    if position == "*":
        return "wildcard"
    if position.startswith(("+", "-")):
        return "relative"
    return "absolute"


def parse_raw(path: Path, allow_empty: bool = False) -> dict[str, Any]:
    """Parse one raw Callgrind file without invoking callgrind_annotate."""

    _check_profile_parser_source()
    path = Path(path)
    require(path.is_file() and not path.is_symlink(), f"missing Callgrind file: {path}")
    label = str(path)
    try:
        lines = path.read_text(encoding="utf-8", errors="strict").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"cannot read {label}: {error}")
    header = _header_values(lines, label)
    names: dict[int, str] = {}
    declarations = 0
    for line in lines:
        match = FUNCTION_RE.match(line.strip())
        if match and match.group(3):
            declarations += 1
            _add_name(names, int(match.group(2)), match.group(3), label)
    functions: dict[int, dict[str, Any]] = {}
    current_id: int | None = None
    pending: dict[str, Any] | None = None
    costs = 0
    edges = 0
    position_counts: dict[str, int] = {}
    for line_number, line in enumerate(lines, 1):
        stripped = line.strip()
        function_match = FUNCTION_RE.match(stripped)
        if function_match:
            kind, text_id, name = function_match.groups()
            function_id = int(text_id)
            if name:
                _add_name(names, function_id, name, f"{label}:{line_number}")
            if kind == "fn":
                require(pending is None or pending.get("calls") is not None,
                        f"{label}:{line_number}: cfn is missing calls")
                current_id = function_id
                pending = None
                function = functions.setdefault(
                    function_id,
                    {"id": function_id, "name": "", "records": 0,
                     "self_ir": 0, "edges": []},
                )
                function["records"] += 1
            else:
                require(current_id is not None, f"{label}:{line_number}: cfn outside fn")
                require(pending is None, f"{label}:{line_number}: nested cfn before cost")
                pending = {"callee_id": function_id, "calls": None, "line": line_number}
            continue
        calls_match = CALLS_RE.match(stripped)
        if calls_match:
            require(current_id is not None and pending is not None,
                    f"{label}:{line_number}: calls outside cfn")
            require(pending["calls"] is None,
                    f"{label}:{line_number}: duplicate calls record")
            calls = int(calls_match.group(1).replace(",", ""))
            require(calls >= 0, f"{label}:{line_number}: negative call count")
            pending["calls"] = calls
            continue
        if current_id is None:
            continue
        parsed = _cost_line(stripped, 0)
        if parsed is None:
            continue
        value, position = parsed
        require(value >= 0, f"{label}:{line_number}: negative Ir cost")
        position_counts[_position_kind(position)] = position_counts.get(_position_kind(position), 0) + 1
        costs += 1
        function = functions[current_id]
        if pending is not None:
            require(pending["calls"] is not None,
                    f"{label}:{line_number}: cfn cost has no calls record")
            edge = {
                "caller_id": current_id,
                "callee_id": pending["callee_id"],
                "calls": pending["calls"],
                "inclusive_ir": value,
                "line": pending["line"],
                "position": position,
            }
            function["edges"].append(edge)
            edges += 1
            pending = None
        else:
            function["self_ir"] += value
    require(pending is None or pending.get("calls") is not None,
            f"{label}: unterminated cfn record")
    for function_id, function in functions.items():
        function["name"] = names.get(function_id, "")
    if allow_empty and header["summary_ir"] == 0 and not names and costs == 0:
        functions = {}
    else:
        require(names, f"{label}: no compressed function names were resolved")
        require(costs > 0, f"{label}: no Ir cost records were parsed")
    return {
        "path": label,
        "sha256": _sha256_file(path),
        "bytes": path.stat().st_size,
        "header": header,
        "names": names,
        "functions": functions,
        "statistics": {
            "function_records": sum(item["records"] for item in functions.values()),
            "unique_functions": len(functions),
            "compressed_name_declarations": declarations,
            "cost_records": costs,
            "call_edge_records": edges,
            "position_kinds": dict(sorted(position_counts.items())),
        },
    }


def _display_name(name: str) -> str:
    return name or "<unnamed>"


def aggregate_edges(function: Mapping[str, Any], names: Mapping[int, str]) -> list[dict[str, Any]]:
    grouped: dict[int, dict[str, Any]] = {}
    for edge in function.get("edges", []):
        if edge["calls"] <= 0 or edge["inclusive_ir"] <= 0:
            continue
        item = grouped.setdefault(edge["callee_id"], {
            "callee_id": edge["callee_id"],
            "callee": names.get(edge["callee_id"], ""),
            "calls": 0,
            "edge_count": 0,
            "inclusive_ir": 0,
        })
        item["calls"] += edge["calls"]
        item["edge_count"] += 1
        item["inclusive_ir"] += edge["inclusive_ir"]
    return sorted(grouped.values(), key=lambda item: (-item["inclusive_ir"], item["callee_id"]))


def function_view(function: Mapping[str, Any], names: Mapping[int, str]) -> dict[str, Any]:
    direct = aggregate_edges(function, names)
    return {
        "id": function["id"],
        "name": _display_name(function.get("name", "")),
        "records": function["records"],
        "self_ir": function["self_ir"],
        "direct_children": direct,
        "direct_children_ir": sum(item["inclusive_ir"] for item in direct),
    }


def incoming_edges(parsed: Mapping[str, Any], target_id: int) -> list[dict[str, Any]]:
    names = parsed["names"]
    result: list[dict[str, Any]] = []
    for function in parsed["functions"].values():
        for edge in function["edges"]:
            if (edge["callee_id"] == target_id and edge["calls"] > 0
                    and edge["inclusive_ir"] > 0):
                result.append({
                    **edge,
                    "caller": _display_name(names.get(edge["caller_id"], "")),
                    "callee": _display_name(names.get(edge["callee_id"], "")),
                })
    return result


def validate_profile(
    numbered: Path,
    terminal: Path,
    shape: str | None = None,
    baseline: Mapping[str, Any] | None = None,
    report_path: Path | None = None,
    owner: str = PROFILE_OWNER,
) -> dict[str, Any]:
    """Validate one numbered/terminal Callgrind publication pair.

    For compatibility with early analysis drafts, a third positional string
    containing ``::`` is treated as ``owner``.  Passing ``report_path`` also
    checks the profile's one-sample JSON report against the sealed 0782
    generator identity.
    """

    if isinstance(shape, str) and "::" in shape and baseline is None:
        owner, shape = shape, None
    require(isinstance(owner, str) and owner, "Callgrind owner is missing")
    number = parse_raw(Path(numbered))
    final = parse_raw(Path(terminal), allow_empty=True)
    label = str(numbered)
    require(number["header"]["part"] == 1, f"{label}: numbered part changed")
    require(number["header"]["trigger"] == "--dump-after=" + owner,
            f"{label}: collection trigger changed")
    require(number["header"]["summary_ir"] > 0, f"{label}: scoped summary is zero")
    require(number["header"]["totals_ir"] == number["header"]["summary_ir"],
            f"{label}: totals do not match summary")
    require(final["header"]["part"] == 2 and final["header"]["trigger"] == "Program termination",
            f"{terminal}: termination dump changed")
    require(final["header"]["summary_ir"] == 0 and final["header"]["totals_ir"] == 0,
            f"{terminal}: termination dump is not zero Ir")
    require(number["header"]["command"] and final["header"]["command"],
            f"{label}: raw command is missing")
    require(number["header"]["command"] == final["header"]["command"],
            f"{label}: numbered and terminal command headers differ")
    command = number["header"]["command"]
    for option in (
        "--tool=callgrind",
        "--collect-atstart=no",
        "--toggle-collect=" + owner,
        "--zero-before=" + owner,
        "--dump-after=" + owner,
    ):
        require(option in command, f"{label}: Callgrind command lacks {option}")
    owner_ids = [fid for fid, function in number["functions"].items()
                 if function["name"] == owner]
    require(len(owner_ids) == 1, f"{label}: exact owner name resolved to {owner_ids}")
    owner_id = owner_ids[0]
    incoming = incoming_edges(number, owner_id)
    require(len(incoming) == 1 and incoming[0]["calls"] == 1,
            f"{label}: owner incoming positive call is not exactly one")
    require(incoming[0]["inclusive_ir"] == number["header"]["summary_ir"],
            f"{label}: owner incoming Ir differs from summary")
    owner_function = number["functions"][owner_id]
    owner_view = function_view(owner_function, number["names"])
    require(owner_view["self_ir"] + owner_view["direct_children_ir"]
            == number["header"]["summary_ir"],
            f"{label}: owner self plus direct children does not reconstruct summary")
    function_rows = [function_view(function, number["names"])
                     for function in number["functions"].values()]
    all_function_self_ir = sum(row["self_ir"] for row in function_rows)
    require(all_function_self_ir == number["header"]["summary_ir"],
            f"{label}: function self-cost totals do not reconstruct summary")
    function_rows.sort(key=lambda row: (-row["self_ir"], row["name"], row["id"]))
    top_self = [row for row in function_rows if row["self_ir"] > 0][:20]
    unexpected_dump = Path(numbered).with_name(Path(numbered).name.rsplit(".1", 1)[0] + ".2")
    require(not unexpected_dump.exists(), f"{label}: unexpected second numbered dump")
    if report_path is not None:
        require(shape in SHAPES, "profile report shape is missing")
        expected = None
        if baseline is not None:
            require(isinstance(baseline, Mapping), "profile baseline is not an object")
            expected = baseline.get("output", baseline)
            _identity(expected, "profile baseline output")
        report_identity = validate_report(Path(report_path), shape, "write", 1, 0,
                                          expected_output=expected)
    else:
        report_identity = None
    return {
        "file": label,
        "sha256": number["sha256"],
        "bytes": number["bytes"],
        "termination": {
            "file": str(terminal),
            "sha256": final["sha256"],
            "bytes": final["bytes"],
            "summary_ir": 0,
        },
        "summary_ir": number["header"]["summary_ir"],
        "totals_ir": number["header"]["totals_ir"],
        "owner": {
            "id": owner_id,
            "name": owner,
            "incoming": incoming[0],
            "self_ir": owner_view["self_ir"],
            "direct_children_ir": owner_view["direct_children_ir"],
            "direct_children": owner_view["direct_children"],
            "partition_equation": "self_ir + sum(immediate child inclusive_ir) = owner inclusive Ir",
            "partition_disjoint": True,
            "nested_inclusive_rows_excluded": True,
        },
        "function_rows": function_rows,
        "top_self_functions": top_self,
        "top_edges": owner_view["direct_children"],
        "function_totals": {
            "summary_ir": number["header"]["summary_ir"],
            "owner_self_ir": owner_view["self_ir"],
            "owner_direct_children_ir": owner_view["direct_children_ir"],
            "all_function_self_ir": all_function_self_ir,
        },
        "parser": number["statistics"],
        "report_identity": report_identity,
        "validation": {
            "exact_trigger": True,
            "summary_nonzero": True,
            "totals_equal_summary": True,
            "termination_zero_ir": True,
            "exactly_one_positive_owner_call": True,
            "owner_inclusive_matches_summary": True,
            "owner_self_plus_direct_children_equals_summary": True,
            "all_function_self_ir_equals_summary": True,
            "compressed_names_resolved": number["statistics"]["compressed_name_declarations"] > 0,
            "relative_or_wildcard_positions_parsed": any(
                number["statistics"]["position_kinds"].get(kind, 0) > 0
                for kind in ("relative", "wildcard")
            ),
            "terminal_zero": True,
            "single_publication": True,
        },
    }


def validate_callgrind(*args: Any, **kwargs: Any) -> dict[str, Any]:
    """Compatibility alias for callers using the descriptive name."""

    return validate_profile(*args, **kwargs)


def validate_profile_pair(*args: Any, **kwargs: Any) -> dict[str, Any]:
    """Compatibility alias used by profile analysis callers."""

    return validate_profile(*args, **kwargs)
