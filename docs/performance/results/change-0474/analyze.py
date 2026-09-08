#!/usr/bin/env python3
"""Derive portable evidence for the PPTX streaming creation experiment.

The capture driver owns process construction and sample collection.  This
module is deliberately a small, fail-closed reader of those captures: it
checks the frozen lane matrix, summarizes elapsed and allocator vectors, and
binds the deterministic output, sink, and writer-counter identities.  Normal
and allocator timings are kept in separate sections; this module never makes
a normal-versus-allocator latency comparison.
"""

from __future__ import annotations

import argparse
import copy
import report_checks
import hashlib
import json
import math
import re
from pathlib import Path
from typing import Any, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0474-evidence-analysis-v1"
PROTOCOL_SCHEMA = "litchi-0474-protocol-v1"
CASE = "pptx_streaming_create"
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
REPEATS = ("R1", "R2")
SAMPLES = 30
WARMUPS = 3
SLIDES = {"tiny": 8, "medium": 256, "large": 8_192}
HEX = re.compile(r"[0-9a-f]{64}\Z")
U64_MAX = (1 << 64) - 1
SIGNED_U64_MAX = (1 << 64) - 1
ALLOCATOR_FIELDS = (
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
)
SINK_FIELDS = ("accepted_bytes", "write_calls")


class AnalysisError(ValueError):
    """Raised when a capture cannot support the declared evidence."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def _reject_nonfinite(value: str) -> None:
    raise AnalysisError(f"non-finite JSON value {value!r}")


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise AnalysisError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(
                stream,
                object_pairs_hook=_reject_duplicate_keys,
                parse_constant=_reject_nonfinite,
            )
    except (OSError, UnicodeError, json.JSONDecodeError, AnalysisError) as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error


def canonical(value: Any) -> bytes:
    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError) as error:
        raise AnalysisError(f"cannot canonicalize JSON: {error}") from error


def canonical_equal(left: Any, right: Any) -> bool:
    return canonical(left) == canonical(right)


def sha256_file(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except OSError as error:
        raise AnalysisError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest(), size


def binding(path: Path, root: Path = ROOT) -> dict[str, Any]:
    digest, size = sha256_file(path)
    try:
        relative = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        relative = path.name
    return {"path": relative, "sha256": digest, "bytes": size}


def require_object(value: Any, label: str) -> Mapping[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label} must be an object")
    return value


def require_hex(value: Any, label: str) -> str:
    if not isinstance(value, str) or HEX.fullmatch(value) is None:
        fail(f"{label} must be a lowercase SHA-256")
    return value


def require_u64(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or not 0 <= value <= U64_MAX:
        fail(f"{label} must be a u64")
    return value


def require_signed(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or not -SIGNED_U64_MAX <= value <= SIGNED_U64_MAX:
        fail(f"{label} must be a checked signed integer")
    return value


def _percentile(ordered: Sequence[float], percentile: int) -> float:
    index = min(((percentile * len(ordered) + 99) // 100) - 1, len(ordered) - 1)
    return ordered[index]


def stats(values: Sequence[int | float], label: str) -> dict[str, Any]:
    if not values:
        fail(f"{label} cannot be empty")
    checked: list[float] = []
    for index, value in enumerate(values):
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            fail(f"{label}[{index}] must be numeric")
        converted = float(value)
        if not math.isfinite(converted) or converted < 0:
            fail(f"{label}[{index}] must be finite and non-negative")
        checked.append(converted)
    ordered = sorted(checked)
    count = len(ordered)
    result: dict[str, Any] = {
        "min": ordered[0],
        "p50": (ordered[(count - 1) // 2] + ordered[count // 2]) // 2,
        "p95": _percentile(ordered, 95),
        "p99": _percentile(ordered, 99),
        "max": ordered[-1],
        "mean": math.fsum(ordered) / count,
        "samples": list(values),
    }
    for key in ("min", "p50", "p95", "p99", "max"):
        if isinstance(result[key], float) and result[key].is_integer():
            result[key] = int(result[key])
    return result


def signed_stats(values: Sequence[int], label: str) -> dict[str, Any]:
    checked = [require_signed(value, f"{label}[{index}]") for index, value in enumerate(values)]
    if not checked:
        fail(f"{label} cannot be empty")
    ordered = sorted(checked)
    count = len(ordered)
    result: dict[str, Any] = {
        "min": ordered[0],
        "p50": (ordered[(count - 1) // 2] + ordered[count // 2]) // 2,
        "p95": _percentile([float(value) for value in ordered], 95),
        "p99": _percentile([float(value) for value in ordered], 99),
        "max": ordered[-1],
        "mean": math.fsum(ordered) / count,
        "samples": checked,
    }
    for key in ("min", "p50", "p95", "p99", "max"):
        if isinstance(result[key], float) and result[key].is_integer():
            result[key] = int(result[key])
    return result


def percent_delta(left: float, right: float) -> float:
    if left == 0:
        return 0.0 if right == 0 else math.inf
    return ((right / left) - 1.0) * 100.0


def protocol_order(protocol: Mapping[str, Any]) -> list[dict[str, Any]]:
    if protocol.get("schema") != PROTOCOL_SCHEMA:
        fail("protocol schema differs")
    if protocol.get("selector") != CASE:
        fail("protocol selector differs")
    shapes = protocol.get("shapes")
    if shapes != SLIDES:
        fail("protocol slide shapes differ")
    if protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS:
        fail("protocol samples or warmups differ")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol CPU or worker binding differs")
    raw = protocol.get("order")
    if not isinstance(raw, list) or len(raw) != len(MODES) * len(SHAPES) * len(REPEATS):
        fail("protocol order must contain twelve lanes")
    rows: list[dict[str, Any]] = []
    seen: set[tuple[str, str, str]] = set()
    for index, item in enumerate(raw):
        lane = dict(require_object(item, f"protocol.order[{index}"))
        mode, shape, repeat = lane.get("mode"), lane.get("shape"), lane.get("repeat")
        if mode not in MODES or shape not in SHAPES or repeat not in REPEATS:
            fail(f"protocol.order[{index}] contains an unknown lane")
        if not isinstance(lane.get("lane"), str) or not lane["lane"]:
            fail(f"protocol.order[{index}].lane is missing")
        identity = (mode, shape, repeat)
        if identity in seen:
            fail(f"protocol order repeats {identity!r}")
        seen.add(identity)
        rows.append(lane)
    expected = {(mode, shape, repeat) for repeat in REPEATS for mode in MODES for shape in SHAPES}
    if seen != expected:
        fail("protocol order does not contain each mode/shape/repeat exactly once")
    first = [(row["mode"], row["shape"]) for row in rows if row["repeat"] == "R1"]
    second = [(row["mode"], row["shape"]) for row in rows if row["repeat"] == "R2"]
    if len(first) != 6 or second != list(reversed(first)):
        fail("R2 protocol order must reverse R1")
    return rows


def lane_map(protocol: Mapping[str, Any]) -> dict[str, dict[str, Any]]:
    rows = protocol_order(protocol)
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        lane = row["lane"]
        if lane in result:
            fail(f"duplicate lane name {lane!r}")
        result[lane] = row
    return result


def lane_dir(root: Path, lane: str) -> Path:
    path = root / "captures" / lane
    if not path.is_dir():
        fail(f"missing capture directory {path}")
    return path


def _vector(value: Any, label: str, count: int = SAMPLES) -> list[int]:
    if isinstance(value, dict):
        if value.get("status") not in (None, "measured"):
            fail(f"{label} is not measured")
        value = value.get("values")
    if not isinstance(value, list) or len(value) != count:
        fail(f"{label} must contain {count} values")
    return [require_u64(item, f"{label}[{index}]") for index, item in enumerate(value)]


def _row(report: Mapping[str, Any], lane: Mapping[str, Any], label: str) -> Mapping[str, Any]:
    if report.get("schema_version") not in (1, "1"):
        fail(f"{label}: unsupported report schema")
    configuration = require_object(report.get("configuration"), f"{label}.configuration")
    if configuration.get("samples_per_case") != SAMPLES:
        fail(f"{label}: expected thirty samples")
    if configuration.get("warmup_iterations_per_case") != WARMUPS:
        fail(f"{label}: expected three warmups")
    cases = configuration.get("cases")
    if cases != [CASE]:
        fail(f"{label}: selector differs")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1:
        fail(f"{label}: expected one result row")
    row = require_object(results[0], f"{label}.results[0]")
    if row.get("case") != CASE:
        fail(f"{label}: result case differs")
    corpus = require_object(row.get("corpus"), f"{label}.corpus")
    if corpus.get("shape") != lane["shape"]:
        fail(f"{label}: corpus shape differs")
    if corpus.get("entry_count") != SLIDES[lane["shape"]]:
        fail(f"{label}: slide count differs")
    elapsed = require_object(row.get("elapsed_ns"), f"{label}.elapsed_ns")
    samples = elapsed.get("samples")
    if not isinstance(samples, list) or len(samples) != SAMPLES:
        fail(f"{label}: elapsed sample count differs")
    for index, value in enumerate(samples):
        if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
            fail(f"{label}.elapsed_ns.samples[{index}] must be a positive integer")
    sample_order = elapsed.get("sample_order")
    if sample_order is not None and sorted(sample_order) != list(range(SAMPLES)):
        fail(f"{label}: elapsed sample_order is not a permutation")
    output = row.get("output_sha256")
    if output is None:
        output_object = row.get("output")
        if isinstance(output_object, dict):
            output = output_object.get("sha256")
    require_hex(output, f"{label}.output_sha256")
    sink = require_object(row.get("sink"), f"{label}.sink")
    for field in SINK_FIELDS:
        require_u64(sink.get(field), f"{label}.sink.{field}")
    # Make the row's core sink identity available to callers without mutating
    # the capture.  PPTX writer counters are retained in source.pptx_slides;
    # the generic discard sink has no slide/text-box counter projection.
    return row


def writer_counters(row: Mapping[str, Any], label: str = "row") -> Mapping[str, Any]:
    source_root = require_object(row.get("source"), f"{label}.source")
    source = require_object(source_root.get("pptx_slides"), f"{label}.source.pptx_slides")
    fields = (
        "slide_count",
        "text_box_count",
        "input_text_bytes",
        "authored_part_bytes",
        "observed_max_slide_xml_bytes",
        "max_slide_xml_bytes",
        "structural_metadata_fixed_member_count",
        "structural_metadata_members_per_slide",
    )
    result = {field: require_u64(source.get(field), f"{label}.source.pptx_slides.{field}")
              for field in fields}
    return result


def allocation_vectors(row: Mapping[str, Any], label: str) -> dict[str, list[int]]:
    operation = row.get("operation_metrics")
    if isinstance(operation, dict):
        allocation = operation.get("allocation")
    else:
        allocation = None
    if allocation is None:
        allocation = row.get("allocation")
    allocation = require_object(allocation, f"{label}.allocation")
    if allocation.get("status") not in (None, "measured"):
        fail(f"{label}.allocation is not measured")
    vectors: dict[str, list[int]] = {}
    for field in ALLOCATOR_FIELDS:
        vectors[field] = _vector(allocation.get(field), f"{label}.allocation.{field}")
    for index, (region, before, after, peak) in enumerate(zip(
        vectors["region_peak_live_bytes"],
        vectors["live_bytes_before"],
        vectors["live_bytes_after"],
        vectors["peak_live_bytes_after"],
    )):
        if region < before or region < after or region > peak:
            fail(f"{label}.allocation region peak invariant fails at sample {index}")
    vectors["live_bytes_delta"] = [
        require_signed(after - before, f"{label}.allocation.live_bytes_delta[{index}]")
        for index, (before, after) in enumerate(zip(vectors["live_bytes_before"], vectors["live_bytes_after"]))
    ]
    vectors["peak_live_bytes_delta"] = [
        require_signed(region - before, f"{label}.allocation.peak_live_bytes_delta[{index}]")
        for index, (region, before) in enumerate(zip(vectors["region_peak_live_bytes"], vectors["live_bytes_before"]))
    ]
    vectors["process_peak_live_bytes_delta"] = [
        require_signed(after - before, f"{label}.allocation.process_peak_live_bytes_delta[{index}]")
        for index, (before, after) in enumerate(zip(vectors["peak_live_bytes_before"], vectors["peak_live_bytes_after"]))
    ]
    return vectors


def validate_report(report: Mapping[str, Any], lane: Mapping[str, Any], label: str) -> Mapping[str, Any]:
    if set(report) != {"schema_version", "tool", "binary_identity", "environment", "configuration", "parallel_metrics", "results", "corpus_catalog"} or report["schema_version"] != 1:
        fail(f"{label}: report envelope differs")
    row = _row(report, lane, label)
    try:
        report_checks.check_tool(dict(report), lane["mode"])
        report_checks.check_binary_identity(dict(report))
        report_checks.check_environment(dict(report), lane["mode"])
        report_checks.check_configuration(dict(report), lane["shape"])
        elapsed, order, _ = report_checks.check_elapsed(row["elapsed_ns"])
        sink = copy.deepcopy(row["sink"])
        buckets = sink["write_size_buckets"]
        if set(buckets) != set(report_checks.SINK_BUCKET_FIELDS):
            fail(f"{label}: sink bucket set differs")
        if sum(require_u64(v, "sink bucket") for v in buckets.values()) != sink["write_calls"]:
            fail(f"{label}: sink buckets do not sum to write calls")
        if not 0 < sink["largest_write"] <= sink["accepted_bytes"] or sink["write_calls"] == 0:
            fail(f"{label}: impossible sink counters")
        sink["buckets"] = buckets
        report_checks.check_operation_metrics(row["operation_metrics"], lane["mode"], elapsed, order, sink)
        report_checks.check_parallel(dict(report), row["corpus"]["archive_sha256"])
    except (report_checks.VerificationError, KeyError, TypeError) as error:
        fail(f"{label}: strict report check: {error}")
    corpus = row["corpus"]
    slide_count = SLIDES[lane["shape"]]
    texts = [f"litchi-perf-pptx-streaming-v1-{i:06}-café-&<>".encode() for i in range(slide_count)]
    input_bytes = sum(map(len, texts))
    semantic = hashlib.sha256(b"litchi-pptx-streaming-semantic-v1\0" + slide_count.to_bytes(8, "little"))
    for index, text in enumerate(texts):
        semantic.update(index.to_bytes(8, "little")); semantic.update(len(text).to_bytes(8, "little")); semantic.update(text)
        for coordinate in (914400, 914400, 7315200, 914400):
            semantic.update(coordinate.to_bytes(8, "little", signed=True))
    full = hashlib.sha256(b"litchi-pptx-streaming-full-text-v1\0" + slide_count.to_bytes(8, "little") + b"".join(texts))
    expected = {"name": f"pptx-streaming-slides-{lane['shape']}", "generator": "litchi-pptx-streaming-plaintext-slides-v1",
                "package_format": "PPTX/OOXML/ZIP", "entry_count": slide_count, "archive_member_count": 37 + 2 * slide_count,
                "entry_bytes": len(texts[0]), "uncompressed_payload_bytes": input_bytes, "target_entry": "ppt/presentation.xml",
                "payload_kind": "deterministic-plain-unicode-xml-significant-text-box-per-slide", "compression": "deflate"}
    for key, value in expected.items():
        if corpus.get(key) != value: fail(f"{label}: corpus {key} differs")
    for key in ("archive_sha256", "target_payload_sha256"): require_hex(corpus.get(key), f"corpus.{key}")
    if row["output_sha256"] != corpus["archive_sha256"]: fail(f"{label}: output does not match corpus")
    expected_sink = {
        "accepted_bytes": corpus["archive_bytes"],
        "input_bytes": input_bytes,
        "authored_part_bytes": corpus["target_payload_bytes"],
        "retained_output_bytes": 0,
    }
    for key, value in expected_sink.items():
        if row["sink"].get(key) != value:
            fail(f"{label}: sink {key} differs")
    if "retained_authoring_window_bytes" in row["sink"]:
        fail(f"{label}: PPTX sink must not publish an authoring-window byte claim")
    source_root = require_object(row.get("source"), f"{label}.source")
    source = require_object(source_root.get("pptx_slides"), f"{label}.source.pptx_slides")
    max_slide_xml_bytes = len(texts[0]) * 5 + 16 * 1024
    expected_source = {
        "role": "streaming",
        "implementation": "litchi_pptx::StreamingPresentationWriter",
        "timing_scope": "fresh deterministic text String generation, public StreamingPresentationWriter creation, forward slide/text-box writes, PPTX package finalization, OPC part-name validation maps, relationship serialization, and writer destruction inside the operation clock; sink setup, materialized corpus construction, semantic reopen, exact 37+2N-member topology gate, graph/geometry/text digest gate, and sink digest finalization are outside",
        "performance_claim": "descriptive timing plus process/RSS and allocator observations for the public plain-text PPTX writer; zero retained sink output and no authoring-window reservation are reported; slide XML byte limits are counters, while ZIP central-directory/member-name storage and OPC part-name validation maps grow with slide count; target presentation-part bytes identify the untimed corpus and are not an all-slide memory bound; no total-RSS, physical-I/O, native-producer, or broad PPTX creation claim",
        "slide_count": slide_count,
        "text_box_count": slide_count,
        "input_text_bytes": input_bytes,
        "authored_part_bytes": corpus["target_payload_bytes"],
        "max_slide_xml_bytes": max_slide_xml_bytes,
        "semantic_sha256": semantic.hexdigest(),
        "full_text_sha256": full.hexdigest(),
        "archive_sha256": corpus["archive_sha256"],
        "target_payload_sha256": corpus["target_payload_sha256"],
        "archive_member_set_verified": True,
        "semantic_reopen_verified": True,
        "deterministic_output_verified": True,
        "structural_metadata_fixed_member_count": 37,
        "structural_metadata_members_per_slide": 2,
        "structural_metadata_scope": "mandatory ZIP package topology and member-name/central-directory metadata; payload, descriptors, and compressor state are runtime work",
        "text_contract": "one plain Unicode text box per slide with deterministic café and XML-significant &<> characters, no title shape, and fixed standard-slide geometry",
    }
    for key, value in expected_source.items():
        if source.get(key) != value:
            fail(f"{label}: source.pptx_slides.{key} differs")
    observed = require_u64(source.get("observed_max_slide_xml_bytes"), f"{label}.source.pptx_slides.observed_max_slide_xml_bytes")
    if observed == 0 or observed > max_slide_xml_bytes:
        fail(f"{label}: observed slide XML bytes exceed the declared finite limit")
    if lane["mode"] == "allocator": allocation_vectors(row, label)
    return row


def _identity(row: Mapping[str, Any], label: str) -> dict[str, Any]:
    sink = require_object(row.get("sink"), f"{label}.sink")
    return {
        "output_sha256": require_hex(row.get("output_sha256"), f"{label}.output_sha256"),
        "sink": dict(sink),
        "corpus": row["corpus"],
        "source": row.get("source"),
        "writer_counters": writer_counters(row, label),
    }


def validate_identity_matrix(rows: Mapping[tuple[str, str, str], Mapping[str, Any]]) -> dict[str, Any]:
    """Require deterministic output/counters within each shape.

    ``rows`` is keyed by ``(mode, repeat, shape)``.  Comparing only the same
    shape avoids treating legitimately different slide counts as a failure.
    """
    identities: dict[str, Any] = {}
    for shape in SHAPES:
        baseline: dict[str, Any] | None = None
        labels: list[str] = []
        for mode in MODES:
            for repeat in REPEATS:
                key = (mode, repeat, shape)
                if key not in rows:
                    fail(f"identity matrix is missing {key!r}")
                label = f"{mode}-{repeat}-{shape}"
                value = _identity(rows[key], label)
                if baseline is None:
                    baseline = value
                elif not canonical_equal(value, baseline):
                    fail(f"deterministic writer identity differs for {shape}")
                labels.append(label)
        assert baseline is not None
        identities[shape] = {"identity": baseline, "lanes": labels, "verified_equal": True}
    return identities


def repeat_summary(
    summaries: Mapping[tuple[str, str, str], Mapping[str, Any]],
    *,
    mode: str | None = None,
    threshold_percent: float = 5.0,
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    modes = (mode,) if mode is not None else MODES
    for selected_mode in modes:
        for shape in SHAPES:
            left = summaries[(selected_mode, "R1", shape)]
            right = summaries[(selected_mode, "R2", shape)]
            deltas = {
                metric: percent_delta(float(left[metric]), float(right[metric]))
                for metric in ("mean", "p50", "p95", "p99")
            }
            result.append({
                "mode": selected_mode,
                "shape": shape,
                "r1": {metric: left[metric] for metric in ("mean", "p50", "p95", "p99")},
                "r2": {metric: right[metric] for metric in ("mean", "p50", "p95", "p99")},
                "deltas_percent_r1_to_r2": deltas,
                "above_threshold": {metric: abs(value) > threshold_percent for metric, value in deltas.items()},
                "threshold_percent": threshold_percent,
            })
    return result


def parse_rss(path: Path, root: Path = ROOT) -> dict[str, Any]:
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise AnalysisError(f"cannot read RSS log {path}: {error}") from error
    marker = "Maximum resident set size (kbytes):"
    values: list[int] = []
    for line in text.splitlines():
        if marker in line:
            value = line.split(":", 1)[1].strip().replace(",", "")
            try:
                parsed = int(value)
            except ValueError as error:
                raise AnalysisError(f"invalid RSS in {path}: {value!r}") from error
            if parsed < 0:
                fail(f"negative RSS in {path}")
            values.append(parsed)
    if len(values) != 1:
        fail(f"{path}: expected exactly one GNU maximum RSS line")
    return {
        "path": binding(path, root),
        "unit": "KiB",
        "maximum_resident_set_kib": values[0],
        "maximum_resident_set_bytes": values[0] * 1024,
        "scope": "whole_process_lifetime_including_setup_and_teardown",
        "latency_comparison": "excluded",
    }


def _load_lane(root: Path, lane: Mapping[str, Any]) -> tuple[Mapping[str, Any], Path]:
    directory = lane_dir(root, lane["lane"])
    report_path = directory / "report.json"
    if not report_path.is_file():
        fail(f"missing {report_path}")
    report = require_object(read_json(report_path), str(report_path))
    validate_report(report, lane, lane["lane"])
    report_checks.check_catalog_sidecar(report_path, dict(report), dict(report["corpus_catalog"]))
    return report, report_path


def build_summary(root: Path = ROOT) -> dict[str, Any]:
    root = root.resolve()
    protocol_path = root / "protocol.json"
    protocol = require_object(read_json(protocol_path), "protocol.json")
    order = protocol_order(protocol)
    rows: dict[tuple[str, str, str], Mapping[str, Any]] = {}
    report_bindings: dict[str, dict[str, Any]] = {}
    rss: dict[str, Any] = {}
    allocation: dict[str, Any] = {}
    elapsed: dict[str, Any] = {}
    for lane in order:
        report, report_path = _load_lane(root, lane)
        key = (lane["mode"], lane["repeat"], lane["shape"])
        rows[key] = require_object(report["results"][0], f"{lane['lane']}.results[0]")
        report_bindings[lane["lane"]] = binding(report_path, root)
        elapsed[lane["lane"]] = stats(rows[key]["elapsed_ns"]["samples"], f"{lane['lane']}.elapsed_ns.samples")
        if lane["mode"] == "allocator":
            vectors = allocation_vectors(rows[key], lane["lane"])
            allocation[lane["lane"]] = {
                field: stats(values, f"{lane['lane']}.allocation.{field}")
                if field not in ("live_bytes_delta", "peak_live_bytes_delta")
                else signed_stats(values, f"{lane['lane']}.allocation.{field}")
                for field, values in vectors.items()
            }
        resource = lane_dir(root, lane["lane"]) / "resource.log"
        if resource.is_file():
            rss[lane["lane"]] = parse_rss(resource, root)
    identities = validate_identity_matrix(rows)
    lane_names = {(
        lane["mode"], lane["repeat"], lane["shape"]
    ): lane["lane"] for lane in order}
    normal_elapsed = {
        ("normal", repeat, shape): elapsed[lane_names[("normal", repeat, shape)]]
        for repeat in REPEATS for shape in SHAPES
    }
    allocator_elapsed = {
        ("allocator", repeat, shape): elapsed[lane_names[("allocator", repeat, shape)]]
        for repeat in REPEATS for shape in SHAPES
    }
    return {
        "schema": SCHEMA,
        "claim": {
            "registered_latency": False,
            "normal_vs_allocator_latency": "excluded",
            "repeat_threshold_percent": 5.0,
            "scope": "fresh PPTX streaming creation; three slide shapes; whole-process captures",
        },
        "protocol": binding(protocol_path, root),
        "normal": {
            "purpose": "uninstrumented fresh-process elapsed distributions",
            "lanes": {lane: elapsed[lane] for lane in elapsed if "normal" in lane},
            "repeat_flags": repeat_summary(normal_elapsed, mode="normal"),
        },
        "allocator": {
            "purpose": "allocator-instrumented fresh-process elapsed distributions; timing is descriptive",
            "lanes": {lane: elapsed[lane] for lane in elapsed if "allocator" in lane},
            "repeat_flags": repeat_summary(allocator_elapsed, mode="allocator"),
            "operation_metrics": allocation,
        },
        "identities": identities,
        "reports": report_bindings,
        "rss": rss,
        "limits": {
            "normal_vs_allocator_timing": "not compared",
            "rss": "whole-process maximum RSS, not operation-local memory",
            "allocator": "requested system-allocator counters aligned to each timed operation sample",
            "writer_counters": "deterministic slide/text-box/XML counters only; no independent performance claim",
            "semantic_window": "bounded input text and public slide XML limit; total heap and ZIP directory growth are not claimed constant",
        },
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--output", type=Path, default=None)
    args = parser.parse_args(argv)
    try:
        root = args.root.resolve()
        result = build_summary(root)
        output = (args.output or root / "summary.json").resolve()
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": "pass", "output": str(output)}))
        return 0
    except (AnalysisError, OSError, TypeError, ValueError, KeyError) as error:
        print(json.dumps({"status": "fail", "error": str(error)}, sort_keys=True))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
