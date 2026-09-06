#!/usr/bin/env python3
"""Independent fixture and report verification for OPC one-Part addition.

Generic report/canonical-catalog checks adapted from change-0439. Archive
inspection uses Python zipfile/XML and an independent payload formula; Rust
refusal gates remain separately identified producer checks.
"""
from __future__ import annotations
import argparse, hashlib, io, json, math, re, struct, sys, zipfile
from pathlib import Path
from typing import Any, Iterable
import xml.etree.ElementTree as ET
ROOT = Path(__file__).resolve().parent.parent
SELECTOR = 'opc_part_add_lifecycle'
SHAPES = {'tiny': 64, 'medium': 1024, 'large': 4096}
ADDED = 'benchmark/added/leaf.bin'
ADDED_TYPE = 'application/vnd.litchi.perf.added-part'
REL_TYPE = 'urn:litchi:perf:relationships:added-part'
CT = '[Content_Types].xml'
RELS = '_rels/.rels'
HEX64 = re.compile(r'^[0-9a-fA-F]{64}$')
class ValidationError(ValueError): pass
def _fail(message: str) -> None:
    raise ValidationError(message)


def _object(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        _fail(f"{label}: expected object")
    return value


def _array(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        _fail(f"{label}: expected array")
    return value


def _string(value: Any, label: str, *, allow_empty: bool = False) -> str:
    if not isinstance(value, str) or (not allow_empty and not value):
        _fail(f"{label}: expected {'non-empty ' if not allow_empty else ''}string")
    return value


def _integer(value: Any, label: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        _fail(f"{label}: expected integer >= {minimum}")
    return value


def _boolean_true(value: Any, label: str) -> None:
    if value is not True:
        _fail(f"{label}: expected true")


def _sha(value: Any, label: str) -> str:
    result = _string(value, label).lower()
    if len(result) != 64 or any(character not in "0123456789abcdef" for character in result):
        _fail(f"{label}: expected lowercase SHA-256")
    return result


def _hash_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _get(mapping: dict[str, Any], key: str, label: str) -> Any:
    if key not in mapping:
        _fail(f"{label}: missing {key}")
    return mapping[key]


def _first(mapping: dict[str, Any], names: Iterable[str], label: str) -> Any:
    for name in names:
        if name in mapping:
            return mapping[name]
    _fail(f"{label}: missing one of {', '.join(names)}")


def _summary_value(summary: dict[str, Any], side: str, name: str, label: str) -> Any:
    """Read the canonical flat field, with nested source/output aliases.

    The final Rust schema uses flat `source_*`/`output_*` fields. The
    nested aliases let the oracle diagnose a draft schema without silently
    treating an absent source or output identity as valid.
    """

    flat_names: list[str]
    if side == "source":
        flat_names = [f"source_{name}"]
    else:
        flat_names = [
            f"output_{name}",
            f"expected_output_{name}",
        ]
        if name == "archive_sha256":
            flat_names.append("expected_output_sha256")
        elif name == "archive_bytes":
            flat_names.append("expected_output_bytes")
        elif name == "content_xml_sha256":
            flat_names.append("expected_output_content_xml_sha256")
        elif name == "content_xml_bytes":
            flat_names.append("expected_output_content_xml_bytes")
    for candidate in flat_names:
        if candidate in summary:
            return summary[candidate]
    nested = summary.get(side)
    if isinstance(nested, dict):
        for candidate in (name, f"{side}_{name}"):
            if candidate in nested:
                return nested[candidate]
    _fail(f"{label}: missing {side} {name}")


def _optional_summary_value(summary: dict[str, Any], side: str, name: str) -> Any:
    names = [f"{side}_{name}"] if side == "source" else [f"output_{name}", f"expected_output_{name}"]
    if side != "source":
        if name == "archive_sha256":
            names.append("expected_output_sha256")
        elif name == "archive_bytes":
            names.append("expected_output_bytes")
        elif name == "content_xml_sha256":
            names.append("expected_output_content_xml_sha256")
        elif name == "content_xml_bytes":
            names.append("expected_output_content_xml_bytes")
    for candidate in names:
        if candidate in summary:
            return summary[candidate]
    nested = summary.get(side)
    if isinstance(nested, dict):
        for candidate in (name, f"{side}_{name}"):
            if candidate in nested:
                return nested[candidate]
    return None


def _gate(summary: dict[str, Any], names: Iterable[str], label: str) -> None:
    for name in names:
        if name in summary:
            _boolean_true(summary[name], f"{label}.{name}")
            return
    _fail(f"{label}: missing one of {', '.join(names)}")


def _summary_field(summary: dict[str, Any], names: Iterable[str], label: str) -> Any:
    for name in names:
        if name in summary:
            return summary[name]
    _fail(f"{label}: missing one of {', '.join(names)}")



from fixtures import verify_fixture
def _validate_catalog(report: dict[str, Any], report_path: Path, corpus: dict[str, Any]) -> None:
    reference = _object(_get(report, "corpus_catalog", "report"), "report.corpus_catalog")
    if _integer(_get(reference, "manifest_version", "report.corpus_catalog"), "report.corpus_catalog.manifest_version") != 2:
        _fail("report.corpus_catalog.manifest_version must be 2")
    if _string(_get(reference, "catalog_id", "report.corpus_catalog"), "report.corpus_catalog.catalog_id") != "litchi-perf-corpus-v2":
        _fail("report.corpus_catalog.catalog_id is unexpected")
    reference_catalog_sha = _sha(_get(reference, "catalog_sha256", "report.corpus_catalog"), "report.corpus_catalog.catalog_sha256")
    reference_content_sha = _sha(_get(reference, "content_set_sha256", "report.corpus_catalog"), "report.corpus_catalog.content_set_sha256")
    candidates = [
        report_path.with_name(report_path.stem + "-catalog.json"),
        report_path.with_name("catalog.json"),
    ]
    sidecar = next((candidate for candidate in candidates if candidate.is_file()), None)
    if sidecar is None:
        _fail("corpus catalog sidecar is missing beside report")
    try:
        catalog = json.loads(sidecar.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        _fail(f"catalog sidecar is invalid JSON: {error}")
    catalog = _object(catalog, "catalog")
    if _integer(_get(catalog, "manifest_version", "catalog"), "catalog.manifest_version") != 2:
        _fail("catalog.manifest_version must be 2")
    if _string(_get(catalog, "manifest_kind", "catalog"), "catalog.manifest_kind") != "corpus-catalog":
        _fail("catalog.manifest_kind is unexpected")
    canonicalization = _object(_get(catalog, "canonicalization", "catalog"), "catalog.canonicalization")
    if canonicalization != {"algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"}:
        _fail("catalog canonicalization is unsupported")
    if _string(_get(catalog, "catalog_id", "catalog"), "catalog.catalog_id") != reference["catalog_id"]:
        _fail("catalog ID does not bind to report reference")
    corpora = _array(_get(catalog, "corpora", "catalog"), "catalog.corpora")
    bindings = _array(_get(catalog, "case_bindings", "catalog"), "catalog.case_bindings")
    if len(corpora) == 0 or len(bindings) == 0:
        _fail("catalog has no corpora or bindings")
    ids = [
        _string(_get(_object(item, f"catalog.corpora[{index}]"), "id", f"catalog.corpora[{index}]"), f"catalog.corpora[{index}].id")
        for index, item in enumerate(corpora)
    ]
    if ids != sorted(ids) or len(set(ids)) != len(ids):
        _fail("catalog corpus IDs are not unique and sorted")
    relevant = []
    for index, raw in enumerate(corpora):
        item = _object(raw, f"catalog.corpora[{index}]")
        legacy = _object(_get(item, "legacy_v1", f"catalog.corpora[{index}]"), f"catalog.corpora[{index}].legacy_v1")
        if legacy.get("archive_sha256") == corpus.get("archive_sha256"):
            relevant.append(item)
    if len(relevant) != 1:
        _fail("catalog does not contain exactly one source corpus binding")
    item = relevant[0]
    bytes_summary = _object(_get(item, "bytes", "catalog corpus"), "catalog corpus.bytes")
    if _sha(_get(bytes_summary, "archive_sha256", "catalog corpus.bytes"), "catalog corpus.bytes.archive_sha256") != _sha(corpus["archive_sha256"], "corpus.archive_sha256"):
        _fail("catalog corpus archive SHA does not match report corpus")
    if _integer(_get(bytes_summary, "archive_bytes", "catalog corpus.bytes"), "catalog corpus.bytes.archive_bytes") != _integer(corpus["archive_bytes"], "corpus.archive_bytes"):
        _fail("catalog corpus archive byte count does not match report corpus")
    if _integer(_get(bytes_summary, "logical_payload_bytes", "catalog corpus.bytes"), "catalog corpus.bytes.logical_payload_bytes") != _integer(corpus["uncompressed_payload_bytes"], "corpus.uncompressed_payload_bytes"):
        _fail("catalog corpus logical payload bytes do not match the canonical projection")
    legacy = _object(_get(item, "legacy_v1", "catalog corpus"), "catalog corpus.legacy_v1")
    for key in ("name", "generator", "package_format", "shape", "payload_kind", "compression", "entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "target_entry", "target_payload_bytes", "target_payload_sha256"):
        if legacy.get(key) != corpus.get(key):
            _fail(f"catalog legacy field {key} does not match report corpus")
    member_set = _object(_get(item, "members", "catalog corpus"), "catalog corpus.members")
    member_status = _string(_get(member_set, "status", "catalog corpus.members"), "catalog corpus.members.status")
    member_items = _array(_get(member_set, "items", "catalog corpus.members"), "catalog corpus.members.items")
    if member_status == 'unavailable':
        if member_items: _fail('unavailable catalog must not contain member claims')
    elif member_status == 'complete':
        archive = ROOT / 'fixtures' / (corpus['shape'] + '-source.zip')
        with zipfile.ZipFile(archive) as source_zip:
            expected = source_zip.namelist()
            if len(member_items) != len(expected):
                _fail('catalog member count differs')
            for ordinal, (member, name) in enumerate(zip(member_items, expected)):
                if member['ordinal'] != ordinal or member['name'] != name or member['sha256'] != _hash_bytes(source_zip.read(name)):
                    _fail('catalog member identity/order differs')
    else:
        _fail('unsupported catalog member status')
    matching_bindings = []
    for index, raw in enumerate(bindings):
        binding = _object(raw, f"catalog.case_bindings[{index}]")
        if binding.get("case") == SELECTOR:
            matching_bindings.append(binding)
    if len(matching_bindings) != 1:
        _fail("catalog must have exactly one append selector binding")
    binding = matching_bindings[0]
    if binding.get("corpus_id") != item.get("id") or binding.get("role") != "timed":
        _fail("append selector catalog binding does not identify the source corpus")
    if binding.get("legacy_archive_sha256") != corpus.get("archive_sha256"):
        _fail("append selector legacy archive binding differs")
    catalog_for_hash = json.loads(json.dumps(catalog))
    catalog_for_hash.pop("catalog_sha256", None)
    computed_catalog_sha = _hash_bytes(_canonical_json(catalog_for_hash))
    if computed_catalog_sha != reference_catalog_sha or computed_catalog_sha != catalog.get("catalog_sha256"):
        _fail("catalog SHA does not match sorted canonical JSON")
    content_value = {
        "corpora": [
            {
                "id": entry.get("id"),
                "archive_sha256": _object(entry.get("bytes"), "catalog corpus.bytes").get("archive_sha256"),
                "members": [
                    {
                        "ordinal": member.get("ordinal"),
                        "name": member.get("name"),
                        "sha256": member.get("sha256"),
                    }
                    for member in _object(entry.get("members"), "catalog corpus.members").get("items", [])
                ],
            }
            for entry in corpora
        ],
        "case_bindings": [
            {"case": entry.get("case"), "corpus_id": entry.get("corpus_id"), "role": entry.get("role")}
            for entry in bindings
        ],
    }
    if _hash_bytes(_canonical_json(content_value)) != reference_content_sha or _hash_bytes(_canonical_json(content_value)) != catalog.get("content_set_sha256"):
        _fail("catalog content-set SHA does not match sorted canonical JSON")


def _canonical_json(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def _metric_vector(value: Any, label: str, samples: int) -> tuple[str, list[int] | None]:
    vector = _object(value, label)
    status = _string(_get(vector, "status", label), f"{label}.status")
    values = vector.get("values")
    if status == "measured":
        values = _array(values, f"{label}.values")
        if len(values) != samples:
            _fail(f"{label}.values has {len(values)} values, expected {samples}")
        checked = [_integer(item, f"{label}.values[{index}]") for index, item in enumerate(values)]
        return status, checked
    if status not in {'unavailable', 'not_applicable'}:
        _fail(f'{label}: unsupported metric status')
    if values is not None:
        _fail(f"{label}.values must be absent for {status}")
    return status, None


def _walk_metric_vectors(value: Any, label: str, samples: int) -> list[tuple[str, list[int] | None, str]]:
    found: list[tuple[str, list[int] | None, str]] = []
    if isinstance(value, dict):
        if "status" in value and "scope" in value and ("values" in value or set(value).issubset({"status", "scope"})):
            status, values = _metric_vector(value, label, samples)
            found.append((label, values, status))
        for key, child in value.items():
            if key not in {"values", "status", "scope"}:
                found.extend(_walk_metric_vectors(child, f"{label}.{key}", samples))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            found.extend(_walk_metric_vectors(child, f"{label}[{index}]", samples))
    return found


def _validate_metrics(result: dict[str, Any], summary: dict[str, Any], fixture: dict[str, Any], mode: str, samples: int) -> None:
    elapsed = _object(_get(result, "elapsed_ns", "result"), "result.elapsed_ns")
    if _string(_get(elapsed, "unit", "result.elapsed_ns"), "result.elapsed_ns.unit") != "ns":
        _fail("elapsed_ns unit is not ns")
    elapsed_values = [_integer(item, f"result.elapsed_ns.samples[{index}]", minimum=1) for index, item in enumerate(_array(_get(elapsed, "samples", "result.elapsed_ns"), "result.elapsed_ns.samples"))]
    if len(elapsed_values) != samples or elapsed_values != sorted(elapsed_values):
        _fail("elapsed_ns.samples must be a sorted vector of the requested length")
    sample_order = [_integer(item, f"result.elapsed_ns.sample_order[{index}]") for index, item in enumerate(_array(_get(elapsed, "sample_order", "result.elapsed_ns"), "result.elapsed_ns.sample_order"))]
    if len(sample_order) != samples or sorted(sample_order) != list(range(samples)):
        _fail("elapsed_ns.sample_order must be a permutation of retained sample indices")
    for key in ("min", "p50", "p95", "p99", "max"):
        _integer(_get(elapsed, key, "result.elapsed_ns"), f"result.elapsed_ns.{key}", minimum=1)
    for key in ("mean", "standard_deviation"):
        value = _get(elapsed, key, "result.elapsed_ns")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)) or float(value) < 0:
            _fail(f"result.elapsed_ns.{key}: expected finite non-negative number")
    interval = _object(_get(elapsed, "confidence_interval_95", "result.elapsed_ns"), "result.elapsed_ns.confidence_interval_95")
    if not isinstance(_get(interval, "method", "confidence interval"), str):
        _fail("confidence interval method is missing")
    for key in ("lower", "upper"):
        value = _get(interval, key, "confidence interval")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)) or float(value) < 0:
            _fail(f"confidence interval {key} is invalid")

    lifecycle = _array(_get(summary, "lifecycle_ns", "opc_part_add"), "opc_part_add.lifecycle_ns")
    lifecycle_values = [_integer(item, f"opc_part_add.lifecycle_ns[{index}]", minimum=1) for index, item in enumerate(lifecycle)]
    if len(lifecycle_values) != samples or sorted(lifecycle_values) != elapsed_values:
        _fail("opc_part_add.lifecycle_ns does not bind to elapsed_ns.samples")
    if [lifecycle_values[index] for index in sample_order] != elapsed_values:
        _fail("lifecycle_ns does not bind to original sample indices")
    output_vector = [_sha(item, f"opc_part_add.output_sha256[{index}]") for index, item in enumerate(_array(_get(summary, "output_sha256", "opc_part_add"), "opc_part_add.output_sha256"))]
    expected_output_sha = _sha(_summary_value(summary, "output", "archive_sha256", "opc_part_add"), "opc_part_add.expected_output_sha256")
    if len(output_vector) != samples or any(item != expected_output_sha for item in output_vector):
        _fail("opc_part_add.output_sha256 is not a stable retained output vector")
    if _sha(_get(result, "output_sha256", "result"), "result.output_sha256") != expected_output_sha:
        _fail("result.output_sha256 does not bind to expected output")

    metrics = _object(_get(result, "operation_metrics", "result"), "result.operation_metrics")
    if _integer(_get(metrics, "sample_count", "operation_metrics"), "operation_metrics.sample_count") != samples:
        _fail("operation_metrics.sample_count does not match samples")
    if _string(_get(metrics, "alignment", "operation_metrics"), "operation_metrics.alignment") != "elapsed_ns.samples_by_elapsed_then_sample_index":
        _fail("operation_metrics alignment is not the declared elapsed/sample-index order")
    metrics_indices = [_integer(item, f"operation_metrics.sample_indices[{index}]") for index, item in enumerate(_array(_get(metrics, "sample_indices", "operation_metrics"), "operation_metrics.sample_indices"))]
    if metrics_indices != sample_order:
        _fail("operation_metrics.sample_indices do not match elapsed_ns.sample_order")
    latency_claim = _string(_get(metrics, "latency_claim", "operation_metrics"), "operation_metrics.latency_claim").lower()
    if "physical" in latency_claim or "cold" in latency_claim or "source_backed" in latency_claim:
        _fail("operation_metrics latency claim overstates the owned lifecycle scope")
    vectors = _walk_metric_vectors(metrics, "operation_metrics", samples)
    if not vectors:
        _fail("operation_metrics contains no metric vectors")
    source_metrics = _object(_get(metrics, "source", "operation_metrics"), "operation_metrics.source")
    if source_metrics['status'] != 'measured' or source_metrics['counter_scope'] != 'in_process_instrumented_source_read_at':
        _fail('instrumented source metrics must be measured')
    source = result['source']
    for metric, counter in [('logical_read_calls','read_calls'), ('logical_read_returned_bytes','read_bytes'), ('max_concurrent_reads','max_in_flight_reads')]:
        status, values = _metric_vector(source_metrics[metric], metric, samples)
        if status != 'measured' or values != [source[counter][i] for i in sample_order]:
            _fail('source metrics differ from chronological observations')
    for key in ('logical_read_requested_bytes','logical_read_largest_requested_bytes','logical_read_largest_returned_bytes','compressed_bytes','decompressed_bytes','recompressed_bytes'):
        if source_metrics[key]['status'] != 'unavailable':
            _fail('unobserved source byte flow must be unavailable')
    for section_name in ("publication", "materialization", "cfb_phases"):
        section = _object(_get(metrics, section_name, "operation_metrics"), f"operation_metrics.{section_name}")
        if _string(_get(section, "status", f"operation_metrics.{section_name}"), f"operation_metrics.{section_name}.status") != "not_applicable":
            _fail(f"operation_metrics.{section_name} must be not_applicable")
    sink_metrics = _object(_get(metrics, "sink", "operation_metrics"), "operation_metrics.sink")
    if _string(_get(sink_metrics, "write_status", "operation_metrics.sink"), "operation_metrics.sink.write_status") != "measured":
        _fail("operation_metrics sink write status must be measured")
    accepted_status, accepted = _metric_vector(_get(sink_metrics, "accepted_bytes", "operation_metrics.sink"), "operation_metrics.sink.accepted_bytes", samples)
    expected_output_bytes = _integer(_summary_value(summary, "output", "archive_bytes", "opc_part_add"), "opc_part_add.expected_output_bytes")
    if accepted_status != "measured" or accepted is None or any(value != expected_output_bytes for value in accepted):
        _fail("operation_metrics sink accepted bytes do not equal committed output bytes")
    _, write_calls = _metric_vector(_get(sink_metrics, "write_calls", "operation_metrics.sink"), "operation_metrics.sink.write_calls", samples)
    _, largest = _metric_vector(_get(sink_metrics, "largest_write", "operation_metrics.sink"), "operation_metrics.sink.largest_write", samples)
    sink = _object(_get(result, "sink", "result"), "result.sink")
    sink_accepted = _integer(_get(sink, "accepted_bytes", "result.sink"), "result.sink.accepted_bytes")
    sink_calls = _integer(_get(sink, "write_calls", "result.sink"), "result.sink.write_calls")
    sink_largest = _integer(_get(sink, "largest_write", "result.sink"), "result.sink.largest_write")
    if sink_accepted != expected_output_bytes or sink_calls <= 0 or sink_largest <= 0:
        _fail("result.sink does not describe the committed output write")
    if write_calls is None or any(value != sink_calls for value in write_calls) or largest is None or any(value != sink_largest for value in largest):
        _fail("operation_metrics sink counters do not bind to result.sink")
    buckets = _object(_get(sink, "write_size_buckets", "result.sink"), "result.sink.write_size_buckets")
    bucket_total = 0
    for key in ("bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536"):
        bucket_total += _integer(_get(buckets, key, "result.sink.write_size_buckets"), f"result.sink.write_size_buckets.{key}")
    if bucket_total != sink_calls:
        _fail("result.sink write buckets do not sum to write_calls")
    if sink.get("retained_output_bytes") != 0:
        _fail("append sink must report zero retained output bytes")
    if sink.get("retained_authoring_window_bytes") is not None:
        _fail("append sink must not report an authoring window")

    allocation = metrics.get("allocation")
    if mode == "allocator":
        if not isinstance(allocation, dict) or allocation.get("status") != "measured":
            _fail("allocator report must contain measured allocation metrics")
        allocator_vectors = {}
        for key in ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes"):
            status, values = _metric_vector(_get(allocation, key, "operation_metrics.allocation"), f"operation_metrics.allocation.{key}", samples)
            if status != "measured" or values is None:
                _fail(f"allocator vector {key} is not measured")
            allocator_vectors[key] = values
        for index in range(samples):
            values = {key: vector[index] for key, vector in allocator_vectors.items()}
            if values['failed_allocation_calls'] != 0:
                _fail('allocator failure calls must be zero')
            if values['live_bytes_before'] + values['allocated_bytes'] - values['deallocated_bytes'] != values['live_bytes_after']:
                _fail('allocator live-byte balance differs')
            if values['region_peak_live_bytes'] < max(values['live_bytes_before'], values['live_bytes_after']):
                _fail('allocator region peak is below an endpoint')
        if allocation.get("scope") != "operation_global_system_allocator":
            _fail("allocator scope is not operation_global_system_allocator")
    elif allocation is not None:
        if not isinstance(allocation, dict) or allocation.get("status") == "measured":
            _fail("normal report must not publish measured allocator metrics")


def _validate_identity(report: dict[str, Any], mode: str) -> None:
    tool = _object(_get(report, 'tool', 'report'), 'tool')
    expected_binary = 'litchi-perf-baseline' + ('-alloc' if mode == 'allocator' else '')
    if tool.get('name') != 'litchi-perf-baseline' or tool.get('binary') != expected_binary or tool.get('profile') != 'release':
        _fail('tool binary/profile identity differs')
    expected_instrumentation = 'system_allocator_operation_scoped' if mode == 'allocator' else 'none'
    if tool.get('instrumentation') != expected_instrumentation:
        _fail('tool instrumentation differs from requested mode')
    if tool.get('allocator_counter_revision') != ('serialized_region_peak_v3' if mode == 'allocator' else None):
        _fail('allocator counter revision differs')
    identity = _object(_get(report, 'binary_identity', 'report'), 'binary_identity')
    _sha(identity.get('binary_sha256'), 'binary_identity.binary_sha256')
    _integer(identity.get('binary_bytes'), 'binary_identity.binary_bytes', minimum=1)
    if identity.get('profile') != 'release' or identity.get('executable') is not True:
        _fail('binary identity must describe a release executable')
    environment = _object(_get(report, 'environment', 'report'), 'environment')
    if environment.get('rustc_version') != 'rustc 1.98.1 (48a229cea 2026-09-01)' or environment.get('rustflags') != '-Cforce-frame-pointers=yes':
        _fail('toolchain or frame-pointer build flags differ')
    if environment.get('cpu_affinity') != '2' or environment.get('logical_cpus_available') != 1:
        _fail('CPU 2 single-worker affinity differs')
    expected_allocator = 'CountingSystemAllocator(std::alloc::System)' if mode == 'allocator' else 'Rust system allocator'
    if environment.get('allocator') != expected_allocator:
        _fail('allocator identity differs from mode')


def validate_report(path: str | Path, mode: str, shape: str, *, samples: int = 30, warmups: int = 3) -> dict[str, Any]:
    """Validate one serialized 0444 report and return its decoded object."""

    if mode not in {"normal", "allocator"}:
        _fail("mode must be normal or allocator")
    if shape not in SHAPES:
        _fail("shape must be tiny, medium, or large")
    if isinstance(samples, bool) or not isinstance(samples, int) or samples <= 0:
        _fail("samples must be a positive integer")
    if isinstance(warmups, bool) or not isinstance(warmups, int) or warmups < 0:
        _fail("warmups must be a non-negative integer")
    report_path = Path(path)
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        _fail(f"report is invalid JSON: {error}")
    report = _object(report, "report")
    _validate_identity(report, mode)
    if _integer(_get(report, "schema_version", "report"), "report.schema_version") != 1:
        _fail("report.schema_version must be 1")
    configuration = _object(_get(report, "configuration", "report"), "report.configuration")
    if _integer(_get(configuration, "samples_per_case", "configuration"), "configuration.samples_per_case") != samples:
        _fail("configuration.samples_per_case differs from --samples")
    if _integer(_get(configuration, "warmup_iterations_per_case", "configuration"), "configuration.warmup_iterations_per_case") != warmups:
        _fail("configuration.warmup_iterations_per_case differs from --warmups")
    if _get(configuration, "cases", "configuration") != [SELECTOR]:
        _fail("configuration.cases must contain only the append selector")
    if _get(configuration, "semantic_shapes", "configuration") != [shape]:
        _fail("configuration.semantic_shapes must contain only --shape")
    if _get(configuration, "execution_workers", "configuration") != [1]:
        _fail("configuration.execution_workers must be [1]")
    results = _array(_get(report, "results", "report"), "report.results")
    if len(results) != 1:
        _fail("one report invocation must contain exactly one result")
    result = _object(results[0], "report.results[0]")
    if _string(_get(result, "case", "result"), "result.case") != SELECTOR:
        _fail("result.case is not the append selector")
    corpus = _object(_get(result, "corpus", "result"), "result.corpus")
    source = _object(_get(result, "source", "result"), "result.source")
    summary = _object(_get(source, "opc_part_add", "result.source"), "result.source.opc_part_add")
    fixture = verify_fixture(shape)
    expected_summary = json.loads((ROOT / 'fixtures' / (shape + '.json')).read_text())
    for key, value in expected_summary.items():
        if key not in {'lifecycle_ns', 'output_sha256'} and summary.get(key) != value:
            _fail('static fixture summary differs: ' + key)
    for key in ('read_calls','read_bytes','ordinary_payload_read_calls','ordinary_payload_read_bytes','max_in_flight_reads'):
        values = source.get(key)
        if not isinstance(values, list) or len(values) != samples or any(type(v) is not int or v <= 0 for v in values):
            _fail('source observation missing or invalid: ' + key)
    if any(v != 1 for v in source['max_in_flight_reads']):
        _fail('serial source exceeded one read in flight')
    if source.get('ordinary_payload_materializations') is not None:
        _fail('Part materialization count is not measured here')
    for key, value in {'entry_count':SHAPES[shape], 'archive_member_count':SHAPES[shape]+2,
                       'entry_bytes':1024, 'uncompressed_payload_bytes':SHAPES[shape]*1024,
                       'archive_bytes':fixture['source_bytes'], 'archive_sha256':fixture['source_sha'],
                       'generator':'litchi-opc-part-add-lifecycle-v1', 'shape':shape,
                       'package_format':'OPC/ZIP'}.items():
        if corpus.get(key) != value: _fail('corpus differs: ' + key)
    _validate_catalog(report, report_path, corpus)
    _validate_metrics(result, summary, fixture, mode, samples)
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--mode", choices=("normal", "allocator"), required=True)
    parser.add_argument("--shape", choices=tuple(SHAPES), required=True)
    parser.add_argument("--samples", type=int, default=30)
    parser.add_argument("--warmups", type=int, default=3)
    # Evidence capture passes this role field; 0444 has one current-revision
    # role, so it is checked for compatibility and never used to relax gates.
    parser.add_argument("--role", default="after")
    args = parser.parse_args(argv)
    if args.role != "after":
        print("INVALID: only the current after role is supported", file=sys.stderr)
        return 2
    try:
        validate_report(args.report, args.mode, args.shape, samples=args.samples, warmups=args.warmups)
    except (OSError, TypeError, ValueError, ValidationError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 2
    print("VALID")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
