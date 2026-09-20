#!/usr/bin/env python3
"""Independently verify the 0709 DOCX ordinary-save baseline packet.

The analyzer consumes only the frozen plan, build/source manifests, fixture
identity, child receipts and child JSON reports.  It recomputes elapsed
statistics from the retained sample vectors, checks every measured metric
vector's cardinality, proves native/allocator semantic parity, and emits
descriptive repeat and phase comparisons.  It intentionally contains no
speedup gate: the packet is a current-source baseline refresh.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import math
from pathlib import Path
import statistics
from typing import Any, Iterable
import zipfile


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PHASES = ("lifecycle", "edit", "atomic_publish", "counting_publish")
PHASE_LABELS = {
    "lifecycle": "open+edit+save",
    "edit": "edit",
    "atomic_publish": "save-to-path",
    "counting_publish": "serialize-to-counting-sink",
}
PHASE_TIMING_SCOPES = {
    "lifecycle": "open the path through the documented reader, make one semantic edit, save to a path; destination preparation, readback, digest and cleanup are outside the clock",
    "edit": "the semantic edit and its commit only; the documented open and every verification are outside the clock",
    "atomic_publish": "one documented save-to-path only: sibling creation in the destination directory, the publication write, permission preservation, the temporary's data sync, the rename that replaces the destination, and the parent-directory sync; the open, the edit, the readback and the cleanup are outside the clock",
    "counting_publish": "the documented sequential serialization into a bounded counting sink only; the open, the edit and the byte accounting are outside the clock",
}
EDIT_MARKER = 'Package::document_mut().add_paragraph_with_text("litchi-perf-0638-ordinary-save")'
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
)
HEX = set("0123456789abcdef")


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def json_digest(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        fail(f"invalid JSON in {path}: {error}")


def write(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n")


def check_hex(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def close(left: float, right: float, label: str, tolerance: float = 1e-9) -> None:
    require(math.isclose(left, right, rel_tol=tolerance, abs_tol=1e-6),
            f"{label}: {left!r} != {right!r}")


def midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + ((left % 2 + right % 2) // 2)


def nearest_rank(values: list[int], percentile: int) -> int:
    index = ((percentile * len(values) + 99) // 100) - 1
    return values[min(index, len(values) - 1)]


def student_t_critical_95(degrees_of_freedom: int) -> float:
    values = [
        12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306,
        2.262, 2.228, 2.201, 2.179, 2.160, 2.145, 2.131, 2.120,
        2.110, 2.101, 2.093, 2.086, 2.080, 2.074, 2.069, 2.064,
        2.060, 2.056, 2.052, 2.048, 2.045, 2.042,
    ]
    if degrees_of_freedom == 0:
        return 0.0
    if degrees_of_freedom <= len(values):
        return values[degrees_of_freedom - 1]
    z = 1.959963984540054
    degrees = float(degrees_of_freedom)
    z2 = z * z
    z3 = z2 * z
    z5 = z3 * z2
    z7 = z5 * z2
    return (
        z + (z3 + z) / (4.0 * degrees)
        + (5.0 * z5 + 16.0 * z3 + 3.0 * z) / (96.0 * degrees * degrees)
        + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z)
        / (384.0 * degrees * degrees * degrees)
    )


def elapsed_stats(values: Iterable[int]) -> dict[str, Any]:
    values = list(values)
    require(values, "elapsed vector is empty")
    for index, value in enumerate(values):
        require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
                f"elapsed[{index}] is not a positive integer")
    ordered = sorted(values)
    mean = statistics.mean(values)
    deviation = statistics.stdev(values) if len(values) > 1 else 0.0
    margin = student_t_critical_95(len(values) - 1) * deviation / math.sqrt(len(values))
    return {
        "count": len(values),
        "min": ordered[0],
        "p50": midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
        "p95": nearest_rank(ordered, 95),
        "p99": nearest_rank(ordered, 99),
        "max": ordered[-1],
        "mean": mean,
        "standard_deviation": deviation,
        "confidence_interval_95": {
            "method": "two-sided Student's t interval for the mean",
            "lower": max(0.0, mean - margin),
            "upper": mean + margin,
        },
    }


def integer_stats(values: Iterable[int], label: str) -> dict[str, Any]:
    """Describe an integer metric vector without treating it as elapsed time."""

    values = list(values)
    require(values, f"{label} is empty")
    for index, value in enumerate(values):
        require(isinstance(value, int) and not isinstance(value, bool),
                f"{label}[{index}] is not an integer")
    ordered = sorted(values)
    return {
        "count": len(values),
        "min": ordered[0],
        "p50": midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
        "p95": nearest_rank(ordered, 95),
        "p99": nearest_rank(ordered, 99),
        "max": ordered[-1],
        "mean": statistics.mean(values),
    }


def validate_elapsed(value: Any, samples: int, label: str) -> tuple[list[int], dict[str, Any]]:
    require(isinstance(value, dict), f"{label}.elapsed_ns is missing")
    require(value.get("unit") == "ns", f"{label}.elapsed_ns unit changed")
    vector = value.get("samples")
    require(isinstance(vector, list) and len(vector) == samples,
            f"{label}.elapsed_ns.samples length changed")
    require(vector == sorted(vector), f"{label}.elapsed_ns.samples are not sorted")
    order = value.get("sample_order")
    require(isinstance(order, list) and sorted(order) == list(range(samples)),
            f"{label}.elapsed_ns.sample_order is not a permutation")
    expected = elapsed_stats(vector)
    for key in ("count", "min", "p50", "p95", "p99", "max"):
        if key == "count":
            continue
        require(value.get(key) == expected[key], f"{label}.elapsed_ns.{key} is stale")
    close(float(value.get("mean")), float(expected["mean"]), f"{label}.elapsed_ns.mean")
    close(float(value.get("standard_deviation")), float(expected["standard_deviation"]),
          f"{label}.elapsed_ns.standard_deviation")
    interval = value.get("confidence_interval_95")
    require(isinstance(interval, dict), f"{label} confidence interval is missing")
    require(interval.get("method") == expected["confidence_interval_95"]["method"],
            f"{label} confidence interval method changed")
    finite(interval.get("lower"), f"{label} confidence interval lower")
    finite(interval.get("upper"), f"{label} confidence interval upper")
    close(float(interval["lower"]), float(expected["confidence_interval_95"]["lower"]),
          f"{label} confidence interval lower", tolerance=1e-8)
    close(float(interval["upper"]), float(expected["confidence_interval_95"]["upper"]),
          f"{label} confidence interval upper", tolerance=1e-8)
    return vector, expected


def source_census() -> dict[str, str]:
    paths: list[Path] = [REPO / "Cargo.toml", REPO / "Cargo.lock"]
    paths.extend(path for path in (REPO / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path
            for path in (REPO / folder).rglob("*")
            if path.is_file() and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {str(path.relative_to(REPO)): sha(path) for path in sorted(set(paths))}


def load_plan() -> dict[str, Any]:
    p = read(HERE / "plan.json")
    require(isinstance(p, dict), "plan is not an object")
    require(p.get("native") == {"repeats": 3, "samples": 100, "warmup": 10},
            "native plan changed")
    require(p.get("allocator") == {"repeats": 2, "samples": 3, "warmup": 0},
            "allocator plan changed")
    order_description = p.get("order", "")
    require(isinstance(order_description, str)
            and "repeat 1" in order_description
            and "repeat 2" in order_description
            and "reverse" in order_description,
            "repeat order description is missing")
    require(p.get("review_threshold_percent") == 5,
            "repeat review threshold changed")
    require(p.get("phase_order") in (None, list(PHASES)), "phase order changed")
    corpora = p.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 3, "corpus count changed")
    ids: set[str] = set()
    generated = real_admitted = real_refusal = 0
    for corpus in corpora:
        require(isinstance(corpus, dict), "corpus entry is malformed")
        identity = corpus.get("id")
        require(isinstance(identity, str) and identity and identity not in ids,
                "corpus ids are not unique")
        ids.add(identity)
        origin = corpus.get("origin")
        admitted = corpus.get("expected_edit_admitted")
        require(isinstance(admitted, bool), f"{identity}: expected edit outcome missing")
        if origin == "generated-harness-corpus":
            generated += 1
            require(corpus.get("path") is None and corpus.get("sha256") is None,
                    f"{identity}: generated corpus has fixture binding")
        elif origin == "caller-named-real-file":
            path = corpus.get("path")
            digest = corpus.get("sha256")
            require(isinstance(path, str) and path and isinstance(digest, str),
                    f"{identity}: real fixture binding is missing")
            check_hex(digest, f"{identity}.sha256")
            if admitted:
                real_admitted += 1
            else:
                real_refusal += 1
        else:
            fail(f"{identity}: unsupported corpus origin {origin!r}")
    require((generated, real_admitted, real_refusal) == (1, 1, 1),
            "plan must contain generated, admitted-real, and refusal-real DOCX corpora")
    return p


def normalized_corpus(corpus: dict[str, Any]) -> dict[str, Any]:
    value = dict(corpus)
    fixture = value.get("fixture")
    if isinstance(fixture, dict):
        value.setdefault("path", fixture.get("path"))
        value.setdefault("sha256", fixture.get("sha256"))
    return value


def corpus_fixture_binding(corpus: dict[str, Any]) -> dict[str, Any] | None:
    corpus = normalized_corpus(corpus)
    if corpus["origin"] == "generated-harness-corpus":
        return None
    raw = corpus["path"]
    path = (REPO / raw).resolve() if not Path(raw).is_absolute() else Path(raw).resolve()
    require(path.is_file() and not path.is_symlink(), f"missing fixture {path}")
    digest = sha(path)
    require(digest == corpus["sha256"], f"fixture digest changed for {corpus['id']}")
    if corpus.get("bytes") is not None:
        require(path.stat().st_size == corpus["bytes"],
                f"fixture byte count changed for {corpus['id']}")
    return {"plan_path": raw, "resolved_path": str(path), "bytes": path.stat().st_size,
            "sha256": digest}


def cleanup_witnesses() -> list[dict[str, Any]]:
    witnesses: list[dict[str, Any]] = []
    for filename in ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json"):
        path = HERE / filename
        if not path.is_file():
            continue
        value = read(path)

        def walk(item: Any) -> None:
            if isinstance(item, dict):
                raw_path = item.get("path")
                digest = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes"))
                if isinstance(raw_path, str) and isinstance(digest, str):
                    check_hex(digest, f"{filename}:{raw_path}")
                    witnesses.append({"path": raw_path, "sha256": digest, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return witnesses


def build_info(lane: str) -> dict[str, Any]:
    records_path = HERE / "build-baseline.json"
    records = read(records_path)
    require(isinstance(records, list), "build-baseline.json is not a list")
    binary_name = f"baseline-{lane}"
    matches = [item for item in records if isinstance(item, dict)
               and Path(str(item.get("binary", ""))).name == binary_name]
    require(len(matches) == 1, f"build record for {binary_name} is not unique")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{binary_name} build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    digest = record.get("binary_sha256")
    size = record.get("binary_bytes")
    check_hex(digest, f"{binary_name}.binary_sha256")
    require(isinstance(size, int) and size > 0, f"{binary_name}.binary_bytes is invalid")
    if binary.is_file() and not binary.is_symlink():
        require(sha(binary) == digest and binary.stat().st_size == size,
                f"{binary_name} live binary identity changed")
    else:
        matched = False
        target = str(binary)
        for witness in cleanup_witnesses():
            candidate = Path(witness["path"])
            candidate = ((REPO / candidate).resolve() if not candidate.is_absolute()
                         else candidate.resolve())
            if str(candidate) == target and witness["sha256"] == digest \
                    and (witness.get("bytes") is None or witness["bytes"] == size):
                matched = True
                break
        require(matched, f"{binary_name} missing without an exact cleanup witness")
    manifest = HERE / "source-baseline.json"
    source = read(manifest)
    require(isinstance(source, dict) and source, "source-baseline.json is invalid")
    require(record.get("source_manifest_sha256") == sha(manifest),
            f"{binary_name} source manifest binding changed")
    return {
        "name": binary_name,
        "binary": str(binary),
        "binary_sha256": digest,
        "binary_bytes": size,
        "build_record_sha256": sha(records_path),
        "source_manifest_sha256": sha(manifest),
        "source": source,
    }


def check_source_artifact(path: Path, expected: dict[str, str], label: str) -> None:
    value = read(path)
    require(value == expected, f"{label}: source census differs from retained baseline")


def check_metric_vectors(value: Any, samples: int, label: str) -> None:
    """Check every serialized measured metric vector recursively."""

    if isinstance(value, dict):
        if "status" in value and "values" in value:
            status = value.get("status")
            vector = value.get("values")
            require(isinstance(vector, list) and len(vector) == samples,
                    f"{label}.values length changed")
            require(status == "measured", f"{label} has values without measured status")
            for index, item in enumerate(vector):
                require(isinstance(item, (int, float, str, bool)) and item is not None,
                        f"{label}.values[{index}] is invalid")
        for key, child in value.items():
            if key != "values":
                check_metric_vectors(child, samples, f"{label}.{key}")
    elif isinstance(value, list):
        # Lists in operation metrics are vectors only when carried by the
        # status/values envelope; ordinary-save evidence lists are checked by
        # validate_ordinary_save with their phase-specific cardinality.
        for index, child in enumerate(value):
            if isinstance(child, dict):
                check_metric_vectors(child, samples, f"{label}[{index}]")


def validate_operation_metrics(value: Any, elapsed: dict[str, Any], samples: int,
                              lane: str, label: str) -> None:
    require(isinstance(value, dict), f"{label}.operation_metrics is missing")
    require(value.get("sample_count") == samples, f"{label} operation sample count changed")
    require(value.get("sample_indices") == elapsed["sample_order"],
            f"{label} operation sample identity is not aligned to elapsed samples")
    require(value.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{label} operation alignment changed")
    check_metric_vectors(value, samples, f"{label}.operation_metrics")
    allocation = value.get("allocation")
    require(isinstance(allocation, dict), f"{label} allocation envelope is missing")
    expected_status = "measured" if lane == "allocator" else "unavailable"
    require(allocation.get("status") == expected_status,
            f"{label} allocation status is {allocation.get('status')!r}, expected {expected_status}")
    require(allocation.get("scope") == "operation_global_system_allocator",
            f"{label} allocation scope changed")
    for field in ALLOCATION_FIELDS:
        metric = allocation.get(field)
        require(isinstance(metric, dict), f"{label}.allocation.{field} is missing")
        require(metric.get("status") == expected_status,
                f"{label}.allocation.{field} status changed")
        require(metric.get("scope") == "operation_global_system_allocator",
                f"{label}.allocation.{field} scope changed")
        if expected_status == "measured":
            vector = metric.get("values")
            require(isinstance(vector, list) and len(vector) == samples,
                    f"{label}.allocation.{field} vector length changed")
            for index, item in enumerate(vector):
                nonnegative_int(item, f"{label}.allocation.{field}.values[{index}]")
            if field == "failed_allocation_calls":
                require(all(item == 0 for item in vector),
                        f"{label} has failed allocation calls")
        else:
            require("values" not in metric, f"{label} unavailable allocation has values")


def validate_byte_split(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label} byte split is missing")
    fields = (
        "output_total_bytes", "payload_bytes_deflated", "payload_bytes_stored",
        "payload_bytes_identical_to_source", "payload_bytes_regenerated",
        "uncompressed_payload_bytes_compressed", "uncompressed_payload_bytes_regenerated",
        "framing_bytes", "output_member_count", "deflate_member_count",
        "stored_member_count", "members_identical_to_source", "members_regenerated",
    )
    for field in fields:
        nonnegative_int(value.get(field), f"{label}.{field}")
    require(value["payload_bytes_deflated"] + value["payload_bytes_stored"]
            == value["payload_bytes_identical_to_source"] + value["payload_bytes_regenerated"],
            f"{label} payload split does not balance")
    require(value["output_member_count"] == value["members_identical_to_source"]
            + value["members_regenerated"], f"{label} member split does not balance")
    require(value["output_member_count"] == value["deflate_member_count"]
            + value["stored_member_count"], f"{label} compression member split does not balance")
    require(value["output_total_bytes"] >= value["framing_bytes"],
            f"{label} framing exceeds output")


def result_corpus_identity(result: dict[str, Any], ordinary: dict[str, Any],
                           corpus: dict[str, Any], label: str) -> None:
    manifest = result.get("corpus")
    evidence = ordinary.get("corpus")
    require(isinstance(manifest, dict) and isinstance(evidence, dict),
            f"{label} corpus evidence is missing")
    require(manifest.get("package_format") == "DOCX/OPC/ZIP",
            f"{label} package format changed")
    require(manifest.get("archive_sha256") == evidence.get("source_archive_sha256"),
            f"{label} result/source archive digest mismatch")
    require(manifest.get("archive_bytes") == evidence.get("source_archive_bytes"),
            f"{label} result/source archive size mismatch")
    require(manifest.get("archive_member_count") == evidence.get("source_member_count"),
            f"{label} result/source member count mismatch")
    check_hex(manifest.get("archive_sha256"), f"{label}.corpus.archive_sha256")
    check_hex(evidence.get("source_archive_sha256"), f"{label}.source_archive_sha256")
    require(evidence.get("format") == "DOCX", f"{label} ordinary format changed")
    require(evidence.get("origin") == corpus["origin"], f"{label} ordinary origin changed")
    require(evidence.get("save_entry_point") == "litchi_docx::Package::save",
            f"{label} save entry point changed")
    require(evidence.get("sink_entry_point") == "litchi_docx::Package::to_stream",
            f"{label} sink entry point changed")
    if corpus["origin"] == "caller-named-real-file":
        # The real-file manifest has one unit for both entry_count and
        # archive_member_count: the non-directory ZIP members.  Keep its
        # producer identity explicit so a generated semantic-unit count can
        # never be mistaken for a physical archive-member count.
        require(manifest.get("name") == "docx-real-file-save",
                f"{label} real-file manifest name changed")
        require(manifest.get("generator") == "litchi-docx-real-file-v1",
                f"{label} real-file manifest generator changed")
        require(manifest.get("shape") == "real-file",
                f"{label} real-file manifest shape changed")
        require(manifest.get("payload_kind") == "real-producer-ooxml",
                f"{label} real-file manifest payload kind changed")
        fixture = corpus_fixture_binding(corpus)
        real_file = evidence.get("real_file")
        require(isinstance(real_file, dict) and fixture is not None,
                f"{label} real-file provenance missing")
        require(real_file.get("path") == corpus.get("path"),
                f"{label} real-file plan spelling changed")
        raw_real_path = Path(str(real_file.get("path")))
        resolved_real_path = ((REPO / raw_real_path).resolve()
                              if not raw_real_path.is_absolute() else raw_real_path.resolve())
        require(resolved_real_path == Path(fixture["resolved_path"]),
                f"{label} real-file path changed")
        require(real_file.get("sha256") == fixture["sha256"]
                and real_file.get("bytes") == fixture["bytes"],
                f"{label} real-file identity changed")
        # Independently recompute the corrected real-file manifest quantities
        # from the fixture archive.  In particular, target_payload_bytes and
        # target_payload_sha256 are decoded XML facts, while
        # uncompressed_payload_bytes is the checked total of all non-directory
        # ZIP members; neither is a compressed member-range measurement.
        with zipfile.ZipFile(fixture["resolved_path"]) as archive:
            members = [info for info in archive.infolist() if not info.is_dir()]
            require(len(members) == evidence["source_member_count"],
                    f"{label} fixture member count changed")
            require(manifest.get("entry_count") == len(members),
                    f"{label} real-file entry count is not the ZIP member count")
            main_info = next((info for info in members if info.filename == "word/document.xml"), None)
            require(main_info is not None, f"{label} fixture main DOCX member is missing")
            main_payload = archive.read(main_info)
            require(manifest.get("target_entry") == "word/document.xml"
                    and manifest.get("target_payload_bytes") == len(main_payload)
                    and manifest.get("target_payload_sha256") == hashlib.sha256(main_payload).hexdigest(),
                    f"{label} decoded target manifest identity changed")
            require(manifest.get("uncompressed_payload_bytes")
                    == sum(info.file_size for info in members),
                    f"{label} checked uncompressed member total changed")
    else:
        # The generated DOCX uses semantic paragraph units for entry_count;
        # archive_member_count remains the physical ZIP-member count checked
        # above.  These source-defined Medium-shape constants deliberately
        # keep the two dimensions separate (200 paragraphs versus 12 ZIP
        # members, 50-byte paragraph entries versus 10,000 text bytes).
        generated_manifest = {
            "name": "docx-semantic-medium",
            "generator": "litchi-docx-semantic-v1",
            "shape": "medium",
            "payload_kind": "deterministic-semantic-text",
            "entry_count": 200,
            "entry_bytes": 50,
            "uncompressed_payload_bytes": 10_000,
            "target_entry": "paragraph:0",
            "target_payload_bytes": 50,
        }
        for key, expected in generated_manifest.items():
            require(manifest.get(key) == expected,
                    f"{label} generated manifest {key} changed")
        require(evidence.get("real_file") is None, f"{label} generated corpus has real provenance")


def validate_ordinary_save(result: dict[str, Any], corpus: dict[str, Any], phase: str,
                           lane: str, samples: int, label: str) -> dict[str, Any]:
    source = result.get("source")
    require(isinstance(source, dict), f"{label}.source is missing")
    ordinary = source.get("ordinary_save")
    require(isinstance(ordinary, dict), f"{label}.source.ordinary_save is missing")
    require(ordinary.get("format") == "DOCX", f"{label} format changed")
    require(ordinary.get("origin") == corpus["origin"], f"{label} origin changed")
    require(ordinary.get("phase") == PHASE_LABELS[phase], f"{label} phase label changed")
    require(ordinary.get("timing_scope") == PHASE_TIMING_SCOPES[phase],
            f"{label} timing scope changed")
    require(isinstance(ordinary.get("atomic_publication_steps"), str)
            and ordinary["atomic_publication_steps"], f"{label} atomic steps missing")
    result_corpus_identity(result, ordinary, corpus, label)
    evidence = ordinary["corpus"]
    require(evidence.get("edit_description") == EDIT_MARKER,
            f"{label} edit description changed")
    require(evidence.get("edit_admitted") == corpus["expected_edit_admitted"],
            f"{label} edit admission differs from plan")
    outcome = evidence.get("edit_outcome")
    require(isinstance(outcome, str) and outcome, f"{label} edit outcome missing")
    if corpus["expected_edit_admitted"]:
        require(outcome == "admitted", f"{label} admitted corpus outcome changed")
    else:
        require(outcome.startswith("refused:"), f"{label} refusal outcome changed")
    check_hex(evidence.get("published_sha256"), f"{label}.published_sha256")
    require(evidence.get("repeated_cycles_identical") is True
            and evidence.get("repeated_saves_identical") is True,
            f"{label} determinism proof failed")
    require(evidence.get("repeated_cycle_sha256") == evidence.get("published_sha256")
            and evidence.get("repeated_save_sha256") == evidence.get("published_sha256"),
            f"{label} repeated publication digest changed")
    validate_byte_split(evidence.get("byte_split"), f"{label}.corpus")
    require(evidence["byte_split"]["output_total_bytes"] == evidence["published_bytes"],
            f"{label} corpus byte split/output size mismatch")

    published = ordinary.get("published_sha256")
    edit_hashes = ordinary.get("edit_outcome_sha256")
    require(isinstance(published, list) and isinstance(edit_hashes, list),
            f"{label} sample evidence vectors are missing")
    require(len(edit_hashes) == samples, f"{label} edit outcome vector length changed")
    outcome_digest = hashlib.sha256(outcome.encode()).hexdigest()
    require(all(item == outcome_digest for item in edit_hashes),
            f"{label} edit outcome evidence is not deterministic")
    require(ordinary.get("edit_outcomes_identical") is True,
            f"{label} edit outcome identity flag failed")
    expected_published = evidence["published_sha256"]
    if phase == "edit":
        require(published == [] and result.get("output_sha256") is None,
                f"{label} edit phase unexpectedly published output evidence")
    else:
        require(len(published) == samples and all(item == expected_published for item in published),
                f"{label} publication digest vector changed")
        require(result.get("output_sha256") == expected_published,
                f"{label} output digest changed")
    require(ordinary.get("publications_identical") is True,
            f"{label} publication identity flag failed")
    sample_split = ordinary.get("sample_byte_split")
    if phase == "counting_publish":
        validate_byte_split(sample_split, f"{label}.sample")
        require(sample_split == evidence["byte_split"], f"{label} sample byte split changed")
        sink = result.get("sink")
        require(isinstance(sink, dict), f"{label} counting sink is missing")
        nonnegative_int(sink.get("accepted_bytes"), f"{label}.sink.accepted_bytes")
        nonnegative_int(sink.get("write_calls"), f"{label}.sink.write_calls")
        require(sink["accepted_bytes"] == sample_split["output_total_bytes"],
                f"{label} sink/output byte count mismatch")
    else:
        require(sample_split is None, f"{label} non-counting phase has byte split")
        require(result.get("sink") is None, f"{label} non-counting phase has sink evidence")
    return ordinary


def check_report_metadata(report: dict[str, Any], build: dict[str, Any], lane: str,
                          case: str, samples: int, warmup: int, label: str) -> None:
    require(report.get("schema_version") == 1, f"{label} report schema changed")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{label} binary identity is missing")
    require(identity.get("binary_sha256") == build["binary_sha256"]
            and identity.get("binary_bytes") == build["binary_bytes"],
            f"{label} binary identity is not bound to build record")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label} configuration is missing")
    require(configuration.get("samples_per_case") == samples
            and configuration.get("warmup_iterations_per_case") == warmup,
            f"{label} sample configuration changed")
    require(configuration.get("cases") == [case], f"{label} report mixes cases")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label} report must contain exactly one result")


def expected_jobs(p: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    order_index = 0
    for repeat in range(1, p[lane]["repeats"] + 1):
        order = list(p.get("phase_order", list(PHASES)))
        if repeat % 2 == 0:
            order.reverse()
        for phase in order:
            corpora = list(p["corpora"])
            if repeat % 2 == 0:
                corpora.reverse()
            for corpus in corpora:
                result.append({
                    "lane": lane,
                    "repeat": repeat,
                    "phase": phase,
                    "corpus": normalized_corpus(corpus),
                    "name": f"{lane}-r{repeat}-{corpus['id']}-{phase}",
                    "case": ("docx_ordinary_save_" if corpus["origin"] == "generated-harness-corpus"
                             else "docx_real_file_ordinary_save_") + phase,
                    "samples": p[lane]["samples"],
                    "warmup": p[lane]["warmup"],
                    "lane_order_index": order_index,
                })
                order_index += 1
    return result


def expected_command(job: dict[str, Any], build: dict[str, Any], plan: dict[str, Any]) -> list[str]:
    """Reconstruct the one-child argv from the frozen plan and job identity."""

    name = job["name"]
    command = [
        "taskset", "-c", str(plan["cpu"]), str(build["binary"]),
        "--warmup", str(job["warmup"]),
        "--samples", str(job["samples"]),
        "--case", job["case"],
        "--json", str(HERE / f"{name}.json"),
    ]
    filesystem_root = plan.get("filesystem_root")
    if filesystem_root:
        command += ["--filesystem-root", str(filesystem_root)]
    if job["corpus"]["origin"] == "caller-named-real-file":
        command += ["--ooxml-file", str(job["corpus"]["path"])]
    return command


def validate_receipt(job: dict[str, Any], build: dict[str, Any], plan: dict[str, Any],
                     expected_source: dict[str, str], script_digest: str,
                     plan_digest: str, constraints_digest: str) -> dict[str, Any]:
    name = job["name"]
    receipt_path = HERE / f"{name}.receipt.json"
    receipt = read(receipt_path)
    require(receipt.get("schema_version") == 1, f"{name} receipt schema changed")
    for key in ("lane", "repeat", "corpus_id", "phase", "case", "samples", "warmup",
                "lane_order_index"):
        expected = job["corpus"]["id"] if key == "corpus_id" else job[key]
        require(receipt.get(key) == expected, f"{name} receipt {key} changed")
    require(receipt.get("exit_code") == 0, f"{name} child failed")
    require(receipt.get("cpu") == job.get("cpu", receipt.get("cpu")),
            f"{name} cpu binding is absent")
    require(receipt.get("binary_sha256") == build["binary_sha256"]
            and receipt.get("binary_bytes") == build["binary_bytes"],
            f"{name} binary receipt is not bound to build")
    require(receipt.get("build_record_sha256") == build["build_record_sha256"],
            f"{name} build record binding changed")
    require(receipt.get("build_source_manifest_sha256") == build["source_manifest_sha256"],
            f"{name} build/source binding changed")
    require(receipt.get("plan_sha256") == plan_digest, f"{name} plan binding changed")
    require(receipt.get("script_sha256") == script_digest, f"{name} capture script changed")
    require(receipt.get("constraints_sha256") == constraints_digest,
            f"{name} constraints binding changed")
    retained = receipt.get("retained_binary_source")
    require(isinstance(retained, dict)
            and retained.get("manifest") == "source-baseline.json"
            and retained.get("manifest_sha256") == build["source_manifest_sha256"]
            and retained.get("source_census_sha256") == json_digest(expected_source),
            f"{name} retained source binding changed")
    current = receipt.get("current_checkout_source")
    require(isinstance(current, dict) and current.get("unchanged_during_child") is True,
            f"{name} current source custody failed")
    for which in ("before_artifact", "after_artifact"):
        expected_filename = (
            f"{name}.source-before.json" if which == "before_artifact"
            else f"{name}.source-after.json"
        )
        require(current.get(which) == expected_filename,
                f"{name} {which} filename changed")
        source_path = HERE / str(current.get(which))
        check_source_artifact(source_path, expected_source, f"{name} {which}")
        require(current.get(which.replace("artifact", "file_sha256")) == sha(source_path),
                f"{name} {which} digest binding changed")
    require(current.get("before_sha256") == json_digest(expected_source)
            and current.get("after_sha256") == json_digest(expected_source),
            f"{name} source census digest changed")
    fixture = receipt.get("fixture")
    planned_fixture = corpus_fixture_binding(job["corpus"])
    require(isinstance(fixture, dict), f"{name} fixture receipt is missing")
    if planned_fixture is None:
        require(fixture.get("plan_path") is None and fixture.get("before") is None
                and fixture.get("after") is None, f"{name} generated fixture binding exists")
    else:
        require(fixture.get("plan_path") == planned_fixture["plan_path"]
                and fixture.get("plan_sha256") == planned_fixture["sha256"],
                f"{name} fixture plan binding changed")
        require(fixture.get("before") == fixture.get("after"),
                f"{name} fixture changed during child")
        require(fixture["before"].get("path") == planned_fixture["plan_path"]
                and fixture["before"].get("resolved_path") == planned_fixture["resolved_path"],
                f"{name} fixture path binding changed")
        require(fixture["before"]["sha256"] == planned_fixture["sha256"]
                and fixture["before"]["bytes"] == planned_fixture["bytes"],
                f"{name} fixture digest/size binding changed")
    command = receipt.get("command")
    require(command == expected_command(job, build, plan),
            f"{name} command argv is not the frozen taskset/binary/plan binding")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{name} artifact inventory is missing")
    expected_names = {
        f"{name}.json", f"{name}.stdout", f"{name}.stderr",
        f"{name}.source-before.json", f"{name}.source-after.json",
    }
    require(set(artifacts) == expected_names, f"{name} artifact inventory changed")
    for filename, digest in artifacts.items():
        path = HERE / filename
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{name} artifact digest changed: {filename}")
    return receipt


def stable_semantics(result: dict[str, Any]) -> dict[str, Any]:
    value = copy.deepcopy(result)
    value.pop("elapsed_ns", None)
    value.pop("operation_metrics", None)
    # Native and allocator lanes retain different sample counts (100 versus
    # 3). The vectors themselves are validated lane-locally; semantic parity
    # compares the frozen outcome represented by those vectors.
    ordinary = value.get("source", {}).get("ordinary_save")
    if isinstance(ordinary, dict):
        ordinary.pop("published_sha256", None)
        ordinary.pop("edit_outcome_sha256", None)
    return value


def stable_evidence(result: dict[str, Any]) -> dict[str, Any]:
    ordinary = result["source"]["ordinary_save"]
    return {
        "corpus": result["corpus"],
        "ordinary_corpus": ordinary["corpus"],
        "published_sha256": ordinary["corpus"]["published_sha256"],
        "byte_split": ordinary["corpus"]["byte_split"],
        "sample_byte_split": ordinary.get("sample_byte_split"),
    }


def percent_spread(values: list[float]) -> float:
    require(values and all(value > 0 for value in values), "spread values must be positive")
    return (max(values) - min(values)) * 100.0 / min(values)


def analyze_repeats(records: dict[tuple[str, str, str], dict[str, Any]],
                    p: dict[str, Any], lane: str) -> dict[str, Any]:
    metrics = ("p50", "mean", "p95", "p99")
    by_corpus: dict[str, Any] = {}
    flags: list[dict[str, Any]] = []
    for corpus in p["corpora"]:
        cid = corpus["id"]
        phase_data: dict[str, Any] = {}
        for phase in PHASES:
            repeat_stats = []
            for repeat in range(1, p[lane]["repeats"] + 1):
                result = records[(cid, phase, str(repeat))]
                stats = elapsed_stats(result["elapsed_ns"]["samples"])
                repeat_stats.append({"repeat": repeat, **stats})
            drift: dict[str, Any] = {}
            for metric in metrics:
                values = [float(item[metric]) for item in repeat_stats]
                spread = percent_spread(values)
                drift[metric] = {"values": values, "spread_percent": spread,
                                 "flag_over_5_percent": spread > 5.0}
                if spread > 5.0:
                    flags.append({"lane": lane, "corpus_id": cid, "phase": phase,
                                  "metric": metric, "spread_percent": spread})
            phase_data[phase] = {"repeats": repeat_stats, "repeat_drift": drift}
        by_corpus[cid] = phase_data

    comparisons: dict[str, Any] = {}
    for corpus in p["corpora"]:
        cid = corpus["id"]
        pairs: dict[str, Any] = {}
        for index, left in enumerate(PHASES):
            for right in PHASES[index + 1:]:
                left_values = by_corpus[cid][left]["repeats"]
                right_values = by_corpus[cid][right]["repeats"]
                pair: dict[str, Any] = {}
                for metric in metrics:
                    left_median = statistics.median(float(item[metric]) for item in left_values)
                    right_median = statistics.median(float(item[metric]) for item in right_values)
                    pair[metric] = {
                        "left_median": left_median,
                        "right_median": right_median,
                        "difference": right_median - left_median,
                        "ratio": right_median / left_median,
                    }
                pairs[f"{left}_vs_{right}"] = pair
        comparisons[cid] = pairs
    return {
        "per_corpus_phase": by_corpus,
        "repeat_flags_over_5_percent": flags,
        "phase_comparisons": comparisons,
        "phase_comparison_note": "Each phase is timed independently; phase values are descriptive comparisons and are not additive or a decomposition of lifecycle time.",
    }


def signed_metric_spread(values: list[float]) -> float:
    """Return a bounded relative spread even when a derived metric crosses zero."""

    require(values, "metric spread vector is empty")
    scale = min((abs(value) for value in values if value != 0.0), default=0.0)
    if scale == 0.0:
        return 0.0 if max(values) == min(values) else float("inf")
    return (max(values) - min(values)) * 100.0 / scale


def analyze_allocator_allocation(
    records: dict[tuple[str, str, str], dict[str, Any]], p: dict[str, Any],
) -> dict[str, Any]:
    """Summarize six allocator samples per corpus/phase independently.

    The phase vectors are pooled only within their own phase across the two
    allocator repeats.  In particular, absolute live/peak counters are never
    added across phases or repeats; their per-sample observations and
    per-repeat variation remain visible.
    """

    fields = list(ALLOCATION_FIELDS) + ["net_live", "peak_above_start"]
    per_corpus_phase: dict[str, Any] = {}
    repeat_flags: list[dict[str, Any]] = []
    for corpus in p["corpora"]:
        cid = corpus["id"]
        per_phase: dict[str, Any] = {}
        for phase in PHASES:
            repeats: list[dict[str, Any]] = []
            pooled: dict[str, list[int]] = {field: [] for field in fields}
            for repeat in range(1, p["allocator"]["repeats"] + 1):
                result = records[(cid, phase, str(repeat))]
                allocation = result["operation_metrics"]["allocation"]
                repeat_values: dict[str, list[int]] = {
                    field: list(allocation[field]["values"])
                    for field in ALLOCATION_FIELDS
                }
                repeat_values["net_live"] = [
                    after - before
                    for before, after in zip(
                        repeat_values["live_bytes_before"],
                        repeat_values["live_bytes_after"],
                    )
                ]
                repeat_values["peak_above_start"] = [
                    peak - before
                    for peak, before in zip(
                        repeat_values["region_peak_live_bytes"],
                        repeat_values["live_bytes_before"],
                    )
                ]
                for field in fields:
                    pooled[field].extend(repeat_values[field])
                repeats.append({
                    "repeat": repeat,
                    "sample_count": len(next(iter(repeat_values.values()))),
                    "values": repeat_values,
                    "stats": {
                        field: integer_stats(values, f"{cid}/{phase}/r{repeat}/{field}")
                        for field, values in repeat_values.items()
                    },
                })

            variation: dict[str, Any] = {}
            for field in fields:
                medians = [float(item["stats"][field]["p50"]) for item in repeats]
                spread = signed_metric_spread(medians)
                entry = {
                    "repeat_p50_values": medians,
                    "spread_percent": spread,
                    "flag_over_5_percent": spread > 5.0,
                }
                variation[field] = entry
                if spread > 5.0:
                    repeat_flags.append({
                        "corpus_id": cid,
                        "phase": phase,
                        "metric": field,
                        "spread_percent": spread,
                    })
            per_phase[phase] = {
                "repeats": repeats,
                "pooled_six_samples": {
                    field: {
                        "values": pooled[field],
                        "stats": integer_stats(pooled[field], f"{cid}/{phase}/{field}/pooled"),
                    }
                    for field in fields
                },
                "repeat_variation": variation,
                "aggregation_note": "Two repeats are pooled only within this phase; phase values and absolute peak/live counters are never summed across phases.",
            }
        per_corpus_phase[cid] = per_phase
    return {
        "metrics": fields,
        "samples_per_repeat": p["allocator"]["samples"],
        "repeats": p["allocator"]["repeats"],
        "pooled_samples_per_corpus_phase": p["allocator"]["samples"] * p["allocator"]["repeats"],
        "per_corpus_phase": per_corpus_phase,
        "repeat_flags_over_5_percent": repeat_flags,
        "aggregation_note": "Allocator metrics are operation-scoped observations; six samples are pooled only inside each corpus/phase, and phase peaks are never summed.",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", default=str(HERE / "analysis.json"))
    args = parser.parse_args()
    try:
        p = load_plan()
        expected_source = read(HERE / "source-baseline.json")
        require(isinstance(expected_source, dict), "source-baseline.json is invalid")
        current_source = source_census()
        require(current_source == expected_source,
                "current source is not the frozen baseline census")
        constraints_path = HERE / "constraints.json"
        constraints = read(constraints_path)
        require(isinstance(constraints, dict), "constraints is invalid")
        for name, digest in constraints.items():
            path = REPO / name
            require(path.is_file() and sha(path) == digest, f"constraint changed: {name}")

        builds = {lane: build_info("native" if lane == "native" else "alloc")
                  for lane in ("native", "allocator")}
        plan_digest = sha(HERE / "plan.json")
        script_digest = sha(Path(__file__).resolve().parent / "capture.py")
        constraints_digest = sha(constraints_path)
        jobs: list[dict[str, Any]] = []
        for lane in ("native", "allocator"):
            lane_jobs = expected_jobs(p, lane)
            require(len(lane_jobs) == (36 if lane == "native" else 24),
                    f"{lane} child count changed")
            for job in lane_jobs:
                # Retain the cpu expectation without duplicating it in every
                # plan child: capture records it as an identity witness.
                job["cpu"] = p["cpu"]
                receipt = validate_receipt(
                    job, builds[lane], p, expected_source, script_digest,
                    plan_digest, constraints_digest,
                )
                report = read(HERE / f"{job['name']}.json")
                check_report_metadata(report, builds[lane], lane, job["case"],
                                      job["samples"], job["warmup"], job["name"])
                result = report["results"][0]
                require(result.get("case") == job["case"], f"{job['name']} case changed")
                elapsed, _ = validate_elapsed(result.get("elapsed_ns"), job["samples"], job["name"])
                ordinary = validate_ordinary_save(
                    result, job["corpus"], job["phase"], lane,
                    job["samples"], job["name"],
                )
                validate_operation_metrics(result.get("operation_metrics"),
                                           result["elapsed_ns"], job["samples"],
                                           lane, job["name"])
                jobs.append({**job, "report": report, "result": result,
                             "elapsed": elapsed, "ordinary": ordinary,
                             "receipt": receipt})

        by_key = {(job["lane"], job["corpus"]["id"], job["phase"], str(job["repeat"])): job
                  for job in jobs}
        require(len(by_key) == len(jobs), "duplicate child identity")

        # Semantic parity is checked between native and allocator children
        # after removing only timed elapsed and instrumentation envelopes.
        shared_repeats = min(p["native"]["repeats"], p["allocator"]["repeats"])
        for job in jobs:
            if job["lane"] != "native":
                continue
            if job["repeat"] > shared_repeats:
                continue
            peer = by_key[("allocator", job["corpus"]["id"], job["phase"], str(job["repeat"]))]
            require(stable_semantics(job["result"]) == stable_semantics(peer["result"]),
                    f"native/allocator semantic parity failed for {job['name']}")

        # Corpus and publication evidence must remain deterministic across
        # phases, repeats, and the allocator lane.  Phase labels and elapsed
        # metric envelopes are intentionally excluded from the stable key.
        stable_by_corpus: dict[str, dict[str, Any]] = {}
        for job in jobs:
            cid = job["corpus"]["id"]
            value = stable_evidence(job["result"])
            if cid not in stable_by_corpus:
                stable_by_corpus[cid] = value
            else:
                require(value["corpus"] == stable_by_corpus[cid]["corpus"],
                        f"{job['name']} corpus manifest parity failed")
                require(value["ordinary_corpus"] == stable_by_corpus[cid]["ordinary_corpus"],
                        f"{job['name']} ordinary corpus evidence parity failed")

        native_records = {
            (job["corpus"]["id"], job["phase"], str(job["repeat"])): job["result"]
            for job in jobs if job["lane"] == "native"
        }
        allocator_records = {
            (job["corpus"]["id"], job["phase"], str(job["repeat"])): job["result"]
            for job in jobs if job["lane"] == "allocator"
        }

        # The tuple map above intentionally keeps the two lanes separate; the
        # report has no cross-lane timing arithmetic.
        native_analysis = analyze_repeats(native_records, p, "native")
        allocator_analysis = analyze_repeats(allocator_records, p, "allocator")
        allocator_allocation = analyze_allocator_allocation(allocator_records, p)
        output = {
            "schema_version": 1,
            "packet": "change-0709-docx-ordinary-save",
            "revision": p["revision"],
            "performance_claim": "descriptive-current-source-baseline-only; no speedup claim",
            "native": {
                "build": {key: builds["native"][key] for key in
                           ("name", "binary_sha256", "binary_bytes", "build_record_sha256", "source_manifest_sha256")},
                "children": 36,
                **native_analysis,
            },
            "allocator": {
                "build": {key: builds["allocator"][key] for key in
                           ("name", "binary_sha256", "binary_bytes", "build_record_sha256", "source_manifest_sha256")},
                "children": 24,
                "elapsed_claim": "instrumented allocator elapsed evidence; not latency-comparable to native",
                "allocation_metrics": allocator_allocation,
                **allocator_analysis,
            },
            "corpora": [
                {
                    "id": corpus["id"],
                    "label": corpus["label"],
                    "origin": corpus["origin"],
                    "path": corpus.get("path"),
                    "sha256": corpus.get("sha256"),
                    "expected_edit_admitted": corpus["expected_edit_admitted"],
                    "fixture": corpus_fixture_binding(corpus),
                    "stable_evidence_sha256": json_digest(stable_by_corpus[corpus["id"]]),
                }
                for corpus in p["corpora"]
            ],
            "verification": {
                "current_source_exact_baseline": True,
                "child_receipts_verified": len(jobs),
                "native_allocator_semantic_parity_verified": True,
                "deterministic_corpus_output_evidence_verified": True,
                "metric_vector_lengths_verified": True,
                "raw_elapsed_statistics_recomputed": True,
                "phase_values_additive": False,
                "speedup_claims": [],
            },
            "source_manifest_sha256": sha(HERE / "source-baseline.json"),
            "plan_sha256": plan_digest,
            "constraints_sha256": constraints_digest,
            "capture_script_sha256": script_digest,
        }
        write(Path(args.output), output)
        print(f"verified {len(jobs)} children; wrote {args.output}")
    except (AssertionError, OSError, RuntimeError, ValueError, KeyError) as error:
        print(f"analysis failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
