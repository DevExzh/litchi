"""Validate current-head XLSX native phase evidence.

The benchmark keeps its phase vectors in acquisition order while the native
runner sorts ``elapsed_ns.samples`` and retains the original indexes in
``sample_order``.  This analyzer independently reconstructs that alignment,
recomputes the descriptive statistics, and checks the source, sink, cache,
budget, corpus, and output invariants before emitting a report.  It is a
diagnostic for the current head; it never compares this capture with an older
build.
"""

from __future__ import annotations

import datetime
import hashlib
import json
import math
from pathlib import Path
import statistics
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]

PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
ALL_PHASES = PHASES + ("reopen_ns",)
ALLOCATION_FIELDS = (
    "plan_allocation_metrics",
    "staging_allocation_metrics",
    "commit_allocation_metrics",
    "commit_core_allocation_metrics",
    "publication_allocation_metrics",
)
HEX64 = set("0123456789abcdef")
UNSET = object()


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read_json(root: Path, name: str) -> dict[str, Any]:
    path = root / name
    check(path.is_file(), f"missing evidence file: {path}")
    value = json.loads(path.read_text())
    check(isinstance(value, dict), f"{path} is not a JSON object")
    return value


def source_check(root: Path = HERE) -> None:
    """Verify the source census for a real current-head analysis."""

    manifest = read_json(root, "source-manifest.json")
    for name, digest in manifest.items():
        path = REPO / name
        check(path.is_file(), f"source-manifest path is absent: {name}")
        check(sha(path) == digest, f"source-manifest digest mismatch: {name}")


def midpoint(left: int, right: int) -> int:
    # This is the overflow-safe integer midpoint used by the Rust harness.
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def nearest_rank(values: list[int], percentile: int) -> int:
    index = ((percentile * len(values) + 99) // 100) - 1
    return values[min(index, len(values) - 1)]


def stats(values: list[int]) -> dict[str, int | float]:
    check(values, "a timing vector is empty")
    check(all(isinstance(value, int) and not isinstance(value, bool) for value in values),
          "a timing vector contains a non-integer")
    check(all(value >= 0 for value in values), "a timing vector contains a negative value")
    ordered = sorted(values)
    return {
        "p50": midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
        "p95": nearest_rank(ordered, 95),
        "p99": nearest_rank(ordered, 99),
        "mean": statistics.mean(values),
        "min": ordered[0],
        "max": ordered[-1],
    }


def close(left: float, right: float, message: str) -> None:
    check(math.isfinite(left) and math.isfinite(right), message + " is not finite")
    check(math.isclose(left, right, rel_tol=1e-12, abs_tol=1e-9),
          f"{message}: {left!r} != {right!r}")


def check_hex(value: Any, label: str) -> None:
    check(isinstance(value, str) and len(value) == 64 and set(value) <= HEX64,
          f"{label} is not a lowercase SHA-256 digest")


def check_vector(values: Any, count: int, label: str, *, nonnegative: bool = True) -> list[Any]:
    check(isinstance(values, list), f"{label} is not a vector")
    check(len(values) == count, f"{label} has {len(values)} values; expected {count}")
    if nonnegative:
        check(all(isinstance(value, int) and not isinstance(value, bool) and value >= 0
                  for value in values),
              f"{label} contains a non-negative integer violation")
    return values


def check_reported_statistics(elapsed: dict[str, Any], samples: int) -> dict[str, int | float]:
    check(elapsed.get("unit") == "ns", "elapsed_ns unit is not ns")
    values = check_vector(elapsed.get("samples"), samples, "elapsed_ns.samples")
    check(all(value > 0 for value in values), "elapsed samples must be positive")
    order = check_vector(elapsed.get("sample_order"), samples, "elapsed_ns.sample_order")
    # The values are sorted by the runner, while equal values retain their
    # original indexes.  Reconstructing this pair catches both a reordered
    # vector and a stale alignment vector.
    check(sorted(order) == list(range(samples)), "sample_order is not a permutation")
    check(values == sorted(values), "elapsed_ns.samples is not sorted")
    expected = stats(values)
    for key, value in expected.items():
        reported = elapsed.get(key)
        check(reported is not None, f"elapsed_ns is missing {key}")
        if isinstance(value, int):
            check(reported == value, f"elapsed_ns.{key}: {reported!r} != {value!r}")
        else:
            close(float(reported), float(value), f"elapsed_ns.{key}")

    standard_deviation = statistics.stdev(values) if samples > 1 else 0.0
    reported_standard_deviation = elapsed.get("standard_deviation")
    check(isinstance(reported_standard_deviation, (int, float)),
          "elapsed_ns.standard_deviation is missing or non-numeric")
    close(float(reported_standard_deviation), standard_deviation,
          "elapsed_ns.standard_deviation")
    interval = elapsed.get("confidence_interval_95")
    check(isinstance(interval, dict), "elapsed_ns confidence interval is not an object")
    check(interval.get("method") == "two-sided Student's t interval for the mean",
          "elapsed_ns confidence interval method changed")
    degrees = samples - 1
    check(degrees > 30, "native sample count is too small for the retained t interval")
    z = 1.959963984540054
    df = float(degrees)
    critical = (z + (z**3 + z) / (4 * df)
                + (5 * z**5 + 16 * z**3 + 3 * z) / (96 * df**2)
                + (3 * z**7 + 19 * z**5 + 17 * z**3 - 15 * z) / (384 * df**3))
    margin = critical * standard_deviation / math.sqrt(samples)
    lower = interval.get("lower")
    upper = interval.get("upper")
    check(isinstance(lower, (int, float)) and isinstance(upper, (int, float)),
          "elapsed_ns confidence interval bounds are missing or non-numeric")
    close(float(lower), max(0.0, float(expected["mean"]) - margin),
          "elapsed_ns confidence_interval_95.lower")
    close(float(upper), float(expected["mean"]) + margin,
          "elapsed_ns confidence_interval_95.upper")
    check(float(lower) <= float(upper),
          "elapsed_ns confidence interval is inverted")
    return expected


def check_budget_snapshot(snapshot: Any, label: str, *, post: bool) -> None:
    check(isinstance(snapshot, dict), f"{label} is not an object")
    required = {
        "input_bytes_used", "input_bytes_limit", "output_bytes_used", "output_bytes_limit",
        "work_used", "work_limit", "objects_used", "objects_limit",
        "catalog_reserved_objects", "cache_reserved_objects",
    }
    check(set(snapshot) == required, f"{label} fields changed")
    for field in ("input_bytes_limit", "output_bytes_limit", "work_limit", "objects_limit"):
        check(snapshot[field] is None, f"{label}.{field} is non-null for unmanaged evidence")
    for field in ("input_bytes_used", "output_bytes_used", "work_used", "objects_used"):
        check(snapshot[field] == 0, f"{label}.{field} is nonzero for unmanaged evidence")
    for field in ("catalog_reserved_objects", "cache_reserved_objects"):
        expected = None if post else 0
        check(snapshot[field] == expected,
              f"{label}.{field} is {snapshot[field]!r}; expected {expected!r}")


def constant_value(values: list[Any], label: str) -> Any:
    check(values, f"{label} is empty")
    check(all(value == values[0] for value in values), f"{label} varies across samples")
    return values[0]


def validate_corpus(corpus: dict[str, Any], shape: str) -> dict[str, Any]:
    check(corpus.get("shape") == shape, f"corpus shape is not {shape}")
    check(corpus.get("package_format") == "XLSX/OPC/ZIP", f"{shape} is not an XLSX corpus")
    check(corpus.get("generator") == "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1",
          f"{shape} uses an unexpected corpus generator")
    for field in ("entry_count", "archive_member_count", "entry_bytes",
                  "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        check(isinstance(corpus.get(field), int) and corpus[field] > 0,
              f"corpus.{field} is not a positive integer")
    check_hex(corpus.get("archive_sha256"), f"{shape} corpus.archive_sha256")
    check_hex(corpus.get("target_payload_sha256"), f"{shape} corpus.target_payload_sha256")
    xlsx = corpus.get("xlsx")
    check(isinstance(xlsx, dict), f"{shape} corpus.xlsx is missing")
    for field in ("sheet_count", "rows_per_sheet", "columns_per_sheet", "one_percent_update_count"):
        check(isinstance(xlsx.get(field), int) and xlsx[field] > 0,
              f"{shape} corpus.xlsx.{field} is invalid")
    members = xlsx.get("source_members")
    check(isinstance(members, dict), f"{shape} source member manifest is missing")
    check(isinstance(members.get("workbook"), str) and members["workbook"],
          f"{shape} workbook member is missing")
    worksheets = members.get("worksheets")
    check(isinstance(worksheets, list) and len(worksheets) == xlsx["sheet_count"],
          f"{shape} worksheet member count does not match sheet_count")
    check(len(set(worksheets)) == len(worksheets), f"{shape} worksheet members are duplicated")
    check(all(isinstance(name, str) and name for name in worksheets),
          f"{shape} worksheet member name is invalid")
    for field in ("workbook", "styles"):
        check(isinstance(members.get(field), str) and members[field],
              f"{shape} source member {field} is missing")
    for field in ("shared_strings",):
        check(members.get(field) is None or isinstance(members[field], str),
              f"{shape} source member {field} is invalid")
    check(xlsx["one_percent_update_count"] <= corpus["entry_count"],
          f"{shape} update count exceeds the corpus entry count")
    return corpus


def validate_sink(sink: Any, label: str) -> None:
    check(isinstance(sink, dict), f"{label} sink is not an object")
    for field in ("accepted_bytes", "write_calls", "largest_write"):
        check(isinstance(sink.get(field), int) and sink[field] >= 0,
              f"{label} sink.{field} is invalid")
    buckets = sink.get("write_size_buckets")
    check(isinstance(buckets, dict), f"{label} sink buckets are missing")
    bucket_names = {
        "bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
        "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536",
    }
    check(set(buckets) == bucket_names, f"{label} sink bucket fields changed")
    check(all(isinstance(value, int) and value >= 0 for value in buckets.values()),
          f"{label} sink bucket count is invalid")
    check(sum(buckets.values()) == sink["write_calls"],
          f"{label} sink bucket count does not equal write_calls")
    check(sink["accepted_bytes"] > 0, f"{label} sink accepted no output bytes")
    check(sink["largest_write"] <= 65536, f"{label} sink exceeded the sequential write bound")
    check(buckets["bytes_over_65536"] == 0, f"{label} sink has an oversized write")


def validate_source(result: dict[str, Any], corpus: dict[str, Any], shape: str,
                    samples: int, *, compatibility: bool) -> tuple[dict[str, Any], dict[str, Any]]:
    source = result.get("source")
    check(isinstance(source, dict), f"{shape} source summary is missing")
    summary = source.get("xlsx_cell_values")
    check(isinstance(summary, dict), f"{shape} XLSX cell-values source summary is missing")
    check(summary.get("implementation") == "source-backed",
          f"{shape} is not source-backed evidence")
    check(summary.get("cache_mode") == "unmanaged-control",
          f"{shape} is not the unmanaged control")
    check(summary.get("cache_budget_managed") is False,
          f"{shape} cache budget unexpectedly reports managed mode")
    check(isinstance(summary.get("timing_scope"), str)
          and "open" in summary["timing_scope"]
          and "commit" in summary["timing_scope"]
          and "publication" in summary["timing_scope"]
          and "reopen" in summary["timing_scope"],
          f"{shape} timing scope does not document the phase boundary")
    xlsx = corpus["xlsx"]
    check(summary.get("update_count") == xlsx["one_percent_update_count"],
          f"{shape} update count does not match the corpus manifest")
    check(summary.get("selected_worksheet_count") == xlsx["sheet_count"],
          f"{shape} selected worksheet count does not match the corpus manifest")
    untouched = summary.get("untouched_member_count")
    check(isinstance(untouched, int) and 0 < untouched < corpus["archive_member_count"],
          f"{shape} untouched member count is invalid")
    partial = summary.get("partial_sink_verified", UNSET)
    if partial is not UNSET:
        check(partial is True, f"{shape} partial sink gate is false")
    # The selected medium and dense-sparse corpora do not carry the vendor
    # extension that enables the harness's partial-sink replay.  A future
    # corpus may serialize the gate as true; omission is correct for these
    # two native shapes.

    for field in ("payload_memory_limit", "publication_planning_memory_headroom",
                  "cache_budget_memory_limit"):
        check(summary.get(field) is None, f"{shape} unmanaged {field} is non-null")

    phase_vectors: dict[str, list[int]] = {}
    constants: dict[str, Any] = {}
    for key, value in summary.items():
        if key in ALLOCATION_FIELDS:
            # The current harness emits one explicit unavailable marker per
            # native sample.  Older 0520 output omitted these optional fields;
            # either representation carries no allocator measurement.
            if value:
                values = check_vector(value, samples,
                                      f"{shape} source.xlsx_cell_values.{key}",
                                      nonnegative=False)
                check(all(isinstance(item, dict)
                          and item.get("status") == "unavailable"
                          for item in values),
                      f"{shape} native {key} contains measured allocation data")
            continue
        if isinstance(value, list):
            values = check_vector(value, samples, f"{shape} source.xlsx_cell_values.{key}",
                                  nonnegative=False)
            if key in ALL_PHASES:
                check(all(isinstance(item, int) and not isinstance(item, bool) and item >= 0
                          for item in values),
                      f"{shape} phase {key} is not a non-negative integer vector")
                phase_vectors[key] = values
            else:
                constants[key] = constant_value(values,
                                                 f"{shape} source.xlsx_cell_values.{key}")
        else:
            constants[key] = value

    check(set(phase_vectors) == set(ALL_PHASES),
          f"{shape} phase vector set is incomplete: {sorted(phase_vectors)}")
    for key in ("output_sha256", "semantic_sha256", "untouched_member_sha256"):
        check_hex(constants[key], f"{shape} source.xlsx_cell_values.{key}")
    check(constants["output_sha256"] == result.get("output_sha256"),
          f"{shape} source output digest disagrees with result output digest")
    check(constants["payload_materializations"] == constants["cache_successful_loads"],
          f"{shape} payload materializations disagree with successful cache loads")
    check(constants["payload_materializations"] > 0,
          f"{shape} source produced no payload materializations")
    check(constants["cache_cold_loads"] == constants["cache_successful_loads"],
          f"{shape} cold-load and successful-load counts disagree")
    check(constants["cache_retained_entries"] == constants["cache_successful_loads"],
          f"{shape} retained-entry and successful-load counts disagree")
    check(constants["cache_retained_bytes"] > 0, f"{shape} cache retained no bytes")

    for key in ("cache_hits", "cache_waiter_joins", "cache_failed_loads", "cache_evictions",
                "cache_bypasses", "cache_oversized_bypasses", "cache_allocation_bypasses",
                "cache_in_flight_loads", "cache_budget_memory_used",
                "cache_budget_reserved_bytes", "cache_budget_reservation_failures",
                "budget_used_after_package_drop", "budget_used_after_handles_drop",
                "budget_objects_used_after_handles_drop", "unselected_worksheet_read_calls",
                "unselected_worksheet_read_bytes"):
        check(constants[key] == 0, f"{shape} resource invariant {key} is nonzero")

    for key in ("pre_publication_budget", "post_publication_budget"):
        check_budget_snapshot(constants[key], f"{shape} {key}", post=key.startswith("post"))
    refusal = constants["output_budget_refusal"]
    check(isinstance(refusal, dict), f"{shape} output-budget refusal evidence is missing")
    refusal_fields = {
        "successful_output_ceiling", "first_output_request_bytes", "one_under_output_limit",
        "accepted_output_bytes", "source_read_calls", "source_read_bytes", "output_bytes_used",
        "typed_output_resource_refusal", "zero_output_verified", "source_identity_preserved",
    }
    check(set(refusal) == refusal_fields, f"{shape} output-budget refusal fields changed")
    check(all(refusal[field] == 0 for field in refusal_fields
              if field not in {"typed_output_resource_refusal", "zero_output_verified",
                               "source_identity_preserved"}),
          f"{shape} unmanaged output-budget refusal has nonzero evidence")
    check(all(refusal[field] is False for field in
              {"typed_output_resource_refusal", "zero_output_verified", "source_identity_preserved"}),
          f"{shape} unmanaged output-budget refusal has a true gate")

    generic_pairs = (
        ("read_calls", "source_read_calls"),
        ("read_bytes", "source_read_bytes"),
        ("ordinary_payload_materializations", "payload_materializations"),
    )
    for generic, specific in generic_pairs:
        check(source.get(generic) == summary[specific],
              f"{shape} generic {generic} disagrees with source summary {specific}")
    generic_constants: dict[str, Any] = {}
    for key, value in source.items():
        if isinstance(value, list):
            values = check_vector(value, samples, f"{shape} source.{key}", nonnegative=False)
            check(all(item == values[0] for item in values),
                  f"{shape} generic source counter {key} varies across samples")
            generic_constants[key] = values[0]
    check(constants["source_read_calls"] > 0 and constants["source_read_bytes"] > 0,
          f"{shape} source recorded no reads")
    check(constants["workbook_read_calls"] == 1,
          f"{shape} workbook read count is not one")
    check(constants["selected_worksheet_read_calls"] >= summary["selected_worksheet_count"],
          f"{shape} selected worksheet reads are fewer than selected worksheets")
    check(constants["selected_worksheet_read_bytes"] > 0,
          f"{shape} selected worksheet reads have no bytes")
    check(constants["workbook_read_bytes"] > 0,
          f"{shape} workbook reads have no bytes")
    check(generic_constants["ordinary_payload_materializations"]
          == constants["payload_materializations"],
          f"{shape} generic materialization count disagrees with source summary")
    check(generic_constants["max_in_flight_reads"] == 1,
          f"{shape} source was not serialized to one in-flight read")

    # Every serialized source vector is acquisition-ordered.  The phase sum
    # must follow that same order before it is matched against sorted elapsed
    # samples through sample_order.
    for phase in ALL_PHASES:
        check(len(phase_vectors[phase]) == samples, f"{shape} {phase} length changed")
    phase_sums = [sum(phase_vectors[phase][index] for phase in PHASES)
                  for index in range(samples)]
    elapsed = result["elapsed_ns"]
    check([phase_sums[index] for index in elapsed["sample_order"]] == elapsed["samples"],
          f"{shape} phase sums do not align with sorted elapsed samples")
    check(sorted(range(samples), key=lambda index: (phase_sums[index], index))
          == elapsed["sample_order"],
          f"{shape} sample_order is not the phase-sum ordering")
    return phase_vectors, constants


def drift_flags(rows: list[dict[str, Any]], shapes: list[str]) -> list[dict[str, Any]]:
    flags = []
    for shape in shapes:
        pair = [row for row in rows if row["shape"] == shape]
        check(len(pair) == 2, f"{shape} does not have exactly two native repeats")
        first, second = sorted(pair, key=lambda row: row["repeat"])
        for phase in ("elapsed_ns",) + ALL_PHASES:
            left = first["elapsed_ns"] if phase == "elapsed_ns" else first["phases"][phase]
            right = second["elapsed_ns"] if phase == "elapsed_ns" else second["phases"][phase]
            for metric in ("p50", "p95", "p99", "mean"):
                check(left[metric] != 0, f"{shape} {phase}.{metric} has zero repeat-1 value")
                change = (right[metric] / left[metric] - 1.0) * 100.0
                if abs(change) > 5.0:
                    flags.append({
                        "shape": shape,
                        "phase": phase,
                        "metric": metric,
                        "repeat1": left[metric],
                        "repeat2": right[metric],
                        "change_percent": change,
                    })
    return flags


def analyze(root: Path = HERE, *, check_sources: bool = True,
            compatibility: bool = False) -> dict[str, Any]:
    """Analyze one evidence directory without writing to it.

    ``compatibility=True`` is only for the in-memory replay of the sealed 0520
    fixture: that historical report predates the current partial-sink field.
    It does not weaken the default current-head analysis.
    """

    root = Path(root).resolve()
    if check_sources:
        source_check(root)
    plan = read_json(root, "plan.json")
    build = read_json(root, "build.json")
    check(plan.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save",
          "plan selects an unexpected case")
    shapes = plan.get("shapes")
    check(shapes == ["medium", "dense-sparse"],
          f"plan shapes are not the required native pair: {shapes!r}")
    repeats = plan.get("native_repeats")
    check(repeats == 2, f"plan native_repeats is {repeats!r}; expected two")
    samples = plan.get("samples")
    check(isinstance(samples, int) and samples > 30, "plan sample count is too small")
    warmup = plan.get("warmup")
    check(isinstance(warmup, int) and warmup >= 0, "plan warmup count is invalid")
    cpu = plan.get("cpu")
    check(isinstance(cpu, int) and cpu >= 0, "plan CPU affinity is invalid")
    check("none" in str(plan.get("performance_claim", "")).lower(),
          "plan permits a performance claim")
    check("speed" not in str(plan.get("performance_claim", "")).lower()
          or "no" in str(plan.get("performance_claim", "")).lower(),
          "plan performance claim is not diagnostic-only")
    check(build.get("exit_code") == 0, "native build did not pass")
    check_hex(build.get("binary_sha256"), "build binary_sha256")
    check(build.get("source_manifest_sha256") == sha(root / "source-manifest.json"),
          "build source manifest custody mismatch")
    build_revision = build.get("revision")
    if build_revision is None:
        check(compatibility, "build receipt omitted its revision")
    else:
        check(build_revision == plan.get("revision"),
              "build and plan revisions disagree")

    expected_names = {
        f"native-r{repeat}-{shape}"
        for repeat in range(1, repeats + 1)
        for shape in shapes
    }
    receipts = []
    rows = []
    identities: dict[str, dict[str, Any]] = {}
    for repeat in range(1, repeats + 1):
        for shape in shapes:
            name = f"native-r{repeat}-{shape}"
            receipt = read_json(root, name + ".receipt.json")
            receipts.append(receipt)
            check(receipt.get("exit_code") == 0, f"{name} capture failed")
            check(receipt.get("binary_sha256") == build["binary_sha256"],
                  f"{name} binary custody mismatch")
            check(receipt.get("script_sha256") == sha(root / "capture.py"),
                  f"{name} capture script custody mismatch")
            check(receipt.get("plan_sha256") == sha(root / "plan.json"),
                  f"{name} plan custody mismatch")
            check(receipt.get("source_manifest_sha256") == sha(root / "source-manifest.json"),
                  f"{name} source-manifest custody mismatch")
            command = receipt.get("command")
            check(isinstance(command, list), f"{name} command is not a list")
            check(command[:3] == ["taskset", "-c", str(cpu)],
                  f"{name} command is not pinned to the plan CPU")
            for option, value in (("--warmup", warmup), ("--samples", samples),
                                  ("--case", plan["case"]),
                                  ("--xlsx-cell-crud-shape", shape)):
                check(command.count(option) == 1,
                      f"{name} command has an unexpected {option} count")
                index = command.index(option)
                check(index + 1 < len(command) and command[index + 1] == str(value),
                      f"{name} command has an unexpected {option} value")
            artifacts = receipt.get("artifacts")
            check(isinstance(artifacts, dict), f"{name} receipt artifacts are missing")
            expected_artifacts = {name + suffix for suffix in (".json", ".stdout", ".stderr")}
            check(set(artifacts) == expected_artifacts, f"{name} artifact set changed")
            for filename, digest in artifacts.items():
                path = root / filename
                check(path.is_file(), f"{name} artifact is absent: {filename}")
                check(sha(path) == digest, f"{name} artifact digest mismatch: {filename}")

            raw = read_json(root, name + ".json")
            check(raw.get("schema_version") == 1, f"{name} schema version changed")
            check(raw.get("binary_identity", {}).get("binary_sha256")
                  == build["binary_sha256"], f"{name} raw binary identity mismatch")
            environment = raw.get("environment", {})
            check(environment.get("git_revision") == plan["revision"],
                  f"{name} raw revision mismatch")
            check(environment.get("cpu_affinity") == str(cpu),
                  f"{name} raw CPU affinity mismatch")
            tool = raw.get("tool", {})
            check(tool.get("profile") == "release", f"{name} is not release evidence")
            check(tool.get("instrumentation") == "none", f"{name} is instrumented native evidence")
            configuration = raw.get("configuration", {})
            check(configuration.get("cases") == [plan["case"]],
                  f"{name} raw case configuration mismatch")
            check(configuration.get("xlsx_cell_crud_shapes") == [shape],
                  f"{name} raw shape configuration mismatch")
            check(configuration.get("samples_per_case") == samples,
                  f"{name} raw sample configuration mismatch")
            check(configuration.get("warmup_iterations_per_case") == warmup,
                  f"{name} raw warmup configuration mismatch")
            results = raw.get("results")
            check(isinstance(results, list) and len(results) == 1,
                  f"{name} does not contain exactly one result")
            result = results[0]
            check(result.get("case") == plan["case"], f"{name} result case mismatch")
            corpus = validate_corpus(result.get("corpus", {}), shape)
            check(result.get("sink") is not None, f"{name} sink summary is missing")
            validate_sink(result["sink"], name)
            elapsed = result.get("elapsed_ns")
            check(isinstance(elapsed, dict), f"{name} elapsed summary is missing")
            measured = check_reported_statistics(elapsed, samples)
            phase_vectors, constants = validate_source(
                result, corpus, shape, samples, compatibility=compatibility
            )
            phases = {phase: stats(phase_vectors[phase]) for phase in ALL_PHASES}
            denominator = sum(elapsed["samples"])
            check(denominator > 0, f"{name} elapsed denominator is zero")
            shares = {phase: sum(phase_vectors[phase]) / denominator for phase in PHASES}
            close(sum(shares.values()), 1.0, f"{name} aggregate phase shares")
            digest = constants["output_sha256"]
            semantic = constants["semantic_sha256"]
            untouched_digest = constants["untouched_member_sha256"]
            identity = {
                "corpus": corpus,
                "sink": result["sink"],
                "output_sha256": digest,
                "semantic_sha256": semantic,
                "untouched_member_count": constants["untouched_member_count"],
                "untouched_member_sha256": untouched_digest,
                "resource_constants": {
                    key: value for key, value in constants.items() if key not in ALL_PHASES
                },
            }
            if shape in identities:
                check(identities[shape] == identity,
                      f"{shape} corpus/output/resource identity changed between repeats")
            else:
                identities[shape] = identity
            rows.append({
                "name": name,
                "shape": shape,
                "repeat": repeat,
                "samples": samples,
                "elapsed_ns": measured,
                "phases": phases,
                "aggregate_time_share": shares,
            })

    actual_receipt_names = {
        path.name[:-len(".receipt.json")]
        for path in root.glob("native-r*-*.receipt.json")
    }
    check(actual_receipt_names == expected_names, "native receipt set is incomplete or has extras")
    ordered_receipts = []
    for receipt in receipts:
        try:
            start = datetime.datetime.fromisoformat(receipt["start_utc"])
            end = datetime.datetime.fromisoformat(receipt["end_utc"])
        except (KeyError, TypeError, ValueError) as error:
            raise AssertionError("receipt timestamp is invalid") from error
        check(start <= end, "receipt end precedes receipt start")
        check(isinstance(receipt.get("seconds"), (int, float)) and receipt["seconds"] > 0,
              "receipt duration is not positive")
        ordered_receipts.append((start, end, receipt))
    ordered_receipts.sort(key=lambda item: item[0])
    check(all(left[1] <= right[0] for left, right in zip(ordered_receipts, ordered_receipts[1:])),
          "native captures overlap")

    flags = drift_flags(rows, shapes)
    return {
        "status": "pass",
        "claim": "current-head phase attribution; no historical before-after speed claim",
        "case": plan["case"],
        "revision": plan["revision"],
        "native_repeats": repeats,
        "total_samples": sum(row["samples"] for row in rows),
        "rows": rows,
        "repeat_variation_over_five_percent": flags,
        "repeat_drift_over_five_percent": flags,
        "identities": identities,
        "checks": {
            "receipt_custody_and_serialization": True,
            "phase_sums_recomputed_in_sample_order": True,
            "percentiles_and_intervals_recomputed": True,
            "aggregate_phase_shares_recomputed": True,
            "correctness_and_resource_invariants": True,
            "corpus_and_output_consistency_per_shape": True,
            "repeat_drift_threshold_percent": 5.0,
        },
        "uncertainty": (
            "Two fresh serialized children per shape; phase percentiles and aggregate shares are "
            "descriptive. Within-child samples are not independent host repetitions. The report "
            "contains no historical before-after comparison."
        ),
        "unavailable": [
            "operation-local allocations",
            "peak RSS",
            "physical provider I/O",
            "cold cache",
            "parallel scaling",
            "native Office producer",
        ],
    }


if __name__ == "__main__":
    output = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "analysis.json"
    report = analyze()
    output.write_text(json.dumps(report, indent=2) + "\n")
    print("Current-head native evidence and phase sums verified:", output)
