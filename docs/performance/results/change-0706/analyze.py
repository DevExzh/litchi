#!/usr/bin/env python3
"""Independently verify and compare the 0706 XLSX ABBA evidence packet.

The verifier is intentionally self-contained.  It does not import an older
packet helper, consult the current git revision, build a binary, or run a
benchmark.  It checks the retained source maps and every child receipt first,
then recomputes phase sums, descriptive statistics, paired deltas, parity,
repeat drift, the AA floor, and the frozen admission gates.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
from pathlib import Path
import statistics
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]

TIMING_PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
ALL_PHASES = TIMING_PHASES + ("reopen_ns",)
ALLOCATION_FIELDS = (
    "plan_allocation_metrics",
    "staging_allocation_metrics",
    "commit_allocation_metrics",
    "commit_core_allocation_metrics",
    "publication_allocation_metrics",
)
ALLOCATION_METRICS = (
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

NATIVE_PHASES = (
    "baseline-noise1",
    "baseline-noise2",
    "baseline-A1",
    "candidate-B1",
    "candidate-B2",
    "baseline-A2",
)
FULL_PHASES = ("baseline-A1", "candidate-B1", "candidate-B2", "baseline-A2")
ALLOC_PHASES = ("allocator-baseline", "allocator-candidate")
ROLE_FOR_PHASE = {
    "baseline-noise1": "baseline",
    "baseline-noise2": "baseline",
    "baseline-A1": "baseline",
    "candidate-B1": "candidate",
    "candidate-B2": "candidate",
    "baseline-A2": "baseline",
    "allocator-baseline": "baseline",
    "allocator-candidate": "candidate",
}
ALIASES = {
    "baseline-noise-1": "baseline-noise1",
    "baseline-noise-2": "baseline-noise2",
}

# A rejected candidate may leave an independently useful integration test in
# the candidate crate after production files are restored.  The final source
# check accepts that narrow, explicit residue while requiring every production
# source entry to equal either the retained candidate or the restored baseline.
RESTORED_TEST_ROOTS = ("crates/xml-minifier/tests/",)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def json_digest(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read(path: Path) -> Any:
    require(path.is_file(), f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        raise AssertionError(f"invalid JSON in {path}: {error}") from error


def check_hex(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def close(left: float, right: float, label: str) -> None:
    require(math.isclose(left, right, rel_tol=1e-12, abs_tol=1e-7),
            f"{label}: {left!r} != {right!r}")


def midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def nearest_rank(values: list[int], percentile: int) -> int:
    index = ((percentile * len(values) + 99) // 100) - 1
    return values[min(index, len(values) - 1)]


def numeric_stats(values: Iterable[int]) -> dict[str, Any]:
    values = list(values)
    require(values, "numeric vector is empty")
    for index, value in enumerate(values):
        nonnegative_int(value, f"numeric vector[{index}]")
    ordered = sorted(values)
    return {
        "count": len(values),
        "p50": midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
        "p95": nearest_rank(ordered, 95),
        "p99": nearest_rank(ordered, 99),
        "mean": statistics.mean(values),
        "min": ordered[0],
        "max": ordered[-1],
    }


def elapsed_stats(values: list[int]) -> dict[str, Any]:
    result = numeric_stats(values)
    deviation = statistics.stdev(values) if len(values) > 1 else 0.0
    result["standard_deviation"] = deviation
    # The runner emits a Student-t interval.  The packet includes both large
    # (native) and tiny (allocator) samples; the latter need an exact t table
    # rather than the large-df approximation, so interval bounds are checked
    # for shape and containment below while all descriptive statistics are
    # recomputed exactly.
    result["confidence_interval_95"] = {
        "method": "two-sided Student's t interval for the mean",
    }
    return result


def vector(value: Any, count: int, label: str, *, nonnegative: bool = True) -> list[Any]:
    require(isinstance(value, list), f"{label} is not a list")
    require(len(value) == count, f"{label} has {len(value)} values; expected {count}")
    if nonnegative:
        for index, item in enumerate(value):
            nonnegative_int(item, f"{label}[{index}]")
    return value


def report_stats(elapsed: Any, samples: int, label: str) -> tuple[list[int], list[int]]:
    require(isinstance(elapsed, dict), f"{label}.elapsed_ns is not an object")
    require(elapsed.get("unit") == "ns", f"{label}.elapsed_ns unit changed")
    values = vector(elapsed.get("samples"), samples, f"{label}.elapsed_ns.samples")
    require(all(value > 0 for value in values), f"{label} elapsed sample is zero")
    order = vector(elapsed.get("sample_order"), samples, f"{label}.sample_order")
    require(sorted(order) == list(range(samples)), f"{label} sample_order is not a permutation")
    require(values == sorted(values), f"{label} elapsed samples are not sorted")
    expected = elapsed_stats(values)
    for key in ("p50", "p95", "p99", "min", "max"):
        require(elapsed.get(key) == expected[key], f"{label}.elapsed_ns.{key} is stale")
    close(float(elapsed.get("mean")), float(expected["mean"]), f"{label}.elapsed_ns.mean")
    close(float(elapsed.get("standard_deviation")),
          float(expected["standard_deviation"]), f"{label}.elapsed_ns.standard_deviation")
    interval = elapsed.get("confidence_interval_95")
    require(isinstance(interval, dict), f"{label} confidence interval is missing")
    require(interval.get("method") == expected["confidence_interval_95"]["method"],
            f"{label} confidence interval method changed")
    lower = interval.get("lower")
    upper = interval.get("upper")
    finite(lower, f"{label} confidence interval lower")
    finite(upper, f"{label} confidence interval upper")
    require(float(lower) <= float(upper), f"{label} confidence interval is inverted")
    require(float(lower) <= float(expected["mean"]) <= float(upper),
            f"{label} confidence interval does not contain the mean")
    return values, order


def acquisition(sorted_values: list[int], order: list[int]) -> list[int]:
    result: list[int | None] = [None] * len(sorted_values)
    for sorted_index, original_index in enumerate(order):
        require(result[original_index] is None, "sample_order repeats an index")
        result[original_index] = sorted_values[sorted_index]
    require(all(value is not None for value in result), "sample_order leaves a gap")
    return [int(value) for value in result]


def source_census() -> dict[str, str]:
    paths: list[Path] = [REPO / "Cargo.toml", REPO / "Cargo.lock"]
    paths.extend(path for path in (REPO / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path for path in (REPO / folder).rglob("*")
            if path.is_file() and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {str(path.relative_to(REPO)): sha(path) for path in sorted(set(paths))}


def load_plan() -> dict[str, Any]:
    p = read(HERE / "plan.json")
    require(isinstance(p, dict), "plan is not an object")
    primary = p.get("primary")
    require(isinstance(primary, dict), "plan primary is missing")
    require(primary.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save",
            "unexpected primary case")
    require(primary.get("shapes") == ["medium", "dense-sparse"],
            "unexpected primary shape order")
    require(primary.get("repeats") == 2 and primary.get("warmup") == 20
            and primary.get("samples") == 200, "primary sample plan changed")
    require(p.get("guard_samples") == 30 and p.get("guard_warmup") == 10,
            "guard sample plan changed")
    require(p.get("producer") == "xlsx_producer_medium_source_one_edit_save",
            "producer case changed")
    expanded = expand_guards(p)
    require(len(expanded) == 7 and len(set(expanded)) == 7,
            "guard matrix is not seven unique children")
    allocation = p.get("allocation")
    require(isinstance(allocation, dict), "allocation plan is missing")
    require(allocation.get("shapes") == ["medium", "dense-sparse"],
            "allocation shapes changed")
    require(allocation.get("repeats") == 2 and allocation.get("samples") == 5
            and allocation.get("warmup") == 0, "allocation sample plan changed")
    gates = p.get("gates")
    require(isinstance(gates, dict), "admission gates are missing")
    expected_gates = {
        "total_p50_reduction_percent": 2.0,
        "total_mean_reduction_percent": 2.0,
        "publication_p50_reduction_percent": 5.0,
        "allocation_calls_reduction_percent": 8.0,
        "require_every_shape_repeat": True,
    }
    for key, expected in expected_gates.items():
        require(gates.get(key) == expected, f"plan gate {key} changed")
    require(p.get("candidate_roots") == ["crates/xml-minifier/"],
            "candidate source roots changed")
    require(p.get("order") == [
        "baseline-noise-1", "baseline-noise-2", "baseline-A1",
        "candidate-B1", "candidate-B2", "baseline-A2",
    ], "ABBA order changed")
    micro = p.get("audit_microguards")
    require(isinstance(micro, dict), "audit_microguards plan is missing")
    require(micro.get("order") == ["A1", "B1", "B2", "A2"],
            "microguard order changed")
    return p


def expand_guards(p: dict[str, Any]) -> list[tuple[str, str]]:
    guards = p.get("guards")
    require(isinstance(guards, list), "guards is not a list")
    result: list[tuple[str, str]] = []
    for guard in guards:
        require(isinstance(guard, dict), "guard entry is not an object")
        case = guard.get("case")
        shapes = guard.get("shapes")
        require(isinstance(case, str) and isinstance(shapes, list), "malformed guard entry")
        for shape in shapes:
            require(isinstance(shape, str) and shape, "malformed guard shape")
            result.append((case, shape))
    return result


def expected_jobs(p: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    primary = p["primary"]
    if phase in NATIVE_PHASES:
        full = phase in FULL_PHASES
        jobs: list[dict[str, Any]] = []
        for repeat in range(1, primary["repeats"] + 1):
            for shape in (primary["shapes"] if repeat == 1 else list(reversed(primary["shapes"]))):
                jobs.append({
                    "kind": "primary", "repeat": repeat, "shape": shape,
                    "case": primary["case"], "samples": primary["samples"],
                    "warmup": primary["warmup"],
                    "name": f"native-{phase}-r{repeat}-{shape}",
                })
        if not full:
            return jobs
        for case, shape in expand_guards(p):
            jobs.append({
                "kind": "guard", "repeat": None, "shape": shape, "case": case,
                "samples": p["guard_samples"], "warmup": p["guard_warmup"],
                "name": f"guard-{phase}-{case}-{shape}",
            })
        for repeat in range(1, primary["repeats"] + 1):
            jobs.append({
                "kind": "producer", "repeat": repeat, "shape": "medium",
                "case": p["producer"], "samples": primary["samples"],
                "warmup": primary["warmup"],
                "name": f"producer-{phase}-r{repeat}",
            })
        return jobs
    require(phase in ALLOC_PHASES, f"unknown phase {phase}")
    config = p["allocation"]
    jobs = []
    for repeat in range(1, config["repeats"] + 1):
        for shape in (config["shapes"] if repeat == 1 else list(reversed(config["shapes"]))):
            jobs.append({
                "kind": "allocation", "repeat": repeat, "shape": shape,
                "case": primary["case"], "samples": config["samples"],
                "warmup": config["warmup"],
                "name": f"alloc-{phase}-r{repeat}-{shape}",
            })
    return jobs


def cleanup_binary_witnesses() -> list[dict[str, Any]]:
    """Read optional exact witnesses for binaries removed after capture.

    A cleanup record is deliberately treated as an identity witness only.  It
    cannot make a changed or missing receipt valid; it supplies the digest and
    byte count for the frozen executable when the external copy has been
    removed after the measurement packet was sealed.
    """

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
                    if size is not None:
                        nonnegative_int(size, f"{filename}:{raw_path}.bytes")
                    witnesses.append({"path": raw_path, "sha256": digest, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return witnesses


def binary_witness_matches(path: Path, digest: str, size: int,
                           witnesses: list[dict[str, Any]]) -> bool:
    target = str(path.resolve())
    for witness in witnesses:
        raw_path = witness["path"]
        candidate = Path(raw_path)
        resolved = str((REPO / candidate).resolve()
                       if not candidate.is_absolute() else candidate.resolve())
        if (resolved == target and witness["sha256"] == digest
                and (witness["bytes"] is None or witness["bytes"] == size)):
            return True
    return False


def build_info(role: str, lane: str,
               witnesses: list[dict[str, Any]] | None = None,
               ) -> tuple[dict[str, Any], dict[str, str]]:
    path = HERE / f"build-{role}.json"
    records = read(path)
    require(isinstance(records, list), f"{path} is not a list")
    expected_binary = f"{role}-{'native' if lane == 'native' else 'alloc'}"
    matches = [record for record in records if isinstance(record, dict)
               and Path(str(record.get("binary", ""))).name == expected_binary]
    require(len(matches) == 1, f"{path} has no unique {expected_binary} record")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{expected_binary} build failed")
    binary = Path(str(record.get("binary", "")))
    digest = record.get("binary_sha256")
    check_hex(digest, f"{expected_binary} build digest")
    if binary.is_file() and not binary.is_symlink():
        require(sha(binary) == digest, f"{expected_binary} binary digest changed")
        binary_custody = "live-file"
    else:
        if witnesses is None:
            witnesses = cleanup_binary_witnesses()
        recorded_bytes = record.get("binary_bytes")
        require(isinstance(recorded_bytes, int) and recorded_bytes > 0,
                f"{expected_binary} build byte count is invalid")
        require(binary_witness_matches(binary, digest, recorded_bytes,
                                       witnesses),
                f"missing binary {binary} has no exact cleanup witness")
        binary_custody = "post-cleanup-witness"
    manifest_path = HERE / f"source-{role}.json"
    expected = read(manifest_path)
    require(isinstance(expected, dict), f"{manifest_path} is not a source map")
    require(record.get("source_manifest_sha256") == sha(manifest_path),
            f"{expected_binary} build/source custody mismatch")
    return {
        "role": role,
        "lane": lane,
        "binary": str(binary.resolve()),
        "binary_sha256": digest,
        "binary_bytes": (binary.stat().st_size if binary.is_file() else record["binary_bytes"]),
        "binary_custody": binary_custody,
        "build_path": str(path),
        "build_sha256": sha(path),
        "source_manifest": f"source-{role}.json",
        "source_manifest_sha256": sha(manifest_path),
        "expected_source": expected,
    }, expected


def source_relation(role: str, expected: dict[str, str], current: dict[str, str],
                    p: dict[str, Any]) -> dict[str, Any]:
    changed = sorted(name for name in set(expected) | set(current)
                     if expected.get(name) != current.get(name))
    allowed = p["candidate_roots"]
    if role == "candidate":
        require(not changed, f"candidate source differs from retained map: {changed}")
        mode = "exact"
    else:
        require(all(any(name.startswith(root) for root in allowed) for name in changed),
                f"baseline source changed outside allowed roots: {changed}")
        mode = "exact" if not changed else "baseline-retained-under-allowed-candidate-delta"
    return {"mode": mode, "changed_paths": changed,
            "allowed_roots": list(allowed),
            "expected_entry_count": len(expected), "current_entry_count": len(current)}


def restored_test_path(name: str) -> bool:
    return any(name.startswith(root) for root in RESTORED_TEST_ROOTS)


def explicit_outcome() -> str | None:
    """Return a coordinator's final retention outcome when one is recorded."""

    values: list[str] = []
    raw_env = os.environ.get("LITCHI_0706_OUTCOME")
    if raw_env:
        values.append(raw_env)
    for filename in ("decision.json", "retention.json", "outcome.json"):
        path = HERE / filename
        if not path.is_file():
            continue
        value = read(path)
        require(isinstance(value, dict), f"{filename} is not an object")
        for key in ("retained", "candidate_retained", "keep_candidate"):
            if key in value:
                require(isinstance(value[key], bool), f"{filename}.{key} is not boolean")
                values.append("retained" if value[key] else "rejected")
        for key in ("outcome", "candidate_outcome", "decision", "retention"):
            item = value.get(key)
            if isinstance(item, str):
                values.append(item)
    if not values:
        return None
    normalized: list[str] = []
    for value in values:
        item = value.lower().replace("_", "-").replace(" ", "-")
        if item in {"retained", "retain", "accepted", "kept", "keep", "applied"}:
            normalized.append("retained")
        elif item in {"rejected", "reject", "restored", "restore", "discarded", "drop"}:
            normalized.append("rejected")
        elif item in {"retained-candidate", "candidate-retained"}:
            normalized.append("retained")
        elif item in {"rejected-candidate", "candidate-rejected", "restored-baseline"}:
            normalized.append("rejected")
        else:
            # A decision artifact may contain a descriptive pilot decision in
            # addition to a boolean outcome.  Ignore unrelated prose, but do
            # not silently accept a value that looks like an outcome and is
            # contradictory or malformed.
            continue
    if not normalized:
        return None
    require(all(item == normalized[0] for item in normalized),
            "retention outcome artifacts disagree")
    return normalized[0]


def final_source_state(
    baseline: dict[str, str], candidate: dict[str, str], current: dict[str, str],
) -> dict[str, Any]:
    """Validate either final source state without confusing it with the pilot gate.

    The analyzer may run while the candidate checkout is live, or later after
    production files have been restored and an independent integration test is
    retained.  In both cases the historical binary/source bindings are checked
    from the receipts; this function checks only the final checkout relation.
    """

    outcome = explicit_outcome()
    inferred = outcome is None
    if outcome is None:
        candidate_delta = sorted(name for name in set(candidate) | set(current)
                                 if candidate.get(name) != current.get(name))
        baseline_delta = sorted(name for name in set(baseline) | set(current)
                                if baseline.get(name) != current.get(name))
        if all(restored_test_path(name) for name in candidate_delta):
            outcome = "retained"
        elif all(restored_test_path(name) for name in baseline_delta):
            outcome = "rejected"
        else:
            require(False, "final checkout matches neither candidate nor restored baseline")
    if outcome == "retained":
        differences = sorted(name for name in set(candidate) | set(current)
                             if candidate.get(name) != current.get(name))
        require(all(restored_test_path(name) for name in differences),
                f"retained final checkout differs outside independent tests: {differences}")
        production_differences = [name for name in differences if not restored_test_path(name)]
        return {
            "outcome": outcome,
            "inferred": inferred,
            "production_matches": True,
            "independent_test_paths": [name for name in differences if restored_test_path(name)],
            "production_difference_paths": production_differences,
        }
    require(outcome == "rejected", f"unknown final checkout outcome: {outcome}")
    differences = sorted(name for name in set(baseline) | set(current)
                         if baseline.get(name) != current.get(name))
    require(all(restored_test_path(name) for name in differences),
            f"restored final checkout differs outside independent tests: {differences}")
    production_differences = [name for name in differences if not restored_test_path(name)]
    return {
        "outcome": outcome,
        "inferred": inferred,
        "production_matches": True,
        "independent_test_paths": [name for name in differences if restored_test_path(name)],
        "production_difference_paths": production_differences,
    }


def validate_sink(sink: Any, label: str) -> None:
    require(isinstance(sink, dict), f"{label} sink is missing")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        nonnegative_int(sink.get(key), f"{label}.sink.{key}")
    buckets = sink.get("write_size_buckets")
    require(isinstance(buckets, dict), f"{label} sink buckets are missing")
    expected = {"bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
                "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536"}
    require(set(buckets) == expected, f"{label} sink bucket set changed")
    for key, value in buckets.items():
        nonnegative_int(value, f"{label}.sink.{key}")
    require(sum(buckets.values()) == sink["write_calls"],
            f"{label} sink bucket total differs from write_calls")
    require(sink["accepted_bytes"] > 0 and sink["largest_write"] <= 65536
            and buckets["bytes_over_65536"] == 0,
            f"{label} sequential sink bound failed")


def validate_corpus(corpus: Any, shape: str, label: str) -> dict[str, Any]:
    require(isinstance(corpus, dict), f"{label}.corpus is missing")
    require(corpus.get("package_format") == "XLSX/OPC/ZIP", f"{label} package format changed")
    require(isinstance(corpus.get("generator"), str) and corpus["generator"],
            f"{label} corpus generator is missing")
    for key in ("entry_count", "archive_member_count", "entry_bytes",
                "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        nonnegative_int(corpus.get(key), f"{label}.corpus.{key}")
        require(corpus[key] > 0, f"{label}.corpus.{key} is zero")
    for key in ("archive_sha256", "target_payload_sha256"):
        check_hex(corpus.get(key), f"{label}.corpus.{key}")
    xlsx = corpus.get("xlsx")
    require(isinstance(xlsx, dict), f"{label}.corpus.xlsx is missing")
    for key in ("sheet_count", "rows_per_sheet", "columns_per_sheet"):
        nonnegative_int(xlsx.get(key), f"{label}.corpus.xlsx.{key}")
        require(xlsx[key] > 0, f"{label}.corpus.xlsx.{key} is zero")
    members = xlsx.get("source_members")
    require(isinstance(members, dict), f"{label} source members are missing")
    require(isinstance(members.get("workbook"), str) and members["workbook"],
            f"{label} workbook member is missing")
    worksheets = members.get("worksheets")
    require(isinstance(worksheets, list) and len(worksheets) == xlsx["sheet_count"],
            f"{label} worksheet member count changed")
    require(len(set(worksheets)) == len(worksheets) and all(isinstance(x, str) and x for x in worksheets),
            f"{label} worksheet member identity changed")
    return corpus


def allocation_sample(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an allocation object")
    require(value.get("status") == "measured", f"{label} is not measured")
    require(value.get("scope") == "operation_global_system_allocator",
            f"{label} allocation scope changed")
    for key in ALLOCATION_METRICS:
        nonnegative_int(value.get(key), f"{label}.{key}")
    require(value["failed_allocation_calls"] == 0,
            f"{label} recorded failed allocation calls")
    return {key: value[key] for key in ALLOCATION_METRICS}


def validate_budget_summary(summary: dict[str, Any], label: str, managed: bool) -> None:
    if managed:
        require(summary.get("cache_budget_managed") is True, f"{label} is not managed evidence")
        require(summary.get("payload_memory_limit") is not None,
                f"{label} managed payload limit is missing")
    else:
        require(summary.get("cache_budget_managed") is False, f"{label} is not unmanaged evidence")
        for key in ("payload_memory_limit", "publication_planning_memory_headroom",
                    "cache_budget_memory_limit"):
            require(summary.get(key) is None, f"{label}.{key} is unexpectedly managed")


def validate_source(
    result: dict[str, Any], job: dict[str, Any], *, allocator: bool,
) -> tuple[dict[str, Any] | None, dict[str, Any] | None]:
    source = result.get("source")
    label = job["name"]
    if source is None:
        require(job["kind"] == "producer", f"{label} omitted source evidence")
        return None, None
    require(isinstance(source, dict), f"{label}.source is not an object")
    summary = source.get("xlsx_cell_values")
    require(isinstance(summary, dict), f"{label} XLSX source summary is missing")
    managed = "managed" in job["case"]
    require(summary.get("implementation") in {"source-backed", "managed-source-backed"},
            f"{label} source implementation changed")
    require(summary.get("cache_mode") in {"unmanaged-control", "managed-budget"},
            f"{label} source cache mode changed")
    validate_budget_summary(summary, label, managed)
    samples = job["samples"]
    phase_vectors: dict[str, list[int]] = {}
    allocation_vectors: dict[str, list[dict[str, Any]]] = {}
    constants: dict[str, Any] = {}
    for key, value in summary.items():
        if key in ALLOCATION_FIELDS:
            if value is None:
                require(not allocator, f"{label} allocator omitted {key}")
                continue
            values = vector(value, samples, f"{label}.{key}", nonnegative=False)
            if allocator:
                allocation_vectors[key] = [
                    allocation_sample(item, f"{label}.{key}[{index}]")
                    for index, item in enumerate(values)
                ]
            else:
                require(all(isinstance(item, dict) and item.get("status") == "unavailable"
                             for item in values), f"{label} native {key} is measured")
            continue
        if isinstance(value, list):
            values = vector(value, samples, f"{label}.{key}", nonnegative=False)
            if key in ALL_PHASES:
                require(all(isinstance(item, int) and not isinstance(item, bool) and item >= 0
                            for item in values), f"{label}.{key} is not a phase vector")
                phase_vectors[key] = values
            else:
                require(all(item == values[0] for item in values),
                        f"{label}.{key} varies across samples")
                constants[key] = values[0]
        else:
            constants[key] = value
    require(set(phase_vectors) == set(ALL_PHASES), f"{label} phase vector set is incomplete")
    for key in ("output_sha256", "semantic_sha256", "untouched_member_sha256"):
        check_hex(constants.get(key), f"{label}.source.{key}")
    check_hex(result.get("output_sha256"), f"{label}.output_sha256")
    require(constants["output_sha256"] == result["output_sha256"],
            f"{label} output digest disagrees with source digest")
    require(constants.get("payload_materializations", 0) > 0,
            f"{label} recorded no payload materialization")
    require(constants.get("source_read_calls", 0) > 0 and constants.get("source_read_bytes", 0) > 0,
            f"{label} recorded no source reads")
    for key in ("read_calls", "read_bytes", "ordinary_payload_materializations"):
        require(source.get(key) == summary.get({
            "read_calls": "source_read_calls",
            "read_bytes": "source_read_bytes",
            "ordinary_payload_materializations": "payload_materializations",
        }[key]), f"{label} generic source {key} disagrees")
    source_vectors: dict[str, Any] = {}
    for key, value in source.items():
        if isinstance(value, list):
            values = vector(value, samples, f"{label}.source.{key}", nonnegative=False)
            require(all(item == values[0] for item in values),
                    f"{label}.source.{key} varies across samples")
            source_vectors[key] = values[0]
    elapsed_values, elapsed_order = report_stats(result["elapsed_ns"], samples, label)
    elapsed_acq = acquisition(elapsed_values, elapsed_order)
    phase_sums = [sum(phase_vectors[phase][index] for phase in TIMING_PHASES)
                  for index in range(samples)]
    require(phase_sums == elapsed_acq, f"{label} native phase sums do not equal elapsed samples")
    return {
        "phase_vectors": phase_vectors,
        "phase_stats": {phase: elapsed_stats(values) for phase, values in phase_vectors.items()},
        "phase_allocation": allocation_vectors,
        "elapsed_acquisition": elapsed_acq,
        "source_constants": constants,
        "source_generic": {
            **{key: value for key, value in source.items()
               if key != "xlsx_cell_values" and not isinstance(value, list)},
            **source_vectors,
        },
    }, {
        "implementation": summary.get("implementation"),
        "cache_mode": summary.get("cache_mode"),
        "source_constants": constants,
    }


def logical_identity(result: dict[str, Any], source_info: dict[str, Any] | None,
                    corpus: dict[str, Any], sink: dict[str, Any]) -> dict[str, Any]:
    source = None
    if source_info is not None:
        source = {
            "generic": source_info.get("source_generic"),
            "constants": source_info.get("source_constants"),
        }
    return {
        "case": result.get("case"),
        "corpus": corpus,
        "sink": sink,
        "output_sha256": result.get("output_sha256"),
        "source": source,
    }


def validate_raw(
    raw: Any, job: dict[str, Any], build: dict[str, Any], *, allocator: bool,
) -> dict[str, Any]:
    label = job["name"]
    require(isinstance(raw, dict), f"{label} raw report is not an object")
    require(raw.get("schema_version") == 1, f"{label} schema version changed")
    identity = raw.get("binary_identity")
    require(isinstance(identity, dict), f"{label} binary identity is missing")
    require(identity.get("binary_sha256") == build["binary_sha256"],
            f"{label} raw binary identity differs from receipt")
    tool = raw.get("tool")
    require(isinstance(tool, dict) and tool.get("profile") == "release",
            f"{label} tool profile changed")
    if allocator:
        require(tool.get("instrumentation") == "system_allocator_operation_scoped",
                f"{label} allocator instrumentation changed")
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{label} allocator counter revision changed")
    else:
        require(tool.get("instrumentation") == "none", f"{label} native binary is instrumented")
    config = raw.get("configuration")
    require(isinstance(config, dict), f"{label} configuration is missing")
    require(config.get("samples_per_case") == job["samples"]
            and config.get("warmup_iterations_per_case") == job["warmup"],
            f"{label} sample configuration differs from receipt")
    require(config.get("cases") == [job["case"]], f"{label} case configuration differs")
    shapes = config.get("xlsx_cell_crud_shapes")
    require(isinstance(shapes, list) and job["shape"] in shapes,
            f"{label} shape configuration omits its shape")
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label} does not have one result")
    result = results[0]
    require(result.get("case") == job["case"], f"{label} result case differs")
    check_hex(result.get("output_sha256"), f"{label}.output_sha256")
    corpus = validate_corpus(result.get("corpus"), job["shape"], label)
    sink = result.get("sink")
    validate_sink(sink, label)
    elapsed_values, elapsed_order = report_stats(result.get("elapsed_ns"), job["samples"], label)
    source_info, identity_source = validate_source(result, job, allocator=allocator)
    identity = logical_identity(result, source_info, corpus, sink)
    return {
        "name": label,
        "kind": job["kind"],
        "repeat": job["repeat"],
        "shape": job["shape"],
        "case": job["case"],
        "samples": job["samples"],
        "raw": result,
        "corpus": corpus,
        "sink": sink,
        "elapsed_values": elapsed_values,
        "elapsed_order": elapsed_order,
        "elapsed_acquisition": acquisition(elapsed_values, elapsed_order),
        "elapsed_stats": elapsed_stats(elapsed_values),
        "source_info": source_info,
        "identity_source": identity_source,
        "identity": identity,
    }


def validate_receipt(
    receipt: Any, job: dict[str, Any], phase: str, role: str, lane: str,
    p: dict[str, Any], build: dict[str, Any], expected: dict[str, str],
) -> dict[str, Any]:
    label = job["name"]
    require(isinstance(receipt, dict), f"{label} receipt is not an object")
    require(receipt.get("schema_version") == 2 and receipt.get("name") == label,
            f"{label} receipt schema/name changed")
    for key, value in (("phase", phase), ("role", role), ("lane", lane),
                       ("kind", job["kind"]), ("repeat", job["repeat"]),
                       ("shape", job["shape"]), ("case", job["case"])):
        require(receipt.get(key) == value, f"{label} receipt {key} differs")
    require(receipt.get("exit_code") == 0, f"{label} child failed")
    require(receipt.get("cpu") == p["cpu"], f"{label} CPU affinity differs")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{label} retained binary digest differs")
    require(receipt.get("binary_path") == build["binary"], f"{label} binary path differs")
    require(receipt.get("build_record_sha256") == build["build_sha256"],
            f"{label} build receipt binding differs")
    require(receipt.get("build_source_manifest_sha256") == build["source_manifest_sha256"],
            f"{label} build source binding differs")
    require(receipt.get("plan_sha256") == sha(HERE / "plan.json"),
            f"{label} plan binding differs")
    require(receipt.get("script_sha256") == sha(HERE / "capture.py"),
            f"{label} capture script binding differs")
    require(receipt.get("constraints_sha256") == sha(HERE / "constraints.json"),
            f"{label} constraint binding differs")
    command = receipt.get("command")
    require(isinstance(command, list), f"{label} command is not a list")
    require(command[:3] == ["taskset", "-c", str(p["cpu"])],
            f"{label} command is not CPU-pinned")
    require(command[3] == build["binary"], f"{label} command binary differs")
    for option, value in (("--warmup", job["warmup"]), ("--samples", job["samples"]),
                          ("--case", job["case"]), ("--xlsx-cell-crud-shape", job["shape"])):
        require(command.count(option) == 1, f"{label} command has duplicate {option}")
        index = command.index(option)
        require(command[index + 1] == str(value), f"{label} command {option} differs")
    require("--features" not in command, f"{label} command leaks build features")
    json_option = command.index("--json")
    require(Path(command[json_option + 1]).name == label + ".json",
            f"{label} command JSON destination differs")
    if job["kind"] == "producer":
        require("--producer-evidence" in command, f"{label} producer oracle is missing")
        producer_option = command.index("--producer-evidence")
        require(Path(command[producer_option + 1]).name == label + ".producer.json",
                f"{label} producer oracle destination differs")
    else:
        require("--producer-evidence" not in command, f"{label} has unexpected producer oracle")

    retained = receipt.get("retained_binary_source")
    require(isinstance(retained, dict), f"{label} retained source block is missing")
    require(retained.get("manifest") == build["source_manifest"],
            f"{label} retained source manifest differs")
    require(retained.get("manifest_sha256") == build["source_manifest_sha256"],
            f"{label} retained source manifest hash differs")
    require(retained.get("source_census_sha256") == json_digest(expected),
            f"{label} retained source census digest differs")
    current = receipt.get("current_checkout_source")
    require(isinstance(current, dict), f"{label} current source block is missing")
    before_name = current.get("before_artifact")
    after_name = current.get("after_artifact")
    require(before_name == label + ".source-before.json" and after_name == label + ".source-after.json",
            f"{label} source census artifact names differ")
    before = read(HERE / before_name)
    after = read(HERE / after_name)
    require(isinstance(before, dict) and isinstance(after, dict),
            f"{label} source census is not a map")
    # A retained baseline child may have run before the candidate checkout was
    # applied.  Its before/after maps must match each other, while the current
    # checkout is checked independently against the role's retained source
    # map below.  Requiring the historical baseline map to equal today's map
    # would incorrectly reject the explicitly allowed xml-minifier delta.
    require(before == after, f"{label} current source census changed during child")
    require(current.get("before_sha256") == json_digest(before)
            and current.get("after_sha256") == json_digest(after),
            f"{label} current source digest differs")
    require(current.get("before_file_sha256") == sha(HERE / before_name)
            and current.get("after_file_sha256") == sha(HERE / after_name),
            f"{label} source census file hash differs")
    relation = source_relation(role, expected, before, p)
    for key in ("relation_before", "relation_after"):
        require(current.get(key) == relation, f"{label} {key} relation differs")
    require(current.get("unchanged_during_child") is True,
            f"{label} source did not remain unchanged during child")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label} artifact hash map is missing")
    expected_artifacts = {
        label + ".json", label + ".stdout", label + ".stderr",
        label + ".source-before.json", label + ".source-after.json",
    }
    if job["kind"] == "producer":
        expected_artifacts.add(label + ".producer.json")
    require(set(artifacts) == expected_artifacts, f"{label} artifact inventory differs")
    for name, digest in artifacts.items():
        path = HERE / name
        require(path.is_file() and path.parent == HERE and sha(path) == digest,
                f"{label} artifact digest differs: {name}")
    raw = read(HERE / (label + ".json"))
    return validate_raw(raw, job, build, allocator=lane == "alloc")


def row_key(row: dict[str, Any]) -> tuple[Any, ...]:
    return row["kind"], row["case"], row["shape"], row["repeat"]


def source_parity(left: dict[str, Any], right: dict[str, Any], label: str) -> None:
    require(left["identity"] == right["identity"], f"{label} logical output/source identity differs")


def percent_reduction(baseline: float, candidate: float, label: str) -> float:
    finite(baseline, label + ".baseline")
    finite(candidate, label + ".candidate")
    require(baseline > 0.0, f"{label} baseline is zero")
    return (baseline - candidate) / baseline * 100.0


def paired_delta(left: dict[str, Any], right: dict[str, Any], metric: str) -> dict[str, Any]:
    baseline = left["elapsed_acquisition"] if metric == "elapsed_ns" else left["source_info"]["phase_vectors"][metric]
    candidate = right["elapsed_acquisition"] if metric == "elapsed_ns" else right["source_info"]["phase_vectors"][metric]
    require(len(baseline) == len(candidate), f"paired vector length differs for {metric}")
    deltas = [candidate[index] - baseline[index] for index in range(len(baseline))]
    return {"delta": numeric_stats([abs(value) for value in deltas]),
            "signed_delta": {
                "count": len(deltas), "mean": statistics.mean(deltas),
                "min": min(deltas), "max": max(deltas),
            },
            "candidate_minus_baseline": deltas}


def timing_summary(row: dict[str, Any], metric: str) -> dict[str, Any]:
    if metric == "elapsed_ns":
        return row["elapsed_stats"]
    return row["source_info"]["phase_stats"][metric]


def compare_rows(left: dict[str, Any], right: dict[str, Any], *, baseline_phase: str,
                 candidate_phase: str, gate: bool) -> dict[str, Any]:
    source_parity(left, right, f"{baseline_phase}/{candidate_phase}/{row_key(left)}")
    metrics = (("elapsed_ns",) + ALL_PHASES
               if left["source_info"] is not None
               else ("elapsed_ns",))
    values: dict[str, Any] = {}
    flags: list[dict[str, Any]] = []
    for metric in metrics:
        baseline_summary = timing_summary(left, metric)
        candidate_summary = timing_summary(right, metric)
        stats_out: dict[str, Any] = {}
        for statistic in ("p50", "p95", "p99", "mean", "min", "max"):
            base_value = float(baseline_summary[statistic])
            candidate_value = float(candidate_summary[statistic])
            reduction = percent_reduction(base_value, candidate_value,
                                          f"{baseline_phase}/{candidate_phase}/{metric}.{statistic}")
            stats_out[statistic] = {
                "baseline": baseline_summary[statistic],
                "candidate": candidate_summary[statistic],
                "candidate_minus_baseline": candidate_summary[statistic] - baseline_summary[statistic],
                "reduction_percent": reduction,
            }
            if abs(reduction) > 5.0:
                flags.append({"metric": metric, "statistic": statistic,
                              "reduction_percent": reduction,
                              "baseline": baseline_summary[statistic],
                              "candidate": candidate_summary[statistic]})
        values[metric] = stats_out
        if metric == "elapsed_ns":
            values[metric]["paired"] = paired_delta(left, right, metric)
        else:
            values[metric]["paired"] = paired_delta(left, right, metric)
    total = values["elapsed_ns"]
    gates = {
        "total_p50": total["p50"]["reduction_percent"],
        "total_mean": total["mean"]["reduction_percent"],
    }
    if left["source_info"] is not None:
        gates["publication_p50"] = values["publication_ns"]["p50"]["reduction_percent"]
    return {
        "baseline_phase": baseline_phase,
        "candidate_phase": candidate_phase,
        "kind": left["kind"],
        "case": left["case"],
        "shape": left["shape"],
        "repeat": left["repeat"],
        "identity_equal": True,
        "timing": values,
        "gate_metrics": gates,
        "gate_used": gate,
        "over_five_percent_flags": flags,
    }


def primary_map(rows: dict[str, dict[tuple[Any, ...], dict[str, Any]]], phase: str) -> dict[tuple[int, str], dict[str, Any]]:
    result: dict[tuple[int, str], dict[str, Any]] = {}
    for key, row in rows[phase].items():
        if row["kind"] == "primary":
            result[(row["repeat"], row["shape"])] = row
    return result


def admission(comparisons: list[dict[str, Any]], p: dict[str, Any]) -> dict[str, Any]:
    gates = p["gates"]
    rows: list[dict[str, Any]] = []
    for comparison in comparisons:
        if comparison["kind"] != "primary":
            continue
        values = comparison["gate_metrics"]
        checks = {
            "total_p50": {
                "observed_reduction_percent": values["total_p50"],
                "required_reduction_percent": gates["total_p50_reduction_percent"],
                "passed": values["total_p50"] >= gates["total_p50_reduction_percent"],
            },
            "total_mean": {
                "observed_reduction_percent": values["total_mean"],
                "required_reduction_percent": gates["total_mean_reduction_percent"],
                "passed": values["total_mean"] >= gates["total_mean_reduction_percent"],
            },
            "publication_p50": {
                "observed_reduction_percent": values["publication_p50"],
                "required_reduction_percent": gates["publication_p50_reduction_percent"],
                "passed": values["publication_p50"] >= gates["publication_p50_reduction_percent"],
            },
        }
        rows.append({**{key: comparison[key] for key in
                        ("baseline_phase", "candidate_phase", "shape", "repeat")},
                     "checks": checks, "passed": all(item["passed"] for item in checks.values())})
    require(len(rows) == 8, f"primary admission has {len(rows)} rows; expected 8")
    return {
        "rows": rows,
        "passed": all(row["passed"] for row in rows),
        "scope": "every primary shape/repeat in both ABBA pairs",
        "thresholds": {
            "total_p50_reduction_percent": gates["total_p50_reduction_percent"],
            "total_mean_reduction_percent": gates["total_mean_reduction_percent"],
            "publication_p50_reduction_percent": gates["publication_p50_reduction_percent"],
        },
    }


def aa_floor(rows: dict[str, dict[tuple[Any, ...], dict[str, Any]]]) -> dict[str, Any]:
    pairs = [("baseline-noise1", "baseline-noise2"), ("baseline-A1", "baseline-A2")]
    values: list[dict[str, Any]] = []
    for left_phase, right_phase in pairs:
        left_rows = primary_map(rows, left_phase)
        right_rows = primary_map(rows, right_phase)
        require(set(left_rows) == set(right_rows), f"AA primary matrix differs for {left_phase}/{right_phase}")
        for key in sorted(left_rows):
            left, right = left_rows[key], right_rows[key]
            source_parity(left, right, f"AA {left_phase}/{right_phase}/{key}")
            metrics: dict[str, Any] = {}
            for metric in ("elapsed_ns",) + ALL_PHASES:
                item: dict[str, Any] = {}
                for statistic in ("p50", "mean", "p95", "p99"):
                    baseline = float(timing_summary(left, metric)[statistic])
                    other = float(timing_summary(right, metric)[statistic])
                    item[statistic] = {
                        "left": timing_summary(left, metric)[statistic],
                        "right": timing_summary(right, metric)[statistic],
                        "absolute_percent": abs(other - baseline) / baseline * 100.0,
                    }
                metrics[metric] = item
            values.append({"left_phase": left_phase, "right_phase": right_phase,
                           "shape": key[1], "repeat": key[0], "metrics": metrics})
    maximum = max(
        item["absolute_percent"]
        for row in values for metric in row["metrics"].values()
        for item in metric.values()
    )
    return {"pairs": values, "maximum_absolute_percent": maximum,
            "scope": "baseline noise pair and baseline A1/A2 pair; no fixed pass threshold"}


def allocator_report(
    rows: dict[str, dict[tuple[Any, ...], dict[str, Any]]], p: dict[str, Any],
) -> dict[str, Any]:
    left_rows = rows["allocator-baseline"]
    right_rows = rows["allocator-candidate"]
    require(set(left_rows) == set(right_rows), "allocator baseline/candidate matrix differs")
    comparisons: list[dict[str, Any]] = []
    for key in sorted(left_rows):
        left, right = left_rows[key], right_rows[key]
        source_parity(left, right, f"allocator/{key}")
        left_alloc = left["source_info"]["phase_allocation"]
        right_alloc = right["source_info"]["phase_allocation"]
        require(set(left_alloc) == set(ALLOCATION_FIELDS) and set(right_alloc) == set(ALLOCATION_FIELDS),
                f"allocator/{key} allocation phase set incomplete")
        metrics: dict[str, Any] = {}
        flags: list[dict[str, Any]] = []
        for scope in ALLOCATION_FIELDS:
            scope_metrics: dict[str, Any] = {}
            for metric in ALLOCATION_METRICS:
                baseline = [sample[metric] for sample in left_alloc[scope]]
                candidate = [sample[metric] for sample in right_alloc[scope]]
                base_stats = numeric_stats(baseline)
                candidate_stats = numeric_stats(candidate)
                if base_stats["p50"] == 0:
                    reduction: float | None = None
                    require(candidate_stats["p50"] == 0,
                            f"allocator/{key}/{scope}/{metric} has zero baseline and nonzero candidate")
                else:
                    reduction = percent_reduction(
                        float(base_stats["p50"]), float(candidate_stats["p50"]),
                        f"allocator/{key}/{scope}/{metric}")
                scope_metrics[metric] = {
                    "baseline": base_stats, "candidate": candidate_stats,
                    "candidate_minus_baseline": [candidate[index] - baseline[index]
                                                  for index in range(len(baseline))],
                    "p50_reduction_percent": reduction,
                }
                if reduction is not None and abs(reduction) > 5.0:
                    flags.append({"scope": scope, "metric": metric,
                                  "p50_reduction_percent": reduction})
            metrics[scope] = scope_metrics
        publication = metrics["publication_allocation_metrics"]["allocation_calls"]
        gate = {
            "observed_reduction_percent": publication["p50_reduction_percent"],
            "required_reduction_percent": p["gates"]["allocation_calls_reduction_percent"],
            "passed": publication["p50_reduction_percent"]
            >= p["gates"]["allocation_calls_reduction_percent"],
        }
        comparisons.append({"shape": key[2], "repeat": key[3],
                            "identity_equal": True, "metrics": metrics,
                            "publication_allocation_calls_gate": gate,
                            "over_five_percent_flags": flags})
    require(len(comparisons) == 4, f"allocator comparison has {len(comparisons)} rows; expected 4")
    return {
        "comparisons": comparisons,
        "passed": all(row["publication_allocation_calls_gate"]["passed"] for row in comparisons),
        "threshold_percent": p["gates"]["allocation_calls_reduction_percent"],
        "scope": "all allocator metrics retained; only publication allocation_calls p50 gates",
    }


def validate_phase_matrix(
    p: dict[str, Any], phase: str, build: dict[str, Any], expected: dict[str, str],
) -> dict[tuple[Any, ...], dict[str, Any]]:
    role = ROLE_FOR_PHASE[phase]
    lane = "alloc" if phase in ALLOC_PHASES else "native"
    jobs = expected_jobs(p, phase)
    actual_receipts = sorted(
        path.name[:-len(".receipt.json")]
        for path in HERE.glob(f"*{phase}*.receipt.json")
        if path.is_file()
    )
    expected_names = {job["name"] for job in jobs}
    # The glob is deliberately broad so an accidentally duplicated or stale
    # receipt cannot hide behind the expected subset.
    require(set(actual_receipts) == expected_names,
            f"{phase} receipt set differs: actual={sorted(actual_receipts)} expected={sorted(expected_names)}")
    rows: dict[tuple[Any, ...], dict[str, Any]] = {}
    for job in jobs:
        receipt = read(HERE / f"{job['name']}.receipt.json")
        row = validate_receipt(receipt, job, phase, role, lane, p, build, expected)
        key = (job["kind"], job["case"], job["shape"], job["repeat"])
        require(key not in rows, f"{phase} duplicate logical row {key}")
        rows[key] = row
    return rows


def repeat_flags(rows: dict[str, dict[tuple[Any, ...], dict[str, Any]]], phases: Iterable[str]) -> list[dict[str, Any]]:
    flags: list[dict[str, Any]] = []
    for phase in phases:
        by_shape: dict[tuple[str, str], list[dict[str, Any]]] = {}
        for row in rows[phase].values():
            if row["kind"] == "primary":
                by_shape.setdefault((row["case"], row["shape"]), []).append(row)
        for (case, shape), values in by_shape.items():
            require(len(values) == 2, f"{phase}/{case}/{shape} does not have two repeats")
            first, second = sorted(values, key=lambda row: row["repeat"])
            for metric in ("elapsed_ns",) + ALL_PHASES:
                for statistic in ("p50", "p95", "p99", "mean"):
                    left = timing_summary(first, metric)[statistic]
                    right = timing_summary(second, metric)[statistic]
                    change = (float(right) / float(left) - 1.0) * 100.0
                    if abs(change) > 5.0:
                        flags.append({"phase": phase, "case": case, "shape": shape,
                                      "metric": metric, "statistic": statistic,
                                      "repeat1": left, "repeat2": right,
                                      "change_percent": change})
    return flags


def analyze() -> dict[str, Any]:
    p = load_plan()
    constraints = read(HERE / "constraints.json")
    require(isinstance(constraints, dict), "constraints is not an object")
    for name, digest in constraints.items():
        path = REPO / name
        require(path.is_file() and sha(path) == digest, f"constraint changed: {name}")
    census_at_start = source_census()
    require(source_census() == census_at_start, "source census is unstable")

    cleanup_witnesses = cleanup_binary_witnesses()
    builds = {
        (role, lane): build_info(role, lane, cleanup_witnesses)[0]
        for role, lane in (("baseline", "native"), ("candidate", "native"),
                           ("baseline", "alloc"), ("candidate", "alloc"))
    }
    expected_sources = {
        role: builds[(role, "native")]["expected_source"]
        for role in ("baseline", "candidate")
    }
    rows: dict[str, dict[tuple[Any, ...], dict[str, Any]]] = {}
    for phase in NATIVE_PHASES + ALLOC_PHASES:
        lane = "alloc" if phase in ALLOC_PHASES else "native"
        build = builds[(ROLE_FOR_PHASE[phase], lane)]
        expected = build["expected_source"]
        rows[phase] = validate_phase_matrix(p, phase, build, expected)
    require(source_census() == census_at_start,
            "source census changed while analyzing the evidence packet")
    _final_checkout = final_source_state(
        expected_sources["baseline"], expected_sources["candidate"], census_at_start,
    )

    # Every logical full-matrix row must preserve exactly the same output,
    # corpus and source counters between all ABBA legs.  This is stronger than
    # comparing only primary digests and catches a guard-specific drift.
    parity_rows: list[dict[str, Any]] = []
    reference = rows["baseline-A1"]
    for phase in FULL_PHASES[1:]:
        require(set(reference) == set(rows[phase]), f"full matrix logical rows differ for {phase}")
        for key in sorted(reference):
            source_parity(reference[key], rows[phase][key], f"full parity/{phase}/{key}")
            parity_rows.append({"reference_phase": "baseline-A1", "phase": phase,
                                "kind": key[0], "case": key[1], "shape": key[2],
                                "repeat": key[3], "identity_equal": True})
    for phase in ALLOC_PHASES:
        for key, row in rows[phase].items():
            primary_key = ("primary", p["primary"]["case"], key[2], key[3])
            require(primary_key in reference, f"allocator/native parity row missing for {phase}/{key}")
            source_parity(reference[primary_key], row, f"allocator/native parity/{phase}/{key}")

    comparisons: list[dict[str, Any]] = []
    pairings = (("baseline-A1", "candidate-B1"), ("baseline-A2", "candidate-B2"))
    for baseline_phase, candidate_phase in pairings:
        require(set(rows[baseline_phase]) == set(rows[candidate_phase]),
                f"paired matrix differs for {baseline_phase}/{candidate_phase}")
        for key in sorted(rows[baseline_phase]):
            comparisons.append(compare_rows(rows[baseline_phase][key], rows[candidate_phase][key],
                                            baseline_phase=baseline_phase,
                                            candidate_phase=candidate_phase,
                                            gate=key[0] == "primary"))

    native_admission = admission(comparisons, p)
    allocation = allocator_report(rows, p)
    aa = aa_floor(rows)
    drift = repeat_flags(rows, FULL_PHASES + ALLOC_PHASES)
    all_flags = [
        {"scope": "paired", **flag}
        for comparison in comparisons for flag in comparison["over_five_percent_flags"]
    ] + [{"scope": "repeat", **flag} for flag in drift]
    pilot_passed = native_admission["passed"] and allocation["passed"]
    pilot_decision = "retain-for-profiling-only" if pilot_passed else "reject"
    retention_requirements = {
        "publication_ir_reduction_percent": {
            "required": p["gates"].get("publication_ir_reduction_percent"),
            "status": "deferred; no IR profile is part of this measurement-only matrix",
            "passed": None,
        },
        "full_correctness_and_consumer_quality": {
            "status": "separate final audit evidence required",
            "passed": None,
        },
        "final_audit_decision": {
            "status": "separate final audit decision required",
            "passed": None,
        },
    }
    return {
        "schema_version": 1,
        "status": "pass",
        "decision": pilot_decision,
        "decision_scope": "pilot-native-and-allocator-gates-only",
        "pilot_decision": pilot_decision,
        "retention_requires": retention_requirements,
        "revision": p["revision"],
        "scope": "0706 XLSX source-backed XML-minifier matched-source ABBA evidence",
        "claim": "No performance claim is admitted unless every frozen primary and allocator gate passes.",
        "phase_order": p["order"],
        "native": {
            "phases": list(NATIVE_PHASES),
            "rows": [
                {"phase": phase, **{
                    "kind": row["kind"], "case": row["case"], "shape": row["shape"],
                    "repeat": row["repeat"], "samples": row["samples"],
                    "elapsed": row["elapsed_stats"],
                    "phases": row["source_info"]["phase_stats"] if row["source_info"] else None,
                }}
                for phase in NATIVE_PHASES for row in rows[phase].values()
            ],
        },
        "comparisons": comparisons,
        "aa_floor": aa,
        "repeat_drift_over_five_percent": drift,
        "over_five_percent_flags": all_flags,
        "parity": {
            "full_matrix_rows": parity_rows,
            "exact_output_corpus_source_counter_identity": True,
            "scope": "all primary, guard and producer rows across full native legs; allocator rows against native primary identity",
        },
        "admission": {
            "native": native_admission,
            "allocator": allocation,
            "passed": pilot_passed,
            "scope": "pilot eligibility for profiling only; not final retention",
            "require_every_shape_repeat": True,
        },
        "source_proof": {
            "baseline_manifest": "source-baseline.json",
            "candidate_manifest": "source-candidate.json",
            "baseline_allowed_delta": p["candidate_roots"],
            "baseline_binary_may_run_under_candidate_checkout": True,
            "binary_source_and_current_checkout_bound_separately": True,
            "all_child_receipts_verified": True,
            "final_checkout_validated": True,
            "final_checkout_modes": [
                "candidate production with optional independent tests",
                "restored baseline production with optional independent tests",
            ],
            "post_cleanup_binary_witness_supported": True,
        },
        "checks": {
            "constraints_unchanged": True,
            "source_census_stable_during_children": True,
            "native_phase_sums_recomputed": True,
            "elapsed_statistics_recomputed": True,
            "paired_sample_deltas_recomputed": True,
            "exact_output_corpus_source_parity": True,
            "allocator_full_metric_vectors_verified": True,
            "allocator_native_identity_parity": True,
            "aa_floor_reported": True,
            "over_five_percent_flags_reported": True,
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_positional", nargs="?", type=Path)
    parser.add_argument("--output", dest="output_option", type=Path)
    args = parser.parse_args()
    require_output = args.output_option or args.output_positional or HERE / "analysis.json"
    try:
        report = analyze()
        require_output.parent.mkdir(parents=True, exist_ok=True)
        require_output.write_text(json.dumps(report, indent=2) + "\n")
    except (AssertionError, OSError, ValueError, KeyError, TypeError) as error:
        print(f"analysis failed: {error}", file=__import__("sys").stderr)
        return 1
    print(f"0706 evidence verified: {require_output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
