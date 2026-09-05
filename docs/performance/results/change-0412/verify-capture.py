#!/usr/bin/env python3
"""Validate the 0412 observer-only XLS captures and print bounded summaries.

The capture consists of fresh sequential children in the order control,
candidate, candidate, control.  This checker treats elapsed time as a
descriptive per-repetition observation.  It checks the semantic and logical
source evidence strictly, with one deliberate result-level exception:
``result.source.xls.source_counter_scope`` is expected to change between the
control and candidate source observers.  Binary/build identity and the
elapsed/resource vectors are capture metadata, rather than a production
performance claim.

The only required argument is ``--root``.  Reports and catalog sidecars may
be plain JSON or JSON compressed as ``.json.zst``.  The declared ``owned``
and ``owned-allocator`` capture sets are required alongside the core
``normal`` and ``allocator`` sets.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import math
import re
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable


def _repository_operation_helper() -> Any:
    """Use the repository's strict operation validator when available.

    The checker remains runnable from a copied ``/tmp`` location: the local
    cardinality checks below are the fallback.  During normal repository
    runs, reusing ``perf_compare._validate_operation_metrics`` keeps this
    capture verifier on the same schema contract as the other evidence
    checkers.
    """

    candidates = [Path.cwd(), Path("/home/zhuhe/code/litchi")]
    for candidate in candidates:
        if not (candidate / "tools" / "perf_compare.py").is_file():
            continue
        if str(candidate) not in sys.path:
            sys.path.insert(0, str(candidate))
        try:
            from tools import perf_compare  # type: ignore

            return perf_compare
        except ImportError:
            continue
    return None


_PERF_COMPARE = _repository_operation_helper()


CASES = (
    "xls_semantic_open",
    "xls_source_backed_open",
    "xls_eager_open_list_worksheets",
    "xls_source_backed_open_list_worksheets",
    "xls_eager_open_one_cell",
    "xls_source_backed_open_one_cell",
)
OWNED_CASES = (
    "xls_owned_source_open",
    "xls_owned_source_open_list_worksheets",
    "xls_owned_source_open_one_cell",
)
SOURCE_CASES = frozenset(
    {
        "xls_source_backed_open",
        "xls_source_backed_open_list_worksheets",
        "xls_source_backed_open_one_cell",
    }
)
OPEN_CASES = frozenset({"xls_semantic_open", "xls_source_backed_open"})
LIST_CASES = frozenset(
    {"xls_eager_open_list_worksheets", "xls_source_backed_open_list_worksheets"}
)
ONE_CELL_CASES = frozenset(
    {"xls_eager_open_one_cell", "xls_source_backed_open_one_cell"}
)
ALLOWED_SCOPE_PATH = ("source", "xls", "source_counter_scope")
FLAGS = "-C force-frame-pointers=yes -C force-unwind-tables=yes"
EXPECTED_ORDER = ("control", "candidate", "candidate", "control")
EXPECTED_INSTRUMENTATION = {
    "normal": "none",
    "allocator": "system_allocator_operation_scoped",
    "owned": "none",
    "owned-allocator": "system_allocator_operation_scoped",
}
EXPECTED_BINARY = {
    "normal": "litchi-perf-baseline",
    "allocator": "litchi-perf-baseline-alloc",
    "owned": "litchi-perf-baseline",
    "owned-allocator": "litchi-perf-baseline-alloc",
}
# Independent corpus/output oracles.  These are intentionally checked per
# report row, before repeat-parity projections, so a capture that is
# consistently bound to the wrong generated fixture cannot pass by agreeing
# with itself.
EXPECTED_ARCHIVE_SHA256 = "6a57231ba681bc7bdd38d447ebd5348ef3b1fefedeefb1e61c97f22faa074e53"
EXPECTED_WORKBOOK_SHA256 = "c78e03ba3743132e04b08bf6f4579ceb1c112a22c441c1e036381d3e06c6d041"
EXPECTED_ARCHIVE_BYTES = 16995840
EXPECTED_WORKBOOK_BYTES = 80946
EXPECTED_OUTPUT_SHA256 = {
    "xls_semantic_open": EXPECTED_ARCHIVE_SHA256,
    "xls_source_backed_open": EXPECTED_ARCHIVE_SHA256,
    "xls_eager_open_list_worksheets": "e9f1a47d927dddab1c37d497365644e5533a2dbd97d662a471f481d7d384c505",
    "xls_source_backed_open_list_worksheets": "e9f1a47d927dddab1c37d497365644e5533a2dbd97d662a471f481d7d384c505",
    "xls_eager_open_one_cell": "e726a50d216e6d71d7c53aabd23ab5e0d4677c3ef1f41fc35410143ebe6381c1",
    "xls_source_backed_open_one_cell": "e726a50d216e6d71d7c53aabd23ab5e0d4677c3ef1f41fc35410143ebe6381c1",
    "xls_owned_source_open": EXPECTED_ARCHIVE_SHA256,
    "xls_owned_source_open_list_worksheets": "e9f1a47d927dddab1c37d497365644e5533a2dbd97d662a471f481d7d384c505",
    "xls_owned_source_open_one_cell": "e726a50d216e6d71d7c53aabd23ab5e0d4677c3ef1f41fc35410143ebe6381c1",
}
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
)
ALLOCATION_COUNT_FIELDS = frozenset(
    {
        "allocation_calls",
        "deallocation_calls",
        "reallocation_calls",
        "failed_allocation_calls",
    }
)


class ValidationError(Exception):
    """An input did not satisfy the capture contract."""


class _Missing:
    pass


MISSING = _Missing()
OMIT = _Missing()


def _fail(path: str, message: str) -> None:
    raise ValidationError(f"{path}: {message}")


def _dict(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        _fail(path, "expected an object")
    return value


def _list(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        _fail(path, "expected an array")
    return value


def _string(value: Any, path: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        _fail(path, "expected a non-empty string")
    return value


def _integer(value: Any, path: str, *, minimum: int = 0) -> int:
    if type(value) is not int or value < minimum:
        _fail(path, f"expected an integer >= {minimum}")
    return value


def _number(value: Any, path: str) -> float | int:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        _fail(path, "expected a number")
    if isinstance(value, float) and not math.isfinite(value):
        _fail(path, "must be finite")
    return value


def _digest(value: Any, path: str) -> str:
    text = _string(value, path)
    if not re.fullmatch(r"[0-9a-fA-F]{64}", text):
        _fail(path, "expected a 64-digit hexadecimal SHA-256 digest")
    return text.lower()


def _canonical(value: Any) -> bytes:
    return json.dumps(
        value, sort_keys=True, separators=(",", ":"), ensure_ascii=False
    ).encode("utf-8")


def _sha256(value: Any) -> str:
    return hashlib.sha256(_canonical(value)).hexdigest()


def _read_json(path: Path, *, required: bool = True) -> dict[str, Any] | list[Any] | None:
    """Read plain JSON, falling back to a sibling ``.zst`` file."""

    candidates: list[Path] = []
    if path.exists():
        candidates.append(path)
    if path.suffix == ".zst":
        if path.exists():
            candidates.append(path)
    else:
        compressed = Path(str(path) + ".zst")
        if compressed.exists():
            candidates.append(compressed)
    # A caller occasionally supplies ``foo.json.zst`` while the plain file is
    # present.  The first existing candidate is still the requested artifact.
    if not candidates:
        if required:
            raise ValidationError(f"missing capture artifact: {path} (or {path}.zst)")
        return None
    chosen = candidates[0]
    try:
        if chosen.suffix == ".zst":
            raw = subprocess.check_output(["zstd", "-q", "-dc", str(chosen)])
        else:
            raw = chosen.read_bytes()
        value = json.loads(raw.decode("utf-8"))
    except FileNotFoundError as error:
        raise ValidationError(f"cannot read {chosen}: zstd is unavailable") from error
    except (OSError, subprocess.CalledProcessError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read JSON artifact {chosen}: {error}") from error
    return value


def _read_json_with_meta(path: Path, *, required: bool = True) -> dict[str, Any] | None:
    """Read JSON and retain the artifact path and raw-file digest."""

    candidates: list[Path] = []
    if path.exists():
        candidates.append(path)
    if path.suffix != ".zst" and Path(str(path) + ".zst").exists():
        candidates.append(Path(str(path) + ".zst"))
    if not candidates:
        if required:
            raise ValidationError(f"missing capture artifact: {path} (or {path}.zst)")
        return None
    chosen = candidates[0]
    try:
        raw = (
            subprocess.check_output(["zstd", "-q", "-dc", str(chosen)])
            if chosen.suffix == ".zst"
            else chosen.read_bytes()
        )
        value = json.loads(raw.decode("utf-8"))
    except FileNotFoundError as error:
        raise ValidationError(f"cannot read {chosen}: zstd is unavailable") from error
    except (OSError, subprocess.CalledProcessError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read JSON artifact {chosen}: {error}") from error
    return {
        "value": value,
        "path": chosen,
        "raw_sha256": hashlib.sha256(raw).hexdigest(),
        "compressed": chosen.suffix == ".zst",
    }


def _option(argv: list[Any], name: str, path: str) -> str:
    values = [str(value) for value in argv]
    for index, value in enumerate(values):
        if value == name:
            if index + 1 >= len(values):
                _fail(path, f"{name} has no value")
            return values[index + 1]
        if value.startswith(name + "="):
            return value[len(name) + 1 :]
    _fail(path, f"missing {name}")


def _parse_time(value: Any, path: str) -> _datetime.datetime:
    text = _string(value, path)
    try:
        return _datetime.datetime.fromisoformat(text.replace("Z", "+00:00"))
    except ValueError as error:
        raise ValidationError(f"{path}: invalid ISO-8601 timestamp {text!r}") from error


def _midpoint(left: int, right: int) -> int:
    # Match the Rust harness' overflow-safe integer midpoint.
    return left // 2 + right // 2 + ((left % 2 + right % 2) // 2)


def _nearest_rank(values: list[int], percentile: int) -> int:
    index = (percentile * len(values) + 99) // 100 - 1
    return values[min(index, len(values) - 1)]


def _summary(values: list[int]) -> dict[str, Any]:
    if not values:
        raise ValidationError("statistics: empty sample vector")
    ordered = sorted(values)
    n = len(ordered)
    return {
        "samples": n,
        "min": ordered[0],
        "p50": _midpoint(ordered[(n - 1) // 2], ordered[n // 2]),
        "p95": _nearest_rank(ordered, 95),
        "p99": _nearest_rank(ordered, 99),
        "max": ordered[-1],
        "mean": statistics.fmean(ordered),
    }


def _assert_close(actual: Any, expected: float | int, path: str) -> None:
    value = _number(actual, path)
    if not math.isclose(float(value), float(expected), rel_tol=2e-9, abs_tol=1e-6):
        _fail(path, f"reported value {value!r} differs from recomputed {expected!r}")


def _validate_elapsed(value: Any, path: str, expected_count: int) -> dict[str, Any]:
    elapsed = _dict(value, path)
    if elapsed.get("unit") != "ns":
        _fail(f"{path}.unit", "must be 'ns'")
    samples = _list(elapsed.get("samples", MISSING), f"{path}.samples")
    if len(samples) != expected_count:
        _fail(f"{path}.samples", f"expected {expected_count} samples, got {len(samples)}")
    numeric_samples = [
        _integer(sample, f"{path}.samples[{index}]", minimum=1)
        for index, sample in enumerate(samples)
    ]
    order = _list(elapsed.get("sample_order", MISSING), f"{path}.sample_order")
    if len(order) != expected_count:
        _fail(
            f"{path}.sample_order",
            f"expected {expected_count} entries, got {len(order)}",
        )
    numeric_order = [
        _integer(item, f"{path}.sample_order[{index}]")
        for index, item in enumerate(order)
    ]
    if sorted(numeric_order) != list(range(expected_count)):
        _fail(f"{path}.sample_order", "must be an exact permutation of retained samples")
    if list(zip(numeric_samples, numeric_order)) != sorted(
        zip(numeric_samples, numeric_order)
    ):
        _fail(
            f"{path}.sample_order",
            "must describe the elapsed-then-original-index ordering emitted by the harness",
        )
    expected = _summary(numeric_samples)
    for field in ("min", "p50", "p95", "p99", "max"):
        if field not in elapsed:
            _fail(path, f"missing elapsed statistic {field!r}")
        if elapsed[field] != expected[field]:
            _fail(
                f"{path}.{field}",
                f"reported {elapsed[field]!r}, recomputed {expected[field]!r}",
            )
    if "mean" not in elapsed:
        _fail(path, "missing elapsed statistic 'mean'")
    _assert_close(elapsed["mean"], expected["mean"], f"{path}.mean")
    if "standard_deviation" in elapsed:
        if len(numeric_samples) > 1:
            _assert_close(
                elapsed["standard_deviation"],
                statistics.stdev(numeric_samples),
                f"{path}.standard_deviation",
            )
        elif elapsed["standard_deviation"] != 0:
            _fail(f"{path}.standard_deviation", "one-sample deviation must be zero")
    interval = elapsed.get("confidence_interval_95", MISSING)
    if interval is not MISSING:
        interval = _dict(interval, f"{path}.confidence_interval_95")
        _string(interval.get("method", MISSING), f"{path}.confidence_interval_95.method")
        lower = _number(interval.get("lower", MISSING), f"{path}.confidence_interval_95.lower")
        upper = _number(interval.get("upper", MISSING), f"{path}.confidence_interval_95.upper")
        if lower > upper:
            _fail(f"{path}.confidence_interval_95", "lower exceeds upper")
        if lower < 0:
            _fail(f"{path}.confidence_interval_95.lower", "must be non-negative")
    return {
        "samples": numeric_samples,
        "sample_order": numeric_order,
        "summary": expected,
        "sample_order_sha256": _sha256(numeric_order),
    }


def _is_metric_vector(value: Any) -> bool:
    if not isinstance(value, dict):
        return False
    keys = set(value)
    return {"status", "scope"}.issubset(keys) and keys.issubset(
        {"status", "scope", "values"}
    )


def _validate_metric_tree(value: Any, path: str, expected_count: int) -> None:
    if _is_metric_vector(value):
        status = _string(value.get("status", MISSING), f"{path}.status")
        if status not in {"measured", "not_applicable", "unavailable", "overflow"}:
            _fail(f"{path}.status", f"unknown metric status {status!r}")
        _string(value.get("scope", MISSING), f"{path}.scope")
        if "values" not in value:
            if status == "measured":
                _fail(path, "measured metric vector omitted values")
            return
        values = value["values"]
        if values is None:
            if status == "measured":
                _fail(f"{path}.values", "measured metric vector cannot be null")
            return
        values = _list(values, f"{path}.values")
        if len(values) != expected_count:
            _fail(
                f"{path}.values",
                f"expected {expected_count} values, got {len(values)}",
            )
        if status != "measured":
            _fail(f"{path}.values", "non-measured metric must omit values")
        for index, item in enumerate(values):
            if isinstance(item, bool):
                continue
            if isinstance(item, (int, float, str)):
                if isinstance(item, float) and not math.isfinite(item):
                    _fail(f"{path}.values[{index}]", "must be finite")
                continue
            _fail(f"{path}.values[{index}]", "contains an unsupported scalar")
        return
    if isinstance(value, dict):
        for key, child in value.items():
            _validate_metric_tree(child, f"{path}.{key}", expected_count)
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _validate_metric_tree(child, f"{path}[{index}]", expected_count)


def _validate_operation_metrics(
    value: Any,
    path: str,
    expected_count: int,
    elapsed_order: list[int],
    *,
    allocator: bool,
) -> dict[str, Any]:
    metrics = _dict(value, path)
    if _integer(metrics.get("sample_count", MISSING), f"{path}.sample_count") != expected_count:
        _fail(f"{path}.sample_count", "does not match elapsed sample count")
    indices = _list(metrics.get("sample_indices", MISSING), f"{path}.sample_indices")
    normalized_indices = [
        _integer(item, f"{path}.sample_indices[{index}]")
        for index, item in enumerate(indices)
    ]
    if normalized_indices != elapsed_order:
        _fail(
            f"{path}.sample_indices",
            "does not equal elapsed_ns.sample_order",
        )
    alignment = _string(metrics.get("alignment", MISSING), f"{path}.alignment")
    if alignment not in {
        "elapsed_ns.samples_by_elapsed_then_sample_index",
        "elapsed_ns.samples",
    }:
        _fail(f"{path}.alignment", f"unknown alignment {alignment!r}")
    _string(metrics.get("latency_claim", MISSING), f"{path}.latency_claim")
    _validate_metric_tree(metrics, path, expected_count)
    allocation = _dict(metrics.get("allocation", MISSING), f"{path}.allocation")
    allocation_status = _string(
        allocation.get("status", MISSING), f"{path}.allocation.status"
    )
    if allocator:
        if allocation_status != "measured":
            _fail(f"{path}.allocation.status", "allocator capture must be measured")
        _string(allocation.get("scope", MISSING), f"{path}.allocation.scope")
        allocation_values: dict[str, list[int]] = {}
        for field in ALLOCATION_FIELDS:
            vector = _dict(
                allocation.get(field, MISSING), f"{path}.allocation.{field}"
            )
            if vector.get("status") != "measured":
                _fail(
                    f"{path}.allocation.{field}.status",
                    "allocator capture must measure every allocation vector",
                )
            values = _list(
                vector.get("values", MISSING), f"{path}.allocation.{field}.values"
            )
            if len(values) != expected_count:
                _fail(
                    f"{path}.allocation.{field}.values",
                    f"expected {expected_count} values, got {len(values)}",
                )
            allocation_values[field] = [
                _integer(item, f"{path}.allocation.{field}.values[{index}]")
                for index, item in enumerate(values)
            ]
        if any(allocation_values["failed_allocation_calls"]):
            _fail(
                f"{path}.allocation.failed_allocation_calls",
                "failed allocation calls must remain zero",
            )
        return {"allocation": allocation_values}
    if allocation_status == "measured":
        _fail(f"{path}.allocation.status", "normal capture unexpectedly measures allocations")
    for field in ALLOCATION_FIELDS:
        vector = allocation.get(field, MISSING)
        if vector is MISSING:
            _fail(f"{path}.allocation.{field}", "missing normal allocation status vector")
        if _is_metric_vector(vector) and "values" in vector:
            _fail(f"{path}.allocation.{field}", "normal capture must omit values")
    return {"allocation": None}


def _validate_scalar_sample_vector(value: Any, path: str, expected_count: int) -> list[Any]:
    values = _list(value, path)
    if len(values) != expected_count:
        _fail(path, f"expected {expected_count} values, got {len(values)}")
    return values


def _validate_source(
    value: Any,
    path: str,
    case: str,
    expected_count: int,
    corpus: dict[str, Any],
) -> dict[str, Any] | None:
    if case not in SOURCE_CASES:
        if value is not None:
            _fail(path, "non-source XLS cases must emit source=null")
        return None
    source = _dict(value, path)
    for field in (
        "read_calls",
        "read_bytes",
        "ordinary_payload_read_calls",
        "ordinary_payload_read_bytes",
        "max_in_flight_reads",
    ):
        vector = _validate_scalar_sample_vector(
            source.get(field, MISSING), f"{path}.{field}", expected_count
        )
        for index, item in enumerate(vector):
            _integer(item, f"{path}.{field}[{index}]")
    xls = _dict(source.get("xls", MISSING), f"{path}.xls")
    if xls.get("implementation") != "source-backed":
        _fail(f"{path}.xls.implementation", "must be source-backed")
    expected_operation = {
        "xls_source_backed_open": "open",
        "xls_source_backed_open_list_worksheets": "open+list",
        "xls_source_backed_open_one_cell": "open+one-cell",
    }[case]
    if xls.get("operation") != expected_operation:
        _fail(f"{path}.xls.operation", f"must be {expected_operation!r}")
    if xls.get("timing_scope") != expected_operation:
        _fail(f"{path}.xls.timing_scope", f"must be {expected_operation!r}")
    _string(xls.get("source_counter_scope", MISSING), f"{path}.xls.source_counter_scope")
    _string(xls.get("materialization_scope", MISSING), f"{path}.xls.materialization_scope")
    if _digest(xls.get("archive_sha256", MISSING), f"{path}.xls.archive_sha256") != _digest(
        corpus.get("archive_sha256", MISSING), "corpus.archive_sha256"
    ):
        _fail(f"{path}.xls.archive_sha256", "does not match result.corpus.archive_sha256")
    if _digest(
        xls.get("workbook_stream_sha256", MISSING), f"{path}.xls.workbook_stream_sha256"
    ) != _digest(
        corpus.get("target_payload_sha256", MISSING), "corpus.target_payload_sha256"
    ):
        _fail(
            f"{path}.xls.workbook_stream_sha256",
            "does not match result.corpus.target_payload_sha256",
        )
    scalar_xls = {
        "implementation",
        "operation",
        "timing_scope",
        "source_counter_scope",
        "materialization_scope",
        "archive_sha256",
        "workbook_stream_sha256",
    }
    for field, field_value in xls.items():
        if field in scalar_xls:
            continue
        if not isinstance(field_value, list):
            _fail(f"{path}.xls.{field}", "expected a per-sample vector")
        vector = _validate_scalar_sample_vector(
            field_value, f"{path}.xls.{field}", expected_count
        )
        for index, item in enumerate(vector):
            if isinstance(item, bool):
                continue
            _integer(item, f"{path}.xls.{field}[{index}]")
    for field in (
        "source_retained_bytes",
        "source_version_checks",
        "cfb_structural_read_bytes",
        "workbook_global_read_bytes",
    ):
        values = xls.get(field, MISSING)
        if values is not MISSING and any(
            item == 0 for item in _validate_scalar_sample_vector(values, f"{path}.xls.{field}", expected_count)
        ):
            _fail(f"{path}.xls.{field}", "must be positive for source-backed reads")
    for field in (
        "opaque_payload_read_bytes",
        "unselected_worksheet_read_bytes",
        "complete_archive_materialized_bytes",
    ):
        values = _validate_scalar_sample_vector(
            xls.get(field, MISSING), f"{path}.xls.{field}", expected_count
        )
        if any(item != 0 for item in values):
            _fail(f"{path}.xls.{field}", "must remain zero for this locality probe")
    stable = _validate_scalar_sample_vector(
        xls.get("source_version_stability_verified", MISSING),
        f"{path}.xls.source_version_stability_verified",
        expected_count,
    )
    if stable != [True] * expected_count:
        _fail(f"{path}.xls.source_version_stability_verified", "source version was not stable")
    parsed_sheets = _validate_scalar_sample_vector(
        xls.get("parsed_sheet_counts", MISSING), f"{path}.xls.parsed_sheet_counts", expected_count
    )
    if any(item <= 0 for item in parsed_sheets):
        _fail(f"{path}.xls.parsed_sheet_counts", "must contain positive sheet counts")
    parsed_cells = _validate_scalar_sample_vector(
        xls.get("parsed_cell_counts", MISSING), f"{path}.xls.parsed_cell_counts", expected_count
    )
    expected_cells = 1 if case.endswith("one_cell") else 0
    if parsed_cells != [expected_cells] * expected_count:
        _fail(
            f"{path}.xls.parsed_cell_counts",
            f"must contain the operation oracle value {expected_cells}",
        )
    selected_bytes = _validate_scalar_sample_vector(
        xls.get("selected_worksheet_read_bytes", MISSING),
        f"{path}.xls.selected_worksheet_read_bytes",
        expected_count,
    )
    open_zero = _validate_scalar_sample_vector(
        xls.get("open_reads_zero_worksheet_payload", MISSING),
        f"{path}.xls.open_reads_zero_worksheet_payload",
        expected_count,
    )
    selected_only = _validate_scalar_sample_vector(
        xls.get("selected_query_reads_only_selected_worksheet", MISSING),
        f"{path}.xls.selected_query_reads_only_selected_worksheet",
        expected_count,
    )
    if case.endswith("one_cell"):
        if any(item <= 0 for item in selected_bytes) or selected_only != [True] * expected_count:
            _fail(f"{path}.xls", "one-cell locality must read selected worksheet only")
        if open_zero != [False] * expected_count:
            _fail(f"{path}.xls.open_reads_zero_worksheet_payload", "one-cell query is not an open-only probe")
    else:
        if any(item != 0 for item in selected_bytes) or open_zero != [True] * expected_count:
            _fail(f"{path}.xls", "open/list locality must avoid worksheet payload reads")
        if selected_only != [False] * expected_count:
            _fail(f"{path}.xls.selected_query_reads_only_selected_worksheet", "open/list must not claim a selected query")
    root_bytes = source["read_bytes"]
    root_calls = source["read_calls"]
    if any(item <= 0 for item in root_bytes) or any(item <= 0 for item in root_calls):
        _fail(path, "source logical reads must be positive")
    return {
        "source_counter_scope": xls["source_counter_scope"],
        "implementation": xls["implementation"],
        "operation": xls["operation"],
    }


def _validate_corpus(value: Any, path: str) -> dict[str, Any]:
    corpus = _dict(value, path)
    _digest(corpus.get("archive_sha256", MISSING), f"{path}.archive_sha256")
    _digest(corpus.get("target_payload_sha256", MISSING), f"{path}.target_payload_sha256")
    for field in ("archive_bytes", "target_payload_bytes"):
        if field in corpus:
            _integer(corpus[field], f"{path}.{field}", minimum=1)
    return corpus


def _validate_parallel_metrics(
    value: Any, path: str, cases: tuple[str, ...], corpus_sha256: str
) -> None:
    metrics = _dict(value, path)
    if "schema_version" in metrics and metrics["schema_version"] != 1:
        _fail(f"{path}.schema_version", "unsupported parallel-metrics schema")
    listed = _list(metrics.get("cases", MISSING), f"{path}.cases")
    if len(listed) != len(cases):
        _fail(f"{path}.cases", "does not match result count")
    for index, (case, item) in enumerate(zip(cases, listed)):
        obj = _dict(item, f"{path}.cases[{index}]")
        if obj.get("case") != case:
            _fail(f"{path}.cases[{index}].case", f"expected {case!r}")
        if "corpus_sha256" in obj and obj["corpus_sha256"] != corpus_sha256:
            _fail(f"{path}.cases[{index}].corpus_sha256", "does not match result corpus")


def _validate_report_header(
    report: Any,
    path: str,
    family: str,
    expected_cases: tuple[str, ...],
    expected_count: int,
    expected_warmup: int,
) -> dict[str, Any]:
    obj = _dict(report, path)
    if obj.get("schema_version") != 1:
        _fail(f"{path}.schema_version", "expected report schema version 1")
    tool = _dict(obj.get("tool", MISSING), f"{path}.tool")
    if tool.get("name") != "litchi-perf-baseline":
        _fail(f"{path}.tool.name", "unexpected benchmark tool")
    if tool.get("binary") != EXPECTED_BINARY[family]:
        _fail(f"{path}.tool.binary", "unexpected executable")
    if tool.get("instrumentation") != EXPECTED_INSTRUMENTATION[family]:
        _fail(f"{path}.tool.instrumentation", "does not match capture family")
    _string(tool.get("version", MISSING), f"{path}.tool.version")
    if tool.get("profile") != "release":
        _fail(f"{path}.tool.profile", "capture must use the release profile")
    binary = _dict(obj.get("binary_identity", MISSING), f"{path}.binary_identity")
    binary_sha = _digest(binary.get("binary_sha256", MISSING), f"{path}.binary_identity.binary_sha256")
    binary_path = _string(binary.get("path", MISSING), f"{path}.binary_identity.path")
    if Path(binary_path).name != EXPECTED_BINARY[family]:
        _fail(f"{path}.binary_identity.path", "basename does not match capture family")
    if binary.get("executable") is not True:
        _fail(f"{path}.binary_identity.executable", "benchmark binary is not marked executable")
    if binary.get("profile") != "release":
        _fail(f"{path}.binary_identity.profile", "binary identity is not release")
    _integer(binary.get("binary_bytes", MISSING), f"{path}.binary_identity.binary_bytes", minimum=1)
    environment = _dict(obj.get("environment", MISSING), f"{path}.environment")
    revision = _string(environment.get("git_revision", MISSING), f"{path}.environment.git_revision")
    if environment.get("git_worktree_dirty") is not False:
        _fail(f"{path}.environment.git_worktree_dirty", "capture source must be clean")
    rustc = _string(environment.get("rustc_version", MISSING), f"{path}.environment.rustc_version")
    if not rustc.startswith("rustc 1.98.1 "):
        _fail(f"{path}.environment.rustc_version", "capture must use Rust 1.98.1")
    if environment.get("rustflags") != FLAGS:
        _fail(f"{path}.environment.rustflags", "does not match the frozen profiling flags")
    if environment.get("cpu_affinity") != "2":
        _fail(f"{path}.environment.cpu_affinity", "capture must be pinned to CPU 2")
    configuration = _dict(obj.get("configuration", MISSING), f"{path}.configuration")
    if configuration.get("cases") != list(expected_cases):
        _fail(f"{path}.configuration.cases", "scenario set/order differs from protocol")
    if configuration.get("samples_per_case") != expected_count:
        _fail(f"{path}.configuration.samples_per_case", "does not match protocol")
    if configuration.get("warmup_iterations_per_case") != expected_warmup:
        _fail(f"{path}.configuration.warmup_iterations_per_case", "does not match protocol")
    if "filesystem_cache_states" in configuration and configuration["filesystem_cache_states"] != ["warm"]:
        _fail(f"{path}.configuration.filesystem_cache_states", "must be the warm capture")
    if "execution_workers" in configuration and configuration["execution_workers"] != [1]:
        _fail(f"{path}.configuration.execution_workers", "must use one worker")
    results = _list(obj.get("results", MISSING), f"{path}.results")
    if len(results) != len(expected_cases):
        _fail(f"{path}.results", "does not match scenario count")
    corpus_values: list[dict[str, Any]] = []
    case_details: dict[str, Any] = {}
    for index, (case, result_value) in enumerate(zip(expected_cases, results)):
        result_path = f"{path}.results[{index}]"
        result = _dict(result_value, result_path)
        if result.get("case") != case:
            _fail(f"{result_path}.case", f"expected {case!r}")
        corpus = _validate_corpus(result.get("corpus", MISSING), f"{result_path}.corpus")
        corpus_values.append(corpus)
        output = _digest(result.get("output_sha256", MISSING), f"{result_path}.output_sha256")
        archive_sha = _digest(corpus["archive_sha256"], f"{result_path}.corpus.archive_sha256")
        if archive_sha != EXPECTED_ARCHIVE_SHA256:
            _fail(
                f"{result_path}.corpus.archive_sha256",
                "does not match the independent generated-archive oracle",
            )
        workbook_sha = _digest(
            corpus["target_payload_sha256"],
            f"{result_path}.corpus.target_payload_sha256",
        )
        if workbook_sha != EXPECTED_WORKBOOK_SHA256:
            _fail(
                f"{result_path}.corpus.target_payload_sha256",
                "does not match the independent Workbook-stream oracle",
            )
        if corpus.get("archive_bytes") != EXPECTED_ARCHIVE_BYTES:
            _fail(
                f"{result_path}.corpus.archive_bytes",
                f"does not match the independent archive-size oracle ({EXPECTED_ARCHIVE_BYTES})",
            )
        if corpus.get("target_payload_bytes") != EXPECTED_WORKBOOK_BYTES:
            _fail(
                f"{result_path}.corpus.target_payload_bytes",
                f"does not match the independent Workbook-size oracle ({EXPECTED_WORKBOOK_BYTES})",
            )
        expected_output = EXPECTED_OUTPUT_SHA256[case]
        if output != expected_output:
            _fail(
                f"{result_path}.output_sha256",
                "does not match the independent operation-output oracle",
            )
        elapsed = _validate_elapsed(result.get("elapsed_ns", MISSING), f"{result_path}.elapsed_ns", expected_count)
        operation = _validate_operation_metrics(
            result.get("operation_metrics", MISSING),
            f"{result_path}.operation_metrics",
            expected_count,
            elapsed["sample_order"],
            allocator=family.endswith("allocator"),
        )
        if _PERF_COMPARE is not None:
            try:
                _PERF_COMPARE._validate_operation_metrics(
                    result["operation_metrics"],
                    f"{result_path}.operation_metrics",
                    elapsed["samples"],
                    obj["schema_version"],
                    elapsed_sample_order=elapsed["sample_order"],
                )
            except Exception as error:
                _fail(
                    f"{result_path}.operation_metrics",
                    f"repository perf_compare validation failed: {error}",
                )
        source = _validate_source(
            result.get("source"),
            f"{result_path}.source",
            case,
            expected_count,
            corpus,
        )
        case_details[case] = {
            "result": result,
            "corpus": corpus,
            "output_sha256": output,
            "elapsed": elapsed,
            "operation": operation,
            "source": source,
        }
    if len({_canonical(item) for item in corpus_values}) != 1:
        _fail(path, "scenario rows do not share one corpus manifest")
    if "parallel_metrics" not in obj:
        _fail(path, "missing parallel_metrics envelope")
    _validate_parallel_metrics(
        obj["parallel_metrics"],
        f"{path}.parallel_metrics",
        expected_cases,
        _digest(corpus_values[0]["archive_sha256"], "corpus.archive_sha256"),
    )
    return {
        "report": obj,
        "revision": revision,
        "binary_sha256": binary_sha,
        "binary": binary,
        "tool": tool,
        "environment": environment,
        "configuration": configuration,
        "cases": case_details,
    }


def _walk_paths(value: Any, path: tuple[str, ...] = ()) -> Iterable[tuple[tuple[str, ...], Any]]:
    if isinstance(value, dict):
        for key, child in value.items():
            child_path = path + (str(key),)
            yield child_path, child
            yield from _walk_paths(child, child_path)
    elif isinstance(value, list):
        for index, child in enumerate(value):
            child_path = path + (str(index),)
            yield child_path, child
            yield from _walk_paths(child, child_path)


def _scope_values(result: dict[str, Any]) -> list[tuple[tuple[str, ...], str]]:
    values: list[tuple[tuple[str, ...], str]] = []
    for path, value in _walk_paths(result):
        if path and path[-1] == "source_counter_scope":
            values.append((path, _string(value, "result." + ".".join(path))))
    return values


def _project_operation(value: Any, path: tuple[str, ...] = ()) -> Any:
    """Keep operation shape/status while omitting child-sensitive vectors."""

    if _is_metric_vector(value):
        output = {
            "status": value["status"],
            "scope": value["scope"],
        }
        if "values" in value:
            output["values"] = "null" if value["values"] is None else "present"
        return output
    if isinstance(value, dict):
        output: dict[str, Any] = {}
        for key, child in value.items():
            if key == "sample_indices":
                output[key] = "<elapsed_sample_order>"
            else:
                output[key] = _project_operation(child, path + (str(key),))
        return output
    if isinstance(value, list):
        return [_project_operation(child, path) for child in value]
    return value


def _project_result(value: Any, path: tuple[str, ...] = ()) -> Any:
    if isinstance(value, dict):
        output: dict[str, Any] = {}
        for key, child in value.items():
            child_path = path + (str(key),)
            if key == "elapsed_ns":
                # Elapsed vectors/statistics are intentionally descriptive and
                # differ between fresh children.
                continue
            if child_path == ALLOWED_SCOPE_PATH:
                output[key] = "<allowed_xls_source_counter_scope>"
            elif key == "operation_metrics":
                output[key] = _project_operation(child, child_path)
            else:
                output[key] = _project_result(child, child_path)
        return output
    if isinstance(value, list):
        return [_project_result(child, path) for child in value]
    return value


def _project_report(value: dict[str, Any]) -> dict[str, Any]:
    output: dict[str, Any] = {}
    for key, child in value.items():
        if key in {"binary_identity", "corpus_catalog"}:
            continue
        if key == "environment" and isinstance(child, dict):
            output[key] = {k: v for k, v in child.items() if k != "git_revision"}
        elif key == "results":
            output[key] = [_project_result(item) for item in child]
        else:
            output[key] = child
    return output


def _differences(left: Any, right: Any, path: tuple[str, ...] = (), limit: int = 20) -> list[str]:
    if limit <= 0:
        return []
    if type(left) is not type(right):
        return [".".join(path) or "<root>"]
    if isinstance(left, dict):
        differences: list[str] = []
        keys = sorted(set(left) | set(right))
        for key in keys:
            if key not in left or key not in right:
                differences.append(".".join(path + (str(key),)))
            else:
                differences.extend(_differences(left[key], right[key], path + (str(key),), limit - len(differences)))
            if len(differences) >= limit:
                return differences[:limit]
        return differences
    if isinstance(left, list):
        if len(left) != len(right):
            return [".".join(path) + f"[length {len(left)} != {len(right)}]"]
        differences = []
        for index, (left_item, right_item) in enumerate(zip(left, right)):
            differences.extend(_differences(left_item, right_item, path + (str(index),), limit - len(differences)))
            if len(differences) >= limit:
                return differences[:limit]
        return differences
    if left != right:
        return [".".join(path) or "<root>"]
    return []


def _catalog_projection(catalog: dict[str, Any], path: str) -> dict[str, Any]:
    """Return catalog content identity while ignoring build/catalog metadata."""

    manifest = catalog.get("manifest_version")
    if manifest != 2:
        _fail(f"{path}.manifest_version", "expected corpus catalog manifest version 2")
    catalog_id = _string(catalog.get("catalog_id", MISSING), f"{path}.catalog_id")
    content_set = _digest(catalog.get("content_set_sha256", MISSING), f"{path}.content_set_sha256")
    digest_pairs: list[tuple[str, str]] = []

    def visit(value: Any, location: tuple[str, ...]) -> None:
        if isinstance(value, dict):
            for key, child in value.items():
                lower = str(key).lower()
                if lower == "catalog_sha256":
                    continue
                child_location = location + (str(key),)
                if isinstance(child, str) and (
                    lower.endswith("sha256") or "digest" in lower
                ):
                    digest_pairs.append((".".join(child_location), _digest(child, f"{path}." + ".".join(child_location))))
                else:
                    visit(child, child_location)
        elif isinstance(value, list):
            for index, child in enumerate(value):
                visit(child, location + (str(index),))

    visit(catalog, ())
    if not digest_pairs:
        _fail(path, "catalog has no corpus digest fields")
    bindings: list[dict[str, Any]] = []
    raw_bindings = catalog.get("case_bindings", [])
    if not isinstance(raw_bindings, list):
        _fail(f"{path}.case_bindings", "expected an array")
    for index, item in enumerate(raw_bindings):
        binding = _dict(item, f"{path}.case_bindings[{index}]")
        case = _string(binding.get("case", MISSING), f"{path}.case_bindings[{index}].case")
        corpus_id = _string(binding.get("corpus_id", MISSING), f"{path}.case_bindings[{index}].corpus_id")
        archive = binding.get("legacy_archive_sha256")
        if archive is not None:
            archive = _digest(archive, f"{path}.case_bindings[{index}].legacy_archive_sha256")
        bindings.append({"case": case, "corpus_id": corpus_id, "legacy_archive_sha256": archive})
    return {
        "manifest_version": manifest,
        "catalog_id": catalog_id,
        "content_set_sha256": content_set,
        "digest_fields": sorted(digest_pairs),
        "case_bindings": sorted(bindings, key=lambda item: item["case"]),
    }


def _validate_catalog_binding(
    catalog: dict[str, Any],
    path: str,
    cases: tuple[str, ...],
    corpus: dict[str, Any],
) -> dict[str, Any]:
    projection = _catalog_projection(catalog, path)
    expected_archive = _digest(corpus["archive_sha256"], "corpus.archive_sha256")
    expected_payload = _digest(corpus["target_payload_sha256"], "corpus.target_payload_sha256")
    if not any(
        value == expected_archive
        for location, value in projection["digest_fields"]
        if "archive_sha256" in location.lower()
    ):
        _fail(path, "catalog does not carry the result archive digest")
    if not any(
        value == expected_payload
        for location, value in projection["digest_fields"]
        if any(token in location.lower() for token in ("target_payload_sha256", "workbook_sha256"))
    ):
        _fail(path, "catalog does not carry the result workbook/target digest")
    binding_by_case = {item["case"]: item for item in projection["case_bindings"]}
    for case in cases:
        binding = binding_by_case.get(case)
        if binding is None:
            _fail(f"{path}.case_bindings", f"missing binding for {case!r}")
        if binding["legacy_archive_sha256"] not in {None, expected_archive}:
            _fail(
                f"{path}.case_bindings.{case}.legacy_archive_sha256",
                "does not match the result archive digest",
            )
    return projection


def _validate_protocol(protocol: Any) -> dict[str, Any]:
    obj = _dict(protocol, "protocol")
    if obj.get("change") != "0412":
        _fail("protocol.change", "expected 0412")
    revisions: dict[str, str] = {}
    for field in ("control_revision", "candidate_revision"):
        revision = _string(obj.get(field, MISSING), f"protocol.{field}")
        if not re.fullmatch(r"[0-9a-fA-F]{40}", revision):
            _fail(f"protocol.{field}", "expected a 40-digit hexadecimal revision")
        revisions[field] = revision.lower()
    if revisions["control_revision"] == revisions["candidate_revision"]:
        _fail("protocol", "control_revision and candidate_revision must differ")
    if obj.get("cases") != list(CASES):
        _fail("protocol.cases", "does not contain the six XLS scenarios in order")
    for family in ("normal", "allocator"):
        section = _dict(obj.get(family, MISSING), f"protocol.{family}")
        if section.get("order") != list(EXPECTED_ORDER):
            _fail(f"protocol.{family}.order", "must be control,candidate,candidate,control")
        _integer(section.get("samples", MISSING), f"protocol.{family}.samples", minimum=1)
        _integer(section.get("warmup", MISSING), f"protocol.{family}.warmup")
    return obj


def _family_specs(protocol: dict[str, Any], root: Path) -> list[dict[str, Any]]:
    specs = [
        {
            "name": "normal",
            "cases": CASES,
            "count": 4,
            "roles": EXPECTED_ORDER,
            "samples": protocol["normal"]["samples"],
            "warmup": protocol["normal"]["warmup"],
        },
        {
            "name": "allocator",
            "cases": CASES,
            "count": 4,
            "roles": EXPECTED_ORDER,
            "samples": protocol["allocator"]["samples"],
            "warmup": protocol["allocator"]["warmup"],
        },
    ]
    owned_section = protocol.get("owned_source")
    if isinstance(owned_section, dict):
        owned_count = owned_section.get("repetitions", 4)
        owned_allocator_count = owned_section.get("allocation_repetitions", 2)
        owned_samples = owned_section.get("samples", protocol["normal"]["samples"])
        owned_warmup = owned_section.get("warmup", protocol["normal"]["warmup"])
        owned_alloc_samples = owned_section.get("allocation_samples", protocol["allocator"]["samples"])
        owned_alloc_warmup = owned_section.get("allocation_warmup", protocol["allocator"]["warmup"])
        for field, value in {
            "protocol.owned_source.repetitions": owned_count,
            "protocol.owned_source.allocation_repetitions": owned_allocator_count,
            "protocol.owned_source.samples": owned_samples,
            "protocol.owned_source.warmup": owned_warmup,
            "protocol.owned_source.allocation_samples": owned_alloc_samples,
            "protocol.owned_source.allocation_warmup": owned_alloc_warmup,
        }.items():
            _integer(value, field, minimum=1 if "samples" in field or "repetitions" in field else 0)
        declared_owned = [
            {
                "name": "owned",
                "cases": OWNED_CASES,
                "count": owned_count,
                "roles": ("candidate",) * owned_count,
                "samples": owned_samples,
                "warmup": owned_warmup,
            },
            {
                "name": "owned-allocator",
                "cases": OWNED_CASES,
                "count": owned_allocator_count,
                "roles": ("candidate",) * owned_allocator_count,
                "samples": owned_alloc_samples,
                "warmup": owned_alloc_warmup,
            },
        ]
        # The 0412 protocol declares these as part of the deliverable.  Keep
        # them in the required family list even though older pre-0412
        # directories may have omitted them; a declared owned baseline must
        # fail closed when any report is missing.
        specs.extend(declared_owned)
    return specs


def _validate_commands(
    root: Path,
    specs: list[dict[str, Any]],
    loaded: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    commands_value = _read_json(root / "commands.json")
    if isinstance(commands_value, dict):
        commands_value = commands_value.get("commands", commands_value.get("entries", MISSING))
    commands = _list(commands_value, "commands.json")
    by_label: dict[str, dict[str, Any]] = {}
    positions: dict[str, int] = {}
    for index, item_value in enumerate(commands):
        item = _dict(item_value, f"commands.json[{index}]")
        label = _string(item.get("label", MISSING), f"commands.json[{index}].label")
        if label in by_label:
            _fail("commands.json", f"duplicate label {label!r}")
        by_label[label] = item
        positions[label] = index
    selected: dict[str, Any] = {}
    for spec in specs:
        labels = [f"{spec['name']}-{index}" for index in range(1, spec["count"] + 1)]
        previous_position = -1
        selected_records = []
        for index, label in enumerate(labels):
            path = f"commands.json.{label}"
            command = by_label.get(label)
            if command is None:
                _fail(path, "missing command journal record")
            position = positions[label]
            if position <= previous_position:
                _fail(path, "command records are not in sequential capture order")
            previous_position = position
            if command.get("exit_code") != 0:
                _fail(path, "child command did not exit successfully")
            expected_role = spec["roles"][index]
            variant = command.get("variant", command.get("role", command.get("implementation")))
            if variant is not None and str(variant).lower() != expected_role:
                _fail(f"{path}.variant", f"expected {expected_role!r}")
            revision = _string(command.get("revision", MISSING), f"{path}.revision")
            binary_sha = _digest(command.get("binary_sha256", MISSING), f"{path}.binary_sha256")
            argv = _list(command.get("argv", MISSING), f"{path}.argv")
            argv_text = [str(value) for value in argv]
            if _option(argv, "--case", path) != ",".join(spec["cases"]):
                _fail(f"{path}.argv", "case selection differs from protocol")
            if _option(argv, "--filesystem-cache", path) != "warm":
                _fail(f"{path}.argv", "filesystem cache must be warm")
            if _option(argv, "--samples", path) != str(spec["samples"]):
                _fail(f"{path}.argv", "sample count differs from protocol")
            if _option(argv, "--warmup", path) != str(spec["warmup"]):
                _fail(f"{path}.argv", "warmup count differs from protocol")
            executable = EXPECTED_BINARY[spec["name"]]
            if not any(Path(value).name == executable for value in argv_text):
                _fail(f"{path}.argv", f"does not invoke {executable}")
            report_path = Path(_option(argv, "--json", path))
            expected_report_path = root / f"{label}.json"
            if report_path.name != expected_report_path.name:
                _fail(f"{path}.argv", "--json path does not identify the command label")
            catalog_path = Path(_option(argv, "--corpus-manifest", path))
            if catalog_path.name != f"{label}.catalog.json":
                _fail(f"{path}.argv", "--corpus-manifest path does not identify the command label")
            if label not in loaded:
                _fail(path, "internal report index is missing this command label")
            report_info = loaded[label]["info"]
            if report_info["revision"] != revision:
                _fail(f"{path}.revision", "does not match report environment.git_revision")
            if report_info["binary_sha256"] != binary_sha:
                _fail(f"{path}.binary_sha256", "does not match report binary identity")
            if "started_utc" in command:
                _parse_time(command["started_utc"], f"{path}.started_utc")
            if "wall_seconds" in command:
                _number(command["wall_seconds"], f"{path}.wall_seconds")
            selected_records.append({
                "label": label,
                "role": expected_role,
                "revision": revision,
                "binary_sha256": binary_sha,
                "argv": argv_text,
                "position": position,
                "started_utc": command.get("started_utc"),
            })
        timestamps = [
            _parse_time(item["started_utc"], f"commands.json.{item['label']}.started_utc")
            for item in selected_records
            if item["started_utc"] is not None
        ]
        if timestamps and len(timestamps) == len(selected_records) and any(
            right <= left for left, right in zip(timestamps, timestamps[1:])
        ):
            _fail(f"commands.json.{spec['name']}", "timestamps do not prove sequential children")
        selected[spec["name"]] = selected_records
    return {"selected": selected, "record_count": len(commands)}


def _validate_identity_sets(
    specs: list[dict[str, Any]], loaded: dict[str, dict[str, Any]]
) -> dict[str, Any]:
    roles: dict[str, dict[str, Any]] = {}
    for spec in specs:
        identities: dict[str, list[dict[str, Any]]] = {}
        for index in range(1, spec["count"] + 1):
            label = f"{spec['name']}-{index}"
            info = loaded[label]["info"]
            role = spec["roles"][index - 1]
            identities.setdefault(role, []).append(info)
        for role, entries in identities.items():
            first = entries[0]
            for entry in entries[1:]:
                if (entry["revision"], entry["binary_sha256"]) != (
                    first["revision"],
                    first["binary_sha256"],
                ):
                    _fail(
                        f"{spec['name']}.{role}",
                        "repeated children do not use one stable revision/binary identity",
                    )
        if set(identities) == {"control", "candidate"}:
            control = identities["control"][0]
            candidate = identities["candidate"][0]
            if (control["revision"], control["binary_sha256"]) == (
                candidate["revision"],
                candidate["binary_sha256"],
            ):
                _fail(
                    f"{spec['name']}.identity",
                    "control and candidate identities are not distinct",
                )
        roles[spec["name"]] = {
            role: {
                "revision": entries[0]["revision"],
                "binary_sha256": entries[0]["binary_sha256"],
                "binary_bytes": entries[0]["binary"].get("binary_bytes"),
                "binary_path": entries[0]["binary"].get("path"),
            }
            for role, entries in identities.items()
        }
    # The normal and allocator binaries are separate files, but both families
    # must have been built from the same control/candidate source revisions.
    by_role: dict[str, set[str]] = {}
    for spec in specs:
        for role, identity in roles[spec["name"]].items():
            by_role.setdefault(role, set()).add(identity["revision"])
    for role, revisions in by_role.items():
        if len(revisions) != 1:
            _fail(f"identity.{role}", "normal/allocator revisions disagree")
    return roles


def _validate_build_identity_sidecars(
    root: Path,
    protocol: dict[str, Any],
    identities: dict[str, Any],
) -> dict[str, Any]:
    """Bind every report role to its recorded build and protocol revision.

    The build identity files are the durable evidence for the binaries used
    by the short-lived capture children.  The checker does not require those
    executable paths to remain present after capture; it verifies that each
    report's recorded hash and size equal the corresponding sidecar record.
    """

    expected_revisions = {
        "control": protocol["control_revision"].lower(),
        "candidate": protocol["candidate_revision"].lower(),
    }
    sidecars: dict[str, Any] = {}
    for role in ("control", "candidate"):
        path = root / f"{role}-build-identity.json"
        value = _read_json(path)
        obj = _dict(value, f"{role}-build-identity.json")
        revision = _string(
            obj.get("revision", MISSING), f"{role}-build-identity.json.revision"
        ).lower()
        if revision != expected_revisions[role]:
            _fail(
                f"{role}-build-identity.json.revision",
                "does not match protocol revision",
            )
        if obj.get("exit_code") != 0:
            _fail(f"{role}-build-identity.json.exit_code", "build record is not successful")
        binaries = _dict(
            obj.get("binaries", MISSING), f"{role}-build-identity.json.binaries"
        )
        recorded: dict[str, dict[str, Any]] = {}
        for executable in ("litchi-perf-baseline", "litchi-perf-baseline-alloc"):
            binary_path = f"{role}-build-identity.json.binaries.{executable}"
            metadata = _dict(binaries.get(executable, MISSING), binary_path)
            binary_bytes = _integer(
                metadata.get("bytes", MISSING), f"{binary_path}.bytes", minimum=1
            )
            binary_sha = _digest(metadata.get("sha256", MISSING), f"{binary_path}.sha256")
            recorded[executable] = {
                "bytes": binary_bytes,
                "sha256": binary_sha,
            }

        # Every report role must agree with both the protocol and its build
        # record.  The normal/owned binaries and allocator binaries are
        # checked separately because their hashes are intentionally distinct.
        matched_executables: set[str] = set()
        for family, role_identities in identities.items():
            identity = role_identities.get(role)
            if identity is None:
                continue
            executable = EXPECTED_BINARY[family]
            expected_binary = recorded[executable]
            if identity["revision"].lower() != revision:
                _fail(
                    f"identity.{family}.{role}.revision",
                    f"does not match {role}-build-identity.json.revision",
                )
            if identity["binary_sha256"] != expected_binary["sha256"]:
                _fail(
                    f"identity.{family}.{role}.binary_sha256",
                    f"does not match {role}-build-identity.json.binaries.{executable}.sha256",
                )
            if identity["binary_bytes"] != expected_binary["bytes"]:
                _fail(
                    f"identity.{family}.{role}.binary_bytes",
                    f"does not match {role}-build-identity.json.binaries.{executable}.bytes",
                )
            matched_executables.add(executable)
        if matched_executables != set(recorded):
            _fail(
                f"{role}-build-identity.json.binaries",
                "does not cover every captured binary family",
            )
        sidecars[role] = {
            "revision": revision,
            "binaries": recorded,
        }
    return sidecars


def _validate_corpus_and_catalogs(
    specs: list[dict[str, Any]], loaded: dict[str, dict[str, Any]], root: Path
) -> dict[str, Any]:
    corpus_identity: dict[str, Any] | None = None
    catalog_identity: dict[str, str] = {}
    for spec in specs:
        for index in range(1, spec["count"] + 1):
            label = f"{spec['name']}-{index}"
            info = loaded[label]["info"]
            first_case = next(iter(info["cases"].values()))
            corpus = first_case["corpus"]
            projection = {
                "archive_sha256": corpus["archive_sha256"],
                "target_payload_sha256": corpus["target_payload_sha256"],
                "archive_bytes": corpus.get("archive_bytes"),
                "target_payload_bytes": corpus.get("target_payload_bytes"),
                "corpus": corpus,
            }
            if corpus_identity is None:
                corpus_identity = projection
            elif _canonical(projection) != _canonical(corpus_identity):
                _fail(label, "corpus manifest/digests differ from the capture baseline")
            catalog = _read_json(root / f"{label}.catalog.json")
            catalog_obj = _dict(catalog, f"{label}.catalog.json")
            catalog_projection = _validate_catalog_binding(
                catalog_obj,
                f"{label}.catalog.json",
                spec["cases"],
                corpus,
            )
            reference = info["report"].get("corpus_catalog")
            if reference is not None:
                reference = _dict(reference, f"{label}.corpus_catalog")
                if reference.get("content_set_sha256") != catalog_projection["content_set_sha256"]:
                    _fail(f"{label}.corpus_catalog.content_set_sha256", "does not match sidecar")
                if reference.get("catalog_sha256") != catalog_obj.get("catalog_sha256"):
                    _fail(f"{label}.corpus_catalog.catalog_sha256", "does not match sidecar")
            catalog_identity[label] = catalog_projection["content_set_sha256"]
    if corpus_identity is None:  # pragma: no cover - core specs are required
        _fail("capture", "no reports were loaded")
    return {"corpus": corpus_identity, "catalog_content_set_sha256": catalog_identity}


def _validate_scope_contract(
    specs: list[dict[str, Any]], loaded: dict[str, dict[str, Any]]
) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for spec in specs:
        family_name = spec["name"]
        family_result: dict[str, Any] = {}
        for case in spec["cases"]:
            values_by_role: dict[str, list[str]] = {}
            paths_by_role: dict[str, list[tuple[str, ...]]] = {}
            for index in range(1, spec["count"] + 1):
                label = f"{family_name}-{index}"
                role = spec["roles"][index - 1]
                raw_result = loaded[label]["info"]["cases"][case]["result"]
                values = _scope_values(raw_result)
                paths = [path for path, _value in values]
                if case in SOURCE_CASES:
                    if paths != [ALLOWED_SCOPE_PATH]:
                        _fail(
                            f"{label}.{case}",
                            "source-backed row must expose exactly result.source.xls.source_counter_scope",
                        )
                elif paths:
                    _fail(f"{label}.{case}", "non-source row unexpectedly exposes source_counter_scope")
                values_by_role.setdefault(role, []).extend(value for _path, value in values)
                paths_by_role.setdefault(role, []).extend(paths)
            if case in SOURCE_CASES and set(values_by_role) >= {"control", "candidate"}:
                controls = values_by_role["control"]
                candidates = values_by_role["candidate"]
                if not controls or len(set(controls)) != 1:
                    _fail(f"{family_name}.{case}.control_scope", "control scope is not stable")
                if not candidates or len(set(candidates)) != 1:
                    _fail(f"{family_name}.{case}.candidate_scope", "candidate scope is not stable")
                control = controls[0]
                candidate = candidates[0]
                if control == candidate:
                    _fail(
                        f"{family_name}.{case}.source_counter_scope",
                        "candidate observer marker did not differ from control",
                    )
                if not re.search(r"(?<![A-Za-z])v[0-9]+(?![0-9])", candidate, re.IGNORECASE):
                    _fail(
                        f"{family_name}.{case}.candidate_scope",
                        "candidate source scope does not carry a version marker",
                    )
                family_result[case] = {
                    "path": ".".join(ALLOWED_SCOPE_PATH),
                    "control": control,
                    "candidate": candidate,
                    "candidate_version_marker": True,
                    "permitted_only_difference": True,
                }
            elif case in SOURCE_CASES:
                # Optional owned captures have no control/candidate pair; this
                # branch is retained for a future single-role source family.
                family_result[case] = {"status": "not_applicable"}
        result[family_name] = family_result
    return result


def _validate_equivalence(
    specs: list[dict[str, Any]], loaded: dict[str, dict[str, Any]]
) -> dict[str, Any]:
    neutrality: dict[str, Any] = {}
    for spec in specs:
        family = spec["name"]
        labels = [f"{family}-{index}" for index in range(1, spec["count"] + 1)]
        projections = [_project_report(loaded[label]["info"]["report"]) for label in labels]
        for label, projection in zip(labels[1:], projections[1:]):
            differences = _differences(projections[0], projection)
            if differences:
                _fail(
                    f"{family}.{label}",
                    "semantic/logical report projection differs outside the explicit XLS scope marker: "
                    + ", ".join(differences),
                )
        per_case: dict[str, Any] = {}
        for case in spec["cases"]:
            outputs = [loaded[label]["info"]["cases"][case]["output_sha256"] for label in labels]
            if len(set(outputs)) != 1:
                _fail(f"{family}.{case}.output_sha256", "semantic output digest differs between repetitions")
            corpora = [loaded[label]["info"]["cases"][case]["corpus"] for label in labels]
            if len({_canonical(item) for item in corpora}) != 1:
                _fail(f"{family}.{case}.corpus", "corpus identity differs between repetitions")
            source_values = [loaded[label]["info"]["cases"][case]["result"].get("source") for label in labels]
            # Project with the result-level path so the one intentional
            # discriminator exception reaches ``source.xls``.  Calling the
            # helper on the source object alone would make its local path
            # start at ``xls`` and incorrectly report the marker as a data
            # difference.
            source_projection = [
                _project_result(item, ("source",)) for item in source_values
            ]
            if len({_canonical(item) for item in source_projection}) != 1:
                _fail(f"{family}.{case}.source", "logical source/locality evidence differs outside scope marker")
            per_case[case] = {
                "corpus_digest_equal": True,
                "output_sha256_equal": True,
                "logical_io_and_locality_equal": True,
                "semantic_output_equal": True,
                "only_allowed_source_scope_difference": case in SOURCE_CASES and set(spec["roles"]) == {"control", "candidate"},
            }
        neutrality[family] = per_case
    return neutrality


def _validate_owned_output_pair(
    loaded: dict[str, dict[str, Any]], specs: list[dict[str, Any]]
) -> None:
    spec_names = {spec["name"] for spec in specs}
    if not {"owned", "normal"}.issubset(spec_names):
        return
    normal_info = loaded["normal-1"]["info"]
    for owned_family in ("owned", "owned-allocator"):
        if owned_family not in spec_names:
            continue
        spec = next(spec for spec in specs if spec["name"] == owned_family)
        for index in range(1, spec["count"] + 1):
            label = f"{owned_family}-{index}"
            owned_cases = loaded[label]["info"]["cases"]
            for owned_case, normal_case in zip(OWNED_CASES, ("xls_source_backed_open", "xls_source_backed_open_list_worksheets", "xls_source_backed_open_one_cell")):
                owned_output = owned_cases[owned_case]["output_sha256"]
                normal_output = normal_info["cases"][normal_case]["output_sha256"]
                if owned_output != normal_output:
                    _fail(
                        f"{label}.{owned_case}.output_sha256",
                        f"owned-source output does not match {normal_case} oracle",
                    )
                if owned_cases[owned_case]["result"].get("source") is not None:
                    _fail(f"{label}.{owned_case}.source", "owned-source runner must leave source=null")


def _build_report_summary(
    specs: list[dict[str, Any]], loaded: dict[str, dict[str, Any]], catalogs: dict[str, str], identities: dict[str, Any]
) -> dict[str, Any]:
    reports: dict[str, Any] = {}
    for spec in specs:
        family = spec["name"]
        rows: list[dict[str, Any]] = []
        for index in range(1, spec["count"] + 1):
            label = f"{family}-{index}"
            info = loaded[label]["info"]
            case_rows: dict[str, Any] = {}
            for case in spec["cases"]:
                detail = info["cases"][case]
                elapsed = detail["elapsed"]
                row: dict[str, Any] = {
                    "elapsed_ns": {
                        **elapsed["summary"],
                        "sample_order_sha256": elapsed["sample_order_sha256"],
                        "claimable": False,
                    },
                    "output_sha256": detail["output_sha256"],
                }
                allocation = detail["operation"]["allocation"]
                if allocation is not None:
                    row["allocator"] = {
                        field: {
                            "unit": "calls" if field in ALLOCATION_COUNT_FIELDS else "bytes",
                            **_summary(values),
                        }
                        for field, values in allocation.items()
                    }
                case_rows[case] = row
            rows.append(
                {
                    "label": label,
                    "role": spec["roles"][index - 1],
                    "revision": info["revision"],
                    "binary_sha256": info["binary_sha256"],
                    "catalog_content_set_sha256": catalogs[label],
                    "cases": case_rows,
                }
            )
        reports[family] = rows
    return reports


def verify(root: Path) -> dict[str, Any]:
    if not root.is_dir():
        raise ValidationError(f"capture directory does not exist: {root}")
    protocol_value = _read_json(root / "protocol.json")
    protocol = _validate_protocol(protocol_value)
    specs = _family_specs(protocol, root)
    loaded: dict[str, dict[str, Any]] = {}
    for spec in specs:
        for index in range(1, spec["count"] + 1):
            label = f"{spec['name']}-{index}"
            report_meta = _read_json_with_meta(root / f"{label}.json")
            if report_meta is None:  # core families are always required
                raise ValidationError(f"missing report artifact for {label}")
            info = _validate_report_header(
                report_meta["value"],
                label,
                spec["name"],
                spec["cases"],
                spec["samples"],
                spec["warmup"],
            )
            loaded[label] = {"info": info, "artifact": report_meta}
    command_info = _validate_commands(root, specs, loaded)
    identities = _validate_identity_sets(specs, loaded)
    build_identity = _validate_build_identity_sidecars(root, protocol, identities)
    corpus_info = _validate_corpus_and_catalogs(specs, loaded, root)
    scope_info = _validate_scope_contract(specs, loaded)
    neutrality = _validate_equivalence(specs, loaded)
    _validate_owned_output_pair(loaded, specs)
    reports = _build_report_summary(
        specs, loaded, corpus_info["catalog_content_set_sha256"], identities
    )
    capture_order = {
        spec["name"]: [
            {"label": f"{spec['name']}-{index}", "role": spec["roles"][index - 1]}
            for index in range(1, spec["count"] + 1)
        ]
        for spec in specs
    }
    return {
        "schema_version": "0412-observer-comparison-v1",
        "status": "pass",
        "performance_claim": "none",
        "comparison_mode": "observer-only fixed warm in-memory XLS capture",
        "allowed_result_difference_paths": ["result.source.xls.source_counter_scope"],
        "equivalence_exclusions": {
            "capture_identity": [
                "binary_identity",
                "environment.git_revision",
                "corpus_catalog",
            ],
            "descriptive_measurements": [
                "result.elapsed_ns",
                "result.operation_metrics.*.values",
            ],
            "reason": "Fresh-child timing, process snapshots, and allocator vectors are summarized per repetition; they are not semantic or logical-I/O equivalence fields.",
        },
        "protocol": protocol,
        "capture_order": capture_order,
        "identity": identities,
        "build_identity": build_identity,
        "commands": command_info,
        "validation": {
            "operation_metrics": "repository perf_compare._validate_operation_metrics"
            if _PERF_COMPARE is not None
            else "local strict cardinality/alignment fallback",
            "elapsed_sample_order": "exact elapsed-then-original-index permutation",
            "catalog_binding": "catalog corpus/archive/workbook digests and selected-case bindings",
        },
        "corpus": {
            "archive_sha256": corpus_info["corpus"]["archive_sha256"],
            "target_payload_sha256": corpus_info["corpus"]["target_payload_sha256"],
            "archive_bytes": corpus_info["corpus"].get("archive_bytes"),
            "target_payload_bytes": corpus_info["corpus"].get("target_payload_bytes"),
            "all_report_corpus_digests_equal": True,
            "catalog_content_set_equal": len(set(corpus_info["catalog_content_set_sha256"].values())) == 1,
        },
        "source_counter_scope": scope_info,
        "neutrality": neutrality,
        "reports": reports,
        "limitations": [
            "Elapsed p50/p95/p99 and allocator vectors are per-child descriptive observations.",
            "This observer-only comparison makes no production speedup or regression claim.",
            "The warm generated corpus does not establish cold-file, physical-I/O, remote-source, scaling, or cache behavior.",
            "The source scope marker difference is explicitly permitted only at result.source.xls.source_counter_scope; logical ReadAt vectors, locality flags, output digests, and corpus identity remain equal.",
        ],
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root",
        type=Path,
        required=True,
        help="0412 capture directory containing protocol.json, reports, catalogs, and commands.json",
    )
    parser.add_argument(
        "--output",
        "--json-out",
        type=Path,
        help="summary JSON path (default: ROOT/0412-comparison.json)",
    )
    args = parser.parse_args(argv)
    output = args.output or args.root / "0412-comparison.json"
    try:
        result = verify(args.root)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(
            json.dumps(result, indent=2, sort_keys=True, allow_nan=False) + "\n",
            encoding="utf-8",
        )
    except (OSError, ValidationError, ValueError, TypeError) as error:
        print(f"litchi-goal-0412-compare: FAIL: {error}", file=sys.stderr)
        return 2
    print(f"litchi-goal-0412-compare: PASS: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
