#!/usr/bin/env python3
"""Create the sealed 0471 memory experiment summary.

The benchmark harness and the repository comparators remain the authorities for
sample validation and full-matrix comparisons.  This file only selects the
predeclared seven-row ABBA probe, parses process-lifetime RSS and Heaptrack's
retained process totals, and records immutable input bindings.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import math
import re
import sys
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0471-evidence-analysis-v1"
LANES = ("A1", "B1", "B2", "A2")
FULL_LANES = ("A-full", "B-full")
HEAP_LANES = ("A-heap", "B-heap")
PROBE_CASES = (
    "xlsx_one_cell_commit_save",
    "xlsx_one_percent_commit_save",
    "ppt_fresh_write_to",
)
PROBE_SHAPES = ("tiny", "medium", "dense-wide")
WRITER_SHAPES = ("payload-heavy",)
PROBE_SAMPLES = 100
PROBE_WARMUPS = 5
FULL_SAMPLES = 15
FULL_WARMUPS = 3
HEAP_SAMPLES = 5
HEAP_WARMUPS = 1


class AnalysisError(ValueError):
    """Raised when the retained evidence is incomplete or incomparable."""


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
    try:
        with path.open("rb") as stream:
            size = 0
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except OSError as error:
        raise AnalysisError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest(), size


def file_binding(path: Path, root: Path = ROOT) -> dict[str, Any]:
    digest, size = sha256_file(path)
    try:
        name = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        name = f"tools/{path.name}" if path.parent.name == "tools" else path.name
    return {"path": name, "sha256": digest, "bytes": size}


def _load_module(path: Path, name: str) -> Any:
    if not path.is_file():
        raise AnalysisError(f"missing dependency: {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise AnalysisError(f"cannot load dependency: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except (OSError, SyntaxError, TypeError, ValueError) as error:
        raise AnalysisError(f"cannot load {path}: {error}") from error
    return module


def _repository_file(root: Path, name: str) -> Path:
    for parent in (root, *root.parents):
        candidate = parent / "tools" / name
        if candidate.is_file():
            return candidate
    raise AnalysisError(f"portable dependency tools/{name} is missing")


def load_comparator(root: Path = ROOT) -> tuple[Any, Path]:
    path = _repository_file(root, "perf_compare.py")
    return _load_module(path, "litchi_change0471_perf_compare"), path


def load_abba(root: Path = ROOT) -> tuple[Any, Path]:
    comparator, _ = load_comparator(root)
    sys.modules["perf_compare"] = comparator
    path = _repository_file(root, "perf_abba_summary.py")
    return _load_module(path, "litchi_change0471_perf_abba_summary"), path


def _require_object(value: Any, label: str) -> Mapping[str, Any]:
    if not isinstance(value, dict):
        raise AnalysisError(f"{label} must be an object")
    return value


def expected_rows(protocol: Mapping[str, Any]) -> set[tuple[str, str]]:
    cases = tuple(protocol.get("probe_cases", ()))
    shapes = tuple(protocol.get("probe_shapes", ()))
    writer_shapes = tuple(protocol.get("writer_shapes", ()))
    if cases != PROBE_CASES or shapes != PROBE_SHAPES or writer_shapes != WRITER_SHAPES:
        raise AnalysisError("protocol probe selectors differ from the fixed 0471 matrix")
    rows = {
        (case, shape)
        for case in PROBE_CASES[:2]
        for shape in ("tiny", "medium", "dense-wide")
    }
    rows.add(("ppt_fresh_write_to", "payload-heavy"))
    return rows


def _report_rows(report: Mapping[str, Any], label: str) -> dict[tuple[str, str], Mapping[str, Any]]:
    raw = report.get("results")
    if not isinstance(raw, list):
        raise AnalysisError(f"{label}.results must be an array")
    rows: dict[tuple[str, str], Mapping[str, Any]] = {}
    for index, row in enumerate(raw):
        item = _require_object(row, f"{label}.results[{index}]")
        case = item.get("case")
        corpus = _require_object(item.get("corpus"), f"{label}.results[{index}].corpus")
        shape = corpus.get("shape")
        if not isinstance(case, str) or not isinstance(shape, str):
            raise AnalysisError(f"{label}.results[{index}] lacks case/shape")
        key = (case, shape)
        if key in rows:
            raise AnalysisError(f"{label}: duplicate result row {key!r}")
        rows[key] = item
    return rows


def validate_probe_reports(
    reports: Sequence[Mapping[str, Any]], protocol: Mapping[str, Any]
) -> set[tuple[str, str]]:
    if len(reports) != 4:
        raise AnalysisError("the normal probe requires A1, B1, B2 and A2")
    expected = expected_rows(protocol)
    for index, report in enumerate(reports):
        label = LANES[index]
        if report.get("schema_version") != 1:
            raise AnalysisError(f"{label}: unsupported report schema")
        configuration = _require_object(report.get("configuration"), f"{label}.configuration")
        if configuration.get("samples_per_case") != PROBE_SAMPLES:
            raise AnalysisError(f"{label}: sample count is not 100")
        if configuration.get("warmup_iterations_per_case") != PROBE_WARMUPS:
            raise AnalysisError(f"{label}: warmup count is not 5")
        if tuple(configuration.get("cases", ())) != PROBE_CASES:
            raise AnalysisError(f"{label}: case selector differs from protocol")
        if tuple(configuration.get("writer_shapes", ())) != WRITER_SHAPES:
            raise AnalysisError(f"{label}: writer selector differs from protocol")
        rows = _report_rows(report, label)
        if set(rows) != expected:
            raise AnalysisError(f"{label}: normal row set differs from protocol")
        for key, row in rows.items():
            elapsed = _require_object(row.get("elapsed_ns"), f"{label}.{key}.elapsed_ns")
            samples = elapsed.get("samples")
            if not isinstance(samples, list) or len(samples) != PROBE_SAMPLES:
                raise AnalysisError(f"{label}.{key}: expected 100 elapsed samples")
    return expected


def comparison_projection(report: Mapping[str, Any]) -> tuple[dict[str, Any], list[str]]:
    """Remove empty optional source vectors only in a comparator copy."""

    projected = copy.deepcopy(dict(report))
    removed: list[str] = []

    def visit(value: Any, path: str) -> Any:
        if isinstance(value, dict):
            output: dict[str, Any] = {}
            for key, item in value.items():
                child = f"{path}/{key}"
                if isinstance(item, list) and not item:
                    removed.append(child)
                else:
                    output[key] = visit(item, child)
            return output
        if isinstance(value, list):
            return [visit(item, f"{path}/{index}") for index, item in enumerate(value)]
        return value

    rows = projected.get("results")
    if not isinstance(rows, list):
        raise AnalysisError("full guard report lacks results")
    for index, row in enumerate(rows):
        if isinstance(row, dict) and isinstance(row.get("source"), dict):
            row["source"] = visit(row["source"], f"results/{index}/source")
    return projected, sorted(removed)


def parse_gnu_rss(path: Path, *, root: Path = ROOT, latency: str = "excluded") -> dict[str, Any]:
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise AnalysisError(f"cannot read RSS log {path}: {error}") from error
    values = []
    for line in text.splitlines():
        marker = "Maximum resident set size (kbytes):"
        if marker in line:
            raw = line.split(":", 1)[1].strip().replace(",", "")
            try:
                value = int(raw)
            except ValueError as error:
                raise AnalysisError(f"invalid maximum RSS in {path}: {raw!r}") from error
            if value < 0:
                raise AnalysisError(f"negative maximum RSS in {path}")
            values.append(value)
    if len(values) != 1:
        raise AnalysisError(f"expected exactly one maximum RSS value in {path}")
    kib = values[0]
    return {
        "path": file_binding(path, root),
        "unit": "KiB",
        "maximum_resident_set_kib": kib,
        "maximum_resident_set_bytes": kib * 1024,
        "scope": "whole_process_lifetime_including_setup_and_teardown",
        "latency_comparison": latency,
    }


def parse_heaptrack(path: Path, *, root: Path = ROOT) -> dict[str, Any]:
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise AnalysisError(f"cannot read Heaptrack output {path}: {error}") from error

    def integer(pattern: str, label: str) -> int:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        if match is None:
            raise AnalysisError(f"Heaptrack output lacks {label}: {path}")
        try:
            value = int(match.group(1).replace(",", ""))
        except ValueError as error:
            raise AnalysisError(f"invalid Heaptrack {label}: {path}") from error
        if value < 0:
            raise AnalysisError(f"negative Heaptrack {label}: {path}")
        return value

    peak = re.search(r"^peak heap memory consumption:\s*([^\r\n]+)", text, re.IGNORECASE | re.MULTILINE)
    if peak is None or not peak.group(1).strip():
        raise AnalysisError(f"Heaptrack output lacks peak heap display: {path}")
    return {
        "path": file_binding(path, root),
        "scope": "whole_process",
        "allocation_calls": integer(r"^calls to allocation functions:\s*([0-9]+)", "allocation calls"),
        "temporary_allocations": integer(r"^temporary memory allocations:\s*([0-9]+)", "temporary allocations"),
        "peak_heap_display": peak.group(1).strip(),
        "latency_comparison": "excluded",
        "claim": "diagnostic_only; no allocation-count or latency claim",
    }


def _rss_delta(
    rss: Mapping[str, Mapping[str, Any]],
    name: str,
    baseline_lane: str,
    current_lane: str,
) -> dict[str, Any]:
    baseline = rss[baseline_lane]
    current = rss[current_lane]
    baseline_bytes = baseline["maximum_resident_set_bytes"]
    current_bytes = current["maximum_resident_set_bytes"]
    delta = ((current_bytes / baseline_bytes) - 1.0) * 100.0 if baseline_bytes else None
    return {
        "name": name,
        "baseline_lane": baseline_lane,
        "current_lane": current_lane,
        "baseline_kib": baseline["maximum_resident_set_kib"],
        "current_kib": current["maximum_resident_set_kib"],
        "candidate_delta_percent": delta,
        "claim": "descriptive_only; no RSS claim",
    }


def _full_guard_flags(
    control: Mapping[str, Any], candidate: Mapping[str, Any], comparator: Any
) -> dict[str, Any]:
    """Recompute raw full-guard latency flags at the explicit five-percent level."""

    def rows(report: Mapping[str, Any], label: str) -> dict[tuple[str, bytes], Mapping[str, Any]]:
        result: dict[tuple[str, bytes], Mapping[str, Any]] = {}
        raw = report.get("results")
        if not isinstance(raw, list):
            raise AnalysisError(f"{label}.results is not an array")
        for index, row in enumerate(raw):
            if not isinstance(row, dict) or not isinstance(row.get("case"), str):
                raise AnalysisError(f"{label}.results[{index}] is malformed")
            corpus = row.get("corpus")
            key = (row["case"], canonical(corpus))
            if key in result:
                raise AnalysisError(f"{label}: duplicate full-guard row")
            result[key] = row
        return result

    before = rows(control, "A-full")
    after = rows(candidate, "B-full")
    if set(before) != set(after):
        raise AnalysisError("full guard row identities differ")

    def stats(row: Mapping[str, Any], label: str) -> dict[str, float]:
        elapsed = row.get("elapsed_ns")
        if not isinstance(elapsed, dict) or not isinstance(elapsed.get("samples"), list):
            raise AnalysisError(f"{label}: elapsed samples missing")
        samples = elapsed["samples"]
        if len(samples) != FULL_SAMPLES or any(isinstance(value, bool) or not isinstance(value, int) or value <= 0 for value in samples):
            raise AnalysisError(f"{label}: elapsed samples malformed")
        validated = comparator._latencies(dict(row), label, FULL_SAMPLES)
        return {"mean": math.fsum(samples) / len(samples), **validated}

    individual: list[dict[str, Any]] = []
    uniform: list[dict[str, Any]] = []
    for key in sorted(before):
        left = stats(before[key], f"A-full.{key[0]}")
        right = stats(after[key], f"B-full.{key[0]}")
        deltas = {
            metric: ((right[metric] / left[metric]) - 1.0) * 100.0
            if left[metric] != 0
            else (0.0 if right[metric] == 0 else math.inf)
            for metric in ("mean", "p50", "p95", "p99")
        }
        corpus = before[key]["corpus"]
        if any(value > 5.0 for value in deltas.values()):
            individual.append({"case": key[0], "corpus": corpus, "deltas_percent": deltas})
        if all(value > 5.0 for value in deltas.values()):
            uniform.append({"case": key[0], "corpus": corpus, "deltas_percent": deltas})
    return {
        "threshold_percent": 5.0,
        "individual_rows_above_threshold": individual,
        "uniform_rows_above_threshold_all_mean_p50_p95_p99": uniform,
    }


def _binding(path: Path, root: Path) -> dict[str, Any]:
    return file_binding(path, root)


def build_summary(root: Path = ROOT) -> dict[str, Any]:
    root = root.resolve()
    protocol = _require_object(read_json(root / "protocol.json"), "protocol.json")
    reports = [
        _require_object(read_json(root / lane / "report.json"), f"{lane}/report.json")
        for lane in LANES
    ]
    expected = validate_probe_reports(reports, protocol)
    abba, abba_path = load_abba(root)
    try:
        normal = abba.summarize_reports(
            reports=reports,
            profile="current-v1",
            cases=protocol["probe_cases"],
            shapes=tuple(protocol["probe_shapes"]) + tuple(protocol["writer_shapes"]),
        )
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        raise AnalysisError(f"canonical ABBA summary failed: {error}") from error
    selected = {(row.get("case"), row.get("shape")) for row in normal.get("results", [])}
    if selected != expected:
        raise AnalysisError("canonical ABBA selector returned an unexpected row set")

    policy_path = root / "report-policy.json"
    policy = _require_object(read_json(policy_path), "report-policy.json")
    full_reports = [
        _require_object(read_json(root / lane / "report.json"), f"{lane}/report.json")
        for lane in FULL_LANES
    ]
    full_before, before_empty = comparison_projection(full_reports[0])
    full_after, after_empty = comparison_projection(full_reports[1])
    if before_empty != after_empty:
        raise AnalysisError("full guard optional source vector presence differs")
    comparator, comparator_path = load_comparator(root)
    try:
        full_comparison = comparator.compare_reports(full_before, full_after, policy)
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        raise AnalysisError(f"canonical full guard comparison failed: {error}") from error
    full_flags = _full_guard_flags(full_reports[0], full_reports[1], comparator)

    rss = {
        lane: parse_gnu_rss(root / lane / "resource.log", root=root, latency="descriptive_only")
        for lane in (*LANES, *FULL_LANES)
    }
    heap: dict[str, Any] = {}
    for lane in HEAP_LANES:
        heap[lane] = {
            "rss": parse_gnu_rss(root / lane / "resource.log", root=root, latency="excluded_instrumented"),
            "heaptrack": parse_heaptrack(root / lane / "heaptrack-print.stdout", root=root),
        }

    control_heap = heap["A-heap"]
    candidate_heap = heap["B-heap"]
    heap_comparison = {
        "scope": "whole_process; Heaptrack instrumentation overhead is included in these heap runs",
        "instrumented_resource_rss": {
            "control_kib": control_heap["rss"]["maximum_resident_set_kib"],
            "candidate_kib": candidate_heap["rss"]["maximum_resident_set_kib"],
            "comparison": "excluded",
            "reason": "GNU time measured the Heaptrack-instrumented process; these values are not retention evidence",
        },
        "peak_heap_display": {
            "control": control_heap["heaptrack"]["peak_heap_display"],
            "candidate": candidate_heap["heaptrack"]["peak_heap_display"],
            "equal_display": control_heap["heaptrack"]["peak_heap_display"] == candidate_heap["heaptrack"]["peak_heap_display"],
            "numeric_claim": "withheld; retained Heaptrack display is rounded",
        },
        "allocation_counts": {
            "control": control_heap["heaptrack"]["allocation_calls"],
            "candidate": candidate_heap["heaptrack"]["allocation_calls"],
            "temporary_control": control_heap["heaptrack"]["temporary_allocations"],
            "temporary_candidate": candidate_heap["heaptrack"]["temporary_allocations"],
            "claim": "diagnostic_only",
        },
    }

    return {
        "schema": SCHEMA,
        "claim": {
            "registered": False,
            "latency": "none",
            "allocation_counts": "diagnostic_only",
            "scope": "whole_process_memory_evidence; instrumented time and RSS excluded",
        },
        "protocol": _binding(root / "protocol.json", root),
        "normal_abba": {
            "purpose": "100-sample diagnostic ABBA; no registered latency claim",
            "inputs": {
                "reports": [_binding(root / lane / "report.json", root) for lane in LANES],
                "abba_tool": _binding(abba_path, root),
            },
            "summary": normal,
        },
        "full_guard": {
            "purpose": "15-sample full-matrix descriptive guard; flags retained for review",
            "inputs": {
                "reports": [_binding(root / lane / "report.json", root) for lane in FULL_LANES],
                "policy": _binding(policy_path, root),
                "comparator": _binding(comparator_path, root),
            },
            "empty_optional_source_vectors": {
                "control": before_empty,
                "candidate": after_empty,
            },
            "comparison": full_comparison,
            "raw_five_percent_flags": full_flags,
        },
        "rss": rss,
        "normal_rss_pairs": {
            "scope": "uninstrumented normal-process GNU time RSS; descriptive paired readings",
            "pairs": [
                _rss_delta(rss, "control_repetition", "A1", "A2"),
                _rss_delta(rss, "candidate_repetition", "B1", "B2"),
                _rss_delta(rss, "A1_control_to_B1_candidate", "A1", "B1"),
                _rss_delta(rss, "A2_control_to_B2_candidate", "A2", "B2"),
            ],
            "claim": "none",
        },
        "heap": {
            "purpose": "Heaptrack process-total peak-memory diagnostic; no allocation-count or latency claim",
            "control": control_heap,
            "candidate": candidate_heap,
            "comparison": heap_comparison,
        },
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--output", type=Path, default=None)
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    root = args.root.resolve()
    output = (args.output or root / "summary.json").resolve()
    try:
        result = build_summary(root)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": "pass", "output": str(output)}))
        return 0
    except (AnalysisError, OSError, TypeError, ValueError, KeyError) as error:
        print(json.dumps({"status": "fail", "error": str(error)}, sort_keys=True))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
