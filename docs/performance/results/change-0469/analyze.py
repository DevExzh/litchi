#!/usr/bin/env python3
"""Replay the 0469 diagnostic evidence with the repository comparators.

This module intentionally contains no benchmark math.  The four probe reports
are passed to ``tools/perf_abba_summary.py`` and the full guard reports are
passed to ``tools/perf_compare.py``.  The resulting document is descriptive
evidence only; in particular, a passing probe or guard does not register a
latency claim.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import sys
from typing import Any, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
SCHEMA = "litchi-0469-evidence-analysis-v1"
ABBA_SCHEMA = "current-v1"
PROBE_LANES = ("A1", "B1", "B2", "A2")
PROBE_CASES = (
    "xlsx_one_cell_commit_save",
    "xlsx_one_percent_commit_save",
)
PROBE_SHAPES = ("tiny", "medium", "dense-wide")
PROBE_SAMPLES = 100
PROBE_WARMUPS = 5
HEAP_LANES = ("A-heap", "B-heap")
FULL_LANES = ("A-full", "B-full")


class AnalysisError(ValueError):
    """Raised when evidence cannot be replayed safely."""


def _reject_nonfinite(value: str) -> None:
    raise ValueError(f"non-finite JSON value {value!r}")


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def load_json(path: Path) -> dict[str, Any]:
    try:
        with path.open(encoding="utf-8") as stream:
            value = json.load(
                stream,
                object_pairs_hook=_reject_duplicate_keys,
                parse_constant=_reject_nonfinite,
            )
    except (OSError, UnicodeError, ValueError) as error:
        raise AnalysisError(f"cannot read {path}: {error}") from error
    if not isinstance(value, dict):
        raise AnalysisError(f"{path}: expected a JSON object")
    return value


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
        if path.parent.name == "tools":
            name = f"tools/{path.name}"
        else:
            name = path.name
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


def repository_root(root: Path = ROOT) -> Path:
    for candidate in root.parents:
        if (candidate / "tools").is_dir():
            return candidate
    raise AnalysisError("cannot locate repository tools directory")


def load_abba(root: Path = ROOT) -> tuple[Any, Path]:
    # The ABBA wrapper imports ``perf_compare`` when loaded as a standalone
    # file.  Load and register that exact module name first so this remains
    # portable without modifying sys.path globally.
    load_comparator(root)
    path = repository_root(root) / "tools" / "perf_abba_summary.py"
    return _load_module(path, "litchi_change0469_perf_abba_summary"), path


def load_comparator(root: Path = ROOT) -> tuple[Any, Path]:
    path = repository_root(root) / "tools" / "perf_compare.py"
    module = _load_module(path, "litchi_change0469_perf_compare")
    sys.modules.setdefault("perf_compare", module)
    return module, path


def _report_paths(root: Path, lanes: Sequence[str]) -> list[Path]:
    return [root / lane / "report.json" for lane in lanes]


def _require_current_reports(reports: Sequence[Mapping[str, Any]], minimum: int) -> None:
    if len(reports) != 4:
        raise AnalysisError("ABBA requires exactly four reports")
    for index, report in enumerate(reports):
        if report.get("schema_version") != 1:
            raise AnalysisError(f"ABBA report {index} has an unsupported schema")
        configuration = report.get("configuration")
        if not isinstance(configuration, dict):
            raise AnalysisError(f"ABBA report {index} lacks configuration")
        if configuration.get("samples_per_case", 0) < minimum:
            raise AnalysisError(
                f"ABBA report {index} has fewer than {minimum} samples per case"
            )
        binary = report.get("binary_identity")
        if not isinstance(binary, dict):
            raise AnalysisError(f"ABBA report {index} lacks current-v1 binary identity")


def _validate_probe_shape(summary: Mapping[str, Any]) -> None:
    results = summary.get("results")
    if not isinstance(results, list) or len(results) != len(PROBE_CASES) * len(PROBE_SHAPES):
        raise AnalysisError("ABBA summary does not contain the six probe rows")
    observed = {(row.get("case"), row.get("shape")) for row in results if isinstance(row, dict)}
    expected = {(case, shape) for case in PROBE_CASES for shape in PROBE_SHAPES}
    if observed != expected:
        raise AnalysisError("ABBA summary probe case/shape set differs")
    for row in results:
        elapsed = row.get("elapsed_ns") if isinstance(row, dict) else None
        if not isinstance(elapsed, dict) or elapsed.get("sample_count", 0) < PROBE_SAMPLES:
            raise AnalysisError("ABBA summary probe row is below the 100-sample floor")


def summarize_probe(
    reports: Sequence[Mapping[str, Any]],
    *,
    root: Path = ROOT,
    report_paths: Sequence[Path] = (),
) -> dict[str, Any]:
    """Return the canonical current-v1 ABBA summary plus evidence metadata."""

    _require_current_reports(reports, PROBE_SAMPLES)
    abba, abba_path = load_abba(root)
    try:
        summary = abba.summarize_reports(
            reports=reports,
            profile=ABBA_SCHEMA,
            cases=PROBE_CASES,
            shapes=PROBE_SHAPES,
        )
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        raise AnalysisError(f"canonical ABBA summary failed: {error}") from error
    _validate_probe_shape(summary)
    return {
        "schema": SCHEMA,
        "purpose": "descriptive 100-sample ABBA probe; no registered performance claim",
        "claim": {
            "registered": False,
            "kind": "descriptive_diagnostic",
            "latency_claim": "none",
        },
        "probe": {
            "order": list(PROBE_LANES),
            "cases": list(PROBE_CASES),
            "shapes": list(PROBE_SHAPES),
            "samples_per_case_floor": PROBE_SAMPLES,
            "warmups": PROBE_WARMUPS,
        },
        "inputs": {
            "reports": [file_binding(path, root) for path in report_paths]
            if report_paths
            else [],
            "abba_summary": file_binding(abba_path, root),
        },
        "abba": summary,
    }


def _load_policy(path: Path) -> dict[str, Any]:
    return load_json(path)


def comparison_projection(report: Mapping[str, Any]) -> tuple[dict[str, Any], list[str]]:
    """Treat empty optional source vectors as absent, as in the 0467 guard.

    Retain raw reports. Never touch elapsed samples or operation MetricVectors.
    Return every removed path so both roles must have the same empty fields.
    """
    projected = copy.deepcopy(dict(report))
    removed = []

    def visit(value: Any, path: str) -> Any:
        if isinstance(value, dict):
            output = {}
            for key, item in value.items():
                child = f"{path}/{key}"
                if isinstance(item, list) and not item:
                    removed.append(child)
                else:
                    output[key] = visit(item, child)
            return output
        if isinstance(value, list):
            return [visit(item, f"{path}/{i}") for i, item in enumerate(value)]
        return value

    for index, row in enumerate(projected["results"]):
        if "source" in row:
            row["source"] = visit(row["source"], f"results/{index}/source")
    return projected, sorted(removed)


def compare_guard(
    control: Mapping[str, Any],
    candidate: Mapping[str, Any],
    policy: Mapping[str, Any],
    *,
    root: Path = ROOT,
    report_paths: Sequence[Path] = (),
) -> dict[str, Any]:
    """Run the canonical full-matrix comparator and label it as a guard."""

    comparator, comparator_path = load_comparator(root)
    try:
        control_projected, control_empty = comparison_projection(control)
        candidate_projected, candidate_empty = comparison_projection(candidate)
        if control_empty != candidate_empty:
            raise AnalysisError("full guard empty optional source vector presence differs")
        result = comparator.compare_reports(control_projected, candidate_projected, policy)
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        raise AnalysisError(f"canonical full guard comparison failed: {error}") from error
    return {
        "purpose": "descriptive correctness and regression guard; no registered latency claim",
        "latency_claim": "none",
        "inputs": {
            "reports": [file_binding(path, root) for path in report_paths]
            if report_paths
            else [],
            "comparator": file_binding(comparator_path, root),
        },
        "empty_optional_source_vectors": {
            "scope": "comparison copy only; unchanged raw reports and measured values",
            "control": control_empty,
            "candidate": candidate_empty,
        },
        "comparison": result,
    }


def parse_gnu_rss(
    path: Path,
    *,
    root: Path = ROOT,
    latency_comparison: str = "excluded",
) -> dict[str, Any]:
    """Extract GNU ``time -v``'s process-lifetime maximum RSS."""

    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise AnalysisError(f"cannot read GNU time output {path}: {error}") from error
    values: list[int] = []
    for line in text.splitlines():
        if "Maximum resident set size (kbytes):" not in line:
            continue
        raw = line.split(":", 1)[1].strip().replace(",", "")
        try:
            value = int(raw)
        except ValueError as error:
            raise AnalysisError(f"invalid maximum RSS in {path}: {raw!r}") from error
        if value < 0:
            raise AnalysisError(f"negative maximum RSS in {path}")
        values.append(value)
    if len(values) != 1:
        raise AnalysisError(f"expected one GNU maximum RSS value in {path}")
    kib = values[0]
    return {
        "path": file_binding(path, root),
        "unit": "KiB",
        "maximum_resident_set_kib": kib,
        "maximum_resident_set_bytes": kib * 1024,
        "scope": "whole_process_lifetime_including_setup_and_teardown",
        "latency_comparison": latency_comparison,
    }


def _heap_number(token: str) -> int | None:
    token = token.strip().replace(",", "")
    try:
        return int(token)
    except ValueError:
        return None


def parse_heaptrack(path: Path, *, root: Path = ROOT) -> dict[str, Any]:
    """Parse process-total fields from the retained heaptrack_print output."""

    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise AnalysisError(f"cannot read heaptrack output {path}: {error}") from error
    patterns: tuple[tuple[str, str, Any], ...] = (
        ("allocation_calls", r"^calls to allocation functions:[ \t]*([0-9]+)(?:[ \t]+\([0-9]+/s\))?[ \t]*$", _heap_number),
        ("temporary_allocations", r"^temporary memory allocations:[ \t]*([0-9]+)(?:[ \t]+\([0-9]+/s\))?[ \t]*$", _heap_number),
        (
            "peak_heap_display",
            r"^peak heap memory consumption:\s*([^\r\n]+)",
            lambda token: token.strip(),
        ),
    )
    parsed: dict[str, Any] = {"path": file_binding(path, root), "scope": "whole_process"}
    for key, pattern, converter in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
        parsed[key] = converter(match.group(1)) if match else None
    if parsed["allocation_calls"] is None or parsed["temporary_allocations"] is None:
        raise AnalysisError(f"heaptrack output lacks process allocation totals: {path}")
    if parsed["peak_heap_display"] is None:
        raise AnalysisError(f"heaptrack output lacks peak heap total: {path}")
    parsed["rss_source"] = "gnu_time_resource_log"
    parsed["latency_comparison"] = "excluded"
    return parsed


def build_summary(
    *,
    root: Path = ROOT,
    probe_paths: Sequence[Path] | None = None,
    full_paths: Sequence[Path] | None = None,
    policy_path: Path | None = None,
    heap_paths: Sequence[Path] | None = None,
) -> dict[str, Any]:
    root = root.resolve()
    probe_paths = list(probe_paths or _report_paths(root, PROBE_LANES))
    reports = [load_json(path) for path in probe_paths]
    result = summarize_probe(reports, root=root, report_paths=probe_paths)
    result["normal_process_rss"] = {
        lane: parse_gnu_rss(
            root / lane / "resource.log",
            root=root,
            latency_comparison="descriptive_only",
        )
        for lane in (*PROBE_LANES, *FULL_LANES)
    }
    if full_paths is not None:
        if len(full_paths) != 2:
            raise AnalysisError("full guard requires control and candidate reports")
        full_reports = [load_json(path) for path in full_paths]
        policy_path = policy_path or root / "report-policy.json"
        guard = compare_guard(
            full_reports[0],
            full_reports[1],
            _load_policy(policy_path),
            root=root,
            report_paths=full_paths,
        )
        guard["inputs"]["policy"] = file_binding(policy_path, root)
        result["full_guard"] = guard
    if heap_paths is not None:
        if len(heap_paths) != 2:
            raise AnalysisError("heap evidence requires control and candidate resource logs")
        result["heap"] = {
            "purpose": "whole-process allocation diagnostic; heaptrack time and RSS are not latency evidence",
            "latency_claim": "none",
            "control": {
                "rss": parse_gnu_rss(heap_paths[0], root=root),
                "heaptrack": parse_heaptrack(
                    heap_paths[0].with_name("heaptrack-print.stdout"), root=root
                ),
            },
            "candidate": {
                "rss": parse_gnu_rss(heap_paths[1], root=root),
                "heaptrack": parse_heaptrack(
                    heap_paths[1].with_name("heaptrack-print.stdout"), root=root
                ),
            },
        }
    return result


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=ROOT)
    parser.add_argument("--a1", type=Path, default=None)
    parser.add_argument("--b1", type=Path, default=None)
    parser.add_argument("--b2", type=Path, default=None)
    parser.add_argument("--a2", type=Path, default=None)
    parser.add_argument("--full-control", type=Path)
    parser.add_argument("--full-candidate", type=Path)
    parser.add_argument("--heap-control", type=Path)
    parser.add_argument("--heap-candidate", type=Path)
    parser.add_argument("--policy", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    root = args.root.resolve()
    probe = [args.a1, args.b1, args.b2, args.a2]
    if any(path is None for path in probe):
        probe = _report_paths(root, PROBE_LANES)
    full = None
    if args.full_control is not None or args.full_candidate is not None:
        if args.full_control is None or args.full_candidate is None:
            parser.error("--full-control and --full-candidate must be provided together")
        full = [args.full_control, args.full_candidate]
    heap = None
    if args.heap_control is not None or args.heap_candidate is not None:
        if args.heap_control is None or args.heap_candidate is None:
            parser.error("--heap-control and --heap-candidate must be provided together")
        heap = [args.heap_control, args.heap_candidate]
    try:
        result = build_summary(
            root=root,
            probe_paths=probe,
            full_paths=full,
            policy_path=args.policy,
            heap_paths=heap,
        )
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(
            json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
        print(json.dumps({"output": str(args.output.resolve()), "status": "pass"}, sort_keys=True))
        return 0
    except (AnalysisError, OSError, ValueError, TypeError, KeyError) as error:
        print(json.dumps({"status": "fail", "error": str(error)}, sort_keys=True))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
