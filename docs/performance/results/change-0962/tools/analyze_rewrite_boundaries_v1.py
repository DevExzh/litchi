#!/usr/bin/env python3
"""Offline boundary attribution for selected-versus-rewritten source paths.

The helper reads only retained ``perf script`` text.  It imports the frozen
native CPU analyzer for parsing and qualification, then partitions each
qualified physical callchain once by exact full symbol identity.  It never
starts ``perf``, a workload, or a build.
"""

from __future__ import annotations

import argparse
import concurrent.futures
import datetime
import hashlib
import json
import re
import sys
from pathlib import Path
from typing import Any

sys.path.insert(0, str(Path(__file__).parent))
import analyze_native_cpu_v1 as frozen  # noqa: E402


ROOT = Path("/home/zhuhe/litchi-performance-0962")
DEFAULT_WAVE_ROOT = ROOT / "fp-wave-v1"
FROZEN_ANALYZER_SHA256 = "f5f7a387f99ba3a587eeb8f27c8f422a9fa696306c2ae47724952b861910bd3d"
EXPECTED_SELECTED_SAMPLES = 12319
SOURCE_SYMBOL = "litchi_xlsx::cell_values::snapshot::Snapshot::from_source_selected"
REWRITTEN_SYMBOL = "litchi_xlsx::cell_values::snapshot::Snapshot::from_rewritten_value_source"
VALIDATE_XML_SYMBOL = "litchi_xlsx::cell_values::validation::validate_xml"
SCHEMA = "litchi.performance.0962.rewrite-boundaries.v1"


class BoundaryError(RuntimeError):
    """A qualification, boundary, or conservation failure."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise BoundaryError(message)


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    require(path.is_absolute() and path.is_file() and not path.is_symlink(), f"invalid input file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_file(path)}


def write_create(path: Path, value: dict[str, Any]) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to overwrite {path}")
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def stats(samples: int, period: int, denominator: int) -> dict[str, Any]:
    return {
        "samples": samples,
        "period": period,
        "period_share": period / denominator if denominator else 0.0,
    }


def add_stats(target: dict[str, int], sample: frozen.Sample) -> None:
    target["samples"] += 1
    target["period"] += sample.period


def empty_stats() -> dict[str, int]:
    return {"samples": 0, "period": 0}


def classify_boundary(physical_sample: frozen.Sample) -> tuple[str, str, bool, list[str]]:
    """Return a disjoint exact boundary class and nested validate evidence."""
    symbols = [frozen.function_symbol(frame.symbol) for frame in physical_sample.frames]
    source_indices = [index for index, symbol in enumerate(symbols) if symbol == SOURCE_SYMBOL]
    rewritten_indices = [index for index, symbol in enumerate(symbols) if symbol == REWRITTEN_SYMBOL]
    validate_indices = [index for index, symbol in enumerate(symbols) if symbol == VALIDATE_XML_SYMBOL]
    unknown = any(not frame.known for frame in physical_sample.frames)
    both = bool(source_indices and rewritten_indices)
    if unknown:
        category = "ambiguous_unknown"
    elif both:
        category = "ambiguous_source_and_rewritten"
    elif source_indices:
        category = "source_selected"
    elif rewritten_indices:
        category = "rewritten_value_source"
    else:
        category = "neither"
    nested = bool(
        rewritten_indices
        and any(validate < rewritten for validate in validate_indices for rewritten in rewritten_indices)
    )
    return category, "physical_stack_exact_symbol_boundary", nested, symbols


def analyze_child(index: int, case_dir: Path, binary_dso: str) -> dict[str, Any]:
    inline_path = case_dir / "offline-decode-v1" / "perf-script-inline.txt"
    physical_path = case_dir / "offline-decode-v2" / "perf-script-physical.txt"
    analyzer_output_path = case_dir / "offline-decode-v2" / "analyzer.json"
    output_path = case_dir / "offline-decode-v2" / "boundary-attribution-v1.json"
    with inline_path.open("r", encoding="utf-8", errors="strict") as stream:
        inline_samples, inline_diag = frozen.parse_samples(stream, frozen.DEFAULT_EVENT)
    with physical_path.open("r", encoding="utf-8", errors="strict") as stream:
        physical_samples, physical_diag = frozen.parse_samples(stream, frozen.DEFAULT_EVENT)
    prior = json.loads(analyzer_output_path.read_text(encoding="utf-8"))
    prior_qualified = prior["denominators"]["qualified_ancestor"]
    pairs, unmatched_inline, unmatched_physical, _ = frozen.pair_samples(inline_samples, physical_samples)
    require(not unmatched_inline and not unmatched_physical, f"child {index} pair mismatch")
    require(len(pairs) == len(inline_samples) == len(physical_samples), f"child {index} pair count mismatch")
    require(prior["pairing"]["exact_pairing_pass"] is True, f"child {index} prior pairing gate failed")

    partition: dict[str, dict[str, int]] = {}
    selected = empty_stats()
    nested = empty_stats()
    validate_any = empty_stats()
    physical_self: dict[str, dict[str, int]] = {}
    classification_counts: dict[str, dict[str, int]] = {}
    qualified_pairs = 0
    for inline_sample, physical_sample in pairs:
        classification, _, _ = frozen.classify_pair(
            inline_sample,
            physical_sample,
            source_suffix=frozen.DEFAULT_SOURCE_SUFFIX,
            stage_range=frozen.DEFAULT_STAGE_RANGE,
            timed_range=frozen.DEFAULT_TIMED_RANGE,
            physical_prefix=frozen.DEFAULT_PHYSICAL_TIMED_PREFIX,
            binary_dso=binary_dso,
        )
        classification_counts.setdefault(classification, {"samples": 0, "period": 0})
        classification_counts[classification]["samples"] += 1
        classification_counts[classification]["period"] += inline_sample.period
        if classification != "qualified":
            continue
        qualified_pairs += 1
        add_stats(selected, inline_sample)
        category, _, is_nested, symbols = classify_boundary(physical_sample)
        add_stats(partition.setdefault(category, empty_stats()), inline_sample)
        if is_nested:
            add_stats(nested, inline_sample)
        if VALIDATE_XML_SYMBOL in symbols:
            add_stats(validate_any, inline_sample)
        leaf = frozen.normalized_physical_leaf(physical_sample.frames[0])
        add_stats(physical_self.setdefault(leaf, empty_stats()), inline_sample)

    require(selected == {"samples": prior_qualified["samples"], "period": prior_qualified["period"]}, f"child {index} qualification drift")
    # The six scheduled children have intentionally different sample counts;
    # the prior analyzer's qualified denominator is the per-child reference.
    require(qualified_pairs == prior_qualified["samples"], f"child {index} selected sample drift")
    # Every boundary partition is required to conserve the child's selected
    # population exactly.
    partition_samples = sum(value["samples"] for value in partition.values())
    partition_period = sum(value["period"] for value in partition.values())
    require(partition_samples == selected["samples"], f"child {index} boundary sample conservation failed")
    require(partition_period == selected["period"], f"child {index} boundary period conservation failed")
    require(nested["samples"] <= partition.get("rewritten_value_source", empty_stats())["samples"], f"child {index} nested sample subset failed")
    require(nested["period"] <= partition.get("rewritten_value_source", empty_stats())["period"], f"child {index} nested period subset failed")
    rows = {
        key: {**value, "period_share": value["period"] / selected["period"] if selected["period"] else 0.0}
        for key, value in sorted(partition.items())
    }
    whole_period = sum(sample.period for sample in inline_samples)
    result = {
        "schema": SCHEMA,
        "status": "passed",
        "index": index,
        "case_directory": str(case_dir),
        "inputs": {"inline": descriptor(inline_path), "physical": descriptor(physical_path), "prior_analyzer": descriptor(analyzer_output_path)},
        "frozen_analyzer_sha256": FROZEN_ANALYZER_SHA256,
        "physical_binary_dso": binary_dso,
        "selected_qualified": stats(selected["samples"], selected["period"], selected["period"]),
        "boundary_partition": rows,
        "partition_conservation": {
            "selected_samples": selected["samples"],
            "selected_period": selected["period"],
            "partition_samples": partition_samples,
            "partition_period": partition_period,
            "sample_sum_pass": partition_samples == selected["samples"],
            "period_sum_pass": partition_period == selected["period"],
        },
        "validate_xml": {
            "exact_symbol": VALIDATE_XML_SYMBOL,
            "any_selected": stats(validate_any["samples"], validate_any["period"], selected["period"]),
            "nested_under_rewritten": stats(nested["samples"], nested["period"], selected["period"]),
            "nested_subset_pass": nested["samples"] <= partition.get("rewritten_value_source", empty_stats())["samples"] and nested["period"] <= partition.get("rewritten_value_source", empty_stats())["period"],
        },
        "physical_self": {
            "scope": "deepest physical frame only; sampled period is charged once and callee frames are excluded",
            "groups": [
                {"key": key, **value, "period_share": value["period"] / selected["period"] if selected["period"] else 0.0}
                for key, value in sorted(physical_self.items(), key=lambda item: (-item[1]["period"], item[0]))
            ],
        },
        "qualification_replay": {
            "classification": {
                key: {**value, "period_share": value["period"] / whole_period if whole_period else 0.0}
                for key, value in sorted(classification_counts.items())
            },
            "inline_parser": inline_diag,
            "physical_parser": physical_diag,
        },
        "interpretation": "Boundary rows are disjoint presence classes over qualified physical stacks; nested validate_xml is reported as a subset and no frame-level callee periods are added.",
    }
    write_create(output_path, result)
    return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--wave-root", type=Path, default=DEFAULT_WAVE_ROOT)
    parser.add_argument("--index", type=int, action="append")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        analyzer_path = Path(frozen.__file__).resolve().with_suffix(".py")
        require(sha256_file(analyzer_path) == FROZEN_ANALYZER_SHA256, "frozen analyzer hash changed")
        wave_root = args.wave_root.expanduser().absolute()
        indexes = sorted(set(args.index if args.index is not None else range(6)))
        require(indexes and all(0 <= index < 6 for index in indexes), "index must be 0 through 5")
        case_dirs = sorted(wave_root.glob("[0-9][0-9]-*"))
        require(len(case_dirs) == 6, "expected six wave child directories")
        binary_dso = "/home/zhuhe/litchi-performance-0962/executables/xlsx_managed_compression_benchmark.fp-debug2.v1"
        with concurrent.futures.ThreadPoolExecutor(max_workers=len(indexes)) as pool:
            futures = [pool.submit(analyze_child, index, case_dirs[index], binary_dso) for index in indexes]
            results = [future.result() for future in futures]
        if len(indexes) == 6:
            total_samples = sum(result["selected_qualified"]["samples"] for result in results)
            total_period = sum(result["selected_qualified"]["period"] for result in results)
            require(total_samples == EXPECTED_SELECTED_SAMPLES, f"wave selected sample count {total_samples} != {EXPECTED_SELECTED_SAMPLES}")
            aggregate = {
                "schema": SCHEMA + ".receipt",
                "status": "passed",
                "generated_utc": now(),
                "frozen_analyzer_sha256": FROZEN_ANALYZER_SHA256,
                "expected_selected_samples": EXPECTED_SELECTED_SAMPLES,
                "selected_samples": total_samples,
                "selected_period": total_period,
                "children": [result["index"] for result in sorted(results, key=lambda item: item["index"])],
                "outputs": [str(case_dirs[index] / "offline-decode-v2" / "boundary-attribution-v1.json") for index in indexes],
                "boundary_partition_is_disjoint": True,
                "physical_self_excludes_callee_costs": True,
            }
            write_create(wave_root / "boundary-attribution-receipt-v1.json", aggregate)
        print(json.dumps({"schema": SCHEMA, "status": "passed", "children": indexes}, sort_keys=True))
        return 0
    except (BoundaryError, OSError, UnicodeError, ValueError, json.JSONDecodeError) as exc:
        print(json.dumps({"schema": SCHEMA, "status": "failed", "error": str(exc)}, sort_keys=True), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
