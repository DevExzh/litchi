#!/usr/bin/env python3
"""Adapt the record-0359a W5 analyzer's JSON stdout to ``summary.json``.

The historical analyzer intentionally prints its analysis to stdout; it does
not write a summary file.  The safe runner needs a named, fresh output for
its correctness receipt, so this adapter executes the real analyzer, parses
and checks its result, and writes that unchanged analysis as the qualified
summary payload.  It never replaces the analyzer's calculations.
"""

from __future__ import annotations

import argparse
import json
import math
from pathlib import Path
import subprocess
import sys
from typing import Any


SUMMARY_SCHEMA = "litchi.native-qualification.w5-analysis.v1"
REQUIRED_KEYS = frozenset({"source", "ops", "doc_batch1k_unified_over_leaf", "max_rss_kb"})
ROUTE_KEYS = frozenset({"scenario", "before_us", "after_us", "ratio"})


class AdapterError(RuntimeError):
    """The historical analysis did not produce a qualified result."""


def _finite_number(value: Any, label: str) -> None:
    if not isinstance(value, (int, float)) or isinstance(value, bool) or not math.isfinite(value):
        raise AdapterError(f"{label} must be a finite number")


def validate_analysis(value: Any) -> dict[str, Any]:
    """Validate the stable shape emitted by record-0359a ``analyze.py``."""

    if not isinstance(value, dict):
        raise AdapterError("analyzer stdout must be one JSON object")
    missing = sorted(REQUIRED_KEYS - value.keys())
    if missing:
        raise AdapterError(f"analyzer JSON is missing required keys: {missing}")
    for name in ("source", "ops"):
        rows = value[name]
        if not isinstance(rows, list) or not rows:
            raise AdapterError(f"analyzer {name} rows must be a non-empty list")
        for index, row in enumerate(rows):
            if not isinstance(row, dict) or not ROUTE_KEYS <= row.keys():
                raise AdapterError(f"analyzer {name}[{index}] is not a qualified route row")
            if not isinstance(row["scenario"], str) or not row["scenario"]:
                raise AdapterError(f"analyzer {name}[{index}].scenario must be non-empty")
            for key in ("before_us", "after_us", "ratio"):
                _finite_number(row[key], f"analyzer {name}[{index}].{key}")
    if not isinstance(value["doc_batch1k_unified_over_leaf"], dict):
        raise AdapterError("analyzer doc batch section must be an object")
    if not isinstance(value["max_rss_kb"], dict):
        raise AdapterError("analyzer max_rss_kb section must be an object")
    return value


def _parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("analyzer", type=Path)
    parser.add_argument("output_dir", type=Path)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = _parse_args(argv)
    analyzer = args.analyzer
    output_dir = args.output_dir
    if analyzer.is_symlink() or not analyzer.is_file():
        print(f"w5 analyzer is missing or symlinked: {analyzer}", file=sys.stderr)
        return 2
    if output_dir.is_symlink() or not output_dir.is_dir():
        print(f"w5 analyzer output directory is missing or symlinked: {output_dir}", file=sys.stderr)
        return 2
    summary = output_dir / "summary.json"
    if summary.is_symlink() or summary.exists():
        print(f"refusing to overwrite or follow summary: {summary}", file=sys.stderr)
        return 2
    try:
        result = subprocess.run(
            [sys.executable, str(analyzer), str(output_dir)],
            cwd=str(output_dir),
            check=False,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
        )
    except OSError as error:
        print(f"could not execute W5 analyzer: {error}", file=sys.stderr)
        return 2
    if result.stderr:
        sys.stderr.write(result.stderr)
    if result.returncode:
        if result.stdout:
            sys.stdout.write(result.stdout)
        return result.returncode
    try:
        parsed = validate_analysis(json.loads(result.stdout))
    except (AdapterError, json.JSONDecodeError) as error:
        print(f"W5 analyzer output is not a qualified JSON analysis: {error}", file=sys.stderr)
        return 2
    try:
        with summary.open("x", encoding="utf-8") as stream:
            json.dump(
                {"schema": SUMMARY_SCHEMA, "lane": "w5-legacy", "analysis": parsed},
                stream,
                indent=2,
                sort_keys=True,
            )
            stream.write("\n")
    except OSError as error:
        print(f"could not write W5 qualified summary: {error}", file=sys.stderr)
        return 2
    # Keep the analyzer's actual JSON visible in the runner's validator log;
    # the qualified file above is the parsed, schema-checked copy.
    if result.stdout:
        sys.stdout.write(result.stdout)
        if not result.stdout.endswith("\n"):
            sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
