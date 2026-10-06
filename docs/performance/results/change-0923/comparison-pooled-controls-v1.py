#!/usr/bin/env python3
"""Negative control for raw-sample pooling.

The fixture has two child vectors with different sizes.  Pooling child
medians produces a different p50 from pooling the raw values, so silently
substituting child summaries is rejected.
"""

from __future__ import annotations

import argparse
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Iterable


def nearest_rank(values: list[int], fraction: float) -> int:
    ordered = sorted(values)
    index = min(max(int((len(ordered) * fraction) + 0.999999999) - 1, 0), len(ordered) - 1)
    return ordered[index]


def run_control() -> dict[str, object]:
    child_vectors = [[1, 1], [100, 100, 100, 100]]
    raw_values = [value for child in child_vectors for value in child]
    child_medians = [nearest_rank(child, 0.50) for child in child_vectors]
    raw_p50 = nearest_rank(raw_values, 0.50)
    summary_p50 = nearest_rank(child_medians, 0.50)
    if raw_p50 == summary_p50:
        raise RuntimeError("negative control did not distinguish raw pooling from child-summary pooling")
    return {
        "schema": "litchi.performance.0923.raw-pooled-controls.v1",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "fixture": {"child_vectors": child_vectors, "raw_sample_count": len(raw_values)},
        "expected_raw_p50": raw_p50,
        "naive_child_summary_p50": summary_p50,
        "negative_control": {
            "name": "child_summary_substitution",
            "status": "rejected",
            "reason": "raw p50 differs from p50 of child medians; raw pooling must retain every sample",
        },
    }


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(list(argv) if argv is not None else None)
    value = run_control()
    data = json.dumps(value, indent=2, sort_keys=True) + "\n"
    if args.output:
        output = args.output.resolve()
        if output.exists():
            raise RuntimeError(f"refusing to replace retained receipt: {output}")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(data)
    else:
        print(data, end="")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
