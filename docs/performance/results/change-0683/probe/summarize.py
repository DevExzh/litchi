#!/usr/bin/env python3
"""Summarize retained 0683 timing and allocator rows.

This is intentionally a small standard-library reader. It reports only the
observed samples and semantic agreement; it does not infer a performance
claim or normalize a refusal into a success.
"""

from __future__ import annotations

import pathlib
import statistics
import sys


def key_values(path: pathlib.Path) -> dict[str, str]:
    fields: dict[str, str] = {}
    with path.open() as stream:
        for field in stream.read().strip().split("\t"):
            key, value = field.split("=", 1)
            fields[key] = value
    return fields


def p50(values: list[int]) -> int:
    return int(statistics.median(values))


def percentile(values: list[int], percent: int) -> int:
    ordered = sorted(values)
    index = max(0, (len(ordered) * percent + 99) // 100 - 1)
    return ordered[index]


def timing_rows(path: pathlib.Path) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    with path.open() as stream:
        for line in stream:
            if not line.strip():
                continue
            rows.append(
                {
                    key: value
                    for key, value in (
                        field.split("=", 1) for field in line.strip().split("\t")
                    )
                }
            )
    return rows


def timing(path: pathlib.Path) -> None:
    rows = timing_rows(path)
    if len({(row["result"], row["callbacks"]) for row in rows}) != 1:
        raise SystemExit(f"mixed semantics in {path}")
    if [int(row["sample"]) for row in rows] != list(range(len(rows))):
        raise SystemExit(f"missing/reordered samples in {path}")
    values = [int(row["nanos"]) for row in rows]
    results = {row["result"] for row in rows}
    callbacks = {row["callbacks"] for row in rows}
    if not values:
        raise SystemExit(f"no timing samples in {path}")
    print(
        f"timing\t{path.name}\tsamples={len(values)}\tp50_ns={p50(values)}"
        f"\tmean_ns={int(statistics.mean(values))}"
        f"\tp95_ns={percentile(values, 95)}\tp99_ns={percentile(values, 99)}"
        f"\tcallbacks={','.join(sorted(callbacks))}\tresults={','.join(sorted(results))}"
    )


def allocation(path: pathlib.Path) -> None:
    fields = key_values(path)
    required = (
        "operation",
        "callbacks",
        "allocation_calls",
        "allocated_bytes",
        "peak_live_delta",
        "retained_live_delta",
    )
    missing = [field for field in required if field not in fields]
    if missing:
        raise SystemExit(f"{path}: missing {missing}")
    print(
        "allocation\t{}\toperation={}\tcallbacks={}\tallocation_calls={}"
        "\tallocated_bytes={}\tpeak_live_delta={}\tretained_live_delta={}"
        .format(
            path.name,
            fields["operation"],
            fields["callbacks"],
            fields["allocation_calls"],
            fields["allocated_bytes"],
            fields["peak_live_delta"],
            fields["retained_live_delta"],
        )
    )


def timing_key(path: pathlib.Path) -> tuple[str, str, str] | None:
    operations = (
        "visit-cold",
        "cells-cold",
        "visit-selected",
        "cells-selected",
        "visit-warm",
        "cells-warm",
    )
    labels = ("a1", "a2", "b1", "b2")
    for label in labels:
        prefix = f"{label}-"
        if not path.name.startswith(prefix):
            continue
        body = path.stem[len(prefix) :]
        for operation in operations:
            suffix = f"-{operation}"
            if body.endswith(suffix):
                return body[: -len(suffix)], operation, label
    return None


def paired(directory: pathlib.Path) -> None:
    groups: dict[tuple[str, str], dict[str, list[int]]] = {}
    for path in directory.glob("*.tsv"):
        key = timing_key(path)
        if key is None:
            continue
        case, operation, label = key
        rows = timing_rows(path)
        groups.setdefault((case, operation), {})[label] = [
            int(row["nanos"])
            for row in rows
        ]
    required = {"a1", "a2", "b1", "b2"} if any("b1" in legs or "b2" in legs for legs in groups.values()) else {"a1", "a2"}
    for (case, operation), legs in sorted(groups.items()):
        if set(legs) != required or len({len(v) for v in legs.values()}) != 1:
            raise SystemExit(f"incomplete group: {case} {operation}")
        if required == {"a1", "a2"}:
            print(f"aa\t{case}\t{operation}\tdrift_pct={(p50(legs["a2"])/p50(legs["a1"])-1)*100:.3f}")
            continue
        medians = {label: p50(values) for label, values in legs.items()}
        aa = abs(medians["a2"] - medians["a1"]) * 100 / medians["a1"]
        b1 = (medians["b1"] - medians["a1"]) * 100 / medians["a1"]
        b2 = (medians["b2"] - medians["a2"]) * 100 / medians["a2"]
        print(
            f"pair\t{case}\toperation={operation}\taa_abs_pct={aa:.3f}"
            f"\tb1_vs_a1_pct={b1:.3f}\tb2_vs_a2_pct={b2:.3f}"
        )


def main() -> int:
    if len(sys.argv) != 2:
        print("usage: summarize.py DIRECTORY", file=sys.stderr)
        return 2
    directory = pathlib.Path(sys.argv[1])
    for path in sorted(directory.glob("*.tsv")):
        with path.open() as stream:
            first = stream.readline()
        if first.startswith("sample="):
            timing(path)
        elif first.startswith("operation="):
            allocation(path)
    paired(directory)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
