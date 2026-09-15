#!/usr/bin/env python3
"""Analysis for change 0620: XLS edit-and-save attribution and the reuse delta.

Three independent readers, each over the raw outputs retained beside this file:

* ``callgrind``  differences the small/large ``callgrind_annotate --inclusive``
  pairs and divides by the extra operations, giving per-operation inclusive Ir
  for every function *and caller chain* (``--separate-callers=2``), which is
  what makes the three complete ``Workbook::new`` parses of a source-backed
  commit separable from one another.
* ``perf``       does the same differencing over the native ``perf stat`` CSVs.
* ``latency``    reads the probe's per-sample phase vectors and reports p50,
  mean, p95 and p99 per round, the paired A1/B1 and B2/A2 deltas, and the
  same-binary A1/A2 and B1/B2 floors.

Usage:
    analyze.py callgrind <dir> <leg> [<leg> ...]
    analyze.py perf      <dir> <leg> [<leg> ...]
    analyze.py latency   <dir>
    analyze.py counters  <dir> <leg> [<leg> ...]
"""

from __future__ import annotations

import json
import re
import statistics
import sys
from pathlib import Path

ANNOTATION = re.compile(r"^\s*([\d,]+) \(([\d.]+)%\)\s+(.*)$")


def read_annotation(path: Path) -> dict[str, int]:
    """Returns {function-with-caller-chain: inclusive Ir} for one annotation."""
    totals: dict[str, int] = {}
    for line in path.read_text(errors="replace").splitlines():
        if line.startswith(" ") and "=>" in line:
            continue  # a call-edge line, not a function total
        match = ANNOTATION.match(line)
        if not match:
            continue
        name = match.group(3).strip()
        if name.startswith("=>"):
            continue
        # Strip the binary location suffix and the source file prefix; keep the
        # function plus its caller chain, which is the part we attribute by.
        name = re.sub(r"\s+\[[^\]]*\]$", "", name)
        if ":" in name:
            name = name.split(":", 1)[1]
        # callgrind_annotate prints one inclusive line per (source file,
        # function): a function whose code was inlined across several files
        # appears several times, the outermost line carrying the whole
        # inclusive cost and the others its parts. Take the maximum, never the
        # sum, or a function that spans two files is counted twice.
        value = int(match.group(1).replace(",", ""))
        totals[name] = max(totals.get(name, 0), value)
    return totals


def shorten(name: str) -> str:
    name = name.replace("core::io::cursor::Cursor<alloc::vec::Vec<u8>>", "Cursor<Vec<u8>>")
    name = name.replace("core::io::cursor::Cursor<&[u8]>", "Cursor<&[u8]>")
    name = re.sub(r"<litchi_xls::workbook::model::Workbook<[^>]*>>", "Workbook", name)
    name = name.replace("litchi_xls::cell_values::", "")
    name = name.replace("<litchi_xls::cell_values::", "<")
    name = name.replace("litchi_xls::", "")
    name = name.replace("litchi_cfb::", "cfb::")
    name = name.replace("litchi_ole_common::", "ole::")
    name = name.replace("<litchi_ole_common::", "<ole::")
    name = name.replace("alloc::vec::Vec<u8>", "Vec<u8>")
    return name


def callgrind(directory: Path, legs: list[str]) -> None:
    for leg in legs:
        pairs = (directory / f"pairs-{leg}.txt").read_text().split("\n")
        result: dict[str, dict[str, int]] = {}
        for entry in pairs:
            if not entry.strip():
                continue
            stem, op, small, large = entry.split()
            extra = int(large) - int(small)
            low = read_annotation(directory / f"ann-{leg}-{stem}-{op}-small.txt")
            high = read_annotation(directory / f"ann-{leg}-{stem}-{op}-large.txt")
            per_op = {}
            for name in set(low) | set(high):
                delta = high.get(name, 0) - low.get(name, 0)
                if delta > 0:
                    per_op[shorten(name)] = round(delta / extra)
            result[f"{stem}/{op}"] = per_op
        out = directory / f"callgrind-{leg}.json"
        out.write_text(json.dumps(result, indent=1, sort_keys=True))
        print(f"wrote {out}")


def perf(directory: Path, legs: list[str]) -> None:
    for leg in legs:
        pairs = (directory / f"perf-pairs-{leg}.txt").read_text().split("\n")
        result: dict[str, dict[str, float]] = {}
        for entry in pairs:
            if not entry.strip():
                continue
            stem, op, small, large = entry.split()
            extra = int(large) - int(small)
            counts = {}
            for size, samples in (("small", small), ("large", large)):
                path = directory / f"perf-{leg}-{stem}-{op}-s{samples}.csv"
                values: dict[str, float] = {}
                for line in path.read_text().splitlines():
                    if line.startswith("#") or not line.strip():
                        continue
                    fields = line.split(",")
                    if len(fields) < 3:
                        continue
                    try:
                        values[fields[2]] = float(fields[0])
                    except ValueError:
                        continue
                counts[size] = values
            per_op = {}
            for event in counts["large"]:
                delta = counts["large"][event] - counts["small"].get(event, 0.0)
                per_op[event] = round(delta / extra, 2)
            result[f"{stem}/{op}"] = per_op
        out = directory / f"perf-{leg}.json"
        out.write_text(json.dumps(result, indent=1, sort_keys=True))
        print(f"wrote {out}")


def summarize(values: list[int]) -> dict[str, float]:
    values = sorted(values)
    if not values:
        return {}

    def percentile(fraction: float) -> float:
        index = min(len(values) - 1, int(fraction * (len(values) - 1) + 0.5))
        return float(values[index])

    return {
        "samples": len(values),
        "p50": float(statistics.median(values)),
        "mean": round(statistics.fmean(values), 1),
        "p95": percentile(0.95),
        "p99": percentile(0.99),
    }


def latency(directory: Path) -> None:
    rounds = ["a1", "b1", "b2", "a2"]
    table: dict[str, dict[str, dict[str, dict[str, float]]]] = {}
    for round_name in rounds:
        for path in sorted((directory / round_name).glob("*/*.json")):
            stem = path.parent.name
            op = path.stem
            payload = json.loads(path.read_text())
            key = f"{stem}/{op}"
            entry = table.setdefault(key, {})
            for phase in ("open_ns", "stage_ns", "commit_ns", "publish_ns"):
                if not payload.get(phase):
                    continue
                entry.setdefault(phase, {})[round_name] = summarize(payload[phase])

    report: dict[str, dict[str, dict[str, object]]] = {}
    for key, phases in sorted(table.items()):
        for phase, per_round in phases.items():
            if not {"a1", "a2", "b1", "b2"} <= set(per_round):
                continue
            row: dict[str, object] = {r: per_round[r] for r in rounds}
            def ratio(new: float, base: float) -> float | None:
                return None if base == 0 else round((new - base) / base * 100, 2)

            for statistic in ("p50", "mean", "p95", "p99"):
                a1, b1 = per_round["a1"][statistic], per_round["b1"][statistic]
                b2, a2 = per_round["b2"][statistic], per_round["a2"][statistic]
                row[f"delta_forward_{statistic}_pct"] = ratio(b1, a1)
                row[f"delta_reverse_{statistic}_pct"] = ratio(b2, a2)
                row[f"floor_a_{statistic}_pct"] = ratio(a2, a1)
                row[f"floor_b_{statistic}_pct"] = ratio(b2, b1)
            report.setdefault(key, {})[phase] = row
    out = directory / "latency-summary.json"
    out.write_text(json.dumps(report, indent=1, sort_keys=True))
    print(f"wrote {out}")


def counters(directory: Path, legs: list[str]) -> None:
    for leg in legs:
        rows = []
        for line in (directory / f"counters-{leg}.jsonl").read_text().splitlines():
            if line.strip():
                rows.append(json.loads(line))
        out = directory / f"counters-{leg}.json"
        out.write_text(json.dumps(rows, indent=1))
        print(f"wrote {out}")


if __name__ == "__main__":
    mode = sys.argv[1]
    where = Path(sys.argv[2])
    if mode == "callgrind":
        callgrind(where, sys.argv[3:])
    elif mode == "perf":
        perf(where, sys.argv[3:])
    elif mode == "latency":
        latency(where)
    elif mode == "counters":
        counters(where, sys.argv[3:])
    else:
        raise SystemExit(f"unknown mode {mode}")
