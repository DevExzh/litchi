#!/usr/bin/env python3
"""Render aggregate results from fresh process-local theme-family profiles."""

from __future__ import annotations

import argparse
import json
import re
from collections import defaultdict
from pathlib import Path


LANES = {(fixture, operation) for fixture in ("small", "opaque") for operation in ("read", "clone", "noop", "change")}
RUN_PATTERN = re.compile(r"^(small|opaque)-(read|clone|noop|change)-p([0-9]+)\.json$")


def percentile(values: list[int], percent: int) -> int:
    values = sorted(values)
    rank = (len(values) * percent + 99) // 100
    return values[max(0, min(len(values) - 1, rank - 1))]


def read_max_rss(path: Path) -> int | None:
    if not path.exists():
        return None
    for line in path.read_text().splitlines():
        if line.lstrip().startswith("Maximum resident set size (kbytes):"):
            return int(line.split(":", 1)[1].strip())
    return None


def require(value: bool, message: str) -> None:
    if not value:
        raise SystemExit(message)


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, default=Path(__file__).with_name("results"))
    parser.add_argument("--output", type=Path, default=Path(__file__).with_name("report.md"))
    parser.add_argument("--expected-processes", type=int, default=3)
    parser.add_argument("--samples-per-process", type=int, default=30)
    args = parser.parse_args()
    require(args.expected_processes > 0, "--expected-processes must be positive")
    require(args.samples_per_process > 0, "--samples-per-process must be positive")

    runs: dict[tuple[str, str], list[tuple[int, Path, dict[str, object]]]] = defaultdict(list)
    for path in sorted(args.results.glob("*-p*.json")):
        match = RUN_PATTERN.match(path.name)
        if not match:
            continue
        fixture, operation, process_text = match.groups()
        value = json.loads(path.read_text())
        process = int(process_text)
        require(value.get("schema") == "drawingml-theme-family-profile-v1", f"unexpected schema in {path}")
        require(value.get("fixture") == fixture and value.get("operation") == operation, f"lane mismatch in {path}")
        require(value.get("allocator_instrumented") is True, f"allocator instrumentation missing in {path}")
        require(value.get("allocator_self_test") is True, f"allocator self-test missing in {path}")
        require(value.get("peak_live_is_delta") is True, f"peak-live baseline metadata missing in {path}")
        require(value.get("fixture_shape_ok") is True, f"fixture shape gate failed in {path}")
        require(value.get("sample_count") == args.samples_per_process, f"sample count mismatch in {path}")
        require(len(value.get("samples", [])) == args.samples_per_process, f"raw sample count mismatch in {path}")
        runs[(fixture, operation)].append((process, path, value))

    require(set(runs) == LANES, f"profile lanes are incomplete: {sorted(set(runs))}")
    rows: list[dict[str, object]] = []
    for lane in sorted(LANES):
        fixture, operation = lane
        lane_runs = sorted(runs[lane])
        process_ids = [item[0] for item in lane_runs]
        require(process_ids == list(range(1, args.expected_processes + 1)), f"fresh process set mismatch for {lane}: {process_ids}")
        input_sizes = {int(item[2]["input_bytes"]) for item in lane_runs}
        input_hashes = {int(item[2]["input_hash"]) for item in lane_runs}
        require(len(input_sizes) == 1 and len(input_hashes) == 1, f"fixture changed between processes for {lane}")
        elapsed: list[int] = []
        allocated: list[int] = []
        peak: list[int] = []
        source_shared: list[bool] = []
        for _, path, value in lane_runs:
            require(value.get("semantic_ok_all") is True, f"semantic gate failed in {path}")
            require(value.get("opaque_preserved_all") is True, f"opaque gate failed in {path}")
            require(value.get("changed_ok_all") is True, f"changed gate failed in {path}")
            require(value.get("inverse_ok_all") is True, f"inverse gate failed in {path}")
            source_shared.append(bool(value["source_shared_all"]))
            for sample in value["samples"]:
                require(
                    sample["alloc_invalid"] is False
                    and sample["alloc_failed"] == 0
                    and sample["alloc_balance_ok"] is True,
                    f"allocator gate failed in {path}",
                )
                elapsed.append(int(sample["elapsed_ns"]))
                allocated.append(int(sample["allocated_bytes"]))
                peak.append(int(sample["peak_live_bytes"]))
        rss = [read_max_rss(args.results / f"{fixture}-{operation}-p{process}.time.txt") for process in process_ids]
        require(all(value is not None for value in rss), f"whole-process RSS is missing for {lane}")
        expected_sharing = operation in {"clone", "noop"}
        require(all(source_shared) == expected_sharing, f"source-sharing gate failed for {lane}")
        rss_values = [int(value) for value in rss if value is not None]
        rows.append(
            {
                "fixture": fixture,
                "operation": operation,
                "processes": len(lane_runs),
                "samples": len(elapsed),
                "bytes": next(iter(input_sizes)),
                "p50": percentile(elapsed, 50),
                "p95": percentile(elapsed, 95),
                "p99": percentile(elapsed, 99),
                "alloc50": percentile(allocated, 50),
                "alloc95": percentile(allocated, 95),
                "peak50": percentile(peak, 50),
                "peak95": percentile(peak, 95),
                "rss": f"{min(rss_values)}–{max(rss_values)} KiB",
                "share": all(source_shared),
            }
        )

    opaque_bytes = next(int(row["bytes"]) for row in rows if row["fixture"] == "opaque")
    process_run_count = sum(len(run_list) for run_list in runs.values())
    lines = [
        "# DrawingML `themeFamily` bounded profile",
        "",
        f"This report aggregates {args.expected_processes} fresh processes with {args.samples_per_process} measured samples",
        "per process for each lane. Elapsed time is measured with the harness’s",
        "process-local counting allocator installed, so it is allocator-instrumented",
        "operation time rather than an uninstrumented latency claim. Allocation",
        "counts and peak-live deltas come from the timed closure. A peak-live",
        "delta is above the closure’s live-before baseline; it is not total retained",
        "source/result memory. Maximum RSS is",
        "whole-process startup/warm-up/sample RSS from `/usr/bin/time`, not per-op",
        "memory.",
        "",
        "The evidence is absolute for the shared fragment owner on two deterministic",
        "synthetic inputs. It makes no before/after or whole-library speedup claim.",
        "Timers exclude fixture construction, source hashing, byte comparisons,",
        "pointer checks, and inverse checks; those checks run after the timer and",
        "allocation snapshot.",
        "",
        "| fixture | operation | fresh processes | samples | bytes | allocator-instrumented p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | max RSS per process | source share |",
        "|---|---:|---:|---:|---:|---:|---:|---:|---:|---|",
        ]
    for row in rows:
        lines.append(
            "| {fixture} | {operation} | {processes} | {samples} | {bytes} | "
            "{p50} / {p95} / {p99} | {alloc50} / {alloc95} | {peak50} / {peak95} | "
            "{rss} | {share} |".format(**row)
        )

    lines.extend(
        [
            "",
            f"All {process_run_count} process runs passed fixture-shape, semantic, opaque-preservation, changed-edit, inverse, and allocator-validity gates.",
            "Requested allocation bytes charge direct allocation sizes plus each successful realloc’s new size; raw JSON also records realloc old/new sizes and the exact live-balance gate.",
            "",
            "The `small` input is the extracted LibreOffice `themeFamily` empty",
            f"element. The `opaque` input is generated by the harness and is {opaque_bytes:,} bytes",
            "with 96 `a:ext` entries, comments, and a vendor attribute.",
            "Both inputs stay below the owner’s one MiB source bound. The change",
            "lane edits only the typed `name` field and checks the opaque marker and",
            "exact inverse outside the measured closure. The harness independently",
            "checks the opaque fragment’s family root, required attributes,",
            "family-namespace `extLst`, and DrawingML `a:ext` child namespace before",
            "measuring it; the vendor root attribute is intentionally outside strict",
            "XSD validation.",
            "",
            "A zero allocation count in the `clone` and `noop` lanes is the observed",
            "source-sharing result for these snapshots; it is not a general",
            "allocation guarantee for unrelated DrawingML owners. The opaque change",
            "lane scales with retained source size because it rebuilds the changed",
            "fragment while preserving unknown markup.",
            "",
            "The exact build/run commands, host/toolchain data, binary SHA-256,",
            "pre/post Cargo package source manifests, dirty theme hashes, and raw",
            "per-process JSON are retained beside this report. Optional `perf stat`",
            "output is retained only when the kernel permits hardware counters.",
        ]
    )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
