#!/usr/bin/env python3
"""Aggregate the process-isolated XLSB theme-family profile."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from collections import defaultdict
from pathlib import Path


FIXTURES = ("native", "opaque")
OPERATIONS = (
    "codec_read",
    "metadata_read",
    "source_read",
    "family_clone",
    "noop",
    "add",
    "update",
    "remove",
    "base_edit",
)
LANES = {(fixture, operation) for fixture in FIXTURES for operation in OPERATIONS}
RUN_PATTERN = re.compile(
    r"^(native|opaque)-(codec_read|metadata_read|source_read|family_clone|noop|add|update|remove|base_edit)-p([0-9]+)\.json$"
)


def percentile(values: list[int], percent: int) -> int:
    ordered = sorted(values)
    rank = (len(ordered) * percent + 99) // 100
    return ordered[max(0, min(len(ordered) - 1, rank - 1))]


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


def sample_accounting(sample: dict[str, object], path: Path) -> None:
    direct = int(sample["direct_allocated_bytes"])
    realloc_new = int(sample["realloc_new_bytes"])
    realloc_old = int(sample["realloc_old_bytes"])
    deallocated = int(sample["deallocated_bytes"])
    live_before = int(sample["live_before"])
    live_after = int(sample["live_after"])
    expected_live = live_before + direct + realloc_new - realloc_old - deallocated
    require(
        int(sample["requested_alloc_bytes"]) == direct + realloc_new,
        f"requested allocation accounting failed in {path}",
    )
    require(live_after == expected_live, f"live balance failed in {path}")
    require(
        int(sample["peak_live_delta"]) >= max(0, live_after - live_before),
        f"peak-live delta is below live growth in {path}",
    )
    require(sample.get("alloc_balance_ok") is True, f"allocator balance gate failed in {path}")
    require(sample.get("alloc_invalid") is False, f"invalid allocator observation in {path}")
    require(int(sample.get("alloc_failed", 0)) == 0, f"allocator failure in {path}")
    require(sample.get("forward_preservation_ok") is True, f"forward preservation gate failed in {path}")


def lane_summary(
    fixture: str,
    operation: str,
    lane_runs: list[tuple[int, Path, dict[str, object]]],
    results: Path,
    expected_processes: int,
    samples_per_process: int,
) -> dict[str, object]:
    process_ids = [item[0] for item in sorted(lane_runs)]
    require(
        process_ids == list(range(1, expected_processes + 1)),
        f"fresh process set mismatch for {(fixture, operation)}: {process_ids}",
    )
    first = lane_runs[0][2]
    for _, path, value in lane_runs:
        require(value.get("schema") == "xlsb-theme-family-profile-v1", f"unexpected schema in {path}")
        require(
            value.get("fixture") == fixture and value.get("operation") == operation,
            f"lane mismatch in {path}",
        )
        require(value.get("allocator_instrumented") is True, f"allocator instrumentation missing in {path}")
        require(value.get("allocator_self_test") is True, f"allocator self-test missing in {path}")
        require(value.get("peak_live_is_incremental_delta") is True, f"peak-live label missing in {path}")
        require(value.get("fixture_shape_ok") is True, f"fixture-shape gate failed in {path}")
        expected_foreign_namespaces = 96 if fixture == "opaque" else 0
        require(
            int(value.get("foreign_namespace_declarations", -1)) == expected_foreign_namespaces,
            f"foreign namespace stress count mismatch in {path}",
        )
        require(value.get("sample_count") == samples_per_process, f"sample count mismatch in {path}")
        samples = value.get("samples")
        require(isinstance(samples, list) and len(samples) == samples_per_process, f"raw sample count mismatch in {path}")
        for key in (
            "semantic_ok_all",
            "forward_preservation_ok_all",
            "preservation_ok_all",
            "inverse_ok_all",
            "changed_ok_all",
            "allocation_balance_all",
            "source_observation_stable",
            "source_sharing_gate",
        ):
            require(value.get(key) is True, f"{key} failed in {path}")
        for sample in samples:
            require(isinstance(sample, dict), f"invalid raw sample in {path}")
            sample_accounting(sample, path)
        elapsed_one = [int(sample["elapsed_ns"]) for sample in samples]
        allocated_one = [int(sample["requested_alloc_bytes"]) for sample in samples]
        peak_one = [int(sample["peak_live_delta"]) for sample in samples]
        for percent in (50, 95, 99):
            require(
                int(value[f"p{percent}_ns"]) == percentile(elapsed_one, percent),
                f"per-process time percentile failed in {path}",
            )
        for prefix, values in (("requested_alloc", allocated_one), ("peak_live_delta", peak_one)):
            for percent in (50, 95):
                require(
                    int(value[f"{prefix}_p{percent}"]) == percentile(values, percent),
                    f"per-process {prefix} percentile failed in {path}",
                )

    fields = ("package_bytes", "theme_bytes", "family_bytes", "package_digest", "theme_digest", "family_digest")
    for field in fields:
        require(
            len({int(value[field]) for _, _, value in lane_runs}) == 1,
            f"fixture field changed across processes for {(fixture, operation)}: {field}",
        )

    source_applicable = operation in {"family_clone", "noop"}
    source_shared = [bool(sample["source_shared"]) for _, _, value in lane_runs for sample in value["samples"]]
    require(
        all(source_shared) if source_applicable else not source_applicable,
        f"source-sharing applicability mismatch in {(fixture, operation)}",
    )
    if operation == "source_read":
        read_tuples = {
            (int(sample["read_calls"]), int(sample["read_requested_bytes"]), int(sample["read_returned_bytes"]))
            for _, _, value in lane_runs
            for sample in value["samples"]
        }
        read_tuple = next(iter(read_tuples)) if len(read_tuples) == 1 else (0, 0, 0)
        require(
            len(read_tuples) == 1
            and read_tuple[0] > 0
            and read_tuple[1] >= read_tuple[2] > 0,
            f"source I/O observation failed in {(fixture, operation)}",
        )
    else:
        read_tuples = {
            (int(sample["read_calls"]), int(sample["read_requested_bytes"]), int(sample["read_returned_bytes"]))
            for _, _, value in lane_runs
            for sample in value["samples"]
        }
        require(read_tuples == {(0, 0, 0)}, f"unexpected source I/O in {(fixture, operation)}")
        read_tuple = (0, 0, 0)

    elapsed = [int(sample["elapsed_ns"]) for _, _, value in lane_runs for sample in value["samples"]]
    allocated = [int(sample["requested_alloc_bytes"]) for _, _, value in lane_runs for sample in value["samples"]]
    peak = [int(sample["peak_live_delta"]) for _, _, value in lane_runs for sample in value["samples"]]
    rss = [read_max_rss(results / f"{fixture}-{operation}-p{process}.time.txt") for process in process_ids]
    require(all(value is not None for value in rss), f"whole-process RSS is missing for {(fixture, operation)}")
    rss_values = [int(value) for value in rss if value is not None]
    return {
        "fixture": fixture,
        "operation": operation,
        "processes": len(lane_runs),
        "samples": len(elapsed),
        "bytes": int(first["package_bytes"]),
        "theme_bytes": int(first["theme_bytes"]),
        "family_bytes": int(first["family_bytes"]),
        "p50": percentile(elapsed, 50),
        "p95": percentile(elapsed, 95),
        "p99": percentile(elapsed, 99),
        "alloc50": percentile(allocated, 50),
        "alloc95": percentile(allocated, 95),
        "peak50": percentile(peak, 50),
        "peak95": percentile(peak, 95),
        "rss": f"{min(rss_values)}–{max(rss_values)} KiB",
        "share": "yes" if source_applicable else "n/a",
        "read_io": f"{read_tuple[0]} / {read_tuple[1]} / {read_tuple[2]}",
    }


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
        if match is None:
            continue
        fixture, operation, process_text = match.groups()
        runs[(fixture, operation)].append((int(process_text), path, json.loads(path.read_text())))
    require(set(runs) == LANES, f"profile lanes are incomplete: {sorted(set(runs))}")

    rows = [
        lane_summary(fixture, operation, sorted(runs[(fixture, operation)]), args.results, args.expected_processes, args.samples_per_process)
        for fixture in FIXTURES
        for operation in OPERATIONS
    ]
    manifest_path = args.results / "source-manifest-before.txt"
    manifest_sha = hashlib.sha256(manifest_path.read_bytes()).hexdigest() if manifest_path.exists() else "unavailable"
    binary_sha = "unavailable"
    binary_sha_path = args.results / "binary.sha256"
    if binary_sha_path.exists():
        binary_sha = binary_sha_path.read_text().split()[0]
    lines = [
        "# XLSB `themeFamily` host profile",
        "",
        f"This report aggregates {args.expected_processes} fresh processes with {args.samples_per_process} measured samples per process for each of {len(rows)} lanes.",
        "Elapsed time is measured with the harness's process-local counting allocator installed. It is allocator-instrumented operation time for the stated prepared scope, not a production latency claim.",
        "Requested allocation bytes charge direct allocation sizes plus each successful realloc's new size. Raw JSON records realloc old/new sizes, exact live balance, and peak live as an incremental delta above the closure's live-before baseline; it is not total process or retained source/result memory.",
        "Maximum RSS is whole-process startup/warm-up/sample RSS from `/usr/bin/time`, not per-operation memory. Source checks, semantic hashes, pointer checks, and inverse checks are outside timed/allocation regions.",
        "",
        "`metadata_read` is the new eager XLSB host read including family discovery. `source_read` uses a counted `ReadAt` source. `codec_read` is a shared DrawingML Theme codec component reference, while `base_edit` is a host base-Theme edit reference; neither is a HEAD-before baseline or a whole-library comparison.",
        "",
        "| fixture | operation | processes | samples | package B | Theme B | family B | allocator-instrumented p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live delta p50 / p95 B | RSS range | ReadAt calls / requested / returned B | source share |",
        "|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---|---:|---|",
    ]
    for row in rows:
        lines.append(
            "| {fixture} | {operation} | {processes} | {samples} | {bytes} | {theme_bytes} | {family_bytes} | "
            "{p50} / {p95} / {p99} | {alloc50} / {alloc95} | {peak50} / {peak95} | {rss} | {read_io} | {share} |".format(**row)
        )
    lines.extend(
        [
            "",
            "All lanes passed fixture shape, semantic, preservation, changed-edit, inverse, allocator-accounting, and source-observation gates. The `family_clone` and `noop` lanes additionally require the observed Family source pointer to be shared; the gate does not generalize to unrelated owners.",
            "",
            "The `native` fixture is `test-data/ooxml/xlsb/date.xlsb`, a small native workbook containing the Office `themeFamily` fragment. The `opaque` fixture is generated in memory from that same package: it retains the native family attributes and adds a family-namespace `extLst` with 96 valid DrawingML `a:ext` children, comments, vendor attributes/payload, and one distinct foreign namespace declaration plus nested payload per child. Its family shape is independently checked before timing.",
            "",
            "The synthetic opaque case exercises source-preserving edits at a larger fragment size. It is not a corpus-wide representativeness claim. No broad speedup or regression claim is made from these absolute lanes; a comparison to existing behavior would require a separately frozen control build using the same harness and inputs.",
            "",
            "Exact commands, toolchain/host data, binary SHA-256, pre/post Cargo package source manifests, dirty source hashes, raw per-process JSON, and `/usr/bin/time` output are retained beside this report.",
            f"Final provenance identifiers: binary SHA-256 `{binary_sha}`; matching source manifest SHA-256 `{manifest_sha}`.",
        ]
    )
    args.output.write_text("\n".join(lines) + "\n")


if __name__ == "__main__":
    main()
