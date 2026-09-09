#!/usr/bin/env python3
"""Render the verified formal summary as complete tables and a scaling figure."""

import argparse
import csv
import hashlib
import json
from pathlib import Path
import statistics


ROUTES = ("deterministic", "memory_store", "file_store")
COLORS = ("#276FBF", "#D97706", "#168A75")


def meta(path):
    with path.open("rb") as stream:
        return {"bytes": path.stat().st_size, "sha256": hashlib.file_digest(stream, "sha256").hexdigest()}


def paired_rows(rows):
    allocation = {(r["identity"], r["case"], r["repeat"]): r for r in rows if r["role"] == "allocator"}
    output = []
    for row in rows:
        if row["role"] != "normal":
            continue
        other = allocation[(row["identity"], row["case"], row["repeat"])]
        elapsed = row["metrics"]["elapsed_ns"]
        output.append({
            "case": row["case"], "profile": row["identity"], "repeat": row["repeat"],
            "p50_ms": elapsed["p50"] / 1e6, "p95_ms": elapsed["p95"] / 1e6,
            "p99_ms": elapsed["p99"] / 1e6,
            "operation_heap_increment_bytes": other["allocator"]["operation_peak_increment_bytes"]["p50"],
            "allocated_bytes": other["allocator"]["allocated_bytes"]["p50"],
            "allocation_calls": other["allocator"]["allocation_calls"]["p50"],
            "normal_process_max_rss_bytes": row["time_max_rss_bytes"],
            "source_MiB_per_s": row["metrics"]["source_throughput_bytes_per_second"]["p50"] / 2**20,
            "authored_MiB_per_s": row["metrics"]["authored_throughput_bytes_per_second"]["p50"] / 2**20,
        })
    return sorted(output, key=lambda r: (r["case"], str(r["profile"]), r["repeat"]))


def table(rows):
    lines = [
        "| Workload | Profile | Repeat | p50 ms | p95 ms | p99 ms | Operation heap increment KiB | Requested allocation KiB | Allocation calls | Process RSS MiB |",
        "| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for row in rows:
        lines.append(f"| {row['case']} | {row['profile']} | {row['repeat']} | {row['p50_ms']:.4f} | {row['p95_ms']:.4f} | {row['p99_ms']:.4f} | {row['operation_heap_increment_bytes']/1024:.3f} | {row['allocated_bytes']/1024:.3f} | {row['allocation_calls']:g} | {row['normal_process_max_rss_bytes']/2**20:.3f} |")
    return "\n".join(lines)


def plot(summary, destination):
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt

    indexed = {(r["case"], r["route"], r["role"], r["repeat"]): r for r in summary["route_processes"]}
    fig, axes = plt.subplots(2, 2, figsize=(11, 7), layout="constrained")
    groups = (
        ("Existing source paragraphs · 64 authored", [64, 8192, 131072], lambda n: f"s{n}-a64-short-c64"),
        ("Authored paragraphs · 64 existing", [64, 256, 4096, 16384], lambda n: f"s64-a{n}-short-c64"),
    )
    for col, (label, counts, case_label) in enumerate(groups):
        for route, color in zip(ROUTES, COLORS):
            for row, role, metric in ((0, "normal", "latency"), (1, "allocator", "heap")):
                values = []
                for count in counts:
                    repeats = [indexed[(case_label(count), route, role, repeat)] for repeat in (1, 2)]
                    values.append([r["metrics"]["elapsed_ns"]["p50"] / 1e6 if metric == "latency" else r["allocator"]["operation_peak_increment_bytes"]["p50"] / 1024 for r in repeats])
                centers = [statistics.mean(pair) for pair in values]
                errors = [[center - min(pair) for center, pair in zip(centers, values)], [max(pair) - center for center, pair in zip(centers, values)]]
                axes[row, col].errorbar(counts, centers, yerr=errors, marker="o", capsize=3, color=color, label=route)
        for row in (0, 1):
            axes[row, col].set_xscale("log", base=2)
            axes[row, col].set_yscale("log")
            axes[row, col].set_xticks(counts, [f"{n:,}" for n in counts])
            axes[row, col].grid(True, which="both", alpha=0.2)
            axes[row, col].set_xlabel(label)
    axes[0, 0].set_ylabel("Normal operation latency (ms)")
    axes[1, 0].set_ylabel("Allocator operation peak increment (KiB)")
    axes[0, 0].legend(fontsize=8)
    fig.suptitle("DOCX paragraph append · separate source and authored scaling\nTwo process medians: marker average and repeat range, not confidence intervals", fontsize=12)
    fig.savefig(destination, dpi=180, metadata={"Software": "Litchi 0484 verified formal summary renderer"})
    plt.close(fig)
    return {"matplotlib": matplotlib.__version__, "backend": matplotlib.get_backend(), "dpi": 180}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--summary", type=Path, required=True)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    summary = json.loads(args.summary.read_text())
    assert summary["schema"] == "docx-replayable-tail-append-route-analysis-v1"
    assert summary["inventory"]["route_processes"] == 120 and summary["inventory"]["axis_processes"] == 108
    assert summary["inventory"]["pilots_included"] is False
    args.output_dir.mkdir(parents=True, exist_ok=False)
    route_rows = paired_rows(summary["route_processes"])
    axis_rows = paired_rows(summary["axis_processes"])
    assert len(route_rows) == 60 and len(axis_rows) == 54
    all_rows = route_rows + axis_rows
    csv_path = args.output_dir / "measurements.csv"
    with csv_path.open("x", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=list(all_rows[0]))
        writer.writeheader()
        writer.writerows(all_rows)
    plotting = plot(summary, args.output_dir / "route-scaling.png")
    text = """# DOCX replay route and input-profile measurements

This table covers all 228 formal child processes and 6,840 measured operations.
Each row pairs one normal process with its separate allocator process for the
same workload/profile/repeat. The 114 pilots are excluded. Two process repeats
provide descriptive repeat spread, not a confidence interval or proof of a
historical optimization speedup. The full performance goal remains open.

Latency percentiles come from 30 normal-build samples. Heap increment is the
median of each allocator sample's region peak minus its starting live bytes;
requested bytes and calls are allocator medians. RSS is one whole-normal-child
GNU time maximum, including fixture setup and report work. It is not an
operation-only memory measurement. Tail percentiles are descriptive with this
sample count. Allocation requests are not physical memory-copy counts.

The source and authored axes are varied separately. Input files are prepared
and fingerprinted before timing; cache eviction is not performed. File replay
uses explicit data sync and same-operation page-cache participation. These
measurements do not establish cold-cache or atomic-save performance.

![Source and authored scaling](route-scaling.png)

## Replay routes

""" + table(route_rows) + "\n\n## Input, sink and compression profiles\n\n" + table(axis_rows)
    text += "\n\n[Complete CSV](measurements.csv) also includes source and authored throughput. "
    text += "The machine-readable verified summary retains individual repeat statistics, detailed counters, comparison flags and input identities. "
    text += "Changes above 5% trigger review; favorable and adverse differences remain visible individually.\n"
    (args.output_dir / "measurements.md").write_text(text)
    receipt = {
        "schema": "docx-route-render-v1", "summary": {"path": str(args.summary.resolve()), **meta(args.summary)},
        "driver_sha256": meta(Path(__file__))["sha256"], "route_rows": len(route_rows), "axis_rows": len(axis_rows),
        "plotting": plotting,
        "artifacts": {p.name: meta(p) for p in sorted(args.output_dir.iterdir()) if p.is_file()},
    }
    with (args.output_dir / "render-receipt.json").open("x") as stream:
        json.dump(receipt, stream, indent=2, sort_keys=True, allow_nan=False)
        stream.write("\n")
    print(f"Rendered {len(all_rows)} paired process rows and one scaling figure")


if __name__ == "__main__":
    main()
