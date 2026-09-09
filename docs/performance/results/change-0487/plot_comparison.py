#!/usr/bin/env python3
"""Render each normal-process p50 without averaging the two repeats."""
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

from support import ROOT, meta, read, write


def main():
    source = ROOT / "comparison-summary.json"
    summary = read(source)
    assert summary["inventory"]["formal_processes"] == 144
    rows = [row for row in summary["processes"] if row["role"] == "normal"]
    workloads = ["s64-a64-short-c64", "s64-a16384-short-c64", "s131072-a64-short-c64"]
    titles = ["64 source / 64 authored", "64 source / 16,384 authored", "131,072 source / 64 authored"]
    arms = [("deterministic", mode) for mode in ("owned", "file", "short-read", "latency")]
    arms += [("memory_store", "owned"), ("file_store", "owned")]
    labels = ["Owned input", "File input", "Short-read input", "Latency input", "Memory replay store", "File replay store"]
    destination = ROOT / "latency-comparison.svg"
    preview = ROOT / "latency-comparison.png"
    assert not destination.exists() and not preview.exists()
    with plt.rc_context({"font.size": 10, "svg.hashsalt": "litchi-0487"}):
        fig, axes = plt.subplots(1, 3, figsize=(14, 4.8), sharey=True)
        for ax, workload, title in zip(axes, workloads, titles):
            for index, (route, input_mode) in enumerate(arms):
                for repeat, offset, marker in ((1, -.12, "o"), (2, .12, "^") ):
                    pair = []
                    for phase in ("before", "after"):
                        matched = [r for r in rows if r["workload"] == workload
                                   and r["route"] == route and r["input_mode"] == input_mode
                                   and r["repeat"] == repeat and r["phase"] == phase]
                        assert len(matched) == 1
                        pair.append(matched[0]["metrics"]["elapsed_ns"]["p50"] / 1e6)
                    y = index + offset
                    ax.plot(pair, [y, y], color="#a8adb3", linewidth=1.1, zorder=1)
                    ax.scatter(pair, [y, y], c=["#b45309", "#087f8c"], marker=marker, s=38, zorder=2)
            ax.set_title(title, fontsize=11, pad=14)
            ax.set_xlim(left=0)
            ax.set_xlabel("Normal-process p50 (ms)")
            ax.grid(axis="x", alpha=.2)
            ax.spines[["top", "right", "left"]].set_visible(False)
        axes[0].set_yticks(range(len(labels)), labels)
        axes[0].invert_yaxis()
        fig.suptitle("OPC replay buffering: matched DOCX tail append", fontsize=14)
        legend = [Line2D([], [], marker="o", linestyle="", color="#b45309", label="Before"),
                  Line2D([], [], marker="o", linestyle="", color="#087f8c", label="After"),
                  Line2D([], [], marker="o", linestyle="", color="#555", label="Repeat 1"),
                  Line2D([], [], marker="^", linestyle="", color="#555", label="Repeat 2")]
        fig.legend(handles=legend, loc="lower center", ncol=4, frameon=False, bbox_to_anchor=(.5, .06))
        fig.text(.5, .025, "30 samples after 3 warmups per point; separate x scales. Lines connect matched repeats; no confidence interval.", ha="center", fontsize=9)
        fig.tight_layout(rect=(0, .17, 1, .92))
        fig.savefig(destination, metadata={"Date": None})
        fig.savefig(preview, dpi=160)
        plt.close(fig)
    write(ROOT / "plot-receipt.json", {"driver": meta(Path(__file__)), "summary": meta(source),
                                        "artifacts": {p.name: meta(p) for p in (destination, preview)},
                                        "matplotlib_version": matplotlib.__version__,
                                        "scope": "Normal-process p50, both repeats shown independently; tails and memory remain in full tables."})


if __name__ == "__main__":
    main()
