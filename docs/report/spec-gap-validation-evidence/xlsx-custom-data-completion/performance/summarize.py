#!/usr/bin/env python3
"""Summarize matched XLSX Custom Data profile receipts."""

from __future__ import annotations

import argparse
import json
import statistics
from pathlib import Path

LEVELS = ("small", "medium", "large")
LANES = (
    "read",
    "noop-commit-save",
    "payload-replacement",
    "rename-binding-rewrite",
    "remove-inverse",
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline", type=Path, required=True)
    parser.add_argument("--candidate", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--samples", type=int, default=15)
    return parser.parse_args()


def read_capture(root: Path, expected_samples: int) -> dict:
    provenance = json.loads((root / "provenance.json").read_text(encoding="utf-8"))
    if provenance["samples"] != expected_samples or provenance["warmups"] != 3:
        raise SystemExit(
            f"{root}: expected {expected_samples} samples and 3 warmups, "
            f"got {provenance['samples']} and {provenance['warmups']}"
        )
    stability = json.loads(
        (root / "source-manifest-stability.json").read_text(encoding="utf-8")
    )
    if stability.get("equal") is not True:
        raise SystemExit(f"{root}: source manifest was not stable")

    groups: dict[tuple[str, str], list[dict]] = {}
    for level in LEVELS:
        for lane in LANES:
            path = root / f"{level}-{lane}.jsonl"
            rows = [json.loads(line) for line in path.read_text().splitlines() if line]
            if len(rows) != expected_samples:
                raise SystemExit(f"{path}: expected {expected_samples} rows, got {len(rows)}")
            for sample, row in enumerate(rows, start=1):
                if row["sample"] != sample or row["warmups"] != 3:
                    raise SystemExit(f"{path}: sample/warmup sequence is invalid")
                if not row["allocator"]["balanced"]:
                    raise SystemExit(f"{path}: allocator balance failed")
                flags = row["validation"]
                if not all(flags[key] for key in ("payloads_exact", "bindings_exact")):
                    raise SystemExit(f"{path}: semantic validation failed")
                time_path = root / f"{level}-{lane}-sample{sample}.time.txt"
                row["rss_kib"] = read_rss(time_path)
            groups[(level, lane)] = rows

    source_hashes = {
        row["fixture"]["source_sha256"]
        for rows in groups.values()
        for row in rows
    }
    if len(source_hashes) != 3:
        raise SystemExit(f"{root}: expected one fixture hash per size class, got {source_hashes}")
    return {"root": root, "provenance": provenance, "groups": groups}


def read_rss(path: Path) -> int:
    prefix = "Maximum resident set size (kbytes):"
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.strip().startswith(prefix):
            return int(line.split(":", 1)[1].strip())
    raise SystemExit(f"{path}: missing /usr/bin/time RSS field")


def median(values: list[int | float]) -> int | float:
    value = statistics.median(values)
    return int(value) if isinstance(value, float) and value.is_integer() else value


def percentile(values: list[int | float], fraction: float) -> int | float:
    ordered = sorted(values)
    index = max(0, min(len(ordered) - 1, int((len(ordered) - 1) * fraction + 0.5)))
    return ordered[index]


def group_summary(rows: list[dict]) -> dict:
    elapsed = [row["elapsed_ns"] / 1_000_000 for row in rows]
    requested = [row["allocator"]["requested_bytes"] for row in rows]
    peak = [row["allocator"]["peak_live_delta"] for row in rows]
    rss = [row["rss_kib"] for row in rows]
    first = rows[0]
    return {
        "source_bytes": first["fixture"]["source_bytes"],
        "source_sha256": first["fixture"]["source_sha256"],
        "storages": first["fixture"]["storages"],
        "references": first["fixture"]["references"],
        "payload_bytes": first["fixture"]["payload_bytes"],
        "samples": len(rows),
        "elapsed_ms_median": median(elapsed),
        "elapsed_ms_p90": percentile(elapsed, 0.90),
        "requested_bytes_median": median(requested),
        "requested_bytes_p90": percentile(requested, 0.90),
        "peak_live_bytes_median": median(peak),
        "peak_live_bytes_p90": percentile(peak, 0.90),
        "rss_kib_median": median(rss),
        "rss_kib_p90": percentile(rss, 0.90),
    }


def percent_delta(baseline: float | int, candidate: float | int) -> str:
    if baseline == 0:
        return "n/a"
    return f"{(candidate - baseline) / baseline * 100:+.1f}%"


def render(baseline: dict, candidate: dict | None) -> str:
    lines = [
        "# XLSX Custom Data completion performance",
        "",
        "This report summarizes fresh-process matched observations from the authored, bounded fixtures. It does not establish native Office acceptance, tail-latency guarantees, or asymptotic complexity.",
        "",
        f"Baseline source: `{baseline['provenance']['source_commit']}`; samples per lane/size: `{baseline['provenance']['samples']}`; warmups: `{baseline['provenance']['warmups']}`.",
    ]
    if candidate is None:
        lines.append("Candidate capture: pending the root freeze receipt.")
    else:
        lines.append(f"Candidate source: `{candidate['provenance']['source_commit']}`; both captures use the same harness profile and input matrix.")
    lines.extend(
        [
            "",
            "| size | lane | storages | bindings | payload | median ms | p90 ms | median requested bytes | median peak live bytes | median RSS KiB | candidate elapsed | candidate requested | candidate peak | candidate RSS |",
            "| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
        ]
    )
    for level in LEVELS:
        for lane in LANES:
            base = group_summary(baseline["groups"][(level, lane)])
            if candidate is None:
                comparison = ["pending"] * 4
            else:
                cand = group_summary(candidate["groups"][(level, lane)])
                if cand["source_sha256"] != base["source_sha256"]:
                    raise SystemExit(f"fixture hash differs for {level}/{lane}")
                comparison = [
                    percent_delta(base["elapsed_ms_median"], cand["elapsed_ms_median"]),
                    percent_delta(base["requested_bytes_median"], cand["requested_bytes_median"]),
                    percent_delta(base["peak_live_bytes_median"], cand["peak_live_bytes_median"]),
                    percent_delta(base["rss_kib_median"], cand["rss_kib_median"]),
                ]
            lines.append(
                f"| {level} | {lane} | {base['storages']} | {base['references']} | {base['payload_bytes']} B | "
                f"{base['elapsed_ms_median']} | {base['elapsed_ms_p90']} | {base['requested_bytes_median']} | "
                f"{base['peak_live_bytes_median']} | {base['rss_kib_median']} | "
                f"{comparison[0]} | {comparison[1]} | {comparison[2]} | {comparison[3]} |"
            )
    lines.extend(
        [
            "",
            "The timed boundary includes package ingress, the public Custom Data operation, XLSX serialization, and result reopening. Allocator values include the observer's atomic accounting overhead. `requested bytes` is direct allocation plus new realloc bytes; `peak live bytes` is aggregate allocator accounting; RSS comes from `/usr/bin/time -v`.",
            "",
            "The three sizes vary storage count, connection count, and payload bytes together. Their rows characterize the authored workloads and should not be read as a proof of an asymptotic slope. Percent deltas compare medians only; p90 columns are descriptive sample percentiles.",
        ]
    )
    return "\n".join(lines) + "\n"


def main() -> None:
    args = parse_args()
    baseline = read_capture(args.baseline, args.samples)
    candidate = read_capture(args.candidate, args.samples) if args.candidate else None
    args.output.write_text(render(baseline, candidate), encoding="utf-8")


if __name__ == "__main__":
    main()
