#!/usr/bin/env python3
"""Analyze the final change-0498 source-backed batch captures.

The capture runner writes one CSV and one ``/usr/bin/time -v`` receipt for
each corpus/source/API/worker configuration below ``final/``::

    {corpus}-{source}-serial-w1.csv
    {corpus}-{source}-batch-w{1,2,4,8}.csv

Only sample rows with index >= 3 are measured.  The first three rows in each
repeat are warmups.  This script intentionally uses only the Python standard
library so that the result can be regenerated on a machine without the
benchmark workspace's Rust dependencies.
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import math
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Iterable


CORPORA = ("few-large", "many-small")
SOURCES = ("owned", "file", "instrumented")
BATCH_WORKERS = (1, 2, 4, 8)
EXPECTED_STEMS = tuple(
    f"{corpus}-{source}-{mode}"
    for corpus in CORPORA
    for source in SOURCES
    for mode in ("serial-w1", *(f"batch-w{worker}" for worker in BATCH_WORKERS))
)
MEASURED_START = 3
MEASURED_END = 32
EXPECTED_REPEATS = (0, 1)
EXPECTED_SAMPLES_PER_REPEAT = MEASURED_END - MEASURED_START + 1
ADVERSE_THRESHOLD_PCT = 5.0
RSS_PATTERN = re.compile(
    r"Maximum resident set size \(kbytes\):\s*(?P<value>[0-9]+)"
)
STEM_PATTERN = re.compile(
    r"^(?P<corpus>few-large|many-small)-"
    r"(?P<source>owned|file|instrumented|short-delay)-"
    r"(?P<mode>serial-w1|batch-w(?P<workers>[1248]))$"
)


def percentile(values: list[float], percentile_value: int) -> float:
    """Return the nearest-rank percentile used by the Rust harness."""

    if not values:
        raise ValueError("cannot calculate a percentile of an empty sample set")
    ordered = sorted(values)
    rank = math.ceil(len(ordered) * percentile_value / 100) - 1
    return ordered[max(0, min(rank, len(ordered) - 1))]


def mean(values: Iterable[float]) -> float:
    values = list(values)
    if not values:
        raise ValueError("cannot calculate a mean of an empty sample set")
    return sum(values) / len(values)


def percent_delta(value: float, reference: float) -> float | None:
    if reference == 0:
        return None
    return (value - reference) * 100.0 / reference


def ratio(value: float, reference: float) -> float | None:
    if reference == 0:
        return None
    return value / reference


def finite(value: float | None) -> float | None:
    if value is None or not math.isfinite(value):
        return None
    return value


def parse_stem(path: Path) -> dict[str, Any]:
    match = STEM_PATTERN.fullmatch(path.stem)
    if match is None:
        raise ValueError(f"unexpected final capture name: {path.name}")
    source = match.group("source")
    if source == "short-delay":
        source = "instrumented"
    mode = match.group("mode")
    return {
        "corpus": match.group("corpus"),
        "source": source,
        "mode": "serial" if mode == "serial-w1" else "batch",
        "workers": 1 if mode == "serial-w1" else int(match.group("workers")),
        "stem": path.stem,
    }


def parse_rss(path: Path) -> dict[str, Any]:
    text = path.read_text(encoding="utf-8", errors="replace")
    matches = RSS_PATTERN.findall(text)
    if not matches:
        raise ValueError(f"{path.name} has no Maximum resident set size line")
    values = [int(value) for value in matches]
    return {
        "rss_kib": max(values),
        "rss_source": "usr-bin-time-v",
        "receipt": path.name,
    }


def load_capture(csv_path: Path) -> dict[str, Any]:
    identity = parse_stem(csv_path)
    time_path = csv_path.with_suffix(".time.stderr")
    if not time_path.is_file():
        raise FileNotFoundError(f"missing time receipt for {csv_path.name}: {time_path.name}")

    by_repeat: dict[int, list[dict[str, float]]] = {}
    with csv_path.open(newline="", encoding="utf-8") as handle:
        reader = csv.DictReader(handle)
        required = {"record", "repeat", "index", "elapsed_ns", "logical_bytes"}
        missing = required - set(reader.fieldnames or ())
        if missing:
            raise ValueError(f"{csv_path.name} is missing columns: {sorted(missing)}")
        for row in reader:
            if row.get("record") != "sample":
                continue
            repeat = int(row["repeat"])
            index = int(row["index"])
            if index < MEASURED_START:
                continue
            if index > MEASURED_END:
                raise ValueError(
                    f"{csv_path.name} has sample index {index}; expected at most {MEASURED_END}"
                )
            elapsed_ns = float(row["elapsed_ns"])
            logical_bytes = float(row["logical_bytes"])
            if elapsed_ns <= 0 or logical_bytes < 0:
                raise ValueError(f"{csv_path.name} has invalid sample {repeat}/{index}")
            by_repeat.setdefault(repeat, []).append(
                {
                    "index": index,
                    "elapsed_ns": elapsed_ns,
                    "logical_bytes": logical_bytes,
                }
            )

    if tuple(sorted(by_repeat)) != EXPECTED_REPEATS:
        raise ValueError(
            f"{csv_path.name} repeats {sorted(by_repeat)}; expected {list(EXPECTED_REPEATS)}"
        )

    repeats: dict[str, Any] = {}
    for repeat, rows in sorted(by_repeat.items()):
        if len(rows) != EXPECTED_SAMPLES_PER_REPEAT:
            raise ValueError(
                f"{csv_path.name} repeat {repeat} has {len(rows)} measured rows; "
                f"expected {EXPECTED_SAMPLES_PER_REPEAT}"
            )
        indexes = [int(row["index"]) for row in rows]
        expected_indexes = list(range(MEASURED_START, MEASURED_END + 1))
        if sorted(indexes) != expected_indexes:
            raise ValueError(f"{csv_path.name} repeat {repeat} has noncanonical sample indexes")
        elapsed_us = [row["elapsed_ns"] / 1_000.0 for row in rows]
        total_ns = sum(row["elapsed_ns"] for row in rows)
        total_bytes = sum(row["logical_bytes"] for row in rows)
        repeats[str(repeat)] = {
            "sample_count": len(rows),
            "p50_us": percentile(elapsed_us, 50),
            "p95_us": percentile(elapsed_us, 95),
            "p99_us": percentile(elapsed_us, 99),
            "mean_us": mean(elapsed_us),
            "throughput_bytes_s": total_bytes * 1_000_000_000.0 / total_ns,
            "logical_bytes": total_bytes,
        }

    all_rows = [row for rows in by_repeat.values() for row in rows]
    all_elapsed_us = [row["elapsed_ns"] / 1_000.0 for row in all_rows]
    all_total_ns = sum(row["elapsed_ns"] for row in all_rows)
    all_total_bytes = sum(row["logical_bytes"] for row in all_rows)
    aggregate = {
        "sample_count": len(all_rows),
        "p50_us": percentile(all_elapsed_us, 50),
        "p95_us": percentile(all_elapsed_us, 95),
        "p99_us": percentile(all_elapsed_us, 99),
        "mean_us": mean(all_elapsed_us),
        "throughput_bytes_s": all_total_bytes * 1_000_000_000.0 / all_total_ns,
        "logical_bytes": all_total_bytes,
    }
    return {
        **identity,
        "csv": csv_path.name,
        "repeats": repeats,
        "aggregate": aggregate,
        "time": parse_rss(time_path),
    }


def metric_delta(batch: dict[str, Any], serial: dict[str, Any]) -> dict[str, Any]:
    latency_metrics = ("p50_us", "p95_us", "p99_us", "mean_us")
    latency = {
        metric: {
            "serial": serial[metric],
            "batch": batch[metric],
            "delta_pct": percent_delta(batch[metric], serial[metric]),
        }
        for metric in latency_metrics
    }
    throughput_delta = percent_delta(
        batch["throughput_bytes_s"], serial["throughput_bytes_s"]
    )
    return {
        "latency": latency,
        "throughput": {
            "serial_bytes_s": serial["throughput_bytes_s"],
            "batch_bytes_s": batch["throughput_bytes_s"],
            "delta_pct": throughput_delta,
        },
    }


def adverse_flags(
    comparison: dict[str, Any], batch_rss: int | None = None, serial_rss: int | None = None
) -> list[str]:
    flags: list[str] = []
    for metric, values in comparison["latency"].items():
        delta = values["delta_pct"]
        if delta is not None and delta > ADVERSE_THRESHOLD_PCT:
            flags.append(f"latency_{metric}_regression_gt_5pct")
    throughput_delta = comparison["throughput"]["delta_pct"]
    if throughput_delta is not None and throughput_delta < -ADVERSE_THRESHOLD_PCT:
        flags.append("throughput_regression_gt_5pct")
    if batch_rss is not None and serial_rss is not None:
        rss_delta = percent_delta(float(batch_rss), float(serial_rss))
        if rss_delta is not None and rss_delta > ADVERSE_THRESHOLD_PCT:
            flags.append("rss_regression_gt_5pct")
    return flags


def compare_capture(batch: dict[str, Any], serial: dict[str, Any]) -> dict[str, Any]:
    repeats: dict[str, Any] = {}
    for repeat in EXPECTED_REPEATS:
        serial_stats = serial["repeats"][str(repeat)]
        batch_stats = batch["repeats"][str(repeat)]
        comparison = metric_delta(batch_stats, serial_stats)
        # RSS is one whole-child receipt per capture, not a per-repeat
        # observation. Keep it at aggregate scope below so it is not counted
        # twice as if each repeat had an independent process RSS measurement.
        comparison["adverse_flags"] = adverse_flags(comparison)
        repeats[str(repeat)] = comparison

    aggregate = metric_delta(batch["aggregate"], serial["aggregate"])
    aggregate["rss"] = {
        "serial_kib": serial["time"]["rss_kib"],
        "batch_kib": batch["time"]["rss_kib"],
        "delta_pct": percent_delta(
            float(batch["time"]["rss_kib"]), float(serial["time"]["rss_kib"])
        ),
    }
    aggregate["adverse_flags"] = adverse_flags(
        aggregate, batch["time"]["rss_kib"], serial["time"]["rss_kib"]
    )
    return {"repeats": repeats, "aggregate": aggregate}


def amdahl_fraction(speedup: float | None, workers: int) -> float | None:
    """Return the raw simple-Amdahl fraction for a useful positive speedup.

    The caller decides whether the result is meaningful.  In particular, a
    speedup greater than the worker count is retained as a negative raw value
    but is labelled superlinear and omitted from the reported model value.
    """

    if speedup is None or workers <= 1 or speedup <= 1.0:
        return None
    return (workers / speedup - 1.0) / (workers - 1.0)


def amdahl_report(
    speedup: float | None, workers: int, selected_parts: int
) -> dict[str, Any]:
    raw = amdahl_fraction(speedup, workers)
    if speedup is None:
        return {
            "amdahl_serial_fraction": None,
            "amdahl_raw_serial_fraction": None,
            "amdahl_meaningful": False,
            "amdahl_status": "unavailable",
        }
    if workers > selected_parts:
        return {
            "amdahl_serial_fraction": None,
            "amdahl_raw_serial_fraction": raw,
            "amdahl_meaningful": False,
            "amdahl_status": "parallel_work_cap_simple_amdahl_invalid",
        }
    if speedup > workers:
        return {
            "amdahl_serial_fraction": None,
            "amdahl_raw_serial_fraction": raw,
            "amdahl_meaningful": False,
            "amdahl_status": "superlinear_observed_simple_amdahl_invalid",
        }
    if speedup <= 1.0:
        return {
            "amdahl_serial_fraction": None,
            "amdahl_raw_serial_fraction": raw,
            "amdahl_meaningful": False,
            "amdahl_status": "no_speedup_simple_amdahl_invalid",
        }
    return {
        "amdahl_serial_fraction": raw,
        "amdahl_raw_serial_fraction": raw,
        "amdahl_meaningful": True,
        "amdahl_status": "descriptive_valid_range",
    }


def scaling_capture(
    batch: dict[str, Any], baseline: dict[str, Any], selected_parts: int
) -> dict[str, Any]:
    batch_aggregate = batch["aggregate"]
    baseline_aggregate = baseline["aggregate"]
    speedup_p50 = ratio(baseline_aggregate["p50_us"], batch_aggregate["p50_us"])
    throughput_speedup = ratio(
        batch_aggregate["throughput_bytes_s"], baseline_aggregate["throughput_bytes_s"]
    )
    workers = batch["workers"]
    effective_workers = min(workers, selected_parts)
    amdahl = amdahl_report(speedup_p50, workers, selected_parts)
    result = {
        "workers": workers,
        "baseline_workers": 1,
        "selected_parts": selected_parts,
        "effective_parallel_width": effective_workers,
        "aggregate": {
            "p50_speedup": speedup_p50,
            "throughput_speedup": throughput_speedup,
            "efficiency": ratio(speedup_p50, float(workers)),
            "effective_efficiency": ratio(speedup_p50, float(effective_workers)),
            **amdahl,
        },
        "repeats": {},
        "interpretation": (
            "Descriptive same corpus/source comparison. Amdahl is reported only "
            "when p50 latency improves; it is not a causal attribution."
        ),
    }
    for repeat in EXPECTED_REPEATS:
        target = batch["repeats"][str(repeat)]
        base = baseline["repeats"][str(repeat)]
        speedup = ratio(base["p50_us"], target["p50_us"])
        amdahl = amdahl_report(speedup, workers, selected_parts)
        result["repeats"][str(repeat)] = {
            "p50_speedup": speedup,
            "throughput_speedup": ratio(
                target["throughput_bytes_s"], base["throughput_bytes_s"]
            ),
            "efficiency": ratio(speedup, float(workers)),
            "effective_efficiency": ratio(speedup, float(effective_workers)),
            **amdahl,
        }
    return result


def markdown_number(value: Any, digits: int = 2) -> str:
    if value is None:
        return "—"
    if isinstance(value, (int, float)):
        return f"{value:.{digits}f}"
    return str(value)


def build_markdown(result: dict[str, Any]) -> str:
    lines = [
        "# Change 0498 final benchmark analysis",
        "",
        f"Generated: `{result['generated_at_utc']}`",
        "",
        "The analysis uses measured sample indexes 3–32 from each of two repeats; "
        "indexes 0–2 are warmups. Percentiles use the nearest-rank rule used by "
        "the Rust harness. RSS is the whole child maximum from `/usr/bin/time -v`. "
        "The measured process was pinned to the configured benchmark CPU set on a "
        "shared host.",
        "",
        "## Serial versus batch",
        "",
        "Positive latency or RSS deltas and negative throughput deltas are adverse. "
        "The flag threshold is greater than 5 percent.",
        "",
        "| corpus | source | workers | p50 batch/serial (us) | p95 batch/serial (us) | p99 batch/serial (us) | mean batch/serial (us) | throughput delta | RSS delta | flags |",
        "| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |",
    ]
    for item in result["comparisons"]:
        aggregate = item["comparison"]["aggregate"]
        lines.append(
            "| {corpus} | {source} | {workers} | {p50} | {p95} | {p99} | {mean} | {throughput}% | {rss}% | {flags} |".format(
                corpus=item["corpus"],
                source=item["source"],
                workers=item["workers"],
                p50=(
                    f"{markdown_number(aggregate['latency']['p50_us']['batch'])}/"
                    f"{markdown_number(aggregate['latency']['p50_us']['serial'])}"
                ),
                p95=(
                    f"{markdown_number(aggregate['latency']['p95_us']['batch'])}/"
                    f"{markdown_number(aggregate['latency']['p95_us']['serial'])}"
                ),
                p99=(
                    f"{markdown_number(aggregate['latency']['p99_us']['batch'])}/"
                    f"{markdown_number(aggregate['latency']['p99_us']['serial'])}"
                ),
                mean=(
                    f"{markdown_number(aggregate['latency']['mean_us']['batch'])}/"
                    f"{markdown_number(aggregate['latency']['mean_us']['serial'])}"
                ),
                throughput=markdown_number(aggregate["throughput"]["delta_pct"]),
                rss=markdown_number(aggregate["rss"]["delta_pct"]),
                flags=", ".join(aggregate["adverse_flags"]) or "—",
            )
        )

    lines.extend(
        [
            "",
            "## Flag counts",
            "",
            "RSS flags are counted only at aggregate comparison scope because each "
            "capture has one whole-child receipt; per-repeat counts cover latency and "
            "throughput flags only.",
            "",
            "| scope | flag | count |",
            "| --- | --- | ---: |",
        ]
    )
    flag_summary = result["flag_summary"]
    for scope, key in (
        ("aggregate", "aggregate_flags"),
        ("per-repeat, RSS excluded", "per_repeat_flags_excluding_rss"),
    ):
        flags = flag_summary[key]
        if flags:
            for flag, count in flags.items():
                lines.append(f"| {scope} | {flag} | {count} |")
        else:
            lines.append(f"| {scope} | — | 0 |")

    lines.extend(
        [
            "",
            "The JSON contains the corresponding per-repeat latency, throughput, and "
            "RSS comparisons. The rows below use batch worker 1 as the scaling "
            "reference; requested efficiency is p50 latency speedup divided by "
            "requested worker count, and effective efficiency uses the selected-Part "
            "work cap.",
            "",
            "## Batch scaling",
            "",
            "| corpus | source | workers | p50 speedup | throughput speedup | efficiency requested/effective | simple Amdahl result |",
            "| --- | --- | ---: | ---: | ---: | ---: | ---: |",
        ]
    )
    for item in result["scaling"]:
        aggregate = item["scaling"]["aggregate"]
        amdahl = aggregate["amdahl_status"]
        if aggregate["amdahl_meaningful"]:
            amdahl = markdown_number(aggregate["amdahl_serial_fraction"])
        elif amdahl == "superlinear_observed_simple_amdahl_invalid":
            amdahl = "invalid: superlinear"
        elif amdahl == "parallel_work_cap_simple_amdahl_invalid":
            amdahl = f"invalid: {item['scaling']['selected_parts']}-Part cap"
        elif amdahl == "no_speedup_simple_amdahl_invalid":
            amdahl = "invalid: no speedup"
        lines.append(
            "| {corpus} | {source} | {workers} | {speedup} | {throughput} | {efficiency} | {amdahl} |".format(
                corpus=item["corpus"],
                source=item["source"],
                workers=item["workers"],
                speedup=markdown_number(aggregate["p50_speedup"]),
                throughput=markdown_number(aggregate["throughput_speedup"]),
                efficiency=(
                    f"{markdown_number(aggregate['efficiency'])}/"
                    f"{markdown_number(aggregate['effective_efficiency'])}"
                ),
                amdahl=amdahl,
            )
        )

    lines.extend(
        [
            "",
            "Simple-Amdahl values are descriptive model outputs for matched captures. "
            "Superlinear observations and widths above the selected-Part work cap are "
            "labelled invalid for this model rather than being fitted or clamped; raw "
            "fractions remain in JSON for traceability. No scaling row establishes a "
            "causal mechanism.",
            "",
            "## Historical baseline boundary",
            "",
            result["historical_baseline_note"],
            "",
            "The final files are process-level measurements: setup, repeated package "
            "opens, and post-timer verification are included in the whole-child RSS "
            "receipt, while the timed CSV interval excludes corpus construction and "
            "post-timer digest verification. These results support descriptive "
            "matched comparisons and do not establish operation-local causal costs.",
            "",
        ]
    )
    return "\n".join(lines)


def summarize_flags(comparisons: list[dict[str, Any]]) -> dict[str, Any]:
    aggregate = Counter()
    per_repeat = Counter()
    for item in comparisons:
        aggregate.update(item["comparison"]["aggregate"]["adverse_flags"])
        for repeat in item["comparison"]["repeats"].values():
            # There is one RSS receipt per process/capture.  Per-repeat counts
            # intentionally cover timed latency/throughput flags only.
            per_repeat.update(
                flag for flag in repeat["adverse_flags"] if not flag.startswith("rss_")
            )
    return {
        "aggregate_comparison_rows": len(comparisons),
        "aggregate_flags": dict(sorted(aggregate.items())),
        "per_repeat_comparison_rows": len(comparisons) * len(EXPECTED_REPEATS),
        "per_repeat_flags_excluding_rss": dict(sorted(per_repeat.items())),
        "rss_scope": "aggregate comparison rows only; one whole-child receipt per capture",
    }


def analyze(final_dir: Path) -> dict[str, Any]:
    expected = {f"{stem}.csv" for stem in EXPECTED_STEMS}
    found = {path.name for path in final_dir.glob("*.csv")}
    missing = sorted(expected - found)
    unexpected = sorted(found - expected)
    if missing:
        raise FileNotFoundError(
            f"final capture directory is incomplete; missing {', '.join(missing)}"
        )

    captures = [load_capture(final_dir / f"{stem}.csv") for stem in EXPECTED_STEMS]
    by_key = {(capture["corpus"], capture["source"], capture["mode"], capture["workers"]): capture for capture in captures}

    comparisons: list[dict[str, Any]] = []
    scaling: list[dict[str, Any]] = []
    for corpus in CORPORA:
        for source in SOURCES:
            serial = by_key[(corpus, source, "serial", 1)]
            for workers in BATCH_WORKERS:
                batch = by_key[(corpus, source, "batch", workers)]
                comparisons.append(
                    {
                        "corpus": corpus,
                        "source": source,
                        "workers": workers,
                        "serial_csv": serial["csv"],
                        "batch_csv": batch["csv"],
                        "comparison": compare_capture(batch, serial),
                    }
                )
            batch_one = by_key[(corpus, source, "batch", 1)]
            for workers in BATCH_WORKERS:
                if workers == 1:
                    continue
                batch = by_key[(corpus, source, "batch", workers)]
                scaling.append(
                    {
                        "corpus": corpus,
                        "source": source,
                        "workers": workers,
                        "baseline_csv": batch_one["csv"],
                        "batch_csv": batch["csv"],
                        "scaling": scaling_capture(
                            batch,
                            batch_one,
                            4 if corpus == "few-large" else 64,
                        ),
                    }
                )

    return {
        "generated_at_utc": dt.datetime.now(dt.UTC).replace(microsecond=0).isoformat(),
        "final_dir": str(final_dir),
        "expected_csv_count": len(expected),
        "found_csv_count": len(found),
        "unexpected_csv_files": unexpected,
        "warmup_indexes_excluded": [0, 1, 2],
        "measured_indexes": [MEASURED_START, MEASURED_END],
        "repeats": list(EXPECTED_REPEATS),
        "adverse_threshold_pct": ADVERSE_THRESHOLD_PCT,
        "captures": captures,
        "comparisons": comparisons,
        "scaling": scaling,
        "flag_summary": summarize_flags(comparisons),
        "historical_baseline_note": (
            "The retained before executable used data_with_accounting on its primary "
            "serial path, while the final serial control uses plain data() by default. "
            "Any before-versus-final latency comparison is therefore descriptive and "
            "has an accounting-path confound; this script does not merge historical "
            "baseline numbers into the matched batch/scaling tables."
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--final-dir",
        type=Path,
        default=Path(__file__).with_name("final"),
        help="directory containing final CSV and .time.stderr captures",
    )
    parser.add_argument(
        "--output-json",
        type=Path,
        default=Path(__file__).with_name("final-analysis.json"),
    )
    parser.add_argument(
        "--output-md",
        type=Path,
        default=Path(__file__).with_name("final-analysis.md"),
    )
    args = parser.parse_args()
    try:
        result = analyze(args.final_dir)
    except (FileNotFoundError, ValueError, KeyError) as error:
        print(f"analyze-final.py: {error}", file=sys.stderr)
        return 2
    args.output_json.write_text(
        json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    args.output_md.write_text(build_markdown(result), encoding="utf-8")
    print(f"wrote {args.output_json}")
    print(f"wrote {args.output_md}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
