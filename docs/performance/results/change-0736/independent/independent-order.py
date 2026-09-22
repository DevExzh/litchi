#!/usr/bin/env python3
"""Read-only audit of sample position, lag, and oracle cadence in 0735.

This script never starts a probe, Cargo, a profiler, or a native process.  It
validates the sealed 0735 packet, reads its 36 native reports, and writes a
descriptive result beside this script.  Samples inside one process and
overlapping windows are retained as sensitivity observations; the nine paired
processes per fixture remain the comparison units.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
from pathlib import Path
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
PACKET = HERE.parents[1] / "change-0735"
CAPTURES = PACKET / "captures"
OUTPUT = HERE / "independent-analysis.json"
PACKET_MANIFEST_SHA256 = "dbe26c50532792cba24871862c96dc0889b3fc21626cb57357fab6a08ff03955"
CASES = ("primary", "secondary")
VARIANTS = ("baseline", "candidate")
WINDOWS = {
    "all": (0, 50),
    "first10": (0, 10),
    "middle30": (10, 40),
    "last10": (40, 50),
    "first25": (0, 25),
    "last25": (25, 50),
}
INDEX_BINS = {
    "00-04": (0, 5),
    "05-09": (5, 10),
    "10-14": (10, 15),
    "15-19": (15, 20),
    "20-24": (20, 25),
    "25-29": (25, 30),
    "30-34": (30, 35),
    "35-39": (35, 40),
    "40-44": (40, 45),
    "45-49": (45, 50),
}
LAGS = (1, 2, 5, 10)


class Failure(Exception):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise Failure(message)


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    return sha_bytes(path.read_bytes())


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise Failure(f"invalid JSON {path}: {error}") from error


def median(values: Iterable[float | int]) -> float | int:
    values = list(values)
    require(values, "median received no values")
    return statistics.median(values)


def percent(before: float | int, after: float | int) -> float:
    require(before > 0, "percentage denominator is not positive")
    return 100.0 * (after / before - 1.0)


def summary(values: Iterable[float | int]) -> dict[str, Any]:
    values = list(values)
    require(values, "summary received no values")
    return {
        "n": len(values),
        "minimum": min(values),
        "median": median(values),
        "maximum": max(values),
        "mean": math.fsum(values) / len(values),
        "positive": sum(value > 0 for value in values),
        "above_five": sum(value > 5 for value in values),
        "below_minus_five": sum(value < -5 for value in values),
    }


def pearson(left: list[float | int], right: list[float | int]) -> float:
    require(len(left) == len(right) and len(left) >= 2, "invalid correlation vectors")
    left_mean = statistics.fmean(left)
    right_mean = statistics.fmean(right)
    numerator = math.fsum((a - left_mean) * (b - right_mean)
                          for a, b in zip(left, right))
    left_ss = math.fsum((a - left_mean) ** 2 for a in left)
    right_ss = math.fsum((b - right_mean) ** 2 for b in right)
    return numerator / math.sqrt(left_ss * right_ss) if left_ss and right_ss else 0.0


def line_number(text: str, needle: str) -> int:
    position = text.find(needle)
    require(position >= 0, f"source marker not found: {needle!r}")
    return text.count("\n", 0, position) + 1


def verify_packet() -> dict[str, Any]:
    manifest_path = PACKET / "artifact-manifest.json"
    require(sha(manifest_path) == PACKET_MANIFEST_SHA256,
            "0735 artifact manifest digest changed")
    manifest = read_json(manifest_path)
    files = manifest.get("files")
    require(isinstance(files, dict), "0735 artifact manifest files is not an object")
    actual = set()
    for path in PACKET.rglob("*"):
        require(not path.is_symlink(), f"0735 packet contains symlink: {path}")
        if path.is_file():
            actual.add(str(path.relative_to(PACKET)))
    require(actual == set(files) | {"artifact-manifest.json"},
            "0735 packet inventory differs from its sealed manifest")
    for relative, row in files.items():
        path = PACKET / relative
        require(isinstance(row, dict), f"manifest row is not an object: {relative}")
        require(path.stat().st_size == row.get("bytes"), f"byte count changed: {relative}")
        require(sha(path) == row.get("sha256"), f"digest changed: {relative}")
    return {"sha256": PACKET_MANIFEST_SHA256, "files": len(files)}


def source_lifecycle() -> dict[str, Any]:
    source_path = PACKET / "probe/src/lib.rs"
    source = source_path.read_text(encoding="utf-8")
    timed_start = source.index("fn timed_format")
    timed_end = source.index("fn timed_container", timed_start)
    timed_format = source[timed_start:timed_end]
    run_start = source.index("pub fn run")
    run_source = source[run_start:]
    warmup_start = run_source.index("for _ in 0..args.warmups")
    measured_start = run_source.index("for index in 0..args.samples")
    warmup_source = run_source[warmup_start:measured_start]
    expected_position = run_source.index("let expected_bytes = public_format_edit")
    warmup_position = run_source.index("for _ in 0..args.warmups")
    timed_output_position = run_source.index("let result = timed_format")
    timed_sample_position = run_source.index("samples.push(", timed_output_position)
    require("output_sample" not in timed_format,
            "timed_format now invokes output_sample inside its timer")
    require("drop(result.output)" in warmup_source,
            "warmup path no longer drops the timed result directly")
    require("output_sample" not in warmup_source,
            "warmup path now invokes output_sample")
    require(expected_position < warmup_position,
            "expected output/oracle preparation moved into the warmup loop")
    require(timed_output_position < timed_sample_position,
            "sample output validation is not after timed_format")
    require("let source_bytes = std::fs::read(&args.input)?;" in run_source,
            "source read marker disappeared")
    source_read_position = run_source.index("let source_bytes = std::fs::read(&args.input)?;")
    require(source_read_position < warmup_position,
            "source read moved into the warmup loop")
    return {
        "path": "probe/src/lib.rs",
        "sha256": sha(source_path),
        "timed_format_line": line_number(source, "fn timed_format"),
        "warmup_loop_line": line_number(source, "for _ in 0..args.warmups"),
        "measured_loop_line": line_number(source, "for index in 0..args.samples"),
        "output_sample_line": line_number(source, "samples.push("),
        "source_read_before_warmups": True,
        "timed_format_excludes_output_sample": True,
        "warmups_drop_without_output_sample": True,
        "measured_sample_validates_after_timer": True,
        "oracle_and_expected_prepared_before_warmups": True,
    }


def load_native() -> tuple[dict[tuple[str, int, int, str], dict[str, Any]], list[dict[str, Any]]]:
    manifest = read_json(CAPTURES / "manifest.json")
    runs = [row for row in manifest.get("runs", []) if row.get("lane") == "native"]
    require(len(runs) == 36, f"expected 36 native runs, got {len(runs)}")
    data: dict[tuple[str, int, int, str], dict[str, Any]] = {}
    processes: list[dict[str, Any]] = []
    for process_order, row in enumerate(runs):
        case = row.get("case")
        variant = row.get("variant")
        cycle = row.get("cycle")
        repeat = row.get("repeat")
        require(case in CASES and variant in VARIANTS, "invalid native run identity")
        require(isinstance(cycle, int) and isinstance(repeat, int), "invalid native run index")
        key = (case, cycle, repeat, variant)
        require(key not in data, f"duplicate native run: {key}")
        output_name = row.get("output")
        require(isinstance(output_name, str) and Path(output_name).name == output_name,
                f"unsafe native output name: {output_name!r}")
        output_path = CAPTURES / output_name
        require(sha(output_path) == row.get("sha256"), f"capture digest changed: {output_name}")
        report = read_json(output_path)
        require(report.get("timing_claim") is True and report.get("operation") == "format",
                f"timing contract changed: {output_name}")
        require(report.get("warmups") == 3 and report.get("samples_requested") == 50,
                f"sample plan changed: {output_name}")
        samples = report.get("samples")
        require(isinstance(samples, list) and len(samples) == 50,
                f"sample count changed: {output_name}")
        expected_sha = report.get("expected_output_sha256")
        expected_inventory = report.get("expected_output_inventory")
        expected_oracle = report.get("expected_oracle")
        times: list[int] = []
        serialized_sample_bytes = 0
        for index, sample in enumerate(samples):
            require(sample.get("index") == index, f"sample index changed: {output_name}:{index}")
            phase = sample.get("phase_ns")
            require(isinstance(phase, dict) and set(phase) == {"whole_ns"},
                    f"sample timing shape changed: {output_name}:{index}")
            value = phase["whole_ns"]
            require(type(value) is int and value > 0, f"invalid whole_ns: {output_name}:{index}")
            require(sample.get("output_sha256") == expected_sha
                    and sample.get("output_inventory") == expected_inventory,
                    f"output identity changed: {output_name}:{index}")
            oracle = sample.get("oracle")
            require(oracle == expected_oracle and oracle.get("oracle_ok") is True,
                    f"oracle changed or failed: {output_name}:{index}")
            times.append(value)
            serialized_sample_bytes += len(json.dumps(
                sample, separators=(",", ":"), ensure_ascii=True).encode("utf-8"))
        data[key] = {"times": times, "process_order": process_order, "seconds": row.get("seconds")}
        processes.append({
            "case": case,
            "cycle": cycle,
            "repeat": repeat,
            "variant": variant,
            "process_order": process_order,
            "wall_seconds": row.get("seconds"),
            "serialized_sample_bytes": serialized_sample_bytes,
        })
    require(len(data) == 36, "native process identity census is incomplete")
    return data, processes


def process_windows(times: list[int]) -> dict[str, dict[str, Any]]:
    return {
        name: {"p50_ns": median(times[start:end]),
               "mean_ns": math.fsum(times[start:end]) / (end - start)}
        for name, (start, end) in WINDOWS.items()
    }


def lag_values(times: list[int]) -> dict[str, float]:
    return {str(lag): pearson(times[:-lag], times[lag:]) for lag in LAGS}


def linear_slope(times: list[int]) -> float:
    x = list(range(len(times)))
    x_mean = statistics.fmean(x)
    y_mean = statistics.fmean(times)
    denominator = math.fsum((value - x_mean) ** 2 for value in x)
    return math.fsum((i - x_mean) * (value - y_mean) for i, value in enumerate(times)) / denominator


def grouped_summary(values: list[float]) -> dict[str, Any]:
    return summary(values)


def main() -> None:
    packet = verify_packet()
    lifecycle = source_lifecycle()
    data, processes = load_native()

    for process in processes:
        key = (process["case"], process["cycle"], process["repeat"], process["variant"])
        times = data[key]["times"]
        process["windows"] = process_windows(times)
        process["last10_vs_first10_p50_percent"] = percent(
            median(times[:10]), median(times[40:50]))
        process["linear_slope_ns_per_index"] = linear_slope(times)
        process["lag_pearson"] = lag_values(times)
        timed_sum = math.fsum(times) / 1_000_000_000
        wall = process["wall_seconds"]
        require(isinstance(wall, (int, float)) and wall > 0, "invalid manifest wall time")
        process["timed_sum_seconds"] = timed_sum
        process["wall_minus_timed_seconds"] = wall - timed_sum
        process["wall_minus_timed_percent"] = 100.0 * (wall - timed_sum) / wall

    pairs: list[dict[str, Any]] = []
    for case in CASES:
        for cycle in range(3):
            for repeat in range(3):
                baseline = data[(case, cycle, repeat, "baseline")]
                candidate = data[(case, cycle, repeat, "candidate")]
                windows = {}
                for name, (start, end) in WINDOWS.items():
                    before = baseline["times"][start:end]
                    after = candidate["times"][start:end]
                    windows[name] = {
                        "p50_percent": percent(median(before), median(after)),
                        "mean_percent": percent(statistics.fmean(before), statistics.fmean(after)),
                    }
                pairs.append({
                    "case": case,
                    "cycle": cycle,
                    "repeat": repeat,
                    "first_variant": (
                        "baseline" if baseline["process_order"] < candidate["process_order"]
                        else "candidate"
                    ),
                    "windows": windows,
                })

    window_groups: dict[str, dict[str, dict[str, Any]]] = {}
    for case in CASES:
        window_groups[case] = {}
        for order in ("all", "baseline", "candidate"):
            selected = [pair for pair in pairs
                        if pair["case"] == case
                        and (order == "all" or pair["first_variant"] == order)]
            window_groups[case][order] = {}
            for window in WINDOWS:
                window_groups[case][order][window] = {
                    metric: grouped_summary([pair["windows"][window][metric] for pair in selected])
                    for metric in ("p50_percent", "mean_percent")
                }

    drift: dict[str, dict[str, Any]] = {}
    lags: dict[str, dict[str, Any]] = {}
    slopes: dict[str, dict[str, Any]] = {}
    cadence: dict[str, dict[str, Any]] = {}
    for case in CASES:
        drift[case] = {}
        lags[case] = {}
        slopes[case] = {}
        cadence[case] = {}
        for variant in VARIANTS:
            selected = [process for process in processes
                        if process["case"] == case and process["variant"] == variant]
            drift[case][variant] = summary(
                [process["last10_vs_first10_p50_percent"] for process in selected])
            slopes[case][variant] = summary(
                [process["linear_slope_ns_per_index"] for process in selected])
            lags[case][variant] = {
                lag: summary([process["lag_pearson"][lag] for process in selected])
                for lag in (str(value) for value in LAGS)
            }
            cadence[case][variant] = {
                metric: summary([process[metric] for process in selected])
                for metric in ("wall_seconds", "timed_sum_seconds",
                               "wall_minus_timed_seconds", "wall_minus_timed_percent")
            }

    paired_index_bins: dict[str, dict[str, Any]] = {}
    for case in CASES:
        paired_index_bins[case] = {}
        for name, (start, end) in INDEX_BINS.items():
            values = []
            for cycle in range(3):
                for repeat in range(3):
                    before = data[(case, cycle, repeat, "baseline")]["times"]
                    after = data[(case, cycle, repeat, "candidate")]["times"]
                    values.extend(percent(before[index], after[index])
                                  for index in range(start, end))
            paired_index_bins[case][name] = summary(values)

    # Independently confirm that the original full-window paired medians were
    # not changed by this post-hoc read.  This is a consistency check, not a
    # new acceptance rule.
    original = read_json(PACKET / "analysis.json")
    full_reproduction = {}
    for case in CASES:
        original_row = next(row for row in original["native_case_summaries"]
                            if row["case"] == case)
        calculated = window_groups[case]["all"]["all"]
        full_reproduction[case] = {}
        for metric in ("p50", "mean"):
            value = calculated[metric + "_percent"]["median"]
            expected = original_row["metrics"][metric]["median_percent"]
            require(abs(value - expected) < 1e-10,
                    f"0735 full-window {case}/{metric} no longer reproduces")
            full_reproduction[case][metric] = {"calculated": value, "0735": expected}

    result = {
        "status": "passed",
        "scope": "Independent post-hoc descriptive audit of sealed 0735 native samples; original disposition is unchanged.",
        "execution": "read-only Python; no Cargo, Rust, native, allocator, or profiler execution",
        "packet": packet,
        "captures": {"native_processes": 36, "process_pairs_per_case": 9,
                     "samples_per_process": 50, "warmups_per_process": 3,
                     "sample_count": 1800},
        "quantile": "midpoint median within each process; ordinary median across the nine paired process percentages",
        "windows": {name: [start, end] for name, (start, end) in WINDOWS.items()},
        "index_bins": {name: [start, end] for name, (start, end) in INDEX_BINS.items()},
        "source_lifecycle": lifecycle,
        "full_window_reproduction": full_reproduction,
        "processes": processes,
        "pairs": pairs,
        "window_groups": window_groups,
        "within_process_drift": drift,
        "within_process_lag_pearson": lags,
        "within_process_linear_slope": slopes,
        "paired_sample_index_bins": paired_index_bins,
        "manifest_wall_vs_timed_sum": cadence,
        "interpretation_limits": [
            "Warmup timings are not retained, so cold-to-warm transition is not directly observed.",
            "Manifest wall time minus timed sums includes expected-output preparation, oracle controls, output validation, JSON serialization, and process overhead; it is not an oracle-only duration.",
            "Lag correlations, window changes, and process-order strata are descriptive correlations, not causal allocator or cache evidence.",
            "Samples within a process and overlapping windows are not independent process repeats.",
        ],
    }
    OUTPUT.write_text(json.dumps(result, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": "passed", "output": str(OUTPUT),
                      "processes": 36, "samples": 1800, "pairs": 18}))


if __name__ == "__main__":
    try:
        main()
    except (Failure, OSError, KeyError, TypeError, ValueError) as error:
        raise SystemExit(f"FAIL: {error}") from error
