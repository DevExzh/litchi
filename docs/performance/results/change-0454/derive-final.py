#!/usr/bin/env python3
"""Derive individual lifecycle percentiles and descriptive ABBA comparisons."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import random
import re
import statistics
from typing import Any


ROOT = Path(__file__).resolve().parent
BOOTSTRAP_RESAMPLES = 10_000
REVIEW_PERCENT = 5.0


def load(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def sha(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def quantile(values: list[float], q: float) -> float:
    if not values:
        raise ValueError("quantile requires at least one sample")
    ordered = sorted(values)
    position = (len(ordered) - 1) * q
    lower = int(position)
    upper = min(lower + 1, len(ordered) - 1)
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def bootstrap_median_ci(values: list[float], seed: int) -> list[float]:
    rng = random.Random(seed)
    medians = [statistics.median(rng.choices(values, k=len(values))) for _ in range(BOOTSTRAP_RESAMPLES)]
    return [quantile(medians, 0.025), quantile(medians, 0.975)]


def distribution(values_ns: list[int], seed: int) -> dict[str, Any]:
    values_ms = [value / 1_000_000 for value in values_ns]
    return {
        "samples": len(values_ms),
        "p50_ms": statistics.median(values_ms),
        "p95_ms": quantile(values_ms, 0.95),
        "p99_ms": quantile(values_ms, 0.99),
        "max_ms": max(values_ms),
        "median_bootstrap_ci95_ms": bootstrap_median_ci(values_ms, seed),
    }


RESOURCE_KEYS: dict[str, str] = {
    "Maximum resident set size (kbytes)": "max_rss_kib",
    "Minor (reclaiming a frame) page faults": "minor_page_faults",
    "Major (requiring I/O) page faults": "major_page_faults",
    "File system inputs": "filesystem_inputs",
    "File system outputs": "filesystem_outputs",
    "Voluntary context switches": "voluntary_context_switches",
    "Involuntary context switches": "involuntary_context_switches",
}
TIME_KEYS = {
    "User time (seconds)": "user_time_ns",
    "System time (seconds)": "system_time_ns",
}


def parse_resource(path: Path) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if ":" not in line:
            continue
        key, value = line.split(":", 1)
        key, value = key.strip(), value.strip()
        if key in RESOURCE_KEYS:
            match = re.fullmatch(r"\d+", value)
            if not match:
                raise ValueError(f"resource counter is not an integer: {key}: {value}")
            result[RESOURCE_KEYS[key]] = int(value)
        elif key in TIME_KEYS:
            # GNU time emits a decimal number of seconds.  Preserve it as an
            # integer nanosecond counter so derived JSON has no floating point
            # resource ambiguity.
            seconds = float(value)
            if seconds < 0:
                raise ValueError(f"resource time is negative: {key}")
            result[TIME_KEYS[key]] = round(seconds * 1_000_000_000)
    for key in RESOURCE_KEYS.values():
        result.setdefault(key, None)
    for key in TIME_KEYS.values():
        result.setdefault(key, None)
    result["path"] = str(path.relative_to(ROOT))
    result["scope"] = "one fresh child process; includes setup and all retained samples"
    return result


READ_FIELDS = (
    "logical_calls",
    "requested_bytes",
    "returned_bytes",
    "short_reads",
    "delayed_calls",
    "transfer_paced_calls",
    "transfer_delay_ns",
)
BUDGET_FIELDS = (
    "memory_used",
    "objects_used",
    "depth_used",
    "input_bytes_used",
    "output_bytes_used",
    "work_used",
)


def read_summary(point: dict[str, Any]) -> dict[str, Any] | None:
    if point.get("availability") == "unavailable":
        return None
    if point.get("delta") is not None:
        point = point["delta"]
    return {field: point.get(field) for field in READ_FIELDS} | {
        "request_size_counts": point.get("request_size_counts"),
        "min_request_bytes": point.get("min_request_bytes"),
        "max_request_bytes": point.get("max_request_bytes"),
    }


def budget_summary(point: dict[str, Any]) -> dict[str, Any]:
    return {field: point[field] for field in BUDGET_FIELDS}


def external_budget_summary(point: dict[str, Any]) -> dict[str, Any]:
    return {
        "memory": point["memory"],
        "objects": point["objects"],
        "depth": point["depth"],
        "input_bytes": point["input_bytes"],
        "output_bytes": point["output_bytes"],
        "work": point["work"],
    }


def validate_sample_counters(
    report: dict[str, Any],
    protocol: dict[str, Any],
    *,
    warmup_key: str,
    label: str,
) -> list[dict[str, Any]]:
    """Require the report counters and retained rows to describe one lane."""
    samples = report.get("samples_raw")
    if not isinstance(samples, list):
        raise ValueError(f"{label} report samples_raw is not a list")
    if report.get("samples") != protocol["samples"]:
        raise ValueError(f"{label} report sample count differs from the protocol")
    if report.get(warmup_key) != protocol["warmups"]:
        raise ValueError(f"{label} report warmup count differs from the protocol")
    if len(samples) != report["samples"]:
        raise ValueError(f"{label} report samples_raw length differs from samples")
    for number, sample in enumerate(samples):
        timing = sample.get("timings", sample)
        try:
            api_sum = timing["api_sum_ns"]
            components = [timing[name] for name in ("open_ns", "plan_ns", "publication_ns")]
        except (KeyError, TypeError) as error:
            raise ValueError(f"{label} sample {number} is missing timing counters") from error
        if any(type(value) is not int or value <= 0 for value in [*components, api_sum]):
            raise ValueError(f"{label} sample {number} has a non-positive timing counter")
        if api_sum != sum(components):
            raise ValueError(f"{label} sample {number} api_sum_ns is not the phase sum")
    return samples


def provider_io_work(report: dict[str, Any]) -> dict[str, Any]:
    first = report["samples_raw"][0]
    phases = {phase["label"]: phase for phase in first["phases"]}
    result: dict[str, Any] = {"phases": {}}
    for label in ("opened", "planned", "published"):
        phase = phases[label]
        result["phases"][label] = {
            "source_reads": read_summary(phase["source_reads"]),
            "destination_reads": read_summary(phase["destination_reads"]),
            "source_budget": budget_summary(phase["source_budget"]),
            "destination_budget": budget_summary(phase["destination_budget"]),
        }
    result["final_source_budget"] = budget_summary(phases["drop_sink"]["source_budget"])
    result["final_destination_budget"] = budget_summary(phases["drop_sink"]["destination_budget"])
    # Work/read counters are expected to be deterministic for this harness.
    for sample in report["samples_raw"][1:]:
        sample_phases = {phase["label"]: phase for phase in sample["phases"]}
        sample_value = {
            "phases": {
                label: {
                    "source_reads": read_summary(sample_phases[label]["source_reads"]),
                    "destination_reads": read_summary(sample_phases[label]["destination_reads"]),
                    "source_budget": budget_summary(sample_phases[label]["source_budget"]),
                    "destination_budget": budget_summary(sample_phases[label]["destination_budget"]),
                }
                for label in ("opened", "planned", "published")
            },
            "final_source_budget": budget_summary(sample_phases["drop_sink"]["source_budget"]),
            "final_destination_budget": budget_summary(sample_phases["drop_sink"]["destination_budget"]),
        }
        if sample_value != result:
            raise ValueError("provider I/O/work counters changed across retained samples")
    return result


def external_io_work(report: dict[str, Any]) -> dict[str, Any]:
    first = report["samples_raw"][0]

    def phase_value(label: str) -> dict[str, Any]:
        phase = first[label]
        result = {owner: read_summary(phase[owner]) for owner in ("source", "destination")}
        if label == "planned":
            result["source_budget"] = external_budget_summary(phase["source_budget"])
            result["destination_budget"] = external_budget_summary(phase["destination_budget"])
        return result

    result = {
        "phases": {label: phase_value(label) for label in ("opened", "planned", "published")},
        "final_source_budget": external_budget_summary(first["final_source_budget"]),
        "final_destination_budget": external_budget_summary(first["final_destination_budget"]),
    }
    # External reports expose monotonic snapshots rather than the provider
    # journal's checked deltas.  Preserve the first report's full counters and
    # require the output/semantic identities to stay stable per the oracle.
    for sample in report["samples_raw"][1:]:
        if sample["output_sha256"] != first["output_sha256"] or sample["output_bytes"] != first["output_bytes"]:
            raise ValueError("external output identity changed across retained samples")
    return result


def timing_row(report: dict[str, Any], *, seed: int) -> dict[str, Any]:
    sample_timing = report["samples_raw"][0].get("timings", report["samples_raw"][0])
    names = (
        "open_source_ns",
        "open_destination_ns",
        "open_ns",
        "plan_ns",
        "publication_ns",
        "api_sum_ns",
    )
    if "open_source_ns" not in sample_timing:
        names = ("open_ns", "plan_ns", "publication_ns", "api_sum_ns")
    def value(sample: dict[str, Any], name: str) -> int:
        return sample.get("timings", sample)[name]
    return {
        name.removesuffix("_ns"): distribution(
            [value(sample, name) for sample in report["samples_raw"]], seed + index
        )
        for index, name in enumerate(names)
    }


def provider_row(index: int, lane: dict[str, Any], protocol: dict[str, Any]) -> dict[str, Any]:
    report_path = ROOT / "provider-runs" / str(index) / "report.json"
    receipt_path = ROOT / "provider-runs" / str(index) / "receipt.json"
    if not report_path.is_file() or not receipt_path.is_file():
        raise FileNotFoundError(f"missing provider capture lane {index}")
    receipt = load(receipt_path)
    if receipt["status"] != "pass" or receipt["lane"] != index or receipt.get("pilot") is not False:
        raise ValueError(f"provider receipt {index} is not a passing lane receipt")
    report = load(report_path)
    if report["schema"] != "pptx_provider_lifecycle_v1":
        raise ValueError(f"unexpected provider report schema in lane {index}")
    if report["provider"] != lane["provider"] or report["corpus"] != lane["corpus"]:
        raise ValueError(f"provider lane {index} identity differs from the protocol")
    samples = validate_sample_counters(
        report, protocol, warmup_key="warmup", label=f"provider lane {index}"
    )
    return {
        "lane": index,
        **lane,
        "report": str(report_path.relative_to(ROOT)),
        "receipt": str(receipt_path.relative_to(ROOT)),
        "output": {
            "sha256": report["expected_output_sha256"],
            "bytes": report["expected_output_bytes"],
        },
        "timings": timing_row(report, seed=454_000 + index * 10),
        "process_resource": parse_resource(ROOT / "provider-runs" / str(index) / "resource.log"),
        "io_work": provider_io_work(report),
        "scope": report["timing_scope"],
        "samples": len(samples),
        "warmups": report["warmup"],
    }


def external_row(index: int, lane: dict[str, Any], protocol: dict[str, Any]) -> dict[str, Any]:
    report_path = ROOT / "external-runs" / str(index) / "report.json"
    receipt_path = ROOT / "external-runs" / str(index) / "receipt.json"
    if not report_path.is_file() or not receipt_path.is_file():
        raise FileNotFoundError(f"missing external capture lane {index}")
    receipt = load(receipt_path)
    if receipt["status"] != "pass" or receipt["lane"] != index or receipt.get("pilot") is not False:
        raise ValueError(f"external receipt {index} is not a passing lane receipt")
    report = load(report_path)
    if report["schema"] != "pptx-external-cross-copy-v1":
        raise ValueError(f"unexpected external report schema in lane {index}")
    if report["provider"] != lane["provider"]:
        raise ValueError(f"external lane {index} provider differs from the protocol")
    samples = validate_sample_counters(
        report, protocol, warmup_key="warmups", label=f"external lane {index}"
    )
    first = samples[0]
    return {
        "lane": index,
        **lane,
        "report": str(report_path.relative_to(ROOT)),
        "receipt": str(receipt_path.relative_to(ROOT)),
        "fixture": {
            "bytes": report["fixture_bytes"],
            "sha256": report["fixture_sha256"],
        },
        "output": {
            "sha256": first["output_sha256"],
            "bytes": first["output_bytes"],
            "oracle": first["oracle"],
        },
        "timings": timing_row(report, seed=454_100 + index * 10),
        "process_resource": parse_resource(ROOT / "external-runs" / str(index) / "resource.log"),
        "io_work": external_io_work(report),
        "scope": report["timing_scope"],
        "range_scope": report["range_scope"],
        "samples": len(samples),
        "warmups": report["warmups"],
    }


def percent_change(before: float, after: float) -> float | None:
    if before == 0:
        return None if after else 0.0
    return (after - before) / abs(before) * 100


def compare_provider_rows(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    groups: dict[tuple[str, str, str], dict[str, dict[str, Any]]] = {}
    for row in rows:
        groups.setdefault((row["provider"], row["corpus"], row["repeat"]), {})[row["build"]] = row
    pairs: list[dict[str, Any]] = []
    flags: list[dict[str, Any]] = []
    for key in sorted(groups):
        group = groups[key]
        if set(group) != {"baseline", "candidate"}:
            raise ValueError(f"incomplete provider pair {key}")
        baseline, candidate = group["baseline"], group["candidate"]
        metrics: dict[str, Any] = {}
        for phase, values in baseline["timings"].items():
            metrics[phase] = {}
            for percentile in ("p50_ms", "p95_ms", "p99_ms"):
                before = values[percentile]
                after = candidate["timings"][phase][percentile]
                change = percent_change(before, after)
                metrics[phase][percentile] = {
                    "baseline": before,
                    "candidate": after,
                    "percent_change": change,
                }
                if change is not None and abs(change) > REVIEW_PERCENT:
                    flags.append(
                        {
                            "kind": "provider_timing",
                            "provider": key[0],
                            "corpus": key[1],
                            "repeat": key[2],
                            "metric": f"{phase}.{percentile}",
                            "percent_change": change,
                        }
                    )
        for field in ("max_rss_kib", "minor_page_faults", "major_page_faults"):
            before = baseline["process_resource"].get(field)
            after = candidate["process_resource"].get(field)
            if before is None or after is None:
                continue
            change = percent_change(float(before), float(after))
            metrics[f"resource.{field}"] = {
                "baseline": before,
                "candidate": after,
                "percent_change": change,
            }
            if change is not None and abs(change) > REVIEW_PERCENT:
                flags.append(
                    {
                        "kind": "provider_resource",
                        "provider": key[0],
                        "corpus": key[1],
                        "repeat": key[2],
                        "metric": f"resource.{field}",
                        "percent_change": change,
                    }
                )
        if baseline["io_work"] != candidate["io_work"]:
            flags.append(
                {
                    "kind": "provider_io_work",
                    "provider": key[0],
                    "corpus": key[1],
                    "repeat": key[2],
                    "metric": "io_work",
                    "percent_change": None,
                    "detail": "logical read or budget work identity changed",
                }
            )
        pairs.append(
            {
                "provider": key[0],
                "corpus": key[1],
                "repeat": key[2],
                "baseline_lane": baseline["lane"],
                "candidate_lane": candidate["lane"],
                "interpretation": "descriptive matched control; review threshold is not an acceptance gate",
                "metrics": metrics,
            }
        )
    return pairs, flags


def render(measurements: dict[str, Any]) -> str:
    lines = [
        "# Change 0454 performance measurements",
        "",
        "The provider lifecycle table retains the individual p50, p95 and p99 for every",
        "phase and the complete API sum. The two external fixture rows are candidate-only",
        "descriptive evidence: the preserved baseline typed-refuses the unnamed slide, so",
        "no speedup against that refusal is claimed.",
        "",
        "## Provider lifecycle controls",
        "",
        "| Provider | Corpus | Repeat | Baseline API p50/p95/p99 ms | Candidate API p50/p95/p99 ms |",
        "|---|---|---|---:|---:|",
    ]
    for pair in measurements["provider_pairs"]:
        base = next(row for row in measurements["provider_rows"] if row["lane"] == pair["baseline_lane"])
        candidate = next(row for row in measurements["provider_rows"] if row["lane"] == pair["candidate_lane"])
        def fmt(row: dict[str, Any]) -> str:
            timing = row["timings"]["api_sum"]
            return f"{timing['p50_ms']:.3f}/{timing['p95_ms']:.3f}/{timing['p99_ms']:.3f}"
        lines.append(
            f"| {pair['provider']} | {pair['corpus']} | {pair['repeat']} | {fmt(base)} | {fmt(candidate)} |"
        )
    lines += [
        "",
        "## External fixture (candidate-only)",
        "",
        "| Provider | API p50/p95/p99 ms | RSS KiB | Logical read/work evidence |",
        "|---|---:|---:|---|",
    ]
    for row in measurements["external_rows"]:
        timing = row["timings"]["api_sum"]
        resource = row["process_resource"].get("max_rss_kib")
        work = row["io_work"]["phases"]["planned"]
        source_reads = work.get("source_reads", work.get("source"))
        destination_reads = work.get("destination_reads", work.get("destination"))
        source_calls = source_reads["logical_calls"] if source_reads else 0
        destination_calls = destination_reads["logical_calls"] if destination_reads else 0
        lines.append(
            f"| {row['provider']} | {timing['p50_ms']:.3f}/{timing['p95_ms']:.3f}/{timing['p99_ms']:.3f} | "
            f"{resource if resource is not None else 'unavailable'} | planned source/destination calls {source_calls}/{destination_calls} |"
        )
    lines += [
        "",
        "## Scope and limits",
        "",
        f"- Every retained lane has {measurements['protocol']['samples']} samples after {measurements['protocol']['warmups']} warmups in a fresh child, pinned to CPU {measurements['protocol']['cpu']} with one worker.",
        f"- Median intervals use {BOOTSTRAP_RESAMPLES:,} deterministic bootstrap resamples within each lane; p95 and p99 use linear interpolation over the retained samples.",
        "- RSS and page-fault counters come from GNU `/usr/bin/time -v` around the whole child and include setup, fixture construction and retained samples.",
        "- Logical read counters and budget work are copied from the Rust reports and checked for per-lane stability; they do not assert physical device I/O.",
        "- Allocation calls/bytes, live bytes and allocator-region peaks are unavailable from these ordinary binaries. RSS is not an allocation substitute.",
        "- Range lanes use a 64 KiB logical cap, 200 microseconds per request and 25 MiB/s separate sleeps. The external range lane has its own 256-byte/100 microsecond logical adapter settings.",
        f"- Review flags above {REVIEW_PERCENT:.0f}% are retained for individual inspection; no flag is silently converted into an acceptance claim.",
        "",
    ]
    return "\n".join(lines)


def derive() -> dict[str, Any]:
    protocol = load(ROOT / "protocol.json")
    provider_rows = [provider_row(i, lane, protocol) for i, lane in enumerate(protocol["provider_lanes"])]
    external_rows = [external_row(i, lane, protocol) for i, lane in enumerate(protocol["external_lanes"])]
    provider_pairs, flags = compare_provider_rows(provider_rows)
    machine = ROOT / "machine.json"
    measurements = {
        "change": 454,
        "schema": "pptx_change0454_measurements_v1",
        "protocol": {
            "path": "protocol.json",
            "sha256": sha(ROOT / "protocol.json"),
            "cpu": protocol["cpu"],
            "workers": protocol["workers"],
            "samples": protocol["samples"],
            "warmups": protocol["warmups"],
        },
        "machine": {
            "path": "machine.json",
            "sha256": sha(machine) if machine.exists() else None,
            "status": "captured" if machine.exists() else "missing until capture release",
        },
        "provider_rows": provider_rows,
        "external_rows": external_rows,
        "provider_pairs": provider_pairs,
        "review_flags": flags,
        "uncertainty": {
            "bootstrap": "median within each fresh child report",
            "resamples": BOOTSTRAP_RESAMPLES,
            "seed": "454000 + lane*10 + metric index",
            "percentile_interpolation": "linear",
        },
        "allocation": protocol["allocation"],
        "external_comparison": {
            "status": "withheld",
            "speedup_claim": False,
            "reason": protocol["external_fixture"]["comparison"],
            "refusal_evidence": "name-only-outcome-comparison.json",
        },
        "claims": protocol["claims"],
    }
    return measurements


if __name__ == "__main__":
    value = derive()
    (ROOT / "measurements.json").write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")
    (ROOT / "measurements.md").write_text(render(value), encoding="utf-8")
    print(json.dumps({"status": "pass", "provider_rows": len(value["provider_rows"]), "external_rows": len(value["external_rows"]), "review_flags": len(value["review_flags"])}))
