"""Independent reconstruction and raw receipt audit for the 0816 packet.

This reader deliberately does not import the derived analysis. It rebuilds the
deterministic 0786 corpus, checks every retained report and receipt, and
derives paired width/source curves directly from native JSON samples.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import statistics
import struct
import sys
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
SEED = 816816
RESAMPLES = 10_000
WIDTHS = (1, 2, 4, 8)
SHAPES = ("large", "mixed")
CASE_FAMILIES = (("large", "fresh"), ("mixed", "fresh"), ("large", "primed"))
LANE_SAMPLES = {"qualification": 1, "native": 30, "observer": 2}
LANE_REPORTS = {"qualification": 72, "native": 432, "observer": 144}
NORMAL_SCHEMA = "litchi.execution-baseline.v1"
RANGE_SCHEMA = "litchi.execution-range-baseline.v1"


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    return sha256_bytes(path.read_bytes())


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON {path}: {error}")


def resolve_artifact(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label} is not an artifact")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    candidates = [Path(raw), P / raw]
    marker = "/docs/performance/results/change-0816/"
    if marker in raw:
        candidates.append(P / raw.split(marker, 1)[1])
    candidates.append(P / Path(raw).name)
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            path = candidate.resolve()
            require(path.is_relative_to(P.resolve()), f"{label} escaped packet: {raw}")
            require(path.stat().st_size == value.get("bytes"), f"{label} byte count changed")
            require(sha256_file(path) == value.get("sha256"), f"{label} digest changed")
            return path
    fail(f"missing artifact {label}: {raw}")


def payload(shape: str, index: int) -> bytes:
    size = 4 * 1024 if shape == "mixed" and index == 31 else 256 * 1024
    label = f"litchi-0786-member-{index:02}-".encode()
    offset = index * 11 if index * 11 <= 255 else 0
    return bytes(
        (label[position % len(label)] + (position % 97) * 3 + offset) % 256
        for position in range(size)
    )


def corpus(shape: str) -> dict[str, Any]:
    members = [payload(shape, index) for index in range(32)]
    sequence = hashlib.sha256()
    for index, item in enumerate(members):
        sequence.update(struct.pack("<Q", index))
        sequence.update(item)
    return {
        "shape": shape,
        "members": [sha256_bytes(item) for item in members],
        "sizes": [len(item) for item in members],
        "sequence": sequence.hexdigest(),
        "logical_bytes": sum(len(item) for item in members),
    }


CORPORA = {shape: corpus(shape) for shape in SHAPES}


def expected_rows(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for block, order in enumerate(plan[lane]["orders"]):
        cases = plan["cases"] if order == "forward" else list(reversed(plan["cases"]))
        rows.extend({"lane": lane, "block": block, **case} for case in cases)
    return rows


def nearest_rank(values: list[float], quantile: float) -> float:
    ordered = sorted(values)
    require(ordered, "empty nearest-rank input")
    index = max(0, min(len(ordered) - 1, math.ceil(quantile * len(ordered)) - 1))
    return ordered[index]


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(values, "empty bootstrap input")
    rng = random.Random(SEED)
    estimates = [statistics.median(rng.choice(values) for _ in values)
                 for _ in range(RESAMPLES)]
    return {
        "estimate": statistics.median(values),
        "lower": nearest_rank(estimates, 0.025),
        "upper": nearest_rank(estimates, 0.975),
        "seed": SEED, "resamples": RESAMPLES, "confidence": 0.95,
        "statistic": "median of six paired block ratios", "block_ratios": values,
    }


def check_resources(sample: dict[str, Any], case: dict[str, Any], label: str) -> None:
    resources = sample.get("resources")
    require(isinstance(resources, dict), f"{label} resources missing")
    limits = resources.get("limits")
    require(isinstance(limits, dict), f"{label} resource limits missing")
    require(limits.get("workers") == case["workers"]
            and limits.get("io_concurrency") == case["workers"]
            and limits.get("cpu_tasks") == 1_000_000,
            f"{label} resource limits changed")
    snapshots = [resources.get(name) for name in
                 ("before_operation", "after_operation", "after_drop")]
    require(all(isinstance(item, dict) for item in snapshots),
            f"{label} resource snapshots missing")
    for point, snapshot in zip(("before", "after", "drop"), snapshots):
        assert isinstance(snapshot, dict)
        for key in ("workers", "io_concurrency", "cpu_tasks"):
            value = snapshot.get(key)
            require(isinstance(value, int) and 0 <= value <= limits[key],
                    f"{label} {point} {key} exceeds bound")
    require(snapshots[2]["workers"] == 0 and snapshots[2]["io_concurrency"] == 0,
            f"{label} worker/I/O permits leaked")
    require(resources.get("worker_and_io_released") is True
            and resources.get("cpu_tasks_within_limit") is True,
            f"{label} release witness is false")
    expected_before_cpu = 32 if case["state"] == "primed" else 0
    require(snapshots[0]["cpu_tasks"] == expected_before_cpu
            and snapshots[1]["cpu_tasks"] == expected_before_cpu + 32
            and snapshots[2]["cpu_tasks"] == expected_before_cpu + 32,
            f"{label} CPU-task accounting changed")


def source_tuple(metrics: dict[str, Any], label: str) -> tuple[Any, ...]:
    fields = ("logical_calls", "requested_bytes", "returned_bytes", "short_reads",
              "active_reads_after_operation", "request_size_histogram")
    for field in fields[:-1]:
        require(isinstance(metrics.get(field), int) and metrics[field] >= 0,
                f"{label} source metric {field} is invalid")
    histogram = metrics.get("request_size_histogram")
    require(isinstance(histogram, list) and all(isinstance(item, int) and item >= 0
                                                for item in histogram),
            f"{label} request histogram is invalid")
    require(sum(histogram) == metrics["logical_calls"],
            f"{label} request histogram does not conserve calls")
    require(metrics["active_reads_after_operation"] == 0,
            f"{label} active source reads were not released")
    return tuple(metrics[field] if field != "request_size_histogram"
                 else tuple(histogram) for field in fields)


def audit_report(report: dict[str, Any], receipt: dict[str, Any], lane: str,
                 sample_label: str) -> tuple[Any, ...] | None:
    case = receipt
    expected_schema = (RANGE_SCHEMA if case["source_max_read_bytes"] or case["source_delay_us"]
                       else NORMAL_SCHEMA)
    require(report.get("schema") == expected_schema, f"{sample_label} schema changed")
    config = report.get("config")
    require(isinstance(config, dict), f"{sample_label} config missing")
    for key in ("route", "shape", "state", "task_floor", "workers",
                "source_max_read_bytes", "source_delay_us"):
        require(config.get(key) == case[key], f"{sample_label} config {key} changed")
    require(config.get("samples") == LANE_SAMPLES[lane]
            and config.get("warmup") == {"native": 3, "observer": 0,
                                          "qualification": 0}[lane]
            and config.get("aggregate_parallel_bytes") == 65536
            and config.get("cpu_task_limit") == 1_000_000,
            f"{sample_label} benchmark configuration changed")
    shape = case["shape"]
    expected = CORPORA[shape]
    corpus_value = report.get("corpus")
    require(isinstance(corpus_value, dict), f"{sample_label} corpus missing")
    require(corpus_value.get("shape") == shape
            and corpus_value.get("selected_payload_member_count") == 32,
            f"{sample_label} corpus shape changed")
    members = corpus_value.get("members")
    require(isinstance(members, list) and len(members) == 32,
            f"{sample_label} corpus member count changed")
    require([member.get("sha256") for member in members] == expected["members"],
            f"{sample_label} payload digest/order changed")
    for index, member in enumerate(members):
        require(member.get("index") == index
                and member.get("bytes") == expected["sizes"][index],
                f"{sample_label} member metadata changed")
    require(isinstance(corpus_value.get("opc_sha256"), str)
            and isinstance(corpus_value.get("cfb_sha256"), str),
            f"{sample_label} container identity missing")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == LANE_SAMPLES[lane],
            f"{sample_label} sample count changed")
    source_rows: list[tuple[Any, ...]] = []
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"{sample_label} sample {index} malformed")
        verification = sample.get("verification")
        require(isinstance(verification, dict)
                and verification.get("ordered") is True
                and verification.get("all_member_sha256_match") is True
                and verification.get("members") == 32
                and verification.get("logical_bytes") == expected["logical_bytes"]
                and verification.get("sequence_sha256") == expected["sequence"],
                f"{sample_label} byte verification changed")
        check_resources(sample, case, f"{sample_label} sample {index}")
        metrics = sample.get("source_metrics")
        if lane == "native":
            require(isinstance(metrics, dict)
                    and metrics.get("availability") == "unavailable-normal-build"
                    and all(metrics.get(key) is None for key in (
                        "logical_calls", "requested_bytes", "returned_bytes", "short_reads",
                        "active_reads_after_operation", "max_simultaneous_reads",
                        "request_size_histogram")),
                    f"{sample_label} native source counters are present")
        else:
            require(isinstance(metrics, dict)
                    and metrics.get("availability") == "source-metrics-feature",
                    f"{sample_label} observer source counters are missing")
            metric_row = source_tuple(metrics, f"{sample_label} sample {index}")
            require(metrics["max_simultaneous_reads"] <= case["workers"],
                    f"{sample_label} max active reads exceed width")
            if case["route"] == "cfb":
                sizes = expected["sizes"]
                cap = case["source_max_read_bytes"]
                calls = sum(math.ceil(size / cap) if cap else 1 for size in sizes)
                requested = sum(
                    sum(range(size, 0, -cap)) if cap else size for size in sizes
                )
                require(metric_row[:4] == (
                    calls, requested, expected["logical_bytes"], calls - 32
                ), f"{sample_label} CFB source counters changed")
            elif case["state"] == "fresh":
                expected_parts = 146041 if shape == "large" else 143782
                require(metric_row[:4] == (64, expected_parts, expected_parts, 0),
                        f"{sample_label} Parts source counters changed")
            if case["route"] == "parts" and case["state"] == "primed":
                require(metric_row[:4] == (0, 0, 0, 0)
                        and metric_row[5] == (0,) * len(metric_row[5]),
                        f"{sample_label} primed Parts performed source reads")
            source_rows.append(metric_row)
    return None if lane == "native" else source_rows[0]


def main() -> int:
    plan = read_json(P / "plan.json")
    require(isinstance(plan, dict)
            and plan.get("schema") == "litchi.performance.0816.plan.v1",
            "plan schema changed")
    source_observations: dict[tuple[Any, ...], tuple[Any, ...]] = {}
    identities: dict[str, tuple[str, str]] = {}
    report_count = 0
    sample_count = 0
    p50: dict[tuple[Any, ...], float] = {}
    for lane in ("qualification", "native", "observer"):
        receipts = read_json(P / lane / "receipts.json")
        require(isinstance(receipts, list) and len(receipts) == LANE_REPORTS[lane],
                f"{lane} receipt cardinality changed")
        wanted = expected_rows(plan, lane)
        require(len(wanted) == len(receipts), f"{lane} schedule cardinality changed")
        for index, (receipt, expected_case) in enumerate(zip(receipts, wanted)):
            require(isinstance(receipt, dict), f"{lane} receipt {index} malformed")
            for key, value in expected_case.items():
                require(receipt.get(key) == value, f"{lane} receipt {index} {key} changed")
            report_path = resolve_artifact(receipt.get("report"), f"{lane} report {index}")
            report = read_json(report_path)
            source_row = audit_report(report, receipt, lane, f"{lane} report {index}")
            shape = receipt["shape"]
            ids = (report["corpus"]["opc_sha256"], report["corpus"]["cfb_sha256"])
            previous = identities.setdefault(shape, ids)
            require(previous == ids, f"{lane} container identity changed for {shape}")
            if source_row is not None:
                source_key = (receipt["route"], shape, receipt["state"],
                              receipt["source_max_read_bytes"], receipt["source_delay_us"])
                old = source_observations.setdefault(source_key, source_row)
                require(old == source_row, f"source byte counters changed for {source_key}")
            if lane == "native":
                p50_key = (receipt["route"], shape, receipt["state"], receipt["task_floor"],
                           receipt["source_max_read_bytes"], receipt["source_delay_us"],
                           receipt["workers"], receipt["block"])
                p50[p50_key] = nearest_rank(
                    [float(item["wall_ns"]) for item in report["samples"]], 0.50)
            report_count += 1
            sample_count += len(report["samples"])
    require(report_count == 648 and sample_count == 13_320,
            "raw report aggregate cardinality changed")

    # A delayed source is the same capped source with a sleep inserted. The
    # sleep may change timing, but it must not alter calls, request sizes or
    # returned bytes. Local and capped reads must return the same bytes even
    # though capped reads are allowed to produce short reads.
    for route in ("cfb", "parts"):
        for shape, state in CASE_FAMILIES:
                local = source_observations[(route, shape, state, 0, 0)]
                capped = source_observations[(route, shape, state, 65536, 0)]
                delayed = source_observations[(route, shape, state, 65536, 250)]
                require(local[2] == capped[2] == delayed[2],
                        f"returned source bytes changed across controls: {route}/{shape}/{state}")
                require(capped == delayed,
                        f"delayed source counters changed: {route}/{shape}/{state}")
                if route == "parts" and state == "primed":
                    require(local[:4] == capped[:4] == delayed[:4] == (0, 0, 0, 0),
                            f"primed Parts source control is nonzero: {route}/{shape}")

    rows: list[dict[str, Any]] = []
    families = sorted({key[:6] for key in p50})
    require(len(families) == 18, f"native family cardinality changed: {len(families)}")
    for family in families:
        route, shape, state, floor, max_read, delay = family
        for width in WIDTHS:
            ratios = [
                p50[(route, shape, state, floor, max_read, delay, 1, block)]
                / p50[(route, shape, state, floor, max_read, delay, width, block)]
                for block in range(6)
            ]
            boot = bootstrap(ratios)
            throughput = [
                CORPORA[shape]["logical_bytes"]
                / p50[(route, shape, state, floor, max_read, delay, width, block)]
                * 1_000_000_000.0
                for block in range(6)
            ]
            rows.append({
                "case": [route, shape, state, floor, max_read, delay, width],
                "paired_speedups": ratios,
                "median": statistics.median(ratios),
                "ci95": [boot["lower"], boot["upper"]],
                "block_payload_throughput_bytes_s": throughput,
                "payload_throughput_bytes_s": statistics.median(throughput),
            })
    require(len(rows) == 72, "raw paired curve cardinality changed")
    source_control_rows = []
    for route in ("cfb", "parts"):
        for shape, state in CASE_FAMILIES:
                for width in WIDTHS:
                    cap_local = [
                        p50[(route, shape, state, 65536, 65536, 0, width, block)]
                        / p50[(route, shape, state, 65536, 0, 0, width, block)]
                        for block in range(6)
                    ]
                    delay_cap = [
                        p50[(route, shape, state, 65536, 65536, 250, width, block)]
                        / p50[(route, shape, state, 65536, 65536, 0, width, block)]
                        for block in range(6)
                    ]
                    source_control_rows.append({
                        "case": [route, shape, state, width],
                        "cap_to_local": bootstrap(cap_local),
                        "delay_to_cap": bootstrap(delay_cap),
                    })
    result = {
        "schema": "litchi.performance.0816.raw-audit.v1",
        "reports": report_count,
        "samples": sample_count,
        "independently_reconstructed_payloads": len(SHAPES) * 32,
        "corpora": CORPORA,
        "container_identities": identities,
        "rows": rows,
        "source_control_rows": source_control_rows,
        "source_metrics": {
            "observations": len(source_observations),
            "delayed_counters_equal_capped": True,
            "returned_bytes_equal_across_controls": True,
            "primed_parts_zero_reads": True,
        },
    }
    encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
    output = P / "raw-audit.json"
    if "--check" in sys.argv:
        require(output.read_text() == encoded, "raw-audit.json does not replay byte-for-byte")
    else:
        require(not output.exists(), "refusing to overwrite retained raw-audit.json")
        output.write_text(encoded)
    print("independent 0816 raw audit PASS: 648 reports, 13,320 samples, 72 paired curves")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AssertionError, KeyError, OSError, TypeError, ValueError) as error:
        print(f"0816 raw audit failed: {error}", file=sys.stderr)
        raise SystemExit(1)
