"""Independently analyze the 0835 filesystem route/cache baseline.

This program is deliberately descriptive.  It consumes the content-addressed
capture produced by ``capture.py`` and writes ``analysis.json``; it does not
run the native harness and it does not turn eager/source route ratios into an
optimization claim.
"""

from __future__ import annotations

import argparse
from collections import defaultdict
import hashlib
import json
import math
from pathlib import Path
import random
import statistics
from typing import Any


P = Path(__file__).resolve().parent
CASES = (
    "opc_file_eager_open",
    "opc_file_source_open",
    "opc_file_eager_one_part_atomic_save",
    "opc_file_source_one_part_atomic_save",
    "pptx_file_eager_open_selected_slide_lifecycle",
    "pptx_file_source_open_selected_slide_lifecycle",
)
STATES = ("warm", "cold-verified")
BLOCKS = tuple(range(6))
SAMPLES_PER_ROW = 30
WARMUPS_PER_ROW = 3
EXPECTED_ROWS = len(BLOCKS) * len(CASES) * len(STATES)
EXPECTED_SAMPLES = EXPECTED_ROWS * SAMPLES_PER_ROW
BOOTSTRAP_SEED = 835083
BOOTSTRAP_RESAMPLES = 10_000


class AnalysisError(RuntimeError):
    """Raised when a capture is incomplete, altered, or internally inconsistent."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        fail(f"cannot read JSON {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(type(value) is int, f"{label}: expected integer")
    require(value >= (1 if positive else 0), f"{label}: expected non-negative value")
    return value


def finite_number(value: Any, label: str) -> float | int:
    require(type(value) in (int, float), f"{label}: expected number")
    require(math.isfinite(float(value)), f"{label}: expected finite number")
    require(value >= 0, f"{label}: expected non-negative number")
    return value


def packet_path(raw_path: Any, label: str) -> Path:
    require(isinstance(raw_path, str) and raw_path, f"{label}.path: missing path")
    candidate = Path(raw_path)
    if not candidate.is_absolute():
        candidate = P / candidate
    try:
        resolved = candidate.resolve(strict=True)
        resolved.relative_to(P.resolve())
    except (OSError, RuntimeError, ValueError) as error:
        fail(f"{label}.path: outside packet or missing: {error}")
    require(not candidate.is_symlink(), f"{label}.path: symlink is not allowed")
    require(resolved.is_file(), f"{label}.path: expected regular file")
    return resolved


def verify_descriptor(value: Any, label: str) -> tuple[Path, str]:
    require(isinstance(value, dict), f"{label}: descriptor is missing")
    path = packet_path(value.get("path"), label)
    expected_bytes = integer(value.get("bytes"), f"{label}.bytes")
    expected_sha = value.get("sha256")
    require(
        isinstance(expected_sha, str)
        and len(expected_sha) == 64
        and all(character in "0123456789abcdef" for character in expected_sha),
        f"{label}.sha256: invalid digest",
    )
    actual_bytes = path.stat().st_size
    actual_sha = sha256(path)
    require(actual_bytes == expected_bytes, f"{label}: byte count changed")
    require(actual_sha == expected_sha, f"{label}: SHA-256 changed")
    return path, expected_sha


def relative(path: Path) -> str:
    return str(path.relative_to(P.resolve()))


def midpoint(left: int, right: int) -> int:
    # Matches the Rust integer midpoint used by the native harness.
    return left // 2 + right // 2 + ((left % 2 + right % 2) // 2)


def nearest_rank(values: list[int], quantile: float) -> int:
    ordered = sorted(values)
    rank = max(1, math.ceil(len(ordered) * quantile))
    return ordered[rank - 1]


def summary(values: list[int | float]) -> dict[str, int | float]:
    require(values, "statistics: empty sample vector")
    ordered = sorted(values)
    p50: int | float
    if len(ordered) % 2:
        p50 = ordered[len(ordered) // 2]
    else:
        p50 = midpoint(int(ordered[len(ordered) // 2 - 1]), int(ordered[len(ordered) // 2]))
    p95 = nearest_rank([int(value) for value in ordered], 0.95)
    p99 = nearest_rank([int(value) for value in ordered], 0.99)
    return {
        "min": ordered[0],
        "p50": p50,
        "p95": p95,
        "p99": p99,
        "max": ordered[-1],
        "mean": statistics.mean(ordered),
    }


def spread(values: list[int | float]) -> dict[str, int | float | None]:
    require(values, "spread: empty vector")
    low = min(values)
    high = max(values)
    return {
        "min": low,
        "median": statistics.median(values),
        "max": high,
        "max_over_min": high / low if low > 0 else None,
    }


def compare_native_stats(native: Any, values: list[int], label: str) -> None:
    require(isinstance(native, dict), f"{label}: elapsed statistics are missing")
    expected = summary(values)
    for field in ("min", "p50", "p95", "p99", "max"):
        require(native.get(field) == expected[field], f"{label}.{field}: differs from raw samples")
    actual_mean = finite_number(native.get("mean"), f"{label}.mean")
    require(
        math.isclose(float(actual_mean), float(expected["mean"]), rel_tol=0.0, abs_tol=1e-6),
        f"{label}.mean: differs from raw samples",
    )


def descriptor_inputs(admission: dict[str, Any]) -> dict[str, str]:
    require(admission.get("status") == "pass", "capture admission is not pass")
    inputs = admission.get("inputs")
    require(isinstance(inputs, dict) and inputs, "capture admission has no inputs")
    result: dict[str, str] = {}
    for name, raw_descriptor in sorted(inputs.items()):
        require(isinstance(name, str) and name, "capture admission has an invalid input name")
        descriptor = raw_descriptor
        if isinstance(raw_descriptor, dict) and "descriptor" in raw_descriptor:
            descriptor = raw_descriptor["descriptor"]
        path, digest = verify_descriptor(descriptor, f"capture-admission.inputs[{name!r}]")
        result[name] = digest
        # A named packet input must resolve to the descriptor's actual file.
        named = packet_path(name, f"capture-admission.inputs[{name!r}]")
        require(named == path, f"capture admission input {name!r} names a different file")
    return result


def validate_plan(plan: dict[str, Any]) -> list[dict[str, Any]]:
    rows = plan.get("rows")
    require(isinstance(rows, list) and len(rows) == EXPECTED_ROWS, "measurement plan must have 72 rows")
    require(plan.get("cpu") == 12, "measurement plan must pin CPU 12")
    base = plan.get("base")
    require(isinstance(base, str) and len(base) == 40, "measurement plan base is not a commit hash")
    statistics_plan = plan.get("statistics")
    require(isinstance(statistics_plan, dict), "measurement plan statistics contract is missing")
    require(statistics_plan.get("bootstrap_resamples") == BOOTSTRAP_RESAMPLES,
            "measurement plan bootstrap count differs")
    require(statistics_plan.get("seed") == BOOTSTRAP_SEED,
            "measurement plan bootstrap seed differs")
    require(statistics_plan.get("sorted_endpoints") == [249, 9749],
            "measurement plan bootstrap endpoints differ")
    require(statistics_plan.get("process_quantiles") == "p50 integer midpoint; p95/p99 nearest rank",
            "measurement plan process quantile contract differs")
    require(statistics_plan.get("summary") == "midpoint median of six process p50 values",
            "measurement plan summary contract differs")
    require(statistics_plan.get("paired_comparisons") ==
            "eager/source per operation and cache state; block-paired p50 ratios",
            "measurement plan paired comparison contract differs")
    require(statistics_plan.get("spread_flag") ==
            "max/min block p50, p95, p99 or peak RSS exceeds 1.2",
            "measurement plan spread contract differs")
    require(statistics_plan.get("tail_flag") ==
            "none; p99 equals block maximum with 30 samples",
            "measurement plan tail contract differs")
    expected_pairs = {(case, state) for case in CASES for state in STATES}
    by_block: dict[int, list[tuple[str, str]]] = defaultdict(list)
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"measurement plan row {index}: malformed")
        block = integer(row.get("block"), f"measurement plan row {index}.block")
        case = row.get("case")
        state = row.get("cache_state")
        require(block in BLOCKS, f"measurement plan row {index}: invalid block")
        require(case in CASES, f"measurement plan row {index}: invalid case")
        require(state in STATES, f"measurement plan row {index}: invalid state")
        require(row.get("samples") == SAMPLES_PER_ROW, f"measurement plan row {index}: sample count")
        require(row.get("warmup") == WARMUPS_PER_ROW, f"measurement plan row {index}: warmup count")
        by_block[block].append((case, state))
    require(set(by_block) == set(BLOCKS), "measurement plan does not cover all six blocks")
    for block in BLOCKS:
        pairs = by_block[block]
        require(len(pairs) == len(expected_pairs), f"measurement plan block {block}: row count")
        require(set(pairs) == expected_pairs, f"measurement plan block {block}: not counterbalanced")
        require(len(pairs) == len(set(pairs)), f"measurement plan block {block}: duplicate pair")
    return rows


def metric_values(samples: list[dict[str, Any]], key: str) -> list[int]:
    values: list[int] = []
    for index, sample in enumerate(samples):
        metrics = sample.get("process_metrics")
        require(isinstance(metrics, dict), f"sample {index}: process_metrics missing")
        values.append(integer(metrics.get(key), f"sample {index}.process_metrics.{key}"))
    return values


def report_samples(
    report: dict[str, Any],
    plan_row: dict[str, Any],
    label: str,
    seen_pids: set[int],
) -> tuple[list[dict[str, Any]], dict[str, str], dict[str, Any]]:
    case = plan_row["case"]
    state = plan_row["cache_state"]
    require(isinstance(report, dict), f"{label}: report is not an object")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}: configuration missing")
    require(configuration.get("samples_per_case") == SAMPLES_PER_ROW, f"{label}: sample count differs")
    require(configuration.get("warmup_iterations_per_case") == WARMUPS_PER_ROW, f"{label}: warmup differs")
    require(configuration.get("filesystem_cache_states") == [state], f"{label}: cache state differs")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1, f"{label}: expected one timed result")
    result = results[0]
    require(isinstance(result, dict), f"{label}: timed result malformed")
    require(result.get("case") == case, f"{label}: timed case differs")
    require(result.get("cache_state") == state, f"{label}: timed state differs")
    evidence_list = report.get("filesystem_evidence")
    require(isinstance(evidence_list, list) and len(evidence_list) == 1, f"{label}: evidence missing")
    evidence = evidence_list[0]
    require(isinstance(evidence, dict), f"{label}: evidence malformed")
    require(evidence.get("case") == case, f"{label}: evidence case differs")
    require(evidence.get("cache_states") == [state], f"{label}: evidence state differs")
    require(evidence.get("sample_count") == SAMPLES_PER_ROW, f"{label}: evidence sample count")
    raw_samples = evidence.get("samples")
    require(isinstance(raw_samples, list) and len(raw_samples) == SAMPLES_PER_ROW, f"{label}: samples missing")
    by_index: dict[int, dict[str, Any]] = {}
    for position, sample in enumerate(raw_samples):
        require(isinstance(sample, dict), f"{label}.samples[{position}]: malformed")
        index = integer(sample.get("sample_index"), f"{label}.samples[{position}].sample_index")
        require(index < SAMPLES_PER_ROW, f"{label}.samples[{position}]: sample index out of range")
        require(index not in by_index, f"{label}: duplicate sample index {index}")
        require(sample.get("cache_state") == state, f"{label}.samples[{position}]: state differs")
        elapsed = integer(sample.get("elapsed_ns"), f"{label}.samples[{position}].elapsed_ns", positive=True)
        pid = integer(sample.get("child_process_id"), f"{label}.samples[{position}].child_process_id", positive=True)
        require(pid not in seen_pids, f"{label}: child process id {pid} was reused")
        seen_pids.add(pid)
        require(isinstance(sample.get("process_metrics"), dict), f"{label}.samples[{position}]: process metrics missing")
        by_index[index] = sample
        require(elapsed > 0, f"{label}.samples[{position}]: elapsed time is zero")
        if state == "cold-verified":
            proof = sample.get("cold_verified")
            require(isinstance(proof, dict) and proof.get("status") == "eligible", f"{label}.samples[{position}]: cold proof missing")
    require(set(by_index) == set(range(SAMPLES_PER_ROW)), f"{label}: sample indices are not complete")

    elapsed_stats = result.get("elapsed_ns")
    require(isinstance(elapsed_stats, dict), f"{label}: elapsed statistics missing")
    sorted_values = elapsed_stats.get("samples")
    sample_order = elapsed_stats.get("sample_order")
    require(isinstance(sorted_values, list) and len(sorted_values) == SAMPLES_PER_ROW, f"{label}: elapsed vector length")
    require(isinstance(sample_order, list) and len(sample_order) == SAMPLES_PER_ROW, f"{label}: sample order length")
    require(sorted(sample_order) == list(range(SAMPLES_PER_ROW)), f"{label}: sample order is not a permutation")
    reconstructed = [by_index[index]["elapsed_ns"] for index in sample_order]
    require(reconstructed == sorted_values, f"{label}: elapsed sample order does not match evidence")
    compare_native_stats(elapsed_stats, list(reconstructed), f"{label}.elapsed_ns")

    process_keys: set[str] | None = None
    for sample in by_index.values():
        metrics = sample["process_metrics"]
        keys = set(metrics)
        process_keys = keys if process_keys is None else process_keys & keys
        require("clock_ticks_per_second" in keys, f"{label}: process metric clock factor missing")
        for key, value in metrics.items():
            integer(value, f"{label}.process_metrics.{key}")
    require(process_keys is not None, f"{label}: process metric keys missing")
    process_keys.discard("clock_ticks_per_second")
    require(process_keys, f"{label}: no process metrics remain after clock factor")

    binary = report.get("binary_identity")
    environment = report.get("environment")
    require(isinstance(binary, dict) and isinstance(environment, dict), f"{label}: identity missing")
    binary_digest = binary.get("binary_sha256")
    require(
        isinstance(binary_digest, str)
        and len(binary_digest) == 64
        and all(character in "0123456789abcdef" for character in binary_digest),
        f"{label}: binary digest missing",
    )
    return list(by_index.values()), {"binary_sha256": binary_digest}, {
        "environment": environment,
        "process_keys": sorted(process_keys),
        "elapsed": [by_index[index]["elapsed_ns"] for index in range(SAMPLES_PER_ROW)],
        "report": report,
    }


def bootstrap_median(values: list[float]) -> tuple[list[float], list[float]]:
    require(len(values) == len(BLOCKS), "route ratio bootstrap requires six blocks")
    randomizer = random.Random(BOOTSTRAP_SEED)
    samples = sorted(
        statistics.median(randomizer.choices(values, k=len(BLOCKS)))
        for _ in range(BOOTSTRAP_RESAMPLES)
    )
    return samples, [samples[249], samples[9749]]


def compute() -> dict[str, Any]:
    # The formal packet name is capture-admission.json.  Keep a read-only
    # compatibility fallback for the root driver's earlier admission.json
    # spelling so a fully retained packet cannot be rejected solely for that
    # filename; the descriptor and status checks remain identical.
    admission_path = P / "capture-admission.json"
    if not admission_path.is_file():
        admission_path = P / "admission.json"
    plan_path = P / "measurement-plan.json"
    capture_path = P / "capture.json"
    admission = read_json(admission_path)
    admission_inputs = descriptor_inputs(admission)
    plan = read_json(plan_path)
    plan_rows = validate_plan(plan)
    capture = read_json(capture_path)
    require(capture.get("status") == "commands_pass", "formal capture is not commands_pass")
    capture_rows = capture.get("rows")
    require(isinstance(capture_rows, list) and len(capture_rows) == EXPECTED_ROWS, "capture must have 72 rows")
    require(capture.get("report_count") == EXPECTED_ROWS, "capture report count differs")
    require(capture.get("sample_count") == EXPECTED_SAMPLES, "capture sample count differs")

    seen_pids: set[int] = set()
    groups: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
    environments: list[dict[str, Any]] = []
    binaries: set[str] = set()
    process_keys: set[str] | None = None
    report_descriptors: list[dict[str, Any]] = []
    receipt_descriptors: list[dict[str, Any]] = []
    for index, row in enumerate(capture_rows):
        label = f"capture.rows[{index}]"
        require(isinstance(row, dict), f"{label}: malformed")
        plan_row = row.get("plan")
        require(plan_row == plan_rows[index], f"{label}.plan: differs from measurement plan")
        result_descriptor = row.get("result")
        require(isinstance(result_descriptor, dict), f"{label}.result: missing")
        require(result_descriptor.get("case") == plan_row["case"], f"{label}: case differs")
        require(result_descriptor.get("state") == plan_row["cache_state"], f"{label}: state differs")
        require(result_descriptor.get("exit_code") == 0, f"{label}: workload did not pass")
        report_path, report_digest = verify_descriptor(result_descriptor.get("report"), f"{label}.report")
        receipt_path, receipt_digest = verify_descriptor(result_descriptor.get("receipt"), f"{label}.receipt")
        receipt = read_json(receipt_path)
        require(receipt.get("exit_code") == 0, f"{label}: receipt exit code is nonzero")
        require(receipt.get("error") is None, f"{label}: receipt has an error")
        report = read_json(report_path)
        samples, binary_info, details = report_samples(report, plan_row, label, seen_pids)
        binary_digest = binary_info["binary_sha256"]
        binaries.add(binary_digest)
        environments.append(details["environment"])
        keys = set(details["process_keys"])
        process_keys = keys if process_keys is None else process_keys & keys
        block_record = {
            "block": plan_row["block"],
            "report": relative(report_path),
            "latency_ns": summary(details["elapsed"]),
            "process_metrics": {
                key: summary(metric_values(samples, key))
                for key in sorted(keys)
            },
            "report_sha256": report_digest,
            "receipt": relative(receipt_path),
            "receipt_sha256": receipt_digest,
            "sample_count": len(samples),
        }
        groups[(plan_row["case"], plan_row["cache_state"])].append(block_record)
        report_descriptors.append({"path": relative(report_path), "bytes": report_path.stat().st_size, "sha256": report_digest})
        receipt_descriptors.append({"path": relative(receipt_path), "bytes": receipt_path.stat().st_size, "sha256": receipt_digest})

    require(len(groups) == len(CASES) * len(STATES), "capture does not cover all case/state groups")
    require(len(binaries) == 1, "formal reports use more than one binary")
    require(len(set(json.dumps(env, sort_keys=True) for env in environments)) == 1, "formal reports use more than one environment")
    require(process_keys is not None, "formal reports expose no process metrics")

    distributions: list[dict[str, Any]] = []
    lookup: dict[tuple[str, str], dict[str, Any]] = {}
    spread_flags: list[dict[str, Any]] = []
    for case in CASES:
        for state in STATES:
            rows = sorted(groups[(case, state)], key=lambda value: value["block"])
            require([row["block"] for row in rows] == list(BLOCKS), f"{case}/{state}: block coverage differs")
            latency_spread = {
                metric: spread([row["latency_ns"][metric] for row in rows])
                for metric in ("min", "p50", "p95", "p99", "max", "mean")
            }
            process_spread = {
                key: {
                    metric: spread([row["process_metrics"][key][metric] for row in rows])
                    for metric in ("min", "p50", "p95", "p99", "max", "mean")
                }
                for key in sorted(process_keys)
            }
            for metric in ("p50", "p95", "p99"):
                value = latency_spread[metric]["max_over_min"]
                if value is not None and value > 1.2:
                    spread_flags.append({
                        "case": case,
                        "cache_state": state,
                        "metric": metric,
                        "reason": "six-block spread exceeds 20%",
                        "max_over_min": value,
                    })
            peak_rss_p50 = process_spread.get("peak_rss_bytes", {}).get("p50", {})
            rss_spread = peak_rss_p50.get("max_over_min")
            if rss_spread is not None and rss_spread > 1.2:
                spread_flags.append({
                    "case": case,
                    "cache_state": state,
                    "metric": "peak_rss_bytes.p50",
                    "reason": "six-block spread exceeds 20%",
                    "max_over_min": rss_spread,
                })
            distribution = {
                "case": case,
                "cache_state": state,
                "blocks": rows,
                "six_block_spread": {
                    "latency_ns": latency_spread,
                    "process_metrics": process_spread,
                },
            }
            distributions.append(distribution)
            lookup[(case, state)] = distribution

    route_pairs = (
        ("opc_open", "opc_file_eager_open", "opc_file_source_open"),
        ("opc_one_part_atomic_save", "opc_file_eager_one_part_atomic_save", "opc_file_source_one_part_atomic_save"),
        ("pptx_selected_slide_lifecycle", "pptx_file_eager_open_selected_slide_lifecycle", "pptx_file_source_open_selected_slide_lifecycle"),
    )
    route_ratios: list[dict[str, Any]] = []
    for pair_name, eager, source in route_pairs:
        for state in STATES:
            eager_blocks = lookup[(eager, state)]["blocks"]
            source_blocks = lookup[(source, state)]["blocks"]
            eager_p50 = [row["latency_ns"]["p50"] for row in eager_blocks]
            source_p50 = [row["latency_ns"]["p50"] for row in source_blocks]
            ratios = [eager_value / source_value for eager_value, source_value in zip(eager_p50, source_p50)]
            require(all(value > 0 for value in ratios), f"{pair_name}/{state}: invalid route ratio")
            bootstrap, interval = bootstrap_median(ratios)
            route_ratios.append({
                "pair": pair_name,
                "eager_case": eager,
                "source_case": source,
                "cache_state": state,
                "eager_p50_ns_by_block": eager_p50,
                "source_p50_ns_by_block": source_p50,
                "eager_over_source_p50_by_block": ratios,
                "median": statistics.median(ratios),
                "bootstrap_95": interval,
                "bootstrap": {
                    "method": "median of six block ratios, six draws with replacement",
                    "seed": BOOTSTRAP_SEED,
                    "resamples": BOOTSTRAP_RESAMPLES,
                    "lower_index": 249,
                    "upper_index": 9749,
                },
                "scope": "descriptive route ratio; not an optimization effect",
            })
    base = plan.get("base")
    require(isinstance(base, str) and len(base) == 40, "analysis base is not a commit hash")
    return {
        "schema": "litchi.0835.descriptive-analysis.v1",
        "base": base,
        "measurement_plan": {
            "path": relative(plan_path),
            "sha256": sha256(plan_path),
            "blocks": len(BLOCKS),
            "rows": EXPECTED_ROWS,
            "samples_per_row": SAMPLES_PER_ROW,
            "warmups_per_row": WARMUPS_PER_ROW,
            "samples": EXPECTED_SAMPLES,
            "cpu": 12,
            "cache_states": list(STATES),
            "cases": list(CASES),
        },
        "capture": {
            "path": relative(capture_path),
            "sha256": sha256(capture_path),
            "status": capture["status"],
            "reports": EXPECTED_ROWS,
            "samples": EXPECTED_SAMPLES,
        },
        "admission": {
            "path": relative(admission_path),
            "sha256": sha256(admission_path),
            "status": admission["status"],
            "inputs": admission_inputs,
        },
        "binary_sha256": next(iter(binaries)),
        "environment": environments[0],
        "reports": EXPECTED_ROWS,
        "samples": EXPECTED_SAMPLES,
        "distributions": distributions,
        "route_ratios": route_ratios,
        "spread_flags": spread_flags,
        "report_descriptors": report_descriptors,
        "receipt_descriptors": receipt_descriptors,
        "bootstrap_seed": BOOTSTRAP_SEED,
        "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
        "limitations": [
            "Synthetic fixed corpora and one host are represented by this capture.",
            "The OPC eager-open timer drops its package, while source-open retains the package through post-timer diagnostics; their ratio is descriptive and has asymmetric lifetime scope.",
            "PPTX logical source counters come from an untimed source replay and are not latency attribution.",
            "Verified-cold evidence observes page-cache residency and process read_bytes; it is not a physical-device I/O measurement.",
            "Each row uses three independent warmups and fresh process children for measured samples; counterbalancing reduces order effects but does not remove host noise.",
            "All route ratios and spread flags describe this current-source baseline; no before/after production optimization claim is authorized.",
            "iWork is outside the scope of this packet.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--write", action="store_true", help="write analysis.json exactly once")
    parser.add_argument("--check", action="store_true", help="compare an existing analysis.json byte-for-byte")
    args = parser.parse_args()
    try:
        result = compute()
        encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
        output = P / "analysis.json"
        if args.write:
            with output.open("x") as stream:
                stream.write(encoded)
        if args.check:
            require(output.read_text() == encoded, "analysis.json is not a deterministic recomputation")
        print(json.dumps({
            "reports": result["reports"],
            "samples": result["samples"],
            "route_ratios": len(result["route_ratios"]),
            "spread_flags": len(result["spread_flags"]),
        }, indent=2, sort_keys=True))
        return 0
    except AnalysisError as error:
        parser.error(str(error))
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
