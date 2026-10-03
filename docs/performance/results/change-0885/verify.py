#!/usr/bin/env python3
"""Verify and summarize raw 0885 producer-baseline children.

This verifier is intentionally descriptive.  It validates the semantic
contract and recomputes latency statistics from each raw elapsed vector; it
does not pool children, compare against another build, or reject a child for
being slow.  A wide native p50 spread is retained as a flag so the root
coordinator can inspect it without losing evidence.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
from pathlib import Path
import statistics
from typing import Any

import capture


SCHEMA = "litchi.xlsx-producer-full-cell-verification-0885.v1"
CASE = capture.CASE


def _read(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def floor_midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def nearest_rank(values: list[int], percentile: int) -> int:
    require(values, "cannot calculate a percentile for an empty vector")
    rank = (len(values) * percentile + 99) // 100
    return values[max(0, min(len(values) - 1, rank - 1))]


def raw_statistics(samples: list[int]) -> dict[str, int]:
    require(samples and all(type(value) is int and value > 0 for value in samples), "elapsed samples must be positive integers")
    ordered = sorted(samples)
    return {
        "min": ordered[0],
        "p50": floor_midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
        "p95": nearest_rank(ordered, 95),
        "p99": nearest_rank(ordered, 99),
        "max": ordered[-1],
    }


def _parse_time(value: str) -> _datetime.datetime:
    return _datetime.datetime.fromisoformat(value)


def _descriptor_matches(path: Path, record: dict[str, Any]) -> None:
    actual = capture.descriptor(path)
    for key in ("bytes", "sha256", "mode", "executable"):
        if key in record.get("identity", record) and record.get("identity", record)[key] != actual[key]:
            raise ValueError(f"artifact {path} {key} mismatch")


def _check_plan(plan: dict[str, Any]) -> None:
    capture.validate_plan(plan)


def _validate_sidecar(sidecar: dict[str, Any]) -> dict[str, Any]:
    require(sidecar.get("schema") == "litchi.perf-baseline.producer-shape-evidence.v1", "producer evidence schema mismatch")
    corpora = sidecar.get("corpora")
    require(type(corpora) is list and len(corpora) == 1, "producer evidence must contain one corpus")
    corpus = corpora[0]
    require(corpus.get("generator") == "litchi-xlsx-producer-shape-v1", "producer generator mismatch")
    require(corpus.get("shape") == "producer-medium-read", "producer shape mismatch")
    require(corpus.get("variant") == "read", "producer variant mismatch")
    require(corpus.get("selected_sheet") == 1, "selected sheet mismatch")
    require(corpus.get("selected_target") == "Bench01!D17", "selected target mismatch")
    require(corpus.get("shared_string_part") == "xl/sharedStrings.xml", "shared string part mismatch")
    require(corpus.get("shared_string_unique_count") == 64, "shared string unique count mismatch")
    require(corpus.get("shared_string_cell_percent") == 40, "shared string cell percentage mismatch")
    parts = corpus.get("parts")
    require(type(parts) is list and len(parts) == 6, "producer evidence part census mismatch")
    worksheets = [part for part in parts if str(part.get("part", "")).startswith("xl/worksheets/")]
    require(len(worksheets) == 4, "producer evidence worksheet census mismatch")
    require(sum(part.get("numeric_cells", -1) for part in worksheets) == capture.COUNTS["numeric"], "numeric census mismatch")
    require(sum(part.get("shared_string_cells", -1) for part in worksheets) == capture.COUNTS["text"], "text census mismatch")
    expected_worksheets = [f"xl/worksheets/sheet{i}.xml" for i in range(1, 5)]
    require([part.get("part") for part in worksheets] == expected_worksheets, "worksheet order mismatch")
    for part in worksheets:
        require(part.get("markup_compatibility_namespace") is True, "MCE namespace census mismatch")
        require(part.get("x14ac_namespace") is True, "x14ac namespace census mismatch")
        require(part.get("cols_block") is True, "cols census mismatch")
        require(part.get("selected_cell_stream_ineligible_on_cols") is True, "selected-cell route census mismatch")
        require(part.get("source_stream_eligible") is False, "source route census mismatch")
        require("x14ac-namespace" in part.get("source_stream_ineligible_reasons", []), "x14ac refusal missing")
    refusals = corpus.get("proven_refusals")
    require(type(refusals) is list and len(refusals) == 6, "producer refusal census mismatch")
    require(all(row.get("message") == "admitted by the value-only editor" for row in refusals), "producer admission mismatch")
    return corpus


def _validate_report(report: dict[str, Any], evidence: dict[str, Any], mode: str, lane: dict[str, Any]) -> dict[str, Any]:
    require(report.get("schema_version") == 1, "report schema version mismatch")
    tool = report.get("tool", {})
    expected_binary = "litchi-perf-baseline" if mode == "native" else "litchi-perf-baseline-alloc"
    require(tool.get("name") == "litchi-perf-baseline", "tool name mismatch")
    require(tool.get("version") == "0.1.0", "tool version mismatch")
    require(tool.get("binary") == expected_binary, "tool binary mismatch")
    require(tool.get("profile") == "release", "debug or unknown profile was captured")
    require(tool.get("target_os") == "linux" and tool.get("target_arch") == "x86_64", "tool target mismatch")
    if mode == "native":
        require(tool.get("instrumentation") == "none", "native instrumentation mismatch")
        require("allocator_counter_revision" not in tool or tool.get("allocator_counter_revision") is None, "native allocator revision unexpectedly present")
    else:
        require(tool.get("instrumentation") == "system_allocator_operation_scoped", "allocator instrumentation mismatch")
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3", "allocator counter revision mismatch")
    configuration = report.get("configuration", {})
    require(configuration.get("samples_per_case") == lane["samples"], "report sample count mismatch")
    require(configuration.get("warmup_iterations_per_case") == lane["warmup"], "report warmup mismatch")
    require(configuration.get("filesystem_cache_states") == ["warm"], "filesystem cache mode mismatch")
    require(configuration.get("filesystem_fresh_child_per_sample") is True, "fresh-child report flag missing")
    require(configuration.get("filesystem_process_isolated") is True, "process-isolated report flag missing")
    require(configuration.get("cases") == [CASE], "report case selection mismatch")
    results = report.get("results")
    require(type(results) is list and len(results) == 1, "report result count mismatch")
    result = results[0]
    require(result.get("case") == CASE, "result case mismatch")
    elapsed = result.get("elapsed_ns", {})
    samples = elapsed.get("samples")
    require(type(samples) is list and len(samples) == lane["samples"], "elapsed sample vector length mismatch")
    require(all(type(value) is int and value > 0 for value in samples), "elapsed samples must be positive integers")
    require(samples == sorted(samples), "elapsed samples are not in report order")
    order = elapsed.get("sample_order")
    # The binary sorts elapsed samples and retains the original sample index
    # permutation.  Native runs commonly produce a non-identity order.
    require(type(order) is list and sorted(order) == list(range(lane["samples"])), "elapsed sample order is not a permutation")
    stats = raw_statistics(samples)
    for key, expected in stats.items():
        require(elapsed.get(key) == expected, f"elapsed {key} was not recomputed from raw samples")
    corpus = result.get("corpus", {})
    require(corpus.get("archive_sha256") == capture.ARCHIVE_SHA256, "report archive digest mismatch")
    require(corpus.get("archive_bytes") == capture.ARCHIVE_BYTES, "report archive size mismatch")
    require(corpus.get("entry_count") == capture.COUNTS["cell"], "report entry count mismatch")
    scan = result.get("source", {}).get("producer_full_cell_scan")
    require(isinstance(scan, dict), "producer scan evidence missing")
    require(scan.get("sample_indices") == order, "producer scan sample alignment mismatch")
    require(scan.get("archive_sha256") == capture.ARCHIVE_SHA256 and scan.get("archive_bytes") == capture.ARCHIVE_BYTES, "producer scan archive mismatch")
    require(scan.get("expected_cell_count") == capture.COUNTS["cell"], "producer expected cell count mismatch")
    require(scan.get("expected_numeric_count") == capture.COUNTS["numeric"], "producer expected numeric count mismatch")
    require(scan.get("expected_text_count") == capture.COUNTS["text"], "producer expected text count mismatch")
    require(scan.get("observed_cell_counts") == [capture.COUNTS["cell"]] * lane["samples"], "producer observed cell counts mismatch")
    require(scan.get("observed_numeric_counts") == [capture.COUNTS["numeric"]] * lane["samples"], "producer observed numeric counts mismatch")
    require(scan.get("observed_text_counts") == [capture.COUNTS["text"]] * lane["samples"], "producer observed text counts mismatch")
    require(scan.get("oracle_digest_sha256") == capture.ORACLE_SHA256, "producer oracle digest mismatch")
    require(scan.get("observed_digest_sha256") == [capture.ORACLE_SHA256] * lane["samples"], "producer observed oracle mismatch")
    require(scan.get("marker_refusal_census_sha256") == capture.CENSUS_SHA256, "producer refusal census digest mismatch")
    require(scan.get("source_version_stable") == [True] * lane["samples"], "source version fence failed")
    require(scan.get("physical_cold_eligible") is False and scan.get("runtime_route_observed") is False, "cold/runtime route was unexpectedly claimed")
    routes = scan.get("worksheet_route_eligibility")
    require(type(routes) is list and [row.get("worksheet") for row in routes] == [f"xl/worksheets/sheet{i}.xml" for i in range(1, 5)], "worksheet route order mismatch")
    require(all(row.get("source_stream_eligible") is False for row in routes), "worksheet source route unexpectedly eligible")
    metrics = result.get("operation_metrics")
    require(isinstance(metrics, dict), "operation metrics missing")
    require(metrics.get("sample_count") == lane["samples"], "operation metric count mismatch")
    require(metrics.get("sample_indices") == order, "operation metric sample alignment mismatch")
    require(metrics.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index", "operation metric alignment contract mismatch")
    process = metrics.get("process")
    process_peak: list[int] | None = None
    if mode == "native":
        require(isinstance(process, dict), "native process metrics missing")
        peak = process.get("peak_rss_bytes")
        require(isinstance(peak, dict), "native peak RSS metric missing")
        require(peak.get("status") == "measured", "native peak RSS is not measured")
        require(peak.get("scope") == "process_lifetime_high_water_after_not_operation_peak", "native peak RSS scope mismatch")
        process_peak = peak.get("values")
        require(type(process_peak) is list and len(process_peak) == lane["samples"], "native peak RSS vector alignment mismatch")
        require(all(type(value) is int and value > 0 for value in process_peak), "native peak RSS vector is invalid")
    if mode == "native":
        require("allocation" not in metrics, "native report contains allocator metrics")
    else:
        _validate_allocation(metrics.get("allocation"), lane["samples"])
    _validate_sidecar(evidence)
    return {"elapsed": stats, "scan": scan, "corpus": corpus, "allocation": metrics.get("allocation"), "process_peak_rss": process_peak}


def _validate_allocation(allocation: Any, count: int) -> None:
    require(isinstance(allocation, dict), "allocator metrics missing")
    require(allocation.get("status") == "measured", "allocator metrics are not measured")
    require(allocation.get("scope") == "operation_global_system_allocator", "allocator scope mismatch")
    primary = {"allocation_calls", "allocated_bytes", "region_peak_live_bytes"}
    for key, value in allocation.items():
        if key in ("status", "scope"):
            continue
        require(isinstance(value, dict), f"allocator metric {key} is not a vector")
        require(value.get("status") == "measured" and value.get("scope") == "operation_global_system_allocator", f"allocator metric {key} status/scope mismatch")
        values = value.get("values")
        require(type(values) is list and len(values) == count, f"allocator metric {key} vector length mismatch")
        require(all(type(item) is int and item >= 0 for item in values), f"allocator metric {key} contains an invalid value")
    failed = allocation.get("failed_allocation_calls", {}).get("values")
    require(failed == [0] * count, "failed allocator calls were observed")
    for key in primary:
        values = allocation.get(key, {}).get("values")
        require(values and all(value > 0 for value in values), f"allocator primary metric {key} is not positive")


def _median_and_range(values: list[int]) -> dict[str, Any]:
    require(values, "summary vector is empty")
    ordered = sorted(values)
    return {
        "median": statistics.median(ordered),
        "min": ordered[0],
        "max": ordered[-1],
        "values": values,
    }


def _float_summary(values: list[float]) -> dict[str, Any]:
    require(values, "summary vector is empty")
    return {"median": statistics.median(values), "min": min(values), "max": max(values), "values": values}


def verify(plan_path: Path, capture_root: Path, output: Path | None = None) -> dict[str, Any]:
    plan = capture.load_plan(plan_path)
    root = capture_root.resolve()
    capture_receipt_path = root / "capture.json"
    require(capture_receipt_path.is_file(), "capture.json is missing")
    capture_receipt = _read(capture_receipt_path)
    require(capture_receipt.get("schema") == capture.CAPTURE_SCHEMA, "capture receipt schema mismatch")
    require(capture_receipt.get("status") == "pass" and capture_receipt.get("schedule") == ["native", "allocator"], "capture receipt status/schedule mismatch")
    require(capture_receipt.get("plan_sha256") == _sha(plan_path), "capture receipt plan binding mismatch")
    all_rows: dict[str, list[dict[str, Any]]] = {}
    sidecar_digest: str | None = None
    starts: list[tuple[_datetime.datetime, _datetime.datetime, str]] = []
    summaries: dict[str, Any] = {}
    for mode in ("native", "allocator"):
        lane = capture._lane(plan, mode)
        binary = capture._build_descriptor(plan_path.resolve().parent, plan, mode)
        lane_receipt_path = root / f"capture-{mode}.json"
        require(lane_receipt_path.is_file(), f"{mode} lane receipt missing")
        lane_receipt = _read(lane_receipt_path)
        require(lane_receipt.get("schema") == capture.CAPTURE_SCHEMA and lane_receipt.get("status") == "pass", f"{mode} lane receipt invalid")
        require(lane_receipt.get("plan_sha256") == _sha(plan_path), f"{mode} lane plan binding mismatch")
        records = lane_receipt.get("records")
        require(type(records) is list and len(records) == lane["children"], f"{mode} child record count mismatch")
        rows: list[dict[str, Any]] = []
        child_p50: list[int] = []
        child_p95: list[int] = []
        child_p99: list[int] = []
        child_peak_rss: list[int] = []
        child_cells_per_second: list[float] = []
        allocation_vectors: dict[str, list[int]] = {}
        for expected_index, record in enumerate(records):
            require(record.get("path") == f"{mode}/child-{expected_index:02d}/receipt.json", f"{mode} child path/order mismatch")
            receipt_path = root / record["path"]
            require(receipt_path.is_file(), f"missing {mode} child receipt")
            receipt = _read(receipt_path)
            require(receipt.get("schema") == capture.RECEIPT_SCHEMA, f"{mode} receipt schema mismatch")
            require(receipt.get("mode") == mode and receipt.get("child_index") == expected_index, f"{mode} receipt identity mismatch")
            require(receipt.get("exit_code") == 0, f"{mode} child did not exit successfully")
            require(receipt.get("profile") == "release", f"{mode} child profile mismatch")
            require(receipt.get("plan_sha256") == _sha(plan_path), f"{mode} receipt plan binding mismatch")
            require(receipt.get("binary") == binary, f"{mode} receipt binary/build binding mismatch")
            argv = receipt.get("argv")
            require(type(argv) is list and len(argv) == 16, f"{mode} child argv shape mismatch")
            expected_argv = capture._command(mode, receipt["binary"]["path"], lane["warmup"], lane["samples"], Path(argv[13]), Path(argv[15]))
            require(argv == expected_argv, f"{mode} child command mismatch")
            folder = receipt_path.parent
            command_path = folder / "command.json"
            require(command_path.is_file(), f"{mode} child command receipt missing")
            command = _read(command_path)
            require(command.get("environment") == capture.CHILD_ENV, f"{mode} child environment mismatch")
            require(command.get("argv") == argv and command.get("exit_code") == 0, f"{mode} command receipt mismatch")
            require(receipt.get("command", {}).get("sha256") == _sha(command_path), f"{mode} command descriptor mismatch")
            for key in ("mode", "child_index", "cwd", "profile", "plan_sha256", "binary", "environment", "started_utc", "finished_utc", "exit_code"):
                require(command.get(key) == receipt.get(key), f"{mode} command/receipt {key} mismatch")
            report_path = folder / "report.json"
            evidence_path = folder / "producer-evidence.json"
            for stream in ("stdout.txt", "stderr.txt"):
                stream_path = folder / stream
                require(stream_path.is_file(), f"{mode} child {stream} missing")
                require(receipt.get("artifacts", {}).get(stream, {}).get("sha256") == _sha(stream_path), f"{mode} {stream} descriptor mismatch")
            require(report_path.is_file() and evidence_path.is_file(), f"{mode} child outputs missing")
            require(receipt.get("artifacts", {}).get("report.json", {}).get("sha256") == _sha(report_path), f"{mode} report descriptor mismatch")
            require(receipt.get("artifacts", {}).get("producer-evidence.json", {}).get("sha256") == _sha(evidence_path), f"{mode} evidence descriptor mismatch")
            report = _read(report_path)
            evidence = _read(evidence_path)
            checked = _validate_report(report, evidence, mode, lane)
            corpus = checked["corpus"]
            sidecar_json = json.dumps(evidence, sort_keys=True, separators=(",", ":"))
            digest = hashlib.sha256(sidecar_json.encode()).hexdigest()
            if sidecar_digest is None:
                sidecar_digest = digest
            require(digest == sidecar_digest, "producer evidence differs between children")
            child_p50.append(checked["elapsed"]["p50"])
            child_p95.append(checked["elapsed"]["p95"])
            child_p99.append(checked["elapsed"]["p99"])
            if mode == "native":
                child_peak_rss.append(max(checked["process_peak_rss"]))
                child_cells_per_second.append(capture.COUNTS["cell"] * 1_000_000_000 / checked["elapsed"]["p50"])
            if mode == "allocator":
                for key, vector in checked["allocation"].items():
                    if isinstance(vector, dict) and "values" in vector:
                        allocation_vectors.setdefault(key, []).append(statistics.median(vector["values"]))
            started = _parse_time(receipt["started_utc"])
            finished = _parse_time(receipt["finished_utc"])
            require(finished >= started, f"{mode} child time interval is invalid")
            starts.append((started, finished, f"{mode}/{expected_index:02d}"))
            rows.append({"child_index": expected_index, "report": str(report_path), "evidence": str(evidence_path), "stats": checked["elapsed"]})
        require(all(rows[index]["stats"]["p50"] <= rows[index]["stats"]["p95"] <= rows[index]["stats"]["p99"] for index in range(len(rows))), f"{mode} percentile order mismatch")
        summary: dict[str, Any] = {
            "children": len(rows),
            "p50_ns": _median_and_range(child_p50),
            "p95_ns": _median_and_range(child_p95),
            "p99_ns": _median_and_range(child_p99),
            "rows": rows,
        }
        if mode == "native":
            summary["p50_spread_ratio"] = max(child_p50) / min(child_p50)
            summary["p50_spread_flag"] = summary["p50_spread_ratio"] > 1.20
            summary["process_peak_rss_bytes"] = _median_and_range(child_peak_rss)
            summary["cells_per_second_from_p50"] = _float_summary(child_cells_per_second)
        else:
            summary["allocation"] = {key: _median_and_range(values) for key, values in allocation_vectors.items()}
        summaries[mode] = summary
        all_rows[mode] = rows
    ordered = sorted(starts)
    require(all(a[1] <= b[0] for a, b in zip(ordered, ordered[1:])), "child processes were not sequential")
    native_finish = max(item[1] for item in starts if item[2].startswith("native/"))
    allocator_start = min(item[0] for item in starts if item[2].startswith("allocator/"))
    require(native_finish <= allocator_start, "allocator lane started before native lane finished")
    result = {
        "schema": SCHEMA,
        "status": "pass",
        "plan": str(plan_path.resolve()),
        "plan_sha256": _sha(plan_path),
        "capture": str(capture_receipt_path),
        "capture_sha256": _sha(capture_receipt_path),
        "profile": "release",
        "cpu_affinity": capture.CPU_AFFINITY,
        "case": CASE,
        "children": {mode: len(rows) for mode, rows in all_rows.items()},
        "summaries": summaries,
        "producer_evidence_sha256": sidecar_digest,
        "performance_claim": None,
        "comparison": None,
    }
    if output is not None:
        with output.open("x", encoding="utf-8") as handle:
            json.dump(result, handle, indent=2, sort_keys=True)
            handle.write("\n")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--capture-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = verify(args.plan.resolve(), args.capture_root.resolve(), args.output.resolve())
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
