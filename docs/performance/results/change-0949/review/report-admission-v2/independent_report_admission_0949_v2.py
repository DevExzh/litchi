#!/usr/bin/env python3
"""Independent admission and conservation audit for the sealed 0949 packet.

This script is offline-only.  It does not launch a workload, heaptrack, or
heaptrack_print.  It verifies capture custody, revalidates the four fresh
application reports against the retained 0948 report contract, and recomputes
raw-decoder totals from its independent aggregate files.  Heaptrack histogram
and footer reconciliation is deliberately reported as pending for the
separately owned heaptrack_print export step.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
from typing import Any

CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
CHILD_IDS = (
    "h00-baseline-primary-percent-medium",
    "h01-candidate-primary-percent-medium",
    "h02-baseline-primary-percent-dense-sparse",
    "h03-candidate-primary-percent-dense-sparse",
)
ENV = {"LANG": "C", "LC_ALL": "C", "PATH": "/usr/bin:/bin", "RUST_BACKTRACE": "0", "TZ": "UTC"}
PUB_COUNTERS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
)
SOURCE_VECTOR_FIELDS = (
    "source_read_calls", "source_read_bytes", "workbook_read_calls",
    "workbook_read_bytes", "selected_worksheet_read_calls",
    "selected_worksheet_read_bytes", "unselected_worksheet_read_calls",
    "unselected_worksheet_read_bytes", "payload_materializations", "cache_hits",
    "cache_cold_loads", "cache_waiter_joins", "cache_successful_loads",
    "cache_failed_loads", "cache_evictions", "cache_bypasses",
    "cache_oversized_bypasses", "cache_allocation_bypasses", "cache_in_flight_loads",
    "cache_retained_entries", "cache_retained_bytes", "cache_budget_memory_used",
    "cache_budget_reserved_bytes", "cache_budget_reservation_failures",
    "budget_used_after_package_drop", "budget_used_after_handles_drop",
    "budget_objects_used_after_handles_drop",
)
HEX = set("0123456789abcdef")


class AuditError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"not a regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def desc(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"not a regular file: {path}")
    return {"path": str(path.absolute()), "bytes": path.stat().st_size, "sha256": sha(path)}


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")

    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicates)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"cannot parse JSON {path}: {error}") from error


def rel_path(packet: Path, value: str) -> Path:
    path = Path(value)
    require(not path.is_absolute() and path.as_posix() == value and ".." not in path.parts,
            f"artifact path is not normalized relative path: {value!r}")
    result = packet / path
    require(result.resolve().is_relative_to(packet.resolve()), f"artifact escapes packet: {value}")
    return result


def descriptor(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not a descriptor")
    path = value.get("path")
    require(isinstance(path, str) and Path(path).is_absolute(), f"{label}.path is not absolute")
    actual = desc(Path(path))
    for key in ("bytes", "sha256"):
        require(value.get(key) == actual[key], f"{label}.{key} mismatch")
    return actual


def check_hex(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256")


def same_descriptor(actual: Any, expected: dict[str, Any], label: str) -> None:
    require(isinstance(actual, dict), f"{label} is not a descriptor")
    for key in ("path", "bytes", "sha256"):
        require(actual.get(key) == expected[key], f"{label}.{key} mismatch")


def vector(value: Any, label: str, length: int = 1) -> list[Any]:
    require(isinstance(value, list) and len(value) == length, f"{label} is not length {length}")
    return value


def unique_old(values: Any, label: str) -> Any:
    require(isinstance(values, list) and values, f"old {label} is not a nonempty vector")
    unique = {json.dumps(value, sort_keys=True, separators=(",", ":")) for value in values}
    require(len(unique) == 1, f"old deterministic {label} varies across repeats")
    return values[0]


def child_path(packet: Path, child_id: str, name: str) -> Path:
    path = packet / "capture-v1" / "children" / child_id / name
    require(path.resolve().is_relative_to((packet / "capture-v1").resolve()),
            f"child path escapes packet: {path}")
    return path


def artifact_index_check(packet: Path) -> tuple[dict[str, dict[str, Any]], dict[str, Any]]:
    path = packet / "receipts/capture-artifact-index-v1.json"
    index = read_json(path)
    require(isinstance(index, dict) and index.get("schema") == "litchi.performance.0949.root-capture-terminal.v1",
            "capture artifact index schema changed")
    artifacts = index.get("artifacts")
    require(isinstance(artifacts, list), "capture artifact index artifacts is not a list")
    rows: dict[str, dict[str, Any]] = {}
    for row in artifacts:
        require(isinstance(row, dict) and isinstance(row.get("path"), str),
                "malformed artifact index row")
        name = row["path"]
        require(name not in rows, f"duplicate artifact index path: {name}")
        check_hex(row.get("sha256"), f"artifact {name} sha256")
        require(isinstance(row.get("bytes"), int) and row["bytes"] >= 0,
                f"artifact {name} byte count invalid")
        physical = rel_path(packet, name)
        actual = desc(physical)
        require(actual["bytes"] == row["bytes"] and actual["sha256"] == row["sha256"],
                f"artifact index mismatch: {name}")
        rows[name] = row
    require(index.get("exit_code") == 0 and index.get("retries") == 0,
            "capture artifact index does not record clean terminal/no-retry")
    expected: set[str] = {"capture-v1/runner-start.json", "capture-v1/runner-terminal.json"}
    for child_id in CHILD_IDS:
        child = packet / "capture-v1/children" / child_id
        terminal = read_json(child / "terminal.json")
        actual_name = terminal.get("artifacts", {}).get("profile_actual_name")
        require(actual_name in {"heaptrack.data.gz", "heaptrack.data.gz.zst"},
                f"{child_id} has invalid actual profile name")
        expected.update(
            f"capture-v1/children/{child_id}/{name}"
            for name in ("start.json", "terminal.json", "stdout.log", "stderr.log",
                         "time.txt", "report.json", actual_name)
        )
    require(set(rows) == expected,
            f"artifact index census differs: missing={sorted(expected - set(rows))}, extra={sorted(set(rows) - expected)}")
    return rows, index


def plan_and_outer(packet: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    plan_path = packet / "heaptrack-plan-v3.json"
    runner_path = packet / "tooling/run_heaptrack_0949_v3.py"
    plan = read_json(plan_path)
    require(plan.get("plan_revision") == "v3" and plan.get("status") == "prospective_no_capture",
            "final plan is not v3 prospective plan")
    require(plan.get("capture_authorization") == "root_GO_required", "capture authorization changed")
    matrix = plan.get("matrix")
    require(isinstance(matrix, dict), "plan matrix missing")
    for key, expected in (("children", 4), ("warmup", 0), ("samples", 1), ("cpu", 12),
                          ("deadline_seconds", 600), ("no_retry", True), ("serial", True)):
        require(matrix.get(key) == expected, f"plan matrix {key} changed")
    require(matrix.get("jobs") and [x.get("name") for x in matrix["jobs"]] == list(CHILD_IDS),
            "plan child order changed")
    plan_desc = desc(plan_path)
    runner_desc = desc(runner_path)
    outer = read_json(packet / "capture-v1/runner-terminal.json")
    require(outer.get("status") == "terminal" and outer.get("capture_status") == "pass",
            "runner terminal did not pass")
    require(outer.get("children") == list(CHILD_IDS)
            and outer.get("expected_children") == list(CHILD_IDS)
            and outer.get("child_terminals_written") == 4
            and outer.get("failures") == []
            and outer.get("runner_error") is None
            and outer.get("no_retry") is True
            and outer.get("analysis_started") is False
            and outer.get("terminal_before_any_report_or_profile_parse") is True,
            "runner terminal custody/schedule fields changed")
    same_descriptor(outer.get("plan"), plan_desc, "runner terminal plan")
    same_descriptor(outer.get("runner"), runner_desc, "runner terminal runner")
    return plan, plan_desc, runner_desc


def check_child(packet: Path, plan: dict[str, Any], plan_desc: dict[str, Any],
                runner_desc: dict[str, Any], child_id: str, index_rows: dict[str, dict[str, Any]]) -> dict[str, Any]:
    child = packet / "capture-v1/children" / child_id
    start = read_json(child / "start.json")
    terminal = read_json(child / "terminal.json")
    require(start.get("status") == "started" and terminal.get("status") == "pass",
            f"{child_id} start/terminal status invalid")
    for record, name in ((start, "start"), (terminal, "terminal")):
        require(record.get("schema") == "litchi.performance.0949.heaptrack-child-terminal.v1",
                f"{child_id} {name} schema changed")
        require(record.get("id") == child_id and record.get("case") == CASE,
                f"{child_id} {name} identity changed")
        same_descriptor(record.get("plan"), plan_desc, f"{child_id} {name} plan")
        same_descriptor(record.get("runner"), runner_desc, f"{child_id} {name} runner")
        require(record.get("environment") == ENV and record.get("warmup") == 0
                and record.get("sample_index") == 0 and record.get("deadline_seconds") == 600,
                f"{child_id} {name} fixed child configuration changed")
    require(start.get("argv") == terminal.get("argv") and start.get("cwd") == terminal.get("cwd"),
            f"{child_id} start/terminal command binding changed")
    require(terminal.get("returncode") == 0 and terminal.get("timed_out") is False
            and terminal.get("launch_error") is None and terminal.get("process_group_clear") is True
            and terminal.get("terminal_before_any_report_or_profile_parse") is True,
            f"{child_id} terminal process status invalid")
    role = terminal.get("role")
    shape = terminal.get("xlsx_cell_crud_shape")
    require(role in {"baseline", "candidate"} and shape in {"medium", "dense-sparse"},
            f"{child_id} role/shape invalid")
    wanted = next(row for row in plan["matrix"]["jobs"] if row["name"] == child_id)
    require(wanted["role"] == role and wanted["shape"] == shape,
            f"{child_id} schedule role/shape mismatch")
    source = terminal.get("source")
    binary = terminal.get("binary")
    require(isinstance(source, dict) and isinstance(binary, dict), f"{child_id} source/binary missing")
    require(source.get("root") == plan["sources"][role]["root"]
            and source.get("manifest") == plan["sources"][role]["manifest"]
            and source.get("source_commit") == plan["sources"][role]["source_commit"],
            f"{child_id} source binding mismatch")
    require(binary == plan["binaries"][role]["binary"], f"{child_id} binary binding mismatch")
    require(terminal.get("cwd") == source["root"], f"{child_id} cwd/source-root mismatch")
    require(terminal.get("environment_sha256") == hashlib.sha256(
        json.dumps(ENV, sort_keys=True, separators=(",", ":")).encode()).hexdigest(),
            f"{child_id} environment digest mismatch")
    argv = terminal.get("argv")
    require(isinstance(argv, list) and argv[-1].endswith("/report.json"),
            f"{child_id} argv/report binding missing")
    require(binary["path"] in argv,
            f"{child_id} argv binary mismatch")
    artifacts = terminal.get("artifacts")
    require(isinstance(artifacts, dict) and artifacts.get("profile_match") == "exactly_one",
            f"{child_id} profile admission failed")
    actual_name = artifacts.get("profile_actual_name")
    actual_path = Path(artifacts.get("profile_actual_path", ""))
    require(actual_name in {"heaptrack.data.gz", "heaptrack.data.gz.zst"}
            and actual_path == (child / actual_name).absolute(),
            f"{child_id} actual profile path mismatch")
    require(actual_path.is_file() and not actual_path.is_symlink(), f"{child_id} profile is not regular")
    expected_files = {"start.json", "terminal.json", "stdout.log", "stderr.log", "time.txt", "report.json", actual_name}
    actual_files = {p.name for p in child.iterdir()}
    require(actual_files == expected_files, f"{child_id} unexpected child files: {actual_files ^ expected_files}")
    for filename in expected_files:
        rel = f"capture-v1/children/{child_id}/{filename}"
        require(rel in index_rows, f"{child_id} missing artifact index row {filename}")
    return {"id": child_id, "role": role, "shape": shape, "terminal": terminal}


def source_counter_projection(summary: dict[str, Any]) -> dict[str, Any]:
    source = summary["source"]
    xlsx = source["xlsx_cell_values"]
    values: dict[str, Any] = {}
    for field in SOURCE_VECTOR_FIELDS:
        values[field] = xlsx[field][0]
    values["generic_read_calls"] = source["read_calls"][0]
    values["generic_read_bytes"] = source["read_bytes"][0]
    values["ordinary_payload_read_calls"] = source["ordinary_payload_read_calls"][0]
    values["ordinary_payload_read_bytes"] = source["ordinary_payload_read_bytes"][0]
    values["max_in_flight_reads"] = source["max_in_flight_reads"][0]
    values["output_budget_refusal"] = xlsx["output_budget_refusal"]
    values["pre_publication_budget"] = xlsx["pre_publication_budget"][0]
    values["post_publication_budget"] = xlsx["post_publication_budget"][0]
    return values


def report_admission(packet: Path, plan: dict[str, Any], child_meta: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for meta in child_meta:
        child_id, role, shape = meta["id"], meta["role"], meta["shape"]
        report_path = packet / "capture-v1/children" / child_id / "report.json"
        report = read_json(report_path)
        require(report.get("schema_version") == 1, f"{child_id} report schema changed")
        tool = report.get("tool")
        require(isinstance(tool, dict) and tool.get("name") == "litchi-perf-baseline"
                and tool.get("binary") == "litchi-perf-baseline-alloc"
                and tool.get("profile") == "release"
                and tool.get("instrumentation") == "system_allocator_operation_scoped"
                and tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{child_id} report tool identity changed")
        terminal = meta["terminal"]
        binary_identity = report.get("binary_identity")
        require(isinstance(binary_identity, dict)
                and binary_identity.get("binary_sha256") == terminal["binary"]["sha256"]
                and binary_identity.get("binary_bytes") == terminal["binary"]["bytes"]
                and binary_identity.get("path") == terminal["binary"]["path"]
                and binary_identity.get("executable") is True,
                f"{child_id} report binary identity mismatch")
        configuration = report.get("configuration")
        require(isinstance(configuration, dict)
                and configuration.get("samples_per_case") == 1
                and configuration.get("warmup_iterations_per_case") == 0
                and configuration.get("cases") == [CASE]
                and configuration.get("xlsx_cell_crud_shapes") == [shape]
                and configuration.get("filesystem_process_isolated") is True,
                f"{child_id} report configuration mismatch")
        results = report.get("results")
        require(isinstance(results, list) and len(results) == 1, f"{child_id} result cardinality changed")
        result = results[0]
        require(result.get("case") == CASE, f"{child_id} result case changed")
        corpus = result.get("corpus")
        require(isinstance(corpus, dict) and corpus.get("shape") == shape
                and corpus.get("package_format") == "XLSX/OPC/ZIP"
                and isinstance(corpus.get("archive_sha256"), str),
                f"{child_id} corpus identity invalid")
        sink = result.get("sink")
        require(isinstance(sink, dict) and sink.get("accepted_bytes", 0) > 0
                and sink.get("write_calls", 0) > 0 and sink.get("largest_write") <= 65536,
                f"{child_id} sink contract invalid")
        source = result.get("source")
        xlsx = source.get("xlsx_cell_values") if isinstance(source, dict) else None
        require(isinstance(source, dict) and isinstance(xlsx, dict), f"{child_id} source report missing")
        require(xlsx.get("implementation") == "source-backed"
                and xlsx.get("cache_mode") == "unmanaged-control"
                and xlsx.get("cache_budget_managed") is False
                and xlsx.get("update_count") == corpus["xlsx"]["one_percent_update_count"]
                and xlsx.get("selected_worksheet_count") == corpus["xlsx"]["sheet_count"],
                f"{child_id} source semantic contract invalid")
        for field in ("publication_ns", "publication_allocation_metrics"):
            vector(xlsx.get(field), field, 1)
        publication = xlsx["publication_allocation_metrics"][0]
        require(isinstance(publication, dict) and publication.get("status") == "measured",
                f"{child_id} publication allocation metric unavailable")
        for field in PUB_COUNTERS:
            require(isinstance(publication.get(field), int) and publication[field] >= 0,
                    f"{child_id} publication counter {field} invalid")
        for field in ("output_sha256", "semantic_sha256", "untouched_member_sha256"):
            value = xlsx.get(field)
            value = value[0] if isinstance(value, list) and len(value) == 1 else value
            check_hex(value, f"{child_id} {field}")
        require(xlsx.get("output_sha256", [None])[0] == result.get("output_sha256"),
                f"{child_id} output digest mismatch")
        old_path = Path("/home/zhuhe/litchi-performance-0948/allocator") / (
            f"r00-{role}-primary-percent-{shape}/report.json")
        old = read_json(old_path)
        old_result = old["results"][0]
        require(corpus == old_result["corpus"], f"{child_id} corpus differs from retained 0948 report")
        require(sink == old_result["sink"], f"{child_id} sink contract differs from retained 0948 report")
        old_xlsx = old_result["source"]["xlsx_cell_values"]
        for field in SOURCE_VECTOR_FIELDS:
            current = xlsx[field][0]
            expected = unique_old(old_xlsx[field], f"{child_id}.{field}")
            require(current == expected, f"{child_id} deterministic source counter differs: {field}")
        for field in ("read_calls", "read_bytes", "ordinary_payload_read_calls",
                      "ordinary_payload_read_bytes", "max_in_flight_reads"):
            current = source[field][0]
            expected = unique_old(old_result["source"][field], f"{child_id}.{field}")
            require(current == expected, f"{child_id} deterministic generic source counter differs: {field}")
        for field in ("output_budget_refusal", "pre_publication_budget", "post_publication_budget"):
            current = xlsx[field][0] if isinstance(xlsx[field], list) else xlsx[field]
            expected = unique_old(old_xlsx[field], f"{child_id}.{field}") if isinstance(old_xlsx[field], list) else old_xlsx[field]
            require(current == expected, f"{child_id} deterministic budget evidence differs: {field}")
        for field in ("output_sha256", "semantic_sha256", "untouched_member_sha256"):
            current = xlsx[field][0] if isinstance(xlsx[field], list) else xlsx[field]
            old_values = old_xlsx[field] if isinstance(old_xlsx[field], list) else [old_xlsx[field]]
            require(current in old_values, f"{child_id} retained 0948 {field} identity differs")
        old_publication = old_xlsx["publication_allocation_metrics"]
        for field in PUB_COUNTERS:
            expected = unique_old([item[field] for item in old_publication], f"{child_id}.publication.{field}")
            require(publication[field] == expected,
                    f"{child_id} exact 0948 publication counter differs: {field}")
        rows.append({
            "id": child_id, "role": role, "shape": shape,
            "report": desc(report_path), "retained_0948_report": desc(old_path),
            "output_sha256": result["output_sha256"],
            "corpus_archive_sha256": corpus["archive_sha256"],
            "sink": sink,
            "publication_counters": {field: publication[field] for field in PUB_COUNTERS},
            "publication_ns": xlsx["publication_ns"][0],
            "whole_process_report_scope": "application report operation/publication counters; not heaptrack interval",
        })
    return rows


def decoder_conservation(packet: Path, child_meta: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for meta in child_meta:
        label = {"h00-baseline-primary-percent-medium": "baseline-medium",
                 "h01-candidate-primary-percent-medium": "candidate-medium",
                 "h02-baseline-primary-percent-dense-sparse": "baseline-dense",
                 "h03-candidate-primary-percent-dense-sparse": "candidate-dense"}[meta["id"]]
        root = packet / "raw-decoder-v1" / label
        summary = read_json(root / "summary.json")
        require(summary.get("status") == "completed"
                and summary.get("schema") == "litchi.performance.0949.heaptrack-raw-decode.v1",
                f"{label} decoder summary not completed")
        profile = summary.get("profile")
        require(isinstance(profile, dict), f"{label} decoder profile descriptor missing")
        actual = desc(Path(meta["terminal"]["artifacts"]["profile_actual_path"]))
        require(profile.get("path") == actual["path"] and profile.get("bytes") == actual["bytes"]
                and profile.get("sha256") == actual["sha256"],
                f"{label} decoder raw profile binding differs from terminal")
        sizes: dict[int, tuple[int, int]] = {}
        with (root / "size-aggregates.tsv").open(encoding="utf-8") as stream:
            require(stream.readline().rstrip("\n") == "size_bytes\tcalls\trequested_bytes",
                    f"{label} size header changed")
            for line_number, line in enumerate(stream, 2):
                parts = line.rstrip("\n").split("\t")
                require(len(parts) == 3, f"{label} malformed size row {line_number}")
                size, calls, requested = map(int, parts)
                require(size >= 0 and calls > 0 and requested == size * calls,
                        f"{label} size arithmetic mismatch {line_number}")
                require(size not in sizes, f"{label} duplicate size row {size}")
                sizes[size] = (calls, requested)
        trace_sizes: dict[int, dict[int, int]] = {}
        with (root / "trace-size-aggregates.tsv").open(encoding="utf-8") as stream:
            require(stream.readline().rstrip("\n") == "trace_id\tsize_bytes\tcalls\trequested_bytes",
                    f"{label} trace-size header changed")
            for line_number, line in enumerate(stream, 2):
                parts = line.rstrip("\n").split("\t")
                require(len(parts) == 4, f"{label} malformed trace-size row {line_number}")
                trace_id = int(parts[0], 16); size, calls, requested = map(int, parts[1:])
                require(size >= 0 and calls > 0 and requested == size * calls,
                        f"{label} trace-size arithmetic mismatch {line_number}")
                require(size not in trace_sizes.setdefault(trace_id, {}),
                        f"{label} duplicate trace-size row {line_number}")
                trace_sizes[trace_id][size] = calls
        trace_rows: dict[int, tuple[int, int]] = {}
        with (root / "trace-aggregates.jsonl").open(encoding="utf-8") as stream:
            for line_number, line in enumerate(stream, 1):
                row = json.loads(line)
                trace_id = int(row["trace_id"], 16)
                require(trace_id not in trace_rows, f"{label} duplicate trace row {trace_id:x}")
                calls = int(row["calls"]); requested = int(row["requested_bytes"])
                require(calls > 0 and requested >= 0, f"{label} invalid trace totals")
                trace_rows[trace_id] = (calls, requested)
        require(set(trace_rows) == set(trace_sizes), f"{label} trace/trace-size key sets differ")
        trace_calls = sum(calls for calls, _ in trace_rows.values())
        trace_bytes = sum(bytes_ for _, bytes_ in trace_rows.values())
        size_calls = sum(calls for calls, _ in sizes.values())
        size_bytes = sum(bytes_ for _, bytes_ in sizes.values())
        trace_size_calls = sum(sum(values.values()) for values in trace_sizes.values())
        trace_size_bytes = sum(size * count for values in trace_sizes.values() for size, count in values.items())
        for trace_id, (calls, requested) in trace_rows.items():
            values = trace_sizes[trace_id]
            require(sum(values.values()) == calls and
                    sum(size * count for size, count in values.items()) == requested,
                    f"{label} trace-size conservation failed for {trace_id:x}")
        require(trace_calls == size_calls == trace_size_calls == summary["allocation_events"],
                f"{label} call conservation failed")
        require(trace_bytes == size_bytes == trace_size_bytes == summary["requested_bytes_total"],
                f"{label} byte conservation failed")
        require(summary["raw_marker_counts"].get("+") == summary["allocation_events"]
                and summary["raw_marker_counts"].get("-") == summary["deallocation_events"]
                and summary["plus_without_metadata"] == 0
                and summary["a_plus_adjacency_mismatches"] == 0,
                f"{label} raw marker conservation failed")
        rows.append({
            "label": label, "summary": desc(root / "summary.json"),
            "allocation_events": summary["allocation_events"],
            "requested_bytes_total": summary["requested_bytes_total"],
            "deallocation_events": summary["deallocation_events"],
            "allocation_minus_deallocation": summary["allocation_minus_deallocation"],
            "trace_rows": len(trace_rows), "size_rows": len(sizes),
            "unknown_any_frame": summary["unknown_any_frame"],
            "unknown_leaf_frame": summary["unknown_leaf_frame"],
            "exact_disjoint_conservation": True,
        })
    return rows


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet", type=Path, default=Path("/home/zhuhe/litchi-performance-0949"))
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--attribution-root", type=Path)
    parser.add_argument("--attribution-tool", type=Path)
    parser.add_argument("--attribution-config", type=Path)
    args = parser.parse_args()
    packet = args.packet.absolute(); output = args.output.absolute()
    attribution_root = (args.attribution_root or (packet / "attribution-v1")).absolute()
    attribution_tool = (args.attribution_tool or (packet / "tools/attribute_heaptrack_0949_v1.py")).absolute()
    attribution_config = (args.attribution_config or (packet / "tools/attribution-config-v1.json")).absolute()
    require(not output.exists() and not output.is_symlink(), f"refusing existing output: {output}")
    plan, plan_desc, runner_desc = plan_and_outer(packet)
    index_rows, index = artifact_index_check(packet)
    child_meta = [check_child(packet, plan, plan_desc, runner_desc, child_id, index_rows)
                  for child_id in CHILD_IDS]
    reports = report_admission(packet, plan, child_meta)
    decoder = decoder_conservation(packet, child_meta)
    attr_terminal = attribution_root / "attribution-terminal.json"
    attr = read_json(attr_terminal)
    require(attr.get("status") == "terminal" and attr.get("schema") == "litchi.performance.0949.heaptrack-attribution-run.v1",
            "attribution terminal is not complete")
    require(sorted(attr.get("profiles_requested", [])) == sorted([r["label"] for r in decoder]),
            "attribution profile set differs from decoder profile set")
    decoder_by_label = {row["label"]: row for row in decoder}
    attribution_profiles = attr.get("profiles")
    require(isinstance(attribution_profiles, list), "attribution profiles are not a list")
    require({row.get("label") for row in attribution_profiles} == set(decoder_by_label),
            "attribution profile labels differ from decoder labels")
    attribution_checks = []
    config_descriptor = desc(attribution_config)
    for profile in attribution_profiles:
        label = profile["label"]
        raw = decoder_by_label[label]
        whole = profile.get("whole_process")
        conservation = profile.get("conservation")
        profile_config = profile.get("config")
        require(isinstance(whole, dict) and isinstance(conservation, dict),
                f"{label}: attribution summary missing whole-process/conservation")
        require(profile_config == config_descriptor,
                f"{label}: attribution config descriptor mismatch")
        require(whole.get("calls") == raw["allocation_events"]
                and whole.get("requested_bytes") == raw["requested_bytes_total"],
                f"{label}: attribution whole-process totals differ from decoder")
        require(conservation.get("all_exact") is True
                and conservation.get("decoder_allocation_events") == raw["allocation_events"]
                and conservation.get("decoder_requested_bytes") == raw["requested_bytes_total"]
                and conservation.get("size_aggregate_calls") == raw["allocation_events"]
                and conservation.get("trace_size_aggregate_calls") == raw["allocation_events"]
                and conservation.get("size_aggregate_requested_bytes") == raw["requested_bytes_total"]
                and conservation.get("trace_size_aggregate_requested_bytes") == raw["requested_bytes_total"],
                f"{label}: attribution conservation differs from decoder")
        external = conservation.get("external_heaptrack_histogram")
        require(isinstance(external, dict)
                and external.get("status") == "pass"
                and external.get("decoder_calls") == raw["allocation_events"]
                and external.get("heaptrack_print_calls") == raw["allocation_events"]
                and external.get("mismatch_count") == 0,
                f"{label}: attribution heaptrack histogram reconciliation failed")
        attribution_checks.append({
            "label": label,
            "profile": whole,
            "conservation": conservation,
            "config": profile_config,
            "exact_decoder_and_histogram_reconciliation": True,
        })
    result = {
        "schema": "litchi.performance.0949.independent-report-admission.v2",
        "attribution_root": str(attribution_root),
        "attribution_tool": desc(attribution_tool),
        "attribution_config": desc(attribution_config),
        "status": "pass",
        "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(timespec="microseconds"),
        "workload_launched": False,
        "plan": plan_desc,
        "runner": runner_desc,
        "artifact_index": desc(packet / "receipts/capture-artifact-index-v1.json"),
        "runner_terminal": desc(packet / "capture-v1/runner-terminal.json"),
        "child_count": len(child_meta),
        "reports": reports,
        "decoder": decoder,
        "decoder_terminal": desc(packet / "raw-decoder-v1/decoder-terminal.json"),
        "attribution_terminal": desc(attr_terminal),
        "attribution_profiles": attribution_checks,
        "attribution_seal": (desc(packet / "receipts/root-offline-attribution-seal-v1.json")
                             if (packet / "receipts/root-offline-attribution-seal-v1.json").is_file() else None),
        "heaptrack_print_reconciliation": {
            "status": "pending_separate_export",
            "scope": "This audit does not invoke heaptrack_print; the later owner must compare histogram/footer calls and weighted bytes to decoder totals and retain unknown/unattributed buckets.",
        },
        "scope": {
            "report_counters": "Exact deterministic 0948 publication operation counters (calls, deallocations, reallocations, failed calls, allocated and deallocated bytes) plus corpus/sink/output identities; one-sample publication timing and live gauges are descriptive only.",
            "heaptrack": "Whole-process raw requested allocations before frees; decoder totals are not publication-interval totals.",
            "no_claim": "No latency, speedup, adoption, or publication-only heaptrack attribution claim.",
        },
        "checks": {
            "artifact_index_full_sha_size_census": True,
            "terminal_first_and_no_retry": True,
            "report_binary_source_case_identity": True,
            "exact_0948_publication_counter_projection": True,
            "corpus_sink_output_semantic_identity": True,
            "raw_decoder_size_trace_trace_size_conservation": True,
            "unknown_and_deallocation_retained": True,
        },
    }
    data=(json.dumps(result, indent=2, sort_keys=True)+"\n").encode()
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("xb") as stream:
        stream.write(data); stream.flush(); os.fsync(stream.fileno())
    print(json.dumps({"output":str(output),"bytes":len(data),"sha256":hashlib.sha256(data).hexdigest()}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (AuditError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"0949 independent report admission failed: {error}")
        raise SystemExit(1)
