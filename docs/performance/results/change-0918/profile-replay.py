#!/usr/bin/env python3
"""Verify retained 0918 callgrind captures without running Valgrind."""

from __future__ import annotations

import hashlib
import json
import re
import sys
from pathlib import Path


PACKET = Path(__file__).resolve().parent
RECEIPT_PATH = PACKET / "profile-receipt.json"
REPOSITORY = PACKET.parents[3]
EVENT_NAMES = ("Ir", "Dr", "Dw", "I1mr", "D1mr", "D1mw", "ILmr", "DLmr", "DLmw")
EVENT_HEADER = " ".join(EVENT_NAMES)
RUN_SYMBOL = "doc_retention_probe_0917::run_case"
ENTRY_SYMBOL = "doc_retention_probe_0917::measure_one"
CALL_EDGE = "calls=1 0xf5d30"
CORRECTNESS_KEYS = (
    "all_stages_change_text",
    "exact_cross_route_outputs",
    "final_paragraph_semantic_readback",
    "final_public_reader_reopen",
    "exact_noop_source",
    "same_value_preserves_retention",
    "operation_limit_refusal_atomic",
    "noop_shares_source_allocation",
    "forward_patch_exact_output",
    "inverse_patch_exact_source",
    "foreign_patch_refused",
    "stale_patch_refused",
    "retained_capacities_within_ceiling",
    "zero_route_clears_retention",
    "release_route_clears_every_stage",
    "untouched_streams_preserved",
    "directory_metadata_preserved",
    "semantic_oracles_passed",
    "semantic_change_count_matches_edit_count",
    "allocator_retained_matches_output_capacity",
)
ORACLE_TRUE_KEYS = (
    "oracle_ok",
    "source_structurally_valid",
    "expected_structurally_valid",
    "output_structurally_valid",
    "semantic_reopen_ok",
    "semantic_directory_metadata_matches_expected",
    "unchanged_stream_bytes_match_expected",
    "unchanged_source_output_stream_bytes_match",
    "output_stream_paths_match_expected",
    "output_stream_bytes_match_expected",
    "root_clsid_preserved",
    "storage_clsids_preserved",
)
EXPECTED_TEXTS = (
    "litchi copy-through baseline replacement text",
    "litchi retained render second replacement text",
)


class VerificationError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def check_file(path: Path, expected_bytes: int, expected_sha: str, label: str) -> None:
    require(path.is_file(), f"{label}: missing file {path}")
    actual_bytes = path.stat().st_size
    require(actual_bytes == expected_bytes, f"{label}: bytes {actual_bytes} != {expected_bytes}")
    actual_sha = sha256(path)
    require(actual_sha == expected_sha, f"{label}: sha256 {actual_sha} != {expected_sha}")


def argument(argv: list[str], option: str) -> str:
    try:
        index = argv.index(option)
    except ValueError as error:
        raise VerificationError(f"argv: missing {option}") from error
    require(index + 1 < len(argv), f"argv: missing value after {option}")
    return argv[index + 1]


def parse_summary(raw: str, label: str) -> dict[str, int]:
    headers = re.findall(r"^events:\s*(.+)$", raw, re.MULTILINE)
    require(headers == [EVENT_HEADER], f"{label}: event order is not exactly {EVENT_HEADER}")
    summaries = re.findall(r"^summary:\s+([0-9 ]+)$", raw, re.MULTILINE)
    require(len(summaries) == 1, f"{label}: expected one callgrind summary")
    values = [int(value) for value in summaries[0].split()]
    require(len(values) == len(EVENT_NAMES), f"{label}: unexpected event count")
    require(values[0] > 0, f"{label}: Ir summary is zero")
    return dict(zip(EVENT_NAMES, values))


def integer(value: object, label: str, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{label}: expected integer >= {minimum}")
    return value


def validate_active_stage(stage: object, index: int, retention: str, ceiling: int, label: str) -> None:
    require(isinstance(stage, dict), f"{label}: active stage {index} is not an object")
    require(type(stage.get("stage")) is int and stage.get("stage") == index,
            f"{label}: active stage {index} sequence mismatch")
    for key in ("before", "after", "after_release"):
        value = stage.get(key)
        require(value is None or (type(value) is int and 0 <= value <= ceiling),
                f"{label}: active stage {index} invalid {key}")
    before = stage.get("before")
    after = stage.get("after")
    after_release = stage.get("after_release")
    if retention == "zero":
        require(before is None and after is None and after_release is None,
                f"{label}: zero route retained render at stage {index}")
    else:
        require(after is not None and after_release is None,
                f"{label}: {retention} route stage {index} missing retained render")
    if index == 1:
        require(before is None, f"{label}: first stage has a predecessor")
    elif retention == "default_8mib":
        require(before is not None, f"{label}: default stage {index} lost predecessor")
    else:
        require(before is None, f"{label}: cleared stage {index} retained predecessor")


def validate_inactive_stage(stage: object, index: int, label: str) -> None:
    require(isinstance(stage, dict), f"{label}: inactive stage {index} is not an object")
    require(type(stage.get("stage")) is int and stage.get("stage") == 0,
            f"{label}: inactive stage {index} has a nonzero frame")
    require(type(stage.get("requested_text")) is str and stage.get("requested_text") == "",
            f"{label}: inactive stage {index} is not the empty string marker")
    require(stage.get("before") is None and stage.get("after") is None
            and stage.get("after_release") is None,
            f"{label}: inactive stage {index} contains retained values")


def validate_oracle(oracle: object, label: str) -> None:
    require(isinstance(oracle, dict), f"{label}: oracle is not an object")
    require(set(oracle) == set(ORACLE_TRUE_KEYS) | {"failure_reasons"},
            f"{label}: oracle key set changed")
    require(all(oracle.get(key) is True for key in ORACLE_TRUE_KEYS),
            f"{label}: oracle gate failed")
    require(type(oracle.get("failure_reasons")) is list and oracle["failure_reasons"] == [],
            f"{label}: oracle has failure reasons")


def validate_report(report: object, run_id: str, case_name: str, edits: int,
                    input_path: Path, fixture: dict, probe: dict) -> None:
    require(isinstance(report, dict), f"{run_id}: JSON report is not an object")
    require(report.get("schema_version") == 1
            and report.get("schema") == "litchi.performance.0917.doc_retention.v1",
            f"{run_id}: schema mismatch")
    require(report.get("change") == "0917", f"{run_id}: probe change mismatch")
    require(report.get("case") == case_name and type(report.get("edits")) is int
            and report.get("edits") == edits and edits in (2, 32),
            f"{run_id}: case or edit count mismatch")
    require(type(report.get("selected_paragraph")) is int
            and report.get("selected_paragraph") == 0,
            f"{run_id}: selected paragraph mismatch")
    require(report.get("input") == str(input_path), f"{run_id}: JSON input mismatch")
    require(report.get("source_sha256") == fixture["sha256"]
            and report.get("source_bytes") == fixture["bytes"],
            f"{run_id}: source identity mismatch")
    require(report.get("alternating_texts") == list(EXPECTED_TEXTS),
            f"{run_id}: edit sequence changed")
    require(report.get("retention") == "default_8mib"
            and report.get("retention_ceiling_bytes") == 8 * 1024 * 1024,
            f"{run_id}: retention identity mismatch")
    require(report.get("warmups") == 0 and report.get("samples_requested") == 1,
            f"{run_id}: warmup/sample request mismatch")
    require(report.get("timing_claim") is True
            and report.get("allocator_instrumented") is False
            and report.get("allocator_timing_claim") is False,
            f"{run_id}: timing/allocator gate mismatch")
    require(isinstance(report.get("region_definition"), str),
            f"{run_id}: region definition missing")
    require(report.get("allocation") is None, f"{run_id}: native report carries allocation metrics")

    output_hash = report.get("output_sha256")
    reference_hash = report.get("reference_output_sha256")
    output_bytes = report.get("output_bytes")
    reference_bytes = report.get("reference_output_bytes")
    output_capacity = report.get("output_capacity")
    require(output_hash == reference_hash and type(output_bytes) is int
            and type(reference_bytes) is int and output_bytes == reference_bytes
            and type(output_capacity) is int and output_capacity >= output_bytes
            and report.get("exact_output_equal_reference") is True,
            f"{run_id}: output equality gate failed")
    require(output_hash == probe["output_sha256"] and output_bytes == probe["output_bytes"]
            and output_capacity == probe["output_capacity"],
            f"{run_id}: output differs from receipt")

    correctness = report.get("correctness")
    require(isinstance(correctness, dict) and set(correctness) == set(CORRECTNESS_KEYS)
            and all(type(correctness[key]) is bool and correctness[key]
                    for key in CORRECTNESS_KEYS),
            f"{run_id}: correctness key set or gate changed")

    elapsed = report.get("elapsed_ns")
    require(isinstance(elapsed, list) and len(elapsed) == 1
            and type(elapsed[0]) is int and elapsed[0] > 0,
            f"{run_id}: elapsed_ns vector is not exactly one positive sample")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1,
            f"{run_id}: post-sample frame count is not exactly one")
    sample = samples[0]
    require(isinstance(sample, dict) and sample.get("index") == 0
            and sample.get("elapsed_ns") == elapsed[0],
            f"{run_id}: sample frame mismatch")
    require(sample.get("output_sha256") == reference_hash
            and sample.get("output_bytes") == output_bytes
            and sample.get("output_capacity") == output_capacity
            and sample.get("exact_output_equal_reference") is True
            and sample.get("semantic_change_count") == edits,
            f"{run_id}: sample output gate failed")
    for key in ("final_paragraph_semantic_readback", "final_public_reader_reopen",
                "retained_capacities_within_ceiling", "allocator_retained_matches_output_capacity"):
        require(sample.get(key) is True, f"{run_id}: sample gate failed: {key}")
    require(sample.get("allocation") is None, f"{run_id}: sample carries allocation metrics")
    validate_oracle(sample.get("oracle"), f"{run_id}: sample")

    expected_routes = {"zero": 0, "default_8mib": 8 * 1024 * 1024,
                       "release_8mib": 8 * 1024 * 1024}
    route_outputs = report.get("route_outputs")
    require(isinstance(route_outputs, list) and len(route_outputs) == 3,
            f"{run_id}: route output set mismatch")
    route_by_name: dict[str, dict] = {}
    for route_output in route_outputs:
        require(isinstance(route_output, dict), f"{run_id}: route output malformed")
        route = route_output.get("retention")
        require(route in expected_routes and route not in route_by_name,
                f"{run_id}: duplicate or unknown route output")
        route_by_name[route] = route_output
        ceiling = expected_routes[route]
        require(route_output.get("retention_ceiling_bytes") == ceiling,
                f"{run_id}: route ceiling mismatch: {route}")
        require(route_output.get("output_sha256") == reference_hash
                and route_output.get("output_bytes") == reference_bytes
                and type(route_output.get("output_capacity")) is int
                and route_output["output_capacity"] >= reference_bytes,
                f"{run_id}: route output mismatch: {route}")
        require(route_output.get("semantic_change_count") == edits,
                f"{run_id}: route edit count mismatch: {route}")
        validate_oracle(route_output.get("oracle"), f"{run_id}: route {route}")
        stages = route_output.get("stages")
        require(isinstance(stages, list) and len(stages) == 32,
                f"{run_id}: route stage count mismatch: {route}")
        for index, stage in enumerate(stages[:edits], start=1):
            require(isinstance(stage, dict)
                    and stage.get("requested_text") == EXPECTED_TEXTS[(index - 1) % 2],
                    f"{run_id}: active stage text mismatch: {route}/{index}")
            validate_active_stage(stage, index, route, ceiling, run_id)
        for index, stage in enumerate(stages[edits:], start=edits + 1):
            validate_inactive_stage(stage, index, f"{run_id}: {route}")
        route_final = route_output.get("final_retained_before_commit")
        require(route_final is None or (type(route_final) is int and 0 <= route_final <= ceiling),
                f"{run_id}: route final capacity invalid: {route}")
        if route in ("zero", "release_8mib"):
            require(route_final is None, f"{run_id}: cleared route retained final render: {route}")
        else:
            require(route_final is not None, f"{run_id}: default route final render missing")
    require(set(route_by_name) == set(expected_routes), f"{run_id}: route set incomplete")
    require(route_by_name["default_8mib"]["output_capacity"] == output_capacity,
            f"{run_id}: selected route capacity mismatch")

    sample_stages = sample.get("stages")
    require(isinstance(sample_stages, list) and len(sample_stages) == 32,
            f"{run_id}: sample stage count mismatch")
    for index, stage in enumerate(sample_stages[:edits], start=1):
        require(isinstance(stage, dict)
                and stage.get("requested_text") == EXPECTED_TEXTS[(index - 1) % 2],
                f"{run_id}: sample active stage text mismatch: {index}")
        validate_active_stage(stage, index, "default_8mib", 8 * 1024 * 1024, run_id)
    for index, stage in enumerate(sample_stages[edits:], start=edits + 1):
        validate_inactive_stage(stage, index, f"{run_id}: sample")


def validate_run(run: dict, receipt: dict, fixture_paths: dict[str, Path]) -> dict:
    run_id = run["id"]
    argv = run["argv"]
    require(run["exit_code"] == 0, f"{run_id}: exit code is not zero")
    for required in (
        "--tool=callgrind",
        "--instr-atstart=yes",
        "--collect-atstart=no",
        "--toggle-collect=*measure_one*",
        "--collect-jumps=yes",
        "--simulate-cache=yes",
        "--dump-instr=yes",
    ):
        require(required in argv, f"{run_id}: argv missing {required}")

    case_name = argument(argv, "--case")
    edits = int(argument(argv, "--edits"))
    input_path = Path(argument(argv, "--input"))
    expected_input = fixture_paths[case_name]
    require(input_path == expected_input, f"{run_id}: input path differs from fixture receipt")
    require(argument(argv, "--retention") == "default", f"{run_id}: retention is not default")
    require(argument(argv, "--warmup") == "0", f"{run_id}: warmup is not zero")
    require(argument(argv, "--samples") == "1", f"{run_id}: samples is not one")

    raw_info = run["raw"]
    raw_path = Path(raw_info["path"])
    check_file(raw_path, raw_info["bytes"], raw_info["sha256"], f"{run_id} raw")
    raw_text = raw_path.read_text(errors="replace")
    events = parse_summary(raw_text, run_id)
    require(set(run["events"]) == set(EVENT_NAMES) and events == run["events"],
            f"{run_id}: summary differs from its nine-event receipt")
    require(f") {RUN_SYMBOL}" in raw_text, f"{run_id}: run_case definition missing")
    require(f") {ENTRY_SYMBOL}" in raw_text, f"{run_id}: measure_one definition missing")
    require(CALL_EDGE in raw_text, f"{run_id}: measure_one -> run_case calls=1 edge missing")
    scope = run.get("scope_proof", {})
    scope_calls = scope.get("run_case_calls", 1)
    require(scope_calls == 1 and scope.get("edge", CALL_EDGE) == CALL_EDGE,
            f"{run_id}: receipt scope is not one run_case")

    output_path = Path(run["probe"]["stdout_path"])
    probe = run["probe"]
    check_file(output_path, probe["stdout_bytes"], probe["stdout_sha256"], f"{run_id} stdout")
    stderr_info = run["stderr"]
    check_file(Path(stderr_info["path"]), stderr_info["bytes"], stderr_info["sha256"], f"{run_id} stderr")
    report = json.loads(output_path.read_text())
    fixture = receipt["fixtures"][case_name]
    validate_report(report, run_id, case_name, edits, input_path, fixture, probe)

    expected_raw = f"--callgrind-out-file={raw_path}"
    require(expected_raw in argv, f"{run_id}: argv does not bind raw output path")
    require(str(receipt["binary"]["path"]) in argv, f"{run_id}: argv does not bind retained binary")
    return {
        "id": run_id,
        "case": case_name,
        "edits": edits,
        "raw_path": str(raw_path),
        "raw_bytes": raw_info["bytes"],
        "raw_sha256": raw_info["sha256"],
        "stderr_path": stderr_info["path"],
        "stderr_bytes": stderr_info["bytes"],
        "stderr_sha256": stderr_info["sha256"],
        "run_case_calls": scope_calls,
        "output_sha256": probe["output_sha256"],
        "output_bytes": probe["output_bytes"],
        "events": events,
    }


def validate_receipt() -> tuple[dict, list[dict], list[dict], list[dict]]:
    receipt = json.loads(RECEIPT_PATH.read_text())
    require(receipt["schema"] == "litchi.performance.0918.callgrind.v1", "receipt schema mismatch")
    binary = Path(receipt["binary"]["path"])
    check_file(binary, receipt["binary"]["bytes"], receipt["binary"]["sha256"], "binary")
    fixture_paths = {
        name: REPOSITORY / details["path"] for name, details in receipt["fixtures"].items()
    }
    for name, path in fixture_paths.items():
        details = receipt["fixtures"][name]
        check_file(path, details["bytes"], details["sha256"], f"fixture {name}")

    runs = receipt["runs"] + receipt.get("repeats", [])
    summaries = [validate_run(run, receipt, fixture_paths) for run in runs]
    by_logical_id: dict[str, list[dict]] = {}
    for run, summary in zip(runs, summaries):
        logical_id = run["id"].removesuffix("-reproduce")
        by_logical_id.setdefault(logical_id, []).append({"run": run, "summary": summary})
    output_guards = []
    for logical_id, entries in by_logical_id.items():
        baseline = entries[0]["run"]["probe"]
        for entry in entries[1:]:
            candidate = entry["run"]["probe"]
            require(candidate["output_sha256"] == baseline["output_sha256"], f"{logical_id}: repeat output hash differs")
            require(candidate["output_bytes"] == baseline["output_bytes"], f"{logical_id}: repeat output bytes differ")
            require(candidate["output_capacity"] == baseline["output_capacity"], f"{logical_id}: repeat output capacity differs")
        output_guards.append({"id": logical_id, "runs": [entry["run"]["id"] for entry in entries], "output_sha256": baseline["output_sha256"]})
    return receipt, runs, summaries, output_guards


def main() -> int:
    _, _, summaries, output_guards = validate_receipt()
    print(json.dumps({"verified": True, "runs": summaries, "output_guards": output_guards}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (KeyError, TypeError, ValueError, VerificationError) as error:
        print(f"profile-replay: verification failed: {error}", file=sys.stderr)
        raise SystemExit(1)
