#!/usr/bin/env python3
"""Seal the successful 0889 evidence without rewriting any attempt receipt."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def read_json(path: Path) -> dict[str, Any]:
    require(path.is_file(), f"missing evidence: {path}")
    value = json.loads(path.read_text())
    require(isinstance(value, dict), f"evidence is not an object: {path}")
    return value


def receipt(name: str) -> dict[str, Any]:
    return read_json(PACKET / "attempts" / name / "validation-receipt.json")


def command(record: dict[str, Any], name: str) -> dict[str, Any]:
    found = next((item for item in record.get("commands", []) if item.get("name") == name), None)
    require(found is not None, f"missing command {name}")
    require(found.get("pass") is True, f"command did not pass: {name}")
    return found


def negative_report(
    record: dict[str, Any], index: int, stages: list[str], member: str | None = None
) -> dict[str, Any]:
    reports = record.get("negative_reports", [])
    require(len(reports) > index, f"missing negative report {index}")
    report = reports[index]
    require(report.get("qualification_complete") is False, "negative report is complete")
    require(report.get("control_failure_count") == 0, "negative report has control failures")
    failures = report.get("failures", [])
    require([failure.get("stage") for failure in failures] == stages,
            f"negative stages differ: {failures}")
    if member is not None:
        require(all(failure.get("member") == member for failure in failures),
                "negative member identity differs")
    descriptor = report.get("json")
    require(isinstance(descriptor, dict), "negative JSON descriptor missing")
    path = PACKET / descriptor["path"]
    require(path.is_file() and hashlib.sha256(path.read_bytes()).hexdigest() == descriptor["sha256"],
            f"negative JSON changed or missing: {path}")
    return report


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=PACKET / "qualification-receipt.json")
    parser.add_argument("--broad", default="validation-3")
    parser.add_argument("--unreadable", default="validation-5")
    parser.add_argument("--corrupt", default="validation-7")
    args = parser.parse_args()
    output = args.output if args.output.is_absolute() else ROOT / args.output
    require(not output.exists(), f"aggregate receipt already exists: {output}")

    broad = receipt(args.broad)
    unreadable = receipt(args.unreadable)
    corrupt = receipt(args.corrupt)
    repeated = receipt("validation-6")
    oracle = read_json(PACKET / "oracle-verification.json")
    census = read_json(PACKET / "independent-input-census.json")

    require(broad.get("status") == "fail", "broad attempt should retain its failed terminal receipt")
    require(broad.get("exception", "").startswith("PermissionError"),
            "broad attempt did not stop at the runner preparation failure")
    expected_gates = {
        "fmt_check", "cargo_test_bin", "clippy_bin_no_deps", "clippy_bin_tests_no_deps",
        "rustdoc_bin", "crate_boundary_check", "cargo_build_bin",
    }
    require({item.get("name") for item in broad.get("commands", [])[:7]} == expected_gates,
            "broad gate population differs")
    gates = [command(broad, name) for name in sorted(expected_gates)]
    command(broad, "differential_a")
    command(broad, "differential_b")
    command(broad, "negative_missing_root")
    differential = broad.get("differential")
    require(isinstance(differential, dict) and differential.get("byte_identical") is True,
            "broad repeat reports are not byte-identical")
    missing = negative_report(broad, 0, ["root_stat", "no_ooxml_package"])

    require(unreadable.get("status") == "fail", "unreadable attempt should retain its old failed receipt")
    unreadable_command = command(unreadable, "negative_unreadable_package")
    unreadable_report = negative_report(unreadable, 0, ["package_scan_read", "no_ooxml_package"])
    require(corrupt.get("status") == "pass", "corrupt continuation did not pass")
    corrupt_command = command(corrupt, "negative_corrupt_member")
    corrupt_report = negative_report(corrupt, 0, ["member_read"], "[Content_Types].xml")
    mutation = corrupt.get("corrupt_member", {})
    require(mutation.get("offset") == 569 and mutation.get("original_byte") == 204 and
            mutation.get("mutated_byte") == 206,
            "corrupt mutation is not the reserved-BTYPE recipe")
    require("invalid block type" in mutation.get("raw_read_error", ""),
            "raw ZIP refusal was not retained")

    require(repeated.get("status") == "pass", "repeated targeted attempt did not pass")
    require(broad.get("before", {}).get("source") == corrupt.get("after", {}).get("source"),
            "source identity changed across final evidence")
    runtime_binary = broad.get("runtime_binary")
    require(runtime_binary == corrupt.get("before", {}).get("binary") == corrupt.get("after", {}).get("binary"),
            "binary identity changed across final evidence")
    require(broad.get("environment") == unreadable.get("environment") == corrupt.get("environment"),
            "validation environment changed across evidence")
    require(census.get("package_count") == 3 and census.get("member_count") == 67,
            "independent input census differs")
    require(oracle.get("status") == "pass" and oracle.get("repeat_byte_equal") is True and
            oracle.get("missing_oracle_refused") is True and oracle.get("missing_controls_refused") is True,
            "independent oracle verification did not pass")
    for oracle_report, broad_report in zip(
        oracle.get("reports", []),
        (differential["compressed_first"], differential["compressed_second"]),
    ):
        require(
            oracle_report.get("path") == f"docs/performance/results/change-0889/{broad_report['path']}"
            and oracle_report.get("bytes") == broad_report.get("bytes")
            and oracle_report.get("sha256") == broad_report.get("sha256")
            and oracle_report.get("json_bytes") == broad_report.get("decompressed_bytes")
            and oracle_report.get("json_sha256") == broad_report.get("decompressed_sha256"),
            "oracle report is not bound to broad compressed output",
        )

    aggregate = {
        "schema": "litchi.mce-0889-qualification.v1",
        "status": "pass",
        "purpose": "correctness_and_bounded_input_qualification",
        "performance_claim": False,
        "base_commit": broad.get("before", {}).get("git_head"),
        "toolchain": broad.get("toolchain"),
        "environment": broad.get("environment"),
        "fixtures": broad.get("fixtures"),
        "input_census": {
            "path": "independent-input-census.json",
            "packages": census["package_count"],
            "members": census["member_count"],
        },
        "attempts": {
            "broad_gates_and_differential": {
                "receipt": f"attempts/{args.broad}/validation-receipt.json",
                "status": broad["status"],
                "expected_terminal_exception": broad["exception"],
                "passed_gates": gates,
                "missing_root": missing,
                "differential": differential,
            },
            "unreadable_package": {
                "receipt": f"attempts/{args.unreadable}/validation-receipt.json",
                "status": unreadable["status"],
                "command": unreadable_command,
                "report": unreadable_report,
            },
            "corrupt_member": {
                "receipt": f"attempts/{args.corrupt}/validation-receipt.json",
                "status": corrupt["status"],
                "command": corrupt_command,
                "report": corrupt_report,
                "mutation": mutation,
            },
            "repeated_targeted": {
                "receipt": "attempts/validation-6/validation-receipt.json",
                "status": repeated["status"],
                "purpose": "retained repeat of both negative cases; not used as primary evidence",
            },
        },
        "oracle_verification": {
            "path": "oracle-verification.json",
            "result": oracle,
        },
        "orphan_cleanup": unreadable.get("orphan_cleanup"),
        "source_identity": {
            "before": broad["before"],
            "after": corrupt["after"],
            "equal": broad["before"]["source"] == corrupt["after"]["source"],
        },
        "binary_identity": {
            "runtime": runtime_binary,
            "before": corrupt["before"]["binary"],
            "after": corrupt["after"]["binary"],
            "equal": True,
        },
        "qualification": {
            "gates_passed": True,
            "differential_repeat_byte_equal": True,
            "missing_root_refused": True,
            "unreadable_package_refused": True,
            "corrupt_member_refused": True,
            "control_failure_count": 0,
        },
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x") as destination:
        json.dump(aggregate, destination, indent=2, sort_keys=True)
        destination.write("\n")
    print(json.dumps({"aggregate": str(output), "status": aggregate["status"]}))


if __name__ == "__main__":
    main()
