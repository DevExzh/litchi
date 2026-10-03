#!/usr/bin/env python3
"""Run the final corrupt-member refusal after prior negative evidence is sealed."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import shutil
from typing import Any

import validation_runner as runner


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def read_receipt(attempt: str) -> tuple[Path, dict[str, Any]]:
    directory = runner.PACKET / "attempts" / attempt
    path = directory / "validation-receipt.json"
    require(path.is_file(), f"receipt is missing: {path}")
    return directory, json.loads(path.read_text())


def prove_previous_evidence(
    broad: dict[str, Any], unreadable: dict[str, Any], current: dict[str, Any], env: dict[str, str]
) -> dict[str, Any]:
    broad_before = broad.get("before")
    require(isinstance(broad_before, dict), "broad validation before identity is missing")
    require(current.get("git_head") == broad_before.get("git_head"), "git HEAD changed")
    require(current.get("source") == broad_before.get("source"), "source changed")
    require(runner.retained_env(env) == broad.get("environment"), "environment changed")
    require(current.get("binary") == broad.get("runtime_binary"), "binary changed")
    require(unreadable.get("before", {}).get("source") == broad_before.get("source"),
            "unreadable continuation source differs from broad source")
    require(unreadable.get("before", {}).get("binary") == broad.get("runtime_binary"),
            "unreadable continuation binary differs from broad binary")
    require(unreadable.get("negative_reports"), "unreadable continuation has no report")
    report = unreadable["negative_reports"][0]
    require([failure.get("stage") for failure in report.get("failures", [])] ==
            ["package_scan_read", "no_ooxml_package"],
            "unreadable continuation stages are not sealed")
    report_path = runner.PACKET / report["json"]["path"]
    require(report_path.is_file() and runner.sha256(report_path) == report["json"]["sha256"],
            "unreadable continuation JSON is missing or changed")
    return {
        "broad_receipt": "attempts/validation-3/validation-receipt.json",
        "unreadable_receipt": "attempts/validation-5/validation-receipt.json",
        "unreadable_report": report,
        "broad_differential": broad["differential"],
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", required=True, help="new create-only continuation name")
    parser.add_argument("--broad-attempt", default="validation-3")
    parser.add_argument("--unreadable-attempt", default="validation-5")
    args = parser.parse_args()
    for value in (args.attempt, args.broad_attempt, args.unreadable_attempt):
        if not value or value in {".", ".."} or "/" in value:
            raise SystemExit("attempt names must be one path component")

    attempt_dir = runner.PACKET / "attempts" / args.attempt
    try:
        attempt_dir.mkdir(parents=True, exist_ok=False)
    except FileExistsError as error:
        raise SystemExit(f"attempt already exists: {attempt_dir}") from error
    logs_dir = attempt_dir / "logs"
    runtime_dir = attempt_dir / "runtime"
    logs_dir.mkdir()
    runtime_dir.mkdir()

    broad_dir, broad = read_receipt(args.broad_attempt)
    unreadable_dir, unreadable = read_receipt(args.unreadable_attempt)
    marker = runner.TARGET / ".owner-0889"
    require(
        marker.is_file()
        and not marker.is_symlink()
        and marker.read_text() == "litchi 0889 validation target\n",
        f"owned validation target marker is missing or mismatched: {marker}",
    )
    env = runner.build_env()
    commands: list[dict[str, Any]] = []
    receipt: dict[str, Any] = {
        "schema": "litchi.mce-0889-corrupt-continuation.v1",
        "purpose": "complete_corrupt_member_qualification",
        "performance_claim": False,
        "started_utc": runner.utc_now(),
        "toolchain": runner.TOOLCHAIN,
        "target_dir": str(runner.TARGET),
        "target_marker": str(marker.relative_to(runner.ROOT)),
        "attempt": str(attempt_dir.relative_to(runner.PACKET)),
        "environment": runner.retained_env(env),
        "fixtures": [str(path.relative_to(runner.ROOT)) for path in runner.FIXTURES],
        "commands": commands,
        "status": "running",
    }
    receipt_path = attempt_dir / "validation-receipt.json"
    runner.write_receipt(receipt_path, receipt)
    failure: str | None = None
    try:
        receipt["runner"] = runner.capture_runner(attempt_dir)
        continuation_source = Path(__file__).resolve()
        continuation_copy = attempt_dir / continuation_source.name
        shutil.copy2(continuation_source, continuation_copy)
        receipt["continuation_runner"] = {
            "path": str(continuation_copy.relative_to(runner.PACKET)),
            "bytes": continuation_copy.stat().st_size,
            "sha256": runner.sha256(continuation_copy),
        }
        overlay = runner.capture_source_overlay(attempt_dir)
        receipt["source_overlay"] = {
            "path": str((attempt_dir / "source-overlay.json").relative_to(runner.PACKET)),
            **overlay,
        }
        before = runner.identity()
        receipt["reused_evidence"] = prove_previous_evidence(broad, unreadable, before, env)
        receipt["before"] = before
        runner.write_receipt(receipt_path, receipt)

        corrupt = runtime_dir / "corrupt-member.xlsx"
        mutation = runner.corrupt_one_zip_member(runner.FIXTURES[0], corrupt)
        receipt["corrupt_member"] = mutation
        runner.write_receipt(receipt_path, receipt)
        command = runner.run_command(
            "negative_corrupt_member",
            [
                str(runner.BINARY),
                "differential",
                "--json",
                str(runtime_dir / "corrupt.json"),
                "--generated",
                "0",
                "--aliasing",
                "0",
                str(corrupt),
            ],
            env,
            logs_dir / "negative-corrupt-member.log",
            commands,
            receipt,
            receipt_path,
            expected="fail",
        )
        require(command["pass"], "corrupt member was not refused")
        receipt["negative_reports"] = [
            runner.verify_negative_report(
                runtime_dir / "corrupt.json", ["member_read"], corrupt, mutation["member"]
            )
        ]
        corrupt.unlink()
        receipt["negative_input_cleanup"] = {
            "removed": [str(corrupt.relative_to(runner.PACKET))],
            "remaining": [str(corrupt) for _ in (1,) if corrupt.exists()],
        }
        after = runner.identity()
        receipt["after"] = after
        require(before["source"] == after["source"], "source changed during corrupt continuation")
        require(before["binary"] == after["binary"], "binary changed during corrupt continuation")
        runner.write_receipt(receipt_path, receipt)
    except Exception as error:
        failure = repr(error)
        receipt["exception"] = failure
    finally:
        receipt["finished_utc"] = runner.utc_now()
        receipt["command_count"] = len(commands)
        receipt["all_commands_passed"] = all(command["pass"] for command in commands)
        receipt["status"] = (
            "pass"
            if not failure and receipt["all_commands_passed"] and receipt.get("negative_reports")
            else "fail"
        )
        runner.write_receipt(receipt_path, receipt)

    print(json.dumps({"receipt": str(receipt_path), "commands": len(commands), "status": receipt["status"]}))
    if receipt["status"] != "pass":
        raise SystemExit(1)


if __name__ == "__main__":
    main()
