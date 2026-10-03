#!/usr/bin/env python3
"""Run only the negative cases omitted by a failed 0889 runtime attempt."""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import shutil
import stat
import sys
from typing import Any

import validation_runner as runner


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def load_previous(attempt: str) -> tuple[Path, dict[str, Any]]:
    previous_dir = runner.PACKET / "attempts" / attempt
    receipt_path = previous_dir / "validation-receipt.json"
    require(receipt_path.is_file(), f"previous receipt is missing: {receipt_path}")
    receipt = json.loads(receipt_path.read_text())
    require(receipt.get("status") == "fail", "previous attempt is not a failed diagnostic run")
    require(
        receipt.get("exception", "").startswith("PermissionError"),
        "previous attempt did not stop at the unreadable-fixture preparation",
    )
    expected_commands = {
        "fmt_check",
        "cargo_test_bin",
        "clippy_bin_no_deps",
        "clippy_bin_tests_no_deps",
        "rustdoc_bin",
        "crate_boundary_check",
        "cargo_build_bin",
        "differential_a",
        "differential_b",
        "negative_missing_root",
    }
    commands = receipt.get("commands", [])
    require({command.get("name") for command in commands} == expected_commands,
            "previous attempt does not contain the complete passed prefix")
    require(all(command.get("pass") for command in commands),
            "previous attempt contains a failed command")
    require(receipt.get("gate_failures") is None, "previous attempt has gate failures")
    require(receipt.get("runtime_failures") is None, "previous attempt has runtime failures")
    differential = receipt.get("differential")
    require(isinstance(differential, dict) and differential.get("byte_identical") is True,
            "previous differential reports were not sealed byte-identically")
    for descriptor_name in ("compressed_first", "compressed_second"):
        descriptor = differential.get(descriptor_name)
        require(isinstance(descriptor, dict), f"missing previous {descriptor_name}")
        path = runner.PACKET / descriptor["path"]
        require(path.is_file() and runner.sha256(path) == descriptor["sha256"],
                f"previous compressed report is missing or changed: {path}")
    return previous_dir, receipt


def prove_same_inputs(previous: dict[str, Any], current: dict[str, Any], env: dict[str, str]) -> None:
    previous_before = previous.get("before")
    require(isinstance(previous_before, dict), "previous before identity is missing")
    require(current.get("git_head") == previous_before.get("git_head"),
            "git HEAD changed since the broad validation")
    require(current.get("source") == previous_before.get("source"),
            "diagnostic source changed since the broad validation")
    require(runner.retained_env(env) == previous.get("environment"),
            "validation environment changed since the broad validation")
    require(current.get("binary") == previous.get("runtime_binary"),
            "diagnostic binary changed since the broad validation")


def cleanup_orphan(previous_dir: Path) -> dict[str, Any]:
    orphan = previous_dir / "runtime/unreadable/fixture.xlsx"
    directory = orphan.parent
    if not orphan.exists():
        return {"path": str(orphan.relative_to(runner.PACKET)), "present": False}
    metadata = orphan.lstat()
    require(stat.S_ISREG(metadata.st_mode) and not stat.S_ISLNK(metadata.st_mode),
            f"orphan unreadable fixture is not a regular file: {orphan}")
    require(metadata.st_uid == os.getuid(), f"orphan unreadable fixture owner differs: {orphan}")
    source = runner.FIXTURES[0]
    source_bytes = source.stat().st_size
    source_sha256 = runner.sha256(source)
    require(metadata.st_size == source_bytes, "orphan unreadable fixture size differs from source")
    orphan.chmod(stat.S_IRUSR | stat.S_IWUSR)
    restored_mode = oct(stat.S_IMODE(orphan.stat().st_mode))
    require(runner.sha256(orphan) == source_sha256,
            "orphan unreadable fixture content differs from source")
    orphan.unlink()
    require(not orphan.exists(), "orphan unreadable fixture remains after cleanup")
    directory.rmdir()
    return {
        "path": str(orphan.relative_to(runner.PACKET)),
        "present": True,
        "source_fixture": str(source.relative_to(runner.ROOT)),
        "bytes": source_bytes,
        "sha256": source_sha256,
        "restored_mode": restored_mode,
        "removed": True,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", required=True, help="new create-only continuation name")
    parser.add_argument("--from-attempt", required=True, help="failed broad attempt to continue")
    args = parser.parse_args()
    for value in (args.attempt, args.from_attempt):
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

    previous_dir, previous = load_previous(args.from_attempt)
    marker = runner.TARGET / ".owner-0889"
    require(
        marker.is_file()
        and not marker.is_symlink()
        and marker.read_text() == "litchi 0889 validation target\n",
            f"owned validation target marker is missing or mismatched: {marker}")
    env = runner.build_env()
    commands: list[dict[str, Any]] = []
    receipt: dict[str, Any] = {
        "schema": "litchi.mce-0889-negative-continuation.v1",
        "purpose": "complete_negative_input_qualification",
        "performance_claim": False,
        "started_utc": runner.utc_now(),
        "toolchain": runner.TOOLCHAIN,
        "target_dir": str(runner.TARGET),
        "target_marker": str(marker.relative_to(runner.ROOT)),
        "attempt": str(attempt_dir.relative_to(runner.PACKET)),
        "continuation_of": str((previous_dir / "validation-receipt.json").relative_to(runner.PACKET)),
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
        prove_same_inputs(previous, before, env)
        receipt["before"] = before
        receipt["orchestration"] = {
            "broad_attempt": str((previous_dir / "validation-receipt.json").relative_to(runner.PACKET)),
            "broad_differential": previous["differential"],
            "broad_commands_reused": [command["name"] for command in previous["commands"]],
        }
        receipt["orphan_cleanup"] = cleanup_orphan(previous_dir)
        runner.write_receipt(receipt_path, receipt)

        def execute(name: str, argv: list[str], log_name: str) -> dict[str, Any]:
            return runner.run_command(
                name,
                argv,
                env,
                logs_dir / log_name,
                commands,
                receipt,
                receipt_path,
                expected="fail",
            )

        unreadable_dir = runtime_dir / "unreadable"
        unreadable_dir.mkdir()
        unreadable = unreadable_dir / "fixture.xlsx"
        shutil.copy2(runner.FIXTURES[0], unreadable)
        unreadable_bytes = unreadable.stat().st_size
        unreadable_sha256 = runner.sha256(unreadable)
        unreadable.chmod(0)
        receipt["negative_inputs"] = [
            {
                "kind": "unreadable_package",
                "path": str(unreadable.relative_to(runner.PACKET)),
                "bytes": unreadable_bytes,
                "sha256": unreadable_sha256,
                "mode": oct(stat.S_IMODE(unreadable.stat().st_mode)),
                "source_fixture": str(runner.FIXTURES[0].relative_to(runner.ROOT)),
            }
        ]
        unreadable_command = execute(
            "negative_unreadable_package",
            [str(runner.BINARY), "differential", "--json", str(runtime_dir / "unreadable.json"), "--generated", "0", "--aliasing", "0", str(unreadable)],
            "negative-unreadable-package.log",
        )
        require(unreadable_command["pass"], "unreadable package was not refused")
        receipt["negative_reports"] = [
            runner.verify_negative_report(
                runtime_dir / "unreadable.json",
                ["package_scan_read", "no_ooxml_package"],
                unreadable,
            )
        ]

        corrupt = runtime_dir / "corrupt-member.xlsx"
        mutation = runner.corrupt_one_zip_member(runner.FIXTURES[0], corrupt)
        receipt["corrupt_member"] = mutation
        corrupt_command = execute(
            "negative_corrupt_member",
            [str(runner.BINARY), "differential", "--json", str(runtime_dir / "corrupt.json"), "--generated", "0", "--aliasing", "0", str(corrupt)],
            "negative-corrupt-member.log",
        )
        require(corrupt_command["pass"], "corrupt member was not refused")
        receipt["negative_reports"].append(
            runner.verify_negative_report(
                runtime_dir / "corrupt.json", ["member_read"], corrupt, mutation["member"]
            )
        )
        unreadable.unlink()
        unreadable_dir.rmdir()
        corrupt.unlink()
        receipt["negative_input_cleanup"] = {
            "removed": [
                str(unreadable.relative_to(runner.PACKET)),
                str(corrupt.relative_to(runner.PACKET)),
            ],
            "remaining": [str(path) for path in (unreadable, corrupt) if path.exists()],
        }
        after = runner.identity()
        receipt["after"] = after
        require(before["source"] == after["source"], "source changed during negative continuation")
        require(before["binary"] == after["binary"], "binary changed during negative continuation")
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
