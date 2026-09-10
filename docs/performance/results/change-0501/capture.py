#!/usr/bin/env python3
"""Capture one frozen 0501 provider-lifecycle phase serially."""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from custody import HERE, REPO, artifact, canonical_json, now, sha_file, source_identity, source_snapshot


TMP_ROOT = Path("/tmp/litchi-goal-0501")
MAX_REPORT_BYTES = 64 * 1024 * 1024
MAX_LOG_BYTES = 16 * 1024 * 1024


def load_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def lane_name(lane: dict[str, Any]) -> str:
    return f"{lane['corpus']}-{lane['provider_label']}-{lane['repeat'].lower()}"


def lanes(protocol: dict[str, Any], include_range: bool) -> list[dict[str, Any]]:
    result = list(protocol["core_order"])
    if include_range:
        result.extend(protocol["optional_range"]["order"])
    return result


def filesystem_evidence(path: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    for key, command in (
        ("statfs_type", ["stat", "-f", "-c", "%T", str(path)]),
        ("mount", ["findmnt", "-T", str(path), "-n", "-o", "SOURCE,FSTYPE,TARGET"]),
    ):
        try:
            result[key] = subprocess.check_output(command, text=True, stderr=subprocess.STDOUT).strip()
        except (OSError, subprocess.CalledProcessError) as error:
            result[key] = f"unavailable: {error}"
    return result


def check_fresh(path: Path) -> None:
    if path.exists():
        raise RuntimeError(f"refusing to replace existing artifact: {path}")


def invoke(phase: str, include_range: bool) -> None:
    protocol = json.loads((HERE / "protocol.json").read_text())
    freeze = json.loads((HERE / f"{phase}-freeze.json").read_text())
    if freeze["change"] != 501 or freeze["phase"] != phase:
        raise RuntimeError("freeze identity mismatch")
    if sha_file(HERE / "protocol.json") != freeze["protocol_sha256"]:
        raise RuntimeError("protocol changed after freeze")
    if source_identity(source_snapshot()) != freeze["source_manifest"]:
        raise RuntimeError("source changed after executable freeze")
    binary = Path(freeze["binary"]["path"])
    if sha_file(binary) != freeze["binary"]["sha256"] or binary.stat().st_size != freeze["binary"]["bytes"]:
        raise RuntimeError("frozen executable changed")

    verifier = load_module("verify_report", HERE / "verify-report.py")
    destination = HERE / phase
    destination.mkdir(exist_ok=False)
    run_root = TMP_ROOT / "corpora" / phase
    run_root.mkdir(parents=True, exist_ok=False)
    previous_finished: float | None = None
    try:
        selected = lanes(protocol, include_range)
        for lane in selected:
            name = lane_name(lane)
            report_path = destination / f"{name}.report.json"
            receipt_path = destination / f"{name}.receipt.json"
            resource_path = destination / f"{name}.time.txt"
            stdout_path = destination / f"{name}.stdout.txt"
            oracle_path = destination / f"{name}.oracle.txt"
            for path in (report_path, receipt_path, resource_path, stdout_path, oracle_path):
                check_fresh(path)
            scratch = run_root / name
            scratch.mkdir(exist_ok=False)
            tmpdir = scratch / "tmp"
            tmpdir.mkdir()
            output = report_path.resolve()
            command = [
                "taskset",
                "-c",
                protocol["cpu_set"],
                "/usr/bin/time",
                "-v",
                "-o",
                str(resource_path.resolve()),
                str(binary),
                "provider-lifecycle",
                "--corpus",
                lane["corpus"],
                "--provider",
                lane["provider"],
                "--samples",
                str(protocol["samples"]),
                "--warmup",
                str(protocol["warmups"]),
                "--source-revision",
                freeze["revision"],
                "--output",
                str(output),
            ]
            if lane["provider"] == "range":
                command.extend(["--max-range", str(protocol["optional_range"]["max_range_bytes"]), "--delay-us", str(protocol["optional_range"]["delay_us"])])
            environment = os.environ.copy()
            environment.update(
                {
                    "TMPDIR": str(tmpdir.resolve()),
                    "RUSTUP_TOOLCHAIN": protocol["rust_toolchain"],
                    "CARGO_INCREMENTAL": "0",
                    "CARGO_PROFILE_RELEASE_DEBUG": "0",
                    "CARGO_BUILD_JOBS": "2",
                    "PYTHONDONTWRITEBYTECODE": "1",
                    "LC_ALL": "C",
                    "TZ": "UTC",
                }
            )
            receipt = {
                "change": 501,
                "phase": phase,
                "lane": lane,
                "name": name,
                "command": command,
                "environment": {"TMPDIR": str(tmpdir.resolve()), "LC_ALL": "C", "TZ": "UTC"},
                "tmpdir_filesystem": filesystem_evidence(tmpdir),
                "binary": freeze["binary"],
                "revision": freeze["revision"],
                "protocol_sha256": freeze["protocol_sha256"],
                "source_manifest_sha256": freeze["source_manifest"]["sha256"],
                "capture_sha256": sha_file(Path(__file__)),
                "verifier_sha256": sha_file(HERE / "verify-report.py"),
                "started_utc": now(),
                "status": "running",
            }
            receipt_path.write_bytes(canonical_json(receipt))
            started = time.monotonic()
            print(f"START {phase} {name}", flush=True)
            try:
                with stdout_path.open("xb") as stream:
                    result = subprocess.run(command, cwd=REPO, env=environment, stdout=stream, stderr=subprocess.STDOUT)
                receipt["exit_code"] = result.returncode
                if result.returncode != 0:
                    raise RuntimeError(f"provider lifecycle exited {result.returncode}: {name}")
                if not report_path.is_file() or report_path.stat().st_size > MAX_REPORT_BYTES:
                    raise RuntimeError(f"missing or oversized report: {name}")
                with oracle_path.open("xb") as stream:
                    checked = verifier.check_report(verifier.load(report_path))
                    stream.write((json.dumps({"status": "pass", **checked}, sort_keys=True) + "\n").encode())
                receipt["status"] = "pass"
            except Exception as error:
                receipt["error"] = str(error)
                raise
            finally:
                finished = time.monotonic()
                receipt["finished_utc"] = now()
                receipt["elapsed_driver_seconds"] = finished - started
                receipt["artifacts"] = {
                    key: artifact(path)
                    for key, path in (
                        ("report", report_path),
                        ("resource", resource_path),
                        ("stdout", stdout_path),
                        ("oracle", oracle_path),
                    )
                    if path.is_file()
                }
                for path in receipt["artifacts"].values():
                    if path["bytes"] > (MAX_REPORT_BYTES if path["path"].endswith("report.json") else MAX_LOG_BYTES):
                        receipt["status"] = "failed"
                        receipt["error"] = f"oversized retained artifact: {path['path']}"
                receipt["source_after"] = source_identity(source_snapshot())
                receipt["source_unchanged"] = receipt["source_after"] == freeze["source_manifest"]
                if not receipt["source_unchanged"]:
                    receipt["status"] = "failed"
                    receipt["error"] = "source changed during capture"
                receipt_path.write_bytes(canonical_json(receipt))
                if previous_finished is not None and started < previous_finished:
                    raise RuntimeError("capture chronology moved backwards")
                previous_finished = finished
                if receipt["status"] == "pass" and scratch.exists():
                    if tmpdir.exists() and not any(tmpdir.iterdir()):
                        tmpdir.rmdir()
                    leftovers = list(scratch.rglob("*"))
                    if leftovers:
                        receipt["cleanup_verified"] = False
                        receipt["leftovers"] = [str(path) for path in leftovers]
                        receipt_path.write_bytes(canonical_json(receipt))
                        raise RuntimeError(f"temporary provider files remain: {name}")
                    scratch.rmdir()
                    receipt["cleanup_verified"] = True
                    receipt_path.write_bytes(canonical_json(receipt))
            print(f"FINISH {phase} {name}", flush=True)
    finally:
        if run_root.exists() and not any(run_root.iterdir()):
            run_root.rmdir()
    print(json.dumps({"status": "pass", "phase": phase, "reports": len(selected), "include_range": include_range}, sort_keys=True))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("phase", choices=("before", "after"))
    parser.add_argument("--include-range", action="store_true")
    args = parser.parse_args()
    invoke(args.phase, args.include_range)


if __name__ == "__main__":
    main()
