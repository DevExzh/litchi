#!/usr/bin/env python3
"""Capture bounded whole-child perf statistics and SHA callchains."""

from __future__ import annotations

import argparse
import importlib.util
import json
import os
import re
import subprocess
import time
from pathlib import Path
from typing import Any

from custody import HERE, REPO, artifact, canonical_json, now, sha_file, source_identity, source_snapshot


TMP_ROOT = Path("/tmp/litchi-goal-0501")
MAX_RAW_PERF = 64 * 1024 * 1024
MAX_COMPACT = 16 * 1024 * 1024
SHA_TOKENS = ("digest_touched", "digest_bytes", "sha2", "Sha256")


def load_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def profile_lanes(protocol: dict[str, Any]) -> list[dict[str, Any]]:
    result = []
    for corpus in protocol["profile"]["corpora"]:
        for provider in protocol["profile"]["providers"]:
            label = "owned" if provider == "bytes" else "file-warm"
            result.append({"corpus": corpus, "provider": provider, "provider_label": label})
    return result


def name(lane: dict[str, Any]) -> str:
    return f"{lane['corpus']}-{lane['provider_label']}"


def run_one(phase: str, lane: dict[str, Any], include_sha_requirement: bool) -> dict[str, Any]:
    protocol = json.loads((HERE / "protocol.json").read_text())
    freeze = json.loads((HERE / f"{phase}-freeze.json").read_text())
    if sha_file(HERE / "protocol.json") != freeze["protocol_sha256"]:
        raise RuntimeError("protocol changed after freeze")
    if source_identity(source_snapshot()) != freeze["source_manifest"]:
        raise RuntimeError("source changed after executable freeze")
    binary = Path(freeze["binary"]["path"])
    if sha_file(binary) != freeze["binary"]["sha256"]:
        raise RuntimeError("frozen executable changed")
    verifier = load_module("verify_report_profile", HERE / "verify-report.py")
    folder = HERE / "profiles" / phase
    folder.mkdir(parents=True, exist_ok=True)
    profile_name = name(lane)
    receipt_path = folder / f"{profile_name}.receipt.json"
    check_paths = [receipt_path]
    for path in check_paths:
        if path.exists():
            raise RuntimeError(f"refusing to replace {path}")
    raw_root = TMP_ROOT / "profiles" / phase / profile_name
    raw_root.mkdir(parents=True, exist_ok=False)
    tmpdir = raw_root / "tmp"
    tmpdir.mkdir()
    outputs = {
        "stat": folder / f"{profile_name}.perf-stat.txt",
        "report": folder / f"{profile_name}.perf-report.txt",
        "resource_stat": folder / f"{profile_name}.stat.time.txt",
        "resource_record": folder / f"{profile_name}.record.time.txt",
        "stat_stdout": folder / f"{profile_name}.stat.stdout.txt",
        "record_stdout": folder / f"{profile_name}.record.stdout.txt",
        "report_json": folder / f"{profile_name}.report.json",
        "record_report": folder / f"{profile_name}.record.report.json",
    }
    for path in outputs.values():
        if path.exists():
            raise RuntimeError(f"refusing to replace {path}")
    raw_perf = raw_root / "perf.data"
    env = os.environ.copy()
    env.update(
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

    def workload(report: Path) -> list[str]:
        return [
            str(binary),
            "provider-lifecycle",
            "--corpus",
            lane["corpus"],
            "--provider",
            lane["provider"],
            "--samples",
            str(protocol["profile"]["samples"]),
            "--warmup",
            str(protocol["profile"]["warmups"]),
            "--source-revision",
            freeze["revision"],
            "--output",
            str(report.resolve()),
        ]

    stat_command = [
        "taskset",
        "-c",
        protocol["cpu_set"],
        "/usr/bin/time",
        "-v",
        "-o",
        str(outputs["resource_stat"].resolve()),
        "perf",
        "stat",
        "--no-big-num",
        "-x,",
        "-e",
        ",".join(protocol["profile"]["events"]),
        "-o",
        str(outputs["stat"].resolve()),
        "--",
        *workload(outputs["report_json"]),
    ]
    record_command = [
        "taskset",
        "-c",
        protocol["cpu_set"],
        "/usr/bin/time",
        "-v",
        "-o",
        str(outputs["resource_record"].resolve()),
        "perf",
        "record",
        "--no-buildid-cache",
        "-o",
        str(raw_perf),
        "-F",
        str(protocol["profile"]["frequency_hz"]),
        "-e",
        "cycles:u",
        "--call-graph",
        protocol["profile"]["callgraph"],
        "--",
        *workload(outputs["record_report"]),
    ]
    receipt: dict[str, Any] = {
        "change": 501,
        "phase": phase,
        "lane": lane,
        "profile": protocol["profile"],
        "binary": freeze["binary"],
        "revision": freeze["revision"],
        "protocol_sha256": freeze["protocol_sha256"],
        "source_manifest_sha256": freeze["source_manifest"]["sha256"],
        "tmpdir": str(tmpdir.resolve()),
        "raw_perf_path": str(raw_perf),
        "stat_command": stat_command,
        "record_command": record_command,
        "started_utc": now(),
        "status": "running",
    }
    receipt_path.write_bytes(canonical_json(receipt))
    try:
        for command, stdout_path in ((stat_command, outputs["stat_stdout"]), (record_command, outputs["record_stdout"])):
            started = time.monotonic()
            with stdout_path.open("xb") as stream:
                result = subprocess.run(command, cwd=REPO, env=env, stdout=stream, stderr=subprocess.STDOUT)
            if result.returncode != 0:
                raise RuntimeError(f"profile command exited {result.returncode}: {name(lane)}")
            if time.monotonic() - started > 600:
                raise RuntimeError(f"profile command exceeded time bound: {name(lane)}")
        if raw_perf.stat().st_size > MAX_RAW_PERF:
            raise RuntimeError(f"perf.data exceeds {MAX_RAW_PERF} bytes")
        with outputs["report"].open("xb") as stream:
            result = subprocess.run(
                ["perf", "report", "--stdio", "--no-children", "--percent-limit", "0.1", "-i", str(raw_perf)],
                cwd=REPO,
                env=env,
                stdout=stream,
                stderr=subprocess.STDOUT,
            )
        if result.returncode != 0:
            raise RuntimeError(f"perf report exited {result.returncode}: {name(lane)}")
        report = verifier.load(outputs["report_json"])
        record_report = verifier.load(outputs["record_report"])
        verifier.check_report(report)
        verifier.check_report(record_report)
        report_text = outputs["report"].read_text(errors="replace")
        hits = sorted({token for token in SHA_TOKENS if re.search(re.escape(token), report_text)})
        receipt["sha_hotness_tokens_found"] = hits
        receipt["sha_hotness_detected"] = bool(hits)
        if include_sha_requirement and not hits:
            raise RuntimeError("perf report did not resolve any configured SHA hotness token")
        receipt["report_identities"] = {
            "stat": {
                "output_sha256": report["expected_output_sha256"],
                "output_bytes": report["expected_output_bytes"],
                "source_archive_sha256": report["source_archive_sha256"],
                "destination_archive_sha256": report["destination_archive_sha256"],
            },
            "record": {
                "output_sha256": record_report["expected_output_sha256"],
                "output_bytes": record_report["expected_output_bytes"],
                "source_archive_sha256": record_report["source_archive_sha256"],
                "destination_archive_sha256": record_report["destination_archive_sha256"],
            },
        }
        receipt["status"] = "pass"
    except Exception as error:
        receipt["status"] = "failed"
        receipt["error"] = str(error)
        raise
    finally:
        receipt["finished_utc"] = now()
        receipt["artifacts"] = {
            key: artifact(path)
            for key, path in outputs.items()
            if path.is_file()
        }
        for row in receipt["artifacts"].values():
            if row["bytes"] > MAX_COMPACT:
                receipt["status"] = "failed"
                receipt["error"] = f"oversized compact profile artifact: {row['path']}"
        receipt_path.write_bytes(canonical_json(receipt))
        if receipt["status"] == "pass" and raw_root.exists():
            # perf.data remains under the owned tmp namespace until the root
            # coordinator performs the final replay/cleanup audit.
            receipt["raw_perf_bytes"] = raw_perf.stat().st_size if raw_perf.exists() else 0
            receipt["raw_perf_sha256"] = sha_file(raw_perf) if raw_perf.exists() else None
            receipt["raw_retained_under_owned_tmp"] = raw_perf.is_file() and raw_perf.is_relative_to(TMP_ROOT)
            receipt_path.write_bytes(canonical_json(receipt))
    return receipt


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("phase", choices=("before", "after"))
    parser.add_argument("--allow-unresolved-sha", action="store_true")
    args = parser.parse_args()
    protocol = json.loads((HERE / "protocol.json").read_text())
    rows = []
    for lane in profile_lanes(protocol):
        rows.append(run_one(args.phase, lane, not args.allow_unresolved_sha))
    summary = {
        "change": 501,
        "phase": args.phase,
        "profiles": rows,
        "scope": protocol["profile"]["scope"],
        "sha_hotness_requirement": not args.allow_unresolved_sha,
    }
    path = HERE / "profiles" / args.phase / "summary.json"
    if path.exists():
        raise RuntimeError(f"refusing to replace {path}")
    path.write_bytes(canonical_json(summary))
    print(json.dumps({"status": "pass", "phase": args.phase, "profiles": len(rows)}, sort_keys=True))


if __name__ == "__main__":
    main()
