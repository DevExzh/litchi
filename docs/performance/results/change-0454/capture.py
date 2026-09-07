#!/usr/bin/env python3
"""Capture one frozen change-0454 lifecycle lane in a fresh child process.

The driver owns only report custody and process-level resource evidence.  The
Rust programs own the semantic/output oracles; this script invokes the
corresponding verifier before accepting a receipt.  It deliberately has no
profiling or allocator mode: allocation calls/bytes/live peaks are unavailable
from the preserved ordinary binaries and remain an explicit limitation.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def sha_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def sha_path(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def load(path: Path):
    return json.loads(path.read_text(encoding="utf-8"))


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def artifact(path: Path) -> dict[str, object]:
    return {
        "path": str(path.relative_to(ROOT)),
        "bytes": path.stat().st_size,
        "sha256": sha_path(path),
    }


def custody_module():
    spec = importlib.util.spec_from_file_location("change0454_check", ROOT / "check.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load the source-custody driver")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def binary_record(build: dict, kind: str, *, provider_suite: bool) -> tuple[dict, dict]:
    if provider_suite:
        binaries = build.get("binaries")
        record = (binaries.get("normal") if binaries else build)
    else:
        record = build["binaries"]["external"]
    path = Path(record["path"])
    if not path.is_file() or path.is_symlink():
        raise RuntimeError(f"missing or symlinked benchmark binary: {path}")
    if sha_path(path) != record["sha256"] or path.stat().st_size != record["bytes"]:
        raise RuntimeError(f"benchmark binary identity mismatch: {path}")
    return record, {"path": str(path), "kind": kind}


def provider_command(binary: Path, lane: dict, report: Path, protocol: dict, build: dict) -> list[str]:
    command = [
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
        build["revision"],
        "--output",
        str(report),
    ]
    if lane["provider"] == "range":
        range_config = protocol["range"]
        command.extend(
            [
                "--max-range",
                str(range_config["max_range_bytes"]),
                "--delay-us",
                str(range_config["delay_us"]),
                "--transfer-bytes-per-second",
                str(range_config["transfer_bytes_per_second"]),
                "--transfer-delay-policy",
                range_config["transfer_delay_policy"],
            ]
        )
    return command


def external_command(binary: Path, lane: dict, report: Path, protocol: dict, build: dict) -> list[str]:
    fixture = load(ROOT / "external-fixture.json")
    fixture_path = Path(fixture["path"])
    if not fixture_path.is_file() or fixture_path.is_symlink():
        raise RuntimeError(f"missing external fixture: {fixture_path}")
    if sha_path(fixture_path) != fixture["sha256"] or fixture_path.stat().st_size != fixture["bytes"]:
        raise RuntimeError("external fixture identity mismatch")
    return [
        str(binary),
        str(fixture_path),
        str(report),
        str(protocol["samples"]),
        str(protocol["warmups"]),
        build["revision"],
        lane["provider"],
    ]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--suite", choices=("provider", "external"), required=True)
    parser.add_argument("--lane", type=int, required=True)
    parser.add_argument("--pilot", action="store_true", help="one sample and no warmup for a smoke lane")
    args = parser.parse_args()

    protocol = load(ROOT / "protocol.json")
    lanes = protocol["provider_lanes"] if args.suite == "provider" else protocol["external_lanes"]
    if not 0 <= args.lane < len(lanes):
        raise SystemExit("lane index is outside the frozen protocol")
    lane = lanes[args.lane]
    if args.pilot:
        if not protocol["allow_pilot"]:
            raise SystemExit("pilot captures are disabled by the frozen protocol")
        samples, warmups = 1, 0
    else:
        samples, warmups = protocol["samples"], protocol["warmups"]

    build_path = ROOT / lane["build_manifest"]
    build = load(build_path)
    binary_record_value, _ = binary_record(build, lane["suite"], provider_suite=args.suite == "provider")
    binary = Path(binary_record_value["path"])

    custody = custody_module()
    before = custody.sources()
    # Both the preserved baseline executable and the current candidate are
    # run from one frozen checkout epoch.  The baseline's own historical
    # manifest identifies the executable's build inputs; the capture checkout
    # is bound to the current candidate manifest, as in change-0453.
    capture_source = load(ROOT / "candidate-build.json")["source_manifest"]
    expected_source = capture_source
    if expected_source is not None and before != expected_source:
        raise RuntimeError("capture source epoch differs from the bound build manifest")

    if args.pilot:
        directory_name = "provider-pilots" if args.suite == "provider" else "external-pilots"
    else:
        directory_name = "provider-runs" if args.suite == "provider" else "external-runs"
    directory = ROOT / directory_name / str(args.lane)
    if directory.exists():
        raise RuntimeError(f"capture directory already exists: {directory}")
    directory.mkdir(parents=True)
    report = directory / "report.json"
    resource = directory / "resource.log"
    workload_log = directory / "workload.log"
    oracle_log = directory / "oracle.log"
    if args.suite == "provider":
        workload = provider_command(binary, lane, report, dict(protocol, samples=samples, warmups=warmups), build)
    else:
        workload = external_command(binary, lane, report, dict(protocol, samples=samples, warmups=warmups), build)
    cpu = str(protocol["cpu"])
    argv = ["taskset", "-c", cpu, "/usr/bin/time", "-v", "-o", str(resource), *workload]
    record: dict[str, object] = {
        "change": 454,
        "schema": "pptx_change0454_capture_v1",
        "suite": args.suite,
        "lane": args.lane,
        "lane_definition": lane,
        "pilot": args.pilot,
        "status": "running",
        "revision": build["revision"],
        "argv": argv,
        "binary": binary_record_value,
        "build_manifest": {
            "path": str(build_path.relative_to(ROOT)),
            "sha256": sha_path(build_path),
        },
        "source_before": before,
        "capture_source_manifest": capture_source,
        "protocol_sha256": sha_path(ROOT / "protocol.json"),
        "driver_sha256": sha_path(Path(__file__)),
        "capture_sha256": sha_path(Path(__file__)),
        "provider_oracle_sha256": sha_path(ROOT / "verify-report.py"),
        "external_oracle_sha256": sha_path(ROOT / "external-verifier.py"),
        "started_utc": now(),
    }
    if args.suite == "external":
        fixture = load(ROOT / "external-fixture.json")
        record["fixture"] = {
            "path": fixture["path"],
            "bytes": fixture["bytes"],
            "sha256": fixture["sha256"],
        }
    receipt = directory / "receipt.json"
    receipt.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    try:
        with workload_log.open("xb") as output:
            result = subprocess.run(argv, cwd=REPO, stdout=output, stderr=subprocess.STDOUT)
        record["exit_code"] = result.returncode
        if result.returncode != 0:
            raise RuntimeError(f"benchmark child exited with {result.returncode}")
        oracle = ROOT / ("verify-report.py" if args.suite == "provider" else "external-verifier.py")
        with oracle_log.open("xb") as output:
            result = subprocess.run(
                [sys.executable, "-B", str(oracle), str(report)],
                cwd=REPO,
                stdout=output,
                stderr=subprocess.STDOUT,
            )
        record["oracle_exit_code"] = result.returncode
        if result.returncode != 0:
            raise RuntimeError(f"report oracle exited with {result.returncode}")
        record["status"] = "pass"
    finally:
        record["source_after"] = custody.sources()
        record["source_unchanged"] = record["source_after"] == before
        if not record["source_unchanged"] or record["status"] == "running":
            record["status"] = "failed"
        record["finished_utc"] = now()
        artifact_paths = [
            ("report", report),
            ("resource", resource),
            ("workload", workload_log),
            ("oracle", oracle_log),
        ]
        if args.suite == "external":
            artifact_paths.append(("output_artifact", report.with_suffix(".pptx")))
        record["artifacts"] = {
            name: artifact(path) for name, path in artifact_paths if path.exists()
        }
        receipt.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    if record["status"] != "pass":
        raise RuntimeError(f"capture failed; inspect {receipt}")
    print(json.dumps({"status": "pass", "suite": args.suite, "lane": args.lane}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
