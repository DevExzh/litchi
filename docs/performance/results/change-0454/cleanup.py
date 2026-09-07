#!/usr/bin/env python3
"""Inventory and remove this batch's explicitly owned artifacts."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parent
TASK = Path("/tmp/litchi-goal-0454-native-copy")
ALLOWED = {
    "baseline-probe",
    "name-only-probe",
    "baseline-lifecycle",
    "external-smoketest.pptx",
    "integrated-probe",
    "preflight-external",
    "candidate-normal",
    "candidate-external",
    "candidate-probe",
    "oracle-r1-candidate-normal",
    "oracle-r1-candidate-external",
    "oracle-r1-candidate-probe",
}
BINARY_MANIFEST = ROOT / "checks" / "binary-artifacts.json"
BINARY_RECEIPT = ROOT / "checks" / "binary-cleanup.json"
RAW_MANIFEST = ROOT / "checks" / "raw-output-artifacts.json"
RAW_RECEIPT = ROOT / "checks" / "raw-output-cleanup.json"
RAW_SCOPE = "external-runs and external-pilots report.pptx siblings; archived failed external-pilot-attempts report.pptx siblings"


def sha(path: Path) -> str:
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def binary_inventory() -> list[dict[str, object]]:
    assert TASK.is_dir() and not TASK.is_symlink()
    rows: list[dict[str, object]] = []
    for path in sorted(TASK.iterdir()):
        assert path.name in ALLOWED and path.is_file() and not path.is_symlink(), path
        rows.append({"path": path.name, "bytes": path.stat().st_size, "sha256": sha(path)})
    return rows


def raw_output_inventory() -> list[dict[str, object]]:
    protocol = json.loads((ROOT / "protocol.json").read_text(encoding="utf-8"))
    lanes = protocol["external_lanes"]
    expected: set[Path] = set()
    rows: list[dict[str, object]] = []
    roots = (("external-runs", False), ("external-pilots", True))
    for root_name, pilot in roots:
        external_root = ROOT / root_name
        assert external_root.is_dir() and not external_root.is_symlink()
        children = list(external_root.iterdir())
        directories = [path for path in children if path.is_dir()]
        assert len(directories) == len(children) and all(not path.is_symlink() for path in directories)
        actual_directories = {path.name for path in directories}
        assert actual_directories == {str(index) for index in range(len(lanes))}
        for index, lane in enumerate(lanes):
            directory = external_root / str(index)
            report = directory / "report.json"
            receipt = directory / "receipt.json"
            output = directory / "report.pptx"
            assert report.is_file() and not report.is_symlink()
            assert receipt.is_file() and not receipt.is_symlink()
            assert output.is_file() and not output.is_symlink()
            expected.add(output)
            report_value = json.loads(report.read_text(encoding="utf-8"))
            assert report_value["schema"] == "pptx-external-cross-copy-v1"
            output_record = report_value["output_artifact"]
            assert Path(output_record["path"]).resolve() == output.resolve()
            assert output_record["bytes"] == output.stat().st_size
            assert output_record["sha256"] == sha(output)
            receipt_value = json.loads(receipt.read_text(encoding="utf-8"))
            assert receipt_value["status"] == "pass"
            assert receipt_value["suite"] == "external"
            assert receipt_value["lane"] == index
            assert receipt_value["pilot"] is pilot
            assert receipt_value["lane_definition"] == lane
            artifacts = receipt_value["artifacts"]
            report_artifact = artifacts["report"]
            output_artifact = artifacts["output_artifact"]
            assert (ROOT / report_artifact["path"]).resolve() == report.resolve()
            assert report_artifact["bytes"] == report.stat().st_size
            assert report_artifact["sha256"] == sha(report)
            assert (ROOT / output_artifact["path"]).resolve() == output.resolve()
            assert output_artifact["bytes"] == output.stat().st_size
            assert output_artifact["sha256"] == sha(output)
            rows.append(
                {
                    "path": str(output.relative_to(ROOT)),
                    "bytes": output.stat().st_size,
                    "sha256": sha(output),
                    "report_path": str(report.relative_to(ROOT)),
                    "report_bytes": report.stat().st_size,
                    "report_sha256": sha(report),
                    "receipt_path": str(receipt.relative_to(ROOT)),
                    "receipt_bytes": receipt.stat().st_size,
                    "receipt_sha256": sha(receipt),
                    "lane": index,
                    "pilot": pilot,
                }
            )
    attempts_root = ROOT / "external-pilot-attempts"
    if attempts_root.exists():
        assert attempts_root.is_dir() and not attempts_root.is_symlink()
        attempts = [path for path in attempts_root.iterdir()]
        assert all(path.is_dir() and not path.is_symlink() for path in attempts)
        for attempt in sorted(attempts):
            assert re.fullmatch(r"oracle-r[0-9]+", attempt.name), attempt
            lane_children = list(attempt.iterdir())
            lane_directories = [path for path in lane_children if path.is_dir()]
            assert len(lane_directories) == len(lane_children) and all(not path.is_symlink() for path in lane_directories)
            assert {path.name for path in lane_directories} <= {str(index) for index in range(len(lanes))}
            for lane_directory in sorted(lane_directories, key=lambda path: int(path.name)):
                index = int(lane_directory.name)
                receipt = lane_directory / "receipt.json"
                report = lane_directory / "report.json"
                output = lane_directory / "report.pptx"
                assert receipt.is_file() and not receipt.is_symlink()
                receipt_value = json.loads(receipt.read_text(encoding="utf-8"))
                assert receipt_value["status"] == "failed"
                assert receipt_value["suite"] == "external"
                assert receipt_value["pilot"] is True and receipt_value["lane"] == index
                assert receipt_value.get("exit_code", 0) != 0 or receipt_value.get("oracle_exit_code", 0) != 0
                assert receipt_value["lane_definition"] == lanes[index]
                if not report.exists() and not output.exists():
                    continue
                assert report.is_file() and not report.is_symlink()
                assert output.is_file() and not output.is_symlink()
                expected.add(output)
                report_value = json.loads(report.read_text(encoding="utf-8"))
                assert report_value["schema"] == "pptx-external-cross-copy-v1"
                output_record = report_value["output_artifact"]
                recorded_output = Path(output_record["path"])
                assert recorded_output.name == output.name and recorded_output.parent.name == lane_directory.name
                assert output_record["bytes"] == output.stat().st_size
                assert output_record["sha256"] == sha(output)
                artifacts = receipt_value["artifacts"]
                report_artifact = artifacts["report"]
                output_artifact = artifacts["output_artifact"]
                recorded_report = Path(report_artifact["path"])
                recorded_receipt_output = Path(output_artifact["path"])
                assert recorded_report.name == report.name and recorded_report.parent.name == lane_directory.name
                assert recorded_receipt_output.name == output.name and recorded_receipt_output.parent.name == lane_directory.name
                assert report_artifact["bytes"] == report.stat().st_size
                assert report_artifact["sha256"] == sha(report)
                assert output_artifact["bytes"] == output.stat().st_size
                assert output_artifact["sha256"] == sha(output)
                rows.append(
                    {
                        "path": str(output.relative_to(ROOT)),
                        "bytes": output.stat().st_size,
                        "sha256": sha(output),
                        "report_path": str(report.relative_to(ROOT)),
                        "report_bytes": report.stat().st_size,
                        "report_sha256": sha(report),
                        "receipt_path": str(receipt.relative_to(ROOT)),
                        "receipt_bytes": receipt.stat().st_size,
                        "receipt_sha256": sha(receipt),
                        "lane": index,
                        "pilot": True,
                        "historical": True,
                        "attempt": attempt.name,
                        "receipt_status": "failed",
                    }
                )
    actual = {
        path
        for root_name, _pilot in roots
        for path in (ROOT / root_name).rglob("*.pptx")
        if path.is_file() or path.is_symlink()
    }
    if attempts_root.exists():
        actual.update(
            path
            for path in attempts_root.rglob("*.pptx")
            if path.is_file() or path.is_symlink()
        )
    assert actual == expected, (sorted(actual), sorted(expected))
    return rows


def write_json(path: Path, value: dict[str, object]) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


parser = argparse.ArgumentParser()
parser.add_argument("action", choices=("inventory", "cleanup"))
args = parser.parse_args()

if args.action == "inventory":
    assert not BINARY_MANIFEST.exists()
    assert not RAW_MANIFEST.exists()
    binaries = binary_inventory()
    raw_outputs = raw_output_inventory()
    write_json(
        BINARY_MANIFEST,
        {
            "status": "pass",
            "task": str(TASK),
            "artifacts": binaries,
            "files": len(binaries),
            "bytes": sum(int(row["bytes"]) for row in binaries),
            "driver_sha256": sha(Path(__file__)),
        },
    )
    write_json(
        RAW_MANIFEST,
        {
            "status": "pass",
            "scope": RAW_SCOPE,
            "artifacts": raw_outputs,
            "files": len(raw_outputs),
            "bytes": sum(int(row["bytes"]) for row in raw_outputs),
            "driver_sha256": sha(Path(__file__)),
        },
    )
else:
    precleanup = json.loads((ROOT / "checks" / "precleanup.json").read_text(encoding="utf-8"))
    assert precleanup["status"] == "pass"
    frozen = json.loads(BINARY_MANIFEST.read_text(encoding="utf-8"))
    frozen_raw = json.loads(RAW_MANIFEST.read_text(encoding="utf-8"))
    driver_sha256 = sha(Path(__file__))
    assert frozen["task"] == str(TASK) and frozen["driver_sha256"] == driver_sha256
    assert frozen_raw["status"] == "pass"
    assert frozen_raw["scope"] == RAW_SCOPE
    assert frozen_raw["driver_sha256"] == driver_sha256
    binaries = binary_inventory()
    raw_outputs = raw_output_inventory()
    assert binaries == frozen["artifacts"]
    assert raw_outputs == frozen_raw["artifacts"]

    for row in binaries:
        (TASK / str(row["path"])).unlink()
    TASK.rmdir()
    for row in raw_outputs:
        (ROOT / str(row["path"])).unlink()

    assert not TASK.exists()
    assert all(not (ROOT / str(row["path"])).exists() for row in raw_outputs)
    write_json(
        BINARY_RECEIPT,
        {
            "status": "pass",
            "task": str(TASK),
            "temporary_directory_absent": True,
            "files_removed": len(binaries),
            "bytes_removed": sum(int(row["bytes"]) for row in binaries),
            "artifact_manifest_sha256": sha(BINARY_MANIFEST),
            "driver_sha256": driver_sha256,
        },
    )
    write_json(
        RAW_RECEIPT,
        {
            "status": "pass",
            "scope": RAW_SCOPE,
            "raw_outputs_absent": True,
            "files_removed": len(raw_outputs),
            "bytes_removed": sum(int(row["bytes"]) for row in raw_outputs),
            "artifact_manifest_sha256": sha(RAW_MANIFEST),
            "driver_sha256": driver_sha256,
        },
    )

print(json.dumps({"status": "pass", "action": args.action}))
