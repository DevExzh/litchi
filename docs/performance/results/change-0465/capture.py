#!/usr/bin/env python3
"""Capture the frozen descriptive default/ODP matrix from retained binaries.

The runner intentionally executes one lane per invocation.  Every lane gets a
new directory and an exclusive receipt, so a partially completed or reordered
lane cannot silently replace evidence from an earlier run.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
PROTOCOL_PATH = ROOT / "protocol.json"
CHECK_PATH = ROOT / "check.py"
BINDING_PATH = ROOT / "binding.json"


class CaptureError(RuntimeError):
    """A failed custody, workload, or report gate."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CaptureError(message)


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    with path.open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def read_json(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing or symlinked JSON: {path}")
    value = json.loads(path.read_text())
    require(isinstance(value, dict), f"JSON object required: {path}")
    return value


def write_exclusive(path: Path, value: dict[str, Any]) -> None:
    with path.open("x") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def regular(path: Path, label: str) -> None:
    require(path.is_file() and not path.is_symlink(), f"{label} is missing or symlinked: {path}")


def resolve_retained(value: str, *, base: Path = ROOT) -> Path:
    path = Path(value)
    return path if path.is_absolute() else base / path


def relative_to_repo(path: Path) -> str:
    try:
        return str(path.relative_to(REPO))
    except ValueError:
        return str(path)


def artifact(path: Path) -> dict[str, Any]:
    regular(path, "artifact")
    return {
        "path": str(path.relative_to(ROOT)),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }


def load_custody() -> Any:
    spec = importlib.util.spec_from_file_location("change0465_check", CHECK_PATH)
    require(spec is not None and spec.loader is not None, "cannot import check.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    require(hasattr(module, "sources"), "check.py does not expose sources()")
    return module


def source_file(record: dict[str, Any], label: str) -> Path:
    require(set(("path", "sha256", "files")) <= set(record), f"{label} source record is incomplete")
    path = resolve_retained(str(record["path"]))
    regular(path, f"{label} source manifest")
    require(sha(path) == record["sha256"], f"{label} source manifest hash differs")
    require(isinstance(record["files"], int) and record["files"] > 0, f"{label} source file count is invalid")
    return path


def validate_build_receipt(entry: dict[str, Any], source: dict[str, Any], label: str) -> None:
    receipt_value = entry["build_receipt"]
    require(isinstance(receipt_value, str) and receipt_value, f"{label} build receipt path is empty")
    receipt_path = resolve_retained(receipt_value)
    regular(receipt_path, f"{label} build receipt")
    require(sha(receipt_path) == entry["build_receipt_sha256"], f"{label} build receipt hash differs")
    receipt = read_json(receipt_path)
    require(receipt.get("status") == "pass", f"{label} build receipt is not pass")
    require(receipt.get("source_unchanged") is True, f"{label} build changed its source custody")
    if "source_before" in receipt:
        require(receipt["source_before"] == source, f"{label} build source_before differs from binding")
    if "source_after" in receipt:
        require(receipt["source_after"] == source, f"{label} build source_after differs from binding")


def validate_binding(protocol: dict[str, Any], lane: dict[str, Any]) -> tuple[dict[str, Any], str]:
    regular(BINDING_PATH, "binary binding")
    binding = read_json(BINDING_PATH)
    contract = protocol["binding"]
    require(binding.get("schema") == contract["schema"], "binary binding schema differs")
    revision = binding.get("revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision) is not None,
            "binary binding revision is not a commit SHA")
    binaries = binding.get("binaries")
    require(isinstance(binaries, dict), "binary binding binaries map is missing")
    required_names = {lane["binary"]}
    if lane["binary"] == "allocator":
        required_names.add("normal")
    for name in sorted(required_names):
        entry = binaries.get(name)
        require(isinstance(entry, dict), f"binding entry is missing: {name}")
        required = {"path", "sha256", "bytes", "build_receipt", "build_receipt_sha256", "source"}
        require(required <= set(entry), f"binding entry is incomplete: {name}")
        binary_path = resolve_retained(str(entry["path"]), base=ROOT)
        regular(binary_path, f"{name} binary")
        require(re.fullmatch(r"[0-9a-f]{64}", str(entry["sha256"])) is not None,
                f"{name} binary SHA-256 is invalid")
        require(binary_path.stat().st_size == entry["bytes"], f"{name} binary byte count differs")
        require(sha(binary_path) == entry["sha256"], f"{name} binary hash differs from binding")
        source = entry["source"]
        require(isinstance(source, dict), f"{name} source record is missing")
        source_file(source, f"{name}")
        validate_build_receipt(entry, source, name)
    binding_sha = sha(BINDING_PATH)
    return binding, binding_sha


def validate_checked_manifest(protocol: dict[str, Any], lane_name: str) -> Path | None:
    spec = protocol["default_manifest"]
    path = ROOT / spec["path"]
    if lane_name == "preflight" and not path.exists() and not path.is_symlink():
        # The first lane is allowed to establish this checked record.  Its
        # report is still validated against the frozen protocol list below.
        return None
    regular(path, "checked default manifest")
    manifest = read_json(path)
    require(manifest.get("result_count") == spec["result_count"], "checked default manifest result count differs")
    require(manifest.get("case_count") == spec["case_count"], "checked default manifest case count differs")
    require(manifest.get("default_cases") == protocol["default_cases"],
            "checked default manifest case order differs")
    return path


def validate_catalog(
    catalog_path: Path,
    report: dict[str, Any],
    expected_rows: int,
    expected_cases: list[str],
) -> dict[str, Any]:
    catalog = read_json(catalog_path)
    require(catalog.get("manifest_version") == 2, "corpus catalog manifest version differs")
    require(catalog.get("manifest_kind") == "corpus-catalog", "corpus catalog kind differs")
    for key in ("catalog_sha256", "content_set_sha256"):
        require(re.fullmatch(r"[0-9a-f]{64}", str(catalog.get(key))) is not None,
                f"corpus catalog {key} is invalid")
    bindings = catalog.get("case_bindings")
    require(isinstance(bindings, list) and len(bindings) == expected_rows,
            "corpus catalog case binding count differs")
    require(all(isinstance(row, dict) and row.get("case") in expected_cases for row in bindings),
            "corpus catalog contains an unexpected case")
    reference = report.get("corpus_catalog")
    require(isinstance(reference, dict), "report omitted corpus catalog reference")
    for key in ("manifest_version", "catalog_id", "catalog_sha256", "content_set_sha256"):
        require(reference.get(key) == catalog.get(key), f"report/catalog {key} differs")
    return catalog


def validate_report(
    report_path: Path,
    catalog_path: Path,
    lane: dict[str, Any],
    protocol: dict[str, Any],
    binary_entry: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    report = read_json(report_path)
    expected_rows = lane["expected_rows"]
    is_default = lane["case_mode"] == "default"
    expected_cases = protocol["default_cases"] if is_default else [lane["case"]]
    require(report.get("schema_version") == 1, "report schema version differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), "report configuration is missing")
    require(configuration.get("samples_per_case") == lane["samples"], "report sample count differs")
    require(configuration.get("warmup_iterations_per_case") == lane["warmup"], "report warmup count differs")
    require(configuration.get("execution_workers") == [protocol["workers"]], "report worker selection differs")
    require(configuration.get("cases") == expected_cases, "report case selection differs")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == expected_rows,
            f"report result count differs (expected {expected_rows})")
    require(all(isinstance(row, dict) and row.get("case") in expected_cases for row in results),
            "report contains an unexpected case")
    for row in results:
        elapsed = row.get("elapsed_ns")
        require(isinstance(elapsed, dict) and isinstance(elapsed.get("samples"), list),
                "report result has no elapsed samples")
        require(len(elapsed["samples"]) == lane["samples"], "report result sample length differs")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), "report binary identity is missing")
    require(identity.get("binary_sha256") == binary_entry["sha256"], "report binary SHA differs")
    require(identity.get("binary_bytes") == binary_entry["bytes"], "report binary byte count differs")
    catalog = validate_catalog(catalog_path, report, expected_rows, expected_cases)
    return report, catalog


def lane_definition(protocol: dict[str, Any], name: str) -> dict[str, Any]:
    lanes = protocol.get("lanes")
    require(isinstance(lanes, dict) and name in lanes, f"lane is not in frozen protocol: {name}")
    lane = lanes[name]
    require(isinstance(lane, dict), f"lane definition is not an object: {name}")
    return lane


def ensure_order(protocol: dict[str, Any], lane_name: str) -> None:
    if lane_name == "preflight":
        return
    order = protocol["formal_order"]
    position = order.index(lane_name)
    for prior in order[:position]:
        receipt = ROOT / "captures" / prior / "receipt.json"
        regular(receipt, f"preceding lane receipt {prior}")
        record = read_json(receipt)
        require(record.get("status") == "pass", f"preceding lane did not pass: {prior}")


def command_for(protocol: dict[str, Any], lane: dict[str, Any], binary: Path, report: Path, catalog: Path) -> list[str]:
    argv = [
        str(binary),
        "--workers", str(protocol["workers"]),
        "--samples", str(lane["samples"]),
        "--warmup", str(lane["warmup"]),
        "--json", relative_to_repo(report),
        "--corpus-manifest", relative_to_repo(catalog),
    ]
    if lane["case_mode"] == "selection":
        argv.extend(["--case", lane["case"]])
    return argv


def run_lane(lane_name: str) -> int:
    protocol = read_json(PROTOCOL_PATH)
    require(protocol.get("schema") == "litchi-0465-capture-v1", "protocol schema differs")
    require(protocol.get("change") == 465 and protocol.get("status") == "frozen",
            "protocol identity differs")
    require(sha(Path(__file__)) == protocol["runner"]["sha256"], "capture driver differs from protocol")
    require(sha(CHECK_PATH) == protocol["helper"]["sha256"], "check.py differs from protocol")
    require(protocol.get("cpu") == 2 and protocol.get("workers") == 1, "CPU/worker protocol differs")
    lane = lane_definition(protocol, lane_name)
    ensure_order(protocol, lane_name)
    checked_manifest_path = validate_checked_manifest(protocol, lane_name)

    directory = ROOT / "captures" / lane_name
    require(not directory.exists(), f"lane directory already exists: {directory}")
    directory.mkdir(parents=True)
    report_path = directory / "report.json"
    catalog_path = directory / "corpus-catalog.json"
    resource_path = directory / "resource.log"
    workload_path = directory / "workload.log"
    receipt_path = directory / "receipt.json"
    for path in (report_path, catalog_path, resource_path, workload_path, receipt_path):
        require(not path.exists(), f"lane artifact already exists: {path}")

    custody = load_custody()
    binding, binding_sha = validate_binding(protocol, lane)
    source_before = custody.sources()
    required_entry = binding["binaries"][lane["binary"]]
    require(required_entry["source"] == source_before,
            f"{lane_name} binary source does not match current source custody")
    binary_path = resolve_retained(required_entry["path"])
    command = command_for(protocol, lane, binary_path, report_path, catalog_path)
    timed_command = [
        "/usr/bin/time", "-v", "-o", relative_to_repo(resource_path),
        "taskset", "-c", str(protocol["cpu"]), *command,
    ]
    environment = os.environ | protocol["environment"]
    record: dict[str, Any] = {
        "schema": "litchi-0465-capture-v1",
        "change": 465,
        "lane": lane_name,
        "lane_definition": lane,
        "argv": timed_command,
        "cwd": str(REPO),
        "binding_revision": binding["revision"],
        "binding_sha256": binding_sha,
        "protocol_sha256": sha(PROTOCOL_PATH),
        "runner_sha256": sha(Path(__file__)),
        "helper_sha256": sha(CHECK_PATH),
        "binary": {
            "name": lane["binary"],
            "path": str(binary_path),
            "sha256": required_entry["sha256"],
            "bytes": required_entry["bytes"],
        },
        "environment": protocol["environment"],
        "started_utc": now(),
        "source_before": source_before,
        "status": "running",
    }
    if checked_manifest_path is not None:
        record["checked_manifest"] = artifact(checked_manifest_path)
    error: str | None = None
    report: dict[str, Any] | None = None
    catalog: dict[str, Any] | None = None
    try:
        observed_hash = sha(binary_path)
        require(observed_hash == required_entry["sha256"], "binary changed before workload")
        with workload_path.open("xb") as output:
            result = subprocess.run(
                timed_command,
                cwd=REPO,
                env=environment,
                stdout=output,
                stderr=subprocess.STDOUT,
                check=False,
            )
        record["exit_code"] = result.returncode
        require(result.returncode == 0, f"workload exited with status {result.returncode}")
        report, catalog = validate_report(report_path, catalog_path, lane, protocol, required_entry)
        require(sha(binary_path) == required_entry["sha256"], "binary changed after workload")
        record["status"] = "pass"
    except Exception as exc:
        error = str(exc)
        record["status"] = "failed"
    finally:
        try:
            source_after = custody.sources()
            record["source_after"] = source_after
            if source_after != source_before:
                record["status"] = "failed"
                error = "source custody changed during lane"
        except Exception as exc:
            record["source_after_error"] = str(exc)
            record["status"] = "failed"
            error = error or str(exc)
    record["finished_utc"] = now()
    if error:
        record["error"] = error
    record["artifact_hashes"] = {
        name: artifact(path)
        for name, path in (
            ("report", report_path),
            ("corpus_catalog", catalog_path),
            ("resource_log", resource_path),
            ("workload_log", workload_path),
        )
        if path.is_file() and not path.is_symlink()
    }
    record["report_validation"] = {
        "expected_rows": lane["expected_rows"],
        "report_rows": len(report["results"]) if report else None,
        "catalog_case_bindings": len(catalog["case_bindings"]) if catalog else None,
        "validated": record["status"] == "pass",
    }
    write_exclusive(receipt_path, record)
    print(json.dumps({"lane": lane_name, "status": record["status"], "receipt": str(receipt_path.relative_to(ROOT))}))
    return 0 if record["status"] == "pass" else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--lane", choices=("preflight", "R1-normal", "R1-allocator", "R2-allocator", "R2-normal"), required=True)
    args = parser.parse_args()
    try:
        return run_lane(args.lane)
    except Exception as exc:
        print(f"capture failed before immutable receipt: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
