#!/usr/bin/env python3
"""Audit and replay the retained 0918 functional qualification custody.

The qualification is deliberately separate from the timed comparison.  This
script reads the external receipt/raw tree, revalidates every report with the
unchanged 0917 validator, and compares the non-metric report controls between
the two production versions.  It never builds or executes a probe.

Qualification receipts intentionally retain absolute paths for their external
raw files, input fixtures, and binaries.  Replay therefore requires the
original checkout inputs and the external qualification root recorded in the
receipts; it is not a relocation mechanism.  Comparative packet relocation is
handled separately by ``comparison-replay.py`` and requires a reconstructed
repository hierarchy.
"""

from __future__ import annotations

import argparse
import copy
import datetime as dt
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
REPO_ROOT = HERE.parents[3]
PROTOCOL_PATH = HERE / "comparison-protocol.json"
SUMMARY_PATH = HERE / "qualification-summary.json"
VALIDATOR_PATH = REPO_ROOT / "docs/performance/results/change-0917/capture.py"
DEFAULT_EXTERNAL = Path(
    os.environ.get("LITCHI_PERF_0918_QUALIFY_ROOT",
                   "/home/zhuhe/litchi-performance-0918/qualification")
).resolve()
MAX_PACKET_ARTIFACT = 1 << 20
HEX = set("0123456789abcdef")
VERSIONS = ("baseline", "candidate")
LANES = ("native", "allocator")
CASES = ("docnohf", "docfloat")
EDITS = (1, 2, 8, 32)
ROUTES = ("zero", "default", "release")
EXPECTED_BINARY_HASHES = {
    ("baseline", "native"): "2ae48b4f7548c1468c8e7691d56d89b231c4a63f52ed48383f07fe3609c59a13",
    ("baseline", "allocator"): "927606880090f9169e222254944df92366dd1cb0e1aa37fc0861234c466ccf1b",
    ("candidate", "native"): "76dbda76cd6bdbb3e9a219b4ef0c06160c4607a2bc58ad2e4bd21103008f6d2b",
    ("candidate", "allocator"): "b151763ad9613ac2ac6f8349b3ff80383b45f96d27298f1fbb3b3d1ab2fe5fb9",
}
EXPECTED_SOURCE_REVISION = "f5485b3ca8582051247d31946b1c5f45c8a561ad"
EXPECTED_PROTOCOL_SHA256 = "0bfd3669f72b0a491f40c489f9d2cd153275e99add8fd65fba3554e5e50b72de"
EXPECTED_VALIDATOR_SHA256 = "fde5219539cf538d51b8bf92ec5527fd9959ac0f39eb2ea5412b8ddf8cd7a894"


class QualificationError(RuntimeError):
    """Fail-closed qualification custody or semantic validation error."""


def fail(message: str) -> None:
    raise QualificationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def json_bytes(value: object) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def digest(value: object, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label}: invalid SHA-256")
    return value


def integer(value: object, label: str, minimum: int = 0) -> int:
    require(type(value) is int and value >= minimum, f"{label}: invalid integer")
    return value


def resolve_descriptor(identity: object, label: str) -> Path:
    require(isinstance(identity, dict), f"{label}: descriptor missing")
    raw = identity.get("path")
    require(isinstance(raw, str) and raw, f"{label}: descriptor path missing")
    path = Path(raw)
    if not path.is_absolute():
        path = (REPO_ROOT / path).resolve()
    else:
        path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"{label}: missing regular file: {path}")
    expected_bytes = integer(identity.get("bytes"), f"{label}: bytes")
    expected_hash = digest(identity.get("sha256"), f"{label}: sha256")
    actual_bytes = path.stat().st_size
    require(actual_bytes == expected_bytes,
            f"{label}: byte length changed ({actual_bytes} != {expected_bytes})")
    actual_hash = sha256(path)
    require(actual_hash == expected_hash,
            f"{label}: SHA-256 changed ({actual_hash} != {expected_hash})")
    return path


def descriptor(path: Path, *, relative_to_repo: bool = False) -> dict[str, object]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    result: dict[str, object] = {
        "path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path),
    }
    if relative_to_repo:
        try:
            result["path"] = path.relative_to(REPO_ROOT).as_posix()
        except ValueError:
            pass
    return result


def parse_timestamp(value: object, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label}: timestamp missing")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label}: malformed timestamp: {error}")
    require(parsed.tzinfo is not None, f"{label}: timestamp has no timezone")
    return parsed.astimezone(dt.timezone.utc)


def validator_module():
    require(VALIDATOR_PATH.is_file(),
            f"relocation inputs incomplete: copied 0917 validator missing: {VALIDATOR_PATH}")
    spec = importlib.util.spec_from_file_location("litchi_perf_0917_qualification_validator",
                                                  VALIDATOR_PATH)
    require(spec is not None and spec.loader is not None,
            f"cannot load unchanged 0917 validator: {VALIDATOR_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def protocol_and_fixtures(module: Any) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    protocol = read_json(PROTOCOL_PATH)
    require(isinstance(protocol, dict), "0918 comparison protocol is not an object")
    require(protocol.get("schema") == "litchi.performance.0918.doc-comparison.protocol.v1",
            "0918 protocol schema changed")
    require(protocol.get("current_source_head") == EXPECTED_SOURCE_REVISION,
            "0918 candidate source identity changed")
    require(sha256(PROTOCOL_PATH) == EXPECTED_PROTOCOL_SHA256,
            "0918 protocol bytes changed")
    validator_protocol = module.protocol()
    fixtures = module.fixture_rows(validator_protocol)
    require(set(fixtures) == set(CASES), "0917 fixture set changed")
    for case in CASES:
        row = fixtures[case]
        path = Path(str(row["path_abs"])).resolve()
        require(path.is_file(), f"relocation inputs incomplete: fixture missing: {path}")
        require(path.stat().st_size == row["bytes"] and sha256(path) == row["sha256"],
                f"{case}: fixture custody changed")
    return protocol, fixtures


def expected_rows() -> dict[str, dict[str, object]]:
    rows: dict[str, dict[str, object]] = {}
    for version in VERSIONS:
        for lane in LANES:
            # Qualification covers every retention route for both lanes.  The
            # timed native schedule later selects default retention only.
            routes = ROUTES
            for case in CASES:
                for edits in EDITS:
                    for route in routes:
                        directory = f"{version}-{lane}-{case}-e{edits:02d}-{route}"
                        receipt_label = f"qualification-{version}-{case}-e{edits}-{route}"
                        rows[directory] = {
                            "directory": directory, "version": version, "lane": lane,
                            "case": case, "edits": edits, "retention": route,
                            "receipt_label": receipt_label,
                        }
    require(len(rows) == 96, f"qualification schedule has {len(rows)} expected rows")
    return rows


def parse_report(stdout: bytes, label: str) -> dict[str, Any]:
    try:
        text = stdout.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"{label}: stdout is not UTF-8: {error}")
    candidates = [line.strip() for line in text.splitlines() if line.strip()]
    require(candidates, f"{label}: stdout is empty")
    for line in reversed(candidates):
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict) and value.get("schema") == "litchi.performance.0917.doc_retention.v1":
            return value
    fail(f"{label}: stdout contains no 0917 report JSON")


def parse_time_v(data: bytes, label: str) -> dict[str, object]:
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"{label}: process log is not UTF-8: {error}")
    patterns = {
        "user_seconds": r"^\s*User time \(seconds\):\s*([0-9.]+)$",
        "system_seconds": r"^\s*System time \(seconds\):\s*([0-9.]+)$",
        "max_rss_kib": r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)$",
        "elapsed": r"^\s*Elapsed \(wall clock\) time \(h:mm:ss or m:ss\):\s*(\S+)$",
        "exit_status": r"^\s*Exit status:\s*(\d+)$",
    }
    result: dict[str, object] = {}
    for key, pattern in patterns.items():
        match = re.search(pattern, text, re.MULTILINE)
        if match is not None:
            result[key] = (float(match.group(1)) if key.endswith("seconds") else
                           int(match.group(1)) if key in {"max_rss_kib", "exit_status"}
                           else match.group(1))
    require("max_rss_kib" in result and "exit_status" in result,
            f"{label}: process metrics incomplete")
    return result


def semantic_signature(report: dict[str, Any]) -> dict[str, Any]:
    value = copy.deepcopy(report)
    value.pop("elapsed_ns", None)
    value.pop("allocation", None)
    for sample in value.get("samples", []):
        if isinstance(sample, dict):
            sample.pop("elapsed_ns", None)
            sample.pop("allocation", None)
            sample.pop("allocations", None)
    return value


def report_controls(report: dict[str, Any], semantic_hash: str) -> dict[str, object]:
    correctness = report.get("correctness")
    require(isinstance(correctness, dict), "report correctness controls missing")
    require(all(value is True for value in correctness.values()),
            "report contains a false correctness control")
    return {
        "report_sha256": sha256_bytes(json_bytes(report)),
        "semantic_sha256": semantic_hash,
        "source_sha256": report.get("source_sha256"),
        "source_bytes": report.get("source_bytes"),
        "output_sha256": report.get("output_sha256"),
        "output_bytes": report.get("output_bytes"),
        "output_capacity": report.get("output_capacity"),
        "reference_output_sha256": report.get("reference_output_sha256"),
        "reference_output_bytes": report.get("reference_output_bytes"),
        "retention": report.get("retention"),
        "retention_ceiling_bytes": report.get("retention_ceiling_bytes"),
        "correctness_flags": sorted(correctness),
    }


def expected_command(receipt: dict[str, Any], fixture: dict[str, Any],
                     process_path: Path) -> list[str]:
    spec = receipt["spec"]
    binary = receipt["binary"]["path"]
    return [
        "/usr/bin/time", "-v", "-o", str(process_path), "/usr/bin/taskset", "-c", "14",
        str(binary), "--case", str(spec["case"]), "--input", str(fixture["path_abs"]),
        "--edits", str(spec["edits"]), "--retention", str(spec["retention"]),
        "--warmup", "0", "--samples", "1",
    ]


def verify_driver_log(path: Path, receipt_path: Path, expected: dict[str, object]) -> dict[str, object]:
    value = read_json(path)
    require(isinstance(value, dict), f"{path}: driver log is not an object")
    require(value.get("status") == "qualified", f"{path}: driver status is not qualified")
    require(value.get("version") == expected["version"] and value.get("case") == expected["case"]
            and value.get("lane") == expected["lane"] and value.get("edits") == expected["edits"]
            and value.get("retention") == expected["retention"],
            f"{path}: driver identity mismatch")
    receipt_descriptor = value.get("receipt")
    require(isinstance(receipt_descriptor, dict), f"{path}: receipt descriptor missing")
    require(receipt_descriptor.get("path") == str(receipt_path)
            and receipt_descriptor.get("bytes") == receipt_path.stat().st_size
            and receipt_descriptor.get("sha256") == sha256(receipt_path),
            f"{path}: receipt descriptor mismatch")
    return value


def verify_child(root: Path, expected: dict[str, object], module: Any,
                 fixtures: dict[str, dict[str, Any]], protocol_descriptor: dict[str, object],
                 validator_descriptor: dict[str, object]) -> tuple[dict[str, object], dict[str, Any]]:
    directory = str(expected["directory"])
    child = root / directory
    require(child.is_dir() and not child.is_symlink(), f"{directory}: child directory missing")
    receipt_path = child / "qualification.json"
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), f"{directory}: receipt is not an object")
    require(receipt.get("schema") == "litchi.performance.0918.qualification.v1"
            and receipt.get("status") == "passed" and receipt.get("exit_code") == 0,
            f"{directory}: receipt status/schema failed")
    require(receipt.get("version") == expected["version"] and receipt.get("lane") == expected["lane"],
            f"{directory}: version/lane mismatch")
    spec = receipt.get("spec")
    require(isinstance(spec, dict), f"{directory}: qualification spec missing")
    for key in ("version", "lane", "case", "edits", "retention"):
        if key == "version" or key == "lane":
            actual = receipt.get(key)
        else:
            actual = spec.get(key)
        require(actual == expected[key], f"{directory}: {key} mismatch")
    require(spec.get("label") == expected["receipt_label"]
            and spec.get("pair_id") == expected["receipt_label"],
            f"{directory}: qualification label mismatch")
    require(receipt.get("protocol") == protocol_descriptor,
            f"{directory}: protocol descriptor mismatch")
    require(receipt.get("validator") == validator_descriptor,
            f"{directory}: validator descriptor mismatch")
    resolve_descriptor(receipt.get("protocol"), f"{directory}: protocol")
    resolve_descriptor(receipt.get("validator"), f"{directory}: validator")

    fixture = fixtures[str(expected["case"])]
    fixture_identity = receipt.get("fixture")
    require(isinstance(fixture_identity, dict), f"{directory}: fixture descriptor missing")
    fixture_path = resolve_descriptor(fixture_identity, f"{directory}: fixture")
    require(fixture_path == Path(str(fixture["path_abs"])).resolve()
            and fixture_identity.get("bytes") == fixture["bytes"]
            and fixture_identity.get("sha256") == fixture["sha256"],
            f"{directory}: fixture identity mismatch")

    binary = receipt.get("binary")
    require(isinstance(binary, dict), f"{directory}: binary descriptor missing")
    resolve_descriptor(binary, f"{directory}: binary")
    require(binary.get("sha256") == EXPECTED_BINARY_HASHES[(str(expected["version"]), str(expected["lane"]))],
            f"{directory}: binary hash is not the retained version identity")

    raw_paths: dict[str, Path] = {}
    for key, filename in (("stdout", "stdout.jsonl"), ("stderr", "stderr.log"),
                          ("process", "process.time")):
        path = child / filename
        identity = receipt.get(key)
        require(isinstance(identity, dict), f"{directory}: {key} descriptor missing")
        resolved = resolve_descriptor(identity, f"{directory}: {key}")
        require(resolved == path.resolve(), f"{directory}: {key} path is outside child")
        raw_paths[key] = resolved
    stdout_bytes = raw_paths["stdout"].read_bytes()
    report = parse_report(stdout_bytes, directory)
    validator_fixture = fixtures[str(expected["case"])]
    validator_spec = {
        "lane": expected["lane"], "case": expected["case"], "edits": expected["edits"],
        "retention": expected["retention"], "label": expected["receipt_label"],
    }
    validated = module.validate_report(report, validator_spec, validator_fixture,
                                       str(expected["lane"]), expected_warmups=0,
                                       expected_samples=1)
    require(receipt.get("report") == validated,
            f"{directory}: receipt summary differs from unchanged 0917 validator")
    report_sha = sha256_bytes(json_bytes(report))
    require(validated.get("report_sha256") == report_sha,
            f"{directory}: report canonical hash mismatch")

    process_metrics = receipt.get("process_metrics")
    require(isinstance(process_metrics, dict), f"{directory}: process metrics missing")
    actual_metrics = parse_time_v(raw_paths["process"].read_bytes(), f"{directory}: process")
    require({key: value for key, value in actual_metrics.items() if key != "path"}
            == {key: value for key, value in process_metrics.items() if key != "path"},
            f"{directory}: process metrics changed")
    require(actual_metrics.get("exit_status") == 0, f"{directory}: process exit status failed")
    require(receipt.get("command") == expected_command(receipt, fixture, raw_paths["process"]),
            f"{directory}: command provenance mismatch")
    started = parse_timestamp(receipt.get("started_utc"), f"{directory}: started")
    finished = parse_timestamp(receipt.get("finished_utc"), f"{directory}: finished")
    require(started <= finished, f"{directory}: finish precedes start")

    driver_path = root / f"{directory}.driver.log"
    driver = verify_driver_log(driver_path, receipt_path.resolve(), expected)
    semantic = semantic_signature(report)
    semantic_hash = sha256_bytes(json_bytes(semantic))
    controls = report_controls(report, semantic_hash)
    row = {
        "label": directory,
        "version": expected["version"], "lane": expected["lane"],
        "case": expected["case"], "edits": expected["edits"],
        "retention": expected["retention"],
        "receipt_label": expected["receipt_label"],
        "receipt": descriptor(receipt_path),
        "driver_log": descriptor(driver_path),
        "stdout": descriptor(raw_paths["stdout"]),
        "stderr": descriptor(raw_paths["stderr"]),
        "process": descriptor(raw_paths["process"]),
        "binary": dict(binary), "fixture": dict(fixture_identity),
        "process_metrics": {key: value for key, value in actual_metrics.items() if key != "path"},
        "report": controls,
        "correctness": "all-true",
    }
    return row, report


def build_summary(qualification_root: Path) -> dict[str, object]:
    require(qualification_root.is_dir() and not qualification_root.is_symlink(),
            f"qualification root missing: {qualification_root}")
    module = validator_module()
    protocol, fixtures = protocol_and_fixtures(module)
    protocol_identity = descriptor(PROTOCOL_PATH, relative_to_repo=True)
    validator_identity = descriptor(VALIDATOR_PATH, relative_to_repo=True)
    require(validator_identity["sha256"] == EXPECTED_VALIDATOR_SHA256,
            "unchanged 0917 validator bytes changed")
    expected = expected_rows()
    actual_children = {path.name for path in qualification_root.iterdir() if path.is_dir()}
    require(actual_children == set(expected),
            "qualification child set differs: "
            f"missing={sorted(set(expected) - actual_children)} "
            f"extra={sorted(actual_children - set(expected))}")
    actual_logs = {path.name.removesuffix(".driver.log")
                   for path in qualification_root.glob("*.driver.log") if path.is_file()}
    require(actual_logs == set(expected),
            "qualification driver-log set differs: "
            f"missing={sorted(set(expected) - actual_logs)} "
            f"extra={sorted(actual_logs - set(expected))}")

    rows: list[dict[str, object]] = []
    reports: dict[tuple[str, str, int, str, str], tuple[dict[str, object], dict[str, Any]]] = {}
    for label in sorted(expected):
        row, report = verify_child(qualification_root, expected[label], module, fixtures,
                                   protocol_identity, validator_identity)
        rows.append(row)
        key = (str(row["lane"]), str(row["case"]), int(row["edits"]),
               str(row["retention"]), str(row["version"]))
        reports[key] = (row, report)

    pairs: list[dict[str, object]] = []
    for lane in LANES:
        routes = ROUTES
        for case in CASES:
            for edits in EDITS:
                for retention in routes:
                    base_key = (lane, case, edits, retention, "baseline")
                    candidate_key = (lane, case, edits, retention, "candidate")
                    require(base_key in reports and candidate_key in reports,
                            f"missing version pair: {lane}/{case}/e{edits}/{retention}")
                    base_row, base_report = reports[base_key]
                    candidate_row, candidate_report = reports[candidate_key]
                    base_semantic = semantic_signature(base_report)
                    candidate_semantic = semantic_signature(candidate_report)
                    require(base_semantic == candidate_semantic,
                            f"cross-version semantic identity differs: {lane}/{case}/e{edits}/{retention}")
                    semantic_hash = sha256_bytes(json_bytes(base_semantic))
                    pairs.append({
                        "key": {"lane": lane, "case": case, "edits": edits,
                                "retention": retention},
                        "baseline": base_row["label"], "candidate": candidate_row["label"],
                        "semantic_sha256": semantic_hash, "equal": True,
                    })
    require(len(pairs) == 48, f"semantic pair count is {len(pairs)}")
    summary: dict[str, object] = {
        "schema": "litchi.performance.0918.qualification-summary.v1",
        "generated_utc": now(),
        "qualification_root": str(qualification_root),
        "protocol": protocol_identity,
        "validator": validator_identity,
        "source_identity": {"candidate_revision": EXPECTED_SOURCE_REVISION},
        "counts": {"children": 96, "passed": len(rows), "pairs": len(pairs),
                   "semantic_equal": sum(1 for pair in pairs if pair["equal"])},
        "versions": {
            version: {
                lane: {"sha256": EXPECTED_BINARY_HASHES[(version, lane)]}
                for lane in LANES
            }
            for version in VERSIONS
        },
        "children": rows,
        "pairs": pairs,
        "gates": {
            "all_children_passed": len(rows) == 96,
            "all_0917_validator_controls_true": len(rows) == 96,
            "raw_custody_intact": True,
            "exact_qualification_schedule": len(rows) == 96,
            "cross_version_semantics_equal": len(pairs) == 48
            and all(pair["equal"] for pair in pairs),
        },
        "replay": {
            "command": "PYTHONDONTWRITEBYTECODE=1 python3 -B "
                       "docs/performance/results/change-0918/qualification-replay.py "
                       "--qualification-root /home/zhuhe/litchi-performance-0918/qualification "
                       "--summary docs/performance/results/change-0918/qualification-summary.json",
            "mode": "read-only external receipt/raw audit; no build or probe execution",
        },
    }
    return summary


def without_generation(value: object) -> object:
    if isinstance(value, dict):
        return {key: without_generation(item) for key, item in value.items()
                if key != "generated_utc"}
    if isinstance(value, list):
        return [without_generation(item) for item in value]
    return value


def write_exclusive(path: Path, value: object) -> None:
    require(not path.exists(), f"refusing to replace retained summary: {path}")
    data = json_bytes(value)
    require(len(data) < MAX_PACKET_ARTIFACT, f"qualification summary is at least 1 MiB: {path}")
    temporary = path.with_name(path.name + f".tmp-{os.getpid()}")
    require(not temporary.exists(), f"temporary output already exists: {temporary}")
    try:
        with temporary.open("xb") as stream:
            stream.write(data)
        if path.exists():
            fail(f"summary appeared during write: {path}")
        temporary.replace(path)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--qualification-root", type=Path, default=DEFAULT_EXTERNAL)
    parser.add_argument("--summary", type=Path, default=SUMMARY_PATH)
    parser.add_argument("--write-summary", action="store_true",
                        help="write the summary once; existing summaries are never replaced")
    args = parser.parse_args()
    try:
        root = args.qualification_root.resolve()
        summary = build_summary(root)
        summary_path = args.summary.resolve()
        if args.write_summary:
            write_exclusive(summary_path, summary)
            status = "written"
        else:
            retained = read_json(summary_path)
            require(without_generation(retained) == without_generation(summary),
                    "qualification summary replay mismatch")
            status = "verified"
        print(json.dumps({"status": status, "summary": str(summary_path),
                          "children": summary["counts"]["children"],
                          "pairs": summary["counts"]["pairs"]}, sort_keys=True))
    except (QualificationError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"0918 qualification replay failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
