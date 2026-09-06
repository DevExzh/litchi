#!/usr/bin/env python3
"""Portable replay for the separate 0433 ODS process profiles.

The report checker consumes only retained JSON, text, CSV, resource logs, and
raw ``perf.data`` artifacts.  It never invokes Cargo, the profiled executable,
or ``perf``; this keeps replay possible after the temporary before/after
binaries and worktrees have been removed.  It checks all six required jobs:
three semantic roles (before buffered, after buffered, after streaming) times
``stat`` and ``record``.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any


PROFILE_ROOT = Path(__file__).resolve().parent
BUNDLE_ROOT = PROFILE_ROOT.parent
CHANGE = 433
CPU = 2
WORKERS = 1
SHAPE = "large"
NORMAL = "normal"
KINDS = ("stat", "record")
ROLES = {
    "before-buffered": ("before", "ods_buffered_create"),
    "after-buffered": ("after", "ods_buffered_create"),
    "after-streaming": ("after", "ods_streaming_create"),
}
STAT_EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "L1-dcache-load-misses",
)
RECORD_EVENT = "cycles:u"
RECORD_FREQUENCY = 999
CALL_GRAPH = "fp,127"
ENVIRONMENT_OVERRIDES = {"DEBUGINFOD_URLS": "", "RUSTUP_TOOLCHAIN": "1.98.1"}
PROFILE_SCOPE = (
    "whole process including setup/corpus generation, warmups, measured samples, "
    "output hashing, and the harness oracle/report; excludes perf postprocessing "
    "and the external report verifier"
)
WORKLOAD_VERIFY_MARKER = b"VALID\n"
RSS_RE = re.compile(r"^Maximum resident set size \(kbytes\):\s*(\d+)\s*$")
HEX = set("0123456789abcdef")


class Invalid(ValueError):
    """The retained process-profile bundle is invalid."""


def fail(path: str, message: str) -> None:
    raise Invalid(f"{path}: {message}")


def duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise Invalid(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def reject_constant(value: str) -> Any:
    raise Invalid(f"non-finite JSON constant: {value}")


def load_bytes(raw: bytes, label: str) -> Any:
    try:
        return json.loads(raw.decode("utf-8"), object_pairs_hook=duplicate_keys,
                          parse_constant=reject_constant)
    except (UnicodeError, json.JSONDecodeError, Invalid) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def safe_path(name: str) -> Path:
    path = Path(name)
    if path.is_absolute() or not name:
        fail("path", "must be a relative non-empty path")
    resolved = (BUNDLE_ROOT / path).resolve()
    if not resolved.is_relative_to(BUNDLE_ROOT.resolve()):
        fail(name, "escapes the bundle")
    return resolved


def logical_bytes(name: str) -> bytes:
    path = safe_path(name)
    if path.is_file():
        return path.read_bytes()
    compressed = safe_path(name + ".gz")
    if compressed.is_file():
        try:
            return gzip.decompress(compressed.read_bytes())
        except (OSError, EOFError, gzip.BadGzipFile) as error:
            fail(name + ".gz", f"invalid gzip stream: {error}")
    fail(name, "file is missing")
    raise AssertionError("unreachable")


def load(name: str) -> Any:
    return load_bytes(logical_bytes(name), name)


def obj(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "expected an object")
    return value


def array(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "expected an array")
    return value


def text(value: Any, path: str) -> str:
    if not isinstance(value, str) or not value:
        fail(path, "expected non-empty text")
    return value


def uint(value: Any, path: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        fail(path, "expected unsigned integer")
    return value


def digest(value: Any, path: str) -> str:
    value = text(value, path).lower()
    if len(value) != 64 or any(char not in HEX for char in value):
        fail(path, "expected SHA-256 digest")
    return value


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def file_sha(name: str) -> str:
    return sha(logical_bytes(name))


def protocol() -> dict[str, Any]:
    value = obj(load("protocol.json"), "protocol")
    if value.get("change") != CHANGE:
        fail("protocol.change", f"must be {CHANGE}")
    if value.get("cpu") != CPU or value.get("workers") != WORKERS:
        fail("protocol", "CPU/workers differ from the profile lane")
    if value.get("samples") != 30 or value.get("warmups") != 3:
        fail("protocol", "samples/warmups differ from the formal large lane")
    shapes = obj(value.get("shapes"), "protocol.shapes")
    if shapes.get(SHAPE) != 32_768:
        fail("protocol.shapes.large", "must be 32768 rows")
    scope = text(value.get("profile_scope"), "protocol.profile_scope")
    for marker in ("setup", "oracle", "samples", "hash"):
        if marker not in scope.lower():
            fail("protocol.profile_scope", f"omits {marker}")
    roles = obj(value.get("roles"), "protocol.roles")
    if set(roles) != set(ROLES):
        fail("protocol.roles", "role set differs")
    for role, (_build_dir, selector) in ROLES.items():
        spec = obj(roles.get(role), f"protocol.roles.{role}")
        if spec.get("selector") != selector:
            fail(f"protocol.roles.{role}.selector", f"must be {selector}")
    return value


def verify_source_manifest(manifest: dict[str, Any], label: str) -> None:
    path = text(manifest.get("path"), f"{label}.path")
    raw = logical_bytes(path)
    if sha(raw) != digest(manifest.get("sha256"), f"{label}.sha256"):
        fail(label, "source manifest digest differs")
    entries = obj(load_bytes(raw, f"{label}.payload"), f"{label}.payload")
    if len(entries) != uint(manifest.get("files"), f"{label}.files"):
        fail(label, "source manifest file count differs")


def build(role: str, protocol_value: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    build_dir, _selector = ROLES[role]
    name = f"{build_dir}/build.json"
    value = obj(load(name), name)
    if value.get("change") != CHANGE:
        fail(name, "build change differs")
    source = obj(value.get("source_manifest"), f"{name}.source_manifest")
    verify_source_manifest(source, f"{name}.source_manifest")
    if value.get("protocol_sha256") != file_sha("protocol.json"):
        fail(name, "protocol hash differs")
    if value.get("verifier_sha256") != file_sha("verify-report.py"):
        fail(name, "bound report verifier hash differs")
    binaries = obj(value.get("binaries"), f"{name}.binaries")
    binary = obj(binaries.get(NORMAL), f"{name}.binaries.normal")
    if set(binary) != {"path", "sha256", "bytes"}:
        fail(f"{name}.binaries.normal", "binary identity fields differ")
    expected = Path("/tmp/litchi-goal-0433-binaries") / build_dir / NORMAL
    if Path(text(binary["path"], f"{name}.binaries.normal.path")) != expected:
        fail(f"{name}.binaries.normal.path", "does not bind the prescribed normal binary")
    digest(binary["sha256"], f"{name}.binaries.normal.sha256")
    uint(binary["bytes"], f"{name}.binaries.normal.bytes")
    return value, binary


def artifact(name: str, metadata: Any, label: str) -> bytes:
    row = obj(metadata, label)
    if set(row) != {"sha256", "bytes"}:
        fail(label, "artifact metadata fields differ")
    raw = logical_bytes(name)
    if sha(raw) != digest(row.get("sha256"), f"{label}.sha256"):
        fail(label, "artifact digest differs")
    if len(raw) != uint(row.get("bytes"), f"{label}.bytes"):
        fail(label, "artifact byte count differs")
    return raw


def expected_artifact_names(run_prefix: str, kind: str) -> set[str]:
    common = {
        f"{run_prefix}/workload.json",
        f"{run_prefix}/workload-catalog.json",
        f"{run_prefix}/workload-verify.json",
        f"{run_prefix}/workload-verify.stderr.log",
        f"{run_prefix}/process.stdout.log",
        f"{run_prefix}/process.stderr.log",
        f"{run_prefix}/resource.log",
    }
    if kind == "stat":
        return common | {f"{run_prefix}/perf-stat.csv"}
    return common | {
        f"{run_prefix}/perf.data",
        f"{run_prefix}/perf-report.txt",
        f"{run_prefix}/perf-report.stderr.log",
    }


def check_workload_argv(receipt: dict[str, Any], binary: dict[str, Any], selector: str, label: str) -> None:
    argv = array(receipt.get("workload_argv"), f"{label}.workload_argv")
    if not all(isinstance(item, str) for item in argv):
        fail(label, "workload argv contains non-text values")
    expected_prefix = [
        binary["path"], "--case", selector, "--semantic-shape", SHAPE,
        "--workers", str(WORKERS), "--samples", "30", "--warmup", "3",
    ]
    if argv[:len(expected_prefix)] != expected_prefix:
        fail(label, "workload argv does not bind the frozen selector and shape")
    if len(argv) != len(expected_prefix) + 4:
        fail(label, "workload argv has unexpected flags")
    if argv[-4] != "--json" or Path(argv[-3]).name != "workload.json":
        fail(label, "workload report path differs")
    if argv[-2] != "--corpus-manifest" or Path(argv[-1]).name != "workload-catalog.json":
        fail(label, "workload corpus path differs")


def check_profile_argv(receipt: dict[str, Any], kind: str, label: str) -> None:
    argv = array(receipt.get("profile_argv"), f"{label}.profile_argv")
    if len(argv) < 8 or argv[:6] != ["taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o"]:
        fail(label, "profile wrapper is not taskset CPU2 plus GNU time -v")
    if argv[7] != "perf":
        fail(label, "profile command does not invoke perf")
    if Path(argv[6]).name != "resource.log":
        fail(label, "resource log path differs")
    if kind == "stat":
        expected = ["perf", "stat", "--no-big-num", "-x,", "-e", ",".join(STAT_EVENTS)]
        if argv[7:13] != expected or "--" not in argv:
            fail(label, "perf stat argv differs")
        output_flags = [index for index, value in enumerate(argv) if value == "-o"]
        if not output_flags or Path(argv[output_flags[-1] + 1]).name != "perf-stat.csv":
            fail(label, "perf stat output path differs")
    else:
        expected = ["perf", "record", "--no-buildid-cache", "-e", RECORD_EVENT,
                    "-F", str(RECORD_FREQUENCY), "--call-graph", CALL_GRAPH]
        if argv[7:16] != expected or "--" not in argv:
            fail(label, "perf record argv differs")
        output_flags = [index for index, value in enumerate(argv) if value == "-o"]
        if not output_flags or Path(argv[output_flags[-1] + 1]).name != "perf.data":
            fail(label, "perf record output path differs")


def check_report_argv(receipt: dict[str, Any], kind: str, label: str) -> None:
    argv = receipt.get("report_argv")
    if kind == "stat":
        if argv is not None:
            fail(label, "stat profile unexpectedly carries a perf report command")
        return
    values = array(argv, f"{label}.report_argv")
    expected = ["perf", "report", "--stdio", "--no-children", "--percent-limit", "0", "-i"]
    if values[:len(expected)] != expected or len(values) != len(expected) + 1:
        fail(label, "symbolized perf report argv differs")
    if Path(values[-1]).name != "perf.data":
        fail(label, "symbolized report input differs")


def check_resource(raw: bytes, label: str) -> int:
    lines = raw.decode("utf-8", "replace").splitlines()
    values = [int(match.group(1)) for line in lines if (match := RSS_RE.match(line.lstrip()))]
    if len(values) != 1:
        fail(label, "resource log must contain one maximum RSS line")
    return values[0]


def check_stat(raw: bytes, label: str) -> dict[str, str]:
    values: dict[str, str] = {}
    for line in raw.decode("utf-8", "replace").splitlines():
        fields = line.split(",")
        if len(fields) < 3:
            continue
        event = fields[2].strip()
        if event not in STAT_EVENTS:
            continue
        value = fields[0].strip()
        if not value or value.startswith("<"):
            fail(label, f"counter {event} is unavailable: {value}")
        try:
            float(value)
        except ValueError:
            fail(label, f"counter {event} is not numeric: {value}")
        values[event] = value
    if set(values) != set(STAT_EVENTS):
        fail(label, f"counter set differs: {sorted(values)}")
    return values


def verify_job(role: str, kind: str, receipt_name: str, protocol_value: dict[str, Any],
               build_value: dict[str, Any], binary: dict[str, Any]) -> dict[str, Any]:
    label = f"{role}.{kind}"
    receipt = obj(load(receipt_name), receipt_name)
    if receipt.get("schema") != "ods_process_profile_v1":
        fail(label, "profile schema differs")
    if receipt.get("change") != CHANGE or receipt.get("role") != role or receipt.get("kind") != kind:
        fail(label, "role/kind identity differs")
    build_dir, selector = ROLES[role]
    if receipt.get("build_directory") != build_dir or receipt.get("selector") != selector:
        fail(label, "build directory or selector differs")
    if receipt.get("mode") != NORMAL or receipt.get("shape") != SHAPE:
        fail(label, "profile is not the normal large lane")
    if receipt.get("workers") != WORKERS or receipt.get("samples") != 30 or receipt.get("warmups") != 3:
        fail(label, "worker/sample settings differ")
    if receipt.get("status") != "pass" or receipt.get("profile_exit_code") != 0 or receipt.get("verifier_exit_code") != 0:
        fail(label, "profile or workload verifier did not pass")
    if kind == "record" and receipt.get("report_exit_code") != 0:
        fail(label, "perf report did not pass")
    if receipt.get("revision") != build_value.get("revision"):
        fail(label, "revision differs from build")
    if receipt.get("source_manifest") != build_value.get("source_manifest"):
        fail(label, "source manifest differs from build")
    if receipt.get("binary") != binary:
        fail(label, "binary identity differs from build")
    if receipt.get("protocol_sha256") != file_sha("protocol.json"):
        fail(label, "protocol hash differs")
    if receipt.get("profile_driver_sha256") != file_sha("profile/capture.py"):
        fail(label, "profile capture driver hash differs")
    if receipt.get("verifier_sha256") != file_sha("verify-report.py"):
        fail(label, "workload verifier hash differs")
    if receipt.get("scope") != PROFILE_SCOPE:
        fail(label, "process scope differs")
    if receipt.get("protocol_profile_scope") != protocol_value.get("profile_scope"):
        fail(label, "protocol profile scope binding differs")
    if receipt.get("environment_overrides") != ENVIRONMENT_OVERRIDES:
        fail(label, "profiler environment override differs")
    if receipt.get("tracked_tree_clean_before_and_after") is not True:
        fail(label, "tracked-tree cleanliness was not retained")
    for status_key in ("source_status_before", "source_status_after"):
        statuses = array(receipt.get(status_key), f"{label}.{status_key}")
        if any(not isinstance(status, str) or not status.startswith("?? ") for status in statuses):
            fail(label, f"{status_key} contains a tracked worktree change")
    tools = obj(receipt.get("tool_versions"), f"{label}.tool_versions")
    for tool in ("perf", "time", "python"):
        text(tools.get(tool), f"{label}.tool_versions.{tool}")
    check_workload_argv(receipt, binary, selector, label)
    check_profile_argv(receipt, kind, label)
    check_report_argv(receipt, kind, label)
    run_prefix = f"{build_dir}/profiles/{role}/{kind}"
    artifacts = obj(receipt.get("artifacts"), f"{label}.artifacts")
    if set(artifacts) != expected_artifact_names(run_prefix, kind):
        fail(label, "artifact inventory differs")
    raw_artifacts = {
        name: artifact(name, metadata, f"{label}.artifacts.{name}")
        for name, metadata in artifacts.items()
    }
    rss = check_resource(raw_artifacts[f"{run_prefix}/resource.log"], label)
    verification = raw_artifacts[f"{run_prefix}/workload-verify.json"]
    if verification != WORKLOAD_VERIFY_MARKER:
        fail(label, "retained workload verifier output is not the exact VALID marker")
    report = obj(load_bytes(raw_artifacts[f"{run_prefix}/workload.json"], label + ".workload"), label + ".workload")
    environment = obj(report.get("environment"), label + ".workload.environment")
    if environment.get("git_revision") != build_value.get("revision"):
        fail(label, "workload report revision differs")
    identity = obj(report.get("binary_identity"), label + ".workload.binary_identity")
    if identity.get("path") != binary["path"] or identity.get("binary_sha256") != binary["sha256"] or identity.get("binary_bytes") != binary["bytes"]:
        fail(label, "workload binary identity differs")
    configuration = obj(report.get("configuration"), label + ".workload.configuration")
    if configuration.get("samples_per_case") != 30 or configuration.get("warmup_iterations_per_case") != 3:
        fail(label, "workload configuration sample settings differ")
    if configuration.get("execution_workers") != [WORKERS] or configuration.get("semantic_shapes") != [SHAPE]:
        fail(label, "workload configuration shape/worker differs")
    results = array(report.get("results"), label + ".workload.results")
    if len(results) != 1 or obj(results[0], label + ".workload.results[0]").get("case") != selector:
        fail(label, "workload result selector differs")
    if kind == "stat":
        counters = check_stat(raw_artifacts[f"{run_prefix}/perf-stat.csv"], label + ".perf-stat")
        profile_summary = {"kind": kind, "counters": counters}
    else:
        data = raw_artifacts[f"{run_prefix}/perf.data"]
        report_text = raw_artifacts[f"{run_prefix}/perf-report.txt"].decode("utf-8", "replace")
        if not data or not report_text.strip():
            fail(label, "raw perf data or symbolized report is empty")
        if "Samples:" not in report_text and "Overhead" not in report_text:
            fail(label, "symbolized perf report has no sample table")
        profile_summary = {
            "kind": kind,
            "perf_data_sha256": sha(data),
            "perf_data_bytes": len(data),
            "symbolized_report_sha256": sha(report_text.encode("utf-8")),
        }
    return {"role": role, "kind": kind, "revision": build_value["revision"], "rss_kib": rss, **profile_summary}


def verify() -> dict[str, Any]:
    protocol_value = protocol()
    protocol_hash = file_sha("protocol.json")
    index = array(load("profile-index.json"), "profile-index.json")
    expected = {
        f"{ROLES[role][0]}/profiles/{role}/{kind}/receipt.json"
        for role in ROLES for kind in KINDS
    }
    if set(index) != expected or len(index) != len(expected):
        fail("profile-index.json", "does not list exactly the six required profile jobs")
    builds = {role: build(role, protocol_value) for role in ROLES}
    rows = []
    for receipt_name in index:
        parts = Path(receipt_name).parts
        if len(parts) != 5 or parts[1] != "profiles" or parts[-1] != "receipt.json":
            fail(receipt_name, "receipt path shape differs")
        role, kind = parts[2], parts[3]
        if role not in ROLES or kind not in KINDS:
            fail(receipt_name, "receipt role/kind is outside the frozen matrix")
        build_value, binary = builds[role]
        row = verify_job(role, kind, receipt_name, protocol_value, build_value, binary)
        if load(receipt_name).get("protocol_sha256") != protocol_hash:
            fail(receipt_name, "protocol hash differs")
        rows.append(row)
    return {
        "status": "pass",
        "change": CHANGE,
        "jobs": len(rows),
        "rows": rows,
        "scope": PROFILE_SCOPE,
        "performance_claim": None,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="retain explicit check-mode spelling")
    args = parser.parse_args()
    del args
    try:
        print(json.dumps(verify(), indent=2, sort_keys=True))
        return 0
    except (OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
