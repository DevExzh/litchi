"""Validate and compare the matched 0741 baseline/candidate captures.

The native workload itself is deliberately owned by the root coordinator.  This
module emits the deterministic serial matrix and audits receipts and reports
after the workloads have run.  It does not invoke Cargo or a native binary.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import pathlib
import random
import re
import statistics
import subprocess
import sys
from typing import Any, Iterable


P = pathlib.Path(__file__).resolve().parent
ROOT = P.parents[3]
CAPTURES = P / "captures"
PHASES = ("plan_ns", "commit_ns", "publication_ns", "reopen_ns", "lifecycle_ns")
TIMING_METRICS = PHASES
ARMS = ("baseline", "candidate")
LANES = ("native", "allocation")
ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
CPU_AFFINITY = "12"
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_SEED = 741_074
PAIR_FLAG_PERCENT = 5.0
FROZEN_INPUTS = {
    "source_manifest": {"baseline": "source.json", "candidate": "candidate-source.json"},
    "build_manifest": {"baseline": "build.json", "candidate": "candidate-build.json"},
    "oracle": {"baseline": "oracle.json", "candidate": "oracle.json"},
    "workspace_inputs": {"baseline": "workspace-inputs.json", "candidate": "workspace-inputs.json"},
    "constraints": {"baseline": "constraints.json", "candidate": "constraints.json"},
    "compare": {"baseline": "compare.py", "candidate": "compare.py"},
}
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")


class EvidenceError(Exception):
    """Raised for a malformed evidence packet."""


MISSING = object()


def read_json(path: pathlib.Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: pathlib.Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha256_file(path: pathlib.Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def fail(failures: list[str], message: str) -> None:
    failures.append(message)


def packet_path(raw: str | pathlib.Path) -> pathlib.Path:
    path = pathlib.Path(raw)
    resolved = (P / path).resolve() if not path.is_absolute() else path.resolve()
    if resolved != P and P not in resolved.parents:
        raise EvidenceError(f"path escapes packet: {raw}")
    return resolved


def source_hash_at_revision(revision: str, name: str) -> str | None:
    try:
        data = subprocess.check_output(
            ["git", "show", f"{revision}:{name}"], cwd=ROOT, stderr=subprocess.DEVNULL
        )
    except (OSError, subprocess.CalledProcessError):
        return None
    return hashlib.sha256(data).hexdigest()


def source_manifest(path: pathlib.Path, failures: list[str]) -> dict[str, Any]:
    try:
        value = read_json(path)
    except (OSError, json.JSONDecodeError) as error:
        fail(failures, f"source/{path.name}: invalid manifest: {error}")
        return {"head": None, "files": {}}
    if not isinstance(value, dict) or not isinstance(value.get("head"), str):
        fail(failures, f"source/{path.name}: missing string head")
        return {"head": None, "files": {}}
    files = value.get("files")
    if not isinstance(files, dict) or not files:
        fail(failures, f"source/{path.name}: files is empty or malformed")
        return {"head": value["head"], "files": {}}
    valid: dict[str, str] = {}
    for name, expected in files.items():
        if (
            not isinstance(name, str)
            or pathlib.Path(name).is_absolute()
            or ".." in pathlib.Path(name).parts
            or not isinstance(expected, str)
            or len(expected) != 64
        ):
            fail(failures, f"source/{path.name}: malformed entry {name!r}")
            continue
        valid[name] = expected
        current = ROOT / name
        current_hash = sha256_file(current) if current.is_file() else None
        if current_hash == expected:
            continue
        revision_hash = source_hash_at_revision(value["head"], name)
        if revision_hash != expected:
            fail(failures, f"source/{path.name}: custody mismatch {name}")
    return {"head": value["head"], "files": valid}


def cleanup_witness(path: pathlib.Path | None, failures: list[str]) -> dict[tuple[str, str], dict[str, Any]]:
    if path is None:
        return {}
    try:
        value = read_json(path)
    except (OSError, json.JSONDecodeError) as error:
        fail(failures, f"cleanup: invalid witness: {error}")
        return {}
    if not isinstance(value, dict) or value.get("version") != 1 or value.get("removed") is not True:
        fail(failures, "cleanup: witness requires version 1 and removed=true")
        return {}
    entries = value.get("entries")
    if not isinstance(entries, list):
        fail(failures, "cleanup: witness entries must be a list")
        return {}
    result: dict[tuple[str, str], dict[str, Any]] = {}
    for entry in entries:
        if not isinstance(entry, dict):
            fail(failures, "cleanup: witness entry is malformed")
            continue
        arm, lane = entry.get("arm"), entry.get("lane")
        key = (arm, lane)
        if arm not in ARMS or lane not in LANES or key in result:
            fail(failures, f"cleanup: duplicate or invalid entry {key}")
            continue
        if (
            not isinstance(entry.get("path"), str)
            or not isinstance(entry.get("sha256"), str)
            or not HEX64.fullmatch(entry["sha256"])
            or type(entry.get("bytes")) is not int
            or entry["bytes"] < 0
            or entry.get("removed") is not True
        ):
            fail(failures, f"cleanup: incomplete witness entry {key}")
            continue
        result[key] = entry
    expected = {(arm, lane) for arm in ARMS for lane in LANES}
    if set(result) != expected:
        fail(failures, "cleanup: witness must contain exactly both arms and both lanes")
    return result


def build_rows(
    path: pathlib.Path, arm: str, failures: list[str], cleanup: dict[tuple[str, str], dict[str, Any]]
) -> dict[str, dict[str, Any]]:
    try:
        value = read_json(path)
    except (OSError, json.JSONDecodeError) as error:
        fail(failures, f"build/{path.name}: invalid manifest: {error}")
        return {}
    rows: Any = value
    if isinstance(value, dict):
        rows = value.get("builds", value.get("results", value.get("lanes")))
    if not isinstance(rows, list):
        fail(failures, f"build/{path.name}: expected a list of lane records")
        return {}
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        if not isinstance(row, dict) or not isinstance(row.get("lane"), str):
            fail(failures, f"build/{path.name}: malformed lane record")
            continue
        lane = row["lane"]
        if lane in result:
            fail(failures, f"build/{path.name}: duplicate lane {lane}")
        result[lane] = row
        binary = row.get("binary")
        if not isinstance(binary, str) or not isinstance(row.get("binary_sha256"), str):
            fail(failures, f"build/{path.name}/{lane}: incomplete binary identity")
            continue
        binary_path = pathlib.Path(binary)
        witness = cleanup.get((arm, lane))
        if witness is not None:
            if (
                pathlib.Path(witness["path"]).resolve() != binary_path.resolve()
                or witness["sha256"] != row["binary_sha256"]
                or witness["bytes"] != row.get("binary_bytes")
                or binary_path.exists()
            ):
                fail(failures, f"build/{path.name}/{lane}: cleanup witness does not bind removed binary")
        elif not binary_path.is_file():
            fail(failures, f"build/{path.name}/{lane}: missing binary without cleanup witness")
        else:
            if sha256_file(binary_path) != row["binary_sha256"]:
                fail(failures, f"build/{path.name}/{lane}: binary SHA-256 mismatch")
            if row.get("binary_bytes") != binary_path.stat().st_size:
                fail(failures, f"build/{path.name}/{lane}: binary size mismatch")
        if type(row.get("exit")) is not int or row["exit"] != 0:
            fail(failures, f"build/{path.name}/{lane}: build exit was not zero")
    for lane in LANES:
        if lane not in result:
            fail(failures, f"build/{path.name}: missing {lane} lane")
    return result


def packet_bindings(failures: list[str]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for filename in ("workspace-inputs.json", "constraints.json"):
        path = P / filename
        if not path.is_file():
            fail(failures, f"packet: missing {filename}")
            continue
        try:
            value = read_json(path)
        except (OSError, json.JSONDecodeError) as error:
            fail(failures, f"packet/{filename}: invalid: {error}")
            continue
        files = value.get("files") if filename == "workspace-inputs.json" else value
        if not isinstance(files, dict):
            fail(failures, f"packet/{filename}: file map missing")
            continue
        result[filename] = files
        for name, expected in files.items():
            target = ROOT / name
            if not target.is_file() or sha256_file(target) != expected:
                fail(failures, f"packet/{filename}: binding mismatch {name}")
    return result


def frozen_input_hashes(arm: str, failures: list[str]) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    for name, paths in FROZEN_INPUTS.items():
        relative = paths[arm]
        path = P / relative
        if not path.is_file():
            fail(failures, f"frozen-input/{arm}: missing {relative}")
            continue
        result[name] = {"path": relative, "sha256": sha256_file(path)}
    return result


def validate_frozen_inputs(
    receipt: dict[str, Any], expected: dict[str, dict[str, str]], failures: list[str], label: str
) -> None:
    actual = receipt.get("frozen_inputs")
    if not isinstance(actual, dict) or set(actual) != set(expected):
        fail(failures, f"{label}: frozen_inputs key set mismatch")
        return
    for name, binding in expected.items():
        value = actual.get(name)
        if (
            not isinstance(value, dict)
            or set(value) != {"path", "sha256"}
            or value.get("path") != binding["path"]
            or value.get("sha256") != binding["sha256"]
        ):
            fail(failures, f"{label}: frozen_inputs/{name} mismatch")


def cases_from_oracle(path: pathlib.Path, failures: list[str]) -> list[str]:
    try:
        value = read_json(path)
    except (OSError, json.JSONDecodeError) as error:
        fail(failures, f"oracle: invalid: {error}")
        return []
    if not isinstance(value, dict) or set(value) != {
        "pptx_cross_copy_plain_lifecycle",
        "pptx_cross_copy_media_rich_lifecycle",
    }:
        fail(failures, "oracle: expected exactly the two owned PPTX lifecycle cases")
        return list(value) if isinstance(value, dict) else []
    return list(value)


def matrix(cases: Iterable[str]) -> list[dict[str, Any]]:
    """Return the fixed serial schedule, pairing arms at process granularity."""
    result: list[dict[str, Any]] = []
    order_index = 0
    for lane, repeats, samples, warmups in (
        ("native", 6, 30, 3),
        ("allocation", 3, 1, 0),
    ):
        for repeat in range(repeats):
            arms = ARMS if repeat % 2 == 0 else tuple(reversed(ARMS))
            for arm in arms:
                for case in cases:
                    result.append(
                        {
                            "arm": arm,
                            "lane": lane,
                            "repeat": repeat,
                            "case": case,
                            "samples": samples,
                            "warmups": warmups,
                            "pair_id": f"{lane}/{case}/repeat-{repeat:02d}",
                            "order_index": order_index,
                        }
                    )
                    order_index += 1
    return result


def build_for(arm: str, lane: str, builds: dict[str, dict[str, dict[str, Any]]]) -> dict[str, Any]:
    try:
        return builds[arm][lane]
    except KeyError as error:
        raise EvidenceError(f"missing {arm}/{lane} build") from error


def expected_command(row: dict[str, Any], builds: dict[str, dict[str, dict[str, Any]]]) -> list[str]:
    binary = build_for(row["arm"], row["lane"], builds)["binary"]
    return [
        "taskset",
        "-c",
        CPU_AFFINITY,
        binary,
        "--warmup",
        str(row["warmups"]),
        "--samples",
        str(row["samples"]),
        "--case",
        row["case"],
        "--json",
    ]


def option_value(command: list[Any], option: str) -> str | None:
    try:
        index = command.index(option)
    except ValueError:
        return None
    if index + 1 >= len(command):
        return None
    return str(command[index + 1])


def receipt_report_path(receipt: dict[str, Any], failures: list[str], label: str) -> pathlib.Path | None:
    raw = receipt.get("report", receipt.get("json", receipt.get("output")))
    if not isinstance(raw, str):
        fail(failures, f"{label}: receipt has no report path")
        return None
    try:
        path = packet_path(raw)
    except EvidenceError as error:
        fail(failures, f"{label}: {error}")
        return None
    if not path.is_file():
        fail(failures, f"{label}: report is missing {raw}")
        return None
    return path


def receipts(
    rows: list[dict[str, Any]],
    builds: dict[str, dict[str, dict[str, Any]]],
    frozen: dict[str, dict[str, dict[str, str]]],
    failures: list[str],
) -> dict[tuple[str, str, int, str], dict[str, Any]]:
    path = CAPTURES / "manifest.json"
    if not path.is_file():
        fail(failures, f"captures: missing {path}")
        return {}
    try:
        raw = read_json(path)
    except (OSError, json.JSONDecodeError) as error:
        fail(failures, f"captures: invalid manifest: {error}")
        return {}
    if not isinstance(raw, list) or len(raw) != len(rows):
        fail(failures, f"captures: expected {len(rows)} serial receipts")
        return {}
    result: dict[tuple[str, str, int, str], dict[str, Any]] = {}
    seen_files: set[pathlib.Path] = set()
    previous_end = -1
    for index, (expected, receipt) in enumerate(zip(rows, raw)):
        label = f"captures/{index:03d}"
        if not isinstance(receipt, dict):
            fail(failures, f"{label}: receipt is not an object")
            continue
        for key in ("arm", "lane", "repeat", "case", "samples", "warmups", "pair_id", "order_index"):
            if receipt.get(key) != expected[key]:
                fail(failures, f"{label}: {key} does not match matrix")
        validate_frozen_inputs(receipt, frozen[expected["arm"]], failures, label)
        if type(receipt.get("exit")) is not int or receipt["exit"] != 0:
            fail(failures, f"{label}: process exit was not zero")
        start, end = receipt.get("monotonic_start"), receipt.get("monotonic_end")
        if not isinstance(start, int) or not isinstance(end, int) or not start < end:
            fail(failures, f"{label}: invalid monotonic interval")
        elif start < previous_end:
            fail(failures, f"{label}: captures overlap; matrix is not serial")
        else:
            previous_end = end
        report_path = receipt_report_path(receipt, failures, label)
        if report_path is not None:
            try:
                report_path.relative_to(CAPTURES.resolve())
            except ValueError:
                fail(failures, f"{label}: report is outside captures/")
        command = receipt.get("command")
        if report_path is not None:
            try:
                expected_command_value = expected_command(expected, builds) + [str(report_path)]
            except EvidenceError as error:
                fail(failures, f"{label}: {error}")
            else:
                if command != expected_command_value:
                    fail(failures, f"{label}: command does not exactly match expected argv")
        elif not isinstance(command, list):
            fail(failures, f"{label}: command is missing")
        files = receipt.get("files")
        if report_path is None or not isinstance(files, dict):
            if not isinstance(files, dict):
                fail(failures, f"{label}: file hash map is missing")
        else:
            expected_paths = {
                report_path,
                report_path.with_suffix(".stdout"),
                report_path.with_suffix(".stderr"),
            }
            canonical_files: dict[pathlib.Path, str] = {}
            for name, expected_hash in files.items():
                try:
                    file_path = packet_path(name)
                except EvidenceError as error:
                    fail(failures, f"{label}: {error}")
                    continue
                if file_path in canonical_files:
                    fail(failures, f"{label}: duplicate canonical file path {name}")
                canonical_files[file_path] = expected_hash
                if (
                    not isinstance(expected_hash, str)
                    or not HEX64.fullmatch(expected_hash)
                    or not file_path.is_file()
                    or sha256_file(file_path) != expected_hash
                ):
                    fail(failures, f"{label}: file hash mismatch {name}")
            if set(canonical_files) != expected_paths:
                fail(failures, f"{label}: files must exactly be report/stdout/stderr")
            for file_path in expected_paths:
                if file_path in seen_files:
                    fail(failures, f"{label}: duplicate file path across receipts {file_path}")
                seen_files.add(file_path)
        if report_path is not None:
            key = (expected["arm"], expected["lane"], expected["repeat"], expected["case"])
            if key in result:
                fail(failures, f"{label}: duplicate process key")
            result[key] = {"row": expected, "receipt": receipt, "path": report_path}
    if CAPTURES.is_dir():
        for path in CAPTURES.rglob("*"):
            if path.is_file() and path.resolve() != (CAPTURES / "manifest.json").resolve() and path.resolve() not in seen_files:
                fail(failures, f"captures: unmanifested file {path}")
    return result


def expected_cpu_affinity() -> str:
    # 0740 and the frozen harness use CPU 12.  Keep this explicit so a report
    # cannot silently mix an unpinned process into the matched comparison.
    return CPU_AFFINITY


def projection(
    report: dict[str, Any],
    row: dict[str, Any],
    arm: str,
    builds: dict[str, dict[str, dict[str, Any]]],
    sources: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    if report.get("schema_version") != 1 or len(report.get("results", [])) != 1:
        raise EvidenceError("report schema or result count mismatch")
    results = report["results"]
    if not isinstance(results, list) or not isinstance(results[0], dict):
        raise EvidenceError("report result is malformed")
    result = results[0]
    source = result.get("source")
    x = source.get("pptx_cross_copy") if isinstance(source, dict) else None
    if not isinstance(x, dict):
        raise EvidenceError("PPTX cross-copy evidence is missing")
    case = row["case"]
    if result.get("case") != case or result.get("corpus", {}).get("name") != case:
        raise EvidenceError("case/corpus mismatch")
    configuration = report.get("configuration", {})
    if configuration.get("samples_per_case") != row["samples"]:
        raise EvidenceError("sample count mismatch")
    if configuration.get("warmup_iterations_per_case") != row["warmups"]:
        raise EvidenceError("warmup count mismatch")
    if configuration.get("cases") != [case]:
        raise EvidenceError("configuration case list mismatch")
    build = build_for(arm, row["lane"], builds)
    identity = report.get("binary_identity", {})
    for key in ("path", "binary_sha256", "binary_bytes"):
        build_key = "binary" if key == "path" else key
        if identity.get(key) != build.get(build_key):
            raise EvidenceError(f"binary identity mismatch: {key}")
    environment = report.get("environment", {})
    if environment.get("cpu_affinity") != expected_cpu_affinity():
        raise EvidenceError("CPU affinity mismatch")
    if environment.get("git_revision") != sources[arm].get("head"):
        raise EvidenceError("source revision mismatch")
    tool = report.get("tool", {})
    if tool.get("profile") != "release":
        raise EvidenceError("tool profile is not release")
    expected_instrumentation = (
        "none" if row["lane"] == "native" else "system_allocator_operation_scoped"
    )
    if tool.get("instrumentation") != expected_instrumentation:
        raise EvidenceError("instrumentation mismatch")
    if row["lane"] == "allocation" and tool.get("allocator_counter_revision") != "serialized_region_peak_v3":
        raise EvidenceError("allocator counter revision mismatch")
    gates = x.get("gates")
    if not isinstance(gates, dict) or not gates or any(value is not True for value in gates.values()):
        raise EvidenceError("one or more correctness/refusal gates are false")
    output_hash = result.get("output_sha256")
    output_hashes = x.get("output_sha256")
    if not isinstance(output_hash, str) or len(output_hash) != 64:
        raise EvidenceError("output hash identity is malformed")
    if not isinstance(output_hashes, list) or len(output_hashes) != row["samples"]:
        raise EvidenceError("output hash vector length mismatch")
    if any(value != output_hash for value in output_hashes):
        raise EvidenceError("output hash vector is not stable")
    corpus = result.get("corpus")
    if not isinstance(corpus, dict) or corpus.get("archive_sha256") != x.get("destination_archive_sha256"):
        raise EvidenceError("destination archive identity mismatch")
    elapsed = result.get("elapsed_ns")
    if not isinstance(elapsed, dict) or elapsed.get("samples") != x.get("lifecycle_ns"):
        raise EvidenceError("elapsed/lifecycle vector mismatch")
    if len(x.get("lifecycle_ns", [])) != row["samples"]:
        raise EvidenceError("lifecycle vector length mismatch")
    if sorted(elapsed.get("sample_order", [])) != list(range(row["samples"])):
        raise EvidenceError("sample order is not the acquisition permutation")
    if x["lifecycle_ns"] != sorted(x["lifecycle_ns"]):
        raise EvidenceError("lifecycle vector is not sorted as required by the harness")
    for phase in PHASES:
        values = x.get(phase)
        if not isinstance(values, list) or len(values) != row["samples"]:
            raise EvidenceError(f"{phase} vector length mismatch")
        if any(type(value) is not int or value <= 0 for value in values):
            raise EvidenceError(f"{phase} contains a non-positive/non-integer value")
    for index, total in enumerate(x["lifecycle_ns"]):
        if sum(x[name][index] for name in ("plan_ns", "commit_ns", "publication_ns")) > total:
            raise EvidenceError("phase sum exceeds lifecycle")
    allocation = result.get("operation_metrics", {}).get("allocation")
    if not isinstance(allocation, dict):
        raise EvidenceError("allocation metric object is missing")
    expected_status = "unavailable" if row["lane"] == "native" else "measured"
    if allocation.get("status") != expected_status or not isinstance(allocation.get("scope"), str):
        raise EvidenceError("allocation status mismatch")
    if set(allocation) - {"status", "scope"} != set(ALLOCATION_FIELDS):
        raise EvidenceError("allocation field set mismatch")
    if row["lane"] == "allocation":
        for name, metric in allocation.items():
            if name in ("status", "scope"):
                continue
            if not isinstance(metric, dict) or not isinstance(metric.get("values"), list):
                raise EvidenceError(f"allocation field {name} is malformed")
            if (
                metric.get("status") != "measured"
                or not isinstance(metric.get("scope"), str)
                or set(metric) != {"status", "scope", "values"}
                or len(metric["values"]) != row["samples"]
            ):
                raise EvidenceError(f"allocation field {name} status/cardinality mismatch")
            if any(type(value) is not int or value < 0 for value in metric["values"]):
                raise EvidenceError(f"allocation field {name} contains an invalid value")
    else:
        for name in ALLOCATION_FIELDS:
            metric = allocation[name]
            if (
                not isinstance(metric, dict)
                or metric.get("status") != "unavailable"
                or not isinstance(metric.get("scope"), str)
                or set(metric) != {"status", "scope"}
            ):
                raise EvidenceError(f"unavailable allocation field {name} is malformed")
    stable = {name: value for name, value in x.items() if name not in PHASES + ("output_sha256",)}
    return {
        "case": result["case"],
        "corpus": corpus,
        "sink": result.get("sink"),
        "cross_copy": stable,
        "output_sha256": output_hash,
    }


def deep_diff(expected: Any, actual: Any, path: str = "") -> list[dict[str, Any]]:
    if isinstance(expected, dict) and isinstance(actual, dict):
        diffs: list[dict[str, Any]] = []
        for key in sorted(set(expected) | set(actual)):
            child = f"{path}.{key}" if path else str(key)
            if key not in expected:
                diffs.append({"path": child, "expected": "<missing>", "actual": actual[key]})
            elif key not in actual:
                diffs.append({"path": child, "expected": expected[key], "actual": "<missing>"})
            else:
                diffs.extend(deep_diff(expected[key], actual[key], child))
        return diffs
    if isinstance(expected, list) and isinstance(actual, list):
        diffs = []
        for index in range(max(len(expected), len(actual))):
            child = f"{path}[{index}]"
            if index >= len(expected):
                diffs.append({"path": child, "expected": "<missing>", "actual": actual[index]})
            elif index >= len(actual):
                diffs.append({"path": child, "expected": expected[index], "actual": "<missing>"})
            else:
                diffs.extend(deep_diff(expected[index], actual[index], child))
        return diffs
    if type(expected) is not type(actual) or expected != actual:
        return [{"path": path or "$", "expected": expected, "actual": actual}]
    return []


PROTECTED_ALLOWLIST_PATHS = (
    "case",
    "corpus.name",
    "corpus.generator",
    "corpus.package_format",
    "corpus.shape",
    "corpus.payload_kind",
    "corpus.entry_count",
    "corpus.archive_member_count",
    "corpus.entry_bytes",
    "corpus.uncompressed_payload_bytes",
    "corpus.target_entry",
    "corpus.target_payload_bytes",
    "corpus.target_payload_sha256",
    "corpus.xlsx",
    "cross_copy.implementation",
    "cross_copy.timing_scope",
    "cross_copy.performance_claim",
    "cross_copy.source_archive_sha256",
    "cross_copy.source_slide",
    "cross_copy.destination_slide",
    "cross_copy.insertion_position",
    "cross_copy.source_slide_name",
    "cross_copy.destination_slide_name",
    "cross_copy.destination_slide_count_before",
    "cross_copy.destination_slide_count_after",
    "cross_copy.planned_part_count",
    "cross_copy.planned_bytes",
    "cross_copy.external_relationship_count",
    "cross_copy.collision_remapped_parts",
    "cross_copy.gates",
)


def path_under(path: str, prefix: str) -> bool:
    return path == prefix or path.startswith(prefix + ".") or path.startswith(prefix + "[")


def value_at(value: Any, path: str) -> Any:
    tokens = re.findall(r"[^.\[\]]+|\[\d+\]", path)
    current = value
    for token in tokens:
        if token.startswith("["):
            if not isinstance(current, list):
                return MISSING
            index = int(token[1:-1])
            if index >= len(current):
                return MISSING
            current = current[index]
        else:
            if not isinstance(current, dict) or token not in current:
                return MISSING
            current = current[token]
    return current


def binding_value(value: Any) -> Any:
    return {"missing": True} if value is MISSING else value


def allowlist(path: pathlib.Path | None, failures: list[str]) -> dict[str, dict[str, dict[str, Any]]]:
    if path is None:
        return {}
    try:
        value = read_json(path)
    except (OSError, json.JSONDecodeError) as error:
        fail(failures, f"allowlist: invalid: {error}")
        return {}
    if (
        not isinstance(value, dict)
        or value.get("version") != 1
        or value.get("reviewed") is not True
        or value.get("arm") != "candidate"
    ):
        fail(failures, "allowlist: requires version 1, arm=candidate, and reviewed=true")
        return {}
    if not isinstance(value.get("reviewed_by"), str) or not value["reviewed_by"].strip():
        fail(failures, "allowlist: reviewed_by is required")
        return {}
    raw = value.get("paths")
    result: dict[str, dict[str, dict[str, Any]]] = {}
    if not isinstance(raw, dict):
        fail(failures, "allowlist: paths is missing")
        return {}
    for case, paths in raw.items():
        if not isinstance(case, str) or not isinstance(paths, dict):
            fail(failures, "allowlist: paths must map cases to path bindings")
            continue
        result[case] = {}
        for path_name, binding in paths.items():
            if (
                not isinstance(path_name, str)
                or not path_name
                or not isinstance(binding, dict)
                or set(binding) != {"baseline", "candidate"}
            ):
                fail(failures, f"allowlist: malformed binding {case}:{path_name}")
                continue
            if any(
                path_under(path_name, protected) or path_under(protected, path_name)
                for protected in PROTECTED_ALLOWLIST_PATHS
            ):
                fail(failures, f"allowlist: semantic/refusal path is protected: {case}:{path_name}")
                continue
            result[case][path_name] = binding
    return result


def projection_check(
    expected: dict[str, Any],
    actual: dict[str, Any],
    case: str,
    allowed: dict[str, dict[str, dict[str, Any]]],
) -> dict[str, Any]:
    differences = deep_diff(expected, actual)
    bindings = allowed.get(case, {})
    binding_failures: list[dict[str, Any]] = []
    verified_paths: set[str] = set()
    for path, binding in bindings.items():
        old_value = binding_value(value_at(expected, path))
        new_value = binding_value(value_at(actual, path))
        if old_value != binding["baseline"] or new_value != binding["candidate"]:
            binding_failures.append(
                {
                    "path": path,
                    "expected_baseline": binding["baseline"],
                    "actual_baseline": old_value,
                    "expected_candidate": binding["candidate"],
                    "actual_candidate": new_value,
                }
            )
        else:
            verified_paths.add(path)
    accepted: list[dict[str, Any]] = []
    unexpected: list[dict[str, Any]] = []
    for difference in differences:
        if any(path_under(difference["path"], path) for path in verified_paths):
            accepted.append(difference)
        else:
            unexpected.append(difference)
    return {
        "exact": not differences,
        "pass": not unexpected and not binding_failures,
        "differences": differences,
        "allowlisted_differences": accepted,
        "unexpected_differences": unexpected,
        "binding_failures": binding_failures,
    }


def percentile(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    return ordered[max(0, math.ceil(fraction * len(ordered)) - 1)]


def stats(values: list[float]) -> dict[str, float | int]:
    if not values:
        raise EvidenceError("cannot summarize an empty vector")
    return {
        "n": len(values),
        "p50": statistics.median(values),
        "mean": statistics.fmean(values),
        "p95": percentile(values, 0.95),
        "minimum": min(values),
        "maximum": max(values),
    }


def stable_seed(*parts: str) -> int:
    material = "|".join(parts).encode("utf-8")
    return BOOTSTRAP_SEED + int.from_bytes(hashlib.sha256(material).digest()[:8], "big")


def ratio(candidate: float, baseline: float) -> float | None:
    if baseline == 0:
        return 1.0 if candidate == 0 else None
    return candidate / baseline


def bootstrap_ratios(
    pairs: list[tuple[float, float]], statistic: str, seed: int
) -> list[float] | None:
    if not pairs or any(base == 0 for base, _ in pairs):
        return None
    rng = random.Random(seed)
    samples: list[float] = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        chosen = [pairs[rng.randrange(len(pairs))] for _ in pairs]
        base = stats([item[0] for item in chosen])[statistic]
        candidate = stats([item[1] for item in chosen])[statistic]
        value = ratio(float(candidate), float(base))
        if value is not None:
            samples.append(value)
    return [percentile(samples, 0.025), percentile(samples, 0.975)] if samples else None


def ratio_report(
    pairs: list[tuple[float, float]], label: str, failures: list[str]
) -> dict[str, Any]:
    if not pairs:
        raise EvidenceError(f"{label}: no paired process values")
    base_values = [pair[0] for pair in pairs]
    candidate_values = [pair[1] for pair in pairs]
    pairwise = [value for value in (ratio(candidate, baseline) for baseline, candidate in pairs) if value is not None]
    result: dict[str, Any] = {
        "unit": "paired_process",
        "baseline": stats(base_values),
        "candidate": stats(candidate_values),
        "pairwise_ratios": [ratio(candidate, baseline) for baseline, candidate in pairs],
        "pairwise_ratio_stats": stats(pairwise) if pairwise else {"n": 0, "undefined": len(pairs)},
        "statistics": {},
    }
    for statistic in ("p50", "p95", "mean"):
        base = float(result["baseline"][statistic])
        candidate = float(result["candidate"][statistic])
        relative = ratio(candidate, base)
        if relative is None:
            fail(failures, f"{label}: ratio denominator is zero for {statistic}")
        result["statistics"][statistic] = {
            "baseline": base,
            "candidate": candidate,
            "ratio": relative,
            "percent_change": None if relative is None else (relative - 1.0) * 100.0,
            "bootstrap_ratio_95": bootstrap_ratios(
                pairs, statistic, stable_seed(label, statistic)
            ),
            "bootstrap_seed": stable_seed(label, statistic),
            "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
        }
    return result


def process_statistic_ratios(
    process_pairs: list[dict[str, Any]], metric: str, failures: list[str], lane: str, case: str
) -> dict[str, Any]:
    """Compare each per-process p50, p95, and mean at the process unit."""
    result: dict[str, Any] = {}
    for statistic in ("p50", "p95", "mean"):
        pairs = [
            (
                float(stats(pair["baseline"]["metrics"][metric])[statistic]),
                float(stats(pair["candidate"]["metrics"][metric])[statistic]),
            )
            for pair in process_pairs
        ]
        result[statistic] = ratio_report(
            pairs, f"{lane}/{case}/{metric}/{statistic}", failures
        )
        result[statistic]["source_statistic"] = statistic
    return result


def process_record(
    entry: dict[str, Any],
    row: dict[str, Any],
    arm: str,
    builds: dict[str, dict[str, dict[str, Any]]],
    sources: dict[str, dict[str, Any]],
    oracle: dict[str, Any],
    allowed: dict[str, dict[str, dict[str, Any]]],
    failures: list[str],
) -> dict[str, Any] | None:
    label = f"{arm}/{row['lane']}/{row['case']}/repeat-{row['repeat']:02d}"
    try:
        report = read_json(entry["path"])
        projected = projection(report, row, arm, builds, sources)
        if row["case"] not in oracle:
            raise EvidenceError("oracle case is missing")
        check = projection_check(oracle[row["case"]], projected, row["case"], allowed if arm == "candidate" else {})
        if not check["pass"]:
            fail(failures, f"{label}: stable oracle mismatch")
        r = report["results"][0]
        x = r["source"]["pptx_cross_copy"]
        metrics = {name: [float(value) for value in x[name]] for name in TIMING_METRICS}
        allocation = r["operation_metrics"]["allocation"]
        allocation_shape = {
            name: {
                key: value
                for key, value in metric.items()
                if key != "values"
            }
            for name, metric in allocation.items()
            if name not in ("status", "scope") and isinstance(metric, dict)
        }
        allocation_values = (
            {
                name: [int(value) for value in metric["values"]]
                for name, metric in allocation.items()
                if name not in ("status", "scope") and isinstance(metric, dict)
            }
            if row["lane"] == "allocation"
            else {}
        )
        return {
            "row": row,
            "report": report,
            "projection": projected,
            "projection_check": check,
            "metrics": metrics,
            "allocation_status": allocation.get("status"),
            "allocation_scope": allocation.get("scope"),
            "allocation_shape": allocation_shape,
            "allocation_values": allocation_values,
        }
    except (OSError, json.JSONDecodeError, EvidenceError, KeyError, TypeError, ValueError) as error:
        fail(failures, f"{label}: {error}")
        return None


def compare_group(
    lane: str,
    case: str,
    entries: dict[tuple[str, str, int, str], dict[str, Any]],
    failures: list[str],
) -> dict[str, Any]:
    repeats = 6 if lane == "native" else 3
    process_pairs: list[dict[str, Any]] = []
    for repeat in range(repeats):
        key = lambda arm: (arm, lane, repeat, case)
        base = entries.get(key("baseline"))
        candidate = entries.get(key("candidate"))
        if base is None or candidate is None:
            fail(failures, f"{lane}/{case}/repeat-{repeat:02d}: missing arm")
            continue
        if base["allocation_shape"] != candidate["allocation_shape"]:
            fail(failures, f"{lane}/{case}/repeat-{repeat:02d}: allocation fields differ")
        if base["allocation_status"] != candidate["allocation_status"]:
            fail(failures, f"{lane}/{case}/repeat-{repeat:02d}: allocation status differs")
        if base["allocation_scope"] != candidate["allocation_scope"]:
            fail(failures, f"{lane}/{case}/repeat-{repeat:02d}: allocation scope differs")
        process_pairs.append({"repeat": repeat, "baseline": base, "candidate": candidate})
    if len(process_pairs) != repeats:
        return {
            "lane": lane,
            "case": case,
            "repeats": len(process_pairs),
            "timing": {},
            "allocation": {} if lane == "allocation" else None,
            "individual_5_percent_flags": [],
        }
    result: dict[str, Any] = {"lane": lane, "case": case, "repeats": len(process_pairs), "timing": {}, "individual_5_percent_flags": []}
    for metric in TIMING_METRICS:
        for statistic in ("p50", "p95", "mean"):
            pairs = [
                (
                    float(stats(pair["baseline"]["metrics"][metric])[statistic]),
                    float(stats(pair["candidate"]["metrics"][metric])[statistic]),
                )
                for pair in process_pairs
            ]
            for pair, values in zip(process_pairs, pairs):
                relative = ratio(values[1], values[0])
                if relative is None or abs((relative - 1.0) * 100.0) > PAIR_FLAG_PERCENT:
                    result["individual_5_percent_flags"].append(
                        {
                            "repeat": pair["repeat"],
                            "metric": metric,
                            "statistic": statistic,
                            "baseline": values[0],
                            "candidate": values[1],
                            "percent_change": None if relative is None else (relative - 1.0) * 100.0,
                        }
                    )
        result["timing"][metric] = process_statistic_ratios(
            process_pairs, metric, failures, lane, case
        )
    if lane == "allocation":
        names = sorted({name for pair in process_pairs for arm in ("baseline", "candidate") for name in pair[arm]["allocation_values"]})
        result["allocation"] = {"fields": names, "metrics": {}}
        for name in names:
            pairs = []
            for pair in process_pairs:
                before = pair["baseline"]["allocation_values"].get(name, [])
                after = pair["candidate"]["allocation_values"].get(name, [])
                if len(before) != 1 or len(after) != 1:
                    fail(failures, f"{lane}/{case}/{name}: expected one allocator sample")
                    continue
                pairs.append((float(before[0]), float(after[0])))
            result["allocation"]["metrics"][name] = ratio_report(
                pairs, f"{lane}/{case}/allocation/{name}", failures
            )
            for pair, values in zip(process_pairs, pairs):
                relative = ratio(values[1], values[0])
                if relative is None or abs((relative - 1.0) * 100.0) > PAIR_FLAG_PERCENT:
                    result["individual_5_percent_flags"].append(
                        {
                            "repeat": pair["repeat"],
                            "metric": f"allocation.{name}",
                            "statistic": "sample",
                            "baseline": values[0],
                            "candidate": values[1],
                            "percent_change": None if relative is None else (relative - 1.0) * 100.0,
                        }
                    )
    return result


def analyze(
    allowlist_path: pathlib.Path | None = None, cleanup_path: pathlib.Path | None = None
) -> dict[str, Any]:
    failures: list[str] = []
    oracle_path = P / "oracle.json"
    cases = cases_from_oracle(oracle_path, failures)
    oracle = read_json(oracle_path) if oracle_path.is_file() else {}
    packet = packet_bindings(failures)
    cleanup = cleanup_witness(cleanup_path, failures)
    frozen = {arm: frozen_input_hashes(arm, failures) for arm in ARMS}
    sources = {
        "baseline": source_manifest(P / "source.json", failures),
        "candidate": source_manifest(P / "candidate-source.json", failures),
    }
    builds = {
        "baseline": build_rows(P / "build.json", "baseline", failures, cleanup),
        "candidate": build_rows(P / "candidate-build.json", "candidate", failures, cleanup),
    }
    rows = matrix(cases)
    entries = receipts(rows, builds, frozen, failures)
    allowed = allowlist(allowlist_path, failures)
    for case in allowed:
        if case not in cases:
            fail(failures, f"allowlist: unknown oracle case {case}")
    projections: list[dict[str, Any]] = []
    for row in rows:
        key = (row["arm"], row["lane"], row["repeat"], row["case"])
        entry = entries.get(key)
        if entry is None:
            continue
        record = process_record(entry, row, row["arm"], builds, sources, oracle, allowed, failures)
        if record is not None:
            entries[key]["processed"] = record
            projections.append(
                {
                    "arm": row["arm"],
                    "lane": row["lane"],
                    "repeat": row["repeat"],
                    "case": row["case"],
                    "projection": record["projection_check"],
                }
            )
    groups = []
    for lane in LANES:
        for case in cases:
            group_entries = {
                key: value["processed"]
                for key, value in entries.items()
                if value.get("processed") is not None
            }
            groups.append(compare_group(lane, case, group_entries, failures))
    return {
        "status": "passed" if not failures else "failed",
        "matrix": {
            "native": {"paired_repeats_per_case": 6, "samples": 30, "warmups": 3},
            "allocation": {"repeats_per_arm_per_case": 3, "samples": 1, "warmups": 0},
            "arm_order": ["baseline,candidate", "candidate,baseline"],
            "processes": len(rows),
            "serial": True,
        },
        "custody": {
            "source_heads": {arm: sources[arm].get("head") for arm in ARMS},
            "source_file_counts": {arm: len(sources[arm].get("files", {})) for arm in ARMS},
            "binary_sha256": {
                arm: {lane: builds[arm].get(lane, {}).get("binary_sha256") for lane in LANES}
                for arm in ARMS
            },
            "packet_bindings": packet,
            "cleanup_witness": cleanup,
            "frozen_inputs": frozen,
        },
        "oracle": {
            "mode": "exact unless a reviewed explicit allowlist is supplied",
            "allowlist": allowed,
            "projections": projections,
        },
        "groups": groups,
        "failures": failures,
        "scope": "Matched process-level release comparison; bootstrap resamples paired processes, never individual timing samples. No causal speedup claim is made by this analyzer.",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", nargs="?", choices=("matrix", "analyze"), default="analyze")
    parser.add_argument("--allowlist", type=pathlib.Path)
    parser.add_argument("--cleanup-witness", type=pathlib.Path)
    parser.add_argument("--output", type=pathlib.Path, default=P / "analysis.json")
    args = parser.parse_args()
    if args.command == "matrix":
        failures: list[str] = []
        cases = cases_from_oracle(P / "oracle.json", failures)
        value = {"rows": matrix(cases), "failures": failures}
        write_json(args.output if args.output != P / "analysis.json" else P / "matrix.json", value)
        print(json.dumps(value, indent=2, sort_keys=True))
        return 1 if failures else 0
    value = analyze(args.allowlist, args.cleanup_witness)
    output = args.output if args.output.is_absolute() else P / args.output
    write_json(output, value)
    print(f"{value['status'].upper()} 0741 matched comparison ({len(value['failures'])} failures)")
    return 1 if value["failures"] else 0


if __name__ == "__main__":
    sys.exit(main())
