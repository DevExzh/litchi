#!/usr/bin/env python3
"""Fail-closed offline replay for the 0794 DOCX/XLSX control lane.

This script consumes only the frozen plan, capture receipts, and normal
``litchi-perf-baseline`` JSON reports.  It never builds or runs a binary.  The
six paired block medians are the only values used for the cross-format veto;
the per-sample vectors and report identities remain checked and retained.
"""

from __future__ import annotations

import hashlib
import json
import math
import random
import statistics
import sys
import argparse
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
PLAN_PATH = PACKET / "cross-plan.json"
NATIVE_DIR = PACKET / "cross-native"
QUALIFICATION_DIR = PACKET / "cross-qualification"
DEFAULT_OUTPUT = PACKET / "cross-analysis.json"

SCHEMA = "litchi.performance.0794.cross-format-plan.v1"
OUTPUT_SCHEMA = "litchi.performance.0794.cross-format-analysis.v1"
HARNESS_SCHEMA = 1
HARNESS_NAME = "litchi-perf-baseline"
CASES = (
    "docx_semantic_open",
    "docx_semantic_full_text",
    "xlsx_open_owned",
    "xlsx_full_cell_scan",
)
SEMANTIC_SHAPES = ("tiny", "large")
XLSX_SHAPES = ("tiny", "dense-wide")
ORDERS = (
    ("before", "after"),
    ("after", "before"),
    ("before", "after"),
    ("after", "before"),
    ("after", "before"),
    ("before", "after"),
)
LEGS = ("before", "after")
CROSS_BUILD_DIRS = {leg: PACKET / f"cross-build-{leg}" for leg in LEGS}
PRIMARY_SOURCE_PATHS = {
    "before": PACKET / "build-before/source.json",
    "after": PACKET / "build-after/source.json",
}
CROSS_CLEANUP_PATH = PACKET / "cross-cleanup.json"
NATIVE_SAMPLES = 30
NATIVE_WARMUP = 3
QUALIFICATION_SAMPLES = 1
QUALIFICATION_WARMUP = 0
BOOTSTRAP_SEED = 794080
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9_749
U64_MAX = (1 << 64) - 1


class ReplayError(RuntimeError):
    pass


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        fail(f"invalid JSON {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    if not isinstance(value, str) or len(value) != 64:
        return False
    return all(character in "0123456789abcdefABCDEF" for character in value)


def integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(
        isinstance(value, int)
        and not isinstance(value, bool)
        and (value > 0 if positive else value >= 0),
        f"{label} is not a {'positive' if positive else 'non-negative'} integer",
    )
    return value


def finite(value: Any, label: str) -> float:
    require(
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(float(value)),
        f"{label} is not finite",
    )
    return float(value)


def canonical(value: Any) -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        fail(f"cannot canonicalize JSON value: {error}")


def verify_artifact(value: Any, label: str, *, allow_missing: bool = False) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact record")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}.path is missing")
    size = integer(value.get("bytes"), f"{label}.bytes")
    digest = value.get("sha256")
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = Path(raw_path)
    if not path.is_file() or path.is_symlink():
        if allow_missing:
            return None
        fail(f"missing or symlinked {label}: {path}")
    require(path.stat().st_size == size, f"{label} size changed: {path}")
    require(sha256(path) == digest, f"{label} digest changed: {path}")
    return path


def verify_plan(plan: Any) -> dict[str, Any]:
    require(isinstance(plan, dict), "cross-plan.json is not an object")
    require(plan.get("schema") == SCHEMA, "cross-plan schema changed")
    require(plan.get("blocks") == len(ORDERS), "cross-plan block count changed")
    require(tuple(tuple(item) for item in plan.get("orders", ())) == ORDERS,
            "cross-plan order changed")
    require(tuple(plan.get("cases", ())) == CASES, "cross-plan cases changed")
    require(tuple(plan.get("semantic_shapes", ())) == SEMANTIC_SHAPES,
            "cross-plan DOCX shapes changed")
    require(tuple(plan.get("xlsx_shapes", ())) == XLSX_SHAPES,
            "cross-plan XLSX shapes changed")
    require(plan.get("samples") == NATIVE_SAMPLES, "cross-plan sample count changed")
    require(plan.get("warmup") == NATIVE_WARMUP, "cross-plan warmup changed")
    require(plan.get("qualification_samples") == QUALIFICATION_SAMPLES,
            "cross-plan qualification sample count changed")
    require(plan.get("cpu") == 12, "cross-plan CPU changed")
    bootstrap = plan.get("bootstrap")
    require(isinstance(bootstrap, dict), "cross-plan bootstrap is missing")
    require(bootstrap.get("seed") == BOOTSTRAP_SEED,
            "cross-plan bootstrap seed changed")
    require(bootstrap.get("resamples") == BOOTSTRAP_RESAMPLES,
            "cross-plan bootstrap count changed")
    require(tuple(bootstrap.get("zero_based_endpoints", ()))
            == (BOOTSTRAP_LOW_RANK, BOOTSTRAP_HIGH_RANK),
            "cross-plan bootstrap endpoints changed")
    gate = plan.get("additional_rejection_gate")
    require(isinstance(gate, dict)
            and gate.get("any_case_rejects") is True
            and gate.get("bootstrap95_low_above") == 1.0
            and gate.get("paired_process_p50_ratio_above") == 1.05,
            "cross-plan rejection gate changed")
    return plan


def verify_receipt_artifact_set(directory: Path, expected_children: int) -> list[dict[str, Any]]:
    complete = read_json(directory / "complete.json")
    require(isinstance(complete, dict), f"{directory}/complete.json is malformed")
    require(complete.get("children") == expected_children,
            f"{directory} child count changed")
    verify_artifact(complete.get("receipts"), f"{directory}/complete.receipts")
    verify_artifact(complete.get("source"), f"{directory}/complete.source")
    receipts = read_json(directory / "receipts.json")
    require(isinstance(receipts, list) and len(receipts) == expected_children,
            f"{directory}/receipts.json cardinality changed")
    return receipts


def verify_source_manifest(path: Path, label: str) -> dict[str, Any]:
    manifest = read_json(path)
    require(isinstance(manifest, dict), f"{label} source manifest is not an object")
    require(isinstance(manifest.get("revision"), str) and manifest["revision"],
            f"{label} source revision is missing")
    files = manifest.get("files")
    require(isinstance(files, dict) and files, f"{label} source file map is missing")
    for name, digest in files.items():
        require(isinstance(name, str) and name, f"{label} source file name is invalid")
        require(is_sha(digest), f"{label} source digest is invalid: {name}")
    return manifest


def verify_build_receipts() -> dict[str, dict[str, Any]]:
    """Verify each source-bound build and the shared frozen build inputs."""
    builds: dict[str, dict[str, Any]] = {}
    expected_command = [
        "cargo", "build", "--offline", "--locked", "--release", "--manifest-path",
        str(PACKET.parents[3] / "tools/perf-baseline/Cargo.toml"),
        "--bin", HARNESS_NAME,
    ]
    expected_frozen_names = {
        "cross-plan.json", "cross_build.py", "cross_capture.py", "cross-Cargo.lock",
    }
    for leg in LEGS:
        directory = CROSS_BUILD_DIRS[leg]
        receipt = read_json(directory / "receipt.json")
        require(isinstance(receipt, dict), f"{directory}/receipt.json is malformed")
        require(receipt.get("exit_code") == 0, f"{leg} cross build failed")
        require(receipt.get("command") == expected_command,
                f"{leg} cross build command changed")
        verify_artifact(receipt.get("log"), f"{leg} cross build log")
        source_path = verify_artifact(receipt.get("source"),
                                      f"{leg} cross build source")
        require(source_path is not None, f"{leg} cross build source is missing")
        source = verify_source_manifest(source_path, f"{leg} cross build")
        primary_path = PRIMARY_SOURCE_PATHS[leg]
        primary = verify_source_manifest(primary_path, f"{leg} primary build")
        require(source == primary,
                f"{leg} cross build source differs from primary build source")
        tools_path = verify_artifact(receipt.get("tools"), f"{leg} cross build tools")
        require(tools_path is not None, f"{leg} cross build tools are missing")
        tools = read_json(tools_path)
        require(isinstance(tools, dict) and tools,
                f"{leg} cross build tool inventory is malformed")
        for name, digest in tools.items():
            require(isinstance(name, str) and name and is_sha(digest),
                    f"{leg} cross build tool inventory entry is invalid")
        lock_path = verify_artifact(receipt.get("lock"), f"{leg} cross build lock")
        require(lock_path is not None, f"{leg} cross build lock is missing")
        expected_lock = PACKET / "cross-Cargo.lock"
        require(lock_path == expected_lock,
                f"{leg} cross build lock path is not the frozen cross lock")
        require(sha256(lock_path) == sha256(expected_lock),
                f"{leg} cross build lock digest changed")
        frozen_path = directory / "frozen-inputs.json"
        frozen = read_json(frozen_path)
        require(set(frozen) == expected_frozen_names,
                f"{leg} frozen build input set changed")
        for name, digest in frozen.items():
            require(is_sha(digest), f"{leg} frozen input digest is invalid: {name}")
        environment = receipt.get("environment")
        require(isinstance(environment, dict), f"{leg} build environment is missing")
        require(environment.get("CARGO_TARGET_DIR")
                == "/home/zhuhe/code/litchi-target-0794/cross",
                f"{leg} build target cache changed")
        require(environment.get("CARGO_BUILD_JOBS") == "2",
                f"{leg} build job count changed")
        require(environment.get("CARGO_INCREMENTAL") == "0",
                f"{leg} incremental setting changed")
        binary = receipt.get("binary")
        require(isinstance(binary, dict), f"{leg} cross build binary is missing")
        verify_artifact(binary, f"{leg} cross build binary", allow_missing=True)
        builds[leg] = {
            "receipt": receipt,
            "source_path": source_path,
            "source": source,
            "tools_path": tools_path,
            "tools": tools,
            "frozen_path": frozen_path,
            "frozen": frozen,
            "lock_path": lock_path,
            "binary": binary,
        }
    require(builds["before"]["tools"] == builds["after"]["tools"],
            "cross build tool inventories differ")
    require(builds["before"]["frozen"] == builds["after"]["frozen"],
            "cross build frozen inputs differ")
    return builds


def verify_cross_cleanup(builds: dict[str, dict[str, Any]]) -> dict[str, Any]:
    cleanup = read_json(CROSS_CLEANUP_PATH)
    require(isinstance(cleanup, dict), "cross-cleanup.json is malformed")
    require(cleanup.get("binaries_removed") is True,
            "cross cleanup does not certify binary removal")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list) and len(removed) == len(LEGS),
            "cross cleanup binary cardinality changed")
    expected = [builds[leg]["binary"] for leg in LEGS]
    require(all(any(item == identity for item in removed) for identity in expected),
            "cross cleanup does not contain both exact cross binary identities")
    require(len({canonical(item) for item in removed}) == len(LEGS),
            "cross cleanup contains duplicate binary identities")
    for identity in expected:
        require(not Path(identity["path"]).exists(),
                f"cleaned cross binary still exists: {identity['path']}")
    return cleanup


def option_value(command: list[str], option: str) -> str:
    try:
        index = command.index(option)
    except ValueError:
        fail(f"capture command lacks {option}")
    require(index + 1 < len(command), f"capture command truncates {option}")
    return command[index + 1]


def verify_command(receipt: dict[str, Any], *, samples: int, warmup: int,
                   leg: str, block: int | None,
                   expected_binary: dict[str, Any] | None = None) -> Path:
    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            "capture receipt command is malformed")
    require(command[:4] == ["/usr/bin/time", "-f", "%M", "-o"],
            "capture command does not use the declared /usr/bin/time wrapper")
    require("taskset" in command and "-c" in command,
            "capture command does not pin CPU with taskset")
    require(option_value(command, "-c") == "12", "capture CPU changed")
    require(option_value(command, "--samples") == str(samples),
            "capture sample count changed")
    require(option_value(command, "--warmup") == str(warmup),
            "capture warmup changed")
    require(option_value(command, "--case") == ",".join(CASES),
            "capture case selection changed")
    require(option_value(command, "--semantic-shape") == ",".join(SEMANTIC_SHAPES),
            "capture DOCX shape selection changed")
    require(option_value(command, "--xlsx-shape") == ",".join(XLSX_SHAPES),
            "capture XLSX shape selection changed")
    if block is not None:
        require(receipt.get("block") == block, "capture block ordinal changed")
    require(receipt.get("leg") == leg, "capture leg changed")
    require(receipt.get("exit_code") == 0, f"capture child failed for {leg}/{block}")
    report_path = verify_artifact(receipt.get("report"), "capture.report")
    verify_artifact(receipt.get("log"), "capture.log")
    verify_artifact(receipt.get("rss"), "capture.rss")
    binary = receipt.get("binary")
    require(isinstance(binary, dict), "capture binary identity is missing")
    if expected_binary is not None:
        require(binary == expected_binary,
                f"capture binary identity is not bound to the {leg} build")
    integer(binary.get("bytes"), "capture.binary.bytes", positive=True)
    require(is_sha(binary.get("sha256")), "capture.binary.sha256 is invalid")
    # The binary may have been removed after the capture. Its recorded digest
    # is still checked against the harness report below.
    verify_artifact(binary, "capture.binary", allow_missing=True)
    require(report_path is not None, "capture report path is missing")
    taskset_index = command.index("taskset")
    require(taskset_index + 3 < len(command)
            and command[taskset_index + 3] == binary["path"],
            "capture command binary differs from receipt identity")
    require(Path(option_value(command, "--json")) == report_path,
            "capture command report path differs from receipt")
    return report_path


def rust_midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + ((left % 2 + right % 2) // 2)


def nearest_rank(samples: list[int], percentile: int) -> int:
    index = min(((percentile * len(samples) + 99) // 100) - 1, len(samples) - 1)
    return samples[index]


def verify_statistics(value: Any, label: str, expected_count: int) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an object")
    require(value.get("unit") == "ns", f"{label}.unit changed")
    samples_value = value.get("samples")
    require(isinstance(samples_value, list) and len(samples_value) == expected_count,
            f"{label}.samples cardinality changed")
    samples = [integer(item, f"{label}.samples[{index}]")
               for index, item in enumerate(samples_value)]
    require(samples == sorted(samples), f"{label}.samples is not sorted")
    order = value.get("sample_order")
    require(isinstance(order, list) and len(order) == expected_count
            and sorted(order) == list(range(expected_count)),
            f"{label}.sample_order is not a permutation")
    require(value.get("min") == samples[0], f"{label}.min does not match samples")
    require(value.get("p50") == rust_midpoint(samples[(expected_count - 1) // 2],
                                               samples[expected_count // 2]),
            f"{label}.p50 does not match samples")
    require(value.get("p95") == nearest_rank(samples, 95),
            f"{label}.p95 does not match samples")
    require(value.get("p99") == nearest_rank(samples, 99),
            f"{label}.p99 does not match samples")
    require(value.get("max") == samples[-1], f"{label}.max does not match samples")
    finite(value.get("mean"), f"{label}.mean")
    finite(value.get("standard_deviation"), f"{label}.standard_deviation")
    confidence = value.get("confidence_interval_95")
    require(isinstance(confidence, dict), f"{label}.confidence_interval_95 is missing")
    finite(confidence.get("lower"), f"{label}.confidence_interval_95.lower")
    finite(confidence.get("upper"), f"{label}.confidence_interval_95.upper")
    return {"p50": value["p50"], "p95": value["p95"], "p99": value["p99"],
            "mean": value["mean"], "samples": samples}


def verify_binary_identity(report: dict[str, Any], receipt: dict[str, Any], label: str) -> None:
    identity = report.get("binary_identity")
    binary = receipt.get("binary")
    require(isinstance(identity, dict), f"{label}.binary_identity is missing")
    require(isinstance(binary, dict), f"{label}.receipt.binary is missing")
    require(isinstance(identity.get("path"), str) and identity["path"],
            f"{label}.binary_identity.path is invalid")
    require(is_sha(identity.get("binary_sha256")),
            f"{label}.binary_identity.binary_sha256 is invalid")
    integer(identity.get("binary_bytes"), f"{label}.binary_identity.binary_bytes", positive=True)
    require(identity.get("executable") is True, f"{label} binary is not executable")
    require(identity.get("profile") == "release", f"{label} binary profile changed")
    require(identity["binary_sha256"] == binary["sha256"],
            f"{label} report/receipt binary digest differs")
    require(identity["binary_bytes"] == binary["bytes"],
            f"{label} report/receipt binary size differs")


def verify_report(path: Path, receipt: dict[str, Any], *, samples: int, warmup: int,
                  expected_leg: str, block: int | None) -> tuple[dict[str, Any], dict[tuple[str, str], dict[str, Any]]]:
    report = read_json(path)
    label = str(path)
    require(isinstance(report, dict), f"{label} is not an object")
    require(report.get("schema_version") == HARNESS_SCHEMA, f"{label} schema changed")
    tool = report.get("tool")
    require(isinstance(tool, dict) and tool.get("name") == HARNESS_NAME,
            f"{label} harness identity changed")
    require(tool.get("instrumentation") == "none", f"{label} is instrumented")
    verify_binary_identity(report, receipt, label)
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}.configuration is missing")
    require(configuration.get("samples_per_case") == samples,
            f"{label} sample configuration changed")
    require(configuration.get("warmup_iterations_per_case") == warmup,
            f"{label} warmup configuration changed")
    require(tuple(configuration.get("cases", ())) == CASES,
            f"{label} case configuration changed")
    require(tuple(configuration.get("semantic_shapes", ())) == SEMANTIC_SHAPES,
            f"{label} DOCX shape configuration changed")
    require(tuple(configuration.get("xlsx_shapes", ())) == XLSX_SHAPES,
            f"{label} XLSX shape configuration changed")
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{label}.environment is missing")
    for field in ("rustc_version", "allocator", "logical_cpus_available"):
        require(field in environment, f"{label}.environment.{field} is missing")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == len(CASES) * 2,
            f"{label} result count changed")
    rows: dict[tuple[str, str], dict[str, Any]] = {}
    case_shapes: set[tuple[str, str]] = set()
    for index, row in enumerate(results):
        require(isinstance(row, dict), f"{label}.results[{index}] is not an object")
        case = row.get("case")
        require(case in CASES, f"{label}.results[{index}] has an unexpected case")
        corpus = row.get("corpus")
        require(isinstance(corpus, dict), f"{label}.results[{index}].corpus is missing")
        shape = corpus.get("shape")
        if case.startswith("docx_"):
            require(shape in SEMANTIC_SHAPES, f"{label} has an unexpected DOCX corpus shape")
            require(corpus.get("generator") == "litchi-docx-semantic-v1",
                    f"{label} DOCX corpus generator changed")
        else:
            require(shape in XLSX_SHAPES, f"{label} has an unexpected XLSX corpus shape")
            require(corpus.get("generator") == "litchi-xlsx-synthetic-v1",
                    f"{label} XLSX corpus generator changed")
            require(isinstance(corpus.get("xlsx"), dict), f"{label} XLSX manifest is missing")
        require(is_sha(corpus.get("archive_sha256")),
                f"{label} corpus archive digest is invalid")
        case_shapes.add((case, shape))
        key = (case, canonical(corpus))
        require(key not in rows, f"{label} has duplicate result key {case}/{shape}")
        stats = verify_statistics(row.get("elapsed_ns"), f"{label}.{case}.{shape}.elapsed_ns", samples)
        rows[key] = {"case": case, "shape": shape, "corpus": corpus,
                     "corpus_key": key[1], "stats": stats, "report": label,
                     "leg": expected_leg, "block": block,
                     "binary_sha256": report["binary_identity"]["binary_sha256"]}
    expected_case_shapes = {
        (case, shape)
        for case in CASES
        for shape in (SEMANTIC_SHAPES if case.startswith("docx_") else XLSX_SHAPES)
    }
    require(case_shapes == expected_case_shapes,
            f"{label} case/shape matrix changed")
    return report, rows


def capture_source_manifest(directory: Path) -> tuple[Path, dict[str, Any]]:
    complete = read_json(directory / "complete.json")
    path = verify_artifact(complete.get("source"), f"{directory}/complete.source")
    require(path is not None, f"{directory} source artifact is missing")
    return path, verify_source_manifest(path, f"{directory}")


def load_capture(plan: dict[str, Any]) -> tuple[
    list[dict[str, Any]],
    dict[tuple[str, str, int, str], dict[str, Any]],
    dict[str, str],
    dict[str, dict[str, Any]],
    dict[str, Any],
]:
    builds = verify_build_receipts()
    cleanup = verify_cross_cleanup(builds)
    native_receipts = verify_receipt_artifact_set(NATIVE_DIR, len(ORDERS) * 2)
    qualification_receipts = verify_receipt_artifact_set(QUALIFICATION_DIR, 1)
    native_rows: dict[tuple[str, str, int, str], dict[str, Any]] = {}
    native_source_path, native_source = capture_source_manifest(NATIVE_DIR)
    qualification_source_path, qualification_source = capture_source_manifest(QUALIFICATION_DIR)
    require(native_source == builds["after"]["source"],
            "native capture source differs from after build source")
    require(qualification_source == builds["before"]["source"],
            "qualification source differs from before build source")
    source_digests = {
        "build_before": sha256(builds["before"]["source_path"]),
        "build_after": sha256(builds["after"]["source_path"]),
        "native": sha256(native_source_path),
        "qualification": sha256(qualification_source_path),
    }
    require(source_digests["qualification"] == source_digests["build_before"],
            "qualification source digest differs from before build")
    require(source_digests["native"] == source_digests["build_after"],
            "native source digest differs from after build")
    report_records: list[dict[str, Any]] = []
    for index, receipt in enumerate(native_receipts):
        require(isinstance(receipt, dict), f"native receipt {index} is malformed")
        block = index // 2
        expected_leg = ORDERS[block][index % 2]
        path = verify_command(receipt, samples=NATIVE_SAMPLES, warmup=NATIVE_WARMUP,
                              leg=expected_leg, block=block,
                              expected_binary=builds[expected_leg]["binary"])
        report, rows = verify_report(path, receipt, samples=NATIVE_SAMPLES,
                                     warmup=NATIVE_WARMUP, expected_leg=expected_leg,
                                     block=block)
        report_records.append({"path": str(path), "sha256": sha256(path),
                               "leg": expected_leg, "block": block,
                               "binary_sha256": report["binary_identity"]["binary_sha256"]})
        for key, row in rows.items():
            tuple_key = (key[0], key[1], block, expected_leg)
            require(tuple_key not in native_rows, f"duplicate native row {tuple_key}")
            native_rows[tuple_key] = row
    qualification_receipt = qualification_receipts[0]
    require(isinstance(qualification_receipt, dict), "qualification receipt is malformed")
    qualification_path = verify_command(
        qualification_receipt, samples=QUALIFICATION_SAMPLES,
        warmup=QUALIFICATION_WARMUP, leg="before", block=0,
        expected_binary=builds["before"]["binary"],
    )
    qualification_report, qualification_rows = verify_report(
        qualification_path, qualification_receipt, samples=QUALIFICATION_SAMPLES,
        warmup=QUALIFICATION_WARMUP, expected_leg="before", block=0,
    )
    report_records.append({"path": str(qualification_path), "sha256": sha256(qualification_path),
                           "leg": "before", "qualification": True,
                           "binary_sha256": qualification_report["binary_identity"]["binary_sha256"]})
    # Qualification has one row per case/shape. Match it against the before
    # corpus identity from block zero; this catches a stale binary or changed
    # generator before the six-block ratios are read.
    for key, row in qualification_rows.items():
        native_key = (key[0], key[1], 0, "before")
        require(native_key in native_rows, f"qualification row has no native match: {key[0]}")
        require(row["corpus_key"] == native_rows[native_key]["corpus_key"],
                f"qualification corpus differs for {key[0]}/{row['shape']}")
    return report_records, native_rows, source_digests, builds, cleanup


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(values and all(math.isfinite(value) for value in values),
            "cannot bootstrap empty or non-finite ratios")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates = [
        statistics.median(values[rng.randrange(len(values))] for _ in values)
        for _ in range(BOOTSTRAP_RESAMPLES)
    ]
    estimates.sort()
    return {
        "seed": BOOTSTRAP_SEED,
        "resamples": BOOTSTRAP_RESAMPLES,
        "statistic": "median",
        "confidence": 0.95,
        "low_rank": BOOTSTRAP_LOW_RANK,
        "high_rank": BOOTSTRAP_HIGH_RANK,
        "ci_low": estimates[BOOTSTRAP_LOW_RANK],
        "ci_high": estimates[BOOTSTRAP_HIGH_RANK],
    }


def analyze(native_rows: dict[tuple[str, str, int, str], dict[str, Any]]) -> list[dict[str, Any]]:
    keys = sorted({(case, corpus_key) for case, corpus_key, _block, _leg in native_rows})
    output: list[dict[str, Any]] = []
    for case, corpus_key in keys:
        blocks: list[dict[str, Any]] = []
        ratios: list[float] = []
        corpus: dict[str, Any] | None = None
        for block in range(len(ORDERS)):
            before = native_rows[(case, corpus_key, block, "before")]
            after = native_rows[(case, corpus_key, block, "after")]
            require(before["corpus_key"] == after["corpus_key"],
                    f"before/after corpus differs for {case} block {block}")
            corpus = before["corpus"]
            left = before["stats"]["p50"]
            right = after["stats"]["p50"]
            require(left > 0, f"zero before p50 for {case} block {block}")
            ratio = right / left
            require(math.isfinite(ratio), f"non-finite ratio for {case} block {block}")
            ratios.append(ratio)
            blocks.append({"block": block, "before_p50_ns": left,
                           "after_p50_ns": right, "ratio": ratio,
                           "change_percent": (ratio - 1.0) * 100.0,
                           "before_report": before["report"],
                           "after_report": after["report"]})
        interval = bootstrap(ratios)
        median_ratio = statistics.median(ratios)
        rejected = median_ratio > 1.05 and interval["ci_low"] > 1.0
        require(corpus is not None, f"missing corpus for {case}")
        output.append({
            "case": case,
            "shape": corpus["shape"],
            "corpus": corpus,
            "blocks": blocks,
            "ratios": ratios,
            "ratio_median": median_ratio,
            "change_percent_median": (median_ratio - 1.0) * 100.0,
            "bootstrap": interval,
            "reject": rejected,
        })
    require(len(output) == 8, "cross-format analysis row count changed")
    return output


def make_result(plan: dict[str, Any]) -> dict[str, Any]:
    reports, native_rows, source_digests, builds, cleanup = load_capture(plan)
    rows = analyze(native_rows)
    rejected = [row for row in rows if row["reject"]]
    build_records = {
        leg: {
            "source_sha256": sha256(builds[leg]["source_path"]),
            "tools_sha256": sha256(builds[leg]["tools_path"]),
            "frozen_inputs_sha256": sha256(builds[leg]["frozen_path"]),
            "lock_sha256": sha256(builds[leg]["lock_path"]),
            "binary": builds[leg]["binary"],
        }
        for leg in LEGS
    }
    return {
        "schema": OUTPUT_SCHEMA,
        "plan_schema": plan["schema"],
        "counts": {"native_reports": len(reports) - 1, "qualification_reports": 1,
                    "native_samples": len(ORDERS) * 2 * len(CASES) * 2 * NATIVE_SAMPLES,
                    "qualification_samples": len(CASES) * 2 * QUALIFICATION_SAMPLES,
                    "analysis_rows": len(rows)},
        "source_artifacts": source_digests,
        "builds": build_records,
        "cross_cleanup": cleanup,
        "bootstrap": {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                       "statistic": "median", "confidence": 0.95,
                       "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK},
        "reports": reports,
        "rows": rows,
        "decision": {"rejected": bool(rejected),
                      "rejected_rows": [f"{row['case']}/{row['shape']}" for row in rejected],
                      "gate": "candidate p50 median > 1.05 and bootstrap lower bound > 1.0",
                      "benefit_requirement": "none; cross-format lane is a regression veto"},
        "verification": {
            "plan_checked": True,
            "build_receipts_checked": True,
            "cross_cleanup_checked": True,
            "source_custody_by_leg_checked": True,
            "shared_tools_and_frozen_inputs_checked": True,
            "all_receipts_checked": True,
            "all_report_sample_vectors_checked": True,
            "corpus_identity_checked": True,
            "semantic_runtime_gates_required": True,
            "normal_uninstrumented_reports_only": True,
            "no_profiler_or_allocator_claim": True,
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true",
                      help="write the replay result, refusing to overwrite it")
    mode.add_argument("--check", action="store_true",
                      help="recompute and compare without writing")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT,
                        help=f"analysis JSON path (default: {DEFAULT_OUTPUT})")
    args = parser.parse_args(argv)
    output = args.output
    plan = verify_plan(read_json(PLAN_PATH))
    result = make_result(plan)
    if args.check:
        require(output.is_file() and not output.is_symlink(),
                f"analysis output is missing for --check: {output}")
        require(read_json(output) == result,
                f"existing analysis output differs: {output}")
        print(json.dumps({"checked": str(output), "rows": len(result["rows"]),
                          "rejected": result["decision"]["rejected"]}, sort_keys=True))
        return 0
    if output.exists():
        fail(f"refusing to overwrite existing output: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"output": str(output), "rows": len(result["rows"]),
                      "rejected": result["decision"]["rejected"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ReplayError as error:
        print(f"cross-analysis: {error}", file=sys.stderr)
        raise SystemExit(2)
