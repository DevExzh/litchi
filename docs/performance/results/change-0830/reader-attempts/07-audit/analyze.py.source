#!/usr/bin/env python3
"""Deterministic, offline replay for the 0830 XLSX allocation packet.

This reader consumes only retained JSON reports, RSS receipts, Heaptrack's
decoded text stream, and ``heaptrack_print`` output.  It never starts Cargo,
the probe, a profiler, or any build tool.  The output is deliberately
descriptive: native timings are process-local measurements and Heaptrack
metrics count allocation calls and requested bytes.  No live, peak, or
per-edit allocation claim is derived here.  The retained analysis is written
as deterministic gzip (``analysis.json.gz``) so the full lossless parser
traces do not become an oversized Git blob; ``--check`` compares the exact
uncompressed JSON bytes after decompression.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import math
import random
import re
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable

import driver


PACKET = driver.P
ROOT = driver.ROOT
BASE_REVISION = driver.BASE
OWNER = driver.OWNER
ANALYSIS_SCHEMA = "litchi.performance.0830.xlsx-edit-profile-analysis.v1"
REPORT_SCHEMA = "litchi.performance.0830.xlsx-edit-profile.v1"
PLAN_SCHEMA = "litchi.performance.0830.plan.v1"
ARMS = ("direct", "wrapped", "fp")
REPORTS_EXPECTED = 23
SAMPLES_EXPECTED = 559
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_SEED = 830830
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9749
INPUT_BYTES = 8_435
INPUT_SHA256 = "d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4"
REFERENCE_BYTES = 8_521
REFERENCE_SHA256 = "0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68"
MARKER = "litchi-perf-0638-ordinary-save"
TARGET = {"sheet": "Munka1", "address": "A1"}
VERIFICATION_KEYS = frozenset(
    {
        "all_verified",
        "input_hash_verified",
        "reference_hash_verified",
        "output_hash_verified",
        "output_size_verified",
        "output_bytes_verified",
        "reopened",
        "marker_verified",
        "target_verified",
        "semantic_sha256_verified",
        "worksheet_count_verified",
        "stored_cell_count_verified",
    }
)
TIMING_SCOPE = (
    'Workbook::open is outside the clock; Workbook::edit, sheet("Munka1"), '
    'set("A1", marker), commit, patch non-empty check, commit.into_workbook '
    'adoption, and the old Workbook drop at assignment are inside the clock; '
    'serialization, owner drop, hashing, and complete stored-cell semantic '
    'readback are outside the clock'
)
HEX_RE = re.compile(r"^[0-9a-f]{64}$")
PRINT_CALLS_RE = re.compile(r"calls to allocation functions:\s*([0-9][0-9,]*)")


class ReplayError(RuntimeError):
    """Retained evidence is absent, malformed, or contradictory."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1 << 20), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def descriptor(path: Path, label: str, *, packet_bound: bool = False) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    if packet_bound:
        require(path.resolve().is_relative_to(PACKET.resolve()), f"{label} escaped packet")
    return {
        "path": relative(path),
        "bytes": path.stat().st_size,
        "sha256": sha256(path),
    }


def relative(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(PACKET.resolve()))
    except ValueError:
        return str(path)


def valid_sha(value: Any) -> bool:
    return isinstance(value, str) and HEX_RE.fullmatch(value) is not None


def positive_integer(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def nonnegative_integer(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def finite_positive(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)) and float(value) > 0,
            f"{label} is not finite and positive")


def expected_file_identity(value: Any, label: str, path: Path, size: int, digest: str) -> None:
    require(isinstance(value, dict), f"{label} identity is malformed")
    require(value.get("bytes") == size and value.get("sha256") == digest,
            f"{label} pinned identity changed")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label} path is missing")
    try:
        observed = Path(raw_path).resolve()
    except OSError as error:
        fail(f"{label} path cannot be resolved: {error}")
    require(observed == path.resolve(), f"{label} path changed: {raw_path}")


def expected_output_identity(value: Any, label: str = "output") -> None:
    require(value == {"bytes": REFERENCE_BYTES, "sha256": REFERENCE_SHA256},
            f"{label} identity changed")


def require_stage(name: str) -> dict[str, Any]:
    value = read_json(PACKET / f"{name}.json")
    require(value.get("status") == "pass", f"{name} stage is not terminal pass")
    require(value.get("inputs_sha256") == sha256(PACKET / "inputs.json"),
            f"{name} stage input custody changed")
    return {
        "path": relative(PACKET / f"{name}.json"),
        "bytes": (PACKET / f"{name}.json").stat().st_size,
        "sha256": sha256(PACKET / f"{name}.json"),
        "status": value["status"],
        "inputs_sha256": value["inputs_sha256"],
    }


def require_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema changed")
    require(plan.get("base") == BASE_REVISION and plan.get("owner") == OWNER,
            "plan base or owner changed")
    require(plan.get("cpu") == 12, "plan CPU changed")
    require(plan.get("input") == str(driver.INPUT), "plan input changed")
    require(plan.get("reference") == str(driver.REFERENCE), "plan reference changed")
    require(plan.get("qualification") == {
        "arms": ["direct", "wrapped", "fp"], "samples": 3, "warmup": 0,
    }, "qualification plan changed")
    expected_orders = [
        ["direct", "wrapped", "fp"],
        ["direct", "fp", "wrapped"],
        ["wrapped", "direct", "fp"],
        ["wrapped", "fp", "direct"],
        ["fp", "direct", "wrapped"],
        ["fp", "wrapped", "direct"],
    ]
    require(plan.get("native") == {
        "blocks": 6, "samples": 30, "warmup": 3, "order": expected_orders,
    }, "native plan changed")
    require(plan.get("heaptrack") == {
        "repeats": 2, "samples": 5, "warmup": 0, "arm": "fp",
        "requested_bytes_only": True, "profiled_latency_claim": False,
    }, "Heaptrack plan changed")
    require(plan.get("expected_reports") == REPORTS_EXPECTED
            and plan.get("expected_measured_samples") == SAMPLES_EXPECTED,
            "planned cardinality changed")
    require(plan.get("statistics") == {
        "within_process": "nearest-rank",
        "across_process": "midpoint median",
        "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED,
        "endpoints": [BOOTSTRAP_LOW_RANK, BOOTSTRAP_HIGH_RANK],
    }, "statistics plan changed")
    return plan


def require_commands() -> dict[str, Any]:
    commands = PACKET / "commands"
    require(commands.is_dir(), "command receipts are missing")
    expected = {"quality-fmt", "quality-check", "quality-test", "quality-clippy",
                "quality-doc", "build-ordinary", "nm-ordinary", "objdump-ordinary",
                "build-fp", "nm-fp", "objdump-fp"}
    expected.update(f"qualification-{arm}" for arm in ARMS)
    expected.update(f"native-{block:02}-{arm}" for block in range(6) for arm in ARMS)
    for repeat in range(2):
        expected.update({
            f"heaptrack-{repeat}", f"heaptrack-{repeat}-decode",
            f"heaptrack-{repeat}-print-whole", f"heaptrack-{repeat}-print-owner",
        })
    receipts: dict[str, Any] = {}
    for path in sorted(commands.glob("*.json")):
        if path.name.endswith(".started.json"):
            continue
        name = path.stem
        # Reader attempts are intentionally retained separately.  A failed
        # preflight remains useful evidence and must not make the workload
        # command gate look failed; the reader gate below selects the latest
        # successful attempt for the current parser source.
        if name not in expected:
            continue
        row = read_json(path)
        require(isinstance(row, dict), f"command receipt {name} is malformed")
        require(row.get("exit_code") == 0, f"command {name} did not exit zero")
        log_name = row.get("log_sha256")
        log = commands / f"{name}.log"
        require(valid_sha(log_name) and log.is_file(), f"command {name} log receipt missing")
        require(sha256(log) == log_name, f"command {name} log changed")
        receipts[name] = {
            "path": relative(path), "bytes": path.stat().st_size, "sha256": sha256(path),
            "exit_code": row["exit_code"], "log": descriptor(log, f"command {name} log", packet_bound=True),
        }
    require(expected <= set(receipts),
            f"command receipts incomplete: missing {sorted(expected - set(receipts))}")
    return {"expected": len(expected), "recorded": len(receipts), "all_exit_codes": "zero",
            "receipts": receipts}


def require_reader_preflight() -> dict[str, Any]:
    """Require a successful preflight bound to the current reader source.

    ``run_reader.py`` stores a source hash in each numbered attempt and
    delegates the actual command receipt to ``driver.run``.  Failed attempts
    are preserved, but only a successful attempt whose snapshot has exactly
    the current ``heaptrack_reader.py`` hash can authorize attribution.
    """
    reader = PACKET / "heaptrack_reader.py"
    current_hash = sha256(reader)
    attempts_root = PACKET / "reader-attempts"
    require(attempts_root.is_dir() and not attempts_root.is_symlink(),
            "reader-attempts directory is missing")
    candidates: list[tuple[int, Path, Path]] = []
    for folder in sorted(attempts_root.iterdir()):
        if not folder.is_dir() or folder.is_symlink():
            continue
        match = re.fullmatch(r"(\d+)-heaptrack_preflight", folder.name)
        if match is None:
            continue
        index = int(match.group(1))
        sources_path = folder / "sources.json"
        source_snapshot = folder / "heaptrack_reader.py.source"
        if not sources_path.is_file() or not source_snapshot.is_file():
            continue
        sources = read_json(sources_path)
        if not isinstance(sources, dict) or sources.get("heaptrack_reader.py") != current_hash:
            continue
        require(sha256(source_snapshot) == current_hash,
                f"reader attempt {index} source snapshot differs from current reader")
        receipt_path = PACKET / "commands" / f"reader-{index:02}-heaptrack_preflight.json"
        if not receipt_path.is_file():
            continue
        receipt = read_json(receipt_path)
        if receipt.get("exit_code") != 0:
            continue
        log = PACKET / "commands" / f"reader-{index:02}-heaptrack_preflight.log"
        require(log.is_file() and sha256(log) == receipt.get("log_sha256"),
                f"reader attempt {index} log receipt changed")
        candidates.append((index, folder, receipt_path))
    require(candidates, "no successful Heaptrack preflight matches current reader source")
    index, folder, receipt_path = max(candidates, key=lambda row: row[0])
    return {
        "attempt": index,
        "source": descriptor(folder / "heaptrack_reader.py.source",
                             f"reader attempt {index} source", packet_bound=True),
        "sources": descriptor(folder / "sources.json",
                              f"reader attempt {index} source manifest", packet_bound=True),
        "receipt": descriptor(receipt_path, f"reader attempt {index} receipt", packet_bound=True),
        "reader_sha256": current_hash,
    }


def validate_report(path: Path, *, mode: str, samples: int, warmup: int,
                    semantic_sha: str | None, label: str) -> tuple[dict[str, Any], str]:
    report = read_json(path)
    expected_top = {
        "schema", "tool", "base_revision", "mode", "timing_scope", "input", "reference",
        "output", "marker", "target", "target_value", "worksheet_count",
        "stored_cell_count", "semantic_sha256", "warmup", "samples_requested",
        "warmup_verified", "all_verified", "elapsed_ns", "samples",
    }
    require(set(report) == expected_top, f"{label} report fields changed")
    require(report["schema"] == REPORT_SCHEMA and report["tool"] == "xlsx-edit-profile-0830",
            f"{label} report schema/tool changed")
    require(report["base_revision"] == BASE_REVISION and report["mode"] == mode,
            f"{label} base or mode changed")
    require(report["timing_scope"] == TIMING_SCOPE, f"{label} timing scope changed")
    expected_file_identity(report["input"], f"{label} input", driver.INPUT,
                           INPUT_BYTES, INPUT_SHA256)
    expected_file_identity(report["reference"], f"{label} reference", driver.REFERENCE,
                           REFERENCE_BYTES, REFERENCE_SHA256)
    expected_output_identity(report["output"], f"{label} output")
    require(report["marker"] == MARKER and report["target"] == TARGET
            and report["target_value"] == MARKER, f"{label} edit target changed")
    require(report["worksheet_count"] == 1 and report["stored_cell_count"] == 8,
            f"{label} workbook shape changed")
    require(valid_sha(report["semantic_sha256"]), f"{label} semantic hash malformed")
    if semantic_sha is not None:
        require(report["semantic_sha256"] == semantic_sha, f"{label} semantic hash differs")
    require(report["warmup"] == warmup and report["samples_requested"] == samples,
            f"{label} sample/warmup contract changed")
    require(report["warmup_verified"] is True and report["all_verified"] is True,
            f"{label} report verification failed")
    elapsed = report["elapsed_ns"]
    require(set(elapsed) == {"unit", "samples", "sample_order"}
            and elapsed["unit"] == "ns"
            and elapsed["sample_order"] == list(range(samples)),
            f"{label} elapsed schema/order changed")
    require(isinstance(elapsed["samples"], list) and len(elapsed["samples"]) == samples,
            f"{label} elapsed sample count changed")
    require(isinstance(report["samples"], list) and len(report["samples"]) == samples,
            f"{label} sample record count changed")
    for index, sample in enumerate(report["samples"]):
        require(set(sample) == {"index", "elapsed_ns", "output", "verification"},
                f"{label} sample fields changed at {index}")
        require(sample["index"] == index and sample["elapsed_ns"] == elapsed["samples"][index],
                f"{label} sample order changed at {index}")
        positive_integer(sample["elapsed_ns"], f"{label} elapsed_ns[{index}]")
        expected_output_identity(sample["output"], f"{label} sample output {index}")
        verification = sample["verification"]
        require(set(verification) == VERIFICATION_KEYS,
                f"{label} verification keys changed at {index}")
        require(all(value is True for value in verification.values()),
                f"{label} verification failed at {index}")
    values = elapsed["samples"]
    return report, report["semantic_sha256"]


def load_qualification(semantic_sha: str | None) -> tuple[dict[str, Any], str]:
    rows = []
    for arm in ARMS:
        path = PACKET / "qualification" / f"{arm}.json"
        report, semantic_sha = validate_report(path, mode="direct" if arm == "direct" else "wrapped",
                                               samples=3, warmup=0, semantic_sha=semantic_sha,
                                               label=f"qualification/{arm}")
        rows.append({"arm": arm, "report": relative(path), "report_sha256": sha256(path),
                     "values": report["elapsed_ns"]["samples"]})
    return {"reports": 3, "samples": 9, "rows": rows,
            "purpose": "correctness and CLI-shape check; excluded from native summaries"}, semantic_sha


def parse_rss(path: Path, label: str) -> int:
    descriptor(path, label, packet_bound=True)
    text = path.read_text(encoding="utf-8")
    require(re.fullmatch(r"\s*[1-9][0-9]*\s*", text) is not None,
            f"{label} is not one positive RSS value")
    return int(text.strip())


def load_native(semantic_sha: str | None) -> tuple[dict[str, Any], str]:
    rows: list[dict[str, Any]] = []
    expected_modes = {"direct": "direct", "wrapped": "wrapped", "fp": "wrapped"}
    for block in range(6):
        for arm in ARMS:
            path = PACKET / "native" / f"native-{block:02}-{arm}.json"
            report, semantic_sha = validate_report(
                path, mode=expected_modes[arm], samples=30, warmup=3,
                semantic_sha=semantic_sha, label=f"native/{block:02}/{arm}")
            rss_path = path.with_suffix(".rss")
            rss = parse_rss(rss_path, f"native/{block:02}/{arm} RSS")
            values = report["elapsed_ns"]["samples"]
            stats = distribution(values)
            rows.append({
                "block": block, "arm": arm, "report": relative(path),
                "report_sha256": sha256(path), "values": values,
                "stats": stats, "rss_kib": rss,
                "rss_artifact": descriptor(rss_path, f"native/{block:02}/{arm} RSS", packet_bound=True),
            })
    require(len(rows) == 18 and sum(len(row["values"]) for row in rows) == 540,
            "native cardinality changed")
    lookup = {(row["block"], row["arm"]): row for row in rows}
    arms: dict[str, Any] = {}
    spread_flags: list[dict[str, Any]] = []
    tail_flags: list[dict[str, Any]] = []
    metrics_names = ("p50", "p95", "p99", "mean", "rss_kib")
    for arm in ARMS:
        process_rows = [lookup[(block, arm)] for block in range(6)]
        metrics: dict[str, Any] = {}
        for metric in metrics_names:
            values = [row["rss_kib"] if metric == "rss_kib" else row["stats"][metric]
                      for row in process_rows]
            ratio = spread(values)
            metrics[metric] = {
                "values": values,
                "midpoint_median": statistics.median(values),
                "spread_ratio": ratio,
                "spread_flag": ratio > 1.05,
            }
            if ratio > 1.05:
                spread_flags.append({"arm": arm, "metric": metric, "spread_ratio": ratio,
                                     "descriptive_only": True})
        tail_ratio = metrics["p99"]["midpoint_median"] / metrics["p50"]["midpoint_median"]
        tail_flag = tail_ratio > 1.05
        if tail_flag:
            tail_flags.append({"arm": arm, "p99_to_p50_ratio": tail_ratio,
                               "descriptive_only": True})
        arms[arm] = {
            "blocks": 6, "samples_per_report": 30, "metrics": metrics,
            "p99_to_p50_ratio": tail_ratio, "tail_flag": tail_flag,
            "timing_claim": "descriptive native process timing only",
            "rss_source": "/usr/bin/time -f %M; maximum resident set size in KiB",
            "reports": [{"block": row["block"], "report": row["report"],
                         "report_sha256": row["report_sha256"], "stats": row["stats"],
                         "rss_kib": row["rss_kib"], "rss_artifact": row["rss_artifact"]}
                        for row in process_rows],
        }
    pairs: dict[str, Any] = {}
    for name, numerator, denominator in (("wrapped/direct", "wrapped", "direct"),
                                         ("fp/wrapped", "fp", "wrapped")):
        by_block = []
        values = []
        for block in range(6):
            before = lookup[(block, denominator)]["stats"]["p50"]
            after = lookup[(block, numerator)]["stats"]["p50"]
            ratio = after / before
            values.append(ratio)
            by_block.append({"block": block, "numerator": numerator, "denominator": denominator,
                             "before_p50_ns": before, "after_p50_ns": after, "ratio": ratio})
        pairs[name] = {
            "numerator": numerator, "denominator": denominator,
            "by_block": by_block, "ratio_values": values,
            "midpoint_median": statistics.median(values),
            "bootstrap": bootstrap(values),
            "interpretation": "paired diagnostic ratio; no shipping latency or speedup claim",
        }
    return {
        "reports": 18, "samples": 540, "rows": rows, "arms": arms,
        "paired_ratios": pairs, "spread_flags": spread_flags, "tail_flags": tail_flags,
        "quantile_definition": "nearest-rank p50/p95/p99 within each process; midpoint median over six processes",
        "raw_report_quantile_definition": "nearest-rank p50/p95/p99 and arithmetic mean",
        "spread_threshold": 1.05, "timing_claim": "descriptive native process timing only",
    }, semantic_sha


def distribution(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "empty timing vector")
    for index, value in enumerate(vector):
        finite_positive(value, f"timing[{index}]")
    return {
        "count": len(vector), "p50": nearest_rank(vector, 0.50),
        "p95": nearest_rank(vector, 0.95), "p99": nearest_rank(vector, 0.99),
        "mean": statistics.fmean(vector), "min": min(vector), "max": max(vector),
        "values": vector,
    }


def nearest_rank(values: Iterable[int | float], quantile: float) -> int | float:
    ordered = sorted(values)
    require(ordered and 0 < quantile <= 1, "invalid nearest-rank request")
    return ordered[max(1, math.ceil(len(ordered) * quantile)) - 1]


def spread(values: Iterable[int | float]) -> float:
    vector = list(values)
    require(vector and all(float(value) > 0 for value in vector), "invalid spread vector")
    return max(vector) / min(vector)


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(values and all(math.isfinite(value) and value > 0 for value in values),
            "invalid bootstrap vector")
    rng = random.Random(BOOTSTRAP_SEED)
    draws = sorted(statistics.median(rng.choice(values) for _ in values)
                   for _ in range(BOOTSTRAP_RESAMPLES))
    return {
        "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
        "statistic": "midpoint median", "low_rank": BOOTSTRAP_LOW_RANK,
        "high_rank": BOOTSTRAP_HIGH_RANK,
        "estimate": statistics.median(values),
        "ci_low": draws[BOOTSTRAP_LOW_RANK],
        "ci_high": draws[BOOTSTRAP_HIGH_RANK],
    }


def metric(value: Any, label: str) -> dict[str, int]:
    require(isinstance(value, dict), f"{label} metric is malformed")
    require(set(value) == {"calls", "requested_bytes", "owner_calls", "owner_requested_bytes"},
            f"{label} metric fields changed")
    for key, item in value.items():
        nonnegative_integer(item, f"{label}.{key}")
    return dict(value)


def parse_print_calls(path: Path, label: str) -> int:
    text = path.read_text(encoding="utf-8")
    matches = PRINT_CALLS_RE.findall(text)
    require(len(matches) == 1, f"{label} has {len(matches)} allocation-call totals")
    return int(matches[0].replace(",", ""))


def binary_receipt(arm: str) -> dict[str, Any]:
    value = read_json(PACKET / f"binary-{arm}.json")
    require(set(value) == {"path", "bytes", "sha256"},
            f"{arm} binary descriptor fields changed")
    require(isinstance(value["path"], str) and Path(value["path"]).is_absolute(),
            f"{arm} binary path is not absolute")
    positive_integer(value["bytes"], f"{arm} binary bytes")
    require(valid_sha(value["sha256"]), f"{arm} binary hash is malformed")
    return {"path": value["path"], "bytes": value["bytes"], "sha256": value["sha256"]}


def cleanup_witness() -> dict[str, Any] | None:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None
    value = read_json(path)
    require(value.get("status") == "pass"
            and value.get("target") == str(driver.TARGET),
            "cleanup witness status or target changed")
    nonnegative_integer(value.get("removed_files"), "cleanup removed_files")
    nonnegative_integer(value.get("removed_bytes"), "cleanup removed_bytes")
    binaries = value.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"ordinary", "fp"},
            "cleanup binary witness cardinality changed")
    for arm in ("ordinary", "fp"):
        expected = binary_receipt(arm)
        observed = binaries[arm]
        require(observed == expected, f"cleanup {arm} binary witness differs from receipt")
        require(not Path(observed["path"]).exists(),
                f"cleanup witness still has removed {arm} binary")
    return {
        "path": relative(path), "bytes": path.stat().st_size, "sha256": sha256(path),
        "status": value["status"],
    }


def load_binary_fp() -> dict[str, Any]:
    value = binary_receipt("fp")
    path = Path(value["path"])
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == value["bytes"] and sha256(path) == value["sha256"],
                "fp binary identity changed")
    else:
        # Offline replay is also run after root removes its private target.
        # The cleanup witness must name exactly the retained build receipt;
        # no substitute binary or unbound path is accepted.
        cleanup = cleanup_witness()
        require(cleanup is not None, "fp binary is missing without cleanup witness")
    # Keep the descriptor identical before and after cleanup.  In particular,
    # do not expose a live/removed mode in analysis.json.gz, so deterministic
    # replay remains byte-for-byte stable across target cleanup.
    return value


def validate_heaptrack_report(path: Path, semantic_sha: str | None, label: str) -> tuple[dict[str, Any], str]:
    report, semantic_sha = validate_report(path, mode="wrapped", samples=5, warmup=0,
                                           semantic_sha=semantic_sha, label=label)
    return report, semantic_sha


def load_heaptrack(semantic_sha: str | None) -> tuple[dict[str, Any], str]:
    # Import lazily: the analysis module remains importable while a separate
    # agent is still editing the packet-local parser, and no parser executes
    # until a completed capture is explicitly replayed.
    from heaptrack_reader import parse as parse_heaptrack

    binary = load_binary_fp()
    repeats = []
    for repeat in range(2):
        folder = PACKET / f"heaptrack-{repeat}"
        report, semantic_sha = validate_heaptrack_report(folder / "report.json", semantic_sha,
                                                         f"heaptrack/{repeat}/report")
        trace_zst = folder / "trace.zst"
        trace_txt = folder / "trace.txt"
        print_whole = PACKET / "commands" / f"heaptrack-{repeat}-print-whole.log"
        print_owner = PACKET / "commands" / f"heaptrack-{repeat}-print-owner.log"
        trace_zst_desc = descriptor(trace_zst, f"heaptrack/{repeat} compressed trace", packet_bound=True)
        trace_txt_desc = descriptor(trace_txt, f"heaptrack/{repeat} decoded trace", packet_bound=True)
        whole_desc = descriptor(print_whole, f"heaptrack/{repeat} whole print", packet_bound=True)
        owner_desc = descriptor(print_owner, f"heaptrack/{repeat} owner print", packet_bound=True)
        trace_data = trace_txt.read_bytes()
        parsed = parse_heaptrack(trace_data, OWNER, binary["path"])
        require(isinstance(parsed, dict), f"heaptrack/{repeat} parser result is malformed")
        require(parsed.get("owner") == OWNER and parsed.get("binary_path") == binary["path"],
                f"heaptrack/{repeat} parser identity changed")
        whole = metric(parsed.get("whole"), f"heaptrack/{repeat} whole")
        owner_metric = metric(parsed.get("owner_attribution"), f"heaptrack/{repeat} owner")
        official_calls = parse_print_calls(print_whole, f"heaptrack/{repeat} whole print")
        require(official_calls == whole["calls"],
                f"heaptrack/{repeat} official whole calls disagree with parser")
        require(whole["owner_calls"] == owner_metric["calls"]
                and whole["owner_requested_bytes"] == owner_metric["requested_bytes"],
                f"heaptrack/{repeat} owner metric is not a whole-metric projection")
        diagnostics = parsed.get("diagnostics")
        require(isinstance(diagnostics, dict), f"heaptrack/{repeat} parser diagnostics missing")
        wrong_dso = diagnostics.get("owner_symbol_in_wrong_dso")
        require(isinstance(wrong_dso, dict)
                and wrong_dso.get("allocation_calls") == 0
                and wrong_dso.get("requested_bytes") == 0,
                f"heaptrack/{repeat} broad owner filter found a wrong-DSO owner symbol")
        require(isinstance(parsed.get("by_allocation_size"), dict),
                f"heaptrack/{repeat} by-size attribution missing")
        for key in ("leaf", "full_leaf", "caller_trace"):
            require(isinstance(parsed.get(key), dict)
                    and isinstance(parsed[key].get("rows"), list),
                    f"heaptrack/{repeat} {key} attribution missing")
        owner_print_text = print_owner.read_text(encoding="utf-8")
        owner_print_matches = PRINT_CALLS_RE.findall(owner_print_text)
        # heaptrack_print versions differ: some filtered reports retain only
        # the stack rows and omit the aggregate line.  When the summary is
        # present, compare its broad substring-filter count with the strict
        # parser's exact owner count; otherwise retain an explicit unsupported
        # marker rather than inventing a count.
        if owner_print_matches:
            require(len(owner_print_matches) == 1,
                    f"heaptrack/{repeat} owner print has multiple call totals")
            owner_print_calls: int | None = int(owner_print_matches[0].replace(",", ""))
            if owner_print_calls == owner_metric["calls"]:
                owner_print_status = "summary-compared"
                owner_print_explanation = (
                    "heaptrack_print owner-filter summary equals the strict exact-DSO parser count"
                )
            else:
                # Heaptrack versions can retain the global summary while
                # filtering only the rendered stack rows.  Preserve that
                # observed number and explain why it is not an owner count;
                # never force a false equality.
                owner_print_status = "summary-present-but-global-or-different-filter-semantics"
                owner_print_explanation = (
                    "heaptrack_print emitted a summary count different from the strict owner count; "
                    "the printed total is retained as a filter sanity observation, not an owner aggregate"
                )
        else:
            owner_print_calls = None
            owner_print_status = "summary-not-emitted-by-heaptrack_print"
            owner_print_explanation = (
                "heaptrack_print did not emit an allocation-call summary for the owner-filtered view"
            )
        owner_substring_present = OWNER in owner_print_text
        require(owner_substring_present,
                f"heaptrack/{repeat} owner-filter output lacks the requested owner substring")
        repeats.append({
            "repeat": repeat,
            "report": {"path": relative(folder / "report.json"),
                       "bytes": (folder / "report.json").stat().st_size,
                       "sha256": sha256(folder / "report.json")},
            "trace_zst": trace_zst_desc, "trace_txt": trace_txt_desc,
            "heaptrack_print_whole": whole_desc, "heaptrack_print_owner": owner_desc,
            "official_whole_allocation_calls": official_calls,
            "official_owner_filter_allocation_calls": owner_print_calls,
            "official_owner_filter_status": owner_print_status,
            "official_owner_filter_explanation": owner_print_explanation,
            "official_owner_filter_owner_substring_present": owner_substring_present,
            "parser": parsed,
            "metrics": {
                "whole": {"calls": whole["calls"], "requested_bytes": whole["requested_bytes"]},
                "owner": {"calls": owner_metric["calls"],
                          "requested_bytes": owner_metric["requested_bytes"]},
            },
            "attribution_scope": "allocation calls and requested bytes only; no per-edit peak or lifetime metric",
        })
    require(len(repeats) == 2, "Heaptrack repeat cardinality changed")
    return {
        "status": "available", "reports": 2, "samples": 10,
        "binary": binary, "repeats": repeats,
        "reader": "heaptrack_reader.parse interpreted Heaptrack 1.5 format 3",
        "official_cross_check": "heaptrack_print whole allocation-call total equals parser whole.calls for every repeat",
        "allocation_counter_caveat": "Heaptrack malloc interposition is not assumed equal to native Rust allocator counters",
        "profiled_latency_claim": False,
        "peak_or_lifetime_claim": False,
    }, semantic_sha


def build_analysis() -> dict[str, Any]:
    # The packet must have completed all root-owned stages before this offline
    # replay is allowed to interpret any measurements.
    plan = require_plan()
    stages = {name: require_stage(name) for name in ("prepare", "quality", "build", "capture")}
    commands = require_commands()
    reader_preflight = require_reader_preflight()
    inputs = descriptor(PACKET / "inputs.json", "inputs manifest", packet_bound=True)
    source = descriptor(PACKET / "analyze.py", "analysis reader", packet_bound=True)
    semantic_sha: str | None = None
    qualification, semantic_sha = load_qualification(semantic_sha)
    native, semantic_sha = load_native(semantic_sha)
    heaptrack, semantic_sha = load_heaptrack(semantic_sha)
    require(qualification["reports"] + native["reports"] + heaptrack["reports"] == REPORTS_EXPECTED,
            "total report cardinality changed")
    require(qualification["samples"] + native["samples"] + heaptrack["samples"] == SAMPLES_EXPECTED,
            "total measured sample cardinality changed")
    return {
        "schema": ANALYSIS_SCHEMA,
        "status": "pass",
        "base_revision": BASE_REVISION,
        "owner": OWNER,
        "plan": {"path": relative(PACKET / "plan.json"),
                 "bytes": (PACKET / "plan.json").stat().st_size,
                 "sha256": sha256(PACKET / "plan.json"),
                 "schema": plan["schema"]},
        "inputs": inputs,
        "analysis_reader": source,
        "stages": stages,
        "commands": {"expected_receipts": commands["expected"],
                     "recorded_receipts": commands["recorded"],
                     "all_exit_codes": commands["all_exit_codes"]},
        "reader_preflight": reader_preflight,
        "input": {"path": str(driver.INPUT), "bytes": INPUT_BYTES, "sha256": INPUT_SHA256},
        "reference": {"path": str(driver.REFERENCE), "bytes": REFERENCE_BYTES,
                      "sha256": REFERENCE_SHA256},
        "target": {**TARGET, "marker": MARKER, "worksheet_count": 1,
                   "stored_cell_count": 8, "semantic_sha256": semantic_sha},
        "counts": {
            "qualification_reports": qualification["reports"],
            "qualification_samples": qualification["samples"],
            "native_reports": native["reports"], "native_samples": native["samples"],
            "heaptrack_reports": heaptrack["reports"], "heaptrack_samples": heaptrack["samples"],
            "reports": REPORTS_EXPECTED, "samples": SAMPLES_EXPECTED,
            "planned_reports": REPORTS_EXPECTED, "planned_samples": SAMPLES_EXPECTED,
        },
        "qualification": qualification,
        "native": native,
        "heaptrack": heaptrack,
        "claims": {
            "native": "descriptive process timing and RSS evidence only",
            "heaptrack": "allocation calls and requested bytes only",
            "historical_comparison": False,
            "shipping_latency_or_speedup": False,
            "per_edit_peak_or_lifetime": False,
            "allocator_counter_equivalence": False,
        },
    }


def encoded(value: dict[str, Any]) -> str:
    return json.dumps(value, indent=2, sort_keys=True) + "\n"


def write_or_check(value: dict[str, Any], check: bool) -> None:
    path = PACKET / "analysis.json.gz"
    text = encoded(value).encode("utf-8")
    compressed = gzip.compress(text, mtime=0)
    if check:
        require(path.is_file() and not path.is_symlink(), "analysis.json.gz is missing")
        try:
            observed = gzip.decompress(path.read_bytes())
        except (OSError, EOFError) as error:
            fail(f"analysis.json.gz is not valid gzip: {error}")
        require(observed == text, "analysis.json.gz does not replay deterministically")
        return
    # Exclusive creation is intentional: a retained result is never silently
    # replaced by a later replay.
    try:
        with path.open("xb") as stream:
            stream.write(compressed)
    except FileExistsError:
        fail("refusing to overwrite retained analysis.json.gz")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true",
                        help="replay and compare the retained analysis.json.gz")
    args = parser.parse_args(argv)
    try:
        write_or_check(build_analysis(), args.check)
    except (ReplayError, AssertionError, OSError, UnicodeError, ValueError,
            KeyError, TypeError, IndexError) as error:
        print(f"0830 analysis failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps({"status": "pass", "reports": REPORTS_EXPECTED,
                      "samples": SAMPLES_EXPECTED}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
