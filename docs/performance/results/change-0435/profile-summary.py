#!/usr/bin/env python3
"""Summarize the six validated 0435 whole-process profiles.

This is a postprocessor only.  ``verify.verify_matrix`` validates the formal
matrix, receipts, binaries, source custody, and profile artifact inventory
before this file reads the retained perf text.  The resulting counters and
stacks describe the complete fresh-process command, including setup, corpus
generation, warmups, timed calls, hashing, and oracle work.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
ROLES = ("before-buffered", "after-buffered", "after-streaming")
KINDS = ("stat", "record")
EVENTS = ("cycles:u", "instructions:u", "branches:u", "branch-misses:u", "L1-dcache-load-misses:u")
CHANGE = 435
MAX_TEXT_BYTES = 256 * 1024 * 1024
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
STAT_NUMBER = re.compile(r"^[0-9]+$")
LOST_RE = re.compile(r"^#\s*Total Lost Samples:\s*([0-9,]+)\s*$")
EVENT_HEADER_RE = re.compile(r"^#\s*Samples:\s+(.+?)\s+of event '([^']+)'\s*$")
EVENT_COUNT_RE = re.compile(r"^#\s*Event count \(approx\.\):\s*([0-9,]+)\s*$")
REPORT_ROW_RE = re.compile(r"^\s*([0-9]+(?:\.[0-9]+)?)%\s+\S+\s+\S+\s+(.+?)\s*$")
ADDR2LINE_TEXT = "could not read first record"


class ProfileSummaryError(ValueError):
    pass


def load_module(filename: str, name: str):
    spec = importlib.util.spec_from_file_location(name, ROOT / filename)
    if spec is None or spec.loader is None:
        raise ProfileSummaryError(f"cannot load {filename}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


verify = load_module("verify.py", "change0435_verify_for_profile_summary")


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode("utf-8")


def strict_load(path: Path) -> Any:
    def duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ProfileSummaryError(f"{path}: duplicate JSON key {key!r}")
            result[key] = value
        return result

    def reject_constant(value: str) -> Any:
        raise ProfileSummaryError(f"{path}: non-finite JSON value {value!r}")

    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=duplicate_pairs, parse_constant=reject_constant)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ProfileSummaryError(f"{path}: invalid JSON: {error}") from error


def read_artifact(receipt_path: Path, receipt: dict[str, Any], name: str) -> tuple[bytes, dict[str, Any], Path]:
    artifacts = receipt.get("artifacts")
    if not isinstance(artifacts, dict) or name not in artifacts:
        raise ProfileSummaryError(f"{receipt_path}: artifact {name!r} is missing")
    descriptor = artifacts[name]
    selected = verify.artifact_path(receipt_path, descriptor, f"{receipt_path}.artifacts.{name}")
    raw = verify.artifact_bytes(selected, f"{receipt_path}.artifacts.{name}")
    return raw, descriptor, selected


def stable_path(value: Any) -> str:
    if not isinstance(value, str) or not value:
        raise ProfileSummaryError("artifact path is missing")
    return value[:-3] if value.endswith(".gz") else value


def artifact_binding(name: str, descriptor: dict[str, Any]) -> dict[str, Any]:
    path = descriptor.get("path")
    size = descriptor.get("bytes")
    digest = descriptor.get("sha256")
    if not isinstance(path, str) or Path(path).is_absolute() or not isinstance(size, int) or size < 0 or not isinstance(digest, str) or HEX64.fullmatch(digest) is None:
        raise ProfileSummaryError(f"artifact {name}: malformed retained binding")
    return {"path": stable_path(path), "bytes": size, "sha256": digest.lower()}


def parse_count(value: str) -> int | None:
    value = value.strip().replace(",", "")
    return int(value) if STAT_NUMBER.fullmatch(value) else None


def parse_perf_stat(raw: bytes, label: str) -> dict[str, Any]:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        raise ProfileSummaryError(f"{label}: perf stat is not UTF-8: {error}") from error
    events: dict[str, dict[str, Any]] = {}
    lines = text.splitlines()
    for line in lines:
        fields = line.split(",")
        if len(fields) < 3:
            continue
        event = fields[2].strip()
        if event not in EVENTS:
            continue
        if event in events:
            raise ProfileSummaryError(f"{label}: duplicate event {event}")
        count = parse_count(fields[0])
        running = fields[4].strip() if len(fields) > 4 else ""
        if count is None:
            events[event] = {"status": "missing", "reason": fields[0].strip() or "perf did not publish a count", "running_percent": running or None}
        else:
            events[event] = {"status": "available", "count": count, "running_percent": running or None}
    for event in EVENTS:
        events.setdefault(event, {"status": "missing", "reason": "event line was absent from perf stat output"})
    def available(event: str) -> int | None:
        item = events[event]
        return item.get("count") if item.get("status") == "available" else None
    cycles, instructions = available("cycles:u"), available("instructions:u")
    branches, misses = available("branches:u"), available("branch-misses:u")
    ipc = {"status": "available", "value": instructions / cycles} if cycles and instructions is not None else {"status": "missing", "reason": "cycles:u and instructions:u were not both available"}
    branch_rate = {"status": "available", "percent": misses * 100.0 / branches} if branches and misses is not None else {"status": "missing", "reason": "branches:u and branch-misses:u were not both available"}
    return {
        "scope": "perf_stat_whole_process_including_setup_warmups_timed_calls_hashing_oracle",
        "events": events,
        "ipc": ipc,
        "branch_miss_rate": branch_rate,
        "l1_load_count": {"status": "missing", "reason": "L1 load denominator was not collected"},
        "llc_cache_misses": {"status": "missing", "reason": "LLC event was unavailable/not part of the frozen event set"},
        "raw_sha256": sha_bytes(raw),
    }


def parse_perf_report(raw: bytes, label: str) -> dict[str, Any]:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        raise ProfileSummaryError(f"{label}: perf report is not UTF-8: {error}") from error
    lost: int | None = None
    event: str | None = None
    event_count: int | None = None
    rows: list[dict[str, Any]] = []
    warnings = 0
    for line in text.splitlines():
        if ADDR2LINE_TEXT in line:
            warnings += 1
        if match := LOST_RE.match(line):
            lost = int(match.group(1).replace(",", ""))
        if match := EVENT_HEADER_RE.match(line):
            event = match.group(2)
        if match := EVENT_COUNT_RE.match(line):
            event_count = int(match.group(1).replace(",", ""))
        if match := REPORT_ROW_RE.match(line):
            overhead = float(match.group(1))
            symbol = match.group(2).strip()
            if not symbol.startswith("|") and not symbol.startswith("--"):
                rows.append({"overhead_percent": overhead, "symbol": symbol})
    return {
        "scope": "perf_report_no_children_whole_process_samples",
        "event": {"status": "available", "name": event} if event else {"status": "missing", "reason": "perf report event header was absent"},
        "event_count": {"status": "available", "count": event_count} if event_count is not None else {"status": "missing", "reason": "perf report event count was absent"},
        "lost_samples": {"status": "available", "count": lost} if lost is not None else {"status": "missing", "reason": "Total Lost Samples header was absent"},
        "addr2line_warnings": {"status": "available", "count": warnings, "scope": "symbolization_diagnostic_not_lost_sample_count"},
        "top_self_rows": rows[:20],
        "raw_sha256": sha_bytes(raw),
    }


def record_event_summary(raw: bytes, label: str) -> dict[str, Any]:
    """Bind the cycles sampling event without treating samples as counters."""

    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        raise ProfileSummaryError(f"{label}: perf report is not UTF-8: {error}") from error
    event = next((match.group(2) for line in text.splitlines() if (match := EVENT_HEADER_RE.match(line))), None)
    events = {
        "cycles:u": {"status": "available_sampled", "scope": "perf_record_sampling_event"} if event == "cycles:u" else {"status": "missing", "reason": "cycles:u sample event header was absent"},
        "instructions:u": {"status": "missing", "reason": "record lane sampled cycles:u only"},
        "branches:u": {"status": "missing", "reason": "record lane sampled cycles:u only"},
        "branch-misses:u": {"status": "missing", "reason": "record lane sampled cycles:u only"},
        "L1-dcache-load-misses:u": {"status": "missing", "reason": "record lane sampled cycles:u only"},
    }
    return {
        "scope": "perf_record_whole_process_including_setup_warmups_timed_calls_hashing_oracle",
        "events": events,
        "ipc": {"status": "missing", "reason": "instructions and cycles counts are not a paired perf-record counter measurement"},
        "branch_miss_rate": {"status": "missing", "reason": "branch counters were not collected by this record lane"},
        "l1_load_count": {"status": "missing", "reason": "L1 load denominator was not collected"},
        "llc_cache_misses": {"status": "missing", "reason": "LLC event was unavailable/not part of the frozen event set"},
    }


def summarize_profile(role: str, kind: str, row: dict[str, Any]) -> dict[str, Any]:
    receipt = row["receipt"]
    receipt_path = ROOT / "profiles" / role / kind / "receipt.json"
    bindings = {
        name: artifact_binding(name, descriptor)
        for name, descriptor in receipt["artifacts"].items()
    }
    output: dict[str, Any] = {
        "role": role,
        "kind": kind,
        "attempt": receipt.get("attempt"),
        "selector": receipt.get("selector"),
        "shape": receipt.get("shape"),
        "scope": receipt.get("scope"),
        "source_custody": {
            "source_manifest": receipt.get("source_manifest"),
            "ambient_source_manifest": receipt.get("ambient_source_manifest"),
            "source_before": receipt.get("source_before"),
            "source_after": receipt.get("source_after"),
            "source_unchanged": receipt.get("source_unchanged"),
        },
        "binary": receipt.get("binary"),
        "protocol_sha256": receipt.get("protocol_sha256"),
        "driver_sha256": receipt.get("driver_sha256"),
        "oracle_verifier_sha256": receipt.get("oracle_verifier_sha256"),
        "artifacts": bindings,
        "whole_process_rss": {
            **row["resource"],
            "scope": "gnu_time_v_verbose_whole_fresh_process_including_setup_warmups_timed_calls_hashing_oracle",
            "path": stable_path(row["resource"].get("path")),
        },
        "claims": [],
    }
    if kind == "stat":
        raw, _, _ = read_artifact(receipt_path, receipt, "perf_stat")
        output["pmu"] = parse_perf_stat(raw, f"{role}.stat.perf_stat")
    else:
        raw_report, _, _ = read_artifact(receipt_path, receipt, "perf_report")
        output["pmu"] = record_event_summary(raw_report, f"{role}.record.perf_report")
        output["record"] = parse_perf_report(raw_report, f"{role}.record.perf_report")
        output["record"]["perf_data"] = bindings.get("perf_data")
        output["record"]["perf_script"] = bindings.get("perf_script")
    return output


def derive(*, require_binaries: bool = False) -> dict[str, Any]:
    verified = verify.verify_matrix(require_binaries=require_binaries)
    rows = {(row["role"], row["kind"]): row for row in verified.get("profiles", [])}
    expected = {(role, kind) for role in ROLES for kind in KINDS}
    if set(rows) != expected:
        raise ProfileSummaryError(f"formal profile set differs: expected {sorted(expected)}, got {sorted(rows)}")
    profiles = [summarize_profile(role, kind, rows[(role, kind)]) for role in ROLES for kind in KINDS]
    return {
        "schema_version": 1,
        "change": CHANGE,
        "classification": "Whole-process ODT profiles retained for descriptive PMU and caller evidence; no causal hotspot, speedup, cache-miss, or 10x claim.",
        "protocol_sha256": verified["protocol_sha256"],
        "oracle": verified["oracle"],
        "matrix": {"formal_profiles": len(profiles), "required_profiles": 6, "roles": list(ROLES), "kinds": list(KINDS), "shape": "large", "source_matrix_reports": verified["matrix"]["reports"]},
        "profiles": profiles,
        "preparatory_evidence": {"path": "buffered-hypothesis.json", "scope": "retained separately from formal profiles and formal latency matrix"},
        "limitations": [
            "Every counter, IPC result, miss rate, sample count, and self row is whole-command evidence including setup, corpus generation, warmups, timed calls, output hashing, and untimed oracle work.",
            "A zero measured L1-dcache-load-misses count is retained as an observed counter value; it does not prove zero misses, and LLC counters are explicitly unavailable in this event set.",
            "perf-record samples are caller evidence, not a paired instructions/cycles counter measurement; IPC and branch miss rate are unavailable in record lanes.",
            "Total Lost Samples and addr2line warnings have separate scopes: lost samples concern perf data; addr2line warnings concern symbolization and are not lost-sample evidence.",
            "Self rows come from perf report --no-children and may contain unresolved/inlined symbols; they do not isolate the timed operation from setup or oracle work.",
            "No profile establishes causality, optimization benefit, latency improvement, or a 10x/cache-miss claim.",
        ],
        "claims": [],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "formal-profile-summary.json")
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--require-binaries", action="store_true")
    args = parser.parse_args()
    try:
        derived = derive(require_binaries=args.require_binaries)
        if args.check:
            if not args.output.is_file() or canonical(strict_load(args.output)) != canonical(derived):
                raise ProfileSummaryError(f"{args.output}: retained summary differs from fresh derivation")
        else:
            if args.output.exists():
                raise ProfileSummaryError(f"{args.output}: already exists; use --check")
            args.output.write_text(json.dumps(derived, indent=2, sort_keys=True, allow_nan=False) + "\n", encoding="utf-8")
    except (OSError, KeyError, TypeError, ValueError, AssertionError, verify.VerificationError, ProfileSummaryError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    print("VALID")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
