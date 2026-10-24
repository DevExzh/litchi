#!/usr/bin/env python3
"""Validate one DrawingML paired raw-probe output directory.

This validator is intentionally stricter than a JSON-shape check.  It binds
every row to the lane, source role, source revision, fixture, and allocator
mode, checks the fixed ABBA order, and refuses to turn an unavailable adapter
into a performance claim.  Semantic fidelity and replay checks belong in the
lane adapter; when that adapter is absent, its rows must say so explicitly.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any, Mapping


SCHEMA = "litchi.drawingml.raw.probe.v1"
SUMMARY_SCHEMA = "litchi.drawingml.raw.validation.v1"
LANES = (
    "drawingml-chart-raw",
    "drawingml-diagram-raw",
    "drawingml-text-media-raw",
    "drawingml-theme-ink",
    "drawingml-geometry-color-formatcode2-transform",
)
OPERATIONS = (
    "entry_read",
    "read_shared_when_supported",
    "noop_commit",
    "changed_commit",
    "forward_replay",
    "inverse_replay",
    "stale_replay",
    "vec_write",
    "write_to_bounded_sink",
    "preflight_refusal",
    "malformed_refusal",
)
ABBA = ("baseline", "candidate", "candidate", "baseline")
MINIMUM_MS = 150
TARGET_MS = 500
WARMUPS = 3
SAMPLES = 20
BLOCKS = 5


class ValidationError(ValueError):
    """A fail-closed probe output error."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise ValidationError(message)


def _sha256(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink() and path.is_file(), f"not a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _load(path: Path) -> Mapping[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"cannot read probe output {path}: {error}") from error
    _expect(isinstance(value, dict), f"probe output root is not an object: {path}")
    return value


def _ordered_outputs(output_dir: Path, prefix: str) -> list[Path]:
    values = sorted(
        output_dir.glob(f"{prefix}-*.json"),
        key=lambda path: int(re.search(r"-(\d+)\.json$", path.name).group(1))
        if re.search(r"-(\d+)\.json$", path.name)
        else -1,
    )
    _expect(len(values) == 4, f"expected four fresh {prefix} reports, found {len(values)}")
    _expect(
        [int(re.search(r"-(\d+)\.json$", path.name).group(1)) for path in values]
        == [1, 2, 3, 4],
        f"{prefix} report indices are not the four ABBA legs",
    )
    return values


def _validate_rows(report: Mapping[str, Any], path: Path) -> str:
    rows = report.get("rows")
    _expect(isinstance(rows, list), f"{path}: rows are missing")
    _expect(all(isinstance(row, dict) for row in rows), f"{path}: rows contain a non-object")
    _expect([row.get("operation") for row in rows if isinstance(row, dict)] == list(OPERATIONS), f"{path}: operation matrix differs")
    status = report.get("status")
    _expect(status in {"unavailable", "qualified"}, f"{path}: unsupported status {status!r}")
    if status == "unavailable":
        _expect(report.get("qualification_claim") is False, f"{path}: unavailable report claims qualification")
        _expect(isinstance(report.get("availability_reason"), str) and report["availability_reason"], f"{path}: unavailable reason is missing")
        _expect(all(row.get("status") == "unavailable" for row in rows), f"{path}: unavailable report has a non-unavailable row")
        _expect(all(row.get("elapsed_ns") is None for row in rows), f"{path}: unavailable report carries timing")
        return status

    _expect(report.get("qualification_claim") is True, f"{path}: qualified report has no claim marker")
    protocol = report.get("protocol")
    _expect(isinstance(protocol, dict), f"{path}: protocol receipt is missing")
    _expect(protocol.get("minimum_ms") == MINIMUM_MS, f"{path}: minimum duration was weakened")
    _expect(protocol.get("target_ms") == TARGET_MS, f"{path}: target duration differs")
    _expect(protocol.get("warmups") == WARMUPS and protocol.get("samples_per_leg") == SAMPLES and protocol.get("blocks") == BLOCKS, f"{path}: sample shape differs")
    _expect(protocol.get("fresh_process") is True, f"{path}: fresh-process provenance is missing")
    _expect(protocol.get("abba_order") == list(ABBA), f"{path}: ABBA order differs")
    expected_samples = BLOCKS * SAMPLES
    for row in rows:
        _expect(row.get("status") == "qualified", f"{path}: qualified report has an unavailable row")
        elapsed = row.get("elapsed_ns")
        _expect(isinstance(elapsed, int) and elapsed >= MINIMUM_MS * 1_000_000, f"{path}: accepted leg is shorter than {MINIMUM_MS} ms")
        samples = row.get("samples")
        _expect(isinstance(samples, list) and len(samples) == expected_samples, f"{path}: sample receipt shape differs")
        sample_elapsed = 0
        for sample in samples:
            _expect(isinstance(sample, dict), f"{path}: sample receipt is not an object")
            _expect(isinstance(sample.get("elapsed_ns"), int) and sample["elapsed_ns"] > 0, f"{path}: sample timer is missing")
            _expect(isinstance(sample.get("iterations"), int) and sample["iterations"] > 0, f"{path}: sample iteration receipt is missing")
            sample_elapsed += sample["elapsed_ns"]
        _expect(sample_elapsed == elapsed, f"{path}: row timer is not the sum of measured samples")
        _expect(row.get("fidelity_checked") is True, f"{path}: row lacks outside-clock fidelity check")
        details = row.get("details")
        _expect(isinstance(details, dict), f"{path}: semantic details are missing")
        operation = row.get("operation")
        if operation == "entry_read" or operation == "read_shared_when_supported":
            _expect(isinstance(details.get("output_sha256"), str) and len(details["output_sha256"]) == 64, f"{path}: read output identity is missing")
            _expect(isinstance(details.get("output_bytes"), int) and details["output_bytes"] > 0, f"{path}: read output size is missing")
        elif operation == "noop_commit":
            _expect(details.get("source_preserving_noop") is True and isinstance(details.get("source_sha256"), list) and details["source_sha256"], f"{path}: no-op fidelity receipt is incomplete")
            _expect(all(isinstance(value, str) and len(value) == 64 for value in details["source_sha256"]), f"{path}: no-op source identities are incomplete")
        elif operation == "changed_commit":
            _expect(details.get("semantic_changed_output_readback") is True and isinstance(details.get("changed_sha256"), list) and details["changed_sha256"], f"{path}: changed-output fidelity receipt is incomplete")
            _expect(all(isinstance(value, str) and len(value) == 64 for value in details["changed_sha256"]), f"{path}: changed-output identities are incomplete")
        elif operation == "forward_replay":
            _expect(details.get("forward_exact") is True and isinstance(details.get("parts"), list) and details["parts"], f"{path}: forward replay receipt is incomplete")
            for part in details["parts"]:
                _expect(isinstance(part, dict) and all(isinstance(part.get(key), str) and len(part[key]) == 64 for key in ("base_sha256", "actual_sha256", "output_sha256")) and part["base_sha256"] == part["actual_sha256"] and part["output_sha256"] != part["base_sha256"], f"{path}: forward base/output identity is incomplete")
        elif operation == "inverse_replay":
            _expect(details.get("inverse_exact") is True and isinstance(details.get("parts"), list) and details["parts"], f"{path}: inverse replay receipt is incomplete")
            for part in details["parts"]:
                _expect(isinstance(part, dict) and all(isinstance(part.get(key), str) and len(part[key]) == 64 for key in ("changed_sha256", "restored_sha256", "base_sha256")) and part["restored_sha256"] == part["base_sha256"] and part["changed_sha256"] != part["restored_sha256"], f"{path}: inverse restoration identity is incomplete")
        elif operation == "vec_write":
            _expect(isinstance(details.get("output_sha256"), str) and len(details["output_sha256"]) == 64, f"{path}: vector output identity is missing")
            _expect(isinstance(details.get("output_bytes"), int) and details["output_bytes"] > 0, f"{path}: vector output size is missing")
        elif operation == "write_to_bounded_sink":
            _expect(isinstance(row.get("sink"), dict), f"{path}: sink receipt is missing")
            sink = row["sink"]
            _expect(isinstance(sink.get("sha256"), str) and len(sink["sha256"]) == 64, f"{path}: sink hash is missing")
            _expect(isinstance(sink.get("bytes"), int) and sink["bytes"] > 0, f"{path}: sink byte receipt is missing")
            _expect(isinstance(sink.get("writes"), int) and sink["writes"] > 0 and row.get("sink_writes") == sink["writes"], f"{path}: sink write receipt is incomplete")
            _expect(details.get("sink_hash") == sink["sha256"] and details.get("sink_bytes") == sink["bytes"] and details.get("sink_writes") == sink["writes"], f"{path}: sink detail receipt differs from row receipt")
        if operation in {"stale_replay", "preflight_refusal", "malformed_refusal"}:
            _expect(row.get("typed_refusal_checked") is True, f"{path}: typed refusal was not checked")
            refusal = details.get("refusal")
            _expect(isinstance(refusal, dict), f"{path}: typed refusal receipt is missing")
            if operation == "stale_replay":
                _expect(refusal.get("kind") == "patch_conflict", f"{path}: stale replay did not report patch_conflict")
                _expect(refusal.get("resource") is None and isinstance(refusal.get("observed"), int) and refusal["observed"] > 0, f"{path}: stale observed bytes are missing")
                _expect(all(isinstance(details.get(key), str) and len(details[key]) == 64 for key in ("base_sha256", "stale_sha256", "output_sha256")) and details["base_sha256"] != details["stale_sha256"] and details["output_sha256"] != details["base_sha256"], f"{path}: stale patch identity is incomplete")
            elif operation == "preflight_refusal":
                _expect(details.get("preflight_zero_writes") is True and details.get("sink_writes") == 0, f"{path}: preflight wrote before refusal")
                _expect(refusal.get("kind") == "limit", f"{path}: preflight did not report limit")
                _expect(isinstance(refusal.get("resource"), str) and refusal["resource"] and isinstance(refusal.get("limit"), int) and isinstance(refusal.get("observed"), int) and refusal["observed"] > refusal["limit"] and refusal["observed"] == details.get("preflight_bytes"), f"{path}: preflight limit receipt is incomplete")
            else:
                _expect(details.get("malformed_refusal") is True and refusal.get("kind") == "invalid_input", f"{path}: malformed input did not report invalid_input")
                limit = details.get("limit")
                _expect(isinstance(limit, dict) and limit.get("limit_refusal") is True, f"{path}: malformed row lacks typed limit refusal")
                limit_refusal = limit.get("refusal")
                _expect(isinstance(limit_refusal, dict) and limit_refusal.get("kind") == "limit" and isinstance(limit_refusal.get("resource"), str) and limit_refusal["resource"] and isinstance(limit_refusal.get("limit"), int) and isinstance(limit_refusal.get("observed"), int) and limit_refusal["observed"] > limit_refusal["limit"], f"{path}: limit refusal details are incomplete")
    size_classes = report.get("size_classes")
    _expect(isinstance(size_classes, list) and [item.get("class") for item in size_classes if isinstance(item, dict)] == ["small", "near_limit", "at_limit", "over_limit", "malformed"], f"{path}: size-class matrix differs")
    for item in size_classes:
        _expect(isinstance(item, dict) and isinstance(item.get("bytes"), int) and item["bytes"] > 0, f"{path}: size-class byte receipt is missing")
        if item["class"] in {"small", "near_limit", "at_limit"}:
            _expect(item.get("accepted") is True and isinstance(item.get("output_sha256"), str) and len(item["output_sha256"]) == 64 and item.get("refusal") is None, f"{path}: accepted size class lacks readback identity")
        elif item["class"] == "over_limit":
            refusal = item.get("refusal")
            _expect(item.get("accepted") is False and isinstance(refusal, dict) and refusal.get("kind") == "limit" and isinstance(refusal.get("resource"), str) and refusal["resource"] and isinstance(refusal.get("limit"), int) and isinstance(refusal.get("observed"), int) and refusal["observed"] == item["bytes"] and item["bytes"] > refusal["limit"], f"{path}: over-limit size class lacks typed limit refusal")
        else:
            refusal = item.get("refusal")
            _expect(item.get("accepted") is False and isinstance(refusal, dict) and refusal.get("kind") == "invalid_input", f"{path}: malformed size class lacks invalid_input refusal")
    rows_by_operation = {row["operation"]: row for row in rows}
    source_ids = rows_by_operation["noop_commit"]["details"]["source_sha256"]
    changed_ids = rows_by_operation["changed_commit"]["details"]["changed_sha256"]
    forward_parts = rows_by_operation["forward_replay"]["details"]["parts"]
    inverse_parts = rows_by_operation["inverse_replay"]["details"]["parts"]
    _expect(len(source_ids) == len(changed_ids) == len(forward_parts) == len(inverse_parts), f"{path}: replay identity part counts differ")
    for index, (source_id, changed_id, forward, inverse) in enumerate(zip(source_ids, changed_ids, forward_parts, inverse_parts)):
        _expect(forward["base_sha256"] == source_id and forward["actual_sha256"] == source_id and forward["output_sha256"] == changed_id, f"{path}: forward replay identity differs at part {index}")
        _expect(inverse["changed_sha256"] == changed_id and inverse["restored_sha256"] == source_id and inverse["base_sha256"] == source_id, f"{path}: inverse replay identity differs at part {index}")
    return status


def validate(arguments: argparse.Namespace) -> dict[str, Any]:
    output_dir = arguments.output_dir.resolve(strict=False)
    _expect(output_dir.is_dir() and not output_dir.is_symlink(), f"output directory is invalid: {output_dir}")
    _expect(arguments.lane in LANES, f"unsupported lane: {arguments.lane}")
    fixture = arguments.fixture.resolve(strict=False)
    fixture_hash, fixture_bytes = _sha256(fixture)
    try:
        fixture_data = json.loads(fixture.read_text(encoding="utf-8"))
        fixture_root = fixture_data["lanes"][arguments.lane]["root"]
    except (OSError, KeyError, TypeError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ValidationError(f"fixture descriptor cannot provide lane root: {error}") from error
    reports: list[dict[str, Any]] = []
    unavailable = False
    for mode, prefix in (("normal", "dml-native"), ("allocator", "dml-allocator")):
        for index, path in enumerate(_ordered_outputs(output_dir, prefix)):
            report = _load(path)
            _expect(report.get("schema") == SCHEMA, f"{path}: schema differs")
            _expect(report.get("lane") == arguments.lane, f"{path}: lane differs")
            _expect(report.get("mode") == mode, f"{path}: mode differs")
            expected_role = ABBA[index]
            _expect(report.get("role") == expected_role, f"{path}: expected role {expected_role}")
            expected_revision = arguments.baseline_revision if expected_role == "baseline" else arguments.candidate_revision
            _expect(report.get("source_revision") == expected_revision, f"{path}: source revision differs")
            _expect(report.get("fixture_bytes") == fixture_bytes, f"{path}: fixture byte count differs")
            status = _validate_rows(report, path)
            if status == "qualified":
                _expect(report.get("fixture_sha256") == fixture_hash, f"{path}: fixture SHA-256 differs")
                _expect(report.get("fixture_root_sha256") == hashlib.sha256(fixture_root.encode()).hexdigest(), f"{path}: fixture root identity differs")
                _expect(isinstance(report.get("fixture_mutation"), str) and report["fixture_mutation"], f"{path}: fixture mutation receipt is missing")
            unavailable |= status == "unavailable"
            report_hash, report_bytes = _sha256(path)
            reports.append({
                "path": str(path),
                "sha256": report_hash,
                "bytes": report_bytes,
                "mode": mode,
                "role": expected_role,
                "status": status,
            })
            allocator_metrics = report.get("allocator_metrics")
            resource_metrics = report.get("resource_metrics")
            _expect(isinstance(resource_metrics, dict), f"{path}: process RSS receipt is missing")
            for key in ("process_rss_bytes", "peak_rss_bytes"):
                _expect(key in resource_metrics, f"{path}: resource metric {key} is missing")
                if status == "qualified":
                    _expect(isinstance(resource_metrics[key], int) and resource_metrics[key] > 0, f"{path}: qualified {key} is missing")
            if mode == "normal":
                _expect(allocator_metrics is None, f"{path}: normal report exposes allocator metrics")
            else:
                _expect(isinstance(allocator_metrics, dict), f"{path}: allocator report has no allocator receipt")
                _expect(allocator_metrics.get("elapsed_claimable") is False, f"{path}: allocator elapsed was made claimable")
                metric_keys = ("direct_bytes", "realloc_bytes", "live_balance", "peak_live_delta", "process_rss_bytes", "peak_rss_bytes") if status == "unavailable" else ("direct_bytes", "deallocated_bytes", "realloc_old_bytes", "realloc_new_bytes", "live_before", "live_after", "live_balance", "peak_live_delta", "process_rss_bytes", "peak_rss_bytes", "allocation_calls", "deallocation_calls", "reallocation_calls", "failed_calls", "conservation_ok")
                for key in metric_keys:
                    _expect(key in allocator_metrics, f"{path}: allocator metric {key} is missing")
                    if status == "qualified":
                        if key == "conservation_ok":
                            _expect(allocator_metrics[key] is True, f"{path}: allocator conservation is not proven")
                        else:
                            _expect(isinstance(allocator_metrics[key], int) and allocator_metrics[key] >= 0, f"{path}: qualified allocator metric {key} is invalid")
                if status == "qualified":
                    _expect(allocator_metrics.get("allocation_calls", 0) > 0, f"{path}: allocator observed no allocation calls")
                    _expect(allocator_metrics["live_after"] - allocator_metrics["live_before"] == allocator_metrics["live_balance"], f"{path}: allocator live balance is inconsistent")
                    _expect(allocator_metrics["live_balance"] == allocator_metrics["direct_bytes"] + allocator_metrics["realloc_new_bytes"] - allocator_metrics["deallocated_bytes"] - allocator_metrics["realloc_old_bytes"], f"{path}: allocator byte conservation equation is not proven")
                    _expect(allocator_metrics["peak_live_delta"] >= max(0, allocator_metrics["live_balance"]), f"{path}: allocator peak-live receipt is inconsistent")
                    _expect(allocator_metrics["process_rss_bytes"] == resource_metrics["process_rss_bytes"] and allocator_metrics["peak_rss_bytes"] == resource_metrics["peak_rss_bytes"], f"{path}: allocator and process RSS receipts differ")
    return {
        "schema": SUMMARY_SCHEMA,
        "lane": arguments.lane,
        "baseline_revision": arguments.baseline_revision,
        "candidate_revision": arguments.candidate_revision,
        "fixture": {"path": str(fixture), "sha256": fixture_hash, "bytes": fixture_bytes},
        "source_fence": arguments.source_fence,
        "abba_order": list(ABBA),
        "qualification_status": "unavailable" if unavailable else "qualified",
        "accepted": not unavailable,
        "claim_scope": "paired raw DrawingML protocol only; no absolute-profile or whole-library claim",
        "reports": reports,
        "unavailable_reason": "at least one lane adapter is not supplied until the final source freeze" if unavailable else None,
    }


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--lane", required=True)
    parser.add_argument("--fixture", type=Path, required=True)
    parser.add_argument("--baseline-revision", required=True)
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument("--source-fence", required=True)
    arguments = parser.parse_args(argv)
    try:
        summary = validate(arguments)
        target = arguments.output_dir / "summary.json"
        _expect(not target.exists() and not target.is_symlink(), f"refusing to overwrite {target}")
        target.write_text(json.dumps(summary, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    except (OSError, ValidationError) as error:
        print(f"drawingml raw validation refused: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
