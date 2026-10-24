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
    for row in rows:
        _expect(row.get("status") == "qualified", f"{path}: qualified report has an unavailable row")
        elapsed = row.get("elapsed_ns")
        _expect(isinstance(elapsed, int) and elapsed >= MINIMUM_MS * 1_000_000, f"{path}: accepted leg is shorter than {MINIMUM_MS} ms")
        _expect(row.get("fidelity_checked") is True, f"{path}: row lacks outside-clock fidelity check")
        if row.get("operation") in {"stale_replay", "preflight_refusal", "malformed_refusal"}:
            _expect(row.get("typed_refusal_checked") is True, f"{path}: typed refusal was not checked")
        if row.get("operation") == "write_to_bounded_sink":
            _expect(isinstance(row.get("sink_writes"), int) and row["sink_writes"] > 0, f"{path}: sink write receipt is missing")
    return status


def validate(arguments: argparse.Namespace) -> dict[str, Any]:
    output_dir = arguments.output_dir.resolve(strict=False)
    _expect(output_dir.is_dir() and not output_dir.is_symlink(), f"output directory is invalid: {output_dir}")
    _expect(arguments.lane in LANES, f"unsupported lane: {arguments.lane}")
    fixture = arguments.fixture.resolve(strict=False)
    fixture_hash, fixture_bytes = _sha256(fixture)
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
            _expect(isinstance(report.get("fixture_fingerprint"), str) and report["fixture_fingerprint"], f"{path}: fixture fingerprint is missing")
            status = _validate_rows(report, path)
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
                for key in ("direct_bytes", "realloc_bytes", "live_balance", "peak_live_delta", "process_rss_bytes", "peak_rss_bytes"):
                    _expect(key in allocator_metrics, f"{path}: allocator metric {key} is missing")
                    if status == "qualified":
                        _expect(isinstance(allocator_metrics[key], int) and allocator_metrics[key] >= 0, f"{path}: qualified allocator metric {key} is invalid")
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
