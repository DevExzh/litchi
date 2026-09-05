#!/usr/bin/env python3
"""Validate and summarize change-0416 strict ZIP ABBA observations.

The input directory is produced by ``capture.py``.  Every ordinary row keeps
all raw samples and is summarized with the exact median and nearest-rank p95
and p99.  Positive candidate and within-revision deltas above five percent are
retained as review triggers.  Capability fixtures are reported as
control-error/candidate-success statuses and are deliberately excluded from
latency comparisons.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
from typing import Any, NoReturn


LEGS = ("A1", "B1", "B2", "A2")
MODES = ("borrowed", "indexed")
ROLE_FOR_LEG = {"A1": "control", "B1": "candidate", "B2": "candidate", "A2": "control"}
U64_MAX = (1 << 64) - 1
RSS_RE = re.compile(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", re.MULTILINE)


def fail(message: str) -> NoReturn:
    raise ValueError(message)


def integer(
    value: object,
    where: str,
    *,
    minimum: int = 0,
    maximum: int = U64_MAX,
) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or not minimum <= value <= maximum:
        fail(f"{where} must be an integer in [{minimum}, {maximum}]")
    return value


def percentile(samples: list[int], number: int) -> int | float:
    ordered = sorted(samples)
    if number == 50:
        left = ordered[(len(ordered) - 1) // 2]
        right = ordered[len(ordered) // 2]
        total = left + right
        return total // 2 if total % 2 == 0 else total // 2 + 0.5
    return ordered[((number * len(ordered) + 99) // 100) - 1]


def statistics(samples: list[int]) -> dict[str, int | float]:
    return {f"p{number}": percentile(samples, number) for number in (50, 95, 99)}


def sha256_bytes(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def rss_for(root: Path, stem: str, directory: str = "guards") -> tuple[int, Path]:
    candidates = (
        root / directory / f"{stem}.time.txt",
        root / directory / f"{stem}.time-v.txt",
        root / "time-v" / f"{stem}.txt",
        root / "time-v" / f"{stem}.time.txt",
    )
    for path in candidates:
        if path.is_file():
            match = RSS_RE.search(path.read_text(encoding="utf-8"))
            if match:
                return int(match.group(1)), path
            fail(f"{path}: missing Maximum resident set size (kbytes)")
    fail(f"{stem}: required /usr/bin/time -v report is missing")


def relative(path: Path, root: Path) -> str:
    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return str(path.resolve())


def delta(later: int | float, earlier: int | float) -> float:
    if earlier <= 0:
        fail(f"cannot compute a delta from non-positive statistic {earlier!r}")
    return (later - earlier) * 100.0 / earlier


def load_json(path: Path) -> tuple[dict[str, Any], bytes]:
    try:
        raw = path.read_bytes()
        value = json.loads(raw)
    except (OSError, json.JSONDecodeError) as error:
        fail(f"{path}: cannot load JSON: {error}")
    if not isinstance(value, dict):
        fail(f"{path}: expected a JSON object")
    return value, raw


def check_observation(
    row: dict[str, Any],
    *,
    report: Path,
    mode: str,
    oracle: dict[str, Any],
    warmups_expected: int,
    samples_expected: int,
) -> tuple[list[int], dict[str, Any]]:
    if row.get("schema_version") != 1 or row.get("operation") != "index":
        fail(f"{report}: unsupported probe report identity")
    if row.get("mode") != mode:
        fail(f"{report}: mode {row.get('mode')!r}, expected {mode!r}")
    if row.get("payload_decompression") is not False:
        fail(f"{report}: timed operation must not decompress payloads")
    if row.get("borrowed_store_crc_scan") is not (mode == "borrowed"):
        fail(f"{report}: borrowed CRC-scan scope does not match reader mode")
    warmups = integer(row.get("warmups"), f"{report}.warmups")
    if warmups != warmups_expected:
        fail(f"{report}: warmups {warmups}, expected {warmups_expected}")
    sample_count = integer(row.get("sample_count"), f"{report}.sample_count", minimum=1)
    if sample_count != samples_expected:
        fail(f"{report}: sample_count {sample_count}, expected {samples_expected}")
    values = row.get("samples_ns")
    if not isinstance(values, list) or len(values) != sample_count:
        fail(f"{report}: samples_ns must contain exactly sample_count values")
    samples = [
        integer(value, f"{report}.samples_ns[{index}]", minimum=1)
        for index, value in enumerate(values)
    ]
    if row.get("oracle") != oracle:
        fail(f"{report}: oracle differs from the first leg")
    observation = row.get("observation")
    if not isinstance(observation, dict):
        fail(f"{report}: observation must be an object")
    for key in (
        "digest",
        "entries",
        "name_bytes",
        "compressed_bytes",
        "uncompressed_bytes",
    ):
        integer(observation.get(key), f"{report}.observation.{key}")
    integer(observation.get("crc32_xor"), f"{report}.observation.crc32_xor", maximum=0xFFFFFFFF)
    integer(observation.get("borrowed_some"), f"{report}.observation.borrowed_some")
    integer(observation.get("borrowed_none"), f"{report}.observation.borrowed_none")
    integer(
        observation.get("borrowed_crc_scan_bytes"),
        f"{report}.observation.borrowed_crc_scan_bytes",
    )
    if mode == "borrowed" and observation["borrowed_some"] == 0:
        fail(f"{report}: borrowed guard did not exercise a stored CRC scan")
    if mode != "borrowed" and observation["borrowed_crc_scan_bytes"] != 0:
        fail(f"{report}: indexed preservation unexpectedly reports a borrowed CRC scan")
    expected_entries = oracle["file_count"] if mode == "borrowed" else oracle["central_entry_count"]
    if observation["entries"] != expected_entries:
        fail(f"{report}: indexed entry count does not match oracle")
    return samples, {
        "warmups": warmups,
        "sample_count": samples_expected,
        "statistics_ns": statistics(samples),
        "observation": observation,
    }


def check_capture(
    capture: dict[str, Any], *, samples: int, warmups: int
) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    if capture.get("schema_version") != 1:
        fail("capture.json: unsupported schema version")
    if capture.get("abba_order") != list(LEGS):
        fail("capture.json: ABBA order is not A1/B1/B2/A2")
    if capture.get("modes") != list(MODES):
        fail("capture.json: reader mode list is not borrowed/indexed")
    if capture.get("samples") != samples or capture.get("warmups") != warmups:
        fail("capture.json: configured samples/warmups do not match summarizer arguments")
    fixtures = capture.get("fixtures")
    if not isinstance(fixtures, dict) or len(fixtures) < 4:
        fail("capture.json: at least four ordinary fixtures are required")
    for label, identity in fixtures.items():
        if not isinstance(label, str) or not isinstance(identity, dict):
            fail("capture.json: malformed fixture identity")
        if not isinstance(identity.get("sha256"), str) or len(identity["sha256"]) != 64:
            fail(f"capture.json: malformed fixture hash for {label!r}")
        integer(identity.get("bytes"), f"capture.json.fixtures[{label!r}].bytes", minimum=1)
    indexed_fixtures = capture.get("indexed_fixtures") or {}
    if not isinstance(indexed_fixtures, dict):
        fail("capture.json: indexed_fixtures must be an object")
    for label, identity in indexed_fixtures.items():
        if not isinstance(label, str) or not isinstance(identity, dict):
            fail("capture.json: malformed indexed-only fixture identity")
        if label in fixtures:
            fail(f"capture.json: fixture label is duplicated in indexed_fixtures: {label!r}")
        if not isinstance(identity.get("sha256"), str) or len(identity["sha256"]) != 64:
            fail(f"capture.json: malformed indexed-only fixture hash for {label!r}")
        integer(
            identity.get("bytes"),
            f"capture.json.indexed_fixtures[{label!r}].bytes",
            minimum=1,
        )
    runs = capture.get("runs")
    if not isinstance(runs, list):
        fail("capture.json: runs must be a list")
    expected_count = (len(fixtures) * len(MODES) + len(indexed_fixtures)) * len(LEGS)
    if len(runs) != expected_count:
        fail(f"capture.json: {len(runs)} ordinary runs, expected {expected_count}")
    seen: set[tuple[str, str, str]] = set()
    for run in runs:
        if not isinstance(run, dict):
            fail("capture.json: malformed ordinary run")
        leg, fixture, mode = run.get("leg"), run.get("fixture"), run.get("mode")
        key = (str(leg), str(fixture), str(mode))
        valid_fixture = fixture in fixtures or fixture in indexed_fixtures
        valid_mode = mode in MODES and (fixture in fixtures or mode == "indexed")
        if key in seen or leg not in LEGS or not valid_fixture or not valid_mode:
            fail(f"capture.json: duplicate or unexpected ordinary run {key!r}")
        seen.add(key)
        if fixture in indexed_fixtures and not run.get("indexed_only"):
            fail(f"capture.json: indexed-only run {key!r} lacks indexed_only marker")
        if fixture in fixtures and run.get("indexed_only"):
            fail(f"capture.json: ordinary run {key!r} is marked indexed_only")
        if run.get("role") != ROLE_FOR_LEG[leg] or run.get("exit_code") != 0:
            fail(f"capture.json: ordinary run {key!r} did not succeed")
    if len(seen) != expected_count:
        fail("capture.json: ordinary run matrix is incomplete")
    return fixtures, indexed_fixtures, capture.get("capability_fixtures") or {}


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--samples", type=int, default=300)
    parser.add_argument("--warmups", type=int, default=30)
    args = parser.parse_args()
    if args.samples <= 0 or args.warmups < 0:
        fail("--samples must be positive and --warmups must be non-negative")
    root = args.root.resolve()
    capture, _ = load_json(root / "capture.json")
    fixture_identities, indexed_fixture_identities, capability_identities = check_capture(
        capture, samples=args.samples, warmups=args.warmups
    )

    rows: list[dict[str, Any]] = []
    triggers: list[dict[str, Any]] = []
    fixture_specs = [
        (fixture_label, fixture_identity, MODES)
        for fixture_label, fixture_identity in fixture_identities.items()
    ] + [
        (fixture_label, fixture_identity, ("indexed",))
        for fixture_label, fixture_identity in indexed_fixture_identities.items()
    ]
    for fixture_label, fixture_identity, fixture_modes in fixture_specs:
        fixture_bytes = fixture_identity["bytes"]
        for mode in fixture_modes:
            label = f"{fixture_label}/{mode}"
            reports: dict[str, tuple[dict[str, Any], bytes, list[int], dict[str, Any]]] = {}
            oracle: dict[str, Any] | None = None
            for leg in LEGS:
                stem = f"{leg}-{fixture_label}-{mode}"
                report_path = root / "guards" / f"{stem}.json"
                report, raw = load_json(report_path)
                if integer(report.get("input_bytes"), f"{report_path}.input_bytes") != fixture_bytes:
                    fail(f"{report_path}: input_bytes differs from fixture identity")
                current_oracle = report.get("oracle")
                if not isinstance(current_oracle, dict):
                    fail(f"{report_path}: oracle must be an object")
                if oracle is None:
                    oracle = current_oracle
                sample_values, checked = check_observation(
                    report,
                    report=report_path,
                    mode=mode,
                    oracle=oracle,
                    warmups_expected=args.warmups,
                    samples_expected=args.samples,
                )
                reports[leg] = (report, raw, sample_values, checked)
            assert oracle is not None
            expected_observation = reports["A1"][3]["observation"]
            for leg, (_, _, _, checked) in reports.items():
                if checked["observation"] != expected_observation:
                    fail(f"{label}: observation oracle differs in {leg}")

            rss: dict[str, int] = {}
            rss_paths: dict[str, str] = {}
            for leg in LEGS:
                value, path = rss_for(root, f"{leg}-{fixture_label}-{mode}")
                rss[leg] = value
                rss_paths[leg] = relative(path, root)

            pairs = {
                "B1_minus_A1": ("candidate_regression", "B1", "A1"),
                "B2_minus_A2": ("candidate_regression", "B2", "A2"),
                "A2_minus_A1": ("within_revision_drift", "A2", "A1"),
                "B2_minus_B1": ("within_revision_drift", "B2", "B1"),
            }
            pair_deltas: dict[str, Any] = {}
            rss_deltas: dict[str, float] = {}
            for pair, (kind, later, earlier) in pairs.items():
                metrics = {
                    metric: delta(
                        reports[later][3]["statistics_ns"][metric],
                        reports[earlier][3]["statistics_ns"][metric],
                    )
                    for metric in ("p50", "p95", "p99")
                }
                pair_deltas[pair] = {"comparison_kind": kind, "metrics_percent": metrics}
                for metric, value in metrics.items():
                    if value > 5.0:
                        triggers.append(
                            {
                                "row": label,
                                "comparison_kind": kind,
                                "pair": pair,
                                "metric": metric,
                                "regression_percent": value,
                            }
                        )
                rss_delta = delta(rss[later], rss[earlier])
                rss_deltas[pair] = rss_delta
                if rss_delta > 5.0:
                    triggers.append(
                        {
                            "row": label,
                            "comparison_kind": kind,
                            "pair": pair,
                            "metric": "max_rss_kib",
                            "regression_percent": rss_delta,
                        }
                    )

            rows.append(
                {
                    "fixture": fixture_label,
                    "mode": mode,
                    "input_bytes": fixture_bytes,
                    "sample_count": args.samples,
                    "warmups": args.warmups,
                    "raw_sample_unit": "ns",
                    "legs": {
                        leg: {
                            "path": relative(root / "guards" / f"{leg}-{fixture_label}-{mode}.json", root),
                            "sha256": sha256_bytes(reports[leg][1]),
                            "statistics_ns": reports[leg][3]["statistics_ns"],
                            "observation": reports[leg][3]["observation"],
                        }
                        for leg in LEGS
                    },
                    "oracle": oracle,
                    "output_oracle": expected_observation,
                    "paired_delta_percent": pair_deltas,
                    "time_v_max_rss_kib": rss,
                    "time_v_paths": rss_paths,
                    "time_v_rss_delta_percent": rss_deltas,
                }
            )

    capability_rows: list[dict[str, Any]] = []
    capability_runs = capture.get("capability_runs")
    if capability_identities:
        if not isinstance(capability_identities, dict):
            fail("capture.json: capability_fixtures must be an object")
        if not isinstance(capability_runs, list):
            fail("capture.json: capability_runs must be a list")
        expected_count = len(capability_identities) * 2 * len(MODES)
        if len(capability_runs) != expected_count:
            fail(f"capture.json: {len(capability_runs)} capability runs, expected {expected_count}")
        seen_capability: set[tuple[str, str, str]] = set()
        for fixture_label, identity in capability_identities.items():
            if not isinstance(identity, dict):
                fail(f"capability fixture {fixture_label!r}: malformed identity")
            if not isinstance(identity.get("sha256"), str) or len(identity["sha256"]) != 64:
                fail(f"capability fixture {fixture_label!r}: malformed fixture hash")
            fixture_bytes = integer(
                identity.get("bytes"),
                f"capture.json.capability_fixtures[{fixture_label!r}].bytes",
                minimum=1,
            )
            for role in ("control", "candidate"):
                expected_status = "error" if role == "control" else "ok"
                for mode in MODES:
                    stem = f"{role}-{fixture_label}-{mode}"
                    matching = [
                        run
                        for run in capability_runs
                        if isinstance(run, dict)
                        and run.get("role") == role
                        and run.get("fixture") == fixture_label
                        and run.get("mode") == mode
                    ]
                    if len(matching) != 1:
                        fail(f"capability run {stem}: expected exactly one capture record")
                    run = matching[0]
                    report_path = root / "capability" / f"{stem}.json"
                    report, raw = load_json(report_path)
                    if (
                        report.get("schema_version") != 1
                        or report.get("operation") != "capability"
                        or report.get("mode") != mode
                        or report.get("payload_decompression") is not False
                        or report.get("status") != expected_status
                    ):
                        fail(
                            f"{report_path}: status {report.get('status')!r}, expected {expected_status!r}"
                        )
                    if integer(report.get("input_bytes"), f"{report_path}.input_bytes") != fixture_bytes:
                        fail(f"{report_path}: input_bytes differs from capability fixture identity")
                    if role == "candidate" and mode == "borrowed":
                        observation = report.get("observation")
                        if (
                            not isinstance(observation, dict)
                            or integer(observation.get("borrowed_some"), f"{report_path}.observation.borrowed_some")
                            == 0
                        ):
                            fail(f"{report_path}: candidate capability did not exercise a stored CRC scan")
                    capability_rss, capability_rss_path = rss_for(
                        root, stem, directory="capability"
                    )
                    capability_rows.append(
                        {
                            "fixture": fixture_label,
                            "role": role,
                            "mode": mode,
                            "status": report["status"],
                            "elapsed_ns": integer(report.get("elapsed_ns"), f"{report_path}.elapsed_ns", minimum=1),
                            "report": relative(report_path, root),
                            "sha256": sha256_bytes(raw),
                            "time_v": relative(capability_rss_path, root),
                            "time_v_max_rss_kib": capability_rss,
                            "exit_code": run.get("exit_code"),
                        }
                    )
                    seen_capability.add((role, fixture_label, mode))
        if len(seen_capability) != expected_count:
            fail("capability run matrix is incomplete")
    elif capability_runs:
        fail("capture.json has capability runs without capability fixture identities")

    result = {
        "schema_version": 1,
        "tool": "litchi-goal-0416-zip-strict-summary",
        "performance_claim": "none",
        "scope": "process-isolated strict borrowed Store validation and ReaderAt preservation indexing; payload decompression excluded",
        "abba_order": list(LEGS),
        "statistics": {
            "p50": "exact median (integer or .5)",
            "p95": "integer nearest rank",
            "p99": "integer nearest rank",
        },
        "threshold": {
            "regression_percent": 5.0,
            "rule": "retain positive candidate and within-revision latency or RSS deltas above threshold",
        },
        "rows": rows,
        "capability": {
            "rows": capability_rows,
            "latency_comparison": "excluded; statuses only",
        },
        "regression_triggers": triggers,
        "verification": {
            "ordinary_fixture_count": len(fixture_identities),
            "indexed_only_fixture_count": len(indexed_fixture_identities),
            "ordinary_row_count": len(rows),
            "ordinary_process_count": len(rows) * len(LEGS),
            "all_raw_samples_retained": True,
            "all_input_identity_verified": True,
            "all_observation_oracles_equal_per_row": True,
            "all_time_v_rss_present": True,
            "capability_statuses_verified": bool(capability_rows) or not capability_identities,
        },
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")


if __name__ == "__main__":
    try:
        main()
    except ValueError as error:
        raise SystemExit(f"error: {error}") from error
