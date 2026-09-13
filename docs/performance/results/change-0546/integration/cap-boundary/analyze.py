#!/usr/bin/env python3
"""Validate the supplemental 0546 XLSX cap-boundary valid-path guard.

The cap lane is intentionally independent of the main 0546 admission report.
It authenticates the source manifests, retained example binaries, receipts,
receipt-bound fixture output, correctness envelope and raw timing vectors for sizes
160, 164 and 256.  Every size/repeat candidate p50 and mean must remain at or
below 1.05 times its matched baseline.  Missing evidence produces ``pending``;
this analyzer never invents a capture or converts a missing vector to zero.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
from pathlib import Path
import re
import sys
from typing import Any

sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))

import cap_run as CAP  # noqa: E402  (the sibling driver owns custody paths)


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
SCRIPT_PATH = Path(__file__).resolve()
CAP_RUN_PATH = HERE / "cap_run.py"
MAIN_RUN_PATH = CAP.MAIN_BUNDLE / "run.py"
STAGES = ("baseline", "candidate")
REPEATS = (1, 2)
SIZES = (1, 2, 160, 164, 256)
WARMUP = 10
SAMPLES = 100
CPU = 2
EXPECTED_SCHEMA = "litchi.xlsx.cap-boundary-guard.v1"
EXPECTED_TOOL = "perf_cap_boundary"
CASE = "valid"
ADVERSE_THRESHOLD_PERCENT = 5.0
MAX_CANDIDATE_RATIO = 1.05
TIME_FIELDS = ("p50", "mean")
RSS_FIELD = "max_rss_kib"


class EvidenceError(ValueError):
    """A missing, malformed or contradictory cap-boundary artifact."""


class Pending(EvidenceError):
    """A later campaign phase has not produced its evidence yet."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    try:
        path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n",
                        encoding="utf-8")
    except OSError as error:
        raise EvidenceError(f"cannot write {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError as error:
        raise EvidenceError(f"path is outside cap-boundary evidence: {path}") from error


def regular(path: Path, label: str) -> None:
    require(path.is_file() and not path.is_symlink(),
            f"{label} is not a regular file")


def digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and re.fullmatch(r"[0-9a-f]{64}", value),
            f"{label} is not a lowercase SHA-256 digest")
    return value


def positive_integer(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")
    return value


def nonnegative_integer(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")
    return value


def finite_number(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")
    return float(value)


def nonempty_string(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label} is not a nonempty string")
    return value


def plan_data() -> dict[str, Any]:
    value = read_json(PLAN_PATH)
    require(isinstance(value, dict), "cap-boundary plan is not an object")
    revision = value.get("revision")
    require(isinstance(revision, str) and re.fullmatch(r"[0-9a-f]{40}", revision),
            "cap-boundary plan revision is malformed")
    cap = value.get("cap_boundary")
    require(isinstance(cap, dict), "cap-boundary section is missing")
    require(cap.get("sizes") == list(SIZES), "cap-boundary size matrix differs")
    require(cap.get("repeats") == len(REPEATS)
            and cap.get("warmup") == WARMUP
            and cap.get("samples") == SAMPLES,
            "cap-boundary repeat, warmup or sample count differs")
    require(cap.get("case") == CASE and cap.get("binary") == CAP.EXAMPLE,
            "cap-boundary case or binary differs")
    require(value.get("cpu") == CPU, "cap-boundary CPU differs")
    require(value.get("owned_paths") == [str(CAP.TARGET)],
            "cap-boundary owned target differs")
    require(value.get("candidate_source_roots") == [
        "crates/litchi-xlsx/src/cell_values/",
        "crates/litchi-xlsx/src/raw/worksheet/",
        "crates/litchi-xlsx/examples/",
    ], "cap-boundary candidate roots differ")
    require(value.get("candidate_files") == [
        "crates/litchi-xlsx/src/cell_values/snapshot.rs",
        "crates/litchi-xlsx/src/cell_values/validation.rs",
        "crates/litchi-xlsx/src/cell_values/shared_traversal_tests.rs",
        "crates/litchi-xlsx/src/raw/worksheet/codec.rs",
        "crates/litchi-xlsx/src/raw/worksheet/mod.rs",
    ], "cap-boundary candidate file inventory differs")
    require(value.get("shared_test_source_prefixes") == [
        "crates/litchi-xlsx/tests/source_backed_cell_values",
    ], "cap-boundary shared test prefix differs")
    status = value.get("status")
    require(status == "draft" or (isinstance(status, str)
                                   and status.startswith("frozen-before")),
            "cap-boundary plan status is invalid")
    return value


def expected_job_names() -> list[tuple[str, str, int, int]]:
    return [
        (CAP.job_name(stage, repeat, size), stage, repeat, size)
        for stage, repeat in CAP.NATIVE_ORDER
        for size in CAP.ordered_sizes(repeat)
    ]


def _stage_manifest(stage: str) -> tuple[dict[str, str], str]:
    path = HERE / stage / "source-manifest.json"
    regular(path, f"{stage}/source-manifest.json")
    value = read_json(path)
    require(isinstance(value, dict) and value,
            f"{stage} source manifest is empty")
    for name, value_digest in value.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"{stage} source manifest path is invalid")
        digest(value_digest, f"{stage} source manifest {name}")
    example = "crates/litchi-xlsx/examples/perf_cap_boundary.rs"
    require(example in value, f"{stage} source manifest omits cap example")
    return value, sha(path)


def _cleanup_allows_missing(binary: Path) -> bool:
    cleanup = CAP.MAIN_BUNDLE / "cleanup.json"
    if not cleanup.is_file():
        return False
    try:
        value = read_json(cleanup)
    except EvidenceError:
        return False
    return (
        value.get("owned_paths_absent") is True
        and value.get("accessible_process_references") == []
        and value.get("removed") == [str(CAP.TARGET)]
        and not CAP.TARGET.exists()
        and not binary.exists()
    )


def _binary_identity(stage: str, manifest_sha: str) -> dict[str, Any]:
    path = HERE / stage / "binary-cap-boundary.json"
    regular(path, f"{stage}/binary-cap-boundary.json")
    value = read_json(path)
    require(isinstance(value, dict), f"{stage} binary identity is not an object")
    expected_path = CAP.stage_binary(stage)
    actual_path = Path(nonempty_string(value.get("path"), f"{stage} binary path"))
    require(actual_path == expected_path, f"{stage} binary path differs")
    binary_sha = digest(value.get("sha256"), f"{stage} binary SHA")
    binary_bytes = positive_integer(value.get("bytes"), f"{stage} binary bytes")
    if actual_path.exists():
        regular(actual_path, f"{stage} retained binary")
        require(sha(actual_path) == binary_sha,
                f"{stage} retained binary SHA differs")
        require(actual_path.stat().st_size == binary_bytes,
                f"{stage} retained binary size differs")
    else:
        require(_cleanup_allows_missing(actual_path),
                f"{stage} retained binary is missing without cleanup custody")
    identity_manifest = digest(value.get("source_manifest_sha256"),
                               f"{stage} binary source manifest SHA")
    require(identity_manifest == manifest_sha,
            f"{stage} binary source manifest differs")
    build_receipt_sha = digest(value.get("build_receipt_sha256"),
                               f"{stage} binary build receipt SHA")
    return {
        "path": str(actual_path),
        "sha256": binary_sha,
        "bytes": binary_bytes,
        "source_manifest_sha256": identity_manifest,
        "build_receipt_sha256": build_receipt_sha,
    }


def _build_receipt(stage: str, manifest_sha: str, binary: dict[str, Any]) -> dict[str, Any]:
    path = HERE / stage / "build-cap-boundary.receipt.json"
    regular(path, f"{stage}/build-cap-boundary.receipt.json")
    value = read_json(path)
    require(isinstance(value, dict), f"{stage} build receipt is not an object")
    require(value.get("exit_code") == 0, f"{stage} cap build failed")
    require(value.get("binary_sha256") is None,
            f"{stage} cap build unexpectedly binds an input binary")
    require(value.get("source_manifest_sha256") == manifest_sha,
            f"{stage} cap build source manifest differs")
    require(value.get("plan_sha256") == sha(PLAN_PATH),
            f"{stage} cap build plan differs")
    require(value.get("script_sha256") == sha(MAIN_RUN_PATH),
            f"{stage} cap build main runner differs")
    require(value.get("cap_driver_sha256") == sha(CAP_RUN_PATH),
            f"{stage} cap build driver differs")
    require(value.get("command") == CAP.build_command(),
            f"{stage} cap build command differs")
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict)
            and set(artifacts) == {"build-cap-boundary.stdout",
                                   "build-cap-boundary.stderr"},
            f"{stage} cap build artifacts differ")
    for name, value_digest in artifacts.items():
        artifact = HERE / stage / name
        regular(artifact, f"{stage}/{name}")
        require(sha(artifact) == value_digest,
                f"{stage}/{name} digest differs")
    require(binary["build_receipt_sha256"] == sha(path),
            f"{stage} binary build receipt binding differs")
    return {"sha256": sha(path), "command": value.get("command")}


def _rss(path: Path, label: str) -> dict[str, Any]:
    regular(path, label)
    value = read_json(path)
    require(isinstance(value, dict), f"{label} is not an object")
    for field in ("max_rss_kib", "elapsed_seconds", "user_seconds", "system_seconds"):
        finite_number(value.get(field), f"{label}.{field}")
        require(float(value[field]) >= 0.0, f"{label}.{field} is negative")
    return value


def _fixture_identity(report: dict[str, Any], fixture: Path,
                      receipt_artifacts: dict[str, Any], label: str) -> dict[str, Any]:
    regular(fixture, label)
    actual_sha = sha(fixture)
    actual_bytes = fixture.stat().st_size
    # The standalone guard deliberately has no hash dependency.  The parent
    # runner's receipt is the authoritative fixture binding; a report hash is
    # accepted only as a redundant consistency check when present.
    fixture_sha = digest(receipt_artifacts.get(fixture.name),
                         f"{label} receipt fixture SHA")
    require(actual_sha == fixture_sha,
            f"{label} receipt fixture SHA differs")
    reported_sha = report.get("fixture_sha256")
    reported_bytes = report.get("fixture_bytes")
    fixture_record = report.get("fixture")
    if isinstance(fixture_record, dict):
        reported_sha = fixture_record.get("sha256", fixture_record.get(
            "fixture_sha256", reported_sha))
        reported_bytes = fixture_record.get("bytes", fixture_record.get(
            "fixture_bytes", reported_bytes))
    if reported_sha is not None:
        require(actual_sha == digest(reported_sha, f"{label} reported fixture SHA"),
                f"{label} reported fixture SHA differs")
    if reported_bytes is not None:
        require(actual_bytes == positive_integer(reported_bytes,
                                                 f"{label} reported fixture bytes"),
                f"{label} reported fixture byte count differs")
    return {"sha256": actual_sha, "bytes": actual_bytes,
            "file": relative(fixture)}


def _source_identity(report: dict[str, Any], size: int, fixture: dict[str, Any],
                     label: str) -> dict[str, Any]:
    source = report.get("source")
    require(isinstance(source, dict), f"{label}.source is not an object")
    reported_size = source.get("size", source.get("size_bytes", report.get("size")))
    require(reported_size == size, f"{label}.source size differs")
    source_result = {"size": size}
    for field in ("generator", "archive_bytes", "worksheet_bytes", "source_sha256",
                  "worksheet_sha256", "fixture_kind"):
        if field in source:
            source_result[field] = source[field]
    if "generator" in source:
        nonempty_string(source["generator"], f"{label}.source.generator")
    for field in ("archive_bytes", "worksheet_bytes"):
        if field in source:
            positive_integer(source[field], f"{label}.source.{field}")
    for field in ("source_sha256", "worksheet_sha256"):
        if field in source:
            digest(source[field], f"{label}.source.{field}")
    if "source_sha256" in source:
        require(source["source_sha256"] == fixture["sha256"],
                f"{label}.source SHA does not bind fixture output")
    return source_result


def _column_name(column: int) -> str:
    value = column + 1
    output = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        output = chr(ord("A") + remainder) + output
    return output


def _guard_identity(report: dict[str, Any], size: int, binary: dict[str, Any],
                    fixture: dict[str, Any], label: str) -> None:
    """Check deterministic fixture/cap metadata emitted by the guard."""

    cells = size * size
    expected_events = 5 * cells + 2 * size + 6 + int(size <= 2)
    expected_relation = "below" if expected_events <= 131_072 else "above"
    require(report.get("rows") == size and report.get("columns") == size,
            f"{label} row/column shape differs")
    require(report.get("cells") == cells, f"{label}.cells differs")
    require(report.get("last_cell") == f"{_column_name(size - 1)}{size}",
            f"{label}.last_cell differs")
    require(report.get("event_count") == expected_events
            and report.get("expected_event_count") == expected_events,
            f"{label} event count differs")
    require(report.get("event_count_formula") == "5*N*N+2*N+6+I(N<=2) (including EOF)",
            f"{label} event count formula differs")
    require(report.get("comment_bytes") == (1024 * 1024 if size <= 2 else 0), f"{label} comment bytes differ")
    require(report.get("shared_provisional_event_cap") == 131_072
            and report.get("ordinary_parser_event_cap") == 1_000_000,
            f"{label} event caps differ")
    require(report.get("event_cap_relation") == expected_relation,
            f"{label} event-cap relation differs")
    require(report.get("source_stream_byte_limit") == 8 * 1024 * 1024
            and report.get("source_stream_eligible") is True,
            f"{label} source-stream eligibility differs")
    require(report.get("archive_bytes") == fixture["bytes"],
            f"{label} archive bytes do not bind fixture")
    fixture_out = str((HERE / fixture["file"]).resolve())
    require(report.get("fixture_out") == fixture_out,
            f"{label}.fixture_out differs")

    source = report.get("source")
    require(isinstance(source, dict), f"{label}.source is not an object")
    expected_source = {
        "shape": f"{size}x{size}",
        "rows": size,
        "columns": size,
        "format": "OOXML/XLSX",
        "fixture_kind": "single_worksheet_sparse_numeric_with_comment" if size <= 2 else "single_worksheet_dense_numeric_grid",
        "generator": "litchi-xlsx-cap-boundary-stored-grid-v1",
        "worksheet_member": "xl/worksheets/sheet1.xml",
        "compression": "stored",
        "encoding": "UTF-8",
        "marker_free": True,
        "identity_method": "parent binds SHA-256 of fixture_dump bytes",
    }
    for field, expected in expected_source.items():
        require(source.get(field) == expected,
                f"{label}.source.{field} differs")
    for field in ("source_xml_bytes", "worksheet_bytes", "worksheet_xml_bytes",
                  "archive_bytes"):
        positive_integer(source.get(field), f"{label}.source.{field}")
    require(source.get("archive_bytes") == fixture["bytes"],
            f"{label}.source.archive_bytes does not bind fixture")
    require(source.get("fixture_dump") == fixture_out,
            f"{label}.source.fixture_dump differs")

    binary_report = report.get("binary")
    require(isinstance(binary_report, dict), f"{label}.binary is not an object")
    require(binary_report.get("path") == binary["path"],
            f"{label}.binary.path differs")
    require(binary_report.get("bytes") == binary["bytes"],
            f"{label}.binary.bytes differs")
    require(binary_report.get("profile") == "release",
            f"{label}.binary.profile differs")
    require(binary_report.get("identity_method") ==
            "parent binds SHA-256 of the captured binary",
            f"{label}.binary.identity_method differs")


def _correctness(report: dict[str, Any], label: str) -> dict[str, Any]:
    value = report.get("correctness")
    require(isinstance(value, dict), f"{label}.correctness is not an object")
    required = ("source_unchanged", "source_bytes_unchanged",
                "source_xml_unchanged", "valid_snapshot_values",
                "empty_commit_is_noop", "commit_outside_timing",
                "no_op_publication_exact")
    for field in required:
        require(value.get(field) is True, f"{label}.correctness.{field} is not true")
    return {field: True for field in required}


def _phase(report: dict[str, Any], label: str) -> tuple[list[int], dict[str, Any]]:
    phase = report.get("phase")
    require(isinstance(phase, dict), f"{label}.phase is not an object")
    require(phase.get("name") == "edit_sheets", f"{label}.phase name differs")
    values = phase.get("duration_ns")
    require(isinstance(values, list) and len(values) == SAMPLES,
            f"{label}.phase duration cardinality differs")
    durations = [positive_integer(value, f"{label}.phase.duration_ns[{i}]")
                 for i, value in enumerate(values)]
    order = phase.get("sample_order")
    require(order == list(range(SAMPLES)), f"{label}.phase sample order differs")
    samples = phase.get("samples")
    if samples is not None:
        if isinstance(samples, int) and not isinstance(samples, bool):
            require(samples == SAMPLES, f"{label}.phase.samples cardinality differs")
        else:
            require(isinstance(samples, list) and len(samples) == SAMPLES,
                    f"{label}.phase.samples cardinality differs")
            for index, item in enumerate(samples):
                require(isinstance(item, dict),
                        f"{label}.phase.samples[{index}] is not an object")
                require(item.get("order") == index,
                        f"{label}.phase.samples[{index}] order differs")
                require(item.get("duration_ns") == durations[index],
                        f"{label}.phase.samples[{index}] duration differs")
    ordered = sorted(durations)
    p50 = (ordered[49] + ordered[50]) // 2
    mean = sum(durations) / len(durations)
    computed = {"p50": p50, "mean": mean, "min": ordered[0], "max": ordered[-1]}
    reported = phase.get("statistics", report.get("statistics"))
    if reported is not None:
        require(isinstance(reported, dict), f"{label}.statistics is not an object")
        if "p50" in reported:
            require(reported["p50"] == p50, f"{label}.statistics.p50 differs")
        if "mean" in reported:
            finite_number(reported["mean"], f"{label}.statistics.mean")
            require(math.isclose(float(reported["mean"]), mean,
                                 rel_tol=1e-12, abs_tol=1e-9),
                    f"{label}.statistics.mean differs")
    return durations, computed


def _validate_report(stage: str, repeat: int, size: int, binary: dict[str, Any],
                     manifest_sha: str) -> dict[str, Any]:
    name = CAP.job_name(stage, repeat, size)
    label = f"{stage}/{name}"
    report_path = HERE / stage / f"{name}.json"
    fixture_path = HERE / stage / f"{name}.fixture.bin"
    rss_path = HERE / stage / f"{name}.rss.json"
    stdout_path = HERE / stage / f"{name}.stdout"
    stderr_path = HERE / stage / f"{name}.stderr"
    for path, path_label in ((report_path, "report"), (fixture_path, "fixture"),
                             (rss_path, "RSS"), (stdout_path, "stdout"),
                             (stderr_path, "stderr")):
        regular(path, f"{label}/{path_label}")
    receipt_path = HERE / stage / f"{name}.receipt.json"
    regular(receipt_path, f"{label}/receipt")
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            f"{label} receipt is not successful")
    require(receipt.get("binary_sha256") == binary["sha256"],
            f"{label} receipt binary differs")
    require(receipt.get("source_manifest_sha256") == manifest_sha,
            f"{label} receipt source manifest differs")
    candidate_manifest_path = HERE / "candidate" / "source-manifest.json"
    # Only the final retained-baseline ABBA slot runs under the candidate
    # checkout.  Baseline repeat 1 was captured before that checkout existed
    # and must remain bound to the baseline manifest.
    working_sha = (sha(candidate_manifest_path)
                   if stage == "baseline" and repeat == 2
                   and candidate_manifest_path.is_file() else manifest_sha)
    require(receipt.get("working_source_manifest_sha256") == working_sha,
            f"{label} receipt working source differs")
    require(receipt.get("plan_sha256") == sha(PLAN_PATH),
            f"{label} receipt plan differs")
    require(receipt.get("script_sha256") == sha(MAIN_RUN_PATH),
            f"{label} receipt main runner differs")
    require(receipt.get("cap_driver_sha256") == sha(CAP_RUN_PATH),
            f"{label} receipt driver differs")
    report_bytes = report_path.read_bytes()
    require(stdout_path.read_bytes().rstrip(b"\n") == report_bytes,
            f"{label} stdout does not reproduce report")
    expected_command = [
        "taskset", "-c", str(CPU),
        "/usr/bin/time", "-f", CAP.TIME_FORMAT, "-o", str(rss_path),
        str(CAP.stage_binary(stage)),
        "--size", str(size), "--warmup", str(WARMUP),
        "--samples", str(SAMPLES), "--json", str(report_path),
        "--fixture-out", str(fixture_path),
    ]
    require(receipt.get("command") == expected_command,
            f"{label} command differs")
    artifacts = receipt.get("artifacts")
    expected_artifacts = {report_path.name, fixture_path.name, rss_path.name,
                          stdout_path.name, stderr_path.name}
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts,
            f"{label} artifact inventory differs")
    for artifact_name, artifact_sha in artifacts.items():
        artifact = HERE / stage / artifact_name
        regular(artifact, f"{label}/{artifact_name}")
        require(sha(artifact) == digest(artifact_sha,
                                        f"{label}/{artifact_name} receipt SHA"),
                f"{label}/{artifact_name} digest differs")
    report = read_json(report_path)
    require(isinstance(report, dict), f"{label} report is not an object")
    require(report.get("schema") == EXPECTED_SCHEMA,
            f"{label}.schema differs")
    require(report.get("tool") == EXPECTED_TOOL, f"{label}.tool differs")
    require(report.get("case") == CASE and report.get("size") == size,
            f"{label} case/size differs")
    require(report.get("warmup_iterations") == WARMUP
            and report.get("samples") == SAMPLES,
            f"{label} warmup/samples differ")
    fixture = _fixture_identity(report, fixture_path, artifacts, f"{label}.fixture")
    source = _source_identity(report, size, fixture, label)
    _guard_identity(report, size, binary, fixture, label)
    correctness = _correctness(report, label)
    durations, statistics = _phase(report, label)
    rss = _rss(rss_path, f"{label}.rss")
    runner = report.get("runner")
    if runner is not None:
        require(isinstance(runner, dict), f"{label}.runner is not an object")
        if "git_revision" in runner:
            require(runner["git_revision"] == plan_data()["revision"],
                    f"{label}.runner revision differs")
    if "timing_scope" in report:
        nonempty_string(report["timing_scope"], f"{label}.timing_scope")
    if "performance_claim" in report:
        nonempty_string(report["performance_claim"], f"{label}.performance_claim")
    return {
        "name": name,
        "stage": stage,
        "repeat": repeat,
        "size": size,
        "samples": SAMPLES,
        "durations_ns": durations,
        "statistics": statistics,
        "fixture": fixture,
        "source": source,
        "correctness": correctness,
        "rss": rss,
        "report_sha256": sha(report_path),
        "receipt_sha256": sha(receipt_path),
        "rss_sha256": sha(rss_path),
        "binary_sha256": binary["sha256"],
        "manifest_sha256": manifest_sha,
    }


def _stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    manifest, manifest_sha = _stage_manifest(stage)
    binary = _binary_identity(stage, manifest_sha)
    build = _build_receipt(stage, manifest_sha, binary)
    rows = []
    for repeat in REPEATS:
        for size in CAP.ordered_sizes(repeat):
            rows.append(_validate_report(stage, repeat, size, binary, manifest_sha))
    return {
        "stage": stage,
        "status": "pass",
        "manifest_sha256": manifest_sha,
        "manifest_entries": len(manifest),
        "binary": binary,
        "build_receipt_sha256": build["sha256"],
        "rows": rows,
    }


def _percent_change(before: float, after: float) -> float | None:
    if before == 0.0:
        return 0.0 if after == 0.0 else None
    return (after / before - 1.0) * 100.0


def _comparison_record(before: float, after: float) -> dict[str, Any]:
    change = _percent_change(before, after)
    return {
        "baseline": before,
        "candidate": after,
        "candidate_to_baseline_ratio": (after / before) if before else None,
        "change_percent": change,
    }


def _fixture_key(row: dict[str, Any]) -> tuple[str, int]:
    fixture = row["fixture"]
    return fixture["sha256"], fixture["bytes"]


def _fixture_matrix_identity(rows: list[dict[str, Any]], size: int) -> dict[str, Any]:
    """Require one fixture identity for all four stage/repeat rows of a size."""

    identities = {_fixture_key(row) for row in rows}
    require(len(identities) == 1,
            f"cap-boundary fixture SHA/bytes differ across size {size} rows")
    fixture_sha, fixture_bytes = next(iter(identities))
    return {
        "sha256": fixture_sha,
        "bytes": fixture_bytes,
        "files": [row["fixture"]["file"] for row in rows],
    }


def _append_adverse(adverse: list[dict[str, Any]], *, stage: str,
                    repeat: int, size: int, metric: str,
                    value: dict[str, Any], phase: str) -> None:
    change = value.get("change_percent")
    if change is not None and change > ADVERSE_THRESHOLD_PERCENT:
        adverse.append({
            "id": f"cap-native-{stage}-r{repeat}-{size}-{metric}",
            "stage": stage,
            "repeat": repeat,
            "size": size,
            "phase": phase,
            "metric": metric,
            "change_percent": change,
            "baseline": value["baseline"],
            "candidate": value["candidate"],
            "classification": "adverse_candidate_increase",
            "disposition": "review individually before retention",
        })


def _append_drift(drift: list[dict[str, Any]], *, stage: str, size: int,
                  metric: str, value: dict[str, Any], phase: str) -> None:
    change = value.get("change_percent")
    if change is not None and abs(change) > ADVERSE_THRESHOLD_PERCENT:
        drift.append({
            "id": f"cap-native-{stage}-drift-{size}-{metric}",
            "stage": stage,
            "size": size,
            "phase": phase,
            "metric": metric,
            "change_percent": change,
            "first": value["baseline"],
            "second": value["candidate"],
            "classification": "same_build_repeat_drift",
            "disposition": "review individually before retention",
        })


def _compare(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    left = {(row["repeat"], row["size"]): row for row in baseline["rows"]}
    right = {(row["repeat"], row["size"]): row for row in candidate["rows"]}
    require(set(left) == set(right), "cap-boundary baseline/candidate matrices differ")
    rows = []
    adverse = []
    drift = []
    fixture_identities = {}
    for size in SIZES:
        # The four rows are baseline r1, candidate r1, candidate r2 and the
        # retained baseline r2.  Compare only SHA/bytes here; their evidence
        # file paths necessarily include the stage directory.
        fixture_identities[size] = _fixture_matrix_identity([
            left[(1, size)], right[(1, size)], right[(2, size)],
            left[(2, size)],
        ], size)
    for key in sorted(left):
        before, after = left[key], right[key]
        fixture_identity = fixture_identities[key[1]]
        require(_fixture_key(before) == _fixture_key(after),
                f"cap-boundary fixture identity differs for {key}")
        metrics = {}
        for field in TIME_FIELDS:
            base_value = float(before["statistics"][field])
            cand_value = float(after["statistics"][field])
            change = _percent_change(base_value, cand_value)
            ratio = cand_value / base_value if base_value else None
            passed = ratio is not None and ratio <= MAX_CANDIDATE_RATIO
            metrics[field] = {
                "baseline": before["statistics"][field],
                "candidate": after["statistics"][field],
                "candidate_to_baseline_ratio": ratio,
                "change_percent": change,
                "max_allowed_ratio": MAX_CANDIDATE_RATIO,
                "passed": passed,
            }
            if change is not None and change > ADVERSE_THRESHOLD_PERCENT:
                adverse.append({
                    "id": f"cap-native-candidate-r{key[0]}-{key[1]}-{field}",
                    "stage": "candidate",
                    "repeat": key[0],
                    "size": key[1],
                    "metric": field,
                    "change_percent": change,
                    "baseline": before["statistics"][field],
                    "candidate": after["statistics"][field],
                    "classification": "adverse_candidate_increase",
                    "disposition": "review individually before retention",
                })
        rss = _comparison_record(float(before["rss"][RSS_FIELD]),
                                 float(after["rss"][RSS_FIELD]))
        _append_adverse(adverse, stage="candidate", repeat=key[0], size=key[1],
                        metric=RSS_FIELD, value=rss, phase="rss")
        rows.append({
            "repeat": key[0],
            "size": key[1],
            "baseline": before["name"],
            "candidate": after["name"],
            "fixture": fixture_identity,
            "metrics": metrics,
            "rss": rss,
        })
    for stage_name, stage_data in (("baseline", baseline), ("candidate", candidate)):
        by_size = {(row["size"], row["repeat"]): row for row in stage_data["rows"]}
        for size in SIZES:
            first = by_size[(size, 1)]
            second = by_size[(size, 2)]
            for field in TIME_FIELDS:
                before = float(first["statistics"][field])
                after = float(second["statistics"][field])
                change = _percent_change(before, after)
                if change is not None and abs(change) > ADVERSE_THRESHOLD_PERCENT:
                    drift.append({
                        "id": f"cap-native-{stage_name}-drift-{size}-{field}",
                        "stage": stage_name,
                        "size": size,
                        "metric": field,
                        "change_percent": change,
                        "first": first["statistics"][field],
                        "second": second["statistics"][field],
                        "classification": "same_build_repeat_drift",
                        "disposition": "review individually before retention",
                    })
            rss = _comparison_record(float(first["rss"][RSS_FIELD]),
                                     float(second["rss"][RSS_FIELD]))
            _append_drift(drift, stage=stage_name, size=size, metric=RSS_FIELD,
                          value=rss, phase="rss")
    gate_rows = [row for row in rows]
    passed = all(
        metrics["passed"]
        for row in gate_rows
        for metrics in row["metrics"].values()
    )
    return {
        "status": "pass" if passed else "reject",
        "admission_passed": passed,
        "max_candidate_to_baseline_ratio": MAX_CANDIDATE_RATIO,
        "fixture_identities_by_size": fixture_identities,
        "rows": rows,
        "adverse": adverse,
        "drift": drift,
        "gate": "every size/repeat candidate planning p50 and mean <= 1.05x matched baseline",
    }


def _missing(stage: str) -> list[str]:
    folder = HERE / stage
    required = [folder / "source-manifest.json",
                folder / "binary-cap-boundary.json",
                folder / "build-cap-boundary.receipt.json"]
    for _, stage_name, repeat, size in expected_job_names():
        if stage_name != stage:
            continue
        name = CAP.job_name(stage, repeat, size)
        required.extend(folder / f"{name}{suffix}"
                        for suffix in (".json", ".fixture.bin", ".rss.json",
                                       ".receipt.json", ".stdout", ".stderr"))
    return [relative(path) for path in required if not path.exists()]


def pending_document(plan: dict[str, Any], selected: tuple[str, ...],
                     missing: dict[str, list[str]], reason: str | None = None) -> dict[str, Any]:
    value: dict[str, Any] = {
        "schema": "litchi.xlsx.cap-boundary-analysis.v1",
        "status": "pending",
        "plan": "plan.json",
        "plan_sha256": sha(PLAN_PATH),
        "driver": "cap_run.py",
        "driver_sha256": sha(CAP_RUN_PATH),
        "stages": {
            stage: {"status": "pending", "missing_artifacts": missing.get(stage, [])}
            for stage in selected
        },
        "scope": plan["scope"],
        "limitations": [
            "No capture or derived timing result is fabricated while evidence is incomplete.",
            "The cap lane is supplemental and does not alter the main 0546 gates.",
            "The Region/allocator and profile lanes are intentionally not measured here.",
        ],
    }
    if reason is not None:
        value["reason"] = reason
    return value


def analyze(stage_selection: str = "both") -> dict[str, Any]:
    require(stage_selection in ("baseline", "candidate", "both"),
            f"unknown cap-boundary stage selection: {stage_selection}")
    plan = plan_data()
    selected = (("baseline",) if stage_selection == "baseline" else
                ("candidate",) if stage_selection == "candidate" else STAGES)
    if plan.get("status") == "draft":
        return pending_document(plan, selected, {stage: [] for stage in selected},
                                "cap-boundary plan is still draft")
    missing = {stage: _missing(stage) for stage in selected}
    if any(missing[stage] for stage in selected):
        return pending_document(plan, selected, missing)
    stages = {stage: _stage(stage, plan) for stage in selected}
    result: dict[str, Any] = {
        "schema": "litchi.xlsx.cap-boundary-analysis.v1",
        "status": "pass",
        "plan": "plan.json",
        "plan_sha256": sha(PLAN_PATH),
        "driver": "cap_run.py",
        "driver_sha256": sha(CAP_RUN_PATH),
        "stages": stages,
        "scope": plan["scope"],
        "supplemental": True,
    }
    if set(stages) == set(STAGES):
        comparison = _compare(stages["baseline"], stages["candidate"])
        result["comparison"] = comparison
        result["status"] = "pass" if comparison["admission_passed"] else "reject"
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("baseline", "candidate", "both"),
                        default="both")
    parser.add_argument("--write", action="store_true",
                        help="write cap-analysis.json after read-only validation")
    parser.add_argument("--output", type=Path,
                        help="JSON output path; defaults to cap-analysis.json")
    args = parser.parse_args()
    try:
        result = analyze(args.stage)
        if args.write:
            output = args.output or HERE / "cap-analysis.json"
            output.parent.mkdir(parents=True, exist_ok=True)
            write_json(output, result)
            print(f"0546 cap-boundary analysis {result['status']}: {output}")
        else:
            print(json.dumps(result, indent=2, sort_keys=True))
    except (EvidenceError, OSError, ValueError, KeyError) as error:
        print(f"cap-boundary analyze.py: error: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
