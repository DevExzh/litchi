#!/usr/bin/env python3
"""Validate and summarize one archived database-function capture.

The capture archive is read directly; no member is extracted to a temporary
directory.  The analyzer checks the outer custody receipt, every archived
child artifact, the child result/status pairing, and exact non-timing metrics
across the three rounds before producing a compact per-case summary.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import tarfile
from pathlib import Path
from typing import Any, NoReturn


TIMING_FIELDS = {
    "elapsed_ns",
    "elapsed_ns_mean",
    "elapsed_ns_p95",
    "elapsed_ns_p99",
    "elapsed_ns_per_repeat",
}
RSS_PATTERN = re.compile(r"Maximum resident set size \(kbytes\): (\d+)")
EXIT_PATTERN = re.compile(r"Exit status: (\d+)")


def fail(message: str) -> NoReturn:
    raise SystemExit(f"validation failed: {message}")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_path(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def load_json(data: bytes, name: str) -> Any:
    try:
        return json.loads(data)
    except json.JSONDecodeError as error:
        fail(f"{name} is not JSON: {error}")


def member_bytes(archive: tarfile.TarFile, name: str) -> bytes:
    try:
        member = archive.getmember(name)
    except KeyError:
        fail(f"capture is missing {name}")
    if not member.isfile():
        fail(f"capture member {name} is not a regular file")
    stream = archive.extractfile(member)
    if stream is None:
        fail(f"capture member {name} cannot be read")
    return stream.read()


def integer_median(values: list[int]) -> int:
    ordered = sorted(values)
    return ordered[len(ordered) // 2]


def parse_rss(data: bytes, name: str) -> int:
    match = RSS_PATTERN.search(data.decode("utf-8", errors="strict"))
    if match is None:
        fail(f"{name} has no maximum RSS record")
    return int(match.group(1))


def validate_build(build_path: Path, expected_sha: str | None) -> dict[str, Any]:
    if not build_path.is_file():
        fail(f"build archive not found: {build_path}")
    actual_sha = sha256_path(build_path)
    if expected_sha is not None and actual_sha != expected_sha:
        fail(f"build archive hash {actual_sha} != outer receipt {expected_sha}")

    build_receipt_bytes = b""
    with tarfile.open(build_path, "r:gz") as archive:
        names = archive.getnames()
        if len(names) != len(set(names)):
            fail("build archive contains duplicate member names")
        build_receipt_bytes = member_bytes(archive, "receipt.json")
        receipt = load_json(build_receipt_bytes, "build receipt")
        if receipt.get("status") != 0 or not receipt.get("sources_unchanged"):
            fail("build receipt does not report a successful unchanged-source build")

        harness_hashes = receipt.get("harness_sha256")
        if not isinstance(harness_hashes, dict) or not harness_hashes:
            fail("build receipt has no harness hashes")
        for relative, expected in harness_hashes.items():
            name = f"harness/{relative}"
            actual = sha256_bytes(member_bytes(archive, name))
            if actual != expected:
                fail(f"{name} hash {actual} != build receipt {expected}")
        if sha256_bytes(member_bytes(archive, "build.log")) != receipt.get("log_sha256"):
            fail("build.log hash does not match build receipt")
        if sha256_bytes(member_bytes(archive, "build.py")) != receipt.get("runner_sha256"):
            fail("build.py hash does not match build receipt")

        expected_names = {"build.log", "build.py", "receipt.json"}
        expected_names.update(f"harness/{relative}" for relative in harness_hashes)
        if set(names) != expected_names:
            fail("build archive contains an unexpected or missing member")

    return {
        "archive_sha256": actual_sha,
        "receipt_sha256": sha256_bytes(build_receipt_bytes),
        "source_receipt": receipt.get("source_receipt"),
        "source_receipt_sha256": receipt.get("source_receipt_sha256"),
        "binary": receipt.get("binary"),
        "binary_sha256": receipt.get("binary_sha256"),
        "harness_sha256": harness_hashes,
        "build_status": receipt.get("status"),
        "rustc": receipt.get("rustc"),
        "cargo": receipt.get("cargo"),
    }


def validate_capture(
    capture_path: Path, outer_path: Path, build_path: Path | None
) -> dict[str, Any]:
    if not capture_path.is_file():
        fail(f"capture archive not found: {capture_path}")
    capture_sha = sha256_path(capture_path)
    outer = load_json(outer_path.read_bytes(), str(outer_path))
    expected_capture_sha = outer.get("archives", {}).get("capture.tar.gz")
    if capture_sha != expected_capture_sha:
        fail(f"capture archive hash {capture_sha} != outer receipt {expected_capture_sha}")

    expected_build_sha = outer.get("archives", {}).get("build.tar.gz")
    build_summary = None
    if build_path is not None:
        build_summary = validate_build(build_path, expected_build_sha)

    with tarfile.open(capture_path, "r:gz") as archive:
        names = archive.getnames()
        if len(names) != len(set(names)):
            fail("capture archive contains duplicate member names")
        inner = load_json(member_bytes(archive, "receipt.json"), "capture receipt")
        if not inner.get("binary_unchanged"):
            fail("capture receipt does not report an unchanged binary")

        cases = inner.get("cases")
        records = inner.get("records")
        artifacts = inner.get("artifacts")
        if not isinstance(cases, list) or not cases or len(set(cases)) != len(cases):
            fail("capture receipt has no unique case list")
        if not isinstance(records, list) or not isinstance(artifacts, dict):
            fail("capture receipt has malformed records or artifacts")
        if len(records) != len(cases) * 3:
            fail(f"expected three records per case, found {len(records)} records")
        if outer.get("cases") != len(cases) or outer.get("children") != len(records):
            fail("outer receipt counts do not match capture receipt")
        if outer.get("all_child_statuses_zero") is not True:
            fail("outer receipt does not assert zero child statuses")
        if inner.get("runner_sha256") != sha256_bytes(member_bytes(archive, "run.py")):
            fail("run.py hash does not match capture receipt")
        if build_summary is not None:
            binary_before = inner.get("binary_before")
            binary_after = inner.get("binary_after")
            build_binary = build_summary.get("binary_sha256")
            if not (binary_before == binary_after == build_binary):
                fail(
                    "capture and build binary hashes disagree: "
                    f"before={binary_before}, after={binary_after}, build={build_binary}"
                )

        artifact_names = set(artifacts)
        archive_names = set(names)
        expected_artifact_names = {
            f"{round_number:02}-{case}{suffix}"
            for round_number in range(1, 4)
            for case in cases
            for suffix in (".jsonl", ".stderr", ".time")
        }
        if artifact_names != expected_artifact_names:
            fail("capture receipt artifact list does not match the case/round matrix")
        if archive_names != expected_artifact_names | {"receipt.json", "run.py"}:
            fail("capture archive contains an unexpected or missing member")
        for name, expected in artifacts.items():
            actual = sha256_bytes(member_bytes(archive, name))
            if actual != expected:
                fail(f"{name} hash {actual} != capture receipt {expected}")

        by_case: dict[str, list[dict[str, Any]]] = {case: [] for case in cases}
        seen: set[tuple[int, str]] = set()
        result_keys: set[str] | None = None
        for record in records:
            case = record.get("case")
            round_number = record.get("round")
            if case not in by_case or round_number not in (1, 2, 3):
                fail(f"record has unknown case/round: {case!r}/{round_number!r}")
            key = (round_number, case)
            if key in seen:
                fail(f"duplicate child record {round_number:02}-{case}")
            seen.add(key)
            if record.get("status") != 0:
                fail(f"child {round_number:02}-{case} status is {record.get('status')}")

            stem = f"{round_number:02}-{case}"
            row_data = member_bytes(archive, f"{stem}.jsonl").decode("utf-8")
            lines = [line for line in row_data.splitlines() if line.strip()]
            if len(lines) != 1:
                fail(f"{stem}.jsonl has {len(lines)} result lines, expected one")
            row = load_json(lines[0].encode(), f"{stem}.jsonl")
            if not isinstance(row, dict) or not isinstance(record.get("result"), dict):
                fail(f"{stem}.jsonl or receipt result is not an object")
            if result_keys is None:
                result_keys = set(row)
            elif set(row) != result_keys:
                fail(f"{stem}.jsonl result key set differs from the first child")
            if row != record.get("result"):
                fail(f"{stem}.jsonl differs from its receipt result")
            time_data = member_bytes(archive, f"{stem}.time")
            exit_match = EXIT_PATTERN.search(time_data.decode("utf-8", errors="strict"))
            if exit_match is None or int(exit_match.group(1)) != 0:
                fail(f"{stem}.time does not report exit status 0")
            row["_rss_kib"] = parse_rss(time_data, f"{stem}.time")
            by_case[case].append({"round": round_number, "result": row})

        if len(seen) != len(records) or len(seen) != len(cases) * 3:
            fail("capture does not contain exactly one child for every case and round")

        summaries = []
        deterministic_total = 0
        deterministic_mismatches: list[dict[str, Any]] = []
        for case in cases:
            entries = sorted(by_case[case], key=lambda entry: entry["round"])
            rows = [entry["result"] for entry in entries]
            if [entry["round"] for entry in entries] != [1, 2, 3]:
                fail(f"case {case} is missing a round")
            field_names = set(rows[0]) - {"_rss_kib"}
            for row in rows[1:]:
                field_names &= set(row) - {"_rss_kib"}
            deterministic_fields = sorted(field_names - TIMING_FIELDS)
            deterministic_total += len(deterministic_fields)
            for field in deterministic_fields:
                values = [row[field] for row in rows]
                if any(value != values[0] for value in values[1:]):
                    deterministic_mismatches.append(
                        {"case": case, "field": field, "values": values}
                    )

            first = rows[0]
            repeat = int(first["repeat"])
            if repeat <= 0:
                fail(f"case {case} has non-positive repeat")
            if any(row["requested_bytes"] != row["released_bytes"] for row in rows):
                fail(f"case {case} does not balance requested and released allocator bytes")
            if any(row["live_before"] != row["live_after"] for row in rows):
                fail(f"case {case} does not restore live allocator bytes")
            summaries.append(
                {
                    "case": case,
                    "function": first["function"],
                    "size": first["size"],
                    "phase": first["phase"],
                    "expected": first["expected"],
                    "failure": first["failure"],
                    "shape": first["shape"],
                    "repeat": repeat,
                    "warmups": first["warmups"],
                    "iterations": first["iterations"],
                    "input_bytes": first["input_bytes"],
                    "median_ns_per_evaluation": integer_median(
                        [int(row["elapsed_ns_per_repeat"]) for row in rows]
                    ),
                    "median_p95_ns_per_evaluation": integer_median(
                        [int(row["elapsed_ns_p95"]) // repeat for row in rows]
                    ),
                    "rss_kib_rounds": [int(row["_rss_kib"]) for row in rows],
                    "rss_kib_median": integer_median(
                        [int(row["_rss_kib"]) for row in rows]
                    ),
                    "work_per_evaluation": first["work_per_repeat"],
                    "allocator_calls_per_evaluation": first["allocator_calls"] // repeat,
                    "deallocator_calls_per_evaluation": first["deallocator_calls"] // repeat,
                    "requested_bytes_per_evaluation": first["requested_bytes"] // repeat,
                    "released_bytes_per_evaluation": first["released_bytes"] // repeat,
                    "live_before": first["live_before"],
                    "live_after": first["live_after"],
                    "peak_live_delta": first["peak_live_delta"],
                    "memory_retained": first["memory_retained"],
                    "provider_reads_per_evaluation": first["provider_reads"] // repeat,
                    "provider_database_reads_per_evaluation": first[
                        "provider_database_reads"
                    ]
                    // repeat,
                    "provider_criteria_reads_per_evaluation": first[
                        "provider_criteria_reads"
                    ]
                    // repeat,
                    "provider_extent_calls_per_evaluation": first["provider_extent_calls"]
                    // repeat,
                    "deterministic_fields": len(deterministic_fields),
                }
            )

        if deterministic_mismatches:
            fail(
                "non-timing metrics differ across rounds: "
                + json.dumps(deterministic_mismatches, sort_keys=True)
            )

        source_receipt = None
        source_receipt_sha256 = None
        source_verification: dict[str, Any] = {
            "basis": outer.get("source_basis"),
            "receipt_checked": False,
            "archive_checked": False,
        }
        if build_summary is not None:
            source_receipt = build_summary.get("source_receipt")
            source_receipt_sha256 = build_summary.get("source_receipt_sha256")
            basis = outer.get("source_basis")
            if isinstance(basis, str):
                source_path = (outer_path.parent / basis).resolve()
                if source_path.is_file():
                    source_actual = sha256_path(source_path)
                    if source_actual != source_receipt_sha256:
                        fail(
                            "source-basis receipt hash "
                            f"{source_actual} != build receipt {source_receipt_sha256}"
                        )
                    source_verification["receipt_checked"] = True
                    source_verification["receipt_sha256"] = source_actual
                    source_data = load_json(source_path.read_bytes(), str(source_path))
                    source_archive_sha = source_data.get("archive_sha256")
                    source_archive_path = source_path.with_name("source.tar.gz")
                    if source_archive_path.is_file() and source_archive_sha:
                        archive_actual = sha256_path(source_archive_path)
                        if archive_actual != source_archive_sha:
                            fail(
                                "source archive hash "
                                f"{archive_actual} != source receipt {source_archive_sha}"
                            )
                        source_verification["archive_checked"] = True
                        source_verification["archive_sha256"] = archive_actual
        return {
            "capture_archive_sha256": capture_sha,
            "outer_receipt": str(outer_path),
            "outer_receipt_sha256": sha256_path(outer_path),
            "build": build_summary,
            "capture_receipt": {
                "binary": inner.get("binary"),
                "binary_before": inner.get("binary_before"),
                "binary_after": inner.get("binary_after"),
                "binary_unchanged": inner.get("binary_unchanged"),
                "binary_sha256": inner.get("binary_before"),
                "runner_sha256": inner.get("runner_sha256"),
                "cpu": inner.get("cpu"),
                "cases": len(cases),
                "children": len(records),
                "rounds": 3,
                "source_receipt": source_receipt,
                "source_receipt_sha256": source_receipt_sha256,
            },
            "source_verification": source_verification,
            "validation": {
                "artifact_hashes_verified": len(artifacts),
                "child_statuses_verified": len(records),
                "deterministic_fields_checked": deterministic_total,
                "deterministic_fields_match": True,
                "allocator_bytes_balanced_cases": len(cases),
                "live_bytes_balanced_cases": len(cases),
                "timing_fields_excluded": sorted(TIMING_FIELDS),
                "rss_from": "/usr/bin/time -v Maximum resident set size",
            },
            "cases": summaries,
        }


def main() -> None:
    parser = argparse.ArgumentParser(
        description="validate and summarize an archived database-function capture"
    )
    parser.add_argument("capture", type=Path, help="release capture.tar.gz")
    parser.add_argument(
        "--receipt",
        type=Path,
        help="outer release receipt (default: capture sibling receipt.json)",
    )
    parser.add_argument(
        "--build",
        type=Path,
        help="build archive (default: capture sibling build.tar.gz)",
    )
    args = parser.parse_args()
    outer_path = args.receipt or args.capture.with_name("receipt.json")
    build_path = args.build or args.capture.with_name("build.tar.gz")
    if not outer_path.is_file():
        fail(f"outer release receipt not found: {outer_path}")
    summary = validate_capture(args.capture, outer_path, build_path)
    print(json.dumps(summary, indent=2, sort_keys=False))


if __name__ == "__main__":
    main()
