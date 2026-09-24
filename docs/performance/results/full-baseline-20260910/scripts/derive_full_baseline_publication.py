#!/usr/bin/env python3
"""Derive and verify the review-safe historical full-baseline report.

The capture itself is immutable.  This tool removes only the 25 normal-report
``operation_metrics`` envelopes whose sample-index vectors do not match the
elapsed sample order.  It preserves the report-level metadata, every result's
timing, and every top-level result sink object.  It deliberately fails closed
if the raw report is not the captured 201-row report shape.
"""

from __future__ import annotations

import argparse
import copy
import gzip
import hashlib
import json
from pathlib import Path
from typing import Any


RAW_RESULT_COUNT = 201
RAW_OPERATION_METRICS_COUNT = 28
REMOVED_OPERATION_METRICS_COUNT = 25

# (result index, case, corpus archive SHA-256).  Keeping this identity list
# fixed prevents a changed input from silently changing the derived scope.
EXPECTED_REMOVALS: tuple[tuple[int, str, str], ...] = (
    (4, "opc_noop_save", "1e28b8a9049a82f07e8ea88b2d492ef522d2da793d22fa50e2fe7f354dca3e2a"),
    (5, "opc_mutated_save", "1e28b8a9049a82f07e8ea88b2d492ef522d2da793d22fa50e2fe7f354dca3e2a"),
    (22, "opc_noop_save", "08dc9000ef567838ec5c6d9121cdb35bccb68110c38d1c68eb308e3f060dee40"),
    (23, "opc_mutated_save", "08dc9000ef567838ec5c6d9121cdb35bccb68110c38d1c68eb308e3f060dee40"),
    (40, "opc_noop_save", "879fd74fed6342b1a9f3f695e550e740bd34e42554961edf490e6efef7a93ee4"),
    (41, "opc_mutated_save", "879fd74fed6342b1a9f3f695e550e740bd34e42554961edf490e6efef7a93ee4"),
    (58, "opc_noop_save", "183178dec5b0fd578e5af04279032368598eec79da7caf0441fc979ce8fc14a0"),
    (59, "opc_mutated_save", "183178dec5b0fd578e5af04279032368598eec79da7caf0441fc979ce8fc14a0"),
    (76, "opc_noop_save", "c34438f16692121822ee71e6cd2103777f9d3c552eaf1c201db37a5c516ba4fb"),
    (77, "opc_mutated_save", "c34438f16692121822ee71e6cd2103777f9d3c552eaf1c201db37a5c516ba4fb"),
    (94, "opc_noop_save", "a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6"),
    (95, "opc_mutated_save", "a0c1af9e2c7a19148b44fc2a8c594c7a274131d74f9f042d55b487d5337cd1e6"),
    (112, "opc_noop_save", "c89c13841d987233e00cbf0fd9f40f35a56c463c5eb3aadd551fce1e008a9d35"),
    (113, "opc_mutated_save", "c89c13841d987233e00cbf0fd9f40f35a56c463c5eb3aadd551fce1e008a9d35"),
    (130, "opc_noop_save", "b8ed826dc5627314aa1f2442b6313158ef48bab7622dce0994b94afe8cc2cfd7"),
    (131, "opc_mutated_save", "b8ed826dc5627314aa1f2442b6313158ef48bab7622dce0994b94afe8cc2cfd7"),
    (159, "xlsx_noop_commit_save", "69ef199769a316eaa465a41ebf08f7a1b501f708775fabd7a084a90dc6a9b428"),
    (161, "xlsx_one_cell_commit_save", "69ef199769a316eaa465a41ebf08f7a1b501f708775fabd7a084a90dc6a9b428"),
    (163, "xlsx_one_percent_commit_save", "69ef199769a316eaa465a41ebf08f7a1b501f708775fabd7a084a90dc6a9b428"),
    (174, "xlsx_noop_commit_save", "9574867b4f1ab4d30ce150de32d2a0b01267d15399ec9edd2c0d57b4bc60fab6"),
    (176, "xlsx_one_cell_commit_save", "9574867b4f1ab4d30ce150de32d2a0b01267d15399ec9edd2c0d57b4bc60fab6"),
    (178, "xlsx_one_percent_commit_save", "9574867b4f1ab4d30ce150de32d2a0b01267d15399ec9edd2c0d57b4bc60fab6"),
    (189, "xlsx_noop_commit_save", "5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714"),
    (191, "xlsx_one_cell_commit_save", "5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714"),
    (193, "xlsx_one_percent_commit_save", "5dd3ad701eb686f6d2d14e9f177a4e9433445728b57b484d53f663b2f87a7714"),
)


def load_json(path: Path) -> Any:
    if path.suffix == ".gz":
        with gzip.open(path, "rt", encoding="utf-8") as stream:
            return json.load(stream)
    with path.open(encoding="utf-8") as stream:
        return json.load(stream)


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    opener = gzip.open if path.suffix == ".gz" else open
    with opener(path, "wt", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, ensure_ascii=False, allow_nan=False)
        stream.write("\n")


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def sha256_json_bytes(path: Path) -> str:
    if path.suffix == ".gz":
        with gzip.open(path, "rb") as stream:
            data = stream.read()
    else:
        data = path.read_bytes()
    return sha256_bytes(data)


def canonical_sha256(value: Any) -> str:
    encoded = json.dumps(
        value, ensure_ascii=False, allow_nan=False, sort_keys=True, separators=(",", ":")
    ).encode("utf-8")
    return sha256_bytes(encoded)


def row_identity(index: int, row: dict[str, Any]) -> tuple[int, str, str]:
    corpus = row.get("corpus")
    if not isinstance(corpus, dict) or not isinstance(corpus.get("archive_sha256"), str):
        raise ValueError(f"results[{index}] has no corpus archive SHA-256")
    case = row.get("case")
    if not isinstance(case, str):
        raise ValueError(f"results[{index}] has no case")
    return index, case, corpus["archive_sha256"]


def mismatched_operation_rows(report: dict[str, Any]) -> list[tuple[int, str, str]]:
    results = report.get("results")
    if not isinstance(results, list):
        raise ValueError("report.results is not a list")
    mismatches: list[tuple[int, str, str]] = []
    for index, row in enumerate(results):
        if not isinstance(row, dict) or "operation_metrics" not in row:
            continue
        metrics = row["operation_metrics"]
        if not isinstance(metrics, dict):
            raise ValueError(f"results[{index}].operation_metrics is not an object")
        sample_indices = metrics.get("sample_indices")
        elapsed = row.get("elapsed_ns")
        if not isinstance(elapsed, dict) or "sample_order" not in elapsed:
            raise ValueError(f"results[{index}] has no elapsed sample order")
        sample_order = elapsed["sample_order"]
        if not isinstance(sample_indices, list) or not isinstance(sample_order, list):
            raise ValueError(f"results[{index}] has malformed sample vectors")
        if sample_indices != sample_order:
            mismatches.append(row_identity(index, row))
    return mismatches


def check_raw_shape(report: dict[str, Any]) -> list[tuple[int, str, str]]:
    results = report.get("results")
    if not isinstance(results, list) or len(results) != RAW_RESULT_COUNT:
        raise ValueError(f"expected {RAW_RESULT_COUNT} normal results")
    operation_count = sum("operation_metrics" in row for row in results)
    if operation_count != RAW_OPERATION_METRICS_COUNT:
        raise ValueError(f"expected {RAW_OPERATION_METRICS_COUNT} operation-metrics rows")
    actual = mismatched_operation_rows(report)
    if actual != list(EXPECTED_REMOVALS):
        raise ValueError(
            "raw mismatch identities differ from the frozen 25-row scope:\n"
            f"expected={list(EXPECTED_REMOVALS)!r}\nactual={actual!r}"
        )
    return actual


def derive(report: dict[str, Any]) -> tuple[dict[str, Any], list[tuple[int, str, str]]]:
    removed = check_raw_shape(report)
    derived = copy.deepcopy(report)
    for index, _, _ in removed:
        derived["results"][index].pop("operation_metrics")
    return derived, removed


def preserved_projection(report: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        {
            "index": index,
            "case": row["case"],
            "corpus_archive_sha256": row["corpus"]["archive_sha256"],
            "elapsed_ns": row["elapsed_ns"],
            "sink": row.get("sink"),
        }
        for index, row in enumerate(report["results"])
    ]


def verify(raw: dict[str, Any], derived: dict[str, Any]) -> dict[str, Any]:
    removed = check_raw_shape(raw)
    if raw.keys() != derived.keys():
        raise ValueError("derived report changed top-level keys")
    for key in raw:
        if key != "results" and raw[key] != derived[key]:
            raise ValueError(f"derived report changed top-level field {key!r}")

    raw_results = raw["results"]
    derived_results = derived.get("results")
    if not isinstance(derived_results, list) or len(derived_results) != RAW_RESULT_COUNT:
        raise ValueError("derived report changed result count")
    removed_indexes = {index for index, _, _ in removed}
    for index, (raw_row, derived_row) in enumerate(zip(raw_results, derived_results)):
        expected = copy.deepcopy(raw_row)
        if index in removed_indexes:
            expected.pop("operation_metrics")
        if derived_row != expected:
            raise ValueError(f"derived row {index} differs beyond the allowed removal")
        if derived_row["elapsed_ns"] != raw_row["elapsed_ns"]:
            raise ValueError(f"derived row {index} changed elapsed timing")
        if derived_row.get("sink") != raw_row.get("sink"):
            raise ValueError(f"derived row {index} changed top-level sink")

    derived_operation_count = sum("operation_metrics" in row for row in derived_results)
    if derived_operation_count != RAW_OPERATION_METRICS_COUNT - REMOVED_OPERATION_METRICS_COUNT:
        raise ValueError("derived operation-metrics count is unexpected")
    projection = preserved_projection(raw)
    if canonical_sha256(projection) != canonical_sha256(preserved_projection(derived)):
        raise ValueError("timing/sink preservation projection changed")
    return {
        "schema_version": 1,
        "transformation": "remove_mismatched_operation_metrics_only",
        "raw_result_count": len(raw_results),
        "raw_operation_metrics_count": RAW_OPERATION_METRICS_COUNT,
        "removed_operation_metrics_count": len(removed),
        "derived_operation_metrics_count": derived_operation_count,
        "removed_rows": [
            {"index": index, "case": case, "corpus_archive_sha256": archive_sha}
            for index, case, archive_sha in removed
        ],
        "preserved_timing_and_sink_projection_sha256": canonical_sha256(projection),
        "non_removed_rows_unchanged": True,
    }


def command_derive(args: argparse.Namespace) -> None:
    raw = load_json(args.raw)
    if not isinstance(raw, dict):
        raise ValueError("raw report is not an object")
    derived, removed = derive(raw)
    write_json(args.derived, derived)
    verification = verify(raw, load_json(args.derived))
    verification.update(
        {
            "raw_file_sha256": sha256_file(args.raw),
            "raw_json_sha256": sha256_json_bytes(args.raw),
            "derived_file_sha256": sha256_file(args.derived),
            "derived_json_sha256": sha256_json_bytes(args.derived),
        }
    )
    if args.receipt is not None:
        write_json(args.receipt, verification)
    print(json.dumps(verification, sort_keys=True))
    print(f"removed operation_metrics envelopes: {len(removed)}")


def command_verify(args: argparse.Namespace) -> None:
    raw = load_json(args.raw)
    derived = load_json(args.derived)
    if not isinstance(raw, dict) or not isinstance(derived, dict):
        raise ValueError("reports must be objects")
    verification = verify(raw, derived)
    verification.update(
        {
            "raw_file_sha256": sha256_file(args.raw),
            "raw_json_sha256": sha256_json_bytes(args.raw),
            "derived_file_sha256": sha256_file(args.derived),
            "derived_json_sha256": sha256_json_bytes(args.derived),
        }
    )
    if args.receipt is not None:
        write_json(args.receipt, verification)
    print(json.dumps(verification, indent=2, sort_keys=True))


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    for command, function in (("derive", command_derive), ("verify", command_verify)):
        subparser = subparsers.add_parser(command)
        subparser.add_argument("--raw", type=Path, required=True)
        subparser.add_argument("--derived", type=Path, required=True)
        subparser.add_argument("--receipt", type=Path)
        subparser.set_defaults(function=function)
    args = parser.parse_args()
    try:
        args.function(args)
    except (OSError, ValueError, json.JSONDecodeError) as error:
        parser.error(str(error))


if __name__ == "__main__":
    main()
