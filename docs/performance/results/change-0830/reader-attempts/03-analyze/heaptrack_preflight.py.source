#!/usr/bin/env python3
"""Independent synthetic preflight for :mod:`heaptrack_reader`.

The cases exercise the indexing and rejection boundaries that matter for the
082? heap attribution: sized UTF-8 strings, one-based string/IP/trace IDs,
zero-based allocation descriptors, exact DSO ownership, Rust hash matching,
duplicate owner frames, supported controls, invalid references, unknown
records, cycles, and corrupt string lengths.  The script writes only a small
JSON receipt to stdout when invoked by the coordinator.
"""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import sys
from typing import Callable


HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("heaptrack_reader_0830", HERE / "heaptrack_reader.py")
if SPEC is None or SPEC.loader is None:
    raise RuntimeError("cannot load heaptrack_reader.py")
reader = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = reader
SPEC.loader.exec_module(reader)


OWNER = "xlsx_edit_profile_0830::edit_region_0830"
BINARY = "/tmp/litchi/bin/xlsx-edit-profile-fp"


def _sized(text: str) -> str:
    return f"s {len(text.encode('utf-8')):x} {text}"


def _valid_stream() -> bytes:
    lines = [
        "v 10500 3",
        "X /tmp/xlsx-edit-profile-fp --synthetic",
        "I 1000 20",
        "c 1",
        "R 100",
        "A",
        "S suppression-pattern",
        _sized(BINARY),
        _sized(OWNER),
        _sized("caller::outer"),
        _sized("fixture.rs"),
        _sized(f"{OWNER}::h0123456789abcdef"),
        _sized("/other/lib.so"),
        _sized("fíxture-λ.rs"),
        "i 1 1 2 7 1",
        "i 2 1 3 7 2",
        "i 3 1 5 7 3",
        "i 4 6 2 7 4",
        "i 5 0",
        "t 1 0",
        "t 2 1",
        "t 3 0",
        "t 4 0",
        "t 5 0",
        "t 1 1",
        "a 10 1",
        "a 20 2",
        "a 30 3",
        "a 40 4",
        "a 50 5",
        "a 60 6",
        "a 70 0",  # a zero trace ID is a valid unresolved descriptor
        "+ 0",
        "+ 1",
        "+ 2",
        "+ 3",
        "+ 4",
        "+ 5",
        "c 2",
        "- 0",
        "- 1",
        "- 2",
        "- 3",
        "- 4",
        "- 5",
    ]
    return ("\n".join(lines) + "\n").encode("utf-8")


def _expect_error(name: str, payload: bytes, failures: dict[str, str]) -> None:
    try:
        reader.parse(payload, OWNER, BINARY)
    except reader.HeaptrackError as error:
        failures[name] = type(error).__name__
        return
    raise AssertionError(f"{name}: malformed stream was accepted")


def _replace_once(payload: bytes, old: bytes, new: bytes) -> bytes:
    if payload.count(old) != 1:
        raise AssertionError(f"fixture replacement is not unique: {old!r}")
    return payload.replace(old, new, 1)


def run_preflight() -> dict[str, object]:
    valid = _valid_stream()
    first = reader.parse(valid, OWNER, BINARY)
    second = reader.parse(valid, OWNER, BINARY)
    if first != second:
        raise AssertionError("parser output is not deterministic")

    whole = first["whole"]
    owner = first["owner_attribution"]
    if whole["calls"] != 6 or whole["requested_bytes"] != 0x150:
        raise AssertionError(f"whole totals mismatch: {whole!r}")
    if owner["calls"] != 3 or owner["requested_bytes"] != 0x60:
        raise AssertionError(f"owner totals mismatch: {owner!r}")
    if first["records"]["allocation_events"] != 6:
        raise AssertionError("+ records were not counted one-for-one")
    if first["diagnostics"]["duplicate_owner_frames"]["allocation_calls"] != 1:
        raise AssertionError("duplicate owner frames were not diagnosed")
    if first["diagnostics"]["ambiguous_owner"]["allocation_calls"] != 1:
        raise AssertionError("duplicate owner frame was not excluded as ambiguous")
    if first["diagnostics"]["owner_symbol_in_wrong_dso"]["allocation_calls"] != 1:
        raise AssertionError("wrong-DSO owner was not diagnosed")
    if first["diagnostics"]["rust_hash_owner_matches"]["allocation_calls"] != 1:
        raise AssertionError("strict Rust hash owner was not recognized")
    if first["diagnostics"]["unknown_frames"]["missing_function"]["allocation_calls"] != 1:
        raise AssertionError("zero/unresolved instruction pointer was not diagnosed")
    if first["diagnostics"]["unknown_frames"]["missing_file"]["allocation_calls"] != 1:
        raise AssertionError("missing source file was not diagnosed")
    if not any(
        row["function"] == OWNER and row["file"] == "fíxture-λ.rs"
        for row in first["leaf"]["rows"]
    ):
        raise AssertionError("UTF-8 sized string was not decoded")
    if not first["caller_trace"]["inclusive"] or not first["caller_trace"]["non_additive"]:
        raise AssertionError("inclusive caller projection is not explicitly non-additive")

    failures: dict[str, str] = {}
    _expect_error(
        "missing_instruction_pointer",
        _replace_once(valid, b"t 1 0\n", b"t 9 0\n"),
        failures,
    )
    _expect_error(
        "missing_trace_parent",
        _replace_once(valid, b"t 1 0\n", b"t 1 9\n"),
        failures,
    )
    _expect_error(
        "missing_allocation_descriptor",
        _replace_once(valid, b"+ 0\n", b"+ 9\n"),
        failures,
    )
    _expect_error(
        "missing_string_reference",
        _replace_once(valid, b"i 1 1 2 7 1\n", b"i 1 9 2 7 1\n"),
        failures,
    )
    _expect_error(
        "unknown_record",
        _replace_once(valid, b"c 1\n", b"q 1\n"),
        failures,
    )
    _expect_error(
        "corrupt_sized_string",
        _replace_once(valid, _sized(BINARY).encode() + b"\n", b"s 1 " + BINARY.encode() + b"\n"),
        failures,
    )
    _expect_error(
        "trace_cycle",
        _replace_once(valid, b"t 1 0\n", b"t 1 1\n"),
        failures,
    )
    _expect_error(
        "backwards_timestamp",
        _replace_once(valid, b"c 2\n", b"c 0\n"),
        failures,
    )

    wrong_dso = _replace_once(
        valid,
        _sized(BINARY).encode() + b"\n",
        _sized("/tmp/otherx/bin/xlsx-edit-profile-fp").encode() + b"\n",
    )
    wrong_dso_report = reader.parse(wrong_dso, OWNER, BINARY)
    if wrong_dso_report["owner_attribution"]["calls"] != 0:
        raise AssertionError("module comparison was not exact")

    upper_hash = _replace_once(
        valid,
        f"{OWNER}::h0123456789abcdef".encode(),
        f"{OWNER}::h0123456789ABCDEf".encode(),
    )
    upper_report = reader.parse(upper_hash, OWNER, BINARY)
    if upper_report["diagnostics"]["rust_hash_owner_matches"]["allocation_calls"] != 0:
        raise AssertionError("non-strict Rust hash suffix matched")

    malformed_utf8 = valid.replace(
        _sized("fíxture-λ.rs").encode(), b"s 1 \xff", 1
    )
    _expect_error("invalid_utf8", malformed_utf8, failures)

    return {
        "schema": "litchi.heaptrack.0830.preflight.v1",
        "valid": {
            "whole_calls": whole["calls"],
            "whole_requested_bytes": whole["requested_bytes"],
            "owner_calls": owner["calls"],
            "owner_requested_bytes": owner["requested_bytes"],
            "unicode": "fíxture-λ.rs",
            "zero_based_allocation_descriptor": True,
            "one_based_string_ip_trace_ids": True,
        },
        "negative_cases": sorted(failures),
        "negative_case_error_types": failures,
    }


def main() -> int:
    receipt = run_preflight()
    print(json.dumps(receipt, ensure_ascii=False, sort_keys=True, separators=(",", ":")))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
