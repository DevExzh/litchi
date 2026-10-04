#!/usr/bin/env python3
"""Targeted offline checks for the 0908 Callgrind reducer.

This helper reads the retained packet and writes only inside a temporary
directory.  It never launches the workload, a profiler, or a build.
"""

from __future__ import annotations

import argparse
import copy
import importlib.util
import json
from pathlib import Path
import sys
import tempfile


sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent
ANALYZER_PATH = HERE / "analyze.py"
RAW_PATH = HERE / "capture" / "eager-a" / "callgrind.out.1"
TERMINAL_PATH = HERE / "capture" / "eager-a" / "callgrind.out"
REFERENCE_PATH = HERE.parent / "change-0907" / "eager-a.json"


class CheckError(RuntimeError):
    """A checker invariant failed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CheckError(message)


def load_analyzer():
    spec = importlib.util.spec_from_file_location("analyzer_0908_check", ANALYZER_PATH)
    require(spec is not None and spec.loader is not None, "cannot load analyze.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def read_json(path: Path):
    require(path.is_file() and not path.is_symlink(), f"missing regular JSON file: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def expect_rejection(call, label: str,
                     exception_type: type[BaseException] | tuple[type[BaseException], ...]) -> None:
    try:
        call()
    except exception_type:
        return
    except Exception as error:
        raise CheckError(f"negative mutation raised the wrong exception for {label}: {error}") from error
    raise CheckError(f"negative mutation was accepted: {label}")


def fresh_output_equals_canonical(analyzer, canonical: Path) -> None:
    require(canonical.is_file() and not canonical.is_symlink(),
            f"missing canonical analysis output: {canonical}")
    canonical_bytes = canonical.read_bytes()
    with tempfile.TemporaryDirectory(prefix="litchi-0908-analyzer-check-") as directory:
        fresh = Path(directory) / "analysis.json"
        analyzer.analyze(fresh, first_only=False)
        require(fresh.read_bytes() == canonical_bytes,
                "fresh offline analysis differs byte-for-byte from canonical analysis")


def parse_raw(analyzer):
    require(RAW_PATH.is_file() and not RAW_PATH.is_symlink(),
            f"missing retained first raw dump: {RAW_PATH}")
    return analyzer.READER.parse_raw(RAW_PATH)


def reject_wrong_owner_callsite(analyzer, parsed) -> None:
    mutated = copy.deepcopy(parsed)
    owner_ids = [function_id for function_id, name in mutated["names"].items()
                 if name == analyzer.OWNER]
    require(len(owner_ids) == 1, "first raw owner is not unique")
    incoming = analyzer.READER.incoming_edges(mutated, owner_ids[0])
    require(len(incoming) == 1, "first raw owner edge is not unique")
    caller_id = incoming[0]["caller_id"]
    edges = mutated["functions"][caller_id]["edges"]
    matching = [edge for edge in edges if edge["callee_id"] == owner_ids[0]]
    require(len(matching) == 1, "first raw owner edge cannot be located for mutation")
    matching[0]["position"] = analyzer.EXPECTED_CALLSITE
    expect_rejection(lambda: analyzer._owner_row(mutated, "negative-owner-callsite"),
                     "wrong measured owner callsite", analyzer.EvidenceError)


def reject_ir_conservation(analyzer, parsed) -> None:
    mutated = copy.deepcopy(parsed)
    owner_ids = [function_id for function_id, name in mutated["names"].items()
                 if name == analyzer.OWNER]
    require(len(owner_ids) == 1, "first raw owner is not unique")
    mutated["functions"][owner_ids[0]]["self_ir"] += 1
    expect_rejection(lambda: analyzer._owner_row(mutated, "negative-ir-conservation"),
                     "owner Ir conservation", analyzer.EvidenceError)


def reject_terminal_nonzero(analyzer, parsed) -> None:
    require(TERMINAL_PATH.is_file() and not TERMINAL_PATH.is_symlink(),
            f"missing retained terminal dump: {TERMINAL_PATH}")
    source = TERMINAL_PATH.read_text(encoding="utf-8")
    require("summary: 0" in source and "totals: 0" in source,
            "retained terminal does not contain zero summary/totals")
    mutated = source.replace("summary: 0", "summary: 1", 1).replace("totals: 0", "totals: 1", 1)
    with tempfile.TemporaryDirectory(prefix="litchi-0908-terminal-negative-") as directory:
        path = Path(directory) / "callgrind.out"
        path.write_text(mutated, encoding="utf-8")
        command = parsed["header"]["command"]
        expect_rejection(
            lambda: analyzer.parse_terminal(path, 1, command, "negative-terminal"),
            "nonzero terminal dump",
            (analyzer.EvidenceError, analyzer.READER.ReaderError),
        )


def reject_native_nonconstant(analyzer) -> None:
    reference = read_json(REFERENCE_PATH)
    values = copy.deepcopy(reference["results"][0]["source"]["read_calls"])
    require(isinstance(values, list) and len(values) == 200,
            "native reference read_calls vector is not 200 samples")
    values[1] += 1
    expect_rejection(lambda: analyzer._one_native_vector(values, "negative-native-vector"),
                     "nonconstant native 200-sample vector", analyzer.EvidenceError)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--canonical", type=Path, default=HERE / "analysis.json",
                        help="canonical analysis JSON to compare with a fresh replay")
    args = parser.parse_args()
    try:
        analyzer = load_analyzer()
        fresh_output_equals_canonical(analyzer, args.canonical)
        parsed = parse_raw(analyzer)
        reject_wrong_owner_callsite(analyzer, parsed)
        reject_ir_conservation(analyzer, parsed)
        reject_terminal_nonzero(analyzer, parsed)
        reject_native_nonconstant(analyzer)
    except (CheckError, OSError, TypeError, ValueError, json.JSONDecodeError) as error:
        print(f"check-analyzer.py: error: {error}", file=sys.stderr)
        return 2
    print("0908 analyzer checks passed: fresh equality plus four fail-closed negative mutations")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
