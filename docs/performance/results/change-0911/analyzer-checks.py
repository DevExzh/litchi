#!/usr/bin/env python3
"""In-memory fail-closed mutation checks for the frozen 0911 reducer.

This review helper never changes a retained packet artifact.  It deep-copies
each parsed Callgrind object, injects one contradiction, and requires the
frozen analyzer to reject that object.  The output is create-only.
"""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
from pathlib import Path
from typing import Any, Callable


HERE = Path(__file__).resolve().parent
ANALYZER_PATH = HERE / "analyze.py"
ANALYSIS_PATH = HERE / "analysis.json"
OUTPUT_PATH = HERE / "analyzer-checks.json"
OWNER = "litchi_pptx::package::model::Package::opened_presentation_transaction"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_analyzer() -> Any:
    spec = importlib.util.spec_from_file_location("frozen_analyzer_0911", ANALYZER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load frozen 0911 analyzer")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def owner_id(parsed: dict[str, Any]) -> int:
    ids = [
        function_id
        for function_id, function in parsed["functions"].items()
        if function.get("name") == OWNER
    ]
    if len(ids) != 1:
        raise RuntimeError(f"expected one owner in mutation fixture, got {ids}")
    return ids[0]


def mutate_unknown_caller(path: Path, parsed: dict[str, Any]) -> None:
    if path.name != "callgrind.out.1":
        return
    incoming = parsed["functions"]
    target = owner_id(parsed)
    edges = [
        edge
        for function in incoming.values()
        for edge in function["edges"]
        if edge["callee_id"] == target and edge["calls"] > 0
    ]
    if len(edges) != 1:
        raise RuntimeError(f"expected one incoming edge, got {edges}")
    parsed["names"][edges[0]["caller_id"]] = "synthetic::unknown_owner_caller"


def mutate_extra_owner(path: Path, parsed: dict[str, Any]) -> None:
    if path.name != "callgrind.out.1":
        return
    target = owner_id(parsed)
    caller = next(function for function_id, function in parsed["functions"].items() if function_id != target)
    caller["edges"].append(
        {
            "caller_id": caller["id"],
            "callee_id": target,
            "calls": 1,
            "inclusive_ir": 1,
            "line": 0,
            "position": "0",
            "file": caller.get("file", "???"),
        }
    )


def mutate_counter_mismatch(path: Path, parsed: dict[str, Any]) -> None:
    if path.name == "callgrind.out.1":
        parsed["header"]["totals_ir"] += 1


def mutate_nonzero_terminal(path: Path, parsed: dict[str, Any]) -> None:
    if path.name == "callgrind.out":
        parsed["header"]["summary_ir"] = 1
        parsed["header"]["totals_ir"] = 1


Mutation = Callable[[Path, dict[str, Any]], None]


def run_case(analyzer: Any, name: str, mutation: Mutation) -> dict[str, Any]:
    original = analyzer.READER.parse_raw

    def wrapped(path: Path, *args: Any, **kwargs: Any) -> dict[str, Any]:
        parsed = copy.deepcopy(original(path, *args, **kwargs))
        mutation(Path(path), parsed)
        return parsed

    analyzer.READER.parse_raw = wrapped
    try:
        analyzer.reduce_packet(False)
    except analyzer.EvidenceError as error:
        return {"name": name, "status": "rejected", "error": str(error)}
    finally:
        analyzer.READER.parse_raw = original
    raise RuntimeError(f"mutation {name} was accepted")


def main() -> int:
    analyzer = load_analyzer()
    cases = [
        ("unknown_caller", mutate_unknown_caller),
        ("extra_owner_edge", mutate_extra_owner),
        ("counter_mismatch", mutate_counter_mismatch),
        ("nonzero_terminal", mutate_nonzero_terminal),
    ]
    results = [run_case(analyzer, name, mutation) for name, mutation in cases]
    output = {
        "schema": "litchi.pptx-snapshot-analyzer-mutation-checks-0911.v1",
        "status": "pass",
        "scope": "All mutations exist only in deep-copied in-memory Callgrind parse trees; retained packet files are never edited.",
        "analyzer": descriptor(ANALYZER_PATH),
        "canonical_analysis": descriptor(ANALYSIS_PATH),
        "cases": results,
        "checks": {
            "unknown_caller_rejected": True,
            "extra_owner_edge_rejected": True,
            "counter_mismatch_rejected": True,
            "nonzero_terminal_rejected": True,
        },
    }
    with OUTPUT_PATH.open("x", encoding="utf-8") as stream:
        json.dump(output, stream, indent=2, sort_keys=True)
        stream.write("\n")
    print("0911 analyzer mutation checks passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
