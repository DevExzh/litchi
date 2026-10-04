"""Run fail-closed negative checks against in-memory GDB dispatch copies.

No retained dispatch, workload, plan, or analysis file is modified.  The
script reads the canonical witnesses, deep-copies each JSON object, injects one
contradiction at a time, and requires the same scope invariants to reject it.
"""

from __future__ import annotations

import argparse
import copy
import importlib.util
import json
from pathlib import Path
from typing import Any, Callable


PACKET = Path(__file__).resolve().parent
VERIFY = PACKET / "verify_dispatch.py"
ANALYSIS = PACKET / "dispatch-analysis.json"
SYMBOL_RANGES = PACKET / "symbol-ranges.json"
OWNER = "litchi_pptx::package::model::Package::opened_presentation_transaction"
PARENT = "litchi_perf_baseline::pptx_owned_benchmark::timed_iteration"
SETUP = "litchi_perf_baseline::build_pptx_source_edit_corpus"
HARDWARE = "sha2::sha256::x86_sha::compress"
SOFTWARE = "sha2::sha256::soft::unroll::compress"


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def load_verify_module():
    spec = importlib.util.spec_from_file_location("verify_dispatch_0912_checks", VERIFY)
    require(spec is not None and spec.loader is not None, "cannot load dispatch verifier")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def validate_state(dispatch: dict[str, Any], symbol_rows: dict[str, dict[str, Any]], label: str) -> None:
    """Validate the GDB state-machine invariants used by the negative tests."""

    require(dispatch.get("status") == "pass", f"{label}: dispatch status is not pass")
    require(dispatch.get("errors") == [], f"{label}: dispatch errors are present")
    require(dispatch.get("exit_code") == 0, f"{label}: target exit code is not zero")
    bias = dispatch.get("mapping_load_bias")
    require(isinstance(bias, int) and bias > 0, f"{label}: missing mapping bias")
    runtime = dispatch.get("runtime_symbols")
    require(set(runtime) == set(symbol_rows), f"{label}: runtime symbol set changed")
    for name, row in symbol_rows.items():
        value = runtime[name]
        require(value["file_address"] == row["address"], f"{label}: {name} file address changed")
        require(value["size"] == row["size"], f"{label}: {name} size changed")
        require(value["load_bias"] == bias, f"{label}: {name} load bias changed")
        require(value["start"] == bias + row["address"], f"{label}: {name} runtime start changed")
        require(value["end"] == bias + row["end"], f"{label}: {name} runtime end changed")

    entries = dispatch.get("owner_entries")
    returns = dispatch.get("owner_returns")
    require(isinstance(entries, list) and len(entries) == 2, f"{label}: owner entry count changed")
    require(isinstance(returns, list) and len(returns) == 2, f"{label}: owner return count changed")
    require([row["context"] for row in entries] == ["setup", "measured"], f"{label}: owner entry order changed")
    require([row["context"] for row in returns] == ["setup", "measured"], f"{label}: owner return order changed")
    require([row["index"] for row in entries] == [0, 1], f"{label}: owner entry indexes changed")
    require([row["owner_entry"] for row in returns] == [0, 1], f"{label}: owner return indexes changed")
    expected_parents = {"setup": SETUP, "measured": PARENT}
    for entry in entries:
        context = entry["context"]
        parent = symbol_rows[expected_parents[context]]
        return_offset = entry["return_address"] - bias
        require(parent["address"] <= return_offset < parent["end"], f"{label}: {context} return caller is outside expected range")
        require(entry["entry_pc"] == runtime[OWNER]["start"], f"{label}: owner entry PC changed")
    for entry, row in zip(entries, returns):
        require(row["context"] == entry["context"] and row["owner_entry"] == entry["index"], f"{label}: return guard does not match entry")
        require(row["pc"] == entry["return_address"], f"{label}: return guard PC mismatch")
        require(row["rsp"] == entry["entry_rsp"] + 8, f"{label}: return guard RSP mismatch")
        require(row["backend_breakpoints_disabled"] is True, f"{label}: return leaves backend breakpoint enabled")

    events = dispatch.get("events")
    require(isinstance(events, list) and [event.get("kind") for event in events] == ["owner_entry", "owner_return", "owner_entry", "backend", "owner_return"], f"{label}: event order changed")
    require(events[0] == {"backend_breakpoints_enabled": {"hardware": False, "software": False}, "context": "setup", "kind": "owner_entry", "owner_entry": 0}, f"{label}: setup enable state changed")
    require(events[1] == {"backend_breakpoints_disabled": True, "context": "setup", "kind": "owner_return", "owner_entry": 0}, f"{label}: setup return state changed")
    require(events[2] == {"backend_breakpoints_enabled": {"hardware": True, "software": True}, "context": "measured", "kind": "owner_entry", "owner_entry": 1}, f"{label}: measured enable state changed")
    require(events[3] == {"backend": "hardware", "breakpoint_disabled_after_hit": True, "kind": "backend", "owner_entry": 1}, f"{label}: backend disable state changed")
    require(events[4] == {"backend_breakpoints_disabled": True, "context": "measured", "kind": "owner_return", "owner_entry": 1}, f"{label}: measured return state changed")

    hits = dispatch.get("backend_hits")
    require(isinstance(hits, list) and len(hits) == 1, f"{label}: backend hit count changed")
    hit = hits[0]
    require(hit["kind"] == "hardware" and hit["symbol"] == HARDWARE, f"{label}: backend symbol changed")
    require(hit["owner_entry"] == 1, f"{label}: backend hit owner changed")
    require(runtime[HARDWARE]["start"] <= hit["pc"] < runtime[HARDWARE]["end"], f"{label}: backend PC outside frozen range")
    require(hit["pc"] == runtime[HARDWARE]["start"], f"{label}: backend PC is not exact breakpoint")
    require(hit["frame_scope_available"] is True and hit["unwind_error"] is None, f"{label}: backend frame scope unavailable")
    names = [frame.get("name") for frame in hit["frames"]]
    require(names.count(HARDWARE) == 1 and names.count(OWNER) == 1 and names.count(PARENT) == 1, f"{label}: backend frame scope changed")
    require(SOFTWARE not in names and SETUP not in names, f"{label}: unexpected backend/setup frame")
    require(names.index(OWNER) < names.index(PARENT), f"{label}: owner/parent frame order changed")
    require(OWNER in hit["backtrace"] and PARENT in hit["backtrace"], f"{label}: textual backtrace scope changed")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=PACKET / "dispatch-checks.json")
    args = parser.parse_args()

    verifier = load_verify_module()
    symbol_data = read_json(SYMBOL_RANGES)
    symbol_rows = verifier.verify_symbols(symbol_data)
    canonical = {
        "materialized": read_json(PACKET / "dispatch" / "materialized" / "dispatch.json"),
        "owned": read_json(PACKET / "dispatch" / "owned" / "dispatch.json"),
    }
    for kind, dispatch in canonical.items():
        validate_state(dispatch, symbol_rows, f"canonical {kind}")

    mutators: dict[str, Callable[[dict[str, Any]], None]] = {
        "caller_ra_outside_expected_range": lambda d: d["owner_entries"][1].update(
            return_address=d["runtime_symbols"][OWNER]["start"]
        ),
        "backend_pc_outside_frozen_range": lambda d: d["backend_hits"][0].update(
            pc=d["runtime_symbols"][HARDWARE]["end"]
        ),
        "missing_return_guard": lambda d: d["owner_returns"].pop(),
        "mismatched_return_guard": lambda d: d["owner_returns"][1].update(
            pc=d["owner_entries"][0]["return_address"]
        ),
        "out_of_order_enable_disable": lambda d: d["events"].__setitem__(
            slice(2, 4), [d["events"][3], d["events"][2]]
        ),
    }
    cases: list[dict[str, Any]] = []
    for kind, original in canonical.items():
        for name, mutate in mutators.items():
            mutated = copy.deepcopy(original)
            mutate(mutated)
            case = {"name": f"{kind}_{name}", "mutation": name}
            try:
                validate_state(mutated, symbol_rows, case["name"])
            except AssertionError as error:
                case.update(status="rejected", error=str(error))
            else:
                case.update(status="accepted", error=None)
            cases.append(case)

    expected_names = {f"{kind}_{name}" for kind in canonical for name in mutators}
    require({case["name"] for case in cases} == expected_names, "mutation case set changed")
    require(all(case["status"] == "rejected" for case in cases), "a fail-closed mutation was accepted")

    args.output = args.output.resolve()
    analysis = {
        "schema": "litchi.native-sha-dispatch-mutation-checks-0912.v1",
        "status": "pass",
        "scope": "Canonical dispatch JSON is deep-copied in memory. Mutations are never written to retained dispatch, analysis, plan, or workload files.",
        "canonical": {
            "dispatch_analysis": verifier.descriptor(ANALYSIS),
            "symbol_ranges": verifier.descriptor(SYMBOL_RANGES),
            "materialized_dispatch": verifier.descriptor(PACKET / "dispatch" / "materialized" / "dispatch.json"),
            "owned_dispatch": verifier.descriptor(PACKET / "dispatch" / "owned" / "dispatch.json"),
        },
        "canonical_state_pass": True,
        "checks": {
            "caller_ra_outside_expected_range_rejected": True,
            "backend_pc_outside_frozen_range_rejected": True,
            "missing_return_guard_rejected": True,
            "mismatched_return_guard_rejected": True,
            "out_of_order_enable_disable_rejected": True,
        },
        "cases": cases,
        "hash_bindings": {
            "verify_dispatch": verifier.descriptor(VERIFY),
            "dispatch_checks": verifier.descriptor(Path(__file__)),
        },
    }
    args.output.write_text(json.dumps(analysis, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "cases": len(cases), "output": str(args.output)}, sort_keys=True))


if __name__ == "__main__":
    main()

