#!/usr/bin/env python3
"""Controls for the decoder-side full-lifetime evlist guard."""
from __future__ import annotations

import importlib.util
import json
import re
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0978")
DECODER_PATH = ROOT / "tools/focuseddecode_0978_v3.py"
SCANNER_PATH = ROOT / "tools/raw_stack_0978.py"
EVLIST_PATH = ROOT / "fresh-offline-child02-v1/child-02/child-02.perf-evlist-v.txt"
COMMAND_PATH = ROOT / "fresh-offline-child02-v1/child-02/child-02.perf-evlist-v.command.json"
RAW_PATH = ROOT / "fresh-offline-child02-v1/child-02/child-02.perf.data"
PACKET_PATH = ROOT / "fresh-offline-child02-v1/child-02/child-02.packet.json"
LABELS = ("comm", "comm_exec", "mmap_data", "mmap2", "build_id")


def load_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def mutate(line: str, label: str, mode: str) -> str:
    token = rf"(?P<prefix>{re.escape(label)}:\s*)1(?P<suffix>(?=,\s*|$))"
    if mode == "missing":
        mutated, count = re.subn(rf",\s*{re.escape(label)}:\s*1(?=,\s*|$)", "", line, count=1)
    elif mode == "zero":
        mutated, count = re.subn(token, rf"\g<prefix>0\g<suffix>", line, count=1)
    elif mode == "duplicate":
        mutated, count = re.subn(token, rf"\g<prefix>1\g<suffix>, {label}: 1", line, count=1)
    else:
        raise AssertionError(mode)
    assert count == 1, (mode, label, line)
    return mutated


def main() -> int:
    decoder = load_module("focuseddecode_full_lifetime_controls", DECODER_PATH)
    scanner = load_module("raw_stack_full_lifetime_controls", SCANNER_PATH)
    text = EVLIST_PATH.read_text(encoding="utf-8", errors="strict")
    contract = scanner.parse_evlist_contract(text, expected_period=1_000_000, expected_stack_cap=32_768)
    packet = json.loads(PACKET_PATH.read_text(encoding="utf-8"))
    reducer_contract = packet["raw_stack"]["event_contract"]
    assert contract == reducer_contract
    canonical_before = dict(contract)
    assert decoder._require_full_lifetime_evlist(contract) is None
    assert contract == canonical_before
    assert contract == reducer_contract
    positive = {"status": "pass", "required": {label: "1" for label in LABELS}}

    negatives: list[dict[str, Any]] = []
    line = text.strip()
    for label in LABELS:
        for mode in ("missing", "zero", "duplicate"):
            mutated = mutate(line, label, mode)
            mutated_contract = scanner.parse_evlist_contract(mutated, expected_period=1_000_000, expected_stack_cap=32_768)
            try:
                decoder._require_full_lifetime_evlist(mutated_contract)
            except decoder.DecodeFailure as exc:
                assert exc.code == "evlist_lifetime_contract", (label, mode, exc.as_dict())
                assert exc.details.get("attribute") == label, (label, mode, exc.as_dict())
                negatives.append({"attribute": label, "mutation": mode, "code": exc.code})
            else:
                raise AssertionError((label, mode, "guard accepted mutation"))

    command = json.loads(COMMAND_PATH.read_text(encoding="utf-8"))
    expected_argv = ["/usr/bin/perf", "evlist", "-v", "-i", str(RAW_PATH.absolute())]
    assert command["argv"] == expected_argv
    assert command["returncode"] == 0
    assert command["native_workload_started"] is False
    print(json.dumps({
        "schema": "litchi.performance.0978.full_lifetime_evlist_controls.v1",
        "status": "pass",
        "positive": {"status": positive["status"], "required": LABELS, "contract_unchanged": True, "reducer_contract_equal": True},
        "negative_count": len(negatives),
        "negatives": negatives,
        "command_generation": {"status": "unchanged_bound_receipt", "argv": command["argv"], "returncode": command["returncode"], "native_workload_started": command["native_workload_started"]},
    }, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
