#!/usr/bin/env python3
"""Reduce and replay the frozen 0941 96-child XLSX packet.

The reducer is independent of capture and never starts a workload.  It
accepts a packet whose absolute paths may have moved, using an optional
``packet-manifest.json`` path map before falling back to the recorded path.
Every child receipt is checked against the frozen cell/order matrix, terminal
receipt, exact source and binary witnesses, RSS sidecar, clean environment,
and probe report before any statistic is emitted.

The output keeps every individual child row.  It computes midpoint medians,
nearest-rank p95/p99, and within-child spread ``(max-min)/min``.  Ratios and
5% flags are descriptive packet review fields; the native p50 adoption rule
and allocator review rules remain the gates frozen in the plan.
"""

from __future__ import annotations

import argparse
from copy import deepcopy
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import sys


ROOT = Path(__file__).resolve().parent
PLAN_PATH = ROOT / "xlsx-plan-v4.json"
RUNNER_PATH = ROOT / "xlsx-capture-v4.py"
DEFAULT_PACKET = ROOT / "captures"


def load_runner():
    spec = importlib.util.spec_from_file_location("xlsx_capture_v3_reducer", RUNNER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot import {RUNNER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


CAP = load_runner()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read(path: Path) -> object:
    try:
        return json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"cannot read JSON {path}: {exc}") from exc


def read_object(path: Path) -> dict:
    value = read(path)
    require(isinstance(value, dict), f"JSON object required: {path}")
    return value


def artifact(path: Path) -> dict:
    require(path.is_file(), f"artifact missing: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def module_hash(path: Path) -> dict:
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def midpoint(values: list[float | int]) -> float | int:
    require(values, "empty statistic")
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2


def nearest_rank(values: list[float | int], percentile: float) -> float | int:
    require(values, "empty percentile")
    ordered = sorted(values)
    rank = max(1, math.ceil(percentile * len(ordered)))
    return ordered[rank - 1]


def spread(values: list[float | int]) -> float | None:
    require(values, "empty spread")
    low = min(values)
    high = max(values)
    if low == 0:
        return 0.0 if high == 0 else None
    return (high - low) / low


def ratio(candidate: float | int | None, baseline: float | int | None) -> float | None:
    if candidate is None or baseline is None:
        return None
    if baseline == 0:
        return 1.0 if candidate == 0 else None
    return candidate / baseline


def ratio_flag(value: float | None) -> bool | None:
    return None if value is None else value > 1.05


class Resolver:
    """Resolve recorded artifact paths against a moved packet root."""

    def __init__(self, packet_root: Path):
        self.packet_root = packet_root.resolve()
        manifest_path = self.packet_root / "packet-manifest.json"
        manifest = read_object(manifest_path) if manifest_path.is_file() else {}
        raw_map = manifest.get("path_map", {})
        require(isinstance(raw_map, dict), "packet path_map must be an object")
        self.path_map = {str(key): self.packet_root / str(value) for key, value in raw_map.items()}

    def path(self, recorded: str | Path, anchor: Path | None = None) -> Path:
        recorded_text = str(recorded)
        mapped = self.path_map.get(recorded_text)
        if mapped is not None:
            candidate = mapped
        else:
            raw = Path(recorded_text)
            if raw.is_absolute():
                candidate = raw
                if not candidate.is_file():
                    candidate = self.packet_root / raw.name
            else:
                candidate = (anchor / raw) if anchor is not None else self.packet_root / raw
                if not candidate.is_file():
                    candidate = self.packet_root / raw
        candidate = candidate.resolve()
        require(candidate.is_file(), f"recorded artifact unavailable after path resolution: {recorded_text}")
        return candidate

    def checked_artifact(self, recorded: dict, anchor: Path | None = None, label: str = "artifact") -> tuple[Path, dict]:
        require(isinstance(recorded, dict), f"{label} reference is not an object")
        path = self.path(recorded.get("path", ""), anchor)
        actual = artifact(path)
        require(actual["bytes"] == recorded.get("bytes") and actual["sha256"] == recorded.get("sha256"), f"{label} digest changed")
        return path, actual


EXPECTED_CELLS = [
    ("generated-xlsx/open", "generated-xlsx", "open", "Sheet1", "primary"),
    ("generated-xlsx/lifecycle", "generated-xlsx", "lifecycle", "Sheet1", "primary"),
    ("conditional-formatting/open", "conditional-formatting", "open", "Home", "primary"),
    ("conditional-formatting/lifecycle", "conditional-formatting", "lifecycle", "Home", "primary"),
    ("small-xlsx/open", "small-xlsx", "open", "Sheet1", "boundary-control"),
    ("boundary-xlsx/open", "boundary-xlsx", "open", "Sheet1", "boundary-control"),
]
EXPECTED_RUNTIME_ENV = {
    "PATH": "/usr/bin:/bin",
    "LANG": "C",
    "LC_ALL": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
}
EXPECTED_UNSET = ["RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "CARGO_HOME", "RUSTUP_TOOLCHAIN", "CARGO_INCREMENTAL"]


def validate_plan_shape(plan: dict) -> None:
    require(plan.get("schema") == "litchi.performance.0941.xlsx-controls-plan.v4", "plan schema changed")
    require(plan.get("status") in {"prospective_only_pending_root_approval", "authorized"}, "plan status changed")
    scope = plan.get("scope", {})
    require(scope.get("cpu") == 12 and scope.get("deadline_seconds") == 900, "scope changed")
    require(scope.get("no_retry_or_replacement") is True and scope.get("terminal_first") is True, "capture policy changed")
    require(scope.get("iwork") == "excluded", "iWork scope changed")
    require(scope.get("native") == {"blocks": 6, "warmup": 3, "samples": 30, "orders": [["baseline", "candidate"], ["candidate", "baseline"], ["baseline", "candidate"], ["candidate", "baseline"], ["candidate", "baseline"], ["baseline", "candidate"]], "children_per_cell": 12, "total_children": 72}, "native matrix changed")
    require(scope.get("allocator") == {"blocks": 2, "warmup": 0, "samples": 3, "orders": [["baseline", "candidate"], ["candidate", "baseline"]], "children_per_cell": 4, "total_children": 24}, "allocator matrix changed")
    require([(cell.get("id"), cell.get("fixture"), cell.get("phase"), cell.get("sheet"), cell.get("group")) for cell in plan.get("cells", [])] == EXPECTED_CELLS, "cell matrix changed")
    require(plan.get("runtime_environment", {}).get("PATH") == EXPECTED_RUNTIME_ENV["PATH"], "runtime PATH changed")
    require(plan.get("runtime_environment", {}).get("LANG") == "C" and plan.get("runtime_environment", {}).get("LC_ALL") == "C" and plan.get("runtime_environment", {}).get("TZ") == "UTC" and plan.get("runtime_environment", {}).get("RUST_BACKTRACE") == "0", "runtime environment changed")
    require(plan.get("runtime_environment", {}).get("unset") == EXPECTED_UNSET, "runtime unset policy changed")


def cell_by_id(plan: dict) -> dict[str, dict]:
    return {cell["id"]: cell for cell in plan["cells"]}


def expected_source(plan: dict, fixture: str) -> dict:
    return plan["fixtures"][fixture]


def resolve_plan_artifact(resolver: Resolver, witness: dict, label: str) -> tuple[Path, dict]:
    return resolver.checked_artifact(witness, label=label)


def validate_terminal(
    resolver: Resolver,
    plan: dict,
    row: dict,
    terminal_ref: dict,
    terminal_path: Path,
    source_expected: dict,
    binary_expected: dict,
) -> dict:
    terminal = read_object(terminal_path)
    lane = row["lane"]
    require(terminal.get("schema") == "litchi.performance.0941.xlsx-child-terminal.v3", f"{row['cell']} terminal schema changed")
    for key in ("block", "cell", "leg", "lane"):
        require(terminal.get(key) == row[key], f"{row['cell']} terminal {key} changed")
    require(terminal.get("command") == row["command"], f"{row['cell']} terminal command changed")
    require(terminal.get("cwd") == plan["source_and_probe"]["legs"][row["leg"]]["probe_root"], f"{row['cell']} terminal cwd changed")
    require(terminal.get("environment") == EXPECTED_RUNTIME_ENV, f"{row['cell']} terminal environment changed")
    require(terminal.get("unset", EXPECTED_UNSET) == EXPECTED_UNSET, f"{row['cell']} terminal unset policy changed")
    require(terminal.get("return_code") == 0 and terminal.get("timed_out") is False and terminal.get("process_group_clear") is True, f"{row['cell']} terminal status changed")
    require(terminal.get("terminal_recorded_before_report_parse") is True, f"{row['cell']} terminal ordering changed")
    binary_before_path, binary_before = resolver.checked_artifact(terminal.get("binary_before", {}), terminal_path.parent, f"{row['cell']} binary before")
    binary_after_path, binary_after = resolver.checked_artifact(terminal.get("binary_after", {}), terminal_path.parent, f"{row['cell']} binary after")
    require(binary_before["bytes"] == binary_expected["bytes"] and binary_before["sha256"] == binary_expected["sha256"], f"{row['cell']} terminal binary before witness changed")
    require(binary_after == binary_before and binary_after_path == binary_before_path, f"{row['cell']} terminal binary changed")
    source_path, source_artifact = resolver.checked_artifact(terminal.get("source", {}), terminal_path.parent, f"{row['cell']} terminal source")
    require(source_artifact["bytes"] == source_expected["bytes"] and source_artifact["sha256"] == source_expected["sha256"], f"{row['cell']} terminal source witness changed")
    script_path, script_artifact = resolver.checked_artifact(terminal.get("script", {}), terminal_path.parent, f"{row['cell']} terminal script")
    require(script_artifact == {"path": script_artifact["path"], "bytes": plan["script_witness"]["bytes"], "sha256": plan["script_witness"]["sha256"]}, f"{row['cell']} terminal script witness changed")
    for name in ("time", "taskset"):
        tool = plan["tool_witnesses"][name]
        _, tool_artifact = resolver.checked_artifact(terminal.get(name, {}), terminal_path.parent, f"{row['cell']} terminal {name}")
        require(tool_artifact["bytes"] == tool["bytes"] and tool_artifact["sha256"] == tool["sha256"], f"{row['cell']} terminal {name} witness changed")
    stdout_path, stdout_artifact = resolver.checked_artifact(terminal.get("stdout", {}), terminal_path.parent, f"{row['cell']} terminal stdout")
    stderr_path, stderr_artifact = resolver.checked_artifact(terminal.get("stderr", {}), terminal_path.parent, f"{row['cell']} terminal stderr")
    rss_path, rss_artifact = resolver.checked_artifact(terminal.get("rss", {}), terminal_path.parent, f"{row['cell']} terminal RSS")
    report_ref = terminal.get("report", {})
    report_path = resolver.path(report_ref.get("path", ""), terminal_path.parent)
    require(report_path == resolver.path(row["report"]["path"]), f"{row['cell']} terminal report binding changed")
    return {"terminal": terminal, "source_path": str(source_path), "binary_path": str(binary_before_path), "stdout_path": str(stdout_path), "stderr_path": str(stderr_path), "rss_path": str(rss_path), "rss_artifact": rss_artifact, "report_path": str(report_path)}


def validate_row(resolver: Resolver, plan: dict, cells_by_id: dict[str, dict], row: dict) -> dict:
    require(row.get("schema") == "litchi.performance.0941.xlsx-child-receipt.v3", f"receipt schema changed: {row}")
    cell = cells_by_id.get(row.get("cell"))
    require(cell is not None, f"unknown cell: {row.get('cell')}")
    lane = row.get("lane")
    require(lane in ("native", "allocator"), f"unknown lane: {lane}")
    leg = row.get("leg")
    require(leg in ("baseline", "candidate"), f"unknown leg: {leg}")
    require(row.get("fixture") == cell["fixture"] and row.get("phase") == cell["phase"], f"row cell binding changed: {row}")
    source_expected = expected_source(plan, cell["fixture"])
    binary_expected = plan["binary_witnesses"][leg]["normal" if lane == "native" else "allocator"]
    source_path, source_artifact = resolver.checked_artifact(row.get("source", {}), label=f"{row['cell']} source")
    require(source_artifact["bytes"] == source_expected["bytes"] and source_artifact["sha256"] == source_expected["sha256"], f"{row['cell']} source digest changed")
    binary_path, binary_artifact = resolver.checked_artifact(row.get("binary", {}), label=f"{row['cell']} binary")
    require(binary_artifact["bytes"] == binary_expected["bytes"] and binary_artifact["sha256"] == binary_expected["sha256"], f"{row['cell']} binary digest changed")
    report_path, report_artifact = resolver.checked_artifact(row.get("report", {}), label=f"{row['cell']} report")
    terminal_path, terminal_artifact = resolver.checked_artifact(row.get("terminal", {}), label=f"{row['cell']} terminal")
    terminal_data = validate_terminal(resolver, plan, row, row["terminal"], terminal_path, source_expected, binary_expected)
    require(terminal_artifact["sha256"] == row["terminal"]["sha256"], f"{row['cell']} terminal receipt digest changed")
    rss_path, rss_artifact = resolver.checked_artifact(row.get("rss", {}), label=f"{row['cell']} RSS")
    rss_kib = CAP.read_rss(rss_path)
    require(isinstance(row.get("max_rss_kib"), int) and not isinstance(row.get("max_rss_kib"), bool) and row.get("max_rss_kib") >= 0, f"{row['cell']} RSS value is not a nonnegative integer")
    require(row.get("max_rss_kib") == rss_kib, f"{row['cell']} RSS value changed")
    report = read_object(report_path)
    CAP.validate_report(report, plan, cell, leg, lane, plan["scope"][lane]["samples"], plan["scope"][lane]["warmup"])
    require(report_artifact["sha256"] == row["report"]["sha256"], f"{row['cell']} report digest changed after parse")
    values = {"elapsed_ns": [sample["elapsed_ns"] for sample in report["samples"]], "rss_kib": [rss_kib]}
    if lane == "allocator":
        values["live_bytes"] = [sample["allocation"]["live_bytes_after"] for sample in report["samples"]]
        values["region_peak_live_bytes"] = [sample["allocation"]["region_peak_live_bytes"] for sample in report["samples"]]
    child_metrics = {name: midpoint(items) for name, items in values.items()}
    child_spreads = {name: spread(items) for name, items in values.items()}
    return {"ordinal": row["ordinal"], "lane": lane, "block": row["block"], "cell": row["cell"], "leg": leg, "fixture": row["fixture"], "phase": row["phase"], "source": {"path": str(source_path), "bytes": source_artifact["bytes"], "sha256": source_artifact["sha256"]}, "binary": {"path": str(binary_path), "bytes": binary_artifact["bytes"], "sha256": binary_artifact["sha256"]}, "report": {"path": str(report_path), "bytes": report_artifact["bytes"], "sha256": report_artifact["sha256"]}, "terminal": {"path": str(terminal_path), "bytes": terminal_artifact["bytes"], "sha256": terminal_artifact["sha256"]}, "rss": {"path": str(rss_path), "bytes": rss_artifact["bytes"], "sha256": rss_artifact["sha256"], "max_rss_kib": rss_kib}, "samples": len(report["samples"]), "sample_values": values, "child_metrics": child_metrics, "within_child_spread": child_spreads}


def expected_rows(plan: dict) -> list[tuple[str, int, str, str]]:
    rows = []
    for lane in ("native", "allocator"):
        for cell in plan["cells"]:
            for block, order in enumerate(plan["scope"][lane]["orders"]):
                for leg in order:
                    rows.append((lane, block, cell["id"], leg))
    return rows


def load_receipts(resolver: Resolver, packet_root: Path, lane: str) -> list[dict]:
    path = packet_root / lane / "receipts-v3.json"
    require(path.is_file(), f"{lane} receipts missing: {path}")
    value = read(path)
    require(isinstance(value, list), f"{lane} receipts must be a list")
    return value


def reduce_packet(plan: dict, packet_root: Path) -> dict:
    validate_plan_shape(plan)
    resolver = Resolver(packet_root)
    cells_by_id = cell_by_id(plan)
    raw_rows = load_receipts(resolver, packet_root, "native") + load_receipts(resolver, packet_root, "allocator")
    expected = expected_rows(plan)
    actual_keys = [(row.get("lane"), row.get("block"), row.get("cell"), row.get("leg")) for row in raw_rows]
    require(actual_keys == expected, "receipt order or child matrix changed")
    rows = []
    for ordinal, row in enumerate(raw_rows):
        row = dict(row)
        row.setdefault("ordinal", ordinal)
        rows.append(validate_row(resolver, plan, cells_by_id, row))
    by_key = {(row["lane"], row["block"], row["cell"], row["leg"]): row for row in rows}
    reductions: list[dict] = []
    for cell in plan["cells"]:
        for lane in ("native", "allocator"):
            metrics = ["elapsed_ns", "rss_kib"] + (["live_bytes", "region_peak_live_bytes"] if lane == "allocator" else [])
            for metric in metrics:
                gate = "informational" if lane == "allocator" and metric == "elapsed_ns" else ("adoption_and_review" if lane == "native" and metric == "elapsed_ns" else "review")
                leg_summaries = {}
                for leg in ("baseline", "candidate"):
                    child_rows = [by_key[(lane, block, cell["id"], leg)] for block in range(plan["scope"][lane]["blocks"])]
                    child_values = [row["child_metrics"][metric] for row in child_rows]
                    child_spreads = [row["within_child_spread"][metric] for row in child_rows]
                    leg_summaries[leg] = {"child_values": child_values, "child_spreads": child_spreads, "p50": midpoint(child_values), "p95": nearest_rank(child_values, 0.95), "p99": nearest_rank(child_values, 0.99), "spread": spread(child_values), "within_child_spread_p50": midpoint([value for value in child_spreads if value is not None]) if any(value is not None for value in child_spreads) else None}
                ratios = {name: ratio(leg_summaries["candidate"][name], leg_summaries["baseline"][name]) for name in ("p50", "p95", "p99", "spread", "within_child_spread_p50")}
                flags = {name: ratio_flag(value) for name, value in ratios.items()}
                reductions.append({"cell": cell["id"], "lane": lane, "metric": metric, "gate": gate, "baseline": leg_summaries["baseline"], "candidate": leg_summaries["candidate"], "candidate_over_baseline": ratios, "over_5pct_flags": flags, "hard_native_p50_adoption_reject": gate == "adoption_and_review" and flags["p50"] is True})
    return {"schema": "litchi.performance.0941.xlsx-reduction.v4", "plan_schema": plan["schema"], "packet_root": str(packet_root), "children": len(rows), "native_children": sum(row["lane"] == "native" for row in rows), "allocator_children": sum(row["lane"] == "allocator" for row in rows), "individual_rows": rows, "reductions": reductions, "gates_from_plan": plan["gates"], "status": "validated_and_reduced"}


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--packet-root", type=Path, default=DEFAULT_PACKET)
    parser.add_argument("--plan", type=Path, default=PLAN_PATH)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    plan = read_object(args.plan)
    result = reduce_packet(plan, args.packet_root)
    output = args.output or args.packet_root / "reduction-v3.json"
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"status": result["status"], "children": result["children"], "native_children": result["native_children"], "allocator_children": result["allocator_children"], "output": str(output)}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except RuntimeError as exc:
        raise SystemExit(f"reduction refused: {exc}") from exc
