#!/usr/bin/env python3
"""Replay and reduce one completed 0925 packet without starting children.

The reducer loads the packet-local capture and comparison scripts, verifies the
moved packet first, then validates the same child reports while collecting the
raw phase vectors, provider observations, RSS receipts, and child descriptors.
It never captures, reruns, or selectively drops a child.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import sys
from pathlib import Path
from typing import Any


MAX_ARTIFACT_BYTES = 1 << 20


class ReductionError(RuntimeError):
    """A packet replay or reduction contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReductionError(message)


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load packet script: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_packet_tools(packet: Path) -> tuple[Any, Any]:
    scripts = packet / "scripts"
    capture = load_module(scripts / "comparison-capture.py", "litchi_perf_0925_packet_capture")
    comparison = capture.load_comparison(scripts / "comparison.py")
    return capture, comparison


def reduce_packet(packet: Path) -> dict[str, Any]:
    packet = packet.resolve()
    capture, comparison = load_packet_tools(packet)
    replay = capture.replay(argparse.Namespace(packet=packet))
    manifest = capture.load_json(packet / "manifest.json")
    state = capture.load_json(packet / "capture-state.json")
    children = state.get("children")
    require(isinstance(children, list) and len(children) == 42,
            "completed packet child list changed after replay")

    reports: dict[str, list[dict[str, Any]]] = {route: [] for route in comparison.ROUTES}
    child_receipts: dict[str, list[dict[str, Any]]] = {route: [] for route in comparison.ROUTES}
    for receipt in children:
        schedule = receipt.get("schedule")
        require(isinstance(schedule, dict), "child schedule is missing")
        child_id = str(schedule.get("id"))
        route = schedule.get("route")
        require(route in reports, f"unknown child route: {route}")
        report_path = capture.verify_local_descriptor(
            packet, receipt.get("report"), f"{child_id}: report"
        )
        report = capture.load_json(report_path)
        summary = comparison.validate_report(
            report,
            route,
            warmups=schedule["warmups"],
            samples=schedule["samples"],
            label=child_id,
            binary=manifest["binary"],
        )
        rss = receipt.get("rss_bytes")
        require(type(rss) is int and rss > 0, f"{child_id}: RSS receipt is unavailable")
        reports[route].append(summary)
        child_receipts[route].append({
            "id": child_id,
            "block": schedule["block"],
            "position": schedule["position"],
            "order": schedule["order"],
            "report": receipt["report"],
            "time": receipt["time"],
            "rss_bytes": rss,
        })

    reduced = comparison.reduce_reports(reports)
    for route in comparison.ROUTES:
        values = reports[route]
        rss_values = [row["rss_bytes"] for row in child_receipts[route]]
        require(len(values) == 6 and len(rss_values) == 6,
                f"{route}: expected six completed children")
        reduced_route = reduced["routes"][route]
        reduced_route["child_phase_statistics_ns"] = [
            {phase: comparison._statistics(summary["phase_ns"][phase])
             for phase in comparison.PHASES}
            for summary in values
        ]
        reduced_route["child_rss_bytes"] = {
            "values": rss_values,
            "statistics": comparison._statistics(rss_values),
        }
        reduced_route["child_receipts"] = child_receipts[route]

    output = {
        "schema": "litchi.performance.0925.packet-reduction.v1",
        "status": "verified_and_reduced",
        "claim_scope": "descriptive provider baseline only; no ranking, optimization, or adoption",
        "performance_children_run": 0,
        "packet": str(packet),
        "binary": manifest["binary"],
        "replay": replay,
        "routes": reduced["routes"],
    }
    return output


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        output = reduce_packet(args.packet)
        data = (json.dumps(output, indent=2, sort_keys=True) + "\n").encode()
        require(len(data) < MAX_ARTIFACT_BYTES, "reduction receipt is at least 1 MiB")
        output_path = args.output.resolve()
        require(not output_path.exists(), f"refusing to replace retained receipt: {output_path}")
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_bytes(data)
        print(json.dumps({
            "status": output["status"],
            "output": str(output_path),
            "bytes": len(data),
            "routes": len(output["routes"]),
            "performance_children_run": 0,
        }, sort_keys=True))
    except Exception as error:
        print(f"0925 reduction failed: {type(error).__name__}: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
