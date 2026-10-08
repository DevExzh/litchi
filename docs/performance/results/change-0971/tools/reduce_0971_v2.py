#!/usr/bin/env python3
"""Offline reducer for the 0971 instruction attribution packets."""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

import protocol_common_0971_v2 as p


def self_test() -> int:
    rows = [1.0, 2.0, 3.0, 4.0]
    assert p.percentile(rows, 0.5) == 2.5
    assert p.pearson(rows, rows) == 1.0
    assert p.spearman(rows, list(reversed(rows))) == -1.0
    assert p.percentile([1.0, 2.0, 3.0, 4.0], 0.99) < 4.0
    print(json.dumps({"schema": "litchi.performance.0971.reducer_self_test.v2", "status": "pass"}, indent=2))
    return 0


def load_plan(path: Path) -> dict[str, Any]:
    plan = p.load_json(path.resolve())
    p.check_plan_bindings(plan, path.resolve())
    return plan


def load_packets(plan_path: Path, manifest_path: Path, plan: dict[str, Any]) -> list[dict[str, Any]]:
    manifest = p.load_json(manifest_path.resolve())
    if manifest.get("schema") != "litchi.performance.0971.capture_manifest.v2" or manifest.get("status") != "complete" or manifest.get("no_retry") is not True:
        raise p.ProtocolError("capture manifest is not a complete no-retry v2 manifest")
    plan_row = manifest.get("plan") or {}
    if plan_row.get("path") != str(plan_path.resolve()) or plan_row.get("sha256") != p.sha256(plan_path):
        raise p.ProtocolError("capture manifest plan binding mismatch")
    paths = manifest.get("packets")
    if not isinstance(paths, list) or len(paths) != p.CHILDREN or len(set(paths)) != p.CHILDREN:
        raise p.ProtocolError("capture manifest must list exactly 12 unique packets")
    packets: list[dict[str, Any]] = []
    for index, raw in enumerate(paths, 1):
        packet_path = Path(raw).absolute()
        packet = p.load_json(packet_path)
        if packet.get("schema") != p.PACKET_SCHEMA or packet.get("status") != "complete" or packet.get("packet_id") != f"child-{index:02d}" or packet.get("child_number") != index:
            raise p.ProtocolError(f"packet identity/status mismatch at child {index}")
        artifacts = p.check_artifact_bindings(packet, packet_path)
        raw_terminal = p.load_json(artifacts["terminal"])
        if raw_terminal != packet.get("terminal"):
            raise p.ProtocolError(f"child {index}: terminal packet/raw mismatch")
        terminal = raw_terminal
        if terminal.get("returncode") != 0 or terminal.get("timed_out") is not False:
            raise p.ProtocolError(f"child {index}: terminal result failed")
        raw_command = p.load_json(artifacts["command"])
        if raw_command.get("argv") != terminal.get("command") or raw_command.get("environment") != {"LC_ALL": "C", "LANG": "C", "TZ": "UTC"}:
            raise p.ProtocolError(f"child {index}: command custody mismatch")
        context = packet.get("context") or {}
        if context != {"binary_sha256": plan["binary"]["sha256"], "cell": {"case": "docfloat", "edits": 1, "retention": "default"}, "fixture_sha256": plan["fixture"]["sha256"], "marker_calls": p.CALLS, "source_manifest_sha256": plan["source_manifest"]["sha256"]}:
            raise p.ProtocolError(f"child {index}: context identity mismatch")
        raw_report = p.load_json(artifacts["report"])
        report = packet.get("report")
        if raw_report != report:
            raise p.ProtocolError(f"child {index}: report packet/raw mismatch")
        if not isinstance(report, dict):
            raise p.ProtocolError(f"child {index}: report missing")
        p.validate_report(report)
        raw_launch = p.check_launch_metadata(artifacts["launch"], packet_id=f"child-{index:02d}", binary=Path(plan["binary"]["path"]))
        launch = packet.get("launch")
        if raw_launch != launch:
            raise p.ProtocolError(f"child {index}: launch packet/raw mismatch")
        decoded = packet.get("decoded") or {}
        for name in ("no_inline", "inline", "header", "buildid"):
            if decoded.get(name) != str(artifacts[name]):
                raise p.ProtocolError(f"child {index}: decoded {name} path mismatch")
        script_path = artifacts["no_inline"]
        reparsed = p.parse_perf_script(script_path.read_text(encoding="utf-8"), marker_dso=str(Path(plan["binary"]["path"]).absolute()), expected_pid=launch["exec_pid"], expected_cpu=p.CPU)
        perf_summary = packet.get("perf")
        if not isinstance(perf_summary, dict) or perf_summary != reparsed:
            raise p.ProtocolError(f"child {index}: full authoritative perf summary mismatch")
        p.check_buildid_export(artifacts["buildid"].read_text(encoding="utf-8"), binary_path=plan["binary"]["path"], build_id=plan["binary"]["build_id"])
        # The inline decode is retained and hashed as raw evidence; the
        # physical no-inline decode is the authoritative owner parser.
        packet["_packet_path"] = str(packet_path)
        packet["_artifacts"] = artifacts
        packets.append(packet)
    return packets


def reduce(plan_path: Path, manifest_path: Path, output: Path) -> int:
    plan = load_plan(plan_path)
    packets = load_packets(plan_path, manifest_path, plan)
    p50s: list[float] = []
    marker_periods: list[float] = []
    marker_shares: list[float] = []
    owner_period: dict[str, float] = {}
    child_rows: list[dict[str, Any]] = []
    for packet in packets:
        report = packet["report"]
        elapsed = [float(value) for value in report["elapsed_ns"]]
        p50 = p.percentile(elapsed, 0.5)
        perf = packet["perf"]
        p50s.append(p50)
        marker_periods.append(float(perf["marker_period"]))
        marker_shares.append(float(perf["marker_period_share"]) if perf.get("marker_period_share") is not None else 0.0)
        for owner, value in (perf.get("exclusive_deepest_resolved_period") or {}).items():
            owner_period[owner] = owner_period.get(owner, 0.0) + float(value)
        child_rows.append({
            "packet_id": packet["packet_id"],
            "artifacts": packet["artifacts"],
            "p50_ns": p50,
            "p95_ns": p.percentile(elapsed, 0.95),
            "p99_ns": p.percentile(elapsed, 0.99),
            "max_ns": max(elapsed),
            "elapsed_sample_count": len(elapsed),
            "perf_sample_count": perf["sample_count"],
            "total_period": perf["total_period"],
            "marker_sample_count": perf["marker_sample_count"],
            "marker_period": perf["marker_period"],
            "marker_period_share": perf["marker_period_share"],
            "lost_events": perf["lost_events"],
            "unresolved_frames": perf["unresolved_frames"],
            "stack_truncated": perf["stack_truncated"],
            "exclusive_deepest_resolved_period": perf.get("exclusive_deepest_resolved_period", {}),
            "inclusive_marker_ancestor_period": perf.get("inclusive_marker_ancestor_period", {}),
        })
    total_marker_period = sum(marker_periods)
    owner_share = {name: value / total_marker_period for name, value in sorted(owner_period.items())} if total_marker_period else {}
    owner_associations: dict[str, Any] = {}
    owner_names = sorted({owner for packet in packets for owner in (packet["perf"].get("exclusive_deepest_resolved_period") or {})})
    for owner in owner_names:
        periods = [float((packet["perf"].get("exclusive_deepest_resolved_period") or {}).get(owner, 0)) for packet in packets]
        shares = [period / float(packet["perf"]["marker_period"]) if packet["perf"]["marker_period"] > 0 else 0.0 for period, packet in zip(periods, packets)]
        owner_associations[owner] = {
            "marker_period": {"pearson": p.pearson(p50s, periods), "spearman": p.spearman(p50s, periods)},
            "marker_period_share": {"pearson": p.pearson(p50s, shares), "spearman": p.spearman(p50s, shares)},
            "periods_all_12": periods,
            "shares_all_12": shares,
            "scope": "all 12 children; missing owner in a child is zero coverage; descriptive only",
        }
    result = {
        "schema": "litchi.performance.0971.instruction_profile_analysis.v2",
        "status": "complete",
        "analysis_only": True,
        "diagnostic_only": True,
        "qualification_claim": False,
        "speed_claim": False,
        "plan": {"path": str(plan_path.resolve()), "sha256": p.sha256(plan_path)},
        "manifest": {"path": str(manifest_path.resolve()), "sha256": p.sha256(manifest_path)},
        "custody": {"binary": plan["binary"], "fixture": plan["fixture"], "source_manifest": plan["source_manifest"], "marker": plan["marker"]},
        "schedule": {"children": p.CHILDREN, "warmups": p.WARMUPS, "measured_samples": p.SAMPLES, "marker_calls": p.CALLS, "serial": True, "cpu": p.CPU},
        "marker_scope": {
            "combined_calls": "5 warmups + 80 measured calls",
            "p50": "80 measured samples only",
            "lifecycle": "85-call lifecycle subset",
            "owner_percentage": "not exact timed-region attribution",
            "missing_membership": "unknown",
            "unresolved_ancestry": "reported descriptively",
        },
        "children": child_rows,
        "summary": {
            "p50_ns": {"median": p.percentile(p50s, 0.5), "p95": p.percentile(p50s, 0.95), "min": min(p50s), "max": max(p50s)},
            "marker_period": {"median": p.percentile(marker_periods, 0.5), "total": sum(marker_periods)},
            "marker_period_share": {"median": p.percentile(marker_shares, 0.5), "min": min(marker_shares), "max": max(marker_shares)},
            "exclusive_deepest_resolved_period": owner_period,
            "exclusive_deepest_resolved_share": owner_share,
            "associations_descriptive": {
                "p50_vs_marker_period": {"pearson": p.pearson(p50s, marker_periods), "spearman": p.spearman(p50s, marker_periods)},
                "p50_vs_marker_period_share": {"pearson": p.pearson(p50s, marker_shares), "spearman": p.spearman(p50s, marker_shares)},
                "per_exclusive_deepest_owner": owner_associations,
                "scope": "all 12 children; no fast/slow subset; descriptive only",
            },
        },
        "pilot": {"required": plan["pilot_gate"], "observed": {"marker_samples": packets[0]["perf"]["marker_sample_count"], "lost_events": packets[0]["perf"]["lost_events"], "marker_period_share": packets[0]["perf"]["marker_period_share"]}, "passed": packets[0]["perf"]["marker_sample_count"] >= 100 and packets[0]["perf"]["lost_events"] == 0 and bool(packets[0]["perf"]["marker_period_share"] and packets[0]["perf"]["marker_period_share"] > 0)},
        "interpretation": "A descriptive process-instruction association may support a mechanism hypothesis; it cannot locate an exact timed-region cause or support speed/adoption claims.",
    }
    p.write_json_exclusive(output.resolve(), result)
    print(json.dumps({"status": "pass", "output": str(output.resolve()), "sha256": p.sha256(output.resolve())}, indent=2, sort_keys=True))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--self-test", action="store_true")
    parser.add_argument("--plan", type=Path)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        if args.self_test:
            if any(value is not None for value in (args.plan, args.manifest, args.output)):
                raise p.ProtocolError("--self-test is standalone")
            return self_test()
        if args.plan is None or args.manifest is None or args.output is None:
            raise p.ProtocolError("--plan, --manifest, and --output are required")
        return reduce(args.plan.resolve(), args.manifest.resolve(), args.output.resolve())
    except Exception as exc:
        print(f"0971 reducer rejected: {exc}", file=__import__("sys").stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
