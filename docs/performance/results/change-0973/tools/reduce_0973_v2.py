#!/usr/bin/env python3
"""Offline, fail-closed reducer for 0973 dynamic-bias packets."""
from __future__ import annotations

import argparse
import json
import math
import sys
from pathlib import Path
from typing import Any

import protocol_common_0973_v2 as p


def _without_control(row: dict[str, Any]) -> dict[str, Any]:
    return {key: value for key, value in row.items() if key != "control_stdout_sha256"}


def _check_log(path: Path, *, mode: str, layout: dict[str, Any], binary: Path) -> dict[str, int]:
    counts = {"rows": 0, "reports": 0, "infos": 0, "adjusted": 0, "target_reports": 0, "target_infos": 0}
    corrected_base = layout["corrected_base"]
    expected_start = layout["perf_expected_start"]
    for line in path.read_text(encoding="utf-8").splitlines():
        if not line:
            continue
        row = json.loads(line)
        counts["rows"] += 1
        schema = row.get("schema", "")
        if schema.endswith("dwfl_report.v1"):
            counts["reports"] += 1
            target = bool(row.get("target_match"))
            if target:
                counts["target_reports"] += 1
            if row.get("replacement_applied"):
                if mode != "corrected" or not target or row.get("forwarded_base") != corrected_base or row.get("config_valid") is not True:
                    raise p.ProtocolError(f"{mode}: invalid report correction row")
            elif mode == "noop" and target and row.get("original_base") != row.get("forwarded_base"):
                raise p.ProtocolError("no-op target report changed base")
        elif schema.endswith("dwfl_info.v1"):
            counts["infos"] += 1
            target = bool(row.get("target_match"))
            if target:
                counts["target_infos"] += 1
            if row.get("start_adjusted"):
                counts["adjusted"] += 1
                if mode != "corrected" or not target or row.get("caller_is_perf") is not True or row.get("caller_file") != "/usr/bin/perf" or row.get("before_start") != corrected_base or row.get("after_start") != expected_start or row.get("config_valid") is not True:
                    raise p.ProtocolError(f"{mode}: invalid info correction row")
            elif mode == "noop" and row.get("target_match") and row.get("before_start") != row.get("after_start"):
                raise p.ProtocolError("no-op target info changed start")
        else:
            raise p.ProtocolError(f"unexpected shim log schema: {schema}")
    if mode in {"noop", "corrected"} and counts["rows"] == 0:
        raise p.ProtocolError(f"{mode}: missing shim log rows")
    if mode == "noop" and counts["adjusted"] != 0:
        raise p.ProtocolError("no-op adjustment was observed")
    return counts


def _child_gate(corrected: dict[str, Any]) -> dict[str, Any]:
    reasons: list[str] = []
    if corrected.get("marker_sample_count") < 100:
        reasons.append("marker_sample_count_lt_100")
    if corrected.get("lost_events") != 0:
        reasons.append("lost_events_nonzero")
    share = corrected.get("marker_period_share")
    if share is None or share <= 0:
        reasons.append("marker_period_share_nonpositive")
    return {"passed": not reasons, "reasons": reasons}


def _population_stats(values: list[float]) -> dict[str, Any]:
    if not values:
        raise p.ProtocolError("cannot summarize an empty metric")
    numbers = [float(value) for value in values]
    mean = sum(numbers) / len(numbers)
    variance = sum((value - mean) ** 2 for value in numbers) / len(numbers)
    stddev = math.sqrt(variance)
    return {
        "count": len(numbers),
        "min": min(numbers),
        "median": p.percentile(numbers, 0.5),
        "max": max(numbers),
        "mean": mean,
        "stddev_population": stddev,
        "cv_population": stddev / mean if mean != 0 else None,
    }


def _population_covariance(xs: list[float], ys: list[float]) -> float:
    if len(xs) != len(ys) or not xs:
        raise p.ProtocolError("covariance vectors are not paired")
    mean_x = sum(xs) / len(xs)
    mean_y = sum(ys) / len(ys)
    return sum((x - mean_x) * (y - mean_y) for x, y in zip(xs, ys)) / len(xs)


def _marker_owner_maps(decoded: dict[str, Any], binary_path: str) -> dict[str, Any]:
    """Derive disjoint frame-0 and overlapping inclusive maps from marker rows."""
    inclusive_period: dict[str, float] = {}
    inclusive_samples: dict[str, int] = {}
    physical_period: dict[str, float] = {}
    physical_samples: dict[str, int] = {}
    unknown_period = 0.0
    unknown_samples = 0
    marker_period = 0.0
    marker_count = 0
    for sample in decoded.get("samples", []):
        frames = sample.get("frames") or []
        marker = any(
            frame.get("symbol_base") == p.MARKER and frame.get("dso") == binary_path
            for frame in frames
        )
        if not marker:
            continue
        period = float(sample["period"])
        marker_count += 1
        marker_period += period
        resolved = [frame for frame in frames if frame.get("resolved") is True]
        owners = {
            f"{frame['dso']}::{frame['symbol_base']}"
            for frame in resolved
        }
        for owner in owners:
            inclusive_period[owner] = inclusive_period.get(owner, 0.0) + period
            inclusive_samples[owner] = inclusive_samples.get(owner, 0) + 1
        if not frames or frames[0].get("resolved") is not True:
            unknown_period += period
            unknown_samples += 1
        else:
            owner = f"{frames[0]['dso']}::{frames[0]['symbol_base']}"
            physical_period[owner] = physical_period.get(owner, 0.0) + period
            physical_samples[owner] = physical_samples.get(owner, 0) + 1
    if marker_count != decoded.get("marker_sample_count") or marker_period != float(decoded.get("marker_period", -1)):
        raise p.ProtocolError("decoded marker rows disagree with marker summary")
    if physical_period != {key: float(value) for key, value in decoded.get("physical_leaf_owner_period", {}).items()}:
        raise p.ProtocolError("decoded physical leaf owner map disagrees with marker rows")
    if unknown_period != float(decoded.get("physical_leaf_unknown_period", -1)) or unknown_samples != decoded.get("physical_leaf_unknown_samples"):
        raise p.ProtocolError("decoded physical leaf unknown bucket disagrees with marker rows")
    return {
        "inclusive_period": inclusive_period,
        "inclusive_samples": inclusive_samples,
        "physical_period": physical_period,
        "physical_samples": physical_samples,
        "unknown_period": unknown_period,
        "unknown_samples": unknown_samples,
        "marker_period": marker_period,
        "marker_count": marker_count,
    }


def _load_packets(plan_path: Path, manifest_path: Path, plan: dict[str, Any]) -> list[dict[str, Any]]:
    manifest = p.load_json(manifest_path.resolve())
    if manifest.get("schema") != p.MANIFEST_SCHEMA or manifest.get("status") != "complete" or manifest.get("no_retry") is not True:
        raise p.ProtocolError("manifest is not complete/no-retry")
    binding = manifest.get("plan") or {}
    if binding.get("path") != str(plan_path.resolve()) or binding.get("sha256") != p.sha256(plan_path):
        raise p.ProtocolError("manifest plan binding mismatch")
    paths = manifest.get("packets")
    if not isinstance(paths, list) or len(paths) != p.CHILDREN or len(set(paths)) != p.CHILDREN:
        raise p.ProtocolError("manifest must list exactly 12 unique packets")
    packets: list[dict[str, Any]] = []
    for index, raw_path in enumerate(paths, 1):
        packet_path = Path(raw_path).absolute()
        packet = p.load_json(packet_path)
        child_id = f"child-{index:02d}"
        if packet.get("schema") != p.PACKET_SCHEMA or packet.get("status") != "complete" or packet.get("packet_id") != child_id or packet.get("child_number") != index:
            raise p.ProtocolError(f"packet identity/status mismatch: {child_id}")
        artifacts = p.check_artifact_bindings_0973(packet, packet_path)
        terminal = p.load_json(artifacts["terminal"])
        if terminal != packet.get("terminal") or terminal.get("schema") != "litchi.performance.0973.terminal.v2" or terminal.get("returncode") != 0 or terminal.get("timed_out") is not False or terminal.get("native_ld_preload") != "absent":
            raise p.ProtocolError(f"{child_id}: terminal mismatch/failure")
        command = p.load_json(artifacts["command"])
        if command.get("argv") != terminal.get("command") or command.get("environment") != {**p.ENVIRONMENT, "LD_PRELOAD": "absent"}:
            raise p.ProtocolError(f"{child_id}: native command environment mismatch")
        report = p.load_json(artifacts["report"])
        if report != packet.get("report"):
            raise p.ProtocolError(f"{child_id}: report custody mismatch")
        p.validate_report(report)
        launch = p.check_launch_metadata(artifacts["launch"], packet_id=child_id, binary=Path(plan["binary"]["path"]))
        if launch != packet.get("launch") or launch.get("exec_pid") != launch.get("pid"):
            raise p.ProtocolError(f"{child_id}: launcher custody mismatch")
        layout = p.load_json(artifacts["layout"])
        if layout != packet.get("layout"):
            raise p.ProtocolError(f"{child_id}: layout custody mismatch")
        p.check_layout(layout, perf_data=artifacts["perf_data"], binary=Path(plan["binary"]["path"]), expected_pid=int(launch["exec_pid"]))
        if p.sha256(artifacts["mmap_dump"]) != layout["perf_mmap_dump"]["sha256"]:
            raise p.ProtocolError(f"{child_id}: MMAP2 dump/layout hash mismatch")
        raw_stack = p.load_json(artifacts["raw_stack"])
        p.check_raw_stack_summary(raw_stack, plan=plan)
        config = packet.get("shim_config") or {}
        if config != {"target_path": plan["binary"]["path"], "target_name": Path(plan["binary"]["path"]).name, "corrected_base": layout["corrected_base"], "perf_expected_start": layout["perf_expected_start"]}:
            raise p.ProtocolError(f"{child_id}: shim config mismatch")
        controls = packet.get("perf") or {}
        if set(controls) != {"stock", "noop", "corrected", "control", "marker_scope", "physical_leaf_scope"}:
            raise p.ProtocolError(f"{child_id}: perf summary shape mismatch")
        parsed: dict[str, dict[str, Any]] = {}
        for mode, artifact_key in (("stock", "stock_script"), ("noop", "noop_script"), ("corrected", "corrected_script")):
            parsed[mode] = p.parse_perf_script_strict(artifacts[artifact_key].read_text(encoding="utf-8"), marker_dso=plan["binary"]["path"], expected_pid=int(launch["exec_pid"]), expected_cpu=p.CPU)
            if _without_control(controls[mode]) != parsed[mode]:
                raise p.ProtocolError(f"{child_id}: {mode} parsed summary mismatch")
        expected_gate = _child_gate(parsed["corrected"])
        if packet.get("gate") != expected_gate or expected_gate["passed"] is not True:
            raise p.ProtocolError(f"{child_id}: per-child marker gate failed or was altered")
        if artifacts["stock_script"].read_bytes() != artifacts["noop_script"].read_bytes() or artifacts["stock_stderr"].read_bytes() != artifacts["noop_stderr"].read_bytes():
            raise p.ProtocolError(f"{child_id}: stock/no-op output is not byte-identical")
        logs = {mode: _check_log(artifacts[f"{mode}_log"], mode=mode, layout=layout, binary=Path(plan["binary"]["path"])) for mode in ("noop", "corrected")}
        if parsed["corrected"]["marker_ancestry"] not in {"observed", "no observed marker ancestry"}:
            raise p.ProtocolError(f"{child_id}: marker ancestry label invalid")
        expected_context = {"cell": {"case": "docfloat", "edits": 1, "retention": "default"}, "binary_sha256": plan["binary"]["sha256"], "fixture_sha256": plan["fixture"]["sha256"], "marker_calls": p.CALLS}
        if packet.get("context") != expected_context:
            raise p.ProtocolError(f"{child_id}: context mismatch")
        packet["_packet_path"] = str(packet_path)
        packet["_artifacts"] = artifacts
        packet["_parsed"] = parsed
        packet["_raw_stack"] = raw_stack
        packet["_log_stats"] = logs
        packets.append(packet)
    return packets


def _reduce_packets(plan_path: Path, manifest_path: Path, plan: dict[str, Any], packets: list[dict[str, Any]]) -> dict[str, Any]:
    if len(packets) != p.CHILDREN:
        raise p.ProtocolError(f"reducer requires exactly {p.CHILDREN} children")
    p50s: list[float] = []
    p95s: list[float] = []
    d_periods: list[float] = []
    m_periods: list[float] = []
    r_periods: list[float] = []
    marker_counts: list[float] = []
    marker_shares: list[float] = []
    residual_shares: list[float] = []
    corrected_rows: list[dict[str, Any]] = []
    physical_owner_period: dict[str, float] = {}
    physical_owner_samples: dict[str, int] = {}
    inclusive_owner_period: dict[str, float] = {}
    inclusive_owner_samples: dict[str, int] = {}
    proxy_owner_period: dict[str, float] = {}
    physical_unknown_periods: list[float] = []
    physical_unknown_samples: list[int] = []
    marker_conservation_rows: list[dict[str, Any]] = []

    for packet in packets:
        report = packet["report"]
        elapsed = [float(x) for x in report["elapsed_ns"]]
        p50 = p.percentile(elapsed, 0.5)
        p95 = p.percentile(elapsed, 0.95)
        decoded = packet["_parsed"]["corrected"]
        maps = _marker_owner_maps(decoded, plan["binary"]["path"])
        d_value = float(decoded["total_period"])
        m_value = float(decoded["marker_period"])
        r_value = d_value - m_value
        if r_value < 0:
            raise p.ProtocolError(f"{packet['packet_id']}: marker period exceeds total sampled period")
        classified_period = sum(maps["physical_period"].values()) + maps["unknown_period"]
        if classified_period != m_value:
            raise p.ProtocolError(
                f"{packet['packet_id']}: physical leaf marker-period conservation failed "
                f"({classified_period} != {m_value})"
            )
        if decoded.get("physical_leaf_scope") != "marker-containing samples only; frame-0 physical leaf or explicit unknown":
            raise p.ProtocolError(f"{packet['packet_id']}: physical leaf scope label drift")
        if d_value != m_value + r_value:
            raise p.ProtocolError(f"{packet['packet_id']}: D=M+R conservation failed")
        d_share = None  # filled after the all-child D denominator is known
        m_share = m_value / d_value if d_value else None
        r_share = r_value / d_value if d_value else None
        p50s.append(p50)
        p95s.append(p95)
        d_periods.append(d_value)
        m_periods.append(m_value)
        r_periods.append(r_value)
        marker_counts.append(float(maps["marker_count"]))
        marker_shares.append(m_share if m_share is not None else 0.0)
        residual_shares.append(r_share if r_share is not None else 0.0)
        physical_unknown_periods.append(maps["unknown_period"])
        physical_unknown_samples.append(maps["unknown_samples"])
        marker_conservation_rows.append({
            "packet_id": packet["packet_id"], "D": d_value, "M": m_value, "R": r_value,
            "R_label": "marker ancestry not observed", "passed": True,
        })
        for owner, period in maps["physical_period"].items():
            physical_owner_period[owner] = physical_owner_period.get(owner, 0.0) + period
            physical_owner_samples[owner] = physical_owner_samples.get(owner, 0) + maps["physical_samples"].get(owner, 0)
        for owner, period in maps["inclusive_period"].items():
            inclusive_owner_period[owner] = inclusive_owner_period.get(owner, 0.0) + period
            inclusive_owner_samples[owner] = inclusive_owner_samples.get(owner, 0) + maps["inclusive_samples"].get(owner, 0)
        for owner, period in decoded["deepest_resolved_proxy_period"].items():
            proxy_owner_period[owner] = proxy_owner_period.get(owner, 0.0) + float(period)
        corrected_rows.append({
            "packet_id": packet["packet_id"],
            "gate": packet["gate"],
            "artifacts": packet["artifacts"],
            "p50_ns": p50, "p95_ns": p95, "p99_ns": p.percentile(elapsed, 0.99),
            "elapsed_sample_count": len(elapsed),
            "sample_count": decoded["sample_count"], "total_period": decoded["total_period"],
            "marker_sample_count": decoded["marker_sample_count"], "marker_period": decoded["marker_period"],
            "marker_period_share": decoded["marker_period_share"], "non_marker_period": r_value,
            "non_marker_period_share": r_share, "lost_events": decoded["lost_events"],
            "explicit_stack_truncated_count": decoded["explicit_stack_truncated_count"],
            "empty_callchain_count": decoded["empty_callchain_count"], "unresolved_frames": decoded["unresolved_frames"],
            "period_metrics": {
                "D": d_value, "M": m_value, "R": r_value,
                "R_label": "marker ancestry not observed",
                "total_period": d_value, "marker_period": m_value, "non_marker_period": r_value,
                "D_share_of_all_children_period": d_share,
                "M_share_of_D": m_share, "marker_period_share": m_share,
                "R_share_of_D": r_share, "non_marker_period_share": r_share,
            },
            "physical_leaf_owner_period": maps["physical_period"],
            "physical_leaf_owner_sample_count": maps["physical_samples"],
            "physical_leaf_unknown_period": maps["unknown_period"],
            "physical_leaf_unknown_samples": maps["unknown_samples"],
            "marker_unknown_physical_leaf_samples": decoded["marker_unknown_physical_leaf_samples"],
            "marker_no_resolved_frame_samples": decoded["marker_no_resolved_frame_samples"],
            "physical_leaf_scope": decoded["physical_leaf_scope"],
            "physical_leaf_classified_period": decoded["physical_leaf_classified_period"],
            "physical_leaf_conservation": decoded["physical_leaf_conservation"],
            "inclusive_owner_period": maps["inclusive_period"],
            "inclusive_owner_sample_count": maps["inclusive_samples"],
            "inclusive_owner_share_of_M": {owner: period / m_value for owner, period in maps["inclusive_period"].items()} if m_value else {},
            "inclusive_marker_ancestor_period": maps["inclusive_period"],
            "inclusive_marker_ancestor_sample_count": maps["inclusive_samples"],
            "inclusive_marker_ancestor_share_of_M": {owner: period / m_value for owner, period in maps["inclusive_period"].items()} if m_value else {},
            "deepest_resolved_proxy_period": decoded["deepest_resolved_proxy_period"],
            "raw_stack": packet["_raw_stack"],
            "marker_ancestry": decoded["marker_ancestry"],
            "log_statistics": packet["_log_stats"],
            "layout": packet["layout"],
        })

    total_d = sum(d_periods)
    if total_d <= 0:
        raise p.ProtocolError("total sampled period is not positive")
    d_shares = [value / total_d for value in d_periods]
    for row, d_share in zip(corrected_rows, d_shares):
        row["period_metrics"]["D_share_of_all_children_period"] = d_share

    def correlations(values: list[float]) -> dict[str, Any]:
        return {
            "values_all_12": values,
            "vs_T50": {"pearson": p.pearson(values, p50s), "spearman": p.spearman(values, p50s)},
            "vs_T95": {"pearson": p.pearson(values, p95s), "spearman": p.spearman(values, p95s)},
        }

    physical_names = sorted(physical_owner_period)
    physical_assoc: dict[str, Any] = {}
    for owner in physical_names:
        periods = [float(row["physical_leaf_owner_period"].get(owner, 0.0)) for row in corrected_rows]
        counts = [int(row["physical_leaf_owner_sample_count"].get(owner, 0)) for row in corrected_rows]
        shares = [period / m if m > 0 else 0.0 for period, m in zip(periods, m_periods)]
        presence = [count > 0 for count in counts]
        physical_assoc[owner] = {
            "sample_counts_all_12": counts, "periods_all_12": periods, "shares_all_12": shares,
            "present_all_12": presence, "presence_count": sum(presence),
            "missing_children": [row["packet_id"] for row, present in zip(corrected_rows, presence) if not present],
            "period_association": {"T50": correlations(periods), "T95": correlations(periods)},
            "share_association": {"T50": correlations(shares), "T95": correlations(shares)},
            "p50_vs_period": {"pearson": p.pearson(p50s, periods), "spearman": p.spearman(p50s, periods)},
            "p95_vs_period": {"pearson": p.pearson(p95s, periods), "spearman": p.spearman(p95s, periods)},
            "p50_vs_share": {"pearson": p.pearson(p50s, shares), "spearman": p.spearman(p50s, shares)},
            "p95_vs_share": {"pearson": p.pearson(p95s, shares), "spearman": p.spearman(p95s, shares)},
            "scope": "all 12 children; frame-0 physical leaf only; descriptive; owner key=<DSO>::<symbol>; no owner sum is a phase fraction",
        }
    physical_unknown_shares = [value / m if m > 0 else 0.0 for value, m in zip(physical_unknown_periods, m_periods)]
    physical_unknown_assoc = {
        "sample_counts_all_12": physical_unknown_samples,
        "periods_all_12": physical_unknown_periods,
        "shares_all_12": physical_unknown_shares,
        "present_all_12": [count > 0 for count in physical_unknown_samples],
        "presence_count": sum(count > 0 for count in physical_unknown_samples),
        "period_association": {"T50": correlations(physical_unknown_periods), "T95": correlations(physical_unknown_periods)},
        "share_association": {"T50": correlations(physical_unknown_shares), "T95": correlations(physical_unknown_shares)},
        "scope": "marker-containing samples whose physical frame 0 was absent or unresolved; marker ancestry not observed does not mean outside operation",
    }

    inclusive_names = sorted(inclusive_owner_period)
    inclusive_assoc: dict[str, Any] = {}
    for owner in inclusive_names:
        periods = [float(row["inclusive_owner_period"].get(owner, 0.0)) for row in corrected_rows]
        counts = [int(row["inclusive_owner_sample_count"].get(owner, 0)) for row in corrected_rows]
        shares = [period / m if m > 0 else 0.0 for period, m in zip(periods, m_periods)]
        presence = [count > 0 for count in counts]
        inclusive_assoc[owner] = {
            "sample_counts_all_12": counts, "periods_all_12": periods, "shares_all_12": shares,
            "present_all_12": presence, "presence_count": sum(presence),
            "missing_children": [row["packet_id"] for row, present in zip(corrected_rows, presence) if not present],
            "period_association": {"T50": correlations(periods), "T95": correlations(periods)},
            "share_association": {"T50": correlations(shares), "T95": correlations(shares)},
            "scope": "all 12 children; inclusive resolved marker-chain owner; M-normalized descriptive share; repeated frames deduplicated per sample; owners overlap and are never summed",
        }

    proxy_assoc = {}
    for owner in sorted({owner for packet in packets for owner in packet["_parsed"]["corrected"]["deepest_resolved_proxy_period"]}):
        periods = [float(packet["_parsed"]["corrected"]["deepest_resolved_proxy_period"].get(owner, 0)) for packet in packets]
        proxy_assoc[owner] = {"periods_all_12": periods, "scope": "all 12 children; deepest resolved frame proxy, not physical leaf; owner key=<DSO>::<symbol>"}

    metric_vectors = {
        "D": d_periods,
        "D_share_of_all_children_period": d_shares,
        "M": m_periods,
        "M_share_of_D": marker_shares,
        "R": r_periods,
        "R_share_of_D": residual_shares,
    }
    metric_distributions = {name: _population_stats(values) for name, values in metric_vectors.items()}
    metric_distributions.update({
        "marker_count": _population_stats(marker_counts),
        "T50": _population_stats(p50s),
        "T95": _population_stats(p95s),
    })
    period_correlations = {name: correlations(values) for name, values in metric_vectors.items()}
    variance_d = _population_covariance(d_periods, d_periods)
    variance_m = _population_covariance(m_periods, m_periods)
    variance_r = _population_covariance(r_periods, r_periods)
    covariance_m_r = _population_covariance(m_periods, r_periods)
    variance_rhs = variance_m + variance_r + 2.0 * covariance_m_r
    variance_error = variance_d - variance_rhs
    pilot = packets[0]["_parsed"]["corrected"]
    return {
        "schema": "litchi.performance.0973.instruction_profile_analysis.v2",
        "status": "complete", "analysis_only": True, "diagnostic_only": True,
        "qualification_claim": False, "speed_claim": False, "exact_timed_region_claim": False,
        "plan": {"path": str(plan_path.resolve()), "sha256": p.sha256(plan_path)},
        "manifest": {"path": str(manifest_path.resolve()), "sha256": p.sha256(manifest_path)},
        "custody": {"binary": plan["binary"], "fixture": plan["fixture"], "source_manifest": plan["source_manifest"], "marker": plan["marker"]},
        "schedule": {"children": p.CHILDREN, "warmups": p.WARMUPS, "measured_samples": p.SAMPLES, "marker_calls": p.CALLS, "serial": True, "cpu": p.CPU, "period": p.PERIOD, "callgraph": p.CALLGRAPH, "stack_cap": p.STACK_CAP},
        "marker_scope": {"combined_calls": "5 warmups + 80 measured calls", "p50": "80 measured samples only", "lifecycle": "85-call lifecycle subset", "owner_percentage": "not exact timed-region attribution", "missing_membership": "unknown", "unresolved_ancestry": "reported descriptively"},
        "children": corrected_rows,
        "summary": {
            "p50_ns": {"median": p.percentile(p50s, 0.5), "p95": p.percentile(p50s, 0.95), "min": min(p50s), "max": max(p50s)},
            "p95_ns": {"median": p.percentile(p95s, 0.5), "p95": p.percentile(p95s, 0.95), "min": min(p95s), "max": max(p95s)},
            "marker_period": {"median": p.percentile(m_periods, 0.5), "total": sum(m_periods)},
            "marker_period_share": {"median": p.percentile(marker_shares, 0.5), "min": min(marker_shares), "max": max(marker_shares)},
            "period_metrics": {
                "D_label": "total sampled instructions:u period",
                "M_label": "marker-containing sampled period",
                "R_label": "marker ancestry not observed",
                "distributions": metric_distributions,
                "correlations": period_correlations,
                "conservation": marker_conservation_rows,
            },
            "metric_distributions": metric_distributions,
            "period_correlations": period_correlations,
            "physical_leaf_owner_period": physical_owner_period,
            "physical_leaf_owner_sample_count": physical_owner_samples,
            "physical_leaf_unknown_period_all_12": physical_unknown_periods,
            "physical_leaf_unknown_samples_all_12": physical_unknown_samples,
            "physical_leaf_conservation": {"scope": "marker-containing samples only", "rows": marker_conservation_rows, "marker_period_all_12": m_periods, "classified_period_all_12": [sum(float(value) for value in row["physical_leaf_owner_period"].values()) + float(row["physical_leaf_unknown_period"]) for row in corrected_rows], "passed": True},
            "variance_ledger": {
                "scope": "all 12 children; population variance/covariance",
                "D_label": "total sampled instructions:u period",
                "M_label": "marker-containing sampled period",
                "R_label": "marker ancestry not observed",
                "variance_D": variance_d, "variance_M": variance_m, "variance_R": variance_r,
                "covariance_M_R": covariance_m_r, "rhs": variance_rhs, "identity_error": variance_error,
                "identity_passed": math.isclose(variance_d, variance_rhs, rel_tol=1e-12, abs_tol=1e-6),
            },
            "inclusive_owner_period": inclusive_owner_period,
            "inclusive_owner_sample_count": inclusive_owner_samples,
            "inclusive_owner_associations": inclusive_assoc,
            "inclusive_marker_ancestor_period": inclusive_owner_period,
            "inclusive_marker_ancestor_sample_count": inclusive_owner_samples,
            "inclusive_marker_ancestor_associations": inclusive_assoc,
            "physical_leaf_associations": physical_assoc,
            "physical_leaf_unknown_association": physical_unknown_assoc,
            "deepest_resolved_proxy_period": proxy_owner_period,
            "associations_descriptive": {"physical_leaf": physical_assoc, "inclusive_owner": inclusive_assoc, "deepest_resolved_proxy": proxy_assoc, "scope": "all 12 children; no fast/slow subset; inclusive owners overlap and are not summed"},
        },
        "pilot": {"required": plan["pilot_gate"], "observed": {"marker_samples": pilot["marker_sample_count"], "lost_events": pilot["lost_events"], "marker_period_share": pilot["marker_period_share"]}, "passed": True},
        "interpretation": "D is total sampled instructions:u period; M is marker-containing period; R is marker ancestry not observed and may include missing/truncated marker ancestry or wrapper/setup work. Inclusive owner maps overlap and are not phase fractions; physical leaf buckets are disjoint over marker-containing samples and conserve M.",
    }

def self_test() -> int:
    values = [1.0, 2.0, 3.0, 4.0]
    assert p.percentile(values, 0.5) == 2.5
    assert p.pearson(values, values) == 1.0
    assert p.spearman(values, list(reversed(values))) == -1.0
    print(json.dumps({"schema": "litchi.performance.0973.reducer_self_test.v2", "status": "pass", "physical_leaf_policy": "frame0-or-unknown", "children": 12}, indent=2, sort_keys=True))
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--plan", type=Path)
    ap.add_argument("--manifest", type=Path)
    ap.add_argument("--output", type=Path)
    args = ap.parse_args()
    try:
        if args.self_test:
            if args.plan or args.manifest or args.output:
                raise p.ProtocolError("--self-test is standalone")
            return self_test()
        if args.plan is None or args.manifest is None or args.output is None:
            raise p.ProtocolError("--plan, --manifest, and --output are required")
        plan_path = args.plan.resolve()
        manifest_path = args.manifest.resolve()
        plan = p.load_json(plan_path)
        p.check_plan_bindings(plan, plan_path)
        packets = _load_packets(plan_path, manifest_path, plan)
        result = _reduce_packets(plan_path, manifest_path, plan, packets)
        p.write_json_exclusive(args.output.resolve(), result)
        print(json.dumps({"status": "pass", "output": str(args.output.resolve()), "sha256": p.sha256(args.output.resolve())}, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"0973 reducer rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
