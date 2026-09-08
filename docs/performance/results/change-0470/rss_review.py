#!/usr/bin/env python3
"""Verify and summarize the separate 0470 normal-process RSS follow-up."""

from __future__ import annotations

import datetime
import json
import math
from pathlib import Path
from typing import Any, Mapping

import analyze
import verify


ROOT = Path(__file__).resolve().parent
# Compare historical command destinations, independently of the replay location.
CAPTURE_ROOT = Path("/home/zhuhe/code/litchi/docs/performance/results/change-0470")
SCHEMA = "litchi-0470-rss-review-v1"
LANES = ("rss-A1", "rss-B1", "rss-B2", "rss-A2")
MAIN_LANES = verify.LANES
ROLE_FOR_LANE = {
    "rss-A1": "control",
    "rss-A2": "control",
    "rss-B1": "candidate",
    "rss-B2": "candidate",
}
RSS_PAIRS = (
    ("control_repetition", "rss-A1", "rss-A2"),
    ("candidate_repetition", "rss-B1", "rss-B2"),
    ("a1_control_to_b1_candidate", "rss-A1", "rss-B1"),
    ("a2_control_to_b2_candidate", "rss-A2", "rss-B2"),
)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise verify.VerificationError(message)


def _timestamp(value: Any, label: str) -> datetime.datetime:
    return verify._timestamp(value, label)


def _delta_percent(baseline: float, current: float) -> float:
    if baseline == 0:
        return 0.0 if current == 0 else math.inf
    return (current / baseline - 1.0) * 100.0


def _rss(root: Path, lane: str) -> dict[str, Any]:
    return analyze.parse_gnu_rss(
        root / lane / "resource.log",
        root=root,
        latency_comparison="descriptive_only",
    )


def _rss_drift(
    rss: Mapping[str, Mapping[str, Any]],
    name: str,
    baseline_lane: str,
    current_lane: str,
) -> dict[str, Any]:
    baseline = rss[baseline_lane]
    current = rss[current_lane]
    baseline_bytes = int(baseline["maximum_resident_set_bytes"])
    current_bytes = int(current["maximum_resident_set_bytes"])
    delta = _delta_percent(float(baseline_bytes), float(current_bytes))
    return {
        "name": name,
        "baseline_lane": baseline_lane,
        "current_lane": current_lane,
        "baseline_kib": baseline["maximum_resident_set_kib"],
        "current_kib": current["maximum_resident_set_kib"],
        "baseline_bytes": baseline_bytes,
        "current_bytes": current_bytes,
        "delta_percent": delta if math.isfinite(delta) else None,
        "delta_is_infinite": math.isinf(delta),
    }


def _verify_receipt(
    root: Path,
    lane: str,
    protocol: Mapping[str, Any],
    binding: Mapping[str, Any],
) -> dict[str, Any]:
    lane_root = root / lane
    started = verify.read_json(lane_root / "started.json", f"{lane}.started.json")
    receipt = verify.read_json(lane_root / "receipt.json", f"{lane}.receipt.json")
    for item, label in ((started, "started"), (receipt, "receipt")):
        require(item.get("schema") == "litchi-0470-rss-capture-v1", f"{lane}.{label}: schema differs")
        require(item.get("lane") == lane and item.get("role") == ROLE_FOR_LANE[lane], f"{lane}.{label}: lane or role differs")
        require(item.get("revision") == binding["revision"], f"{lane}.{label}: revision differs")
        require(item.get("binary_sha256") == binding["binary_sha256"], f"{lane}.{label}: binary differs")
        require(item.get("binding_sha256") == verify.sha256_file(root / f"{ROLE_FOR_LANE[lane]}-binding.json")[0], f"{lane}.{label}: binding differs")
        require(item.get("driver_sha256") == protocol["capture_driver_sha256"], f"{lane}.{label}: driver differs")
        require(item.get("protocol_sha256") == verify.sha256_file(root / "rss-protocol.json")[0], f"{lane}.{label}: RSS protocol differs")
        require(item.get("main_protocol_sha256") == protocol["main_protocol_sha256"], f"{lane}.{label}: main protocol differs")
    require(receipt.get("exit_code") == 0, f"{lane}: RSS capture failed")
    for field in ("clean_before", "clean_after", "binary_unchanged", "report_metadata_matches_clean_role"):
        require(receipt.get(field) is True, f"{lane}: missing {field} attestation")
    samples = int(protocol["samples"])
    warmups = int(protocol["warmups"])
    require(receipt.get("samples") == samples and receipt.get("warmups") == warmups, f"{lane}: sample configuration differs")
    require(started.get("argv") == receipt.get("argv") and started.get("cwd") == receipt.get("cwd"), f"{lane}: started/final command differs")
    argv = receipt.get("argv")
    require(isinstance(argv, list) and all(isinstance(value, str) for value in argv), f"{lane}: argv is malformed")
    required = ["taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o", str(CAPTURE_ROOT / lane / "resource.log")]
    require(argv[: len(required)] == required, f"{lane}: command prefix differs")
    require(str(binding["binary_path"]) in argv, f"{lane}: binary argv differs")
    for token, value in (
        ("--workers", str(protocol["workers"])),
        ("--warmup", str(warmups)),
        ("--samples", str(samples)),
        ("--case", ",".join(protocol["cases"])),
        ("--xlsx-shape", ",".join(protocol["shapes"])),
    ):
        require(argv.count(token) == 1 and argv[argv.index(token) + 1] == value, f"{lane}: {token} differs")
    for name in ("report.json", "corpus-catalog.json", "resource.log", "stdout.log", "stderr.log", "started.json"):
        verify.regular(lane_root / name, f"{lane}/{name}")
    artifacts = verify.obj(receipt.get("artifacts"), f"{lane}.artifacts")
    require(set(artifacts) == {"report.json", "corpus-catalog.json", "resource.log", "stdout.log", "stderr.log", "started.json"}, f"{lane}: artifact set differs")
    for name, metadata in artifacts.items():
        path = verify.bundle_file(root, f"{lane}/{name}", f"{lane}.artifacts.{name}")
        metadata = verify.obj(metadata, f"{lane}.artifacts.{name}")
        require(verify.digest(metadata.get("sha256"), f"{lane}.{name}.sha256") == verify.sha256_file(path)[0], f"{lane}: artifact hash differs for {name}")
        require(verify.integer(metadata.get("bytes"), f"{lane}.{name}.bytes") == path.stat().st_size, f"{lane}: artifact size differs for {name}")
    return receipt


def _verify_chronology(
    root: Path,
    receipts: Mapping[str, Mapping[str, Any]],
) -> dict[str, Any]:
    main_receipts = {
        lane: verify.read_json(root / lane / "receipt.json", f"{lane}.receipt.json")
        for lane in MAIN_LANES
    }
    main_intervals = {
        lane: (_timestamp(item["started_utc"], lane), _timestamp(item["finished_utc"], lane))
        for lane, item in main_receipts.items()
    }
    rss_intervals = {
        lane: (_timestamp(item["started_utc"], lane), _timestamp(item["finished_utc"], lane))
        for lane, item in receipts.items()
    }
    for lane, (start, finish) in rss_intervals.items():
        require(finish >= start, f"{lane}: interval is reversed")
        for main_lane, (main_start, main_finish) in main_intervals.items():
            require(not (start < main_finish and finish > main_start), f"{lane}: overlaps main capture {main_lane}")
    for previous, current in zip(LANES, LANES[1:]):
        require(rss_intervals[current][0] >= rss_intervals[previous][1], f"RSS capture order overlaps: {previous}, {current}")
    last_main_finish = max(finish for _, finish in main_intervals.values())
    require(min(start for start, _ in rss_intervals.values()) >= last_main_finish, "RSS follow-up starts before main heap capture completion")
    return {
        "main_captures_complete_before_followup": True,
        "rss_capture_order": list(LANES),
        "rss_intervals_non_overlapping": True,
        "rss_followup_after_main_captures": True,
    }


def evaluate(root: Path = ROOT) -> dict[str, Any]:
    root = root.resolve()
    protocol = verify.read_json(root / "rss-protocol.json", "rss-protocol.json")
    require(protocol.get("schema") == "litchi-0470-rss-protocol-v1", "RSS protocol schema differs")
    require(protocol.get("order") == list(LANES), "RSS protocol order differs")
    require(protocol.get("samples") == 100 and protocol.get("warmups") == 5, "RSS protocol sample configuration differs")
    require(protocol.get("cpu") == 2 and protocol.get("workers") == 1, "RSS protocol CPU or worker differs")
    require(protocol.get("cases") == ["xlsx_one_cell_commit_save", "xlsx_one_percent_commit_save"], "RSS protocol cases differ")
    require(protocol.get("shapes") == ["tiny", "medium", "dense-wide"], "RSS protocol shapes differ")
    require(protocol.get("performance_claim") == "none; descriptive RSS drift follow-up only", "RSS protocol claim differs")
    require(verify.sha256_file(root / "rss_capture.py")[0] == protocol.get("capture_driver_sha256"), "RSS protocol driver hash differs")
    require(verify.sha256_file(root / "protocol.json")[0] == protocol.get("main_protocol_sha256"), "RSS protocol main hash differs")

    bindings = {role: verify.verify_role_binding(root, role) for role in ("control", "candidate")}
    reports = {
        lane: verify.read_json(root / lane / "report.json", f"{lane}/report.json")
        for lane in LANES
    }
    for lane, report in reports.items():
        verify.verify_report_metadata(
            report,
            f"{lane}.report",
            bindings[ROLE_FOR_LANE[lane]],
            int(protocol["samples"]),
            int(protocol["warmups"]),
            cases=protocol["cases"],
            shapes=protocol["shapes"],
        )
    receipts = {
        lane: _verify_receipt(root, lane, protocol, bindings[ROLE_FOR_LANE[lane]])
        for lane in LANES
    }
    chronology = _verify_chronology(root, receipts)

    ordered_reports = [reports[lane] for lane in LANES]
    abba, abba_path = analyze.load_abba(root)
    try:
        strict_summary = abba.summarize_reports(
            reports=ordered_reports,
            profile="current-v1",
            cases=protocol["cases"],
            shapes=protocol["shapes"],
        )
        strict_status = "valid"
        strict_error = None
    except (AttributeError, KeyError, TypeError, ValueError) as error:
        # Preserve a strict rejection and its reason.  Never erase source
        # counters or project rows to manufacture an ABBA acceptance.
        strict_summary = None
        strict_status = "rejected"
        strict_error = str(error)

    rss = {lane: _rss(root, lane) for lane in LANES}
    paired_rss = {
        name: _rss_drift(rss, name, baseline_lane, current_lane)
        for name, baseline_lane, current_lane in RSS_PAIRS
    }
    return {
        "schema": SCHEMA,
        "scope": "Separate descriptive normal-process RSS follow-up using the original six-row 100-sample ABBA workload; no registered latency or RSS claim",
        "strict_abba_status": strict_status,
        "strict_abba_error": strict_error,
        "abba": strict_summary,
        "rss": rss,
        "paired_normal_rss_drift": paired_rss,
        "report_metadata": {
            lane: {
                "revision": reports[lane]["environment"]["git_revision"],
                "binary_sha256": reports[lane]["binary_identity"]["binary_sha256"],
                "binary_bytes": reports[lane]["binary_identity"]["binary_bytes"],
                "worktree_dirty": reports[lane]["environment"]["git_worktree_dirty"],
            }
            for lane in LANES
        },
        "inputs": {
            "reports": [analyze.file_binding(root / lane / "report.json", root) for lane in LANES],
            "resource_logs": [analyze.file_binding(root / lane / "resource.log", root) for lane in LANES],
            "receipts": [analyze.file_binding(root / lane / "receipt.json", root) for lane in LANES],
            "abba_summary": analyze.file_binding(abba_path, root),
            "protocol": analyze.file_binding(root / "rss-protocol.json", root),
            "capture_driver": analyze.file_binding(root / "rss_capture.py", root),
            "main_protocol": analyze.file_binding(root / "protocol.json", root),
        },
        "chronology": chronology,
    }


if __name__ == "__main__":
    result = evaluate()
    (ROOT / "rss-review-summary.json").write_text(
        json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    print(json.dumps({"status": "pass", "strict_abba_status": result["strict_abba_status"]}, sort_keys=True))
