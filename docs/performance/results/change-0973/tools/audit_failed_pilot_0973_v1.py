#!/usr/bin/env python3
"""Bounded read-only audit of the failed 0973 child-01 pilot.

This helper does not rerun perf, DWARF decoding, a workload, or a build.  It
checks the retained packet bindings, independently streams the corrected/no-op
script and MMAP dump for sample/loss arithmetic, and summarizes the already
decoded owner maps from the packet.  The result is deliberately one-child
descriptive evidence; it cannot qualify the planned all-child diagnostic.
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import math
import platform
import re
import sys
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0973")
CAPTURE = ROOT / "capture/instruction-profile-v2"
CHILD = CAPTURE / "child-01"
PACKET_PATH = CHILD / "child-01.packet.json"
MANIFEST_PATH = CAPTURE / "capture-manifest-v2.json"
ROOT_TERMINAL_PATH = ROOT / "receipts/root-terminal-0973-v1.json"
PLAN_PATH = ROOT / "plans/plan-0973-final-v2.json"
ADR_RECEIPT_PATH = ROOT / "receipts/adr-goal-revalidation-0973-v1.json"
TARGET_DSO = "/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile"
TARGET_BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
MARKER = "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
PERIOD = 1_000_000

HEADER_RE = re.compile(r"^(?P<pid>\d+)/(?P<tid>\d+)\s+\[(?P<cpu>\d+)\]\s+\S+:\s+(?P<period>\d+)\s+instructions:u:")
SCRIPT_LOST_RE = re.compile(r"PERF_RECORD_LOST\s+lost\s+(?P<lost>\d+)")
MMAP_LOST_RE = re.compile(r"PERF_RECORD_LOST:.*\blost:(?P<lost>\d+)")
MMAP_LOST_SAMPLES_RE = re.compile(r"PERF_RECORD_LOST_SAMPLES:.*lost samples\s*:\s*(?P<lost>\d+)")
MMAP_SAMPLE_RE = re.compile(r": PERF_RECORD_SAMPLE(?:\b|:)")


class AuditError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    path = path.resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def verify_binding(label: str, row: dict[str, Any]) -> dict[str, Any]:
    path = Path(row["path"]).resolve()
    if not path.is_file():
        raise AuditError(f"{label}: missing {path}")
    actual = binding(path)
    if actual["path"] != str(path) or actual["bytes"] != row.get("bytes") or actual["sha256"] != row.get("sha256"):
        raise AuditError(f"{label}: binding mismatch")
    return actual


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def percentile(values: list[int], q: float) -> float:
    ordered = sorted(float(value) for value in values)
    if not ordered:
        raise AuditError("empty timing vector")
    position = (len(ordered) - 1) * q
    lower, upper = math.floor(position), math.ceil(position)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def scan_script(path: Path) -> dict[str, Any]:
    digest = hashlib.sha256()
    sample_count = 0
    total_period = 0
    lost_values: list[int] = []
    marker_count = 0
    marker_period = 0
    empty_callchains = 0
    current_period: int | None = None
    current_frames = 0
    current_marker = False

    def finish_sample() -> None:
        nonlocal empty_callchains, marker_count, marker_period
        if current_period is None:
            return
        if current_frames == 0:
            empty_callchains += 1
        if current_marker:
            marker_count += 1
            marker_period += current_period

    with path.open("rb") as raw:
        for line_bytes in raw:
            digest.update(line_bytes)
            line = line_bytes.decode("utf-8", errors="replace").rstrip("\r\n")
            match = HEADER_RE.match(line)
            if match:
                finish_sample()
                sample_count += 1
                current_period = int(match.group("period"))
                total_period += current_period
                current_frames = 0
                current_marker = False
                continue
            lost = SCRIPT_LOST_RE.search(line)
            if lost:
                lost_values.append(int(lost.group("lost")))
                continue
            if current_period is not None and line[:1].isspace() and "(" in line:
                current_frames += 1
                if f" {MARKER}+" in line and f"({TARGET_DSO})" in line:
                    current_marker = True
    finish_sample()
    return {
        "binding": {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": digest.hexdigest()},
        "sample_count": sample_count,
        "total_period": total_period,
        "period_values": {"unique": [PERIOD] if sample_count else [], "all_equal_configured": total_period == sample_count * PERIOD},
        "script_lost_record_count": len(lost_values),
        "script_lost_values": lost_values,
        "script_lost_sum": sum(lost_values),
        "marker_sample_count": marker_count,
        "marker_period": marker_period,
        "marker_period_share": marker_period / total_period if total_period else None,
        "empty_callchain_count": empty_callchains,
    }


def scan_mmap_dump(path: Path) -> dict[str, Any]:
    digest = hashlib.sha256()
    lost_values: list[int] = []
    lost_samples_values: list[int] = []
    sample_records = 0
    with path.open("rb") as raw:
        for line_bytes in raw:
            digest.update(line_bytes)
            line = line_bytes.decode("utf-8", errors="replace")
            match = MMAP_LOST_RE.search(line)
            if match:
                lost_values.append(int(match.group("lost")))
            match = MMAP_LOST_SAMPLES_RE.search(line)
            if match:
                lost_samples_values.append(int(match.group("lost")))
            if MMAP_SAMPLE_RE.search(line):
                sample_records += 1
    return {
        "binding": {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": digest.hexdigest()},
        "sample_record_count": sample_records,
        "lost_record_count": len(lost_values),
        "lost_record_values": lost_values,
        "lost_record_sum": sum(lost_values),
        "lost_samples_record_count": len(lost_samples_values),
        "lost_samples_record_values": lost_samples_values,
        "lost_samples_record_sum": sum(lost_samples_values),
    }


def owner_summary(mapping: dict[str, Any], marker_period: int) -> dict[str, Any]:
    values = {str(key): int(value) for key, value in mapping.items()}
    target_prefix = f"{TARGET_DSO}::"
    target_period = sum(value for key, value in values.items() if key.startswith(target_prefix))
    external_period = sum(value for key, value in values.items() if not key.startswith(target_prefix))
    rows = []
    for owner, period in sorted(values.items(), key=lambda item: (-item[1], item[0]))[:20]:
        rows.append({
            "owner": owner,
            "period": period,
            "sample_count_at_fixed_period": period // PERIOD,
            "share_of_marker": period / marker_period if marker_period else None,
            "dso_class": "target" if owner.startswith(target_prefix) else "external",
        })
    return {
        "bucket_count": len(values),
        "period_sum": sum(values.values()),
        "target_bucket_count": sum(owner.startswith(target_prefix) for owner in values),
        "external_bucket_count": sum(not owner.startswith(target_prefix) for owner in values),
        "target_period": target_period,
        "external_period": external_period,
        "target_share_of_marker": target_period / marker_period if marker_period else None,
        "external_share_of_marker": external_period / marker_period if marker_period else None,
        "top20": rows,
    }


def audit() -> dict[str, Any]:
    packet = load_json(PACKET_PATH)
    manifest = load_json(MANIFEST_PATH)
    root_terminal = load_json(ROOT_TERMINAL_PATH)
    plan = load_json(PLAN_PATH)
    adr_receipt = load_json(ADR_RECEIPT_PATH)
    if packet.get("packet_id") != "child-01" or packet.get("status") != "complete":
        raise AuditError("child-01 packet identity/status mismatch")
    if manifest.get("status") != "failed" or manifest.get("failure", {}).get("child") != 1:
        raise AuditError("capture manifest is not the retained child-01 failure")
    if root_terminal.get("returncode") != 2 or root_terminal.get("children_completed") != 1:
        raise AuditError("root terminal receipt is not the expected stopped run")
    if adr_receipt.get("status") != "pass":
        raise AuditError("ADR/GOAL revalidation receipt is not passing")

    artifact_rows = packet.get("artifacts") or {}
    artifact_bindings: dict[str, Any] = {}
    for name, row in sorted(artifact_rows.items()):
        artifact_bindings[name] = verify_binding(f"packet artifact {name}", row)

    corrected_path = Path(artifact_rows["corrected_script"]["path"])
    stock_path = Path(artifact_rows["stock_script"]["path"])
    noop_path = Path(artifact_rows["noop_script"]["path"])
    corrected = scan_script(corrected_path)
    stock = scan_script(stock_path)
    noop = scan_script(noop_path)
    mmap = scan_mmap_dump(Path(artifact_rows["mmap_dump"]["path"]))
    raw = load_json(Path(artifact_rows["raw_stack"]["path"]))
    report = packet["report"]
    elapsed = [int(value) for value in report["elapsed_ns"]]
    decoded = packet["perf"]["corrected"]

    if corrected["sample_count"] != raw["sample_count"] or corrected["sample_count"] != decoded["sample_count"]:
        raise AuditError("corrected/sample summary count mismatch")
    if corrected["total_period"] != raw["sample_count"] * PERIOD or decoded["total_period"] != corrected["total_period"]:
        raise AuditError("observed period conservation mismatch")
    if mmap["sample_record_count"] != raw["sample_count"]:
        raise AuditError("MMAP sample record count disagrees with raw stack summary")
    if corrected["marker_sample_count"] != decoded["marker_sample_count"] or corrected["marker_period"] != decoded["marker_period"]:
        raise AuditError("corrected marker scan disagrees with decoded summary")
    if corrected["script_lost_sum"] != decoded["lost_events"]:
        raise AuditError("corrected script loss lines disagree with decoded loss count")
    if mmap["lost_record_sum"] != corrected["script_lost_sum"] or mmap["lost_samples_record_sum"] != corrected["script_lost_sum"]:
        raise AuditError("MMAP/script loss evidence disagrees")
    if stock["binding"]["sha256"] != noop["binding"]["sha256"] or stock["sample_count"] != noop["sample_count"] or stock["total_period"] != noop["total_period"]:
        raise AuditError("stock/no-op replay identity changed")

    physical = decoded.get("physical_leaf_owner_period", {})
    inclusive = decoded.get("inclusive_marker_ancestor_period", {})
    marker_period = int(decoded["marker_period"])
    physical_sum = sum(int(value) for value in physical.values()) + int(decoded.get("physical_leaf_unknown_period", 0))
    if physical_sum != marker_period or int(decoded.get("physical_leaf_classified_period", -1)) != marker_period:
        raise AuditError("physical marker ledger does not conserve M")

    input_bindings = {
        "packet": binding(PACKET_PATH),
        "manifest": binding(MANIFEST_PATH),
        "root_terminal": binding(ROOT_TERMINAL_PATH),
        "plan": binding(PLAN_PATH),
        "adr_revalidation": binding(ADR_RECEIPT_PATH),
        "binary": binding(Path(plan["binary"]["path"])),
        "fixture": binding(Path(plan["fixture"]["path"])),
        "perf_data": artifact_bindings["perf_data"],
        "mmap_dump": artifact_bindings["mmap_dump"],
        "corrected_script": corrected["binding"],
        "raw_stack": artifact_bindings["raw_stack"],
        "report": artifact_bindings["report"],
    }
    loss_sum = corrected["script_lost_sum"]
    observed_period = corrected["total_period"]
    return {
        "schema": "litchi.performance.0973.failed_pilot_audit.v1",
        "status": "complete_failed_pilot",
        "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "analysis_only": True,
        "diagnostic_only": True,
        "qualification_claim": False,
        "speed_claim": False,
        "adoption_claim": False,
        "all_child_claim": False,
        "no_new_capture_build_or_workload": True,
        "scope": "child-01 pilot only; retained failed run; descriptive raw/report/decoded audit",
        "inputs": input_bindings,
        "capture_stop": {
            "root_terminal_returncode": root_terminal["returncode"],
            "children_completed": root_terminal["children_completed"],
            "children_not_started": root_terminal["children_not_started"],
            "manifest_status": manifest["status"],
            "manifest_failure": manifest["failure"],
            "no_retry": root_terminal["no_retry"],
        },
        "child_terminal": {
            "returncode": packet["terminal"]["returncode"],
            "timed_out": packet["terminal"]["timed_out"],
            "native_ld_preload": packet["terminal"]["native_ld_preload"],
            "process_pid": packet["terminal"]["process_pid"],
            "exec_pid": packet["launch"]["exec_pid"],
            "actual_affinity": packet["launch"]["actual_affinity"],
        },
        "gate": packet["gate"],
        "loss_evidence": {
            "configured_period": PERIOD,
            "perf_record_lost_record_count": mmap["lost_record_count"],
            "perf_record_lost_values": mmap["lost_record_values"],
            "perf_record_lost_aggregate": mmap["lost_record_sum"],
            "perf_record_lost_samples_record_count": mmap["lost_samples_record_count"],
            "perf_record_lost_samples_values": mmap["lost_samples_record_values"],
            "perf_record_lost_samples_aggregate": mmap["lost_samples_record_sum"],
            "script_lost_record_count": corrected["script_lost_record_count"],
            "script_lost_values": corrected["script_lost_values"],
            "script_lost_aggregate": loss_sum,
            "nominal_period_units_if_one_period_per_lost_sample": loss_sum * PERIOD,
            "observed_period_excludes_lost_samples": True,
            "interpretation": "Two PERF_RECORD_LOST records (4+4) and one PERF_RECORD_LOST_SAMPLES summary (8) agree. The 8 lost samples have no observed instruction period; the nominal configured-period product is not added to D.",
        },
        "sample_evidence": {
            "raw_stack": {
                "sample_count": raw["sample_count"],
                "raw_event_count": raw["raw_event_count"],
                "unconsumed_raw_count": raw["unconsumed_raw_count"],
                "record_type_counts": raw["record_type_counts"],
                "pid_tid_cpu_counts": raw["pid_tid_cpu_counts"],
                "period": raw["period"],
                "period_valid": raw["period_valid"],
                "violations": raw["violations"],
            },
            "mmap_sample_record_count": mmap["sample_record_count"],
            "corrected_script": {
                "sample_count": corrected["sample_count"],
                "total_period": corrected["total_period"],
                "marker_sample_count": corrected["marker_sample_count"],
                "marker_period": corrected["marker_period"],
                "marker_period_share": corrected["marker_period_share"],
                "empty_callchain_count": corrected["empty_callchain_count"],
                "decoded_empty_callchain_count": decoded["empty_callchain_count"],
                "unresolved_frames": decoded["unresolved_frames"],
                "explicit_stack_truncated_count": decoded["explicit_stack_truncated_count"],
                "physical_leaf_unknown_period": decoded["physical_leaf_unknown_period"],
                "physical_leaf_unknown_samples": decoded["physical_leaf_unknown_samples"],
                "physical_leaf_classified_period": decoded["physical_leaf_classified_period"],
                "physical_leaf_conservation": decoded["physical_leaf_conservation"],
            },
            "stock_noop": {
                "stock_sample_count": stock["sample_count"],
                "noop_sample_count": noop["sample_count"],
                "stock_total_period": stock["total_period"],
                "noop_total_period": noop["total_period"],
                "script_sha_equal": stock["binding"]["sha256"] == noop["binding"]["sha256"],
            },
            "observed_period_conservation": {
                "sample_count_times_period": corrected["sample_count"] * PERIOD,
                "corrected_total_period": observed_period,
                "passed": corrected["sample_count"] * PERIOD == observed_period,
            },
        },
        "report": {
            "samples_requested": report["samples_requested"],
            "warmups": report["warmups"],
            "elapsed_count": len(elapsed),
            "elapsed_min_ns": min(elapsed),
            "elapsed_max_ns": max(elapsed),
            "p50_ns": percentile(elapsed, 0.5),
            "p95_ns": percentile(elapsed, 0.95),
            "p99_ns": percentile(elapsed, 0.99),
            "mean_ns": sum(elapsed) / len(elapsed),
            "timing_claim": report["timing_claim"],
            "region_definition": report["region_definition"],
            "source_bytes": report["source_bytes"],
            "output_bytes": report["output_bytes"],
            "output_capacity": report["output_capacity"],
            "exact_output_equal_reference": report["exact_output_equal_reference"],
            "correctness_all_true": all(value is True for value in report["correctness"].values()),
            "route_count": len(report["route_outputs"]),
            "route_output_sha256s": sorted({row["output_sha256"] for row in report["route_outputs"]}),
        },
        "owner_coverage_observed": {
            "marker_period_denominator": marker_period,
            "physical_leaf_proxy": owner_summary(physical, marker_period),
            "physical_leaf_unknown_period": decoded["physical_leaf_unknown_period"],
            "inclusive_marker_ancestors": owner_summary(inclusive, marker_period),
            "inclusive_sum_is_nonadditive": sum(int(value) for value in inclusive.values()) > marker_period,
            "inclusive_sum_period": sum(int(value) for value in inclusive.values()),
            "root_resolved_owner_period": decoded["root_resolved_owner_period"],
            "owner_key_format": decoded["owner_key_format"],
            "scope": "one failed child; physical frame-0 proxy and overlapping resolved marker-chain tags; no source or phase attribution",
        },
        "artifact_bindings": artifact_bindings,
        "audit_helper": {
            "path": str(Path(__file__).resolve()),
            "sha256": sha256(Path(__file__).resolve()),
            "argv": sys.argv,
            "python": platform.python_version(),
        },
        "interpretation": "The pilot has usable descriptive report and corrected marker-owner observations, but lost events make the frozen child gate fail. The retained D is observed sample period only; no all-child reduction, speed, causal, exact timed-region, source-owner, or adoption conclusion is valid.",
    }


def markdown(result: dict[str, Any]) -> str:
    loss = result["loss_evidence"]
    sample = result["sample_evidence"]["corrected_script"]
    physical = result["owner_coverage_observed"]["physical_leaf_proxy"]
    inclusive = result["owner_coverage_observed"]["inclusive_marker_ancestors"]
    lines = [
        "# 0973 failed child-01 pilot audit",
        "",
        "This is a one-child descriptive audit. The pilot gate failed on lost events; it does not qualify the planned all-child diagnostic or support a speed/adoption claim.",
        "",
        f"- Root runner return code: `{result['capture_stop']['root_terminal_returncode']}`; children completed: `{result['capture_stop']['children_completed']}`; retry: `{result['capture_stop']['no_retry']}`.",
        f"- Child gate: `{result['gate']}`.",
        f"- Retained corrected samples: `{sample['sample_count']}` at period `{loss['configured_period']}`; observed D: `{sample['total_period']}`.",
        f"- Marker samples/period/share: `{sample['marker_sample_count']}` / `{sample['marker_period']}` / `{sample['marker_period_share']:.9f}`.",
        f"- Empty callchains: `{sample['empty_callchain_count']}`; unresolved frames: `{sample['unresolved_frames']}`; physical unknown leaf period: `{sample['physical_leaf_unknown_period']}`.",
        "",
        "## Loss accounting",
        "",
        f"The MMAP dump contains `{loss['perf_record_lost_record_count']}` `PERF_RECORD_LOST` records with values `{loss['perf_record_lost_values']}` and one `PERF_RECORD_LOST_SAMPLES` record with value `{loss['perf_record_lost_samples_values']}`. The corrected script contains `{loss['script_lost_record_count']}` loss lines with values `{loss['script_lost_values']}`. All agree on `{loss['script_lost_aggregate']}` lost samples.",
        "",
        "The configured period gives a nominal 8,000,000-period product for eight lost samples, but that is not observed period and is excluded from D.",
        "",
        "## Timing and owner coverage",
        "",
        f"The report retained 80 measured rows plus 5 warmups. Timing p50/p95/p99 are `{result['report']['p50_ns']:.1f}` / `{result['report']['p95_ns']:.1f}` / `{result['report']['p99_ns']:.1f}` ns. Correctness and exact output equality were retained.",
        f"The disjoint physical-leaf proxy has `{physical['bucket_count']}` buckets, target-DSO period `{physical['target_period']}` and external-DSO period `{physical['external_period']}`; its buckets plus unknown conserve marker period.",
        f"The overlapping inclusive marker-chain map has `{inclusive['bucket_count']}` tags and sum `{inclusive['period_sum']}`, while marker M is `{sample['marker_period']}`; these values must not be summed as phases.",
        "",
        "Top observed physical leaf buckets:",
        "",
        "| period | owner | class |",
        "|---:|---|---|",
    ]
    for row in physical["top20"][:12]:
        lines.append(f"| {row['period']} | `{row['owner']}` | {row['dso_class']} |")
    lines += [
        "",
        "Top observed inclusive ancestor tags:",
        "",
        "| period | owner | class |",
        "|---:|---|---|",
    ]
    for row in inclusive["top20"][:12]:
        lines.append(f"| {row['period']} | `{row['owner']}` | {row['dso_class']} |")
    lines += [
        "",
        "All paths and hashes are recorded in the JSON receipt. Raw captures remain external and unchanged.",
        "",
    ]
    return "\n".join(lines)


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--output-json", type=Path, default=ROOT / "analysis/failed-pilot-child01-audit-v1.json")
    ap.add_argument("--output-md", type=Path, default=ROOT / "analysis/failed-pilot-child01-audit-v1.md")
    ap.add_argument("--output-receipt", type=Path, default=ROOT / "receipts/failed-pilot-audit-0973-v1.json")
    args = ap.parse_args()
    result = audit()
    args.output_json.parent.mkdir(parents=True, exist_ok=True)
    args.output_json.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    args.output_md.write_text(markdown(result), encoding="utf-8")
    receipt = {
        "schema": "litchi.performance.0973.failed_pilot_audit_receipt.v1",
        "status": "pass",
        "created_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "audit_status": result["status"],
        "analysis_only": True,
        "diagnostic_only": True,
        "qualification_claim": False,
        "speed_claim": False,
        "adoption_claim": False,
        "no_new_capture_build_or_workload": True,
        "helper": binding(Path(__file__).resolve()),
        "summary_json": binding(args.output_json.resolve()),
        "summary_markdown": binding(args.output_md.resolve()),
        "packet": result["inputs"]["packet"],
        "manifest": result["inputs"]["manifest"],
        "root_terminal": result["inputs"]["root_terminal"],
        "loss_record_and_period_check": result["loss_evidence"],
        "scope": result["scope"],
    }
    args.output_receipt.parent.mkdir(parents=True, exist_ok=True)
    args.output_receipt.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": result["status"], "json": str(args.output_json.resolve()), "markdown": str(args.output_md.resolve()), "receipt": str(args.output_receipt.resolve()), "helper_sha256": result["audit_helper"]["sha256"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
