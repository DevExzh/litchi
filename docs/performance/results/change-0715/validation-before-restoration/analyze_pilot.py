#!/usr/bin/env python3
"""Strictly validate and decide the 0715 paired DOCX publication pilot.

The capture driver owns custody.  This file rechecks every receipt, invokes
the reviewed 0709 elapsed/ordinary-save/allocation validators, compares only
paired native observations, and retains allocator request diagnostics.  A
failed performance gate is emitted as ``decision.accepted: false``; malformed
or contradictory evidence aborts without producing a decision artifact.
"""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import statistics
import sys
from typing import Any

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import pilot as CAPTURE  # noqa: E402


def load(path: Path) -> Any:
    if not path.is_file() or path.is_symlink():
        raise AssertionError(f"missing evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AssertionError(f"invalid JSON {path}: {error}") from error


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def digest(value: Any) -> str:
    raw = (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()
    return hashlib.sha256(raw).hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def legacy_module() -> Any:
    path = REPO / "docs/performance/results/change-0709/analyze.py"
    spec = importlib.util.spec_from_file_location("ordinary_save_0709_for_0715", path)
    require(spec is not None and spec.loader is not None, f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


LEGACY = legacy_module()
PHASES = CAPTURE.PHASES
LANES = CAPTURE.LANES
STAGES = CAPTURE.STAGE_LABELS
PAIRS = {"pair-1": ("baseline-A1", "candidate-B1"),
         "pair-2": ("baseline-A2", "candidate-B2")}


def source_delta(baseline: dict[str, str], candidate: dict[str, str]) -> list[str]:
    return sorted(path for path in set(baseline) | set(candidate)
                  if baseline.get(path) != candidate.get(path))


def validate_receipt(job: dict[str, Any], build: dict[str, Any], plan: dict[str, Any],
                     baseline: dict[str, str], candidate: dict[str, str],
                     constraints_sha: str | None) -> dict[str, Any]:
    name = job["name"]
    receipt = load(HERE / f"{name}.receipt.json")
    metadata = CAPTURE.STAGE_METADATA[job["stage"]]
    for key, expected in {
        "schema_version": 1, "packet": plan["packet"], "name": name,
        "stage": job["stage"], "source_label": metadata["source"],
        "build_label": metadata["build"], "pair": metadata["pair"],
        "repeat": metadata["repeat"], "lane": job["lane"],
        "stage_order": metadata["order"], "stage_order_index": job["stage_order_index"],
        "corpus_id": job["corpus_id"], "phase": job["phase"], "case": job["case"],
        "samples": job["samples"], "warmup": job["warmup"], "cpu": plan["cpu"],
    }.items():
        require(receipt.get(key) == expected, f"{name}: receipt {key} changed")
    require(receipt.get("exit_code") == 0, f"{name}: child failed")
    require(receipt.get("binary_sha256") == build["binary_sha256"]
            and receipt.get("binary_bytes") == build["binary_bytes"],
            f"{name}: binary identity changed")
    require(receipt.get("build_record") == build["build_record"]
            and receipt.get("build_record_sha256") == build["build_record_sha256"],
            f"{name}: build record binding changed")
    require(receipt.get("pilot_plan_sha256") == sha(HERE / "pilot-plan.json"),
            f"{name}: plan binding changed")
    require(receipt.get("script_sha256") == sha(HERE / "pilot.py"),
            f"{name}: capture script binding changed")
    require(receipt.get("constraints_sha256") == constraints_sha,
            f"{name}: constraints binding changed")
    require(receipt.get("command") == CAPTURE.expected_command(job, build, plan),
            f"{name}: command binding changed")

    retained = receipt.get("retained_binary_source")
    require(isinstance(retained, dict)
            and retained.get("manifest") == build["source_manifest"]
            and retained.get("manifest_sha256") == build["source_manifest_sha256"]
            and retained.get("source_census_sha256") == digest(build["source"]),
            f"{name}: retained source binding changed")

    live_label = "baseline" if job["stage"] == "baseline-A1" else "candidate"
    live_source = baseline if live_label == "baseline" else candidate
    current = receipt.get("current_checkout_source")
    require(isinstance(current, dict)
            and current.get("manifest") == f"source-{live_label}.json"
            and current.get("manifest_sha256") == sha(HERE / f"source-{live_label}.json")
            and current.get("unchanged_during_child") is True,
            f"{name}: live source custody failed")
    for suffix, key in (("source-before.json", "before_artifact"),
                        ("source-after.json", "after_artifact")):
        filename = f"{name}.{suffix}"
        require(current.get(key) == filename, f"{name}: source artifact name changed")
        artifact = HERE / filename
        require(load(artifact) == live_source and sha(artifact) == current.get(
            "before_file_sha256" if key == "before_artifact" else "after_file_sha256"),
                f"{name}: source artifact changed")
    require(current.get("before_sha256") == digest(live_source)
            and current.get("after_sha256") == digest(live_source),
            f"{name}: source digest changed")

    delta = receipt.get("source_delta")
    expected_delta = source_delta(baseline, candidate)
    require(isinstance(delta, dict)
            and delta.get("changed_paths") == ([] if live_label == "baseline" else expected_delta)
            and delta.get("within_allowlist") is True,
            f"{name}: source delta custody changed")
    if live_label == "baseline":
        require(delta.get("candidate_available") is False,
                f"{name}: baseline-A1 must precede candidate source capture")
    else:
        require(delta.get("candidate_available") is True,
                f"{name}: candidate source was not available")

    fixture = receipt.get("fixture")
    _, expected_fixture = CAPTURE.fixture_info(job["corpus"])
    require(isinstance(fixture, dict)
            and fixture.get("plan_path") == job["corpus"].get("path")
            and fixture.get("plan_sha256") == job["corpus"].get("sha256")
            and fixture.get("before") == expected_fixture
            and fixture.get("after") == expected_fixture,
            f"{name}: fixture custody changed")
    environment = receipt.get("environment")
    require(environment == {"LC_ALL": "C", "LANG": "C", "TZ": "UTC",
                            "RUSTFLAGS": None, "LD_PRELOAD": None,
                            "MALLOC_CONF": None, "GLIBC_TUNABLES": None,
                            "PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"},
            f"{name}: process environment changed")
    require(isinstance(environment, dict)
            and environment.get("LC_ALL") == "C"
            and environment.get("LANG") == "C"
            and environment.get("TZ") == "UTC"
            and environment.get("PERL_HASH_SEED") == "0"
            and environment.get("PERL_PERTURB_KEYS") == "0",
            f"{name}: deterministic environment binding changed")
    artifacts = receipt.get("artifacts")
    expected_artifacts = {f"{name}.{suffix}" for suffix in (
        "json", "stdout", "stderr", "source-before.json", "source-after.json")}
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts,
            f"{name}: artifact inventory changed")
    for filename, file_digest in artifacts.items():
        path = HERE / filename
        require(path.is_file() and sha(path) == file_digest,
                f"{name}: artifact digest changed: {filename}")
    return receipt


def validate_child(job: dict[str, Any], build: dict[str, Any], plan: dict[str, Any],
                   baseline: dict[str, str], candidate: dict[str, str],
                   constraints_sha: str | None) -> dict[str, Any]:
    receipt = validate_receipt(job, build, plan, baseline, candidate, constraints_sha)
    report = load(HERE / f"{job['name']}.json")
    label = job["name"]
    LEGACY.check_report_metadata(report, build, job["lane"], job["case"],
                                 job["samples"], job["warmup"], label)
    result = report["results"][0]
    require(result.get("case") == job["case"], f"{label}: result case changed")
    elapsed, elapsed_stats = LEGACY.validate_elapsed(result.get("elapsed_ns"),
                                                     job["samples"], label)
    ordinary = LEGACY.validate_ordinary_save(result, job["corpus"], job["phase"],
                                             job["lane"], job["samples"], label)
    LEGACY.validate_operation_metrics(result.get("operation_metrics"),
                                      result["elapsed_ns"], job["samples"],
                                      job["lane"], label)
    allocation = None
    if job["lane"] == "allocator":
        allocation = result["operation_metrics"]["allocation"]
    return {**job, "report": report, "result": result, "ordinary": ordinary,
            "elapsed": elapsed, "elapsed_stats": elapsed_stats,
            "allocation": allocation, "receipt": receipt}


def stable_result(result: dict[str, Any]) -> dict[str, Any]:
    value = copy.deepcopy(result)
    value.pop("elapsed_ns", None)
    value.pop("operation_metrics", None)
    ordinary = value.get("source", {}).get("ordinary_save")
    if isinstance(ordinary, dict):
        ordinary.pop("published_sha256", None)
        ordinary.pop("edit_outcome_sha256", None)
    return value


def stable_evidence(result: dict[str, Any]) -> dict[str, Any]:
    ordinary = result["source"]["ordinary_save"]
    return {"corpus": result["corpus"], "ordinary_corpus": ordinary["corpus"],
            "published_sha256": ordinary["corpus"]["published_sha256"],
            "byte_split": ordinary["corpus"]["byte_split"],
            "sample_byte_split": ordinary.get("sample_byte_split")}


def change_percent(candidate: float, baseline: float) -> float:
    if baseline == 0:
        return 0.0 if candidate == 0 else float("inf")
    return (candidate - baseline) * 100.0 / baseline


def compare(candidate: dict[str, Any], baseline: dict[str, Any], metric: str) -> dict[str, Any]:
    left, right = float(baseline[metric]), float(candidate[metric])
    delta = change_percent(right, left)
    return {"baseline": left, "candidate": right, "delta_percent": delta,
            "improvement_percent": -delta}


def allocation_stats(row: dict[str, Any], field: str) -> dict[str, Any]:
    values = list(row["allocation"][field]["values"])
    return LEGACY.integer_stats(values, f"{row['name']}.{field}")


def repeat_flags(rows: dict[tuple[str, str, str], dict[str, Any]],
                 metric_names: tuple[str, ...], threshold: float) -> list[dict[str, Any]]:
    flags: list[dict[str, Any]] = []
    for lane in LANES:
        for source in ("baseline", "candidate"):
            for corpus in ("generated", "numbered-list"):
                for phase in PHASES:
                    selected = [row for (stage, row_lane, cid, row_phase), row in rows.items()
                                if row_lane == lane and cid == corpus and row_phase == phase
                                and CAPTURE.STAGE_METADATA[stage]["source"] == source]
                    if len(selected) != 2:
                        continue
                    stats = [row["elapsed_stats"] for row in selected]
                    for metric in metric_names:
                        values = [float(item[metric]) for item in stats]
                        spread = (max(values) - min(values)) * 100.0 / min(values)
                        if spread > threshold:
                            flags.append({"lane": lane, "source": source,
                                          "corpus_id": corpus, "phase": phase,
                                          "metric": metric, "spread_percent": spread,
                                          "flag_over_5_percent": True})
    return flags


def analyze() -> dict[str, Any]:
    plan = CAPTURE.load_plan()
    require(plan["freeze_status"] == "frozen" and plan.get("revision"),
            "pilot plan is not frozen")
    CAPTURE.verify_freeze(plan)
    baseline, candidate = CAPTURE.load_source_pair(plan)
    require(candidate is not None, "source-candidate.json is required for final analysis")
    require(CAPTURE.source_census() == candidate,
            "current checkout does not match source-candidate.json")
    constraints_sha = CAPTURE.constraints_digest(plan)
    delta = source_delta(baseline, candidate)
    require(set(delta) <= set(plan["source_delta_allowlist"]),
            f"source delta exceeds allowlist: {delta}")

    rows: dict[tuple[str, str, str, str], dict[str, Any]] = {}
    builds: dict[tuple[str, str], dict[str, Any]] = {}
    for stage in STAGES:
        for lane in LANES:
            build = CAPTURE.build_info(stage, lane, plan)
            builds[(stage, lane)] = build
            for job in CAPTURE.ordered_jobs(plan, stage, lane):
                key = (stage, lane, job["corpus_id"], job["phase"])
                require(key not in rows, f"duplicate child identity: {key}")
                rows[key] = validate_child(job, build, plan, baseline, candidate,
                                           constraints_sha)
    require(len(rows) == 32, f"expected 32 children, found {len(rows)}")

    parity_failures: list[dict[str, str]] = []
    parity_hashes: dict[str, str] = {}
    for corpus in ("generated", "numbered-list"):
        for phase in PHASES:
            reference = stable_result(rows[("baseline-A1", "native", corpus, phase)]["result"])
            key = f"{corpus}/{phase}"
            parity_hashes[key] = digest(reference)
            for stage in STAGES:
                for lane in LANES:
                    value = stable_result(rows[(stage, lane, corpus, phase)]["result"])
                    if value != reference:
                        parity_failures.append({"key": key, "stage": stage, "lane": lane})
    require(not parity_failures, f"deterministic output parity failed: {parity_failures}")

    hard_gates: list[dict[str, Any]] = []
    tail_flags: list[dict[str, Any]] = []
    paired: dict[str, Any] = {}
    for pair, (baseline_stage, candidate_stage) in PAIRS.items():
        pair_data: dict[str, Any] = {"baseline_stage": baseline_stage,
                                     "candidate_stage": candidate_stage, "corpora": {}}
        for corpus in ("generated", "numbered-list"):
            pair_data["corpora"][corpus] = {}
            for phase in PHASES:
                bn = rows[(baseline_stage, "native", corpus, phase)]
                cn = rows[(candidate_stage, "native", corpus, phase)]
                timing = {metric: compare(cn["elapsed_stats"], bn["elapsed_stats"], metric)
                          for metric in ("p50", "mean", "p95", "p99")}
                pair_data["corpora"][corpus][phase] = {"native": timing}
                for metric in plan["thresholds"]["tail_metrics"]:
                    if timing[metric]["delta_percent"] > plan["thresholds"]["repeat_flag_percent"]:
                        tail_flags.append({"pair": pair, "corpus_id": corpus,
                                           "phase": phase, "metric": metric,
                                           "delta_percent": timing[metric]["delta_percent"],
                                           "flag_over_5_percent": True})
                if corpus == "generated" and phase == "counting_publish":
                    for metric in ("p50", "mean"):
                        item = timing[metric]
                        accepted = item["improvement_percent"] >= plan["thresholds"]["generated_counting_improvement_percent"]
                        hard_gates.append({"name": f"{pair}/{corpus}/{phase}/{metric}/improvement",
                                           "pass": accepted, "threshold_percent": 10,
                                           "observed_improvement_percent": item["improvement_percent"]})
                else:
                    for metric in ("p50", "mean"):
                        item = timing[metric]
                        accepted = item["delta_percent"] <= plan["thresholds"]["nonregression_percent"]
                        hard_gates.append({"name": f"{pair}/{corpus}/{phase}/{metric}/nonregression",
                                           "pass": accepted, "threshold_percent": 3,
                                           "observed_delta_percent": item["delta_percent"]})

                ba, ca = rows[(baseline_stage, "allocator", corpus, phase)], rows[(candidate_stage, "allocator", corpus, phase)]
                allocation: dict[str, Any] = {}
                for field in ("allocation_calls", "allocated_bytes"):
                    allocation[field] = {}
                    for metric in ("p50", "mean"):
                        item = compare(allocation_stats(ca, field), allocation_stats(ba, field), metric)
                        allocation[field][metric] = item
                        accepted = item["delta_percent"] <= plan["thresholds"]["allocation_nonregression_percent"]
                        hard_gates.append({"name": f"{pair}/{corpus}/{phase}/allocation/{field}/{metric}/nonregression",
                                           "pass": accepted, "threshold_percent": 3,
                                           "observed_delta_percent": item["delta_percent"]})
                pair_data["corpora"][corpus][phase]["allocator"] = {"allocation": allocation}
        paired[pair] = pair_data

    flat_rows = {(stage, lane, cid, phase): row for (stage, lane, cid, phase), row in rows.items()}
    flags = repeat_flags({(stage, lane, cid, phase): row for (stage, lane, cid, phase), row in flat_rows.items()},
                         ("p50", "mean", "p95", "p99"), plan["thresholds"]["repeat_flag_percent"])
    output_rows = []
    for (stage, lane, corpus, phase), row in sorted(rows.items()):
        allocation = ({field: LEGACY.integer_stats(
            row["allocation"][field]["values"], f"{row['name']}.{field}")
                       for field in ("allocation_calls", "allocated_bytes")}
                       if row["allocation"] is not None else None)
        if allocation is not None:
            raw = row["allocation"]
            before = raw["live_bytes_before"]["values"]
            for field, endpoint in (("net_live_bytes", "live_bytes_after"),
                                    ("peak_above_start_bytes", "region_peak_live_bytes")):
                values = [end - start for start, end in zip(before, raw[endpoint]["values"], strict=True)]
                allocation[field] = LEGACY.integer_stats(values, f"{row['name']}.{field}")
        output_rows.append({"stage": stage, "lane": lane, "corpus_id": corpus,
                            "phase": phase, "repeat": row["repeat"],
                            "elapsed_ns": row["elapsed_stats"],
                            "allocation": allocation,
                            "report_sha256": sha(HERE / f"{row['name']}.json"),
                            "receipt_sha256": sha(HERE / f"{row['name']}.receipt.json")})
    return {
        "schema_version": 1,
        "status": "accepted" if all(item["pass"] for item in hard_gates) else "rejected",
        "packet": plan["packet"],
        "revision": plan["revision"],
        "performance_claim": "paired pilot decision only; no hardware, RSS, cold-cache, throughput, or scaling claim",
        "scope": {"corpora": ["generated", "numbered-list"], "phases": list(PHASES),
                  "stages": list(STAGES), "children": len(rows),
                  "native_samples": plan["native"]["samples"],
                  "allocator_samples": plan["allocator"]["samples"]},
        "source_delta": {"changed_paths": delta, "allowlist": plan["source_delta_allowlist"],
                         "within_allowlist": True},
        "rows": output_rows,
        "normalized_output_parity": {"verified": True, "reference": "baseline-A1/native",
                                      "keys": parity_hashes,
                                      "excluded_fields": [
                                          "result.elapsed_ns", "result.operation_metrics",
                                          "result.source.ordinary_save.published_sha256",
                                          "result.source.ordinary_save.edit_outcome_sha256"]},
        "paired_comparisons": paired,
        "review_flags": {"native_tail_regressions_over_5_percent": tail_flags,
                         "repeat_drift_over_5_percent": flags,
                         "tail_metrics": plan["thresholds"]["tail_metrics"],
                         "note": "Repeat and tail flags are retained separately from hard gates."},
        "decision": {"hard_gates": hard_gates,
                     "all_hard_gates_pass": all(item["pass"] for item in hard_gates),
                     "deterministic_output_parity_pass": True,
                     "accepted": all(item["pass"] for item in hard_gates),
                     "acceptance_rule": "Accept only when output parity and every timing/allocation hard gate pass."},
        "verification": {"child_receipts_verified": len(rows),
                         "strict_0709_validators_used": True,
                         "raw_elapsed_statistics_recomputed": True,
                         "allocation_request_and_bytes_gated": True,
                         "phase_values_additive": False},
        "binary_custody": {f"{stage}/{lane}": {
            "binary": builds[(stage, lane)]["binary"],
            "sha256": builds[(stage, lane)]["binary_sha256"],
            "bytes": builds[(stage, lane)]["binary_bytes"],
            "validation": "exact-live-binary-or-exact-cleanup-witness",
        } for stage in STAGES for lane in LANES},
        "pilot_plan_sha256": sha(HERE / "pilot-plan.json"),
        "capture_script_sha256": sha(HERE / "pilot.py"),
        "constraints_sha256": constraints_sha,
    }


def preflight_baseline_a1() -> int:
    """Validate the eight baseline-A1 children before candidate source exists."""

    plan = CAPTURE.load_plan()
    require(plan["freeze_status"] == "frozen" and plan.get("revision"),
            "pilot plan is not frozen")
    CAPTURE.verify_freeze(plan)
    baseline, candidate = CAPTURE.load_source_pair(plan)
    constraints_sha = CAPTURE.constraints_digest(plan)
    count = 0
    for lane in LANES:
        build = CAPTURE.build_info("baseline-A1", lane, plan)
        for job in CAPTURE.ordered_jobs(plan, "baseline-A1", lane):
            # Passing baseline twice makes the expected source delta empty;
            # validate_receipt still requires candidate_available=false from
            # the capture that genuinely preceded candidate admission.
            validate_child(job, build, plan, baseline, baseline, constraints_sha)
            count += 1
    require(candidate is None or isinstance(candidate, dict),
            "candidate source manifest is malformed")
    return count


def main() -> int:
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", default=str(HERE / "pilot-analysis.json"))
    parser.add_argument("--check", action="store_true",
                        help="validate and replay evidence without writing output")
    parser.add_argument("--preflight-baseline-a1", action="store_true",
                        help="validate only the eight baseline-A1 children")
    args = parser.parse_args()
    try:
        if args.preflight_baseline_a1:
            print(f"validated {preflight_baseline_a1()} baseline-A1 children")
            return 0
        output = analyze()
        path = Path(args.output)
        expected_bytes = (json.dumps(output, indent=2) + "\n").encode()
        if args.check:
            require(path.is_file() and not path.is_symlink(),
                    f"missing retained analysis output: {path}")
            require(path.read_bytes() == expected_bytes,
                    f"analysis replay differs from retained output: {path}")
            print(f"verified {output['scope']['children']} children; exact replay PASS")
        else:
            require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
            path.write_bytes(expected_bytes)
            print(f"verified {output['scope']['children']} children; wrote {path}")
    except (AssertionError, KeyError, OSError, RuntimeError, TypeError, ValueError) as error:
        print(f"analysis failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
