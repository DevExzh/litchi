#!/usr/bin/env python3
"""Verify and compare the 0711 paired DOCX ordinary-save pilot.

The analyzer recomputes elapsed statistics from raw sample vectors, validates
every report with the strict 0709 ordinary-save validators (including decoded
DOCX manifest quantities), checks source/binary/fixture custody for all four
stages, and compares only paired corpus/phase observations.  A performance
decision is emitted only after deterministic output parity and every hard
threshold gate pass.  Allocator live and peak values are retained as
per-operation diagnostics; they are never pooled across phases.
"""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import statistics
import sys
from typing import Any

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

import capture as CAPTURE  # noqa: E402


LEGACY_PATH = REPO / "docs/performance/results/change-0709/analyze.py"
_legacy_spec = importlib.util.spec_from_file_location("ordinary_save_0709", LEGACY_PATH)
if _legacy_spec is None or _legacy_spec.loader is None:
    raise RuntimeError(f"cannot load strict validator: {LEGACY_PATH}")
LEGACY = importlib.util.module_from_spec(_legacy_spec)
_legacy_spec.loader.exec_module(LEGACY)

PHASES = ("edit", "lifecycle")
LANES = ("native", "allocator")
STAGES = ("baseline-A1", "candidate-B1", "candidate-B2", "baseline-A2")
PAIR_BASELINE = {"pair-1": "baseline-A1", "pair-2": "baseline-A2"}
PAIR_CANDIDATE = {"pair-1": "candidate-B1", "pair-2": "candidate-B2"}
ALLOCATION_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def json_digest(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        fail(f"invalid JSON in {path}: {error}")


def write(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    path.write_text(json.dumps(value, indent=2) + "\n")


def check_hex(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64
            and set(value) <= set("0123456789abcdef"),
            f"{label} is not a lowercase SHA-256 digest")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def integer_stats(values: list[int], label: str, *, allow_negative: bool = False) -> dict[str, Any]:
    require(values, f"{label} is empty")
    for index, value in enumerate(values):
        if allow_negative:
            require(isinstance(value, int) and not isinstance(value, bool),
                    f"{label}[{index}] is not an integer")
        else:
            nonnegative_int(value, f"{label}[{index}]")
    ordered = sorted(values)
    return {
        "count": len(values),
        "min": ordered[0],
        "p50": ordered[(len(ordered) - 1) // 2] // 2 + ordered[len(ordered) // 2] // 2
        + ((ordered[(len(ordered) - 1) // 2] % 2
            + ordered[len(ordered) // 2] % 2) // 2),
        "p95": ordered[min(((95 * len(ordered) + 99) // 100) - 1, len(ordered) - 1)],
        "p99": ordered[min(((99 * len(ordered) + 99) // 100) - 1, len(ordered) - 1)],
        "max": ordered[-1],
        "mean": statistics.mean(values),
    }


def cleanup_witnesses() -> list[dict[str, Any]]:
    witnesses: list[dict[str, Any]] = []
    for filename in ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json"):
        path = HERE / filename
        if not path.is_file():
            continue
        value = read(path)

        def walk(item: Any) -> None:
            if isinstance(item, dict):
                raw_path = item.get("path")
                digest = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes"))
                if isinstance(raw_path, str) and isinstance(digest, str):
                    check_hex(digest, f"{filename}:{raw_path}")
                    witnesses.append({"path": raw_path, "sha256": digest, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return witnesses


def build_info(stage: str, lane: str) -> dict[str, Any]:
    metadata = CAPTURE.STAGE_BY_LABEL[stage]
    build_path = HERE / f"build-{metadata['build']}.json"
    records = read(build_path)
    require(isinstance(records, list), f"{build_path.name} is not a list")
    binary_name = f"{metadata['build']}-{'native' if lane == 'native' else 'alloc'}"
    matches = [item for item in records if isinstance(item, dict)
               and Path(str(item.get("binary", ""))).name == binary_name]
    require(len(matches) == 1, f"build record for {binary_name} is not unique")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{binary_name} build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    digest = record.get("binary_sha256")
    size = record.get("binary_bytes")
    check_hex(digest, f"{binary_name}.binary_sha256")
    require(isinstance(size, int) and size > 0, f"{binary_name}.binary_bytes is invalid")
    if binary.is_file() and not binary.is_symlink():
        require(sha(binary) == digest and binary.stat().st_size == size,
                f"{binary_name} live binary identity changed")
    else:
        target = str(binary)
        found = any(str((REPO / Path(item["path"])).resolve()
                      if not Path(item["path"]).is_absolute()
                      else Path(item["path"]).resolve()) == target
                    and item["sha256"] == digest
                    and (item.get("bytes") is None or item["bytes"] == size)
                    for item in cleanup_witnesses())
        require(found, f"{binary_name} missing without an exact cleanup witness")
    source_path = HERE / f"source-{metadata['source']}.json"
    source = read(source_path)
    require(isinstance(source, dict) and source, f"{source_path.name} is invalid")
    require(record.get("source_manifest_sha256") == sha(source_path),
            f"{binary_name} source manifest binding changed")
    return {
        "name": binary_name,
        "binary": str(binary),
        "binary_sha256": digest,
        "binary_bytes": size,
        "build_record": build_path.name,
        "build_record_sha256": sha(build_path),
        "source_manifest": source_path.name,
        "source_manifest_sha256": sha(source_path),
        "source": source,
        # Keep this output independent of whether cleanup has already removed
        # the owned binary.  The branch above still performs the strict live
        # binary or exact post-cleanup-witness check on every analysis run.
        "binary_custody_validation": "exact-live-binary-or-exact-postcleanup-witness",
    }


def expected_command(stage: str, lane: str, corpus: dict[str, Any], phase: str,
                     build: dict[str, Any], plan: dict[str, Any], name: str) -> list[str]:
    command = [
        "taskset", "-c", str(plan["cpu"]), build["binary"],
        "--warmup", str(plan[lane]["warmup"]),
        "--samples", str(plan[lane]["samples"]),
        "--case", CAPTURE.phase_case(corpus, phase),
        "--json", str(HERE / f"{name}.json"),
        "--filesystem-root", str(plan["filesystem_root"]),
    ]
    if corpus["origin"] == "caller-named-real-file":
        command += ["--ooxml-file", str(corpus["path"])]
    return command


def source_artifact(path: Path, expected: dict[str, str], label: str) -> None:
    value = read(path)
    require(value == expected, f"{label}: source census differs from stage manifest")


def fixture_binding(corpus: dict[str, Any]) -> dict[str, Any] | None:
    if corpus["origin"] == "generated-harness-corpus":
        return None
    path = (REPO / corpus["path"]).resolve()
    require(path.is_file() and sha(path) == corpus["sha256"]
            and path.stat().st_size == corpus["bytes"],
            f"fixture changed for {corpus['id']}")
    return {"plan_path": corpus["path"], "resolved_path": str(path),
            "bytes": path.stat().st_size, "sha256": sha(path)}


def ordered_jobs(plan: dict[str, Any], stage: str) -> list[tuple[dict[str, Any], str, int]]:
    metadata = CAPTURE.STAGE_BY_LABEL[stage]
    phases = list(plan["phase_order"])
    corpora = list(plan["corpora"])
    if metadata["order"] == "reverse":
        phases.reverse()
        corpora.reverse()
    return [(corpus, phase, index)
            for index, (phase, corpus) in enumerate((phase, corpus)
                for phase in phases for corpus in corpora)]


def validate_receipt(plan: dict[str, Any], stage: str, lane: str,
                     corpus: dict[str, Any], phase: str, order_index: int,
                     build: dict[str, Any], expected_source: dict[str, str],
                     plan_digest: str, script_digest: str,
                     constraints_digest: str) -> dict[str, Any]:
    name = CAPTURE.child_name(stage, lane, corpus, phase)
    receipt_path = HERE / f"{name}.receipt.json"
    receipt = read(receipt_path)
    metadata = CAPTURE.STAGE_BY_LABEL[stage]
    expected = {
        "schema_version": 1, "packet": plan["packet"], "name": name,
        "stage": stage, "source_label": metadata["source"],
        "build_label": metadata["build"], "pair": metadata["pair"],
        "lane": lane, "stage_order": metadata["order"],
        "stage_order_index": order_index, "corpus_id": corpus["id"],
        "corpus_label": corpus["label"], "corpus_origin": corpus["origin"],
        "phase": phase, "case": CAPTURE.phase_case(corpus, phase),
        "samples": plan[lane]["samples"], "warmup": plan[lane]["warmup"],
        "cpu": plan["cpu"],
    }
    for key, value in expected.items():
        require(receipt.get(key) == value, f"{name} receipt {key} changed")
    require(receipt.get("exit_code") == 0, f"{name} child failed")
    require(receipt.get("binary_path") == build["binary"]
            and receipt.get("binary_sha256") == build["binary_sha256"]
            and receipt.get("binary_bytes") == build["binary_bytes"],
            f"{name} binary custody changed")
    for key, value in (("build_record", build["build_record"]),
                       ("build_record_sha256", build["build_record_sha256"]),
                       ("build_source_manifest", build["source_manifest"]),
                       ("build_source_manifest_sha256", build["source_manifest_sha256"]),
                       ("plan_sha256", plan_digest),
                       ("script_sha256", script_digest),
                       ("constraints_sha256", constraints_digest)):
        require(receipt.get(key) == value, f"{name} receipt {key} changed")
    retained = receipt.get("retained_binary_source")
    require(isinstance(retained, dict)
            and retained.get("manifest") == build["source_manifest"]
            and retained.get("manifest_sha256") == build["source_manifest_sha256"]
            and retained.get("source_census_sha256") == json_digest(expected_source)
            and retained.get("source_entry_count") == len(expected_source),
            f"{name} retained source custody changed")
    current = receipt.get("current_checkout_source")
    require(isinstance(current, dict) and current.get("unchanged_during_child") is True,
            f"{name} source custody failed")
    before = HERE / str(current.get("before_artifact"))
    after = HERE / str(current.get("after_artifact"))
    require(before.name == f"{name}.source-before.json"
            and after.name == f"{name}.source-after.json", f"{name} source names changed")
    source_artifact(before, expected_source, f"{name} before")
    source_artifact(after, expected_source, f"{name} after")
    require(current.get("before_file_sha256") == sha(before)
            and current.get("after_file_sha256") == sha(after)
            and current.get("before_sha256") == json_digest(expected_source)
            and current.get("after_sha256") == json_digest(expected_source),
            f"{name} source digest custody changed")
    for relation_key in ("relation_before", "relation_after"):
        relation = current.get(relation_key)
        require(isinstance(relation, dict)
                and relation.get("mode") == "exact"
                and relation.get("changed_paths") == []
                and relation.get("expected_entry_count") == len(expected_source)
                and relation.get("current_entry_count") == len(expected_source),
                f"{name} {relation_key} changed")
    planned_fixture = fixture_binding(corpus)
    fixture = receipt.get("fixture")
    require(isinstance(fixture, dict), f"{name} fixture receipt missing")
    if planned_fixture is None:
        require(fixture.get("plan_path") is None and fixture.get("before") is None
                and fixture.get("after") is None, f"{name} generated fixture binding exists")
    else:
        require(fixture.get("plan_path") == planned_fixture["plan_path"]
                and fixture.get("plan_sha256") == planned_fixture["sha256"]
                and fixture.get("before") == fixture.get("after")
                and fixture.get("before", {}).get("sha256") == planned_fixture["sha256"]
                and fixture.get("before", {}).get("bytes") == planned_fixture["bytes"],
                f"{name} fixture custody changed")
    require(receipt.get("command") == expected_command(stage, lane, corpus, phase,
                                                        build, plan, name),
            f"{name} command argv changed")
    artifacts = receipt.get("artifacts")
    expected_names = {f"{name}.json", f"{name}.stdout", f"{name}.stderr",
                      f"{name}.source-before.json", f"{name}.source-after.json"}
    require(isinstance(artifacts, dict) and set(artifacts) == expected_names,
            f"{name} artifact inventory changed")
    for filename, digest in artifacts.items():
        path = HERE / filename
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{name} artifact digest changed: {filename}")
    return receipt


def normalized_result(result: dict[str, Any]) -> dict[str, Any]:
    """Remove only timing/instrumentation and sample-cardinality envelopes."""

    # Keep this list deliberately identical to the already reviewed 0709
    # stable_semantics helper.  Scalar publication hashes, decoded manifests,
    # edit outcomes, phase identity, and all deterministic output evidence stay.
    return LEGACY.stable_semantics(result)


def percent_delta(candidate: float, baseline: float) -> float:
    require(baseline > 0, "baseline metric must be positive")
    return (candidate / baseline - 1.0) * 100.0


def metric_comparison(candidate: dict[str, Any], baseline: dict[str, Any],
                      metric: str, statistic: str | None = None) -> dict[str, Any]:
    c = candidate["stats"][metric]
    b = baseline["stats"][metric]
    if isinstance(c, dict):
        require(statistic is not None, f"{metric} requires an allocation statistic")
        require(isinstance(b, dict), f"{metric} baseline statistic is missing")
        c = c[statistic]
        b = b[statistic]
    else:
        require(statistic is None, f"{metric} is not a nested allocation statistic")
    return {"baseline": b, "candidate": c,
            "delta_percent": percent_delta(float(c), float(b)),
            "improvement_percent": (1.0 - float(c) / float(b)) * 100.0,
            **({"statistic": statistic} if statistic is not None else {})}


def elapsed_summary(vector: list[int], label: str) -> dict[str, Any]:
    stats = LEGACY.elapsed_stats(vector)
    return stats


def allocation_summary(result: dict[str, Any], label: str) -> dict[str, Any]:
    envelope = result["operation_metrics"]["allocation"]
    values: dict[str, list[int]] = {}
    for field in ALLOCATION_FIELDS:
        vector = envelope[field]["values"]
        require(isinstance(vector, list), f"{label} allocation {field} missing")
        values[field] = list(vector)
    values["net_live"] = [after - before for before, after in zip(
        values["live_bytes_before"], values["live_bytes_after"])]
    values["peak_above_start"] = [peak - before for peak, before in zip(
        values["region_peak_live_bytes"], values["live_bytes_before"])]
    return {
        "sample_count": len(values["allocation_calls"]),
        "values": values,
        "stats": {field: integer_stats(
            items, f"{label}/{field}", allow_negative=field in {"net_live", "peak_above_start"})
            for field, items in values.items()},
        "diagnostic_note": "net_live and peak_above_start are per-operation observations; no phase or sample sums are valid",
    }


def repeat_drift(jobs: dict[tuple[str, str, str, str], dict[str, Any]],
                 plan: dict[str, Any]) -> list[dict[str, Any]]:
    flags: list[dict[str, Any]] = []
    for lane in LANES:
        for corpus in plan["corpora"]:
            for phase in PHASES:
                for source_label, stages in (("baseline", ("baseline-A1", "baseline-A2")),
                                             ("candidate", ("candidate-B1", "candidate-B2"))):
                    values = []
                    for stage in stages:
                        job = jobs[(stage, lane, corpus["id"], phase)]
                        values.append(job["elapsed_stats"])
                    for metric in ("p50", "mean", "p95", "p99"):
                        left, right = float(values[0][metric]), float(values[1][metric])
                        spread = abs(right - left) * 100.0 / min(left, right)
                        if spread > plan["thresholds"]["repeat_drift_flag_percent"]:
                            flags.append({"lane": lane, "corpus_id": corpus["id"],
                                          "phase": phase, "source": source_label,
                                          "metric": metric, "spread_percent": spread,
                                          "flag_over_5_percent": True})
    return flags


def main() -> int:
    parser = __import__("argparse").ArgumentParser(description=__doc__)
    parser.add_argument("--output", default=str(HERE / "analysis.json"))
    args = parser.parse_args()
    try:
        plan = CAPTURE.load_plan()
        require(tuple(item["label"] for item in plan["stages"]) == STAGES,
                "stage labels changed")
        expected_sources = {
            label: read(HERE / f"source-{CAPTURE.STAGE_BY_LABEL[label]['source']}.json")
            for label in STAGES
        }
        for label, source in expected_sources.items():
            require(isinstance(source, dict) and source, f"{label} source manifest invalid")
        current_source = CAPTURE.source_census()
        require(current_source in expected_sources.values(),
                "current checkout does not match either retained stage source manifest")
        constraints_path = HERE / "constraints.json"
        constraints = read(constraints_path)
        require(isinstance(constraints, dict), "constraints invalid")
        for name, digest in constraints.items():
            path = REPO / name
            require(path.is_file() and sha(path) == digest, f"constraint changed: {name}")
        builds = {(stage, lane): build_info(stage, lane)
                  for stage in STAGES for lane in LANES}
        plan_digest = sha(HERE / "plan.json")
        script_digest = sha(HERE / "capture.py")
        constraints_digest = sha(constraints_path)
        jobs: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        normalized_by_key: dict[tuple[str, str, str, str], dict[str, Any]] = {}
        for stage in STAGES:
            for corpus, phase, order_index in ordered_jobs(plan, stage):
                for lane in LANES:
                    build = builds[(stage, lane)]
                    receipt = validate_receipt(
                        plan, stage, lane, corpus, phase, order_index, build,
                        expected_sources[stage], plan_digest, script_digest,
                        constraints_digest,
                    )
                    name = CAPTURE.child_name(stage, lane, corpus, phase)
                    report = read(HERE / f"{name}.json")
                    LEGACY.check_report_metadata(
                        report, build, lane, CAPTURE.phase_case(corpus, phase),
                        plan[lane]["samples"], plan[lane]["warmup"], name,
                    )
                    result = report["results"][0]
                    require(result.get("case") == CAPTURE.phase_case(corpus, phase),
                            f"{name} case changed")
                    elapsed, _ = LEGACY.validate_elapsed(
                        result.get("elapsed_ns"), plan[lane]["samples"], name)
                    ordinary = LEGACY.validate_ordinary_save(
                        result, corpus, phase, lane, plan[lane]["samples"], name)
                    LEGACY.validate_operation_metrics(
                        result.get("operation_metrics"), result["elapsed_ns"],
                        plan[lane]["samples"], lane, name)
                    jobs[(stage, lane, corpus["id"], phase)] = {
                        "stage": stage, "lane": lane, "corpus": corpus, "phase": phase,
                        "receipt": receipt, "report": report, "result": result,
                        "elapsed": elapsed, "elapsed_stats": elapsed_summary(elapsed, name),
                        "ordinary": ordinary,
                        "allocation": (allocation_summary(result, name)
                                        if lane == "allocator" else None),
                    }
                    normalized_by_key[(stage, lane, corpus["id"], phase)] = normalized_result(result)
        require(len(jobs) == 32, f"expected 32 isolated children, found {len(jobs)}")

        # Every stage/lane must produce exactly the same deterministic result
        # for a corpus/phase.  The only dropped fields are timed envelopes,
        # allocator/operation instrumentation, and repeated sample arrays whose
        # cardinality differs between native and allocator lanes.
        parity_digests: dict[str, str] = {}
        parity_failures: list[dict[str, Any]] = []
        stable_output_witnesses: dict[str, list[dict[str, Any]]] = {}
        for corpus in plan["corpora"]:
            for phase in PHASES:
                reference = normalized_by_key[("baseline-A1", "native", corpus["id"], phase)]
                key = f"{corpus['id']}/{phase}"
                parity_digests[key] = json_digest(reference)
                stable_output_witnesses[key] = []
                for stage in STAGES:
                    for lane in LANES:
                        job = jobs[(stage, lane, corpus["id"], phase)]
                        value = normalized_result(job["result"])
                        if value != reference:
                            parity_failures.append({"key": key, "stage": stage, "lane": lane})
                        stable_output_witnesses[key].append({
                            "stage": stage,
                            "lane": lane,
                            "report": f"{CAPTURE.child_name(stage, lane, corpus, phase)}.json",
                            "report_sha256": sha(HERE / f"{CAPTURE.child_name(stage, lane, corpus, phase)}.json"),
                            "normalized_sha256": json_digest(value),
                            "published_sha256": job["ordinary"]["corpus"]["published_sha256"],
                        })
        require(not parity_failures, f"deterministic output parity failed: {parity_failures}")

        paired: dict[str, Any] = {}
        hard_gates: list[dict[str, Any]] = []
        tail_flags: list[dict[str, Any]] = []
        for pair, baseline_stage in PAIR_BASELINE.items():
            candidate_stage = PAIR_CANDIDATE[pair]
            pair_data: dict[str, Any] = {"baseline_stage": baseline_stage,
                                         "candidate_stage": candidate_stage,
                                         "corpora": {}}
            for corpus in plan["corpora"]:
                cid = corpus["id"]
                pair_data["corpora"][cid] = {}
                for phase in PHASES:
                    bn = jobs[(baseline_stage, "native", cid, phase)]
                    cn = jobs[(candidate_stage, "native", cid, phase)]
                    elapsed = {metric: metric_comparison(
                        {"stats": cn["elapsed_stats"]}, {"stats": bn["elapsed_stats"]}, metric)
                               for metric in ("p50", "mean", "p95", "p99")}
                    pair_data["corpora"][cid][phase] = {"native": elapsed}
                    if phase == "edit":
                        for metric in ("p50", "mean"):
                            gate = elapsed[metric]["improvement_percent"] >= plan["thresholds"]["edit_improvement_percent"]
                            hard_gates.append({"name": f"{pair}/{cid}/edit/{metric}/improvement",
                                               "pass": gate, "threshold_percent": 3,
                                               "observed_improvement_percent": elapsed[metric]["improvement_percent"]})
                    else:
                        for metric in ("p50", "mean"):
                            gate = elapsed[metric]["delta_percent"] <= plan["thresholds"]["lifecycle_regression_percent"]
                            hard_gates.append({"name": f"{pair}/{cid}/lifecycle/{metric}/nonregression",
                                               "pass": gate, "threshold_percent": 3,
                                               "observed_delta_percent": elapsed[metric]["delta_percent"]})
                    for metric in ("p95", "p99"):
                        if elapsed[metric]["delta_percent"] > plan["thresholds"]["tail_flag_percent"]:
                            tail_flags.append({"pair": pair, "corpus_id": cid, "phase": phase,
                                               "metric": metric, "delta_percent": elapsed[metric]["delta_percent"]})
                    ba = jobs[(baseline_stage, "allocator", cid, phase)]["allocation"]
                    ca = jobs[(candidate_stage, "allocator", cid, phase)]["allocation"]
                    allocation = {}
                    for metric in ("allocation_calls", "allocated_bytes"):
                        allocation[metric] = {}
                        for statistic in ("p50", "mean"):
                            comparison = metric_comparison(ca, ba, metric, statistic)
                            allocation[metric][statistic] = comparison
                            gate = comparison["delta_percent"] <= plan["thresholds"]["allocation_regression_percent"]
                            hard_gates.append({"name": f"{pair}/{cid}/{phase}/allocation/{metric}/{statistic}/nonregression",
                                               "pass": gate, "threshold_percent": 3,
                                               "observed_delta_percent": comparison["delta_percent"]})
                    pair_data["corpora"][cid][phase]["allocator"] = {
                        "allocation": allocation,
                        "baseline_diagnostics": ba,
                        "candidate_diagnostics": ca,
                    }
            paired[pair] = pair_data

        drift_flags = repeat_drift(jobs, plan)
        output = {
            "schema_version": 1,
            "packet": plan["packet"],
            "revision": plan["revision"],
            "performance_claim": "paired pilot decision only; no hardware, RSS, cold-cache, throughput, or scaling claim",
            "scope": {
                "corpora": [corpus["id"] for corpus in plan["corpora"]],
                "phases": list(PHASES), "stages": list(STAGES),
                "children": len(jobs), "native_samples": plan["native"]["samples"],
                "allocator_samples": plan["allocator"]["samples"],
            },
            "raw_sample_statistics": {
                "/".join((stage, lane, cid, phase)): {
                    "elapsed_ns": item["elapsed_stats"],
                    "allocation": item["allocation"],
                }
                for (stage, lane, cid, phase), item in sorted(jobs.items())
            },
            "normalized_output_parity": {
                "verified": True,
                "reference": "baseline-A1/native",
                "keys": parity_digests,
                "stable_output_witnesses": stable_output_witnesses,
                "excluded_fields": [
                    "result.elapsed_ns",
                    "result.operation_metrics",
                    "result.source.ordinary_save.published_sha256",
                    "result.source.ordinary_save.edit_outcome_sha256",
                ],
                "note": "Scalar publication hashes, decoded package manifests, phase identity, edit outcomes, and all non-repeated deterministic evidence remain compared.",
            },
            "paired_comparisons": paired,
            "review_flags": {
                "native_tail_regressions_over_5_percent": tail_flags,
                "repeat_drift_over_5_percent": drift_flags,
                "note": "Flags require review; they are reported separately from the hard threshold gates."
            },
            "decision": {
                "hard_gates": hard_gates,
                "all_hard_gates_pass": all(item["pass"] for item in hard_gates),
                "deterministic_output_parity_pass": True,
                "accepted": all(item["pass"] for item in hard_gates),
                "acceptance_rule": "accept only when deterministic output parity and every native edit, lifecycle, and allocator hard gate pass",
            },
            "verification": {
                "current_source_matches_retained_stage": True,
                "child_receipts_verified": len(jobs),
                "strict_0709_validators_used": True,
                "strict_decoded_real_file_manifest_verified": True,
                "raw_elapsed_statistics_recomputed": True,
                "allocation_net_live_and_peak_above_start_are_not_summed": True,
                "postcleanup_witnesses": {
                    "required": True,
                    "validation": "Every retained build binary is checked as the exact live file or by an exact post-cleanup witness with matching path, SHA-256, and byte count.",
                    "output_mode": "normalized-invariant",
                },
                "postcleanup_witness_requirement": "Final seal must retain exact binary/output witnesses after owned scratch cleanup.",
            },
            "binary_custody": {
                f"{stage}/{lane}": {
                    "binary": builds[(stage, lane)]["binary"],
                    "sha256": builds[(stage, lane)]["binary_sha256"],
                    "bytes": builds[(stage, lane)]["binary_bytes"],
                    "validation": builds[(stage, lane)]["binary_custody_validation"],
                }
                for stage in STAGES for lane in LANES
            },
            "source_manifests": {
                label: {"path": f"source-{CAPTURE.STAGE_BY_LABEL[label]['source']}.json",
                        "sha256": sha(HERE / f"source-{CAPTURE.STAGE_BY_LABEL[label]['source']}.json"),
                        "entry_count": len(expected_sources[label])}
                for label in STAGES
            },
            "plan_sha256": plan_digest,
            "capture_script_sha256": script_digest,
            "constraints_sha256": constraints_digest,
            "strict_validator_script_sha256": sha(LEGACY_PATH),
        }
        write(Path(args.output), output)
        print(f"verified {len(jobs)} children; wrote {args.output}")
    except (AssertionError, OSError, RuntimeError, ValueError, KeyError) as error:
        print(f"analysis failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
