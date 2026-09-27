"""Offline replay and analysis for the ordinary-save durability capture.

The capture process is deliberately outside this module.  This program only
reads the frozen plan, receipts, harness JSON and retained process gauges.  A
missing or changed receipt is an error; the analyzer never fills in a missing
child from the plan and never runs Cargo, a benchmark, or a profiler.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import math
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
REVIEW_PERCENT = 5.0
TIMED_PHASES = ("lifecycle", "atomic_publish")
POLICIES = ("default", "full", "file-only", "no-sync")
CONTROL_PHASES = ("edit", "counting_publish")
METRICS = ("p50", "mean", "p95", "p99")
ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
HEX = frozenset("0123456789abcdefABCDEF")


class ReplayError(RuntimeError):
    pass


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def json_sha256(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def capture_json_sha256(value: Any) -> str:
    """Digest used by capture.py for report/qualification identities."""

    encoded = json.dumps(value, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def relative_packet(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def _relocated_candidates(raw: Path) -> list[Path]:
    """Resolve paths retained from a removed checkout without basename lookup."""

    candidates: list[Path] = []
    parts = raw.parts
    markers = ("change-0778", "native-0", "allocation-0", "qualification-0")
    for marker in markers:
        if marker not in parts:
            continue
        index = parts.index(marker)
        suffix = parts[index + 1 :]
        if marker == "change-0778":
            candidates.append(PACKET.joinpath(*suffix))
        else:
            candidates.append(PACKET / marker / Path(*suffix))
    return candidates


def resolve_path(value: Any, *, capture_bound: bool = False) -> Path:
    require(isinstance(value, str) and value, f"invalid artifact path: {value!r}")
    raw = Path(value)
    candidates: list[Path] = []
    if raw.is_absolute():
        candidates.extend(_relocated_candidates(raw))
        candidates.append(raw)
    else:
        text = value.replace("\\", "/")
        if text.startswith("change-0778/"):
            candidates.append(PACKET / text.split("/", 1)[1])
        if capture_bound:
            for directory in ("native-0", "allocation-0", "qualification-0", "capture-0"):
                if text == directory or text.startswith(directory + "/"):
                    candidates.append(PACKET / text)
        candidates.extend((PACKET / raw, PACKET / "capture-0" / raw))
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            return candidate.resolve()
    return candidates[0].resolve(strict=False) if candidates else raw.resolve(strict=False)


def artifact_receipt(
    value: Any,
    label: str,
    *,
    capture_bound: bool = True,
    allow_missing: bool = False,
) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact receipt")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label} has no path")
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = resolve_path(raw, capture_bound=capture_bound)
    if capture_bound:
        try:
            path.relative_to(PACKET.resolve())
        except ValueError:
            fail(f"{label} escaped packet: {raw}")
    if not path.is_file():
        if allow_missing and not capture_bound:
            return None
        fail(f"missing {label} after path relocation: {raw}")
    require(path.stat().st_size == size, f"{label}.bytes changed")
    require(sha256(path) == digest, f"{label}.sha256 changed")
    return path


def path_digest(value: Any, label: str, *, capture_bound: bool = True) -> Path:
    path = artifact_receipt(value, label, capture_bound=capture_bound)
    assert path is not None
    return path


def find_packet_file(names: Iterable[str]) -> Path | None:
    for name in names:
        path = PACKET / name
        if path.is_file() and not path.is_symlink():
            return path
    return None


def find_lane(name: str) -> Path:
    for candidate in (PACKET / f"{name}-0", PACKET / name, PACKET / "capture-0"):
        if candidate.is_dir():
            return candidate
    fail(f"missing {name} capture directory")


def load_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "litchi-0778-durability-plan-v1", "plan schema changed")
    base = plan.get("base")
    require(isinstance(base, str) and 7 <= len(base) <= 64
            and all(c in HEX for c in base), "plan base revision is invalid")
    require(plan.get("phases") == list(TIMED_PHASES), "timed phase order changed")
    require(plan.get("policies") == list(POLICIES), "policy order changed")
    require(plan.get("controls") == list(CONTROL_PHASES), "control phase set changed")
    require(plan.get("policy_orders") == [
        ["default", "full", "no-sync", "file-only"],
        ["full", "file-only", "default", "no-sync"],
        ["file-only", "no-sync", "full", "default"],
        ["no-sync", "default", "file-only", "full"],
    ], "block policy order changed")
    require(plan.get("review_percent") == 5, "review threshold changed")
    for lane, expected in (
        ("native", {"blocks": 4, "samples": 100, "warmup": 10}),
        ("allocation", {"blocks": 2, "samples": 3, "warmup": 0}),
        ("qualification", {"samples": 1, "warmup": 0}),
    ):
        value = plan.get(lane)
        require(isinstance(value, dict), f"plan {lane} is missing")
        for key, expected_value in expected.items():
            require(value.get(key) == expected_value, f"plan {lane}.{key} changed")
    corpora = plan.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 7, "plan corpus count changed")
    seen: set[str] = set()
    for index, corpus in enumerate(corpora):
        require(isinstance(corpus, dict), f"corpus {index} is not an object")
        cid = corpus.get("id")
        require(isinstance(cid, str) and cid and cid not in seen, f"corpus {index} id is invalid")
        seen.add(cid)
        require(corpus.get("format") in {"docx", "xlsx", "pptx"}, f"{cid}: format changed")
        require(isinstance(corpus.get("expected_edit_admitted"), bool), f"{cid}: admission missing")
        raw = corpus.get("path")
        if raw is None:
            require(corpus.get("sha256") is None and corpus.get("bytes") is None,
                    f"{cid}: generated corpus has fixture identity")
        else:
            require(isinstance(raw, str) and raw, f"{cid}: fixture path is missing")
            require(is_sha(corpus.get("sha256")), f"{cid}: fixture SHA is invalid")
            positive_int(corpus.get("bytes"), f"{cid}: fixture bytes")
            fixture = (PACKET.parents[3] / raw).resolve() if not Path(raw).is_absolute() else Path(raw)
            require(fixture.is_file() and not fixture.is_symlink(), f"{cid}: fixture is missing")
            require(fixture.stat().st_size == corpus["bytes"], f"{cid}: fixture bytes changed")
            require(sha256(fixture) == corpus["sha256"], f"{cid}: fixture SHA changed")
    expected = plan.get("expected_children")
    require(isinstance(expected, dict), "plan expected_children is missing")
    require(expected == {"native": 280, "allocation": 140, "qualification": 28},
            "plan expected child cardinalities changed")
    return plan


def plan_corpora(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {item["id"]: item for item in plan["corpora"]}


def field(row: dict[str, Any], *names: str) -> Any:
    for name in names:
        if name in row:
            return row[name]
    nested = row.get("job")
    if isinstance(nested, dict):
        for name in names:
            if name in nested:
                return nested[name]
    return None


def row_identity(
    row: dict[str, Any], plan: dict[str, Any], label: str, *, allow_none_block: bool = False
) -> dict[str, Any]:
    corpora = plan_corpora(plan)
    corpus_id = field(row, "corpus_id", "corpus", "corpus_name")
    phase = field(row, "phase")
    policy = field(row, "policy", "durability", "save_durability")
    block = field(row, "block", "block_index", "block_id")
    control = field(row, "control", "control_id")
    require(isinstance(corpus_id, str) and corpus_id in corpora, f"{label}: corpus identity missing")
    require(isinstance(phase, str), f"{label}: phase identity missing")
    require(isinstance(policy, str), f"{label}: policy identity missing")
    require(policy in POLICIES, f"{label}: unknown policy {policy!r}")
    if allow_none_block and block is None:
        pass
    else:
        require(isinstance(block, int) and not isinstance(block, bool), f"{label}: block identity missing")
        require(block >= 0, f"{label}: block is negative")
    if phase in TIMED_PHASES:
        require(control is None, f"{label}: timed row is marked control")
    else:
        require(phase in CONTROL_PHASES, f"{label}: unknown control phase {phase!r}")
        require(policy == "default", f"{label}: controls must use default policy")
        require(control in (None, phase), f"{label}: control identity changed")
        control = phase
    return {"corpus_id": corpus_id, "phase": phase, "policy": policy,
            "block": block, "control": control}


def scheduled_jobs(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    """Reconstruct capture.py's frozen child schedule for strict replay."""

    corpora = [item["id"] for item in plan["corpora"]]
    if lane == "qualification":
        def qualification_case(corpus_id: str) -> str:
            corpus = plan_corpora(plan)[corpus_id]
            prefix = f"{corpus['format']}_ordinary_save_"
            if corpus.get("path") is not None:
                prefix = f"{corpus['format']}_real_file_ordinary_save_"
            return prefix + "atomic_publish"

        return [
            {"lane": lane, "corpus_id": corpus_id, "phase": "atomic_publish",
             "policy": policy, "block": None, "repeat": None,
             "case": qualification_case(corpus_id),
             "samples": plan["qualification"]["samples"],
             "warmup": plan["qualification"]["warmup"], "order_index": index}
            for index, (corpus_id, policy) in enumerate(
                ( (corpus_id, policy) for corpus_id in corpora for policy in POLICIES )
            )
        ]
    require(lane in {"native", "allocation"}, f"unknown capture lane {lane}")
    jobs: list[dict[str, Any]] = []
    for block in range(plan[lane]["blocks"]):
        require(isinstance(plan.get("policy_orders"), list)
                and block < len(plan["policy_orders"]),
                f"plan policy order is missing block {block}")
        order = plan["policy_orders"][block]
        require(isinstance(order, list) and set(order) == set(POLICIES)
                and len(order) == len(POLICIES),
                f"plan policy order is invalid for block {block}")
        for corpus_id in corpora:
            corpus = plan_corpora(plan)[corpus_id]
            prefix = f"{corpus['format']}_ordinary_save_"
            if corpus.get("path") is not None:
                prefix = f"{corpus['format']}_real_file_ordinary_save_"
            for phase in TIMED_PHASES:
                for policy in order:
                    jobs.append({"lane": lane, "corpus_id": corpus_id, "phase": phase,
                                 "policy": policy, "block": block,
                                 "repeat": None if lane == "native" else block,
                                 "case": prefix + phase, "samples": plan[lane]["samples"],
                                 "warmup": plan[lane]["warmup"], "order_index": len(jobs)})
            for phase in CONTROL_PHASES:
                jobs.append({"lane": lane, "corpus_id": corpus_id, "phase": phase,
                             "policy": "default", "block": block,
                             "repeat": None if lane == "native" else block,
                             "case": prefix + phase, "samples": plan[lane]["samples"],
                             "warmup": plan[lane]["warmup"], "order_index": len(jobs)})
    return jobs


def check_scheduled_rows(rows: list[dict[str, Any]], plan: dict[str, Any], lane: str) -> None:
    expected = scheduled_jobs(plan, lane)
    require(len(rows) == len(expected), f"{lane} schedule cardinality changed")
    for index, (row, job) in enumerate(zip(rows, expected)):
        label = f"{lane} run {index}"
        actual = row_identity(row, plan, label, allow_none_block=(lane == "qualification"))
        expected_identity = {"corpus_id": job["corpus_id"], "phase": job["phase"],
                            "policy": job["policy"], "block": job["block"],
                            "control": job["phase"] if job["phase"] in CONTROL_PHASES else None}
        require(actual == expected_identity, f"{label} identity differs from frozen schedule")
        captured_job = row.get("job")
        require(isinstance(captured_job, dict), f"{label} job receipt is missing")
        for key, value in job.items():
            require(captured_job.get(key) == value,
                    f"{label} job.{key} differs from frozen schedule")


def load_runs(directory: Path, label: str) -> list[dict[str, Any]]:
    # capture.py writes a lane-level manifest beside the numbered lane
    # directory.  Accept the older runs.json spelling only for replaying a
    # retained attempt; both forms carry the same row objects.
    path = directory.parent / f"{label}.json"
    if not path.is_file():
        path = directory / "runs.json"
    value = read_json(path)
    if isinstance(value, dict) and isinstance(value.get("rows"), list):
        require(value.get("schema") in {
                    "litchi-0778-durability-runs-v1",
                    "litchi-0778-durability-qualification-runs-v1",
                }, f"{label} runs schema changed")
        value = value["rows"]
    require(isinstance(value, list), f"{label} runs manifest is not an array")
    rows: list[dict[str, Any]] = []
    for index, row in enumerate(value):
        require(isinstance(row, dict), f"{label} run {index} is not an object")
        require(field(row, "exit", "exit_code") == 0, f"{label} run {index} failed")
        rows.append(row)
    return rows


def load_complete(directory: Path, label: str) -> dict[str, Any]:
    complete_path = directory / "complete.json"
    require(complete_path.is_file(), f"{label} completion marker is missing")
    complete = read_json(complete_path)
    require(isinstance(complete, dict), f"{label} completion marker is invalid")
    complete_marker = complete.get("complete") is True
    lane_marker = isinstance(complete.get("lane"), str) and complete.get("serial") is True
    require(complete_marker or lane_marker, f"{label} capture is incomplete")
    if complete.get("rows") is not None and complete.get("expected_rows") is not None:
        require(complete.get("rows") == complete.get("expected_rows"),
                f"{label} completion row count is incomplete")
    require(complete.get("source_unchanged", True) is True, f"{label} source changed")
    require(complete.get("fixtures_unchanged", True) is True, f"{label} fixtures changed")
    for key in ("binaries_unchanged", "binary_unchanged", "runner_unchanged"):
        if key in complete:
            require(complete[key] is True, f"{label} {key} guard failed")
    return complete


def harness_quantile(values: list[int], percentile: int) -> int:
    """Match perf-baseline's exact nearest-rank percentile definition."""

    ordered = sorted(values)
    index = ((percentile * len(ordered) + 99) // 100) - 1
    return ordered[min(index, len(ordered) - 1)]


def midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + ((left % 2 + right % 2) // 2)


def elapsed_stats(values: Iterable[int]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "timing vector is empty")
    for index, value in enumerate(vector):
        positive_int(value, f"elapsed_ns.samples[{index}]")
    ordered = sorted(vector)
    return {
        "count": len(vector),
        "min": ordered[0],
        "p50": midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
        "p95": harness_quantile(vector, 95),
        "p99": harness_quantile(vector, 99),
        "max": ordered[-1],
        "mean": statistics.mean(vector),
    }


def check_elapsed(value: Any, samples: int, label: str) -> tuple[list[int], dict[str, Any]]:
    require(isinstance(value, dict), f"{label}.elapsed_ns is missing")
    require(value.get("unit") == "ns", f"{label}.elapsed_ns unit changed")
    vector = value.get("samples")
    require(isinstance(vector, list) and len(vector) == samples,
            f"{label}.elapsed_ns.samples length changed")
    require(vector == sorted(vector), f"{label}.elapsed_ns.samples are not sorted")
    order = value.get("sample_order")
    require(isinstance(order, list) and sorted(order) == list(range(samples)),
            f"{label}.elapsed_ns.sample_order is not a permutation")
    expected = elapsed_stats(vector)
    for key in ("count", "min", "p50", "p95", "p99", "max"):
        if key in value:
            require(value[key] == expected[key], f"{label}.elapsed_ns.{key} is stale")
    if "mean" in value:
        finite_number(value["mean"], f"{label}.elapsed_ns.mean")
        require(math.isclose(float(value["mean"]), float(expected["mean"]), rel_tol=1e-12, abs_tol=1e-6),
                f"{label}.elapsed_ns.mean is stale")
    return vector, expected


def _result_from_report(report: dict[str, Any], label: str) -> dict[str, Any]:
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1, f"{label} must contain one result")
    result = results[0]
    require(isinstance(result, dict), f"{label} result is not an object")
    return result


def check_report_metadata(report: dict[str, Any], row: dict[str, Any], plan: dict[str, Any],
                          samples: int, warmup: int, label: str) -> dict[str, Any]:
    require(report.get("schema_version") == 1, f"{label} report schema changed")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{label} binary identity missing")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label} configuration missing")
    require(configuration.get("samples_per_case") == samples,
            f"{label} sample configuration changed")
    require(configuration.get("warmup_iterations_per_case") == warmup,
            f"{label} warmup configuration changed")
    require(configuration.get("filesystem_root_selected") is True,
            f"{label} filesystem root was not selected")
    case = field(row, "case")
    require(isinstance(case, str) and case, f"{label} case receipt is missing")
    require(configuration.get("cases") == [case], f"{label} mixes cases")
    return _result_from_report(report, label)


def check_report_binary_identity(report: dict[str, Any], row: dict[str, Any],
                                 label: str, *, allocation: bool = False) -> None:
    identity = report.get("binary_identity")
    binary = field(row, "binary")
    require(isinstance(identity, dict) and isinstance(binary, dict),
            f"{label} binary identity receipt is missing")
    require(identity.get("binary_sha256") == binary.get("sha256") and
            identity.get("binary_bytes", identity.get("bytes")) == binary.get("bytes"),
            f"{label} report binary identity differs from retained binary")
    if identity.get("path") is not None and binary.get("path") is not None:
        require(identity.get("path") == binary.get("path"),
                f"{label} report binary path differs from retained binary")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{label} tool identity is missing")
    expected_tool_binary = "litchi-perf-baseline-alloc" if allocation else "litchi-perf-baseline"
    require(tool.get("binary") == expected_tool_binary,
            f"{label} tool binary identity changed")
    if allocation:
        require(tool.get("instrumentation") == "system_allocator_operation_scoped" and
                tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{label} allocator instrumentation identity changed")
    else:
        require(tool.get("instrumentation") == "none" and
                tool.get("allocator_counter_revision") in (None, "none"),
                f"{label} non-allocator instrumentation identity changed")


def check_operation_alignment(result: dict[str, Any], samples: int, label: str,
                              *, allocation: bool = False) -> None:
    metrics = result.get("operation_metrics")
    require(isinstance(metrics, dict), f"{label}.operation_metrics missing")
    require(metrics.get("sample_count") == samples, f"{label} operation sample count changed")
    require(metrics.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{label} operation-metric alignment changed")
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict) and metrics.get("sample_indices") == elapsed.get("sample_order"),
            f"{label} operation vectors are not aligned to elapsed sample order")
    expected_claim = ("allocator_instrumented_elapsed_not_latency_claim"
                      if allocation else "comparable_timed_operation")
    require(metrics.get("latency_claim") == expected_claim,
            f"{label} operation latency claim changed")
    envelope = metrics.get("allocation")
    require(isinstance(envelope, dict), f"{label} allocation envelope is missing")
    require(envelope.get("scope") == "operation_global_system_allocator",
            f"{label} allocation scope changed")
    if allocation:
        require(envelope.get("status") == "measured",
                f"{label} allocator envelope is not measured")
    else:
        require(envelope.get("status") == "unavailable",
                f"{label} non-allocator envelope unexpectedly measured")
        for name in ALLOCATION_FIELDS:
            vector = envelope.get(name)
            require(isinstance(vector, dict) and "values" not in vector,
                    f"{label} non-allocator {name} unexpectedly has values")


def corpus_evidence(result: dict[str, Any], plan_corpus: dict[str, Any],
                    identity: dict[str, Any], label: str) -> dict[str, Any]:
    source = result.get("source")
    require(isinstance(source, dict), f"{label}.source is missing")
    ordinary = source.get("ordinary_save")
    require(isinstance(ordinary, dict), f"{label}.source.ordinary_save is missing")
    expected_format = plan_corpus["format"].upper()
    require(ordinary.get("format") == expected_format, f"{label} format changed")
    expected_origin = "caller-named-real-file" if plan_corpus.get("path") else "generated-harness-corpus"
    require(ordinary.get("origin") == expected_origin, f"{label} origin changed")
    phase = identity["phase"]
    phase_labels = {"lifecycle": "open+edit+save", "atomic_publish": "save-to-path",
                    "edit": "edit", "counting_publish": "serialize-to-counting-sink"}
    require(ordinary.get("phase") == phase_labels[phase], f"{label} phase label changed")
    require(isinstance(ordinary.get("timing_scope"), str) and ordinary["timing_scope"],
            f"{label} timing scope missing")
    policy = identity["policy"]
    durability = ordinary.get("save_durability")
    if phase in CONTROL_PHASES:
        require(policy == "default" and durability is None, f"{label} control durability changed")
    elif policy == "default":
        require(durability is None, f"{label} default unexpectedly names durability")
    else:
        require(durability == policy, f"{label} durability policy changed")
    evidence = ordinary.get("corpus")
    require(isinstance(evidence, dict), f"{label} corpus evidence missing")
    require(evidence.get("edit_admitted") == plan_corpus["expected_edit_admitted"],
            f"{label} edit admission differs from plan")
    require(isinstance(evidence.get("edit_description"), str) and evidence["edit_description"],
            f"{label} edit description is missing")
    outcome = evidence.get("edit_outcome")
    require(isinstance(outcome, str) and outcome, f"{label} edit outcome missing")
    if plan_corpus["expected_edit_admitted"]:
        require(outcome == "admitted", f"{label} admitted outcome changed")
    else:
        require(outcome.startswith("refused:"), f"{label} refusal outcome changed")
    source_bytes = evidence.get("source_archive_bytes")
    source_sha = evidence.get("source_archive_sha256")
    positive_int(source_bytes, f"{label} source archive bytes")
    require(is_sha(source_sha), f"{label} source archive SHA missing")
    corpus_published_sha = evidence.get("published_sha256")
    require(is_sha(corpus_published_sha), f"{label} corpus publication SHA missing")
    if plan_corpus.get("path"):
        require(source_bytes == plan_corpus["bytes"] and source_sha == plan_corpus["sha256"],
                f"{label} source fixture identity changed")
    published = ordinary.get("published_sha256")
    require(isinstance(published, list), f"{label} publication vector missing")
    require(ordinary.get("publications_identical") is True, f"{label} publication identity failed")
    expected_count = 0 if phase == "edit" else _expected_samples_from_result(result)
    require(len(published) == expected_count, f"{label} publication vector length changed")
    require(all(is_sha(item) for item in published), f"{label} publication digest is invalid")
    if published:
        require(len(set(published)) == 1, f"{label} publication is not deterministic")
        require(all(item == corpus_published_sha for item in published),
                f"{label} publication vector differs from corpus identity")
        require(result.get("output_sha256") == published[0], f"{label} output SHA changed")
    else:
        require(result.get("output_sha256") is None, f"{label} edit unexpectedly published")
    edit_hashes = ordinary.get("edit_outcome_sha256")
    require(isinstance(edit_hashes, list)
            and len(edit_hashes) == _expected_samples_from_result(result),
            f"{label} edit outcome vector length changed")
    require(ordinary.get("edit_outcomes_identical") is True, f"{label} edit identity failed")
    expected_outcome_sha = hashlib.sha256(outcome.encode()).hexdigest()
    require(all(item == expected_outcome_sha for item in edit_hashes),
            f"{label} edit outcome digest does not match its outcome text")
    if phase == "counting_publish":
        require(isinstance(ordinary.get("sample_byte_split"), dict), f"{label} counting split missing")
        require(isinstance(result.get("sink"), dict), f"{label} counting sink missing")
    else:
        require(ordinary.get("sample_byte_split") is None, f"{label} non-counting split present")
    return {"ordinary": ordinary, "corpus": evidence, "outcome": outcome,
            "published": tuple(published), "source_sha": source_sha,
            "source_bytes": source_bytes, "corpus_published_sha": corpus_published_sha}


def check_row_outcome_bindings(row: dict[str, Any], evidence: dict[str, Any], label: str) -> None:
    """Bind capture's duplicated top-level outcome receipts back to the report."""

    for key, expected in (("publication_sha256", list(evidence["published"])),
                          ("edit_outcome_sha256", list(evidence["ordinary"].get("edit_outcome_sha256", []))),
                          ("edit_admitted", evidence["corpus"].get("edit_admitted")),
                          ("edit_outcome", evidence["outcome"]),
                          ("source_archive_sha256", evidence["source_sha"])):
        require(key in row, f"{label} {key} receipt is missing")
        require(row[key] == expected, f"{label} {key} differs from report evidence")


def check_oracle_outcome(evidence: dict[str, Any], oracle_entry: dict[str, Any], label: str) -> None:
    require(evidence["corpus"].get("edit_admitted") == oracle_entry["edit_admitted"],
            f"{label} edit admission differs from export oracle")
    require(evidence["outcome"] == oracle_entry["edit_outcome"],
            f"{label} edit outcome differs from export oracle")
    require(evidence["corpus"].get("edit_description") == oracle_entry["edit_description"],
            f"{label} edit description differs from export oracle")


def _expected_samples_from_result(result: dict[str, Any]) -> int:
    elapsed = result.get("elapsed_ns")
    if isinstance(elapsed, dict) and isinstance(elapsed.get("samples"), list):
        return len(elapsed["samples"])
    fail("result has no elapsed sample vector")


def check_allocation_metrics(result: dict[str, Any], samples: int, label: str) -> dict[str, list[int]]:
    check_operation_alignment(result, samples, label, allocation=True)
    metrics = result["operation_metrics"]
    allocation = metrics.get("allocation")
    require(isinstance(allocation, dict), f"{label} allocation envelope missing")
    require(allocation.get("status") == "measured", f"{label} allocation is not measured")
    require(allocation.get("scope") == "operation_global_system_allocator",
            f"{label} allocation scope changed")
    vectors: dict[str, list[int]] = {}
    for field_name in ALLOCATION_FIELDS:
        metric = allocation.get(field_name)
        require(isinstance(metric, dict), f"{label}.allocation.{field_name} missing")
        require(metric.get("status") == "measured", f"{label}.{field_name} status changed")
        require(metric.get("scope") == "operation_global_system_allocator",
                f"{label}.{field_name} scope changed")
        vector = metric.get("values")
        require(isinstance(vector, list) and len(vector) == samples,
                f"{label}.{field_name} vector length changed")
        for index, value in enumerate(vector):
            nonnegative_int(value, f"{label}.{field_name}[{index}]")
        vectors[field_name] = vector
    require(all(value == 0 for value in vectors["failed_allocation_calls"]),
            f"{label} failed allocation call observed")
    for index, (peak, before) in enumerate(zip(vectors["region_peak_live_bytes"], vectors["live_bytes_before"])):
        require(peak >= before, f"{label} region peak below live-at-start at {index}")
    for index, (before, after, high_before, high_after, region_peak) in enumerate(zip(
        vectors["live_bytes_before"], vectors["live_bytes_after"],
        vectors["peak_live_bytes_before"], vectors["peak_live_bytes_after"],
        vectors["region_peak_live_bytes"])):
        require(after == before + vectors["allocated_bytes"][index]
                - vectors["deallocated_bytes"][index],
                f"{label} live-byte balance does not reconcile at {index}")
        require(high_before >= before and high_after >= after,
                f"{label} process high-water is below live boundary at {index}")
        require(high_after >= high_before,
                f"{label} process high-water regressed at {index}")
        require(region_peak >= after,
                f"{label} region peak below live-at-end at {index}")
        require(high_after >= region_peak,
                f"{label} region peak exceeds process high-water at {index}")
    vectors["net_live"] = [after - before for before, after in
                            zip(vectors["live_bytes_before"], vectors["live_bytes_after"])]
    vectors["peak_above_start"] = [peak - before for peak, before in
                                    zip(vectors["region_peak_live_bytes"], vectors["live_bytes_before"])]
    for index, value in enumerate(vectors["peak_above_start"]):
        require(value >= 0, f"{label} peak_above_start is negative at {index}")
    # Allocation metrics describe the region only.  Their elapsed vector is
    # checked independently and is never used as a latency statistic here.
    return vectors


def check_rss(row: dict[str, Any], label: str) -> int:
    value = field(row, "rss", "rss_kib", "peak_rss")
    require(isinstance(value, dict), f"{label} whole-process RSS receipt is missing")
    path = artifact_receipt(value, f"{label} RSS", capture_bound=True)
    assert path is not None
    text = path.read_text().strip()
    require(text.isdigit(), f"{label} RSS is not numeric")
    measured = int(text)
    positive_int(measured, f"{label} RSS")
    require(value.get("rss_kib") == measured, f"{label} RSS receipt value changed")
    return measured


def check_artifacts(row: dict[str, Any], label: str) -> dict[str, Path]:
    paths: dict[str, Path] = {}
    # capture.py retains process stdout/stderr separately; a child-level
    # ``log`` is optional and is accepted only when the lane recorded one.
    for key in ("report", "stdout", "stderr"):
        value = field(row, key)
        require(isinstance(value, dict), f"{label}.{key} receipt is missing")
        path = artifact_receipt(value, f"{label}.{key}", capture_bound=True)
        assert path is not None
        paths[key] = path
    optional_log = field(row, "log")
    if optional_log is not None:
        require(isinstance(optional_log, dict), f"{label}.log receipt is invalid")
        path = artifact_receipt(optional_log, f"{label}.log", capture_bound=True)
        assert path is not None
        paths["log"] = path
    return paths


def check_row_custody(row: dict[str, Any], label: str, build: dict[str, Any],
                      binary_name: str) -> None:
    for key, path in (("plan_sha256", PACKET / "plan.json"),
                      ("source_sha256", PACKET / "source.json"),
                      ("build_sha256", PACKET / "build.json"),
                      ("runner_sha256", PACKET / "capture.py")):
        value = field(row, key)
        require(is_sha(value), f"{label} {key} receipt is missing")
        require(value == sha256(path), f"{label} {key} changed")
    before = row.get("fixture_before")
    after = row.get("fixture_after")
    require(isinstance(before, dict) and isinstance(after, dict),
            f"{label} fixture custody is missing")
    require(before == after, f"{label} fixture changed during child")
    binary = field(row, "binary")
    require(isinstance(binary, dict), f"{label} binary receipt is missing")
    expected_binary = build.get("binaries", {}).get(binary_name)
    require(isinstance(expected_binary, dict), f"{label} expected {binary_name} binary is missing")
    require(binary == expected_binary, f"{label} binary receipt differs from build.json")


def check_fixture_custody(row: dict[str, Any], plan: dict[str, Any], label: str) -> None:
    before = row["fixture_before"]
    after = row["fixture_after"]
    expected = {
        corpus["id"]: (None if corpus.get("path") is None else {
            "path": corpus["path"], "bytes": corpus["bytes"], "sha256": corpus["sha256"]
        })
        for corpus in plan["corpora"]
    }
    require(before == expected and after == expected,
            f"{label} fixture identity differs from frozen plan")


def check_lane_bindings(complete: dict[str, Any], label: str, *, timed: bool = False) -> None:
    identities_path = PACKET / "qualification-identities.json"
    if "qualification_identities_sha256" in complete:
        require(identities_path.is_file() and
                complete["qualification_identities_sha256"] == sha256(identities_path),
                f"{label} qualification identity binding changed")
    if not timed:
        return
    admission = complete.get("admission")
    require(isinstance(admission, dict), f"{label} admission binding is missing")
    admission_receipt = admission.get("admission")
    path = artifact_receipt(admission_receipt, f"{label} admission receipt", capture_bound=True)
    assert path is not None
    require(path == PACKET / "admission.json", f"{label} admission path changed")
    for key in ("oracle_report", "export_receipt", "export_manifest"):
        artifact_receipt(admission.get(key), f"{label} {key}", capture_bound=True)
    require(admission.get("qualification_identities_sha256") == sha256(identities_path),
            f"{label} admission qualification binding changed")
    before_path = PACKET / "admission-before.json"
    if before_path.is_file():
        require(read_json(before_path) == admission,
                f"{label} admission differs from the timed-lane freeze")


def load_build_and_cleanup() -> tuple[dict[str, Any], bool]:
    build_path = find_packet_file(("build.json", "quality-build.json"))
    require(build_path is not None, "build.json is missing")
    build = read_json(build_path)
    require(isinstance(build, dict), "build.json is not an object")
    if "source_sha256" in build:
        require(build["source_sha256"] == sha256(PACKET / "source.json"),
                "build.json is not bound to the retained source census")
    cleanup_path = PACKET / "cleanup.json"
    cleanup = False
    cleanup_value: Any = None
    if cleanup_path.is_file():
        cleanup_value = read_json(cleanup_path)
        require(isinstance(cleanup_value, dict), "cleanup.json is not an object")
        cleanup = cleanup_value.get("verified") is True or cleanup_value.get("executables_verified_before_removal") is True
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and binaries, "build binaries are missing")
    for name, receipt in binaries.items():
        path = resolve_path(receipt.get("path"), capture_bound=False)
        if path.is_file():
            artifact_receipt(receipt, f"build binary {name}", capture_bound=False)
            continue
        require(cleanup, f"build binary {name} is missing without cleanup verification")
        found = False

        def walk(value: Any) -> None:
            nonlocal found
            if found:
                return
            if isinstance(value, dict):
                if (value.get("path") == receipt.get("path")
                        and value.get("sha256") == receipt.get("sha256")
                        and value.get("bytes", value.get("size")) == receipt.get("bytes", receipt.get("size"))):
                    found = True
                for child in value.values():
                    walk(child)
            elif isinstance(value, list):
                for child in value:
                    walk(child)

        walk(cleanup_value)
        require(found, f"build binary {name} lacks exact cleanup witness")
    return build, cleanup


def _manifest_file(value: Any, base: Path, label: str) -> tuple[Path, int, str]:
    """Validate an artifact named by the export manifest itself."""

    require(isinstance(value, dict), f"{label} is not an artifact record")
    raw = value.get("path")
    require(isinstance(raw, str) and raw and not Path(raw).is_absolute(),
            f"{label}.path is invalid")
    candidate = base / raw
    # Do not let a manifest symlink turn a relative archive path into an
    # unrelated file.  Checking every existing component also catches a
    # symlinked directory in the exported artifact tree.
    cursor = candidate
    while cursor != base and cursor != cursor.parent:
        require(not cursor.is_symlink(), f"{label}.path traverses a symlink")
        cursor = cursor.parent
    path = candidate.resolve()
    try:
        path.relative_to(base.resolve())
    except ValueError:
        fail(f"{label}.path escapes the export directory")
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {raw}")
    size = value.get("bytes", value.get("size"))
    digest = value.get("sha256", value.get("digest"))
    nonnegative_int(size, f"{label}.bytes")
    require(is_sha(digest), f"{label}.sha256 is invalid")
    require(path.stat().st_size == size and sha256(path) == digest,
            f"{label} bytes or SHA changed")
    return path, size, digest


def _oracle_case_id(entry: dict[str, Any], plan: dict[str, Any], label: str) -> str:
    corpora = plan_corpora(plan)
    raw_id = entry.get("id", entry.get("case_id", entry.get("corpus_id", entry.get("case"))))
    require(isinstance(raw_id, str) and raw_id, f"{label} has no corpus id")
    if raw_id in corpora:
        return raw_id
    raw_path = entry.get("input_path", entry.get("path"))
    if isinstance(raw_path, str):
        normalized = raw_path.replace("\\", "/")
        matches = [cid for cid, corpus in corpora.items()
                   if corpus.get("path") and corpus["path"].replace("\\", "/") == normalized]
        if len(matches) == 1:
            return matches[0]
    fmt = str(entry.get("format", entry.get("kind", ""))).lower()
    if raw_id.startswith("generated-"):
        candidate = f"generated-{fmt}" if fmt else raw_id.removesuffix("-medium")
        if candidate in corpora:
            return candidate
    if raw_id.startswith("real-"):
        try:
            ordinal = int(raw_id.split("-", 2)[1])
        except (IndexError, ValueError):
            ordinal = -1
        real = [cid for cid, corpus in corpora.items() if corpus.get("path")]
        if 0 <= ordinal < len(real):
            candidate = real[ordinal]
            if not fmt or corpora[candidate]["format"] == fmt:
                return candidate
    fail(f"{label} corpus id {raw_id!r} is not in the frozen plan")


def _record_digest(value: Any, names: tuple[str, ...]) -> Any:
    if isinstance(value, dict):
        for name in names:
            if name in value:
                return value[name]
        for name in ("output", "source_archive", "source", "published", "result"):
            nested = value.get(name)
            if isinstance(nested, dict):
                found = _record_digest(nested, names)
                if found is not None:
                    return found
    return None


def _policy_records(entry: dict[str, Any], label: str) -> dict[str, dict[str, Any]]:
    raw = entry.get("policy_outputs", entry.get("outputs", entry.get("policies")))
    records: dict[str, dict[str, Any]] = {}
    if isinstance(raw, dict):
        iterator = raw.items()
    elif isinstance(raw, list):
        iterator = []
        for item in raw:
            require(isinstance(item, dict), f"{label} policy output is invalid")
            policy = item.get("policy", item.get("durability", item.get("level")))
            require(isinstance(policy, str), f"{label} policy output has no policy")
            iterator.append((policy, item))
    else:
        fail(f"{label} has no policy output records")
    for policy, value in iterator:
        if policy == "stream":
            # The independent oracle also records the sequential sink
            # control.  The durability publication matrix is exactly the
            # four named policies below; stream is checked by the oracle and
            # does not create a fifth timed policy.
            continue
        require(policy in POLICIES and policy not in records,
                f"{label} has invalid or duplicate policy {policy!r}")
        require(isinstance(value, dict), f"{label}.{policy} policy output is invalid")
        records[policy] = value
    require(set(records) == set(POLICIES), f"{label} policy output set changed")
    return records


def _stream_record(entry: dict[str, Any], label: str, *, manifest_base: Path | None = None,
                   strict_manifest: bool = False, source_sha: str | None = None,
                   source_bytes: int | None = None, published_sha: str | None = None,
                   published_bytes: int | None = None, admitted: bool | None = None) -> dict[str, Any]:
    stream = entry.get("stream_output", entry.get("stream"))
    if stream is None and isinstance(entry.get("policies"), dict):
        stream = entry["policies"].get("stream")
    require(isinstance(stream, dict), f"{label} stream output record is missing")
    stream_sha = _record_digest(stream, ("sha256", "output_sha256", "archive_sha256"))
    stream_bytes = _record_digest(stream, ("bytes", "output_bytes", "archive_bytes", "size"))
    require(is_sha(stream_sha), f"{label} stream output SHA is missing")
    positive_int(stream_bytes, f"{label} stream output bytes")
    require(stream_sha == published_sha and stream_bytes == published_bytes,
            f"{label} stream publication differs from the common publication")
    if strict_manifest:
        require(manifest_base is not None, f"{label} stream manifest base is missing")
        output = stream.get("output", stream)
        _, actual_bytes, actual_sha = _manifest_file(output, manifest_base, f"{label}.stream.output")
        require(actual_bytes == stream_bytes and actual_sha == stream_sha,
                f"{label} stream artifact identity is inconsistent")
        if "matches_reference" in stream:
            require(stream.get("matches_reference") is True,
                    f"{label} stream does not match the common publication")
        if admitted is False:
            require(stream.get("matches_source") is True and
                    stream_sha == source_sha and stream_bytes == source_bytes,
                    f"{label} refusal stream is not byte-exact source")
    return {"sha256": stream_sha, "bytes": stream_bytes}


def _normalize_oracle_case(
    entry: dict[str, Any], plan: dict[str, Any], label: str, *,
    manifest_base: Path | None = None, strict_manifest: bool = False,
) -> dict[str, Any]:
    cid = _oracle_case_id(entry, plan, label)
    corpus = plan_corpora(plan)[cid]
    source_record = entry.get("source_archive")
    if not isinstance(source_record, dict):
        source_record = entry.get("source") if isinstance(entry.get("source"), dict) else entry
    source_sha = _record_digest(source_record, ("sha256", "source_sha256", "archive_sha256"))
    source_bytes = _record_digest(source_record, ("bytes", "source_archive_bytes", "archive_bytes", "size"))
    published_sha = _record_digest(entry, ("published_sha256", "output_sha256"))
    published_bytes = _record_digest(entry, ("published_bytes", "output_bytes"))
    require(is_sha(source_sha), f"{label} source SHA is missing")
    positive_int(source_bytes, f"{label} source bytes")
    require(is_sha(published_sha), f"{label} published SHA is missing")
    positive_int(published_bytes, f"{label} published bytes")
    admitted = entry.get("admitted", entry.get("edit_admitted", entry.get("expected_edit_admitted")))
    require(isinstance(admitted, bool) and admitted is corpus["expected_edit_admitted"],
            f"{label} admission differs from the plan")

    source_path: Path | None = None
    if strict_manifest:
        require(manifest_base is not None, f"{label} manifest base is missing")
        source_path, actual_bytes, actual_sha = _manifest_file(source_record, manifest_base, f"{label}.source")
        require(actual_bytes == source_bytes and actual_sha == source_sha,
                f"{label} source artifact identity is inconsistent")
    policies = _policy_records(entry, label)
    policy_identity: dict[str, dict[str, Any]] = {}
    for policy, record in policies.items():
        policy_sha = _record_digest(record, ("sha256", "output_sha256", "archive_sha256"))
        policy_bytes = _record_digest(record, ("bytes", "output_bytes", "archive_bytes", "size"))
        require(is_sha(policy_sha), f"{label}.{policy} output SHA is missing")
        if policy_bytes is None:
            policy_bytes = published_bytes
        positive_int(policy_bytes, f"{label}.{policy} output bytes")
        require(policy_sha == published_sha and policy_bytes == published_bytes,
                f"{label}.{policy} publication differs from the common publication")
        if strict_manifest:
            _, actual_bytes, actual_sha = _manifest_file(
                record.get("output", record), manifest_base, f"{label}.{policy}.output")
            require(actual_bytes == policy_bytes and actual_sha == policy_sha,
                    f"{label}.{policy} output artifact identity is inconsistent")
        policy_identity[policy] = {"sha256": policy_sha, "bytes": policy_bytes}
    stream_identity = _stream_record(
        entry, label, manifest_base=manifest_base, strict_manifest=strict_manifest,
        source_sha=source_sha, source_bytes=source_bytes,
        published_sha=published_sha, published_bytes=published_bytes, admitted=admitted)
    require(str(entry.get("format", corpus["format"])).lower() == corpus["format"],
            f"{label} format differs from the plan")
    if strict_manifest:
        require(entry.get("origin") == ("generated-harness-corpus" if not corpus.get("path")
                                         else "caller-named-real-file"),
                f"{label} origin differs from the plan")
        if corpus.get("path"):
            require(source_sha == corpus["sha256"] and source_bytes == corpus["bytes"],
                    f"{label} fixture source identity changed")
        if not admitted:
            require(entry.get("refused_output_source_exact") is True,
                    f"{label} refusal source-byte closure is absent")
            require(published_sha == source_sha and published_bytes == source_bytes,
                    f"{label} refusal output is not byte-exact source")
            for policy, record in policies.items():
                require(record.get("matches_source") is True,
                        f"{label}.{policy} refusal is not marked source-exact")
        outcome = entry.get("edit_outcome")
        require(isinstance(outcome, str) and
                ((admitted and outcome == "admitted") or
                 (not admitted and outcome.startswith("refused:"))),
                f"{label} edit outcome changed")
    else:
        outcome = entry.get("edit_outcome")
    description = entry.get("edit_description")
    target = entry.get("edit_target")
    require(isinstance(description, str) and description,
            f"{label} edit description is missing")
    require(isinstance(target, dict), f"{label} edit target is missing")
    return {"id": cid, "entry": entry, "source_sha": source_sha,
            "source_bytes": source_bytes, "published_sha": published_sha,
            "published_bytes": published_bytes, "policies": policy_identity,
            "stream": stream_identity,
            "edit_admitted": admitted, "edit_outcome": outcome,
            "edit_description": description, "edit_target": target,
            "source_path": relative_packet(source_path) if source_path else None}


def _manifest_entries(value: dict[str, Any], label: str) -> list[dict[str, Any]]:
    raw = value.get("cases", value.get("corpora", value.get("entries", value.get("exports"))))
    if isinstance(raw, dict):
        require(all(isinstance(item, dict) for item in raw.values()),
                f"{label} corpus entry is invalid")
        return [{"id": cid, **item} for cid, item in raw.items()]
    require(isinstance(raw, list), f"{label} corpus entries are missing")
    require(all(isinstance(item, dict) for item in raw), f"{label} corpus entry is invalid")
    return raw


def _independent_oracle() -> Any:
    """Load the packet's independent ZIP/XML oracle without executing it as a CLI."""

    module_name = "litchi_0778_independent_oracle"
    cached = sys.modules.get(module_name)
    if cached is not None:
        return cached
    path = PACKET / "oracle.py"
    require(path.is_file() and not path.is_symlink(), "independent oracle runner is missing")
    spec = importlib.util.spec_from_file_location(module_name, path)
    require(spec is not None and spec.loader is not None,
            "cannot load independent oracle runner")
    module = importlib.util.module_from_spec(spec)
    # dataclasses used by oracle.py resolve their module through sys.modules
    # while the file is being executed.
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    return module


_ORACLE_PATH_KEYS = frozenset({
    "manifest", "path", "output_directory", "input_path", "source_path",
    "archive_path", "file_path",
})


def _oracle_path_value(value: str, plan: dict[str, Any]) -> str:
    """Canonicalize only fields whose schema names them as filesystem paths."""

    candidate = Path(value)
    if not candidate.is_absolute():
        return value.replace("\\", "/")
    candidate = candidate.resolve(strict=False)
    current_root = PACKET.parents[3].resolve()
    raw_root = plan.get("root")
    frozen_root = Path(raw_root).resolve(strict=False) if isinstance(raw_root, str) else current_root
    roots = ((PACKET.resolve(), "@packet"),
             ((frozen_root / "docs/performance/results/change-0778").resolve(strict=False), "@packet"),
             (current_root, "@repo"), (frozen_root, "@repo"))
    for root, marker in roots:
        try:
            return f"{marker}/{candidate.relative_to(root).as_posix()}"
        except ValueError:
            continue
    # The oracle itself rejects paths outside the frozen/current roots.  Keep
    # this fail-closed for an unexpected path-valued field rather than
    # turning it into a basename comparison.
    fail(f"independent oracle path escaped frozen/current roots: {value}")


def _normalize_oracle_json(value: Any, plan: dict[str, Any], key: str | None = None) -> Any:
    """Make fresh and retained oracle JSON comparable across checkout moves.

    JSON serialization turns tuples into arrays.  Path relocation is applied
    only beneath explicitly path-valued keys; ordinary strings such as edit
    descriptions, ZIP member names, and digests remain byte-for-byte values.
    """

    if isinstance(value, tuple):
        return [_normalize_oracle_json(item, plan) for item in value]
    if isinstance(value, list):
        return [_normalize_oracle_json(item, plan) for item in value]
    if isinstance(value, dict):
        return {name: _normalize_oracle_json(item, plan, name)
                for name, item in value.items()}
    if key in _ORACLE_PATH_KEYS and isinstance(value, str):
        return _oracle_path_value(value, plan)
    return value


def _check_oracle_bounded_controls(oracle_runner: Any, manifest: dict[str, Any],
                                   manifest_base: Path) -> dict[str, bool]:
    """Replay the small negative controls that guard the independent oracle."""

    prefix = (
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"'
        '><Relationship Id="rId1" Type="urn:test" Target="part.xml"'
    )
    internal = oracle_runner.relationship_map(
        (prefix + '/></Relationships>').encode(), "control.internal")
    explicit = oracle_runner.relationship_map(
        (prefix + ' TargetMode="Internal"/></Relationships>').encode(),
        "control.explicit-internal")
    external = oracle_runner.relationship_map(
        (prefix + ' TargetMode="External"/></Relationships>').encode(),
        "control.external")
    require(internal == explicit and internal != external,
            "independent oracle does not distinguish relationship target modes")
    try:
        oracle_runner.relationship_map(
            (prefix + ' TargetMode="invalid"/></Relationships>').encode(),
            "control.invalid-target-mode")
    except oracle_runner.OracleError:
        pass
    else:
        fail("independent oracle accepted an invalid relationship target mode")

    cases = manifest.get("cases")
    require(isinstance(cases, list) and cases and isinstance(cases[0], dict),
            "export manifest cases are missing for stream-control replay")
    case = dict(cases[0])
    had_stream = case.pop("stream_output", None) is not None
    case.pop("stream", None)
    require(had_stream, "export manifest has no explicit stream output to control")
    try:
        oracle_runner.output_records(case, manifest_base, "control.missing-stream")
    except oracle_runner.OracleError:
        pass
    else:
        fail("independent oracle accepted a manifest without the stream control")
    controls = [
        {"control": "internal-and-external-distinguished", "rejected": True},
        {"control": "invalid-target-mode", "rejected": True},
        {"control": "missing-stream-control", "rejected": True},
    ]
    receipt_path = PACKET / "oracle-extra-controls.json"
    if receipt_path.is_file():
        require(read_json(receipt_path) == controls,
                "oracle-extra-controls.json does not match the bounded controls")
    return {"relationship_target_modes": True, "invalid_target_mode_rejected": True,
            "mandatory_stream_rejected_when_missing": True,
            "retained_receipt": receipt_path.is_file()}


def _bind_packet_artifact(value: Any, expected: Path, label: str) -> None:
    require(isinstance(value, (dict, str)), f"{label} is missing")
    if isinstance(value, dict):
        path = artifact_receipt(value, label, capture_bound=True)
    else:
        path = resolve_path(value, capture_bound=True)
        require(path.is_file(), f"{label} is missing")
    assert path is not None
    require(path == expected and path.stat().st_size == expected.stat().st_size
            and sha256(path) == sha256(expected), f"{label} is not bound to {expected.name}")


def load_oracle(plan: dict[str, Any], build: dict[str, Any], cleanup: bool) -> dict[str, dict[str, Any]]:
    """Replay export.json, its raw manifest, the independent oracle report and admission."""

    export_path = PACKET / "export.json"
    require(export_path.is_file(), "export.json is missing")
    export_value = read_json(export_path)
    require(isinstance(export_value, dict) and export_value.get("exit_code") == 0,
            "exporter did not complete")
    expected_bindings = {
        "source_sha256": sha256(PACKET / "source.json"),
        "plan_sha256": sha256(PACKET / "plan.json"),
        "build_sha256": sha256(PACKET / "build.json"),
        "runner_sha256": sha256(PACKET / "export.py"),
    }
    for key, expected in expected_bindings.items():
        require(export_value.get(key) == expected, f"export.json {key} is not current-source bound")
    require(export_value.get("binary") == build.get("binaries", {}).get("export"),
            "export.json binary is not the retained export build")
    artifact_receipt(export_value.get("binary"), "export binary", capture_bound=False,
                     allow_missing=cleanup)
    for key in ("stdout", "stderr"):
        artifact_receipt(export_value.get(key), f"export {key}", capture_bound=True)
    manifest_value = export_value.get("manifest")
    manifest_path = artifact_receipt(manifest_value, "export manifest", capture_bound=True)
    assert manifest_path is not None

    manifest = read_json(manifest_path)
    require(isinstance(manifest, dict), "export manifest is not an object")
    require(manifest.get("schema_version") == 1 and
            manifest.get("kind") == "ordinary-save-artifact-export" and
            manifest.get("generator") == "litchi-perf-ordinary-save-artifacts-v1",
            "export manifest schema changed")
    require(manifest.get("filesystem_root") == plan.get("filesystem_root"),
            "export manifest filesystem root changed")
    output_directory = manifest.get("output_directory")
    require(isinstance(output_directory, str) and output_directory,
            "export manifest output directory is missing")
    if Path(output_directory).is_absolute():
        output_path = resolve_path(output_directory, capture_bound=True)
    else:
        output_path = (manifest_path.parent / output_directory).resolve()
    require(output_path == manifest_path.parent.resolve(),
            "export manifest output directory is not its artifact directory")
    raw_cases = _manifest_entries(manifest, "export manifest")
    raw_result: dict[str, dict[str, Any]] = {}
    for index, entry in enumerate(raw_cases):
        normalized = _normalize_oracle_case(
            entry, plan, f"export manifest case {index}",
            manifest_base=manifest_path.parent, strict_manifest=True)
        cid = normalized["id"]
        require(cid not in raw_result, f"export manifest duplicates {cid}")
        raw_result[cid] = normalized
    require(set(raw_result) == set(plan_corpora(plan)), "export manifest corpus set changed")

    admission_path = PACKET / "admission.json"
    admission = read_json(admission_path) if admission_path.is_file() else None
    require(isinstance(admission, dict) and admission.get("oracle_pass") is True,
            "admission.json is missing or oracle did not pass")
    oracle_digest = admission.get("oracle_report_sha256")
    require(is_sha(oracle_digest), "admission.json has no oracle report digest")
    require(admission.get("oracle_runner_sha256") == sha256(PACKET / "oracle.py"),
            "admission oracle runner binding changed")
    identities_path = PACKET / "qualification-identities.json"
    require(identities_path.is_file(), "qualification identities are missing")
    identities_digest = admission.get("qualification_identities_sha256")
    require(identities_digest == sha256(identities_path),
            "admission qualification identity digest changed")
    oracle_binding = admission.get("oracle_report")
    oracle_path = artifact_receipt(oracle_binding, "admission oracle report", capture_bound=True)
    assert oracle_path is not None
    require(sha256(oracle_path) == oracle_digest, "oracle report digest differs from admission")

    export_binding = admission.get("export", admission.get("export_receipt", admission.get("export_binding")))
    require(isinstance(export_binding, dict), "admission export binding is missing")
    receipt_binding = export_binding.get("receipt")
    _bind_packet_artifact(receipt_binding, export_path, "admission export receipt")
    require(export_binding.get("binary") == build["binaries"]["export"],
            "admission export binary changed")
    manifest_binding = export_binding.get("manifest")
    _bind_packet_artifact(manifest_binding, manifest_path, "admission export manifest")
    # admission.py binds the export receipt, binary, source revision, and
    # plan/build digests here.  The export receipt itself carries the runner
    # digest and was checked above; it is not duplicated in this binding.
    for key in ("source_sha256", "plan_sha256", "build_sha256"):
        require(export_binding.get(key) == expected_bindings[key],
                f"admission export {key} changed")
    source_value = read_json(PACKET / "source.json")
    require(export_binding.get("source_revision") == source_value.get("revision"),
            "admission export source revision changed")

    oracle_value = read_json(oracle_path)
    require(isinstance(oracle_value, dict), "oracle report is not an object")
    independent = _independent_oracle().check_manifest(
        manifest_path, require_qualification=True)
    oracle_controls = _check_oracle_bounded_controls(
        _independent_oracle(), manifest, manifest_path.parent)
    require(
        _normalize_oracle_json(independent, plan)
        == _normalize_oracle_json(oracle_value, plan),
        "retained oracle report differs from a fresh independent manifest replay",
    )
    require(oracle_value.get("schema") == "litchi-0778-export-v1",
            "oracle report schema changed")
    require(oracle_value.get("status") == "pass", "oracle report did not pass")
    require(oracle_value.get("oracle_runner_sha256") == sha256(PACKET / "oracle.py"),
            "oracle report runner binding changed")
    require(oracle_value.get("manifest_sha256") == sha256(manifest_path),
            "oracle report manifest binding changed")
    require(oracle_value.get("case_count") == len(plan_corpora(plan)) and
            oracle_value.get("policy_count") == len(POLICIES),
            "oracle report cardinality changed")
    qualification_binding = oracle_value.get("qualification_binding")
    require(isinstance(qualification_binding, dict) and
            qualification_binding.get("status") == "bound" and
            qualification_binding.get("sha256") == identities_digest,
            "oracle report qualification binding changed")
    oracle_cases = _manifest_entries(oracle_value, "oracle report")
    oracle_result: dict[str, dict[str, Any]] = {}
    for index, entry in enumerate(oracle_cases):
        normalized = _normalize_oracle_case(entry, plan, f"oracle report case {index}")
        cid = normalized["id"]
        require(cid not in oracle_result, f"oracle report duplicates {cid}")
        oracle_result[cid] = normalized
    require(set(oracle_result) == set(raw_result), "oracle report corpus set changed")
    result: dict[str, dict[str, Any]] = {}
    for cid in plan_corpora(plan):
        raw = raw_result[cid]
        checked = oracle_result[cid]
        for key in ("source_sha", "source_bytes", "published_sha", "published_bytes"):
            require(raw[key] == checked[key], f"oracle report {cid} {key} differs from export manifest")
        require(raw["policies"] == checked["policies"],
                f"oracle report {cid} policy publication differs from export manifest")
        require(raw["stream"] == checked["stream"],
                f"oracle report {cid} stream publication differs from export manifest")
        for key in ("edit_admitted", "edit_description", "edit_target"):
            require(raw[key] == checked[key],
                    f"oracle report {cid} {key} differs from export manifest")
        result[cid] = raw

    fixtures = admission.get("real_fixtures")
    require(isinstance(fixtures, dict), "admission real fixture map is malformed")
    expected_real = {cid for cid, corpus in plan_corpora(plan).items() if corpus.get("path")}
    require(set(fixtures) == expected_real, "admission real fixture set changed")
    for cid, corpus in plan_corpora(plan).items():
        if not corpus.get("path"):
            continue
        row = fixtures.get(cid)
        require(isinstance(row, dict) and row.get("path") == corpus["path"]
                and row.get("bytes") == corpus["bytes"]
                and row.get("sha256") == corpus["sha256"],
                f"admission fixture identity changed for {cid}")
    generated = admission.get("generated")
    require(isinstance(generated, list), "admission generated identity map is malformed")
    generated_result: dict[str, dict[str, Any]] = {}
    for index, entry in enumerate(generated):
        require(isinstance(entry, dict), f"admission generated row {index} is invalid")
        checked = _normalize_oracle_case(entry, plan, f"admission generated row {index}")
        require(checked["id"] not in generated_result,
                f"admission generated identity duplicates {checked['id']}")
        generated_result[checked["id"]] = checked
    expected_generated = {cid for cid, corpus in plan_corpora(plan).items()
                          if not corpus.get("path")}
    require(set(generated_result) == expected_generated,
            "admission generated identity set changed")
    identities = read_json(identities_path).get("corpora", {})
    require(isinstance(identities, dict), "qualification identities corpus map is malformed")
    for cid in expected_generated:
        require(generated_result[cid]["source_sha"] == result[cid]["source_sha"] and
                generated_result[cid]["published_sha"] == result[cid]["published_sha"] and
                generated_result[cid]["policies"] == result[cid]["policies"] and
                generated_result[cid]["stream"] == result[cid]["stream"],
                f"admission generated identity differs from export oracle for {cid}")
        ordinary = identities.get(cid, {}).get("ordinary_corpus", {})
        require(ordinary.get("source_archive_sha256") == result[cid]["source_sha"] and
                ordinary.get("published_sha256") == result[cid]["published_sha"],
                f"admission generated identity differs from qualification for {cid}")
    # Keep the control result attached to the returned map without making it
    # look like a seventh corpus to callers that iterate the oracle map.
    result["__oracle_controls__"] = oracle_controls
    return result


def load_source_custody(plan: dict[str, Any]) -> dict[str, str] | None:
    source_paths = ("source.json", "source-baseline.json", "final-source.json")
    path = find_packet_file(source_paths)
    require(path is not None, "source custody manifest is missing")
    value = read_json(path)
    if isinstance(value, dict) and isinstance(value.get("files"), dict):
        files = value["files"]
    elif isinstance(value, dict):
        files = value
    else:
        fail("source custody is not an object")
    require(files, "source custody is empty")
    result: dict[str, str] = {}
    for name, digest in files.items():
        require(isinstance(name, str) and is_sha(digest), "source custody entry is invalid")
        source_file = (PACKET.parents[3] / name).resolve()
        try:
            source_file.relative_to(PACKET.parents[3].resolve())
        except ValueError:
            fail(f"source custody path escapes repository: {name}")
        require(source_file.is_file() and not source_file.is_symlink(),
                f"source custody file is missing: {name}")
        require(sha256(source_file) == digest, f"source custody file changed: {name}")
        result[name] = digest
    return result


def _quality_attempts() -> list[tuple[Path, Any, list[dict[str, Any]]]]:
    candidates: list[Path] = []
    aggregate = PACKET / "quality.json"
    if aggregate.is_file() and not aggregate.is_symlink():
        candidates.append(aggregate)
    numbered: list[tuple[int, Path]] = []
    for directory in PACKET.glob("quality-*"):
        if not directory.is_dir() or directory.is_symlink():
            continue
        suffix = directory.name.removeprefix("quality-")
        if suffix.isdigit() and (directory / "checks.json").is_file():
            numbered.append((int(suffix), directory / "checks.json"))
    candidates.extend(path for _, path in sorted(numbered))
    attempts: list[tuple[Path, Any, list[dict[str, Any]]]] = []
    for path in candidates:
        value = read_json(path)
        if isinstance(value, list):
            rows = value
        elif isinstance(value, dict):
            rows = value.get("rows")
        else:
            rows = None
        if isinstance(rows, list) and all(isinstance(row, dict) for row in rows):
            attempts.append((path, value, rows))
    return attempts


def quality_receipt() -> tuple[Path, Any, list[dict[str, Any]], list[dict[str, Any]]]:
    attempts = _quality_attempts()
    require(attempts, "quality receipt is missing")
    successful = [(path, value, rows) for path, value, rows in attempts
                  if len(rows) == 9 and all(field(row, "exit", "exit_code") == 0 for row in rows)]
    require(successful, "no retained quality attempt completed all nine commands")
    # A later numbered attempt supersedes an earlier retry.  A root wrapper
    # is selected only when it is the sole successful receipt.
    def attempt_number(path: Path) -> int:
        if path.parent.name.startswith("quality-") and path.parent.name[8:].isdigit():
            return int(path.parent.name[8:])
        return -1
    selected = max(successful, key=lambda item: attempt_number(item[0]))
    failed = [{"path": relative_packet(path), "gates": len(rows),
               "failed": [index + 1 for index, row in enumerate(rows)
                          if field(row, "exit", "exit_code") != 0]}
              for path, _, rows in attempts if not (len(rows) == 9 and
                                                   all(field(row, "exit", "exit_code") == 0 for row in rows))]
    return selected[0], selected[1], selected[2], failed


def _quality_log(row: dict[str, Any], label: str) -> Path:
    log_value = row.get("log")
    if isinstance(log_value, dict):
        path = artifact_receipt(log_value, f"{label} log", capture_bound=True)
        assert path is not None
        return path
    require(isinstance(log_value, str), f"{label} log is missing")
    path = resolve_path(log_value, capture_bound=True)
    try:
        path.relative_to(PACKET.resolve())
    except ValueError:
        fail(f"{label} log escaped packet")
    require(path.is_file() and not path.is_symlink(), f"{label} log is missing")
    expected_sha = row.get("sha256", row.get("log_sha256"))
    require(is_sha(expected_sha) and sha256(path) == expected_sha,
            f"{label} log SHA changed")
    return path


def _check_library_quality(selected_path: Path) -> dict[str, Any] | None:
    path = PACKET / "library-quality.json"
    if not path.is_file():
        return None
    value = read_json(path)
    require(isinstance(value, dict), "library-quality.json is not an object")
    log = _quality_log(value, "library quality")
    require(value.get("overall_exit_code") != 0,
            "library-quality.json no longer records the retained failed attempt")
    text = log.read_text(errors="replace")
    matches = list(__import__("re").finditer(
        r"test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored", text))
    require(any(status == "ok" and int(passed) == value.get("library_passed")
                and int(failed) == 0 and int(ignored) == value.get("library_ignored")
                for status, passed, failed, ignored in (match.groups() for match in matches)),
            "library-quality counts are not bound to its retained log")
    # The failed full-library run belongs to the attempt that owns its log;
    # a later successful retry may be built from a two-line source revision.
    # Bind the diagnostic to that archived predecessor rather than claiming
    # its 556 passing tests for the final source.
    source_path = log.parent / "source.json"
    if source_path.is_file() and value.get("source_sha256") is not None:
        require(value["source_sha256"] == sha256(source_path),
                "library-quality source census changed")
    return {"path": relative_packet(path), "sha256": sha256(path),
            "log": relative_packet(log), "log_sha256": sha256(log),
            "library_passed": value.get("library_passed"),
            "library_ignored": value.get("library_ignored"),
            "overall_exit_code": value.get("overall_exit_code"),
            "scope": value.get("scope")}


def check_library_followup() -> dict[str, Any]:
    """Replay the focused final-source retest and its two-line lineage diff."""

    path = PACKET / "library-followup.json"
    require(path.is_file(), "library-followup.json is missing")
    value = read_json(path)
    require(isinstance(value, dict) and value.get("exit_code") == 0,
            "final-source library follow-up did not pass")
    command = value.get("command")
    expected_command = [
        "cargo", "test", "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--offline", "--locked", "--features", "allocator-metrics", "--lib",
        "ordinary_save::tests::artifact_export_generated_matrix_matches_reference_and_refuses_reuse",
        "--", "--exact", "--test-threads=1",
    ]
    require(command == expected_command, "final-source follow-up command changed")
    log = artifact_receipt(value.get("log"), "library follow-up log", capture_bound=True)
    diff = artifact_receipt(value.get("only_source_difference"),
                            "library follow-up source diff", capture_bound=True)
    assert log is not None and diff is not None
    require(value.get("source_sha256") == sha256(PACKET / "source.json"),
            "library follow-up final source binding changed")
    predecessor = PACKET / "quality-0" / "source.json"
    require(predecessor.is_file() and value.get("predecessor_source_sha256") == sha256(predecessor),
            "library follow-up predecessor source binding changed")
    require(value.get("replacement_count") == 2, "library follow-up replacement count changed")
    diff_text = diff.read_text(errors="replace")
    require("tools/perf-baseline/src/ordinary_save.rs" in diff_text,
            "library follow-up diff has no ordinary_save source path")
    removed = [line for line in diff_text.splitlines()
               if line.startswith("-") and not line.startswith("---")
               and ".then(|| corpus.pptx_target)" in line]
    added = [line for line in diff_text.splitlines()
             if line.startswith("+") and not line.startswith("+++")
             and ".then_some(corpus.pptx_target)" in line]
    require(len(removed) == 2,
            "library follow-up diff does not retain both predecessor lines")
    require(len(added) == 2,
            "library follow-up diff does not retain both final replacements")
    require(value.get("scope") and "predecessor" in str(value["scope"]).lower(),
            "library follow-up scope does not disclose predecessor-only full-suite evidence")
    return {"path": relative_packet(path), "sha256": sha256(path),
            "log": relative_packet(log), "log_sha256": sha256(log),
            "diff": relative_packet(diff), "diff_sha256": sha256(diff),
            "source_sha256": value["source_sha256"],
            "predecessor_source_sha256": value["predecessor_source_sha256"],
            "replacement_count": value["replacement_count"], "scope": value["scope"]}


def optional_library_followup() -> dict[str, Any] | None:
    """Replay the legacy focused follow-up only when this packet retains it.

    A current quality attempt may include the library/export regression in its
    nine-command matrix.  In that case the archived predecessor exception and
    its follow-up are deliberately outside the final packet and are not
    required here.
    """

    if not (PACKET / "library-followup.json").is_file():
        return None
    return check_library_followup()


def check_quality() -> dict[str, Any]:
    quality_path, quality, rows, failed_attempts = quality_receipt()
    # Retained failed attempts are evidence too: verify every log receipt so
    # the disclosure cannot be edited while the successful retry is replayed.
    for attempt_path, _, attempt_rows in _quality_attempts():
        if attempt_path == quality_path:
            continue
        for index, row in enumerate(attempt_rows):
            _quality_log(row, f"retained quality {relative_packet(attempt_path)} gate {index + 1}")
    counts = {"passed": 0, "failed": 0, "ignored": 0, "suites": 0}
    for index, row in enumerate(rows):
        require(field(row, "exit", "exit_code") == 0, f"quality gate {index + 1} failed")
        log = _quality_log(row, f"quality gate {index + 1}")
        text = log.read_text(errors="replace")
        for match in __import__("re").finditer(
                r"test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored", text):
            status, passed, failed, ignored = match.groups()
            require(status == "ok" and int(failed) == 0, f"quality gate {index + 1} test failed")
            counts["passed"] += int(passed)
            counts["failed"] += int(failed)
            counts["ignored"] += int(ignored)
            counts["suites"] += 1
    metadata = quality if isinstance(quality, dict) else {}
    stored = metadata.get("counts", metadata.get("test_counts"))
    if stored is not None:
        require(isinstance(stored, dict), "quality counts are malformed")
        for key, value in counts.items():
            if key in stored:
                require(stored[key] == value, f"quality {key} count changed")
    commands = []
    for index, row in enumerate(rows):
        log = _quality_log(row, f"quality gate {index + 1}")
        commands.append({"index": index + 1, "command": row.get("command"),
                         "exit_code": field(row, "exit", "exit_code"),
                         "log": relative_packet(log), "log_sha256": sha256(log)})
    build_path = PACKET / "build.json"
    if build_path.is_file():
        build_value = read_json(build_path)
        build_rows = build_value.get("rows") if isinstance(build_value, dict) else None
        require(isinstance(build_rows, list) and len(build_rows) == len(commands),
                "build quality rows are missing or incomplete")
        for index, (build_row, command_row) in enumerate(zip(build_rows, commands)):
            require(isinstance(build_row, dict), f"build quality row {index + 1} is invalid")
            require(build_row.get("command") == command_row["command"],
                    f"build quality command {index + 1} differs from selected quality")
            require(field(build_row, "exit", "exit_code") == 0,
                    f"build quality row {index + 1} failed")
            build_log = build_row.get("log")
            if isinstance(build_log, dict):
                build_log_path = artifact_receipt(build_log, f"build quality gate {index + 1} log",
                                                  capture_bound=True)
                assert build_log_path is not None
            else:
                require(build_log == command_row["log"],
                        f"build quality log {index + 1} differs from selected quality")
                build_log_path = resolve_path(build_log, capture_bound=True)
                require(build_log_path.is_file(), f"build quality log {index + 1} is missing")
            require(sha256(build_log_path) == command_row["log_sha256"],
                    f"build quality log {index + 1} hash differs from selected quality")
    docx_quality = check_docx_quality()
    return {"path": relative_packet(quality_path), "gates": len(rows), "counts": counts,
            "sha256": sha256(quality_path), "failed_attempts": failed_attempts,
            "commands": commands,
            "library_followup": _check_library_quality(quality_path),
            "final_source_followup": optional_library_followup(),
            "docx_owner": docx_quality}


def _docx_quality_attempts() -> list[tuple[Path, Any, list[dict[str, Any]]]]:
    candidates: list[Path] = []
    aggregate = PACKET / "docx-quality.json"
    if aggregate.is_file() and not aggregate.is_symlink():
        candidates.append(aggregate)
    numbered: list[tuple[int, Path]] = []
    for directory in PACKET.glob("docx-quality-*"):
        if not directory.is_dir() or directory.is_symlink():
            continue
        suffix = directory.name.removeprefix("docx-quality-")
        if suffix.isdigit() and (directory / "checks.json").is_file():
            numbered.append((int(suffix), directory / "checks.json"))
    candidates.extend(path for _, path in sorted(numbered))
    attempts: list[tuple[Path, Any, list[dict[str, Any]]]] = []
    for path in candidates:
        value = read_json(path)
        rows = value if isinstance(value, list) else value.get("rows") if isinstance(value, dict) else None
        if isinstance(rows, list) and all(isinstance(row, dict) for row in rows):
            attempts.append((path, value, rows))
    return attempts


def check_docx_quality() -> dict[str, Any]:
    attempts = _docx_quality_attempts()
    require(attempts, "DOCX owner quality receipt is missing")
    successful = [(path, value, rows) for path, value, rows in attempts
                  if len(rows) == 5 and all(field(row, "exit", "exit_code") == 0 for row in rows)]
    require(successful, "no retained DOCX owner quality attempt completed five commands")

    def attempt_number(path: Path) -> int:
        return int(path.parent.name.removeprefix("docx-quality-")) \
            if path.parent.name.startswith("docx-quality-") \
            and path.parent.name.removeprefix("docx-quality-").isdigit() else -1

    selected_path, selected_value, selected_rows = max(successful, key=lambda item: attempt_number(item[0]))
    selected_source_path = selected_path.parent / "source.json"
    require(selected_source_path.is_file(), "selected DOCX quality source census is missing")
    source_value = read_json(PACKET / "source.json")
    require(read_json(selected_source_path) == source_value,
            "selected DOCX quality source census differs from final source custody")
    metadata = selected_value if isinstance(selected_value, dict) else {}
    aggregate_path = PACKET / "docx-quality.json"
    require(aggregate_path.is_file(), "docx-quality.json is missing")
    aggregate = read_json(aggregate_path)
    require(isinstance(aggregate, dict) and aggregate.get("rows") == selected_rows,
            "docx-quality.json does not bind the selected five-command attempt")
    require(aggregate.get("source_sha256") == sha256(selected_source_path),
            "docx-quality.json source binding changed")
    require(aggregate.get("source_revision") == source_value.get("revision"),
            "docx-quality.json source revision changed")

    failed_attempts: list[dict[str, Any]] = []
    superseded_attempts: list[dict[str, Any]] = []
    for path, _, rows in attempts:
        if path == selected_path or path == aggregate_path:
            continue
        for index, row in enumerate(rows):
            _quality_log(row, f"retained DOCX quality {relative_packet(path)} gate {index + 1}")
        attempt_source = path.parent / "source.json"
        superseded = {"path": relative_packet(path), "gates": len(rows),
                      "source_sha256": sha256(attempt_source) if attempt_source.is_file() else None}
        superseded_attempts.append(superseded)
        if len(rows) != 5 or any(field(row, "exit", "exit_code") != 0 for row in rows):
            superseded["failed"] = [index + 1 for index, row in enumerate(rows)
                                     if field(row, "exit", "exit_code") != 0]
            failed_attempts.append(superseded)
    commands: list[dict[str, Any]] = []
    counts = {"passed": 0, "failed": 0, "ignored": 0, "suites": 0}
    import re
    for index, row in enumerate(selected_rows):
        require(field(row, "exit", "exit_code") == 0,
                f"DOCX owner quality gate {index + 1} failed")
        log = _quality_log(row, f"DOCX owner quality gate {index + 1}")
        text = log.read_text(errors="replace")
        for match in re.finditer(r"test result: (\w+)\. (\d+) passed; (\d+) failed; (\d+) ignored", text):
            status, passed, failed, ignored = match.groups()
            require(status == "ok" and int(failed) == 0,
                    f"DOCX owner quality gate {index + 1} test failed")
            counts["passed"] += int(passed)
            counts["failed"] += int(failed)
            counts["ignored"] += int(ignored)
            counts["suites"] += 1
        commands.append({"index": index + 1, "command": row.get("command"),
                         "exit_code": field(row, "exit", "exit_code"),
                         "log": relative_packet(log), "log_sha256": sha256(log)})
    return {"path": relative_packet(aggregate_path), "sha256": sha256(aggregate_path),
            "attempt": relative_packet(selected_path), "gates": len(selected_rows),
            "counts": counts, "commands": commands, "failed_attempts": failed_attempts,
            "superseded_attempts": superseded_attempts,
            "source_sha256": sha256(selected_source_path),
            "source_revision": source_value.get("revision")}


def load_qualification_identities(plan: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "qualification-identities.json"
    require(path.is_file(), "qualification-identities.json is missing")
    value = read_json(path)
    require(isinstance(value, dict), "qualification identities are not an object")
    require(value.get("schema") == "litchi-0778-durability-qualification-v1",
            "qualification identity schema changed")
    for key, source_path in (("plan_sha256", PACKET / "plan.json"),
                             ("source_sha256", PACKET / "source.json"),
                             ("build_sha256", PACKET / "build.json"),
                             ("runner_sha256", PACKET / "capture.py")):
        if key in value:
            require(value[key] == sha256(source_path),
                    f"qualification identities {key} changed")
    corpora = value.get("corpora")
    require(isinstance(corpora, dict), "qualification identities corpus map is missing")
    expected = set(plan_corpora(plan))
    require(set(corpora) == expected, "qualification identity corpus set changed")
    for cid, identity in corpora.items():
        require(isinstance(identity, dict), f"qualification identity {cid} is invalid")
        require(isinstance(identity.get("result_corpus"), dict),
                f"qualification identity {cid} result corpus is missing")
        require(isinstance(identity.get("ordinary_corpus"), dict),
                f"qualification identity {cid} ordinary corpus is missing")
    return {"path": relative_packet(path), "sha256": sha256(path), "corpora": corpora}


def check_report_identity(
    result: dict[str, Any], row: dict[str, Any], expected: dict[str, Any], label: str
) -> None:
    source = result.get("source")
    require(isinstance(source, dict) and isinstance(source.get("ordinary_save"), dict),
            f"{label} report identity source is missing")
    stable = {"result_corpus": result.get("corpus"),
              "ordinary_corpus": source["ordinary_save"].get("corpus")}
    require(stable == expected, f"{label} corpus identity differs from qualification")
    digest = field(row, "report_identity_sha256")
    require(is_sha(digest), f"{label} report identity digest is missing")
    require(digest == capture_json_sha256(stable), f"{label} report identity digest changed")


def check_native(plan: dict[str, Any], oracle: dict[str, dict[str, Any]],
                 build: dict[str, Any], cleanup: bool,
                 qualification_identities: dict[str, dict[str, Any]]) -> dict[str, Any]:
    directory = find_lane("native")
    complete = load_complete(directory, "native")
    check_lane_bindings(complete, "native", timed=True)
    rows = load_runs(directory, "native")
    expected_count = plan["expected_children"]["native"]
    require(len(rows) == expected_count, f"native child count changed: {len(rows)} != {expected_count}")
    check_scheduled_rows(rows, plan, "native")
    corpora = plan_corpora(plan)
    entries: list[dict[str, Any]] = []
    identity_seen: set[tuple[Any, ...]] = set()
    for index, row in enumerate(rows):
        label = f"native run {index}"
        identity = row_identity(row, plan, label)
        key = (identity["corpus_id"], identity["phase"], identity["policy"], identity["block"], identity["control"])
        require(key not in identity_seen, f"duplicate native identity: {key}")
        identity_seen.add(key)
        check_row_custody(row, label, build, "native")
        check_fixture_custody(row, plan, label)
        artifacts = check_artifacts(row, label)
        rss = check_rss(row, label)
        report = read_json(artifacts["report"])
        require(isinstance(report, dict), f"{label} report is not an object")
        check_report_binary_identity(report, row, label)
        result = check_report_metadata(report, row, plan, plan["native"]["samples"], plan["native"]["warmup"], label)
        case = field(row, "case")
        if case is not None:
            require(result.get("case") == case, f"{label} case changed")
        elapsed, stats = check_elapsed(result.get("elapsed_ns"), plan["native"]["samples"], label)
        check_operation_alignment(result, plan["native"]["samples"], label)
        evidence = corpus_evidence(result, corpora[identity["corpus_id"]], identity, label)
        check_row_outcome_bindings(row, evidence, label)
        check_report_identity(result, row, qualification_identities[identity["corpus_id"]], label)
        oracle_entry = oracle[identity["corpus_id"]]
        check_oracle_outcome(evidence, oracle_entry, label)
        require(evidence["source_sha"] == oracle_entry["source_sha"],
                f"{label} source archive differs from export oracle")
        require(evidence["corpus_published_sha"] == oracle_entry["published_sha"],
                f"{label} corpus publication differs from export oracle")
        require((not evidence["published"])
                or all(item == oracle[identity["corpus_id"]]["published_sha"]
                       for item in evidence["published"]),
                f"{label} published hash differs from export oracle")
        binary = field(row, "binary")
        if binary is not None:
            artifact_receipt(binary, f"{label} binary", capture_bound=False, allow_missing=cleanup)
        entries.append({"identity": identity, "row": row, "result": result,
                        "elapsed": elapsed, "stats": stats, "rss_kib": rss,
                        "evidence": evidence, "report": relative_packet(artifacts["report"]),
                        "log": relative_packet(artifacts.get("log", artifacts["report"])),
                        "report_sha256": sha256(artifacts["report"])})
    # Every timed combination must have all four policies and all four blocks.
    blocks = set(range(plan["native"]["blocks"]))
    for cid in corpora:
        for phase in TIMED_PHASES:
            for policy in POLICIES:
                got = {item["identity"]["block"] for item in entries
                       if item["identity"]["corpus_id"] == cid and item["identity"]["phase"] == phase
                       and item["identity"]["policy"] == policy}
                require(got == blocks, f"native combination incomplete: {cid}/{phase}/{policy}")
    controls = [item for item in entries if item["identity"]["phase"] in CONTROL_PHASES]
    require(controls, "native default controls are missing")
    require(all(item["identity"]["policy"] == "default" for item in controls),
            "native control uses a non-default policy")
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for item in entries:
        i = item["identity"]
        if i["phase"] in TIMED_PHASES or i["phase"] in CONTROL_PHASES:
            grouped.setdefault((i["corpus_id"], i["phase"], i["policy"]), []).append(item)
    for key, values in grouped.items():
        publication = {item["evidence"]["published"] for item in values}
        require(len(publication) <= 1, f"policy publication hashes diverged: {key}")
    return {"directory": relative_packet(directory), "complete": complete, "rows": entries,
            "children": len(entries), "grouped": grouped, "controls": len(controls)}


def check_allocation(plan: dict[str, Any], oracle: dict[str, dict[str, Any]],
                     build: dict[str, Any], cleanup: bool,
                     qualification_identities: dict[str, dict[str, Any]]) -> dict[str, Any]:
    directory = find_lane("allocation")
    complete = load_complete(directory, "allocation")
    check_lane_bindings(complete, "allocation", timed=True)
    rows = load_runs(directory, "allocation")
    expected_count = plan["expected_children"]["allocation"]
    require(len(rows) == expected_count, f"allocation child count changed: {len(rows)} != {expected_count}")
    check_scheduled_rows(rows, plan, "allocation")
    corpora = plan_corpora(plan)
    entries: list[dict[str, Any]] = []
    identity_seen: set[tuple[Any, ...]] = set()
    for index, row in enumerate(rows):
        label = f"allocation run {index}"
        identity = row_identity(row, plan, label)
        key = (identity["corpus_id"], identity["phase"], identity["policy"], identity["block"], identity["control"])
        require(key not in identity_seen, f"duplicate allocation identity: {key}")
        identity_seen.add(key)
        check_row_custody(row, label, build, "allocation")
        check_fixture_custody(row, plan, label)
        artifacts = check_artifacts(row, label)
        rss = check_rss(row, label)
        report = read_json(artifacts["report"])
        require(isinstance(report, dict), f"{label} report is not an object")
        check_report_binary_identity(report, row, label, allocation=True)
        result = check_report_metadata(report, row, plan, plan["allocation"]["samples"], plan["allocation"]["warmup"], label)
        elapsed, stats = check_elapsed(result.get("elapsed_ns"), plan["allocation"]["samples"], label)
        vectors = check_allocation_metrics(result, plan["allocation"]["samples"], label)
        evidence = corpus_evidence(result, corpora[identity["corpus_id"]], identity, label)
        check_row_outcome_bindings(row, evidence, label)
        check_report_identity(result, row, qualification_identities[identity["corpus_id"]], label)
        oracle_entry = oracle[identity["corpus_id"]]
        check_oracle_outcome(evidence, oracle_entry, label)
        require(evidence["source_sha"] == oracle_entry["source_sha"],
                f"{label} source archive differs from export oracle")
        require(evidence["corpus_published_sha"] == oracle_entry["published_sha"],
                f"{label} corpus publication differs from export oracle")
        require((not evidence["published"])
                or all(item == oracle[identity["corpus_id"]]["published_sha"]
                       for item in evidence["published"]),
                f"{label} published hash differs from export oracle")
        binary = field(row, "binary")
        if binary is not None:
            artifact_receipt(binary, f"{label} binary", capture_bound=False, allow_missing=cleanup)
        entries.append({"identity": identity, "row": row, "result": result, "elapsed": elapsed,
                        "stats": stats, "rss_kib": rss, "allocation": vectors,
                        "evidence": evidence, "report": relative_packet(artifacts["report"]),
                        "log": relative_packet(artifacts.get("log", artifacts["report"])),
                        "report_sha256": sha256(artifacts["report"])})
    blocks = set(range(plan["allocation"]["blocks"]))
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    controls = [item for item in entries if item["identity"]["phase"] in CONTROL_PHASES]
    require(controls, "allocation default controls are missing")
    require(all(item["identity"]["policy"] == "default" for item in controls),
            "allocation control uses a non-default policy")
    for item in entries:
        i = item["identity"]
        if i["phase"] in TIMED_PHASES or i["phase"] in CONTROL_PHASES:
            grouped.setdefault((i["corpus_id"], i["phase"], i["policy"]), []).append(item)
    for key, values in grouped.items():
        require({item["identity"]["block"] for item in values} == blocks,
                f"allocation combination incomplete: {key}")
    return {"directory": relative_packet(directory), "complete": complete, "rows": entries,
            "children": len(entries), "grouped": grouped, "controls": len(controls)}


def spread(values: Iterable[float]) -> float:
    numbers = list(values)
    require(numbers and all(value > 0 for value in numbers), "spread values must be positive")
    return (max(numbers) - min(numbers)) * 100.0 / min(numbers)


def distribution(items: list[dict[str, Any]], metric: str) -> dict[str, Any]:
    values = [float(item["stats"][metric]) for item in items]
    return {"values": values, "min": min(values), "max": max(values),
            "median": statistics.median(values), "spread_percent": spread(values),
            "flag_over_5_percent": spread(values) > REVIEW_PERCENT}


def native_analysis(native: dict[str, Any]) -> dict[str, Any]:
    grouped = native["grouped"]
    by_group: dict[str, Any] = {}
    flags: list[dict[str, Any]] = []
    for key in sorted(grouped):
        items = grouped[key]
        metrics = {metric: distribution(items, metric) for metric in METRICS}
        for metric, value in metrics.items():
            if value["flag_over_5_percent"]:
                flags.append({"group": list(key), "metric": metric,
                              "spread_percent": value["spread_percent"]})
        rss_values = [float(item["rss_kib"]) for item in items]
        rss = {"values": rss_values, "min": min(rss_values), "max": max(rss_values),
               "median": statistics.median(rss_values), "spread_percent": spread(rss_values),
               "flag_over_5_percent": spread(rss_values) > REVIEW_PERCENT}
        by_group["/".join(key)] = {
            "corpus_id": key[0], "phase": key[1], "policy": key[2],
            "processes": [{"block": item["identity"]["block"], **item["stats"],
                           "rss_kib": item["rss_kib"], "report": item["report"]}
                          for item in sorted(items, key=lambda item: item["identity"]["block"])],
            "elapsed_distribution_across_4_processes": metrics,
            "whole_process_rss_distribution_across_4_processes": rss,
        }
    paired: list[dict[str, Any]] = []
    for cid in sorted({key[0] for key in grouped}):
        for phase in TIMED_PHASES:
            for policy in POLICIES:
                if policy == "default":
                    continue
                left = grouped.get((cid, phase, "default"), [])
                right = grouped.get((cid, phase, policy), [])
                require(len(left) == len(right) == 4, f"paired policy rows incomplete: {cid}/{phase}/{policy}")
                for block in range(4):
                    a = next(item for item in left if item["identity"]["block"] == block)
                    b = next(item for item in right if item["identity"]["block"] == block)
                    metrics: dict[str, Any] = {}
                    for metric in METRICS:
                        base = float(a["stats"][metric])
                        value = float(b["stats"][metric])
                        metrics[metric] = {"default": base, "policy": value,
                                           "change_percent": (value / base - 1.0) * 100.0,
                                           "over_5_percent": abs(value / base - 1.0) > REVIEW_PERCENT / 100.0}
                    paired.append({"corpus_id": cid, "phase": phase, "policy": policy,
                                   "block": block, "metrics": metrics,
                                   "comparison": "paired by block; descriptive policy comparison, not a time-causal claim"})
    paired_full: list[dict[str, Any]] = []
    for cid in sorted({key[0] for key in grouped}):
        for phase in TIMED_PHASES:
            reference = grouped.get((cid, phase, "full"), [])
            require(len(reference) == 4, f"paired full-reference rows incomplete: {cid}/{phase}")
            for policy in POLICIES:
                if policy == "full":
                    continue
                compared = grouped.get((cid, phase, policy), [])
                require(len(compared) == 4,
                        f"paired full-reference rows incomplete: {cid}/{phase}/{policy}")
                for block in range(4):
                    base_row = next(item for item in reference
                                    if item["identity"]["block"] == block)
                    value_row = next(item for item in compared
                                     if item["identity"]["block"] == block)
                    metrics: dict[str, Any] = {}
                    for metric in METRICS:
                        base = float(base_row["stats"][metric])
                        value = float(value_row["stats"][metric])
                        metrics[metric] = {"full": base, "policy": value,
                                           "change_percent": (value / base - 1.0) * 100.0,
                                           "over_5_percent": abs(value / base - 1.0) > REVIEW_PERCENT / 100.0}
                    paired_full.append({
                        "corpus_id": cid, "phase": phase, "policy": policy,
                        "block": block, "metrics": metrics,
                        "comparison": "paired by block against full; descriptive policy comparison, not a time-causal claim",
                    })
    aa: list[dict[str, Any]] = []
    for cid in sorted({key[0] for key in grouped}):
        for phase in TIMED_PHASES:
            default = grouped.get((cid, phase, "default"), [])
            full = grouped.get((cid, phase, "full"), [])
            require(len(default) == len(full) == 4, f"default/full control rows incomplete: {cid}/{phase}")
            metrics = {}
            for metric in METRICS:
                values = [(float(next(item for item in full if item["identity"]["block"] == block)["stats"][metric]) /
                           float(next(item for item in default if item["identity"]["block"] == block)["stats"][metric]) - 1.0) * 100.0
                          for block in range(4)]
                metrics[metric] = {"change_percent_by_block": values,
                                   "flag_over_5_percent": any(abs(value) > REVIEW_PERCENT for value in values)}
            aa.append({"corpus_id": cid, "phase": phase, "metrics": metrics,
                       "comparison": "default versus explicit full A/A-like control; timing evidence only, not causal"})
    return {"groups": by_group, "spread_flags_over_5_percent": flags,
            "paired_by_block_default": paired, "paired_by_block_full": paired_full,
            "default_vs_full_aa_like_controls": aa,
            "rss_is_separate_from_elapsed": True}


def allocation_analysis(allocation: dict[str, Any]) -> dict[str, Any]:
    grouped = allocation["grouped"]
    fields = list(ALLOCATION_FIELDS) + ["net_live", "peak_above_start"]
    output: dict[str, Any] = {}
    flags: list[dict[str, Any]] = []
    for key in sorted(grouped):
        items = grouped[key]
        metrics: dict[str, Any] = {}
        for name in fields:
            per_block = [{"block": item["identity"]["block"],
                          "values": item["allocation"][name],
                          "stats": integer_stats(item["allocation"][name])}
                         for item in sorted(items, key=lambda item: item["identity"]["block"])]
            medians = [float(item["stats"]["p50"]) for item in per_block]
            value = {"per_block": per_block, "repeat_p50_values": medians,
                     "spread_percent": signed_spread(medians),
                     "flag_over_5_percent": signed_spread(medians) > REVIEW_PERCENT}
            metrics[name] = value
            if value["flag_over_5_percent"]:
                flags.append({"group": list(key), "metric": name,
                              "spread_percent": value["spread_percent"]})
        output["/".join(key)] = {"corpus_id": key[0], "phase": key[1], "policy": key[2],
                                  "blocks": len(items), "metrics": metrics,
                                  "elapsed_not_mixed": True}
    return {"groups": output, "spread_flags_over_5_percent": flags,
            "metrics": fields,
            "note": "Allocation values are operation-region observations; they are not elapsed latency and are never summed across phases."}


_TRACE_PARSER: Any = None


def _trace_parser() -> Any:
    """Load the retained independent syscall parser without executing a lane."""

    global _TRACE_PARSER
    if _TRACE_PARSER is not None:
        return _TRACE_PARSER
    path = PACKET / "trace_analysis.py"
    require(path.is_file() and not path.is_symlink(), "trace analysis parser is missing")
    module_name = "litchi_0778_trace_analysis"
    cached = sys.modules.get(module_name)
    if cached is not None:
        _TRACE_PARSER = cached
        return cached
    spec = importlib.util.spec_from_file_location(module_name, path)
    require(spec is not None and spec.loader is not None,
            "cannot load trace analysis parser")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    try:
        spec.loader.exec_module(module)
    except Exception:
        sys.modules.pop(module_name, None)
        raise
    _TRACE_PARSER = module
    return module


def _trace_case(corpus: dict[str, Any]) -> str:
    prefix = f"{corpus['format']}_ordinary_save_"
    if corpus.get("path") is not None:
        prefix = f"{corpus['format']}_real_file_ordinary_save_"
    return prefix + "atomic_publish"


def _trace_artifact(value: Any, label: str, *, capture_bound: bool = True,
                    allow_missing: bool = False) -> Path | None:
    return artifact_receipt(value, label, capture_bound=capture_bound,
                            allow_missing=allow_missing)


def _trace_receipt_equal(left: Any, right: Any, label: str,
                         *, capture_bound: bool = True) -> None:
    require(isinstance(left, dict) and isinstance(right, dict),
            f"{label} artifact receipt is invalid")
    for key in ("bytes", "size", "sha256", "digest"):
        if key in left or key in right:
            require(left.get(key) == right.get(key), f"{label}.{key} differs")
    left_path = left.get("path")
    right_path = right.get("path")
    require(isinstance(left_path, str) and isinstance(right_path, str),
            f"{label}.path is missing")
    require(resolve_path(left_path, capture_bound=capture_bound)
            == resolve_path(right_path, capture_bound=capture_bound),
            f"{label}.path differs after relocation")


def _trace_command_path(value: Any, expected: Path, label: str,
                        *, capture_bound: bool) -> None:
    require(isinstance(value, str) and value, f"{label} is missing")
    actual = resolve_path(value, capture_bound=capture_bound)
    require(actual == expected, f"{label} differs after relocation")


def _check_trace_command(command: Any, run: dict[str, Any], plan: dict[str, Any],
                         corpus: dict[str, Any], policy: str,
                         binary_path: Path, trace_path: Path, report_path: Path,
                         parser: Any, label: str) -> list[str]:
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{label} command is invalid")
    options = list(parser.TRACE_OPTIONS)
    prefix = ["taskset", "-c", str(plan["cpu"]), "/usr/bin/strace"] + options
    require(command[:len(prefix)] == prefix, f"{label} trace command options changed")
    index = len(prefix)
    require(index + 1 < len(command) and command[index] == "-o",
            f"{label} trace output option is missing")
    _trace_command_path(command[index + 1], trace_path, f"{label} trace output",
                        capture_bound=True)
    index += 2
    require(index < len(command), f"{label} benchmark binary is missing")
    _trace_command_path(command[index], binary_path, f"{label} benchmark binary",
                        capture_bound=False)
    index += 1
    fixed = [
        ("--case", _trace_case(corpus)),
        ("--samples", "1"),
        ("--warmup", "0"),
        ("--filesystem-root", plan["filesystem_root"]),
    ]
    for option, expected in fixed:
        require(index + 1 < len(command) and command[index] == option,
                f"{label} {option} option is missing")
        require(command[index + 1] == expected,
                f"{label} {option} value changed")
        index += 2
    require(index + 1 < len(command) and command[index] == "--json",
            f"{label} report option is missing")
    _trace_command_path(command[index + 1], report_path, f"{label} report",
                        capture_bound=True)
    index += 2
    if corpus.get("path") is not None:
        require(index + 1 < len(command) and command[index] == "--ooxml-file",
                f"{label} real-file fixture option is missing")
        require(command[index + 1] == corpus["path"],
                f"{label} real-file fixture path changed")
        index += 2
    if policy != "default":
        require(index + 1 < len(command) and command[index] == "--save-durability",
                f"{label} durability option is missing")
        require(command[index + 1] == policy, f"{label} durability policy changed")
        index += 2
    require(index == len(command), f"{label} has unexpected command arguments")
    return list(command)


def _trace_transaction_summary(transaction: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(transaction, dict), f"{label} transaction is invalid")
    destination = transaction.get("destination")
    temp_path = transaction.get("temp_path")
    require(isinstance(destination, str) and isinstance(temp_path, str),
            f"{label} transaction paths are missing")
    return {
        "index": transaction.get("index"),
        "destination": Path(destination).name,
        "event_range": transaction.get("event_range"),
        "line_range": transaction.get("line_range"),
        "temp_name": Path(temp_path).name,
        "temp_fd": transaction.get("temp_fd"),
        "parent_fd": transaction.get("parent_fd"),
        "permission_probe_enoent": transaction.get("permission_probe_enoent"),
        "window_ns": transaction.get("window_ns"),
        "syscall_total_ns": transaction.get("syscall_total_ns"),
        "written_bytes": transaction.get("written_bytes"),
        "write_calls": transaction.get("write_calls"),
        "file_sync_calls": transaction.get("file_sync_calls"),
        "parent_sync_calls": transaction.get("parent_sync_calls"),
        "fchmod_calls": transaction.get("fchmod_calls"),
        "replacement_renames": transaction.get("replacement_renames"),
        "contract": transaction.get("contract"),
    }


def _check_trace_row_summary(retained: dict[str, Any], fresh: dict[str, Any],
                             run: dict[str, Any], corpus: dict[str, Any],
                             index: int, label: str) -> None:
    require(retained.get("valid") is True, f"{label} retained validity changed")
    require(retained.get("corpus") == corpus["id"], f"{label} corpus changed")
    require(retained.get("format") == corpus["format"], f"{label} format changed")
    require(retained.get("edit_admitted") == corpus["expected_edit_admitted"],
            f"{label} edit admission changed")
    require(retained.get("index") == index, f"{label} index changed")
    for key in ("contract", "destination_extension", "destination_sequence",
                "measured_trace_window", "output", "policy", "publication_count"):
        require(retained.get(key) == fresh.get(key), f"{label} {key} differs from fresh replay")

    measured = _trace_transaction_summary(fresh["measured_transaction"], label)
    require(retained.get("measured_transaction") == measured,
            f"{label} measured transaction differs from fresh replay")
    setup = [_trace_transaction_summary(item, f"{label} setup transaction {n}")
             for n, item in enumerate(fresh["setup_transactions"])]
    require(retained.get("setup_transactions") == setup,
            f"{label} setup transactions differ from fresh replay")
    cleanup = fresh["setup_cleanup"]
    require(retained.get("setup_cleanup") == {
        "event_range": cleanup["event_range"],
        "line_range": cleanup["line_range"],
        "paths": [Path(item).name for item in cleanup["paths"]],
        "successful_unlinks": cleanup["successful_unlinks"],
    }, f"{label} setup cleanup differs from fresh replay")
    post = fresh["post_save_scope"]
    require(retained.get("post_save_scope") == {
        "outside_measured_window": post["outside_measured_window"],
        "readback_bytes": post["readback_bytes"],
        "readback_open_line": post["readback_open"]["line"],
        "readback_read_lines": [item["line"] for item in post["readback_reads"]],
        "readback_close_line": post["readback_close"]["line"],
        "destination_unlink_line": post["destination_unlink"]["line"],
    }, f"{label} post-save scope differs from fresh replay")


def _stable_trace_command(command: list[str], run: dict[str, Any],
                          report_path: Path, trace_path: Path) -> list[str]:
    """Retain exact command shape while making packet paths relocatable."""

    output: list[str] = []
    binary_path = resolve_path(run["binary"]["path"], capture_bound=False)
    for token in command:
        try:
            path = resolve_path(token, capture_bound=True)
        except ReplayError:
            output.append(token)
            continue
        if path == trace_path:
            output.append(relative_packet(trace_path))
        elif path == report_path:
            output.append(relative_packet(report_path))
        elif path == binary_path:
            output.append("<native-binary>")
        else:
            output.append(token)
    return output


def _stable_trace_artifact(value: dict[str, Any], path: Path) -> dict[str, Any]:
    return {"path": relative_packet(path),
            "bytes": value.get("bytes", value.get("size")),
            "sha256": value.get("sha256", value.get("digest"))}


def _stable_trace_row(retained: dict[str, Any], run: dict[str, Any],
                      paths: dict[str, Path], command: list[str]) -> dict[str, Any]:
    value = dict(retained)
    value["command"] = _stable_trace_command(command, run, paths["report"], paths["trace"])
    value["receipt"] = {
        "corpus": run["corpus"], "policy": run["policy"],
        "exit_code": run["exit_code"], "command": value["command"],
        "binary": {"bytes": run["binary"].get("bytes"),
                    "sha256": run["binary"].get("sha256")},
        **{key: _stable_trace_artifact(run[key], paths[key])
           for key in ("report", "stderr", "stdout", "trace")},
    }
    value["packet_artifacts"] = {
        key: _stable_trace_artifact(run[key], paths[key])
        for key in ("report", "stderr", "stdout", "trace")
    }
    return value


def check_trace(plan: dict[str, Any], oracle: dict[str, dict[str, Any]],
                build: dict[str, Any], cleanup: bool,
                qualification_identities: dict[str, dict[str, Any]]) -> dict[str, Any]:
    """Replay every retained syscall trace and bind its report identities."""

    parser = _trace_parser()
    trace_directory = PACKET / "trace-0"
    require(trace_directory.is_dir() and not trace_directory.is_symlink(),
            "trace capture directory is missing")
    complete_path = trace_directory / "complete.json"
    runs_path = trace_directory / "runs.json"
    analysis_path = PACKET / "trace-analysis.json"
    complete = read_json(complete_path)
    runs = read_json(runs_path)
    retained = read_json(analysis_path)
    require(isinstance(complete, dict), "trace completion marker is invalid")
    require(isinstance(runs, list), "trace runs manifest is not an array")
    require(isinstance(retained, dict), "trace analysis is not an object")
    require(retained.get("schema") == "litchi-0778-durability-trace-analysis-v1",
            "trace analysis schema changed")
    require(retained.get("children") == plan["expected_children"]["qualification"],
            "trace child count changed")
    require(retained.get("failures") == [], "trace analysis retained failures")
    require(retained.get("plan_sha256") == sha256(PACKET / "plan.json"),
            "trace analysis plan binding changed")
    require(retained.get("runs_sha256") == sha256(runs_path),
            "trace analysis runs binding changed")
    require(retained.get("complete_sha256") == sha256(complete_path),
            "trace analysis completion binding changed")
    require(retained.get("trace_options") == list(parser.TRACE_OPTIONS),
            "trace analysis options changed")
    parser_receipt = retained.get("parser")
    require(parser_receipt == {
        "api": "analyze(trace_path, report, policy, plan=None, corpus=None)",
        "bytes": (PACKET / "trace_analysis.py").stat().st_size,
        "path": "trace_analysis.py",
        "sha256": sha256(PACKET / "trace_analysis.py"),
    }, "trace parser receipt changed")
    require(complete == retained.get("complete"),
            "trace analysis completion differs from trace lane")
    require(complete.get("supported") is True and complete.get("runs") == 28
            and complete.get("source_unchanged") is True
            and complete.get("fixtures_unchanged") is True
            and complete.get("binary_unchanged") is True,
            "trace completion guards are incomplete")
    require(complete.get("plan_sha256") == sha256(PACKET / "plan.json")
            and complete.get("source_sha256") == sha256(PACKET / "source.json"),
            "trace completion source or plan binding changed")
    require(is_sha(complete.get("runner_sha256")),
            "trace completion runner binding is invalid")
    expected = {
        "children": 28, "corpora": 7,
        "policies": list(POLICIES), "samples": 1, "warmup": 0,
        "setup_policy": "full (the harness default save)",
        "publication_sequence": ["published.<format>", "published.<format>",
                                  "published-repeat.<format>", "published.<format>",
                                  "published.<format>"],
        "sync_counts": {
            "default": {"file_sync_calls": 1, "parent_sync_calls": 1},
            "full": {"file_sync_calls": 1, "parent_sync_calls": 1},
            "file-only": {"file_sync_calls": 1, "parent_sync_calls": 0},
            "no-sync": {"file_sync_calls": 0, "parent_sync_calls": 0},
        },
    }
    require(retained.get("expected") == expected, "trace analysis expected contract changed")
    retained_rows = retained.get("rows")
    require(isinstance(retained_rows, list) and len(retained_rows) == len(runs) == 28,
            "trace analysis rows are incomplete")
    corpora = plan_corpora(plan)
    expected_binary = build.get("binaries", {}).get("native")
    require(isinstance(expected_binary, dict), "trace native binary binding is missing")
    fresh_rows: list[dict[str, Any]] = []
    stable_rows: list[dict[str, Any]] = []
    format_counts: dict[str, int] = {}
    observed_sync: dict[str, set[tuple[int, int]]] = {policy: set() for policy in POLICIES}
    window_values: list[int] = []
    refusal_corpora: list[str] = []
    case_order = [(corpus["id"], policy)
                  for corpus in plan["corpora"] for policy in POLICIES]
    require(len(case_order) == len(runs), "trace schedule cardinality changed")
    for index, (run, retained_row, expected_identity) in enumerate(
            zip(runs, retained_rows, case_order)):
        label = f"trace run {index}"
        require(isinstance(run, dict) and isinstance(retained_row, dict),
                f"{label} is invalid")
        corpus_id, policy = expected_identity
        corpus = corpora[corpus_id]
        require(run.get("corpus") == corpus_id and run.get("policy") == policy,
                f"{label} identity differs from frozen trace schedule")
        require(run.get("exit_code") == 0, f"{label} failed")
        finite_number(run.get("started"), f"{label}.started")
        finite_number(run.get("ended"), f"{label}.ended")
        require(run["ended"] >= run["started"], f"{label} ended before started")
        binary_path = _trace_artifact(run.get("binary"), f"{label} binary",
                                      capture_bound=False, allow_missing=cleanup)
        if binary_path is not None:
            require(binary_path == resolve_path(expected_binary["path"], capture_bound=False),
                    f"{label} binary path differs from build")
        _trace_receipt_equal(run.get("binary"), expected_binary, f"{label} binary",
                             capture_bound=False)
        paths: dict[str, Path] = {}
        for key in ("report", "stderr", "stdout", "trace"):
            path = _trace_artifact(run.get(key), f"{label} {key}", capture_bound=True)
            assert path is not None
            paths[key] = path
        for key in ("binary", "report", "stderr", "stdout", "trace"):
            _trace_receipt_equal(run.get(key), retained_row.get("receipt", {}).get(key),
                                 f"{label} retained receipt {key}",
                                 capture_bound=(key != "binary"))
        packet_artifacts = retained_row.get("packet_artifacts")
        require(isinstance(packet_artifacts, dict), f"{label} packet artifacts are missing")
        for key in ("report", "stderr", "stdout", "trace"):
            _trace_receipt_equal(run.get(key), packet_artifacts.get(key),
                                 f"{label} packet artifact {key}")
        command = _check_trace_command(
            run.get("command"), run, plan, corpus, policy, binary_path or
            resolve_path(expected_binary["path"], capture_bound=False),
            paths["trace"], paths["report"], parser, label)
        report = read_json(paths["report"])
        require(isinstance(report, dict), f"{label} report is not an object")
        trace_row = {"case": _trace_case(corpus), "binary": run["binary"]}
        check_report_binary_identity(report, trace_row, label)
        result = check_report_metadata(report, trace_row, plan, 1, 0, label)
        check_elapsed(result.get("elapsed_ns"), 1, label)
        identity = {"corpus_id": corpus_id, "phase": "atomic_publish",
                    "policy": policy, "block": 0, "control": None}
        evidence = corpus_evidence(result, corpus, identity, label)
        oracle_entry = oracle[corpus_id]
        check_oracle_outcome(evidence, oracle_entry, label)
        require(evidence["source_sha"] == oracle_entry["source_sha"],
                f"{label} source archive differs from export oracle")
        require(evidence["corpus_published_sha"] == oracle_entry["published_sha"],
                f"{label} publication differs from export oracle")
        stable_identity = {
            "result_corpus": result.get("corpus"),
            "ordinary_corpus": result["source"]["ordinary_save"].get("corpus"),
        }
        require(stable_identity == qualification_identities[corpus_id],
                f"{label} corpus identity differs from qualification")
        try:
            fresh = parser.analyze(paths["trace"], paths["report"], policy,
                                   plan, corpus)
        except Exception as error:
            fail(f"{label} independent trace replay failed: {error}")
        require(fresh.get("report") == {
            "path": str(paths["report"]), "bytes": run["report"]["bytes"],
            "sha256": run["report"]["sha256"]},
                f"{label} parser report binding changed")
        require(fresh.get("trace", {}).get("sha256") == run["trace"]["sha256"]
                and fresh.get("trace", {}).get("bytes") == run["trace"]["bytes"],
                f"{label} parser trace binding changed")
        _check_trace_row_summary(retained_row, fresh, run, corpus, index, label)
        window_values.append(fresh["measured_trace_window"]["duration_ns"])
        observed_sync[policy].add((fresh["measured_transaction"]["file_sync_calls"],
                                   fresh["measured_transaction"]["parent_sync_calls"]))
        format_counts[corpus["format"]] = format_counts.get(corpus["format"], 0) + 1
        if not corpus["expected_edit_admitted"]:
            refusal_corpora.append(corpus_id)
        fresh_rows.append(fresh)
        stable_rows.append(_stable_trace_row(retained_row, run, paths, command))

    expected_sync = expected["sync_counts"]
    aggregate = {
        "all_28_rows_valid": True,
        "all_atomic_one_replacement": all(
            row["contract"]["all_atomic_one_replacement"] for row in fresh_rows),
        "all_exit_codes_zero": all(run.get("exit_code") == 0 for run in runs),
        "all_measured_destinations_missing": all(
            row["contract"]["measured_destination_missing"] for row in fresh_rows),
        "all_measured_permission_copies_absent": all(
            row["contract"]["measured_permission_copies_absent"] for row in fresh_rows),
        "all_published_output_same": all(
            row["contract"]["published_output_same"] for row in fresh_rows),
        "all_readback_outside_measured_window": all(
            row["post_save_scope"]["outside_measured_window"] for row in fresh_rows),
        "all_report_output_bound": all(row["contract"]["report_output_bound"] for row in fresh_rows),
        "all_setup_sequences_exact": all(row["contract"]["setup_sequence_exact"] for row in fresh_rows),
        "expected_measured_sync_counts": expected_sync,
        "formats": format_counts,
        "observed_measured_sync_counts": {
            policy: [list(item) for item in sorted(observed_sync[policy])]
            for policy in POLICIES
        },
        "policy_sync_counts_match": all(
            observed_sync[policy] == {(value["file_sync_calls"], value["parent_sync_calls"])}
            for policy, value in expected_sync.items()),
        "refusal_corpora": refusal_corpora,
        "window_ns": {
            "diagnostic_only": True, "max": max(window_values),
            "median": statistics.median(window_values), "min": min(window_values), "unit": "ns",
        },
    }
    require(retained.get("aggregate") == aggregate,
            "trace analysis aggregate differs from fresh replay")
    require(retained.get("limits") == [
        "Trace and syscall durations include strace perturbation and are diagnostic only.",
        "No native-time fraction, cache, device, or durability-cause inference is made.",
        "Output SHA-256 is bound to each harness report and repeated publication identities; strace itself does not prove content bytes.",
    ], "trace analysis limits changed")
    return {
        "schema": retained["schema"],
        "children": len(stable_rows),
        "complete": complete,
        "complete_sha256": sha256(complete_path),
        "expected": expected,
        "aggregate": aggregate,
        "limits": retained["limits"],
        "parser": {"api": parser_receipt["api"], "bytes": parser_receipt["bytes"],
                    "path": "trace_analysis.py", "sha256": parser_receipt["sha256"]},
        "plan_sha256": sha256(PACKET / "plan.json"),
        "runs_sha256": sha256(runs_path),
        "trace_options": list(parser.TRACE_OPTIONS),
        "rows": stable_rows,
    }


def integer_stats(values: Iterable[int]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "integer metric vector is empty")
    ordered = sorted(vector)
    return {"count": len(vector), "min": ordered[0],
            "p50": midpoint(ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
            "p95": harness_quantile(vector, 95), "p99": harness_quantile(vector, 99),
            "max": ordered[-1], "mean": statistics.mean(vector)}


def signed_spread(values: Iterable[float]) -> float:
    vector = list(values)
    require(vector, "signed spread vector is empty")
    scale = min((abs(value) for value in vector if value), default=0.0)
    if scale == 0.0:
        return 0.0 if max(vector) == min(vector) else float("inf")
    return (max(vector) - min(vector)) * 100.0 / scale


def check_qualification(plan: dict[str, Any], oracle: dict[str, dict[str, Any]],
                        build: dict[str, Any], cleanup: bool) -> dict[str, Any]:
    qualification_identities = load_qualification_identities(plan)
    directory = find_lane("qualification")
    complete = load_complete(directory, "qualification")
    check_lane_bindings(complete, "qualification")
    rows = load_runs(directory, "qualification")
    require(len(rows) == plan["expected_children"]["qualification"],
            f"qualification child count changed: {len(rows)}")
    check_scheduled_rows(rows, plan, "qualification")
    result_rows = []
    identities_seen: set[tuple[str, str]] = set()
    for index, row in enumerate(rows):
        label = f"qualification run {index}"
        identity = row_identity(row, plan, label, allow_none_block=True)
        key = (identity["corpus_id"], identity["policy"])
        require(key not in identities_seen, f"duplicate qualification identity: {key}")
        identities_seen.add(key)
        require(identity["phase"] in TIMED_PHASES or identity["phase"] in CONTROL_PHASES,
                f"{label} phase is invalid")
        check_row_custody(row, label, build, "native")
        check_fixture_custody(row, plan, label)
        binary = field(row, "binary")
        if binary is not None:
            artifact_receipt(binary, f"{label} binary", capture_bound=False, allow_missing=cleanup)
        artifacts = check_artifacts(row, label)
        report = read_json(artifacts["report"])
        require(isinstance(report, dict), f"{label} report is not an object")
        check_report_binary_identity(report, row, label)
        result = check_report_metadata(report, row, plan, 1, 0, label)
        _, _ = check_elapsed(result.get("elapsed_ns"), 1, label)
        check_operation_alignment(result, 1, label)
        evidence = corpus_evidence(result, plan_corpora(plan)[identity["corpus_id"]], identity, label)
        check_row_outcome_bindings(row, evidence, label)
        check_report_identity(
            result, row, qualification_identities["corpora"][identity["corpus_id"]], label)
        oracle_entry = oracle[identity["corpus_id"]]
        check_oracle_outcome(evidence, oracle_entry, label)
        require(evidence["source_sha"] == oracle_entry["source_sha"],
                f"{label} source archive differs from export oracle")
        require(evidence["corpus_published_sha"] == oracle_entry["published_sha"],
                f"{label} corpus publication differs from export oracle")
        if evidence["published"]:
            require(evidence["published"] == (oracle[identity["corpus_id"]]["published_sha"],),
                    f"{label} published hash differs from oracle")
        result_rows.append({"identity": identity, "report": relative_packet(artifacts["report"]),
                            "report_sha256": sha256(artifacts["report"])})
    for cid in plan_corpora(plan):
        require({policy for corpus_id, policy in identities_seen if corpus_id == cid} == set(POLICIES),
                f"qualification policy set incomplete for {cid}")
    return {"directory": relative_packet(directory), "complete": complete,
            "children": len(result_rows), "rows": result_rows,
            "identities": qualification_identities["corpora"],
            "identities_receipt": qualification_identities}


def _perf_counter_csv(path: Path, label: str) -> dict[str, dict[str, Any]]:
    values: dict[str, dict[str, Any]] = {}
    for line_number, line in enumerate(path.read_text(errors="replace").splitlines(), 1):
        if not line.strip() or line.lstrip().startswith("#"):
            continue
        parts = line.split(";")
        if len(parts) < 3:
            continue
        raw_value, event = parts[0].strip(), parts[2].strip()
        if raw_value in {"<not counted>", "<not supported>", "<not running>"}:
            fail(f"{label}:{line_number} has unavailable counter {event}")
        try:
            value = float(raw_value.replace(",", ""))
        except ValueError:
            continue
        require(math.isfinite(value) and value >= 0, f"{label}:{line_number} has invalid counter")
        runtime_percent = None
        # perf's semicolon output is value;unit;event;time-running;percent-
        # running;... .  The fourth field is an absolute counter interval,
        # so it must not be mistaken for a fraction.
        if len(parts) > 4 and parts[4].strip():
            try:
                runtime_percent = float(parts[4].strip())
            except ValueError:
                runtime_percent = None
        if runtime_percent is not None:
            require(math.isfinite(runtime_percent) and 0.0 < runtime_percent <= 100.0,
                    f"{label}:{line_number} has invalid runtime fraction")
        values[event] = {"value": int(value) if value.is_integer() else value,
                         "runtime_percent": runtime_percent}
    return values


def diagnostics_summary(plan: dict[str, Any], build: dict[str, Any], cleanup: bool) -> dict[str, Any] | None:
    directory = PACKET / "instructions-0"
    complete_path = directory / "complete.json"
    if not complete_path.is_file():
        require(not directory.exists(), "instruction diagnostics directory is incomplete")
        return None
    complete = read_json(complete_path)
    require(isinstance(complete, dict), "instruction diagnostics completion is invalid")
    if complete.get("supported") is False:
        return {"supported": False, "reason": complete.get("reason", "unsupported"),
                "no_elapsed_division": True}
    require(complete.get("supported") is True, "instruction diagnostics support is unspecified")
    for key, source_path in (("runner_sha256", PACKET / "diagnostics.py"),
                             ("plan_sha256", PACKET / "plan.json"),
                             ("source_sha256", PACKET / "source.json")):
        if key in complete:
            require(complete[key] == sha256(source_path),
                    f"instruction diagnostics {key} changed")
    expected = plan.get("diagnostics", {}).get("instructions", {})
    events = expected.get("events", ["instructions:u", "cycles:u", "branches:u", "branch-misses:u"])
    require(isinstance(events, list) and events, "instruction event list is missing")
    path = directory / "runs.json"
    if not path.is_file():
        path = directory / "records.json"
    require(path.is_file(), "instruction diagnostics records are missing")
    raw = read_json(path)
    rows = (raw.get("rows", raw.get("records")) if isinstance(raw, dict) else raw)
    require(isinstance(rows, list), "instruction diagnostics records are not an array")
    require(len(rows) == expected.get("processes"), "instruction diagnostics process count changed")
    parsed: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"instruction diagnostic row {index} is invalid")
        require(field(row, "exit", "exit_code") == 0,
                f"instruction diagnostic row {index} failed")
        binary = row.get("binary")
        if binary is not None:
            require(binary == build.get("binaries", {}).get("native"),
                    f"instruction diagnostic row {index} binary changed")
            artifact_receipt(binary, f"instruction diagnostic {index} binary",
                             capture_bound=False, allow_missing=cleanup)
        csv_value = row.get("csv")
        csv_path = artifact_receipt(csv_value, f"instruction diagnostic {index} CSV", capture_bound=True)
        assert csv_path is not None
        counters = _perf_counter_csv(csv_path, f"instruction diagnostic {index}")
        require(set(events) <= set(counters), f"instruction diagnostic {index} misses an event")
        sample_count = row.get("samples")
        require(sample_count in {3, 23}, f"instruction diagnostic {index} sample count changed")
        corpus_id = row.get("corpus", row.get("corpus_id"))
        policy = row.get("policy", row.get("durability"))
        require(corpus_id in expected.get("corpora", []),
                f"instruction diagnostic {index} corpus changed")
        require(policy in expected.get("policies", []),
                f"instruction diagnostic {index} policy changed")
        require(isinstance(row.get("repeat"), int) and
                0 <= row["repeat"] < expected.get("repeats"),
                f"instruction diagnostic {index} repeat changed")
        parsed.append({"corpus_id": corpus_id, "policy": policy,
                       "repeat": row.get("repeat"), "samples": sample_count,
                       "counters": counters, "csv": relative_packet(csv_path)})
    grouped: dict[tuple[str, str, int], dict[int, dict[str, Any]]] = {}
    for row in parsed:
        key = (row["corpus_id"], row["policy"], row["repeat"])
        grouped.setdefault(key, {})[row["samples"]] = row
    marginal: list[dict[str, Any]] = []
    for key, by_samples in sorted(grouped.items()):
        require(set(by_samples) == {3, 23}, f"instruction pair incomplete: {key}")
        events_out: dict[str, Any] = {}
        for event in events:
            low = by_samples[3]["counters"][event]
            high = by_samples[23]["counters"][event]
            events_out[event] = {
                "samples_3": low["value"], "samples_23": high["value"],
                "marginal_per_extra_sample": (high["value"] - low["value"]) / 20.0,
                "runtime_percent_samples_3": low["runtime_percent"],
                "runtime_percent_samples_23": high["runtime_percent"],
            }
            finite_number(events_out[event]["marginal_per_extra_sample"],
                          f"instruction marginal {key} {event}")
        marginal.append({"corpus_id": key[0], "policy": key[1], "repeat": key[2],
                         "events": events_out,
                         "scope": "whole-process counters; marginal (23-3)/20 includes report work"})
    require(len(marginal) == expected.get("processes", 32) // 2,
            "instruction marginal pair count changed")
    return {"supported": True, "processes": len(parsed), "events": events,
            "marginal_pairs": marginal,
            "scope": expected.get("scope", "whole-process counters; not timed-region attribution"),
            "no_elapsed_division": True}


def analyze() -> dict[str, Any]:
    plan = load_plan()
    quality = check_quality()
    build, cleanup = load_build_and_cleanup()
    oracle = load_oracle(plan, build, cleanup)
    oracle_controls = oracle.pop("__oracle_controls__")
    qualification = check_qualification(plan, oracle, build, cleanup)
    native = check_native(plan, oracle, build, cleanup, qualification["identities"])
    allocation = check_allocation(plan, oracle, build, cleanup, qualification["identities"])
    trace = check_trace(plan, oracle, build, cleanup, qualification["identities"])
    source = load_source_custody(plan)
    diagnostics = diagnostics_summary(plan, build, cleanup)
    output = {
        "schema": "litchi-0778-durability-analysis-v1",
        "plan_schema": plan["schema"],
        "base": plan["base"],
        "performance_claim": "descriptive current-source policy comparison; no time-causal or physical durability claim",
        "quality": quality,
        "source_manifest": {"entries": len(source), "sha256": sha256(PACKET / "source.json")},
        "native": {"children": native["children"], "controls": native["controls"],
                   "analysis": native_analysis(native)},
        "allocation": {"children": allocation["children"],
                       "analysis": allocation_analysis(allocation)},
        "qualification": {"children": qualification["children"],
                          "identities": qualification["identities_receipt"]},
        "trace": trace,
        "counter_diagnostics": diagnostics,
        "verification": {
            "native_child_receipts": native["children"],
            "allocation_child_receipts": allocation["children"],
            "qualification_child_receipts": qualification["children"],
            "publication_hashes_bound_to_export_oracle": True,
            "refusal_byte_exact_source_checked": True,
            "whole_process_rss_separate": True,
            "allocation_elapsed_not_mixed": True,
            "policy_comparisons_paired_by_block": True,
            "source_binary_custody_checked": True,
            "oracle_bounded_controls": oracle_controls,
            # This is a stable custody contract.  The current cleanup state
            # is intentionally not serialized here: after a reviewer removes
            # the retained binaries, replay must produce the same analysis
            # while checking cleanup.json as the evidence for their absence.
            "cleanup_binary_witness_required_when_missing": True,
            "trace_child_receipts": trace["children"],
            "trace_commands_and_report_identities_replayed": True,
            "counter_marginal_scope_includes_report_work":
                diagnostics is None or diagnostics.get("no_elapsed_division") is True,
        },
        "limits": [
            "Durability levels describe distinct crash/power-loss contracts; timing comparisons are descriptive.",
            "Process RSS includes startup, parsing, filesystem and report work.",
            "Allocation metrics cover the harness allocation region and are not elapsed latency.",
            "No physical cold-cache, device-floor, or time-causal claim is made.",
        ],
    }
    # Retain raw child identities and hashes so a reviewer can trace every
    # summarized process without loading the raw reports through a second tool.
    output["native"]["receipts"] = [
        {"corpus_id": item["identity"]["corpus_id"], "phase": item["identity"]["phase"],
         "policy": item["identity"]["policy"], "block": item["identity"]["block"],
         "report": item["report"], "report_sha256": item["report_sha256"]}
        for item in native["rows"]
    ]
    output["allocation"]["receipts"] = [
        {"corpus_id": item["identity"]["corpus_id"], "phase": item["identity"]["phase"],
         "policy": item["identity"]["policy"], "block": item["identity"]["block"],
         "report": item["report"], "report_sha256": item["report_sha256"]}
        for item in allocation["rows"]
    ]
    output["qualification"]["receipts"] = qualification["rows"]
    return output


if __name__ == "__main__":
    result = analyze()
    output_path = PACKET / "analysis.json"
    require(not output_path.exists(), "refusing to overwrite analysis.json")
    output_path.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"native": result["native"]["children"],
                      "allocation": result["allocation"]["children"],
                      "qualification": result["qualification"]["children"]}, sort_keys=True))
