#!/usr/bin/env python3
"""Capture the conditional eager XLSX read controls for the 0542 pilot.

This lane is deliberately separate from the source-backed primary benchmark.
It exercises the existing ``litchi-perf-baseline`` binary with the ordinary
``xlsx_open_owned`` and ``xlsx_first_cell`` cases.  The source manifests and
retained normal binaries are the ones already frozen by ``run.py``; no eager
specific build is performed.  Captures use the same ``run.run`` custody and
the same retained-baseline source protocol as the primary ABBA lanes.

The plan is created only after both stage manifests and normal binary
identities exist.  It records every job, binary/source digest, and driver
digest before the first eager child is launched.  ``capture`` then enforces
the frozen baseline/candidate/candidate/baseline order.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import sys
from typing import Any

# The campaign invokes this file with ``python3 -B``.  Keep imports safe for
# analyzer replay as well, so importing the runner never leaves a cache file
# beside the sealed evidence.
sys.dont_write_bytecode = True
sys.path.insert(0, str(Path(__file__).resolve().parent))

import run  # noqa: E402  (the sibling campaign runner is the custody authority)


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
EAGER_PLAN_PATH = HERE / "eager-plan.json"
RUN_PATH = HERE / "run.py"
ANALYZER_PATH = HERE / "analyze_eager.py"
SCRIPT_PATH = Path(__file__).resolve()

SCHEMA = "litchi-0542-eager-guard-plan-v1"
BINDING_SCHEMA = "litchi-0542-eager-guard-binding-v1"
CASES = ("xlsx_open_owned", "xlsx_first_cell")
SHAPES = ("medium", "dense-wide")
REPEATS = (1, 2)
NATIVE_ORDER = ("baseline-r1", "candidate-r1", "candidate-r2", "baseline-r2")
WARMUP = 10
SAMPLES = 30
CPU = 2
ADVERSE_THRESHOLD_PERCENT = 5.0
TIME_FORMAT = (
    '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,'
    '"system_seconds":%S}'
)


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory eager-lane artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    try:
        path.write_text(
            json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8"
        )
    except OSError as error:
        raise EvidenceError(f"cannot write JSON {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside eager evidence: {path}") from error


def digest(value: Any, label: str, length: int = 64) -> str:
    require(
        isinstance(value, str)
        and re.fullmatch(r"[0-9a-f]{%d}" % length, value) is not None,
        f"{label} is not a lowercase SHA-{length * 4} digest",
    )
    return value


def nonnegative_integer(value: Any, label: str) -> int:
    require(
        isinstance(value, int) and not isinstance(value, bool) and value >= 0,
        f"{label} is not a nonnegative integer",
    )
    return value


def positive_integer(value: Any, label: str) -> int:
    result = nonnegative_integer(value, label)
    require(result > 0, f"{label} is not positive")
    return result


def finite_number(value: Any, label: str) -> float:
    require(
        isinstance(value, (int, float))
        and not isinstance(value, bool)
        and math.isfinite(float(value)),
        f"{label} is not finite",
    )
    return float(value)


def nonempty_string(value: Any, label: str) -> str:
    require(isinstance(value, str) and value, f"{label} is not a nonempty string")
    return value


def primary_plan() -> dict[str, Any]:
    primary = read_json(PLAN_PATH)
    require(isinstance(primary, dict), "primary plan is not an object")
    digest(primary.get("revision"), "primary plan revision", 40)
    require(
        isinstance(primary.get("status"), str)
        and primary["status"].startswith("frozen-before"),
        "primary plan is not frozen before capture",
    )
    require(primary.get("cpu") == CPU, "primary plan CPU is not 2")
    require(primary.get("owned_paths") == [str(run.TARGET)],
            "primary plan owned target differs")
    primary_case = primary.get("primary")
    require(isinstance(primary_case, dict), "primary plan primary lane is missing")
    require(
        primary_case.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save",
        "primary plan primary case differs",
    )
    conditional = primary.get("conditional_lanes")
    require(isinstance(conditional, dict) and isinstance(conditional.get("eager"), str),
            "primary plan eager conditional lane is missing")
    return primary


def stage_manifest(stage: str) -> tuple[dict[str, str], str]:
    require(stage in ("baseline", "candidate"), f"unsupported stage {stage}")
    path = HERE / stage / "source-manifest.json"
    require(path.is_file() and not path.is_symlink(),
            f"{stage} source manifest is missing")
    value = read_json(path)
    require(isinstance(value, dict) and value, f"{stage} source manifest is empty")
    for name, value_digest in value.items():
        nonempty_string(name, f"{stage} source path")
        digest(value_digest, f"{stage} source digest")
    return value, sha256(path)


def binary_metadata(stage: str, *, require_live: bool = False) -> dict[str, Any]:
    path = HERE / stage / "binary-normal.json"
    value = read_json(path)
    require(isinstance(value, dict), f"{stage} normal binary identity is not an object")
    expected_path = run.SCRATCH / f"{stage}-normal"
    binary_path = Path(nonempty_string(value.get("path"), f"{stage} binary path"))
    require(binary_path == expected_path,
            f"{stage} binary path differs from retained normal path")
    binary_digest = digest(value.get("sha256"), f"{stage} binary SHA")
    byte_count = positive_integer(value.get("bytes"), f"{stage} binary bytes")
    require(not binary_path.is_symlink(), f"{stage} binary is a symlink")
    if binary_path.exists():
        require(binary_path.is_file(), f"{stage} binary is not a regular file")
        require(sha256(binary_path) == binary_digest,
                f"{stage} binary digest differs from identity")
        require(binary_path.stat().st_size == byte_count,
                f"{stage} binary byte count differs from identity")
    elif require_live:
        raise EvidenceError(f"{stage} retained normal binary is missing")
    return {
        "path": str(binary_path),
        "sha256": binary_digest,
        "bytes": byte_count,
        "build_receipt_sha256": digest(
            value.get("build_receipt_sha256"), f"{stage} binary build receipt SHA"
        ),
        "source_manifest_sha256": digest(
            value.get("source_manifest_sha256"),
            f"{stage} binary source manifest SHA",
        ),
    }


def group_key(stage: str, repeat: int) -> str:
    require(stage in ("baseline", "candidate"), f"unsupported stage {stage}")
    require(repeat in REPEATS, f"unsupported eager repeat {repeat}")
    return f"{stage}-r{repeat}"


def source_protocol(stage: str) -> str:
    if stage == "baseline":
        return "retained baseline binary under candidate source checkout"
    if stage == "candidate":
        return "candidate binary under candidate source checkout"
    raise EvidenceError(f"unsupported eager stage {stage}")


def jobs_for(stage: str, repeat: int, *, plan: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    """Return the exact case/shape order for one ABBA group."""

    if plan is not None:
        cases = tuple(plan.get("cases", ()))
        shapes = tuple(plan.get("shapes", ()))
        warmup = plan.get("warmup")
        samples = plan.get("samples")
    else:
        cases, shapes, warmup, samples = CASES, SHAPES, WARMUP, SAMPLES
    require(cases == CASES and shapes == SHAPES,
            "eager case or shape matrix differs from frozen constants")
    require(warmup == WARMUP and samples == SAMPLES,
            "eager iteration counts differ from frozen constants")
    group = group_key(stage, repeat)
    retained = stage == "baseline"
    return [
        {
            "name": f"eager-r{repeat}-{case}-{shape}",
            "stage": stage,
            "repeat": repeat,
            "group": group,
            "case": case,
            "shape": shape,
            "warmup": warmup,
            "samples": samples,
            "binary_stage": stage,
            "retained_baseline": retained,
            "source_checkout": source_protocol(stage),
        }
        for case in cases
        for shape in shapes
    ]


def expected_jobs() -> list[dict[str, Any]]:
    return [
        job
        for group in NATIVE_ORDER
        for stage, repeat in (("baseline", 1), ("candidate", 1),
                              ("candidate", 2), ("baseline", 2))
        if group == group_key(stage, repeat)
        for job in jobs_for(stage, repeat)
    ]


def expected_binary_hashes() -> dict[str, str]:
    return {stage: binary_metadata(stage)["sha256"] for stage in ("baseline", "candidate")}


def expected_manifest_hashes() -> dict[str, str]:
    return {stage: stage_manifest(stage)[1] for stage in ("baseline", "candidate")}


def create_plan() -> dict[str, Any]:
    """Freeze the eager lane after both normal binaries are retained.

    Root calls this once immediately before the first eager capture.  If a
    plan already exists, it is treated as immutable and is validated rather
    than rewritten.
    """

    primary = primary_plan()
    if EAGER_PLAN_PATH.exists():
        value = read_json(EAGER_PLAN_PATH)
        validate_plan(value, primary)
        return value

    require(ANALYZER_PATH.is_file() and not ANALYZER_PATH.is_symlink(),
            f"missing eager analyzer: {ANALYZER_PATH}")
    # The candidate manifest is required because every retained-baseline
    # child is intentionally run under the candidate checkout.
    stage_manifest("baseline")
    stage_manifest("candidate")
    binary_hashes = expected_binary_hashes()
    manifest_hashes = expected_manifest_hashes()
    value: dict[str, Any] = {
        "schema": SCHEMA,
        "primary_plan_sha256": sha256(PLAN_PATH),
        "run_script_sha256": sha256(RUN_PATH),
        "eager_run_script_sha256": sha256(SCRIPT_PATH),
        "analyzer_script_sha256": sha256(ANALYZER_PATH),
        "case_matrix": list(CASES),
        "shape_matrix": list(SHAPES),
        # Short aliases make the frozen values easy to consume in shell
        # receipts while the explicit matrix above remains the job contract.
        "cases": list(CASES),
        "shapes": list(SHAPES),
        "repeats": len(REPEATS),
        "warmup": WARMUP,
        "samples": SAMPLES,
        "cpu": CPU,
        "native_order": list(NATIVE_ORDER),
        "stage_source_protocol": {
            group_key(stage, repeat): source_protocol(stage)
            for stage, repeat in (("baseline", 1), ("candidate", 1),
                                  ("candidate", 2), ("baseline", 2))
        },
        "binary_sha256": binary_hashes,
        "source_manifest_sha256": manifest_hashes,
        "jobs": expected_jobs(),
        "adverse_threshold_percent": ADVERSE_THRESHOLD_PERCENT,
        "no_speedup_requirement": True,
        "purpose": (
            "Conditional ordinary eager XLSX read controls for the 0542 raw "
            "worksheet codec extraction; this lane supplies no speedup claim."
        ),
        "gate": (
            "No speedup threshold. Validate exact report/corpus identity, "
            "elapsed p50/p95/p99/mean, whole-child peak RSS, ABBA custody, "
            "and individually review every adverse or repeat-drift change "
            "over 5%."
        ),
        "result_identity_contract": {
            "source": "ordinary eager read cases omit result.source",
            "output": "ordinary read controls omit result.output_sha256",
            "sink": "ordinary read controls serialize sink as null",
            "phases": "ordinary read controls omit operation_metrics and phase evidence",
        },
        "status": "frozen-before-capture",
    }
    write_json(EAGER_PLAN_PATH, value)
    frozen = read_json(EAGER_PLAN_PATH)
    validate_plan(frozen, primary)
    return frozen


def validate_plan(value: Any, primary: dict[str, Any] | None = None) -> dict[str, Any]:
    """Validate an eager plan and all hashes that can be checked pre-capture."""

    require(isinstance(value, dict), "eager plan is not an object")
    require(value.get("schema") == SCHEMA, "eager plan schema differs")
    if primary is None:
        primary = primary_plan()
    require(value.get("primary_plan_sha256") == sha256(PLAN_PATH),
            "eager plan is bound to a different primary plan")
    require(value.get("run_script_sha256") == sha256(RUN_PATH),
            "eager plan is bound to a different run script")
    require(value.get("eager_run_script_sha256") == sha256(SCRIPT_PATH),
            "eager plan is bound to a different eager runner")
    require(ANALYZER_PATH.is_file() and not ANALYZER_PATH.is_symlink(),
            "eager analyzer is missing")
    require(value.get("analyzer_script_sha256") == sha256(ANALYZER_PATH),
            "eager plan is bound to a different eager analyzer")
    require(value.get("cases") == list(CASES)
            and value.get("case_matrix") == list(CASES),
            "eager case matrix differs")
    require(value.get("shapes") == list(SHAPES)
            and value.get("shape_matrix") == list(SHAPES),
            "eager shape matrix differs")
    require(value.get("repeats") == len(REPEATS)
            and value.get("warmup") == WARMUP
            and value.get("samples") == SAMPLES,
            "eager repeat/warmup/sample counts differ")
    require(value.get("cpu") == primary.get("cpu") == CPU,
            "eager CPU differs from primary plan")
    require(value.get("native_order") == list(NATIVE_ORDER),
            "eager ABBA order differs")
    protocol = value.get("stage_source_protocol")
    expected_protocol = {
        group_key(stage, repeat): source_protocol(stage)
        for stage, repeat in (("baseline", 1), ("candidate", 1),
                              ("candidate", 2), ("baseline", 2))
    }
    require(protocol == expected_protocol, "eager source protocol differs")
    require(value.get("no_speedup_requirement") is True,
            "eager lane unexpectedly declares a speedup gate")
    threshold = finite_number(value.get("adverse_threshold_percent"),
                              "eager adverse threshold")
    require(threshold == ADVERSE_THRESHOLD_PERCENT,
            "eager adverse threshold differs")
    binary_hashes = value.get("binary_sha256")
    manifest_hashes = value.get("source_manifest_sha256")
    require(isinstance(binary_hashes, dict)
            and set(binary_hashes) == {"baseline", "candidate"},
            "eager binary hash inventory differs")
    require(isinstance(manifest_hashes, dict)
            and set(manifest_hashes) == {"baseline", "candidate"},
            "eager source hash inventory differs")
    for stage in ("baseline", "candidate"):
        digest(binary_hashes[stage], f"eager {stage} binary SHA")
        digest(manifest_hashes[stage], f"eager {stage} manifest SHA")
        identity = binary_metadata(stage)
        require(identity["sha256"] == binary_hashes[stage],
                f"eager {stage} binary hash changed")
        _, manifest_sha = stage_manifest(stage)
        require(manifest_sha == manifest_hashes[stage],
                f"eager {stage} source manifest hash changed")
    jobs = value.get("jobs")
    require(isinstance(jobs, list), "eager job inventory is not a list")
    require(jobs == expected_jobs(), "eager job inventory differs")
    return value


def plans() -> tuple[dict[str, Any], dict[str, Any]]:
    primary = primary_plan()
    eager = read_json(EAGER_PLAN_PATH)
    return primary, validate_plan(eager, primary)


def expected_capture_command(binary: Path, job: dict[str, Any], eager: dict[str, Any]) -> list[str]:
    stage_dir = HERE / job["stage"]
    name = job["name"]
    return [
        "taskset", "-c", str(eager["cpu"]),
        "/usr/bin/time", "-f", TIME_FORMAT,
        "-o", str(stage_dir / f"{name}.rss.json"),
        str(binary),
        "--warmup", str(job["warmup"]),
        "--samples", str(job["samples"]),
        "--case", job["case"],
        "--xlsx-shape", job["shape"],
        "--json", str(stage_dir / f"{name}.json"),
    ]


def group_jobs(stage: str, repeat: int, eager: dict[str, Any]) -> list[dict[str, Any]]:
    expected = [job for job in eager["jobs"]
                if job["stage"] == stage and job["repeat"] == repeat]
    require(expected == jobs_for(stage, repeat, plan=eager),
            f"frozen eager jobs for {stage} repeat {repeat} differ")
    return expected


def assert_capture_slot_is_next(stage: str, repeat: int, eager: dict[str, Any]) -> None:
    target = group_key(stage, repeat)
    index = list(NATIVE_ORDER).index(target)
    for previous in NATIVE_ORDER[:index]:
        previous_stage, previous_repeat = previous.rsplit("-r", 1)
        for job in group_jobs(previous_stage, int(previous_repeat), eager):
            require((HERE / previous_stage / f"{job['name']}.receipt.json").is_file(),
                    f"capture prerequisite is incomplete: {previous}/{job['name']}")
    for later in NATIVE_ORDER[index + 1:]:
        later_stage, later_repeat = later.rsplit("-r", 1)
        for job in group_jobs(later_stage, int(later_repeat), eager):
            require(not (HERE / later_stage / f"{job['name']}.receipt.json").exists(),
                    f"capture order already advanced past {target}")


def write_binding(stage: str, repeat: int, job: dict[str, Any], eager: dict[str, Any],
                  binary: dict[str, Any], source_manifest_sha: str,
                  working_manifest_sha: str, receipt_path: Path,
                  command: list[str]) -> None:
    stage_dir = HERE / stage
    binding = {
        "schema": BINDING_SCHEMA,
        "stage": stage,
        "repeat": repeat,
        "group": job["group"],
        "name": job["name"],
        "case": job["case"],
        "shape": job["shape"],
        "source_checkout": job["source_checkout"],
        "retained_baseline": job["retained_baseline"],
        "primary_plan_sha256": sha256(PLAN_PATH),
        "eager_plan_sha256": sha256(EAGER_PLAN_PATH),
        "run_script_sha256": sha256(RUN_PATH),
        "eager_run_script_sha256": sha256(SCRIPT_PATH),
        "analyzer_script_sha256": sha256(ANALYZER_PATH),
        "source_manifest_sha256": source_manifest_sha,
        "working_source_manifest_sha256": working_manifest_sha,
        "binary_sha256": binary["sha256"],
        "receipt_sha256": sha256(receipt_path),
        "command_sha256": hashlib.sha256(
            json.dumps(command, separators=(",", ":")).encode()
        ).hexdigest(),
        "output": relative(stage_dir / f"{job['name']}.json"),
        "rss": relative(stage_dir / f"{job['name']}.rss.json"),
    }
    write_json(stage_dir / f"{job['name']}.binding.json", binding)


def capture(stage: str, repeat: int) -> None:
    primary, eager = plans()
    require(stage in ("baseline", "candidate"), f"unsupported stage {stage}")
    require(repeat in REPEATS, f"unsupported repeat {repeat}")
    assert_capture_slot_is_next(stage, repeat, eager)
    jobs = group_jobs(stage, repeat, eager)
    identity = binary_metadata(stage, require_live=True)
    binary = Path(identity["path"])
    source_value, source_manifest_sha = stage_manifest(stage)
    del source_value
    _, working_manifest_sha = stage_manifest("candidate")
    # Retained baseline children intentionally validate and run under the
    # candidate manifest via run.check_source(stage, retained_baseline=True).
    if stage == "baseline":
        require(working_manifest_sha == sha256(HERE / "candidate" / "source-manifest.json"),
                "candidate working manifest is not available for retained baseline")
    for job in jobs:
        name = job["name"]
        stage_dir = HERE / stage
        for suffix in (".json", ".stdout", ".stderr", ".rss.json",
                       ".receipt.json", ".binding.json"):
            require(not (stage_dir / f"{name}{suffix}").exists(),
                    f"eager artifact already exists: {stage}/{name}{suffix}")
        command = expected_capture_command(binary, job, eager)
        run.run(stage, name, command, binary,
                retained_baseline=(stage == "baseline"))
        receipt_path = stage_dir / f"{name}.receipt.json"
        receipt = read_json(receipt_path)
        require(receipt.get("working_source_manifest_sha256") == working_manifest_sha,
                f"{name} working source manifest does not follow eager protocol")
        write_binding(stage, repeat, job, eager, identity, source_manifest_sha,
                      working_manifest_sha, receipt_path, command)
    print(f"0542 eager capture passed: {stage} repeat {repeat}", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--create-plan", action="store_true",
                       help="freeze eager-plan.json after normal binaries exist")
    modes.add_argument("--capture", action="store_true",
                       help="capture one frozen eager ABBA group")
    parser.add_argument("--stage", choices=("baseline", "candidate"))
    parser.add_argument("--repeat", type=int, choices=REPEATS)
    args = parser.parse_args()
    try:
        if args.create_plan:
            require(args.stage is None and args.repeat is None,
                    "--create-plan does not take --stage or --repeat")
            create_plan()
            print(f"0542 eager plan frozen: {EAGER_PLAN_PATH}")
        else:
            require(args.stage is not None and args.repeat is not None,
                    "--capture requires --stage and --repeat")
            capture(args.stage, args.repeat)
    except (EvidenceError, AssertionError, OSError, json.JSONDecodeError) as error:
        print(f"eager_run.py: error: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
