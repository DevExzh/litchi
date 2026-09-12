#!/usr/bin/env python3
"""Validate and compare the 0522 XLSX scanner Callgrind profiles.

The capture layout and profile owner are the same as 0521: two repeats, two
synthetic numeric shapes, and four Callgrind dumps per child.  The fourth dump
contains one timed ``MultiSourceEdit::commit`` call.  This report checks the
raw lifecycle/direct-runner edges, preserves the final zero-Ir process dump,
and emits deterministic inclusive/self annotations for the timed dump.

The 0521 analyzer supplies the retained receipt, native-parity, annotation,
and 0519 raw-edge machinery.  Its module is loaded by path and its evidence
root is rebound to this directory; neither earlier analyzer nor evidence is
modified.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
PRIOR_ANALYZER = HERE.parent / "change-0521" / "analyze_profiles.py"

OWNER = "litchi_xlsx::cell_values::source::MultiSourceEdit::commit"
LIFECYCLE_PARENT = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
DIRECT_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
SCANNER_START_CELL = (
    "litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::Scanner::start_cell"
)
WIRE_CELL_TAG = "litchi_xlsx::raw::worksheet::edit::codec::wire::cell_tag"
VALIDATOR = "litchi_xlsx::cell_values::validation::validate_xml"
EVENT_INTO_OWNED = "quick_xml::events::Event::into_owned"
PARTS = (1, 2, 3, 4)


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    import hashlib

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


def load_prior() -> Any:
    require(PRIOR_ANALYZER.is_file(), f"missing retained helper {PRIOR_ANALYZER}")
    spec = importlib.util.spec_from_file_location(
        "litchi_0521_profile_helpers", PRIOR_ANALYZER
    )
    require(spec is not None and spec.loader is not None,
            f"cannot load retained helper {PRIOR_ANALYZER}")
    prior = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(prior)
    # All helper functions resolve these names in their defining module.  The
    # 0521 parser is intentionally reused, but all paths and plan reads must
    # point at this fresh 0522 bundle.
    prior.HERE = HERE
    prior.PLAN_PATH = PLAN_PATH
    return prior


PRIOR = load_prior()


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def plan_data() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "plan profile is not an object")
    require(profile.get("owner") == OWNER, "plan profile owner differs")
    require(profile.get("shapes") == ["medium", "dense-sparse"],
            "plan profile shape matrix differs")
    require(profile.get("repeats") == 2, "plan profile repeat count differs")
    require(profile.get("warmup") == 0, "plan profile warmup differs")
    require(profile.get("samples") == 1, "plan profile sample count differs")
    return plan


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    return [
        {
            "name": f"profile-r{repeat}-{shape}",
            "repeat": repeat,
            "shape": shape,
        }
        for repeat in range(1, profile["repeats"] + 1)
        for shape in profile["shapes"]
    ]


def edge_summary(path: Path, target: str, caller: str) -> dict[str, Any]:
    summary = PRIOR.target_edge_summary(path, target, caller)
    return {
        "target": target,
        "caller": caller,
        "positive_edge_count": summary["positive_edge_count"],
        "calls": summary["calls"],
        "inclusive_ir": summary["inclusive_ir"],
        "edges": summary["edges"],
        "present": summary["positive_edge_count"] > 0,
    }


def validate_scanner_edge(path: Path, stage: str) -> dict[str, Any]:
    edge = edge_summary(path, WIRE_CELL_TAG, SCANNER_START_CELL)
    if stage == "baseline":
        require(edge["positive_edge_count"] == 1,
                f"{relative(path)}: baseline scanner -> cell_tag edge missing or duplicated")
        policy = "present"
    else:
        require(edge["positive_edge_count"] == 0,
                f"{relative(path)}: candidate retained scanner -> cell_tag edge")
        policy = "absent"
    return {
        **edge,
        "policy": policy,
        "interpretation": (
            "The baseline calls wire::cell_tag from Scanner::start_cell; the "
            "candidate policy requires that positive edge to be absent after "
            "the checked unprefixed-cell proof."
        ),
    }


def validate_event_edge(path: Path, stage: str) -> dict[str, Any]:
    edge = edge_summary(path, EVENT_INTO_OWNED, VALIDATOR)
    # 0521 already removed this edge in both binaries.  Keep the assertion
    # explicit so 0522 cannot accidentally attribute that older change to the
    # scanner optimization, while avoiding the obsolete baseline-present /
    # candidate-absent policy from the pre-0521 analyzer.
    require(edge["positive_edge_count"] == 0,
            f"{relative(path)}: validate_xml -> Event::into_owned edge reappeared")
    return {
        **edge,
        "policy": "absent-both-stages",
        "stage": stage,
        "interpretation": (
            "This edge is retained only as a diagnostic invariant from 0521; "
            "it is not a 0522 improvement claim."
        ),
    }


def validate_profile(stage: str, plan: dict[str, Any], metadata: dict[str, Any],
                     job: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    name = job["name"]
    repeat = job["repeat"]
    shape = job["shape"]
    receipt_path = stage_dir / f"{name}.receipt.json"
    profile_path = stage_dir / f"{name}.json"
    require(receipt_path.is_file(), f"{relative(receipt_path)}: receipt is missing")
    require(profile_path.is_file(), f"{relative(profile_path)}: result is missing")
    receipt = read_json(receipt_path)
    PRIOR.validate_profile_receipt(stage_dir, name, receipt, metadata, plan, shape)
    _, profile_identity = PRIOR.validate_result_report(
        profile_path, metadata["binary_sha256"], plan, repeat, shape, True
    )

    native_name = f"native-r{repeat}-primary-{shape}"
    native_path = stage_dir / f"{native_name}.json"
    require(native_path.is_file(), f"{relative(native_path)}: native result is missing")
    native_metadata = PRIOR.validate_native_receipt(
        stage_dir, native_name, metadata, plan
    )
    _, native_identity = PRIOR.validate_result_report(
        native_path, metadata["binary_sha256"], plan, repeat, shape, False
    )
    require(profile_identity == native_identity,
            f"{name}: profile/native corpus, output, or counter identity differs")

    raw_dumps = []
    for part in PARTS:
        path = stage_dir / f"{name}.callgrind.{part}"
        require(path.is_file(), f"{relative(path)}: numbered raw dump is missing")
        parent = LIFECYCLE_PARENT if part < 4 else DIRECT_PARENT
        raw_dumps.append(PRIOR.validate_numbered_dump(path, part, parent))
    final_path = stage_dir / f"{name}.callgrind"
    require(final_path.is_file(), f"{relative(final_path)}: final raw dump is missing")
    final_dump = PRIOR.validate_final_dump(final_path)

    timed_path = stage_dir / f"{name}.callgrind.4"
    annotations = PRIOR.annotate_profile(
        stage, stage_dir, name, timed_path, raw_dumps[-1]["summary_ir"]
    )
    scanner_edge = validate_scanner_edge(timed_path, stage)
    event_edge = validate_event_edge(timed_path, stage)
    validator_incoming = PRIOR.target_edge_summary(timed_path, VALIDATOR, None)

    return {
        "name": name,
        "repeat": repeat,
        "shape": shape,
        "receipt": relative(receipt_path),
        "receipt_sha256": sha256(receipt_path),
        "profile_result": relative(profile_path),
        "profile_result_sha256": sha256(profile_path),
        "native": native_metadata,
        "native_result_identity": native_identity,
        "raw_dumps": raw_dumps,
        "final_process_dump": final_dump,
        "annotations": annotations,
        "validator_raw_incoming": validator_incoming,
        "scanner_start_cell_to_cell_tag": scanner_edge,
        "validator_event_into_owned_edge": event_edge,
        "total_ir": annotations["owner"]["inclusive_ir"],
        "self_ir": annotations["owner"]["self_ir"],
        "direct_ir": sum(annotations["owner"]["direct_callee_ir"].values()),
        "direct_callee_ir": annotations["owner"]["direct_callee_ir"],
        "validator_inclusive_ir": annotations["validator"]["inclusive_ir"],
        "validation": {
            "receipt_and_command": True,
            "profile_native_identity_parity": True,
            "all_numbered_raw_dumps_valid": True,
            "final_process_dump_zero_ir": True,
            "scanner_edge_policy": True,
            "event_edge_absent_in_both_stage_policy": True,
            "validator_inclusive_row_diagnostic": True,
            **annotations["validation"],
        },
    }


def validate_stage(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    require(stage_dir.is_dir(), f"stage directory is missing: {stage}")
    metadata = PRIOR.validate_build(stage, plan)
    build_receipt = read_json(stage_dir / "build-normal.receipt.json")
    build_command = build_receipt.get("command")
    require(isinstance(build_command, list),
            f"{stage}/build-normal.receipt.json: command is not a list")
    require(build_command[:3] == ["env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0"],
            f"{stage}: build command does not retain the bounded env prefix")
    require(build_command[3:] == [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", "litchi-perf-baseline",
        "--target-dir", plan["owned_paths"][1],
    ], f"{stage}: normal build command differs")
    metadata["build_command"] = build_command
    profiles = [validate_profile(stage, plan, metadata, job)
                for job in profile_jobs(plan)]
    require(len(profiles) == 4, f"{stage}: profile matrix is incomplete")
    return {
        "stage": stage,
        "metadata": metadata,
        "profile_count": len(profiles),
        "profiles": profiles,
        "validation": {
            "expected_profile_matrix": True,
            "receipts_and_bindings": True,
            "native_output_parity": True,
            "raw_parts_and_final_dump": True,
            "scanner_edge_policy": True,
        },
    }


def metric(baseline: int | None, candidate: int | None) -> dict[str, Any]:
    result: dict[str, Any] = {
        "baseline": baseline,
        "candidate": candidate,
        "delta_ir": (candidate - baseline
                     if baseline is not None and candidate is not None else None),
        "candidate_over_baseline": None,
        "delta_percent": None,
    }
    if baseline not in (None, 0) and candidate is not None:
        result["candidate_over_baseline"] = candidate / baseline
        result["delta_percent"] = (candidate / baseline - 1.0) * 100.0
    return result


def by_key(stage_report: dict[str, Any]) -> dict[tuple[int, str], dict[str, Any]]:
    return {(row["repeat"], row["shape"]): row for row in stage_report["profiles"]}


def edge_for_report(row: dict[str, Any], key: str) -> dict[str, Any]:
    edge = row[key]
    return {
        "present": edge["present"],
        "positive_edge_count": edge["positive_edge_count"],
        "calls": edge["calls"],
        "inclusive_ir": edge["inclusive_ir"] if edge["present"] else None,
        "edges": edge["edges"],
        "policy": edge["policy"],
    }


def compare(baseline: dict[str, Any], candidate: dict[str, Any]) -> dict[str, Any]:
    left = by_key(baseline)
    right = by_key(candidate)
    require(set(left) == set(right), "baseline/candidate profile matrices differ")
    rows = []
    totals = {
        field: {"baseline": 0, "candidate": 0}
        for field in ("total_ir", "self_ir", "direct_ir", "validator_inclusive_ir")
    }
    direct_names: set[str] = set()
    for key in sorted(left):
        base = left[key]
        cand = right[key]
        require(base["native_result_identity"] == cand["native_result_identity"],
                f"{base['name']}: baseline/candidate native identity differs")
        for field in totals:
            totals[field]["baseline"] += base[field]
            totals[field]["candidate"] += cand[field]
        direct_names.update(base["direct_callee_ir"])
        direct_names.update(cand["direct_callee_ir"])
        direct = []
        for name in sorted(
            direct_names,
            key=lambda name: (
                -max(base["direct_callee_ir"].get(name, 0),
                     cand["direct_callee_ir"].get(name, 0)),
                name,
            ),
        ):
            direct.append({
                "name": name,
                "metric": metric(base["direct_callee_ir"].get(name, 0),
                                  cand["direct_callee_ir"].get(name, 0)),
            })
        rows.append({
            "repeat": key[0],
            "shape": key[1],
            "baseline_profile": base["name"],
            "candidate_profile": cand["name"],
            "native_identity_equal": True,
            "metrics": {
                "commit_inclusive_ir": metric(base["total_ir"], cand["total_ir"]),
                "commit_self_ir": metric(base["self_ir"], cand["self_ir"]),
                "commit_direct_ir": metric(base["direct_ir"], cand["direct_ir"]),
                "validator_inclusive_ir": metric(
                    base["validator_inclusive_ir"], cand["validator_inclusive_ir"]
                ),
            },
            "direct_callee_ir": direct,
            "scanner_start_cell_to_cell_tag": {
                "baseline": edge_for_report(base, "scanner_start_cell_to_cell_tag"),
                "candidate": edge_for_report(cand, "scanner_start_cell_to_cell_tag"),
                "interpretation": (
                    "The comparison covers the positive raw edge from the scanner "
                    "proof owner to wire::cell_tag, not every cell_tag call."
                ),
            },
            "validator_event_into_owned_edge": {
                "baseline": edge_for_report(base, "validator_event_into_owned_edge"),
                "candidate": edge_for_report(cand, "validator_event_into_owned_edge"),
                "interpretation": (
                    "Both stages require this pre-0522 edge to remain absent; it is "
                    "not credited to the scanner change."
                ),
            },
        })

    aggregate_direct = []
    for name in sorted(
        direct_names,
        key=lambda name: (
            -max(
                sum(row["direct_callee_ir"].get(name, 0) for row in left.values()),
                sum(row["direct_callee_ir"].get(name, 0) for row in right.values()),
            ),
            name,
        ),
    ):
        aggregate_direct.append({
            "name": name,
            "metric": metric(
                sum(row["direct_callee_ir"].get(name, 0) for row in left.values()),
                sum(row["direct_callee_ir"].get(name, 0) for row in right.values()),
            ),
        })
    return {
        "available": True,
        "profile_count": len(rows),
        "profiles": rows,
        "aggregate_totals": {
            name: metric(value["baseline"], value["candidate"])
            for name, value in totals.items()
        },
        "aggregate_direct_callee_ir": aggregate_direct,
        "validator_inclusive_cost_caveat": (
            "Validator inclusive Ir includes child calls. Child-call metadata can "
            "include calls attributable outside the selected scanner edge, so this "
            "row is diagnostic and does not by itself establish the 0522 gain."
        ),
        "validation": {
            "profile_matrix_matches": True,
            "native_identity_parity": True,
            "commit_totals_compared": True,
            "direct_callee_ir_compared": True,
            "scanner_edge_baseline_present_candidate_absent": all(
                row["scanner_start_cell_to_cell_tag"]["baseline"]["present"]
                and not row["scanner_start_cell_to_cell_tag"]["candidate"]["present"]
                for row in rows
            ),
            "event_edge_absent_both_stages": all(
                not row["validator_event_into_owned_edge"]["baseline"]["present"]
                and not row["validator_event_into_owned_edge"]["candidate"]["present"]
                for row in rows
            ),
        },
    }


def stage_has_profiles(stage: str, plan: dict[str, Any]) -> bool:
    stage_dir = HERE / stage
    return stage_dir.is_dir() and any(
        (stage_dir / f"{job['name']}.callgrind.4").is_file()
        for job in profile_jobs(plan)
    )


def analyze(stage_selection: str = "both") -> dict[str, Any]:
    plan = plan_data()
    if stage_selection in ("baseline", "candidate"):
        selected = [stage_selection]
    else:
        selected = [stage for stage in ("baseline", "candidate")
                    if stage_has_profiles(stage, plan)]
        require(selected, "no completed profile stage is available")
    stages = {stage: validate_stage(stage, plan) for stage in selected}
    comparison = (
        compare(stages["baseline"], stages["candidate"])
        if "baseline" in stages and "candidate" in stages else None
    )
    return {
        "schema": "xlsx_callgrind_scanner_profile_analysis_v1",
        "status": "pass",
        "selected_function": OWNER,
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha256(PLAN_PATH),
        "stage_selection": selected,
        "stages": stages,
        "comparison": comparison,
        "helpers": {
            "change-0521/analyze_profiles.py": sha256(PRIOR_ANALYZER),
            "change-0519/analyze_profiles.py": sha256(
                HERE.parent / "change-0519" / "analyze_profiles.py"
            ),
            "change-0519/compare_profile_lanes.py": sha256(
                HERE.parent / "change-0519" / "compare_profile_lanes.py"
            ),
        },
        "limitations": (
            "Callgrind Ir and inclusive rows are mechanism diagnostics, not native "
            "latencies, hardware cycles, allocation counts, cold-cache, range, "
            "scaling, or native Office-producer measurements."
        ),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", nargs="?", type=Path,
                        help="JSON destination; defaults to profile-analysis.json")
    parser.add_argument("--output", dest="output_option", type=Path,
                        help="JSON destination (alternative to the positional path)")
    parser.add_argument("--stage", choices=("baseline", "candidate", "both"), default="both",
                        help="validate one stage, or both stages when available")
    args = parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error("provide the output path either positionally or with --output")
    output = args.output_option or args.output or HERE / "profile-analysis.json"
    try:
        document = analyze(args.stage)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
    except (EvidenceError, OSError, json.JSONDecodeError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2
    print(f"Profile raw scanner edge, native parity, and annotations verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
