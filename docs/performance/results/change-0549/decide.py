"""Bind reviewed 0549 native, allocation, collector and quality evidence."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from run import HERE, sha


MEASURED_STAGES = ("baseline", "candidate")
COLLECTOR_TARGET = "litchi_cfb::file::SectorChainScratch::collect_exact"


def read(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read_regular(path: Path, label: str) -> Any:
    """Read retained JSON only when its custody path is a regular file."""

    require(path.is_file() and not path.is_symlink(),
            f"{label} is not a retained regular file")
    return read(path)


def write_report(path: Path, value: dict[str, Any]) -> None:
    """Create an analysis artifact once and make replay byte-stable."""

    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    if path.exists():
        require(path.is_file() and not path.is_symlink(),
                f"refusing to read non-regular decision output {path}")
        require(path.read_bytes() == data,
                f"refusing to overwrite non-identical decision output {path}")
        return
    require(not (HERE / "SHA256SUMS").exists()
            or not path.resolve().is_relative_to(HERE.resolve()),
            f"sealed evidence is missing decision output {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)


def profile_row(profiles: dict[str, Any], stage: str, group: str,
                repeat: int, shape: str | None = None) -> dict[str, Any]:
    require(stage in MEASURED_STAGES and isinstance(profiles.get(stage), dict),
            f"{stage} profile analysis is missing")
    rows = profiles[stage].get("profiles")
    require(isinstance(rows, list), f"{stage} profile matrix is missing")
    matches = [item for item in rows
               if isinstance(item, dict) and item.get("group") == group
               and item.get("repeat") == repeat
               and (shape is None or item.get("shape") == shape)]
    require(len(matches) == 1,
            f"{stage} profile {group}/{shape or 'none'} repeat {repeat} is not unique")
    return matches[0]


def constructor_ir(profile: dict[str, Any], label: str) -> int:
    attribution = profile.get("constructor_attribution")
    require(isinstance(attribution, list) and attribution,
            f"{label} constructor attribution is missing")
    total = 0
    for item in attribution:
        require(isinstance(item, dict) and isinstance(item.get("constructor"), dict),
                f"{label} constructor attribution row is malformed")
        value = item["constructor"].get("inclusive_ir")
        require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
                f"{label} constructor Ir is not positive")
        total += value
    require(total > 0, f"{label} constructor Ir total is not positive")
    return total


def profile_ir_gate(profiles: dict[str, Any]) -> tuple[list[dict[str, Any]], bool]:
    """Replay the XLS-owned constructor-inclusive-Ir admission gate."""
    rows: list[dict[str, Any]] = []
    for repeat in (1, 2):
        values = {
            stage: constructor_ir(
                profile_row(profiles, stage, "xls-owned", repeat),
                f"{stage} XLS-owned repeat {repeat}",
            )
            for stage in MEASURED_STAGES
        }
        rows.append({"repeat": repeat, "baseline": values["baseline"],
                     "candidate": values["candidate"],
                     "change_percent": 100 * (values["candidate"] /
                                               values["baseline"] - 1)})
    return rows, all(row["candidate"] < row["baseline"] for row in rows)


def _collector_target(item: dict[str, Any], label: str) -> dict[str, Any]:
    """Find the exact collect_exact function attribution in one dump."""
    functions = item.get("functions")
    require(isinstance(functions, dict), f"{label} function attribution is missing")
    matches = [value for value in functions.values()
               if isinstance(value, dict) and value.get("target") == COLLECTOR_TARGET]
    if not matches:
        # Keep the adapter tolerant of the profile analyzer's named alias while
        # still requiring the target identity before accepting its Ir.
        for key in ("collector", "collect_exact"):
            value = functions.get(key)
            if isinstance(value, dict):
                matches.append(value)
    require(len(matches) == 1,
            f"{label} collect_exact attribution is not unique")
    target = matches[0]
    require(target.get("target") == COLLECTOR_TARGET,
            f"{label} collect_exact target identity differs")
    return target


def collector_ir(profile: dict[str, Any], label: str) -> int:
    """Sum exclusive collect_exact self Ir across timed constructor dumps."""
    attribution = profile.get("constructor_attribution")
    require(isinstance(attribution, list) and attribution,
            f"{label} collector attribution is missing")
    total = 0
    for item in attribution:
        require(isinstance(item, dict), f"{label} collector attribution row is malformed")
        target = _collector_target(item, label)
        require(target.get("out_of_line") is True
                and target.get("inlined_or_absent") is False
                and isinstance(target.get("incoming_edge_count"), int)
                and target["incoming_edge_count"] > 0,
                f"{label} collect_exact is not a positive out-of-line target")
        value = target.get("self_ir")
        require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
                f"{label} collect_exact exclusive Ir is malformed")
        total += value
    return total


def collector_profile_gate(profiles: dict[str, Any]) -> tuple[list[dict[str, Any]], bool]:
    """Replay the paired collect_exact exclusive-Ir admission gate."""
    rows: list[dict[str, Any]] = []
    selected = (("xls-owned", None), ("cfb-few-large", "few-large"))
    for repeat in (1, 2):
        for group, shape in selected:
            values = {
                stage: collector_ir(
                    profile_row(profiles, stage, group, repeat, shape),
                    f"{stage} {group} repeat {repeat}",
                )
                for stage in MEASURED_STAGES
            }
            require(values["baseline"] > 0,
                    f"baseline {group} repeat {repeat} collect_exact Ir is not positive")
            rows.append({
                "repeat": repeat,
                "group": group,
                "shape": shape,
                "baseline": values["baseline"],
                "candidate": values["candidate"],
                "change_percent": 100 * (values["candidate"] /
                                          values["baseline"] - 1),
            })
    return rows, all(row["candidate"] < row["baseline"] for row in rows)


def review_rows(raw: list[Any], checked: Any, label: str) -> None:
    require(isinstance(checked, list) and len(checked) == len(raw),
            f"{label} review length differs")
    remaining = list(checked)
    for item in raw:
        require(isinstance(item, dict), f"{label} raw row is malformed")
        index = next((i for i, candidate in enumerate(remaining)
                      if isinstance(candidate, dict)
                      and all(candidate.get(key) == value
                              for key, value in item.items())), None)
        require(index is not None, f"{label} row is not individually reviewed")
        candidate = remaining.pop(index)
        explanation = candidate.get("review", candidate.get("reason"))
        require(isinstance(explanation, str) and explanation.strip(),
                f"{label} row lacks an explanation")
    require(not remaining, f"{label} review has extra rows")


def evaluate() -> dict[str, Any]:
    comparison_path = HERE / "comparison.json"
    comparison = read(comparison_path)["comparison"]
    admission = comparison["admission"]
    primary = admission["primary_workflow_p50"]["all_four_cases_both_repeats_pass"]
    memory = admission["allocation_guard"]["passes"]
    profiles = {
        stage: read(HERE / stage / "profile-analysis.json")
        for stage in MEASURED_STAGES
    }
    owner, profile_gate = profile_ir_gate(profiles)
    collector, collector_gate = collector_profile_gate(profiles)

    # The public malformed-input lane is an independent hard admission gate.
    # Its analyzer deliberately reports status=pass for a completed analysis;
    # admission_passed is the boolean that binds the threshold result.
    guard_analysis_path = HERE / "guard-analysis.json"
    guard_analysis = read_regular(guard_analysis_path, "guard-analysis.json")
    plan_path = HERE / "plan.json"
    require(guard_analysis.get("status") == "pass"
            and guard_analysis.get("plan_sha256") == sha(plan_path)
            and isinstance(guard_analysis.get("admission_passed"), bool),
            "guard analysis custody or admission result is incomplete")
    require(isinstance(guard_analysis.get("adverse"), list)
            and isinstance(guard_analysis.get("drift"), list),
            "guard analysis adverse/drift rows are missing")
    guard_gate = guard_analysis["admission_passed"]
    guard_analysis_sha256 = sha(guard_analysis_path)

    quality = read(HERE / "quality-summary.json")
    quality_checks = quality.get("checks")
    require(quality.get("status") == "pass"
            and quality.get("stage") in {"candidate", "final"}
            and isinstance(quality_checks, list)
            and len(quality_checks) == 15,
            "quality summary does not contain all 15 passing checks")
    quality_names = [item.get("name") if isinstance(item, dict) else None
                     for item in quality_checks]
    require(len(set(quality_names)) == 15
            and all(isinstance(name, str) and name.startswith("check-")
                    for name in quality_names)
            and all(item.get("receipt_sha256")
                    for item in quality_checks),
            "quality summary contains an incomplete check receipt")

    review = read(HERE / "adverse-review.json")
    require(review.get("comparison_sha256") == sha(comparison_path),
            "adverse review comparison digest differs")
    require(review.get("guard_analysis_sha256") == guard_analysis_sha256,
            "adverse review guard-analysis digest differs")
    require(review.get("complete") is True
            and isinstance(review.get("adoption_allowed"), bool),
            "adverse review is incomplete")
    review_rows(comparison["matched_adverse_flags_over_five_percent"],
                review.get("matched"), "matched")
    review_rows(comparison["same_build_variations_over_five_percent"],
                review.get("same_build"), "same-build")
    # Keep the retained review keys aligned with the prior adverse-review
    # envelope: guard_adverse_flags covers guard-analysis.adverse and
    # guard_same_build_drift_flags covers guard-analysis.drift.
    review_rows(guard_analysis["adverse"],
                review.get("guard_adverse_flags"), "guard-adverse")
    review_rows(guard_analysis["drift"],
                review.get("guard_same_build_drift_flags"), "guard-drift")

    adoption_allowed = bool(primary and memory and profile_gate and
                            collector_gate and guard_gate and
                            review["adoption_allowed"])
    disposition = "accepted" if adoption_allowed else "rejected"
    final_source = "candidate" if adoption_allowed else "final"
    final_manifest_path = HERE / final_source / "source-manifest.json"
    require(final_manifest_path.is_file() and not final_manifest_path.is_symlink(),
            f"{final_source} source manifest is missing from retained custody")
    if disposition == "rejected":
        baseline_manifest_path = HERE / "baseline" / "source-manifest.json"
        require(baseline_manifest_path.is_file()
                and not baseline_manifest_path.is_symlink(),
                "baseline source manifest is missing from retained custody")
        require(sha(final_manifest_path) == sha(baseline_manifest_path),
                "rejected decision does not retain an exact restored baseline manifest")
    require(quality["stage"] == final_source,
            "quality summary stage does not match decision final source")
    return {
        "disposition": disposition,
        "native_primary_gate": primary,
        "memory_gate": memory,
        "profile_gate": profile_gate,
        "collector_profile_gate": collector_gate,
        "guard_gate": guard_gate,
        "adoption_allowed": adoption_allowed,
        "constructor_ir": owner,
        "collector_ir": collector,
        "adverse_review_complete": True,
        "quality_gates": 15,
        "other_rejection_reason": review.get("rejection_reason"),
        "comparison_sha256": sha(comparison_path),
        "profiles_sha256": {
            stage: sha(HERE / stage / "profile-analysis.json")
            for stage in MEASURED_STAGES
        },
        "quality_summary_sha256": sha(HERE / "quality-summary.json"),
        "adverse_review_sha256": sha(HERE / "adverse-review.json"),
        "guard_analysis_sha256": guard_analysis_sha256,
        "final_source": final_source,
        "final_source_manifest_sha256": sha(final_manifest_path),
        "scope": (
            "Matched CFB checked bitset test-and-mark experiment; OLE2/OOXML active, "
            "ODF deferred, iWork excluded"
        ),
    }


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('output', nargs='?', type=Path)
    parser.add_argument('--output', dest='output_option', type=Path)
    args = parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error('provide output either positionally or with --output')
    result = evaluate()
    destination = args.output_option or args.output or HERE / 'decision.json'
    destination.parent.mkdir(parents=True, exist_ok=True)
    write_report(destination, result)
    print(json.dumps(result))
