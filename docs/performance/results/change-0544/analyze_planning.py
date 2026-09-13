"""Replay the 0544 paired planning profiles and attribute their Ir costs.

The profile lane is retained only as a mechanism diagnostic.  Each profile
process emits lifecycle dumps while collection is disabled and one measured
``SourceBackedEditor::edit_sheets`` dump after collection is enabled.  This
analyzer checks that exact raw edge, the lifecycle fallback, the termination
dump, and the inclusive/self annotation accounting before comparing the
matched baseline and candidate Ir totals.

``analyze(False)`` is deliberately read-only.  The command line ``--write``
mode may create the two deterministic callgrind annotations and the report.
No Ir value is converted to latency or presented as a latency claim.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Any


sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
HELPER = HERE.parent / "change-0521" / "analyze_profiles.py"
_helper_spec = importlib.util.spec_from_file_location(
    "xlsx_0521_profile_helpers_for_0544", HELPER
)
if _helper_spec is None or _helper_spec.loader is None:
    raise ImportError(f"cannot load profile helper: {HELPER}")
h = importlib.util.module_from_spec(_helper_spec)
_helper_spec.loader.exec_module(h)
# The helper's annotation command has no evidence-path input other than its
# module global.  Rebind it so this report remains tied to the 0544 bundle.
h.HERE = HERE


def display_name(text: str) -> str:
    """Normalize callgrind names without losing Rust bracketed impl names."""

    # Only the final bracketed field is the executable.  Rust symbols can
    # contain earlier brackets, e.g. core::slice::<impl [T]>::sort_unstable_by.
    text = text.rsplit(" [", 1)[0].strip()
    text = re.sub(r"\s+\([\d,]+x\)$", "", text)
    if text.startswith("???:"):
        return text[4:]
    return text.rsplit(":", 1)[-1] if text.startswith(("./", "/")) else text


h.display_name = display_name

PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
LIFECYCLE = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
STAGES = ("baseline", "candidate")


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def _read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def _profile_config(plan: dict[str, Any]) -> dict[str, Any]:
    primary = plan.get("primary")
    profile = plan.get("profile")
    if not isinstance(primary, dict) or not isinstance(profile, dict):
        raise ValueError("plan primary/profile sections are missing")
    shapes = profile.get("shapes")
    repeats = profile.get("repeats")
    owner = profile.get("owner")
    if not isinstance(shapes, list) or not shapes or not all(
        isinstance(shape, str) and shape for shape in shapes
    ):
        raise ValueError("plan profile shapes are invalid")
    if not isinstance(repeats, int) or isinstance(repeats, bool) or repeats < 1:
        raise ValueError("plan profile repeats are invalid")
    if not isinstance(owner, str) or not owner:
        raise ValueError("plan profile owner is missing")
    # 0544 freezes the exact method name.  An owner-candidate fallback would
    # allow a profile to silently attribute a different symbol.
    if "owner_candidates" in profile:
        raise ValueError("0544 profile must use its direct owner")
    if primary.get("case") != "xlsx_source_backed_cell_values_one_percent_edit_save":
        raise ValueError("plan primary case differs from the profile capture")
    if profile.get("warmup") != 0 or profile.get("samples") != 1:
        raise ValueError("profile warmup/samples differ from the frozen capture")
    return {"shapes": shapes, "repeats": repeats, "owner": owner}


def _check_symbol_observation(plan: dict[str, Any], owner: str) -> str | None:
    """Bind an available frozen nm observation to the direct plan owner.

    The root driver owns creation of this artifact.  Keeping the check
    optional lets a caller replay an individual profile stage before that
    separate nm receipt is materialized, while a present observation is never
    allowed to disagree with the frozen plan.
    """

    path = HERE / "symbol-observation.json"
    if not path.is_file():
        return None
    observation = _read_json(path)
    if not isinstance(observation, dict):
        raise ValueError("symbol observation is not an object")
    if observation.get("owner") != owner:
        raise ValueError("symbol observation owner differs from plan.profile.owner")
    if observation.get("plan_sha256") != sha(HERE / "plan.json"):
        raise ValueError("symbol observation plan binding differs")
    return sha(path)


def _numbered_dumps(stem: Path) -> list[tuple[int, Path]]:
    raw = Path(str(stem) + ".callgrind")
    candidates: list[tuple[int, Path]] = []
    prefix = raw.name + "."
    for path in raw.parent.glob(prefix + "*"):
        suffix = path.name[len(prefix):]
        if suffix.isdigit():
            candidates.append((int(suffix), path))
    candidates.sort(key=lambda item: item[0])
    if not candidates:
        raise ValueError(f"{raw}: no numbered owner dumps captured")
    numbers = [number for number, _ in candidates]
    if numbers != list(range(1, len(numbers) + 1)):
        raise ValueError(f"{raw}: dump parts are not contiguous: {numbers}")
    return candidates


def _edge(path: Path, owner: str, caller: str) -> dict[str, Any]:
    return h.target_edge_summary(path, owner, caller)


def _validate_parts(
    stage: str,
    repeat: int,
    shape: str,
    stem: Path,
    owner: str,
) -> tuple[list[dict[str, Any]], Path, dict[str, Any]]:
    """Validate lifecycle/owner dumps and return the final measured dump."""

    numbered = _numbered_dumps(stem)
    parts: list[dict[str, Any]] = []
    selected: list[Path] = []
    for number, path in numbered:
        text = path.read_text(encoding="utf-8")
        label = str(path.relative_to(HERE))
        if "events: Ir" not in text.splitlines():
            raise ValueError(f"{label}: callgrind event set does not contain Ir")
        if h.part_number(text, label) != number:
            raise ValueError(f"{label}: part number does not match its suffix")
        expected_trigger = f"--dump-after={owner}"
        if h.trigger(text, label) != expected_trigger:
            raise ValueError(f"{label}: trigger is not the exact owner dump")

        total = h.summary_ir(text, label)
        measured = _edge(path, owner, PARENT)
        lifecycle = _edge(path, owner, LIFECYCLE)
        record = {
            "part": number,
            "path": label,
            "sha256": sha(path),
            "summary_ir": total,
            "measured_parent": measured,
            "lifecycle_parent": lifecycle,
        }
        parts.append(record)

        # Collection is expected to switch to the exact measured caller only
        # for the final lifecycle part.  Earlier parts are retained through
        # the lifecycle wrapper edge so setup/guards cannot be mistaken for
        # planning Ir.
        if measured["inclusive_ir"] > 0:
            if measured["positive_edge_count"] != 1 or measured["calls"] != 1:
                raise ValueError(f"{label}: measured parent edge is not one call")
            if measured["inclusive_ir"] != total:
                raise ValueError(f"{label}: measured edge Ir differs from summary")
            selected.append(path)
        else:
            if lifecycle["positive_edge_count"] != 1 or lifecycle["calls"] != 1:
                raise ValueError(f"{label}: lifecycle parent edge is not one call")
            if lifecycle["inclusive_ir"] != total:
                raise ValueError(f"{label}: lifecycle edge Ir differs from summary")

    if len(selected) != 1 or selected[0] != numbered[-1][1]:
        raise ValueError(
            f"{stage} repeat {repeat} {shape}: measured owner call must be unique and final"
        )

    raw = Path(str(stem) + ".callgrind")
    termination_text = raw.read_text(encoding="utf-8")
    if "events: Ir" not in termination_text.splitlines():
        raise ValueError(f"{raw.relative_to(HERE)}: termination event set lacks Ir")
    termination_summary = h.summary_ir(termination_text, str(raw))
    termination_trigger = h.trigger(termination_text, str(raw))
    if termination_summary != 0 or termination_trigger != "Program termination":
        raise ValueError(
            f"{raw.relative_to(HERE)}: termination dump is not zero Ir/Program termination"
        )
    termination = {
        "path": str(raw.relative_to(HERE)),
        "sha256": sha(raw),
        "summary_ir": termination_summary,
        "trigger": termination_trigger,
    }
    return parts, selected[0], termination


def _annotation(
    stem: Path,
    selected: Path,
    inclusive: bool,
    create_annotations: bool,
) -> tuple[Path, str]:
    suffix = ".inclusive.txt" if inclusive else ".self.txt"
    path = Path(str(stem) + suffix)
    output, _command = h.run_annotation(selected, inclusive)
    if create_annotations:
        if path.exists():
            if path.read_text(encoding="utf-8") != output:
                raise ValueError(f"{path.relative_to(HERE)} differs from replayed annotation")
        else:
            path.write_text(output, encoding="utf-8")
    elif not path.is_file() or path.read_text(encoding="utf-8") != output:
        raise ValueError(f"{path.relative_to(HERE)} does not match annotation replay")
    return path, output


def _owners(
    selected: Path,
    inclusive_path: Path,
    inclusive_text: str,
    self_path: Path,
    self_text: str,
    owner: str,
) -> dict[str, Any]:
    """Replay direct annotation accounting and bounded material child edges."""

    inc = h.parse_annotation(inclusive_text, owner, str(inclusive_path))
    own = h.parse_annotation(self_text, owner, str(self_path))
    direct = h.direct_map(inc["direct"])
    if direct != h.direct_map(own["direct"]):
        raise ValueError("inclusive and self annotation child maps differ")
    if inc["selected_ir"] != own["selected_ir"] + sum(direct.values()):
        raise ValueError("owner annotation Ir accounting does not balance")

    owners: dict[str, Any] = {}
    pending: list[tuple[str, int]] = [(owner, 0)]
    while pending:
        name, depth = pending.pop(0)
        if name in owners:
            continue
        child_inc = h.parse_annotation(inclusive_text, name, str(inclusive_path))
        child_self = h.parse_annotation(self_text, name, str(self_path))
        children = h.direct_map(child_inc["direct"])
        if children != h.direct_map(child_self["direct"]):
            raise ValueError(f"{name}: inclusive/self child maps differ")
        if child_inc["selected_ir"] != child_self["selected_ir"] + sum(children.values()):
            raise ValueError(f"{name}: annotation Ir accounting does not balance")
        for child, cost in children.items():
            raw_edge = _edge(selected, child, name)
            if raw_edge["inclusive_ir"] != cost:
                raise ValueError(f"{name} -> {child}: annotation/raw Ir differs")
        owners[name] = {
            "inclusive_ir": child_inc["selected_ir"],
            "self_ir": child_self["selected_ir"],
            "direct": dict(sorted(children.items(), key=lambda item: (-item[1], item[0]))),
        }
        if depth < 6:
            pending.extend(
                (child, depth + 1)
                for child, cost in children.items()
                if cost >= inc["selected_ir"] * 0.05
            )
    return owners


def _stage_row(
    stage: str,
    repeat: int,
    shape: str,
    plan: dict[str, Any],
    owner: str,
    create_annotations: bool,
) -> dict[str, Any]:
    stem = HERE / stage / f"profile-r{repeat}-{shape}"
    parts, selected, termination = _validate_parts(stage, repeat, shape, stem, owner)
    inclusive_path, inclusive_text = _annotation(
        stem, selected, True, create_annotations
    )
    self_path, self_text = _annotation(stem, selected, False, create_annotations)
    inc = h.parse_annotation(inclusive_text, owner, str(inclusive_path))
    own = h.parse_annotation(self_text, owner, str(self_path))
    direct = h.direct_map(inc["direct"])
    if direct != h.direct_map(own["direct"]):
        raise ValueError(f"{stage}/{shape}/r{repeat}: owner child maps differ")
    final_ir = parts[-1]["summary_ir"]
    if inc["selected_ir"] != own["selected_ir"] + sum(direct.values()):
        raise ValueError(f"{stage}/{shape}/r{repeat}: owner annotation Ir does not balance")
    if inc["selected_ir"] != final_ir:
        raise ValueError(f"{stage}/{shape}/r{repeat}: owner Ir differs from final summary")
    owners = _owners(
        selected, inclusive_path, inclusive_text, self_path, self_text, owner
    )
    return {
        "stage": stage,
        "repeat": repeat,
        "shape": shape,
        "owner": owner,
        "planning_ir": final_ir,
        "parts": parts,
        "selected": str(selected.relative_to(HERE)),
        "termination": termination,
        "annotations": {
            str(inclusive_path.relative_to(HERE)): sha(inclusive_path),
            str(self_path.relative_to(HERE)): sha(self_path),
        },
        "owner_accounting": {
            "inclusive_ir": inc["selected_ir"],
            "self_ir": own["selected_ir"],
            "direct": dict(sorted(direct.items(), key=lambda item: (-item[1], item[0]))),
        },
        "owners": owners,
    }


def _reduction(baseline: int, candidate: int, required: float) -> dict[str, Any]:
    if baseline <= 0:
        raise ValueError("baseline planning Ir must be positive for reduction")
    reduction_ir = baseline - candidate
    reduction_percent = 100.0 * reduction_ir / baseline
    passed = reduction_percent >= required
    if not passed:
        raise ValueError(
            f"planning Ir reduction {reduction_percent:.6f}% is below {required:.6f}%"
        )
    return {
        "baseline": baseline,
        "candidate": candidate,
        "reduction_ir": reduction_ir,
        "reduction_percent": reduction_percent,
        "required_reduction_percent": required,
        "passed": passed,
    }


def analyze(create_annotations: bool = False) -> dict[str, Any]:
    """Validate both profile stages and compare each matched shape/repeat."""

    plan_path = HERE / "plan.json"
    plan = _read_json(plan_path)
    if not isinstance(plan, dict):
        raise ValueError("plan is not an object")
    profile = _profile_config(plan)
    owner = profile["owner"]
    observation_sha = _check_symbol_observation(plan, owner)
    required = float(plan.get("gates", {}).get("planning_ir_reduction_percent", 5.0))
    if required < 0:
        raise ValueError("planning Ir reduction gate is negative")

    stages: dict[str, list[dict[str, Any]]] = {}
    keyed: dict[tuple[int, str], dict[str, dict[str, Any]]] = {}
    for stage in STAGES:
        stage_rows: list[dict[str, Any]] = []
        for repeat in range(1, profile["repeats"] + 1):
            for shape in profile["shapes"]:
                row = _stage_row(
                    stage, repeat, shape, plan, owner, create_annotations
                )
                stage_rows.append(row)
                keyed.setdefault((repeat, shape), {})[stage] = row
        stages[stage] = stage_rows

    rows: list[dict[str, Any]] = []
    for repeat in range(1, profile["repeats"] + 1):
        for shape in profile["shapes"]:
            pair = keyed[(repeat, shape)]
            rows.append(
                {
                    "repeat": repeat,
                    "shape": shape,
                    "baseline": pair["baseline"],
                    "candidate": pair["candidate"],
                    "planning_ir": _reduction(
                        pair["baseline"]["planning_ir"],
                        pair["candidate"]["planning_ir"],
                        required,
                    ),
                }
            )

    result: dict[str, Any] = {
        "schema": "xlsx_0544_paired_planning_profile_analysis_v1",
        "status": "pass",
        "plan_sha256": sha(plan_path),
        "helper_sha256": sha(HELPER),
        "owner": owner,
        "stages": stages,
        "rows": rows,
        "symbol_observation_sha256": observation_sha,
        "performance_claim": "none",
        "scope": (
            "Matched baseline/candidate callgrind Ir attribution for the exact "
            "SourceBackedEditor::edit_sheets owner across the frozen two-shape, "
            "two-repeat profile matrix. Ir is diagnostic and is not converted "
            "to latency or used as a native timing claim."
        ),
        "limitations": [
            "Nested owner totals overlap; immediate raw caller edges are the disjoint accounting boundary.",
            "Lifecycle dumps include collection-off work and are retained for edge identity, not allocation counts.",
            "The paired Ir gate is conditional on the native, allocation, and guard pilot gates.",
            "This profile does not establish cold-cache, range, scaling, hardware, or production latency behavior.",
        ],
    }
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--write",
        "--create-annotations",
        dest="write",
        action="store_true",
        help="create deterministic annotations and planning-profile-analysis.json",
    )
    parser.add_argument("--output", type=Path, help="JSON report destination")
    args = parser.parse_args()
    result = analyze(args.write)
    if args.write or args.output is not None:
        output = args.output or HERE / "planning-profile-analysis.json"
        output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(f"0544 paired planning profile verified: {output}")
    else:
        print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
