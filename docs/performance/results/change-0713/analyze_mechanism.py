#!/usr/bin/env python3
"""Audit the gated 0713 shared-MCE Callgrind and RSS mechanism packet.

The analyzer consumes retained artifacts only.  It keeps the exact raw
Callgrind parts and process RSS peaks, and reuses the retained 0709 raw-edge
parser plus the retained 0707 edge cross-check.  Callgrind Ir and RSS are
diagnostics; this packet emits no optimization decision.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import re
import sys
from typing import Any, NoReturn


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
MECHANISM_PATH = HERE / "mechanism.py"
PLAN_PATH = HERE / "mechanism-plan.json"
OWNER = "litchi_perf_baseline::ordinary_save::Owner::edit"
MEASURED_PARENT = "litchi_perf_baseline::ordinary_save::run_case"
TARGET_DELTA = "crates/litchi-ooxml-common/src/mce/codec.rs"
STAGE_BY_LABEL = {
    "baseline-A1": {"source": "baseline", "build": "baseline", "pair": "pair-1", "repeat": 1, "order": "forward"},
    "candidate-B1": {"source": "candidate", "build": "candidate", "pair": "pair-1", "repeat": 1, "order": "forward"},
    "candidate-B2": {"source": "candidate", "build": "candidate", "pair": "pair-2", "repeat": 2, "order": "reverse"},
    "baseline-A2": {"source": "baseline", "build": "baseline", "pair": "pair-2", "repeat": 2, "order": "reverse"},
}
PHASES = ("edit", "lifecycle")
PEAK_RE = re.compile(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", re.MULTILINE)


class EvidenceError(RuntimeError):
    """A missing, malformed, or contradictory retained artifact."""


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest_json(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON input: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON in {path}: {error}")


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        fail(f"cannot read {path}: {error}")


def write_json(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def load_module(path: Path, name: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing helper: {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load helper: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_mechanism() -> Any:
    return load_module(MECHANISM_PATH, "mechanism_0713_for_analysis")


def load_retained_0709() -> Any:
    """Keep the reviewed 0709 parser binding, including its 0705/0521 helpers."""

    path = REPO / "docs/performance/results/change-0709/analyze_profiles.py"
    module = load_module(path, "retained_0709_for_mechanism_analysis")
    module.HERE = HERE
    module.H.HERE = HERE
    return module


def load_retained_0707() -> Any:
    """Keep the reviewed 0707 raw-edge cross-check binding."""

    path = REPO / "docs/performance/results/change-0707/analyze_profiles.py"
    module = load_module(path, "retained_0707_for_mechanism_analysis")
    module.HERE = HERE
    module.H.HERE = HERE
    return module


M = load_mechanism()
PARSER = load_retained_0709()
PARSER_0707 = load_retained_0707()
H = PARSER.H


def load_plan() -> dict[str, Any]:
    return M.load_plan()


def source_state(plan: dict[str, Any]) -> dict[str, Any]:
    value = M.source_state(plan)
    require(value["changed_paths"] == [TARGET_DELTA], "source delta differs")
    return value


def fixture_map(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {item["id"]: item for item in plan["corpora"]}


def fixture_binding(corpus: dict[str, Any]) -> dict[str, Any]:
    return M.fixture_binding(corpus)


def normalized_cleanup_policy() -> dict[str, Any]:
    """Describe binary validation without exposing live/cleanup state.

    ``build_info`` still validates each binary as a live file or an exact
    retained cleanup witness.  The validation mode is deliberately omitted
    from the report so replay after owned scratch cleanup is byte-identical.
    """

    return {
        "required": True,
        "validation": "Every retained build binary is checked as the exact live file or by an exact post-cleanup witness with matching path, SHA-256, and byte count.",
        "output_mode": "normalized-invariant",
    }


def normalized_build_custody(build: dict[str, Any]) -> dict[str, Any]:
    """Return stable binary identity fields, excluding live/witness mode."""

    return {
        "path": build["binary"],
        "sha256": build["binary_sha256"],
        "bytes": build["binary_bytes"],
        "validation": "exact-live-binary-or-exact-postcleanup-witness",
    }


def build_info(plan: dict[str, Any], stage: str, source: dict[str, Any]) -> dict[str, Any]:
    return M.build_info(plan, stage, source)


def stage_order(plan: dict[str, Any], stage: str) -> list[dict[str, Any]]:
    return M.ordered_corpora(plan, stage)


def name(stage: str, corpus_id: str, kind: str, phase: str) -> str:
    corpus = fixture_map(M.load_plan())[corpus_id]
    return (M.profile_name(stage, corpus) if kind == "profile"
            else M.rss_name(stage, corpus, phase))


def validate_source_receipt(receipt: dict[str, Any], source: dict[str, Any], label: str) -> dict[str, Any]:
    custody = receipt.get("current_checkout_source")
    require(isinstance(custody, dict), f"{label}: source custody is missing")
    before_path = HERE / str(custody.get("before_artifact"))
    after_path = HERE / str(custody.get("after_artifact"))
    require(before_path.parent == HERE and after_path.parent == HERE
            and before_path.name == f"{label}.source-before.json"
            and after_path.name == f"{label}.source-after.json"
            and before_path.is_file() and after_path.is_file(),
            f"{label}: source artifacts missing or outside packet")
    before = read_json(before_path)
    after = read_json(after_path)
    require(before == source["candidate"] and after == source["candidate"],
            f"{label}: current source differs from candidate")
    require(before == after and custody.get("unchanged_during_child") is True,
            f"{label}: source changed during child")
    require(custody.get("before_sha256") == digest_json(before)
            and custody.get("after_sha256") == digest_json(after)
            and custody.get("before_file_sha256") == sha(before_path)
            and custody.get("after_file_sha256") == sha(after_path),
            f"{label}: source artifact digest differs")
    delta = receipt.get("source_delta")
    require(isinstance(delta, dict)
            and delta.get("allowed_paths") == [TARGET_DELTA]
            and delta.get("current_checkout_matches_candidate") is True,
            f"{label}: source-delta custody differs")
    return {"before": before_path.name, "after": after_path.name,
            "before_sha256": sha(before_path), "after_sha256": sha(after_path), "equal": True}


def validate_fixture_receipt(receipt: dict[str, Any], corpus: dict[str, Any], label: str) -> dict[str, Any]:
    value = receipt.get("fixture")
    require(isinstance(value, dict), f"{label}: fixture custody is missing")
    before_path = HERE / str(value.get("before_artifact"))
    after_path = HERE / str(value.get("after_artifact"))
    require(before_path.parent == HERE and after_path.parent == HERE
            and before_path.name == f"{label}.fixture-before.json"
            and after_path.name == f"{label}.fixture-after.json"
            and before_path.is_file() and after_path.is_file(),
            f"{label}: fixture artifacts missing or outside packet")
    before = read_json(before_path)
    after = read_json(after_path)
    expected = fixture_binding(corpus)
    require(before == expected and after == expected and before == after,
            f"{label}: fixture identity differs")
    require(value.get("before") == before and value.get("after") == after
            and value.get("plan_path") == corpus.get("path")
            and value.get("plan_sha256") == corpus.get("sha256"),
            f"{label}: receipt fixture values differ")
    return {"before": before_path.name, "after": after_path.name,
            "before_sha256": sha(before_path), "after_sha256": sha(after_path), "equal": True}


def validate_artifact_inventory(receipt: dict[str, Any], label: str) -> dict[str, str]:
    inventory = receipt.get("artifacts")
    require(isinstance(inventory, dict) and inventory, f"{label}: artifact inventory is missing")
    name = receipt.get("name")
    kind = receipt.get("kind")
    require(isinstance(name, str) and kind in {"profile", "rss"},
            f"{label}: receipt name/kind is invalid")
    required = {
        f"{name}.json", f"{name}.stdout", f"{name}.stderr",
        f"{name}.source-before.json", f"{name}.source-after.json",
        f"{name}.fixture-before.json", f"{name}.fixture-after.json",
    }
    if kind == "profile":
        required.add(f"{name}.callgrind")
        numbered = sorted(
            path.name for path in HERE.glob(f"{name}.callgrind.[0-9]*")
            if path.is_file() and not path.is_symlink() and path.suffix[1:].isdigit()
        )
        require(numbered == [f"{name}.callgrind.{index}"
                             for index in range(1, len(numbered) + 1)],
                f"{label}: inventoried Callgrind parts are not contiguous")
        required.update(numbered)
    else:
        required.add(f"{name}.time-v")
    require(required.issubset(inventory),
            f"{label}: receipt omits required artifacts: {sorted(required - set(inventory))}")
    checked: dict[str, str] = {}
    for raw, expected in sorted(inventory.items()):
        require(isinstance(raw, str) and isinstance(expected, str),
                f"{label}: artifact inventory entry is malformed")
        path = (HERE / raw).resolve()
        require(path.parent == HERE.resolve() and path.is_file() and not path.is_symlink(),
                f"{label}: inventoried artifact is invalid: {raw}")
        actual = sha(path)
        require(actual == expected, f"{label}: artifact hash differs: {raw}")
        checked[raw] = actual
    return checked


def expected_command(plan: dict[str, Any], corpus: dict[str, Any], phase: str,
                     kind: str, build: dict[str, Any], report_path: Path) -> list[str]:
    base = M.command_common(build["binary"], plan, corpus, phase, report_path,
                            plan["callgrind"]["samples"] if kind == "profile"
                            else plan["rss"]["samples"],
                            plan["callgrind"]["warmup"] if kind == "profile"
                            else plan["rss"]["warmup"])
    if kind == "profile":
        return ["taskset", "-c", str(plan["cpu"]), plan["callgrind"]["tool"],
                *plan["callgrind"]["options"],
                f"--callgrind-out-file={HERE / (report_path.stem + '.callgrind')}", *base[3:]]
    return ["taskset", "-c", str(plan["cpu"]), plan["rss"]["tool"], "-v", "-o",
            str(HERE / (report_path.stem + ".time-v")), *base[3:]]


def validate_common_receipt(receipt: dict[str, Any], report_path: Path,
                            plan: dict[str, Any], source: dict[str, Any], stage: str,
                            corpus: dict[str, Any], phase: str, kind: str,
                            build: dict[str, Any], gate: dict[str, Any]) -> dict[str, Any]:
    label = report_path.stem
    metadata = STAGE_BY_LABEL[stage]
    require(receipt.get("schema_version") == 1 and receipt.get("kind") == kind,
            f"{label}: receipt schema/kind differs")
    require(receipt.get("name") == label
            and receipt.get("packet") == plan["packet"] and receipt.get("stage") == stage
            and receipt.get("source_label") == metadata["source"]
            and receipt.get("build_label") == metadata["build"]
            and receipt.get("pair") == metadata["pair"]
            and receipt.get("repeat") == metadata["repeat"]
            and receipt.get("stage_order") == metadata["order"]
            and receipt.get("corpus_id") == corpus["id"]
            and receipt.get("corpus_label") == corpus["label"]
            and receipt.get("corpus_origin") == corpus["origin"]
            and receipt.get("phase") == phase
            and receipt.get("case") == M.case_name(corpus, phase),
            f"{label}: job identity differs")
    require(receipt.get("exit_code") == 0 and receipt.get("cpu") == plan["cpu"],
            f"{label}: child status or CPU differs")
    require(receipt.get("binary") == build["binary"]
            and receipt.get("binary_sha256") == build["binary_sha256"]
            and receipt.get("binary_bytes") == build["binary_bytes"],
            f"{label}: binary identity differs")
    require(receipt.get("build_record") == build["record_path"]
            and receipt.get("build_record_sha256") == build["record_sha256"]
            and receipt.get("build_source_manifest") == build["source_path"]
            and receipt.get("build_source_manifest_sha256") == build["source_manifest_sha256"],
            f"{label}: build custody differs")
    require(receipt.get("plan_sha256") == sha(PLAN_PATH)
            and receipt.get("pilot_plan_sha256") == sha(HERE / "plan.json")
            and receipt.get("script_sha256") == sha(MECHANISM_PATH)
            and receipt.get("constraints_sha256") == sha(HERE / "constraints.json"),
            f"{label}: plan/script/constraint hash differs")
    pilot = receipt.get("pilot_gate")
    require(isinstance(pilot, dict) and pilot.get("status") == "pass"
            and pilot.get("artifact") == gate["artifact"]
            and pilot.get("artifact_sha256") == gate["artifact_sha256"],
            f"{label}: pilot gate custody differs")
    source_custody = validate_source_receipt(receipt, source, label)
    fixture_custody = validate_fixture_receipt(receipt, corpus, label)
    artifact_inventory = validate_artifact_inventory(receipt, label)
    source_delta = receipt.get("source_delta")
    expected_changed = [TARGET_DELTA] if build["source_label"] == "baseline" else []
    require(isinstance(source_delta, dict)
            and source_delta.get("changed_paths_from_binary_source") == expected_changed
            and source_delta.get("baseline_manifest") == source["baseline_sha256"]
            and source_delta.get("candidate_manifest") == source["candidate_sha256"],
            f"{label}: source relation differs")
    identity = M.report_identity(report_path)
    require(receipt.get("output_identity") == identity, f"{label}: output identity differs")
    require(receipt.get("command") == expected_command(plan, corpus, phase, kind, build, report_path),
            f"{label}: full command differs")
    return {"label": label, "identity": identity, "source": source_custody,
            "fixture": fixture_custody, "artifacts": artifact_inventory,
            "build": normalized_build_custody(build)}


def callgrind_numbered(stem: Path) -> list[Path]:
    paths = [path for path in HERE.glob(stem.name + ".[0-9]*")
             if path.is_file() and path.suffix[1:].isdigit()]
    paths.sort(key=lambda path: int(path.suffix[1:]))
    numbers = [int(path.suffix[1:]) for path in paths]
    require(numbers and numbers == list(range(1, len(paths) + 1)),
            f"{stem.name}: Callgrind parts are not contiguous: {numbers}")
    return paths


def callgrind_terminal(stem: Path, numbered: list[Path]) -> dict[str, Any]:
    require(stem.is_file() and not stem.is_symlink(), f"{stem.name}: termination part missing")
    text = read_text(stem)
    label = stem.name
    require(H.part_number(text, label) == int(numbered[-1].suffix[1:]) + 1,
            f"{label}: termination part number differs")
    require(H.trigger(text, label) == "Program termination",
            f"{label}: termination trigger differs")
    require(H.summary_ir(text, label) == 0 and "events: Ir" in text.splitlines(),
            f"{label}: termination part is not zero-Ir")
    return {"file": stem.name, "part": int(numbered[-1].suffix[1:]) + 1,
            "sha256": sha(stem), "summary_ir": 0, "trigger": "Program termination"}


def analyze_callgrind(name_value: str, receipt: dict[str, Any], plan: dict[str, Any],
                      source: dict[str, Any], stage: str, corpus: dict[str, Any],
                      gate: dict[str, Any], build: dict[str, Any]) -> dict[str, Any]:
    report = HERE / f"{name_value}.json"
    common = validate_common_receipt(receipt, report, plan, source, stage, corpus,
                                     "edit", "profile", build, gate)
    stem = HERE / f"{name_value}.callgrind"
    numbered = callgrind_numbered(stem)
    require(len(numbered) == plan["callgrind"]["expected_numbered_parts"],
            f"{name_value}: expected five numbered Callgrind parts")
    rows: list[dict[str, Any]] = []
    measured: list[Path] = []
    setup_callers: dict[str, int] = {}
    expected_setup = plan["callgrind"]["setup_parents"]
    for path in numbered:
        text = read_text(path)
        part = int(path.suffix[1:])
        label = path.name
        require(H.part_number(text, label) == part and "events: Ir" in text.splitlines(),
                f"{label}: raw Callgrind identity differs")
        require(H.trigger(text, label) == "--dump-after=" + OWNER,
                f"{label}: exact owner trigger differs")
        summary = H.summary_ir(text, label)
        cross_checked = PARSER_0707.edge(path, OWNER)
        require(cross_checked.get("calls") == 1
                and cross_checked.get("inclusive_ir") == summary,
                f"{label}: retained 0707 edge parser disagrees with raw summary")
        incoming = [edge for edge in PARSER.raw_edges(path)
                    if edge["callee"] == OWNER and edge["calls"] > 0
                    and edge["inclusive_ir"] > 0]
        require(incoming and sum(edge["calls"] for edge in incoming) == 1,
                f"{label}: exact owner incoming edge count differs")
        measured_edges = [edge for edge in incoming if edge["caller"] == MEASURED_PARENT]
        other_edges = [edge for edge in incoming if edge["caller"] != MEASURED_PARENT]
        if measured_edges:
            require(len(measured_edges) == 1 and not other_edges
                    and measured_edges[0]["inclusive_ir"] == summary,
                    f"{label}: measured owner edge is not an exact singleton")
            measured.append(path)
            role, caller = "measured", MEASURED_PARENT
        else:
            require(len(incoming) == 1, f"{label}: setup owner edge is split")
            caller = incoming[0]["caller"]
            require(caller in expected_setup, f"{label}: setup caller is unexpected: {caller}")
            setup_callers[caller] = setup_callers.get(caller, 0) + 1
            role = "setup"
        rows.append({"file": label, "part": part, "sha256": sha(path),
                     "summary_ir": summary, "role": role,
                     "retained_0707_owner_edge": cross_checked,
                     "raw_incoming_owner_edges": incoming, "selected_caller": caller})
    require(len(measured) == 1 and setup_callers == expected_setup,
            f"{name_value}: measured/setup caller counts differ: {setup_callers}")
    terminal = callgrind_terminal(stem, numbered)
    return {"common": common, "name": name_value, "stage": stage,
            "corpus_id": corpus["id"], "phase": "edit", "owner": OWNER,
            "owner_match": "exact", "measured_parent": MEASURED_PARENT,
            "measured_file": measured[0].name,
            "measured_ir": next(row["summary_ir"] for row in rows if row["role"] == "measured"),
            "setup_callers": setup_callers, "raw_parts": rows,
            "program_termination": terminal, "raw_parts_retained": True,
            "zero_ir_retained": True}


def parse_peak(path: Path) -> int:
    text = read_text(path)
    matches = PEAK_RE.findall(text)
    require(len(matches) == 1, f"{path.name}: expected one maximum RSS line")
    return int(matches[0])


def analyze_rss(name_value: str, receipt: dict[str, Any], plan: dict[str, Any],
                source: dict[str, Any], stage: str, corpus: dict[str, Any],
                phase: str, gate: dict[str, Any], build: dict[str, Any]) -> dict[str, Any]:
    report = HERE / f"{name_value}.json"
    common = validate_common_receipt(receipt, report, plan, source, stage, corpus,
                                     phase, "rss", build, gate)
    time_info = receipt.get("time_verbose")
    require(isinstance(time_info, dict), f"{name_value}: time -v custody is missing")
    path = HERE / str(time_info.get("path"))
    require(path.parent == HERE and path.is_file() and not path.is_symlink(),
            f"{name_value}: time -v output is missing or outside packet")
    require(time_info.get("sha256") == sha(path), f"{name_value}: time -v hash differs")
    peak = parse_peak(path)
    require(time_info.get("maximum_resident_set_size_kb") == peak,
            f"{name_value}: RSS peak differs from receipt")
    require(receipt.get("rss_peak_policy") ==
            "per-child maximum RSS; never summed across phases or repeats",
            f"{name_value}: RSS peak policy differs")
    return {"common": common, "name": name_value, "stage": stage,
            "corpus_id": corpus["id"], "phase": phase,
            "peak_rss_kb": peak,
            "time_verbose": {"file": path.name, "sha256": sha(path)},
            "peak_is_per_child": True, "peak_not_summed": True}


def compare_rss(rss_rows: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    comparisons: list[dict[str, Any]] = []
    repeats: list[dict[str, Any]] = []
    for corpus in (item["id"] for item in plan["corpora"]):
        for phase in plan["rss"]["phases"]:
            base_rows = [row for row in rss_rows if row["corpus_id"] == corpus
                         and row["phase"] == phase
                         and STAGE_BY_LABEL[row["stage"]]["source"] == "baseline"]
            cand_rows = [row for row in rss_rows if row["corpus_id"] == corpus
                         and row["phase"] == phase
                         and STAGE_BY_LABEL[row["stage"]]["source"] == "candidate"]
            require(len(base_rows) == 2 and len(cand_rows) == 2,
                    f"RSS repeat matrix is incomplete: {corpus}/{phase}")
            base_values = sorted(row["peak_rss_kb"] for row in base_rows)
            cand_values = sorted(row["peak_rss_kb"] for row in cand_rows)
            base_reference, cand_reference = base_values[0], cand_values[0]
            denominator = min(base_reference, cand_reference)
            delta = None if denominator == 0 else (cand_reference - base_reference) * 100.0 / denominator
            comparisons.append({"corpus_id": corpus, "phase": phase,
                                "baseline_peaks_kb": [row["peak_rss_kb"] for row in base_rows],
                                "candidate_peaks_kb": [row["peak_rss_kb"] for row in cand_rows],
                                "baseline_reference_kb": base_reference,
                                "candidate_reference_kb": cand_reference,
                                "reference_method": "lower_of_two_repeat_peaks",
                                "candidate_minus_baseline_percent": delta,
                                "flag_over_5_percent": delta is not None and abs(delta) > 5,
                                "peak_values_not_summed": True})
            for source, rows in (("baseline", base_rows), ("candidate", cand_rows)):
                ordered = sorted((row["stage"], row["peak_rss_kb"]) for row in rows)
                low, high = min(value for _, value in ordered), max(value for _, value in ordered)
                spread = None if low == 0 else (high - low) * 100.0 / low
                repeats.append({"corpus_id": corpus, "phase": phase, "source": source,
                                "stages": [stage for stage, _ in ordered],
                                "peaks_kb": [value for _, value in ordered],
                                "spread_percent": spread,
                                "flag_over_5_percent": spread is not None and spread > 5,
                                "diagnostic_only": True})
    return {"comparisons": comparisons, "repeat_diagnostics": repeats,
            "threshold_percent": plan["rss"]["peak_flag_percent"],
            "repeat_threshold_percent": plan["rss"]["repeat_flag_percent"],
            "peaks_never_summed": True}


def compare_profiles(rows: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    values: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for row in rows:
        source = STAGE_BY_LABEL[row["stage"]]["source"]
        values.setdefault((row["corpus_id"], source), []).append(row)
    output: list[dict[str, Any]] = []
    for (corpus, source), source_rows in sorted(values.items()):
        require(len(source_rows) == 2, f"Callgrind repeat matrix is incomplete: {corpus}/{source}")
        ordered = sorted(source_rows, key=lambda row: row["stage"])
        irs = [row["measured_ir"] for row in ordered]
        low, high = min(irs), max(irs)
        spread = None if low == 0 else (high - low) * 100.0 / low
        output.append({"corpus_id": corpus, "source": source,
                       "stages": [row["stage"] for row in ordered],
                       "measured_ir": irs, "spread_percent": spread,
                       "flag_over_5_percent": spread is not None and spread > 5,
                       "diagnostic_only": True})
    return {"repeat_diagnostics": output, "threshold_percent": 5}


def output_parity(profile_rows: list[dict[str, Any]], rss_rows: list[dict[str, Any]],
                  plan: dict[str, Any]) -> dict[str, Any]:
    all_rows = profile_rows + rss_rows
    parity: list[dict[str, Any]] = []
    for kind_rows, kind in ((profile_rows, "profile"), (rss_rows, "rss")):
        by_key: dict[tuple[str, str], list[dict[str, Any]]] = {}
        for row in kind_rows:
            by_key.setdefault((row["corpus_id"], row["phase"]), []).append(row)
        for key, values in sorted(by_key.items()):
            require(len(values) == 4, f"output parity matrix is incomplete: {key}")
            identities = [value["common"]["identity"] for value in values]
            require(all(identity == identities[0] for identity in identities[1:]),
                    f"baseline/candidate output parity failed: {key}")
            parity.append({"kind": kind, "corpus_id": key[0], "phase": key[1],
                           "stages": [value["stage"] for value in values], "equal": True})
    pilot_matches: list[dict[str, Any]] = []
    for row in all_rows:
        path = HERE / f"{row['stage']}-native-{row['corpus_id']}-{row['phase']}.json"
        require(path.is_file() and not path.is_symlink(),
                f"native pilot report is missing for {row['name']}")
        identity = M.report_identity(path)
        require(identity == row["common"]["identity"],
                f"native pilot output parity failed: {row['name']}")
        pilot_matches.append({"mechanism": row["name"], "pilot": path.name,
                              "pilot_sha256": sha(path), "equal": True})
    return {"stage_output_parity": parity, "pilot_matches": pilot_matches,
            "pilot_match_when_available": True, "all_outputs_match": True}


def analyze(plan: dict[str, Any]) -> dict[str, Any]:
    source = source_state(plan)
    gate = M.validate_pilot_gate(plan)
    corpus_by_id = fixture_map(plan)
    profile_rows: list[dict[str, Any]] = []
    rss_rows: list[dict[str, Any]] = []
    for stage in STAGE_BY_LABEL:
        build = build_info(plan, stage, source)
        for corpus in stage_order(plan, stage):
            profile = M.profile_name(stage, corpus)
            profile_receipt = read_json(HERE / f"{profile}.receipt.json")
            profile_rows.append(analyze_callgrind(profile, profile_receipt, plan, source,
                                                  stage, corpus, gate, build))
            for phase in plan["rss"]["phases"]:
                rss = M.rss_name(stage, corpus, phase)
                rss_receipt = read_json(HERE / f"{rss}.receipt.json")
                rss_rows.append(analyze_rss(rss, rss_receipt, plan, source, stage,
                                            corpus, phase, gate, build))
    require(len(profile_rows) == 8 and len(rss_rows) == 16,
            "mechanism child counts differ")
    parity = output_parity(profile_rows, rss_rows, plan)
    return {
        "schema_version": 1,
        "packet": plan["packet"],
        "revision": plan["revision"],
        "mechanism_plan_sha256": sha(PLAN_PATH),
        "mechanism_script_sha256": sha(MECHANISM_PATH),
        "source": {
            "baseline_manifest": source["baseline_path"],
            "candidate_manifest": source["candidate_path"],
            "baseline_manifest_sha256": source["baseline_sha256"],
            "candidate_manifest_sha256": source["candidate_sha256"],
            "changed_paths": source["changed_paths"],
            "current_checkout_is_candidate": True,
        },
        "pilot_gate": gate,
        "callgrind": {
            "children": len(profile_rows), "owner": OWNER, "owner_match": "exact",
            "measured_parent": MEASURED_PARENT, "raw_parts_retained": True,
            "program_termination_zero_ir_retained": True, "rows": profile_rows,
            "repeat_diagnostics": compare_profiles(profile_rows, plan),
        },
        "rss": {
            "children": len(rss_rows), "tool": "/usr/bin/time -v", "rows": rss_rows,
            "diagnostics": compare_rss(rss_rows, plan), "peaks_never_summed": True,
        },
        "output_parity": parity,
        "custody": {
            "constraints_sha256": sha(HERE / "constraints.json"),
            "normalized_cleanup_witness": normalized_cleanup_policy(),
            "full_command_source_binary_fixture_script_hashes_validated": True,
            "retained_callgrind_helpers": [
                "docs/performance/results/change-0707/analyze_profiles.py",
                "docs/performance/results/change-0709/analyze_profiles.py",
                "docs/performance/results/change-0705/analyze_profiles.py",
                "docs/performance/results/change-0521/analyze_profiles.py",
            ],
        },
        "claims": plan["claims"],
        "performance_claim": "mechanism-diagnostic-only; no optimization claim is emitted by this analyzer",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan-only", action="store_true",
                        help="validate helper and plan identities without reading captures")
    args = parser.parse_args()
    try:
        plan = load_plan()
        if args.plan_only:
            print(json.dumps({"packet": plan["packet"], "revision": plan["revision"],
                              "owner": OWNER, "profile_children": 8, "rss_children": 16,
                              "pilot_children": 32}, sort_keys=True))
            return 0
        report = analyze(plan)
        write_json(HERE / "mechanism-analysis.json", report)
        print("mechanism evidence analysis passed", flush=True)
        return 0
    except (EvidenceError, OSError, KeyError, ValueError) as error:
        print(f"analyze_mechanism.py: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
