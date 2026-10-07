#!/usr/bin/env python3
"""Deterministic compact reducer for the completed 0951 Callgrind packet.

All reads are rooted at the selected 0951 packet plus the explicitly retained
0944/0950 helper and build paths.  Captured absolute argv/cwd values are
historical metadata only; they are never used as locators.  Raw Callgrind and
capture packets remain external.  The output contains normalized descriptors,
revalidated custody, compact instruction tables, and explicit nested-overlap
flags; it makes no native performance or adoption claim.
"""
from __future__ import annotations
import hashlib
import importlib.util
import json
import os
import pathlib
import stat
import sys
from typing import Any

PACKET = pathlib.Path(__file__).resolve().parents[1]
P0944 = pathlib.Path("/home/zhuhe/litchi-performance-0944")
P0950 = pathlib.Path("/home/zhuhe/litchi-performance-0950")
PLAN_PATH = PACKET / "plan-0951-v1.json"
CAPTURE = PACKET / "tools" / "capture_0951_v1.py"
AGGREGATOR = PACKET / "tools" / "aggregate_current_audit_0951_v2.py"
CURRENT_HELPER = PACKET / "tools" / "attribute_current_audit_0951_v1.py"
EXPECTED_PLAN_SHA = "227322ea28ccadb206f00fee3efdc162db83e33431fbaa430d164b4e269b8bad"
EXPECTED_CAPTURE_SHA = "9c4ab50803150ecb8c05251ad6a484671878bbd60cbdd98148e166ee146a58da"
EXPECTED_AGGREGATOR_SHA = "bd73d26b95e6aafa9f2273b7873c10972dd8ed82ff49306b24f8ea5e86ef413e"
EXPECTED_CURRENT_HELPER_SHA = "c885ba8bc434a611693b192866b487dec2b54667b4cc73a82a2dd05f3d3fffb0"
OWNER = "litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream"
PREPARE = "litchi_opc::source_backed::SourceBackedPackage::prepare_topology_state"
WRITER = "litchi_opc::source_backed::SourceBackedPackage::write_changed_overlays_with_appended_inner"
TARGET = "litchi_opc::source_backed::audit_source_part_replacement_xml"


class EvidenceError(ValueError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def reject_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise EvidenceError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def load(path: pathlib.Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing regular JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot parse {path}: {error}") from error


def digest(path: pathlib.Path) -> tuple[int, str]:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    h = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block)
            size += len(block)
    return size, h.hexdigest()


def descriptor(path: pathlib.Path) -> dict[str, Any]:
    size, sha = digest(path)
    return {"path": str(path), "bytes": size, "sha256": sha,
            "mode": stat.S_IMODE(path.stat().st_mode),
            "executable": os.access(path, os.X_OK)}


def short_descriptor(path: pathlib.Path) -> dict[str, Any]:
    size, sha = digest(path)
    return {"path": str(path), "bytes": size, "sha256": sha}


def normalize_string(value: str) -> str:
    # These are the only roots allowed to be normalized.  Other absolute paths
    # stay visible in historical metadata and are never used for file lookup.
    roots = ((str(PACKET), "$P0951"), (str(P0950), "$P0950"),
             (str(P0944), "$P0944"))
    for root, token in sorted(roots, key=lambda item: len(item[0]), reverse=True):
        if value == root:
            return token
        if value.startswith(root + "/"):
            return token + value[len(root):]
    return value


def normalize(value: Any) -> Any:
    if isinstance(value, str):
        return normalize_string(value)
    if isinstance(value, list):
        return [normalize(item) for item in value]
    if isinstance(value, dict):
        return {key: normalize(item) for key, item in value.items()}
    return value


def normalized_descriptor(path: pathlib.Path) -> dict[str, Any]:
    return normalize(descriptor(path))


def resolve_declared(spec: Any, label: str, base: pathlib.Path = PACKET) -> pathlib.Path:
    require(isinstance(spec, dict), f"{label} descriptor is not an object")
    raw = spec.get("path")
    require(isinstance(raw, str) and raw and "\x00" not in raw, f"{label} descriptor path is invalid")
    path = pathlib.Path(raw) if raw.startswith("/") else base / raw
    path = path.resolve()
    actual = descriptor(path)
    for key in ("bytes", "sha256"):
        require(actual[key] == spec.get(key), f"{label} {key} differs")
    if "mode" in spec:
        require(actual["mode"] == spec["mode"], f"{label} mode differs")
    if "executable" in spec:
        require(actual["executable"] == spec["executable"], f"{label} executable bit differs")
    return path


def load_module(path: pathlib.Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot import helper {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def assert_file_hash(path: pathlib.Path, expected: str, label: str) -> None:
    require(digest(path)[1] == expected, f"{label} hash differs")


def verify_review_and_receipts(plan: dict[str, Any], cap: Any) -> dict[str, Any]:
    # The runner's existing preflight is the source/build/tool/marker custody
    # checker.  Its result is summarized; no workload is started here.
    preflight = cap.validate_plan()
    require(preflight.get("status") == "pass", "0951 validate-only preflight failed")
    assert_file_hash(PLAN_PATH, EXPECTED_PLAN_SHA, "frozen plan")
    assert_file_hash(CAPTURE, EXPECTED_CAPTURE_SHA, "frozen capture runner")
    assert_file_hash(AGGREGATOR, EXPECTED_AGGREGATOR_SHA, "frozen incoming-edge aggregator")
    assert_file_hash(CURRENT_HELPER, EXPECTED_CURRENT_HELPER_SHA, "frozen current-audit helper")

    def receipt(name: str) -> tuple[pathlib.Path, dict[str, Any], dict[str, Any]]:
        path = PACKET / "receipts" / name
        return path, load(path), normalized_descriptor(path)

    start_path, start, start_desc = receipt("root-start-v1.json")
    first_path, first, first_desc = receipt("root-first-child-go-v1.json")
    remaining_path, remaining, remaining_desc = receipt("root-remaining-capture-go-v1.json")
    reuse_path, reuse, reuse_desc = receipt("source-binary-reuse-v1.json")
    review_path, review, review_desc = receipt("comparison-tooling-review-v1.json")
    validation_path, validation, validation_desc = receipt("validate-only-v1.json")
    require(start.get("schema") == "litchi.performance.0951.root-start.v1" and start.get("cpu12_available") is True, "root-start receipt failed")
    require(start.get("changed_since_0950_start") == [], "root-start changed-input receipt failed")
    require(first.get("schema") == "litchi.performance.0951.root-first-child-go.v1" and first.get("status") == "authorized", "first-child GO receipt failed")
    require(first.get("first_child_before_rest") is True and first.get("no_retries") is True, "first-child GO schedule failed")
    require(remaining.get("schema") == "litchi.performance.0951.root-remaining-capture-go.v1" and remaining.get("status") == "authorized", "remaining GO receipt failed")
    require(remaining.get("indices") == [1, 2, 3] and remaining.get("serial") is True and remaining.get("retry") is False and remaining.get("source_or_gate_changes") is False, "remaining GO schedule failed")
    require(reuse.get("schema") == "litchi.performance.0951.source-binary-reuse.v1" and reuse.get("status") == "pass_with_declared_ignored_lock_exception", "source/binary reuse receipt failed")
    require(review.get("schema") == "litchi.xlsx-publication-instruction-profile-0951-comparison-tooling-review.v1" and review.get("status") == "GO", "comparison tooling review failed")
    require(validation.get("status") == "pass" and validation.get("workload_started") is False and validation.get("profiler_started") is False, "validate-only receipt failed")

    # Resolve every packet-local binding in the authorization receipts.  The
    # descriptors are custody checks; no captured absolute argv/cwd is opened.
    for key in ("plan", "runner", "postprocessor_not_part_of_capture", "readiness_review", "reuse"):
        resolve_declared(first[key], f"first-child {key}")
    resolve_declared(remaining["first_child_admission"], "remaining admission")
    resolve_declared(review["frozen_bindings"]["plan"], "review plan")
    resolve_declared(review["frozen_bindings"]["runner"], "review runner")
    resolve_declared(review["frozen_bindings"]["current_audit_helper"], "review current-audit helper")
    resolve_declared(review["frozen_bindings"]["source_binary_reuse_receipt"], "review source/binary receipt")
    resolve_declared(review["frozen_bindings"]["validate_only_receipt"], "review validate-only receipt")
    resolve_declared(review["frozen_bindings"]["native_context"], "review native context")
    resolve_declared(review["frozen_bindings"]["current_audit_nm_output"], "review nm output")
    resolve_declared(review["retained_0944_admission_controls"]["historical_control_receipt"], "0944 control receipt", P0944)
    controls = review["retained_0944_admission_controls"]
    control_refs = []
    for polarity in ("positive", "negative"):
        for item in controls[polarity]:
            refs = {}
            for key in ("stderr", "stdout", "scratch_plan"):
                if key in item:
                    refs[key] = normalized_descriptor(resolve_declared(item[key], f"control {item['name']} {key}"))
            control_refs.append({"polarity": polarity, "name": item["name"], "status": item["status"], "expected_rejection": item.get("expected_rejection", False), "artifacts": refs})
    source_manifest = plan["source_binding"]["manifest"]
    reuse_source = reuse["source_manifest"]
    require(reuse_source["sha256"] == source_manifest["sha256"] and reuse_source["file_count"] == 10790 and reuse_source["bytes"] == 260102337, "source reuse manifest binding differs")
    native_reuse = reuse["build_reuse"]["native"]
    native_path = P0950 / native_reuse["artifact"]["path"]
    native_actual = short_descriptor(native_path)
    require(native_actual["bytes"] == native_reuse["artifact"]["bytes"] and native_actual["sha256"] == native_reuse["artifact"]["sha256"], "retained 0950 native binary differs")
    require(native_actual["path"] == plan["binary"]["path"], "retained native path differs from plan")
    require(native_reuse["source_commit"] == plan["source_binding"]["base_commit"] and native_reuse["terminal_returncode"] == 0 and native_reuse["timed_out"] is False and native_reuse["process_group_clear"] is True, "retained native build custody differs")
    head_compare = reuse["git_object_comparison"]
    require(head_compare["manifest_paths"] == 10790 and head_compare["committed_git_object_matches"] == 10789 and head_compare["declared_ignored_missing"] == ["Cargo.lock"] and head_compare["unexpected_missing_committed_paths"] == [], "HEAD byte-equivalence reference differs")
    require(head_compare["strict_all_files_committed_match"] is False, "ignored-lock exception was not retained")
    return {
        "runner_preflight": {"status": preflight["status"], "source": normalize(preflight["source"]), "build": normalize(preflight["build"]), "helpers": normalize(preflight["helpers"]), "markers": normalize(preflight["markers"])},
        "receipts": {"root_start": start_desc, "first_child_go": first_desc, "remaining_go": remaining_desc, "source_binary_reuse": reuse_desc, "comparison_review": review_desc, "validate_only": validation_desc},
        "review": {"status": review["status"], "decision": review["decision"], "known_limits": review["known_limits"], "controls": control_refs},
        "current_0950_binary": {"descriptor": normalize(plan["binary"]), "reuse_artifact": normalize(native_reuse["artifact"]), "source_commit": native_reuse["source_commit"], "build_terminal_sha256": native_reuse["terminal_sha256"], "source_manifest": normalize(reuse_source)},
        "root_head_byte_equivalence": {"current_head": reuse["current_head"], "status": reuse["status"], "committed_git_object_matches": head_compare["committed_git_object_matches"], "manifest_paths": head_compare["manifest_paths"], "matched_committed_bytes": head_compare["matched_committed_bytes"], "matched_committed_rows_digest": head_compare["matched_committed_rows_digest"], "head_tree_listing_sha256": head_compare["head_tree_listing_sha256"], "declared_ignored_missing": head_compare["declared_ignored_missing"], "strict_all_files_committed_match": head_compare["strict_all_files_committed_match"], "ignored_lock_binding": normalize(reuse["ignored_lock_binding"])},
    }


def verify_artifacts(cap: Any, folder: pathlib.Path, terminal: dict[str, Any], command: dict[str, Any]) -> dict[str, Any]:
    all_artifacts = cap.artifact_list(folder)
    terminal_expected = [item for item in all_artifacts if pathlib.Path(item["path"]).name not in {"semantics.json", "raw.json", "command.json"}]
    command_expected = [item for item in all_artifacts if pathlib.Path(item["path"]).name != "command.json"]
    require(terminal.get("artifacts") == terminal_expected, f"{folder.name}: terminal artifact descriptors differ")
    require(command.get("artifacts") == command_expected, f"{folder.name}: command artifact descriptors differ")
    return {"terminal": [normalize(item) for item in terminal_expected], "command": [normalize(item) for item in command_expected]}


def verify_historical_argv(plan: dict[str, Any], row: dict[str, Any], folder: pathlib.Path, start: dict[str, Any], terminal: dict[str, Any], command: dict[str, Any]) -> dict[str, Any]:
    require(start.get("argv") == terminal.get("argv") == command.get("argv"), f"{folder.name}: captured argv differs across custody receipts")
    argv = start["argv"]
    require(plan["binary"]["path"] in argv, f"{folder.name}: binary absent from captured argv")
    require("--warmup" in argv and argv[argv.index("--warmup") + 1] == "0", f"{folder.name}: warmup metadata differs")
    require("--samples" in argv and argv[argv.index("--samples") + 1] == "1", f"{folder.name}: sample metadata differs")
    require("--case" in argv and argv[argv.index("--case") + 1] == plan["case"], f"{folder.name}: case metadata differs")
    require("--xlsx-cell-crud-shape" in argv and argv[argv.index("--xlsx-cell-crud-shape") + 1] == row["shape"], f"{folder.name}: shape metadata differs")
    callgrind_arg = next((item for item in argv if item.startswith("--callgrind-out-file=")), "")
    json_arg = argv[argv.index("--json") + 1] if "--json" in argv else ""
    require(callgrind_arg.endswith(str(folder / "callgrind.out")) and json_arg == str(folder / "report.json"), f"{folder.name}: captured artifact argv differs")
    require(start.get("cwd") == plan["source_binding"]["root"]["path"] and terminal.get("cwd") == start.get("cwd") and command.get("cwd") == start.get("cwd"), f"{folder.name}: captured cwd differs")
    require(start.get("environment") == plan["environment"] and terminal.get("environment") == plan["environment"] and command.get("environment") == plan["environment"], f"{folder.name}: captured environment differs")
    return {"argv": normalize(argv), "cwd": normalize(start["cwd"]), "environment": start["environment"], "started_utc": start.get("started_utc"), "terminal_utc": terminal.get("terminal_utc"), "pid": terminal.get("pid"), "process_group_id": terminal.get("process_group_id")}


def compact_nested(raw: dict[str, Any], incoming: dict[str, Any]) -> dict[str, Any]:
    retained = []
    for item in raw.get("nested_diagnostics", {}).get("rows", []):
        fn = item.get("function", {})
        retained.append({"name": fn.get("name"), "self_ir": fn.get("self_ir"), "direct_children_ir": fn.get("direct_children_ir"), "incoming_ir": item.get("incoming_ir"), "incoming_calls": item.get("incoming_calls"), "overlap": True, "sum_allowed": False, "implementation_note": "retained 0944 nested name; not the current direct audit symbol"})
    current = []
    for fn in incoming["target_functions"]:
        current.append({"name": fn["name"], "self_ir": fn["self_ir"], "direct_children_ir": fn["direct_children_ir"], "inclusive_ir": fn["self_ir"] + fn["direct_children_ir"], "incoming_edges": normalize(incoming["incoming_edges"]), "overlap": True, "sum_allowed": False})
    multi = []
    for fn in incoming["multi_window_functions"]:
        multi.append({"name": fn["name"], "self_ir": fn["self_ir"], "direct_children_ir": fn["direct_children_ir"], "inclusive_ir": fn["self_ir"] + fn["direct_children_ir"], "overlap": True, "sum_allowed": False})
    paths = [{"functions": item["functions"], "edges": [{"callee": e["callee"], "calls": e["calls"], "inclusive_ir": e["inclusive_ir"]} for e in item["edges"]]} for item in incoming["target_to_multi_paths"]]
    return {"retained_0944_rows": retained, "current_audit": current, "multi_window": multi, "current_audit_to_multi_paths": paths, "rules":{"all_nested_rows_overlap": True,"nested_sum_allowed": False,"current_audit_incoming_edges_conserved": incoming["conservation"]["equal"],"replicates_must_remain_separate": True}}


def accounting_row(raw: dict[str, Any], incoming: dict[str, Any]) -> dict[str, Any]:
    measured = raw["measured"]
    topology = raw["topology"]
    topology_in = topology["incoming_from_publication"]
    require(len(topology_in) == 1, "topology incoming edge is ambiguous")
    topology_ir = topology_in[0]["inclusive_ir"]
    prep_rows = [item for item in topology["immediate_children"] if item["callee"] == PREPARE]
    writer_rows = [item for item in topology["immediate_children"] if item["callee"] == WRITER]
    require(len(prep_rows) == 1 and len(writer_rows) == 1, "preparation/writer direct child is ambiguous")
    preparation_ir = prep_rows[0]["inclusive_ir"]
    writer_ir = writer_rows[0]["inclusive_ir"]
    publication_ir = measured["summary_ir"]
    current_audit_ir = incoming["conservation"]["incoming_edge_ir"]
    require(publication_ir == measured["function"]["self_ir"] + measured["function"]["direct_children_ir"], "publication conservation failed")
    require(topology_ir == topology["function"]["self_ir"] + topology["function"]["direct_children_ir"], "topology conservation failed")
    require(preparation_ir >= current_audit_ir, "current audit exceeds preparation boundary")
    publication_residual = publication_ir - topology_ir
    preparation_residual = preparation_ir - current_audit_ir
    topology_residual = topology_ir - writer_ir - preparation_ir
    require(publication_residual >= 0 and preparation_residual >= 0 and topology_residual >= 0, "negative residual")
    return {"publication_owner_ir": publication_ir, "topology_ir": topology_ir, "preparation_ir": preparation_ir, "current_audit_ir": current_audit_ir, "writer_ir": writer_ir, "publication_boundary_residual_ir": publication_residual, "preparation_boundary_residual_ir": preparation_residual, "topology_other_children_and_self_ir": topology_residual, "publication_owner_call_count": measured["selected"]["calls"], "topology_call_count": topology_in[0]["calls"], "preparation_call_count": prep_rows[0]["calls"], "current_audit_call_count": incoming["conservation"]["incoming_calls"], "writer_call_count": writer_rows[0]["calls"], "relations":{"publication_residual_disjoint_from_topology": True,"writer_and_preparation_are_sibling_boundary_rows": True,"current_audit_is_nested_in_preparation": True,"current_audit_overlaps_preparation_and_publication": True,"current_audit_not_added_to_writer_or_residual": True}}


def verify_child(cap: Any, agg_mod: Any, plan: dict[str, Any], row: dict[str, Any]) -> dict[str, Any]:
    folder = PACKET / "capture" / f"{row['index']:02d}-{row['shape']}"
    require(folder.is_dir(), f"missing capture folder: {folder}")
    start_path, terminal_path, command_path = folder/"start.json", folder/"terminal.json", folder/"command.json"
    report_path, semantics_path, raw_path = folder/"report.json", folder/"semantics.json", folder/"raw.json"
    start, terminal, command = load(start_path), load(terminal_path), load(command_path)
    report, stored_semantic, stored_raw = load(report_path), load(semantics_path), load(raw_path)
    runner_desc = descriptor(CAPTURE); plan_desc = descriptor(PLAN_PATH)
    require(start.get("schema") == "litchi.xlsx-publication-instruction-profile-child-start-0951.v1", f"{folder.name}: start schema differs")
    require(terminal.get("schema") == "litchi.xlsx-publication-instruction-profile-child-terminal-0951.v1", f"{folder.name}: terminal schema differs")
    require(command.get("schema") == "litchi.xlsx-publication-instruction-profile-command-0951.v1", f"{folder.name}: command schema differs")
    for doc, label in ((start,"start"),(terminal,"terminal"),(command,"command")):
        require(doc.get("row") == row and doc.get("plan") == plan_desc and doc.get("runner") == runner_desc and doc.get("binary") == plan["binary"], f"{folder.name}: {label} custody binding differs")
    require(command.get("helper") == runner_desc, f"{folder.name}: command helper differs")
    require(terminal.get("returncode") == 0 and terminal.get("timed_out") is False and terminal.get("launch_error") is None and terminal.get("artifact_error") is None and terminal.get("process_group_clear") is True and terminal.get("terminal_before_parse") is True, f"{folder.name}: terminal gate failed")
    require(command.get("terminal_receipt") == descriptor(terminal_path) and command.get("semantics") == descriptor(semantics_path) and command.get("raw") == descriptor(raw_path), f"{folder.name}: command receipt descriptors differ")
    artifact_custody = verify_artifacts(cap, folder, terminal, command)
    historical = verify_historical_argv(plan, row, folder, start, terminal, command)
    semantic_mod = load_module(P0944/"tools/semantic_v2.py", f"p951_reduce_semantic_{row['index']}")
    semantic_mod.PACKET = PACKET
    recalculated_semantic = semantic_mod.validate(plan, report, row["shape"])
    require(recalculated_semantic == stored_semantic, f"{folder.name}: semantic replay differs")
    raw_mod = load_module(P0944/"tools/raw_analyzer_v2.py", f"p951_reduce_raw_{row['index']}")
    raw_mod.PACKET = PACKET; raw_mod.MOD.HERE = PACKET
    recalculated_raw = raw_mod.analyze(PLAN_PATH, folder, row)
    require(recalculated_raw == stored_raw, f"{folder.name}: raw replay differs")
    raw = stored_raw
    incoming_path = PACKET/"analysis"/f"current-audit-{row['index']:02d}-incoming-v2.json"
    stored_incoming = load(incoming_path)
    temp = pathlib.Path(f"/tmp/p951-reduce-incoming-{row['index']}.json")
    if temp.exists(): temp.unlink()
    recalculated_incoming = agg_mod.aggregate(PLAN_PATH, folder, temp)
    require(recalculated_incoming == stored_incoming, f"{folder.name}: incoming-edge replay differs")
    if temp.exists(): temp.unlink()
    incoming = stored_incoming
    require(incoming.get("status") == "pass" and incoming.get("shape") == row["shape"] and incoming.get("index") == row["index"], f"{folder.name}: incoming analysis identity differs")
    top20 = raw["measured_exclusive_self"]
    top_rows = [{"rank": index + 1, "name": item["name"], "self_ir": item["self_ir"], "file": normalize_string(item.get("file", ""))} for index, item in enumerate(top20["rows"][:20])]
    return {"index": row["index"], "repeat": row["repeat"], "shape": row["shape"], "custody":{"start":normalized_descriptor(start_path),"terminal":normalized_descriptor(terminal_path),"command":normalized_descriptor(command_path),"report":normalized_descriptor(report_path),"semantics":normalized_descriptor(semantics_path),"raw":normalized_descriptor(raw_path),"terminal_artifacts":artifact_custody["terminal"],"command_artifacts":artifact_custody["command"]},"historical_metadata":historical,"semantic":{"status":stored_semantic["status"],"shape":stored_semantic["shape"],"reopen_ns":stored_semantic["reopen_ns"],"preservation":stored_semantic["preservation"],"timing_policy":stored_semantic["timing_policy"]},"publication_table":accounting_row(raw, incoming),"exclusive_top20":{"summary_ir":top20["summary_ir"],"all_exclusive_self_ir":top20["all_self_ir"],"top20_exclusive_self_ir":top20["top_self_ir"],"other_exclusive_self_ir":top20["other_self_ir"],"rows":top_rows},"nested_overlap":compact_nested(raw, incoming),"revalidation":{"terminal":True,"semantic":True,"raw_callgrind":True,"current_audit_incoming_edges":True,"no_retry_or_replacement":True}}


def reduce_packet(output: pathlib.Path) -> dict[str, Any]:
    plan = load(PLAN_PATH)
    require(plan.get("schema") == "litchi.xlsx-publication-instruction-profile-0951.v1", "plan schema differs")
    cap = load_module(CAPTURE, "p951_reduce_capture")
    static = verify_review_and_receipts(plan, cap)
    agg_mod = load_module(AGGREGATOR, "p951_reduce_aggregator")
    captures = [verify_child(cap, agg_mod, plan, row) for row in plan["schedule"]]
    first = load(PACKET/"first-child-admission.json")
    require(first.get("status") == "pass" and first.get("index") == 0, "first-child admission is not passing")
    require(first.get("plan") == descriptor(PLAN_PATH) and first.get("runner") == descriptor(CAPTURE) and first.get("binary") == plan["binary"], "first-child admission binding differs")
    admission = {"status": first["status"], "index": first["index"], "row": first["row"], "terminal": normalized_descriptor(PACKET/"capture/00-medium/terminal.json"), "report": normalized_descriptor(PACKET/"capture/00-medium/report.json"), "semantics": normalized_descriptor(PACKET/"capture/00-medium/semantics.json"), "raw": normalized_descriptor(PACKET/"capture/00-medium/raw.json"), "validation": first["validation"]}
    original_current = PACKET/"analysis/current-audit-00.json"
    require(original_current.is_file(), "first current-audit postprocessor output is missing")
    current_doc = load(original_current); require(current_doc.get("status") == "pass" and current_doc.get("observed", {}).get("current_audit"), "first current-audit postprocessor did not pass")
    result = {"schema":"litchi.xlsx-publication-profile-reduction-0951.v1","status":"pass","scope":"Descriptive Callgrind Ir attribution for four source-backed primary publication children; no native latency/p99, allocation/RSS, speedup, adoption, or source-change conclusion.","frozen_bindings":{"plan":normalized_descriptor(PLAN_PATH),"runner":normalized_descriptor(CAPTURE),"incoming_aggregator":normalized_descriptor(AGGREGATOR),"current_audit_helper":normalized_descriptor(CURRENT_HELPER),"native_context":normalize(plan["native_context"]),"binary":normalize(plan["binary"]),"source_manifest":normalize(plan["source_binding"]["manifest"])},"preflight":static,"first_child_admission":admission,"children":captures,"determinism":{"schedule":[{"index":x["index"],"repeat":x["repeat"],"shape":x["shape"]} for x in captures],"serial":True,"retries":0,"all_raw_and_semantic_replays_equal":True,"all_current_audit_incoming_replays_equal":True,"large_raw_artifacts_external":True},"normalization":{"declared_roots":{"0951": "$P0951","0950":"$P0950","0944":"$P0944"},"captured_absolute_argv_and_cwd":"historical metadata only; never used as filesystem locators","unrecognized_absolute_paths":"retained verbatim; no broad path rewriting"},"limitations":["Exclusive top-20 rows are disjoint self-Ir rows within each selected measured dump.","Publication/preparation/writer boundary rows are reported with their explicit nesting relations; current audit and MultiWindowSearch rows overlap and are never summed.","The current direct audit implementation is derived from the post-raw audit_source_part_replacement_xml graph; the retained 0944 validate_source_part_replacement_xml row is labeled historical and is not used as the current claim.","The 0950 source/binary custody receipt records one ignored root Cargo.lock exception: 10,789 committed Git objects plus the declared ignored lock make the 10,790-file closure."]}
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return result


def main() -> int:
    target = pathlib.Path(sys.argv[1]) if len(sys.argv) == 2 else PACKET/"analysis/reduction-0951-v1.json"
    try:
        result = reduce_packet(target.resolve())
    except (EvidenceError, OSError, ImportError, json.JSONDecodeError) as error:
        print(f"reduce_0951_v1: {error}", file=sys.stderr)
        return 2
    print(json.dumps({"status":result["status"],"output":str(target.resolve()),"bytes":target.resolve().stat().st_size,"children":len(result["children"])},sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
