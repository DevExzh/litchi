#!/usr/bin/env python3
"""Independent offline controls for the frozen 0946 attribution reducer.

The script consumes only copied retained inputs.  It validates the positive
reduction and its moved-input replay, then runs bounded negative controls in
fresh external copies.  It never starts a workload or profiler and never
writes to the retained packet.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Any

EXPECTED_TOOL_SHA = "9fb08db7c5ece8f3c466829651c727b8599a26cca8ae1fcff4a36ac3d7fda6de"
EXPECTED_PLAN_SHA = "9c9700f1798369244bee67b074410c92abfbeb31cfcd163ba52fac94d372dfc9"
EXPECTED_OUTPUT_SHA = "e0ca6ecd41879a1bdfcb0530d76dfae8a92cfe7b87365e91cf7e5ea64dd27a4b"
TARGETS = [
    "litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream",
    "litchi_opc::source_backed::SourceBackedPackage::write_topology_to_stream",
    "litchi_opc::source_backed::validate_source_part_replacement_xml",
    "xml_minifier::audit::verify_source_replacement",
    "xml_minifier::audit::verify_with_policy",
    "xml_minifier::audit::check_start",
    "xml_minifier::audit::inspect_attributes",
    "xml_minifier::audit::check_end",
    "xml_minifier::audit::wellformed::scan_qname",
    "quick_xml::reader::Reader<R>::read_event_impl",
    "quick_xml::reader::state::ReaderState::emit_start",
    "quick_xml::reader::state::ReaderState::emit_end",
]


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def read_json(path: Path) -> Any:
    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result
    return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=unique)


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def strict_relative(root: Path, value: str) -> Path:
    path = Path(value)
    require(not path.is_absolute() and ".." not in path.parts and "." not in path.parts,
            f"unsafe packet path {value!r}")
    current = root
    for component in path.parts:
        current /= component
        require(not current.is_symlink(), f"symlink packet path {value!r}")
    require(current.is_file(), f"missing packet file {value!r}")
    return current


def run_helper(tree: Path, output: Path) -> dict[str, Any]:
    command = [sys.executable, "-B", str(tree / "tools/reduce.py"), "--output", str(output)]
    env = dict(os.environ)
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    completed = subprocess.run(command, cwd=tree, env=env, text=True,
                               stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                               check=False)
    return {
        "argv": command,
        "cwd": str(tree),
        "returncode": completed.returncode,
        "stdout": completed.stdout,
        "stderr": completed.stderr,
        "output_exists": output.exists(),
        "output": str(output),
    }


def copy_case(base: Path, cases: Path, name: str) -> Path:
    case = cases / name
    if case.exists():
        shutil.rmtree(case)
    shutil.copytree(base, case, symlinks=False)
    return case


def check_source_ids(tree: Path, plan: dict[str, Any], result: dict[str, Any]) -> dict[str, Any]:
    focused = read_json(tree / "inputs/focused-source-hashes.json")
    bindings = plan["source_bindings"]
    require(set(bindings) == set(focused), "source binding keys differ from focused source receipt")
    for key, expected in focused.items():
        require(bindings[key] == expected, f"source binding differs for {key}")
    source = result["source_and_binary"]
    require(source["0944_xml_audit_sha256"] == bindings["audit_rs"]["sha256"],
            "output 0944 audit hash differs from plan")
    require(source["current_xml_audit_sha256"] == bindings["current_checkout_audit_rs"]["sha256"],
            "output current audit hash differs from plan")
    require(source["xml_audit_source_unchanged"] is True, "source unchanged flag is false")
    require(source["current_native_or_0945_sharing_claim"] is False,
            "output makes an unbound native/sharing claim")
    return {"focused_source_hashes_sha256": sha256(tree / "inputs/focused-source-hashes.json"),
            "source_binding_keys": sorted(bindings)}


def validate_positive(tree: Path, output: Path, author_output: Path | None) -> dict[str, Any]:
    plan = read_json(tree / "plan-v2.json")
    result = read_json(output)
    require(plan["schema"] == "litchi.xlsx-publication-attribution-0946.plan.v2", "plan schema")
    require(plan["status"] == "frozen", "plan status")
    require(plan["targets"] == TARGETS, "target list")
    require(sha256(tree / "tools/reduce.py") == EXPECTED_TOOL_SHA, "helper hash")
    require(sha256(tree / "plan-v2.json") == EXPECTED_PLAN_SHA, "plan hash")
    require(result["status"] == "pass", "reduction status")
    require(result["tool"]["sha256"] == EXPECTED_TOOL_SHA, "output helper binding")
    require(result["plan"]["sha256"] == EXPECTED_PLAN_SHA, "output plan binding")
    require(sha256(output) == EXPECTED_OUTPUT_SHA, "independent output hash")
    rows = plan["captures"]
    require([(row["shape"], row["repeat"]) for row in rows]
            == [("medium", 0), ("dense-sparse", 0), ("dense-sparse", 1), ("medium", 1)],
            "frozen four-row schedule")
    require([row["index"] for row in rows] == [0, 1, 2, 3], "frozen indices")
    require(len(result["captures"]) == 4, "result capture count")
    prior = read_json(tree / "inputs/root-reduction-v1.json")
    prior_by_index = {int(row["index"]): row for row in prior["captures"]}
    for row, captured in zip(rows, result["captures"]):
        path = strict_relative(tree, row["input_path"])
        require(path.stat().st_size == row["bytes"], f"input bytes {path}")
        require(sha256(path) == row["sha256"], f"input hash {path}")
        require(captured["index"] == row["index"] and captured["shape"] == row["shape"]
                and captured["repeat"] == row["repeat"], "capture identity")
        require(captured["input"]["path"] == row["input_path"], "relative input path")
        require(captured["input"]["bytes"] == row["bytes"]
                and captured["input"]["sha256"] == row["sha256"], "input descriptor")
        require(set(captured["targets"]) == set(TARGETS), "target set")
        prior_row = prior_by_index[row["index"]]
        require(captured["header"]["summary_ir"] == prior_row["exclusive_self"]["all_self_ir"],
                "prior publication denominator")
        require(captured["measured_owner_edge"]["calls"] == 1,
                "owner edge call count")
        for name, group in captured["targets"].items():
            require(group["record_count"] == len(group["records"]), f"record count {name}")
            require(group["exclusive_self_ir"] + group["direct_children_ir"]
                    == group["inclusive_ir"], f"group conservation {name}")
            for record in group["records"]:
                require(record["exclusive_self_ir"] + record["direct_children_ir"]
                        == record["inclusive_ir"], f"record conservation {name}")
                census = record["edge_census"]
                for field in ("direct_count", "direct_zero_cost_count", "direct_zero_call_count",
                              "incoming_count", "incoming_zero_cost_count", "incoming_zero_call_count"):
                    require(isinstance(census[field], int) and census[field] >= 0,
                            f"edge census {name}/{field}")
                require(record["validation"]["exclusive_self_plus_all_direct_edges_equals_all_incoming_edges"],
                        f"record validation {name}")
    source_ids = check_source_ids(tree, plan, result)
    author_match = None
    if author_output is not None:
        author_match = {"path": str(author_output), "sha256": sha256(author_output),
                        "byte_equal": output.read_bytes() == author_output.read_bytes()}
        require(author_match["sha256"] == EXPECTED_OUTPUT_SHA and author_match["byte_equal"],
                "author v2 output differs from independent replay")
    return {"plan_sha256": sha256(tree / "plan-v2.json"),
            "helper_sha256": sha256(tree / "tools/reduce.py"),
            "output_sha256": sha256(output),
            "output_bytes": output.stat().st_size,
            "capture_count": len(result["captures"]),
            "target_count": len(TARGETS),
            "source_ids": source_ids,
            "author_output_match": author_match}


def mutate_json(path: Path, mutate) -> None:
    data = read_json(path)
    mutate(data)
    path.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--base", type=Path, required=True,
                        help="copied frozen input tree containing tools/, plan-v2.json, inputs/")
    parser.add_argument("--moved", type=Path, required=True,
                        help="independent moved copy of the same copied input tree")
    parser.add_argument("--original-output", type=Path, required=True)
    parser.add_argument("--moved-output", type=Path, required=True)
    parser.add_argument("--author-output", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    base = args.base.resolve()
    moved = args.moved.resolve()
    cases = args.output.resolve().parent / "control-cases-v1"
    if cases.exists():
        shutil.rmtree(cases)
    cases.mkdir(parents=True)

    positive = validate_positive(base, args.original_output.resolve(), args.author_output.resolve())
    moved_positive = validate_positive(moved, args.moved_output.resolve(), None)
    require(args.original_output.read_bytes() == args.moved_output.read_bytes(),
            "original and moved reduction bytes differ")

    negatives: dict[str, Any] = {}
    case = copy_case(base, cases, "hash-mismatch")
    target = case / "inputs/callgrind/00-medium-r0.out.6"
    target.write_bytes(target.read_bytes() + b"\n")
    negative = run_helper(case, case / "negative-output.json")
    require(negative["returncode"] != 0 and not negative["output_exists"],
            "hash mismatch was accepted")
    negatives["hash_mismatch"] = negative

    case = copy_case(base, cases, "duplicate-index")
    mutate_json(case / "plan-v2.json", lambda d: d["captures"][1].__setitem__("index", 0))
    negative = run_helper(case, case / "negative-output.json")
    require(negative["returncode"] != 0 and not negative["output_exists"],
            "duplicate schedule index was accepted")
    negatives["duplicate_schedule_index"] = negative

    case = copy_case(base, cases, "path-traversal")
    mutate_json(case / "plan-v2.json",
                lambda d: d["captures"][0].__setitem__("input_path", "../outside.out"))
    negative = run_helper(case, case / "negative-output.json")
    require(negative["returncode"] != 0 and not negative["output_exists"],
            "path traversal was accepted")
    negatives["path_traversal"] = negative

    case = copy_case(base, cases, "descriptor-corruption")
    target = case / "inputs/reader_v2.py"
    original = target.read_bytes()
    target.write_bytes(original[:-1] + bytes([original[-1] ^ 1]))
    negative = run_helper(case, case / "negative-output.json")
    require(negative["returncode"] != 0 and not negative["output_exists"],
            "descriptor corruption was accepted")
    negatives["descriptor_corruption"] = negative

    case = copy_case(base, cases, "symlink-input")
    link = case / "inputs/callgrind/link.out.6"
    link.symlink_to("00-medium-r0.out.6")
    mutate_json(case / "plan-v2.json",
                lambda d: d["captures"][0].__setitem__("input_path", "inputs/callgrind/link.out.6"))
    negative = run_helper(case, case / "negative-output.json")
    require(negative["returncode"] != 0 and not negative["output_exists"],
            "symlink input was accepted")
    negatives["symlink_input"] = negative

    case = copy_case(base, cases, "overwrite-output")
    output = case / "same-output.json"
    first = run_helper(case, output)
    second = run_helper(case, output)
    require(first["returncode"] == 0 and first["output_exists"], "overwrite setup positive failed")
    require(second["returncode"] != 0 and second["output_exists"], "overwrite was accepted")
    negatives["output_overwrite"] = {"first": first, "second": second}

    receipt = {
        "schema": "litchi.xlsx-publication-attribution-0946.independent-controls.v1",
        "status": "pass",
        "scope": "Offline validation of the frozen reducer and copied retained Callgrind inputs; no workload, build, or profiler was run.",
        "command": "PYTHONDONTWRITEBYTECODE=1 python3 -B review/independent-controls-v1.py --base review/original-input-copy --moved review/moved-input-copy --original-output review/reduction-original-independent-v1.json --moved-output review/reduction-moved-independent-v1.json --author-output attribution-summary-v2.json --output review/independent-controls-v1.json",
        "bindings": {
            "helper_sha256": EXPECTED_TOOL_SHA,
            "plan_sha256": EXPECTED_PLAN_SHA,
            "independent_output_sha256": EXPECTED_OUTPUT_SHA,
            "independent_output_bytes": args.original_output.stat().st_size,
            "author_output_sha256": sha256(args.author_output),
            "author_output_bytes": args.author_output.stat().st_size,
            "original_moved_byte_equal": True,
            "original_moved_sha256": sha256(args.original_output),
        },
        "positive": {"original": positive, "moved": moved_positive},
        "negative_controls": {
            name: {"returncode": value.get("returncode", value.get("second", {}).get("returncode")),
                   "rejected": (value.get("returncode", value.get("second", {}).get("returncode")) != 0)}
            for name, value in negatives.items()
        },
        "negative_details": negatives,
        "retained_history": {
            "draft_output_v1_sha256": "fbb1cf5308e9ce98acd77eb594790c01dda774fda33bb3357cbb3a2248c460e5",
            "draft_output_v1_preserved": True,
            "raw_json_replay": "not performed; this review binds the retained part-6 Callgrind descriptors and prior bounded totals, per final scope",
        },
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    require(not args.output.exists(), f"refusing to overwrite receipt: {args.output}")
    args.output.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "output": str(args.output),
                      "output_sha256": sha256(args.output),
                      "output_bytes": args.output.stat().st_size,
                      "negative_count": len(negatives)}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
