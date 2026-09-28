"""Offline replay of the six operation-scoped Callgrind captures."""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

import analysis_common as c
import native_analysis as native


PROFILE_ORDERS = (
    ("tiny", "medium", "large"),
    ("large", "medium", "tiny"),
)


def _build(plan: dict[str, Any]) -> dict[str, Any]:
    _, build = native._build()
    profile = plan.get("profile")
    c.require(isinstance(profile, dict), "profile plan is missing")
    c.require(profile.get("collect_at_start") is False
              and profile.get("events") == ["Ir"]
              and profile.get("repeats") == 2
              and profile.get("samples") == 1
              and profile.get("warmup") == 0,
              "Callgrind policy changed")
    c.require(profile.get("orders") == [list(order) for order in PROFILE_ORDERS],
              "Callgrind shape order changed")
    c.require(profile.get("expected_numbered_parts") == 1,
              "Callgrind numbered-part policy changed")
    c.require(plan.get("owner") == c.OWNER and plan.get("cpu") == 12,
              "Callgrind owner or CPU changed")
    return build


def _artifacts(row: dict[str, Any], stem: str) -> dict[str, Path]:
    values = row.get("artifacts")
    expected = {f"{stem}.json", f"{stem}.log", f"{stem}.callgrind",
                f"{stem}.callgrind.1"}
    c.require(isinstance(values, dict) and set(values) == expected,
              f"{stem}: artifact set changed")
    paths: dict[str, Path] = {}
    for name in sorted(expected):
        c.require(Path(name).name == name, f"{stem}: artifact key escapes packet")
        path = c.artifact(values[name], f"{stem} {name}")
        c.require(path.name == name, f"{stem}: artifact basename changed")
        paths[name] = path
    return paths


def _profile_receipt(row: dict[str, Any], block: int, shape: str,
                     plan: dict[str, Any], build: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    stem = f"{block}-{shape}"
    c.require(row.get("block") == block and row.get("shape") == shape,
              f"{stem}: receipt position changed")
    c.require(row.get("exit_code") == 0, f"{stem}: Callgrind process failed")
    c.require(row.get("binary") == build["binaries"]["profile"],
              f"{stem}: binary identity changed")
    c.require(row.get("driver_sha256") == c.sha256(c.PACKET / "profile.py"),
              f"{stem}: profile driver changed")
    command = row.get("command")
    c.require(isinstance(command, list), f"{stem}: command is malformed")
    raw_path = str(c.PACKET / "profiles" / f"{stem}.callgrind")
    output_path = str(c.PACKET / "profiles" / f"{stem}.json")
    expected = [
        "taskset", "-c", str(plan["cpu"]), "valgrind", "--tool=callgrind",
        "--collect-atstart=no", "--toggle-collect=" + c.OWNER,
        "--zero-before=" + c.OWNER, "--dump-after=" + c.OWNER,
        "--callgrind-out-file=" + raw_path, build["binaries"]["profile"]["path"],
        "--mode", "capture", "--shape", shape, "--samples", "1", "--warmup", "0",
        "--output", output_path,
    ]
    c.require(c.normalize_command(command) == c.normalize_command(expected),
              f"{stem}: Callgrind command changed")
    paths = _artifacts(row, stem)
    return paths, {"command": command, "stem": stem}


def _raw_pair(numbered: Path, terminal: Path, parser: Any, stem: str) -> dict[str, Any]:
    number = parser.parse_raw(numbered)
    final = parser.parse_raw(terminal, allow_empty=True)
    c.require(number["header"]["events"] == ["Ir"], f"{stem}: event set changed")
    c.require(number["header"]["part"] == 1
              and number["header"]["trigger"] == "--dump-after=" + c.OWNER,
              f"{stem}: positive dump header changed")
    c.require(number["header"]["summary_ir"] > 0
              and number["header"]["totals_ir"] == number["header"]["summary_ir"],
              f"{stem}: positive Ir summary changed")
    c.require(final["header"]["part"] == 2
              and final["header"]["trigger"] == "Program termination"
              and final["header"]["summary_ir"] == 0
              and final["header"]["totals_ir"] == 0,
              f"{stem}: termination dump is not empty")
    c.require(number["header"].get("command") and final["header"].get("command"),
              f"{stem}: Callgrind command header is missing")
    unexpected = numbered.with_name(numbered.name.rsplit(".1", 1)[0] + ".2")
    c.require(not unexpected.exists(), f"{stem}: unexpected second positive dump")
    owner_ids = [fid for fid, value in number["functions"].items()
                 if value["name"] == c.OWNER]
    c.require(len(owner_ids) == 1, f"{stem}: exact wrapper owner is not unique")
    owner_id = owner_ids[0]
    incoming = parser.incoming_edges(number, owner_id)
    c.require(len(incoming) == 1 and incoming[0]["calls"] == 1
              and incoming[0]["inclusive_ir"] == number["header"]["summary_ir"],
              f"{stem}: owner incoming edge changed")
    owner_view = parser.function_view(number["functions"][owner_id], number["names"])
    c.require(owner_view["self_ir"] + owner_view["direct_children_ir"]
              == number["header"]["summary_ir"],
              f"{stem}: owner flat partition does not reconstruct summary")
    child_rows = [item for item in owner_view["direct_children"]
                  if item["callee"] == c.CHILD]
    c.require(len(child_rows) == 1, f"{stem}: capture child is missing")
    child_id = child_rows[0]["callee_id"]
    child = number["functions"].get(child_id)
    c.require(child is not None, f"{stem}: capture child function is missing")
    child_view = parser.function_view(child, number["names"])
    nested_rows = [item for item in child_view["direct_children"]
                   if item["callee"] == c.NESTED]
    c.require(len(nested_rows) == 1, f"{stem}: capture_internal child is missing")
    function_views = [parser.function_view(value, number["names"])
                      for value in number["functions"].values()]
    all_self = sum(value["self_ir"] for value in function_views)
    c.require(all_self == number["header"]["summary_ir"],
              f"{stem}: function self Ir does not reconstruct summary")
    function_views.sort(key=lambda value: (-value["self_ir"], value["name"], value["id"]))
    return {
        "file": str(numbered.relative_to(c.PACKET)), "sha256": number["sha256"],
        "bytes": number["bytes"],
        "termination": {"file": str(terminal.relative_to(c.PACKET)),
                         "sha256": final["sha256"], "bytes": final["bytes"],
                         "summary_ir": 0},
        "summary_ir": number["header"]["summary_ir"],
        "owner": {
            "id": owner_id, "name": c.OWNER, "incoming": incoming[0],
            "self_ir": owner_view["self_ir"],
            "direct_children_ir": owner_view["direct_children_ir"],
            "direct_children": owner_view["direct_children"],
            "partition_equation": "wrapper self Ir + immediate child inclusive Ir = owner inclusive Ir",
            "partition_disjoint": True, "nested_inclusive_rows_excluded": True,
        },
        "dominant_child": child_rows[0], "capture_internal": nested_rows[0],
        "ancestry_to_owner": parser.ancestry_to_owner(number, owner_id),
        "dominant_inclusive_path": parser.dominant_path(number, owner_id),
        "top_self_functions": [value for value in function_views if value["self_ir"] > 0][:20],
        "parser": number["statistics"], "all_function_self_ir": all_self,
        "validation": {
            "exact_trigger": True, "summary_nonzero": True,
            "termination_zero_ir": True, "exactly_one_positive_owner_call": True,
            "owner_inclusive_matches_summary": True,
            "self_plus_immediate_children_equals_summary": True,
            "all_function_self_ir_equals_summary": True,
            "expected_direct_child_present": True,
            "expected_nested_child_present": True,
            "compressed_names_resolved": number["statistics"]["compressed_name_declarations"] > 0,
            "relative_or_wildcard_positions_parsed": any(
                number["statistics"]["position_kinds"].get(kind, 0) > 0
                for kind in ("relative", "wildcard")),
        },
    }


def analyze() -> dict[str, Any]:
    plan = c.read_json(c.PACKET / "plan.json")
    c.require(plan.get("schema") == "litchi.performance.0807.v1", "plan schema changed")
    build = _build(plan)
    complete = c.read_json(c.PACKET / "profiles" / "complete.json")
    c.require(complete == {
        "processes": 6,
        "plan_sha256": c.sha256(c.PACKET / "plan.json"),
        "build_sha256": c.sha256(c.PACKET / "build" / "build.json"),
    }, "profile completion receipt changed")
    rows = c.read_json(c.PACKET / "profiles" / "receipts.json")
    c.require(isinstance(rows, list) and len(rows) == 6,
              "profile receipt cardinality changed")
    parser = c.load_profile_parser()
    profiles: list[dict[str, Any]] = []
    expected_jobs = [(block, shape) for block, order in enumerate(PROFILE_ORDERS)
                     for shape in order]
    for row, (block, shape) in zip(rows, expected_jobs):
        paths, receipt = _profile_receipt(row, block, shape, plan, build)
        report = c.read_json(paths[f"{block}-{shape}.json"])
        identity = c.check_report_identity(report, shape, samples=1, warmup=0,
                                           binary="profile")
        raw = _raw_pair(paths[f"{block}-{shape}.callgrind.1"],
                        paths[f"{block}-{shape}.callgrind"], parser,
                        f"{block}-{shape}")
        raw.update({
            "block": block, "shape": shape,
            "report": str(paths[f"{block}-{shape}.json"].relative_to(c.PACKET)),
            "report_sha256": c.sha256(paths[f"{block}-{shape}.json"]),
            "report_identity": {key: identity[key] for key in ("source", "output", "verification")},
            "log": str(paths[f"{block}-{shape}.log"].relative_to(c.PACKET)),
            "command": receipt["command"],
        })
        profiles.append(raw)
    c.require([(item["block"], item["shape"]) for item in profiles] == expected_jobs,
              "profile result order changed")
    return {
        "schema": "litchi-0807-callgrind-profile-analysis-v1",
        "packet": "change-0807",
        "plan": {"path": "plan.json", "sha256": c.sha256(c.PACKET / "plan.json"),
                 "schema": plan["schema"], "owner": c.OWNER, "cpu": plan["cpu"]},
        "build": {"path": "build/build.json",
                   "sha256": c.sha256(c.PACKET / "build" / "build.json"),
                   "binaries": {name: c.external_artifact(value, f"build {name}", c.cleanup_witness())
                                for name, value in build["binaries"].items()}},
        "source": c.source_identity(),
        "profiles": profiles,
        "summary": {
            "profile_count": len(profiles),
            "positive_ir_profiles": sum(item["summary_ir"] > 0 for item in profiles),
            "inclusive_ir_min": min(item["summary_ir"] for item in profiles),
            "inclusive_ir_max": max(item["summary_ir"] for item in profiles),
            "wrapper_self_ir_values": [item["owner"]["self_ir"] for item in profiles],
            "all_exact_scope_checks_pass": True,
            "all_output_fixture_checks_pass": True,
        },
        "scope": "Callgrind Ir is guest-instruction attribution for the exact capture wrapper; it is not native latency, a CPU fraction, or a production speedup.",
        "claims": [
            "Wrapper self Ir plus immediate-child inclusive Ir is a disjoint flat partition.",
            "Nested inclusive rows are retained as path diagnostics and are never added to the flat partition.",
            "Source/output/semantic identities are bound to the sealed 0780 and 0785 capture fixtures.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    c.require(args.check ^ args.write, "choose exactly one of --write or --check")
    c.write_or_check(c.PACKET / "profile-analysis.json", analyze(), args.check)
    print("0807 Callgrind analysis PASS", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
