#!/usr/bin/env python3
"""Create the fresh 0975 plan while reusing the repaired 0974 v5 stack."""
from __future__ import annotations

import argparse
import copy
import importlib.util
import json
import sys
from pathlib import Path

ROOT = Path("/home/zhuhe/litchi-performance-0975")
V5_TOOLS = Path("/home/zhuhe/litchi-performance-0974/tools")
V5_GENERATOR = V5_TOOLS / "prepare_plan_0974_v5.py"
sys.path.insert(0, str(Path(__file__).resolve().parent))
import plan_binding_0975 as custody  # noqa: E402


def load_v5_generator():
    sys.path.insert(0, str(V5_TOOLS))
    spec = importlib.util.spec_from_file_location("litchi_prepare_plan_0974_v5_reused", V5_GENERATOR)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load retained 0974 v5 plan generator")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    module.ROOT = ROOT
    module.a.ROOT = ROOT
    module.a.DESIGN_PATH = ROOT / "design-0975-v1.md"
    return module


GEN = load_v5_generator()


def helper_rows() -> dict[str, dict]:
    return {
        "protocol": GEN.a.file_binding(V5_TOOLS / "protocol_adapter_0974_v5.py"),
        "runner": GEN.a.file_binding(ROOT / "tools/run_0975.py"),
        "reducer": GEN.a.file_binding(ROOT / "tools/reduce_0975.py"),
        "plan_generator": GEN.a.file_binding(ROOT / "tools/prepare_plan_0975.py"),
        "raw_stack": GEN.a.file_binding(V5_TOOLS / "protocol_adapter_0974_v5.py"),
        "shim_source": GEN.a.file_binding(GEN.a.V1.RETAINED_ROOT / "tools-v2/dwfl_dynamic_shim_v2.c"),
        "shim_map": GEN.a.file_binding(GEN.a.V1.RETAINED_ROOT / "tools-v2/dwfl_dynamic_shim_v2.map"),
        "shim_so": GEN.a.file_binding(GEN.a.V1.RETAINED_ROOT / "replay-dwfl-v2/dwfl_dynamic_shim_v2.so"),
    }


GEN.helper_rows = helper_rows


def batch_contract() -> dict:
    return {
        "fresh_cohort": True,
        "not_a_retry_of": "/home/zhuhe/litchi-performance-0974",
        "children": 12,
        "serial": True,
        "cpu": 31,
        "event": "instructions:u",
        "period": 1000000,
        "callgraph": "dwarf,32768",
        "record_mode": ["--per-thread"],
        "record_memory": ["-m", "16M"],
        "record_memory_bytes": 16777216,
        "warmups": 5,
        "samples": 80,
        "minimum_marker_samples_each_child": 100,
        "marker_period_share_positive_each_child": True,
        "required_raw_lost_sum": 0,
        "required_terminal_lost_samples": 0,
        "no_retry": True,
        "stop_on_first_failure": True,
    }


def make_seal() -> dict:
    seal = GEN.make_helper_seal()
    seal.update({
        "version": "0975-reuse-0974-v5",
        "scope": "fresh 0975 pre-capture wrapper; repaired 0974 v5 native protocol and v4 reducer are reused without source/build changes",
        "helpers": helper_rows(),
        "dependency_bindings": custody.expected_dependencies(),
        "fresh_output_root": str((ROOT / "capture/instruction-profile-v1").resolve()),
        "no_build_or_workload": True,
    })
    return seal


def make_plan(helper_seal_path: Path, output: Path) -> dict:
    seal = GEN.a.load_json(helper_seal_path)
    if seal.get("helpers") != helper_rows() or seal.get("dependency_bindings") != custody.expected_dependencies():
        raise GEN.a.ProtocolError("0975 helper seal changed before plan generation")
    plan = GEN.make_plan(helper_seal_path, output)
    plan["batch_0975"] = batch_contract()
    plan["dependency_bindings"] = custody.expected_dependencies()
    plan["reused_0974_v5"] = {
        "plan": custody.binding(custody.V5_PLAN),
        "helper_seal": custody.binding(custody.V5_HELPER_SEAL),
        "scope": "source/build/protocol lineage only; 0975 output paths are fresh",
    }
    plan["interpretation"] = {
        **plan.get("interpretation", {}),
        "fresh_cohort": True,
        "not_a_retry_of": "/home/zhuhe/litchi-performance-0974",
        "diagnostic_only": True,
        "qualification_claim": False,
        "speed_claim": False,
        "exact_timed_region_claim": False,
    }
    plan["no_build_or_workload"] = True
    GEN.a.check_plan_bindings(plan, None)
    return plan


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--seal-output", type=Path)
    parser.add_argument("--helper-seal", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        if args.seal_output:
            if args.helper_seal or args.output:
                raise GEN.a.ProtocolError("--seal-output is standalone")
            output = args.seal_output.resolve()
            GEN.a.BASE.write_json_exclusive(output, make_seal())
            print(json.dumps({"status": "pass", "schema": GEN.a.HELPER_SCHEMA, "path": str(output), "sha256": GEN.a.sha256(output)}, indent=2, sort_keys=True))
            return 0
        if args.helper_seal is None or args.output is None:
            raise GEN.a.ProtocolError("--helper-seal and --output are required")
        output = args.output.resolve()
        plan = make_plan(args.helper_seal.resolve(), output)
        GEN.a.check_plan_bindings(plan, None)
        GEN.a.BASE.write_json_exclusive(output, plan)
        GEN.a.check_plan_bindings(GEN.a.load_json(output), output)
        checked = custody.check(output)
        print(json.dumps({"status": "pass", "schema": GEN.a.PLAN_SCHEMA, "path": str(output), "sha256": GEN.a.sha256(output), "custody": checked}, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        print(f"0975 plan generator rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
