#!/usr/bin/env python3
"""0975 runner wrapper around the repaired 0974 v5 runner."""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

ROOT = Path("/home/zhuhe/litchi-performance-0975")
V5_TOOLS = Path("/home/zhuhe/litchi-performance-0974/tools")
sys.path.insert(0, str(ROOT / "tools"))
import plan_binding_0975 as custody  # noqa: E402


def load_runner():
    sys.path.insert(0, str(V5_TOOLS))
    spec = importlib.util.spec_from_file_location("litchi_run_0974_v5_reused_for_0975", V5_TOOLS / "run_0974_v5.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load repaired 0974 v5 runner")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    adapter = sys.modules["protocol_adapter_0974_v5"]
    adapter.ROOT = ROOT
    adapter.DESIGN_PATH = ROOT / "design-0975-v1.md"
    return module


def plan_arg() -> Path | None:
    values = sys.argv[1:]
    option = "--plan" if "--plan" in values else "--check-plan" if "--check-plan" in values else None
    if option is None:
        return None
    index = values.index(option)
    if index + 1 >= len(values):
        raise RuntimeError("--plan requires a path")
    return Path(values[index + 1]).resolve()


def main() -> int:
    runner = load_runner()
    if "--self-test" not in sys.argv[1:]:
        plan = plan_arg()
        if plan is None:
            raise RuntimeError("0975 runner requires --plan for admission")
        custody.check(plan)
    return runner.main()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"0975 runner rejected before delegation: {exc}", file=sys.stderr)
        raise SystemExit(2)
