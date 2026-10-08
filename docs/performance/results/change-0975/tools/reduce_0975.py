#!/usr/bin/env python3
"""0975 reducer wrapper around the repaired 0974 v4 reducer."""
from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

ROOT = Path("/home/zhuhe/litchi-performance-0975")
V4_TOOLS = Path("/home/zhuhe/litchi-performance-0974/tools")
sys.path.insert(0, str(ROOT / "tools"))
import plan_binding_0975 as custody  # noqa: E402


def load_reducer():
    sys.path.insert(0, str(V4_TOOLS))
    spec = importlib.util.spec_from_file_location("litchi_reduce_0974_v4_reused_for_0975", V4_TOOLS / "reduce_0974_v4.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load repaired 0974 v4 reducer")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    adapter = sys.modules["protocol_adapter_0974_v3"]
    adapter.ROOT = ROOT
    adapter.DESIGN_PATH = ROOT / "design-0975-v1.md"
    return module


def plan_arg() -> Path | None:
    values = sys.argv[1:]
    if "--plan" not in values:
        return None
    index = values.index("--plan")
    if index + 1 >= len(values):
        raise RuntimeError("--plan requires a path")
    return Path(values[index + 1]).resolve()


def main() -> int:
    reducer = load_reducer()
    if "--self-test" not in sys.argv[1:] and "--admission-self-test" not in sys.argv[1:]:
        plan = plan_arg()
        if plan is None:
            raise RuntimeError("0975 reducer requires --plan for admission")
        custody.check(plan)
    return reducer.main()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"0975 reducer rejected before delegation: {exc}", file=sys.stderr)
        raise SystemExit(2)
