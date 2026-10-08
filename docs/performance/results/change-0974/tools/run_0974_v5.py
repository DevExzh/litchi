#!/usr/bin/env python3
"""0974 v5 serial runner: per-thread perf mode plus raw task auditing."""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import protocol_adapter_0974_v5 as a


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--self-test", action="store_true")
    ap.add_argument("--check-plan", type=Path)
    ap.add_argument("--run", action="store_true")
    ap.add_argument("--plan", type=Path)
    ap.add_argument("--go", type=Path)
    args = ap.parse_args()
    try:
        if args.self_test:
            if args.check_plan or args.run or args.plan or args.go:
                raise a.ProtocolError("--self-test is standalone")
            print(json.dumps(a.self_test(), indent=2, sort_keys=True))
            return 0
        a.install_runner_patches()
        a.BASE_RUNNER.load_plan = a.load_plan
        a.BASE_RUNNER.p.check_go = a.check_go
        if args.check_plan:
            if args.run or args.plan or args.go:
                raise a.ProtocolError("--check-plan is standalone")
            plan = a.load_plan(args.check_plan.resolve())
            print(json.dumps({"schema": a.PLAN_SCHEMA, "status": "pass", "plan": str(args.check_plan.resolve()), "sha256": a.sha256(args.check_plan.resolve())}, indent=2, sort_keys=True))
            return 0
        if args.run:
            if args.plan is None or args.go is None:
                raise a.ProtocolError("--run requires --plan and --go")
            return a.BASE_RUNNER.run(args.plan.resolve(), args.go.resolve())
        raise a.ProtocolError("choose --self-test, --check-plan, or --run")
    except Exception as exc:
        print(f"0974 v5 runner rejected: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
