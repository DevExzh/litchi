#!/usr/bin/env python3
"""0975 offline launch-boundary check against the concrete frozen plan.

The retained 0974 v5 runner's real ``BASE_RUNNER.run_child`` is invoked after
the v5 patches are installed.  Only its ``subprocess.Popen`` call is replaced
with a sentinel, so command construction reaches the launch boundary without
starting perf, a workload, or a native child.  The expected command comes
from the independently bound argv template in the plan's record spec.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import importlib
import json
import os
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

ROOT = Path("/home/zhuhe/litchi-performance-0975")
TOOLS = ROOT / "tools"
V5_TOOLS = Path("/home/zhuhe/litchi-performance-0974/tools")
PLAN = ROOT / "plans/plan-0975-pre-capture-v3.json"
RUNNER = TOOLS / "run_0975.py"
PLAN_BINDING = TOOLS / "plan_binding_0975.py"
OLD_0974_REGRESSION = Path("/home/zhuhe/litchi-performance-0974/receipts/launch-boundary-regression-0974-v3.json")
EXPECTED = {
    "plan": "4873e502f09a3f4572ecd28d7545541e7cf7d27738c788d46fecb97293d36f97",
    "helper_seal": "9af5da6061066645d1b7d8e40b712f39113f7c7b53e16c5b3c713793f5ac5687",
    "plan_binding": "489699a6e94bb88eb6d5dd911d6959d4d1859e80590cac1cd2f3d2d92bb547e2",
    "runner": "1dd23f2d38ff9c7333f92d3ab54bd6ae68fb870c5e8d1fc89a743cdec339f125",
    "protocol_v5": "e3725cbc027d2a063f32b926e3fb4024d640aab2a2fbd7493ed5325de9ea2ac6",
}

if str(V5_TOOLS) not in sys.path:
    sys.path.insert(0, str(V5_TOOLS))
adapter = importlib.import_module("protocol_adapter_0974_v5")
runner = adapter.BASE_RUNNER


class FixtureError(RuntimeError):
    pass


class LaunchReached(Exception):
    """The sentinel reached the real native-launch call."""


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def sha256(path: str | Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: str | Path) -> dict[str, Any]:
    resolved = Path(path).resolve()
    if not resolved.is_file():
        raise FixtureError(f"missing binding path: {resolved}")
    return {"path": str(resolved), "sha256": sha256(resolved), "bytes": resolved.stat().st_size}


def load_json(path: str | Path) -> Any:
    try:
        return json.loads(Path(path).resolve().read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise FixtureError(f"cannot load JSON {path}: {exc}") from exc


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def frozen_bindings() -> dict[str, dict[str, Any]]:
    plan = load_json(PLAN)
    rows = {
        "plan": PLAN,
        "helper_seal": plan["helper_seal"]["path"],
        "plan_binding": PLAN_BINDING,
        "runner": RUNNER,
        "protocol_v5": V5_TOOLS / "protocol_adapter_0974_v5.py",
        "reducer": plan["helpers"]["reducer"]["path"],
        "generator": plan["helpers"]["plan_generator"]["path"],
        "record_spec": plan["record_spec"]["path"],
    }
    result = {name: binding(path) for name, path in rows.items()}
    for name, expected in EXPECTED.items():
        if result[name]["sha256"] != expected:
            raise FixtureError(f"frozen {name} hash drift: {result[name]['sha256']} != {expected}")
    return result


def check_plan_and_cli(plan: dict[str, Any]) -> dict[str, Any]:
    """Run the fresh plan checker and the author CLI's read-only admission."""
    sys.path.insert(0, str(TOOLS))
    checker = importlib.import_module("plan_binding_0975")
    checked = checker.check(PLAN)
    env = os.environ.copy()
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    env.update({"LC_ALL": "C", "LANG": "C", "TZ": "UTC"})
    cp = subprocess.run(
        [sys.executable, "-B", str(RUNNER), "--check-plan", str(PLAN)],
        stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, env=env, timeout=60, check=False,
    )
    if cp.returncode != 0:
        raise FixtureError(f"0975 CLI --check-plan failed: {cp.stderr[-500:]}")
    cli = json.loads(cp.stdout)
    if cli.get("status") != "pass" or cli.get("sha256") != sha256(PLAN):
        raise FixtureError(f"0975 CLI --check-plan result drift: {cli}")
    return {
        "plan_binding_checker": checked,
        "cli": {"argv": [sys.executable, "-B", str(RUNNER), "--check-plan", str(PLAN)], "returncode": cp.returncode, "stdout": cli, "stderr": cp.stderr},
    }


def expected_argv_from_spec(plan: dict[str, Any], child_dir: Path) -> dict[str, Any]:
    spec_path = Path(plan["record_spec"]["path"]).resolve()
    spec = load_json(spec_path)
    template = spec.get("argv_template")
    if spec.get("status") != "independent_final_spec" or not isinstance(template, list):
        raise FixtureError("record spec is not a bound independent argv template")
    expected = [str(token).replace("<child_dir>", str(child_dir)) for token in template]
    if expected.count("--per-thread") != 1 or expected[expected.index("--call-graph") - 2:expected.index("--call-graph")] != ["-m", "16M"]:
        raise FixtureError("independent spec does not encode the frozen per-thread/16 MiB command")
    if any(token in expected for token in spec["record_mode"]["forbidden_tokens"]):
        raise FixtureError("independent argv template contains a forbidden token")
    # Preserve the plan's custody spelling in argv.  Resolving the fixture
    # path would collapse a historical symlink and create a false mismatch
    # against the independent command template.
    binary = str(plan["binary"]["path"])
    fixture = str(plan["fixture"]["path"])
    if binary not in expected or fixture not in expected:
        raise FixtureError("independent argv template is not bound to concrete plan inputs")
    return {"binding": binding(spec_path), "expected_argv": expected, "canonical_argv": [token.replace(str(child_dir), "<child_dir>") for token in expected]}


def alias_snapshot() -> dict[str, Any]:
    return {
        "raw_module_alias": runner.raw_stack is adapter.BASE_RAW_STACK,
        "scan_alias": runner.raw_stack.scan_dump is adapter.BASE_RAW_STACK.scan_dump,
        "child_builder_alias": runner.child_record_argv is adapter.native_record_argv,
        "scan_repr": repr(runner.raw_stack.scan_dump),
        "child_repr": repr(runner.child_record_argv),
    }


def run(output: Path) -> dict[str, Any]:
    before = frozen_bindings()
    plan = load_json(PLAN)
    cli = check_plan_and_cli(plan)
    old = binding(OLD_0974_REGRESSION)
    original_write = runner.p.write_json_exclusive
    original_scan = runner.raw_stack.scan_dump
    original_child = runner.child_record_argv
    original_popen = runner.subprocess.Popen
    before_install = alias_snapshot()
    calls: list[dict[str, Any]] = []
    boundary_error: dict[str, Any] | None = None
    first_install: dict[str, Any] | None = None
    second_install: dict[str, Any] | None = None
    clear_report: dict[str, Any] | None = None
    try:
        adapter.install_runner_patches()
        first_install = alias_snapshot()
        with tempfile.TemporaryDirectory(prefix="litchi-0975-launch-boundary-") as temp:
            capture_root = Path(temp) / "capture"
            capture_root.mkdir()
            child_dir = capture_root / "child-01"
            independent = expected_argv_from_spec(plan, child_dir)

            def sentinel(argv: Any, **kwargs: Any) -> Any:
                calls.append({
                    "argv": list(argv),
                    "kwargs": {
                        "start_new_session": kwargs.get("start_new_session"),
                        "preexec_fn": repr(kwargs.get("preexec_fn")),
                        "stdin": repr(kwargs.get("stdin")),
                        "stdout_name": getattr(kwargs.get("stdout"), "name", None),
                        "stderr_name": getattr(kwargs.get("stderr"), "name", None),
                        "env_subset": {key: kwargs.get("env", {}).get(key) for key in ("LC_ALL", "LANG", "TZ", "LD_PRELOAD")},
                    },
                })
                raise LaunchReached("Popen sentinel reached real launch boundary")

            runner.subprocess.Popen = sentinel
            try:
                runner.run_child(plan, 1, capture_root)
            except LaunchReached as exc:
                boundary_error = {"type": type(exc).__name__, "message": str(exc)}
            except Exception as exc:
                boundary_error = {"type": type(exc).__name__, "message": str(exc)}
            if not calls:
                independent = {**independent, "actual_child_dir": str(child_dir)}

        adapter.install_runner_patches()
        second_install = alias_snapshot()
        clear_before = alias_snapshot()
        adapter.clear_runner_patches()
        clear_after = alias_snapshot()
        clear_report = {
            "before": clear_before,
            "after": clear_after,
            "runner_child_restored": runner.child_record_argv is adapter.ORIGINAL_CHILD_RECORD_ARGV,
            "runner_scan_restored": runner.raw_stack.scan_dump is adapter.ORIGINAL_SCAN_DUMP,
            "runner_write_restored": runner.p.write_json_exclusive is adapter.ORIGINAL_WRITE_JSON_EXCLUSIVE,
            "patch_state_cleared": not hasattr(runner, adapter._PATCH_STATE_ATTR),
        }
    finally:
        runner.subprocess.Popen = original_popen
        runner.p.write_json_exclusive = original_write
        runner.raw_stack.scan_dump = original_scan
        runner.child_record_argv = original_child

    after = frozen_bindings()
    boundary_ok = bool(calls and boundary_error and boundary_error["type"] == "LaunchReached")
    argv_ok = boundary_ok and calls[0]["argv"] == independent["expected_argv"]
    aliases_ok = bool(first_install and first_install["raw_module_alias"] and first_install["scan_alias"])
    idempotent = bool(second_install and first_install and second_install["scan_repr"] == first_install["scan_repr"] and second_install["child_repr"] == first_install["child_repr"])
    clear_ok = bool(clear_report and all(clear_report[key] for key in ("runner_child_restored", "runner_scan_restored", "runner_write_restored", "patch_state_cleared")))
    result = {
        "schema": "litchi.performance.0975.launch_boundary_regression.v1",
        "status": "pass" if boundary_ok and argv_ok and aliases_ok and idempotent and clear_ok and before == after else "fail",
        "created_utc": now(),
        "offline_only": True,
        "workload_started": False,
        "build_started": False,
        "native_capture_started": False,
        "plan": before["plan"],
        "helper_seal": before["helper_seal"],
        "old_0974_regression": old,
        "plan_and_cli": cli,
        "independent_argv": independent,
        "boundary": {"popen_calls": len(calls), "launch_reached": boundary_ok, "argv_match": argv_ok, "exception": boundary_error, "sentinel_call": calls[0] if calls else None},
        "patch_composition": {"before_install": before_install, "after_first_install": first_install, "after_second_install": second_install, "raw_aliases_ok": aliases_ok, "install_idempotent": idempotent, "clear": clear_report, "clear_ok": clear_ok},
        "bindings_before": before,
        "bindings_after": after,
        "bindings_unchanged": before == after,
        "scope": "actual retained BASE_RUNNER.run_child with only Popen intercepted; no run_child replacement, native process, workload, build, or capture",
    }
    write_json(output.resolve(), result)
    return result


def main() -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = run(args.output)
    except Exception as exc:
        print(f"0975 launch-boundary regression rejected: {exc}", file=sys.stderr)
        return 2
    print(json.dumps({"status": result["status"], "output": str(args.output.resolve()), "sha256": sha256(args.output.resolve())}, indent=2, sort_keys=True))
    return 0 if result["status"] == "pass" else 2


if __name__ == "__main__":
    raise SystemExit(main())
