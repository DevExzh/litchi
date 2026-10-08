#!/usr/bin/env python3
"""Offline launch-boundary regression for the 0974 runner patch composition.

This invokes the retained ``BASE_RUNNER.run_child`` after installing the
0974 adapter patches.  Only ``subprocess.Popen`` is replaced, by a sentinel
that raises as soon as the real launch boundary is reached.  No process,
perf capture, workload, or decoder can start.  The test also records the
raw-stack module alias and whether a second patch installation reuses the
first wrappers.

The first v3 run used a fixture ``run_child`` replacement and therefore did
not exercise command construction.  This regression keeps that receipt
immutable and makes the missed boundary explicit.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import importlib
import json
import sys
import tempfile
from pathlib import Path
from typing import Any

sys.dont_write_bytecode = True

ROOT = Path("/home/zhuhe/litchi-performance-0974")
TOOLS = ROOT / "tools"
ADAPTER_MODULE = "protocol_adapter_0974_v5"
OLD_PASS_RECEIPT = ROOT / "receipts/synthetic-continuation-reducer-author-v4-v1.json"
EXPECTED_SPEC_SCHEMA = "litchi.performance.0974.native_record_spec.v1"
FROZEN = {
    "protocol_v5": (TOOLS / "protocol_adapter_0974_v5.py", "e3725cbc027d2a063f32b926e3fb4024d640aab2a2fbd7493ed5325de9ea2ac6"),
    "runner_v5": (TOOLS / "run_0974_v5.py", "519a7b88cf6df7688d1b4555fbd3b2d682c2459bc6c070817308ba38031d4fe1"),
    "reducer_v4": (TOOLS / "reduce_0974_v4.py", "7c6e7162e406aecf13868b264da29f2d5b6e07f0212bed1099298a464bb9ade5"),
    "plan_generator_v5": (TOOLS / "prepare_plan_0974_v5.py", "eb5d0f42055b8917078602d1329b1aafdec3eef453370816aec102a9e0ef486e"),
    "plan_v5": (ROOT / "plans/plan-0974-pre-capture-v5.json", "dddd126895d04cd1738719d99aa00ff1b68dcd551233a5c2391f37f7c30e71ee"),
    "helper_seal_v5": (ROOT / "receipts/helper-seal-0974-v5.json", "4cc5ef36a28c6b254a580abd01b21792daef317b383ac5dc718ea544bb3668a3"),
}

if str(TOOLS) not in sys.path:
    sys.path.insert(0, str(TOOLS))
adapter = importlib.import_module(ADAPTER_MODULE)
runner = adapter.BASE_RUNNER


class FixtureError(RuntimeError):
    pass


class LaunchReached(Exception):
    """Raised by the Popen sentinel at the real launch boundary."""


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: str | Path) -> dict[str, Any]:
    resolved = Path(path).resolve()
    if not resolved.is_file():
        raise FixtureError(f"missing binding path: {resolved}")
    return {"path": str(resolved), "sha256": sha256(resolved), "bytes": resolved.stat().st_size}


def load_json(path: str | Path) -> Any:
    resolved = Path(path).resolve()
    try:
        return json.loads(resolved.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise FixtureError(f"cannot load JSON {resolved}: {exc}") from exc


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def production_bindings() -> dict[str, dict[str, Any]]:
    return {name: binding(path) for name, (path, _expected) in FROZEN.items()}


def verify_frozen(rows: dict[str, dict[str, Any]]) -> None:
    for name, (path, expected) in FROZEN.items():
        row = rows.get(name) or {}
        if row.get("path") != str(path.resolve()) or row.get("sha256") != expected:
            raise FixtureError(f"frozen {name} binding drift: {row}")


def independent_expected_argv(spec_path: Path, plan: dict[str, Any], child_dir: Path) -> dict[str, Any]:
    """Derive expected argv before patch installation, then bind the spec."""
    spec = load_json(spec_path)
    if spec.get("schema") != EXPECTED_SPEC_SCHEMA:
        raise FixtureError("native argv spec schema mismatch")
    baseline = list(adapter.V1.native_record_argv(plan, child_dir))
    record = baseline.index("record")
    expected = baseline[:record + 1] + ["--per-thread"] + baseline[record + 1:]
    mode = spec.get("record_mode") or {}
    if mode.get("required_tokens") != ["--per-thread"] or mode.get("inheritance") != "disabled_by_per_thread":
        raise FixtureError("independent spec does not bind per-thread mode")
    if any(token in expected for token in mode.get("forbidden_tokens", [])):
        raise FixtureError("independent spec forbids a token present in expected argv")
    memory = spec.get("record_memory") or {}
    if memory.get("tokens") != ["-m", "16M"] or memory.get("bytes") != 16 * 1024 * 1024:
        raise FixtureError("independent spec does not bind 16 MiB memory selector")
    if expected[expected.index("--call-graph") - 2:expected.index("--call-graph")] != ["-m", "16M"]:
        raise FixtureError("expected argv memory selector is misplaced")
    canonical = [token.replace(str(child_dir), "<child_dir>") for token in expected]
    template = spec.get("argv_template")
    if template is not None and canonical != [str(token) for token in template]:
        raise FixtureError("independent argv template mismatch")
    return {
        "spec": binding(spec_path),
        "baseline_v1_argv": baseline,
        "expected_v3_argv": expected,
        "expected_canonical_argv": canonical,
        "template_checked": template is not None,
    }


def alias_snapshot() -> dict[str, Any]:
    return {
        "runner_raw_stack_is_adapter_base_raw_stack": runner.raw_stack is adapter.BASE_RAW_STACK,
        "runner_scan_is_adapter_scan": runner.raw_stack.scan_dump is adapter.BASE_RAW_STACK.scan_dump,
        "runner_child_record_is_adapter_native": runner.child_record_argv is adapter.native_record_argv,
        "runner_scan_repr": repr(runner.raw_stack.scan_dump),
        "runner_child_record_repr": repr(runner.child_record_argv),
    }


def run_regression(spec_path: Path, output: Path) -> dict[str, Any]:
    before = production_bindings()
    verify_frozen(before)
    old_pass = binding(OLD_PASS_RECEIPT)
    with tempfile.TemporaryDirectory(prefix="litchi-0974-launch-boundary-") as temp:
        root = Path(temp)
        capture_root = root / "capture"
        capture_root.mkdir()
        child_dir = capture_root / "child-01"
        plan = {"binary": {"path": adapter.BASE.BINARY_PATH}, "fixture": {"path": adapter.BASE.FIXTURE_PATH}, "execution": {"timeout_seconds": 900}}
        expected = independent_expected_argv(spec_path.resolve(), plan, child_dir)

        original_write = runner.p.write_json_exclusive
        original_scan = runner.raw_stack.scan_dump
        original_child = runner.child_record_argv
        original_popen = runner.subprocess.Popen
        before_install = alias_snapshot()
        calls: list[dict[str, Any]] = []
        first_install: dict[str, Any] | None = None
        second_install: dict[str, Any] | None = None
        clear_report: dict[str, Any] | None = None
        boundary_error: dict[str, Any] | None = None
        try:
            adapter.install_runner_patches()
            first_install = alias_snapshot()

            def sentinel(argv: Any, **kwargs: Any) -> Any:
                calls.append({"argv": list(argv), "kwargs": {key: repr(value) for key, value in kwargs.items()}})
                raise LaunchReached("Popen sentinel reached real launch boundary")

            runner.subprocess.Popen = sentinel
            try:
                runner.run_child(plan, 1, capture_root)
            except LaunchReached as exc:
                boundary_error = {"type": type(exc).__name__, "message": str(exc)}
            except Exception as exc:
                boundary_error = {"type": type(exc).__name__, "message": str(exc)}

            # A second installation is diagnostic only.  Record whether it
            # reuses aliases instead of silently stacking wrappers.
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

        after = production_bindings()
        verify_frozen(after)
        launch_reached = bool(calls) and boundary_error is not None and boundary_error["type"] == "LaunchReached"
        argv_match = launch_reached and calls[0]["argv"] == expected["expected_v3_argv"]
        raw_alias_ok = bool(first_install and first_install["runner_raw_stack_is_adapter_base_raw_stack"] and first_install["runner_scan_is_adapter_scan"])
        patch_idempotent = bool(
            first_install and second_install and
            second_install["runner_raw_stack_is_adapter_base_raw_stack"] and
            second_install["runner_scan_is_adapter_scan"] and
            second_install["runner_child_record_is_adapter_native"]
        )
        clear_ok = bool(clear_report and all(clear_report[key] for key in ("runner_child_restored", "runner_scan_restored", "runner_write_restored", "patch_state_cleared")))
        result = {
            "schema": "litchi.performance.0974.launch_boundary_regression.v3",
            "status": "pass" if launch_reached and argv_match and raw_alias_ok and clear_ok else "fail",
            "created_utc": now(),
            "offline_only": True,
            "workload_started": False,
            "build_started": False,
            "native_capture_started": False,
            "old_pass_receipt": old_pass,
            "scope": "actual BASE_RUNNER.run_child after adapter.install_runner_patches; only subprocess.Popen replaced by sentinel",
            "independent_argv": expected,
            "boundary": {
                "popen_calls": len(calls),
                "launch_reached": launch_reached,
                "argv_match": argv_match,
                "sentinel_call": calls[0] if calls else None,
                "exception": boundary_error,
                "recursion_gap": boundary_error is not None and boundary_error["type"] == "RecursionError",
            },
            "patch_composition": {
                "before_install": before_install,
                "after_first_install": first_install,
                "after_second_install": second_install,
                "raw_scanner_aliases_valid": raw_alias_ok,
                "second_install_aliases_valid": patch_idempotent,
                "second_install_scan_reused": bool(first_install and second_install and first_install["runner_scan_repr"] == second_install["runner_scan_repr"]),
                "second_install_child_reused": bool(first_install and second_install and first_install["runner_child_record_repr"] == second_install["runner_child_record_repr"]),
                "clear": clear_report,
                "clear_ok": clear_ok,
            },
            "production_helpers_before": before,
            "production_helpers_after": after,
            "production_helpers_unchanged": before == after,
            "interpretation": "The previous PASS receipt used a run_child fixture replacement. This result exercises the real command-construction path and reports whether it reaches Popen without starting a process.",
        }
    write_json(output.resolve(), result)
    return result


def main() -> int:
    import argparse

    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--argv-spec", type=Path, required=True)
    ap.add_argument("--output", type=Path, required=True)
    args = ap.parse_args()
    try:
        result = run_regression(args.argv_spec, args.output)
    except Exception as exc:
        print(f"0974 launch-boundary regression rejected: {exc}", file=sys.stderr)
        return 2
    print(json.dumps({"status": result["status"], "output": str(args.output.resolve()), "sha256": sha256(args.output.resolve())}, indent=2, sort_keys=True))
    return 0 if result["status"] == "pass" else 2


if __name__ == "__main__":
    raise SystemExit(main())
