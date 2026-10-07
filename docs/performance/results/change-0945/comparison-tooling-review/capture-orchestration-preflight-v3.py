#!/usr/bin/env python3
"""Exercise the 0945 capture sealing path without a benchmark workload.

The frozen runner's real binaries and source checks are replaced only inside
this external fixture by a tiny executable that writes a synthetic report.
The old runner is retained to prove the absolute-plan-descriptor failure; a
copy with the two internal plan descriptor call sites made relative is then
run through one complete qualification stage.  Nothing in the repository or
the frozen packet is modified.
"""
from __future__ import annotations

import contextlib
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import shutil
import stat
import sys
from typing import Any


REVIEW = Path(__file__).resolve().parent
SOURCE_RUNNER = Path("/home/zhuhe/litchi-performance-0945/tools/capture_0945.py")
CORRECTED_RUNNER = Path("/home/zhuhe/litchi-performance-0945/tools/capture_0945_v2.py")
FROZEN_PLAN = Path("/home/zhuhe/litchi-performance-0945/plan.json")
CORRECTED_PLAN = Path("/home/zhuhe/litchi-performance-0945/plan-v2.json")
REDUCER_V2 = Path("/home/zhuhe/litchi-performance-0945/tools/reduce_0945_v2.py")
ROOT = REVIEW / "orchestration-preflight-v3"
RESULT = REVIEW / "capture-orchestration-preflight-v3.json"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def import_runner(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def make_fixture(name: str, *, corrected: bool) -> tuple[Any, Path, dict[str, Any]]:
    root = ROOT / name
    (root / "tools").mkdir(parents=True, exist_ok=False)
    runner = root / "tools" / "capture_0945.py"
    source_path = CORRECTED_RUNNER if corrected else SOURCE_RUNNER
    source = source_path.read_text(encoding="utf-8")
    runner.write_text(source, encoding="utf-8")
    runner.chmod(0o755)

    plan_path = root / ("plan-v2.json" if corrected else "plan.json")
    shutil.copy2(CORRECTED_PLAN if corrected else FROZEN_PLAN, plan_path)
    fake = root / "fake-binary"
    fake.write_text(
        "#!/usr/bin/python3\n"
        "import json, sys\n"
        "out = sys.argv[sys.argv.index('--json') + 1]\n"
        "with open(out, 'x', encoding='utf-8') as stream:\n"
        "    json.dump({'fixture': True}, stream)\n",
        encoding="utf-8",
    )
    fake.chmod(fake.stat().st_mode | stat.S_IXUSR)

    module = import_runner(runner, f"capture_0945_fixture_{name}")
    cpu = min(os.sched_getaffinity(0))
    row = {
        "id": "fixture-case",
        "case": "xlsx_fixture_case",
        "xlsx_cell_crud_shape": "medium",
    }
    job = {
        "stage": "qualification", "index": 0, "role": "baseline",
        "kind": "native", "case": "fixture-case", "row": row,
    }
    plan = {
        "cpu": cpu,
        "resource": {
            "enabled": True,
            "deadline_seconds": 30,
            "time_path": "/usr/bin/time",
            "format": '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
        },
    }
    binary = module.file_descriptor(fake)
    source_binding = {"fixture": True, "source_revision": "fixture"}
    env = {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "TZ": "UTC",
           "RUST_BACKTRACE": "0"}

    def role_descriptors(_plan: dict[str, Any], _role: str) -> dict[str, Any]:
        return {"native": {"path": "fake-binary"}, "allocator": {"path": "fake-binary"}}

    def jobs(_plan: dict[str, Any], _stage: str) -> list[dict[str, Any]]:
        return [job]

    module.role_descriptors = role_descriptors
    module.expected_jobs = jobs
    module.source_check = lambda _plan, full=True: {**source_binding, "full": full}
    module.check_binaries = lambda _plan: {
        "baseline": {"native": {"binary": binary}, "allocator": {"binary": binary}},
        "candidate": {"native": {"binary": binary}, "allocator": {"binary": binary}},
    }
    module.environment = lambda _plan: dict(env)
    module.workload_cwd = lambda _plan: root
    return module, root, {"plan": plan, "job": job, "binary": binary}


def run_reducer_custody() -> dict[str, Any]:
    """Exercise the reducer's remaining absolute RUNNER call sites directly."""
    old_root = ROOT / "reducer-old"
    fixed_root = ROOT / "reducer-corrected"
    for root, corrected in ((old_root, False), (fixed_root, True)):
        (root / "tools").mkdir(parents=True, exist_ok=False)
        source = REDUCER_V2.read_text(encoding="utf-8")
        if not corrected:
            old_a = 'record.get("helper") == runner_descriptor()'
            old_b = 'capture.get("helper") == runner_descriptor()'
            if source.count(old_a) != 1 or source.count(old_b) != 1:
                raise AssertionError("current reducer no longer exposes both corrected helper call sites")
            source = source.replace(old_a, 'record.get("helper") == descriptor(RUNNER, "capture helper")')
            source = source.replace(old_b, 'capture.get("helper") == descriptor(RUNNER, "capture helper")')
        (root / "tools" / "reduce_0945_v2.py").write_text(source, encoding="utf-8")
        (root / "tools" / "capture_0945_v2.py").write_text(
            CORRECTED_RUNNER.read_text(encoding="utf-8"), encoding="utf-8")
        shutil.copy2(CORRECTED_PLAN, root / "plan-v2.json")

    old_module = import_runner(old_root / "tools" / "reduce_0945_v2.py", "reduce_0945_v2_fixture_old")
    old_error = None
    try:
        old_module.descriptor(old_module.RUNNER, "capture helper")
    except Exception as exc:
        old_error = repr(exc)
    if not old_error or "path must be packet-relative" not in old_error:
        raise AssertionError(f"old reducer did not retain the expected failure: {old_error}")

    fixed_module = import_runner(fixed_root / "tools" / "reduce_0945_v2.py", "reduce_0945_v2_fixture_fixed")
    descriptor = fixed_module.runner_descriptor()
    if descriptor.get("path") != "tools/capture_0945_v2.py":
        raise AssertionError(f"corrected reducer helper descriptor escaped packet: {descriptor}")
    return {
        "old_runner_sha256": sha256(old_root / "tools" / "reduce_0945_v2.py"),
        "corrected_runner_sha256": sha256(fixed_root / "tools" / "reduce_0945_v2.py"),
        "status": "passed_reducer_relative_descriptor_fixture",
        "old_source": "versioned emulation of the two pre-fix absolute RUNNER call sites",
        "old_failure": old_error,
        "corrected_descriptor": descriptor,
    }


def json_files(root: Path) -> list[Path]:
    return sorted(path for path in root.rglob("*.json") if path.is_file())


def descriptor_paths(root: Path) -> list[dict[str, str]]:
    found: list[dict[str, str]] = []

    def walk(value: Any, source: Path, key: str = "") -> None:
        if isinstance(value, dict):
            for child_key, child in value.items():
                if child_key == "path":
                    if not isinstance(child, str):
                        raise AssertionError(f"non-string descriptor path in {source}")
                    raw = Path(child)
                    if raw.is_absolute() or ".." in raw.parts:
                        raise AssertionError(f"escaped descriptor path {child!r} in {source}")
                    if not (root / raw).resolve().is_relative_to(root.resolve()):
                        raise AssertionError(f"descriptor escapes fixture {child!r}")
                    found.append({"file": str(source.relative_to(root)), "path": child})
                walk(child, source, child_key)
        elif isinstance(value, list):
            for child in value:
                walk(child, source, key)

    for path in json_files(root):
        walk(json.loads(path.read_text(encoding="utf-8")), path)
    return found


def run_old_failure() -> dict[str, Any]:
    module, root, fixture = make_fixture("old-frozen", corrected=False)
    error = None
    captured = io.StringIO()
    try:
        with contextlib.redirect_stdout(captured), contextlib.redirect_stderr(captured):
            module.stage_capture(fixture["plan"], "qualification")
    except Exception as exc:  # CaptureError is intentionally recorded verbatim.
        error = repr(exc)
    child = root / "qualification" / "q00-baseline-native-fixture-case"
    files = sorted(path.name for path in child.iterdir()) if child.is_dir() else []
    if not error or "descriptor path must be packet-relative" not in error:
        raise AssertionError(f"old runner did not retain the expected failure: {error}")
    if files:
        raise AssertionError(f"old runner wrote files before plan binding failed: {files}")
    return {
        "runner_sha256": sha256(root / "tools" / "capture_0945.py"),
        "status": "failed_before_child_start_expected",
        "error": error,
        "child_directory": str(child.relative_to(ROOT)),
        "child_files": files,
        "captured_output": captured.getvalue(),
    }


def run_corrected_success() -> dict[str, Any]:
    module, root, fixture = make_fixture("corrected-relative", corrected=True)
    captured = io.StringIO()
    with contextlib.redirect_stdout(captured), contextlib.redirect_stderr(captured):
        module.stage_capture(fixture["plan"], "qualification")
    child = root / "qualification" / "q00-baseline-native-fixture-case"
    terminal = json.loads((child / "terminal.json").read_text(encoding="utf-8"))
    command = json.loads((child / "command.json").read_text(encoding="utf-8"))
    receipt = json.loads((child / "receipt.json").read_text(encoding="utf-8"))
    capture = json.loads((root / "qualification" / "capture.json").read_text(encoding="utf-8"))
    if terminal.get("terminal_before_parse") is not True or terminal.get("returncode") != 0:
        raise AssertionError("corrected terminal did not seal a successful child")
    if command.get("terminal_receipt", {}).get("path") != "qualification/q00-baseline-native-fixture-case/terminal.json":
        raise AssertionError("command terminal descriptor is not packet-relative")
    if receipt.get("terminal", {}).get("path") != command["terminal_receipt"]["path"]:
        raise AssertionError("receipt terminal descriptor drifted")
    if capture.get("status") != "pass" or capture.get("children") != 1:
        raise AssertionError("capture manifest did not seal the successful child")
    if capture.get("receipts") != [receipt]:
        raise AssertionError("capture manifest receipt differs from retained receipt")
    descriptors = descriptor_paths(root)

    outside = root.parent / "outside-artifact.txt"
    outside.write_text("must remain outside", encoding="utf-8")
    escape_errors: list[str] = []
    for attempt in (
        lambda: module.file_descriptor(outside),
        lambda: module.verify_descriptor({"path": "../outside-artifact.txt"}, "escape"),
        lambda: module.verify_descriptor({"path": str(outside)}, "absolute"),
    ):
        try:
            attempt()
        except Exception as exc:
            escape_errors.append(type(exc).__name__ + ": " + str(exc))
    if len(escape_errors) != 3:
        raise AssertionError(f"path escape controls did not all reject: {escape_errors}")

    return {
        "runner_sha256": sha256(root / "tools" / "capture_0945.py"),
        "status": "passed_fake_child_orchestration",
        "capture": {
            "path": str((root / "qualification" / "capture.json").relative_to(ROOT)),
            "status": capture["status"],
            "children": capture["children"],
            "receipt_count": len(capture["receipts"]),
        },
        "terminal": {
            "returncode": terminal["returncode"],
            "terminal_before_parse": terminal["terminal_before_parse"],
            "process_group_clear": terminal["process_group_clear"],
        },
        "relative_descriptor_count": len(descriptors),
        "relative_descriptors": descriptors,
        "escape_controls_rejected": len(escape_errors),
        "captured_output": captured.getvalue(),
    }


def compare_plan_invariants() -> dict[str, Any]:
    old = json.loads(FROZEN_PLAN.read_text(encoding="utf-8"))
    new = json.loads(CORRECTED_PLAN.read_text(encoding="utf-8"))
    keys = ("schema", "protocol_revision", "seed", "cases",
            "qualification", "native", "allocator", "gate")
    equal = {key: old.get(key) == new.get(key) for key in keys}
    if not all(equal.values()):
        raise AssertionError(f"plan invariants changed: {[key for key, same in equal.items() if not same]}")
    return {"status": "unchanged", "fields": equal,
            "old_plan_sha256": sha256(FROZEN_PLAN),
            "new_plan_sha256": sha256(CORRECTED_PLAN)}


def main() -> None:
    if RESULT.exists() or ROOT.exists():
        raise RuntimeError("versioned preflight output already exists; refusing overwrite")
    if not SOURCE_RUNNER.is_file() or not FROZEN_PLAN.is_file():
        raise RuntimeError("frozen 0945 runner or plan is unavailable")
    old = run_old_failure()
    corrected = run_corrected_success()
    reducer = run_reducer_custody()
    invariants = compare_plan_invariants()
    result = {
        "schema": "litchi.performance.0945.capture-orchestration-preflight.v1",
        "status": "pass",
        "workloads_started": False,
        "measurements_started": False,
        "frozen_plan": {"path": str(FROZEN_PLAN), "sha256": sha256(FROZEN_PLAN)},
        "frozen_runner": {"path": str(SOURCE_RUNNER), "sha256": sha256(SOURCE_RUNNER)},
        "corrected_runner": {"path": str(CORRECTED_RUNNER), "sha256": sha256(CORRECTED_RUNNER)},
        "corrected_plan": {"path": str(CORRECTED_PLAN), "sha256": sha256(CORRECTED_PLAN)},
        "reducer_v2": {"path": str(REDUCER_V2), "sha256": sha256(REDUCER_V2)},
        "finding": "frozen run_child/stage_capture passed absolute PLAN_PATH to strict descriptor_path",
        "old_frozen_failure": old,
        "corrected_relative_fixture": corrected,
        "reducer_relative_fixture": reducer,
        "plan_invariants": invariants,
        "correction_scope": "two internal plan descriptor call sites use plan_descriptor() with packet-relative plan-v2.json; no gates or schedule changed",
    }
    write_json(RESULT, result)
    print(json.dumps({"status": result["status"], "result": str(RESULT),
                      "bytes": RESULT.stat().st_size}, sort_keys=True))


if __name__ == "__main__":
    main()
