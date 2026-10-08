#!/usr/bin/env python3
"""Run the frozen shim through a real libdwfl context, without perf/workload.

The driver uses a dladdr-only caller stub so the final shim's exact
/usr/bin/perf admission branch can be exercised from an isolated executable.
The stub is explicit test plumbing; it cannot change libdwfl arguments or ELF
contents.  A no-stub run exercises the actual wrong-caller rejection.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
import subprocess
from pathlib import Path

ROOT = Path("/home/zhuhe/litchi-performance-0976")
OUT = ROOT / "dynamic-cache-guard-0976-v5"
RECEIPT = ROOT / "receipts/dynamic-cache-guard-0976-v5.json"
DRIVER = ROOT / "tools/cache_guard_driver_0976"
STUB = ROOT / "tools/cache_guard_dladdr_stub_0976.so"
SHIM = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.so"
SOURCE = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.c"
MAP = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.map"
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
LAYOUT = Path("/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01/child-01.target-layout.json")
BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
TARGET_SHA256 = "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb"
DELTA = 0x1000


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def binding(path: Path) -> dict[str, object]:
    return {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": sha256(path)}


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def run_case(name: str, *, caller_stub: bool, build_id: str = BUILD_ID) -> dict[str, object]:
    log = OUT / f"{name}.dwfl.jsonl"
    stderr = OUT / f"{name}.stderr.txt"
    if log.exists() or stderr.exists():
        raise FileExistsError(f"refusing to overwrite {log} or {stderr}")
    layout = json.loads(LAYOUT.read_text())
    expected = int(layout["perf_expected_start"], 16)
    corrected = expected - DELTA
    env = os.environ.copy()
    env.update({
        "LC_ALL": "C",
        "LANG": "C",
        "TZ": "UTC",
        "LITCHI_0973_SHIM_MODE": "corrected",
        "LITCHI_0973_SHIM_LOG": str(log),
        "LITCHI_0973_TARGET_PATH": str(BINARY.resolve()),
        "LITCHI_0973_TARGET_NAME": BINARY.name,
        "LITCHI_0973_CORRECTED_BASE": f"0x{corrected:x}",
        "LITCHI_0973_PERF_EXPECTED_START": f"0x{expected:x}",
        "LITCHI_0976_TARGET_BUILD_ID": build_id,
        "LITCHI_0976_TARGET_SHA256": TARGET_SHA256,
        "LITCHI_0976_EXPECTED_DELTA": f"0x{DELTA:x}",
        "LD_PRELOAD": ":".join(str(p) for p in ([STUB, SHIM] if caller_stub else [SHIM])),
    })
    argv = [str(DRIVER), str(BINARY.resolve()), BINARY.name, f"0x{expected:x}"]
    started = utc()
    with stderr.open("wb") as err:
        cp = subprocess.run(argv, stdout=subprocess.PIPE, stderr=err, env=env,
                            stdin=subprocess.DEVNULL, timeout=120, check=False)
    log_rows = []
    if log.is_file():
        log_rows = [json.loads(line) for line in log.read_text().splitlines() if line]
    report_rows = [row for row in log_rows
                   if row.get("schema") == "litchi.performance.0976.dwfl_report.final"]
    return {
        "name": name,
        "caller_stub": caller_stub,
        "argv": argv,
        "environment": {key: env[key] for key in (
            "LITCHI_0973_SHIM_MODE", "LITCHI_0973_TARGET_PATH",
            "LITCHI_0973_TARGET_NAME", "LITCHI_0973_CORRECTED_BASE",
            "LITCHI_0973_PERF_EXPECTED_START", "LITCHI_0976_TARGET_BUILD_ID",
            "LITCHI_0976_TARGET_SHA256", "LITCHI_0976_EXPECTED_DELTA", "LD_PRELOAD")},
        "started_utc": started,
        "returncode": cp.returncode,
        "stdout": cp.stdout.decode("utf-8", "replace"),
        "stderr": binding(stderr),
        "log": binding(log) if log.exists() else None,
        "report_rows": len(report_rows),
        "target_rows": [row for row in report_rows
                        if row.get("path_match") is True and row.get("name_match") is True],
        "all_reports": report_rows,
    }


def check_stub_case(result: dict[str, object]) -> dict[str, object]:
    if result["returncode"] != 0:
        raise AssertionError(f"dynamic stub driver failed: {result}")
    rows = result["target_rows"]
    if len(rows) != 7:
        raise AssertionError(f"expected 7 exact path/name rows including 3 argument negatives, got {len(rows)}")
    positive_rows = [row for row in rows if row.get("admission") is True]
    if len(positive_rows) != 4:
        raise AssertionError(f"expected 4 admitted target rows, got {len(positive_rows)}")
    initial = [row for row in positive_rows if row.get("initial_rewrite") is True]
    reused = [row for row in positive_rows if row.get("reused_existing") is True]
    if len(initial) != 2 or len(reused) != 2:
        raise AssertionError(f"lifecycle counts differ: initial={len(initial)} reuse={len(reused)}")
    if any(row.get("identity_valid") is not True for row in positive_rows):
        raise AssertionError("admitted target row lacked identity validation")
    negative = [row for row in result["all_reports"]
                if row.get("admission") is False]
    # wrong name/path/base/fd/add each make one actual report call.
    if len(negative) != 5:
        raise AssertionError(f"expected 5 rejected argument reports, got {len(negative)}")
    if any(row.get("forwarded_base") != row.get("original_base") for row in negative):
        raise AssertionError("a rejected dynamic report changed its base")
    if any(row.get("cache_stored") is True or row.get("reused_existing") is True for row in negative):
        raise AssertionError("a rejected dynamic report changed or reused the cache")
    return {
        "status": "pass",
        "exact_target_rows": len(positive_rows),
        "initial_rewrite_rows": len(initial),
        "reuse_rows": len(reused),
        "rejected_argument_rows": len(negative),
        "forwarded_unchanged": True,
        "cache_untouched_by_rejections": True,
    }


def check_wrong_caller(result: dict[str, object]) -> dict[str, object]:
    if result["returncode"] not in (0, 4):
        raise AssertionError(f"wrong-caller driver unexpected exit: {result}")
    exact = [row for row in result["all_reports"]
             if row.get("path_match") is True and row.get("name_match") is True]
    if len(exact) != 3 or any(row.get("admission") is not False for row in exact):
        raise AssertionError("caller stub removal did not reject all exact target calls")
    if any(row.get("forwarded_base") != row.get("original_base") for row in exact):
        raise AssertionError("wrong caller changed a base")
    return {"status": "pass", "exact_target_rows": len(exact), "all_rejected": True}


def check_bad_config(result: dict[str, object]) -> dict[str, object]:
    if result["returncode"] not in (0, 4):
        raise AssertionError(f"bad-config driver unexpected exit: {result}")
    exact = [row for row in result["all_reports"]
             if row.get("path_match") is True and row.get("name_match") is True]
    if not exact or any(row.get("config_valid") is not False or row.get("admission") is not False for row in exact):
        raise AssertionError("malformed build ID was accepted")
    if any(row.get("forwarded_base") != row.get("original_base") for row in exact):
        raise AssertionError("bad config changed a base")
    return {"status": "pass", "exact_target_rows": len(exact), "all_rejected": True}


def main() -> int:
    if OUT.exists():
        raise FileExistsError(f"refusing to overwrite {OUT}")
    for path in (DRIVER, STUB, SHIM, SOURCE, MAP, BINARY, LAYOUT):
        if not path.is_file():
            raise FileNotFoundError(path)
    OUT.mkdir(parents=True)
    input_paths = (DRIVER, STUB, SHIM, SOURCE, MAP, BINARY, LAYOUT)
    before = {str(path): binding(path) for path in input_paths}
    stub_case = run_case("caller-stub", caller_stub=True)
    wrong_caller = run_case("wrong-caller", caller_stub=False)
    bad_config = run_case("bad-build-id", caller_stub=True, build_id="00" * 20)
    checks = {
        "caller_stub_matrix": check_stub_case(stub_case),
        "wrong_caller": check_wrong_caller(wrong_caller),
        "bad_build_id": check_bad_config(bad_config),
    }
    after = {str(path): binding(path) for path in input_paths}
    if before != after:
        raise AssertionError("input custody changed during dynamic guard control")
    receipt = {
        "schema": "litchi.performance.0976.dynamic_cache_guard.v5",
        "status": "pass",
        "created_utc": utc(),
        "purpose": "isolated real-libdwfl shim branch/control test; no perf, workload, or native target execution",
        "previous_attempts": [
            {"path": "/home/zhuhe/litchi-performance-0976/dynamic-cache-guard-0976-v1", "status": "failed_expected_count_due_to_wrong-build-id-admission-found", "preserved": True},
            {"path": "/home/zhuhe/litchi-performance-0976/dynamic-cache-guard-0976-v2", "status": "test_assertion_expected_four_instead_of_seven_path_name_rows", "preserved": True},
            {"path": "/home/zhuhe/litchi-performance-0976/dynamic-cache-guard-0976-v3", "status": "wrong-caller-driver-exit4-was-treated-as-failure-even-though-forwarding-was-correct", "preserved": True},
            {"path": "/home/zhuhe/litchi-performance-0976/dynamic-cache-guard-0976-v4", "status": "bad-config-driver-exit4-was-treated-as-failure-even-though-forwarding-was-correct", "preserved": True},
        ],
        "caller_stub_scope": "cache_guard_dladdr_stub changes only dladdr's reported caller DSO for the driver callsite to /usr/bin/perf; it does not change libdwfl arguments, ELF bytes, or target execution",
        "inputs": {str(path): binding(path) for path in input_paths},
        "binary_custody": {
            "target_sha256_expected": TARGET_SHA256,
            "target_sha256_before": before[str(BINARY)]["sha256"],
            "target_sha256_after": after[str(BINARY)]["sha256"],
            "unchanged": before[str(BINARY)] == after[str(BINARY)],
        },
        "build_provenance": {
            "shim_argv": [
                "cc", "-shared", "-fPIC", "-O2", "-std=c11", "-Wall", "-Wextra", "-Werror", "-pthread",
                "-I/home/zhuhe/litchi-performance-0972/header-v2/root/usr/include",
                "-Wl,--version-script=/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.map",
                "-o", str(SHIM), str(SOURCE), "-ldl", "-pthread",
            ],
            "driver_argv": [
                "cc", "-O2", "-std=c11", "-Wall", "-Wextra", "-Werror",
                "-I/home/zhuhe/litchi-performance-0972/header-v2/root/usr/include",
                "-o", str(DRIVER), str(ROOT / "tools/cache_guard_driver_0976.c"),
                "-L/usr/lib/x86_64-linux-gnu", "-l:libdw.so.1", "-lelf", "-ldl", "-pthread",
            ],
            "stub_argv": ["cc", "-shared", "-fPIC", "-O2", "-std=c11", "-Wall", "-Wextra", "-Werror", "-o", str(STUB), str(ROOT / "tools/cache_guard_dladdr_stub_0976.c"), "-ldl"],
            "warning_clean": True,
        },
        "checks": checks,
        "cases": {
            "caller_stub": stub_case,
            "wrong_caller": wrong_caller,
            "bad_build_id": bad_config,
        },
        "limitations": [
            "The caller stub is an isolated admission-branch control, not proof that an arbitrary DSO is perf.",
            "The target executable is opened by libdwfl for offline identity checks and never executed.",
            "The retained raw translated control separately tests the full perf decoder path.",
        ],
        "native_workload_started": False,
        "native_capture_started": False,
        "native_target_executed": False,
    }
    RECEIPT.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"status": "pass", "receipt": str(RECEIPT), "receipt_sha256": sha256(RECEIPT), "checks": checks}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
