#!/usr/bin/env python3
"""Static admission controls for the 0976 offline cache shim.

This test does not load the shared object or invoke perf.  It independently
checks the exact admission truth table and source-level lifecycle/identity
contracts before a final helper build is allowed.
"""
from __future__ import annotations

import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.c"
MAP = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.map"
OUT = ROOT / "receipts/cache-guard-static-0976-v2.json"


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def admission(v: dict[str, object]) -> bool:
    return bool(
        v["config_ok"]
        and v["mode"] == "corrected"
        and v["path_match"]
        and v["name_match"]
        and v["caller_perf"]
        and v["caller_file"] == "/usr/bin/perf"
        and v["fd"] == -1
        and not v["add_p_vaddr"]
        and v["base"] in (v["expected"], v["corrected"])
    )


def main() -> int:
    source = SOURCE.read_text()
    map_text = MAP.read_text()
    required = [
        '"ELFUTILS_0.127"',  # checked below with map text too
        '"LITCHI_0976_EXPECTED_DELTA"',
        'parse_build_id',
        'REQUIRED_TARGET_BUILD_ID',
        'memcmp(build_id, REQUIRED_TARGET_BUILD_ID',
        'parse_sha256',
        "*text == '-'",
        "*text == '+'",
        "*perf_expected_start - *corrected_base != delta",
        "fd == -1",
        "!add_p_vaddr",
        'strcmp(caller_file, "/usr/bin/perf") == 0',
        "base == perf_expected_start || base == corrected_base",
        "initial_rewrite && result != NULL",
        "dwfl_module_getelf",
        "dwfl_module_build_id",
        "stat_identity",
        "dwfl_addrmodule",
        "pthread_mutex_lock(&cache_mutex)",
        "clear_cache_locked(dwfl)",
        "void dwfl_end(Dwfl *dwfl)",
    ]
    missing = [needle for needle in required if needle not in source]
    if missing:
        raise AssertionError(f"missing source guards: {missing}")
    for needle in (
        "ELFUTILS_0.156",
        "ELFUTILS_0.127",
        "ELFUTILS_0.122",
    ):
        if needle not in map_text:
            raise AssertionError(f"missing map version {needle}")
    if "dwfl_report_begin_add" not in map_text:
        raise AssertionError("begin_add is not exported")
    if "dwfl_end" not in map_text:
        raise AssertionError("dwfl_end is not exported")
    if "if (locked && result != NULL)" in source:
        raise AssertionError("generic cache store remains")

    common = {
        "config_ok": True,
        "mode": "corrected",
        "path_match": True,
        "name_match": True,
        "caller_perf": True,
        "caller_file": "/usr/bin/perf",
        "fd": -1,
        "add_p_vaddr": False,
        "expected": 0x1000,
        "corrected": 0x2000,
    }
    controls = {
        "baseline_initial": {**common, "base": 0x1000},
        "baseline_reuse": {**common, "base": 0x2000},
        "wrong_caller": {**common, "base": 0x1000, "caller_perf": False, "caller_file": "/tmp/helper"},
        "wrong_name": {**common, "base": 0x1000, "name_match": False},
        "wrong_path": {**common, "base": 0x1000, "path_match": False},
        "wrong_base": {**common, "base": 0x3000},
        "fd_reuse": {**common, "base": 0x1000, "fd": 17},
        "add_p_vaddr": {**common, "base": 0x1000, "add_p_vaddr": True},
        "bad_build_or_delta": {**common, "base": 0x1000, "config_ok": False},
    }
    expected = {
        "baseline_initial": True,
        "baseline_reuse": True,
        "wrong_caller": False,
        "wrong_name": False,
        "wrong_path": False,
        "wrong_base": False,
        "fd_reuse": False,
        "add_p_vaddr": False,
        "bad_build_or_delta": False,
    }
    observed = {name: admission(value) for name, value in controls.items()}
    if observed != expected:
        raise AssertionError({"expected": expected, "observed": observed})
    if not admission(controls["baseline_initial"]):
        raise AssertionError("initial baseline was rejected")
    if controls["baseline_initial"]["base"] == controls["baseline_reuse"]["base"]:
        raise AssertionError("paired initial/reuse controls are not distinct")

    result = {
        "schema": "litchi.performance.0976.cache_guard_static.v2",
        "offline_only": True,
        "native_workload_started": False,
        "source": {"path": str(SOURCE), "bytes": SOURCE.stat().st_size, "sha256": sha256(SOURCE)},
        "map": {"path": str(MAP), "bytes": MAP.stat().st_size, "sha256": sha256(MAP)},
        "controls": {name: {"admitted": observed[name], "expected": expected[name]} for name in controls},
        "lifecycle_contracts": {
            "report_begin": "clear before underlying call",
            "report_begin_add": "ELFUTILS_0.127 and clear before underlying call",
            "report_end": "clear after underlying call",
            "dwfl_end": "clear before underlying free",
        },
        "limitations": [
            "source/static controls do not prove runtime libdw symbol behavior",
            "positive replay and translated raw control remain pending reviewer release",
        ],
    }
    OUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"receipt": str(OUT), "sha256": sha256(OUT), "controls": observed}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
