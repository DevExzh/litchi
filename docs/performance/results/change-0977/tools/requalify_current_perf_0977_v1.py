#!/usr/bin/env python3
"""One bounded offline requalification pass with the current /usr/bin/perf.

This consumes two immutable retained perf.data files and runs exactly four
``perf script`` decodes: corrected original, corrected translated, stock
original, and final-v2 noop original.  It never starts a workload or records
new samples.  Every output is written under a fresh external directory and
the current profiler hash is recorded separately from the prior sealed hash.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import importlib.util
import json
import os
import re
import subprocess
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0977")
OUT = ROOT / "requalify-current-perf-v1"
PERF = Path("/usr/bin/perf")
OLD_PERF_SHA256 = "174229fb48b8557aafa821dd828d883dce47c35c31067e1f31e4a01a91c6c98b"
FINAL_SO = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.so")
FINAL_SO_SHA256 = "2acdc5c06ac672dca0e5a34bf204e0a45451b4bfe2d31c304f4c791f5b2a0a12"
FINAL_SOURCE = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.c")
FINAL_SOURCE_SHA256 = "a30536989f1fcc3e28c5bbe5876b9f01159841ba0a6d6c3aa22706cdcfd23d1a"
FINAL_MAP = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.map")
FINAL_MAP_SHA256 = "173dd1902d46a378e713cd794ede6c4fdac2a184124e8aefecc0c2fda7494883"
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
BINARY_SHA256 = "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb"
BINARY_BUILD_ID = "e5c457384eeebd057932c444efda77618f137399"
ORIGINAL_DATA = Path("/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01/child-01.perf.data")
TRANSLATED_DATA = Path("/home/zhuhe/litchi-performance-0976/aslr-cache-final-v2-0975-v3/child-01.translated.perf.data")
TRANSLATED_LAYOUT = Path("/home/zhuhe/litchi-performance-0976/aslr-cache-final-v2-0975-v3/translated.target-layout.json")
ORIGINAL_CORRECTED_LAYOUT = {"corrected_base": "0x5f33d295c000", "perf_expected_start": "0x5f33d295d000"}
TARGET_PID_RE = re.compile(r"^\s*(?:\S+\s+)?(?P<pid>\d+)(?:/(?P<tid>\d+))?\s+\[(?P<cpu>\d+)\]")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    return {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load_decoder():
    path = ROOT / "tools/focuseddecode_0977.py"
    spec = importlib.util.spec_from_file_location("requalify_current_perf_decoder", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot import focused decoder")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def first_pid(path: Path) -> tuple[int, int]:
    for line in path.open(encoding="utf-8", errors="strict"):
        match = TARGET_PID_RE.match(line)
        if match:
            return int(match.group("pid")), int(match.group("cpu"))
    raise RuntimeError(f"no sample header in {path}")


def env_base() -> dict[str, str]:
    env = os.environ.copy()
    for key in list(env):
        if key.startswith("LITCHI_0973_") or key.startswith("LITCHI_0976_"):
            env.pop(key, None)
    env.update({"LC_ALL": "C", "LANG": "C", "TZ": "UTC"})
    return env


def decoder_env(layout: dict[str, str], log: Path, mode: str) -> dict[str, str]:
    env = env_base()
    env.update({
        "LD_PRELOAD": str(FINAL_SO),
        "LITCHI_0973_SHIM_MODE": mode,
        "LITCHI_0973_SHIM_LOG": str(log.resolve()),
        "LITCHI_0973_TARGET_PATH": str(BINARY.resolve()),
        "LITCHI_0973_TARGET_NAME": BINARY.name,
        "LITCHI_0973_CORRECTED_BASE": layout["corrected_base"],
        "LITCHI_0973_PERF_EXPECTED_START": layout["perf_expected_start"],
        "LITCHI_0976_TARGET_BUILD_ID": BINARY_BUILD_ID,
        "LITCHI_0976_TARGET_SHA256": BINARY_SHA256,
        "LITCHI_0976_EXPECTED_DELTA": "0x1000",
    })
    return env


def parse_summary(decoder: Any, script: Path) -> dict[str, Any]:
    pid, cpu = first_pid(script)
    return decoder._stream_perf_script(
        script, marker=decoder.PARSER.MARKER, marker_dso=decoder.PARSER.BINARY_PATH,
        expected_pid=pid, expected_cpu=cpu, expected_period=1_000_000,
    )


def run_one(name: str, data: Path, layout: dict[str, str], *, mode: str | None, decoder: Any) -> dict[str, Any]:
    script = OUT / f"{name}.perf-script-no-inline.txt"
    stderr = OUT / f"{name}.stderr.txt"
    log = OUT / f"{name}.dwfl.jsonl"
    argv = [str(PERF), "script", "--show-lost-events", "--no-inline", "-F", "cpu,pid,tid,time,event,period,ip,sym,symoff,dso", "-i", str(data)]
    env = env_base() if mode is None else decoder_env(layout, log, mode)
    started = dt.datetime.now(dt.timezone.utc).isoformat()
    with script.open("xb") as out, stderr.open("xb") as err:
        completed = subprocess.run(argv, stdin=subprocess.DEVNULL, stdout=out, stderr=err, env=env, timeout=900, check=False)
    row: dict[str, Any] = {
        "name": name, "argv": argv, "environment": {key: env.get(key, "<unset>") for key in ("LD_PRELOAD", "LITCHI_0973_SHIM_MODE", "LITCHI_0973_SHIM_LOG", "LITCHI_0973_CORRECTED_BASE", "LITCHI_0973_PERF_EXPECTED_START")},
        "started_utc": started, "returncode": completed.returncode, "script": binding(script), "stderr": binding(stderr),
        "input": binding(data), "log": binding(log) if log.is_file() else None,
    }
    if completed.returncode == 0:
        try:
            summary = parse_summary(decoder, script)
            row["summary"] = {key: summary[key] for key in ("sample_count", "total_period", "marker_sample_count", "marker_period", "marker_period_share", "empty_callchain_count", "lost_events", "unresolved_frames")}
            row["quality"] = {
                "sample_count_6940": summary["sample_count"] == 6940,
                "period_6940000000": summary["total_period"] == 6_940_000_000,
                "marker_minimum_100": summary["marker_sample_count"] >= 100,
                "empty_zero": summary["empty_callchain_count"] == 0,
                "script_lost_zero": summary["lost_events"] == 0,
            }
        except Exception as exc:  # retain command output and classify parse separately
            row["parse_error"] = str(exc)
    else:
        row["stderr_tail"] = stderr.read_text(encoding="utf-8", errors="replace")[-4000:]
    return row


def main() -> int:
    if OUT.exists():
        raise SystemExit(f"refusing to reuse output directory: {OUT}")
    for path in (PERF, FINAL_SO, FINAL_SOURCE, FINAL_MAP, BINARY, ORIGINAL_DATA, TRANSLATED_DATA, TRANSLATED_LAYOUT):
        if not path.is_file():
            raise SystemExit(f"missing bound input: {path}")
    before = {str(path): binding(path) for path in (PERF, FINAL_SO, FINAL_SOURCE, FINAL_MAP, BINARY, ORIGINAL_DATA, TRANSLATED_DATA, TRANSLATED_LAYOUT)}
    if before[str(PERF)]["sha256"] == OLD_PERF_SHA256:
        raise SystemExit("current perf unexpectedly equals prior sealed hash; drift receipt would be stale")
    if before[str(FINAL_SO)]["sha256"] != FINAL_SO_SHA256 or before[str(FINAL_SOURCE)]["sha256"] != FINAL_SOURCE_SHA256 or before[str(FINAL_MAP)]["sha256"] != FINAL_MAP_SHA256 or before[str(BINARY)]["sha256"] != BINARY_SHA256:
        raise SystemExit("final decoder/binary custody changed")
    decoder = load_decoder()
    OUT.mkdir(parents=True)
    translated_layout = json.loads(TRANSLATED_LAYOUT.read_text(encoding="utf-8"))
    results = [
        run_one("original-corrected", ORIGINAL_DATA, ORIGINAL_CORRECTED_LAYOUT, mode="corrected", decoder=decoder),
        run_one("translated-corrected", TRANSLATED_DATA, {"corrected_base": translated_layout["corrected_base"], "perf_expected_start": translated_layout["perf_expected_start"]}, mode="corrected", decoder=decoder),
        run_one("original-stock", ORIGINAL_DATA, ORIGINAL_CORRECTED_LAYOUT, mode=None, decoder=decoder),
        run_one("original-noop", ORIGINAL_DATA, ORIGINAL_CORRECTED_LAYOUT, mode="noop", decoder=decoder),
    ]
    by_name = {row["name"]: row for row in results}
    identity: dict[str, Any] = {}
    stock = OUT / "original-stock.perf-script-no-inline.txt"
    noop = OUT / "original-noop.perf-script-no-inline.txt"
    stock_err = OUT / "original-stock.stderr.txt"
    noop_err = OUT / "original-noop.stderr.txt"
    if stock.is_file() and noop.is_file():
        identity.update({"stock_noop_script_byte_identical": sha256(stock) == sha256(noop), "stock_noop_script_sha256": sha256(stock), "noop_script_sha256": sha256(noop)})
    if stock_err.is_file() and noop_err.is_file():
        identity["stock_noop_stderr_byte_identical"] = sha256(stock_err) == sha256(noop_err)
    receipt = {
        "schema": "litchi.performance.0977.current_perf_requalification.v1",
        "status": "complete",
        "native_workload_started": False,
        "build_started": False,
        "scope": "four offline perf-script commands over retained raw captures; no native capture/build",
        "tool": {"current": before[str(PERF)], "prior_sealed_sha256": OLD_PERF_SHA256},
        "final_decoder_custody": {key: before[str(key)] for key in (str(FINAL_SO), str(FINAL_SOURCE), str(FINAL_MAP), str(BINARY))},
        "raw_inputs_before": {key: before[str(key)] for key in (str(ORIGINAL_DATA), str(TRANSLATED_DATA), str(TRANSLATED_LAYOUT))},
        "commands": results,
        "identity_controls": identity,
        "comparison_reference": {"retained_final_v2_sample_count": 6940, "retained_final_v2_total_period": 6_940_000_000, "retained_final_v2_marker_sample_count": 951, "retained_final_v2_empty_callchain_count": 0, "retained_final_v2_raw_loss": 0, "exact_marker_count_is_not_an_admission_gate": True},
        "raw_inputs_after": {key: binding(Path(key)) for key in (str(ORIGINAL_DATA), str(TRANSLATED_DATA), str(TRANSLATED_LAYOUT))},
    }
    receipt_path = ROOT / "receipts/current-perf-requalification-0977-v1.json"
    receipt_path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    receipt["receipt"] = {"path": str(receipt_path), "sha256": sha256(receipt_path), "bytes": receipt_path.stat().st_size}
    # The receipt is written once above; report its binding after the write.
    print(json.dumps({"status": receipt["status"], "receipt": str(receipt_path), "receipt_sha256": sha256(receipt_path), "commands": [{"name": row["name"], "returncode": row["returncode"], "summary": row.get("summary"), "parse_error": row.get("parse_error")} for row in results], "identity": identity}, sort_keys=True))
    return 0 if all(row["returncode"] == 0 and "parse_error" not in row for row in results) else 2


if __name__ == "__main__":
    raise SystemExit(main())
