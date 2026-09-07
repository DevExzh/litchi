#!/usr/bin/env python3
"""Capture one serialized 0464 PPTX pair-lifecycle repeat.

The outer change-0464/check.py wrapper owns source custody.  This driver owns
lane ordering, copied binary/pair/input identity checks, GNU time receipts,
portable report/output retention, and the independent package oracle.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import platform
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
PROTOCOL_PATH = ROOT / "protocol.json"
PAIR_PATH = ROOT / "pair.json"
ORACLE_PATH = ROOT / "pair-oracle.py"


class CaptureError(RuntimeError):
    pass


def fail(message: str) -> None:
    raise CaptureError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def artifact(path: Path) -> dict[str, Any]:
    return {
        "path": str(path.relative_to(ROOT)),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }


def import_custody() -> Any:
    spec = importlib.util.spec_from_file_location("change0464_custody", ROOT / "check.py")
    if spec is None or spec.loader is None:
        fail("cannot load change-0464/check.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if not callable(getattr(module, "sources", None)):
        fail("change-0464/check.py has no sources()")
    return module


def machine() -> dict[str, Any]:
    affinity = sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else None
    return {
        "platform": dict(zip(("system", "node", "release", "version", "machine", "processor"), platform.uname())),
        "python": platform.python_version(),
        "logical_cpus": os.cpu_count(),
        "process_affinity": affinity,
        "page_size": os.sysconf("SC_PAGE_SIZE"),
        "kernel_perf_event_paranoid": (
            Path("/proc/sys/kernel/perf_event_paranoid").read_text(encoding="utf-8").strip()
            if Path("/proc/sys/kernel/perf_event_paranoid").is_file() else None
        ),
    }


def fixed_environment() -> dict[str, str]:
    env = os.environ.copy()
    env.update({
        "DEBUGINFOD_URLS": "",
        "PYTHONDONTWRITEBYTECODE": "1",
        "RUSTUP_TOOLCHAIN": "1.98.1",
    })
    return env


def environment_receipt(env: dict[str, str]) -> dict[str, Any]:
    return {
        key: env.get(key)
        for key in sorted(env)
        if key in {
            "RUSTUP_TOOLCHAIN", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
            "PYTHONDONTWRITEBYTECODE", "DEBUGINFOD_URLS", "RUSTFLAGS", "GLIBC_TUNABLES",
        } or key.startswith("MALLOC_")
    }


def pair_inputs(pair: dict[str, Any]) -> tuple[Path, Path]:
    source = ROOT / pair["source"]["path"]
    destination = ROOT / pair["destination"]["path"]
    for path, identity, label in ((source, pair["source"], "source"), (destination, pair["destination"], "destination")):
        require(path.is_file() and not path.is_symlink(), f"pair {label} input missing or symlinked: {path}")
        require(path.stat().st_size == identity["bytes"] and sha(path) == identity["sha256"], f"pair {label} input identity differs")
    return source, destination


def input_identity(pair_path: Path, source: Path, destination: Path) -> dict[str, Any]:
    return {
        "pair_manifest": {"path": str(pair_path.relative_to(ROOT)), "bytes": pair_path.stat().st_size, "sha256": sha(pair_path)},
        "source": {"path": str(source.relative_to(ROOT)), "bytes": source.stat().st_size, "sha256": sha(source)},
        "destination": {"path": str(destination.relative_to(ROOT)), "bytes": destination.stat().st_size, "sha256": sha(destination)},
    }


def load_binding(protocol: dict[str, Any], pair: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any], Any]:
    binding_spec = protocol["binding"]
    binding_path = ROOT / binding_spec["path"]
    require(binding_path.is_file() and not binding_path.is_symlink(), f"binary binding missing: {binding_path}")
    binding_sha = sha(binding_path)
    if binding_spec["sha256"] not in ("bound-at-capture", binding_sha):
        fail("binary binding hash differs from protocol")
    binding = load(binding_path)
    require(binding.get("schema") == protocol["binding_contract"]["schema"], "binary binding schema differs")
    require(isinstance(binding.get("revision"), str) and binding["revision"], "binary binding revision is missing")
    require(binding["revision"] == pair.get("source_revision"), "binary revision differs from pair source_revision")

    binaries = binding.get("binaries")
    require(isinstance(binaries, dict), "binary binding binaries is missing")
    expected = protocol["binaries"]
    source_binding = None
    for mode in ("normal", "allocator"):
        item = binaries.get(mode)
        require(isinstance(item, dict), f"binary binding {mode} entry is missing")
        require(item.get("path") == expected[mode]["path"], f"{mode} binary path differs from protocol")
        path = Path(item["path"])
        require(path.is_file() and not path.is_symlink(), f"{mode} binary missing or symlinked")
        require(item.get("bytes") == path.stat().st_size and item.get("sha256") == sha(path), f"{mode} binary identity differs from binding")
        if "executable" in item:
            require(item.get("executable") is True, f"{mode} binary is not marked executable")
        source = item.get("source")
        require(isinstance(source, dict), f"{mode} binary source binding is missing")
        source_path = ROOT / source["path"]
        require(source_path.is_file() and not source_path.is_symlink(), f"{mode} binary source manifest is missing")
        require(sha(source_path) == source["sha256"] and source.get("files", 0) > 0, f"{mode} binary source identity differs")
        if source_binding is None:
            source_binding = source
        else:
            require(source == source_binding, "normal and allocator source bindings differ")
        build_path = ROOT / item["build_receipt"]
        require(build_path.is_file() and sha(build_path) == item["build_receipt_sha256"], f"{mode} build receipt identity differs")
        build = load(build_path)
        require(build.get("status") == "pass" and build.get("source_unchanged") is True, f"{mode} build receipt is not a successful source-unchanged build")
        require(build.get("source_before") == source and build.get("source_after") == source, f"{mode} build receipt source binding differs")
    require(source_binding is not None, "binary source binding is missing")
    return binding, binaries, source_binding, binding_sha


def workload_argv(protocol: dict[str, Any], lane: dict[str, Any], binary: dict[str, Any], revision: str, report: Path, output: Path, resource: Path) -> list[str]:
    argv = [
        "taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o", str(resource),
        binary["path"], "pptx-pair-lifecycle",
        "--manifest", str(PAIR_PATH), "--provider", lane["provider"],
        "--samples", str(protocol["samples"]), "--warmup", str(protocol["warmups"]),
        "--repeat", lane["repeat"], "--output", str(report), "--output-pptx", str(output),
        "--source-revision", revision,
    ]
    if lane["provider"] == "range":
        argv += ["--max-range", str(protocol["range"]["max_range_bytes"]), "--delay-us", str(protocol["range"]["delay_us"])]
    return argv


def validate_report(report: dict[str, Any], lane: dict[str, Any], binary: dict[str, Any], revision: str, source_identity: dict[str, Any], destination_identity: dict[str, Any], output: Path, protocol: dict[str, Any]) -> None:
    require(report.get("schema") == "pptx_pair_lifecycle_v1", "report schema differs")
    require(report.get("pair_id") == protocol["pair_id"], "report pair_id differs")
    require(report.get("provider") == lane["provider"] and report.get("repeat") == lane["repeat"], "report provider/repeat differs")
    require(report.get("samples") == protocol["samples"] and report.get("warmup") == protocol["warmups"], "report sample configuration differs")
    require(report.get("checked_iteration_count") == protocol["samples"] + protocol["warmups"], "report checked iteration count differs")
    require(report.get("source_revision") == revision, "report source revision differs")
    require(report.get("binary_sha256") == binary["sha256"] and report.get("binary_bytes") == binary["bytes"], "report binary identity differs")
    expected_instrumentation = "system_allocator_operation_scoped" if lane["instrumentation"] == "allocator" else "none"
    require(report.get("instrumentation") == expected_instrumentation, "report allocator instrumentation differs")
    for key, expected in (("source", source_identity), ("destination", destination_identity)):
        value = report.get(key)
        require(isinstance(value, dict) and value.get("sha256") == expected["sha256"] and value.get("bytes") == expected["bytes"], f"report {key} identity differs")
    artifact_value = report.get("output_artifact")
    require(isinstance(artifact_value, dict) and Path(artifact_value.get("path", "")).resolve() == output.resolve(), "report output path differs")
    require(artifact_value.get("bytes") == output.stat().st_size and artifact_value.get("sha256") == sha(output), "report output identity differs")
    rows = report.get("samples_raw")
    require(isinstance(rows, list) and len(rows) == protocol["samples"], "report sample rows differ")
    for row in rows:
        require(row.get("output_sha256") == artifact_value["sha256"] and row.get("output_bytes") == artifact_value["bytes"], "report sample output identity differs")
    limits = report.get("configured_limits")
    require(isinstance(limits, dict) and limits.get("max_range_bytes") == protocol["range"]["max_range_bytes"] and limits.get("fixed_delay_us") == protocol["range"]["delay_us"], "report range limits differ")
    expected_output = protocol["expected_output"]
    require(artifact_value.get("bytes") == expected_output["bytes"], "report output byte count differs from expected pair output")
    operation = report.get("operation")
    require(isinstance(operation, dict) and operation.get("destination_slide_count_after") == expected_output["slides"], "report output slide count differs from expected pair output")


def oracle_argv(protocol: dict[str, Any], pair: dict[str, Any], source: Path, destination: Path, output: Path, report: dict[str, Any]) -> list[str]:
    artifact_value = report["output_artifact"]
    return [
        sys.executable, "-B", str(ORACLE_PATH),
        "--source", str(source), "--destination", str(destination), "--output", str(output),
        "--source-index", str(pair["operation"]["source_slide"]),
        "--insertion-index", str(pair["operation"]["insertion_position"]),
        "--source-sha256", pair["source"]["sha256"], "--source-bytes", str(pair["source"]["bytes"]),
        "--destination-sha256", pair["destination"]["sha256"], "--destination-bytes", str(pair["destination"]["bytes"]),
        "--output-sha256", artifact_value["sha256"], "--output-bytes", str(artifact_value["bytes"]),
    ]


def run_oracle(argv: list[str], env: dict[str, str], log: Path) -> int:
    with log.open("xb") as stream:
        stream.write(("argv=" + json.dumps(argv) + "\n").encode())
        result = subprocess.run(argv, cwd=REPO, env=env, stdout=stream, stderr=subprocess.STDOUT, check=False)
    return result.returncode


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repeat", choices=("R1", "R2"), required=True)
    args = parser.parse_args()

    protocol = load(PROTOCOL_PATH)
    require(protocol.get("schema") == "litchi-0464-pptx-pair-capture-v1", "protocol schema differs")
    require(protocol.get("change") == 464 and protocol.get("status") == "frozen", "protocol identity differs")
    require(sha(ROOT / "capture.py") == protocol["runner"]["sha256"], "capture driver differs from protocol")
    require(sha(ORACLE_PATH) == protocol["oracle"]["sha256"], "pair oracle differs from protocol")
    require(sha(ROOT / "check.py") == protocol["custody"]["check_sha256"], "custody driver differs from protocol")
    require(sha(PAIR_PATH) == protocol["pair_manifest"]["sha256"], "pair manifest differs from protocol")
    host_spec = protocol["host"]
    host_path = ROOT / host_spec["path"]
    require(host_path.is_file() and not host_path.is_symlink(), "host receipt is missing or symlinked")
    require(sha(host_path) == host_spec["sha256"], "host receipt differs from protocol")
    host_artifact = artifact(host_path)
    probe_spec = protocol["expected_output"]["probe"]
    probe_path = ROOT / probe_spec["path"]
    require(probe_path.is_file() and not probe_path.is_symlink(), "expected-output probe is missing or symlinked")
    require(sha(probe_path) == probe_spec["sha256"], "expected-output probe differs from protocol")
    probe_artifact = artifact(probe_path)
    pair = load(PAIR_PATH)
    require(pair.get("pair_id") == protocol["pair_id"], "pair_id differs from protocol")
    source, destination = pair_inputs(pair)
    source_identity = pair["source"]
    destination_identity = pair["destination"]
    binding, binaries, source_binding, binding_sha = load_binding(protocol, pair)
    custody = import_custody()
    source_before = custody.sources()
    require(source_before == source_binding, "current source custody differs from binary binding")
    env = fixed_environment()
    repeat_dir = ROOT / "captures" / args.repeat
    require(not repeat_dir.exists(), f"capture output already exists: {repeat_dir}")
    repeat_dir.mkdir(parents=True)
    host = machine()
    state_path = repeat_dir / "capture-state.json"
    state: dict[str, Any] = {
        "schema": "litchi-0464-pptx-pair-capture-state-v1",
        "change": 464,
        "repeat": args.repeat,
        "status": "running",
        "started_utc": now(),
        "protocol_sha256": sha(PROTOCOL_PATH),
        "pair_manifest_sha256": sha(PAIR_PATH),
        "pair_oracle_sha256": sha(ORACLE_PATH),
        "binding_sha256": binding_sha,
        "binding": binding,
        "host_artifact": host_artifact,
        "expected_output_probe": probe_artifact,
        "source_before": source_before,
        "host": host,
        "environment": environment_receipt(env),
        "expected_lanes": protocol["orders"][args.repeat],
        "completed_lanes": [],
    }
    write_json(state_path, state)
    failures: list[str] = []
    try:
        for lane in protocol["orders"][args.repeat]:
            lane_name = lane["lane"]
            lane_dir = repeat_dir / lane_name
            lane_dir.mkdir()
            report_path = lane_dir / "report.json"
            output_path = lane_dir / "output.pptx"
            resource_path = lane_dir / "resource.log"
            workload_log = lane_dir / "workload.log"
            oracle_log = lane_dir / "oracle.log"
            receipt_path = lane_dir / "receipt.json"
            binary = binaries[lane["instrumentation"]]
            input_before = input_identity(PAIR_PATH, source, destination)
            argv = workload_argv(protocol, lane, binary, binding["revision"], report_path, output_path, resource_path)
            row: dict[str, Any] = {
                "schema": "litchi-0464-pptx-pair-capture-receipt-v1",
                "change": 464,
                "repeat": args.repeat,
                "lane": lane,
                "status": "running",
                "started_utc": now(),
                "argv": argv,
                "cwd": str(REPO),
                "host": host,
                "environment": environment_receipt(env),
                "protocol_sha256": sha(PROTOCOL_PATH),
                "pair_manifest_sha256": sha(PAIR_PATH),
                "pair_oracle_sha256": sha(ORACLE_PATH),
                "binding_sha256": binding_sha,
                "binary": binary,
                "host_artifact": host_artifact,
                "expected_output_probe": probe_artifact,
                "source_binding": source_binding,
                "input_before": input_before,
            }
            write_json(receipt_path, row)
            print(json.dumps({"status": "running", "repeat": args.repeat, "lane": lane_name}), flush=True)
            try:
                with workload_log.open("xb") as stream:
                    result = subprocess.run(argv, cwd=REPO, env=env, stdout=stream, stderr=subprocess.STDOUT, check=False)
                row["exit_code"] = result.returncode
                require(result.returncode == 0, f"workload exited {result.returncode}")
                report = load(report_path)
                validate_report(report, lane, binary, binding["revision"], source_identity, destination_identity, output_path, protocol)
                oracle_command = oracle_argv(protocol, pair, source, destination, output_path, report)
                row["oracle_argv"] = oracle_command
                oracle_exit = run_oracle(oracle_command, env, oracle_log)
                row["oracle_exit_code"] = oracle_exit
                require(oracle_exit == 0, f"independent pair oracle exited {oracle_exit}")
                row["report"] = {"path": str(report_path.relative_to(ROOT)), "sha256": sha(report_path), "bytes": report_path.stat().st_size}
                row["output"] = artifact(output_path)
                row["status"] = "pass"
            except Exception as error:
                row["status"] = "failed"
                row["error"] = repr(error)
                failures.append(f"{lane_name}: {error}")
            finally:
                row["finished_utc"] = now()
                try:
                    row["input_after"] = input_identity(PAIR_PATH, source, destination)
                    row["inputs_unchanged"] = row["input_after"] == row["input_before"]
                except Exception as error:
                    row["input_after_error"] = repr(error)
                    row["inputs_unchanged"] = False
                    failures.append(f"{lane_name}: input identity after lane unavailable")
                if not row["inputs_unchanged"]:
                    row["status"] = "failed"
                    failures.append(f"{lane_name}: pair inputs changed")
                row["source_after"] = custody.sources()
                row["source_unchanged"] = row["source_after"] == source_before
                if not row["source_unchanged"]:
                    row["status"] = "failed"
                    failures.append(f"{lane_name}: source custody changed")
                row["artifacts"] = {
                    name: artifact(path)
                    for name, path in {
                        "report": report_path,
                        "output_pptx": output_path,
                        "resource": resource_path,
                        "workload": workload_log,
                        "oracle": oracle_log,
                    }.items()
                    if path.is_file()
                }
                write_json(receipt_path, row)
            state["completed_lanes"].append(lane_name)
            write_json(state_path, state)
            if failures:
                break
    except Exception as error:
        failures.append(repr(error))
    finally:
        state["finished_utc"] = now()
        state["source_after"] = custody.sources()
        state["source_unchanged"] = state["source_after"] == source_before
        if not state["source_unchanged"]:
            failures.append("source custody changed during repeat")
        state["failures"] = failures
        state["status"] = "pass" if not failures and len(state["completed_lanes"]) == len(state["expected_lanes"]) else "failed"
        write_json(state_path, state)
    print(json.dumps({"status": state["status"], "repeat": args.repeat, "completed_lanes": state["completed_lanes"]}))
    return 0 if state["status"] == "pass" else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except CaptureError as error:
        print(f"capture failed: {error}", file=sys.stderr)
        raise SystemExit(1)
