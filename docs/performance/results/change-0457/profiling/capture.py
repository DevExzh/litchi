#!/usr/bin/env python3
"""Capture one bounded 0457 large normal endpoint CPU profile.

The command is deliberately small: one fresh process, 100 measured workload
operations, and one retained perf.data file.  Invoke it through change-0457's
check.py so the parent source manifest is recorded before and after the run.
"""

from __future__ import annotations

import argparse
import collections
import datetime as dt
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Any


SCRIPT = Path(__file__).resolve()
PROFILE_ROOT = SCRIPT.parent
BUNDLE = PROFILE_ROOT.parent
REPO = BUNDLE.parents[3]
PROTOCOL_PATH = PROFILE_ROOT / "protocol.json"


class ProfileError(RuntimeError):
    pass


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ProfileError(f"{path}: invalid JSON: {error}") from error


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
        "path": str(path.relative_to(PROFILE_ROOT)),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ProfileError(message)


def import_custody() -> Any:
    custody_path = BUNDLE / "check.py"
    spec = importlib.util.spec_from_file_location("change0457_profile_custody", custody_path)
    if spec is None or spec.loader is None:
        raise ProfileError(f"cannot import custody driver {custody_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if not callable(getattr(module, "sources", None)):
        raise ProfileError("check.py has no callable sources()")
    return module


def role_bindings(protocol: dict[str, Any], role_name: str) -> tuple[dict[str, Any], dict[str, Any], Any, dict[str, Any]]:
    role = protocol["roles"][role_name]
    for key in ("workload_protocol", "binary_binding", "build_receipt", "source_manifest"):
        bound = role[key]
        path = BUNDLE / bound["path"]
        require(path.is_file() and not path.is_symlink(), f"missing bound {role_name} {key}: {path}")
        require(sha(path) == bound["sha256"], f"{role_name} {key} hash differs from profiling protocol")

    source_path = BUNDLE / role["source_manifest"]["path"]
    source_map = load(source_path)
    require(isinstance(source_map, dict) and len(source_map) == role["source_manifest"]["files"],
            f"{role_name} source manifest file count differs from profiling protocol")
    source_record = {
        "path": f"sources/{role['source_manifest']['sha256']}.json",
        "sha256": role["source_manifest"]["sha256"],
        "files": role["source_manifest"]["files"],
    }

    build = load(BUNDLE / role["build_receipt"]["path"])
    require(build.get("change") == 457 and build.get("status") == "pass" and build.get("source_unchanged") is True,
            f"{role_name} build receipt is not a successful source-unchanged build")
    require(build.get("source_before") == source_record and build.get("source_after") == source_record,
            f"{role_name} build receipt does not bind the frozen source manifest")

    bindings = load(BUNDLE / role["binary_binding"]["path"])
    binary = bindings.get("binaries", {}).get(role["binary"]["mode"])
    require(isinstance(binary, dict), f"{role_name} normal binary binding is missing")
    expected_binary = {
        "copy_path": role["binary"]["path"],
        "bytes": role["binary"]["bytes"],
        "sha256": role["binary"]["sha256"],
    }
    for key, expected in expected_binary.items():
        require(binary.get(key) == expected, f"{role_name} binary {key} differs from profiling protocol")
    binary_path = Path(binary["copy_path"])
    require(binary_path.is_file() and not binary_path.is_symlink(), f"{role_name} copied binary is missing or symlinked")
    require(binary_path.stat().st_size == binary["bytes"] and sha(binary_path) == binary["sha256"],
            f"{role_name} copied binary bytes or hash differ from binding")
    return role, bindings, binary, source_record


def ambient_allocator() -> dict[str, str]:
    return {
        key: value
        for key, value in sorted(os.environ.items())
        if key.startswith("MALLOC_") or key == "GLIBC_TUNABLES"
    }


def fixed_environment() -> dict[str, str]:
    require(os.environ.get("RUSTFLAGS") in (None, ""), "RUSTFLAGS must be unset for the normal profile")
    require(not ambient_allocator(), "allocator environment must be empty; profiling never tunes it")
    env = os.environ.copy()
    env.update({
        "DEBUGINFOD_URLS": "",
        "PYTHONDONTWRITEBYTECODE": "1",
        "RUSTUP_TOOLCHAIN": "1.98.1",
    })
    env.pop("RUSTFLAGS", None)
    return env


def run_to_files(argv: list[str], cwd: Path, env: dict[str, str], stdout_path: Path, stderr_path: Path) -> int:
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        result = subprocess.run(argv, cwd=cwd, env=env, stdout=stdout, stderr=stderr, check=False)
    return result.returncode


def run_oracle(argv: list[str], cwd: Path, env: dict[str, str], output: Path) -> int:
    with output.open("wb") as stream:
        stream.write(("argv=" + json.dumps(argv) + "\n").encode())
        result = subprocess.run(argv, cwd=cwd, env=env, stdout=stream, stderr=subprocess.STDOUT, check=False)
    return result.returncode


HEADER = re.compile(r"^\S.*:\s+\d+\s+\S+:\s*$")
PERIOD = re.compile(r":\s*(\d+)\s+\S+:\s*$")


def collapse_perf_script(path: Path) -> tuple[str, dict[str, int]]:
    """Collapse default perf-script blocks without requiring FlameGraph tools."""
    weights: collections.Counter[tuple[str, ...]] = collections.Counter()
    frames: list[str] = []
    period = 0
    blocks = 0
    malformed = 0

    def flush() -> None:
        nonlocal frames, period, blocks, malformed
        if not frames or period <= 0:
            if frames:
                malformed += 1
            frames = []
            period = 0
            return
        weights[tuple(reversed(frames))] += period
        blocks += 1
        frames = []
        period = 0

    for raw in path.read_text(encoding="utf-8", errors="replace").splitlines():
        if HEADER.match(raw):
            flush()
            match = PERIOD.search(raw)
            period = int(match.group(1)) if match else 0
            continue
        if not raw.strip():
            flush()
            continue
        if raw[:1].isspace():
            frame = raw.strip().replace(";", ",")
            if frame and not frame.startswith("#"):
                frames.append(frame)
            continue
        if frames:
            flush()
    flush()

    lines = [
        "# generated from perf script; stacks are root-to-leaf and weighted by perf period",
        f"# parsed_blocks={blocks} unique_stacks={len(weights)} malformed_blocks={malformed}",
    ]
    lines.extend(f"{' ; '.join(stack).replace(' ; ', ';')} {weight}" for stack, weight in sorted(weights.items(), key=lambda item: (-item[1], item[0])))
    return "\n".join(lines) + "\n", {
        "parsed_blocks": blocks,
        "unique_stacks": len(weights),
        "malformed_blocks": malformed,
        "total_period": sum(weights.values()),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--role", choices=("candidate", "control"), required=True)
    args = parser.parse_args()

    protocol = load(PROTOCOL_PATH)
    require(protocol.get("schema") == "litchi-0457-odp-source-tail-profiling-v1", "profiling protocol schema differs")
    require(protocol.get("change") == 457 and protocol.get("status") == "frozen-diagnostic-plan", "profiling protocol identity differs")
    require(sha(SCRIPT) == protocol["custody"]["runner_sha256"], "profiling runner differs from frozen runner hash")
    require(sha(BUNDLE / protocol["custody"]["wrapper_path"]) == protocol["custody"]["wrapper_sha256"], "custody wrapper differs from protocol")

    role, bindings, binary, source_record = role_bindings(protocol, args.role)
    verifier = BUNDLE / role["oracle"]["verifier_path"]
    require(verifier.is_file() and not verifier.is_symlink(), f"missing {args.role} oracle verifier")
    require(sha(verifier) == role["oracle"]["verifier_sha256"], f"{args.role} oracle verifier hash differs")
    custody = import_custody()
    source_before = custody.sources()
    require(source_before == source_record, f"ambient source manifest differs from bound {args.role} source")
    env = fixed_environment()
    role_dir = PROFILE_ROOT / "runs" / f"{args.role}-large"
    require(not role_dir.exists(), f"profile output already exists; refusing to overwrite: {role_dir}")
    role_dir.mkdir(parents=True)

    receipt_path = role_dir / "profile-receipt.json"
    protocol_sha = sha(PROTOCOL_PATH)
    state: dict[str, Any] = {
        "schema": "litchi-0457-large-profile-receipt-v1",
        "change": 457,
        "role": args.role,
        "label": role["label"],
        "status": "running",
        "started_utc": now(),
        "check_tag": f"profile-{args.role}-large",
        "protocol_sha256": protocol_sha,
        "runner_sha256": sha(SCRIPT),
        "custody_wrapper_sha256": sha(BUNDLE / "check.py"),
        "source_before": source_before,
        "binary": dict(binary),
        "bound_files": {
            "workload_protocol": {"path": role["workload_protocol"]["path"], "sha256": role["workload_protocol"]["sha256"]},
            "binary_binding": {"path": role["binary_binding"]["path"], "sha256": role["binary_binding"]["sha256"]},
            "build_receipt": {"path": role["build_receipt"]["path"], "sha256": role["build_receipt"]["sha256"]},
            "source_manifest": {"path": role["source_manifest"]["path"], "sha256": role["source_manifest"]["sha256"]},
        },
        "allocator_environment": ambient_allocator(),
        "artifacts": {},
    }
    write_json(receipt_path, state)

    report = role_dir / "report.json"
    catalog = role_dir / "catalog.json"
    resource = role_dir / "resource.log"
    perf_data = role_dir / "perf.data"
    record_stdout = role_dir / "perf-record.stdout"
    record_stderr = role_dir / "perf-record.stderr"
    help_stdout = role_dir / "perf-help.stdout"
    help_stderr = role_dir / "perf-help.stderr"
    top_symbols = role_dir / "top-symbols.txt"
    top_stderr = role_dir / "top-symbols.stderr"
    perf_script = role_dir / "perf-script.txt"
    perf_script_stderr = role_dir / "perf-script.stderr"
    folded = role_dir / "stacks.folded"
    oracle_log = role_dir / "oracle.log"
    commands: dict[str, list[str]] = {}
    failures: list[str] = []
    postprocess: dict[str, Any] = {}

    try:
        require(shutil.which("perf") is not None, "perf is unavailable on PATH")
        probe = ["perf", "record", "--call-graph", "help"]
        commands["call_graph_probe"] = probe
        probe_rc = run_to_files(probe, REPO, env, help_stdout, help_stderr)
        probe_text = (help_stdout.read_text(encoding="utf-8", errors="replace") + "\n" +
                      help_stderr.read_text(encoding="utf-8", errors="replace"))
        dwarf_supported = "dwarf" in probe_text.lower()
        call_graph = protocol["scope"]["call_graph"]["dwarf"] if dwarf_supported else None
        state["call_graph"] = {
            "requested": protocol["scope"]["call_graph"]["dwarf"],
            "selected": call_graph,
            "dwarf_advertised": dwarf_supported,
            "probe_exit_code": probe_rc,
        }
        cli = protocol["workload_cli"]
        record_argv = [
            "taskset", "-c", str(protocol["scope"]["cpu"]), "/usr/bin/time", "-v", "-o", str(resource),
            "perf", "record", "--no-buildid-cache", "-F", str(protocol["scope"]["frequency_hz"]),
            "-e", protocol["scope"]["event"],
        ]
        if call_graph:
            record_argv += ["--call-graph", call_graph]
        record_argv += [
            "-o", str(perf_data), "--", str(binary["copy_path"]), cli["case_flag"], role["selector"],
            cli["shape_flag"], protocol["scope"]["shape"], cli["workers_flag"], str(protocol["scope"]["workers"]),
            cli["samples_flag"], str(protocol["scope"]["samples"]), cli["warmup_flag"], str(protocol["scope"]["warmups"]),
            cli["report_flag"], str(report), cli["catalog_flag"], str(catalog),
        ]
        commands["record"] = record_argv
        record_rc = run_to_files(record_argv, REPO, env, record_stdout, record_stderr)
        state["record_exit_code"] = record_rc
        if record_rc != 0:
            failures.append(f"perf record exited {record_rc}")

        if perf_data.is_file() and perf_data.stat().st_size:
            report_argv = ["perf", "report", "--stdio", "--no-inline", "--no-children", "--call-graph", "none", "--percent-limit", "0", "-i", str(perf_data)]
            commands["flat_report"] = report_argv
            report_rc = run_to_files(report_argv, REPO, env, top_symbols, top_stderr)
            postprocess["flat_report_exit_code"] = report_rc
            if report_rc != 0:
                failures.append(f"perf report exited {report_rc}")
            script_argv = ["perf", "script", "--no-inline", "-i", str(perf_data)]
            commands["script"] = script_argv
            script_rc = run_to_files(script_argv, REPO, env, perf_script, perf_script_stderr)
            postprocess["script_exit_code"] = script_rc
            if script_rc != 0:
                failures.append(f"perf script exited {script_rc}")
            if perf_script.is_file() and perf_script.stat().st_size:
                folded_text, folded_stats = collapse_perf_script(perf_script)
                folded.write_text(folded_text, encoding="utf-8")
                postprocess["folded"] = folded_stats
            else:
                folded.write_text("# unavailable: perf script produced no text\n", encoding="utf-8")
                postprocess["folded"] = {"status": "unavailable", "reason": "perf script produced no text"}
        else:
            folded.write_text("# unavailable: perf.data was not produced\n", encoding="utf-8")
            postprocess["data"] = {"status": "unavailable", "reason": "perf record produced no perf.data"}
            failures.append("perf.data was not produced")

        if report.is_file():
            oracle_argv = [
                sys.executable, "-B", str(verifier), "--report", str(report),
                "--mode", "normal", "--shape", "large",
            ]
            commands["oracle"] = oracle_argv
            oracle_rc = run_oracle(oracle_argv, REPO, env, oracle_log)
            state["oracle_exit_code"] = oracle_rc
            if oracle_rc != 0:
                failures.append(f"oracle exited {oracle_rc}")
        else:
            failures.append("workload did not produce report.json")
    except Exception as error:
        failures.append(repr(error))
    finally:
        try:
            source_after = custody.sources()
        except Exception as error:  # preserve the profile failure if custody itself cannot be read
            source_after = {"error": repr(error)}
            failures.append("could not read source custody after profile")
        state["source_after"] = source_after
        state["source_unchanged"] = source_after == source_before
        if not state["source_unchanged"]:
            failures.append("source custody changed during profile")
        state["commands"] = commands
        state["postprocess"] = postprocess
        state["counters"] = {
            "status": "unavailable",
            "reason": "sampled cycles:u profile only; perf stat counters were not collected",
        }
        state["rss"] = {
            "status": "retained",
            "source": "GNU /usr/bin/time -v Maximum resident set size (kbytes) in resource.log",
            "scope": "perf-recorded process tree, including profiler overhead",
        }
        state["allocation"] = {
            "status": "unavailable",
            "reason": "ordinary allocator counters were not collected and allocator variables were not tuned",
        }
        state["artifacts"] = {
            name: artifact(path)
            for name, path in {
                "perf_data": perf_data,
                "perf_record_stdout": record_stdout,
                "perf_record_stderr": record_stderr,
                "resource": resource,
                "report": report,
                "catalog": catalog,
                "top_symbols": top_symbols,
                "top_symbols_stderr": top_stderr,
                "perf_script": perf_script,
                "perf_script_stderr": perf_script_stderr,
                "stacks_folded": folded,
                "oracle": oracle_log,
                "perf_help_stdout": help_stdout,
                "perf_help_stderr": help_stderr,
            }.items()
            if path.is_file()
        }
        state["failures"] = failures
        state["finished_utc"] = now()
        state["status"] = "pass" if not failures else "failed"
        write_json(receipt_path, state)

    print(json.dumps({"status": state["status"], "role": args.role, "profile": str(role_dir.relative_to(BUNDLE))}))
    return 0 if state["status"] == "pass" else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ProfileError as error:
        print(f"profile failed: {error}", file=sys.stderr)
        raise SystemExit(1)
