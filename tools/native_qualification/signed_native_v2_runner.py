#!/usr/bin/env python3
"""Execute the prepared signed-native v2 probe and emit bound receipts.

The generated Rust probe is the producer of route and batch semantics.  This
module owns only the process boundary: it invokes the exact built binary,
captures bounded stdout/stderr and process ownership, and writes canonical
child/capabilities receipts for the consumer to validate.  It never invents a
route result from a status field.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import selectors
import signal
import subprocess
import sys
import time
from typing import Any, Mapping, Sequence

TOOLS_ROOT = Path(__file__).resolve().parents[1]
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

from native_qualification import signed_native_v2_consumer as consumer  # noqa: E402


RUNNER_SCHEMA = "litchi.signed-native.runner.v2"
PROBE_SCHEMA = consumer.PROBE_SCHEMA
CHILD_SCHEMA = consumer.CHILD_SCHEMA
BATCH_PROOF_SCHEMA = consumer.BATCH_PROOF_SCHEMA


class RunnerError(ValueError):
    """The native process did not produce a trustworthy v2 receipt."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise RunnerError(message)


def _canonical(value: Any) -> bytes:
    return consumer._canonical(value)


def _load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"{label} is not valid JSON: {error}") from error
    _expect(isinstance(value, dict), f"{label} is not an object")
    return value


def _write_new(path: Path, payload: bytes, label: str) -> dict[str, Any]:
    consumer._reject_symlink_components(path, label)
    _expect(not path.exists() and not path.is_symlink(), f"refusing to overwrite {label}: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(payload)
    return consumer._artifact(path, label)


def _write_json(path: Path, value: Mapping[str, Any], label: str) -> dict[str, Any]:
    return _write_new(path, _canonical(value), label)


def _process_group_pids(process_group: int) -> list[int]:
    """Return live members of a process group without following children."""

    members: list[int] = []
    proc_root = Path("/proc")
    try:
        entries = list(proc_root.iterdir())
    except OSError as error:
        raise RunnerError(f"cannot inspect process ownership: {error}") from error
    for entry in entries:
        if not entry.name.isdigit():
            continue
        try:
            stat = (entry / "stat").read_text(encoding="ascii")
            # ``comm`` may contain spaces, so split only after its final
            # closing parenthesis.  The remaining fields begin with state,
            # ppid, pgrp (procfs fields 3, 4, 5).
            _comm, fields = stat.rsplit(") ", 1)
            observed = int(fields.split()[2])
        except (OSError, UnicodeError, ValueError, IndexError):
            continue
        if observed == process_group:
            members.append(int(entry.name))
    return sorted(members)


def _terminate_group(process: subprocess.Popen[bytes]) -> None:
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except (ProcessLookupError, PermissionError):
        pass
    try:
        process.wait(timeout=1.0)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except (ProcessLookupError, PermissionError):
            pass
        try:
            process.wait(timeout=2.0)
        except subprocess.TimeoutExpired:
            pass


def _run_process(
    command: Sequence[str],
    *,
    cwd: Path,
    environment: Mapping[str, str],
    stdout_path: Path,
    stderr_path: Path,
) -> dict[str, Any]:
    """Run one owned process with bounded stream capture and reaping."""

    consumer._reject_symlink_components(stdout_path, "native stdout")
    consumer._reject_symlink_components(stderr_path, "native stderr")
    _expect(not stdout_path.exists() and not stderr_path.exists(), "native output paths must be fresh")
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    stderr_path.parent.mkdir(parents=True, exist_ok=True)
    started = time.monotonic()
    deadline = started + consumer.CHILD_TIMEOUT_SECONDS
    started_ns = time.time_ns()
    process: subprocess.Popen[bytes] | None = None
    timed_out = False
    guard_error: str | None = None
    returncode: int | None = None
    captured = {"stdout": 0, "stderr": 0}
    drain_deadline: float | None = None
    selector = selectors.DefaultSelector()
    streams: dict[str, Any] = {}
    try:
        process = subprocess.Popen(
            list(command),
            cwd=str(cwd),
            env=dict(environment),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            start_new_session=True,
        )
        _expect(process.stdout is not None and process.stderr is not None, "native pipes were not created")
        output_files = {
            "stdout": stdout_path.open("wb"),
            "stderr": stderr_path.open("wb"),
        }
        streams = output_files
        selector.register(process.stdout, selectors.EVENT_READ, "stdout")
        selector.register(process.stderr, selectors.EVENT_READ, "stderr")
        while selector.get_map():
            now = time.monotonic()
            if process.poll() is None and now >= deadline:
                timed_out = True
                guard_error = "child timeout exceeded"
                _terminate_group(process)
                drain_deadline = now + 1.0
            elif process.poll() is not None:
                # A descendant can inherit a pipe after the leader exits.  Do
                # not wait forever for that descriptor: the process-group
                # cleanup below is the ownership proof, and this grace period
                # bounds receipt production even for a malformed child.
                if drain_deadline is None:
                    drain_deadline = now + 1.0
                elif now >= drain_deadline:
                    guard_error = guard_error or "child output pipes remained open after exit"
                    for key in list(selector.get_map().values()):
                        pipe = key.fileobj
                        selector.unregister(pipe)
                        pipe.close()
                    break
            select_timeout = 0.05
            if process.poll() is None:
                select_timeout = min(select_timeout, max(0.0, deadline - now))
            elif drain_deadline is not None:
                select_timeout = min(select_timeout, max(0.0, drain_deadline - now))
            events = selector.select(select_timeout)
            for key, _ in events:
                stream_name = str(key.data)
                pipe = key.fileobj
                chunk = os.read(pipe.fileno(), 64 * 1024)
                if not chunk:
                    selector.unregister(pipe)
                    pipe.close()
                    continue
                remaining = consumer.CHILD_OUTPUT_LIMIT - captured[stream_name]
                if remaining <= 0:
                    guard_error = guard_error or f"{stream_name} exceeded output limit"
                    _terminate_group(process)
                    continue
                accepted = chunk[:remaining]
                streams[stream_name].write(accepted)
                streams[stream_name].flush()
                captured[stream_name] += len(accepted)
                if len(accepted) != len(chunk):
                    guard_error = guard_error or f"{stream_name} exceeded output limit"
                    _terminate_group(process)
            if process.poll() is not None and not selector.get_map():
                break
        returncode = process.wait(timeout=2.0)
    except (OSError, subprocess.TimeoutExpired) as error:
        guard_error = guard_error or str(error)
        if process is not None and process.poll() is None:
            _terminate_group(process)
        if process is not None:
            returncode = process.returncode
    finally:
        selector.close()
        for stream in streams.values():
            stream.close()
        if process is not None:
            for pipe in (process.stdout, process.stderr):
                if pipe is not None and not pipe.closed:
                    pipe.close()
    _expect(process is not None, "native process was not started")
    if returncode is None:
        returncode = process.returncode
    process_group = process.pid
    survivors: list[int] = []
    for _ in range(20):
        survivors = [pid for pid in _process_group_pids(process_group) if pid != os.getpid()]
        if not survivors:
            break
        _terminate_group(process)
        time.sleep(0.05)
    after_guard_error = None if not survivors else "owned process group survived exit"
    ended_ns = time.time_ns()
    return {
        "pid": process.pid,
        "process_group": process_group,
        "start_new_session": True,
        "started_unix_ns": started_ns,
        "ended_unix_ns": ended_ns,
        "elapsed_ns": max(0, ended_ns - started_ns),
        "exit_code": returncode,
        "timed_out": timed_out,
        "guard_error": guard_error,
        "after_guard_error": after_guard_error,
        "process_group_empty": not survivors,
        "stdout": consumer._artifact(stdout_path, "native stdout"),
        "stderr": consumer._artifact(stderr_path, "native stderr"),
    }


def _output_json(process: Mapping[str, Any], label: str) -> dict[str, Any]:
    stdout = process["stdout"]
    _expect(isinstance(stdout, Mapping), f"{label} stdout binding is missing")
    _expect(stdout.get("bytes", 0) <= consumer.CHILD_OUTPUT_LIMIT, f"{label} stdout exceeds the output limit")
    raw = Path(str(stdout["path"])).read_bytes()
    try:
        value = json.loads(raw.decode("utf-8"))
    except (UnicodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"{label} stdout is not JSON: {error}") from error
    _expect(isinstance(value, dict), f"{label} stdout is not an object")
    return value


def _environment(state: Mapping[str, Any], endpoint: str) -> dict[str, str]:
    fixture = state["fixture"]
    manifest = fixture["manifest"]
    return {
        "LITCHI_SIGNED_FIXTURE_ROOT": str(fixture["root"]),
        "LITCHI_SIGNED_FIXTURE_MANIFEST": str(manifest["path"]),
        "LITCHI_SIGNED_FIXTURE_MANIFEST_SHA256": str(manifest["sha256"]),
        "LITCHI_SIGNED_ENDPOINT": endpoint,
    }


def _child(
    *,
    route: str,
    state: Mapping[str, Any],
    build: Mapping[str, Any],
    build_artifact: Mapping[str, Any],
    runner_artifact: Mapping[str, Any],
    command: Sequence[str],
    environment: Mapping[str, str],
    process: Mapping[str, Any],
    result: Mapping[str, Any],
    output: Path,
) -> dict[str, Any]:
    module = consumer._policy()
    endpoint = str(state["endpoint"])
    expected = consumer._expected_route_status(endpoint, route)
    _expect(result.get("schema") == PROBE_SCHEMA, f"{route} probe schema is not v2")
    _expect(result.get("route") == route, f"{route} probe route differs")
    _expect(result.get("capability") == expected, f"{route} probe capability differs")
    _expect(isinstance(result.get("fixtures"), list), f"{route} probe fixture results are missing")
    status = expected
    value = {
        "schema": CHILD_SCHEMA,
        "probe_schema": PROBE_SCHEMA,
        "mode": "validate",
        "route": route,
        "status": status,
        "endpoint": endpoint,
        "api_branch": build["api_branch"],
        "preparation_receipt": state["receipt_artifact"],
        "build_receipt": build_artifact,
        "source_identity": consumer._concrete_source_identity(state),
        "source_archive": state["receipt"]["source"]["archive"],
        "lock": state["receipt"]["source"]["locks"],
        "lock_graph": state["receipt"]["source"].get("lock_graph"),
        "source_manifests": state["receipt"]["source"].get("manifests"),
        "adapter": state["adapter"],
        "binary": build["binary"],
        "fixture": state["receipt"]["fixture"],
        "runner": runner_artifact,
        "argv": list(command),
        "command": list(command),
        "environment": dict(environment),
        "cwd": str(state["prepared"] / "consumer"),
        "process": {
            key: process[key]
            for key in (
                "pid",
                "process_group",
                "start_new_session",
                "started_unix_ns",
                "ended_unix_ns",
                "elapsed_ns",
                "process_group_empty",
            )
        },
        "stdout": process["stdout"],
        "stderr": process["stderr"],
        "timeout_seconds": consumer.CHILD_TIMEOUT_SECONDS,
        "exit_code": process["exit_code"],
        "timed_out": process["timed_out"],
        "guard_error": process["guard_error"],
        "after_guard_error": process["after_guard_error"],
        "result": dict(result),
    }
    path = output / "children" / f"{route.replace('/', '__')}.json"
    artifact = _write_json(path, value, f"child receipt {route}")
    _expect(process["exit_code"] == 0 and not process["timed_out"], f"{route} probe failed")
    _expect(process["guard_error"] is None and process["after_guard_error"] is None, f"{route} probe guard failed")
    return {"value": value, "artifact": artifact}


def _batch_proof(
    *,
    state: Mapping[str, Any],
    build: Mapping[str, Any],
    build_artifact: Mapping[str, Any],
    runner_artifact: Mapping[str, Any],
    command: Sequence[str],
    environment: Mapping[str, str],
    process: Mapping[str, Any],
    result: Mapping[str, Any],
    output: Path,
) -> dict[str, Any]:
    _expect(result.get("schema") == BATCH_PROOF_SCHEMA, "signed-batch producer schema differs")
    _expect(result.get("selection") == "immediate_refusal", "signed-batch producer selection differs")
    _expect(result.get("status") == "expected_refusal", "signed-batch producer did not refuse")
    error = result.get("error")
    _expect(isinstance(error, Mapping), "signed-batch producer error is missing")
    _expect(error.get("kind") == "DocumentProtected", "signed-batch producer kind differs")
    _expect(error.get("code") == "document_protected" and error.get("item") is None, "signed-batch producer refusal differs")
    _expect(result.get("item") is None, "signed-batch producer returned an item")
    _expect(result.get("iterator_reads") == 0, "signed-batch producer consumed its iterator")
    _expect(result.get("publication_attempts") == 0, "signed-batch producer published output")
    _expect(result.get("source_bytes_equal") is True, "signed-batch producer changed source bytes")
    try:
        before_bytes = bytes.fromhex(str(result.pop("source_before_hex")))
        after_bytes = bytes.fromhex(str(result.pop("source_after_hex")))
    except ValueError as error:
        raise RunnerError(f"signed-batch producer source bytes are not hex: {error}") from error
    _expect(before_bytes == after_bytes, "signed-batch producer source bytes differ")
    before_path = output / "batch" / "source-before.bin"
    after_path = output / "batch" / "source-after.bin"
    before_artifact = _write_new(before_path, before_bytes, "signed-batch source before")
    after_artifact = _write_new(after_path, after_bytes, "signed-batch source after")
    value = {
        **dict(result),
        "schema": BATCH_PROOF_SCHEMA,
        "source_before": before_artifact,
        "source_after": after_artifact,
        "endpoint": state["endpoint"],
        "source_identity": consumer._concrete_source_identity(state),
        "source_archive": state["receipt"]["source"]["archive"],
        "lock": state["receipt"]["source"]["locks"],
        "lock_graph": state["receipt"]["source"].get("lock_graph"),
        "source_manifests": state["receipt"]["source"].get("manifests"),
        "binary": build["binary"],
        "fixture": state["receipt"]["fixture"],
        "runner": runner_artifact,
        "argv": list(command),
        "command": list(command),
        "environment": dict(environment),
        "cwd": str(state["prepared"] / "consumer"),
        "process": {
            key: process[key]
            for key in (
                "pid",
                "process_group",
                "start_new_session",
                "started_unix_ns",
                "ended_unix_ns",
                "elapsed_ns",
                "process_group_empty",
            )
        },
        "stdout": process["stdout"],
        "stderr": process["stderr"],
        "timeout_seconds": consumer.CHILD_TIMEOUT_SECONDS,
        "exit_code": process["exit_code"],
        "timed_out": process["timed_out"],
        "guard_error": process["guard_error"],
        "after_guard_error": process["after_guard_error"],
        "preparation_receipt": state["receipt_artifact"],
        "build_receipt": build_artifact,
    }
    _expect(process["exit_code"] == 0 and not process["timed_out"], "signed-batch producer failed")
    _expect(process["guard_error"] is None and process["after_guard_error"] is None, "signed-batch producer guard failed")
    return value


def run(
    state: Mapping[str, Any],
    *,
    output: Path | None = None,
    endpoint: str | None = None,
) -> Path:
    """Run every route and the signed-batch proof, returning capabilities."""

    state_endpoint = str(state.get("endpoint"))
    _expect(state_endpoint in {"before", "after"}, "prepared state endpoint is missing")
    if endpoint is None:
        endpoint = state_endpoint
    _expect(endpoint == state_endpoint, "native endpoint does not match the prepared source binary")
    prepared = Path(state["prepared"])
    build_path = prepared / "build-receipt.json"
    build, build_artifact = consumer._load_canonical(build_path, "build receipt")
    _expect(build.get("schema") == consumer.BUILD_SCHEMA and build.get("status") == "built", "native runner needs a built v2 receipt")
    _expect(build.get("endpoint") == state_endpoint, "native build receipt belongs to a different endpoint")
    _expect(build.get("api_branch") == consumer._policy().ENDPOINT_API_BRANCH[state_endpoint], "native build API branch belongs to a different endpoint")
    binary = consumer._verify_artifact(build.get("binary"), "signed-native consumer binary", executable=True)
    _expect(binary == build["binary"], "native runner binary identity changed")
    runner_artifact = consumer._artifact(Path(__file__).resolve(), "signed-native v2 runner")
    environment = _environment(state, endpoint)
    project = prepared / "consumer"
    destination = consumer._safe_external(output or (prepared / "protocol"), "native protocol output")
    _expect(not destination.exists() and not destination.is_symlink(), f"native protocol output is not fresh: {destination}")
    destination.mkdir(parents=True)
    routes: dict[str, str] = {}
    evidence: dict[str, dict[str, Any]] = {}
    for route in consumer._policy().ROUTES:
        command = [binary["path"], "--mode", "validate", "--route", route, "--blocks", "1"]
        process = _run_process(
            command,
            cwd=project,
            environment={**os.environ, **environment},
            stdout_path=destination / "children" / f"{route.replace('/', '__')}.stdout",
            stderr_path=destination / "children" / f"{route.replace('/', '__')}.stderr",
        )
        result = _output_json(process, f"{route} probe")
        routes[route] = consumer._expected_route_status(endpoint, route)
        child = _child(
            route=route,
            state=state,
            build=build,
            build_artifact=build_artifact,
            runner_artifact=runner_artifact,
            command=command,
            environment=environment,
            process=process,
            result=result,
            output=destination,
        )
        evidence[route] = child["artifact"]
    batch_command = [binary["path"], "--mode", "batch_refusal", "--blocks", "1"]
    batch_process = _run_process(
        batch_command,
        cwd=project,
        environment={**os.environ, **environment},
        stdout_path=destination / "batch" / "batch.stdout",
        stderr_path=destination / "batch" / "batch.stderr",
    )
    batch_result = _output_json(batch_process, "signed-batch probe")
    batch = _batch_proof(
        state=state,
        build=build,
        build_artifact=build_artifact,
        runner_artifact=runner_artifact,
        command=batch_command,
        environment=environment,
        process=batch_process,
        result=batch_result,
        output=destination,
    )
    capabilities = {
        "schema": consumer.CAPABILITIES_SCHEMA,
        "protocol_schema": RUNNER_SCHEMA,
        "preparation_receipt": state["receipt_artifact"],
        "build_receipt": build_artifact,
        "source_identity": consumer._concrete_source_identity(state),
        "endpoint": endpoint,
        "source_identity_receipt": state["receipt"]["source_identity"],
        "source": state["receipt"]["source"],
        "source_archive": state["receipt"]["source"]["archive"],
        "lock": state["receipt"]["source"]["locks"],
        "lock_graph": state["receipt"]["source"].get("lock_graph"),
        "source_manifests": state["receipt"]["source"].get("manifests"),
        "binary": binary,
        "api_branch": build["api_branch"],
        "build_receipt_binding": {
            "preparation_receipt": state["receipt_artifact"],
            "build_receipt": build_artifact,
        },
        "adapter": state["adapter"],
        "fixture": state["receipt"]["fixture"],
        "runner": runner_artifact,
        "routes": routes,
        "evidence": evidence,
        "signed_batch": batch,
    }
    return_path = destination / "capabilities.json"
    _write_json(return_path, capabilities, "v2 capabilities receipt")
    return return_path


def run_prepared(prepared_root: Path, *, output: Path | None = None, endpoint: str | None = None) -> Path:
    state = consumer._revalidate_preparation(prepared_root)
    return run(state, output=output, endpoint=endpoint)


def main(argv: Sequence[str]) -> int:
    import argparse

    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--prepared", type=Path, required=True)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--endpoint", choices=("before", "after"))
    args = parser.parse_args(list(argv))
    receipt = run_prepared(args.prepared, output=args.output, endpoint=args.endpoint)
    print(json.dumps({"status": "completed", "capabilities": str(receipt)}, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
