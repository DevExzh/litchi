#!/usr/bin/env python3
"""Guarded, receipt-producing launcher for the W8 File ABBA window.

This is deliberately a wrapper around the frozen 0362h runner.  ``prepare``
binds the reviewed V1lU2E artifacts and creates a no-argument launcher.  The
``run`` path adds the host/cgroup/deadline/output lifecycle and invokes the
existing analyzer only after every guard passes.  It never builds a binary or
changes the native probe protocol.

The generated launcher sets RLIMIT_FSIZE with Python's byte-valued resource
API.  Bash ``ulimit -f`` is intentionally not used: its units are shell/
platform dependent and are easy to misread as bytes.
"""
from __future__ import annotations

import argparse
import ast
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import resource
import signal
import shutil
import subprocess
import sys
import time
from typing import Any, Iterable


SCHEMA = "litchi_w8_quiet_abba.v1"
ROOT = Path(__file__).resolve().parents[4]
ROUNDS = 5
TARGET_ROUTE_COUNT = 109
DEFAULT_DEADLINE_SECONDS = 12 * 60 * 60
STOP_GRACE_SECONDS = 60
MEMORY_MAX_BYTES = 20_000_000_000
SWAP_MAX_BYTES = 0
OUTPUT_FILE_LIMIT_BYTES = 64 * 1024 * 1024
AGGREGATE_OUTPUT_LIMIT_BYTES = 8 * 1024 * 1024 * 1024
POLL_SECONDS = 0.25
UNIT_PROPERTIES = (
    "ActiveState", "SubState", "Result", "ExecMainStatus", "MainPID",
    "ControlGroup", "MemoryCurrent", "MemoryPeak",
)
FULL_COMMIT = re.compile(r"^[0-9a-f]{40}$")
RUN_ID = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]{0,63}$")

V1_EXPECTED = {
    "before_binary": (67_879_576, "be7e3dbd9c31403d3bc017c07299f391e2b4b5b8f889107f9e0bfd686ada75c6"),
    "after_binary": (70_239_360, "7c9061725c74c2fc255d33694cbbe1f6de0dd5f750cc8d1017dc82811bef58f6"),
    "before_manifest": (3_193, "72080e5313b5ba183a45769d2e78a9116d8b3ea7527f7af8b9aae059f42e2633"),
    "after_manifest": (3_152, "500c156c1fe1b9962a94ff0f5f8c03752ca1fdd08b3725438f2811c52466ab63"),
    "pair": (8_949, "f59c22c65737495b7adc88a5c0c6286680f92ed4e101009f66ca81e4c3d72275"),
}
TOOL_EXPECTED = {
    "policy": (7_609, "e56d549c733629ac73492a3511a45a1ed6a4cec97f1a4705ceea4ba9b67d97d2"),
    "runner": (39_866, "d8e708a0b0d01461e3c0d0597b312c7c514ea9369119c99ad0d67847f3296283"),
    "analyzer": (30_190, "5e0f6ae124f70a1bf07dd0f9fb6b1450445182f77d15bc0a66d716e0bd77e4d6"),
    "pair_tool": (38_986, "121761a850a3c280cece5de4424ef3557ce1220994421c83c82e030ecbcca43a"),
}


class GuardFailure(RuntimeError):
    """A fail-closed lifecycle or provenance failure."""


def canonical(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest(path: Path) -> dict[str, object]:
    path = path.resolve()
    if not path.is_file() or path.is_symlink():
        raise GuardFailure(f"artifact is not a regular file: {path}")
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(block)
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": hasher.hexdigest()}


def verify_digest(path: Path, expected: dict[str, object] | tuple[int, str], label: str) -> dict[str, object]:
    actual = digest(path)
    if isinstance(expected, tuple):
        want = {"bytes": expected[0], "sha256": expected[1]}
    else:
        want = expected
    if (actual["bytes"], actual["sha256"]) != (want.get("bytes"), want.get("sha256")):
        raise GuardFailure(f"{label} identity changed: {actual} != {want}")
    return actual


def load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise GuardFailure(f"{label} is not valid JSON: {error}") from error
    if not isinstance(value, dict):
        raise GuardFailure(f"{label} is not a JSON object")
    return value


def write_json(path: Path, value: object) -> None:
    path.write_bytes(canonical(value))


def sha256_text(path: Path) -> str:
    return str(digest(path)["sha256"])


def root_paths(root: Path, v1: Path, output: Path, tmp: Path) -> dict[str, Path]:
    return {
        "wrapper": Path(__file__).resolve(),
        "before_binary": v1 / "target-before/x86_64-unknown-linux-gnu/release/examples/file_w8_original_baseline_probe",
        "after_binary": v1 / "target-after/x86_64-unknown-linux-gnu/release/examples/file_w8_qualified_probe",
        "before_manifest": v1 / "endpoint-before.json",
        "after_manifest": v1 / "endpoint-after.json",
        "accepted_pair": v1 / "original-to-final-pair.json",
        "policy": root / "docs/unified-ops/results/0362yh/original-to-final-policy.json",
        "runner": root / "docs/unified-ops/results/0362h/run.py",
        "analyzer": root / "docs/unified-ops/results/0362h/analyze.py",
        "pair_tool": root / "tools/w8_projection_policy.py",
        "output": output,
        "tmp": tmp,
    }


def endpoint_inputs(manifest: Path, label: str) -> dict[str, dict[str, object]]:
    value = load_json(manifest, label)
    result: dict[str, dict[str, object]] = {}
    for name in (
        "binary", "build_receipt", "dependency_projection", "fixture", "lock",
        "package_manifest", "probe", "source", "source_assembly", "toolchain",
        "validation", "workspace_manifest",
    ):
        row = value.get(name)
        if not isinstance(row, dict) or not isinstance(row.get("path"), str):
            raise GuardFailure(f"{label}.{name} is not a bound artifact")
        if not isinstance(row.get("bytes"), int) or not isinstance(row.get("sha256"), str):
            raise GuardFailure(f"{label}.{name} has no complete identity")
        result[f"{label}_{name}"] = row
    source_commit = value.get("source_commit")
    if not isinstance(source_commit, str) or not FULL_COMMIT.fullmatch(source_commit):
        raise GuardFailure(f"{label} source commit is not a full commit")
    if value.get("target") != "x86_64-unknown-linux-gnu" or value.get("profile") != "release":
        raise GuardFailure(f"{label} target/profile is not the reviewed pair")
    return result


def tool_path(name: str, candidate: str | None = None) -> Path:
    value = candidate or shutil.which(name)
    if not value:
        raise GuardFailure(f"required tool is unavailable: {name}")
    path = Path(value).resolve()
    if not path.is_file() or not os.access(path, os.X_OK):
        raise GuardFailure(f"tool is not executable: {path}")
    return path


def expected_tool(path: Path, expected: tuple[int, str] | None, label: str) -> dict[str, object]:
    actual = digest(path)
    if expected is not None and (actual["bytes"], actual["sha256"]) != expected:
        raise GuardFailure(f"{label} tool identity differs: {actual} != {expected}")
    return actual


def forbidden_rows() -> tuple[list[dict[str, object]], list[dict[str, object]]]:
    spec = importlib.util.spec_from_file_location("w8_frozen_runner_for_guard", ROOT / "docs/unified-ops/results/0362h/run.py")
    if spec is None or spec.loader is None:
        raise GuardFailure("cannot load frozen runner guard")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    rows = module.process_rows()
    allowed = module.descendants(rows, {os.getpid()})
    blockers = [row for row in rows if int(row["pid"]) not in allowed and module.forbidden(row)]
    return rows, blockers


def quiet_snapshot(path: Path, label: str) -> dict[str, object]:
    rows, blockers = forbidden_rows()
    value = {"label": label, "time_ns": time.time_ns(), "blockers": blockers, "processes": rows}
    write_json(path, value)
    return value


def directory_bytes(paths: Iterable[Path]) -> int:
    total = 0
    seen: set[tuple[int, int]] = set()
    for root in paths:
        if not root.exists():
            continue
        if root.is_file():
            total += root.stat().st_size
            continue
        for path in root.rglob("*"):
            try:
                stat = path.stat()
            except OSError:
                continue
            if not path.is_file() or path.is_symlink():
                continue
            identity = (stat.st_dev, stat.st_ino)
            if identity not in seen:
                seen.add(identity)
                total += stat.st_size
    return total


def proc_identity(pid: int) -> dict[str, object] | None:
    try:
        stat = Path(f"/proc/{pid}/stat").read_text()
        close = stat.rfind(")")
        fields = stat[close + 2 :].split()
        cmdline = Path(f"/proc/{pid}/cmdline").read_bytes().replace(b"\0", b" ").decode(errors="replace").strip()
        return {
            "pid": pid,
            "starttime": fields[19],
            "ppid": int(fields[1]),
            "pgrp": int(fields[2]),
            "session": int(fields[3]),
            "cmdline": cmdline,
        }
    except (OSError, IndexError, ValueError):
        return None


def proc_state(pid: int) -> str | None:
    """Return the Linux process state without changing identity evidence."""
    try:
        stat = Path(f"/proc/{pid}/stat").read_text()
        close = stat.rfind(")")
        fields = stat[close + 2 :].split()
        return fields[0]
    except (OSError, IndexError, ValueError):
        return None


def reap_zombie(identity: dict[str, object]) -> bool:
    """Reap an owned orphan adopted by this process when it is a zombie."""
    pid = int(identity["pid"])
    current = proc_identity(pid)
    if current is None or current["starttime"] != identity["starttime"] or proc_state(pid) != "Z":
        return False
    try:
        waited, _status = os.waitpid(pid, os.WNOHANG)
    except (ChildProcessError, OSError):
        return False
    return waited == pid


def reap_process_group_zombies(pgid: int, protected_pids: set[int] | frozenset[int] = frozenset()) -> None:
    """Reap adopted zombies in one owned process group, if possible."""
    try:
        candidates = list(Path("/proc").iterdir())
    except OSError:
        return
    for candidate in candidates:
        if not candidate.name.isdigit():
            continue
        identity = proc_identity(int(candidate.name))
        if (identity is not None and identity["pid"] not in protected_pids
                and identity["pgrp"] == pgid):
            reap_zombie(identity)


def alive(identity: dict[str, object]) -> bool:
    current = proc_identity(int(identity["pid"]))
    return current is not None and current["starttime"] == identity["starttime"]


def capture_descendants(root_pid: int, seen: dict[str, dict[str, object]]) -> None:
    """Record only processes whose live parent chain starts at root_pid."""
    root = proc_identity(root_pid)
    if root is None:
        return
    rows: dict[int, dict[str, object]] = {}
    try:
        candidates = list(Path("/proc").iterdir())
    except OSError:
        return
    for candidate in candidates:
        if not candidate.name.isdigit():
            continue
        identity = proc_identity(int(candidate.name))
        if identity is not None:
            rows[int(identity["pid"])] = identity
    owned = {root_pid}
    changed = True
    while changed:
        changed = False
        for pid, identity in rows.items():
            if pid not in owned and identity.get("ppid") in owned:
                seen[f"{identity['pid']}:{identity['starttime']}"] = identity
                owned.add(pid)
                changed = True


def parse_properties(text: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for line in text.splitlines():
        if "=" in line:
            key, value = line.split("=", 1)
            result[key] = value
    return result


def show_unit(systemctl: Path, unit: str) -> dict[str, object]:
    command = [str(systemctl), "--user", "show", "--no-pager", f"--property={','.join(UNIT_PROPERTIES)}", unit]
    try:
        completed = subprocess.run(command, capture_output=True, text=True, timeout=5, check=False)
    except (OSError, subprocess.TimeoutExpired) as error:
        return {"available": False, "monitor_error_kind": "systemctl", "monitor_error": str(error), "command": command}
    if completed.returncode != 0:
        stderr = completed.stderr[-4096:]
        # A scope is allowed to disappear only after a valid final sample.  A
        # generic non-zero systemctl result remains a monitor failure; this
        # narrow classification is what lets the caller distinguish unit GC
        # from an intermittent command failure.
        unit_gone = completed.returncode == 1 and re.search(
            r"(?:not found|not loaded|no such unit|could not be found)", stderr, re.IGNORECASE
        )
        return {
            "available": False,
            "monitor_error_kind": "unit-gone" if unit_gone else "systemctl",
            "monitor_error": "systemctl failed",
            "returncode": completed.returncode,
            "stderr": stderr,
            "command": command,
        }
    properties = parse_properties(completed.stdout)
    real_gc_shape = (
        properties.get("ActiveState") == "inactive"
        and properties.get("SubState") == "dead"
        and properties.get("Result") == "success"
        and properties.get("ControlGroup") == ""
        and properties.get("MemoryCurrent") == "[not set]"
        and properties.get("MemoryPeak") == "[not set]"
        and "MainPID" not in properties
        and "ExecMainStatus" not in properties
    )
    if real_gc_shape:
        properties["available"] = False
        properties["monitor_error_kind"] = "unit-gone"
        properties["monitor_error"] = "transient scope is inactive with an empty ControlGroup"
        properties["command"] = command
        return properties
    missing = [name for name in UNIT_PROPERTIES if name not in properties]
    if missing:
        return {
            "available": False,
            "monitor_error_kind": "schema",
            "monitor_error": f"systemctl omitted properties: {','.join(missing)}",
            "missing_properties": missing,
            "command": command,
        }
    if not properties.get("ControlGroup"):
        return {
            "available": False,
            "monitor_error_kind": "schema",
            "monitor_error": "systemctl returned an empty ControlGroup",
            "command": command,
        }
    properties["available"] = True
    properties["command"] = command
    return properties


def cgroup_pids(properties: dict[str, object]) -> tuple[str | None, list[int], str | None]:
    raw = properties.get("ControlGroup")
    if not isinstance(raw, str) or not raw.startswith("/") or ".." in raw.split("/"):
        return None, [], "ControlGroup is missing or unsafe"
    path = Path("/sys/fs/cgroup") / raw.lstrip("/") / "cgroup.procs"
    try:
        pids = [int(line) for line in path.read_text().splitlines() if line.strip()]
    except FileNotFoundError as error:
        return raw, [], f"cgroup.procs is gone: {error}"
    except (OSError, ValueError) as error:
        return raw, [], f"cannot read cgroup.procs: {error}"
    return raw, pids, None


def kill_unit(systemctl: Path, unit: str) -> subprocess.CompletedProcess[str] | None:
    try:
        return subprocess.run(
            [str(systemctl), "--user", "kill", "--kill-who=all", unit],
            capture_output=True,
            text=True,
            timeout=5,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None


def terminate_process_group(process: subprocess.Popen[Any] | None, grace_seconds: float = 10.0) -> str | None:
    if process is None:
        return None
    def group_exists() -> bool:
        # Keep the top-level Popen child for its owner to wait: an external
        # waitpid would discard its real return code.  Only adopted
        # descendants in this owned process group are eligible for reaping.
        reap_process_group_zombies(process.pid, {process.pid})
        try:
            os.killpg(process.pid, 0)
            return True
        except (ProcessLookupError, OSError):
            return False

    if process.poll() is not None and not group_exists():
        try:
            process.wait(timeout=0)
        except (OSError, subprocess.TimeoutExpired):
            pass
        return None
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    except OSError as error:
        return f"SIGTERM process-group cleanup failed: {error}"
    try:
        process.wait(timeout=grace_seconds)
    except subprocess.TimeoutExpired:
        pass
    if not group_exists():
        return None
    try:
        os.killpg(process.pid, signal.SIGKILL)
    except ProcessLookupError:
        return None
    except OSError as error:
        return f"SIGKILL process-group cleanup failed: {error}"
    try:
        process.wait(timeout=grace_seconds)
    except subprocess.TimeoutExpired:
        return "process group did not reap after SIGKILL"
    deadline = time.monotonic() + grace_seconds
    while group_exists() and time.monotonic() < deadline:
        time.sleep(0.05)
    return None if not group_exists() else "process group survived SIGKILL"


def terminate_owned_identity(identity: dict[str, object], sig: int) -> str | None:
    current = proc_identity(int(identity["pid"]))
    if current is None or current["starttime"] != identity["starttime"]:
        return None
    try:
        os.kill(int(identity["pid"]), sig)
    except ProcessLookupError:
        return None
    except OSError as error:
        return f"signal {sig} to owned pid {identity['pid']} failed: {error}"
    return None


def cleanup_owned(
    systemctl: Path | None,
    unit: str | None,
    systemd_process: subprocess.Popen[Any] | None,
    analyzer_process: subprocess.Popen[Any] | None,
    state: dict[str, Any],
) -> dict[str, object]:
    """Stop and reap the scope, analyzer, and observed escaped identities."""
    errors: list[str] = []
    seen: dict[str, dict[str, object]] = {}
    for key in ("seen_scope_processes", "seen_analyzer_processes"):
        values = state.get(key, {})
        if isinstance(values, dict):
            seen.update(values)
    cgroup_path = state.get("last_cgroup_path")
    unit_gc_expected = bool(state.get("unit_gc_expected"))
    reported_cgroup_errors: set[str] = set()

    def observed_alive() -> list[dict[str, object]]:
        popen_pids = {
            process.pid for process in (systemd_process, analyzer_process)
            if process is not None
        }
        for identity in seen.values():
            if int(identity["pid"]) not in popen_pids:
                reap_zombie(identity)
        return [identity for identity in seen.values() if alive(identity)] if isinstance(seen, dict) else []

    def current_cgroup_pids() -> list[int]:
        if not isinstance(cgroup_path, str):
            return []
        _, pids, error = cgroup_pids({"ControlGroup": cgroup_path})
        gone_after_gc = unit_gc_expected and error is not None and error.startswith("cgroup.procs is gone:")
        if error is not None and not gone_after_gc and error not in reported_cgroup_errors:
            reported_cgroup_errors.add(error)
            errors.append(error)
        return pids

    if analyzer_process is not None:
        error = terminate_process_group(analyzer_process)
        if error:
            errors.append(error)

    owned_alive = observed_alive()
    cgroup_alive = current_cgroup_pids()
    if systemctl is not None and unit is not None and (
        (systemd_process is not None and systemd_process.poll() is None)
        or owned_alive
        or cgroup_alive
        or (isinstance(cgroup_path, str) and not unit_gc_expected)
    ):
        result = kill_unit(systemctl, unit)
        if result is None:
            errors.append("systemd cgroup kill did not execute")
        elif result.returncode != 0:
            errors.append(f"systemd cgroup kill failed: {result.returncode}: {result.stderr[-1024:]}")
    error = terminate_process_group(systemd_process)
    if error:
        errors.append(error)
    for identity in owned_alive:
        error = terminate_owned_identity(identity, signal.SIGTERM)
        if error:
            errors.append(error)
    deadline = time.monotonic() + 10.0
    while time.monotonic() < deadline:
        remaining = observed_alive()
        current_pids = current_cgroup_pids()
        if not remaining and not current_pids:
            break
        time.sleep(0.05)

    remaining = observed_alive()
    current_pids = current_cgroup_pids()
    for identity in remaining:
        error = terminate_owned_identity(identity, signal.SIGKILL)
        if error:
            errors.append(error)

    # A TERM-resistant escaped child is killed above, but the receipt must not
    # claim cleanup until a post-KILL observation proves that it disappeared.
    kill_deadline = time.monotonic() + 10.0
    while time.monotonic() < kill_deadline:
        remaining = observed_alive()
        current_pids = current_cgroup_pids()
        if not remaining and not current_pids:
            break
        time.sleep(0.05)
    remaining = observed_alive()
    current_pids = current_cgroup_pids()
    if remaining:
        errors.append(f"owned descendants survived cleanup: {remaining}")
    if current_pids:
        errors.append(f"cgroup pids survived cleanup: {current_pids}")
    return {"errors": errors, "remaining_processes": remaining, "remaining_cgroup_pids": current_pids}


def systemd_command(config: dict[str, Any], paths: dict[str, Path]) -> list[str]:
    systemd_run = Path(config["artifacts"]["systemd_run"]["path"])
    timeout = Path(config["artifacts"]["timeout"]["path"])
    return [
        str(systemd_run), "--user", "--scope", f"--unit={config['unit']}",
        "--property=MemoryAccounting=yes", f"--property=MemoryMax={config['memory_max_bytes']}",
        f"--property=MemorySwapMax={config['swap_max_bytes']}", "--property=TasksMax=4096",
        "--property=KillMode=control-group", "--", str(timeout), "--foreground", "--signal=TERM",
        f"--kill-after={config['stop_grace_seconds']}s", f"{config['deadline_seconds']}s", str(paths["launcher"]),
    ]


def runner_arguments(paths: dict[str, Path], cpu: int, pair_receipt: Path) -> list[str]:
    return [
        str(paths["runner"]), str(paths["before_binary"]), str(paths["after_binary"]),
        str(paths["output"]), str(ROUNDS), str(cpu),
        "--before-manifest", str(paths["before_manifest"]),
        "--after-manifest", str(paths["after_manifest"]),
        "--original-to-final-policy", str(paths["policy"]),
        "--original-to-final-receipt", str(pair_receipt),
    ]


def launcher_binding(path: Path) -> dict[str, object]:
    """Read the literal CPU/argv embedded in the generated launcher."""
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    except (OSError, UnicodeError, SyntaxError) as error:
        raise GuardFailure(f"launcher is not parseable: {error}") from error
    values: dict[str, object] = {}
    for node in ast.walk(tree):
        if not isinstance(node, ast.Assign) or len(node.targets) != 1 or not isinstance(node.targets[0], ast.Name):
            continue
        name = node.targets[0].id
        if name in {"LAUNCHER_CPU", "LAUNCHER_ARGV"}:
            if name in values:
                raise GuardFailure(f"launcher contains duplicate {name} binding")
            try:
                values[name] = ast.literal_eval(node.value)
            except (ValueError, SyntaxError) as error:
                raise GuardFailure(f"launcher binding {name} is not literal: {error}") from error
    cpu = values.get("LAUNCHER_CPU")
    argv = values.get("LAUNCHER_ARGV")
    if not isinstance(cpu, int) or not isinstance(argv, list) or not all(isinstance(item, str) for item in argv):
        raise GuardFailure("launcher does not contain literal CPU and argv bindings")
    return {"cpu": cpu, "argv": argv}


def verify_launcher_binding(path: Path, cpu: int, argv: list[str]) -> None:
    binding = launcher_binding(path)
    if binding["cpu"] != cpu or binding["argv"] != argv:
        raise GuardFailure("launcher embedded CPU/argv does not match configuration")


def launcher_source(paths: dict[str, Path], cpu: int, pair_receipt: Path, tmp: Path) -> str:
    python = str(Path(sys.executable).resolve())
    argv = [python, "-B", *runner_arguments(paths, cpu, pair_receipt)]
    return (
        "#!" + python + "\n"
        f"LAUNCHER_CPU = {cpu!r}\n"
        f"LAUNCHER_ARGV = {argv!r}\n"
        "import os\n"
        "import resource\n"
        "import sys\n"
        "if len(sys.argv) != 1:\n"
        "    raise SystemExit('the reviewed W8 launcher accepts no arguments')\n"
        f"resource.setrlimit(resource.RLIMIT_FSIZE, ({OUTPUT_FILE_LIMIT_BYTES}, {OUTPUT_FILE_LIMIT_BYTES}))\n"
        "resource.setrlimit(resource.RLIMIT_CORE, (0, 0))\n"
        f"os.environ['TMPDIR'] = {str(tmp)!r}\n"
        "os.execve(LAUNCHER_ARGV[0], LAUNCHER_ARGV, os.environ.copy())\n"
    )


def make_paths(args: argparse.Namespace) -> dict[str, Path]:
    root = Path(args.root).resolve()
    v1 = Path(args.v1).resolve()
    output = Path(args.abba_output).resolve()
    tmp = Path(args.tmp_dir).resolve()
    return root_paths(root, v1, output, tmp)


def prepare(args: argparse.Namespace) -> int:
    paths = make_paths(args)
    config_path = Path(args.config).resolve()
    if config_path.exists():
        raise GuardFailure(f"configuration already exists: {config_path}")
    if paths["output"].exists() or paths["tmp"].exists():
        raise GuardFailure("ABBA output and TMP must both be fresh and absent")
    if not RUN_ID.fullmatch(args.run_id):
        raise GuardFailure(f"invalid run id: {args.run_id}")
    paths["output"].mkdir(parents=True, mode=0o700)
    paths["tmp"].mkdir(parents=True, mode=0o700)
    fixed: dict[str, dict[str, object]] = {}
    fixed["wrapper"] = digest(paths["wrapper"])
    for name in ("before_binary", "after_binary", "before_manifest", "after_manifest", "accepted_pair"):
        fixed[name] = verify_digest(paths[name], V1_EXPECTED[name.removeprefix("accepted_")], name)
    fixed["policy"] = verify_digest(paths["policy"], TOOL_EXPECTED["policy"], "policy")
    fixed["runner"] = verify_digest(paths["runner"], TOOL_EXPECTED["runner"], "runner")
    fixed["analyzer"] = verify_digest(paths["analyzer"], TOOL_EXPECTED["analyzer"], "analyzer")
    fixed["pair_tool"] = verify_digest(paths["pair_tool"], TOOL_EXPECTED["pair_tool"], "pair tool")
    for label, manifest in (("before", paths["before_manifest"]), ("after", paths["after_manifest"])):
        for key, row in endpoint_inputs(manifest, label).items():
            fixed[key] = verify_digest(Path(str(row["path"])), row, key)
    python = tool_path("python3", sys.executable)
    systemd_run = tool_path("systemd-run", "/usr/bin/systemd-run")
    systemctl = tool_path("systemctl", "/usr/bin/systemctl")
    timeout = tool_path("timeout", "/usr/bin/timeout")
    fixed["python"] = expected_tool(python, None, "python")
    fixed["systemd_run"] = expected_tool(systemd_run, None, "systemd-run")
    fixed["systemctl"] = expected_tool(systemctl, None, "systemctl")
    fixed["timeout"] = expected_tool(timeout, None, "timeout")

    pair_receipt = paths["tmp"] / "pair-revalidated.json"
    pair_command = [
        str(python), str(paths["pair_tool"]), "pair", "--policy", str(paths["policy"]),
        "--before-manifest", str(paths["before_manifest"]), "--after-manifest", str(paths["after_manifest"]),
        "--output", str(pair_receipt),
    ]
    pair_log = paths["tmp"] / "pair-revalidation.log"
    with pair_log.open("wb") as stream:
        completed = subprocess.run(pair_command, stdout=stream, stderr=subprocess.STDOUT, check=False)
    if completed.returncode != 0:
        raise GuardFailure(f"pair revalidation failed; see {pair_log}")
    fixed["pair_revalidated"] = verify_digest(pair_receipt, V1_EXPECTED["pair"], "revalidated pair receipt")

    cpu = min(os.sched_getaffinity(0))
    unit = f"w8-file-abba-{args.run_id}.scope"
    launcher = paths["tmp"] / "w8-noarg-launcher.py"
    launcher.write_text(launcher_source(paths, cpu, pair_receipt, paths["tmp"]), encoding="utf-8")
    launcher.chmod(0o700)
    fixed["launcher"] = digest(launcher)
    paths["launcher"] = launcher
    paths["pair_revalidated"] = pair_receipt
    paths["python"] = python
    paths["systemd_run"] = systemd_run
    paths["systemctl"] = systemctl
    paths["timeout"] = timeout
    launcher_argv = [str(python), "-B", *runner_arguments(paths, cpu, pair_receipt)]
    config = {
        "schema": SCHEMA,
        "run_id": args.run_id,
        "paths": {key: str(value) for key, value in paths.items()},
        "artifacts": fixed,
        "pair_command": pair_command,
        "launcher_command": [str(launcher)],
        "runner_argv": runner_arguments(paths, cpu, pair_receipt),
        "launcher_argv": launcher_argv,
        "launcher_cpu": cpu,
        "cpu": cpu,
        "rounds": ROUNDS,
        "deadline_seconds": int(args.deadline_seconds),
        "stop_grace_seconds": STOP_GRACE_SECONDS,
        "memory_max_bytes": MEMORY_MAX_BYTES,
        "swap_max_bytes": SWAP_MAX_BYTES,
        "file_limit_bytes": OUTPUT_FILE_LIMIT_BYTES,
        "aggregate_output_limit_bytes": AGGREGATE_OUTPUT_LIMIT_BYTES,
        "unit": unit,
        "prepared_at_ns": time.time_ns(),
    }
    config_path.parent.mkdir(parents=True, exist_ok=True)
    write_json(config_path, config)
    print(json.dumps({"config": str(config_path), "launcher": str(launcher), "cpu": cpu, "rounds": ROUNDS}))
    return 0


def validate_config(config_path: Path) -> dict[str, Any]:
    config = load_json(config_path, "configuration")
    if config.get("schema") != SCHEMA:
        raise GuardFailure("configuration schema is invalid")
    unit = config.get("unit")
    if not isinstance(unit, str) or not unit.endswith(".scope") or not RUN_ID.fullmatch(unit.removeprefix("w8-file-abba-").removesuffix(".scope")):
        raise GuardFailure("configuration unit name is invalid")
    paths_value = config.get("paths")
    artifacts = config.get("artifacts")
    if not isinstance(paths_value, dict) or not isinstance(artifacts, dict):
        raise GuardFailure("configuration paths/artifacts are incomplete")
    for name, expected in V1_EXPECTED.items():
        actual = artifacts.get("accepted_pair" if name == "pair" else name)
        if actual is None or (actual.get("bytes"), actual.get("sha256")) != expected:
            raise GuardFailure(f"configuration does not bind the reviewed {name} identity")
    for name, expected in TOOL_EXPECTED.items():
        actual = artifacts.get(name)
        if actual is None or (actual.get("bytes"), actual.get("sha256")) != expected:
            raise GuardFailure(f"configuration does not bind the reviewed {name} tool")
    paths = {key: Path(value) for key, value in paths_value.items()}
    for name, row in artifacts.items():
        if name == "launcher":
            path = paths.get("launcher")
        elif name in paths:
            path = paths[name]
        else:
            path = None
        if path is not None:
            verify_digest(path, row, name)
    if int(config.get("rounds", 0)) != ROUNDS:
        raise GuardFailure("round count is not the reviewed five")
    if int(config.get("deadline_seconds", 0)) != DEFAULT_DEADLINE_SECONDS:
        raise GuardFailure("deadline is not the reviewed twelve-hour bound")
    if int(config.get("stop_grace_seconds", 0)) != STOP_GRACE_SECONDS:
        raise GuardFailure("stop grace is not the reviewed sixty-second bound")
    if int(config.get("memory_max_bytes", 0)) != MEMORY_MAX_BYTES or int(config.get("swap_max_bytes", -1)) != SWAP_MAX_BYTES:
        raise GuardFailure("memory/swap limits are not the reviewed exact byte limits")
    if int(config.get("file_limit_bytes", 0)) != OUTPUT_FILE_LIMIT_BYTES or int(config.get("aggregate_output_limit_bytes", 0)) != AGGREGATE_OUTPUT_LIMIT_BYTES:
        raise GuardFailure("output limits are not the reviewed exact byte limits")
    cpu = int(config.get("cpu", -1))
    if cpu not in os.sched_getaffinity(0):
        raise GuardFailure(f"selected CPU {cpu} is no longer available")
    output = paths.get("output")
    tmp = paths.get("tmp")
    launcher = paths.get("launcher")
    if output is None or tmp is None or launcher is None or output == tmp:
        raise GuardFailure("output/TMP/launcher paths are incomplete")
    if output.is_symlink() or tmp.is_symlink() or not output.is_dir() or any(output.iterdir()):
        raise GuardFailure("ABBA output must be an existing empty directory")
    if not tmp.is_dir():
        raise GuardFailure("TMP directory is missing")
    if not os.access(launcher, os.X_OK):
        raise GuardFailure("no-argument launcher is not executable")
    pair_receipt = paths.get("pair_revalidated")
    if pair_receipt is None:
        raise GuardFailure("revalidated pair receipt path is missing")
    expected_runner_argv = runner_arguments(paths, cpu, pair_receipt)
    configured_runner_argv = config.get("runner_argv")
    if configured_runner_argv != expected_runner_argv:
        raise GuardFailure("configured runner argv differs from the bound paths/CPU")
    expected_launcher_argv = [str(Path(config["artifacts"]["python"]["path"])), "-B", *expected_runner_argv]
    if config.get("launcher_cpu") != cpu or config.get("launcher_argv") != expected_launcher_argv:
        raise GuardFailure("configured launcher CPU/argv differs from the runner binding")
    verify_launcher_binding(launcher, cpu, expected_launcher_argv)
    return config


def scope_state(
    systemctl: Path,
    unit: str,
    seen: dict[str, dict[str, object]],
    monitor_errors: list[str],
) -> dict[str, object]:
    props = show_unit(systemctl, unit)
    if not props.get("available"):
        error = str(props.get("monitor_error", "systemctl observation failed"))
        monitor_errors.append(f"{props.get('monitor_error_kind', 'systemctl')}: {error}")
        return props
    cgroup, pids, error = cgroup_pids(props)
    props["_cgroup_path"] = cgroup
    props["_cgroup_pids"] = pids
    if error is not None:
        props["available"] = False
        props["monitor_error_kind"] = "cgroup"
        props["monitor_error"] = error
        monitor_errors.append(f"cgroup: {error}")
        return props
    for pid in pids:
        identity = proc_identity(pid)
        if identity is not None:
            seen[f"{identity['pid']}:{identity['starttime']}"] = identity
    return props


def post_process_checks(config: dict[str, Any], state: dict[str, Any]) -> list[str]:
    paths = {key: Path(value) for key, value in config["paths"].items()}
    failures: list[str] = []
    if state.get("runner_exit_code") != 0:
        failures.append(f"systemd-run exit={state.get('runner_exit_code')}")
    if state.get("timeout_exit_code") not in (None, 0):
        failures.append(f"timeout exit={state.get('timeout_exit_code')}")
    if state.get("deadline_exceeded"):
        failures.append("external deadline exceeded")
    if state.get("output_limit_exceeded"):
        failures.append("aggregate output limit exceeded")
    if state.get("monitor_errors"):
        failures.append("active unit monitor errors: " + repr(state["monitor_errors"]))
    fatal_post_errors = [
        error for error in state.get("post_monitor_errors", [])
        if not str(error).startswith("unit-gone:")
    ]
    if fatal_post_errors:
        failures.append("post-run unit monitor errors: " + repr(fatal_post_errors))
    if state.get("post_monitor_errors") and not state.get("unit_gc_expected"):
        failures.append("post-run unit monitor errors: " + repr(state["post_monitor_errors"]))
    if state.get("unit_metadata_missing"):
        failures.append("transient unit status/MemoryPeak was never observed")
    peak = state.get("memory_peak_bytes")
    if not isinstance(peak, int) or peak < 0:
        failures.append("MemoryPeak was not retained")
    elif peak > int(config["memory_max_bytes"]):
        failures.append(f"MemoryPeak exceeded cap: {peak}")
    current_max = state.get("memory_current_max_bytes")
    if not isinstance(current_max, int) or current_max < 0:
        failures.append("MemoryCurrent was not retained")
    elif current_max > int(config["memory_max_bytes"]):
        failures.append(f"MemoryCurrent exceeded cap: {current_max}")
    if state.get("leftover_processes"):
        failures.append("scope descendants survived: " + repr(state["leftover_processes"]))
    if state.get("leftover_cgroup_pids"):
        failures.append("cgroup processes survived: " + repr(state["leftover_cgroup_pids"]))
    if directory_bytes((paths["output"], paths["tmp"])) > int(config["aggregate_output_limit_bytes"]):
        failures.append("final aggregate output exceeds cap")
    for root in (paths["output"], paths["tmp"]):
        for path in root.rglob("*"):
            if path.is_file() and path.stat().st_size > int(config["file_limit_bytes"]):
                failures.append(f"lifecycle file exceeds RLIMIT_FSIZE: {path}")
    rows, blockers = forbidden_rows()
    state["post_quiet_blockers"] = blockers
    state["post_process_count"] = len(rows)
    if blockers:
        failures.append("forbidden process remains after run")
    return failures


def write_failure_receipt(config: dict[str, Any], receipt: dict[str, Any]) -> Path:
    tmp = Path(config["paths"]["tmp"])
    path = tmp / "w8-quiet-abba-receipt.json"
    write_json(path, receipt)
    return path


def invoke_analyzer(config: dict[str, Any], state: dict[str, Any], owned: dict[str, Any]) -> None:
    paths = {key: Path(value) for key, value in config["paths"].items()}
    output = paths["output"]
    manifest_path = output / "manifest.json"
    if not manifest_path.is_file():
        raise GuardFailure("runner did not produce manifest.json")
    manifest = load_json(manifest_path, "runner manifest")
    if manifest.get("schema") != "litchi_file_qualification.v3" or manifest.get("status") != "complete; analyze.py review pending":
        raise GuardFailure("runner manifest is not complete")
    routes = manifest.get("routes")
    if not isinstance(routes, list) or len(routes) != TARGET_ROUTE_COUNT or len(set(routes)) != TARGET_ROUTE_COUNT:
        raise GuardFailure("runner manifest does not contain exactly 109 routes")
    if not isinstance(manifest.get("calibration"), list) or len(manifest["calibration"]) != TARGET_ROUTE_COUNT:
        raise GuardFailure("runner manifest calibration row count is incomplete")
    if not isinstance(manifest.get("abba"), list) or len(manifest["abba"]) != TARGET_ROUTE_COUNT:
        raise GuardFailure("runner manifest ABBA row count is incomplete")
    receipts = list(output.rglob("*.receipt.json"))
    if len(receipts) != 2 + 2 * TARGET_ROUTE_COUNT + 4 * ROUNDS * TARGET_ROUTE_COUNT:
        raise GuardFailure(f"runner receipt count is not 2400: {len(receipts)}")
    stdout = paths["tmp"] / "analyzer.stdout"
    stderr = paths["tmp"] / "analyzer.stderr"
    command = [str(Path(config["artifacts"]["python"]["path"])), "-B", str(paths["analyzer"]), str(output)]
    started = time.monotonic()
    with stdout.open("wb") as out, stderr.open("wb") as err:
        process = subprocess.Popen(command, stdout=out, stderr=err, start_new_session=True)
        owned["analyzer"] = process
        analyzer_seen = state.setdefault("seen_analyzer_processes", {})
        capture_descendants(process.pid, analyzer_seen)
        while process.poll() is None:
            capture_descendants(process.pid, analyzer_seen)
            if directory_bytes((paths["output"], paths["tmp"])) > int(config["aggregate_output_limit_bytes"]):
                state["analyzer_output_limit_exceeded"] = True
                cleanup_error = terminate_process_group(process)
                if cleanup_error:
                    state["analyzer_cleanup_error"] = cleanup_error
                    raise GuardFailure(
                        "aggregate output limit exceeded during analyzer; "
                        f"analyzer cleanup failed: {cleanup_error}"
                    )
                raise GuardFailure("aggregate output limit exceeded during analyzer")
            if time.monotonic() - started > 600:
                state["analyzer_timeout"] = True
                cleanup_error = terminate_process_group(process)
                if cleanup_error:
                    state["analyzer_cleanup_error"] = cleanup_error
                    raise GuardFailure(f"analyzer exceeded its ten-minute deadline; cleanup failed: {cleanup_error}")
                raise GuardFailure("analyzer exceeded its ten-minute deadline")
            time.sleep(POLL_SECONDS)
        completed_returncode = process.wait()
        # Keep the Popen handle until the outer finally block.  A successful
        # analyzer may have left a child in its owned process group; the final
        # cleanup must still inspect and reap that group before qualification.
    state["analyzer_command"] = command
    state["analyzer_exit_code"] = completed_returncode
    state["analyzer_stdout"] = str(stdout)
    state["analyzer_stderr"] = str(stderr)
    if completed_returncode != 0 or not (output / "analysis.json").is_file():
        raise GuardFailure("final analyzer failed; see analyzer stdout/stderr")


def run_config(config_path: Path) -> int:
    try:
        config = validate_config(config_path)
    except Exception as error:
        try:
            raw = load_json(config_path, "configuration")
            tmp_value = raw.get("paths", {}).get("tmp")
            if isinstance(tmp_value, str):
                tmp = Path(tmp_value)
                tmp.mkdir(parents=True, exist_ok=True)
                receipt = {
                    "schema": "litchi_w8_quiet_abba_receipt.v1",
                    "status": "failed",
                    "failure": str(error),
                    "config": digest(config_path),
                    "finished_ns": time.time_ns(),
                }
                path = tmp / "w8-quiet-abba-receipt.json"
                write_json(path, receipt)
                print(json.dumps({"status": "failed", "receipt": str(path), "receipt_sha256": sha256_text(path), "error": str(error)}), file=sys.stderr)
                return 1
        except Exception:
            pass
        print(f"0362ul32 failed closed before receipt creation: {error}", file=sys.stderr)
        return 2
    paths = {key: Path(value) for key, value in config["paths"].items()}
    tmp = paths["tmp"]
    receipt: dict[str, Any] = {
        "schema": "litchi_w8_quiet_abba_receipt.v1",
        "config": digest(config_path),
        "status": "failed",
        "started_ns": time.time_ns(),
        "unit": config["unit"],
        "cpu": config["cpu"],
        "rounds": config["rounds"],
        "runner_command": config["runner_argv"],
        "launcher": config["artifacts"]["launcher"],
        "unit_samples": [],
    }
    systemctl = Path(config["artifacts"]["systemctl"]["path"])
    owned: dict[str, Any] = {"systemd": None, "analyzer": None}
    previous_handlers: dict[int, Any] = {}
    state: dict[str, Any] = {"seen_scope_processes": {}, "last_cgroup_path": None}
    outcome_error: BaseException | None = None

    def interrupt(signum: int, _frame: Any) -> None:
        receipt["signal"] = signum
        raise GuardFailure(f"received signal {signum}; owned lifecycle cleanup required")

    try:
        for signum in (signal.SIGINT, signal.SIGTERM, signal.SIGHUP):
            previous_handlers[signum] = signal.getsignal(signum)
            signal.signal(signum, interrupt)
        preflight = quiet_snapshot(tmp / "preflight.host.json", "preflight")
        receipt["preflight"] = preflight
        if preflight["blockers"]:
            raise GuardFailure("host is not quiet before launch")
        command = systemd_command(config, paths)
        receipt["systemd_command"] = command
        systemd_stdout = tmp / "systemd-run.stdout"
        systemd_stderr = tmp / "systemd-run.stderr"
        state: dict[str, Any] = {
            "started_monotonic": time.monotonic(),
            "unit_metadata_missing": True,
            "deadline_exceeded": False,
            "output_limit_exceeded": False,
            "memory_peak_bytes": None,
            "memory_current_max_bytes": None,
            "leftover_processes": [],
            "leftover_cgroup_pids": [],
            "monitor_errors": [],
            "post_monitor_errors": [],
            "post_monitor_available_samples": 0,
            "seen_scope_processes": {},
            "last_cgroup_path": None,
        }

        def stop_scope() -> None:
            if state.get("stop_requested"):
                return
            state["stop_requested"] = True
            result = kill_unit(systemctl, str(config["unit"]))
            if result is None:
                state["monitor_errors"].append("systemd-kill: command did not execute")
            elif result.returncode != 0:
                state["monitor_errors"].append(
                    f"systemd-kill: exit={result.returncode}: {result.stderr[-1024:]}"
                )

        with systemd_stdout.open("wb") as out, systemd_stderr.open("wb") as err:
            child = subprocess.Popen(command, stdout=out, stderr=err, start_new_session=True)
            owned["systemd"] = child
            while child.poll() is None:
                props = scope_state(systemctl, str(config["unit"]), state["seen_scope_processes"], state["monitor_errors"])
                receipt["unit_samples"].append(props)
                if props.get("_cgroup_path"):
                    state["last_cgroup_path"] = props["_cgroup_path"]
                if not props.get("available"):
                    stop_scope()
                if props.get("available"):
                    state["unit_metadata_missing"] = False
                    if props.get("MemoryPeak", "").isdigit():
                        state["memory_peak_bytes"] = max(int(state["memory_peak_bytes"] or 0), int(props["MemoryPeak"]))
                    if props.get("MemoryCurrent", "").isdigit():
                        state["memory_current_max_bytes"] = max(int(state["memory_current_max_bytes"] or 0), int(props["MemoryCurrent"]))
                tracked = directory_bytes((paths["output"], paths["tmp"]))
                receipt["aggregate_bytes"] = tracked
                if tracked > int(config["aggregate_output_limit_bytes"]):
                    state["output_limit_exceeded"] = True
                    stop_scope()
                if time.monotonic() - state["started_monotonic"] > int(config["deadline_seconds"]) + int(config["stop_grace_seconds"]):
                    state["deadline_exceeded"] = True
                    stop_scope()
                time.sleep(POLL_SECONDS)
            state["runner_exit_code"] = child.wait()
            receipt["runner_exit_code"] = state["runner_exit_code"]
        for _ in range(3):
            props = scope_state(systemctl, str(config["unit"]), state["seen_scope_processes"], state["post_monitor_errors"])
            receipt["unit_samples"].append(props)
            if props.get("_cgroup_path"):
                state["last_cgroup_path"] = props["_cgroup_path"]
            if props.get("available"):
                state["post_monitor_available_samples"] += 1
            if props.get("available"):
                state["unit_metadata_missing"] = False
                if props.get("MemoryPeak", "").isdigit():
                    state["memory_peak_bytes"] = max(int(state["memory_peak_bytes"] or 0), int(props["MemoryPeak"]))
                if props.get("MemoryCurrent", "").isdigit():
                    state["memory_current_max_bytes"] = max(int(state["memory_current_max_bytes"] or 0), int(props["MemoryCurrent"]))
            time.sleep(0.05)
        receipt["systemd_stdout"] = digest(systemd_stdout)
        receipt["systemd_stderr"] = digest(systemd_stderr)
        receipt["memory_peak_bytes"] = state["memory_peak_bytes"]
        receipt["memory_current_max_bytes"] = state["memory_current_max_bytes"]
        receipt["unit_metadata_missing"] = state["unit_metadata_missing"]
        receipt["seen_scope_processes"] = list(state["seen_scope_processes"].values())
        if receipt["unit_samples"] and not receipt["unit_samples"][-1].get("available", False):
            receipt["unit_gone_after_last_sample"] = True
        receipt["timeout_exit_code"] = state.get("runner_exit_code")
        state["timeout_exit_code"] = state.get("runner_exit_code")
        state["leftover_processes"] = [item for item in state["seen_scope_processes"].values() if alive(item)]
        last_available = next((item for item in reversed(receipt["unit_samples"]) if item.get("available")), {})
        last_pids = list(last_available.get("_cgroup_pids", []))
        state["leftover_cgroup_pids"] = last_pids
        receipt["leftover_processes"] = state["leftover_processes"]
        receipt["leftover_cgroup_pids"] = last_pids
        available_units = [item for item in receipt["unit_samples"] if item.get("available")]
        last_unit = available_units[-1] if available_units else {}
        if isinstance(last_unit.get("ExecMainStatus"), str) and last_unit["ExecMainStatus"].isdigit():
            state["timeout_exit_code"] = int(last_unit["ExecMainStatus"])
        else:
            state["timeout_exit_code"] = state.get("runner_exit_code")
        receipt["unit_last"] = last_unit
        receipt["timeout_exit_code"] = state["timeout_exit_code"]
        receipt["monitor_errors"] = state["monitor_errors"]
        receipt["post_monitor_errors"] = state["post_monitor_errors"]
        receipt["post_monitor_available_samples"] = state["post_monitor_available_samples"]
        state["unit_gc_expected"] = bool(
            not state["unit_metadata_missing"]
            and state["post_monitor_errors"]
            and state["post_monitor_available_samples"] == 0
            and all(str(error).startswith("unit-gone:") for error in state["post_monitor_errors"])
        )
        receipt["unit_gc_expected"] = state["unit_gc_expected"]
        failures = post_process_checks(config, state)
        receipt["guard_failures"] = failures
        if failures:
            raise GuardFailure("; ".join(failures))
        invoke_analyzer(config, state, owned)
        post_analyzer = quiet_snapshot(tmp / "post-analyzer.host.json", "post-analyzer")
        receipt["post_analyzer"] = post_analyzer
        if post_analyzer["blockers"]:
            raise GuardFailure("host is not quiet after final analysis")
        final_failures = post_process_checks(config, state)
        receipt["final_guard_failures"] = final_failures
        if final_failures:
            raise GuardFailure("; ".join(final_failures))
        receipt.update(state)
        receipt["status"] = "qualified"
    except BaseException as error:
        outcome_error = error
    finally:
        for signum in (signal.SIGINT, signal.SIGTERM, signal.SIGHUP):
            signal.signal(signum, signal.SIG_IGN)
        for key in ("seen_scope_processes", "seen_analyzer_processes"):
            values = state.get(key)
            if isinstance(values, dict):
                receipt[key] = list(values.values())
        for key in ("analyzer_command", "analyzer_exit_code", "analyzer_stdout", "analyzer_stderr"):
            if key in state:
                receipt[key] = state[key]
        cleanup = cleanup_owned(systemctl, str(config["unit"]), owned.get("systemd"), owned.get("analyzer"), state)
        receipt["cleanup"] = cleanup
        if cleanup.get("errors") and outcome_error is None:
            outcome_error = GuardFailure("lifecycle cleanup failed: " + repr(cleanup["errors"]))
        for signum, handler in previous_handlers.items():
            signal.signal(signum, handler)
        receipt["finished_ns"] = time.time_ns()
        receipt["status"] = "failed" if outcome_error is not None else "qualified"
        if outcome_error is not None:
            receipt["failure"] = str(outcome_error)
        receipt_path = write_failure_receipt(config, receipt)
    if outcome_error is not None:
        print(json.dumps({"status": "failed", "receipt": str(receipt_path), "receipt_sha256": sha256_text(receipt_path), "error": str(outcome_error)}), file=sys.stderr)
        return 1
    print(json.dumps({"status": "qualified", "receipt": str(receipt_path), "receipt_sha256": sha256_text(receipt_path), "analysis": str(paths["output"] / "analysis.json")}))
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest="command", required=True)
    prepare_parser = commands.add_parser("prepare")
    prepare_parser.add_argument("--root", required=True)
    prepare_parser.add_argument("--v1", required=True)
    prepare_parser.add_argument("--config", required=True)
    prepare_parser.add_argument("--abba-output", required=True)
    prepare_parser.add_argument("--tmp-dir", required=True)
    prepare_parser.add_argument("--run-id", required=True)
    prepare_parser.add_argument("--deadline-seconds", type=int, default=DEFAULT_DEADLINE_SECONDS)
    run_parser = commands.add_parser("run")
    run_parser.add_argument("--config", required=True, type=Path)
    return parser


def main() -> int:
    args = build_parser().parse_args()
    try:
        if args.command == "prepare":
            return prepare(args)
        if args.command == "run":
            return run_config(args.config.resolve())
        raise GuardFailure(f"unknown command: {args.command}")
    except GuardFailure as error:
        print(f"0362ul32 failed closed: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
