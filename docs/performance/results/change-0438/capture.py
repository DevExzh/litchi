#!/usr/bin/env python3
"""Capture one serialized 0438 same-API ODP ABBA phase.

The driver owns process execution, immutable receipt writing, and source
custody.  ODP report semantics remain in the copied external oracle; this
driver does not duplicate that schema.  It is intentionally path-parameterized
so an exported evidence bundle does not inherit stale repository or cleanup
paths from an earlier change.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
PHASES = {
    "A1": ("before-streaming", 0, 6),
    "B1": ("after-streaming", 0, 6),
    "B2": ("after-streaming", 6, 12),
    "A2": ("before-streaming", 6, 12),
}
SHAPES = {"tiny": 64, "medium": 4_096, "large": 8_192}
ATTEMPT_RE = re.compile(r"^[a-z0-9][a-z0-9_-]{0,63}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")


class CaptureError(ValueError):
    pass


def fail(message: str) -> None:
    raise CaptureError(message)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def write(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def record(path: Path) -> dict[str, Any]:
    return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size, "sha256": sha(path)}


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty text")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label)
    if HEX64.fullmatch(value) is None:
        fail(f"{label}: expected SHA-256")
    return value.lower()


def import_custody(path: Path):
    spec = importlib.util.spec_from_file_location("change0438_custody", path)
    if spec is None or spec.loader is None:
        fail(f"cannot load custody driver {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if not callable(getattr(module, "sources", None)):
        fail(f"custody driver {path} has no sources()")
    return module


def resolve_repo(value: Any, repo: Path, label: str) -> Path:
    path = Path(text(value, label))
    return path if path.is_absolute() else (repo / path)


def status_outside_bundle(repo: Path) -> list[str]:
    raw = subprocess.check_output(
        ["git", "status", "--short", "--untracked-files=all"], cwd=repo, text=True
    )
    try:
        relative_bundle = ROOT.resolve().relative_to(repo.resolve()).as_posix() + "/"
    except ValueError:
        relative_bundle = None
    if relative_bundle is None:
        return raw.splitlines()
    return [line for line in raw.splitlines() if not line[3:].startswith(relative_bundle)]


def build_for(role: str, protocol: dict[str, Any], require_binary: bool, repo: Path,
              protocol_path: Path | None = None) -> tuple[Path, dict[str, Any]]:
    role_spec = obj(obj(protocol.get("roles"), "protocol.roles").get(role), f"protocol.roles.{role}")
    build_dir = ROOT / text(role_spec.get("build_directory"), f"protocol.roles.{role}.build_directory")
    build_path = build_dir / "build.json"
    build = obj(load(build_path), str(build_path))
    if build.get("change") != 438 or build.get("role") not in {None, role, build_dir.name}:
        fail(f"{build_path}: wrong change or role")
    bound_protocol = protocol_path or (ROOT / "protocol.json")
    if not bound_protocol.is_file():
        fail(f"{bound_protocol}: protocol input is missing")
    if build.get("protocol_sha256") != sha(bound_protocol):
        fail(f"{build_path}: protocol hash is stale")
    binaries = obj(build.get("binaries"), f"{build_path}.binaries")
    if set(binaries) != {"normal", "allocator"}:
        fail(f"{build_path}: normal and allocator binaries are required")
    if require_binary:
        for mode in ("normal", "allocator"):
            identity = obj(binaries[mode], f"{build_path}.binaries.{mode}")
            path = resolve_repo(identity.get("path"), repo, f"{build_path}.binaries.{mode}.path")
            size = identity.get("bytes")
            expected = digest(identity.get("sha256"), f"{build_path}.binaries.{mode}.sha256")
            if not path.is_file() or isinstance(size, bool) or not isinstance(size, int):
                fail(f"{build_path}: {mode} binary is missing")
            if path.stat().st_size != size or sha(path) != expected:
                fail(f"{build_path}: {mode} binary identity is stale")
    return build_dir, build


def oracle_role(protocol: dict[str, Any], role: str) -> str:
    spec = obj(obj(protocol.get("roles"), "protocol.roles").get(role), f"protocol.roles.{role}")
    if isinstance(spec.get("oracle_role"), str):
        return spec["oracle_role"]
    oracle = protocol.get("oracle")
    if isinstance(oracle, dict) and isinstance(oracle.get("roles"), dict):
        mapped = oracle["roles"]
        if isinstance(mapped.get(role), str):
            return mapped[role]
    return role


def oracle_command(protocol: dict[str, Any], report: Path, mode: str, shape: str, role: str) -> list[str]:
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    verifier_value = oracle.get("verifier_path", oracle.get("path"))
    verifier = ROOT / text(verifier_value, "protocol.oracle.verifier_path")
    if not verifier.is_file():
        fail(f"{verifier}: copied ODP oracle verifier is missing")
    template = oracle.get("argv")
    if template is None:
        template = [
            sys.executable, "-B", "{verifier}", "--report", "{report}",
            "--mode", "{mode}", "--shape", "{shape}", "--role", "{role}",
        ]
    if not isinstance(template, list) or any(not isinstance(item, str) for item in template):
        fail("protocol.oracle.argv: expected a string array")
    values = {
        "verifier": str(verifier), "report": str(report), "mode": mode,
        "shape": shape, "role": oracle_role(protocol, role),
    }
    try:
        return [item.format(**values) for item in template]
    except (KeyError, ValueError) as error:
        fail(f"protocol.oracle.argv: invalid template: {error}")
    raise AssertionError("unreachable")


def lane_artifacts(directory: Path, name: str) -> dict[str, Path]:
    return {
        "report": directory / f"{name}.json",
        "catalog": directory / f"{name}-catalog.json",
        "workload_log": directory / f"{name}.log",
        "resource_log": directory / f"{name}-resource.log",
        "oracle_log": directory / f"{name}-oracle.log",
        "receipt": directory / f"{name}-receipt.json",
    }


def cli_flag(protocol: dict[str, Any], name: str, default: str) -> str:
    cli = protocol.get("workload_cli", {})
    if cli is None:
        cli = {}
    cli = obj(cli, "protocol.workload_cli")
    return text(cli.get(name, default), f"protocol.workload_cli.{name}")


def workload_argv(protocol: dict[str, Any], binary: Path, selector: str, shape: str, mode: str,
                  report: Path, catalog: Path) -> list[str]:
    del mode  # allocator mode is selected by the executable identity.
    return [
        str(binary), cli_flag(protocol, "case_flag", "--case"), selector,
        cli_flag(protocol, "shape_flag", "--semantic-shape"), shape,
        cli_flag(protocol, "workers_flag", "--workers"), str(protocol["workers"]),
        cli_flag(protocol, "samples_flag", "--samples"), str(protocol["samples"]),
        cli_flag(protocol, "warmup_flag", "--warmup"), str(protocol["warmups"]),
        cli_flag(protocol, "report_flag", "--json"), str(report),
        cli_flag(protocol, "catalog_flag", "--corpus-manifest"), str(catalog),
    ]


def expected_lanes(protocol: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    order = protocol.get("order")
    if not isinstance(order, list) or len(order) != 12:
        fail("protocol.order: expected 12 lanes")
    _, first, last = PHASES[phase]
    lanes = order[first:last]
    if len(lanes) != 6:
        fail(f"protocol.order: phase {phase} does not contain six lanes")
    checked = [obj(value, f"protocol.order[{first + index}]") for index, value in enumerate(lanes)]
    expected_repeat = "R1" if phase in {"A1", "B1"} else "R2"
    expected = {
        (mode, shape, expected_repeat)
        for mode in ("normal", "allocator")
        for shape in ("tiny", "medium", "large")
    }
    actual = {
        (value.get("mode"), value.get("shape"), value.get("repeat"))
        for value in checked
    }
    if actual != expected or len(actual) != len(checked):
        fail(f"protocol.order: phase {phase} is not the six-lane {expected_repeat} matrix")
    return checked


def run_phase(phase: str, attempt: str, repo: Path, protocol_path: Path, custody_path: Path) -> int:
    protocol = obj(load(protocol_path), str(protocol_path))
    if protocol.get("change") != 438:
        fail("protocol.change: must be 438")
    role, _, _ = PHASES[phase]
    lanes = expected_lanes(protocol, phase)
    build_root, build = build_for(role, protocol, True, repo, protocol_path)
    _, ambient_build = build_for("after-streaming", protocol, True, repo, protocol_path)
    custody = import_custody(custody_path)
    source_before = custody.sources()
    if source_before != ambient_build.get("source_manifest"):
        fail("ambient source manifest differs from after-streaming build")
    status_before = status_outside_bundle(repo)
    directory = ROOT / "runs" / phase / attempt
    if directory.exists() and any(directory.iterdir()):
        fail(f"capture directory already contains evidence: {directory}")
    directory.mkdir(parents=True, exist_ok=True)
    cpu = protocol.get("cpu")
    workers = protocol.get("workers")
    if isinstance(cpu, bool) or not isinstance(cpu, int) or cpu != 2 or cpu not in os.sched_getaffinity(0):
        fail("protocol.cpu: CPU 2 is required and unavailable")
    if workers != 1 or protocol.get("samples") != 30 or protocol.get("warmups") != 3:
        fail("protocol: expected workers=1, samples=30, warmups=3")
    if protocol.get("shapes") != SHAPES:
        fail("protocol.shapes: expected tiny=64, medium=4096, large=8192")
    outer_protocol_sha = sha(protocol_path)
    failed = False
    index: list[str] = []
    for lane in lanes:
        mode = text(lane.get("mode"), "lane.mode")
        shape = text(lane.get("shape"), "lane.shape")
        repeat = text(lane.get("repeat"), "lane.repeat")
        if mode not in {"normal", "allocator"} or shape not in {"tiny", "medium", "large"} or repeat not in {"R1", "R2"}:
            fail(f"lane {lane}: mode, shape, or repeat is outside the ODP matrix")
        name = f"{phase}-{mode}-{shape}-{repeat.lower()}"
        artifacts = lane_artifacts(directory, name)
        if any(path.exists() for path in artifacts.values()):
            fail(f"lane already has an artifact: {name}")
        role_spec = obj(protocol["roles"][role], f"protocol.roles.{role}")
        selector = text(role_spec.get("selector"), f"protocol.roles.{role}.selector")
        binary_desc = obj(build["binaries"][mode], f"{build_root}/build.json.binaries.{mode}")
        binary_path = resolve_repo(binary_desc["path"], repo, f"{name}.binary")
        argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(artifacts["resource_log"])] + workload_argv(
            protocol, binary_path, selector, shape, mode, artifacts["report"], artifacts["catalog"]
        )
        row: dict[str, Any] = {
            "schema": "litchi-0438-capture-receipt-v1", "change": 438,
            "phase": phase, "attempt": attempt, "role": role, "build_directory": build_root.name,
            "selector": selector, "source_field": role_spec.get("source_field"), "name": name,
            "lane": lane, "argv": argv, "cwd": str(repo), "revision": build.get("revision"),
            "source_manifest": build.get("source_manifest"), "binary": binary_desc,
            "protocol_sha256": outer_protocol_sha, "oracle_role": oracle_role(protocol, role),
            "driver_sha256": sha(Path(__file__)), "custody_driver": str(custody_path),
            "source_before": source_before, "ambient_source_manifest": ambient_build.get("source_manifest"),
            "status_before": status_before, "started_utc": now(), "status": "running",
        }
        write(artifacts["receipt"], row)
        print(json.dumps({"status": "running", "phase": phase, "lane": name}), flush=True)
        try:
            env = os.environ.copy()
            env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": "", "PYTHONDONTWRITEBYTECODE": "1"})
            with artifacts["workload_log"].open("xb") as output:
                result = subprocess.run(argv, cwd=repo, env=env, stdout=output, stderr=subprocess.STDOUT)
            row["exit_code"] = result.returncode
            if result.returncode != 0:
                raise RuntimeError(f"producer exited {result.returncode}")
            oracle_argv = oracle_command(protocol, artifacts["report"], mode, shape, role)
            oracle_result = subprocess.run(oracle_argv, cwd=repo, env=env, capture_output=True, text=True)
            artifacts["oracle_log"].write_text(
                "argv=" + json.dumps(oracle_argv) + "\nstdout=" + oracle_result.stdout + "stderr=" + oracle_result.stderr,
                encoding="utf-8",
            )
            row["oracle_exit_code"] = oracle_result.returncode
            row["oracle_stdout_sha256"] = sha(artifacts["oracle_log"])
            if oracle_result.returncode != 0 or oracle_result.stdout.strip() != "VALID":
                raise RuntimeError("copied ODP oracle rejected the report")
            row["status"] = "pass"
        except Exception as error:
            row["status"] = "failed"
            row["error"] = repr(error)
            failed = True
        finally:
            row["finished_utc"] = now()
            row["source_after"] = custody.sources()
            row["status_after"] = status_outside_bundle(repo)
            row["source_unchanged"] = row["source_before"] == row["source_after"]
            row["outside_bundle_status_unchanged"] = row["status_before"] == row["status_after"]
            row["artifacts"] = {
                key: record(path) for key, path in artifacts.items() if key != "receipt" and path.is_file()
            }
            write(artifacts["receipt"], row)
        index.append(str(artifacts["receipt"].relative_to(ROOT)))
        if failed:
            break
    state = {
        "schema": "litchi-0438-capture-state-v1", "change": 438, "phase": phase,
        "attempt": attempt, "role": role, "expected_lanes": len(lanes), "completed_lanes": len(index),
        "status": "failed" if failed else "pass", "source_before": source_before,
        "source_after": custody.sources(), "status_before": status_before,
        "status_after": status_outside_bundle(repo), "source_manifest": build.get("source_manifest"),
        "ambient_source_manifest": ambient_build.get("source_manifest"), "protocol_sha256": outer_protocol_sha,
        "driver_sha256": sha(Path(__file__)), "index": index, "finished_utc": now(),
    }
    write(directory / "capture-index.json", index)
    write(directory / "capture-state.json", state)
    if failed or len(index) != len(lanes):
        return 1
    if custody.sources() != source_before or status_outside_bundle(repo) != status_before:
        fail("source custody or outside-bundle status changed during capture")
    print(json.dumps({"status": "pass", "phase": phase, "role": role, "reports": len(index)}))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--phase", choices=tuple(PHASES), required=True)
    parser.add_argument("--attempt", default="formal")
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol.json")
    parser.add_argument("--custody-driver", type=Path, required=True)
    args = parser.parse_args()
    if ATTEMPT_RE.fullmatch(args.attempt) is None:
        parser.error("--attempt must match [a-z0-9][a-z0-9_-]{0,63}")
    try:
        return run_phase(args.phase, args.attempt, args.repo_root.resolve(), args.protocol.resolve(), args.custody_driver.resolve())
    except (OSError, KeyError, TypeError, ValueError, subprocess.CalledProcessError, CaptureError) as error:
        print(f"CAPTURE INVALID: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
