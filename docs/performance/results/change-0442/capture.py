#!/usr/bin/env python3
"""Capture one serialized 0442 ABBA phase.

This driver only runs the declared workload and copied oracle.  It records
source custody, binary/report identity, exact commands, GNU-time output, and
artifact hashes; semantic and append correctness remain oracle-owned.
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
    "A1": ("before", 0, 6, "R1"),
    "B1": ("after", 6, 12, "R1"),
    "B2": ("after", 12, 18, "R2"),
    "A2": ("before", 18, 24, "R2"),
}
SHAPES = {"tiny": 64, "medium": 4_096, "large": 8_192}
HEX40 = re.compile(r"^[0-9a-fA-F]{40}$")
HEX64 = re.compile(r"^[0-9a-fA-F]{64}$")
ATTEMPT_RE = re.compile(r"^[a-z0-9][a-z0-9_-]{0,63}$")


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
        fail(f"{label}: expected non-empty string")
    return value


def digest(value: Any, label: str) -> str:
    value = text(value, label)
    if HEX64.fullmatch(value) is None:
        fail(f"{label}: expected SHA-256")
    return value.lower()


def import_custody(path: Path):
    spec = importlib.util.spec_from_file_location("change0442_custody", path)
    if spec is None or spec.loader is None:
        fail(f"cannot load custody driver {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if not callable(getattr(module, "sources", None)):
        fail(f"custody driver {path} has no sources()")
    return module


def resolve_repo(value: Any, repo: Path, label: str) -> Path:
    path = Path(text(value, label))
    return (repo / path).resolve() if not path.is_absolute() else path.resolve()


def status_outside_bundle(repo: Path) -> list[str]:
    """Return worktree status while ignoring evidence written by this bundle.

    The formal workload is normally run from a clean role worktree, but the
    same driver is also copied into a repository checkout during review.  In
    that case the driver's own receipts must not turn a clean-custody check
    into a false failure.  Every other tracked or untracked path remains
    visible and is therefore part of the custody contract.
    """
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


def build_for(role: str, protocol: dict[str, Any], repo: Path, protocol_path: Path) -> tuple[Path, dict[str, Any]]:
    role_spec = obj(obj(protocol.get("roles"), "protocol.roles").get(role), f"protocol.roles.{role}")
    directory = ROOT / text(role_spec.get("build_directory"), f"protocol.roles.{role}.build_directory")
    path = directory / "build.json"
    build = obj(load(path), str(path))
    if build.get("change") != 442 or build.get("role") not in {None, role, directory.name}:
        fail(f"{path}: build identity differs")
    if build.get("protocol_sha256") != sha(protocol_path):
        fail(f"{path}: protocol binding is stale")
    if build.get("capture_driver_sha256") != sha(ROOT / "capture.py"):
        fail(f"{path}: capture driver binding is stale")
    binaries = obj(build.get("binaries"), f"{path}.binaries")
    if set(binaries) != {"normal", "allocator"}:
        fail(f"{path}: normal and allocator binaries are required")
    for mode, identity in binaries.items():
        identity = obj(identity, f"{path}.binaries.{mode}")
        binary = resolve_repo(identity.get("path"), repo, f"{path}.binaries.{mode}.path")
        size = identity.get("bytes")
        expected = digest(identity.get("sha256"), f"{path}.binaries.{mode}.sha256")
        if isinstance(size, bool) or not isinstance(size, int) or size <= 0 or not binary.is_file():
            fail(f"{path}.binaries.{mode}: missing binary")
        if binary.stat().st_size != size or sha(binary) != expected:
            fail(f"{path}.binaries.{mode}: binary hash differs")
    revision = build.get("revision")
    if not isinstance(revision, str) or HEX40.fullmatch(revision) is None:
        fail(f"{path}.revision: expected Git SHA-1")
    return directory, build


def oracle_files(protocol: dict[str, Any]) -> tuple[Path, Path, str, str]:
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    verifier = ROOT / text(oracle.get("verifier_path"), "protocol.oracle.verifier_path")
    oracle_protocol = ROOT / text(oracle.get("protocol_path"), "protocol.oracle.protocol_path")
    verifier_sha = digest(oracle.get("verifier_sha256"), "protocol.oracle.verifier_sha256")
    oracle_protocol_sha = digest(oracle.get("protocol_sha256"), "protocol.oracle.protocol_sha256")
    if not verifier.is_file() or sha(verifier) != verifier_sha:
        fail("protocol.oracle.verifier_sha256: copied verifier is missing or differs from its frozen hash")
    if not oracle_protocol.is_file() or sha(oracle_protocol) != oracle_protocol_sha:
        fail("protocol.oracle.protocol_sha256: copied oracle protocol is missing or differs from its frozen hash")
    return verifier, oracle_protocol, verifier_sha, oracle_protocol_sha


def oracle_command(protocol: dict[str, Any], report: Path, mode: str, shape: str, role: str) -> list[str]:
    oracle = obj(protocol.get("oracle"), "protocol.oracle")
    verifier, _oracle_protocol, _verifier_sha, _oracle_protocol_sha = oracle_files(protocol)
    template = oracle.get("argv")
    if not isinstance(template, list) or any(not isinstance(item, str) for item in template):
        fail("protocol.oracle.argv: expected string array")
    role_map = obj(oracle.get("role_map", {}), "protocol.oracle.role_map")
    if role not in role_map or not isinstance(role_map[role], str) or not role_map[role]:
        fail(f"protocol.oracle.role_map: missing role {role}")
    values = {"python": sys.executable, "verifier": str(verifier), "report": str(report), "mode": mode, "shape": shape, "role": role_map[role]}
    try:
        return [item.format(**values) for item in template]
    except (KeyError, ValueError) as error:
        fail(f"protocol.oracle.argv: invalid template: {error}")
    raise AssertionError("unreachable")


def cli(protocol: dict[str, Any], key: str, default: str) -> str:
    values = obj(protocol.get("workload_cli", {}), "protocol.workload_cli")
    return text(values.get(key, default), f"protocol.workload_cli.{key}")


def workload_argv(protocol: dict[str, Any], binary: Path, selector: str, shape: str, report: Path, catalog: Path) -> list[str]:
    return [
        str(binary), cli(protocol, "case_flag", "--case"), selector,
        cli(protocol, "shape_flag", "--semantic-shape"), shape,
        cli(protocol, "workers_flag", "--workers"), str(protocol["workers"]),
        cli(protocol, "samples_flag", "--samples"), str(protocol["samples"]),
        cli(protocol, "warmup_flag", "--warmup"), str(protocol["warmups"]),
        cli(protocol, "report_flag", "--json"), str(report),
        cli(protocol, "catalog_flag", "--corpus-manifest"), str(catalog),
    ]


def expected_lanes(protocol: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    role, first, last, repeat = PHASES[phase]
    order = protocol.get("order")
    if not isinstance(order, list) or len(order) != 24:
        fail("protocol.order: expected 24 lanes")
    lanes = [obj(item, f"protocol.order[{index}]") for index, item in enumerate(order[first:last])]
    expected = {("normal", shape, repeat) for shape in SHAPES} | {("allocator", shape, repeat) for shape in SHAPES}
    actual = {(lane.get("mode"), lane.get("shape"), lane.get("repeat")) for lane in lanes}
    if actual != expected or any(lane.get("role") != role or lane.get("phase") != phase for lane in lanes):
        fail(f"protocol.order: {phase} is not the declared six-lane matrix")
    return lanes


def verify_report_identity(path: Path, binary: dict[str, Any], revision: str) -> None:
    report = obj(load(path), str(path))
    identity = obj(report.get("binary_identity"), f"{path}.binary_identity")
    if identity.get("binary_sha256") != binary.get("sha256") or identity.get("binary_bytes") != binary.get("bytes") or identity.get("path") != binary.get("path"):
        fail(f"{path}: report binary identity differs")
    if identity.get("profile") != "release" or identity.get("executable") is not True:
        fail(f"{path}: report executable flags differ")
    environment = obj(report.get("environment"), f"{path}.environment")
    if environment.get("git_revision") != revision:
        fail(f"{path}: report revision differs")


def run_phase(phase: str, attempt: str, repo: Path, protocol_path: Path, custody_path: Path) -> int:
    protocol = obj(load(protocol_path), str(protocol_path))
    if protocol.get("change") != 442 or protocol.get("status") != "frozen":
        fail("protocol must be the frozen change-0442 protocol")
    oracle_files(protocol)
    role, _, _, _ = PHASES[phase]
    lanes = expected_lanes(protocol, phase)
    build_dir, build = build_for(role, protocol, repo, protocol_path)
    custody = import_custody(custody_path)
    source_before = custody.sources()
    if source_before != build.get("source_manifest"):
        fail(f"{phase}: source manifest differs from build")
    status_before = status_outside_bundle(repo)
    directory = ROOT / "runs" / phase / attempt
    if directory.exists() and any(directory.iterdir()):
        fail(f"{directory}: evidence already exists")
    directory.mkdir(parents=True, exist_ok=True)
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1 or protocol.get("samples") != 30 or protocol.get("warmups") != 3 or protocol.get("shapes") != SHAPES:
        fail("protocol: CPU, worker, sample, warmup, or shape binding differs")
    index: list[str] = []
    failed = False
    for lane in lanes:
        mode, shape = text(lane.get("mode"), "lane.mode"), text(lane.get("shape"), "lane.shape")
        name = f"{phase}-{role}-{mode}-{shape}-{lane['repeat'].lower()}"
        paths = {
            "report": directory / f"{name}.json", "catalog": directory / f"{name}-catalog.json",
            "workload_log": directory / f"{name}.log", "resource_log": directory / f"{name}-resource.log",
            "oracle_log": directory / f"{name}-oracle.log", "receipt": directory / f"{name}-receipt.json",
        }
        if any(path.exists() for path in paths.values()):
            fail(f"{name}: artifact already exists")
        role_spec = obj(protocol["roles"][role], f"protocol.roles.{role}")
        binary_desc = obj(build["binaries"][mode], f"{build_dir}/build.json.binaries.{mode}")
        binary = resolve_repo(binary_desc.get("path"), repo, f"{name}.binary")
        argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(paths["resource_log"])] + workload_argv(protocol, binary, role_spec["selector"], shape, paths["report"], paths["catalog"])
        row: dict[str, Any] = {
            "schema": "litchi-0442-capture-receipt-v1", "change": 442, "phase": phase, "attempt": attempt,
            "role": role, "build_directory": build_dir.name, "selector": role_spec["selector"],
            "source_field": role_spec["source_field"], "name": name, "lane": lane, "argv": argv,
            "cwd": str(repo), "revision": build["revision"], "source_manifest": build["source_manifest"],
            "binary": binary_desc, "protocol_sha256": sha(protocol_path), "driver_sha256": sha(Path(__file__)),
            "source_before": source_before, "status_before": status_before, "started_utc": now(), "status": "running",
        }
        write(paths["receipt"], row)
        try:
            env = os.environ.copy()
            env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": "", "PYTHONDONTWRITEBYTECODE": "1"})
            with paths["workload_log"].open("xb") as output:
                result = subprocess.run(argv, cwd=repo, env=env, stdout=output, stderr=subprocess.STDOUT)
            row["exit_code"] = result.returncode
            if result.returncode != 0:
                raise RuntimeError(f"workload exited {result.returncode}")
            verify_report_identity(paths["report"], binary_desc, build["revision"])
            oracle_argv = oracle_command(protocol, paths["report"], mode, shape, role)
            oracle_result = subprocess.run(oracle_argv, cwd=repo, env=env, capture_output=True, text=True)
            paths["oracle_log"].write_text(
                "argv=" + json.dumps(oracle_argv) + "\nstdout=" + oracle_result.stdout + "stderr=" + oracle_result.stderr,
                encoding="utf-8",
            )
            row["oracle_argv"] = oracle_argv
            row["oracle_exit_code"] = oracle_result.returncode
            if oracle_result.returncode != 0 or oracle_result.stdout.strip() != "VALID":
                raise RuntimeError("append oracle rejected report")
            row["status"] = "pass"
        except Exception as error:
            row["status"], row["error"], failed = "failed", repr(error), True
        finally:
            row["finished_utc"] = now()
            row["source_after"] = custody.sources()
            row["status_after"] = status_outside_bundle(repo)
            row["source_unchanged"] = row["source_before"] == row["source_after"]
            row["outside_bundle_status_unchanged"] = row["status_before"] == row["status_after"]
            if not row["source_unchanged"] or not row["outside_bundle_status_unchanged"]:
                row["status"] = "failed"
                row["error"] = "source custody or outside-bundle worktree status changed"
                failed = True
            row["artifacts"] = {key: record(path) for key, path in paths.items() if key != "receipt" and path.is_file()}
            write(paths["receipt"], row)
        index.append(str(paths["receipt"].relative_to(ROOT)))
        if failed:
            break
    source_after = custody.sources()
    status_after = status_outside_bundle(repo)
    if source_after != source_before or status_after != status_before:
        failed = True
    state = {"schema": "litchi-0442-capture-state-v1", "change": 442, "phase": phase, "role": role, "attempt": attempt, "status": "failed" if failed else "pass", "completed_lanes": len(index), "index": index, "protocol_sha256": sha(protocol_path), "driver_sha256": sha(Path(__file__)), "source_before": source_before, "source_after": source_after, "status_before": status_before, "status_after": status_after, "finished_utc": now()}
    write(directory / "capture-index.json", index)
    write(directory / "capture-state.json", state)
    return 1 if failed or len(index) != 6 else 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--phase", choices=tuple(PHASES), required=True)
    parser.add_argument("--attempt", default="formal")
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol.json")
    parser.add_argument("--custody-driver", type=Path, required=True)
    args = parser.parse_args()
    if ATTEMPT_RE.fullmatch(args.attempt) is None:
        parser.error("invalid attempt")
    try:
        return run_phase(args.phase, args.attempt, args.repo_root.resolve(), args.protocol.resolve(), args.custody_driver.resolve())
    except (OSError, KeyError, TypeError, ValueError, subprocess.CalledProcessError, CaptureError) as error:
        print(f"CAPTURE INVALID: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
