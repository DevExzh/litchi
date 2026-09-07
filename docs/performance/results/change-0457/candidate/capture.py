#!/usr/bin/env python3
"""Capture one serialized 0457 source-backed ODP tail-append candidate phase.

The driver executes only the retained candidate binaries bound after the candidate build.  It records the
ambient allocator environment without changing it, retains every failure
receipt, and delegates report semantics to the independent source-tail candidate oracle.
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
PHASES = {"R1": (0, 6), "R2": (6, 12)}
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


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected non-empty text")
    return value


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected object")
    return value


def import_custody(path: Path):
    spec = importlib.util.spec_from_file_location("change0457_custody", path)
    if spec is None or spec.loader is None:
        fail(f"cannot load custody driver {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    if not callable(getattr(module, "sources", None)):
        fail(f"custody driver {path} has no sources()")
    return module


def allocator_environment() -> dict[str, str]:
    return {
        key: value
        for key, value in sorted(os.environ.items())
        if key.startswith("MALLOC_") or key == "GLIBC_TUNABLES"
    }


def status_outside_bundle(repo: Path) -> list[str]:
    raw = subprocess.check_output(
        ["git", "status", "--short", "--untracked-files=all"], cwd=repo, text=True
    )
    prefix = ROOT.resolve().relative_to(repo.resolve()).as_posix() + "/"
    return [line for line in raw.splitlines() if not line[3:].startswith(prefix)]


def binding(protocol: dict[str, Any], repo: Path) -> tuple[dict[str, Any], dict[str, Any], Any, dict[str, Any], str, str]:
    """Authenticate candidate artifacts without guessing the post-build hash.

    The candidate binary is bound by the retained binary-bindings receipt after
    the candidate build.  The frozen protocol binds the paths and this driver;
    the binding file binds each copied/source executable and build receipt.
    """
    if protocol.get("schema") != "litchi-0457-odp-source-tail-candidate-v1":
        fail("protocol schema is not the 0457 source-tail candidate")
    if protocol.get("capture_driver_sha256") != sha(Path(__file__)):
        fail("capture driver differs from protocol")
    if protocol.get("change") != 457:
        fail("protocol change is not 457")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("candidate requires CPU 2 and one worker")
    if protocol.get("samples") != 30 or protocol.get("warmups") != 3:
        fail("candidate requires 30 samples and three warmups")
    if protocol.get("selector") != "odp_source_tail_append_lifecycle":
        fail("selector differs from the source-tail candidate")
    if protocol.get("shapes") != {"tiny": 64, "medium": 4096, "large": 8192}:
        fail("shape matrix differs from the source-tail candidate")
    role = obj(protocol["roles"]["candidate"], "protocol.roles.candidate")
    build_receipt_path = ROOT / text(role["build_receipt"]["path"], "build receipt path")
    binary_file = ROOT / text(role["binary_binding"]["path"], "binary binding path")
    source_file = ROOT / text(role["source_manifest"]["path"], "source manifest path")
    build = obj(load(build_receipt_path), "candidate build receipt")
    if build.get("change") != 457 or build.get("status") != "pass" or build.get("source_unchanged") is not True:
        fail("candidate build receipt is not a successful source-unchanged build")
    revision = text(build.get("revision"), "candidate build revision")
    source_map = load(source_file)
    if not isinstance(source_map, dict) or not source_map:
        fail("candidate source manifest must be a non-empty file map")
    source_digest = sha(source_file)
    source_record = {
        "path": f"sources/{source_digest}.json",
        "sha256": source_digest,
        "files": len(source_map),
    }
    for key in ("source_before", "source_after"):
        if build.get(key) != source_record:
            fail(f"candidate build receipt {key} does not bind source-manifest.json")
    expected_source = source_record
    binaries = obj(load(binary_file), "candidate binary bindings")
    if binaries.get("schema") != "litchi-0457-source-tail-candidate-binary-binding-v1":
        fail("candidate binary binding schema differs")
    expected_build_record = {
        "path": str(build_receipt_path.relative_to(ROOT)),
        "sha256": sha(build_receipt_path),
    }
    expected_source_record = {
        "path": str(source_file.relative_to(ROOT)),
        "sha256": source_digest,
        "files": len(source_map),
    }
    if binaries.get("build_receipt") != expected_build_record:
        fail("candidate binary bindings do not bind build-receipt.json")
    if binaries.get("source_manifest") != expected_source_record:
        fail("candidate binary bindings do not bind source-manifest.json")
    for mode in ("normal", "allocator"):
        item = obj(obj(binaries["binaries"], "binary bindings.binaries")[mode], f"binary bindings.binaries.{mode}")
        if item.get("profile") != "release" or item.get("executable") is not True:
            fail(f"{mode} candidate binding is not a release executable")
        copy_path = Path(text(item["copy_path"], f"{mode} copy path"))
        source_path = repo / text(item["source_path"], f"{mode} source path")
        for path, label in ((copy_path, f"{mode} copied binary"), (source_path, f"{mode} source binary")):
            if not path.is_file() or path.is_symlink() or path.stat().st_size != item["bytes"] or sha(path) != item["sha256"]:
                fail(f"{label} identity differs from binary binding")
    oracle = obj(protocol["oracle"], "protocol.oracle")
    for key in ("verifier_path", "protocol_path", "corpus_bindings_path"):
        path = ROOT / text(oracle[key], f"protocol.oracle.{key}")
        hash_key = {"verifier_path": "verifier_sha256", "protocol_path": "protocol_sha256", "corpus_bindings_path": "corpus_bindings_sha256"}[key]
        if not path.is_file() or sha(path) != oracle[hash_key]:
            fail(f"candidate oracle {key} differs from protocol")
    output_binding = obj(protocol.get("output_binding"), "protocol.output_binding")
    output_binding_path = ROOT / text(output_binding.get("path"), "protocol.output_binding.path")
    if not output_binding_path.is_file() or output_binding_path.is_symlink():
        fail("independently verified candidate output binding is missing")
    output_binding_sha = sha(output_binding_path)
    expected_output_binding_sha = output_binding.get("sha256")
    if expected_output_binding_sha not in ("pending-output-preflight", "bound-at-capture") and output_binding_sha != expected_output_binding_sha:
        fail("candidate output binding differs from protocol")
    output_binder_path = ROOT / text(output_binding.get("binder_path"), "protocol.output_binding.binder_path")
    expected_binder_sha = output_binding.get("binder_sha256")
    if not output_binder_path.is_file() or output_binder_path.is_symlink() or (expected_binder_sha not in ("pending-output-preflight", "bound-at-capture") and sha(output_binder_path) != text(expected_binder_sha, "protocol.output_binding.binder_sha256")):
        fail("candidate output binding helper differs from protocol")
    native_output = obj(output_binding.get("native_oracle"), "protocol.output_binding.native_oracle")
    native_output_path = (ROOT / text(native_output.get("path"), "protocol.output_binding.native_oracle.path")).resolve()
    if not native_output_path.is_file() or native_output_path.is_symlink() or sha(native_output_path) != text(native_output.get("sha256"), "protocol.output_binding.native_oracle.sha256"):
        fail("native output oracle differs from protocol output binding")
    custody_path = (ROOT / text(protocol["custody"]["driver_path"], "custody driver path")).resolve()
    if sha(custody_path) != protocol["custody"]["driver_sha256"]:
        fail("parent custody driver differs from protocol")
    ambient = obj(protocol["environment"], "protocol.environment")
    if allocator_environment() != ambient.get("allocator_ambient", {}):
        fail("ambient allocator environment differs from the frozen empty observation")
    if os.environ.get("RUSTFLAGS") not in (None, ""):
        fail("RUSTFLAGS must be unset or empty for the null report field")
    role = dict(role)
    role["revision"] = revision
    return role, binaries, import_custody(custody_path), expected_source, sha(binary_file), output_binding_sha

def expected_lanes(protocol: dict[str, Any], phase: str) -> list[dict[str, Any]]:
    order = protocol.get("order")
    if not isinstance(order, list) or len(order) != 12:
        fail("protocol order must contain 12 lanes")
    first, last = PHASES[phase]
    lanes = [obj(item, f"protocol.order[{index}]") for index, item in enumerate(order[first:last], first)]
    expected = {
        (mode, shape, phase)
        for mode in ("normal", "allocator")
        for shape in ("tiny", "medium", "large")
    }
    actual = {(lane.get("mode"), lane.get("shape"), lane.get("repeat")) for lane in lanes}
    if actual != expected or any(lane.get("phase") != phase for lane in lanes):
        fail(f"{phase} is not the frozen six-lane matrix")
    expected_order = (
        [("normal", "tiny"), ("normal", "medium"), ("normal", "large"),
         ("allocator", "tiny"), ("allocator", "medium"), ("allocator", "large")]
        if phase == "R1" else
        [("allocator", "large"), ("allocator", "medium"), ("allocator", "tiny"),
         ("normal", "large"), ("normal", "medium"), ("normal", "tiny")]
    )
    if [(lane["mode"], lane["shape"]) for lane in lanes] != expected_order:
        fail(f"{phase} order is not forward/reverse")
    return lanes


def oracle_command(protocol: dict[str, Any], report: Path, mode: str, shape: str) -> list[str]:
    oracle = obj(protocol["oracle"], "protocol.oracle")
    verifier = ROOT / text(oracle["verifier_path"], "protocol.oracle.verifier_path")
    template = oracle["argv"]
    values = {"python": sys.executable, "verifier": str(verifier), "report": str(report), "mode": mode, "shape": shape}
    return [item.format(**values) for item in template]


def workload_command(protocol: dict[str, Any], binary: Path, lane: dict[str, Any], report: Path, catalog: Path, resource: Path) -> list[str]:
    cli = obj(protocol["workload_cli"], "protocol.workload_cli")
    return [
        "taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o", str(resource),
        str(binary), cli["case_flag"], protocol["selector"], cli["shape_flag"], lane["shape"],
        cli["workers_flag"], str(protocol["workers"]), cli["samples_flag"], str(protocol["samples"]),
        cli["warmup_flag"], str(protocol["warmups"]), cli["report_flag"], str(report),
        cli["catalog_flag"], str(catalog),
    ]


def verify_report_identity(report: Path, binary: dict[str, Any], revision: str, lane: dict[str, Any], protocol: dict[str, Any]) -> None:
    value = obj(load(report), str(report))
    identity = obj(value.get("binary_identity"), "report.binary_identity")
    if identity.get("path") != binary["copy_path"] or identity.get("binary_sha256") != binary["sha256"] or identity.get("binary_bytes") != binary["bytes"] or identity.get("profile") != "release" or identity.get("executable") is not True:
        fail("report executable identity differs from copied candidate binary")
    environment = obj(value.get("environment"), "report.environment")
    if environment.get("git_revision") != revision:
        fail("report revision differs from candidate build")
    configuration = obj(value.get("configuration"), "report.configuration")
    if configuration.get("samples_per_case") != protocol["samples"] or configuration.get("warmup_iterations_per_case") != protocol["warmups"] or configuration.get("execution_workers") != [protocol["workers"]]:
        fail("report sample or worker configuration differs from control")
    results = value.get("results")
    if not isinstance(results, list) or len(results) != 1 or obj(results[0], "report.results[0]").get("case") != protocol["selector"] or obj(results[0], "report.results[0]").get("corpus", {}).get("shape") != lane["shape"]:
        fail("report selector or shape differs from lane")


def run_phase(phase: str, attempt: str, repo: Path, protocol_path: Path, custody_path: Path) -> int:
    directory = ROOT / "runs" / phase / attempt
    if directory.exists() and any(directory.iterdir()):
        fail(f"capture directory already contains evidence: {directory}")
    directory.mkdir(parents=True, exist_ok=True)
    state_path = directory / "capture-state.json"
    state: dict[str, Any] = {"schema": "litchi-0457-source-tail-candidate-capture-state-v1", "change": 457, "phase": phase, "attempt": attempt, "status": "running", "driver_sha256": sha(Path(__file__))}
    write(state_path, state)
    try:
        protocol = obj(load(protocol_path), str(protocol_path))
        protocol_sha = sha(protocol_path)
        role, binaries, custody, expected_source, binary_binding_sha, output_binding_sha = binding(protocol, repo)
        lanes = expected_lanes(protocol, phase)
        source_before = custody.sources()
        if source_before != expected_source:
            fail("ambient source manifest differs from the bound baseline")
        status_before = status_outside_bundle(repo)
        allocator_before = allocator_environment()
        state.update(protocol_sha256=protocol_sha, binary_binding_sha256=binary_binding_sha, output_binding_sha256=output_binding_sha, source_before=source_before, allocator_environment=allocator_before, expected_lanes=len(lanes))
        write(state_path, state)
        env = os.environ.copy()
        env.update(protocol["environment"]["fixed"])
        env.pop("RUSTFLAGS", None)
        index: list[str] = []
        failed = False
        for lane in lanes:
            mode, shape = lane["mode"], lane["shape"]
            name = f"{phase}-{mode}-{shape}-{lane['repeat'].lower()}"
            report = directory / f"{name}.json"
            catalog = directory / f"{name}-catalog.json"
            workload_log = directory / f"{name}.log"
            resource_log = directory / f"{name}-resource.log"
            oracle_log = directory / f"{name}-oracle.log"
            receipt = directory / f"{name}-receipt.json"
            paths = [report, catalog, workload_log, resource_log, oracle_log, receipt]
            if any(path.exists() for path in paths):
                fail(f"lane already has an artifact: {name}")
            lane_source_before = custody.sources()
            if lane_source_before != expected_source:
                fail(f"source manifest changed before lane {name}")
            binary = obj(obj(binaries["binaries"], "binary bindings.binaries")[mode], f"binary bindings.binaries.{mode}")
            binary_path = Path(binary["copy_path"])
            argv = workload_command(protocol, binary_path, lane, report, catalog, resource_log)
            row: dict[str, Any] = {
                "schema": "litchi-0457-source-tail-candidate-capture-receipt-v1", "change": 457,
                "phase": phase, "attempt": attempt, "name": name, "lane": lane,
                "selector": protocol["selector"], "role": "candidate", "revision": role["revision"],
                "argv": argv, "cwd": str(repo), "binary": binary, "protocol_sha256": protocol_sha, "binary_binding_sha256": binary_binding_sha, "output_binding_sha256": output_binding_sha,
                "driver_sha256": sha(Path(__file__)), "custody_driver": str(custody_path),
                "custody_driver_sha256": protocol["custody"]["driver_sha256"], "source_before": lane_source_before,
                "allocator_environment": allocator_before, "started_utc": now(), "status": "running",
            }
            write(receipt, row)
            print(json.dumps({"status": "running", "phase": phase, "lane": name}), flush=True)
            try:
                with workload_log.open("xb") as output:
                    result = subprocess.run(argv, cwd=repo, env=env, stdout=output, stderr=subprocess.STDOUT)
                row["exit_code"] = result.returncode
                if result.returncode != 0:
                    raise RuntimeError(f"workload exited {result.returncode}")
                verify_report_identity(report, binary, role["revision"], lane, protocol)
                oracle_argv = oracle_command(protocol, report, mode, shape)
                oracle_result = subprocess.run(oracle_argv, cwd=repo, env=env, capture_output=True, text=True)
                oracle_log.write_text("argv=" + json.dumps(oracle_argv) + "\nstdout=" + oracle_result.stdout + "stderr=" + oracle_result.stderr, encoding="utf-8")
                row["oracle_argv"] = oracle_argv
                row["oracle_exit_code"] = oracle_result.returncode
                row["oracle_stdout_sha256"] = sha(oracle_log)
                if oracle_result.returncode != 0 or oracle_result.stdout.strip() != protocol["oracle"]["success_stdout"]:
                    raise RuntimeError("copied ODP oracle rejected the report")
                row["status"] = "pass"
            except Exception as error:
                row["status"] = "failed"
                row["error"] = repr(error)
                failed = True
            finally:
                row["finished_utc"] = now()
                row["source_after"] = custody.sources()
                row["source_unchanged"] = row["source_before"] == row["source_after"]
                row["allocator_environment_after"] = allocator_environment()
                if not row["source_unchanged"] and row["status"] == "pass":
                    row["status"] = "failed"
                    row["error"] = "source custody changed during lane"
                row["artifacts"] = {key: record(path) for key, path in {"report": report, "catalog": catalog, "workload_log": workload_log, "resource_log": resource_log, "oracle_log": oracle_log}.items() if path.is_file()}
                write(receipt, row)
                if not row["source_unchanged"]:
                    failed = True
            index.append(str(receipt.relative_to(ROOT)))
            if failed:
                break
        source_after = custody.sources()
        status_after = status_outside_bundle(repo)
        state.update(status="failed" if failed or source_after != source_before or status_after != status_before else "pass", completed_lanes=len(index), index=index, source_after=source_after, source_unchanged=source_after == source_before, status_before=status_before, status_after=status_after, outside_bundle_status_unchanged=status_after == status_before, allocator_environment_after=allocator_environment(), finished_utc=now())
        write(state_path, state)
        write(directory / "capture-index.json", index)
        return 1 if state["status"] != "pass" else 0
    except Exception as error:
        state.update(status="failed", error=repr(error), finished_utc=now())
        write(state_path, state)
        return 1


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
