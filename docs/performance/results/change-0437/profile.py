#!/usr/bin/env python3
"""Capture one 0437 whole-process ODP profile.

The formal profile set is six fresh processes: stat and record for each of
the three ODP roles, using the large 8,192-slide normal workload.  A separate
``--preparatory`` invocation may profile the before-buffered executable before
the final descriptors exist.  Report correctness belongs to the copied ODP
oracle; this driver records the complete process and artifact custody only.
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
KINDS = ("stat", "record")
ROLES = ("before-buffered", "after-buffered", "after-streaming")
EVENTS = ("cycles:u", "instructions:u", "branches:u", "branch-misses:u", "L1-dcache-load-misses:u")
ATTEMPT_RE = re.compile(r"^[a-z0-9][a-z0-9_-]{0,63}$")


class ProfileError(ValueError):
    pass


def fail(message: str) -> None:
    raise ProfileError(message)


def load(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def write(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def record(path: Path) -> dict[str, Any]:
    return {"path": str(path.relative_to(ROOT)), "bytes": path.stat().st_size, "sha256": sha(path)}


def import_capture():
    spec = importlib.util.spec_from_file_location("change0437_capture_for_profile", ROOT / "capture.py")
    if spec is None or spec.loader is None:
        fail("cannot load capture helper")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


capture = import_capture()


def build_for(role: str, protocol: dict[str, Any], repo: Path,
              protocol_path: Path) -> tuple[Path, dict[str, Any]]:
    return capture.build_for(role, protocol, True, repo, protocol_path)


def retained_file(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value:
        fail(f"{label}: expected a retained bundle-relative path")
    path = Path(value)
    if path.is_absolute():
        fail(f"{label}: path must be bundle-relative")
    resolved = (ROOT / path).resolve()
    if not resolved.is_relative_to(ROOT.resolve()) or not resolved.is_file():
        fail(f"{label}: retained file is missing or escapes the evidence root")
    return resolved


def checked_source_manifest(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label}: expected a source manifest object")
    path = retained_file(value.get("path"), f"{label}.path")
    expected = value.get("sha256")
    if not isinstance(expected, str) or re.fullmatch(r"[0-9a-fA-F]{64}", expected) is None:
        fail(f"{label}.sha256: expected SHA-256")
    files = value.get("files")
    if isinstance(files, bool) or not isinstance(files, int) or files < 1:
        fail(f"{label}.files: expected a positive count")
    if sha(path) != expected.lower():
        fail(f"{label}: manifest hash differs from the receipt")
    manifest = load(path)
    if not isinstance(manifest, dict) or len(manifest) != files:
        fail(f"{label}: manifest contents do not match its file count")
    return {"path": str(path.relative_to(ROOT)), "sha256": expected.lower(), "files": files}


def checked_build_receipt() -> tuple[Path, dict[str, Any], dict[str, Any]]:
    path = ROOT / "checks" / "before-build.json"
    receipt = load(path)
    if not isinstance(receipt, dict):
        fail(f"{path}: expected an object")
    if receipt.get("change") != 437 or receipt.get("status") != "pass" or receipt.get("exit_code") != 0:
        fail(f"{path}: before build must have change=437, status=pass, exit_code=0")
    if receipt.get("source_unchanged") is not True:
        fail(f"{path}: before build does not prove source custody")
    before = checked_source_manifest(receipt.get("source_before"), f"{path}.source_before")
    after = checked_source_manifest(receipt.get("source_after"), f"{path}.source_after")
    if before != after:
        fail(f"{path}: source manifest changed during the build")
    revision = receipt.get("revision")
    if not isinstance(revision, str) or re.fullmatch(r"[0-9a-fA-F]{40}", revision) is None:
        fail(f"{path}.revision: expected a Git SHA-1")
    driver_sha = receipt.get("driver_sha256")
    if not isinstance(driver_sha, str) or re.fullmatch(r"[0-9a-fA-F]{64}", driver_sha) is None:
        fail(f"{path}.driver_sha256: expected SHA-256")
    check_driver = ROOT / "check.py"
    if not check_driver.is_file() or sha(check_driver) != driver_sha.lower():
        fail(f"{path}.driver_sha256: retained custody driver hash differs")
    log = receipt.get("log")
    if not isinstance(log, dict) or not isinstance(log.get("path"), str):
        fail(f"{path}.log: build log binding is required")
    log_path = retained_file(log["path"], f"{path}.log.path")
    if log.get("bytes") != log_path.stat().st_size or log.get("sha256") != sha(log_path):
        fail(f"{path}.log: retained build log hash differs")
    return path, receipt, before


def checked_binary_copies() -> tuple[Path, dict[str, dict[str, Any]]]:
    path = ROOT / "before" / "binary-copies.json"
    copies = load(path)
    if not isinstance(copies, dict) or set(copies) != {"normal", "allocator"}:
        fail(f"{path}: normal and allocator copies are required")
    checked: dict[str, dict[str, Any]] = {}
    for mode in ("normal", "allocator"):
        identity = copies[mode]
        if not isinstance(identity, dict):
            fail(f"{path}.{mode}: expected an object")
        binary_value = identity.get("path")
        if not isinstance(binary_value, str) or not binary_value:
            fail(f"{path}.{mode}.path: expected an absolute executable path")
        binary = Path(binary_value)
        if not binary.is_absolute():
            fail(f"{path}.{mode}.path: expected an absolute executable path")
        binary = binary.resolve()
        size = identity.get("bytes")
        expected = identity.get("sha256")
        if isinstance(size, bool) or not isinstance(size, int) or size <= 0:
            fail(f"{path}.{mode}.bytes: expected a positive count")
        if not isinstance(expected, str) or re.fullmatch(r"[0-9a-fA-F]{64}", expected) is None:
            fail(f"{path}.{mode}.sha256: expected SHA-256")
        if not binary.is_file() or binary.stat().st_size != size or sha(binary) != expected.lower():
            fail(f"{path}.{mode}: copied executable is stale or missing")
        checked[mode] = {"path": str(binary), "bytes": size, "sha256": expected.lower()}
    return path, checked


def preparatory_inputs(protocol_path: Path, protocol: dict[str, Any]) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    """Bind the raw before build before formal descriptors exist."""

    if protocol.get("status") != "draft":
        fail("preparatory profiles require a draft protocol")
    build_path, raw_build, source_manifest = checked_build_receipt()
    copies_path, binaries = checked_binary_copies()
    build = {
        "change": 437,
        "role": "before-buffered",
        "revision": raw_build["revision"],
        "source_manifest": source_manifest,
        "binaries": binaries,
        "build_receipt": str(build_path.relative_to(ROOT)),
        "build_receipt_sha256": sha(build_path),
        "binary_copies_receipt": str(copies_path.relative_to(ROOT)),
        "binary_copies_receipt_sha256": sha(copies_path),
        "check_driver_sha256": raw_build["driver_sha256"],
        "source_scope": raw_build.get("source_scope"),
        "preparatory_raw_build": True,
    }
    return ROOT / "before", build, {
        "build_path": build_path,
        "copies_path": copies_path,
        "source_manifest": source_manifest,
        "raw_build": raw_build,
        "protocol_sha256": sha(protocol_path),
    }


def oracle_binding(protocol: dict[str, Any], preparatory: bool) -> tuple[Path, str]:
    oracle = protocol.get("oracle")
    if preparatory and oracle is None:
        verifier = ROOT / "verify-report.py"
    else:
        if not isinstance(oracle, dict):
            fail("protocol.oracle: frozen profile requires an oracle object")
        verifier_value = oracle.get("verifier_path", oracle.get("path"))
        if not isinstance(verifier_value, str) or not verifier_value:
            fail("protocol.oracle.verifier_path: expected a copied verifier path")
        verifier = Path(verifier_value)
        verifier = verifier if verifier.is_absolute() else ROOT / verifier
        verifier = verifier.resolve()
    if not verifier.is_file():
        fail(f"{verifier}: copied ODP oracle verifier is missing")
    return verifier, sha(verifier)


def oracle_command(protocol: dict[str, Any], report: Path, role: str,
                   verifier: Path, preparatory: bool) -> list[str]:
    if preparatory and protocol.get("oracle") is None:
        return [sys.executable, "-B", str(verifier), "--report", str(report),
                "--mode", "normal", "--shape", "large", "--role", role,
                "--samples", str(protocol["samples"]), "--warmups", str(protocol["warmups"])]
    return capture.oracle_command(protocol, report, "normal", "large", role)


def run_profile(role: str, kind: str, preparatory: bool, attempt: str, repo: Path,
                protocol_path: Path, custody_path: Path) -> int:
    if role not in ROLES or kind not in KINDS:
        fail("role or profile kind is invalid")
    protocol = load(protocol_path)
    if protocol.get("change") != 437:
        fail("protocol.change must be 437")
    if protocol.get("shapes") != capture.SHAPES:
        fail("protocol.shapes must bind tiny=64, medium=4096, large=8192")
    if preparatory and role != "before-buffered":
        fail("preparatory profiles are only supported for before-buffered")
    if preparatory:
        build_dir, build, raw_bindings = preparatory_inputs(protocol_path, protocol)
        profile_root = ROOT / "profiles" / "preparatory-before-buffered" / attempt
        ambient_dir = build_dir.name
        ambient_source = build.get("source_manifest")
    else:
        build_role = role
        build_dir, build = build_for(build_role, protocol, repo, protocol_path)
        profile_root = ROOT / "profiles" / role
        _, ambient_build = build_for("after-streaming", protocol, repo, protocol_path)
        ambient_dir = "after"
        ambient_source = ambient_build.get("source_manifest")
        raw_bindings = {
            "build_path": ROOT / build["build_receipt"],
            "copies_path": ROOT / build["binary_copies_receipt"],
            "source_manifest": build.get("source_manifest"),
            "protocol_sha256": sha(protocol_path),
        }
    custody_spec = importlib.util.spec_from_file_location("change0437_profile_custody", custody_path)
    if custody_spec is None or custody_spec.loader is None:
        fail(f"cannot load custody driver {custody_path}")
    custody = importlib.util.module_from_spec(custody_spec)
    custody_spec.loader.exec_module(custody)
    source_before = custody.sources()
    if source_before != ambient_source:
        fail(f"ambient source differs from {role} build")
    profile_dir = profile_root / kind
    if profile_dir.exists() and any(profile_dir.iterdir()):
        fail(f"profile directory already contains evidence: {profile_dir}")
    profile_dir.mkdir(parents=True, exist_ok=True)
    report = profile_dir / "report.json"
    catalog = profile_dir / "report-catalog.json"
    resource = profile_dir / "resource.log"
    workload_log = profile_dir / "workload.log"
    stat_output = profile_dir / "perf-stat.txt"
    data = profile_dir / "perf.data"
    script_output = profile_dir / "perf-script.txt"
    perf_report = profile_dir / "perf-report.txt"
    oracle_log = profile_dir / "oracle.log"
    receipt_path = profile_dir / "receipt.json"
    binary_desc = build["binaries"]["normal"]
    binary_path = capture.resolve_repo(binary_desc["path"], repo, "profile binary")
    if not binary_path.is_file() or binary_path.stat().st_size != binary_desc["bytes"] or sha(binary_path) != binary_desc["sha256"]:
        fail("normal binary identity is stale")
    role_spec = protocol["roles"][role]
    selector = role_spec["selector"]
    verifier, verifier_sha = oracle_binding(protocol, preparatory)
    if not preparatory:
        recorded_oracle = build.get("oracle")
        if not isinstance(recorded_oracle, dict):
            fail("formal build descriptor has no oracle binding")
        recorded_sha = recorded_oracle.get("verifier_sha256", recorded_oracle.get("sha256"))
        if recorded_sha != verifier_sha:
            fail("formal build descriptor oracle hash differs from the frozen verifier")
    base = capture.workload_argv(protocol, binary_path, selector, "large", "normal", report, catalog)
    if kind == "stat":
        profiler = ["perf", "stat", "--no-big-num", "-x,", "-e", ",".join(EVENTS), "-o", str(stat_output), "--"] + base
    else:
        profiler = ["perf", "record", "--no-buildid-cache", "-o", str(data), "-F", "999", "-e", "cycles:u", "--call-graph", "fp,127", "--"] + base
    argv = ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(resource)] + profiler
    row: dict[str, Any] = {
        "schema": "litchi-0437-profile-receipt-v1", "change": 437, "role": role, "kind": kind,
        "attempt": attempt, "preparatory": preparatory, "scope": "whole fresh process including setup, corpus generation, warmups, samples, output hashing, and oracle",
        "selector": selector, "source_field": role_spec.get("source_field"), "shape": "large", "argv": argv,
        "cwd": str(repo), "revision": build.get("revision"), "source_manifest": build.get("source_manifest"),
        "ambient_source_manifest": ambient_source, "ambient_build_directory": ambient_dir,
        "binary": binary_desc, "protocol_path": str(protocol_path.relative_to(ROOT)), "protocol_sha256": sha(protocol_path),
        "oracle_verifier": str(verifier.relative_to(ROOT)) if verifier.is_relative_to(ROOT) else str(verifier),
        "oracle_verifier_sha256": verifier_sha,
        "build_receipt": raw_bindings["build_path"].relative_to(ROOT).as_posix(),
        "build_receipt_sha256": sha(raw_bindings["build_path"]),
        "binary_copies_receipt": raw_bindings["copies_path"].relative_to(ROOT).as_posix(),
        "binary_copies_receipt_sha256": sha(raw_bindings["copies_path"]),
        "driver_sha256": sha(Path(__file__)), "capture_helper_sha256": sha(ROOT / "capture.py"),
        "custody_driver": str(custody_path), "source_before": source_before,
        "stat_events": list(EVENTS) if kind == "stat" else [], "record_event": "cycles:u" if kind == "record" else None,
        "record_frequency_hz": 999 if kind == "record" else None, "call_graph": "fp,127" if kind == "record" else None,
        "started_utc": now(), "status": "running",
    }
    write(receipt_path, row)
    env = os.environ.copy()
    env.update({"RUSTUP_TOOLCHAIN": "1.98.1", "DEBUGINFOD_URLS": "", "PYTHONDONTWRITEBYTECODE": "1"})
    try:
        with workload_log.open("xb") as output:
            result = subprocess.run(argv, cwd=repo, env=env, stdout=output, stderr=subprocess.STDOUT)
        row["exit_code"] = result.returncode
        if result.returncode != 0:
            raise RuntimeError(f"profiler workload exited {result.returncode}")
        oracle_argv = oracle_command(protocol, report, role, verifier, preparatory)
        oracle_result = subprocess.run(oracle_argv, cwd=repo, env=env, capture_output=True, text=True)
        oracle_log.write_text(
            "argv=" + json.dumps(oracle_argv) + "\nstdout=" + oracle_result.stdout + "stderr=" + oracle_result.stderr,
            encoding="utf-8",
        )
        row["oracle_exit_code"] = oracle_result.returncode
        row["oracle_stdout_sha256"] = sha(oracle_log)
        if oracle_result.returncode != 0 or oracle_result.stdout.strip() != "VALID":
            raise RuntimeError("copied ODP oracle rejected profile report")
        if kind == "record":
            for command, output in (
                (["perf", "report", "--stdio", "--no-children", "--percent-limit", "0", "-i", str(data)], perf_report),
                (["perf", "script", "-i", str(data)], script_output),
            ):
                with output.open("xb") as stream:
                    subprocess.run(command, cwd=repo, env=env, stdout=stream, stderr=subprocess.STDOUT, check=True)
        row["status"] = "pass"
    except Exception as error:
        row["status"] = "failed"
        row["error"] = repr(error)
    finally:
        row["source_after"] = custody.sources()
        row["source_unchanged"] = row["source_before"] == row["source_after"]
        if not row["source_unchanged"]:
            row["status"] = "failed"
            row["error"] = "source custody changed during profile"
        row["finished_utc"] = now()
        row["artifacts"] = {
            key: record(path) for key, path in {
                "report": report, "catalog": catalog, "resource": resource, "workload_log": workload_log,
                "oracle_log": oracle_log, "perf_stat": stat_output, "perf_data": data,
                "perf_script": script_output, "perf_report": perf_report,
            }.items() if path.is_file()
        }
        write(receipt_path, row)
    print(json.dumps({"status": row["status"], "role": role, "kind": kind}))
    return 0 if row["status"] == "pass" else 1


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--role", choices=ROLES, required=True)
    parser.add_argument("--kind", choices=KINDS, required=True)
    parser.add_argument("--preparatory", action="store_true")
    parser.add_argument("--attempt", default="preparatory")
    parser.add_argument("--repo-root", type=Path, required=True)
    parser.add_argument("--protocol", type=Path, default=ROOT / "protocol.json")
    parser.add_argument("--custody-driver", type=Path, required=True)
    args = parser.parse_args()
    if ATTEMPT_RE.fullmatch(args.attempt) is None:
        parser.error("--attempt must match [a-z0-9][a-z0-9_-]{0,63}")
    if not args.preparatory and args.attempt != "formal":
        parser.error("formal profiles require --attempt formal")
    try:
        return run_profile(
            args.role, args.kind, args.preparatory, args.attempt,
            args.repo_root.resolve(), args.protocol.resolve(), args.custody_driver.resolve(),
        )
    except (OSError, KeyError, TypeError, ValueError, subprocess.CalledProcessError, ProfileError) as error:
        print(f"PROFILE INVALID: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
