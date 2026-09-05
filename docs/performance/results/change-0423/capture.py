#!/usr/bin/env python3
"""Capture the serialized 0423 same-revision lifecycle matrix.

Each row is a fresh process pinned to CPU 2.  The same copied normal and
allocator binaries are used for both public API roles; role selects the owned
or source-backed lifecycle case.  Reports, corpus catalogs, raw GNU
``time -v`` output, verifier output, and compact custody journals are retained
under ``runs/{lane}/{corpus}/{role}-{repeat}/``.  A failed or partial root is
never resumed or overwritten.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Any


CHANGE = 423
CPU = 2
WORKERS = 1
TOOLCHAIN = "1.98.1"
MODES = ("normal", "allocator")
CORPORA = ("plain", "media_rich")
REPEATS = ("R1", "R2")
ORDER = (
    ("owned", "R1"),
    ("source", "R1"),
    ("source", "R2"),
    ("owned", "R2"),
)
SELECTORS = (
    "pptx_cross_copy_plain_lifecycle",
    "pptx_source_backed_cross_copy_plain_lifecycle",
    "pptx_cross_copy_media_rich_lifecycle",
    "pptx_source_backed_cross_copy_media_rich_lifecycle",
)
SELECTOR_BY_ROLE_CORPUS = {
    ("owned", "plain"): "pptx_cross_copy_plain_lifecycle",
    ("source", "plain"): "pptx_source_backed_cross_copy_plain_lifecycle",
    ("owned", "media_rich"): "pptx_cross_copy_media_rich_lifecycle",
    ("source", "media_rich"): "pptx_source_backed_cross_copy_media_rich_lifecycle",
}
SAFE_NAME = re.compile(r"^[A-Za-z0-9_.-]+$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")
GIT_REVISION = re.compile(r"^[0-9a-f]{40}$")
CAPTURE_FLAGS = frozenset(
    {"--case", "--cases", "--json", "--output", "--corpus-manifest", "--samples", "--warmup", "--warmups"}
)
SOURCE_SPECS = (
    "Cargo.toml",
    "Cargo.lock",
    "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/Cargo.lock",
    "tools/perf-baseline/src",
    "crates/litchi-opc/Cargo.toml",
    "crates/litchi-opc/src",
    "crates/litchi-pptx/Cargo.toml",
    "crates/litchi-pptx/src",
)


class CaptureError(RuntimeError):
    """An input, identity, artifact, verifier, or subprocess error."""


def fail(message: str) -> None:
    raise CaptureError(message)


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON value {value!r}")


def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def load_json(path: Path, label: str) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=reject_duplicates,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"cannot load {label} {path}: {error}")


def write_json(path: Path, value: Any) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    try:
        temporary.write_text(
            json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False) + "\n",
            encoding="utf-8",
        )
        temporary.replace(path)
    except (OSError, TypeError, ValueError, OverflowError) as error:
        try:
            temporary.unlink()
        except OSError:
            pass
        fail(f"cannot write {path}: {error}")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def digest_json(value: Any) -> str:
    try:
        encoded = json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False).encode()
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot hash JSON identity: {error}")
    return hashlib.sha256(encoded).hexdigest()


def string_value(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label} must be a non-empty string")
    return value


def integer_value(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label} must be an integer >= {minimum}")
    return value


def object_value(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label} must be an object")
    return value


def relative_path(root: Path, value: Any, label: str) -> Path:
    raw = string_value(value, label)
    path = Path(raw)
    if path.is_absolute() or ".." in path.parts:
        fail(f"{label} must be relative and cannot contain '..'")
    resolved = (root / path).resolve()
    try:
        resolved.relative_to(root.resolve())
    except ValueError:
        fail(f"{label} escapes the capture root")
    return resolved


def source_identity(worktree: Path) -> dict[str, Any]:
    worktree = worktree.expanduser().resolve()
    if not worktree.is_dir():
        fail(f"source worktree is missing: {worktree}")

    def git(*arguments: str) -> str:
        try:
            process = subprocess.run(
                ["git", *arguments], cwd=worktree, capture_output=True,
                text=True, check=False,
            )
        except OSError as error:
            fail(f"cannot run git in {worktree}: {error}")
        if process.returncode != 0:
            fail(f"git {' '.join(arguments)} failed: {(process.stderr or process.stdout).strip()}")
        return process.stdout

    revision = git("rev-parse", "HEAD").strip()
    status = git("status", "--porcelain=v1", "--untracked-files=all")
    tracked = git("ls-files", "-z", "--", *SOURCE_SPECS)
    files: list[dict[str, Any]] = []
    for relative in sorted(item for item in tracked.split("\0") if item):
        path = worktree / relative
        if not path.is_file():
            fail(f"tracked source file is missing: {path}")
        files.append({"path": relative, "bytes": path.stat().st_size, "sha256": sha256_file(path)})
    if not files:
        fail(f"source binding matched no files in {worktree}")
    return {
        "worktree": str(worktree),
        "revision": revision,
        "git_status_porcelain": status,
        "clean": status == "",
        "source_files": files,
    }


def require_source_unchanged(source: dict[str, Any]) -> dict[str, Any]:
    observed = source_identity(Path(source["worktree"]))
    if observed != source:
        fail("source worktree or bound source files changed during capture")
    return observed


def require_baseline_ancestor(source: dict[str, Any], baseline: str) -> None:
    try:
        process = subprocess.run(
            ["git", "merge-base", "--is-ancestor", baseline, source["revision"]],
            cwd=source["worktree"], capture_output=True, text=True, check=False,
        )
    except OSError as error:
        fail(f"cannot check protocol baseline ancestry: {error}")
    if process.returncode != 0:
        fail(
            f"protocol baseline {baseline} is not an ancestor of source revision "
            f"{source['revision']}: {(process.stderr or process.stdout).strip()}"
        )


def validate_binary(value: Any, label: str) -> dict[str, Any]:
    binary = object_value(value, label)
    raw_path = string_value(binary.get("path"), f"{label}.path")
    path = Path(raw_path)
    if not path.is_absolute() or not raw_path.startswith("/tmp/"):
        fail(f"{label}.path must be an absolute copied /tmp binary")
    if not path.is_file() or not os.access(path, os.X_OK):
        fail(f"{label} is missing or not executable: {path}")
    observed_hash = sha256_file(path)
    expected_hash = string_value(binary.get("sha256", binary.get("binary_sha256")), f"{label}.sha256")
    alias_hash = binary.get("binary_sha256")
    if alias_hash is not None and alias_hash != expected_hash:
        fail(f"{label} sha256 and binary_sha256 disagree")
    if SHA256.fullmatch(expected_hash) is None or expected_hash != observed_hash:
        fail(f"{label} hash differs from its build identity")
    observed_bytes = path.stat().st_size
    expected_bytes = integer_value(binary.get("bytes", binary.get("binary_bytes")), f"{label}.bytes", 1)
    alias_bytes = binary.get("binary_bytes")
    if alias_bytes is not None and alias_bytes != expected_bytes:
        fail(f"{label} bytes and binary_bytes disagree")
    if observed_bytes != expected_bytes:
        fail(f"{label} size differs from its build identity")
    expected_mode_bits = binary.get("mode_bits")
    if expected_mode_bits is not None and expected_mode_bits != path.stat().st_mode & 0o7777:
        fail(f"{label} mode bits differ from its build identity")
    if binary.get("executable") is not None and binary.get("executable") is not True:
        fail(f"{label} build identity is not executable")
    return {
        "path": str(path),
        "sha256": observed_hash,
        "bytes": observed_bytes,
        "mode_bits": binary.get("mode_bits"),
        "executable": binary.get("executable", True),
        "profile": binary.get("profile"),
    }


def load_protocol(root: Path) -> dict[str, Any]:
    path = root / "protocol.json"
    protocol = object_value(load_json(path, "protocol"), str(path))
    if protocol.get("change") != CHANGE:
        fail("protocol.change must be 423")
    if protocol.get("cpu") != CPU or protocol.get("workers") != WORKERS:
        fail("protocol must pin CPU 2 and one worker")
    if protocol.get("corpora") != list(CORPORA):
        fail(f"protocol.corpora must be {list(CORPORA)!r}")
    if protocol.get("selectors") != list(SELECTORS):
        fail(f"protocol.selectors must be {list(SELECTORS)!r}")
    raw_lanes = object_value(protocol.get("lanes"), "protocol.lanes")
    lanes: dict[str, dict[str, int]] = {}
    for mode, expected in (("normal", (100, 10)), ("allocator", (30, 3))):
        value = object_value(raw_lanes.get(mode), f"protocol.lanes.{mode}")
        contract = {
            "samples": integer_value(value.get("samples"), f"lanes.{mode}.samples", 1),
            "warmups": integer_value(value.get("warmups"), f"lanes.{mode}.warmups", 0),
        }
        if (contract["samples"], contract["warmups"]) != expected:
            fail(f"protocol.lanes.{mode} must be {expected[0]}/{expected[1]}")
        lanes[mode] = contract
    if set(raw_lanes) != set(MODES):
        fail("protocol.lanes must contain only normal and allocator")
    raw_order = protocol.get("order")
    expected_order = [{"role": role, "repeat": repeat} for role, repeat in ORDER]
    if raw_order != expected_order:
        fail(f"protocol.order must be {expected_order!r}")
    if protocol.get("expected_run_count") != 16:
        fail("protocol.expected_run_count must be 16")
    flags = protocol.get("common_flags")
    if not isinstance(flags, list) or not flags or any(not isinstance(item, str) or not item for item in flags):
        fail("protocol.common_flags must be a non-empty string list")
    if any(item in CAPTURE_FLAGS for item in flags):
        fail("protocol.common_flags contains a capture-owned flag")
    workers_index = [index for index, item in enumerate(flags) if item == "--workers"]
    if len(workers_index) != 1:
        fail("protocol.common_flags must contain exactly one --workers")
    if workers_index[0] + 1 >= len(flags) or flags[workers_index[0] + 1] != "1":
        fail("protocol.common_flags must pin --workers 1")
    return {
        "path": "protocol.json",
        "sha256": sha256_file(path),
        "value": protocol,
        "common_flags": list(flags),
        "lanes": lanes,
        "corpora": list(CORPORA),
        "selectors": list(SELECTORS),
        "order": [list(row) for row in ORDER],
    }


def load_build(root: Path, protocol: dict[str, Any]) -> dict[str, Any]:
    path = root / "build-candidate.json"
    build = object_value(load_json(path, "candidate build"), str(path))
    if build.get("change") != CHANGE or build.get("role") != "candidate":
        fail("build-candidate.json has the wrong change or role")
    if build.get("status") != "pass" or build.get("exit_code") != 0:
        fail("build-candidate.json is not a successful build record")
    build_hash = string_value(build.get("protocol_sha256"), "build.protocol_sha256")
    if build_hash != protocol["sha256"]:
        fail("candidate build is not bound to the frozen protocol")
    source_before = object_value(build.get("source_before"), "build.source_before")
    source_after = object_value(build.get("source_after"), "build.source_after")
    if source_before != source_after or source_before.get("clean") is not True:
        fail("build source_before/source_after must be the same clean identity")
    revision = string_value(source_before.get("revision"), "build source revision")
    if GIT_REVISION.fullmatch(revision) is None:
        fail("build source revision must be a full lowercase git object")
    baseline = string_value(protocol["value"].get("baseline_revision"), "protocol.baseline_revision")
    require_baseline_ancestor(source_before, baseline)
    source = source_identity(Path(string_value(source_before.get("worktree"), "build source worktree")))
    if source != source_before:
        fail("source worktree differs from build source_before identity")
    binaries = object_value(build.get("binaries"), "build.binaries")
    checked = {mode: validate_binary(binaries.get(mode), f"build.binaries.{mode}") for mode in MODES}
    return {
        "path": "build-candidate.json",
        "sha256": sha256_file(path),
        "value": build,
        "source": source_before,
        "source_identity_sha256": digest_json(source_before),
        "binaries": checked,
    }


def artifact_record(path: Path, root: Path, *, allow_empty: bool = False) -> dict[str, Any]:
    if not path.is_file():
        fail(f"expected capture artifact is missing: {path}")
    size = path.stat().st_size
    if size == 0 and not allow_empty:
        fail(f"expected capture artifact is empty: {path}")
    try:
        relative = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        fail(f"capture artifact escapes root: {path}")
    return {"path": relative, "bytes": size, "sha256": sha256_file(path), "allow_empty": allow_empty}


def validate_report_identity(
    report: dict[str, Any], source: dict[str, Any], binary: dict[str, Any], label: str
) -> None:
    environment = object_value(report.get("environment"), f"{label}.environment")
    if environment.get("git_revision") != source["revision"]:
        fail(f"{label}.environment.git_revision does not match the build source")
    if environment.get("git_worktree_dirty") is not False:
        fail(f"{label}.environment.git_worktree_dirty must be false")
    identity = object_value(report.get("binary_identity"), f"{label}.binary_identity")
    if identity.get("binary_sha256") != binary["sha256"]:
        fail(f"{label}.binary_identity.binary_sha256 does not match the copied binary")
    if identity.get("binary_bytes") != binary["bytes"]:
        fail(f"{label}.binary_identity.binary_bytes does not match the copied binary")


def verify_output(
    path: Path,
    label: str,
    *,
    selector: str,
    lane: str,
    samples: int,
    warmups: int,
) -> dict[str, Any]:
    value = object_value(load_json(path, label), label)
    # This is deliberately an exact common single-report envelope.  Merely
    # accepting verifier exit 0 or an arbitrary JSON object would allow a
    # wrapper to discard the semantic gate while still looking successful.
    required = {
        "status", "claim_authorized", "performance_claim", "selector", "lane",
        "samples", "warmups", "report_count", "reports",
    }
    missing = sorted(required.difference(value))
    if missing:
        fail(f"{label} is missing required verifier fields: {missing}")
    if value["status"] != "pass":
        fail(f"{label}.status must be 'pass'")
    if value["claim_authorized"] is not False:
        fail(f"{label}.claim_authorized must be false")
    if value["performance_claim"] is not None:
        fail(f"{label}.performance_claim must be null")
    if value["selector"] != selector or value["lane"] != lane:
        fail(f"{label} selector/lane does not match the run")
    if value["samples"] != samples or value["warmups"] != warmups:
        fail(f"{label} samples/warmups do not match the run")
    if value["report_count"] != 1:
        fail(f"{label}.report_count must be 1")
    if not isinstance(value["reports"], list) or len(value["reports"]) != 1:
        fail(f"{label}.reports must contain exactly one report")
    return value


def run_one(
    *, root: Path, protocol: dict[str, Any], build: dict[str, Any], verifier: Path,
    repo_root: Path, lane: str, corpus: str, role: str, repeat: str,
    index: int, total: int,
) -> dict[str, Any]:
    selector = SELECTOR_BY_ROLE_CORPUS[(role, corpus)]
    folder = root / "runs" / lane / corpus / f"{role}-{repeat}"
    if folder.exists():
        fail(f"refusing to overwrite run directory: {folder}")
    folder.mkdir(parents=True)
    report = folder / "report.json"
    catalog = folder / "catalog.json"
    time_v = folder / "time-v.txt"
    stdout = folder / "stdout.txt"
    stderr = folder / "stderr.txt"
    verify_stdout = folder / "verify-stdout.txt"
    verify_stderr = folder / "verify-stderr.txt"
    journal = folder / "journal.json"
    contract = protocol["lanes"][lane]
    binary = build["binaries"][lane]
    source = build["source"]
    command = [
        "taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o", str(time_v),
        binary["path"], "--case", selector, *protocol["common_flags"],
        "--samples", str(contract["samples"]), "--warmup", str(contract["warmups"]),
        "--json", str(report), "--corpus-manifest", str(catalog),
    ]
    verify_command = [
        sys.executable, str(verifier), "--repo-root", str(repo_root),
        "--report", str(report), "--catalog", str(catalog), "--selector", selector,
        "--lane", lane, "--contract", "formal", "--samples", str(contract["samples"]),
        "--warmups", str(contract["warmups"]),
    ]
    record: dict[str, Any] = {
        "schema_version": 1,
        "change": CHANGE,
        "status": "running",
        "index": index,
        "lane": lane,
        "corpus": corpus,
        "role": role,
        "repeat": repeat,
        "selector": selector,
        "fresh_process": True,
        "cpu": CPU,
        "workers": WORKERS,
        "source_revision": source["revision"],
        "source_worktree": source["worktree"],
        "source_identity_sha256": build["source_identity_sha256"],
        "build": {"path": build["path"], "sha256": build["sha256"]},
        "build_sha256": build["sha256"],
        "binary": binary,
        "binary_sha256": binary["sha256"],
        "binary_bytes": binary["bytes"],
        "protocol": {"path": protocol["path"], "sha256": protocol["sha256"]},
        "protocol_sha256": protocol["sha256"],
        "verifier": {"path": str(verifier), "sha256": sha256_file(verifier)},
        "verifier_sha256": sha256_file(verifier),
        "common_flags": protocol["common_flags"],
        "contract": {"lane": lane, "samples": contract["samples"], "warmups": contract["warmups"]},
        "environment": {"RUSTUP_TOOLCHAIN": TOOLCHAIN},
        "argv": command,
        "verify_argv": verify_command,
        "artifacts": {
            "report": str(report.relative_to(root)),
            "catalog": str(catalog.relative_to(root)),
            "time_v": str(time_v.relative_to(root)),
            "stdout": str(stdout.relative_to(root)),
            "stderr": str(stderr.relative_to(root)),
            "verify_stdout": str(verify_stdout.relative_to(root)),
            "verify_stderr": str(verify_stderr.relative_to(root)),
        },
        "started_utc": utc_now(),
    }
    write_json(journal, record)
    print(f"[{index}/{total}] {lane}/{corpus}/{role}-{repeat} ({selector})", flush=True)
    environment = os.environ.copy()
    environment["RUSTUP_TOOLCHAIN"] = TOOLCHAIN
    try:
        if sha256_file(root / "protocol.json") != protocol["sha256"]:
            fail("protocol.json changed before the workload")
        if sha256_file(root / "build-candidate.json") != build["sha256"]:
            fail("build-candidate.json changed before the workload")
        if sha256_file(verifier) != record["verifier_sha256"]:
            fail("semantic verifier changed before the workload")
        with stdout.open("wb") as out_stream, stderr.open("wb") as err_stream:
            process = subprocess.run(
                command, cwd=source["worktree"], env=environment,
                stdout=out_stream, stderr=err_stream, check=False,
            )
    except (OSError, subprocess.SubprocessError) as error:
        record.update({"status": "failed", "error": str(error), "finished_utc": utc_now()})
        write_json(journal, record)
        raise CaptureError(f"{lane}/{corpus}/{role}-{repeat} failed to start: {error}") from error
    record["exit_code"] = process.returncode
    record["finished_utc"] = utc_now()
    try:
        if sha256_file(root / "protocol.json") != protocol["sha256"]:
            fail("protocol.json changed during capture")
        if sha256_file(root / "build-candidate.json") != build["sha256"]:
            fail("build-candidate.json changed during capture")
        if sha256_file(verifier) != record["verifier_sha256"]:
            fail("semantic verifier changed during capture")
        after = require_source_unchanged(source)
        record["source_after_identity_sha256"] = digest_json(after)
        observed_binary_hash = sha256_file(Path(binary["path"]))
        observed_binary_bytes = Path(binary["path"]).stat().st_size
        if observed_binary_hash != binary["sha256"] or observed_binary_bytes != binary["bytes"]:
            fail("copied binary changed during capture")
        record["binary_after"] = {"sha256": observed_binary_hash, "bytes": observed_binary_bytes}
        if process.returncode != 0:
            fail(f"workload exited {process.returncode}; see {stderr}")
        artifacts = {
            "report": artifact_record(report, root),
            "catalog": artifact_record(catalog, root),
            "time_v": artifact_record(time_v, root),
            "stdout": artifact_record(stdout, root, allow_empty=True),
            "stderr": artifact_record(stderr, root, allow_empty=True),
        }
        report_value = object_value(load_json(report, f"{lane}/{corpus}/{role}-{repeat} report"), "report")
        validate_report_identity(
            report_value, source, binary, f"{lane}/{corpus}/{role}-{repeat} report"
        )
        results = report_value.get("results")
        if not isinstance(results, list) or len(results) != 1:
            fail("report.results must contain exactly one selected result")
        result = object_value(results[0], "report.results[0]")
        if result.get("case") != selector:
            fail("report result case does not match the selected selector")
        with verify_stdout.open("wb") as verify_out, verify_stderr.open("wb") as verify_err:
            verification = subprocess.run(
                verify_command, cwd=repo_root, env=environment,
                stdout=verify_out, stderr=verify_err, check=False,
            )
        artifacts["verify_stdout"] = artifact_record(verify_stdout, root)
        artifacts["verify_stderr"] = artifact_record(verify_stderr, root, allow_empty=True)
        if verification.returncode != 0:
            fail(f"semantic verifier rejected the run; see {verify_stderr}")
        verification_value = verify_output(
            verify_stdout, f"{lane}/{corpus}/{role}-{repeat} verifier",
            selector=selector, lane=lane, samples=contract["samples"],
            warmups=contract["warmups"],
        )
        record["verification"] = {
            "status": "pass",
            "exit_code": verification.returncode,
            "stdout_sha256": artifacts["verify_stdout"]["sha256"],
            "stderr_sha256": artifacts["verify_stderr"]["sha256"],
            "summary_sha256": digest_json(verification_value),
        }
        record["verification_exit_code"] = verification.returncode
        record["artifacts"] = artifacts
    except CaptureError as error:
        record.update({"status": "failed", "error": str(error)})
        write_json(journal, record)
        raise
    record["status"] = "pass"
    write_json(journal, record)
    print(f"[{index}/{total}] pass", flush=True)
    return {
        "index": index,
        "lane": lane,
        "corpus": corpus,
        "role": role,
        "repeat": repeat,
        "selector": selector,
        "journal": str(journal.relative_to(root)),
        "journal_sha256": sha256_file(journal),
        "status": "pass",
    }


def copy_seed(source: Path | None, target: Path, label: str) -> None:
    if source is None:
        return
    source = source.expanduser().resolve()
    if not source.is_file():
        fail(f"{label} seed is missing: {source}")
    if target.exists():
        fail(f"refusing to overwrite seeded {label}: {target}")
    shutil.copy2(source, target)


def capture(
    root: Path, *, repo_root: Path | None, protocol_seed: Path | None,
    build_seed: Path | None, verifier: Path | None,
) -> None:
    root = root.expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    copy_seed(protocol_seed, root / "protocol.json", "protocol")
    copy_seed(build_seed, root / "build-candidate.json", "candidate build")
    capture_path = root / "capture.json"
    runs_path = root / "runs"
    if capture_path.exists() or runs_path.exists():
        fail(f"refusing to overwrite existing capture output: {capture_path} or {runs_path}")
    protocol = load_protocol(root)
    build = load_build(root, protocol)
    repo_root = (repo_root or Path(__file__).resolve().parents[4]).expanduser().resolve()
    if verifier is None:
        verifier = root / "verify.py"
        if not verifier.is_file():
            verifier = Path(__file__).resolve().with_name("verify.py")
    verifier = verifier.expanduser().resolve()
    if not verifier.is_file():
        fail(f"semantic verifier is missing: {verifier}")
    runs_path.mkdir()
    manifest: dict[str, Any] = {
        "schema_version": 1,
        "change": CHANGE,
        "status": "running",
        "performance_claim": None,
        "claim_authorized": False,
        "started_utc": utc_now(),
        "protocol": {"path": protocol["path"], "sha256": protocol["sha256"]},
        "build": {"path": build["path"], "sha256": build["sha256"]},
        "source": {
            "revision": build["source"]["revision"],
            "worktree": build["source"]["worktree"],
            "identity_sha256": build["source_identity_sha256"],
        },
        "binaries": build["binaries"],
        "cpu": CPU,
        "workers": WORKERS,
        "corpora": list(CORPORA),
        "selectors": list(SELECTORS),
        "lanes": protocol["lanes"],
        "order": [{"role": role, "repeat": repeat} for role, repeat in ORDER],
        "common_flags": protocol["common_flags"],
        "runtime_toolchain": TOOLCHAIN,
        "verifier": {"path": str(verifier), "sha256": sha256_file(verifier)},
        "runs": [],
        "expected_run_count": 16,
    }
    write_json(capture_path, manifest)
    total = 16
    try:
        index = 0
        for lane in MODES:
            for corpus in CORPORA:
                for role, repeat in ORDER:
                    index += 1
                    run = run_one(
                        root=root, protocol=protocol, build=build, verifier=verifier,
                        repo_root=repo_root, lane=lane, corpus=corpus, role=role,
                        repeat=repeat, index=index, total=total,
                    )
                    manifest["runs"].append(run)
                    write_json(capture_path, manifest)
        require_source_unchanged(build["source"])
        for mode in MODES:
            binary = build["binaries"][mode]
            if sha256_file(Path(binary["path"])) != binary["sha256"]:
                fail(f"{mode} binary changed after the final run")
        manifest["status"] = "pass"
        manifest["finished_utc"] = utc_now()
        write_json(capture_path, manifest)
    except CaptureError as error:
        manifest["status"] = "failed"
        manifest["finished_utc"] = utc_now()
        manifest["error"] = str(error)
        write_json(capture_path, manifest)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--protocol", type=Path, help="seed protocol.json into a fresh root")
    parser.add_argument("--build", type=Path, help="seed build-candidate.json into a fresh root")
    parser.add_argument("--verifier", type=Path)
    args = parser.parse_args()
    try:
        capture(
            args.root, repo_root=args.repo_root, protocol_seed=args.protocol,
            build_seed=args.build, verifier=args.verifier,
        )
    except CaptureError as error:
        print(f"0423 capture failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps({"change": CHANGE, "status": "pass", "runs": 16}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
