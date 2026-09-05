#!/usr/bin/env python3
"""Capture the conditional matched 0424 source-backed baseline.

The input bundle contains a frozen ``measurement-protocol.json`` and one
standard 0423-style build record for each role.  This driver runs only the
two source-backed lifecycle selectors, in normal and allocator lanes, with
control R1, candidate R1, candidate R2, control R2 order.  It keeps the old
0424 heaptrack ``runs/`` tree untouched by writing a fresh ``matched/`` root
by default.  Each child is a fresh process on CPU 2; reports, catalogs, raw
GNU ``time -v`` output, verifier output, and compact custody journals are
retained.  It records matched corpus/output identities while allowing the
control and candidate binary hashes to differ.
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


CHANGE = 424
VERIFIER_CHANGE = 423
CPU = 2
WORKERS = 1
TOOLCHAIN = "1.98.1"
MODES = ("normal", "allocator")
CORPORA = ("plain", "media_rich")
SELECTORS = (
    "pptx_source_backed_cross_copy_plain_lifecycle",
    "pptx_source_backed_cross_copy_media_rich_lifecycle",
)
ORDER = (
    ("control", "R1"),
    ("candidate", "R1"),
    ("candidate", "R2"),
    ("control", "R2"),
)
COMMON_FLAGS = (
    "--shape", "many-small", "--payload", "compressible", "--writer-shape", "large",
    "--xlsx-shape", "medium", "--xlsx-cell-crud-shape", "medium",
    "--xlsx-row-visibility-shape", "medium", "--semantic-shape", "medium",
    "--workers", "1", "--filesystem-cache", "warm",
)
CAPTURE_FLAGS = frozenset(
    {"--case", "--cases", "--json", "--output", "--corpus-manifest",
     "--samples", "--warmup", "--warmups"}
)
SOURCE_SPECS = (
    "Cargo.toml", "Cargo.lock", "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/Cargo.lock", "tools/perf-baseline/src",
    "crates/litchi-opc/Cargo.toml", "crates/litchi-opc/src",
    "crates/litchi-pptx/Cargo.toml", "crates/litchi-pptx/src",
)
SHA256 = re.compile(r"^[0-9a-f]{64}$")
GIT_REVISION = re.compile(r"^[0-9a-f]{40}$")


class MeasureError(RuntimeError):
    """An input, identity, artifact, verifier, or subprocess error."""


def fail(message: str) -> None:
    raise MeasureError(message)


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


def object_value(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label} must be an object")
    return value


def string_value(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label} must be a non-empty string")
    return value


def integer_value(value: Any, label: str, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label} must be an integer >= {minimum}")
    return value


def run_git(worktree: Path, *arguments: str) -> str:
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


def source_identity(worktree: Path) -> dict[str, Any]:
    worktree = worktree.expanduser().resolve()
    if not worktree.is_dir():
        fail(f"source worktree is missing: {worktree}")
    revision = run_git(worktree, "rev-parse", "HEAD").strip()
    status = run_git(worktree, "status", "--porcelain=v1", "--untracked-files=all")
    tracked = run_git(worktree, "ls-files", "-z", "--", *SOURCE_SPECS)
    files: list[dict[str, Any]] = []
    for relative in sorted(item for item in tracked.split("\0") if item):
        path = worktree / relative
        if not path.is_file():
            fail(f"tracked source file is missing: {path}")
        files.append({"path": relative, "bytes": path.stat().st_size, "sha256": sha256_file(path)})
    if not files:
        fail(f"source binding matched no files in {worktree}")
    return {
        "worktree": str(worktree), "revision": revision,
        "git_status_porcelain": status, "clean": status == "",
        "source_files": files,
    }


def require_source_unchanged(source: dict[str, Any]) -> dict[str, Any]:
    observed = source_identity(Path(source["worktree"]))
    if observed != source:
        fail("source worktree or bound source files changed during measurement")
    return observed


def require_baseline_ancestor(source: dict[str, Any], baseline: str) -> None:
    try:
        process = subprocess.run(
            ["git", "merge-base", "--is-ancestor", baseline, source["revision"]],
            cwd=source["worktree"], capture_output=True, text=True, check=False,
        )
    except OSError as error:
        fail(f"cannot check baseline ancestry: {error}")
    if process.returncode != 0:
        fail(f"source revision {source['revision']} does not descend from {baseline}")


def validate_binary(value: Any, label: str) -> dict[str, Any]:
    binary = object_value(value, label)
    raw_path = string_value(binary.get("path"), f"{label}.path")
    path = Path(raw_path)
    if not path.is_absolute() or not raw_path.startswith("/tmp/"):
        fail(f"{label}.path must be an absolute copied /tmp binary")
    if not path.is_file() or not os.access(path, os.X_OK):
        fail(f"{label} is missing or not executable: {path}")
    expected_hash = string_value(binary.get("sha256", binary.get("binary_sha256")), f"{label}.sha256")
    alias_hash = binary.get("binary_sha256")
    if alias_hash is not None and alias_hash != expected_hash:
        fail(f"{label} sha256 and binary_sha256 disagree")
    if SHA256.fullmatch(expected_hash) is None or sha256_file(path) != expected_hash:
        fail(f"{label} hash differs from its build identity")
    expected_bytes = integer_value(binary.get("bytes", binary.get("binary_bytes")), f"{label}.bytes", 1)
    alias_bytes = binary.get("binary_bytes")
    if alias_bytes is not None and alias_bytes != expected_bytes:
        fail(f"{label} bytes and binary_bytes disagree")
    if path.stat().st_size != expected_bytes:
        fail(f"{label} size differs from its build identity")
    mode_bits = binary.get("mode_bits")
    if mode_bits is not None and mode_bits != path.stat().st_mode & 0o7777:
        fail(f"{label} mode bits differ from its build identity")
    if binary.get("executable") is not None and binary.get("executable") is not True:
        fail(f"{label} build identity is not executable")
    return {
        "path": str(path), "sha256": expected_hash, "bytes": expected_bytes,
        "binary_sha256": expected_hash, "binary_bytes": expected_bytes,
        "mode_bits": mode_bits, "executable": True, "profile": binary.get("profile"),
        "label": binary.get("label"),
    }


def load_protocol(root: Path) -> dict[str, Any]:
    path = root / "measurement-protocol.json"
    protocol = object_value(load_json(path, "measurement protocol"), str(path))
    if protocol.get("change") != CHANGE:
        fail("measurement-protocol.json.change must be 424")
    baseline = string_value(protocol.get("baseline_revision"), "baseline_revision").lower()
    if GIT_REVISION.fullmatch(baseline) is None:
        fail("baseline_revision must be a full lowercase git revision")
    if protocol.get("cpu") != CPU or protocol.get("workers") != WORKERS:
        fail("measurement protocol must pin CPU 2 and one worker")
    if protocol.get("corpora") != list(CORPORA):
        fail(f"measurement protocol.corpora must be {list(CORPORA)!r}")
    if protocol.get("selectors") != list(SELECTORS):
        fail(f"measurement protocol.selectors must be {list(SELECTORS)!r}")
    if protocol.get("common_flags") != list(COMMON_FLAGS):
        fail("measurement protocol.common_flags differ from the frozen 0423 flags")
    if any(flag in CAPTURE_FLAGS for flag in COMMON_FLAGS):
        fail("measurement protocol common_flags contain capture-owned options")
    lanes_raw = object_value(protocol.get("lanes"), "measurement protocol.lanes")
    lanes: dict[str, dict[str, int]] = {}
    for mode, expected in (("normal", (100, 10)), ("allocator", (30, 3))):
        value = object_value(lanes_raw.get(mode), f"lanes.{mode}")
        contract = {
            "samples": integer_value(value.get("samples"), f"lanes.{mode}.samples", 1),
            "warmups": integer_value(value.get("warmups"), f"lanes.{mode}.warmups", 0),
        }
        if (contract["samples"], contract["warmups"]) != expected:
            fail(f"lanes.{mode} must be {expected[0]}/{expected[1]}")
        lanes[mode] = contract
    if set(lanes_raw) != set(MODES):
        fail("measurement protocol lanes must contain only normal and allocator")
    expected_order = [{"role": role, "repeat": repeat} for role, repeat in ORDER]
    if protocol.get("order") != expected_order:
        fail(f"measurement protocol.order must be {expected_order!r}")
    if protocol.get("expected_run_count") != 16:
        fail("measurement protocol.expected_run_count must be 16")
    return {
        "path": "measurement-protocol.json", "sha256": sha256_file(path),
        "value": protocol, "baseline": baseline, "lanes": lanes,
        "common_flags": list(COMMON_FLAGS), "corpora": list(CORPORA),
        "selectors": list(SELECTORS), "order": expected_order,
    }


def load_build(root: Path, role: str, protocol: dict[str, Any]) -> dict[str, Any]:
    path = root / f"measurement-build-{role}.json"
    build = object_value(load_json(path, f"{role} measurement build"), str(path))
    if build.get("status") != "pass" or build.get("exit_code") != 0:
        fail(f"{path} is not a successful build record")
    if build.get("role") not in (None, role, "candidate"):
        fail(f"{path}.role does not identify the requested {role} role")
    build_protocol = string_value(build.get("protocol_sha256"), f"{role}.protocol_sha256")
    if SHA256.fullmatch(build_protocol) is None:
        fail(f"{role}.protocol_sha256 must be a SHA-256")
    source_before = object_value(build.get("source_before"), f"{role}.source_before")
    source_after = object_value(build.get("source_after"), f"{role}.source_after")
    if source_before != source_after or source_before.get("clean") is not True:
        fail(f"{role} source_before/source_after must be the same clean identity")
    revision = string_value(source_before.get("revision"), f"{role}.source.revision")
    if GIT_REVISION.fullmatch(revision) is None:
        fail(f"{role} source revision must be a full lowercase git revision")
    if role == "control" and revision != protocol["baseline"]:
        fail("control source revision must equal measurement baseline_revision")
    require_baseline_ancestor(source_before, protocol["baseline"])
    source = source_identity(Path(string_value(source_before.get("worktree"), f"{role}.source.worktree")))
    if source != source_before:
        fail(f"{role} source worktree differs from source_before identity")
    binaries_raw = object_value(build.get("binaries"), f"{role}.binaries")
    binaries = {mode: validate_binary(binaries_raw.get(mode), f"{role}.binaries.{mode}") for mode in MODES}
    return {
        "role": role, "path": f"measurement-build-{role}.json",
        "sha256": sha256_file(path), "value": build,
        "build_protocol_sha256": build_protocol, "source": source,
        "source_identity_sha256": digest_json(source), "binaries": binaries,
    }


def artifact(path: Path, root: Path, *, allow_empty: bool = False) -> dict[str, Any]:
    if not path.is_file():
        fail(f"capture artifact is missing: {path}")
    size = path.stat().st_size
    if size == 0 and not allow_empty:
        fail(f"capture artifact is empty: {path}")
    try:
        relative = path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        fail(f"capture artifact escapes output root: {path}")
    return {"path": relative, "bytes": size, "sha256": sha256_file(path), "allow_empty": allow_empty}


def validate_report_identity(report: dict[str, Any], source: dict[str, Any], binary: dict[str, Any], label: str) -> None:
    environment = object_value(report.get("environment"), f"{label}.environment")
    if environment.get("git_revision") != source["revision"]:
        fail(f"{label}.environment.git_revision does not match the role source")
    if environment.get("git_worktree_dirty") is not False:
        fail(f"{label}.environment.git_worktree_dirty must be false")
    identity = object_value(report.get("binary_identity"), f"{label}.binary_identity")
    if identity.get("binary_sha256") != binary["sha256"]:
        fail(f"{label}.binary_identity.binary_sha256 does not match the role binary")
    if identity.get("binary_bytes") != binary["bytes"]:
        fail(f"{label}.binary_identity.binary_bytes does not match the role binary")


def verify_output(
    path: Path, label: str, *, selector: str, lane: str, samples: int, warmups: int,
    report: Path, catalog: Path,
) -> dict[str, Any]:
    value = object_value(load_json(path, label), label)
    required = {
        "change", "status", "claim_authorized", "performance_claim", "selector", "lane",
        "samples", "warmups", "report_count", "reports",
    }
    missing = sorted(required.difference(value))
    if missing:
        fail(f"{label} is missing required fields: {missing}")
    if value["change"] != VERIFIER_CHANGE or value["status"] != "pass":
        fail(f"{label} is not a passing 0423 report verification")
    if value["claim_authorized"] is not False or value["performance_claim"] is not None:
        fail(f"{label} must withhold performance claims")
    if value["selector"] != selector or value["lane"] != lane:
        fail(f"{label} selector/lane does not match the run")
    if value["samples"] != samples or value["warmups"] != warmups:
        fail(f"{label} samples/warmups do not match the run")
    if value["report_count"] != 1 or not isinstance(value["reports"], list) or len(value["reports"]) != 1:
        fail(f"{label} must describe exactly one report")
    proof = object_value(value["reports"][0], f"{label}.reports[0]")
    if proof.get("report_sha256") != sha256_file(report):
        fail(f"{label}.reports[0].report_sha256 does not match report")
    if proof.get("catalog_sha256") != sha256_file(catalog):
        fail(f"{label}.reports[0].catalog_sha256 does not match catalog")
    return value


def run_one(
    *, bundle: Path, output: Path, protocol: dict[str, Any], builds: dict[str, dict[str, Any]],
    verifier: Path, repo_root: Path, lane: str, corpus: str, role: str, repeat: str,
    index: int, total: int,
) -> dict[str, Any]:
    selector = f"pptx_source_backed_cross_copy_{corpus}_lifecycle"
    folder = output / "runs" / lane / corpus / f"{role}-{repeat}"
    if folder.exists():
        fail(f"refusing to overwrite run directory: {folder}")
    folder.mkdir(parents=True)
    report, catalog = folder / "report.json", folder / "catalog.json"
    time_v = folder / "time-v.txt"
    stdout, stderr = folder / "stdout.txt", folder / "stderr.txt"
    verify_stdout, verify_stderr = folder / "verify-stdout.txt", folder / "verify-stderr.txt"
    journal = folder / "journal.json"
    contract = protocol["lanes"][lane]
    build, source, binary = builds[role], builds[role]["source"], builds[role]["binaries"][lane]
    command = [
        "taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o", str(time_v), binary["path"],
        "--case", selector, *protocol["common_flags"], "--samples", str(contract["samples"]),
        "--warmup", str(contract["warmups"]), "--json", str(report), "--corpus-manifest", str(catalog),
    ]
    verify_command = [
        sys.executable, "-B", str(verifier), "--repo-root", str(repo_root), "--report", str(report),
        "--catalog", str(catalog), "--selector", selector, "--lane", lane, "--contract", "formal",
        "--samples", str(contract["samples"]), "--warmups", str(contract["warmups"]),
    ]
    verifier_hash = sha256_file(verifier)
    protocol_ref = os.path.relpath(bundle / "measurement-protocol.json", output)
    build_ref = os.path.relpath(bundle / f"measurement-build-{role}.json", output)
    record: dict[str, Any] = {
        "schema_version": 1, "change": CHANGE, "status": "running", "index": index,
        "lane": lane, "corpus": corpus, "role": role, "repeat": repeat, "selector": selector,
        "fresh_process": True, "cpu": CPU, "workers": WORKERS,
        "source_revision": source["revision"], "source_worktree": source["worktree"],
        "source_identity_sha256": build["source_identity_sha256"],
        "source_before_identity_sha256": build["source_identity_sha256"],
        "build": {"path": build_ref, "sha256": build["sha256"],
                  "protocol_sha256": build["build_protocol_sha256"]},
        "build_sha256": build["sha256"], "binary": binary,
        "binary_sha256": binary["sha256"], "binary_bytes": binary["bytes"],
        "binary_before": {"sha256": binary["sha256"], "bytes": binary["bytes"]},
        "protocol": {"path": protocol_ref, "sha256": protocol["sha256"]},
        "protocol_sha256": protocol["sha256"], "verifier": {"path": str(verifier), "sha256": verifier_hash},
        "verifier_sha256": verifier_hash, "common_flags": protocol["common_flags"],
        "contract": {"lane": lane, "samples": contract["samples"], "warmups": contract["warmups"]},
        "environment": {"RUSTUP_TOOLCHAIN": TOOLCHAIN, "PYTHONDONTWRITEBYTECODE": "1"},
        "argv": command, "verify_argv": verify_command,
        "artifacts": {name: str(path.relative_to(output)) for name, path in {
            "report": report, "catalog": catalog, "time_v": time_v, "stdout": stdout,
            "stderr": stderr, "verify_stdout": verify_stdout, "verify_stderr": verify_stderr,
        }.items()},
        "started_utc": utc_now(),
    }
    write_json(journal, record)
    print(f"[{index}/{total}] {lane}/{corpus}/{role}-{repeat}", flush=True)
    env = os.environ.copy()
    env.update({"RUSTUP_TOOLCHAIN": TOOLCHAIN, "PYTHONDONTWRITEBYTECODE": "1"})
    try:
        if sha256_file(bundle / "measurement-protocol.json") != protocol["sha256"]:
            fail("measurement protocol changed before the workload")
        if sha256_file(bundle / f"measurement-build-{role}.json") != build["sha256"]:
            fail(f"{role} measurement build record changed before the workload")
        if sha256_file(verifier) != verifier_hash:
            fail("pinned verifier changed before the workload")
        with stdout.open("wb") as out, stderr.open("wb") as err:
            process = subprocess.run(command, cwd=source["worktree"], env=env, stdout=out, stderr=err, check=False)
    except (OSError, subprocess.SubprocessError) as error:
        record.update({"status": "failed", "error": str(error), "finished_utc": utc_now()})
        write_json(journal, record)
        raise MeasureError(f"{lane}/{corpus}/{role}-{repeat} failed to start: {error}") from error
    record["exit_code"] = process.returncode
    record["finished_utc"] = utc_now()
    try:
        if sha256_file(bundle / "measurement-protocol.json") != protocol["sha256"]:
            fail("measurement protocol changed during the workload")
        if sha256_file(bundle / f"measurement-build-{role}.json") != build["sha256"]:
            fail(f"{role} measurement build record changed during the workload")
        if sha256_file(verifier) != verifier_hash:
            fail("pinned verifier changed during the workload")
        after = require_source_unchanged(source)
        record["source_after_identity_sha256"] = digest_json(after)
        record["source_unchanged"] = True
        record["protocol_after_sha256"] = sha256_file(bundle / "measurement-protocol.json")
        record["verifier_after_sha256"] = sha256_file(verifier)
        binary_after = sha256_file(Path(binary["path"]))
        if binary_after != binary["sha256"] or Path(binary["path"]).stat().st_size != binary["bytes"]:
            fail(f"{role}/{lane} copied binary changed during the workload")
        record["binary_after"] = {"sha256": binary_after, "bytes": binary["bytes"]}
        record["binary_unchanged"] = True
        if process.returncode != 0:
            fail(f"workload exited {process.returncode}; see {stderr}")
        artifacts = {
            "report": artifact(report, output), "catalog": artifact(catalog, output),
            "time_v": artifact(time_v, output), "stdout": artifact(stdout, output, allow_empty=True),
            "stderr": artifact(stderr, output, allow_empty=True),
        }
        report_value = object_value(load_json(report, "workload report"), "workload report")
        validate_report_identity(report_value, source, binary, "workload report")
        results = report_value.get("results")
        if not isinstance(results, list) or len(results) != 1:
            fail("workload report.results must contain exactly one result")
        result = object_value(results[0], "workload report.results[0]")
        if result.get("case") != selector:
            fail("workload report result case does not match selector")
        output_sha256 = result.get("output_sha256")
        if not isinstance(output_sha256, str) or SHA256.fullmatch(output_sha256) is None:
            fail("workload report result output_sha256 is not a valid digest")
        result_corpus = object_value(result.get("corpus"), "workload report.results[0].corpus")
        with verify_stdout.open("wb") as out, verify_stderr.open("wb") as err:
            checked = subprocess.run(verify_command, cwd=repo_root, env=env, stdout=out, stderr=err, check=False)
        artifacts["verify_stdout"] = artifact(verify_stdout, output)
        artifacts["verify_stderr"] = artifact(verify_stderr, output, allow_empty=True)
        if checked.returncode != 0:
            fail(f"semantic verifier rejected the run; see {verify_stderr}")
        proof = verify_output(
            verify_stdout, "single-report verifier", selector=selector, lane=lane,
            samples=contract["samples"], warmups=contract["warmups"], report=report, catalog=catalog,
        )
        record["verification"] = {
            "status": "pass", "exit_code": checked.returncode,
            "stdout_sha256": artifacts["verify_stdout"]["sha256"],
            "stderr_sha256": artifacts["verify_stderr"]["sha256"],
            "summary_sha256": digest_json(proof),
        }
        record["verification_exit_code"] = checked.returncode
        record["artifacts"] = artifacts
        record["output_sha256"] = output_sha256
        record["corpus_identity_sha256"] = digest_json(result_corpus)
    except (OSError, subprocess.SubprocessError) as error:
        record.update({"status": "failed", "error": str(error)})
        write_json(journal, record)
        raise MeasureError(f"{lane}/{corpus}/{role}-{repeat} verifier failed to start: {error}") from error
    except MeasureError as error:
        record.update({"status": "failed", "error": str(error)})
        write_json(journal, record)
        raise
    record["status"] = "pass"
    write_json(journal, record)
    print(f"[{index}/{total}] pass", flush=True)
    return {
        "index": index, "lane": lane, "corpus": corpus, "role": role, "repeat": repeat,
        "selector": selector, "journal": str(journal.relative_to(output)),
        "journal_sha256": sha256_file(journal), "status": "pass",
        "output_sha256": record["output_sha256"],
        "corpus_identity_sha256": record["corpus_identity_sha256"],
    }


def measure(bundle: Path, output: Path, *, repo_root: Path | None, verifier: Path | None) -> None:
    bundle = bundle.expanduser().resolve()
    output = output.expanduser().resolve()
    if not bundle.is_dir():
        fail(f"measurement bundle is missing: {bundle}")
    if output.exists() and any(output.iterdir()):
        fail(f"refusing to overwrite non-empty measurement output: {output}")
    output.mkdir(parents=True, exist_ok=True)
    protocol = load_protocol(bundle)
    builds = {role: load_build(bundle, role, protocol) for role in ("control", "candidate")}
    verifier = (verifier or (bundle / "pinned" / "verify-report.py")).expanduser().resolve()
    if not verifier.is_file():
        fail(f"pinned verifier is missing: {verifier}")
    repo_root = (repo_root or (bundle / "pinned")).expanduser().resolve()
    if not (repo_root / "tools" / "perf_abba_summary.py").is_file():
        fail(f"verifier repo root lacks pinned tools: {repo_root}")
    runs = output / "runs"
    if runs.exists():
        fail(f"refusing to overwrite existing run tree: {runs}")
    runs.mkdir()
    manifest: dict[str, Any] = {
        "schema_version": 1, "change": CHANGE, "status": "running", "claim_authorized": False,
        "performance_claim": None, "started_utc": utc_now(),
        "protocol": {"path": os.path.relpath(bundle / "measurement-protocol.json", output), "sha256": protocol["sha256"]},
        "builds": {role: {"path": os.path.relpath(bundle / f"measurement-build-{role}.json", output), "sha256": builds[role]["sha256"],
                           "source_revision": builds[role]["source"]["revision"],
                           "source_identity_sha256": builds[role]["source_identity_sha256"],
                           "build_protocol_sha256": builds[role]["build_protocol_sha256"],
                           "binaries": builds[role]["binaries"]} for role in builds},
        "verifier": {"path": str(verifier), "sha256": sha256_file(verifier)},
        "repo_root": str(repo_root), "cpu": CPU, "workers": WORKERS,
        "corpora": list(CORPORA), "selectors": list(SELECTORS),
        "lanes": protocol["lanes"], "order": protocol["order"],
        "common_flags": protocol["common_flags"], "runs": [], "expected_run_count": 16,
    }
    capture_path = output / "capture.json"
    write_json(capture_path, manifest)
    try:
        index = 0
        for lane in MODES:
            for corpus in CORPORA:
                for role, repeat in ORDER:
                    index += 1
                    row = run_one(
                        bundle=bundle, output=output, protocol=protocol, builds=builds,
                        verifier=verifier, repo_root=repo_root, lane=lane, corpus=corpus,
                        role=role, repeat=repeat, index=index, total=16,
                    )
                    manifest["runs"].append(row)
                    write_json(capture_path, manifest)
        for corpus in CORPORA:
            rows = [row for row in manifest["runs"] if row["corpus"] == corpus]
            outputs = {row["output_sha256"] for row in rows}
            corpora = {row["corpus_identity_sha256"] for row in rows}
            if len(outputs) != 1 or len(corpora) != 1:
                fail(f"{corpus} control/candidate corpus or output identity differs")
        for role, build in builds.items():
            require_source_unchanged(build["source"])
            for lane in MODES:
                binary = build["binaries"][lane]
                if sha256_file(Path(binary["path"])) != binary["sha256"]:
                    fail(f"{role}/{lane} binary changed after the final run")
        manifest["status"] = "pass"
        manifest["finished_utc"] = utc_now()
        write_json(capture_path, manifest)
    except MeasureError as error:
        manifest["status"] = "failed"
        manifest["finished_utc"] = utc_now()
        manifest["error"] = str(error)
        write_json(capture_path, manifest)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent,
                        help="bundle root containing measurement protocol/builds")
    parser.add_argument("--output", "--output-root", dest="output", type=Path,
                        help="fresh output root; defaults to ROOT/matched")
    parser.add_argument("--repo-root", type=Path,
                        help="pinned verifier repository root; defaults to ROOT/pinned")
    parser.add_argument("--verifier", type=Path,
                        help="pinned verify-report.py; defaults to ROOT/pinned/verify-report.py")
    args = parser.parse_args()
    try:
        bundle = args.root.expanduser().resolve()
        output = (args.output or (bundle / "matched")).expanduser().resolve()
        measure(bundle, output, repo_root=args.repo_root, verifier=args.verifier)
    except MeasureError as error:
        print(f"0424 matched measurement failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps({"change": CHANGE, "status": "pass", "runs": 16}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
