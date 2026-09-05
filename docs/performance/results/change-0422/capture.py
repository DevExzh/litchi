#!/usr/bin/env python3
"""Capture the bounded 0422 allocator diagnostic.

This driver deliberately owns only the allocator lane.  It runs each of the
two declared selectors twice, in a fresh process on CPU 2, and retains the
report, corpus catalog, GNU ``time -v`` output, command logs, verifier logs,
and a per-run journal.  The source tree and copied allocator binary are
checked before and after every child.  A failed or partial capture is never
silently resumed and an existing capture is never overwritten.

The output root must contain ``protocol.json`` and ``build-candidate.json``.
For a new root, ``--protocol`` and ``--build`` may point at those two files;
they are copied into the canonical names before capture starts.  The verifier
is run once per report with its exact single-report interface.  Replay does
not need the source worktree or binary; ``summarize.py`` reruns the retained
verifier against the report/catalog using the shared repository validators and
checks all custody hashes from the journals.
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


CHANGE = 422
CPU = 2
REPEAT_NAMES = ("R1", "R2")
SELECTORS = (
    "pptx_cross_copy_media_rich_lifecycle",
    "pptx_cross_copy_plain_lifecycle",
)
SAMPLES = 30
WARMUPS = 3
TOOLCHAIN = "1.98.1"
SAFE_NAME = re.compile(r"^[A-Za-z0-9_.-]+$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")
GIT_REVISION = re.compile(r"^[0-9a-f]{40}$")
CAPTURE_OWNED_FLAGS = frozenset(
    {"--case", "--json", "--corpus-manifest", "--samples", "--warmup", "--warmups"}
)


class CaptureError(RuntimeError):
    """A fail-closed capture input, identity, or subprocess error."""


def fail(message: str) -> None:
    raise CaptureError(message)


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def strict_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON constant {value!r}")


def load_json(path: Path, label: str) -> Any:
    try:
        value = json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=strict_pairs,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"cannot load {label} {path}: {error}")
    return value


def write_json(path: Path, value: Any) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    try:
        temporary.write_text(
            json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False)
            + "\n",
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
        payload = json.dumps(
            value, sort_keys=True, separators=(",", ":"),
            ensure_ascii=False, allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot hash JSON identity: {error}")
    return hashlib.sha256(payload).hexdigest()


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


def sha_value(value: Any, label: str) -> str:
    value = string_value(value, label).lower()
    if SHA256.fullmatch(value) is None:
        fail(f"{label} must be a lowercase SHA-256")
    return value


def safe_name(value: Any, label: str) -> str:
    value = string_value(value, label)
    if SAFE_NAME.fullmatch(value) is None:
        fail(f"{label} contains unsafe path characters")
    return value


def relative_path(root: Path, value: Any, label: str) -> Path:
    raw = string_value(value, label)
    path = Path(raw)
    if path.is_absolute() or ".." in path.parts:
        fail(f"{label} must be a relative path without '..'")
    resolved = (root / path).resolve()
    try:
        resolved.relative_to(root.resolve())
    except ValueError:
        fail(f"{label} escapes the output root")
    return resolved


def git_state(worktree: Path) -> dict[str, Any]:
    if not worktree.is_dir():
        fail(f"source worktree is missing: {worktree}")

    def run(arguments: list[str]) -> str:
        try:
            process = subprocess.run(
                ["git", *arguments], cwd=worktree, capture_output=True,
                text=True, check=False,
            )
        except OSError as error:
            fail(f"cannot run git in {worktree}: {error}")
        if process.returncode != 0:
            detail = (process.stderr or process.stdout).strip()
            fail(f"git {' '.join(arguments)} failed in {worktree}: {detail}")
        return process.stdout

    revision = run(["rev-parse", "HEAD"]).strip()
    status = run(["status", "--porcelain=v1", "--untracked-files=all"])
    if not GIT_REVISION.fullmatch(revision):
        fail(f"source revision is not a full lowercase git object: {revision!r}")
    return {
        "worktree": str(worktree.resolve()),
        "revision": revision,
        "git_status_porcelain": status,
        "clean": status == "",
    }


def validate_source(value: Any, label: str) -> dict[str, Any]:
    source = object_value(value, label)
    worktree = Path(string_value(source.get("worktree"), f"{label}.worktree"))
    if not worktree.is_absolute() or not worktree.is_dir():
        fail(f"{label}.worktree must be an existing absolute directory")
    if source.get("clean") is not True or source.get("git_status_porcelain") != "":
        fail(f"{label} must be clean")
    revision = string_value(source.get("revision"), f"{label}.revision").lower()
    if GIT_REVISION.fullmatch(revision) is None:
        fail(f"{label}.revision must be a lowercase full git object")
    return source


def current_source(source: dict[str, Any]) -> dict[str, Any]:
    state = git_state(Path(source["worktree"]))
    if state["revision"] != source["revision"] or not state["clean"]:
        fail(f"source changed or became dirty: {source['worktree']}")
    return state


def binary_identity(value: Any) -> dict[str, Any]:
    binary = object_value(value, "build-candidate.binaries.allocator")
    raw_path = string_value(binary.get("path"), "allocator binary path")
    path = Path(raw_path)
    if not path.is_absolute() or not raw_path.startswith("/tmp/"):
        fail("allocator binary must be an absolute copied /tmp binary")
    if not path.is_file() or not os.access(path, os.X_OK):
        fail(f"allocator binary is missing or not executable: {path}")
    observed_hash = sha256_file(path)
    expected_hash = sha_value(
        binary.get("sha256", binary.get("binary_sha256")),
        "build-candidate.binaries.allocator.sha256",
    )
    if observed_hash != expected_hash:
        fail("allocator binary hash differs from build identity")
    observed_bytes = path.stat().st_size
    expected_bytes = binary.get("bytes", binary.get("binary_bytes"))
    if expected_bytes is not None and integer_value(expected_bytes, "allocator binary bytes", 1) != observed_bytes:
        fail("allocator binary size differs from build identity")
    return {
        "path": str(path),
        "sha256": observed_hash,
        "bytes": observed_bytes,
        "mode_bits": binary.get("mode_bits"),
        "profile": binary.get("profile"),
    }


def load_protocol(root: Path) -> dict[str, Any]:
    path = root / "protocol.json"
    protocol = object_value(load_json(path, "protocol"), str(path))
    if protocol.get("change") != CHANGE:
        fail("protocol.change must be 422")
    if protocol.get("mode") != "allocator":
        fail("protocol.mode must be 'allocator'")
    if protocol.get("cpu") != CPU or protocol.get("workers") != 1:
        fail("protocol must pin CPU 2 and one worker")
    if protocol.get("repeats") != 2 or protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS:
        fail("protocol must declare two repeats of 30 samples and 3 warmups")
    selectors = protocol.get("selectors")
    if selectors != list(SELECTORS):
        fail(f"protocol.selectors must be {list(SELECTORS)!r}")
    flags = protocol.get("common_flags")
    if not isinstance(flags, list) or not flags or any(not isinstance(item, str) or not item for item in flags):
        fail("protocol.common_flags must be a non-empty string list")
    if any(item in CAPTURE_OWNED_FLAGS for item in flags):
        fail("protocol.common_flags contains a capture-owned flag")
    return {
        "path": "protocol.json",
        "sha256": sha256_file(path),
        "value": protocol,
        "common_flags": list(flags),
    }


def load_build(root: Path, protocol: dict[str, Any]) -> dict[str, Any]:
    path = root / "build-candidate.json"
    build = object_value(load_json(path, "candidate build"), str(path))
    if build.get("change") != CHANGE or build.get("role") != "candidate":
        fail("build-candidate.json has the wrong change or role")
    if build.get("status") != "pass" or build.get("exit_code") != 0:
        fail("build-candidate.json is not a successful build record")
    if sha_value(build.get("protocol_sha256"), "build protocol_sha256") != protocol["sha256"]:
        fail("candidate build is not bound to protocol.json")
    before = validate_source(build.get("source_before"), "source_before")
    after = validate_source(build.get("source_after"), "source_after")
    if before != after:
        fail("source_before and source_after differ in the candidate build record")
    current_source(before)
    binary = binary_identity(object_value(build.get("binaries"), "build binaries").get("allocator"))
    return {
        "path": "build-candidate.json",
        "sha256": sha256_file(path),
        "value": build,
        "source": before,
        "binary": binary,
    }


def artifact(path: Path, root: Path, *, allow_empty: bool) -> dict[str, Any]:
    if not path.is_file():
        fail(f"capture artifact is missing: {path}")
    size = path.stat().st_size
    if size == 0 and not allow_empty:
        fail(f"capture artifact is empty: {path}")
    return {
        "path": str(path.resolve().relative_to(root.resolve())),
        "bytes": size,
        "sha256": sha256_file(path),
        "allow_empty": allow_empty,
    }


def verify_result(stdout: Path, label: str) -> dict[str, Any]:
    value = object_value(load_json(stdout, label), label)
    if value.get("claim_authorized") is not False or value.get("performance_claim") is not None:
        fail(f"{label} must withhold all performance claims")
    if value.get("selector") not in SELECTORS or value.get("lane") != "allocator":
        fail(f"{label} has the wrong verifier identity")
    if value.get("samples") != SAMPLES or value.get("warmups") != WARMUPS or value.get("report_count") != 1:
        fail(f"{label} has the wrong single-report contract")
    reports = value.get("reports")
    if not isinstance(reports, list) or len(reports) != 1:
        fail(f"{label}.reports must contain one row")
    return value


def run_one(
    *, root: Path, protocol: dict[str, Any], build: dict[str, Any], verifier: Path,
    repo_root: Path, repeat: str, selector: str, index: int, total: int,
) -> dict[str, Any]:
    folder = root / "runs" / repeat / selector
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
    source = build["source"]
    binary = build["binary"]
    command = [
        "taskset", "-c", str(CPU), "/usr/bin/time", "-v", "-o", str(time_v),
        binary["path"], "--case", selector, *protocol["common_flags"],
        "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
        "--json", str(report), "--corpus-manifest", str(catalog),
    ]
    verify_command = [
        sys.executable, str(verifier), "--repo-root", str(repo_root),
        "--report", str(report), "--catalog", str(catalog),
        "--selector", selector, "--lane", "allocator",
        "--samples", str(SAMPLES), "--warmups", str(WARMUPS),
    ]
    record: dict[str, Any] = {
        "schema_version": 1,
        "change": CHANGE,
        "status": "running",
        "repeat": repeat,
        "selector": selector,
        "index": index,
        "fresh_process": True,
        "cpu": CPU,
        "source_before": source,
        "source_identity_sha256": digest_json(source),
        "source_revision": source["revision"],
        "source_worktree": source["worktree"],
        "build": {"path": build["path"], "sha256": build["sha256"]},
        "build_sha256": build["sha256"],
        "binary": binary,
        "binary_before": {"sha256": binary["sha256"], "bytes": binary["bytes"]},
        "binary_sha256": binary["sha256"],
        "binary_bytes": binary["bytes"],
        "protocol": {"path": protocol["path"], "sha256": protocol["sha256"]},
        "protocol_sha256": protocol["sha256"],
        "common_flags": protocol["common_flags"],
        "contract": {"lane": "allocator", "samples": SAMPLES, "warmups": WARMUPS},
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
    print(f"[{index}/{total}] {repeat}/{selector}", flush=True)
    env = os.environ.copy()
    env["RUSTUP_TOOLCHAIN"] = TOOLCHAIN
    try:
        with stdout.open("wb") as out_stream, stderr.open("wb") as err_stream:
            process = subprocess.run(
                command, cwd=source["worktree"], env=env,
                stdout=out_stream, stderr=err_stream, check=False,
            )
    except (OSError, subprocess.SubprocessError) as error:
        record.update({"status": "failed", "error": str(error), "finished_utc": utc_now()})
        write_json(journal, record)
        raise CaptureError(f"{repeat}/{selector} failed to start: {error}") from error
    record["exit_code"] = process.returncode
    record["finished_utc"] = utc_now()
    try:
        record["source_after"] = current_source(source)
        binary_after_sha = sha256_file(Path(binary["path"]))
        binary_after_bytes = Path(binary["path"]).stat().st_size
        if binary_after_sha != binary["sha256"] or binary_after_bytes != binary["bytes"]:
            fail("allocator binary changed during capture")
        record["binary_after"] = {"sha256": binary_after_sha, "bytes": binary_after_bytes}
        if process.returncode != 0:
            fail(f"{repeat}/{selector} exited {process.returncode}; see {stderr}")
        # The run logs may legitimately be empty.  Reports, catalogs, time-v,
        # and verifier output are evidence and therefore must be non-empty.
        records = {
            "report": artifact(report, root, allow_empty=False),
            "catalog": artifact(catalog, root, allow_empty=False),
            "time_v": artifact(time_v, root, allow_empty=False),
            "stdout": artifact(stdout, root, allow_empty=True),
            "stderr": artifact(stderr, root, allow_empty=True),
        }
        captured_report = object_value(load_json(report, f"{repeat}/{selector} report"), f"{repeat}/{selector} report")
        captured_tool = object_value(captured_report.get("tool"), f"{repeat}/{selector} report.tool")
        if captured_tool.get("allocator_counter_revision") != "serialized_region_peak_v3":
            fail(f"{repeat}/{selector} report lacks serialized_region_peak_v3 allocator identity")
        with verify_stdout.open("wb") as verify_out, verify_stderr.open("wb") as verify_err:
            verification = subprocess.run(
                verify_command, cwd=repo_root, env=os.environ.copy(),
                stdout=verify_out, stderr=verify_err, check=False,
            )
        records["verify_stdout"] = artifact(verify_stdout, root, allow_empty=False)
        records["verify_stderr"] = artifact(verify_stderr, root, allow_empty=True)
        if verification.returncode != 0:
            fail(f"semantic verifier rejected {repeat}/{selector}; see {verify_stderr}")
        verification_value = verify_result(verify_stdout, f"{repeat}/{selector} verifier")
        record["verification"] = {
            "status": "pass",
            "exit_code": verification.returncode,
            "stdout_sha256": records["verify_stdout"]["sha256"],
            "stderr_sha256": records["verify_stderr"]["sha256"],
            "summary_sha256": digest_json(verification_value),
        }
        record["verification_exit_code"] = verification.returncode
        record["artifacts"] = records
    except CaptureError as error:
        record.update({"status": "failed", "error": str(error)})
        write_json(journal, record)
        raise
    record["status"] = "pass"
    write_json(journal, record)
    print(f"[{index}/{total}] pass", flush=True)
    return {
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
    try:
        shutil.copy2(source, target)
    except OSError as error:
        fail(f"cannot copy {label} seed: {error}")


def capture(
    root: Path, *, repo_root: Path | None, protocol_seed: Path | None,
    build_seed: Path | None, verifier: Path | None,
) -> None:
    root = root.expanduser().resolve()
    try:
        root.mkdir(parents=True, exist_ok=True)
    except OSError as error:
        fail(f"cannot create output root {root}: {error}")
    copy_seed(protocol_seed, root / "protocol.json", "protocol")
    copy_seed(build_seed, root / "build-candidate.json", "candidate build")
    capture_path = root / "capture.json"
    runs_path = root / "runs"
    if capture_path.exists() or runs_path.exists():
        fail(f"refusing to overwrite existing capture output: {capture_path} or {runs_path}")
    protocol = load_protocol(root)
    build = load_build(root, protocol)
    if repo_root is None:
        repo_root = Path(__file__).resolve().parents[4]
    repo_root = repo_root.expanduser().resolve()
    if not (repo_root / "tools" / "perf_abba_summary.py").is_file():
        fail(f"repository root lacks stable validators: {repo_root}")
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
        "binary": build["binary"],
        "source": {
            "revision": build["source"]["revision"],
            "identity_sha256": digest_json(build["source"]),
        },
        "cpu": CPU,
        "selectors": list(SELECTORS),
        "repeats": list(REPEAT_NAMES),
        "samples": SAMPLES,
        "warmups": WARMUPS,
        "common_flags": protocol["common_flags"],
        "verifier": {"path": str(verifier), "sha256": sha256_file(verifier)},
        "runs": [],
        "expected_run_count": len(REPEAT_NAMES) * len(SELECTORS),
    }
    write_json(capture_path, manifest)
    total = len(REPEAT_NAMES) * len(SELECTORS)
    try:
        index = 0
        for repeat in REPEAT_NAMES:
            for selector in SELECTORS:
                index += 1
                run = run_one(
                    root=root, protocol=protocol, build=build, verifier=verifier,
                    repo_root=repo_root, repeat=repeat, selector=selector,
                    index=index, total=total,
                )
                manifest["runs"].append(run)
                write_json(capture_path, manifest)
        current_source(build["source"])
        if sha256_file(Path(build["binary"]["path"])) != build["binary"]["sha256"]:
            fail("allocator binary changed after the final run")
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
    default_root = Path(__file__).resolve().parent
    parser.add_argument("--root", type=Path, default=default_root)
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
        print(f"0422 capture failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps({"status": "pass", "change": CHANGE, "runs": 4}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
