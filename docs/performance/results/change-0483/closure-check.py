#!/usr/bin/env python3
"""Run the terminal copied-bundle and corruption checks for change 0483.

This is a documentary closure helper rather than a protocol driver.  It is
intentionally absent from ``common.py`` and ``analyze.py``'s driver lists, so
its own hash does not become a build or validation input.  The command is
guarded by ``--execute`` because it creates a private copied bundle and
mutates only that copy and the requested ``evidence-validation`` output
directory.

The source bundle is verified first while it is still live.  A complete copy
is then checked with the copied verifier, including after the referenced
temporary binaries have disappeared.  Negative structured-evidence cases and
negative SHA256SUMS cases are run against independent copies.  The successful
copy is sealed and replayed before all command output and JSON receipts are
published to ``evidence-validation``.  The real source bundle is never
sealed by this helper; the coordinator seals it afterwards, once this audit
record is complete.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
import subprocess
import sys
import tempfile
from typing import Any, NoReturn


ROOT = Path(__file__).resolve().parent
MANIFEST_NAME = "SHA256SUMS"
DEFAULT_WORK_DIR = Path("/home/zhuhe/.cache/litchi-goal-0483/closure-check")
OUTPUT_DIRECTORY_NAME = "evidence-validation"
SCHEMA = "docx-tail-append-closure-check-v1"
CASE_SCHEMA = "docx-tail-append-closure-case-v1"


class ClosureError(ValueError):
    """Raised when a closure precondition or expected corruption check fails."""


def fail(message: str) -> NoReturn:
    raise ClosureError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def digest_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def metadata(value: bytes) -> dict[str, int | str]:
    return {"bytes": len(value), "sha256": digest_bytes(value)}


def _regular(path: Path, label: str) -> None:
    try:
        info = path.stat(follow_symlinks=False)
    except OSError as error:
        fail(f"{label}: cannot stat: {error}")
    require(stat.S_ISREG(info.st_mode), f"{label}: regular file required")


def _safe_relative(value: str, label: str) -> str:
    require(isinstance(value, str) and value, f"{label}: path is missing")
    require("\\" not in value and "\x00" not in value, f"{label}: unsafe path")
    require("\r" not in value and "\n" not in value, f"{label}: newline in path")
    parts = value.split("/")
    require(not value.startswith("/") and all(part not in {"", ".", ".."} for part in parts), f"{label}: path escapes root")
    return value


def _bundle_root(value: Path | str) -> Path:
    path = Path(value)
    require(not path.is_symlink() and path.is_dir(), "bundle root must be a regular directory")
    try:
        return path.resolve(strict=True)
    except OSError as error:
        fail(f"bundle root cannot be resolved: {error}")


def _read_json(root: Path, relative: str) -> Any:
    safe = _safe_relative(relative, "JSON path")
    path = root.joinpath(*safe.split("/"))
    _regular(path, safe)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{safe}: invalid JSON: {error}")


def _write_exclusive(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError:
        fail(f"refusing to replace existing output: {path}")
    except OSError as error:
        fail(f"cannot write {path}: {error}")


def _copy_source(source: Path, destination: Path) -> None:
    """Copy a complete bundle after rejecting links and special files."""

    require(not destination.exists(), f"copy destination already exists: {destination}")
    require(source != destination and source not in destination.parents, "copy destination must be outside the source bundle")
    source_files: list[Path] = []
    for directory, directories, files in os.walk(source, followlinks=False):
        directory_path = Path(directory)
        for name in sorted(directories + files):
            path = directory_path / name
            relative = path.relative_to(source).as_posix()
            _safe_relative(relative, "source path")
            try:
                info = path.stat(follow_symlinks=False)
            except OSError as error:
                fail(f"{relative}: cannot stat: {error}")
            mode = info.st_mode
            require(not stat.S_ISLNK(mode), f"{relative}: symlinks are forbidden")
            require(not stat.S_ISSOCK(mode) and not stat.S_ISFIFO(mode) and not stat.S_ISCHR(mode) and not stat.S_ISBLK(mode), f"{relative}: special files are forbidden")
            if stat.S_ISDIR(mode):
                require(name != "__pycache__" and "__pycache__" not in path.relative_to(source).parts, f"{relative}: __pycache__ is forbidden")
            else:
                require(stat.S_ISREG(mode), f"{relative}: regular file required")
                require(name != "__pycache__" and "__pycache__" not in path.relative_to(source).parts, f"{relative}: __pycache__ is forbidden")
                require(not name.endswith(".pyc"), f"{relative}: .pyc files are forbidden")
                source_files.append(path)
    try:
        shutil.copytree(source, destination, symlinks=False, copy_function=shutil.copy2)
    except OSError as error:
        fail(f"cannot copy bundle: {error}")
    # The preflight above rejects links, but check the result as well so a
    # race during copying cannot silently turn into an unsealed file boundary.
    for path in destination.rglob("*"):
        require(not path.is_symlink(), f"copied path is a symlink: {path}")
    require(len(source_files) > 0, "source bundle has no regular files")


def _run(root: Path, script: str, *arguments: str) -> dict[str, Any]:
    script_path = root / script
    _regular(script_path, script)
    command = [str(Path(sys.executable).resolve()), "-B", str(script_path), "--root", str(root), *arguments]
    started = now()
    try:
        completed = subprocess.run(command, cwd=root, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
    except OSError as error:
        fail(f"cannot execute {script}: {error}")
    finished = now()
    return {
        "command": command,
        "cwd": str(root),
        "started_utc": started,
        "finished_utc": finished,
        "exit_code": completed.returncode,
        "stdout_bytes": bytes(completed.stdout),
        "stderr_bytes": bytes(completed.stderr),
    }


def _stdout_json(result: dict[str, Any], label: str) -> dict[str, Any]:
    try:
        value = json.loads(result["stdout_bytes"].decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        fail(f"{label} did not emit JSON: {error}")
    require(isinstance(value, dict), f"{label} JSON result is not an object")
    return value


def _record_result(
    result: dict[str, Any],
    output_directory: Path,
    label: str,
    expected_acceptance: bool,
    mutation: str | None = None,
) -> dict[str, Any]:
    """Persist one command's streams and return its JSON-safe receipt."""

    _safe_relative(label, "case label")
    stdout_name = f"{label}.stdout"
    stderr_name = f"{label}.stderr"
    json_name = f"{label}.json"
    stdout = result.pop("stdout_bytes")
    stderr = result.pop("stderr_bytes")
    _write_exclusive(output_directory / stdout_name, stdout)
    _write_exclusive(output_directory / stderr_name, stderr)
    receipt: dict[str, Any] = {
        "schema": CASE_SCHEMA,
        "case": label,
        "expected_acceptance": expected_acceptance,
        "accepted": result["exit_code"] == 0,
        "mutation": mutation,
        **result,
        "stdout": {"path": f"{OUTPUT_DIRECTORY_NAME}/{stdout_name}", **metadata(stdout)},
        "stderr": {"path": f"{OUTPUT_DIRECTORY_NAME}/{stderr_name}", **metadata(stderr)},
    }
    _write_exclusive(output_directory / json_name, (json.dumps(receipt, indent=2, sort_keys=True) + "\n").encode("utf-8"))
    require(receipt["accepted"] is expected_acceptance, f"{label}: verifier acceptance differed from expectation")
    return receipt


def _copy_case(source: Path, work: Path, name: str) -> Path:
    destination = work / name
    _copy_source(source, destination)
    return destination


def _protocol_capture_path(root: Path) -> str:
    captures = sorted((root / "captures").glob("*.report.json"))
    require(captures, "completed bundle has no capture reports")
    return captures[0].relative_to(root).as_posix()


def _required_gate_path(protocol: dict[str, Any]) -> str:
    validation = protocol.get("validation")
    require(isinstance(validation, dict), "protocol.validation is missing")
    labels = validation.get("required_labels")
    require(isinstance(labels, list) and labels and isinstance(labels[0], str), "protocol.validation.required_labels is missing")
    label = _safe_relative(labels[0], "required gate label")
    require("/" not in label, "required gate label is unsafe")
    return f"validation/{label}.json"


def _fuzz_chain_path(protocol: dict[str, Any]) -> str:
    fuzz = protocol.get("fuzz")
    require(isinstance(fuzz, dict), "protocol.fuzz is missing")
    receipts = fuzz.get("receipts")
    require(isinstance(receipts, list), "protocol.fuzz.receipts is missing")
    for reference in receipts:
        if isinstance(reference, dict) and reference.get("kind") == "smoke" and isinstance(reference.get("path"), str):
            return _safe_relative(reference["path"], "fuzz smoke receipt path")
    fail("protocol.fuzz has no smoke receipt")


def _temporary_binary_paths(root: Path, protocol: dict[str, Any]) -> list[str]:
    """Collect the copied executable paths that verify.py treats as optional."""

    paths: set[str] = set()
    binaries = _read_json(root, "builds/binaries-accepted.json")
    require(isinstance(binaries, dict) and isinstance(binaries.get("binaries"), dict), "accepted binary custody is malformed")
    for item in binaries["binaries"].values():
        require(isinstance(item, dict) and isinstance(item.get("path"), str), "accepted binary path is missing")
        paths.add(item["path"])
    fuzz = protocol.get("fuzz")
    require(isinstance(fuzz, dict) and isinstance(fuzz.get("receipts"), list), "protocol.fuzz.receipts is missing")
    for reference in fuzz["receipts"]:
        if not isinstance(reference, dict) or reference.get("kind") != "build":
            continue
        relative = reference.get("path")
        require(isinstance(relative, str), "fuzz build receipt path is missing")
        receipt = _read_json(root, relative)
        binary = receipt.get("binary") if isinstance(receipt, dict) else None
        require(isinstance(binary, dict) and isinstance(binary.get("path"), str), "fuzz build binary path is missing")
        paths.add(binary["path"])
    return sorted(paths)


def _tamper_summary(path: Path) -> str:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"summary cannot be read for tamper case: {error}")
    require(isinstance(value, dict) and isinstance(value.get("sample_count"), int), "summary sample_count is unavailable")
    value["sample_count"] += 1
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return "summary.json.sample_count"


def _run_structured_negative(
    baseline: Path,
    work: Path,
    name: str,
    relative_target: str,
    mutate: Any,
) -> dict[str, Any]:
    case_root = _copy_case(baseline, work, name)
    target = case_root.joinpath(*relative_target.split("/"))
    require(target.exists() and not target.is_symlink(), f"{name}: mutation target is missing: {relative_target}")
    mutation = mutate(target) or relative_target
    result = _run(case_root, "verify.py")
    result["mutation"] = mutation
    return result


def _run_seal_negative(
    sealed: Path,
    work: Path,
    name: str,
    mutate: Any,
) -> dict[str, Any]:
    case_root = _copy_case(sealed, work, name)
    mutation = mutate(case_root)
    result = _run(case_root, "seal.py", "--verify")
    result["mutation"] = mutation
    return result


def _write_summary(output_directory: Path, value: dict[str, Any]) -> None:
    _write_exclusive(output_directory / "closure-check.json", (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8"))


def _publish_staging(staging: Path, output: Path) -> None:
    require(not output.exists(), f"evidence output directory already exists: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    try:
        shutil.copytree(staging, output, symlinks=False, copy_function=shutil.copy2)
    except OSError as error:
        fail(f"cannot publish evidence outputs: {error}")


def execute(source_root: Path, output_directory: Path, work_parent: Path) -> dict[str, Any]:
    """Perform the closure audit, retaining only its final evidence outputs."""

    source = _bundle_root(source_root)
    output = output_directory.resolve()
    require(not (source / MANIFEST_NAME).exists() and not (source / MANIFEST_NAME).is_symlink(), "source bundle is already sealed")
    require(not output.exists(), f"evidence output directory already exists: {output}")
    work_parent = work_parent.resolve()
    work_parent.mkdir(parents=True, exist_ok=True)
    work_root = Path(tempfile.mkdtemp(prefix="run-", dir=work_parent))
    staging = work_root / OUTPUT_DIRECTORY_NAME
    staging.mkdir()
    started = now()
    receipts: dict[str, dict[str, Any]] = {}
    try:
        # This is deliberately the first bundle operation.  It proves that
        # the live completed evidence is valid before any copy or mutation.
        live_result = _run(source, "verify.py")
        receipts["live-verify"] = _record_result(live_result, staging, "live-verify", True)
        require(live_result["exit_code"] == 0, "live structured verification failed")

        protocol = _read_json(source, "protocol.json")
        temporary_binaries = _temporary_binary_paths(source, protocol)
        require(temporary_binaries and all(not Path(path).exists() for path in temporary_binaries), "temporary runtime binaries are still present")
        _copy_source(source, work_root / "portable")
        portable = work_root / "portable"
        copied_result = _run(portable, "verify.py")
        copied_stdout = copied_result["stdout_bytes"]
        receipts["portable-verify"] = _record_result(copied_result, staging, "portable-verify", True)
        require(copied_result["exit_code"] == 0, "copied structured verification failed")
        copied_json = _stdout_json({"stdout_bytes": copied_stdout}, "copied verifier")
        require(isinstance(copied_json, dict), "copied verifier result is missing")
        require(copied_json.get("temporary_binaries_required") is False, "copied verification did not prove runtime binaries are unnecessary")

        # Structured verifier negative cases use the unsealed portable copy.
        structured_cases = [
            ("tampered-summary", "summary.json", _tamper_summary),
            (
                "missing-capture",
                _protocol_capture_path(portable),
                lambda path: path.unlink(),
            ),
            (
                "missing-required-gate",
                _required_gate_path(protocol),
                lambda path: path.unlink(),
            ),
            (
                "missing-fuzz-chain",
                _fuzz_chain_path(protocol),
                lambda path: path.unlink(),
            ),
        ]
        for name, target, mutate in structured_cases:
            result = _run_structured_negative(portable, work_root, name, target, mutate)
            receipts[name] = _record_result(result, staging, name, False)

        # First create and replay a seal on a separate copy.  The source
        # bundle remains unsealed for the coordinator's final seal operation.
        sealed = _copy_case(portable, work_root, "sealed")
        seal_result = _run(sealed, "seal.py", "--seal")
        seal_stdout = seal_result["stdout_bytes"]
        receipts["seal-copy"] = _record_result(seal_result, staging, "seal-copy", True)
        require(seal_result["exit_code"] == 0, "copied bundle could not be sealed")
        sealed_json = _stdout_json({"stdout_bytes": seal_stdout}, "copied seal")
        require(sealed_json.get("mode") == "seal" and sealed_json.get("seal", {}).get("status") == "pass", "copied seal result is incomplete")
        sealed_verify = _run(sealed, "seal.py", "--verify")
        replay_stdout = sealed_verify["stdout_bytes"]
        receipts["sealed-verify"] = _record_result(sealed_verify, staging, "sealed-verify", True)
        require(sealed_verify["exit_code"] == 0, "sealed copied bundle could not be replayed")
        replay_json = _stdout_json({"stdout_bytes": replay_stdout}, "copied seal replay")
        require(replay_json.get("mode") == "verify" and replay_json.get("seal", {}).get("status") == "pass", "copied seal replay result is incomplete")

        def extra_file(root: Path) -> str:
            path = root / "extra.raw"
            path.write_bytes(b"unmanifested extra file\n")
            return "extra.raw"

        def missing_file(root: Path) -> str:
            path = root / "summary.json"
            path.unlink()
            return "summary.json"

        def raw_file(root: Path) -> str:
            path = root / "summary.json"
            original = path.read_bytes()
            path.write_bytes(original + b"\x00")
            return "summary.json"

        def symlink_file(root: Path) -> str:
            path = root / "symlink.raw"
            try:
                path.symlink_to(root / "summary.json")
            except OSError as error:
                fail(f"symlink corruption case is unavailable: {error}")
            return "symlink.raw"

        for name, mutate in (
            ("seal-extra-file", extra_file),
            ("seal-missing-file", missing_file),
            ("seal-raw-file", raw_file),
            ("seal-symlink", symlink_file),
        ):
            result = _run_seal_negative(sealed, work_root, name, mutate)
            receipts[name] = _record_result(result, staging, name, False)

        output_receipts = receipts
        finished = now()
        summary = {
            "schema": SCHEMA,
            "status": "pass",
            "started_utc": started,
            "finished_utc": finished,
            "source_root": str(source),
            "portable_copy": {
                "runtime_binaries_required": False,
                "temporary_binary_paths": temporary_binaries,
                "temporary_binaries_absent": True,
                "temporary_root": str(work_root),
                "removed_after_run": True,
            },
            "cases": output_receipts,
            "case_count": len(output_receipts),
            "negative_case_count": sum(1 for item in output_receipts.values() if item.get("expected_acceptance") is False),
            "source_manifest_created": False,
        }
        _write_summary(staging, summary)
        _publish_staging(staging, output)
        return summary
    except Exception as error:
        # Preserve command streams and receipts even when a precondition or an
        # expected-negative assertion fails before the normal summary exists.
        failure = {
            "schema": SCHEMA,
            "status": "fail",
            "started_utc": started,
            "finished_utc": now(),
            "source_root": str(source),
            "error": str(error),
            "cases": receipts,
            "case_count": len(receipts),
        }
        try:
            _write_exclusive(staging / "closure-check-failure.json", (json.dumps(failure, indent=2, sort_keys=True) + "\n").encode("utf-8"))
            _publish_staging(staging, output)
        except Exception:
            pass
        raise
    finally:
        shutil.rmtree(work_root, ignore_errors=False)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--execute", action="store_true", help="confirm that private copies and evidence outputs may be created")
    parser.add_argument("--root", type=Path, default=ROOT, help="live completed evidence-bundle root")
    parser.add_argument("--output-dir", type=Path, help="evidence-validation output directory (default: ROOT/evidence-validation)")
    parser.add_argument("--work-dir", type=Path, default=DEFAULT_WORK_DIR, help="owned parent directory for the temporary copied bundle")
    options = parser.parse_args()
    if not options.execute:
        parser.error("pass --execute only after the summary exists and the coordinator has authorized the closure run")
    output = options.output_dir or options.root / OUTPUT_DIRECTORY_NAME
    try:
        result = execute(options.root, output, options.work_dir)
    except ClosureError as error:
        print(f"closure check failed: {error}")
        raise SystemExit(1)
    print(json.dumps(result, sort_keys=True))


if __name__ == "__main__":
    main()
