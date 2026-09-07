#!/usr/bin/env python3
"""Run the immutable 0465 finalization phases.

Run ``--precleanup`` against the live workspace with `verify.py`, reseal
the bundle, run ``--portable`` against a fresh copy using the verifier's
default portable mode, reseal again, then run ``--cleanup``.
Receipts are append-only observations.  This driver never rebuilds the root
seal; the caller owns each required reseal between phases and after cleanup.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
from typing import Any


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3] if len(ROOT.parents) > 3 else ROOT
TASK = Path("/tmp/litchi-goal-0465")
VERIFIER = ROOT / "verify.py"
PRE = ROOT / "precleanup.json"
PORTABLE = ROOT / "portable-verification.json"
INVENTORY = ROOT / "temporary-artifacts.json"
CLEANUP = ROOT / "cleanup.json"
SHA256_RE = re.compile(r"[0-9a-f]{64}\Z")
ARCHIVE_AUDITS = ()


def fail(message: str) -> None:
    raise SystemExit(message)


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def sha_file(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def regular(path: Path, label: str) -> Path:
    if not path.is_file() or path.is_symlink():
        fail(f"{label} is missing, symlinked, or not regular: {path}")
    return path


def identity(path: Path, label: str, *, relative_to: Path | None = None) -> dict[str, Any]:
    regular(path, label)
    name = path.name if relative_to is None else path.relative_to(relative_to).as_posix()
    return {"path": name, "bytes": path.stat().st_size, "sha256": sha_file(path)}


def identity_row(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict) or value.get("path") != "SHA256SUMS":
        fail(f"{label} has no SHA256SUMS identity")
    size = value.get("bytes")
    digest = value.get("sha256")
    if isinstance(size, bool) or not isinstance(size, int) or size < 1 or not isinstance(digest, str) or SHA256_RE.fullmatch(digest) is None:
        fail(f"{label} has an invalid identity")
    return {"path": "SHA256SUMS", "bytes": size, "sha256": digest}


def check_seal(root: Path, label: str = "bundle seal") -> dict[str, Any]:
    """Require SHA256SUMS to cover every regular bundle member exactly once."""
    reject_symlinks(root)
    seal = regular(root / "SHA256SUMS", label)
    try:
        lines = seal.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"{label} cannot be read: {error}")
    members: dict[str, str] = {}
    for index, line in enumerate(lines, 1):
        fields = line.split("  ", 1)
        if len(fields) != 2 or SHA256_RE.fullmatch(fields[0]) is None:
            fail(f"{label} has malformed line {index}")
        name = fields[1]
        relative = Path(name)
        if not name or relative.is_absolute() or ".." in relative.parts or relative.as_posix() != name:
            fail(f"{label} has unsafe member {name!r}")
        if name == "SHA256SUMS" or name in members:
            fail(f"{label} has duplicate or self member {name!r}")
        member = root / relative
        regular(member, f"{label} member {name}")
        if sha_file(member) != fields[0]:
            fail(f"{label} hash differs for {name}")
        members[name] = fields[0]
    actual = {
        path.relative_to(root).as_posix()
        for path in root.rglob("*")
        if path.is_file() and path != seal
    }
    if set(members) != actual:
        missing = sorted(actual - set(members))
        extra = sorted(set(members) - actual)
        fail(f"{label} coverage differs (missing={missing[:3]}, extra={extra[:3]})")
    return identity(seal, label, relative_to=root) | {"members": len(members)}


def load(path: Path, label: str) -> dict[str, Any]:
    regular(path, label)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label} is invalid JSON: {error}")
    if not isinstance(value, dict):
        fail(f"{label} must contain an object")
    return value


def reject_symlinks(root: Path) -> None:
    if root.is_symlink():
        fail(f"{root} is a symlink")
    for directory, names, files in os.walk(root, followlinks=False):
        for name in (*names, *files):
            path = Path(directory) / name
            if path.is_symlink():
                fail(f"{root} contains a symlink: {path.relative_to(root)}")


def parse_output(raw: str) -> Any:
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None


def write_once(path: Path, value: dict[str, Any]) -> None:
    if path.exists():
        fail(f"refusing to overwrite retained receipt: {path}")
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, ensure_ascii=False, indent=2)
        stream.write("\n")


def proof_ok(value: dict[str, Any], mode: str) -> None:
    if value.get("schema") not in {
        "litchi-0465-precleanup-v1",
        "litchi-0465-portable-verification-v1",
    }:
        fail("proof receipt schema differs")
    if value.get("change") != 465 or value.get("status") != "pass" or value.get("exit_code") != 0:
        fail("proof receipt is not a passing 0465 result")
    output = value.get("verifier_output")
    if not isinstance(output, dict) or output.get("status") != "pass":
        fail(f"proof receipt does not contain a passing {mode} verifier output")
    if mode == "precleanup" and output.get("precleanup") is not True:
        fail("proof receipt does not contain a passing precleanup verifier output")
    if mode == "portable" and output.get("portable") not in (None, True):
        fail("proof receipt does not contain a default portable verifier output")


def verifier_identity() -> tuple[str, str]:
    regular(VERIFIER, "root verifier")
    return sha_file(VERIFIER), sha_file(Path(__file__).resolve())


def run_precleanup() -> int:
    write_path = PRE
    if write_path.exists():
        fail(f"refusing to overwrite retained receipt: {write_path}")
    seal_identity = check_seal(ROOT)
    verifier_sha, driver_sha = verifier_identity()
    command = [sys.executable, "-B", str(VERIFIER), "--precleanup"]
    record: dict[str, Any] = {
        "schema": "litchi-0465-precleanup-v1",
        "change": 465,
        "status": "failed",
        "argv": command,
        "cwd": str(REPO),
        "verifier_path": VERIFIER.name,
        "verifier_sha256": verifier_sha,
        "driver_sha256": driver_sha,
        "seal": seal_identity,
        "started_utc": now(),
    }
    try:
        result = subprocess.run(command, cwd=REPO, capture_output=True, text=True, errors="replace")
        record.update(
            exit_code=result.returncode,
            stdout=result.stdout,
            stderr=result.stderr,
            verifier_output=parse_output(result.stdout),
            finished_utc=now(),
        )
        if result.returncode == 0 and isinstance(record["verifier_output"], dict) and record["verifier_output"].get("status") == "pass" and record["verifier_output"].get("precleanup") is True:
            record["status"] = "pass"
    except OSError as error:
        record.update(exit_code=None, stdout="", stderr=str(error), verifier_output=None, finished_utc=now())
    write_once(write_path, record)
    print(json.dumps({key: record[key] for key in ("status", "exit_code", "verifier_sha256")}, sort_keys=True))
    return 0 if record["status"] == "pass" else 1


def run_portable() -> int:
    if PORTABLE.exists():
        fail(f"refusing to overwrite retained receipt: {PORTABLE}")
    pre = load(PRE, "precleanup receipt")
    proof_ok(pre, "precleanup")
    verifier_sha, driver_sha = verifier_identity()
    if pre.get("verifier_path") != VERIFIER.name or pre.get("verifier_sha256") != verifier_sha or pre.get("driver_sha256") != driver_sha:
        fail("precleanup receipt is bound to a different root verifier")
    seal_identity = check_seal(ROOT)
    started = now()
    command: list[str] = []
    result: subprocess.CompletedProcess[str] | None = None
    output: Any = None
    temporary_directory_absent = False
    copied_seal_before: dict[str, Any] | None = None
    copied_seal_after: dict[str, Any] | None = None
    copied_verifier_sha: str | None = None
    copied_cwd: str | None = None
    error_text = ""
    try:
        with tempfile.TemporaryDirectory(prefix="litchi-0465-portable-", dir="/tmp") as directory:
            parent = Path(directory)
            target = parent / ROOT.name
            shutil.copytree(ROOT, target, symlinks=False)
            reject_symlinks(target)
            copied_seal_before = check_seal(target, "copied root seal")
            copied_verifier = regular(target / VERIFIER.name, "copied root verifier")
            command = [sys.executable, "-B", str(copied_verifier)]
            environment = os.environ.copy()
            environment["PYTHONPATH"] = ""
            environment["PYTHONDONTWRITEBYTECODE"] = "1"
            result = subprocess.run(command, cwd=target, env=environment, capture_output=True, text=True, errors="replace")
            output = parse_output(result.stdout)
            copied_seal_after = check_seal(target, "copied root seal after verification")
            copied_verifier_sha = sha_file(copied_verifier)
            copied_cwd = str(target)
        temporary_directory_absent = not parent.exists()
    except OSError as error:
        error_text = str(error)
    except SystemExit as error:
        error_text = str(error)
    record: dict[str, Any] = {
        "schema": "litchi-0465-portable-verification-v1",
        "change": 465,
        "status": "pass" if result is not None and result.returncode == 0 and isinstance(output, dict) and output.get("status") == "pass" and output.get("portable") in (None, True) and copied_seal_before == seal_identity and copied_seal_after == copied_seal_before and temporary_directory_absent else "failed",
        "argv": command,
        "cwd": copied_cwd,
        "verifier_path": VERIFIER.name,
        "verifier_sha256": verifier_sha,
        "copied_verifier_sha256": copied_verifier_sha,
        "driver_sha256": driver_sha,
        "root_seal": seal_identity,
        "copy_seal": copied_seal_before,
        "copy_seal_after": copied_seal_after,
        "copy_seal_unchanged": copied_seal_after == copied_seal_before if copied_seal_after is not None else None,
        "started_utc": started,
        "finished_utc": now(),
        "exit_code": result.returncode if result is not None else None,
        "stdout": result.stdout if result is not None else "",
        "stderr": result.stderr if result is not None else error_text,
        "verifier_output": output,
        "temporary_directory_absent": temporary_directory_absent,
        "mode": "complete sealed change-0465 bundle copied to a fresh temporary directory; verifier default is portable",
    }
    write_once(PORTABLE, record)
    print(json.dumps({key: record[key] for key in ("status", "exit_code", "copy_seal_unchanged", "temporary_directory_absent")}, sort_keys=True))
    return 0 if record["status"] == "pass" else 1


def inventory(root: Path) -> tuple[list[dict[str, Any]], int, int]:
    rows: list[dict[str, Any]] = []
    total = 0
    for directory, names, files in os.walk(root, followlinks=False):
        for name in (*names, *files):
            path = Path(directory) / name
            if path.is_symlink():
                fail(f"owned task tree contains a symlink: {path}")
        for name in sorted(files):
            path = Path(directory) / name
            if not path.is_file() or path.is_symlink():
                fail(f"owned task entry is not a regular file: {path}")
            size = path.stat().st_size
            rows.append({"path": str(path.relative_to(root)), "bytes": size, "sha256": sha_file(path)})
            total += size
    rows.sort(key=lambda value: value["path"])
    return rows, len(rows), total


def run_cleanup() -> int:
    if INVENTORY.exists() or CLEANUP.exists():
        fail("refusing to overwrite retained cleanup receipts")
    pre = load(PRE, "precleanup receipt")
    portable = load(PORTABLE, "portable verification receipt")
    proof_ok(pre, "precleanup")
    proof_ok(portable, "portable")
    verifier_sha, driver_sha = verifier_identity()
    for label, value in (("precleanup", pre), ("portable", portable)):
        if value.get("verifier_path") != VERIFIER.name or value.get("verifier_sha256") != verifier_sha or value.get("driver_sha256") != driver_sha:
            fail(f"{label} proof is bound to a different root verifier")
    if portable.get("temporary_directory_absent") is not True or portable.get("copy_seal_unchanged") is not True:
        fail("portable proof did not remove an unchanged temporary copy")
    copy_seal = identity_row(portable.get("copy_seal"), "portable copied seal")
    copy_seal_after = identity_row(portable.get("copy_seal_after"), "portable copied seal after verification")
    if copy_seal_after != copy_seal:
        fail("portable proof copied seal identity is inconsistent")
    seal_identity = check_seal(ROOT)
    receipt_identity = {
        "precleanup": identity(PRE, "precleanup receipt", relative_to=ROOT),
        "portable": identity(PORTABLE, "portable verification receipt", relative_to=ROOT),
    }
    if TASK != Path("/tmp/litchi-goal-0465") or not TASK.is_dir() or TASK.is_symlink():
        fail(f"refusing to remove unexpected owned task root: {TASK}")
    reject_symlinks(TASK)
    rows, count, total = inventory(TASK)
    if count == 0:
        fail("owned task root has no files to inventory")
    audit_records: list[dict[str, Any]] = []
    for archive, temporary in ARCHIVE_AUDITS:
        archive_identity = identity(archive, f"archived audit {archive.name}", relative_to=ROOT)
        temporary_identity = identity(temporary, f"temporary audit {temporary.name}")
        if archive_identity["bytes"] != temporary_identity["bytes"] or archive_identity["sha256"] != temporary_identity["sha256"]:
            fail(f"temporary audit does not match archived copy: {temporary}")
        audit_records.append({
            "archive": archive_identity,
            "temporary_path": str(temporary),
            "temporary": {"bytes": temporary_identity["bytes"], "sha256": temporary_identity["sha256"]},
        })
    inventory_record = {
        "schema": "litchi-0465-temporary-artifacts-v1",
        "change": 465,
        "task": str(TASK),
        "files": count,
        "bytes": total,
        "artifacts": rows,
        "receipts": receipt_identity,
        "seal": seal_identity,
        "audit_archives": audit_records,
    }
    write_once(INVENTORY, inventory_record)
    for _, temporary in ARCHIVE_AUDITS:
        temporary.unlink()
    if any(temporary.exists() or temporary.is_symlink() for _, temporary in ARCHIVE_AUDITS):
        fail("owned temporary audit remains after cleanup")
    shutil.rmtree(TASK)
    if TASK.exists() or TASK.is_symlink():
        fail("owned task root remains after cleanup")
    cleanup_record = {
        "schema": "litchi-0465-cleanup-v1",
        "change": 465,
        "status": "pass",
        "task": str(TASK),
        "temporary_directory_absent": True,
        "temporary_audits_absent": True,
        "files_removed": count,
        "bytes_removed": total,
        "inventory": identity(INVENTORY, "temporary-artifacts receipt", relative_to=ROOT),
        "receipts": receipt_identity,
        "seal": seal_identity,
        "audit_archives": audit_records,
        "verifier_path": VERIFIER.name,
        "verifier_sha256": verifier_sha,
        "driver_sha256": driver_sha,
    }
    write_once(CLEANUP, cleanup_record)
    print(json.dumps({key: cleanup_record[key] for key in ("status", "files_removed", "bytes_removed", "temporary_directory_absent")}, sort_keys=True))
    return 0


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--precleanup", action="store_true", help="run live-copy and current-source proof")
    modes.add_argument("--portable", action="store_true", help="copy and replay the sealed bundle")
    modes.add_argument("--cleanup", action="store_true", help="inventory and remove the owned task tree")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    if args.precleanup:
        return run_precleanup()
    if args.portable:
        return run_portable()
    return run_cleanup()


if __name__ == "__main__":
    raise SystemExit(main())
