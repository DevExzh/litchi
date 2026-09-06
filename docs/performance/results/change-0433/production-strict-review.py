#!/usr/bin/env python3
"""Review the retained production strict-lint diagnostic debt.

The production all-warning-denied command is expected to retain one known
pre-existing ``ArchiveReaderKind`` diagnostic.  This helper compares only the
non-empty ``(message, file)`` diagnostic multiset against the retained ODS
reference run.  It is diagnostic evidence; a passing result does not claim
that the production strict command passed.

Generation reads the two retained check records, logs, and source manifests
under this bundle.  ``--check`` repeats the same work from those retained
artifacts and compares the result JSON byte-for-byte by parsed value.  It
does not inspect the checkout or invoke Git, so it remains usable after the
temporary build worktrees have been removed and logs have been compressed.
"""

from __future__ import annotations

import argparse
from collections import Counter
import gzip
import hashlib
import json
from pathlib import Path
import re
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 433
BASELINE_RECORD = "checks/ods-reference-strict.json"
CURRENT_RECORD = "checks/production-strict-final.json"
BASELINE_LOG = "checks/ods-reference-strict.log"
CURRENT_LOG = "checks/production-strict-final.log"
SOURCE_FILE = "crates/litchi-odf-common/src/package/model.rs"
KNOWN_MESSAGE = "large size difference between variants"
KNOWN_DIAGNOSTIC = (KNOWN_MESSAGE, SOURCE_FILE)
DIAGNOSTIC_RE = re.compile(
    r"^error: (.+)\n\s*--> ([^\n]+?):(\d+):\d+",
    re.MULTILINE,
)
SHA256_RE = re.compile(r"[0-9a-f]{64}")


class Invalid(ValueError):
    pass


def fail(message: str) -> None:
    raise Invalid(message)


def digest(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        fail(f"{path.relative_to(ROOT)}: invalid JSON: {error}")
    raise AssertionError("unreachable")


def relative_path(name: Any, label: str) -> str:
    if not isinstance(name, str) or not name:
        fail(f"{label} must be non-empty text")
    path = Path(name)
    if path.is_absolute() or ".." in path.parts:
        fail(f"{label} must be a bundle-relative path")
    return name


def read_log(name: str) -> tuple[bytes, str]:
    path = ROOT / name
    if path.is_file():
        return path.read_bytes(), name
    compressed = Path(str(path) + ".gz")
    if compressed.is_file():
        try:
            return gzip.decompress(compressed.read_bytes()), name
        except (OSError, EOFError, gzip.BadGzipFile) as error:
            fail(f"{compressed.relative_to(ROOT)}: invalid gzip stream: {error}")
    fail(f"missing {name} and {name}.gz")
    raise AssertionError("unreachable")


def sha256_text(value: Any, label: str) -> str:
    if not isinstance(value, str) or SHA256_RE.fullmatch(value) is None:
        fail(f"{label} must be a lowercase SHA-256 string")
    return value


def uint(value: Any, label: str) -> int:
    if type(value) is not int or value < 0:
        fail(f"{label} must be a non-negative integer")
    return value


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label} must be a JSON object")
    return value


def text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label} must be non-empty text")
    return value


def source_manifest(record: dict[str, Any], record_name: str, side: str) -> dict[str, Any]:
    descriptor = obj(record.get(side), f"{record_name}.{side}")
    manifest_name = relative_path(descriptor.get("path"), f"{record_name}.{side}.path")
    manifest_path = ROOT / manifest_name
    if not manifest_path.is_file():
        fail(f"missing {manifest_name} referenced by {record_name}.{side}")
    manifest_raw = manifest_path.read_bytes()
    manifest_sha = sha256_text(descriptor.get("sha256"), f"{record_name}.{side}.sha256")
    if digest(manifest_raw) != manifest_sha:
        fail(f"{record_name}.{side}.sha256 does not match {manifest_name}")
    files = uint(descriptor.get("files"), f"{record_name}.{side}.files")
    manifest = obj(load_json(manifest_path), manifest_name)
    if len(manifest) != files:
        fail(f"{record_name}.{side}.files does not match {manifest_name}")
    source_sha = sha256_text(manifest.get(SOURCE_FILE), f"{manifest_name}.{SOURCE_FILE}")
    return {
        "path": manifest_name,
        "sha256": manifest_sha,
        "bytes": len(manifest_raw),
        "files": files,
        "source_file": SOURCE_FILE,
        "source_sha256": source_sha,
    }


def record_binding(name: str, log_name: str) -> dict[str, Any]:
    record_path = ROOT / name
    if not record_path.is_file():
        fail(f"missing {name}")
    record_raw = record_path.read_bytes()
    record = obj(load_json(record_path), name)
    if record.get("change") != CHANGE:
        fail(f"{name}.change does not match {CHANGE}")
    if record.get("source_unchanged") is not True:
        fail(f"{name}.source_unchanged must be true")
    log = obj(record.get("log"), f"{name}.log")
    if log.get("path") != log_name:
        fail(f"{name}.log.path must be {log_name}")
    raw, logical_name = read_log(log_name)
    log_sha = sha256_text(log.get("sha256"), f"{name}.log.sha256")
    log_bytes = uint(log.get("bytes"), f"{name}.log.bytes")
    if digest(raw) != log_sha or len(raw) != log_bytes:
        fail(f"{name}.log metadata does not match {logical_name}")
    before = source_manifest(record, name, "source_before")
    after = source_manifest(record, name, "source_after")
    if before["source_sha256"] != after["source_sha256"]:
        fail(f"{name}: {SOURCE_FILE} changed between source_before and source_after")
    return {
        "record_path": name,
        "record_sha256": digest(record_raw),
        "record_bytes": len(record_raw),
        "revision": text(record.get("revision"), f"{name}.revision"),
        "status": text(record.get("status"), f"{name}.status"),
        "log_path": logical_name,
        "log_sha256": digest(raw),
        "log_bytes": len(raw),
        "source_before": before,
        "source_after": after,
        "source_sha256": before["source_sha256"],
    }


def diagnostic_rows(raw: bytes) -> list[dict[str, Any]]:
    text_value = raw.decode("utf-8", errors="replace")
    return [
        {"message": message, "file": path, "line": int(line)}
        for message, path, line in DIAGNOSTIC_RE.findall(text_value)
    ]


def findings(rows: list[dict[str, Any]]) -> Counter[tuple[str, str]]:
    return Counter((row["message"], row["file"]) for row in rows)


def finding_list(counter: Counter[tuple[str, str]]) -> list[dict[str, Any]]:
    return [
        {"message": message, "file": path, "count": count}
        for (message, path), count in sorted(counter.items())
    ]


def derive() -> dict[str, Any]:
    baseline_binding = record_binding(BASELINE_RECORD, BASELINE_LOG)
    current_binding = record_binding(CURRENT_RECORD, CURRENT_LOG)
    baseline_raw, _ = read_log(BASELINE_LOG)
    current_raw, _ = read_log(CURRENT_LOG)
    baseline_rows = diagnostic_rows(baseline_raw)
    current_rows = diagnostic_rows(current_raw)
    baseline_findings = findings(baseline_rows)
    current_findings = findings(current_rows)
    if not baseline_findings or not current_findings:
        fail("strict diagnostic extraction must be non-empty for both records")
    expected = Counter({KNOWN_DIAGNOSTIC: 1})
    if baseline_findings != expected:
        fail(f"retained baseline diagnostics are not the one known finding: {finding_list(baseline_findings)}")
    if current_findings != expected:
        fail(f"production diagnostics are not the one known finding: {finding_list(current_findings)}")
    if baseline_findings != current_findings:
        fail("production diagnostic multiset differs from retained baseline")
    baseline_source = baseline_binding["source_sha256"]
    current_source = current_binding["source_sha256"]
    if baseline_source != current_source:
        fail(f"{SOURCE_FILE} hash differs between baseline and production records")
    return {
        "status": "pass",
        "change": CHANGE,
        "scope": "diagnostic-only; does not claim a passing production strict gate",
        "baseline": baseline_binding,
        "current": current_binding,
        "same_message_and_source_file_multiset": True,
        "known_preexisting_diagnostic": {
            "message": KNOWN_MESSAGE,
            "file": SOURCE_FILE,
            "count": 1,
        },
        "findings": finding_list(current_findings),
        "source_file": SOURCE_FILE,
        "source_sha256_unchanged": True,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true", help="replay retained artifacts without Git")
    args = parser.parse_args()
    target = ROOT / "checks/production-strict-review.json"
    try:
        result = derive()
        if args.check:
            if not target.is_file():
                fail("checks/production-strict-review.json is missing")
            if load_json(target) != result:
                fail("retained production-strict-review.json differs from logs/source bindings")
        else:
            if target.exists():
                fail("refusing to overwrite checks/production-strict-review.json")
            target.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": "pass", "diagnostic_only": True}, sort_keys=True))
        return 0
    except (OSError, UnicodeError, ValueError, AssertionError) as error:
        print(f"INVALID: {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
