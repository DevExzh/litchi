#!/usr/bin/env python3
"""Recheck the sealed 0890 baseline source and input identities."""

from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import stat


PACKET = Path(__file__).resolve().parent
RECEIPT_PATH = PACKET / "attempts/baseline-f790c2228c-retry1/baseline-receipt.json"
OUTPUT = PACKET / "attempts/baseline-f790c2228c-retry1/post-baseline-source-verification.json"
RUNNER_PATH = PACKET / "baseline_runner.py"


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def load_runner():
    spec = importlib.util.spec_from_file_location("baseline_runner_0890", RUNNER_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load runner: {RUNNER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def file_identity(path: Path) -> dict[str, object]:
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "mode": oct(stat.S_IMODE(path.stat().st_mode)),
        "sha256": digest(path),
    }


def main() -> int:
    receipt = json.loads(RECEIPT_PATH.read_text())
    if receipt.get("status") != "pass":
        raise SystemExit("successful baseline receipt is not passing")
    runner = load_runner()
    current_source = runner.source_manifest()
    current_fixtures = [runner.fixture_identity(path) for path in runner.FIXTURES]
    current_binary = runner.identity(runner.BINARY)
    current_copy = runner.identity(runner.BASELINE_COPY)
    if current_source != receipt.get("source"):
        raise SystemExit("source manifest differs from the sealed baseline receipt")
    if current_fixtures != receipt.get("fixtures"):
        raise SystemExit("fixture identities differ from the sealed baseline receipt")
    if current_binary != receipt.get("binary"):
        raise SystemExit("target binary identity differs from the sealed baseline receipt")
    if current_copy != receipt.get("baseline_copy"):
        raise SystemExit("baseline executable copy differs from the sealed baseline receipt")

    lock_receipt_path = RECEIPT_PATH.parent / "root-lock-custody.json"
    lock_receipt = json.loads(lock_receipt_path.read_text())
    original_lock = Path(lock_receipt["original"]["path"])
    custody_lock = Path(lock_receipt["custody"]["path"])
    original_expected = lock_receipt["original"]
    custody_expected = {
        key: lock_receipt["custody"][key]
        for key in ("path", "bytes", "mode", "sha256")
    }
    if file_identity(original_lock) != original_expected:
        raise SystemExit("original root Cargo.lock changed after custody")
    if file_identity(custody_lock) != custody_expected:
        raise SystemExit("root Cargo.lock custody copy changed")
    if original_lock.read_bytes() != custody_lock.read_bytes():
        raise SystemExit("root Cargo.lock custody bytes differ")

    result = {
        "schema": "litchi.mce-0890-post-baseline-source-verification.v1",
        "status": "pass",
        "baseline_receipt": {
            "path": str(RECEIPT_PATH.relative_to(PACKET)),
            "sha256": digest(RECEIPT_PATH),
        },
        "source_manifest_rechecked": True,
        "source_files": current_source["file_count"],
        "fixtures_rechecked": True,
        "binary_rechecked": True,
        "baseline_copy_rechecked": True,
        "root_lock_custody_rechecked": True,
        "source_root": str(runner.SOURCE),
        "target": str(runner.TARGET),
        "fixtures": current_fixtures,
        "binary": current_binary,
        "baseline_copy": current_copy,
        "root_lock_custody": {
            "path": str(lock_receipt_path.relative_to(PACKET)),
            "sha256": digest(lock_receipt_path),
        },
    }
    OUTPUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
