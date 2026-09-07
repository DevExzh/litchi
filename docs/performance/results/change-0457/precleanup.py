#!/usr/bin/env python3
"""Run the exact 0457 bundle verifier before deleting owned build artifacts.

This driver records the verifier identity and process result.  It deliberately
does not create or update the bundle seal: the final verifier owns regeneration
of that seal after all derived proof receipts exist.
"""

from __future__ import annotations

import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
from typing import Any


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
RECEIPT = ROOT / "precleanup.json"
VERIFIER_CANDIDATES = ("bundle-verifier.py", "bundle_verify.py", "verify.py")


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def fail(message: str) -> "NoReturn":
    raise SystemExit(message)


def regular(path: Path, label: str) -> Path:
    if not path.is_file() or path.is_symlink():
        fail(f"{label} is missing, not regular, or a symlink: {path}")
    return path


def verifier_path(requested: str | None) -> Path:
    if requested is not None:
        candidate = (ROOT / requested).resolve()
        if not candidate.is_relative_to(ROOT.resolve()):
            fail("--verifier must stay inside the change-0457 bundle")
        return regular(candidate, "bundle verifier")
    candidates = [regular(ROOT / name, "bundle verifier") for name in VERIFIER_CANDIDATES if (ROOT / name).exists()]
    if len(candidates) != 1:
        names = ", ".join(str(path.relative_to(ROOT)) for path in candidates) or "none"
        fail(f"expected exactly one root bundle verifier ({', '.join(VERIFIER_CANDIDATES)}), found {names}")
    return candidates[0]


def json_status(raw: str) -> Any:
    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        return None


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--verifier", help="bundle-relative root verifier path")
    args = parser.parse_args(sys.argv[1:] if argv is None else argv)
    if RECEIPT.exists():
        fail(f"refusing to overwrite retained receipt: {RECEIPT}")

    verifier = verifier_path(args.verifier)
    verifier_relative = str(verifier.relative_to(ROOT))
    command = [sys.executable, "-B", str(verifier), "--precleanup"]
    record: dict[str, Any] = {
        "schema": "litchi-0457-precleanup-v1",
        "change": 457,
        "argv": command,
        "cwd": str(REPO),
        "verifier_path": verifier_relative,
        "verifier_sha256": sha(verifier),
        "driver_sha256": sha(Path(__file__).resolve()),
        "started_utc": now(),
    }
    try:
        result = subprocess.run(command, cwd=REPO, capture_output=True, text=True, errors="replace")
        record.update(
            exit_code=result.returncode,
            stdout=result.stdout,
            stderr=result.stderr,
            finished_utc=now(),
        )
    except OSError as error:
        record.update(exit_code=None, stdout="", stderr=str(error), finished_utc=now())
    record["verifier_output"] = json_status(record["stdout"])
    record["status"] = "pass" if record["exit_code"] == 0 else "failed"
    with RECEIPT.open("x", encoding="utf-8") as output:
        output.write(json.dumps(record, ensure_ascii=False, indent=2) + "\n")
    print(json.dumps({key: record[key] for key in ("status", "exit_code", "verifier_path", "verifier_sha256")}, sort_keys=True))
    return 0 if record["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
