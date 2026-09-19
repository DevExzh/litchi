#!/usr/bin/env python3
"""Re-run the frozen receipt verifier with its one naming correction in memory."""

from __future__ import annotations

from contextlib import redirect_stdout
import hashlib
import importlib.util
import io
import json
from pathlib import Path
import sys


sys.dont_write_bytecode = True

RESULTS = Path(__file__).resolve().parent
VERIFY_PATH = RESULTS.parent / "verify.py"
EXPECTED_VERIFY_SHA256 = "bd69fffcdaacbb7c8d464e5fcc72c7617bd08070771678fd7ec350756e9f33fc"
RECEIPT_PATH = RESULTS / "verification-receipt.json"
ORIGINAL_DISCRETE_NAMES = (
    "combin",
    "combina",
    "fact",
    "factdouble",
    "gcd",
    "lcm",
    "multinomial",
    "even",
    "odd",
    "delta",
    "gstep",
)
CORRECTED_DISCRETE_NAMES = ORIGINAL_DISCRETE_NAMES[:-1] + ("gestep",)


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def load_frozen_verifier() -> tuple[object, str]:
    observed = digest(VERIFY_PATH)
    if observed != EXPECTED_VERIFY_SHA256:
        raise RuntimeError(
            f"frozen verify.py changed: expected {EXPECTED_VERIFY_SHA256}, observed {observed}"
        )
    verify_directory = str(VERIFY_PATH.parent)
    if verify_directory not in sys.path:
        sys.path.insert(0, verify_directory)
    spec = importlib.util.spec_from_file_location("ods_discrete_frozen_verify", VERIFY_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load frozen verifier: {VERIFY_PATH}")
    module = importlib.util.module_from_spec(spec)
    if module.__file__ != str(VERIFY_PATH):
        raise RuntimeError(f"verifier __file__ changed: {module.__file__}")
    spec.loader.exec_module(module)
    original = tuple(module.DISCRETE_NAMES)
    if original != ORIGINAL_DISCRETE_NAMES:
        raise RuntimeError(f"unexpected frozen DISCRETE_NAMES: {original!r}")
    module.DISCRETE_NAMES = CORRECTED_DISCRETE_NAMES
    return module, observed


def main() -> int:
    arguments = sys.argv[1:]
    module, verify_sha256 = load_frozen_verifier()
    verifier_argv = [str(VERIFY_PATH), *arguments]
    sys.argv = verifier_argv
    output = io.StringIO()
    status = 1
    failure: str | None = None
    try:
        with redirect_stdout(output):
            status = int(module.main())
    except BaseException as error:  # retain a receipt before surfacing the failure
        failure = f"{type(error).__name__}: {error}"
    receipt = {
        "status": "ok" if failure is None and status == 0 else "failed",
        "verify_path": str(VERIFY_PATH),
        "verify_sha256_expected": EXPECTED_VERIFY_SHA256,
        "verify_sha256_observed": verify_sha256,
        "module_file": str(module.__file__),
        "original_discrete_names": list(ORIGINAL_DISCRETE_NAMES),
        "corrected_discrete_names": list(CORRECTED_DISCRETE_NAMES),
        "correction": "one in-memory DISCRETE_NAMES entry: gstep -> gestep",
        "argv": verifier_argv,
        "stdout": output.getvalue(),
        "return_code": status,
    }
    if failure is not None:
        receipt["error"] = failure
    RECEIPT_PATH.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    sys.stdout.write(output.getvalue())
    if failure is not None:
        raise RuntimeError(failure)
    return status


if __name__ == "__main__":
    raise SystemExit(main())
