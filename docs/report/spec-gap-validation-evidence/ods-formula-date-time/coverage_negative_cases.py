#!/usr/bin/env python3
"""Regression checks for date/time coverage false-PASS paths.

All receipt fixtures are temporary.  These checks prove that a nonempty
placeholder, an unbound identifier, stale contract custody, or a shrunk
immutable scope cannot be promoted to verified coverage.
"""

from __future__ import annotations

import copy
import importlib.util
import json
from pathlib import Path
import tempfile
from typing import Callable


HERE = Path(__file__).resolve().parent
VERIFY = HERE / "verify.py"


def load_verifier():
    spec = importlib.util.spec_from_file_location("date_time_verify_negative_cases", VERIFY)
    if spec is None or spec.loader is None:
        raise RuntimeError("unable to load date/time verifier")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def rejected(module, label: str, action: Callable[[], object], fragment: str | None = None) -> None:
    try:
        action()
    except (module.VerificationError, module.PendingReceipt) as error:
        if fragment is not None and fragment not in str(error):
            raise AssertionError(f"{label} failed for unexpected reason: {error}") from error
        return
    raise AssertionError(f"{label} unexpectedly passed")


def oracle_fixture(module, test_source: str, test_receipt: str, test_receipt_hash: str):
    return {
        "schema": "ods-formula-date-time-oracle-receipt-v1",
        "status": "PASS",
        "contract_sha256": module.digest(module.CONTRACT),
        "functions": list(module.FUNCTIONS),
        "vector_count": len(module.FUNCTIONS),
        "execution": {
            "executed": True,
            "exit_code": 0,
            "runner": "cargo-test",
            "command": ["cargo", "test", "--exact", "date_time_oracle"],
            "results_sha256": "1" * 64,
            "test_source": test_source,
            "test_receipt": test_receipt,
            "test": "oracle_vectors_match",
            "test_receipt_sha256": test_receipt_hash,
        },
        "observations": [
            {
                "id": f"fixture.{name.lower()}",
                "function": name,
                "expected": {"kind": "number", "value": 0},
            }
            for name in module.FUNCTIONS
        ],
    }


def main() -> int:
    verifier = load_verifier()
    manifest = verifier.read_json(verifier.MANIFEST)
    contract_hash = verifier.digest(verifier.CONTRACT)

    with tempfile.TemporaryDirectory() as directory:
        placeholder = Path(directory) / "placeholder.json"
        placeholder.write_text(
            json.dumps(
                {
                    "schema": "ods-formula-date-time-oracle-v1",
                    "status": "PASS",
                    "contract_sha256": contract_hash,
                    "functions": list(verifier.FUNCTIONS),
                    "vector_count": 0,
                    "vectors": [],
                }
            ),
            encoding="utf-8",
        )
        rejected(
            verifier,
            "forged nonempty oracle placeholder",
            lambda: verifier.validate_structured_receipt(
                "oracle",
                placeholder,
                "oracle/placeholder.json",
                ["fixture.date"],
                "placeholder oracle",
                contract_hash,
            ),
            "expected-vector corpus",
        )

    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        repo = root / "repo"
        evidence = root / "evidence"
        source = repo / "tests" / "oracle.rs"
        test_receipt = evidence / "gates" / "oracle.log"
        receipt = evidence / "oracle.json"
        source.parent.mkdir(parents=True)
        test_receipt.parent.mkdir(parents=True)
        source.write_text("#[test]\nfn oracle_vectors_match() {}\n", encoding="utf-8")
        test_receipt.write_text(
            "test oracle_vectors_match ... ok\n"
            "test result: ok. 1 passed; 0 failed; 0 ignored\n",
            encoding="utf-8",
        )
        receipt.write_text(
            json.dumps(oracle_fixture(verifier, "tests/oracle.rs", "gates/oracle.log", verifier.digest(test_receipt))),
            encoding="utf-8",
        )
        old_repo, old_here = verifier.REPO, verifier.HERE
        verifier.REPO, verifier.HERE = repo, evidence
        try:
            verifier.validate_structured_receipt(
                "oracle",
                receipt,
                "oracle.json",
                ["fixture.date"],
                "positive oracle identifier",
                contract_hash,
            )
            rejected(
                verifier,
                "nonexistent oracle identifier",
                lambda: verifier.validate_structured_receipt(
                    "oracle",
                    receipt,
                    "oracle.json",
                    ["fixture.missing"],
                    "missing oracle identifier",
                    contract_hash,
                ),
                "identifiers are absent",
            )
        finally:
            verifier.REPO, verifier.HERE = old_repo, old_here

    stale = {
        "schema": verifier.SCHEMA,
        "status": "PASS",
        "contract_sha256": "0" * 64,
        "bindings": [],
    }
    rejected(
        verifier,
        "stale contract custody",
        lambda: verifier.validate_evidence(stale, ["one requirement"], "stale contract", contract_hash, allow_missing=True),
        "contract hash",
    )

    shrunk = copy.deepcopy(verifier.read_json(verifier.SCOPE))
    shrunk["functions"].pop("YEARFRAC")
    rejected(
        verifier,
        "shrunk immutable scope",
        lambda: verifier.validate_scope_value(shrunk, manifest, "shrunk scope"),
        "expected",
    )

    with tempfile.TemporaryDirectory() as directory:
        forged_manifest = Path(directory) / "coverage-requirements.json"
        forged = copy.deepcopy(manifest)
        forged["status"] = "PASS"
        for entry in forged["functions"].values():
            entry["evidence"]["status"] = "PASS"
        for entry in forged["cross_cutting_evidence"]:
            entry["evidence"]["status"] = "PASS"
        forged_manifest.write_text(json.dumps(forged), encoding="utf-8")
        old_manifest = verifier.MANIFEST
        verifier.MANIFEST = forged_manifest
        try:
            rejected(
                verifier,
                "PASS manifest with no semantic bindings",
                verifier.verify_coverage,
                "PASS evidence has no bindings",
            )
        finally:
            verifier.MANIFEST = old_manifest

    print("date/time coverage negative cases passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
