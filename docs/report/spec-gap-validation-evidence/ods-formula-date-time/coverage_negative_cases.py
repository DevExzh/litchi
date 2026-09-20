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


def source_review_fixture(module, root: Path):
    """Build a complete temporary source-review binding without touching evidence."""

    repo = root / "repo"
    evidence = root / "evidence"
    gates = evidence / "gates"
    source_relative = "crates/litchi-ods/src/codec/formula/evaluation/value.rs"
    source = repo / source_relative
    source.parent.mkdir(parents=True)
    gates.mkdir(parents=True)
    source.write_text("// frozen source proof fixture\n", encoding="utf-8")
    report = evidence / "resource-review.md"
    semantic = evidence / "semantic-review.md"
    report.write_text("resource reviewer report\n", encoding="utf-8")
    semantic.write_text("semantic reviewer report\n", encoding="utf-8")
    freeze = gates / "freeze.json"
    source_hash = module.digest(source)
    freeze.write_text(
        json.dumps(
            {
                "schema": "ods-formula-date-time-freeze-v1",
                "selected_files": {source_relative: source_hash},
            },
            sort_keys=True,
        ),
        encoding="utf-8",
    )
    contract_hash = module.digest(module.CONTRACT)
    final_review = evidence / "review-receipt.json"
    final_review.write_text(
        json.dumps(
            {
                "status": "PASS",
                "contract_sha256": contract_hash,
                "freeze_sha256": module.digest(freeze),
                "reviews": {
                    "semantic": {
                        "path": semantic.name,
                        "sha256": module.digest(semantic),
                        "status": "PASS",
                    },
                    "resource": {
                        "path": report.name,
                        "sha256": module.digest(report),
                        "status": "PASS",
                    },
                },
            },
            sort_keys=True,
        ),
        encoding="utf-8",
    )
    receipt = gates / "source-review-receipt.json"
    requirement = "resolver access is read-only"
    receipt_value = {
        "schema": module.SOURCE_REVIEW_SCHEMA,
        "contract_sha256": contract_hash,
        "freeze_sha256": module.digest(freeze),
        "reviewer": "date_time_resource_review",
        "report": {
            "root": "evidence",
            "path": report.name,
            "sha256": module.digest(report),
        },
        "review_receipt": {
            "root": "evidence",
            "path": final_review.name,
            "sha256": module.digest(final_review),
        },
        "proofs": [
            {
                "id": "proof.resolver_read_only",
                "source": {"root": "repo", "path": source_relative, "sha256": source_hash},
                "requirements": [requirement],
            }
        ],
    }
    receipt.write_text(json.dumps(receipt_value, sort_keys=True), encoding="utf-8")
    binding = {
        "kind": "source_review",
        "root": "repo",
        "path": source_relative,
        "sha256": source_hash,
        "identifiers": ["proof.resolver_read_only"],
        "requirements": [requirement],
        "receipt": {
            "root": "evidence",
            "path": "gates/source-review-receipt.json",
            "sha256": module.digest(receipt),
        },
    }
    return {
        "repo": repo,
        "evidence": evidence,
        "gates": gates,
        "source": source,
        "source_relative": source_relative,
        "source_hash": source_hash,
        "freeze": freeze,
        "report": report,
        "final_review": final_review,
        "receipt": receipt,
        "receipt_value": receipt_value,
        "binding": binding,
        "requirements": {requirement},
    }


def write_source_review_receipt(module, fixture) -> None:
    fixture["receipt"].write_text(
        json.dumps(fixture["receipt_value"], sort_keys=True),
        encoding="utf-8",
    )
    fixture["binding"]["receipt"]["sha256"] = module.digest(fixture["receipt"])


def remove_source_from_freeze(module, fixture) -> None:
    fixture["freeze"].write_text(
        json.dumps(
            {
                "schema": "ods-formula-date-time-freeze-v1",
                "selected_files": {"other/frozen.rs": "0" * 64},
            },
            sort_keys=True,
        ),
        encoding="utf-8",
    )
    final_review = json.loads(fixture["final_review"].read_text(encoding="utf-8"))
    final_review["freeze_sha256"] = module.digest(fixture["freeze"])
    fixture["final_review"].write_text(json.dumps(final_review, sort_keys=True), encoding="utf-8")
    fixture["receipt_value"]["freeze_sha256"] = module.digest(fixture["freeze"])
    fixture["receipt_value"]["review_receipt"]["sha256"] = module.digest(fixture["final_review"])
    write_source_review_receipt(module, fixture)


def source_review_case(module, mutate=None):
    with tempfile.TemporaryDirectory() as directory:
        fixture = source_review_fixture(module, Path(directory))
        old = {
            "REPO": module.REPO,
            "HERE": module.HERE,
            "GATES": module.GATES,
            "REVIEW_RECEIPT": module.REVIEW_RECEIPT,
        }
        module.REPO = fixture["repo"]
        module.HERE = fixture["evidence"]
        module.GATES = fixture["gates"]
        module.REVIEW_RECEIPT = fixture["final_review"]
        try:
            if mutate is not None:
                mutate(fixture)
            return module.validate_binding(
                fixture["binding"],
                fixture["requirements"],
                "source-review fixture",
                allow_missing=False,
                contract_hash=module.digest(module.CONTRACT),
            )
        finally:
            module.REPO = old["REPO"]
            module.HERE = old["HERE"]
            module.GATES = old["GATES"]
            module.REVIEW_RECEIPT = old["REVIEW_RECEIPT"]


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

    if source_review_case(verifier) != []:
        raise AssertionError("valid source-review binding did not return an empty pending list")
    rejected(
        verifier,
        "source review without reviewer identity",
        lambda: source_review_case(
            verifier,
            lambda fixture: (
                fixture["receipt_value"].pop("reviewer"),
                write_source_review_receipt(verifier, fixture),
            ),
        ),
        "unexpected fields",
    )
    rejected(
        verifier,
        "source review with invented status",
        lambda: source_review_case(
            verifier,
            lambda fixture: (
                fixture["receipt_value"].update({"status": "PASS"}),
                write_source_review_receipt(verifier, fixture),
            ),
        ),
        "unexpected fields",
    )
    rejected(
        verifier,
        "source review with unmapped proof requirement",
        lambda: source_review_case(
            verifier,
            lambda fixture: (
                fixture["receipt_value"]["proofs"][0]["requirements"].append("unmapped requirement"),
                write_source_review_receipt(verifier, fixture),
            ),
        ),
        "exact mapped subset",
    )
    rejected(
        verifier,
        "source review outside source freeze",
        lambda: source_review_case(
            verifier,
            lambda fixture: remove_source_from_freeze(verifier, fixture),
        ),
        "outside the source freeze",
    )
    rejected(
        verifier,
        "source review with stale reviewed report hash",
        lambda: source_review_case(
            verifier,
            lambda fixture: (
                fixture["receipt_value"]["report"].update({"sha256": "0" * 64}),
                write_source_review_receipt(verifier, fixture),
            ),
        ),
        "source-review report hash",
    )
    rejected(
        verifier,
        "source review without final PASS review receipt",
        lambda: source_review_case(
            verifier,
            lambda fixture: (
                fixture["final_review"].write_text(
                    json.dumps(
                        {
                            "status": "PENDING",
                            "contract_sha256": verifier.digest(verifier.CONTRACT),
                            "freeze_sha256": verifier.digest(fixture["freeze"]),
                            "reviews": {},
                        },
                        sort_keys=True,
                    ),
                    encoding="utf-8",
                ),
                fixture["receipt_value"]["review_receipt"].update(
                    {"sha256": verifier.digest(fixture["final_review"])}
                ),
                write_source_review_receipt(verifier, fixture),
            ),
        ),
        "review status",
    )

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
            entry["evidence"]["bindings"] = []
        for entry in forged["cross_cutting_evidence"]:
            entry["evidence"]["status"] = "PASS"
            entry["evidence"]["bindings"] = []
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
