#!/usr/bin/env python3
"""Small fail-closed regression checks for coverage bindings.

The checks use only temporary fixtures and the verifier's public validation
helpers.  They intentionally exercise receipts that could otherwise look
claim-bearing while being forged, failed, ignored, unbound, or outside the
evidence root.
"""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import tempfile
from typing import Callable


HERE = Path(__file__).resolve().parent
VERIFY = HERE / "verify.py"


def load_verifier():
    spec = importlib.util.spec_from_file_location("lookup_verify_negative_cases", VERIFY)
    if spec is None or spec.loader is None:
        raise RuntimeError("unable to load lookup verifier")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def rejected(
    module,
    label: str,
    action: Callable[[], object],
    error_fragment: str | None = None,
) -> None:
    try:
        action()
    except module.VerificationError as error:
        if error_fragment is not None and error_fragment not in str(error):
            raise AssertionError(f"{label} failed for an unexpected reason: {error}") from error
        return
    raise AssertionError(f"{label} unexpectedly passed")


def main() -> int:
    verifier = load_verifier()
    contract_hash = verifier.digest(verifier.CONTRACT)

    rejected(
        verifier,
        "forged nonempty evidence",
        lambda: verifier.validate_evidence(
            {"notes": "pass"},
            ["requirement"],
            "negative forged evidence",
            contract_hash,
        ),
    )
    rejected(
        verifier,
        "evidence-root traversal",
        lambda: verifier.resolve_bound_file("evidence", "../outside", "negative traversal"),
    )

    source_text = "#[test]\nfn lookup_case() {}\n"
    passing_log = "test lookup_case ... ok\ntest result: ok. 1 passed; 0 failed; 0 ignored\n"
    with tempfile.TemporaryDirectory() as directory:
        root = Path(directory)
        source = root / "lookup.rs"
        receipt = root / "lookup.log"
        source.write_text(source_text, encoding="utf-8")

        receipt.write_text(passing_log, encoding="utf-8")
        verifier.validate_test_receipt(source, receipt, "gates/lookup.log", ["lookup_case"], "positive test binding")

        for state, summary in (
            ("FAILED", "test result: FAILED. 0 passed; 1 failed; 0 ignored\n"),
            ("ignored", "test result: ok. 0 passed; 0 failed; 1 ignored\n"),
        ):
            receipt.write_text(f"test lookup_case ... {state}\n{summary}", encoding="utf-8")
            rejected(
                verifier,
                f"{state} test result",
                lambda: verifier.validate_test_receipt(
                    source, receipt, "gates/lookup.log", ["lookup_case"], f"{state} test"
                ),
            )

        source.write_text("fn helper() {}\n", encoding="utf-8")
        receipt.write_text(passing_log, encoding="utf-8")
        rejected(
            verifier,
            "source without test",
            lambda: verifier.validate_test_receipt(
                source, receipt, "gates/lookup.log", ["lookup_case"], "missing test"
            ),
        )

    rejected(
        verifier,
        "wrong source hash",
        lambda: verifier.validate_binding(
            {
                "kind": "oracle",
                "root": "evidence",
                "path": "contract.md",
                "sha256": "0" * 64,
                "identifiers": ["address.a1.default"],
                "requirements": ["requirement"],
                "receipt": {
                    "root": "evidence",
                    "path": "contract.md",
                    "sha256": contract_hash,
                },
            },
            {"requirement"},
            "negative source hash",
        ),
        error_fragment="source hash",
    )

    with tempfile.TemporaryDirectory() as directory:
        receipt = Path(directory) / "native.json"
        receipt.write_text(
            json.dumps(
                {
                    "status": "PASS",
                    "results": [
                        {
                            "case": "native.case",
                            "native": {"type": "number", "value": 1},
                            "comparison": "native-divergence",
                        }
                    ],
                    "divergences": [{"case": "native.case", "reason": "host token"}],
                }
            ),
            encoding="utf-8",
        )
        rejected(
            verifier,
            "undocumented native divergence",
            lambda: verifier.validate_structured_receipt(
                "native", receipt, "native/native.json", ["native.case"], "native divergence"
            ),
        )

        receipt.write_text(
            json.dumps(
                {
                    "status": "PASS",
                    "results": [{"case": "native.case", "native": {"type": "number", "value": 1}}],
                }
            ),
            encoding="utf-8",
        )
        rejected(
            verifier,
            "native row without comparison disposition",
            lambda: verifier.validate_structured_receipt(
                "native", receipt, "native/native.json", ["native.case"], "native disposition"
            ),
        )

        receipt.write_text(
            json.dumps(
                {
                    "status": "PASS",
                    "results": [
                        {
                            "comparison": "fixture-data",
                            "native": {"type": "number", "value": 1},
                        }
                    ],
                }
            ),
            encoding="utf-8",
        )
        rejected(
            verifier,
            "fixture-data without note",
            lambda: verifier.validate_structured_receipt(
                "native", receipt, "native/native.json", [], "native fixture"
            ),
        )

    print("coverage negative cases passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
