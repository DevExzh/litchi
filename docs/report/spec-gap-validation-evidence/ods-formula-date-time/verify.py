#!/usr/bin/env python3
"""Fail-closed custody verifier for the ODS date/time evidence bundle.

This verifier is intentionally independent of the lookup batch.  During the
preparation phase it verifies only the date/time contract identity, immutable
coverage projection, and the closed evidence schema.  Empty/pending evidence
is reportable with ``--allow-pending``; it can never produce ``verified: true``.
When evidence is promoted to PASS, every requirement must be covered by hashed
source and receipt bindings.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import subprocess
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
CONTRACT = HERE / "contract.md"
SCOPE = HERE / "coverage-scope.json"
MANIFEST = HERE / "coverage-requirements.json"
BASELINE = HERE / "baseline.json"
GATES = HERE / "gates"
NATIVE = HERE / "native"
PERFORMANCE = HERE / "performance"
ORACLE_CORPUS = HERE / "oracle-vectors.json"
ORACLE_EXECUTION_CANDIDATES = (
    HERE / "oracle-execution.json",
    HERE / "oracle-results.json",
    HERE / "oracle-receipt.json",
)
REVIEW_RECEIPT = HERE / "review-receipt.json"

SCHEMA = "ods-formula-date-time-coverage-v1"
SCOPE_SCHEMA = "ods-formula-date-time-coverage-scope-v1"
FUNCTIONS = (
    "DATE",
    "DATEDIF",
    "DATEVALUE",
    "DAY",
    "DAYS",
    "DAYS360",
    "EASTERSUNDAY",
    "EDATE",
    "EOMONTH",
    "HOUR",
    "ISOWEEKNUM",
    "MINUTE",
    "MONTH",
    "NETWORKDAYS",
    "NOW",
    "SECOND",
    "TIME",
    "TIMEVALUE",
    "TODAY",
    "WEEKDAY",
    "WEEKNUM",
    "WORKDAY",
    "YEAR",
    "YEARFRAC",
)
FUNCTION_SET = set(FUNCTIONS)
NORMATIVE = "OpenDocument 1.4 Part 4 section 6.10"
SCOPE_KEYS = {"schema", "normative", "functions", "cross_cutting"}
MANIFEST_KEYS = {
    "schema",
    "status",
    "normative",
    "functions",
    "cross_cutting",
    "cross_cutting_evidence",
}
EVIDENCE_KEYS = {"schema", "status", "contract_sha256", "bindings"}
CROSS_EVIDENCE_KEYS = {"requirement", "evidence"}
BINDING_KEYS = {"kind", "root", "path", "sha256", "identifiers", "requirements", "receipt"}
RECEIPT_KEYS = {"root", "path", "sha256"}
BINDING_KINDS = {"focused_test", "resource", "oracle", "native", "performance", "gate", "source_review"}
SOURCE_REVIEW_SCHEMA = "ods-formula-date-time-source-review-v1"
SOURCE_REVIEW_RECEIPT_KEYS = {
    "schema",
    "contract_sha256",
    "freeze_sha256",
    "reviewer",
    "report",
    "review_receipt",
    "proofs",
}
SOURCE_REVIEW_REPORT_KEYS = {"root", "path", "sha256"}
SOURCE_REVIEW_PROOF_KEYS = {"id", "source", "requirements"}
SOURCE_REVIEW_SOURCE_KEYS = {"root", "path", "sha256"}
PENDING_WORDS = {"pending", "planning", "incomplete", "hold"}
PASS_WORDS = {"pass", "passed", "ok", "verified", "complete", "captured"}
HEX_SHA256 = re.compile(r"[0-9a-f]{64}\Z")
ANSI_ESCAPE = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]")
TEST_RESULT = re.compile(r"^\s*test\s+(\S+)\s+\.\.\.\s+(ok|FAILED|ignored)\s*$", re.MULTILINE)
TEST_SUMMARY = re.compile(
    r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored"
)
RUST_TEST_ATTRIBUTE = re.compile(r"#\[\s*(?:(?:[A-Za-z_][A-Za-z0-9_]*::)*)test(?:\W|$)")
RUST_FUNCTION = re.compile(r"\b(?:pub(?:\([^)]*\))?\s+)?(?:async\s+)?fn\s+([A-Za-z_][A-Za-z0-9_]*)\s*\(")
STRUCTURED_IDENTIFIER_KEYS = {
    "case",
    "id",
    "name",
    "function",
    "formula",
    "label",
    "test",
    "tests",
    "check",
    "checks",
    "command",
    "commands",
}
ORACLE_NORMATIVE_ARCHIVE = "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4"
ORACLE_NORMATIVE_PART4 = "ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1"


class VerificationError(RuntimeError):
    """A present file or record is malformed, stale, or inconsistent."""


class PendingReceipt(RuntimeError):
    """A preparatory evidence file has not been produced yet."""


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def read_json(path: Path) -> Any:
    if not path.is_file():
        raise PendingReceipt(f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise VerificationError(f"invalid JSON evidence {path}: {error}") from error


def equal(label: str, observed: Any, expected: Any) -> None:
    if observed != expected:
        raise VerificationError(f"{label}: expected {expected!r}, observed {observed!r}")


def validate_sha(value: Any, label: str) -> str:
    if not isinstance(value, str) or HEX_SHA256.fullmatch(value) is None:
        raise VerificationError(f"{label} is not a lowercase SHA-256 digest")
    return value


def relative_path(root: Path, value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value or Path(value).is_absolute():
        raise VerificationError(f"{label} is not a relative path: {value!r}")
    path = (root / value).resolve()
    try:
        path.relative_to(root.resolve())
    except ValueError as error:
        raise VerificationError(f"{label} escapes its root: {value!r}") from error
    return path


def resolve_bound(root_name: Any, value: Any, label: str, *, allow_missing: bool) -> Path:
    if root_name == "repo":
        root = REPO
    elif root_name == "evidence":
        root = HERE
    else:
        raise VerificationError(f"{label} has invalid root {root_name!r}")
    path = relative_path(root, value, label)
    if not path.is_file():
        if allow_missing:
            raise PendingReceipt(f"{label} is absent: {value}")
        raise VerificationError(f"{label} is absent: {value}")
    return path


def strings(value: Any, label: str, *, nonempty: bool = True) -> list[str]:
    if not isinstance(value, list):
        raise VerificationError(f"{label} must be a list of strings")
    if nonempty and not value:
        raise VerificationError(f"{label} must not be empty")
    if not all(isinstance(item, str) and item.strip() and "\n" not in item and "\r" not in item for item in value):
        raise VerificationError(f"{label} contains a malformed string")
    if len(value) != len(set(value)):
        raise VerificationError(f"{label} contains duplicates")
    return value


def status_class(value: Any, label: str) -> str:
    if not isinstance(value, str):
        raise VerificationError(f"{label} status is malformed")
    normalized = value.strip().lower().replace("-", " ")
    if normalized == "pass" or normalized in PASS_WORDS:
        return "pass"
    if any(word in normalized for word in PENDING_WORDS):
        return "pending"
    raise VerificationError(f"{label} status is not PASS or pending: {value!r}")


def scope_projection(manifest: dict[str, Any]) -> dict[str, Any]:
    functions = manifest.get("functions")
    if not isinstance(functions, dict):
        raise VerificationError("coverage manifest functions are absent")
    projected: dict[str, list[str]] = {}
    for name in FUNCTIONS:
        entry = functions.get(name)
        if not isinstance(entry, dict) or not isinstance(entry.get("requirements"), list):
            raise VerificationError(f"coverage function {name} has no requirements")
        projected[name] = entry["requirements"]
    cross = manifest.get("cross_cutting")
    if not isinstance(cross, list):
        raise VerificationError("coverage cross-cutting requirements are absent")
    return {
        "schema": SCOPE_SCHEMA,
        "normative": manifest.get("normative"),
        "functions": projected,
        "cross_cutting": cross,
    }


def validate_scope_value(observed: Any, manifest: dict[str, Any], label: str = "coverage scope") -> dict[str, Any]:
    if not isinstance(observed, dict) or set(observed) != SCOPE_KEYS:
        raise VerificationError(f"{label} has unexpected fields")
    expected = scope_projection(manifest)
    equal(label, observed, expected)
    equal(f"{label} schema", observed.get("schema"), SCOPE_SCHEMA)
    equal(f"{label} normative", observed.get("normative"), NORMATIVE)
    functions = observed.get("functions")
    if not isinstance(functions, dict) or tuple(functions) != FUNCTIONS:
        raise VerificationError(f"{label} does not list exactly the 24 functions in order")
    cross = observed.get("cross_cutting")
    strings(cross, f"{label} cross_cutting")
    if len(cross) != len(set(cross)):
        raise VerificationError(f"{label} cross_cutting requirements contain duplicates")
    return {
        "sha256": digest(SCOPE),
        "functions": len(functions),
        "cross_cutting": len(cross),
    }


def validate_scope(manifest: dict[str, Any]) -> dict[str, Any]:
    return validate_scope_value(read_json(SCOPE), manifest)


def rust_test_names(source: str) -> set[str]:
    """Return Rust functions explicitly marked with a test attribute."""

    lines = source.splitlines()
    names: set[str] = set()
    for index, line in enumerate(lines):
        attribute = RUST_TEST_ATTRIBUTE.search(line)
        if attribute is None:
            continue
        window = line[attribute.end() :] + "\n" + "\n".join(lines[index + 1 : index + 9])
        match = RUST_FUNCTION.search(window)
        if match:
            names.add(match.group(1))
    return names


def short_test_name(identifier: str) -> str:
    return identifier.rsplit("::", 1)[-1]


def validate_test_receipt(
    source: Path,
    receipt_path: Path,
    receipt_relative: str,
    identifiers: list[str],
    label: str,
) -> None:
    """Bind focused/resource evidence to defined Rust tests and passing lines."""

    if source.suffix != ".rs":
        raise VerificationError(f"{label} test source is not Rust")
    try:
        defined = rust_test_names(source.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError) as error:
        raise VerificationError(f"{label} test source cannot be read as UTF-8 Rust: {error}") from error
    if not defined:
        raise VerificationError(f"{label} test source defines no #[test] function")
    if not receipt_relative.startswith("gates/") or receipt_path.suffix not in {".log", ".txt"}:
        raise VerificationError(f"{label} test receipt must be a retained gates text log")
    text = ANSI_ESCAPE.sub("", receipt_path.read_bytes().decode("utf-8", errors="replace"))
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    results = TEST_RESULT.findall(text)
    if not results:
        raise VerificationError(f"{label} test receipt has no exact Cargo test result lines")
    summaries = TEST_SUMMARY.findall(text)
    if not summaries or any(
        status != "ok" or int(failed) != 0 or int(ignored) != 0
        for status, _, failed, ignored in summaries
    ):
        raise VerificationError(f"{label} test receipt contains a failed or ignored summary")
    for identifier in identifiers:
        short = short_test_name(identifier)
        if short not in defined:
            raise VerificationError(f"{label} source does not define test {identifier!r}")
        matches = [
            (name, status)
            for name, status in results
            if name == identifier or name == short or name.endswith(f"::{short}")
        ]
        if not matches:
            raise VerificationError(f"{label} receipt has no exact result for test {identifier!r}")
        if any(status != "ok" for _, status in matches):
            raise VerificationError(f"{label} test {identifier!r} is not passing")


def structured_values(value: Any) -> set[str]:
    """Collect identifiers only from named structured receipt fields."""

    found: set[str] = set()
    if isinstance(value, dict):
        for key, child in value.items():
            if key in STRUCTURED_IDENTIFIER_KEYS:
                if isinstance(child, str) and child.strip():
                    found.add(child)
                elif isinstance(child, list):
                    found.update(item for item in child if isinstance(item, str) and item.strip())
            if key in {"reference_reads", "results_by_case", "cases_by_name"} and isinstance(child, dict):
                found.update(name for name in child if isinstance(name, str) and name.strip())
            found.update(structured_values(child))
    elif isinstance(value, list):
        for child in value:
            found.update(structured_values(child))
    return found


def receipt_status(value: dict[str, Any], label: str) -> None:
    if value.get("verified") is False:
        raise VerificationError(f"{label} receipt is not verified")
    status = value.get("status")
    if status is None:
        return
    if not isinstance(status, str):
        raise VerificationError(f"{label} receipt status is malformed")
    normalized = status.lower().replace("-", " ").strip()
    if normalized in PENDING_WORDS or any(word in normalized for word in ("fail", "error", "unsupported")):
        raise VerificationError(f"{label} receipt has failure/pending status {status!r}")


def load_structured_receipt(path: Path, label: str) -> Any:
    if path.suffix == ".jsonl":
        rows = []
        for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if not line.strip():
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError as error:
                raise VerificationError(f"{label} JSONL line {line_number} is malformed") from error
        if not rows:
            raise VerificationError(f"{label} structured receipt is empty")
        return rows
    if path.suffix != ".json":
        raise VerificationError(f"{label} receipt is not JSON/JSONL")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise VerificationError(f"{label} structured receipt is malformed: {error}") from error


def _receipt_rows(value: Any, label: str, keys: tuple[str, ...]) -> list[dict[str, Any]]:
    if isinstance(value, list):
        rows = value
    elif isinstance(value, dict):
        rows = next((value.get(key) for key in keys if key in value), None)
    else:
        rows = None
    if not isinstance(rows, list) or not rows or not all(isinstance(row, dict) for row in rows):
        raise VerificationError(f"{label} has no nonempty structured rows")
    return rows


def _typed_outcome(row: dict[str, Any]) -> bool:
    return any(key in row for key in ("expected", "result", "value", "error", "kind", "type", "expected_type", "native"))


def _nonempty_string(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value.strip() or "\n" in value or "\r" in value:
        raise VerificationError(f"{label} must be a nonempty single-line string")
    return value


def validate_source_review_receipt(
    value: Any,
    receipt_path: Path,
    receipt_relative: str,
    identifiers: list[str],
    mapped_requirements: list[str] | None,
    source_binding: dict[str, Any] | None,
    label: str,
    contract_hash: str | None,
) -> None:
    """Validate a source proof bound to the frozen source map.

    Source review is deliberately a separate receipt kind.  It has no status
    field: the coverage binding becomes usable only when this complete,
    hash-bound receipt exists.  The review report remains mutable evidence and
    is therefore hashed here rather than included in the source freeze.
    """

    if not isinstance(value, dict) or set(value) != SOURCE_REVIEW_RECEIPT_KEYS:
        raise VerificationError(f"{label} source-review receipt has unexpected fields")
    equal(f"{label} source-review schema", value.get("schema"), SOURCE_REVIEW_SCHEMA)
    if contract_hash is None:
        raise VerificationError(f"{label} source-review receipt has no contract identity")
    equal(f"{label} source-review contract hash", value.get("contract_sha256"), contract_hash)
    reviewer = _nonempty_string(value.get("reviewer"), f"{label} source-review reviewer")
    # Keep the local binding alive so a future refactor cannot accidentally
    # make reviewer identity dead metadata.
    if not reviewer.strip():
        raise VerificationError(f"{label} source-review reviewer is empty")
    if mapped_requirements is None or source_binding is None:
        raise VerificationError(f"{label} source-review binding context is absent")
    if len(mapped_requirements) != len(set(mapped_requirements)):
        raise VerificationError(f"{label} source-review binding requirements are duplicated")
    if len(identifiers) != len(set(identifiers)):
        raise VerificationError(f"{label} source-review binding identifiers are duplicated")
    expected_requirements = set(mapped_requirements)
    if not expected_requirements:
        raise VerificationError(f"{label} source-review binding has no mapped requirements")

    freeze_path = GATES / "freeze.json"
    if not freeze_path.is_file():
        raise PendingReceipt(f"{label} source-review freeze is absent")
    freeze = read_json(freeze_path)
    if not isinstance(freeze, dict):
        raise VerificationError(f"{label} source-review freeze is not an object")
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected:
        raise VerificationError(f"{label} source-review freeze has no selected_files map")
    equal(f"{label} source-review freeze hash", value.get("freeze_sha256"), digest(freeze_path))

    source_root = source_binding.get("root")
    source_relative = source_binding.get("path")
    source_hash = source_binding.get("sha256")
    if source_root != "repo" or not isinstance(source_relative, str):
        raise VerificationError(f"{label} source-review source must be repository-rooted")
    frozen_hash = selected.get(source_relative)
    if frozen_hash is None:
        raise VerificationError(f"{label} source-review source is outside the source freeze: {source_relative}")
    validate_sha(frozen_hash, f"{label} frozen source hash")
    equal(f"{label} source-review frozen source hash", frozen_hash, source_hash)
    source = resolve_bound("repo", source_relative, f"{label} source-review source", allow_missing=False)
    equal(f"{label} source-review source hash", digest(source), source_hash)

    report = value.get("report")
    if not isinstance(report, dict) or set(report) != SOURCE_REVIEW_REPORT_KEYS:
        raise VerificationError(f"{label} source-review report identity is malformed")
    if report.get("root") != "evidence":
        raise VerificationError(f"{label} source-review report must be evidence-rooted")
    report_relative = report.get("path")
    relative_path(HERE, report_relative, f"{label} source-review report")
    if report_relative == receipt_relative:
        raise VerificationError(f"{label} source-review report cannot be its own receipt")
    report_hash = validate_sha(report.get("sha256"), f"{label} source-review report hash")
    report_path = resolve_bound("evidence", report_relative, f"{label} source-review report", allow_missing=False)
    equal(f"{label} source-review report hash", digest(report_path), report_hash)

    review_receipt = value.get("review_receipt")
    if not isinstance(review_receipt, dict) or set(review_receipt) != SOURCE_REVIEW_REPORT_KEYS:
        raise VerificationError(f"{label} source-review final review receipt identity is malformed")
    if review_receipt.get("root") != "evidence":
        raise VerificationError(f"{label} source-review final review receipt must be evidence-rooted")
    review_receipt_relative = review_receipt.get("path")
    relative_path(HERE, review_receipt_relative, f"{label} source-review final review receipt")
    try:
        expected_review_receipt = str(REVIEW_RECEIPT.resolve().relative_to(HERE.resolve()))
    except ValueError as error:
        raise VerificationError(f"{label} configured final review receipt escapes evidence root") from error
    equal(
        f"{label} source-review final review receipt path",
        review_receipt_relative,
        expected_review_receipt,
    )
    review_receipt_hash = validate_sha(
        review_receipt.get("sha256"),
        f"{label} source-review final review receipt hash",
    )
    review_receipt_path = resolve_bound(
        "evidence",
        review_receipt_relative,
        f"{label} source-review final review receipt",
        allow_missing=False,
    )
    equal(
        f"{label} source-review final review receipt hash",
        digest(review_receipt_path),
        review_receipt_hash,
    )
    # Reuse the independent final-review validator.  In particular this keeps
    # a source proof from manufacturing a PASS status of its own.
    verify_reviews(contract_hash)

    proofs = value.get("proofs")
    if not isinstance(proofs, list) or not proofs:
        raise VerificationError(f"{label} source-review proofs are absent")
    proof_ids: set[str] = set()
    mapped: set[str] = set()
    for index, proof in enumerate(proofs):
        proof_label = f"{label} source-review proof {index}"
        if not isinstance(proof, dict) or set(proof) != SOURCE_REVIEW_PROOF_KEYS:
            raise VerificationError(f"{proof_label} has unexpected fields")
        proof_id = _nonempty_string(proof.get("id"), f"{proof_label} id")
        if proof_id in proof_ids:
            raise VerificationError(f"{label} source-review proof identifiers are duplicated")
        proof_ids.add(proof_id)
        proof_source = proof.get("source")
        if not isinstance(proof_source, dict) or set(proof_source) != SOURCE_REVIEW_SOURCE_KEYS:
            raise VerificationError(f"{proof_label} source identity is malformed")
        if proof_source.get("root") != "repo":
            raise VerificationError(f"{proof_label} source must be repository-rooted")
        equal(f"{proof_label} source path", proof_source.get("path"), source_relative)
        equal(f"{proof_label} source hash", proof_source.get("sha256"), source_hash)
        proof_requirement_values = strings(proof.get("requirements"), f"{proof_label} requirements")
        if len(proof_requirement_values) != len(set(proof_requirement_values)):
            raise VerificationError(f"{proof_label} requirements are duplicated")
        proof_requirements = set(proof_requirement_values)
        if not proof_requirements <= expected_requirements:
            raise VerificationError(
                f"{proof_label} requirements are not an exact mapped subset: "
                f"{sorted(proof_requirements - expected_requirements)}"
            )
        mapped.update(proof_requirements)
    if proof_ids != set(identifiers):
        raise VerificationError(
            f"{label} source-review proof identifiers do not match binding identifiers: "
            f"expected={sorted(identifiers)}, observed={sorted(proof_ids)}"
        )
    if mapped != expected_requirements:
        raise VerificationError(
            f"{label} source-review proof requirements do not exactly cover binding requirements: "
            f"missing={sorted(expected_requirements - mapped)}, extra={sorted(mapped - expected_requirements)}"
        )


def validate_expected_oracle_corpus(path: Path, contract_hash: str) -> dict[str, Any]:
    """Validate expected vectors while keeping them separate from execution."""

    value = load_structured_receipt(path, "date/time oracle vector corpus")
    if not isinstance(value, dict) or value.get("schema") != "ods-formula-date-time-oracle-v1":
        raise VerificationError("date/time oracle vector corpus has no expected-vector schema")
    equal("date/time oracle corpus contract hash", value.get("contract_sha256"), contract_hash)
    normative = value.get("normative_source")
    if not isinstance(normative, dict):
        raise VerificationError("date/time oracle corpus has no normative-source identity")
    equal("date/time oracle archive hash", normative.get("archive_sha256"), ORACLE_NORMATIVE_ARCHIVE)
    equal("date/time oracle Part 4 hash", normative.get("part4_member_sha256"), ORACLE_NORMATIVE_PART4)
    functions = value.get("functions")
    if not isinstance(functions, list) or set(functions) != FUNCTION_SET or len(functions) != len(FUNCTIONS):
        raise VerificationError("date/time oracle corpus does not cover exactly the 24 functions")
    rows = _receipt_rows(value, "date/time oracle vector corpus", ("vectors", "observations", "results", "rows"))
    vector_count = value.get("vector_count")
    if not isinstance(vector_count, int) or vector_count <= 0 or vector_count != len(rows):
        raise VerificationError("date/time oracle corpus vector_count is not bound to its rows")
    identifiers: set[str] = set()
    for index, row in enumerate(rows):
        row_id = row.get("id")
        if not isinstance(row_id, str) or not row_id or row_id in identifiers:
            raise VerificationError(f"date/time oracle corpus row {index} has a duplicate/missing id")
        identifiers.add(row_id)
        if row.get("function") not in FUNCTION_SET or not _typed_outcome(row):
            raise VerificationError(f"date/time oracle corpus row {index} lacks function or typed outcome")
    return {"sha256": digest(path), "vectors": len(rows), "functions": len(FUNCTIONS), "expected_only": True}


def validate_structured_receipt(
    kind: str,
    receipt_path: Path,
    receipt_relative: str,
    identifiers: list[str],
    label: str,
    contract_hash: str | None = None,
    *,
    mapped_requirements: list[str] | None = None,
    source_binding: dict[str, Any] | None = None,
) -> None:
    """Validate structured evidence receipt semantics."""

    value = load_structured_receipt(receipt_path, label)
    if kind == "source_review":
        validate_source_review_receipt(
            value,
            receipt_path,
            receipt_relative,
            identifiers,
            mapped_requirements,
            source_binding,
            label,
            contract_hash,
        )
    elif kind == "oracle":
        if not isinstance(value, dict) or "oracle" not in str(value.get("schema", "")).lower():
            raise VerificationError(f"{label} schema is not an independent date/time oracle receipt")
        schema_text = str(value.get("schema", "")).lower()
        if value.get("schema") == "ods-formula-date-time-oracle-v1" or not any(
            token in schema_text for token in ("receipt", "execution")
        ):
            raise VerificationError(
                f"{label} is an expected-vector corpus, not an executed oracle receipt"
            )
        if contract_hash is not None:
            equal(f"{label} contract hash", value.get("contract_sha256"), contract_hash)
        execution = value.get("execution")
        if not isinstance(execution, dict):
            raise VerificationError(f"{label} has no execution record")
        if execution.get("executed") is not True or execution.get("exit_code") != 0:
            raise VerificationError(f"{label} execution record is not a successful completed run")
        if execution.get("runner") not in {"cargo-test", "rust-test"}:
            raise VerificationError(f"{label} execution is not a Rust oracle test receipt")
        command = execution.get("command")
        if (
            not isinstance(command, list)
            or not command
            or not all(isinstance(item, str) and item for item in command)
            or "cargo" not in command
            or "test" not in command
        ):
            raise VerificationError(f"{label} execution command is absent")
        result_hash = execution.get("results_sha256")
        validate_sha(result_hash, f"{label} execution results hash")
        test_source_relative = execution.get("test_source")
        test_receipt_relative = execution.get("test_receipt")
        test_name = execution.get("test")
        test_receipt_hash = validate_sha(execution.get("test_receipt_sha256"), f"{label} test receipt hash")
        test_source = resolve_bound("repo", test_source_relative, f"{label} test source", allow_missing=False)
        test_receipt = resolve_bound("evidence", test_receipt_relative, f"{label} test receipt", allow_missing=False)
        equal(f"{label} test receipt hash", digest(test_receipt), test_receipt_hash)
        if not isinstance(test_name, str) or not test_name:
            raise VerificationError(f"{label} execution test name is absent")
        validate_test_receipt(test_source, test_receipt, test_receipt_relative, [test_name], label)
        functions = value.get("functions")
        if not isinstance(functions, list) or set(functions) != FUNCTION_SET:
            raise VerificationError(f"{label} does not cover exactly the 24 functions")
        rows = _receipt_rows(value, label, ("observations", "results", "rows"))
        vector_count = value.get("vector_count")
        if vector_count is not None and vector_count != len(rows):
            raise VerificationError(f"{label} vector_count does not match retained vectors")
        for index, row in enumerate(rows):
            if not any(isinstance(row.get(key), str) and row.get(key) for key in ("id", "case", "formula")):
                raise VerificationError(f"{label} oracle row {index} has no exact id")
            if row.get("function") not in FUNCTION_SET:
                raise VerificationError(f"{label} oracle row {index} has an unknown function")
            if not _typed_outcome(row):
                raise VerificationError(f"{label} oracle row {index} has no typed outcome")
        receipt_status(value, label)
    elif kind == "native":
        if not isinstance(value, (dict, list)):
            raise VerificationError(f"{label} native receipt is not structured")
        if isinstance(value, dict) and contract_hash is not None:
            equal(f"{label} contract hash", value.get("contract_sha256"), contract_hash)
        rows = _receipt_rows(value, label, ("observations", "results", "rows"))
        for index, row in enumerate(rows):
            comparison = row.get("comparison")
            if comparison == "fixture-data":
                if not isinstance(row.get("note"), str) or not row["note"].strip() or not _typed_outcome(row):
                    raise VerificationError(f"{label} fixture-data row {index} lacks note or typed value")
                continue
            if not any(isinstance(row.get(key), str) and row.get(key) for key in ("case", "id", "formula")):
                raise VerificationError(f"{label} native row {index} has no case identity")
            if not _typed_outcome(row):
                raise VerificationError(f"{label} native row {index} has no typed outcome")
            if comparison not in {"parity", "native-divergence", "host-observation"}:
                raise VerificationError(f"{label} native row {index} lacks comparison disposition")
            if comparison in {"native-divergence", "host-observation"} and not isinstance(row.get("divergence_reason"), str):
                raise VerificationError(f"{label} native divergence {index} lacks a reason")
        if isinstance(value, dict):
            receipt_status(value, label)
    elif kind == "performance":
        if not isinstance(value, (dict, list)):
            raise VerificationError(f"{label} performance receipt is not structured")
        if isinstance(value, dict):
            receipt_status(value, label)
            rows = _receipt_rows(value, label, ("captures", "candidate_only", "candidate", "observations", "results", "rows"))
        else:
            rows = _receipt_rows(value, label, ())
        if not any(
            isinstance(row.get("samples", row.get("records")), int)
            and row.get("samples", row.get("records")) > 0
            for row in rows
        ):
            raise VerificationError(f"{label} has no positive retained sample/record count")
        if any(row.get("supported") is False for row in rows):
            raise VerificationError(f"{label} contains unsupported measurements")
    elif kind == "gate":
        if isinstance(value, list):
            if not value or not all(isinstance(row, dict) and row.get("exit_code") == 0 for row in value):
                raise VerificationError(f"{label} gate rows are not all successful")
        elif isinstance(value, dict):
            receipt_status(value, label)
            if value.get("all_required_checks_passed") is not True and value.get("verified") is not True:
                status = value.get("status")
                if not isinstance(status, str) or status.lower() not in PASS_WORDS:
                    raise VerificationError(f"{label} gate receipt has no structured success disposition")
        else:
            raise VerificationError(f"{label} gate receipt is not structured")
    else:
        raise VerificationError(f"{label} has no structured validator for {kind!r}")
    available = structured_values(value)
    missing = [identifier for identifier in identifiers if identifier not in available]
    if missing:
        raise VerificationError(f"{label} identifiers are absent from structured receipt fields: {missing}")


def validate_binding_schema(binding: Any, expected: set[str], label: str) -> dict[str, Any]:
    """Validate all binding fields without requiring retained files."""

    if not isinstance(binding, dict) or set(binding) != BINDING_KEYS:
        raise VerificationError(f"{label} must contain exactly the binding fields")
    kind = binding.get("kind")
    if kind not in BINDING_KINDS:
        raise VerificationError(f"{label} has unsupported kind {kind!r}")
    root_name = binding.get("root")
    if root_name not in {"repo", "evidence"}:
        raise VerificationError(f"{label} has invalid source root {root_name!r}")
    relative_path(REPO if root_name == "repo" else HERE, binding.get("path"), f"{label} source")
    source_hash = validate_sha(binding.get("sha256"), f"{label} source hash")
    identifiers = strings(binding.get("identifiers"), f"{label} identifiers")
    requirements = strings(binding.get("requirements"), f"{label} requirements")
    unknown = set(requirements) - expected
    if unknown:
        raise VerificationError(f"{label} names unknown requirements: {sorted(unknown)}")
    receipt = binding.get("receipt")
    if not isinstance(receipt, dict) or set(receipt) != RECEIPT_KEYS:
        raise VerificationError(f"{label} receipt has malformed fields")
    if receipt.get("root") != "evidence":
        raise VerificationError(f"{label} receipt must be evidence-rooted")
    relative_path(HERE, receipt.get("path"), f"{label} receipt")
    receipt_hash = validate_sha(receipt.get("sha256"), f"{label} receipt hash")
    if kind in {"oracle", "native", "performance", "gate"} and root_name != "evidence":
        raise VerificationError(f"{label} {kind} source must be evidence-rooted")
    if kind in {"focused_test", "resource", "source_review"} and root_name != "repo":
        raise VerificationError(f"{label} {kind} source must be repository-rooted")
    source_prefix = {"native": "native/", "performance": "performance/", "gate": "gates/"}.get(kind)
    if source_prefix and not binding["path"].startswith(source_prefix):
        raise VerificationError(f"{label} {kind} source must be under {source_prefix}")
    receipt_path = receipt["path"]
    if kind in {"focused_test", "resource", "source_review"} and not receipt_path.startswith("gates/"):
        raise VerificationError(f"{label} {kind} receipt must be under gates/")
    expected_prefix = {"native": "native/", "performance": "performance/", "gate": "gates/"}.get(kind)
    if expected_prefix and not receipt_path.startswith(expected_prefix):
        raise VerificationError(f"{label} {kind} receipt must be under {expected_prefix}")
    return {
        "kind": kind,
        "root": root_name,
        "path": binding["path"],
        "sha256": source_hash,
        "identifiers": identifiers,
        "requirements": requirements,
        "receipt": receipt,
        "receipt_sha256": receipt_hash,
    }


def validate_binding(
    binding: Any,
    expected: set[str],
    label: str,
    *,
    allow_missing: bool,
    contract_hash: str | None = None,
) -> list[str]:
    """Validate binding schema, current hashes, and receipt semantics."""

    normalized = validate_binding_schema(binding, expected, label)
    pending: list[str] = []
    try:
        source = resolve_bound(normalized["root"], normalized["path"], f"{label} source", allow_missing=True)
    except PendingReceipt:
        if not allow_missing:
            raise VerificationError(f"{label} source is absent: {normalized['path']}")
        pending.append(f"{label} source is absent: {normalized['path']}")
        source = None
    try:
        receipt_path = resolve_bound("evidence", normalized["receipt"]["path"], f"{label} receipt", allow_missing=True)
    except PendingReceipt:
        if not allow_missing:
            raise VerificationError(f"{label} receipt is absent: {normalized['receipt']['path']}")
        pending.append(f"{label} receipt is absent: {normalized['receipt']['path']}")
        receipt_path = None
    if source is not None:
        equal(f"{label} source hash", digest(source), normalized["sha256"])
    if receipt_path is not None:
        equal(f"{label} receipt hash", digest(receipt_path), normalized["receipt_sha256"])
    if source is not None and receipt_path is not None:
        kind = normalized["kind"]
        if kind in {"focused_test", "resource"}:
            validate_test_receipt(source, receipt_path, normalized["receipt"]["path"], normalized["identifiers"], label)
        else:
            validate_structured_receipt(
                kind,
                receipt_path,
                normalized["receipt"]["path"],
                normalized["identifiers"],
                label,
                contract_hash,
                mapped_requirements=normalized["requirements"],
                source_binding=normalized,
            )
    return pending


def validate_evidence(
    evidence: Any,
    requirements: list[str],
    label: str,
    contract_hash: str,
    *,
    allow_missing: bool,
) -> tuple[set[str], list[str]]:
    if not isinstance(evidence, dict) or set(evidence) != EVIDENCE_KEYS:
        raise VerificationError(f"{label} must contain exactly schema/status/contract_sha256/bindings")
    equal(f"{label} schema", evidence.get("schema"), SCHEMA)
    phase = status_class(evidence.get("status"), label)
    equal(f"{label} contract hash", evidence.get("contract_sha256"), contract_hash)
    bindings = evidence.get("bindings")
    if not isinstance(bindings, list):
        raise VerificationError(f"{label} bindings must be a list")
    expected = set(requirements)
    covered: set[str] = set()
    pending: list[str] = []
    for index, binding in enumerate(bindings):
        try:
            pending.extend(
                validate_binding(
                    binding,
                    expected,
                    f"{label} binding {index}",
                    allow_missing=allow_missing,
                    contract_hash=contract_hash,
                )
            )
        except PendingReceipt as error:
            # This path is reserved for a future resolver implementation that
            # reports absence directly.  Binding schema/path validation above
            # has already run; do not repeat the same missing-file operation.
            if not allow_missing:
                raise
            pending.append(str(error))
        covered.update(binding.get("requirements", []))
    if phase == "pass":
        if not bindings:
            raise VerificationError(f"{label} PASS evidence has no bindings")
        if covered != expected:
            raise VerificationError(
                f"{label} does not cover every requirement; "
                f"missing={sorted(expected - covered)}, unknown={sorted(covered - expected)}"
            )
    elif covered - expected:
        raise VerificationError(f"{label} names unknown requirements: {sorted(covered - expected)}")
    return covered, pending


def command_json(command: list[str], cwd: Path) -> dict[str, Any]:
    completed = subprocess.run(command, cwd=cwd, capture_output=True, text=True)
    if completed.returncode:
        raise VerificationError(
            f"command failed ({completed.returncode}): {' '.join(command)}\n"
            f"{completed.stdout}{completed.stderr}"
        )
    output = completed.stdout.strip()
    for candidate in [output, *reversed(output.splitlines())]:
        try:
            value = json.loads(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    raise VerificationError(f"command did not emit a JSON object: {' '.join(command)}")


def verify_contract() -> dict[str, Any]:
    if not CONTRACT.is_file():
        raise PendingReceipt("date/time contract is absent")
    text = CONTRACT.read_text(encoding="utf-8")
    missing = [name for name in FUNCTIONS if not re.search(rf"\b{re.escape(name)}\b", text)]
    if missing:
        raise VerificationError(f"date/time contract omits functions: {missing}")
    return {"sha256": digest(CONTRACT), "functions": len(FUNCTIONS)}


def verify_reviews(contract_hash: str) -> dict[str, Any]:
    if not REVIEW_RECEIPT.is_file():
        raise PendingReceipt("date/time review-receipt.json is absent")
    receipt = read_json(REVIEW_RECEIPT)
    if not isinstance(receipt, dict) or set(receipt) != {"status", "contract_sha256", "freeze_sha256", "reviews"}:
        raise VerificationError("date/time review receipt has an unexpected schema")
    equal("date/time review status", receipt.get("status"), "PASS")
    equal("date/time review contract hash", receipt.get("contract_sha256"), contract_hash)
    freeze = GATES / "freeze.json"
    if not freeze.is_file():
        raise PendingReceipt("date/time freeze is absent for review receipt")
    equal("date/time review freeze hash", receipt.get("freeze_sha256"), digest(freeze))
    reviews = receipt.get("reviews")
    if not isinstance(reviews, dict) or set(reviews) != {"semantic", "resource"}:
        raise VerificationError("date/time review receipt does not name both independent reviews")
    for name, entry in reviews.items():
        if not isinstance(entry, dict) or set(entry) != {"path", "sha256", "status"}:
            raise VerificationError(f"date/time {name} review receipt is malformed")
        equal(f"date/time {name} review status", entry.get("status"), "PASS")
        path = relative_path(HERE, entry.get("path"), f"date/time {name} review path")
        if not path.is_file():
            raise PendingReceipt(f"date/time {name} review report is absent")
        equal(f"date/time {name} review hash", digest(path), validate_sha(entry.get("sha256"), f"date/time {name} review hash"))
    return {"status": receipt["status"], "contract_sha256": contract_hash}


def verify_oracle_bundle(contract_hash: str) -> dict[str, Any]:
    # The expected corpus is deliberately not enough.  Wait for a distinct
    # executed receipt before reading the corpus into an implementation claim.
    execution_path = next((path for path in ORACLE_EXECUTION_CANDIDATES if path.is_file()), None)
    if execution_path is None:
        raise PendingReceipt("executed date/time oracle receipt is absent; expected vectors are not a receipt")
    if not ORACLE_CORPUS.is_file():
        raise PendingReceipt("date/time oracle expected-vector corpus is absent")
    corpus = validate_expected_oracle_corpus(ORACLE_CORPUS, contract_hash)
    value = load_structured_receipt(execution_path, "date/time executed oracle receipt")
    rows = _receipt_rows(value, "date/time executed oracle receipt", ("observations", "results", "rows"))
    corpus_value = load_structured_receipt(ORACLE_CORPUS, "date/time oracle vector corpus")
    expected_ids = [row["id"] for row in corpus_value["vectors"]]
    validate_structured_receipt(
        "oracle",
        execution_path,
        str(execution_path.relative_to(HERE)),
        expected_ids,
        "date/time executed oracle receipt",
        contract_hash,
    )
    return {"corpus": corpus, "execution_sha256": digest(execution_path), "observations": len(rows)}


def verify_native_bundle(contract_hash: str) -> dict[str, Any]:
    results_path = NATIVE / "native-results.json"
    provenance_path = NATIVE / "provenance.json"
    if not results_path.is_file() or not provenance_path.is_file():
        raise PendingReceipt("date/time native results or provenance is absent")
    value = load_structured_receipt(results_path, "date/time native results")
    rows = _receipt_rows(value, "date/time native results", ("observations", "results", "rows"))
    identifiers = [row.get("id", row.get("case", row.get("formula"))) for row in rows]
    identifiers = [item for item in identifiers if isinstance(item, str) and item]
    validate_structured_receipt(
        "native",
        results_path,
        str(results_path.relative_to(HERE)),
        identifiers,
        "date/time native results",
        contract_hash,
    )
    provenance = read_json(provenance_path)
    if not isinstance(provenance, dict):
        raise VerificationError("date/time native provenance is not an object")
    equal("date/time native provenance contract hash", provenance.get("contract_sha256"), contract_hash)
    if not ORACLE_CORPUS.is_file():
        raise PendingReceipt("date/time oracle corpus is absent for native provenance")
    equal("date/time native provenance oracle hash", provenance.get("oracle_sha256"), digest(ORACLE_CORPUS))
    equal("date/time native provenance results hash", provenance.get("native_results_sha256"), digest(results_path))
    fixture = provenance.get("fixture")
    if not isinstance(fixture, dict):
        raise VerificationError("date/time native provenance fixture identity is absent")
    for key in ("input", "recalculated_output", "content_xml"):
        relative = fixture.get(key)
        path = relative_path(NATIVE, relative, f"date/time native fixture {key}")
        if not path.is_file():
            raise PendingReceipt(f"date/time native fixture is absent: {relative}")
        checksum_key = f"{key}_sha256"
        equal(
            f"date/time native fixture {key} hash",
            fixture.get(checksum_key),
            digest(path),
        )
    receipt_status(provenance, "date/time native provenance")
    return {"results_sha256": digest(results_path), "rows": len(rows)}


def verify_gate_bundle() -> dict[str, Any]:
    gate_verify = GATES / "verify.py"
    if not gate_verify.is_file():
        raise PendingReceipt("date/time gate verifier is absent")
    required = (
        "freeze.json",
        "stage-manifest.json",
        "staged-profile-sources.json",
        "batch-files.json",
        "environment.json",
        "source-before.json",
        "source-after.json",
        "results.json",
        "verification.json",
        "ods-tests.log",
    )
    missing = [name for name in required if not (GATES / name).is_file()]
    if missing:
        raise PendingReceipt("date/time gate receipts are absent: " + ", ".join(missing))
    receipt = command_json(["python3", str(gate_verify)], REPO)
    equal("date/time gate status", receipt.get("status"), "ok")
    equal("date/time gate verified", receipt.get("verified"), True)
    return receipt


def verify_performance_bundle(contract_hash: str) -> dict[str, Any]:
    results = PERFORMANCE / "results"
    required = (
        "capture-summary.json",
        "performance-report.json",
        "profile-inputs-before.json",
        "profile-inputs-after.json",
        "retained-files.json",
    )
    missing = [name for name in required if not (results / name).is_file()]
    if missing:
        raise PendingReceipt("date/time performance receipts are absent: " + ", ".join(missing))
    before = read_json(results / "profile-inputs-before.json")
    after = read_json(results / "profile-inputs-after.json")
    if not isinstance(before, dict) or not before or after != before:
        raise VerificationError("date/time performance profile input custody is malformed")
    for relative, checksum in before.items():
        path = relative_path(PERFORMANCE, relative, "date/time performance profile input")
        if not path.is_file() or digest(path) != checksum:
            raise VerificationError(f"date/time performance profile input hash mismatch: {relative}")
    summary = read_json(results / "capture-summary.json")
    if not isinstance(summary, dict) or not isinstance(summary.get("captures"), list) or not summary["captures"]:
        raise VerificationError("date/time performance capture summary is malformed")
    if any(not isinstance(row, dict) or not isinstance(row.get("records"), int) or row["records"] <= 0 for row in summary["captures"]):
        raise VerificationError("date/time performance capture summary has no positive records")
    retained = read_json(results / "retained-files.json")
    if not isinstance(retained, dict) or not retained:
        raise VerificationError("date/time retained-files manifest is malformed")
    actual = {str(path.relative_to(results)) for path in results.rglob("*") if path.is_file() and path.name != "retained-files.json"}
    equal("date/time retained performance file set", set(retained), actual)
    for relative, checksum in retained.items():
        path = relative_path(results, relative, "date/time retained performance file")
        if not path.is_file() or digest(path) != checksum:
            raise VerificationError(f"date/time retained performance hash mismatch: {relative}")
    report = results / "performance-report.json"
    value = load_structured_receipt(report, "date/time performance report")
    validate_structured_receipt("performance", report, str(report.relative_to(HERE)), [], "date/time performance report", contract_hash)
    audit = HERE / "root-performance-audit.json"
    if not audit.is_file():
        raise PendingReceipt("date/time root performance audit is absent")
    audit_value = read_json(audit)
    if not isinstance(audit_value, dict) or audit_value.get("status") not in {"ok", "verified"}:
        raise VerificationError("date/time root performance audit has no success status")
    return {"report_sha256": digest(report), "captures": len(summary["captures"]), "audit": audit_value.get("status")}


def run_check(name: str, function, checks: dict[str, Any], pending: list[str]) -> None:
    try:
        checks[name] = function()
    except PendingReceipt as error:
        pending.append(f"{name}: {error}")


def verify_coverage() -> dict[str, Any]:
    manifest = read_json(MANIFEST)
    if not isinstance(manifest, dict) or set(manifest) != MANIFEST_KEYS:
        raise VerificationError("coverage-requirements.json has unexpected fields")
    equal("coverage schema", manifest.get("schema"), SCHEMA)
    equal("coverage normative", manifest.get("normative"), NORMATIVE)
    functions = manifest.get("functions")
    if not isinstance(functions, dict) or tuple(functions) != FUNCTIONS:
        raise VerificationError("coverage manifest does not list exactly the 24 functions in order")
    for name in FUNCTIONS:
        entry = functions[name]
        if not isinstance(entry, dict) or set(entry) != {"requirements", "evidence"}:
            raise VerificationError(f"coverage function {name} is malformed")
        requirements = strings(entry.get("requirements"), f"coverage {name} requirements")
        evidence = entry.get("evidence")
        if evidence is None:
            raise PendingReceipt(f"coverage {name} evidence is absent")
    cross = manifest.get("cross_cutting")
    strings(cross, "coverage cross_cutting")
    if len(cross) != len(set(cross)):
        raise VerificationError("coverage cross_cutting requirements contain duplicates")
    cross_evidence = manifest.get("cross_cutting_evidence")
    if not isinstance(cross_evidence, list) or len(cross_evidence) != len(cross):
        raise VerificationError("coverage cross_cutting_evidence does not match scope length")
    if [item.get("requirement") if isinstance(item, dict) else None for item in cross_evidence] != cross:
        raise VerificationError("coverage cross_cutting_evidence order or requirement identity differs")
    scope_receipt = validate_scope(manifest)
    if not CONTRACT.is_file():
        raise PendingReceipt("date/time contract is absent")
    contract_hash = digest(CONTRACT)
    pending: list[str] = []
    function_bindings = 0
    for name in FUNCTIONS:
        evidence = functions[name]["evidence"]
        try:
            covered, missing = validate_evidence(
                evidence,
                functions[name]["requirements"],
                f"coverage {name}",
                contract_hash,
                allow_missing=True,
            )
            function_bindings += len(evidence["bindings"])
            pending.extend(missing)
            if status_class(evidence.get("status"), f"coverage {name}") == "pending":
                pending.append(f"coverage {name} evidence is pending")
        except PendingReceipt as error:
            pending.append(str(error))
    seen: set[str] = set()
    cross_bindings = 0
    for index, item in enumerate(cross_evidence):
        if not isinstance(item, dict) or set(item) != CROSS_EVIDENCE_KEYS:
            raise VerificationError(f"cross-cutting evidence {index} is malformed")
        requirement = item["requirement"]
        if requirement in seen:
            raise VerificationError(f"duplicate cross-cutting evidence {requirement!r}")
        seen.add(requirement)
        try:
            covered, missing = validate_evidence(
                item["evidence"],
                [requirement],
                f"cross-cutting evidence {index}",
                contract_hash,
                allow_missing=True,
            )
            cross_bindings += len(item["evidence"]["bindings"])
            pending.extend(missing)
            if status_class(item["evidence"].get("status"), f"cross-cutting evidence {index}") == "pending":
                pending.append(f"cross-cutting {requirement} evidence is pending")
        except PendingReceipt as error:
            pending.append(str(error))
    if seen != set(cross):
        raise VerificationError("cross-cutting evidence is incomplete")
    manifest_status = status_class(manifest.get("status"), "coverage manifest")
    if manifest_status == "pending":
        pending.append("coverage manifest evidence is pending")
    elif pending:
        raise VerificationError("coverage manifest is PASS but evidence is pending")
    if pending:
        raise PendingReceipt("; ".join(dict.fromkeys(pending)))
    if manifest_status != "pass":
        raise VerificationError("coverage manifest did not reach PASS")
    return {
        "sha256": digest(MANIFEST),
        "scope_sha256": scope_receipt["sha256"],
        "contract_sha256": contract_hash,
        "functions": len(FUNCTIONS),
        "function_bindings": function_bindings,
        "cross_cutting": len(cross),
        "cross_cutting_bindings": cross_bindings,
        "schema": SCHEMA,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--allow-pending",
        action="store_true",
        help="report preparatory gaps as pending; never report verified=true",
    )
    args = parser.parse_args()
    checks: dict[str, Any] = {}
    pending: list[str] = []
    run_check("contract", verify_contract, checks, pending)
    contract = checks.get("contract")
    if isinstance(contract, dict):
        contract_hash = contract["sha256"]
        run_check("coverage", verify_coverage, checks, pending)
        run_check("reviews", lambda: verify_reviews(contract_hash), checks, pending)
        run_check("oracle", lambda: verify_oracle_bundle(contract_hash), checks, pending)
        run_check("native", lambda: verify_native_bundle(contract_hash), checks, pending)
        run_check("gates", verify_gate_bundle, checks, pending)
        run_check("performance", lambda: verify_performance_bundle(contract_hash), checks, pending)
    else:
        pending.extend(
            [
                "coverage: waiting for contract identity",
                "reviews: waiting for contract identity",
                "oracle: waiting for contract identity",
                "native: waiting for contract identity",
                "gates: waiting for contract identity",
                "performance: waiting for contract identity",
            ]
        )
    if pending and not args.allow_pending:
        raise VerificationError("; ".join(dict.fromkeys(pending)))
    if pending:
        result = {"status": "pending", "verified": False, "pending": list(dict.fromkeys(pending)), "checks": checks}
    else:
        result = {"status": "ok", "verified": True, "checks": checks}
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, VerificationError) as error:
        raise SystemExit(f"date/time evidence verification failed: {error}")
