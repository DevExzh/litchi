#!/usr/bin/env python3
"""Fail-closed verifier for the ODS lookup-function evidence bundle.

The bundle is assembled in stages.  ``--allow-pending`` is intended for the
pre-capture state and always reports ``verified: false``.  It records missing
receipts as pending, but never turns a present malformed or failing receipt
into a pass.  Counts and per-case read expectations are taken from retained
inputs; this module contains no copied test, oracle, native, or sample totals.
"""

from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Callable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
GATES = HERE / "gates"
NATIVE = HERE / "native"
PERFORMANCE = HERE / "performance"
CONTRACT = HERE / "contract.md"
SEMANTIC_REVIEW = HERE / "semantic-review.md"
RESOURCE_REVIEW = HERE / "resource-review.md"
REVIEW_RECEIPT = HERE / "review-receipt.json"
COVERAGE_REQUIREMENTS = HERE / "coverage-requirements.json"
COVERAGE_SCOPE = HERE / "coverage-scope.json"


class VerificationError(RuntimeError):
    """A present receipt is malformed, stale, or unsuccessful."""


class PendingReceipt(RuntimeError):
    """A receipt has not been produced yet."""


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def read_json(path: Path) -> Any:
    if not path.is_file():
        raise VerificationError(f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise VerificationError(f"invalid JSON evidence {path}: {error}") from error


def load_baseline() -> dict[str, Any]:
    value = read_json(HERE / "baseline.json")
    if not isinstance(value, dict):
        raise VerificationError("baseline.json is not an object")
    return value


BASELINE = load_baseline()
BASELINE_COMMIT = BASELINE.get("commit")
FUNCTIONS = BASELINE.get("scope")
GATE_LOCK_SHA256 = BASELINE.get("gate_lock_sha256")
AMBIENT_LOCK_SHA256 = BASELINE.get("ambient_lock_sha256")
EXPECTED_FUNCTIONS = {
    "ADDRESS",
    "CHOOSE",
    "HLOOKUP",
    "INDEX",
    "INDIRECT",
    "LOOKUP",
    "MATCH",
    "OFFSET",
    "VLOOKUP",
}
if (
    not isinstance(BASELINE_COMMIT, str)
    or not BASELINE_COMMIT
    or not isinstance(FUNCTIONS, list)
    or set(FUNCTIONS) != EXPECTED_FUNCTIONS
    or len(FUNCTIONS) != len(EXPECTED_FUNCTIONS)
    or not all(isinstance(name, str) and name for name in FUNCTIONS)
    or not isinstance(GATE_LOCK_SHA256, str)
    or not re.fullmatch(r"[0-9a-f]{64}", GATE_LOCK_SHA256)
    or not isinstance(AMBIENT_LOCK_SHA256, str)
    or not re.fullmatch(r"[0-9a-f]{64}", AMBIENT_LOCK_SHA256)
):
    raise VerificationError("baseline.json has an invalid nine-function or lock identity")
FUNCTIONS = tuple(FUNCTIONS)

SUMMARY = re.compile(r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored")
INCLUDE = re.compile(r'include_(?:bytes|str)!\(\s*"([^"\n]+)')

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
CALCEXT = "{urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0}"
FORMULA = TABLE + "formula"
VALUE_TYPE = OFFICE + "value-type"
VALUE = OFFICE + "value"
BOOLEAN_VALUE = OFFICE + "boolean-value"


def equal(label: str, observed: Any, expected: Any) -> None:
    if observed != expected:
        raise VerificationError(f"{label}: expected {expected!r}, observed {observed!r}")


def command_json(command: list[str], *, cwd: Path) -> dict[str, Any]:
    completed = subprocess.run(command, cwd=cwd, capture_output=True, text=True)
    if completed.returncode:
        raise VerificationError(
            f"command failed ({completed.returncode}): {' '.join(command)}\n"
            f"{completed.stdout}{completed.stderr}"
        )
    output = completed.stdout.strip()
    if not output:
        raise VerificationError(f"command emitted no JSON: {' '.join(command)}")
    for candidate in [output, *reversed(output.splitlines())]:
        try:
            value = json.loads(candidate)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    raise VerificationError(f"command did not emit a JSON object: {' '.join(command)}\n{output}")


def first_value(mapping: dict[str, Any], names: tuple[str, ...]) -> Any:
    for name in names:
        if name in mapping:
            return mapping[name]
    return None


def safe_child(root: Path, relative: Any, label: str) -> Path:
    if not isinstance(relative, str) or not relative or Path(relative).is_absolute():
        raise VerificationError(f"{label} is not a relative path: {relative!r}")
    path = (root / relative).resolve()
    try:
        path.relative_to(root.resolve())
    except ValueError as error:
        raise VerificationError(f"{label} escapes its root: {relative!r}") from error
    return path


def validate_repo_relative(relative: Any, label: str) -> str:
    if not isinstance(relative, str) or not relative or Path(relative).is_absolute():
        raise VerificationError(f"{label} is not repository-relative: {relative!r}")
    if ".." in Path(relative).parts:
        raise VerificationError(f"{label} escapes the repository: {relative!r}")
    return relative


# Coverage evidence is deliberately a small, closed schema.  A prose
# requirement is covered only by a binding that identifies the source used,
# the retained receipt used to verify it, and the case/test identifiers found
# in that receipt.  This keeps a planning note or an arbitrary dictionary from
# becoming an accidental approval.
COVERAGE_SCHEMA = "ods-formula-lookups-coverage-v1"
COVERAGE_SCOPE_SCHEMA = "ods-formula-lookups-coverage-scope-v1"
COVERAGE_KINDS = {
    "focused_test",
    "oracle",
    "native",
    "resource",
    "performance",
    "gate",
}
BINDING_KEYS = {"kind", "root", "path", "sha256", "identifiers", "requirements", "receipt"}
RECEIPT_KEYS = {"root", "path", "sha256"}
EVIDENCE_KEYS = {"schema", "status", "contract_sha256", "bindings"}
CROSS_EVIDENCE_KEYS = {"requirement", "evidence"}
COVERAGE_SCOPE_KEYS = {"schema", "normative", "functions", "cross_cutting"}


def resolve_bound_file(root_name: Any, relative: Any, label: str) -> Path:
    """Resolve a retained evidence path without allowing root escape.

    ``repo`` is the checked-out source tree and ``evidence`` is this frozen
    lookup evidence directory.  Receipts must use the latter root, while a
    source binding may use either root.  Missing files remain pending so a
    preparatory run can report incomplete capture; traversal and malformed
    roots are always verification failures.
    """

    if root_name == "repo":
        root = REPO
    elif root_name == "evidence":
        root = HERE
    else:
        raise VerificationError(f"{label} has an invalid root: {root_name!r}")
    path = safe_child(root, relative, label)
    if not path.is_file():
        raise PendingReceipt(f"{label} is absent: {relative}")
    return path


def validate_digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or not re.fullmatch(r"[0-9a-f]{64}", value):
        raise VerificationError(f"{label} is not a lowercase SHA-256 digest")
    return value


def validate_identifiers(value: Any, label: str) -> list[str]:
    if not isinstance(value, list) or not value:
        raise VerificationError(f"{label} must be a non-empty unique string list")
    if not all(
        isinstance(item, str) and item.strip() and "\n" not in item and "\r" not in item
        for item in value
    ):
        raise VerificationError(f"{label} contains an invalid case/test identifier")
    if len(value) != len(set(value)):
        raise VerificationError(f"{label} contains duplicate identifiers")
    return value


def positive_integer(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value > 0


def nonnegative_integer(value: Any) -> bool:
    return isinstance(value, int) and not isinstance(value, bool) and value >= 0


TEST_RESULT = re.compile(r"^\s*test\s+(\S+)\s+\.\.\.\s+(ok|FAILED|ignored)\s*$", re.MULTILINE)
TEST_SUMMARY = re.compile(
    r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored"
)
ANSI_ESCAPE = re.compile(r"\x1b\[[0-?]*[ -/]*[@-~]")
RUST_TEST_ATTRIBUTE = re.compile(
    r"#\[\s*(?:(?:[A-Za-z_][A-Za-z0-9_]*::)*)test(?:\W|$)"
)
RUST_FUNCTION = re.compile(
    r"\b(?:pub(?:\([^)]*\))?\s+)?(?:async\s+)?fn\s+([A-Za-z_][A-Za-z0-9_]*)\s*\("
)
STRUCTURED_IDENTIFIER_KEYS = {
    "case",
    "id",
    "name",
    "function",
    "formula",
    "label",
    "test",
    "tests",
    "functions",
    "scope",
}
PASS_STATUSES = {"ok", "pass", "passed", "verified", "complete", "captured"}
FAIL_STATUS_WORDS = {"fail", "failed", "failure", "error", "errors", "pending", "hold", "incomplete"}


def rust_test_names(source: str) -> set[str]:
    """Return functions explicitly marked with a Rust test attribute."""

    lines = source.splitlines()
    names: set[str] = set()
    for index, line in enumerate(lines):
        attribute = RUST_TEST_ATTRIBUTE.search(line)
        if attribute is None:
            continue
        # Attributes such as cfg or tokio::test may sit between the test
        # attribute and the function declaration.  Keep this window bounded
        # so a later helper function cannot be mistaken for the test.
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
    """Bind focused/resource evidence to real Rust tests and passing lines."""

    if source.suffix != ".rs":
        raise VerificationError(f"{label} test source is not a Rust file")
    try:
        source_text = source.read_text(encoding="utf-8")
    except UnicodeDecodeError as error:
        raise VerificationError(f"{label} test source is not UTF-8 Rust text") from error
    defined = rust_test_names(source_text)
    if not defined:
        raise VerificationError(f"{label} source defines no #[test] function")
    if not receipt_relative.startswith("gates/") or receipt_path.suffix not in {".log", ".txt"}:
        raise VerificationError(f"{label} test receipt must be a retained gates text log")
    text = ANSI_ESCAPE.sub("", receipt_path.read_bytes().decode("utf-8", errors="replace"))
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    results = TEST_RESULT.findall(text)
    if not results:
        raise VerificationError(f"{label} test receipt has no Cargo test result lines")
    summaries = TEST_SUMMARY.findall(text)
    if any(status != "ok" or int(failed) != 0 or int(ignored) != 0 for status, _, failed, ignored in summaries):
        raise VerificationError(f"{label} test receipt contains a failed or ignored test summary")
    for identifier in identifiers:
        short = short_test_name(identifier)
        if short not in defined:
            raise VerificationError(f"{label} source does not define Rust test {identifier!r}")
        matches = [
            (name, status)
            for name, status in results
            if name == identifier or name == short or name.endswith(f"::{short}")
        ]
        if not matches:
            raise VerificationError(f"{label} receipt has no exact result line for test {identifier!r}")
        bad = [(name, status) for name, status in matches if status != "ok"]
        if bad:
            raise VerificationError(f"{label} test {identifier!r} is not passing: {bad}")


def structured_values(value: Any, key: str | None = None) -> set[str]:
    """Collect identifiers only from known structured receipt fields."""

    found: set[str] = set()
    if isinstance(value, dict):
        for child_key, child in value.items():
            if child_key in {"reference_reads", "results_by_case", "cases_by_name"} and isinstance(child, dict):
                found.update(key for key in child if isinstance(key, str) and key.strip())
            if child_key in STRUCTURED_IDENTIFIER_KEYS:
                if isinstance(child, str) and child.strip():
                    found.add(child)
                elif isinstance(child, list):
                    found.update(item for item in child if isinstance(item, str) and item.strip())
            found.update(structured_values(child, child_key))
    elif isinstance(value, list):
        for child in value:
            found.update(structured_values(child, key))
    return found


def receipt_status(value: dict[str, Any], label: str) -> None:
    if value.get("verified") is False:
        raise VerificationError(f"{label} structured receipt is not verified")
    status = value.get("status")
    if status is None:
        return
    if not isinstance(status, str):
        raise VerificationError(f"{label} structured receipt has failure status: {status!r}")
    normalized = status.lower().replace("-", " ").strip()
    failed_prefix = tuple(f"{word} " for word in FAIL_STATUS_WORDS)
    if normalized in FAIL_STATUS_WORDS or normalized.startswith(failed_prefix):
        raise VerificationError(f"{label} structured receipt has failure status: {status!r}")


def load_structured_receipt(path: Path, label: str) -> Any:
    if path.suffix == ".jsonl":
        rows: list[Any] = []
        for line_number, line in enumerate(path.read_text(encoding="utf-8").splitlines(), 1):
            if not line.strip():
                continue
            try:
                rows.append(json.loads(line))
            except json.JSONDecodeError as error:
                raise VerificationError(f"{label} JSONL line {line_number} is malformed: {error}") from error
        if not rows:
            raise VerificationError(f"{label} structured receipt is empty")
        return rows
    if path.suffix != ".json":
        raise VerificationError(f"{label} receipt is not a JSON/JSONL structured receipt")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise VerificationError(f"{label} structured receipt is malformed: {error}") from error


def validate_structured_receipt(
    kind: str,
    receipt_path: Path,
    receipt_relative: str,
    identifiers: list[str],
    label: str,
) -> None:
    """Validate the known JSON receipt shape for non-test evidence kinds."""

    value = load_structured_receipt(receipt_path, label)
    if kind == "oracle":
        if not isinstance(value, dict) or "oracle" not in str(value.get("schema", "")).lower():
            raise VerificationError(f"{label} oracle receipt schema is not independent-oracle JSON")
        observations = value.get("observations")
        functions = value.get("functions")
        if not isinstance(observations, list) or not observations:
            raise VerificationError(f"{label} oracle receipt has no observations")
        if not isinstance(functions, list) or not all(isinstance(item, str) for item in functions) or set(functions) != set(FUNCTIONS):
            raise VerificationError(f"{label} oracle receipt has no function scope")
        for index, row in enumerate(observations):
            if not isinstance(row, dict) or not any(
                isinstance(row.get(key), str) and row.get(key)
                for key in ("case", "id", "formula")
            ):
                raise VerificationError(f"{label} oracle observation {index} has no case identity")
            if not any(
                key in row
                for key in ("expected", "result", "value", "error", "kind", "type", "expected_type")
            ):
                raise VerificationError(f"{label} oracle observation {index} has no typed outcome")
        receipt_status(value, label)
    elif kind == "native":
        if not isinstance(value, (dict, list)):
            raise VerificationError(f"{label} native receipt is not a result object/list")
        if isinstance(value, list):
            native_rows = value
        else:
            native_rows = next((value.get(key) for key in ("observations", "results", "rows") if key in value), None)
        if not isinstance(native_rows, list) or not native_rows:
            raise VerificationError(f"{label} native receipt has no typed result rows")
        for index, row in enumerate(native_rows):
            if not isinstance(row, dict):
                raise VerificationError(f"{label} native result {index} is not a structured row")
            comparison = row.get("comparison")
            # A native workbook may include cells used only to seed the resolver
            # fixture.  They are retained in the capture, but are not function
            # observations and therefore have no case identifier.  Require an
            # explicit disposition and explanation so arbitrary rows cannot
            # silently evade function coverage.
            if comparison == "fixture-data":
                if (
                    row.get("case") not in (None, "")
                    or not isinstance(row.get("note"), str)
                    or not row["note"].strip()
                ):
                    raise VerificationError(f"{label} fixture-data row {index} lacks an explicit note")
                if not any(
                    key in row
                    for key in ("native", "expected", "result", "value", "error", "kind", "type")
                ):
                    raise VerificationError(f"{label} fixture-data row {index} has no typed value")
                continue
            if not any(
                isinstance(row.get(key), str) and row.get(key)
                for key in ("case", "id", "formula")
            ):
                raise VerificationError(f"{label} native result {index} has no case identity")
            if not any(
                key in row
                for key in ("native", "expected", "result", "value", "error", "kind", "type")
            ):
                raise VerificationError(f"{label} native result {index} has no typed outcome")
            if comparison not in {"parity", "native-divergence"}:
                raise VerificationError(
                    f"{label} native result {index} has no recognized comparison disposition"
                )
            reason = row.get("divergence_reason")
            if comparison == "native-divergence" and (not isinstance(reason, str) or not reason.strip()):
                raise VerificationError(f"{label} native result {index} lacks a documented divergence reason")
            if comparison == "parity" and reason not in (None, ""):
                raise VerificationError(f"{label} parity result {index} carries an unexpected divergence reason")
        if isinstance(value, dict):
            receipt_status(value, label)
            divergences = value.get("divergences", value.get("documented_divergences", []))
            if divergences is not None and not isinstance(divergences, list):
                raise VerificationError(f"{label} native divergences are not a list")
            if divergences:
                if not all(isinstance(item, dict) for item in divergences):
                    raise VerificationError(f"{label} native divergences are not structured records")
                disposition = " ".join(
                    str(value.get(key, ""))
                    for key in ("disposition", "divergence_disposition", "comparison_disposition")
                ).lower()
                if "document" not in disposition or (
                    str(value.get("status", "")).lower() in PASS_STATUSES
                    and str(value.get("harness_status", value.get("reproduction_status", ""))).lower()
                    not in PASS_STATUSES
                ):
                    raise VerificationError(
                        f"{label} native divergences require an explicit documented disposition and harness status"
                    )
    elif kind == "performance":
        if isinstance(value, dict):
            receipt_status(value, label)
            captures = value.get("captures")
            if captures is not None:
                if not isinstance(captures, list) or not captures or any(
                    not isinstance(item, dict) or not positive_integer(item.get("records")) for item in captures
                ):
                    raise VerificationError(f"{label} performance capture receipt is incomplete")
            elif "candidate_only" in value or "candidate" in value:
                rows = value.get("candidate_only", value.get("candidate"))
                if not isinstance(rows, list) or not rows or not all(isinstance(item, dict) for item in rows):
                    raise VerificationError(f"{label} performance candidate measurements are malformed")
            elif "reference_reads" in value:
                reads = value.get("reference_reads")
                if not isinstance(reads, dict) or not reads or not all(
                    isinstance(key, str) and nonnegative_integer(item) for key, item in reads.items()
                ):
                    raise VerificationError(f"{label} performance read receipt is malformed")
            elif not any(key in value for key in ("observations", "results", "rows")):
                raise VerificationError(f"{label} performance receipt has no recognized structured measurements")
        elif not isinstance(value, list) or not value or not all(isinstance(item, dict) for item in value):
            raise VerificationError(f"{label} performance receipt is not structured measurement data")
        if isinstance(value, list):
            if any(item.get("supported") is False for item in value):
                raise VerificationError(f"{label} performance receipt contains unsupported measurements")
            for item in value:
                if "status" in item:
                    receipt_status({"status": item["status"]}, label)
    elif kind == "gate":
        if isinstance(value, list):
            if not value or not all(isinstance(item, dict) for item in value):
                raise VerificationError(f"{label} gate receipt is not a result list")
            for item in value:
                if "exit_code" not in item or item.get("exit_code") != 0:
                    raise VerificationError(f"{label} gate receipt contains a nonzero command result")
        elif isinstance(value, dict):
            receipt_status(value, label)
            if value.get("all_required_checks_passed") is False:
                raise VerificationError(f"{label} gate receipt reports failed checks")
            status = value.get("status")
            normalized_status = status.lower().replace("-", " ") if isinstance(status, str) else ""
            if value.get("verified") is not True and value.get("all_required_checks_passed") is not True and normalized_status not in PASS_STATUSES:
                raise VerificationError(f"{label} gate receipt has no success disposition")
        else:
            raise VerificationError(f"{label} gate receipt is not structured JSON")
    else:
        raise VerificationError(f"{label} has no structured receipt validator for kind {kind!r}")

    available = structured_values(value)
    missing = [identifier for identifier in identifiers if identifier not in available]
    if missing:
        raise VerificationError(f"{label} identifiers are absent from structured receipt fields: {missing}")


def validate_kind_receipt(
    kind: str,
    source: Path,
    source_relative: str,
    receipt_path: Path,
    receipt_relative: str,
    identifiers: list[str],
    label: str,
) -> None:
    if kind in {"focused_test", "resource"}:
        validate_test_receipt(source, receipt_path, receipt_relative, identifiers, label)
        return
    if ".." in Path(source_relative).parts or ".." in Path(receipt_relative).parts:
        raise VerificationError(f"{label} path escapes its kind boundary")
    expected_prefix = {
        "oracle": "",
        "native": "native/",
        "performance": "performance/",
        "gate": "gates/",
    }[kind]
    if expected_prefix and not source_relative.startswith(expected_prefix):
        raise VerificationError(f"{label} {kind} source is outside {expected_prefix}")
    if expected_prefix and not receipt_relative.startswith(expected_prefix):
        raise VerificationError(f"{label} {kind} receipt is outside {expected_prefix}")
    if kind == "oracle" and (source.suffix != ".py" or not source.name.endswith("_oracle.py")):
        raise VerificationError(f"{label} oracle source is not Python oracle code")
    validate_structured_receipt(kind, receipt_path, receipt_relative, identifiers, label)


def validate_binding(binding: Any, expected_requirements: set[str], label: str) -> dict[str, Any]:
    if not isinstance(binding, dict) or set(binding) != BINDING_KEYS:
        raise VerificationError(
            f"{label} must contain exactly kind/root/path/sha256/identifiers/requirements/receipt"
        )
    kind = binding.get("kind")
    if kind not in COVERAGE_KINDS:
        raise VerificationError(f"{label} has an unsupported binding kind: {kind!r}")
    root = binding.get("root")
    source = resolve_bound_file(root, binding.get("path"), f"{label} source")
    source_hash = validate_digest(binding.get("sha256"), f"{label} source hash")
    equal(f"{label} source hash", digest(source), source_hash)
    identifiers = validate_identifiers(binding.get("identifiers"), f"{label} identifiers")
    requirements = validate_identifiers(binding.get("requirements"), f"{label} requirements")
    unknown_requirements = set(requirements) - expected_requirements
    if unknown_requirements:
        raise VerificationError(f"{label} names unknown requirements: {sorted(unknown_requirements)}")
    receipt = binding.get("receipt")
    if not isinstance(receipt, dict) or set(receipt) != RECEIPT_KEYS:
        raise VerificationError(f"{label} receipt must contain exactly root/path/sha256")
    if receipt.get("root") != "evidence":
        raise VerificationError(f"{label} receipt must be rooted at the evidence directory")
    receipt_path = resolve_bound_file(receipt.get("root"), receipt.get("path"), f"{label} receipt")
    receipt_hash = validate_digest(receipt.get("sha256"), f"{label} receipt hash")
    equal(f"{label} receipt hash", digest(receipt_path), receipt_hash)
    # These bindings are evidence generated in this bundle.  Keeping their
    # root explicit prevents a repository path with a convenient filename from
    # masquerading as an oracle/native/performance/gate receipt.
    if kind in {"oracle", "native", "performance", "gate"} and root != "evidence":
        raise VerificationError(f"{label} {kind} source must be rooted at the evidence directory")
    source_relative = binding["path"]
    receipt_relative = receipt["path"]
    validate_kind_receipt(
        kind,
        source,
        source_relative,
        receipt_path,
        receipt_relative,
        identifiers,
        label,
    )
    return {
        "kind": kind,
        "root": root,
        "path": binding["path"],
        "identifiers": identifiers,
        "requirements": requirements,
        "receipt": receipt["path"],
    }


def validate_evidence(value: Any, requirements: list[str], label: str, contract_hash: str) -> dict[str, Any]:
    if not isinstance(value, dict) or set(value) != EVIDENCE_KEYS:
        raise VerificationError(
            f"{label} must contain exactly schema/status/contract_sha256/bindings"
        )
    equal(f"{label} schema", value.get("schema"), COVERAGE_SCHEMA)
    equal(f"{label} status", value.get("status"), "PASS")
    equal(f"{label} contract hash", value.get("contract_sha256"), contract_hash)
    bindings = value.get("bindings")
    if not isinstance(bindings, list) or not bindings:
        raise VerificationError(f"{label} bindings are empty or malformed")
    expected = set(requirements)
    covered: set[str] = set()
    normalized: list[dict[str, Any]] = []
    for index, binding in enumerate(bindings):
        normalized_binding = validate_binding(binding, expected, f"{label} binding {index}")
        covered.update(normalized_binding["requirements"])
        normalized.append(normalized_binding)
    if covered != expected:
        raise VerificationError(
            f"{label} does not bind every requirement; missing={sorted(expected - covered)}, "
            f"unknown={sorted(covered - expected)}"
        )
    return {"bindings": len(normalized), "requirements": len(requirements)}


def verify_locks() -> dict[str, str]:
    gate = GATES / "Cargo.lock"
    ambient = REPO / "Cargo.lock"
    if not gate.is_file() or not ambient.is_file():
        raise PendingReceipt("gate or ambient Cargo.lock is absent")
    equal("retained isolated gate lock", digest(gate), GATE_LOCK_SHA256)
    equal("ambient workspace lock", digest(ambient), AMBIENT_LOCK_SHA256)
    return {"gate": GATE_LOCK_SHA256, "ambient": AMBIENT_LOCK_SHA256}


def verify_contract() -> dict[str, Any]:
    if not CONTRACT.is_file():
        raise PendingReceipt("lookup contract.md is absent")
    text = CONTRACT.read_text(encoding="utf-8")
    missing = [name for name in FUNCTIONS if not re.search(rf"\b{re.escape(name)}\b", text)]
    if missing:
        raise VerificationError(f"lookup contract omits functions: {missing}")
    expected_hash = BASELINE.get("contract_sha256")
    if not isinstance(expected_hash, str) or not re.fullmatch(r"[0-9a-f]{64}", expected_hash):
        raise VerificationError("baseline.json does not pin contract.md")
    equal("lookup contract hash", digest(CONTRACT), expected_hash)
    return {"sha256": expected_hash, "functions": list(FUNCTIONS)}


def coverage_scope_projection(value: dict[str, Any]) -> dict[str, Any]:
    """Return the immutable requirement projection of the mutable manifest."""

    functions = value.get("functions")
    if not isinstance(functions, dict):
        raise VerificationError("coverage manifest functions are absent for scope projection")
    projected_functions: dict[str, Any] = {}
    for name in FUNCTIONS:
        entry = functions.get(name)
        if not isinstance(entry, dict) or not isinstance(entry.get("requirements"), list):
            raise VerificationError(f"coverage function {name} has no requirement projection")
        projected_functions[name] = entry["requirements"]
    cross_cutting = value.get("cross_cutting")
    if not isinstance(cross_cutting, list):
        raise VerificationError("coverage cross-cutting requirements are absent for scope projection")
    return {
        "schema": COVERAGE_SCOPE_SCHEMA,
        "normative": value.get("normative"),
        "functions": projected_functions,
        "cross_cutting": cross_cutting,
    }


def validate_coverage_scope(
    observed: Any,
    expected: dict[str, Any],
    label: str = "coverage scope",
) -> None:
    """Reject missing fields or any change to the frozen requirement scope."""

    if not isinstance(observed, dict) or set(observed) != COVERAGE_SCOPE_KEYS:
        raise VerificationError(f"{label} has an unexpected schema or fields")
    equal(f"{label} schema", observed.get("schema"), COVERAGE_SCOPE_SCHEMA)
    equal(f"{label} projection", observed, expected)


def verify_coverage_scope(value: dict[str, Any], path: Path = COVERAGE_SCOPE) -> dict[str, Any]:
    """Verify the retained immutable scope before accepting mutable evidence."""

    if not path.is_file():
        raise PendingReceipt(f"immutable coverage scope is absent: {path.name}")
    observed = read_json(path)
    expected = coverage_scope_projection(value)
    validate_coverage_scope(observed, expected)
    return {
        "sha256": digest(path),
        "schema": COVERAGE_SCOPE_SCHEMA,
        "functions": len(expected["functions"]),
        "cross_cutting": len(expected["cross_cutting"]),
    }


def verify_coverage() -> dict[str, Any]:
    """Require hashed, case-bound evidence for every requirement.

    ``coverage-requirements.json`` starts as a planning input.  At freeze it
    must use :data:`COVERAGE_SCHEMA`: each function has one PASS evidence
    object whose bindings cover the exact requirement strings, and each
    cross-cutting requirement has the same shape.  Every binding hashes its
    source and an evidence-rooted receipt; every identifier must occur in that
    retained receipt.  The manifest therefore cannot be satisfied by prose,
    a non-empty dictionary, or an unbound test name.
    """

    if not COVERAGE_REQUIREMENTS.is_file():
        raise PendingReceipt("coverage-requirements.json is absent")
    value = read_json(COVERAGE_REQUIREMENTS)
    if not isinstance(value, dict):
        raise VerificationError("coverage-requirements.json is not an object")
    allowed_top_level = {
        "schema",
        "status",
        "normative",
        "functions",
        "cross_cutting",
        "cross_cutting_evidence",
    }
    unknown_top_level = set(value) - allowed_top_level
    if unknown_top_level:
        raise VerificationError(f"coverage manifest has unknown fields: {sorted(unknown_top_level)}")
    if not isinstance(value.get("normative"), str) or not value["normative"].strip():
        raise VerificationError("coverage normative source is absent")
    functions = value.get("functions")
    if not isinstance(functions, dict) or set(functions) != set(FUNCTIONS):
        raise VerificationError("coverage requirements do not cover exactly the nine functions")
    pending: list[str] = []
    requirements_by_function: dict[str, list[str]] = {}
    for name in FUNCTIONS:
        entry = functions[name]
        if not isinstance(entry, dict) or set(entry) != {"requirements", "evidence"}:
            raise VerificationError(f"coverage requirement for {name} is malformed")
        requirements = entry.get("requirements")
        if (
            not isinstance(requirements, list)
            or not requirements
            or not all(isinstance(item, str) and item.strip() for item in requirements)
        ):
            raise VerificationError(f"coverage requirements for {name} are malformed")
        if len(requirements) != len(set(requirements)):
            raise VerificationError(f"coverage requirements for {name} contain duplicates")
        requirements_by_function[name] = requirements
        evidence = entry.get("evidence")
        if evidence is None:
            pending.append(f"{name}: evidence is null")
        elif not isinstance(evidence, dict):
            raise VerificationError(f"coverage evidence for {name} is malformed")
        else:
            if set(evidence) != EVIDENCE_KEYS:
                raise VerificationError(
                    f"coverage evidence for {name} must contain exactly "
                    "schema/status/contract_sha256/bindings"
                )
            evidence_status = evidence.get("status")
            if evidence_status in {"pending", "PENDING", "HOLD", "hold", "incomplete"}:
                equal(f"coverage {name} pending schema", evidence.get("schema"), COVERAGE_SCHEMA)
                validate_digest(evidence.get("contract_sha256"), f"coverage {name} contract hash")
                if not isinstance(evidence.get("bindings"), list):
                    raise VerificationError(f"coverage evidence for {name} bindings are malformed")
                pending.append(f"{name}: evidence status is {evidence_status!r}")
            elif evidence_status != "PASS":
                raise VerificationError(f"coverage evidence for {name} has status {evidence_status!r}")
    cross_cutting = value.get("cross_cutting")
    if not isinstance(cross_cutting, list) or not cross_cutting or not all(
        isinstance(item, str) and item.strip() for item in cross_cutting
    ):
        raise VerificationError("coverage cross_cutting requirements are malformed")
    if len(cross_cutting) != len(set(cross_cutting)):
        raise VerificationError("coverage cross_cutting requirements contain duplicates")
    scope_receipt = verify_coverage_scope(value)

    # Validate the shape of any cross-cutting entries even while the manifest
    # is pending.  A pending bundle may leave an entry's evidence null, but it
    # may not hide an arbitrary dictionary behind the pending status.
    cross_evidence = value.get("cross_cutting_evidence")
    if cross_evidence is not None:
        if not isinstance(cross_evidence, list):
            raise VerificationError("coverage cross_cutting_evidence is malformed")
        for index, item in enumerate(cross_evidence):
            if not isinstance(item, dict) or set(item) != CROSS_EVIDENCE_KEYS:
                raise VerificationError(f"cross-cutting evidence {index} is malformed")
            requirement = item.get("requirement")
            if requirement not in cross_cutting:
                raise VerificationError(
                    f"cross-cutting evidence names an unknown requirement: {requirement!r}"
                )
            evidence = item.get("evidence")
            if evidence is None:
                pending.append(f"cross-cutting {requirement}: evidence is null")
            elif not isinstance(evidence, dict) or set(evidence) != EVIDENCE_KEYS:
                raise VerificationError(f"cross-cutting evidence {index} payload is malformed")
            elif evidence.get("status") in {"pending", "PENDING", "HOLD", "hold", "incomplete"}:
                equal(
                    f"cross-cutting {requirement} pending schema",
                    evidence.get("schema"),
                    COVERAGE_SCHEMA,
                )
                validate_digest(
                    evidence.get("contract_sha256"),
                    f"cross-cutting {requirement} contract hash",
                )
                if not isinstance(evidence.get("bindings"), list):
                    raise VerificationError(f"cross-cutting evidence {index} bindings are malformed")
                pending.append(f"cross-cutting {requirement}: evidence is pending")
            elif evidence.get("status") != "PASS":
                raise VerificationError(
                    f"cross-cutting evidence {index} has status {evidence.get('status')!r}"
                )

    status = value.get("status")
    if status != "PASS":
        if status is None or not isinstance(status, str):
            raise VerificationError("coverage manifest status is absent or malformed")
        if any(token in status.lower() for token in ("pending", "planning", "incomplete", "hold")):
            pending.append(f"manifest status is {status!r}")
        else:
            raise VerificationError(f"coverage manifest status is not PASS: {status!r}")

    if status != "PASS":
        raise PendingReceipt("coverage evidence is incomplete: " + ", ".join(pending))

    if value.get("schema") != COVERAGE_SCHEMA:
        raise VerificationError("coverage manifest schema is absent or unexpected")
    if not CONTRACT.is_file():
        raise PendingReceipt("lookup contract.md is absent for coverage bindings")
    contract_hash = digest(CONTRACT)
    function_receipts: dict[str, Any] = {}
    for name in FUNCTIONS:
        evidence = functions[name]["evidence"]
        if evidence is None:
            raise PendingReceipt(f"{name}: evidence is null")
        function_receipts[name] = validate_evidence(
            evidence,
            requirements_by_function[name],
            f"coverage {name}",
            contract_hash,
        )

    cross_evidence = value.get("cross_cutting_evidence")
    if not isinstance(cross_evidence, list):
        raise VerificationError("coverage cross_cutting_evidence is absent or malformed")
    seen_cross: set[str] = set()
    cross_receipts: dict[str, Any] = {}
    for index, item in enumerate(cross_evidence):
        if not isinstance(item, dict) or set(item) != CROSS_EVIDENCE_KEYS:
            raise VerificationError(f"cross-cutting evidence {index} is malformed")
        requirement = item.get("requirement")
        if requirement not in cross_cutting:
            raise VerificationError(f"cross-cutting evidence names an unknown requirement: {requirement!r}")
        if requirement in seen_cross:
            raise VerificationError(f"cross-cutting evidence duplicates requirement: {requirement!r}")
        seen_cross.add(requirement)
        cross_receipts[requirement] = validate_evidence(
            item.get("evidence"),
            [requirement],
            f"cross-cutting evidence {index}",
            contract_hash,
        )
    if seen_cross != set(cross_cutting):
        raise VerificationError(
            "cross-cutting evidence does not cover every requirement: "
            f"missing={sorted(set(cross_cutting) - seen_cross)}"
        )
    return {
        "sha256": digest(COVERAGE_REQUIREMENTS),
        "scope_sha256": scope_receipt["sha256"],
        "functions": len(functions),
        "cross_cutting": len(cross_cutting),
        "function_bindings": sum(item["bindings"] for item in function_receipts.values()),
        "cross_cutting_bindings": sum(item["bindings"] for item in cross_receipts.values()),
        "schema": COVERAGE_SCHEMA,
    }


def verify_reviews(contract: dict[str, Any]) -> dict[str, Any]:
    missing = [
        str(path.relative_to(HERE))
        for path in (SEMANTIC_REVIEW, RESOURCE_REVIEW, REVIEW_RECEIPT)
        if not path.is_file()
    ]
    if missing:
        raise PendingReceipt("independent review evidence is absent: " + ", ".join(missing))
    receipt = read_json(REVIEW_RECEIPT)
    if not isinstance(receipt, dict):
        raise VerificationError("review-receipt.json is not an object")
    schema = receipt.get("schema")
    if schema is not None and schema != "ods-formula-lookups-review-receipt-v1":
        raise VerificationError(f"unexpected lookup review receipt schema: {schema!r}")
    equal("lookup review receipt status", receipt.get("status"), "PASS")
    reviews = receipt.get("reviews")
    if not isinstance(reviews, dict) or set(reviews) != {"semantic", "resource"}:
        raise VerificationError("lookup review receipt kinds are incomplete")
    for kind, report in (("semantic", SEMANTIC_REVIEW), ("resource", RESOURCE_REVIEW)):
        entry = reviews.get(kind)
        if not isinstance(entry, dict):
            raise VerificationError(f"{kind} lookup review receipt entry is malformed")
        equal(f"{kind} review status", entry.get("status"), "PASS")
        equal(f"{kind} review path", entry.get("path"), report.name)
        equal(f"{kind} review hash", entry.get("sha256"), digest(report))
    freeze = GATES / "freeze.json"
    if not freeze.is_file():
        raise PendingReceipt("lookup freeze.json is absent")
    equal("review receipt freeze hash", receipt.get("freeze_sha256"), digest(freeze))
    equal("review receipt contract hash", receipt.get("contract_sha256"), contract["sha256"])
    return {
        "status": receipt["status"],
        "semantic_sha256": digest(SEMANTIC_REVIEW),
        "resource_sha256": digest(RESOURCE_REVIEW),
        "freeze_sha256": digest(freeze),
        "contract_sha256": contract["sha256"],
    }


def discover_oracle() -> tuple[Path, Path]:
    configured_script = BASELINE.get("oracle_script")
    configured_goldens = BASELINE.get("oracle_goldens")
    scripts = (
        [safe_child(HERE, configured_script, "oracle_script")]
        if configured_script is not None
        else sorted(HERE.glob("*_oracle.py"))
    )
    goldens = (
        [safe_child(HERE, configured_goldens, "oracle_goldens")]
        if configured_goldens is not None
        else sorted(set(HERE.glob("*-goldens.json")) | set(HERE.glob("*_goldens.json")))
    )
    if len(scripts) != 1:
        if not scripts:
            raise PendingReceipt("independent lookup oracle script is absent")
        raise VerificationError(f"oracle script discovery is ambiguous: {scripts}")
    if len(goldens) != 1:
        if not goldens:
            raise PendingReceipt("independent lookup oracle goldens are absent")
        raise VerificationError(f"oracle goldens discovery is ambiguous: {goldens}")
    if not scripts[0].is_file() or not goldens[0].is_file():
        raise PendingReceipt("independent lookup oracle inputs are incomplete")
    return scripts[0], goldens[0]


def scope_from_object(value: dict[str, Any], label: str) -> list[str]:
    raw = first_value(value, ("functions", "scope", "selected_functions"))
    if not isinstance(raw, list) or not raw or not all(isinstance(item, str) for item in raw):
        raise VerificationError(f"{label} function scope is malformed")
    if len(raw) != len(set(raw)):
        raise VerificationError(f"{label} function scope contains duplicates")
    equal(f"{label} function scope", sorted(raw), sorted(FUNCTIONS))
    return raw


def observations_from_object(value: dict[str, Any], label: str) -> list[dict[str, Any]]:
    raw = first_value(value, ("observations", "rows", "cases"))
    if not isinstance(raw, list) or not raw:
        raise VerificationError(f"{label} observations are empty or malformed")
    rows: list[dict[str, Any]] = []
    for index, row in enumerate(raw):
        if not isinstance(row, dict):
            raise VerificationError(f"{label} observation {index} is not an object")
        function = first_value(row, ("function", "name"))
        if function not in FUNCTIONS:
            raise VerificationError(f"{label} observation {index} has an unknown function")
        if not any(isinstance(row.get(key), str) for key in ("formula", "case", "id")):
            raise VerificationError(f"{label} observation {index} has no case identity")
        if not any(
            key in row
            for key in (
                "expected", "result", "value", "error", "kind", "type",
                "expected_type", "expected_value", "expected_kind",
            )
        ):
            raise VerificationError(f"{label} observation {index} has no typed outcome")
        if not any(
            key in row
            for key in ("expected_reads", "reference_reads", "cell_reads", "reads")
        ):
            raise VerificationError(f"{label} observation {index} has no read expectation")
        rows.append(row)
    equal(
        f"{label} observed function coverage",
        {first_value(row, ("function", "name")) for row in rows},
        set(FUNCTIONS),
    )
    return rows


def verify_oracle(contract: dict[str, Any]) -> dict[str, Any]:
    script, goldens_path = discover_oracle()
    goldens = read_json(goldens_path)
    if not isinstance(goldens, dict):
        raise VerificationError("lookup oracle goldens are not an object")
    functions = scope_from_object(goldens, "oracle")
    rows = observations_from_object(goldens, "oracle")
    equal("oracle contract hash", first_value(goldens, ("contract_sha256", "contract_hash")), contract["sha256"])
    before = digest(goldens_path)
    receipt = command_json(["python3", str(script), "--check"], cwd=HERE)
    equal("oracle retained goldens", digest(goldens_path), before)
    equal("oracle verification flag", receipt.get("verified"), True)
    reported_functions = first_value(receipt, ("functions", "scope_count"))
    if reported_functions is None:
        raise VerificationError("oracle --check omitted function count")
    equal("oracle function count", reported_functions, len(functions))
    reported_rows = first_value(receipt, ("observations", "rows", "count"))
    if reported_rows is None:
        raise VerificationError("oracle --check omitted observation count")
    equal("oracle observation count", reported_rows, len(rows))
    reported_hash = first_value(receipt, ("oracle_sha256", "goldens_sha256", "golden_sha256"))
    if reported_hash is not None:
        equal("oracle goldens hash", reported_hash, digest(goldens_path))
    return {
        "functions": len(functions),
        "observations": len(rows),
        "contract_sha256": contract["sha256"],
        "goldens_sha256": digest(goldens_path),
        "script_sha256": digest(script),
        "script": str(script.relative_to(HERE)),
        "goldens": str(goldens_path.relative_to(HERE)),
    }


def element_text(element: ET.Element) -> str:
    return "".join(element.itertext())


def typed_native(cell: ET.Element) -> dict[str, Any]:
    raw = element_text(cell)
    if cell.attrib.get(CALCEXT + "value-type") == "error":
        return {"type": "error", "value": raw}
    kind = cell.attrib.get(VALUE_TYPE)
    if kind == "float":
        value = float(cell.attrib.get(VALUE, raw))
        if value.is_integer() and abs(value) < 2**53:
            value = int(value)
        return {"type": "number", "value": value}
    if kind == "boolean":
        return {"type": "logical", "value": cell.attrib.get(BOOLEAN_VALUE, raw).lower() == "true"}
    if kind == "string":
        return {"type": "text", "value": raw}
    if kind in {"date", "time"}:
        return {"type": kind, "value": raw}
    raise VerificationError(f"unexpected native XML value type: {kind!r}")


def formula_rows(xml_bytes: bytes, *, typed: bool) -> list[dict[str, Any]]:
    try:
        root = ET.fromstring(xml_bytes)
    except ET.ParseError as error:
        raise VerificationError(f"native XML is malformed: {error}") from error
    rows: list[dict[str, Any]] = []
    for row in root.iter(TABLE + "table-row"):
        cells = list(row.findall(TABLE + "table-cell"))
        formula_cells = [cell for cell in cells if FORMULA in cell.attrib]
        if not formula_cells:
            continue
        cell = formula_cells[0]
        item: dict[str, Any] = {"case": element_text(cells[0]) if cells else "", "formula": cell.attrib[FORMULA]}
        if typed:
            item["native"] = typed_native(cell)
        rows.append(item)
    return rows


def native_result_rows(value: Any) -> list[dict[str, Any]]:
    raw = value if isinstance(value, list) else value.get("observations", value.get("rows", value.get("results"))) if isinstance(value, dict) else None
    if not isinstance(raw, list) or not raw or not all(isinstance(row, dict) for row in raw):
        raise VerificationError("native results are empty or malformed")
    return raw


def fixture_path(fixture: dict[str, Any], keys: tuple[str, ...], default: str) -> str:
    value = first_value(fixture, keys)
    return value if isinstance(value, str) and value else default


def verify_native(oracle: dict[str, Any], contract: dict[str, Any]) -> dict[str, Any]:
    if not NATIVE.is_dir():
        raise PendingReceipt("native evidence directory is absent")
    provenance_path = NATIVE / "provenance.json"
    if not provenance_path.is_file():
        raise PendingReceipt("native provenance is absent")
    provenance = read_json(provenance_path)
    if not isinstance(provenance, dict) or not isinstance(provenance.get("fixture"), dict):
        raise VerificationError("native provenance fixture is malformed")
    status = provenance.get("status")
    if isinstance(status, str) and status.lower() in {"pending", "planning", "incomplete", "hold"}:
        raise PendingReceipt(f"native provenance status is {status!r}")
    fixture = provenance["fixture"]
    input_path = safe_child(NATIVE, fixture_path(fixture, ("input", "source", "input_fixture"), "native.fods"), "native input")
    output_path = safe_child(NATIVE, fixture_path(fixture, ("recalculated_output", "output", "recalculated"), "recalculated.ods"), "native output")
    results_path = safe_child(NATIVE, fixture_path(fixture, ("native_results", "results"), "native-results.json"), "native results")
    for path in (input_path, output_path, results_path):
        if not path.is_file():
            raise PendingReceipt(f"native receipt file is absent: {path}")
    for label, path, keys in (
        ("native input hash", input_path, ("input_sha256", "source_sha256")),
        ("native output hash", output_path, ("recalculated_output_sha256", "output_sha256")),
        ("native results hash", results_path, ("native_results_sha256", "results_sha256")),
    ):
        expected = first_value(fixture, keys)
        if not isinstance(expected, str):
            raise VerificationError(f"{label} is absent from provenance")
        equal(label, digest(path), expected)
    expected = native_result_rows(read_json(results_path))
    for index, row in enumerate(expected):
        if not isinstance(first_value(row, ("case", "id", "formula")), str):
            raise VerificationError(f"native result {index} has no case identity")
        if not any(key in row for key in ("native", "expected", "result", "value", "error", "kind", "type")):
            raise VerificationError(f"native result {index} has no typed outcome")
    result_functions = {
        first_value(row, ("function", "name"))
        for row in expected
        if isinstance(first_value(row, ("function", "name")), str)
    }
    equal("native function coverage", result_functions, set(FUNCTIONS))
    scope = provenance.get("scope")
    if isinstance(scope, dict):
        scope_functions = first_value(scope, ("functions", "selected_functions"))
        if scope_functions is not None:
            equal("native provenance function scope", sorted(scope_functions), sorted(FUNCTIONS))
        row_count = first_value(scope, ("formula_rows", "rows", "observations"))
        if row_count is not None:
            equal("native provenance row count", row_count, len(expected))
    independent = provenance.get("independent_oracle")
    if isinstance(independent, dict):
        for field, key in (("script_sha256", "script_sha256"), ("goldens_sha256", "goldens_sha256"), ("functions", "functions"), ("observations", "observations")):
            if field in independent:
                equal(f"native oracle link {field}", independent[field], oracle[key])
    normative = provenance.get("normative_source")
    if isinstance(normative, dict) and "contract_sha256" in normative:
        equal("native contract link", normative["contract_sha256"], contract["sha256"])
    if input_path.suffix.lower() in {".fods", ".ods"} and output_path.suffix.lower() == ".ods":
        try:
            with zipfile.ZipFile(output_path) as archive:
                content = archive.read("content.xml")
            source_rows = formula_rows(input_path.read_bytes(), typed=False)
            observed_rows = formula_rows(content, typed=True)
        except (OSError, KeyError, zipfile.BadZipFile) as error:
            raise VerificationError(f"native fixture archive is invalid: {error}") from error
        content_hash = first_value(fixture, ("recalculated_content_xml_sha256", "content_xml_sha256"))
        if not isinstance(content_hash, str):
            raise VerificationError("native content.xml hash is absent from provenance")
        equal("native content.xml hash", digest_bytes(content), content_hash)
        equal("native source/output row count", len(observed_rows), len(source_rows))
        equal("native result row count", len(expected), len(source_rows))
        if all(row.get("case") is not None and row.get("formula") is not None for row in expected):
            equal("native input formula sequence", [(row["case"], row["formula"]) for row in source_rows], [(row["case"], row["formula"]) for row in expected])
        if all("native" in row for row in expected):
            equal("native output formula sequence", observed_rows, [{"case": row.get("case"), "formula": row.get("native_formula", row.get("formula")), "native": row["native"]} for row in expected])
    reproduce = NATIVE / "reproduce.py"
    if not reproduce.is_file():
        raise PendingReceipt("native reproduction script is absent")
    before_results = digest(results_path)
    reproduction = command_json(["python3", str(reproduce)], cwd=NATIVE)
    equal("native results stability", digest(results_path), before_results)
    if reproduction.get("verified") is False or reproduction.get("status") not in (None, "ok", "verified"):
        raise VerificationError(f"native reproduction was not successful: {reproduction}")
    reported_rows = first_value(reproduction, ("rows", "observations", "count"))
    if reported_rows is None:
        raise VerificationError("native reproduction omitted row count")
    equal("native reproduction row count", reported_rows, len(expected))
    return {"rows": len(expected), "functions": len(result_functions), "results_sha256": digest(results_path)}


def git_bytes(relative: str) -> bytes | None:
    try:
        return subprocess.check_output(["git", "show", f"{BASELINE_COMMIT}:{relative}"], cwd=REPO, stderr=subprocess.DEVNULL)
    except subprocess.CalledProcessError:
        return None


def git_digest(relative: str) -> str | None:
    content = git_bytes(relative)
    return digest_bytes(content) if content is not None else None


def include_dependencies(selected: dict[str, str]) -> set[str]:
    listing = subprocess.check_output(["git", "ls-tree", "-r", "--name-only", BASELINE_COMMIT, "crates"], cwd=REPO, text=True)
    paths = {path for path in listing.splitlines() if path.startswith("crates/litchi-ods/") and path.endswith(".rs")}
    paths.update(relative for relative in selected if relative.startswith("crates/litchi-ods/") and relative.endswith(".rs"))
    dependencies: set[str] = set()
    for relative in sorted(paths):
        data = git_bytes(relative)
        candidate = REPO / relative
        if relative in selected and candidate.is_file():
            data = candidate.read_bytes()
        if data is None:
            continue
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            continue
        for included in INCLUDE.findall(text):
            resolved = Path(os.path.normpath(str(Path(relative).parent / included)))
            if resolved.is_absolute() or ".." in resolved.parts:
                continue
            dependency = str(resolved)
            if ((REPO / dependency).is_file() if relative in selected else git_bytes(dependency) is not None):
                dependencies.add(dependency)
    return dependencies


def expected_workspace_paths(selected: dict[str, str]) -> set[str]:
    paths = {"Cargo.toml", "Cargo.lock"}
    listing = subprocess.check_output(["git", "ls-tree", "-r", "--name-only", BASELINE_COMMIT, "crates"], cwd=REPO, text=True)
    paths.update(path for path in listing.splitlines() if path.endswith((".rs", "Cargo.toml", "build.rs")))
    paths.update(include_dependencies(selected))
    return paths


def verify_source_closure() -> dict[str, Any]:
    paths = [GATES / name for name in ("freeze.json", "staged-profile-sources.json", "source-before.json", "source-after.json")]
    if any(not path.is_file() for path in paths):
        raise PendingReceipt("frozen lookup source closure is not present")
    freeze, staged, before, after = (read_json(path) for path in paths)
    if not all(isinstance(value, dict) for value in (freeze, staged, before, after)):
        raise VerificationError("lookup source receipts are malformed")
    equal("source base commit", freeze.get("base_commit"), BASELINE_COMMIT)
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected or "Cargo.lock" not in selected:
        raise VerificationError("frozen selected source map is incomplete")
    scope_relative = str(COVERAGE_SCOPE.relative_to(REPO))
    if scope_relative not in selected:
        raise VerificationError("frozen selected source map omits immutable coverage-scope.json")
    for relative in selected:
        validate_repo_relative(relative, "frozen selected source")
    equal("staged source path set", set(staged), set(selected))
    equal("source manifest stability", before, after)
    equal("source-before baseline commit", before.get("git_head"), BASELINE_COMMIT)
    equal("source-after baseline commit", after.get("git_head"), BASELINE_COMMIT)
    equal("boundary tool stability", before.get("boundary_tool_sha256"), after.get("boundary_tool_sha256"))
    boundary = REPO / "tools/check_crate_boundaries.py"
    if not boundary.is_file():
        raise VerificationError("boundary checker is absent")
    equal("boundary tool hash", before.get("boundary_tool_sha256"), digest(boundary))
    source_map = before.get("source_sha256")
    if not isinstance(source_map, dict):
        raise VerificationError("source manifest selected map is malformed")
    equal("source manifest selected path set", set(source_map), set(selected))
    for relative, expected in selected.items():
        if not isinstance(expected, str) or not re.fullmatch(r"[0-9a-f]{64}", expected):
            raise VerificationError(f"selected source hash is malformed: {relative}")
        equal(f"staged selected source {relative}", staged.get(relative), expected)
        source = GATES / "Cargo.lock" if relative == "Cargo.lock" else REPO / relative
        if not source.is_file():
            raise VerificationError(f"frozen source is absent: {relative}")
        equal(f"current selected source {relative}", digest(source), expected)
        equal(f"manifest selected source {relative}", source_map.get(relative), expected)
    workspace = before.get("workspace_source_sha256")
    if not isinstance(workspace, dict):
        raise VerificationError("workspace source map is malformed")
    expected_paths = expected_workspace_paths(selected)
    expected_paths.update(relative for relative in selected if relative.startswith("crates/") and relative.endswith((".rs", "Cargo.toml", "build.rs")))
    equal("workspace source path set", set(workspace), expected_paths)
    for relative, observed in workspace.items():
        expected = selected.get(relative) if relative in selected else git_digest(relative)
        if relative == "Cargo.lock":
            expected = GATE_LOCK_SHA256
        if expected is None:
            raise VerificationError(f"workspace source lacks expected hash: {relative}")
        equal(f"workspace source {relative}", observed, expected)
    changed = sorted(relative for relative, expected in selected.items() if git_digest(relative) != expected)
    return {"base_commit": BASELINE_COMMIT, "selected_files": len(selected), "changed_paths": changed}


def expected_commands(batch: list[str]) -> dict[str, list[str]]:
    return {
        "ods-tests": ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods"],
        "clippy": ["cargo", "clippy", "--locked", "--offline", "-p", "litchi-ods", "--all-targets", "--", "-D", "warnings"],
        "rustdoc": ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--no-deps"],
        "format": ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"],
        "batch-format": ["rustfmt", "--edition", "2024", "--check", "--config", "skip_children=true", *batch],
        "boundaries": ["python3", "tools/check_crate_boundaries.py"],
        "diff-check": ["git", "diff", "--check"],
    }


def verify_gates() -> dict[str, Any]:
    required = ("freeze.json", "staged-profile-sources.json", "environment.json", "source-before.json", "source-after.json", "batch-files.json", "results.json", "verification.json", "ods-tests.log")
    missing = [name for name in required if not (GATES / name).is_file()]
    if missing:
        raise PendingReceipt("lookup gate receipts are absent: " + ", ".join(missing))
    receipt = command_json(["python3", str(GATES / "verify.py")], cwd=REPO)
    equal("gate verification flag", receipt.get("verified"), True)
    equal("gate verification status", receipt.get("status"), "ok")
    environment = read_json(GATES / "environment.json")
    if not isinstance(environment, dict):
        raise VerificationError("gate environment receipt is malformed")
    equal("gate environment head", environment.get("head"), BASELINE_COMMIT)
    equal("gate environment RUSTFLAGS", environment.get("RUSTFLAGS"), None)
    equal("gate environment RUSTDOCFLAGS", environment.get("RUSTDOCFLAGS"), "-D warnings")
    equal("gate environment lock hash", environment.get("gate_lock_sha256"), GATE_LOCK_SHA256)
    equal("gate lock file hash", digest(GATES / "Cargo.lock"), GATE_LOCK_SHA256)
    batch = read_json(GATES / "batch-files.json")
    if not isinstance(batch, list) or not all(isinstance(path, str) for path in batch) or len(batch) != len(set(batch)):
        raise VerificationError("batch-files.json is malformed")
    results = read_json(GATES / "results.json")
    if not isinstance(results, list) or not results:
        raise VerificationError("gate results are empty or malformed")
    commands = expected_commands(batch)
    names = [row.get("name") for row in results if isinstance(row, dict)]
    equal("gate command name set", set(names), set(commands))
    equal("gate command result count", len(results), len(commands))
    for row in results:
        if not isinstance(row, dict) or row.get("exit_code") != 0:
            raise VerificationError(f"a retained gate did not pass: {row!r}")
        name = row.get("name")
        if name not in commands:
            raise VerificationError(f"unknown gate result: {name!r}")
        equal(f"{name} command", row.get("command"), commands[name])
        log = GATES / f"{name}.log"
        if not log.is_file():
            raise VerificationError(f"missing gate log for {name}")
        equal(f"{name} log hash", digest(log), row.get("log_sha256"))
    text = (GATES / "ods-tests.log").read_text(encoding="utf-8")
    summaries = SUMMARY.findall(text)
    if not summaries or any(status != "ok" or int(failed) != 0 for status, _, failed, _ in summaries):
        raise VerificationError("ods-tests.log has no wholly passing Cargo test summaries")
    totals = {"passed": sum(int(passed) for _, passed, _, _ in summaries), "failed": sum(int(failed) for _, _, failed, _ in summaries), "ignored": sum(int(ignored) for _, _, _, ignored in summaries)}
    verification = read_json(GATES / "verification.json")
    equal("stable source receipt", verification.get("stable_sources"), True)
    equal("required gate receipt", verification.get("all_required_checks_passed"), True)
    return {"commands": len(results), "totals": totals, "receipt": receipt}


def verify_performance() -> dict[str, Any]:
    if not PERFORMANCE.is_dir():
        raise PendingReceipt("lookup performance evidence directory is absent")
    results = PERFORMANCE / "results"
    summary_path = results / "capture-summary.json"
    if not summary_path.is_file():
        raise PendingReceipt("lookup performance capture summary is absent")
    before_path = results / "profile-inputs-before.json"
    after_path = results / "profile-inputs-after.json"
    if not before_path.is_file() or not after_path.is_file():
        raise PendingReceipt("lookup performance profile-input receipts are absent")
    before = read_json(before_path)
    after = read_json(after_path)
    if not isinstance(before, dict) or not before:
        raise VerificationError("lookup performance profile input map is malformed")
    equal("performance profile inputs before/after", after, before)
    for relative, expected in before.items():
        source = safe_child(PERFORMANCE, relative, "performance profile input")
        if not source.is_file():
            raise VerificationError(f"performance profile input is absent: {relative}")
        equal(f"performance profile input {relative}", digest(source), expected)
    summary = read_json(summary_path)
    if not isinstance(summary, dict) or not isinstance(summary.get("captures"), list) or not summary["captures"]:
        raise VerificationError("lookup performance capture summary is malformed")
    labels: set[str] = set()
    for capture in summary["captures"]:
        if not isinstance(capture, dict) or not isinstance(capture.get("label"), str) or capture["label"] in labels:
            raise VerificationError("performance capture labels are malformed or duplicated")
        labels.add(capture["label"])
        if not isinstance(capture.get("records"), int) or capture["records"] <= 0:
            raise VerificationError(f"performance capture {capture['label']} has no records")
    retained_path = results / "retained-files.json"
    if not retained_path.is_file():
        raise PendingReceipt("lookup performance retained-files manifest is absent")
    retained = read_json(retained_path)
    if not isinstance(retained, dict) or not retained:
        raise VerificationError("performance retained-files manifest is malformed")
    actual = {str(path.relative_to(results)) for path in results.rglob("*") if path.is_file() and path != retained_path}
    equal("performance retained file path set", set(retained), actual)
    for relative, expected in retained.items():
        path = safe_child(results, relative, "performance retained file")
        if not path.is_file():
            raise VerificationError(f"performance retained file is absent: {relative}")
        equal(f"performance retained file {relative}", digest(path), expected)
    audit = HERE / "root_performance_audit.py"
    if not audit.is_file():
        raise PendingReceipt("lookup retained performance audit is absent")
    receipt = command_json(["python3", str(audit)], cwd=REPO)
    equal("performance audit status", receipt.get("status"), "ok")
    return receipt


def run_check(name: str, function: Callable[[], dict[str, Any]], checks: dict[str, Any], pending: list[str]) -> None:
    try:
        checks[name] = function()
    except PendingReceipt as error:
        pending.append(f"{name}: {error}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--allow-pending", action="store_true", help="report missing receipts as pending; never verified=true")
    args = parser.parse_args()
    checks: dict[str, Any] = {}
    pending: list[str] = []
    run_check("locks", verify_locks, checks, pending)
    run_check("contract", verify_contract, checks, pending)
    run_check("coverage", verify_coverage, checks, pending)
    contract = checks.get("contract")
    if isinstance(contract, dict):
        run_check("reviews", lambda: verify_reviews(contract), checks, pending)
        run_check("oracle", lambda: verify_oracle(contract), checks, pending)
    else:
        pending.append("reviews: waiting for contract identity")
        pending.append("oracle: waiting for contract identity")
    oracle = checks.get("oracle")
    if isinstance(oracle, dict) and isinstance(contract, dict):
        run_check("native", lambda: verify_native(oracle, contract), checks, pending)
    else:
        pending.append("native: waiting for independent oracle identity")
    run_check("source_closure", verify_source_closure, checks, pending)
    run_check("gates", verify_gates, checks, pending)
    run_check("performance", verify_performance, checks, pending)
    if pending and not args.allow_pending:
        raise VerificationError("; ".join(pending))
    if pending:
        result = {"status": "pending", "verified": False, "pending": pending, "checks": checks}
    else:
        result = {"status": "ok", "verified": True, "checks": checks}
    print(json.dumps(result, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, VerificationError, subprocess.CalledProcessError) as error:
        raise SystemExit(f"lookup evidence verification failed: {error}")
