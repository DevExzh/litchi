#!/usr/bin/env python3
"""Fail-closed verifier for the reference-metadata evidence bundle.

The evidence directory is assembled in stages.  ``--allow-pending`` is useful
while the source, native, or performance receipts are being assembled, but it
always emits ``verified: false`` and names every pending check.  A normal run
requires every receipt and never turns an existing failing receipt into a
pending one.

Counts are taken from the retained inputs.  This verifier intentionally does
not contain a copied test total, oracle total, native row total, or benchmark
sample total.
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
SEMANTIC_REVIEW = HERE / "semantic-review.md"
RESOURCE_REVIEW = HERE / "resource-review.md"
REVIEW_RECEIPT = HERE / "review-receipt.json"


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
if (
    not isinstance(BASELINE_COMMIT, str)
    or not BASELINE_COMMIT
    or not isinstance(FUNCTIONS, list)
    or not FUNCTIONS
    or len(FUNCTIONS) != len(set(FUNCTIONS))
    or not all(isinstance(name, str) and name for name in FUNCTIONS)
    or not isinstance(GATE_LOCK_SHA256, str)
    or not isinstance(AMBIENT_LOCK_SHA256, str)
):
    raise VerificationError("baseline.json has an incomplete commit, scope, or lock identity")
FUNCTIONS = tuple(FUNCTIONS)

SUMMARY = re.compile(r"test result: (ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored")
INCLUDE = re.compile(r'include_(?:bytes|str)!\(\s*"([^"\n]+)"')

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
    candidates = [output]
    candidates.extend(line.strip() for line in reversed(output.splitlines()))
    for candidate in candidates:
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


def safe_child(root: Path, relative: str, label: str) -> Path:
    if not isinstance(relative, str) or not relative or Path(relative).is_absolute():
        raise VerificationError(f"{label} is not a relative path: {relative!r}")
    path = (root / relative).resolve()
    try:
        path.relative_to(root.resolve())
    except ValueError as error:
        raise VerificationError(f"{label} escapes evidence root: {relative!r}") from error
    return path


def validate_repo_relative(relative: Any, label: str) -> str:
    if not isinstance(relative, str) or not relative or Path(relative).is_absolute():
        raise VerificationError(f"{label} is not a repository-relative path: {relative!r}")
    if ".." in Path(relative).parts:
        raise VerificationError(f"{label} escapes the repository: {relative!r}")
    return relative


def verify_locks() -> dict[str, str]:
    gate = GATES / "Cargo.lock"
    ambient = REPO / "Cargo.lock"
    if not gate.is_file() or not ambient.is_file():
        raise PendingReceipt("gate or ambient Cargo.lock is absent")
    equal("retained gate lock", digest(gate), GATE_LOCK_SHA256)
    equal("ambient workspace lock", digest(ambient), AMBIENT_LOCK_SHA256)
    return {"gate": GATE_LOCK_SHA256, "ambient": AMBIENT_LOCK_SHA256}


def verify_reviews() -> dict[str, Any]:
    """Require explicit root acceptance of both independent review reports."""

    missing = [
        str(path.relative_to(HERE))
        for path in (SEMANTIC_REVIEW, RESOURCE_REVIEW, REVIEW_RECEIPT)
        if not path.is_file()
    ]
    if missing:
        raise PendingReceipt("review evidence is absent: " + ", ".join(missing))
    receipt = read_json(REVIEW_RECEIPT)
    if not isinstance(receipt, dict):
        raise VerificationError("review-receipt.json is not an object")
    if "schema" in receipt:
        equal("review receipt schema", receipt.get("schema"), "ods-formula-reference-metadata-review-receipt-v1")
    equal("review receipt status", receipt.get("status"), "PASS")
    reviews = receipt.get("reviews")
    if not isinstance(reviews, dict):
        raise VerificationError("review receipt reviews object is absent")
    expected_reports = {
        "semantic": SEMANTIC_REVIEW,
        "resource": RESOURCE_REVIEW,
    }
    equal("review receipt kinds", set(reviews), set(expected_reports))
    for kind, report in expected_reports.items():
        entry = reviews.get(kind)
        if not isinstance(entry, dict):
            raise VerificationError(f"{kind} review receipt entry is malformed")
        equal(f"{kind} review status", entry.get("status"), "PASS")
        equal(f"{kind} review report", entry.get("path"), report.name)
        equal(f"{kind} review hash", entry.get("sha256"), digest(report))
    freeze = GATES / "freeze.json"
    contract = HERE / "contract.md"
    if not freeze.is_file() or not contract.is_file():
        raise PendingReceipt("review receipt inputs are absent: freeze.json or contract.md")
    equal("review receipt freeze hash", receipt.get("freeze_sha256"), digest(freeze))
    equal("review receipt contract hash", receipt.get("contract_sha256"), digest(contract))
    return {
        "status": receipt["status"],
        "semantic_sha256": digest(SEMANTIC_REVIEW),
        "resource_sha256": digest(RESOURCE_REVIEW),
        "freeze_sha256": digest(freeze),
        "contract_sha256": digest(contract),
    }


def discover_oracle() -> tuple[Path, Path]:
    configured_script = BASELINE.get("oracle_script")
    configured_goldens = BASELINE.get("oracle_goldens")
    if configured_script is not None:
        scripts = [safe_child(HERE, configured_script, "oracle_script")]
    else:
        scripts = sorted(HERE.glob("*_oracle.py"))
    if configured_goldens is not None:
        goldens = [safe_child(HERE, configured_goldens, "oracle_goldens")]
    else:
        goldens = sorted(HERE.glob("*-goldens.json")) + sorted(HERE.glob("*_goldens.json"))
        # The two globs can identify the same file when a future profile uses
        # both spellings.
        goldens = sorted(set(goldens))
    if len(scripts) != 1:
        if not scripts:
            raise PendingReceipt("independent oracle script is absent")
        raise VerificationError(f"oracle script discovery is ambiguous: {scripts}")
    if len(goldens) != 1:
        if not goldens:
            raise PendingReceipt("independent oracle goldens are absent")
        raise VerificationError(f"oracle goldens discovery is ambiguous: {goldens}")
    if not scripts[0].is_file() or not goldens[0].is_file():
        raise PendingReceipt("independent oracle inputs are incomplete")
    return scripts[0], goldens[0]


def scope_from_object(value: dict[str, Any], label: str) -> list[str]:
    raw = first_value(value, ("functions", "scope", "selected_functions"))
    if not isinstance(raw, list) or not raw or not all(isinstance(item, str) for item in raw):
        raise VerificationError(f"{label} function scope is missing or malformed")
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
        if not isinstance(function, str) or function not in FUNCTIONS:
            raise VerificationError(f"{label} observation {index} has an unknown function")
        # Require an actual formula/case identity and an outcome field.  This
        # prevents a receipt containing only a list of function names from
        # masquerading as an executable semantic oracle.
        if not any(isinstance(row.get(key), str) for key in ("formula", "case", "id")):
            raise VerificationError(f"{label} observation {index} has no formula/case identity")
        if not any(
            key in row
            for key in (
                "expected", "result", "value", "error", "kind", "type",
                "expected_type", "expected_value", "expected_kind",
            )
        ):
            raise VerificationError(f"{label} observation {index} has no outcome")
        rows.append(row)
    observed_functions = {first_value(row, ("function", "name")) for row in rows}
    equal(f"{label} observed function coverage", observed_functions, set(FUNCTIONS))
    return rows


def verify_oracle() -> dict[str, Any]:
    contract = HERE / "contract.md"
    if not contract.is_file():
        raise PendingReceipt("independent contract is absent")
    script, goldens_path = discover_oracle()
    goldens = read_json(goldens_path)
    if not isinstance(goldens, dict):
        raise VerificationError("oracle goldens are not an object")
    functions = scope_from_object(goldens, "oracle")
    rows = observations_from_object(goldens, "oracle")
    contract_sha = digest(contract)
    linked_contract = first_value(goldens, ("contract_sha256", "contract_hash"))
    if linked_contract is None:
        raise VerificationError("oracle goldens do not pin contract.md")
    equal("oracle contract hash", linked_contract, contract_sha)
    before = digest(goldens_path)
    receipt = command_json(["python3", str(script), "--check"], cwd=HERE)
    equal("oracle retained goldens", digest(goldens_path), before)
    if receipt.get("verified") is not True:
        raise VerificationError("oracle --check did not report verified=true")
    reported_functions = first_value(receipt, ("functions", "scope_count"))
    if reported_functions is None:
        raise VerificationError("oracle --check omitted its function count")
    equal("oracle function count", reported_functions, len(functions))
    reported_observations = first_value(receipt, ("observations", "rows", "count"))
    if reported_observations is None:
        raise VerificationError("oracle --check omitted its observation count")
    equal("oracle observation count", reported_observations, len(rows))
    reported_hash = first_value(receipt, ("oracle_sha256", "goldens_sha256", "golden_sha256"))
    if reported_hash is None:
        raise VerificationError("oracle --check omitted its goldens hash")
    equal("oracle goldens hash", reported_hash, digest(goldens_path))
    return {
        "functions": len(functions),
        "observations": len(rows),
        "contract_sha256": contract_sha,
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
        return {
            "type": "logical",
            "value": cell.attrib.get(BOOLEAN_VALUE, raw).lower() == "true",
        }
    if kind == "string":
        return {"type": "text", "value": raw}
    if kind == "date" or kind == "time":
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
        formula_cell = formula_cells[0]
        item: dict[str, Any] = {
            "case": element_text(cells[0]) if cells else "",
            "formula": formula_cell.attrib[FORMULA],
        }
        if typed:
            item["native"] = typed_native(formula_cell)
        rows.append(item)
    return rows


def native_result_rows(value: Any) -> list[dict[str, Any]]:
    if isinstance(value, list):
        raw = value
    elif isinstance(value, dict):
        raw = first_value(value, ("observations", "rows", "results"))
    else:
        raw = None
    if not isinstance(raw, list) or not raw:
        raise VerificationError("native results are empty or malformed")
    if not all(isinstance(row, dict) for row in raw):
        raise VerificationError("native results contain a non-object row")
    return raw


def fixture_path(fixture: dict[str, Any], keys: tuple[str, ...], default: str) -> str:
    value = first_value(fixture, keys)
    return value if isinstance(value, str) and value else default


def verify_native(oracle: dict[str, Any], *, allow_stale_link: bool) -> dict[str, Any]:
    if not NATIVE.is_dir():
        raise PendingReceipt("native evidence directory is absent")
    provenance_path = NATIVE / "provenance.json"
    if not provenance_path.is_file():
        raise PendingReceipt("native provenance is absent")
    provenance = read_json(provenance_path)
    if not isinstance(provenance, dict):
        raise VerificationError("native provenance is not an object")
    fixture = provenance.get("fixture")
    if not isinstance(fixture, dict):
        raise VerificationError("native provenance fixture is malformed")
    input_path = safe_child(
        NATIVE,
        fixture_path(fixture, ("input", "source", "input_fixture"), "native.fods"),
        "native input",
    )
    output_path = safe_child(
        NATIVE,
        fixture_path(fixture, ("recalculated_output", "output", "recalculated"), "recalculated.ods"),
        "native output",
    )
    results_path = safe_child(
        NATIVE,
        fixture_path(fixture, ("native_results", "results"), "native-results.json"),
        "native results",
    )
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
        formula_rows_count = first_value(scope, ("formula_rows", "rows", "observations"))
        if formula_rows_count is not None:
            equal("native provenance row count", formula_rows_count, len(expected))
    elif isinstance(provenance.get("selected_functions"), list):
        equal("native provenance function scope", sorted(provenance["selected_functions"]), sorted(FUNCTIONS))

    # The standard native receipt is an ODF fixture.  When it is available,
    # compare the retained input/output formula sequence and typed output with
    # the result manifest.  The checks are conditional only to support a
    # deliberately JSON-only native adapter; such an adapter must still expose
    # hashes, scope, and an executable reproduction receipt above.
    if input_path.suffix.lower() in {".fods", ".ods"} and output_path.suffix.lower() == ".ods":
        try:
            input_bytes = input_path.read_bytes()
            with zipfile.ZipFile(output_path) as archive:
                content = archive.read("content.xml")
        except (OSError, KeyError, zipfile.BadZipFile) as error:
            raise VerificationError(f"native fixture archive is invalid: {error}") from error
        content_hash = first_value(fixture, ("recalculated_content_xml_sha256", "content_xml_sha256"))
        if not isinstance(content_hash, str):
            raise VerificationError("native content.xml hash is absent from provenance")
        equal("native content.xml hash", digest_bytes(content), content_hash)
        source_rows = formula_rows(input_bytes, typed=False)
        observed_rows = formula_rows(content, typed=True)
        equal("native source/output row count", len(observed_rows), len(source_rows))
        equal("native result row count", len(expected), len(source_rows))
        expected_sequence = [
            (row.get("case"), row.get("formula"))
            for row in expected
        ]
        if all(row.get("case") is not None and row.get("formula") is not None for row in expected):
            equal(
                "native input formula sequence",
                [(row["case"], row["formula"]) for row in source_rows],
                expected_sequence,
            )
        if all("native" in row for row in expected):
            projection = [
                {
                    "case": row.get("case"),
                    "formula": row.get("native_formula", row.get("formula")),
                    "native": row["native"],
                }
                for row in expected
            ]
            equal("native output formula sequence", observed_rows, projection)

    links: list[str] = []
    normative = provenance.get("normative_source")
    if isinstance(normative, dict):
        linked = first_value(normative, ("contract_sha256", "contract_hash"))
        if linked is not None and linked != oracle["contract_sha256"]:
            links.append("native provenance contract hash is stale")
    independent = provenance.get("independent_oracle")
    if isinstance(independent, dict):
        for field, oracle_key, label in (
            ("script_sha256", "script_sha256", "oracle script"),
            ("goldens_sha256", "goldens_sha256", "oracle goldens"),
        ):
            linked = independent.get(field)
            if linked is not None and linked != oracle[oracle_key]:
                links.append(f"native provenance {label} hash is stale")
        for field, oracle_key, label in (
            ("functions", "functions", "oracle function count"),
            ("observations", "observations", "oracle observation count"),
        ):
            linked = independent.get(field)
            if linked is not None and linked != oracle[oracle_key]:
                links.append(f"native provenance {label} is stale")
    if links and not allow_stale_link:
        raise VerificationError("; ".join(links))

    reproduce = NATIVE / "reproduce.py"
    if not reproduce.is_file():
        raise PendingReceipt("native reproduction script is absent")
    before_results = digest(results_path)
    reproduction = command_json(["python3", str(reproduce)], cwd=NATIVE)
    equal("native results stability", digest(results_path), before_results)
    if reproduction.get("verified") is False:
        raise VerificationError("native reproduction reported verified=false")
    if reproduction.get("status") not in (None, "ok", "verified"):
        raise VerificationError(f"native reproduction status is not successful: {reproduction}")
    for field in ("rows", "observations", "count"):
        if field in reproduction:
            equal(f"native reproduction {field}", reproduction[field], len(expected))
            break
    else:
        raise VerificationError("native reproduction omitted its row/observation count")
    reported_content_hash = first_value(
        reproduction, ("content_xml_sha256", "recalculated_content_xml_sha256")
    )
    expected_content_hash = first_value(fixture, ("recalculated_content_xml_sha256", "content_xml_sha256"))
    if reported_content_hash is not None and expected_content_hash is not None:
        equal("native reproduction content hash", reported_content_hash, expected_content_hash)
    return {
        "rows": len(expected),
        "functions": len(result_functions),
        "stale_links": links,
        "results_sha256": digest(results_path),
    }


def git_bytes(relative: str) -> bytes | None:
    try:
        return subprocess.check_output(
            ["git", "show", f"{BASELINE_COMMIT}:{relative}"],
            cwd=REPO,
            stderr=subprocess.DEVNULL,
        )
    except subprocess.CalledProcessError:
        return None


def git_digest(relative: str) -> str | None:
    content = git_bytes(relative)
    return digest_bytes(content) if content is not None else None


def include_dependencies(selected: dict[str, str]) -> set[str]:
    source_paths: set[str] = set()
    listing = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", BASELINE_COMMIT, "crates"],
        cwd=REPO,
        text=True,
    )
    source_paths.update(
        path
        for path in listing.splitlines()
        if path.startswith("crates/litchi-ods/") and path.endswith(".rs")
    )
    source_paths.update(
        relative
        for relative in selected
        if relative.startswith("crates/litchi-ods/") and relative.endswith(".rs")
    )
    dependencies: set[str] = set()
    for relative in sorted(source_paths):
        data = git_bytes(relative)
        candidate = REPO / relative
        selected_text = relative in selected and candidate.is_file()
        if selected_text:
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
            exists = (
                (REPO / dependency).is_file()
                if selected_text
                else git_bytes(dependency) is not None
            )
            if exists:
                dependencies.add(dependency)
    return dependencies


def expected_workspace_paths(selected: dict[str, str]) -> set[str]:
    paths = {"Cargo.toml", "Cargo.lock"}
    listing = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", BASELINE_COMMIT, "crates"],
        cwd=REPO,
        text=True,
    )
    for relative in listing.splitlines():
        if relative.endswith((".rs", "Cargo.toml", "build.rs")):
            paths.add(relative)
    paths.update(include_dependencies(selected))
    return paths


def verify_source_closure() -> dict[str, Any]:
    required = [
        GATES / "freeze.json",
        GATES / "staged-profile-sources.json",
        GATES / "source-before.json",
        GATES / "source-after.json",
    ]
    if any(not path.is_file() for path in required):
        raise PendingReceipt("frozen source closure is not present")
    freeze = read_json(GATES / "freeze.json")
    staged = read_json(GATES / "staged-profile-sources.json")
    before = read_json(GATES / "source-before.json")
    after = read_json(GATES / "source-after.json")
    if (
        not isinstance(freeze, dict)
        or not isinstance(staged, dict)
        or not isinstance(before, dict)
        or not isinstance(after, dict)
    ):
        raise VerificationError("freeze/staged source receipts are malformed")
    equal("source base commit", freeze.get("base_commit"), BASELINE_COMMIT)
    selected = freeze.get("selected_files")
    if not isinstance(selected, dict) or not selected:
        raise VerificationError("freeze selected source map is empty")
    for relative in selected:
        validate_repo_relative(relative, "frozen selected source")
    equal("source map path set", set(staged), set(selected))
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
        if not isinstance(expected, str):
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
    expected_paths.update(
        relative
        for relative in selected
        if relative.startswith("crates/")
        and relative.endswith((".rs", "Cargo.toml", "build.rs"))
    )
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


def verify_gates() -> dict[str, Any]:
    required = [
        "freeze.json", "staged-profile-sources.json", "environment.json",
        "source-before.json", "source-after.json", "batch-files.json",
        "results.json", "verification.json", "ods-tests.log",
    ]
    missing = [name for name in required if not (GATES / name).is_file()]
    if missing:
        raise PendingReceipt("gate receipts are absent: " + ", ".join(missing))
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
    results = read_json(GATES / "results.json")
    if not isinstance(results, list) or not results:
        raise VerificationError("gate results are empty or malformed")
    batch = read_json(GATES / "batch-files.json")
    if not isinstance(batch, list) or not all(isinstance(path, str) for path in batch):
        raise VerificationError("batch-files.json is malformed")
    expected_commands = {
        "ods-tests": ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods"],
        "clippy": [
            "cargo", "clippy", "--locked", "--offline", "-p", "litchi-ods",
            "--all-targets", "--", "-D", "warnings",
        ],
        "rustdoc": [
            "cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--no-deps",
        ],
        "format": ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"],
        "batch-format": [
            "rustfmt", "--edition", "2024", "--check", "--config", "skip_children=true", *batch,
        ],
        "boundaries": ["python3", "tools/check_crate_boundaries.py"],
        "diff-check": ["git", "diff", "--check"],
    }
    names = [row.get("name") for row in results if isinstance(row, dict)]
    if (
        len(names) != len(results)
        or not all(isinstance(name, str) for name in names)
        or len(set(names)) != len(names)
    ):
        raise VerificationError("gate result names are not unique")
    equal("gate command name set", set(names), set(expected_commands))
    equal("gate command result count", len(results), len(expected_commands))
    for row in results:
        if not isinstance(row, dict) or row.get("exit_code") != 0:
            raise VerificationError(f"a retained gate did not pass: {row!r}")
        name = row.get("name")
        equal(f"{name} command", row.get("command"), expected_commands[name])
        log = GATES / f"{name}.log"
        if not isinstance(name, str) or not log.is_file():
            raise VerificationError(f"missing gate log for {name!r}")
        equal(f"{name} log hash", digest(log), row.get("log_sha256"))
    log_text = (GATES / "ods-tests.log").read_text(encoding="utf-8")
    summaries = SUMMARY.findall(log_text)
    if not summaries or any(status != "ok" or int(failed) != 0 for status, _, failed, _ in summaries):
        raise VerificationError("ods-tests.log has no wholly passing Cargo test summaries")
    totals = {
        "passed": sum(int(passed) for _, passed, _, _ in summaries),
        "failed": sum(int(failed) for _, _, failed, _ in summaries),
        "ignored": sum(int(ignored) for _, _, _, ignored in summaries),
    }
    verification = read_json(GATES / "verification.json")
    equal("stable source receipt", verification.get("stable_sources"), True)
    equal("required gate receipt", verification.get("all_required_checks_passed"), True)
    return {"commands": len(results), "totals": totals, "receipt": receipt}


def verify_performance(candidate_root: Path | None, candidate_freeze: Path | None) -> dict[str, Any]:
    if not PERFORMANCE.is_dir():
        raise PendingReceipt("performance evidence directory is absent")
    results = PERFORMANCE / "results"
    summary_path = results / "capture-summary.json"
    if not summary_path.is_file():
        raise PendingReceipt("performance capture summary is absent")
    profile_before_path = results / "profile-inputs-before.json"
    profile_after_path = results / "profile-inputs-after.json"
    if not profile_before_path.is_file() or not profile_after_path.is_file():
        raise PendingReceipt("performance profile input receipts are absent")
    profile_before = read_json(profile_before_path)
    profile_after = read_json(profile_after_path)
    equal("performance profile inputs before/after", profile_before, profile_after)
    if not isinstance(profile_before, dict) or not profile_before:
        raise VerificationError("performance profile input map is empty or malformed")
    for relative, expected in profile_before.items():
        if not isinstance(relative, str) or not isinstance(expected, str):
            raise VerificationError("performance profile input map is malformed")
        source = safe_child(PERFORMANCE, relative, "performance profile input")
        if not source.is_file():
            raise VerificationError(f"performance profile input is absent: {relative}")
        equal(f"performance profile input {relative}", digest(source), expected)
    summary = read_json(summary_path)
    if not isinstance(summary, dict):
        raise VerificationError("performance capture summary is malformed")
    captures = summary.get("captures")
    if not isinstance(captures, list) or not captures:
        raise VerificationError("performance capture summary has no captures")
    labels: list[str] = []
    for capture in captures:
        if not isinstance(capture, dict):
            raise VerificationError("performance capture summary contains a non-object capture")
        label = capture.get("label")
        if not isinstance(label, str) or not label or label in labels:
            raise VerificationError("performance capture labels are missing or duplicated")
        labels.append(label)
        records = capture.get("records")
        cases = capture.get("cases")
        phases = capture.get("phases")
        if not isinstance(records, int) or records <= 0:
            raise VerificationError(f"performance capture {label} has no positive record count")
        if not isinstance(cases, list) or not cases or not isinstance(phases, list) or not phases:
            raise VerificationError(f"performance capture {label} has incomplete case/phase scope")
    for field in ("profile_input_sha256_before", "profile_input_sha256_after"):
        if field in summary:
            equal(f"performance {field}", summary[field], profile_before)
    cleanup = summary.get("cleanup")
    if isinstance(cleanup, dict) and "baseline_removed" in cleanup:
        equal("performance baseline cleanup", cleanup["baseline_removed"], True)
    retained_path = results / "retained-files.json"
    if not retained_path.is_file():
        raise PendingReceipt("performance retained-files manifest is absent")
    retained = read_json(retained_path)
    if not isinstance(retained, dict) or not retained:
        raise VerificationError("performance retained-files manifest is empty or malformed")
    actual_paths: set[str] = set()
    for path in results.rglob("*"):
        if not path.is_file() or path == retained_path:
            continue
        relative_parts = path.relative_to(results).parts
        if "__pycache__" in relative_parts or path.suffix == ".pyc":
            raise VerificationError(f"temporary Python bytecode remains in performance results: {path}")
        actual_paths.add(str(path.relative_to(results)))
    equal("performance retained file path set", set(retained), actual_paths)
    for relative, expected in retained.items():
        if not isinstance(expected, str):
            raise VerificationError(f"performance retained hash is malformed: {relative}")
        path = safe_child(results, relative, "performance retained file")
        if not path.is_file():
            raise VerificationError(f"performance retained file is absent: {relative}")
        equal(f"performance retained file {relative}", digest(path), expected)

    if (candidate_root is None) != (candidate_freeze is None):
        raise VerificationError(
            "supply both --candidate-root and --candidate-freeze, or neither for retained-evidence verification"
        )
    audit = HERE / "root_performance_audit.py"
    performance_verify = PERFORMANCE / "verify.py"
    if candidate_root is None:
        if audit.is_file():
            receipt = command_json(["python3", str(audit)], cwd=REPO)
        elif performance_verify.is_file():
            raise PendingReceipt("retained performance audit is absent")
        else:
            raise VerificationError("performance verifier is absent")
    else:
        if not performance_verify.is_file():
            raise VerificationError("performance verifier is absent")
        candidate_root = candidate_root.resolve()
        candidate_freeze = candidate_freeze.resolve() if candidate_freeze is not None else None
        if not candidate_root.is_dir() or candidate_freeze is None or not candidate_freeze.is_file():
            raise PendingReceipt("candidate performance checkout or freeze is absent")
        receipt = command_json(
            [
                "python3", str(performance_verify),
                "--candidate-root", str(candidate_root),
                "--candidate-freeze", str(candidate_freeze),
            ],
            cwd=PERFORMANCE,
        )
    equal("performance verification status", receipt.get("status"), "ok")
    if receipt.get("verified") is False:
        raise VerificationError("performance verifier reported verified=false")
    return receipt


def run_check(
    name: str,
    function: Callable[[], dict[str, Any]],
    checks: dict[str, Any],
    pending: list[str],
) -> None:
    try:
        checks[name] = function()
    except PendingReceipt as error:
        pending.append(f"{name}: {error}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--allow-pending",
        action="store_true",
        help="emit a pending receipt for not-yet-produced evidence; never reports verified=true",
    )
    parser.add_argument("--candidate-root", type=Path)
    parser.add_argument("--candidate-freeze", type=Path)
    args = parser.parse_args()
    checks: dict[str, Any] = {}
    pending: list[str] = []
    run_check("locks", verify_locks, checks, pending)
    run_check("reviews", verify_reviews, checks, pending)
    run_check("oracle", verify_oracle, checks, pending)
    oracle_for_native = checks.get("oracle")
    if isinstance(oracle_for_native, dict):
        run_check(
            "native",
            lambda: verify_native(oracle_for_native, allow_stale_link=args.allow_pending),
            checks,
            pending,
        )
    else:
        pending.append("native: waiting for independent oracle identity")
    run_check("source_closure", verify_source_closure, checks, pending)
    run_check("gates", verify_gates, checks, pending)
    run_check(
        "performance",
        lambda: verify_performance(
            args.candidate_root.resolve() if args.candidate_root else None,
            args.candidate_freeze.resolve() if args.candidate_freeze else None,
        ),
        checks,
        pending,
    )
    native = checks.get("native")
    if isinstance(native, dict) and native.get("stale_links"):
        pending.extend(f"native: {link}" for link in native["stale_links"])
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
        raise SystemExit(f"evidence verification failed: {error}")
