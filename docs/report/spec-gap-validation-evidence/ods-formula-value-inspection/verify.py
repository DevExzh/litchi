#!/usr/bin/env python3
"""Independently verify the ODS value-inspection evidence bundle.

The verifier is usable before the source freeze with ``--allow-pending``.  In
that mode it still checks the independent oracle, retained native XML rows, and
lock identities, then reports missing/stale gate or performance receipts as
pending.  A normal invocation is fail-closed: it requires the seven gate logs,
the frozen source closure, and (when a performance results tree is present) the
raw performance verifier receipt.

Counts are read from the retained data.  This file deliberately does not bake
in a provisional function, oracle-row, focused-test, performance-case, or
changed-file total.
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


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
GATES = HERE / "gates"
NATIVE = HERE / "native"
PERFORMANCE = HERE / "performance"
BASELINE = json.loads((HERE / "baseline.json").read_text(encoding="utf-8"))
BASELINE_COMMIT = BASELINE["commit"]
FUNCTIONS = tuple(BASELINE["scope"])
GATE_LOCK_SHA256 = BASELINE["gate_lock_sha256"]
AMBIENT_LOCK_SHA256 = BASELINE["ambient_lock_sha256"]

TABLE = "{urn:oasis:names:tc:opendocument:xmlns:table:1.0}"
OFFICE = "{urn:oasis:names:tc:opendocument:xmlns:office:1.0}"
CALCEXT = "{urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0}"
FORMULA = TABLE + "formula"
VALUE_TYPE = OFFICE + "value-type"
VALUE = OFFICE + "value"
BOOLEAN_VALUE = OFFICE + "boolean-value"


class VerificationError(RuntimeError):
    pass


class PendingReceipt(RuntimeError):
    pass


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def load(path: Path):
    if not path.is_file():
        raise VerificationError(f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise VerificationError(f"invalid JSON evidence: {path}: {error}") from error


def equal(label: str, observed, expected) -> None:
    if observed != expected:
        raise VerificationError(f"{label}: expected {expected!r}, observed {observed!r}")


def command_json(command: list[str], *, cwd: Path) -> dict:
    completed = subprocess.run(command, cwd=cwd, capture_output=True, text=True)
    if completed.returncode:
        raise VerificationError(
            f"command failed ({completed.returncode}): {' '.join(command)}\n"
            f"{completed.stdout}{completed.stderr}"
        )
    lines = [line for line in completed.stdout.splitlines() if line.strip()]
    if not lines:
        raise VerificationError(f"command emitted no JSON: {' '.join(command)}")
    try:
        return json.loads(lines[-1])
    except json.JSONDecodeError as error:
        raise VerificationError(
            f"command did not emit JSON: {' '.join(command)}\n{completed.stdout}"
        ) from error


def verify_locks() -> dict[str, str]:
    gate = GATES / "Cargo.lock"
    ambient = REPO / "Cargo.lock"
    if not gate.is_file() or not ambient.is_file():
        raise VerificationError("one or both Cargo.lock inputs are absent")
    equal("retained gate lock", digest(gate), GATE_LOCK_SHA256)
    equal("ambient workspace lock", digest(ambient), AMBIENT_LOCK_SHA256)
    return {"gate": GATE_LOCK_SHA256, "ambient": AMBIENT_LOCK_SHA256}


def verify_oracle() -> dict[str, object]:
    contract = HERE / "contract.md"
    script = HERE / "inspection_oracle.py"
    goldens_path = HERE / "inspection-goldens.json"
    contract_sha = digest(contract)
    goldens = load(goldens_path)
    functions = goldens.get("functions")
    rows = goldens.get("observations")
    if not isinstance(functions, list) or not functions:
        raise VerificationError("inspection oracle function list is empty or malformed")
    if sorted(functions) != sorted(FUNCTIONS):
        raise VerificationError(
            f"inspection oracle function scope differs: {functions!r} != {list(FUNCTIONS)!r}"
        )
    if not isinstance(rows, list) or not rows:
        raise VerificationError("inspection oracle observations are empty or malformed")
    for index, row in enumerate(rows):
        if not isinstance(row, dict) or row.get("function") not in functions:
            raise VerificationError(f"oracle row {index} has an unknown function")
        if not isinstance(row.get("formula"), str) or "expected" not in row:
            raise VerificationError(f"oracle row {index} is incomplete")
    equal("oracle contract hash", goldens.get("contract_sha256"), contract_sha)
    receipt = command_json(["python3", str(script), "--check"], cwd=HERE)
    equal("oracle verification flag", receipt.get("verified"), True)
    equal("oracle function count", receipt.get("functions"), len(functions))
    equal("oracle observation count", receipt.get("observations"), len(rows))
    equal("oracle goldens hash", receipt.get("oracle_sha256"), digest(goldens_path))
    return {
        "functions": len(functions),
        "observations": len(rows),
        "contract_sha256": contract_sha,
        "goldens_sha256": digest(goldens_path),
        "script_sha256": digest(script),
    }


def element_text(element: ET.Element) -> str:
    return "".join(element.itertext())


def typed_native(cell: ET.Element) -> dict[str, object]:
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
    raise VerificationError(f"unexpected native XML value type: {kind!r}")


def formula_rows(xml_bytes: bytes, *, typed: bool, require_case: bool = True) -> list[dict[str, object]]:
    root = ET.fromstring(xml_bytes)
    rows: list[dict[str, object]] = []
    for row in root.iter(TABLE + "table-row"):
        cells = list(row.findall(TABLE + "table-cell"))
        formula_cells = [cell for cell in cells if FORMULA in cell.attrib]
        if not formula_cells:
            continue
        case = element_text(cells[0]) if cells else ""
        if require_case and not case:
            continue
        if not case:
            continue
        formula_cell = formula_cells[0]
        item: dict[str, object] = {
            "case": case,
            "formula": formula_cell.attrib[FORMULA],
        }
        if typed:
            item["native"] = typed_native(formula_cell)
        rows.append(item)
    return rows


def verify_native(oracle: dict[str, object], *, allow_stale_link: bool) -> dict[str, object]:
    provenance = load(NATIVE / "provenance.json")
    expected = load(NATIVE / "native-results.json")
    if not isinstance(expected, list) or not expected:
        raise VerificationError("native-results.json is empty or malformed")
    fixture = provenance.get("fixture")
    if not isinstance(fixture, dict):
        raise VerificationError("native provenance fixture is malformed")
    input_path = NATIVE / str(fixture.get("input", ""))
    output_path = NATIVE / str(fixture.get("recalculated_output", ""))
    results_path = NATIVE / str(fixture.get("native_results", "native-results.json"))
    for path in (input_path, output_path, results_path):
        if not path.is_file():
            raise VerificationError(f"native receipt file is absent: {path}")
    equal("native input hash", digest(input_path), fixture.get("input_sha256"))
    equal("native output hash", digest(output_path), fixture.get("recalculated_output_sha256"))
    equal("native results hash", digest(results_path), fixture.get("native_results_sha256"))
    with zipfile.ZipFile(output_path) as archive:
        content = archive.read("content.xml")
    equal(
        "native content.xml hash",
        digest_bytes(content),
        fixture.get("recalculated_content_xml_sha256"),
    )

    source_rows = formula_rows(input_path.read_bytes(), typed=False)
    observed_rows = formula_rows(content, typed=True)
    equal("native source row count", len(source_rows), len(expected))
    equal("native output row count", len(observed_rows), len(expected))
    equal(
        "native input formula sequence",
        [(row["case"], row["formula"]) for row in source_rows],
        [(row.get("case"), row.get("formula")) for row in expected],
    )
    projection = [
        {
            "case": row["case"],
            "formula": row.get("native_formula", row["formula"]),
            "native": row["native"],
        }
        for row in expected
    ]
    equal("native output formula sequence", observed_rows, projection)

    functions = sorted({row.get("function") for row in expected})
    equal("native function scope", functions, sorted(FUNCTIONS))
    comparisons = Counter(row.get("comparison") for row in expected)
    scope = provenance.get("scope")
    if not isinstance(scope, dict):
        raise VerificationError("native provenance scope is malformed")
    equal("native provenance functions", sorted(scope.get("functions", [])), sorted(FUNCTIONS))
    equal("native provenance row count", scope.get("formula_rows"), len(expected))
    raw_count = scope.get("raw_inspection_rows")
    conversion_count = scope.get("conversion_rows")
    if not isinstance(raw_count, int) or not isinstance(conversion_count, int):
        raise VerificationError("native provenance row partitions are malformed")
    equal("native provenance row partition", raw_count + conversion_count, len(expected))
    equal("native comparison counts", scope.get("comparison_counts"), dict(comparisons))

    links: list[str] = []
    normative = provenance.get("normative_source", {})
    if normative.get("contract_sha256") != oracle["contract_sha256"]:
        links.append("native provenance contract hash is stale")
    independent = provenance.get("independent_oracle", {})
    if independent.get("script_sha256") != oracle["script_sha256"]:
        links.append("native provenance oracle script hash is stale")
    if independent.get("goldens_sha256") != oracle["goldens_sha256"]:
        links.append("native provenance oracle goldens hash is stale")
    if independent.get("functions") != oracle["functions"]:
        links.append("native provenance oracle function count is stale")
    if independent.get("observations") != oracle["observations"]:
        links.append("native provenance oracle observation count is stale")
    if links and not allow_stale_link:
        raise VerificationError("; ".join(links))

    reproduction = command_json(["python3", str(NATIVE / "reproduce.py")], cwd=NATIVE)
    equal("native reproduction rows", reproduction.get("rows"), len(expected))
    equal("native reproduction content hash", reproduction.get("content_xml_sha256"), fixture.get("recalculated_content_xml_sha256"))
    equal("native reproduction comparison counts", reproduction.get("comparison_counts"), dict(comparisons))
    return {
        "rows": len(expected),
        "functions": len(functions),
        "comparison_counts": dict(comparisons),
        "stale_links": links,
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


INCLUDE = re.compile(r'include_(?:bytes|str)!\(\s*"([^"\n]+)"')


def include_dependencies(selected: dict[str, str]) -> set[str]:
    """Derive literal include files from baseline plus selected source text."""

    source_paths: set[str] = set()
    listing = subprocess.check_output(
        ["git", "ls-tree", "-r", "--name-only", BASELINE_COMMIT, "crates"],
        cwd=REPO,
        text=True,
    )
    # Keep this in lockstep with gates/run.py: literal dependencies are found
    # in the baseline litchi-ods package and selected candidate litchi-ods
    # source text, while the workspace map itself covers every crate source.
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


def verify_source_closure() -> dict[str, object]:
    freeze_path = GATES / "freeze.json"
    staged_path = GATES / "staged-profile-sources.json"
    before_path = GATES / "source-before.json"
    after_path = GATES / "source-after.json"
    required = [freeze_path, staged_path, before_path, after_path]
    if any(not path.is_file() for path in required):
        raise PendingReceipt("frozen source closure is not present")
    freeze = load(freeze_path)
    staged = load(staged_path)
    before = load(before_path)
    after = load(after_path)
    if not isinstance(freeze.get("selected_files"), dict) or not freeze["selected_files"]:
        raise VerificationError("freeze selected source map is empty")
    selected = freeze["selected_files"]
    equal("source map path set", set(staged), set(selected))
    equal("source manifest stability", before, after)
    equal("boundary tool stability", before.get("boundary_tool_sha256"), after.get("boundary_tool_sha256"))
    equal(
        "boundary tool hash",
        before.get("boundary_tool_sha256"),
        digest(REPO / "tools/check_crate_boundaries.py"),
    )
    source_map = before.get("source_sha256")
    if not isinstance(source_map, dict):
        raise VerificationError("source manifest selected map is malformed")
    for relative, expected in selected.items():
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
    changed = sorted(
        relative for relative, expected in selected.items()
        if git_digest(relative) != expected
    )
    return {
        "base_commit": BASELINE_COMMIT,
        "selected_files": len(selected),
        "changed_paths": changed,
    }


def verify_gates() -> dict[str, object]:
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
    return receipt


def verify_performance(candidate_root: Path | None, candidate_freeze: Path | None) -> dict[str, object]:
    results = PERFORMANCE / "results"
    if not results.is_dir() or not (results / "capture-summary.json").is_file():
        raise PendingReceipt("performance captures are absent")
    if (candidate_root is None) != (candidate_freeze is None):
        raise VerificationError(
            "supply both --candidate-root and --candidate-freeze, or neither for retained-evidence verification"
        )
    matrix = load(PERFORMANCE / "case-matrix.json")
    functions = matrix.get("scope", matrix.get("functions"))
    if not isinstance(functions, list) or sorted(functions) != sorted(FUNCTIONS):
        raise VerificationError("performance case matrix function scope differs from baseline")
    if candidate_root is None:
        # The owned checkout is deliberately removed after acceptance. The
        # retained audit checks baseline source bytes against Git and both
        # candidate closures against the verified gate manifests, then checks
        # every raw receipt. It does not require an ambient lock substitution
        # or a surviving temporary worktree.
        receipt = command_json(["python3", str(HERE / "root_performance_audit.py")], cwd=REPO)
        equal("retained performance verification status", receipt.get("status"), "ok")
        return receipt
    receipt = command_json(
        [
            "python3",
            str(PERFORMANCE / "verify.py"),
            "--candidate-root",
            str(candidate_root),
            "--candidate-freeze",
            str(candidate_freeze),
        ],
        cwd=PERFORMANCE,
    )
    equal("performance verification status", receipt.get("status"), "ok")
    return receipt


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--allow-pending",
        action="store_true",
        help="report incomplete/stale final receipts as pending and exit successfully",
    )
    parser.add_argument("--candidate-root", type=Path)
    parser.add_argument("--candidate-freeze", type=Path)
    args = parser.parse_args()

    checks: dict[str, object] = {}
    pending: list[str] = []
    checks["locks"] = verify_locks()
    checks["oracle"] = verify_oracle()
    try:
        checks["native"] = verify_native(
            checks["oracle"], allow_stale_link=args.allow_pending
        )
        if checks["native"].get("stale_links"):
            pending.extend(checks["native"]["stale_links"])
    except PendingReceipt as error:
        pending.append(str(error))
    try:
        checks["source_closure"] = verify_source_closure()
    except PendingReceipt as error:
        pending.append(str(error))
    try:
        checks["gates"] = verify_gates()
    except PendingReceipt as error:
        pending.append(str(error))
    try:
        checks["performance"] = verify_performance(
            args.candidate_root.resolve() if args.candidate_root else None,
            args.candidate_freeze.resolve() if args.candidate_freeze else None,
        )
    except PendingReceipt as error:
        pending.append(str(error))

    if pending and not args.allow_pending:
        raise VerificationError("; ".join(pending))
    if pending:
        receipt = {
            "status": "pending",
            "verified": False,
            "pending": pending,
            "checks": checks,
        }
    else:
        receipt = {"status": "ok", "verified": True, "checks": checks}
    print(json.dumps(receipt, ensure_ascii=False, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (VerificationError, OSError, subprocess.CalledProcessError) as error:
        raise SystemExit(f"evidence verification failed: {error}")
