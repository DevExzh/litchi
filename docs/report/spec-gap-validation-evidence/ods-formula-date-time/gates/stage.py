#!/usr/bin/env python3
"""Stage the date/time source closure into an isolated preparation checkout.

The date/time batch has no frozen candidate yet.  This preparation tool copies
the current closure and the retained isolated lock, then writes manifests that
the root coordinator can review.  It deliberately never creates ``freeze.json``
and never invokes Cargo.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
from typing import Any


HERE = Path(__file__).resolve().parent
EVIDENCE = HERE.parent
ROOT = HERE.parents[4]
BASELINE_PATH = EVIDENCE / "baseline.json"
SCOPE_PATH = EVIDENCE / "coverage-scope.json"

STAGE_SCHEMA = "ods-formula-date-time-stage-v1"
BASELINE_SCHEMA = "ods-formula-date-time-baseline-v1"
LOCK_RELATIVE = "Cargo.lock"
WORKSPACE_MANIFESTS = ("Cargo.toml", "crates/litchi-ods/Cargo.toml")
FEATURE_MATRIX_FILES = ("docs/FEATURE_MATRIX.md", "crates/litchi-ods/docs/FEATURE_MATRIX.md")
EVALUATION_ROOT = Path("crates/litchi-ods/src/codec/formula/evaluation")
FORMULA_ROOT = Path("crates/litchi-ods/src/codec/formula")
DATE_TIME_TEST_GLOB = "crates/litchi-ods/tests/ods_formula_date_time*.rs"
IMMUTABLE_EVIDENCE_FILES = ("baseline.json", "contract.md", "coverage-scope.json", "oracle-vectors.json")
ORACLE_SOURCE_FILES = ("oracle-plan.md", "oracle-vectors.json", "oracle_verify.py")
NATIVE_DIRECTORY = "native"
PERFORMANCE_FILES = (
    "performance/performance-plan.md",
    "performance/case-matrix.json",
    "performance/run_profile.py",
)
PERFORMANCE_HARNESS_DIRECTORY = "performance/harness"
PERFORMANCE_HARNESS_FILES = (
    "performance/harness/Cargo.toml",
    "performance/harness/Cargo.lock",
)


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"invalid JSON input {path}: {error}") from error


def safe_relative(value: object, label: str) -> str:
    if not isinstance(value, str) or not value or Path(value).is_absolute() or ".." in Path(value).parts:
        raise RuntimeError(f"{label} is not safely repository-relative: {value!r}")
    return value


def safe_checkout_path(checkout: Path, relative: str) -> Path:
    safe_relative(relative, "selected path")
    path = (checkout / relative).resolve()
    try:
        path.relative_to(checkout.resolve())
    except ValueError as error:
        raise RuntimeError(f"selected path escapes isolated checkout: {relative}") from error
    return path


def baseline() -> dict[str, Any]:
    value = read_json(BASELINE_PATH)
    if not isinstance(value, dict) or value.get("schema") != BASELINE_SCHEMA:
        raise RuntimeError("date/time baseline has an unexpected schema")
    for key in ("production_commit", "preparation_commit", "selected_baseline_sources", "isolated_lock"):
        if key not in value:
            raise RuntimeError(f"date/time baseline is missing {key}")
    if not isinstance(value["production_commit"], str) or not value["production_commit"]:
        raise RuntimeError("date/time production_commit is malformed")
    if not isinstance(value["preparation_commit"], str) or not value["preparation_commit"]:
        raise RuntimeError("date/time preparation_commit is malformed")
    selected = value["selected_baseline_sources"]
    if not isinstance(selected, dict) or not selected:
        raise RuntimeError("date/time selected_baseline_sources is empty")
    for relative, expected in selected.items():
        safe_relative(relative, "selected baseline source")
        if not isinstance(expected, str) or len(expected) != 64:
            raise RuntimeError(f"selected baseline hash is malformed: {relative}")
    lock = value["isolated_lock"]
    if not isinstance(lock, dict) or not isinstance(lock.get("path"), str) or not isinstance(lock.get("sha256"), str):
        raise RuntimeError("date/time isolated_lock is malformed")
    safe_relative(lock["path"], "isolated lock path")
    if len(lock["sha256"]) != 64:
        raise RuntimeError("date/time isolated lock hash is malformed")
    return value


def scope_functions() -> tuple[str, ...]:
    scope = read_json(SCOPE_PATH)
    if not isinstance(scope, dict) or scope.get("schema") != "ods-formula-date-time-coverage-scope-v1":
        raise RuntimeError("coverage-scope.json has an unexpected schema")
    functions = scope.get("functions")
    if not isinstance(functions, dict) or not functions:
        raise RuntimeError("coverage-scope.json has no function scope")
    if any(not isinstance(name, str) or not name for name in functions):
        raise RuntimeError("coverage-scope.json has a malformed function name")
    return tuple(functions)


def retained_lock(config: dict[str, Any]) -> tuple[str, Path]:
    lock = config["isolated_lock"]
    relative = safe_relative(lock["path"], "isolated lock path")
    path = (ROOT / relative).resolve()
    try:
        path.relative_to(ROOT.resolve())
    except ValueError as error:
        raise RuntimeError("isolated lock path escapes the repository") from error
    if not path.is_file():
        raise RuntimeError(f"retained isolated lock is absent: {relative}")
    expected = lock["sha256"]
    observed = digest(path)
    if observed != expected:
        raise RuntimeError(f"retained isolated lock hash mismatch: expected {expected}, observed {observed}")
    return relative, path


def relative_files(root: Path, pattern: str) -> list[str]:
    return sorted(str(path.relative_to(ROOT)) for path in root.glob(pattern) if path.is_file())


def evaluation_source_paths(repo: Path = ROOT) -> list[str]:
    root = repo / EVALUATION_ROOT
    if not root.is_dir():
        raise RuntimeError(f"formula evaluation source root is absent: {root}")
    return sorted(str(path.relative_to(repo)) for path in root.rglob("*.rs") if path.is_file())


def formula_dependency_paths(repo: Path = ROOT) -> list[str]:
    """Return expression/reference Rust dependencies used by formula evaluation."""

    root = repo / FORMULA_ROOT
    if not root.is_dir():
        raise RuntimeError(f"formula source root is absent: {root}")
    selected: set[str] = set()
    for path in root.glob("*.rs"):
        if path.is_file():
            selected.add(str(path.relative_to(repo)))
    for name in ("expression", "reference"):
        child = root / name
        if child.is_dir():
            selected.update(str(path.relative_to(repo)) for path in child.rglob("*.rs") if path.is_file())
    return sorted(selected)


def evidence_input_paths(repo: Path = ROOT) -> list[str]:
    """Return only immutable date/time inputs, never review or receipt files."""

    root = repo / EVIDENCE.relative_to(ROOT)
    selected: set[str] = set()

    def add(path: Path, *, required: bool = False) -> None:
        if not path.is_file():
            if required:
                raise RuntimeError(f"date/time input is absent: {path}")
            return
        selected.add(str(path.relative_to(repo)))

    for relative in IMMUTABLE_EVIDENCE_FILES:
        add(root / relative, required=True)
    for relative in ORACLE_SOURCE_FILES:
        add(root / relative, required=True)

    native = root / NATIVE_DIRECTORY
    if not native.is_dir():
        raise RuntimeError(f"date/time native input directory is absent: {native}")
    for path in native.rglob("*"):
        add(path)

    for relative in PERFORMANCE_FILES:
        add(root / relative, required=True)
    for relative in PERFORMANCE_HARNESS_FILES:
        add(root / relative, required=True)
    harness_source = root / PERFORMANCE_HARNESS_DIRECTORY / "src"
    if not harness_source.is_dir():
        raise RuntimeError(f"date/time performance harness source directory is absent: {harness_source}")
    source_paths = list(harness_source.rglob("*.rs"))
    if not source_paths:
        raise RuntimeError(f"date/time performance harness source is empty: {harness_source}")
    for path in source_paths:
        add(path, required=True)
    return sorted(selected)


def source_closure(config: dict[str, Any]) -> tuple[set[str], dict[str, list[str]], list[str]]:
    """Return selected paths, category membership, and preparation warnings.

    Existing baseline paths seed the closure.  The complete formula evaluator
    subtree is included because shared dispatch, cache, reference, and helper
    edits can change a date/time result even when their filenames are generic.
    Date-time tests and authored evidence are discovered from the date/time
    evidence directory rather than copied from another batch.
    """

    selected = {safe_relative(relative, "selected baseline source") for relative in config["selected_baseline_sources"]}
    selected.add(LOCK_RELATIVE)
    categories: dict[str, list[str]] = {
        "baseline_selected": sorted(selected),
        "formula_evaluation_recursive": [],
        "expression_reference_dependencies": [],
        "tests": [],
        "contract_oracle": [],
        "feature_matrix": [],
        "workspace_manifests": [],
        "isolated_lock": [LOCK_RELATIVE],
    }

    for relative in WORKSPACE_MANIFESTS:
        if (ROOT / relative).is_file():
            selected.add(relative)
            categories["workspace_manifests"].append(relative)
        else:
            raise RuntimeError(f"workspace manifest is absent: {relative}")

    evaluation_paths = evaluation_source_paths(ROOT)
    selected.update(evaluation_paths)
    categories["formula_evaluation_recursive"].extend(evaluation_paths)

    dependency_paths = formula_dependency_paths(ROOT)
    selected.update(dependency_paths)
    categories["expression_reference_dependencies"].extend(dependency_paths)

    test_paths = relative_files(ROOT, DATE_TIME_TEST_GLOB)
    selected.update(test_paths)
    categories["tests"].extend(test_paths)

    authored_inputs = evidence_input_paths(ROOT)
    selected.update(authored_inputs)
    categories["contract_oracle"].extend(authored_inputs)

    for relative in FEATURE_MATRIX_FILES:
        path = ROOT / relative
        if not path.is_file():
            raise RuntimeError(f"feature matrix is absent: {path}")
        selected.add(relative)
        categories["feature_matrix"].append(relative)

    warnings: list[str] = []
    if not categories["tests"]:
        warnings.append(f"no files currently match {DATE_TIME_TEST_GLOB}; candidate tests are pending")
    if not any("date_time" in path.lower() or "datetime" in path.lower() for path in categories["formula_evaluation_recursive"]):
        warnings.append("no date/time implementation module is present yet; candidate modules are pending")

    categories = {key: sorted(set(value)) for key, value in categories.items()}
    return selected, categories, warnings


def source_for(relative: str, lock_relative: str, lock_path: Path) -> Path:
    if relative == LOCK_RELATIVE:
        return lock_path
    safe_relative(relative, "source path")
    path = ROOT / relative
    if not path.is_file():
        raise RuntimeError(f"selected source is absent: {relative}")
    return path


def batch_files(selected: set[str]) -> list[str]:
    return sorted(
        relative
        for relative in selected
        if relative.endswith(".rs") and relative.startswith("crates/litchi-ods/")
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("checkout", type=Path, help="clean checkout at baseline preparation_commit")
    args = parser.parse_args()
    checkout = args.checkout.resolve()
    if checkout == ROOT:
        raise RuntimeError("refusing to stage into the working tree")
    if not checkout.is_dir():
        raise RuntimeError(f"isolated checkout is absent: {checkout}")

    config = baseline()
    preparation_commit = config["preparation_commit"]
    observed_head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=checkout, text=True).strip()
    if observed_head != preparation_commit:
        raise RuntimeError(
            f"isolated checkout must be preparation_commit {preparation_commit}, observed {observed_head}"
        )
    if (HERE / "freeze.json").exists():
        raise RuntimeError("refusing to stage while freeze.json exists; root must own freeze promotion")

    lock_relative, lock_path = retained_lock(config)
    selected, categories, warnings = source_closure(config)
    source_map: dict[str, str] = {}
    for relative in sorted(selected):
        source = source_for(relative, lock_relative, lock_path)
        target = safe_checkout_path(checkout, relative)
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source, target)
        source_map[relative] = digest(source)
        if digest(target) != source_map[relative]:
            raise RuntimeError(f"isolated source hash mismatch: {relative}")

    batch = batch_files(selected)
    manifest = {
        "schema": STAGE_SCHEMA,
        "status": "staged; freeze pending",
        "base_commit": preparation_commit,
        "production_commit": config["production_commit"],
        "isolated_checkout_head": observed_head,
        "isolated_lock": {"path": lock_relative, "sha256": config["isolated_lock"]["sha256"]},
        "selected_files": source_map,
        "batch_files": batch,
        "closure": categories,
        "scope_functions": list(scope_functions()),
        "warnings": warnings,
    }
    (HERE / "batch-files.json").write_text(json.dumps(batch, indent=2) + "\n", encoding="utf-8")
    (HERE / "staged-profile-sources.json").write_text(
        json.dumps(source_map, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    (HERE / "stage-manifest.json").write_text(
        json.dumps(manifest, indent=2, sort_keys=True) + "\n", encoding="utf-8"
    )
    print(
        json.dumps(
            {
                "status": manifest["status"],
                "base_commit": preparation_commit,
                "selected_sources": len(source_map),
                "batch_files": len(batch),
                "warnings": warnings,
                "freeze_written": False,
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
