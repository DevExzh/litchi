#!/usr/bin/env python3
"""Make the final, fail-closed disposition for the 0552 XLSX campaign.

This module is deliberately a decision consumer.  The frozen drivers, the two
analyzers, profile verifier, quality verifier, and source custody records are
the evidence authorities.  It does not measure anything, run a build, run a
test, or synthesize a missing result.  A decision is written only after every
required evidence lane is complete and the final checkout is bound to the
resulting disposition.

The main metrics analyzer owns the main workflow gates.  The guard/cap
analyzer owns the planning-refusal and cap gates.  The exact commit profile is
conditional on the complete main plus guard/cap pilot.  Every
candidate-side adverse row and every over-five-percent same-build drift row is
matched to one retained review row.  The review is an input to adoption; it
cannot hide a row or convert an incomplete bundle into a pass.

Use ``--schema`` to inspect the required artifact contract without reading or
writing campaign evidence.  An ordinary invocation reads all evidence and
writes ``decision.json`` with exclusive-create/identical-replay semantics;
missing evidence exits without creating a decision.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Any


sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
VERIFY_PATH = HERE / "verify.py"
DECISION_PATH = HERE / "decision.json"
PROFILE_DECISION_PATH = HERE / "profile-decision.json"
SEAL_PATH = HERE / "SHA256SUMS"
PUBLIC_PATH = HERE / "public-test-sources"
PUBLIC_EXACT_PATH = HERE / "public-exact-test-sources"
CANDIDATE_ATTEMPTS_PATH = HERE / "candidate-attempts"
CHECK_ATTEMPTS_PATH = HERE / "check-attempts"

SCHEMA = "xlsx_0552_decision_v1"
METRICS_SCHEMA = "xlsx_multisource_edit_metrics_0552_v1"
GUARD_SCHEMA = "litchi.xlsx.guard-cap-analysis.v1"
EXACT_SCHEMA = "xlsx_0552_public_exact_custody_v1"
REVIEW_SCHEMA = "xlsx_0552_adverse_review_v1"
PROFILE_DECISION_SCHEMA = "xlsx_0552_profile_decision_v1"
AMENDMENT_SCHEMA = "xlsx_0552_analyzer_serialization_amendment_v1"
ORIGINAL_GUARD_ANALYZER_SHA256 = (
    "c30dbe68e0d0db4ff15ab214f75eeda0a4e5d3ca042f646c6e46d79f5d7bf817"
)
SCOPE = (
    "Matched compact source-cell proof experiment for source-backed XLSX "
    "MultiSourceEdit; OLE2/OOXML first, ODF deferred, iWork excluded"
)

PARENT_TEST = "crates/litchi-xlsx/tests/source_backed_cell_values.rs"
COMPACT_CHILD = (
    "crates/litchi-xlsx/tests/source_backed_cell_values/compact_source_proof.rs"
)
EXACT_CHILD = (
    "crates/litchi-xlsx/tests/source_backed_cell_values/public_exact_output.rs"
)
SHEET_DATA = (
    "crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/write/sheet_data.rs"
)

STAGES = ("baseline", "candidate")
SHAPES = ("medium", "dense-sparse", "noncompact", "vendor-extension")

# The five retained public-exact attempts are a correction history.  The
# first attempt is the root receipt; each later input record names the receipt
# that motivated its correction.
EXACT_ATTEMPTS = ("01", "02", "03", "04", "05")
EXACT_CHILD_HASHES = {
    "01": "18c3e9578a2017cf2d223506aee3c8a548937442efa01ad54b40d7d2d4d5319a",
    "02": "352e820d60544365d91a9925d3133f3b169ef000a79b29dc8e54d6754fc7c5b2",
    "03": "f2696dfc26bc4c90d68396355c6cc2e48c2fd36a1a2bb31301d16d910931f15f",
    "04": "20b2917530a2795b1ee08f488eefdeedd658d2d83435ac80c5cd003c183ea012",
    "05": "2cb65491b9dbaf110fe24912670da829e02d90c829e8e80910426b511b9aea2d",
}

EXACT_INITIAL_COMMAND = [
    "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
    "--all-features", "--test", "source_backed_cell_values",
    "public_multi_edit_matches_baseline_whole_worksheet_output", "--",
    "--test-threads=2",
]
EXACT_FINAL_COMMAND = [
    "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
    "--all-features", "--test", "source_backed_cell_values",
    "public_exact_output", "--", "--test-threads=2",
]
EXACT_EXIT_CODES = {"01": 101, "02": 101, "03": 0, "04": 101, "05": 0}

PRELIGHT_NAMES = (
    "baseline-public-exact-05",
    "candidate-proof-tests-03",
    "candidate-source-integration-02",
    "candidate-clippy-preflight-02",
    "candidate-no-default-preflight-01",
    "candidate-fmt-preflight-01",
    "candidate-boundaries-preflight-01",
)
PRELIGHT_COMMANDS = {
    "candidate-proof-tests-03": [
        "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
        "--all-features", "--lib", "source_proof", "--", "--test-threads=2",
    ],
    "candidate-source-integration-02": [
        "cargo", "test", "--release", "--locked", "-p", "litchi-xlsx",
        "--all-features", "--test", "source_backed_cell_values", "--",
        "--test-threads=2",
    ],
    "candidate-clippy-preflight-02": [
        "cargo", "clippy", "--locked", "-p", "litchi-xlsx",
        "--all-features", "--lib", "--", "-D", "warnings",
    ],
    "candidate-no-default-preflight-01": [
        "cargo", "check", "--locked", "-p", "litchi-xlsx",
        "--no-default-features",
    ],
    "candidate-fmt-preflight-01": ["cargo", "fmt", "--all", "--check"],
    "candidate-boundaries-preflight-01": [
        "python3", "-B", "tools/check_crate_boundaries.py",
    ],
}

SHA256_RE = re.compile(r"^[0-9a-f]{64}$")


class DecisionError(ValueError):
    """Malformed, contradictory, or incomplete evidence."""


class IncompleteDecision(DecisionError):
    """An evidence lane needed for a decision has not arrived."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise DecisionError(message)


def need(path: Path, label: str, *, directory: bool = False) -> Path:
    if not path.exists() or path.is_symlink():
        raise IncompleteDecision(f"{label} is missing or is a symlink")
    if directory:
        require(path.is_dir(), f"{label} is not a directory")
    else:
        require(path.is_file(), f"{label} is not a regular file")
    return path


def read_json(path: Path, label: str) -> Any:
    need(path, label)
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise DecisionError(f"cannot read {label}: {error}") from error


def digest(path: Path, label: str | None = None) -> str:
    label = label or path.as_posix()
    need(path, label)
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                value.update(block)
    except OSError as error:
        raise DecisionError(f"cannot hash {label}: {error}") from error
    return value.hexdigest()


def hash_value(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{label} is not a lowercase SHA-256 digest")
    return value


def nonempty(value: Any, label: str) -> str:
    require(isinstance(value, str) and value.strip(), f"{label} is empty")
    return value


def verify_module() -> Any:
    """Load verify.py without creating bytecode or running its CLI."""

    sys.dont_write_bytecode = True
    spec = importlib.util.spec_from_file_location("xlsx_0552_verify_for_decision", VERIFY_PATH)
    require(spec is not None and spec.loader is not None,
            "cannot load the 0552 verifier")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def relative_to_here(path: Path) -> str:
    try:
        return path.resolve().relative_to(HERE.resolve()).as_posix()
    except ValueError as error:
        raise DecisionError(f"path escapes the 0552 evidence bundle: {path}") from error


def report_document(verify: Any, result: dict[str, Any], label: str,
                    expected_schema: str) -> tuple[dict[str, Any], Path, str]:
    report = result.get("report")
    require(isinstance(report, dict), f"{label} report reference is missing")
    report_name = report.get("path")
    require(isinstance(report_name, str) and report_name,
            f"{label} report path is missing")
    path = HERE / report_name
    document = read_json(path, f"{label} canonical report")
    require(isinstance(document, dict) and document.get("schema") == expected_schema,
            f"{label} canonical report schema differs")
    observed = digest(path, f"{label} canonical report")
    require(report.get("sha256") == observed,
            f"{label} canonical report digest differs")
    require(document.get("status") == "pass",
            f"{label} canonical report is not a completed pass")
    return document, path, observed


def gate_rows_pass(value: Any, label: str) -> None:
    """Require every reported guard row, without aggregate substitution."""

    require(isinstance(value, dict), f"{label} gate object is missing")
    gate = value.get("gates")
    if isinstance(gate, list):
        groups = [("gates", gate)]
    else:
        require(isinstance(gate, dict), f"{label} gate rows are missing")
        groups = list(gate.items())
    for name, rows in groups:
        require(isinstance(rows, list) and rows,
                f"{label}.{name} has no explicit checks")
        for index, row in enumerate(rows):
            require(isinstance(row, dict) and isinstance(row.get("passed"), bool),
                    f"{label}.{name}[{index}] has no explicit passed flag")


def main_gate_rows(metrics: dict[str, Any]) -> dict[str, Any]:
    gates = metrics.get("main_gates")
    require(isinstance(gates, dict), "main metrics gates are missing")
    expected = {
        "primary_one_percent", "one_cell_latency", "workflow_memory",
        "allocation", "correctness_identity", "all_frozen_main_gates_pass",
        "external_controls_required",
    }
    require(set(gates) == expected, "main metrics gate inventory differs")
    for name in ("primary_one_percent", "one_cell_latency", "workflow_memory",
                 "allocation"):
        group = gates[name]
        require(isinstance(group, dict)
                and isinstance(group.get("pass"), bool)
                and isinstance(group.get("checks"), list)
                and group["checks"],
                f"main metrics {name} checks are missing")
        for index, row in enumerate(group["checks"]):
            require(isinstance(row, dict) and isinstance(row.get("pass"), bool),
                    f"main metrics {name}.checks[{index}] is malformed")
    correctness = gates["correctness_identity"]
    require(isinstance(correctness, dict)
            and isinstance(correctness.get("pass"), bool)
            and isinstance(correctness.get("identity_equal"), bool)
            and isinstance(correctness.get("exact_checks"), list)
            and correctness.get("exact_checks"),
            "main correctness identity checks are missing")
    for index, row in enumerate(correctness["exact_checks"]):
        require(isinstance(row, dict) and isinstance(row.get("equal"), bool),
                f"main correctness exact check {index} is malformed")
    require(isinstance(gates["all_frozen_main_gates_pass"], bool),
            "main aggregate gate is not boolean")
    require(gates["external_controls_required"] == {
        "status": "pending", "validated_here": False,
        "required": ["guard", "cap", "quality", "profile"],
        "reason": "main metrics analyzer does not own guard, cap, quality, or profile evidence",
    }, "main external-control gate differs")
    return gates


def extract(value: Any, path: str) -> Any:
    current = value
    for part in path.split("."):
        require(isinstance(current, dict) and part in current,
                f"canonical report field is missing: {path}")
        current = current[part]
    return current


def exact_public_manifest(verify: Any, terminal_child_sha: str) -> dict[str, str]:
    baseline, _ = verify.stage_manifest("baseline")
    old_hashes = read_json(PUBLIC_PATH / "source-hashes.json",
                           "public-test-sources/source-hashes.json")
    require(isinstance(old_hashes, dict)
            and set(old_hashes) == {PARENT_TEST, COMPACT_CHILD},
            "public compact-proof source inventory differs")
    result = dict(baseline)
    require(isinstance(old_hashes[PARENT_TEST], dict)
            and isinstance(old_hashes[COMPACT_CHILD], dict),
            "public compact-proof hash rows are malformed")
    result[PARENT_TEST] = hash_value(
        old_hashes[PARENT_TEST].get("test_sha256"),
        "public compact-proof parent test hash",
    )
    result[COMPACT_CHILD] = hash_value(
        old_hashes[COMPACT_CHILD].get("test_sha256"),
        "public compact-proof child test hash",
    )
    result[PARENT_TEST] = hash_value(
        extract(read_json(PUBLIC_EXACT_PATH / "inputs.json",
                          "public-exact-test-sources/inputs.json"),
                "source_hashes." + PARENT_TEST),
        "public exact parent test hash",
    )
    result[EXACT_CHILD] = hash_value(terminal_child_sha,
                                     "public exact child test hash")
    return dict(sorted(result.items()))


def expected_check_manifest(verify: Any, stage_manifest: dict[str, str]) -> dict[str, str]:
    result = dict(stage_manifest)
    result["Cargo.lock"] = verify.WORKSPACE_LOCK_SHA256
    return dict(sorted(result.items()))


def exact_attempt_manifest(verify: Any, child_sha: str) -> dict[str, str]:
    return expected_check_manifest(verify, exact_public_manifest(verify, child_sha))


def check_named_receipt(verify: Any, name: str, expected_manifest: dict[str, str],
                        expected_command: list[str], expected_exit: int) -> dict[str, Any]:
    path = CHECK_ATTEMPTS_PATH / name
    row = verify.validate_check_attempt(path, expected_manifest)
    receipt = read_json(path / "receipt.json", f"check-attempts/{name}/receipt.json")
    require(receipt.get("command") == expected_command,
            f"{name} command differs")
    require(receipt.get("exit_code") == expected_exit
            and receipt.get("source_stable") is True,
            f"{name} result is not the expected stable result")
    return {
        "name": name,
        "path": f"check-attempts/{name}",
        "receipt_sha256": digest(path / "receipt.json",
                                  f"check-attempts/{name}/receipt.json"),
        "source_manifest_sha256": receipt["source_manifest_sha256"],
        "exit_code": receipt["exit_code"],
        "command": receipt["command"],
    }


def validate_public_exact(verify: Any) -> dict[str, Any]:
    """Validate the retained exact-output correction chain and terminal receipt."""

    root = need(PUBLIC_EXACT_PATH, "public-exact-test-sources", directory=True)
    inputs = read_json(root / "inputs.json", "public-exact-test-sources/inputs.json")
    require(isinstance(inputs, dict)
            and set(inputs) == {"frozen_utc", "scope", "source_hashes",
                                "parent_before_sha256", "patch_sha256"},
            "public exact inputs inventory differs")
    require(inputs["scope"] == (
        "Explicit public whole-worksheet oracle; validate restored baseline first, then candidate."
    ), "public exact scope differs")
    verify.parse_time(inputs["frozen_utc"], "public exact frozen timestamp")
    source_hashes = inputs["source_hashes"]
    require(isinstance(source_hashes, dict)
            and set(source_hashes) == {PARENT_TEST, EXACT_CHILD},
            "public exact root source inventory differs")
    for name, value in source_hashes.items():
        hash_value(value, f"public exact root source hash {name}")
    parent_source = root / "sources" / Path(PARENT_TEST)
    child_source = root / "public_exact_output.rs"
    require(digest(parent_source, "public exact root parent snapshot") == source_hashes[PARENT_TEST],
            "public exact root parent snapshot differs")
    require(digest(child_source, "public exact root child snapshot") == source_hashes[EXACT_CHILD],
            "public exact root child snapshot differs")
    root_patch = root / "public-tests.patch"
    require(digest(root_patch, "public exact root patch") == inputs["patch_sha256"],
            "public exact root patch digest differs")
    hash_value(inputs["parent_before_sha256"], "public exact parent-before hash")
    baseline_manifest = verify.public_augmented_manifest()
    require(baseline_manifest.get(PARENT_TEST) == inputs["parent_before_sha256"],
            "public exact root parent-before hash is not the validated first public-test stage")

    attempt_rows: list[dict[str, Any]] = []
    previous = source_hashes[EXACT_CHILD]
    prior_receipt_names = {"02": "01", "03": "02", "04": "03", "05": "04"}
    for number in EXACT_ATTEMPTS[1:]:
        attempt = root / f"attempt-{number}"
        attempt_inputs = read_json(attempt / "inputs.json",
                                   f"public exact attempt-{number}/inputs.json")
        receipt_field = ("failed_check_receipt_sha256" if number in ("02", "03")
                         else "prior_check_receipt_sha256")
        require(isinstance(attempt_inputs, dict)
                and set(attempt_inputs) == {
                    "frozen_utc", "scope", "path", "before_sha256",
                    "after_sha256", "patch_sha256", receipt_field,
                }, f"public exact attempt-{number} inputs inventory differs")
        verify.parse_time(attempt_inputs["frozen_utc"],
                          f"public exact attempt-{number} timestamp")
        require(attempt_inputs["path"] == EXACT_CHILD,
                f"public exact attempt-{number} path differs")
        require(attempt_inputs["before_sha256"] == previous,
                f"public exact attempt-{number} chain predecessor differs")
        after = hash_value(attempt_inputs["after_sha256"],
                           f"public exact attempt-{number} after hash")
        require(after == EXACT_CHILD_HASHES[number],
                f"public exact attempt-{number} after hash differs")
        snapshot = attempt / "sources" / Path(EXACT_CHILD)
        require(digest(snapshot, f"public exact attempt-{number} snapshot") == after,
                f"public exact attempt-{number} snapshot differs")
        patch = attempt / "public-tests.patch"
        require(digest(patch, f"public exact attempt-{number} patch") ==
                hash_value(attempt_inputs["patch_sha256"],
                           f"public exact attempt-{number} patch hash"),
                f"public exact attempt-{number} patch digest differs")
        prior_name = f"baseline-public-exact-{prior_receipt_names[number]}"
        prior_receipt = CHECK_ATTEMPTS_PATH / prior_name / "receipt.json"
        require(digest(prior_receipt, f"{prior_name} receipt") ==
                hash_value(attempt_inputs[receipt_field],
                           f"public exact attempt-{number} receipt reference"),
                f"public exact attempt-{number} receipt reference differs")
        previous = after

    for number in EXACT_ATTEMPTS:
        name = f"baseline-public-exact-{number}"
        command = EXACT_INITIAL_COMMAND if number in ("01", "02", "03") else EXACT_FINAL_COMMAND
        manifest = exact_attempt_manifest(verify, EXACT_CHILD_HASHES[number])
        row = check_named_receipt(verify, name, manifest, command,
                                  EXACT_EXIT_CODES[number])
        attempt_rows.append(row)
    require(attempt_rows[-1]["exit_code"] == 0,
            "public exact terminal receipt did not pass")
    return {
        "status": "pass",
        "schema": EXACT_SCHEMA,
        "root_inputs_sha256": digest(root / "inputs.json"),
        "root_patch_sha256": digest(root_patch),
        "root_parent_sha256": source_hashes[PARENT_TEST],
        "root_child_sha256": source_hashes[EXACT_CHILD],
        "terminal_child_sha256": previous,
        "attempts": attempt_rows,
    }


def validate_candidate_custody(verify: Any, exact: dict[str, Any]) -> dict[str, Any]:
    """Bind draft-07, the measurement source, and named correctness receipts."""

    candidate_manifest, candidate_manifest_sha = verify.stage_manifest("candidate")
    binding_path = HERE / "candidate-source-binding.json"
    binding = read_json(binding_path, "candidate-source-binding.json")
    require(isinstance(binding, dict) and set(binding) == {
        "schema", "created_utc", "attempt", "inputs_sha256",
        "source_manifest_sha256", "production_source_hashes",
        "public_test_hashes", "preflight_summary_sha256", "status",
    }, "candidate source binding inventory differs")
    require(binding["schema"] == "xlsx_0552_candidate_source_binding_v1"
            and binding["attempt"] == "draft-07"
            and binding["status"] == "frozen for measurement; admission pending",
            "candidate source binding identity differs")
    verify.parse_time(binding["created_utc"], "candidate source binding timestamp")
    require(binding["source_manifest_sha256"] == candidate_manifest_sha,
            "candidate source binding manifest digest differs")
    for field in ("inputs_sha256", "source_manifest_sha256",
                  "preflight_summary_sha256"):
        hash_value(binding[field], f"candidate binding {field}")

    draft07_dir = CANDIDATE_ATTEMPTS_PATH / "draft-07"
    draft07_inputs_path = draft07_dir / "inputs.json"
    draft07_patch_path = draft07_dir / "candidate.patch"
    draft07 = read_json(draft07_inputs_path, "candidate-attempts/draft-07/inputs.json")
    require(isinstance(draft07, dict)
            and draft07.get("schema") == "xlsx_0552_candidate_attempt_v1",
            "draft-07 candidate attempt schema differs")
    require(binding["inputs_sha256"] == digest(draft07_inputs_path),
            "candidate source binding inputs digest differs")
    require(isinstance(draft07.get("revision"), str)
            and draft07["revision"] == verify.validate_plan()["revision"],
            "draft-07 revision differs from plan")
    require(hash_value(draft07["patch_sha256"], "draft-07 patch hash") ==
            digest(draft07_patch_path, "draft-07 candidate patch"),
            "draft-07 patch digest differs")
    require(draft07.get("application_base_attempt") == "draft-06"
            and draft07.get("logical_parent_attempt") == "draft-06",
            "draft-07 lineage differs")
    require(draft07.get("baseline_restore_sha256") ==
            digest(HERE / "baseline-restore-for-public.json",
                   "baseline-restore-for-public.json"),
            "draft-07 baseline-restore binding differs")
    require(draft07.get("adr_manifest_sha256") == digest(verify.ADR),
            "draft-07 ADR binding differs")
    resources = draft07.get("resources")
    require(isinstance(resources, dict)
            and resources.get("cell_slot_bytes") == 8
            and resources.get("max_proof_logical_heap_bytes") == 2 * 1024 * 1024
            and resources.get("source_byte_cap") == 8 * 1024 * 1024
            and resources.get("source_event_cap") == 131_072,
            "draft-07 resource controls are not bound")

    production = binding["production_source_hashes"]
    public = binding["public_test_hashes"]
    require(isinstance(production, dict) and production,
            "candidate production source hashes are missing")
    require(isinstance(public, dict)
            and set(public) == {PARENT_TEST, COMPACT_CHILD, EXACT_CHILD},
            "candidate public source hash inventory differs")
    for name, value in {**production, **public}.items():
        require(candidate_manifest.get(name) == hash_value(value,
                                                          f"candidate source hash {name}"),
                f"candidate source manifest differs for {name}")
    public_hashes = read_json(PUBLIC_PATH / "source-hashes.json",
                              "public-test-sources/source-hashes.json")
    require(public[PARENT_TEST] == exact["root_parent_sha256"]
            and public[COMPACT_CHILD] ==
            public_hashes[COMPACT_CHILD]["test_sha256"]
            and public[EXACT_CHILD] == exact["terminal_child_sha256"],
            "candidate public source custody differs from both public bundles")

    snapshot_hashes = draft07.get("snapshot_hashes")
    changes = draft07.get("changes")
    require(isinstance(snapshot_hashes, dict) and isinstance(changes, dict),
            "draft-07 source snapshots are missing")
    baseline_manifest, _ = verify.stage_manifest("baseline")
    changed = {name for name in set(baseline_manifest) | set(candidate_manifest)
               if baseline_manifest.get(name) != candidate_manifest.get(name)
               and name in production}
    require(set(snapshot_hashes) == changed,
            "draft-07 snapshot inventory does not cover every production change")
    require(set(changes) == {SHEET_DATA},
            "draft-07 incremental change inventory differs")
    for name, value in snapshot_hashes.items():
        require(value == candidate_manifest.get(name),
                f"draft-07 snapshot hash differs for {name}")
    row = changes[SHEET_DATA]
    draft06 = read_json(
        CANDIDATE_ATTEMPTS_PATH / "draft-06" / "inputs.json",
        "candidate-attempts/draft-06/inputs.json",
    )
    draft06_sheet_data = draft06["snapshot_hashes"][SHEET_DATA]
    require(isinstance(row, dict) and set(row) == {"baseline", "candidate"}
            and row["baseline"] == draft06_sheet_data
            and row["candidate"] == candidate_manifest.get(SHEET_DATA),
            "draft-07 changed source row differs")

    preflight_path = HERE / "preflight-summary-draft07.json"
    preflight = read_json(preflight_path, "preflight-summary-draft07.json")
    require(isinstance(preflight, dict)
            and preflight.get("schema") == "xlsx_0552_preflight_summary_v1"
            and preflight.get("candidate_attempt") == "draft-07"
            and preflight.get("admission") == "pending"
            and preflight.get("performance_measured") is False,
            "draft-07 preflight summary identity differs")
    require(binding["preflight_summary_sha256"] == digest(preflight_path),
            "candidate preflight summary digest differs")
    checks = preflight.get("checks")
    require(isinstance(checks, list)
            and [Path(item.get("attempt", "")).name for item in checks] ==
            list(PRELIGHT_NAMES), "draft-07 preflight check inventory differs")
    test_manifest = dict(candidate_manifest)
    test_manifest[SHEET_DATA] = draft06_sheet_data
    test_manifest = expected_check_manifest(verify, test_manifest)
    stage_check_manifest = expected_check_manifest(verify, candidate_manifest)
    exact_manifest = exact_attempt_manifest(verify, exact["terminal_child_sha256"])
    named: list[dict[str, Any]] = []
    for item in checks:
        require(isinstance(item, dict)
                and set(item) == {"attempt", "command", "exit_code",
                                   "receipt_sha256", "source_manifest_sha256",
                                   "source_stable"},
                "draft-07 preflight row inventory differs")
        name = Path(item["attempt"]).name
        expected_manifest = (
            exact_manifest if name == "baseline-public-exact-05"
            else test_manifest if name in {
                "candidate-proof-tests-03", "candidate-source-integration-02"
            }
            else stage_check_manifest
        )
        expected_command = (
            EXACT_FINAL_COMMAND if name == "baseline-public-exact-05"
            else PRELIGHT_COMMANDS[name]
        )
        row = check_named_receipt(verify, name, expected_manifest,
                                  expected_command, 0)
        require(item["attempt"] == f"docs/performance/results/change-0552/check-attempts/{name}"
                and item["command"] == expected_command
                and item["exit_code"] == 0
                and item["receipt_sha256"] == row["receipt_sha256"]
                and item["source_manifest_sha256"] == row["source_manifest_sha256"]
                and item["source_stable"] is True,
                f"draft-07 preflight row differs: {name}")
        named.append(row)
    return {
        "status": "pass",
        "binding_sha256": digest(binding_path),
        "candidate_manifest_sha256": candidate_manifest_sha,
        "attempt": "draft-07",
        "preflight_summary_sha256": digest(preflight_path),
        "named_checks": named,
        "production_source_count": len(production),
        "public_source_hashes": dict(sorted(public.items())),
    }


def optional_verifier_component(verify: Any, name: str) -> dict[str, Any] | None:
    function = getattr(verify, name, None)
    if function is None:
        return None
    require(callable(function), f"verifier component {name} is not callable")
    value = function()
    require(isinstance(value, dict) and value.get("status") == "pass",
            f"verifier component {name} did not pass")
    return value


def reviewed_rows(raw: list[Any], checked: Any, source: str) -> list[dict[str, Any]]:
    require(isinstance(checked, list) and len(checked) == len(raw),
            f"{source} review length differs")
    remaining = list(checked)
    normalized: list[dict[str, Any]] = []
    for index, original in enumerate(raw):
        require(isinstance(original, dict), f"{source} raw row {index} is malformed")
        matches = [
            (position, row) for position, row in enumerate(remaining)
            if isinstance(row, dict) and row.get("original") == original
        ]
        require(len(matches) == 1,
                f"{source} raw row {index} is not individually reviewed")
        position, row = matches[0]
        for field in ("id", "classification", "interpretation", "disposition"):
            require(nonempty(row.get(field), f"{source} row {index}.{field}"),
                    f"{source} row {index} lacks {field}")
        normalized.append(row)
        remaining.pop(position)
    require(not remaining, f"{source} review has extra rows")
    return normalized


REVIEW_SOURCES = {
    "metrics_adverse": "metrics-analysis.json:comparisons.adverse",
    "metrics_drift": "metrics-analysis.json:repeat_drift_over_five_percent",
    "guard_adverse": "guard-cap-analysis.json:comparison.guard.adverse_flags_over_five_percent",
    "guard_drift": "guard-cap-analysis.json:comparison.guard.same_build_drift_over_five_percent",
    "cap_adverse": "guard-cap-analysis.json:comparison.cap.adverse_flags_over_five_percent",
    "cap_drift": "guard-cap-analysis.json:comparison.cap.same_build_drift_over_five_percent",
}


def validate_adverse_review(metrics: dict[str, Any], metrics_sha: str,
                            guards: dict[str, Any], guards_sha: str) -> dict[str, Any]:
    path = HERE / "adverse-review.json"
    review = read_json(path, "adverse-review.json")
    require(isinstance(review, dict)
            and review.get("schema") == REVIEW_SCHEMA
            and review.get("status") == "complete"
            and review.get("complete") is True
            and review.get("all_diagnostic_rows_retained") is True,
            "adverse review completion envelope differs")
    require(review.get("comparison_sha256") == metrics_sha
            and review.get("guard_analysis_sha256") == guards_sha,
            "adverse review analyzer digest binding differs")
    require(isinstance(review.get("adoption_allowed"), bool),
            "adverse review adoption result is missing")
    raw = {
        "metrics_adverse": extract(metrics, "comparisons.adverse"),
        "metrics_drift": extract(metrics, "repeat_drift_over_five_percent"),
        "guard_adverse": extract(
            guards, "comparison.guard.adverse_flags_over_five_percent"),
        "guard_drift": extract(
            guards, "comparison.guard.same_build_drift_over_five_percent"),
        "cap_adverse": extract(
            guards, "comparison.cap.adverse_flags_over_five_percent"),
        "cap_drift": extract(
            guards, "comparison.cap.same_build_drift_over_five_percent"),
    }
    require(all(isinstance(rows, list) for rows in raw.values()),
            "adverse or drift source is not an array")
    groups = review.get("groups")
    retained: dict[str, list[dict[str, Any]]] = {}
    if groups is not None:
        require(isinstance(groups, list) and len(groups) == len(REVIEW_SOURCES),
                "adverse review group inventory differs")
        by_source: dict[str, Any] = {}
        for index, group in enumerate(groups):
            require(isinstance(group, dict) and set(group) == {"source", "rows"},
                    f"adverse review group {index} is malformed")
            source = nonempty(group["source"], f"adverse review group {index}.source")
            require(source not in by_source, f"adverse review repeats source: {source}")
            require(isinstance(group["rows"], list),
                    f"adverse review group {index}.rows is not an array")
            by_source[source] = group["rows"]
        require(set(by_source) == set(REVIEW_SOURCES.values()),
                "adverse review source inventory differs")
        for key, source in REVIEW_SOURCES.items():
            retained[key] = reviewed_rows(raw[key], by_source[source], source)
    else:
        aliases = {
            "metrics_adverse": ("metrics_adverse", "matched"),
            "metrics_drift": ("metrics_drift", "same_build"),
            "guard_adverse": ("guard_adverse", "guard_adverse_flags"),
            "guard_drift": ("guard_drift", "guard_same_build_drift_flags"),
            "cap_adverse": ("cap_adverse", "cap_adverse_flags"),
            "cap_drift": ("cap_drift", "cap_same_build_drift_flags"),
        }
        for key, names in aliases.items():
            present = [name for name in names if name in review]
            require(len(present) == 1,
                    f"adverse review named array is missing or duplicated: {key}")
            retained[key] = reviewed_rows(raw[key], review[present[0]],
                                          REVIEW_SOURCES[key])
    expected_count = sum(len(rows) for rows in raw.values())
    counts = review.get("counts")
    if counts is not None:
        require(isinstance(counts, dict)
                and counts.get("reviewed_flags") == expected_count,
                "adverse review count does not cover every source row")
    coverage = review.get("source_coverage")
    if coverage is not None:
        require(isinstance(coverage, list), "adverse review source coverage is malformed")
        expected_coverage = [
            {"source": REVIEW_SOURCES[key], "count": len(raw[key])}
            for key in REVIEW_SOURCES
        ]
        require(sorted(coverage, key=lambda row: (row.get("source", ""), row.get("count", -1)))
                == sorted(expected_coverage, key=lambda row: (row["source"], row["count"])),
                "adverse review source coverage differs")
    return {
        "status": "pass",
        "path": "adverse-review.json",
        "sha256": digest(path),
        "schema": REVIEW_SCHEMA,
        "complete": True,
        "adoption_allowed": review["adoption_allowed"],
        "counts": {key: len(rows) for key, rows in raw.items()},
        "reviewed_flags": expected_count,
        "groups": [
            {"source": REVIEW_SOURCES[key], "rows": retained[key]}
            for key in REVIEW_SOURCES
        ],
        "document": review,
    }


def validate_final_source(verify: Any, adoption: bool,
                          candidate_manifest: dict[str, str],
                          candidate_sha: str, exact: dict[str, Any]) -> dict[str, Any]:
    final_manifest, final_sha = verify.stage_manifest("final")
    final_patch = need(HERE / "final" / "source.patch", "final/source.patch")
    expected = candidate_manifest if adoption else exact_public_manifest(
        verify, exact["terminal_child_sha256"]
    )
    require(final_manifest == expected,
            "final source manifest does not match the selected disposition")
    if adoption:
        require(final_sha == candidate_sha,
                "accepted final source manifest is not the candidate manifest")
    else:
        require(final_manifest.get(COMPACT_CHILD) is not None
                and final_manifest.get(EXACT_CHILD) == exact["terminal_child_sha256"],
                "rejected final source does not retain both public child suites")
    # The verifier's generic source replay checks the patch and retained
    # witnesses.  Its older disposition helper intentionally is not called:
    # this decision contract adds the exact public-output child to restoration.
    verify.validate_source_patch("final", final_manifest)
    live = verify.current_source_manifest()
    require(live == final_manifest,
            "live checkout does not equal final source custody")
    return {
        "status": "pass",
        "stage": "final",
        "manifest_sha256": final_sha,
        "manifest_entries": len(final_manifest),
        "selected": "candidate" if adoption else "baseline-plus-public-tests",
        "patch_sha256": digest(final_patch),
        "public_children": {
            COMPACT_CHILD: final_manifest.get(COMPACT_CHILD),
            EXACT_CHILD: final_manifest.get(EXACT_CHILD),
        },
        "candidate_manifest_sha256": candidate_sha,
    }


def validate_quality_final(verify: Any, final_sha: str) -> dict[str, Any]:
    quality = verify.validate_quality()
    canonical_path = HERE / "quality.json"
    document = read_json(canonical_path, "quality.json")
    require(document.get("schema") == "xlsx_0552_quality_v1"
            and document.get("status") == "pass"
            and document.get("source_stage") == "final"
            and document.get("source_manifest_sha256") == final_sha,
            "quality result is not bound to the final source")
    require(quality.get("status") == "pass"
            and quality.get("selected_attempt", "").startswith(
                "quality-attempts/"),
            "quality validator did not select a passing final attempt")
    return {
        "status": "pass",
        "path": "quality.json",
        "sha256": digest(canonical_path),
        "selected_attempt": quality["selected_attempt"],
        "completed_utc": document["completed_utc"],
        "source_stage": document["source_stage"],
        "source_manifest_sha256": document["source_manifest_sha256"],
        "commands": document["commands"],
    }


def validate_guard_amendment(guards: dict[str, Any], guards_sha: str,
                             verify: Any) -> dict[str, Any]:
    """Bind the metadata-only guard analyzer serialization correction."""

    path = HERE / "analyzer-amendment.json"
    value = read_json(path, "analyzer-amendment.json")
    expected_keys = {
        "after_path", "amended_sha256", "before_path", "created_utc",
        "failed_receipt_path", "failed_receipt_sha256",
        "original_analysis_inputs_sha256", "original_sha256", "patch_path",
        "patch_sha256", "path", "reason", "schema", "scope",
    }
    require(isinstance(value, dict) and set(value) == expected_keys,
            "analyzer-amendment envelope differs")
    require(value["schema"] == AMENDMENT_SCHEMA
            and value["path"] == "docs/performance/results/change-0552/analyze_guards.py"
            and value["before_path"] ==
            "docs/performance/results/change-0552/analyzer-amendments/guard-serialization-01/before.py"
            and value["after_path"] ==
            "docs/performance/results/change-0552/analyzer-amendments/guard-serialization-01/after.py"
            and value["patch_path"] ==
            "docs/performance/results/change-0552/analyzer-amendments/guard-serialization-01/change.patch"
            and value["failed_receipt_path"] ==
            "docs/performance/results/change-0552/check-attempts/matched-guards-analysis-01/receipt.json",
            "analyzer-amendment identity differs")
    verify.parse_time(value["created_utc"], "analyzer-amendment.created_utc")
    for field in ("amended_sha256", "original_sha256", "patch_sha256",
                  "failed_receipt_sha256", "original_analysis_inputs_sha256"):
        hash_value(value[field], f"analyzer-amendment.{field}")
    require(value["original_sha256"] == ORIGINAL_GUARD_ANALYZER_SHA256,
            "analyzer-amendment original analyzer hash differs")
    require(value["original_analysis_inputs_sha256"] ==
            digest(HERE / "analysis-inputs.json", "analysis-inputs.json"),
            "analyzer-amendment analysis-inputs binding differs")
    require(value["amended_sha256"] == digest(
        HERE / "analyze_guards.py", "analyze_guards.py"
    ) and value["amended_sha256"] == guards.get("analyzer_sha256"),
            "analyzer-amendment current analyzer binding differs")
    require(value["amended_sha256"] == digest(
        HERE / value["after_path"].removeprefix("docs/performance/results/change-0552/"),
        "analyzer-amendments after.py"
    ), "analyzer-amendment after witness differs")
    require(value["original_sha256"] == digest(
        HERE / value["before_path"].removeprefix("docs/performance/results/change-0552/"),
        "analyzer-amendments before.py"
    ), "analyzer-amendment before witness differs")
    require(value["patch_sha256"] == digest(
        HERE / value["patch_path"].removeprefix("docs/performance/results/change-0552/"),
        "analyzer-amendments change.patch"
    ), "analyzer-amendment patch witness differs")
    require(value["failed_receipt_sha256"] == digest(
        HERE / value["failed_receipt_path"].removeprefix("docs/performance/results/change-0552/"),
        "analyzer-amendment failed receipt"
    ), "analyzer-amendment failed receipt witness differs")
    nonempty(value["reason"], "analyzer-amendment.reason")
    nonempty(value["scope"], "analyzer-amendment.scope")
    return {
        "status": "pass",
        "path": "analyzer-amendment.json",
        "sha256": digest(path),
        "schema": AMENDMENT_SCHEMA,
        "original_sha256": value["original_sha256"],
        "amended_sha256": value["amended_sha256"],
        "report_sha256": guards_sha,
        "document": value,
    }


def validate_profile_decision(verify: Any, metrics: dict[str, Any],
                              metrics_sha: str, profiles: dict[str, Any],
                              pilot_passed: bool, guard_gate: bool,
                              cap_gate: bool) -> dict[str, Any]:
    """Bind the conditional profile lane, including an explicit pilot skip."""

    path = need(PROFILE_DECISION_PATH, "profile-decision.json")
    value = read_json(path, "profile-decision.json")
    expected_keys = {
        "schema", "status", "scope", "pilot_passed", "profile_required",
        "profile_gate_passed", "main_analysis_sha256", "pilot_gates",
        "profile_rows", "reason",
    }
    require(isinstance(value, dict) and set(value) == expected_keys,
            "profile-decision envelope differs")
    require(value["schema"] == PROFILE_DECISION_SCHEMA,
            "profile-decision schema differs")
    require(value["pilot_passed"] is pilot_passed
            and value["main_analysis_sha256"] == metrics_sha,
            "profile-decision pilot/report binding differs")
    hash_value(value["main_analysis_sha256"],
               "profile-decision.main_analysis_sha256")
    pilot_gates = value["pilot_gates"]
    require(isinstance(pilot_gates, dict)
            and set(pilot_gates) == {"main", "guard", "cap"}
            and all(isinstance(item, bool) for item in pilot_gates.values()),
            "profile-decision pilot gates are malformed")
    expected_pilot_gates = {
        "main": metrics["main_gates"]["all_frozen_main_gates_pass"],
        "guard": guard_gate,
        "cap": cap_gate,
    }
    require(pilot_gates == expected_pilot_gates,
            "profile-decision pilot gates differ from main metrics")
    for field in ("scope", "reason"):
        nonempty(value[field], f"profile-decision.{field}")
    require(isinstance(value["profile_rows"], list),
            "profile-decision profile_rows is not an array")
    require(value["profile_gate_passed"] is profiles.get("gate_passed"),
            "profile-decision profile gate differs from profile verifier")
    require(value["profile_required"] is profiles.get("required"),
            "profile-decision required flag differs from profile verifier")
    require(profiles.get("decision") == relative_to_here(path),
            "profile verifier selected a different decision record")
    if not pilot_passed:
        require(value["status"] == "skipped"
                and value["profile_required"] is False
                and value["profile_gate_passed"] is True
                and value["profile_rows"] == [],
                "failed pilot does not have an explicit profile skip")
        for stage in STAGES:
            require(not list((HERE / stage).glob("profile-*.receipt.json")),
                    f"{stage} profile receipts exist after a skipped pilot")
    else:
        require(value["status"] in {"pass", "failed", "reject", "rejected"}
                and value["profile_required"] is True,
                "profile-decision required lane is incomplete")
    return {
        "status": "pass",
        "path": relative_to_here(path),
        "sha256": digest(path),
        "schema": PROFILE_DECISION_SCHEMA,
        "pilot_passed": pilot_passed,
        "required": value["profile_required"],
        "gate_passed": value["profile_gate_passed"],
        "pilot_gates": dict(pilot_gates),
        "profile_rows": value["profile_rows"],
        "document": value,
    }


def evaluate() -> dict[str, Any]:
    verify = verify_module()

    # The optional custody components are used when the audit has installed
    # them in verify.py.  The local checks below remain authoritative for the
    # current bundle so this decision cannot silently depend on an older
    # verifier that predates draft-07 or the exact public child.
    optional_exact = optional_verifier_component(verify, "validate_public_exact_tests")
    optional_candidate = optional_verifier_component(verify, "validate_candidate_attempts")

    exact = validate_public_exact(verify)
    candidate_custody = validate_candidate_custody(verify, exact)

    metrics_result = verify.validate_metrics_analysis()
    guards_result = verify.validate_guard_cap_analysis()
    metrics, metrics_path, metrics_sha = report_document(
        verify, metrics_result, "main metrics", METRICS_SCHEMA)
    guards, guards_path, guards_sha = report_document(
        verify, guards_result, "guard/cap", GUARD_SCHEMA)
    main_gates = main_gate_rows(metrics)
    require(metrics_result.get("status") == "pass"
            and guards_result.get("status") == "pass",
            "canonical analyzer result is not complete")
    guard_comparison = extract(guards, "comparison")
    guard_gate = guards_result.get("guard_admission_passed")
    cap_gate = guards_result.get("cap_admission_passed")
    require(isinstance(guard_gate, bool) and isinstance(cap_gate, bool),
            "guard/cap admission booleans are missing")
    guard_report = guard_comparison["guard"]
    cap_report = guard_comparison["cap"]
    require(guard_report.get("admission_passed") is guard_gate
            and cap_report.get("admission_passed") is cap_gate,
            "guard/cap analyzer result aliases differ")
    gate_rows_pass(guard_report, "guard")
    gate_rows_pass(cap_report, "cap")
    guard_amendment = validate_guard_amendment(guards, guards_sha, verify)

    main_gate = bool(main_gates["all_frozen_main_gates_pass"])
    # The exact commit profile is conditional on the complete pilot envelope:
    # all frozen main gates plus both independent planning/cap gates.  This
    # prevents a partial native/memory improvement from creating a profile
    # obligation after an already-failed main or supplemental gate.
    pilot_passed = bool(main_gate and guard_gate and cap_gate)
    profiles = verify.validate_profiles(pilot_expected=pilot_passed)
    require(profiles.get("pilot_passed") is pilot_passed,
            "profile pilot result differs from main pilot gates")
    profile_gate = profiles.get("gate_passed")
    require(isinstance(profile_gate, bool), "profile gate is missing")
    if pilot_passed:
        require(profiles.get("required") is True,
                "exact commit profile gate was not required")
        references = profiles.get("instruction_references")
        require(isinstance(references, list)
                and len(references) == len(STAGES) * 2 * len(SHAPES)
                and all(isinstance(row, dict) and isinstance(row.get("passed"), bool)
                        for row in references),
                "exact commit profile rows are incomplete")
    else:
        require(profiles.get("required") is False and profile_gate is True,
                "skipped profile lane is not explicitly vacuous")

    profile_decision = validate_profile_decision(
        verify, metrics, metrics_sha, profiles, pilot_passed, guard_gate, cap_gate
    )
    review = validate_adverse_review(metrics, metrics_sha, guards, guards_sha)
    review_gate = bool(review["adoption_allowed"])
    adoption = bool(main_gate and guard_gate and cap_gate and profile_gate and review_gate)

    # Quality and final source are deliberately evaluated after the outcome is
    # known.  A failed performance/review gate must still have a passing final
    # quality run on the restored source, and a passing outcome must have a
    # final checkout equal to the candidate source.
    candidate_manifest, candidate_sha = verify.stage_manifest("candidate")
    final_source = validate_final_source(
        verify, adoption, candidate_manifest, candidate_sha, exact,
    )
    quality = validate_quality_final(verify, final_source["manifest_sha256"])
    quality_gate = quality["status"] == "pass"
    if not quality_gate:
        # This branch is defensive; validate_quality_final normally raises on
        # a malformed or non-passing quality result, so no adoption is emitted.
        adoption = False
        raise IncompleteDecision("final quality gate did not pass")

    # Quality/source custody is an adoption prerequisite.  If final quality
    # was available only after an intended accepted source was restored, fail
    # closed rather than rewriting the result into a narrower claim.
    require(final_source["selected"] == ("candidate" if adoption
                                          else "baseline-plus-public-tests"),
            "final source selection does not match outcome")

    selected_status = "accepted" if adoption else "rejected"
    gate_map = {
        "main": main_gate,
        "guard": guard_gate,
        "cap": cap_gate,
        "profile": profile_gate,
        "profile_required": bool(profiles.get("required")),
        "quality": quality_gate,
        "adverse_review": review_gate,
    }
    return {
        "schema": SCHEMA,
        "status": "pass",
        "scope": SCOPE,
        "disposition": selected_status,
        "adoption_allowed": adoption,
        # The quality completion timestamp is retained as a stable evidence
        # timestamp.  It avoids a wall-clock value that would break replay.
        "observed_utc": quality["completed_utc"],
        "gates": gate_map,
        "main_gate": main_gate,
        "native_primary_gate": main_gate,
        "guard_gate": guard_gate,
        "cap_gate": cap_gate,
        "profile_gate": profile_gate,
        "quality_gate": quality_gate,
        "adverse_review_gate": review_gate,
        "main_gates": main_gates,
        "pilot": {
            "passed": pilot_passed,
            "profile_required": profiles["required"],
        },
        "evidence": {
            "metrics": {
                "path": relative_to_here(metrics_path),
                "sha256": metrics_sha,
                "schema": metrics["schema"],
                "status": metrics["status"],
                "matched_identity": metrics.get("matched_identity"),
                "comparisons": metrics.get("comparisons"),
                "repeat_drift": metrics.get("repeat_drift"),
                "repeat_drift_over_five_percent": metrics.get(
                    "repeat_drift_over_five_percent"),
            },
            "guards": {
                "path": relative_to_here(guards_path),
                "sha256": guards_sha,
                "schema": guards["schema"],
                "status": guards["status"],
                "comparison": guards.get("comparison"),
            },
            "guard_amendment": guard_amendment,
            "profiles": profiles,
            "profile_decision": profile_decision,
            "quality": quality,
            "adverse_review": review,
            "candidate_custody": candidate_custody,
            "public_exact_custody": exact,
            "verifier_components": {
                "validate_public_exact_tests": optional_exact,
                "validate_candidate_attempts": optional_candidate,
            },
        },
        "final_source": final_source,
        "source_manifest_sha256": final_source["manifest_sha256"],
        "metrics_analysis_sha256": metrics_sha,
        "main_analysis_sha256": metrics_sha,
        "guard_analysis_sha256": guards_sha,
        "quality_sha256": quality["sha256"],
        "quality_summary_sha256": quality["sha256"],
        "adverse_review_sha256": review["sha256"],
        "profile_analysis_sha256": profiles.get("decision_sha256"),
        "all_metrics_preserved": True,
        "exact_review_sources": list(REVIEW_SOURCES.values()),
    }


def schema_document() -> dict[str, Any]:
    return {
        "schema": SCHEMA,
        "scope": SCOPE,
        "mode": "outcome-neutral; fail closed; no synthetic evidence",
        "required_evidence": {
            "main": {
                "canonical_reports": [
                    "metrics-analysis.json", "main-analysis.json",
                    "metrics-comparison.json",
                ],
                "schema": METRICS_SCHEMA,
                "gate": "main_gates.all_frozen_main_gates_pass",
                "identity": "main_gates.correctness_identity",
            },
            "guard_cap": {
                "canonical_reports": ["guards-analysis.json", "guard-cap-analysis.json"],
                "schema": GUARD_SCHEMA,
                "amendment": {
                    "canonical": "analyzer-amendment.json",
                    "schema": AMENDMENT_SCHEMA,
                    "scope": "metadata-only serialization correction; numerical/gate logic unchanged",
                },
                "gates": [
                    "comparison.guard.admission_passed",
                    "comparison.cap.admission_passed",
                ],
            },
            "profile": {
                "condition": (
                    "main_gates.all_frozen_main_gates_pass and "
                    "comparison.guard.admission_passed and "
                    "comparison.cap.admission_passed"
                ),
                "canonical": "profile-decision.json",
                "schema": PROFILE_DECISION_SCHEMA,
                "fields": [
                    "status", "scope", "pilot_passed", "profile_required",
                    "profile_gate_passed", "main_analysis_sha256",
                    "pilot_gates", "profile_rows", "reason",
                ],
                "gate": "exact commit Ir decreases for every shape/repeat",
                "otherwise": (
                    "status=skipped, pilot_passed=false, profile_required=false, "
                    "profile_gate_passed=true, profile_rows=[] and no profile receipts"
                ),
            },
            "quality": {
                "canonical": "quality.json",
                "schema": "xlsx_0552_quality_v1",
                "binding": "source_stage=final and source_manifest_sha256=final manifest",
            },
            "source_custody": {
                "candidate": ["candidate-source-binding.json", "candidate-attempts/draft-07"],
                "exact_public": "public-exact-test-sources plus baseline-public-exact-01..05",
                "accepted_final": "final manifest exactly equals candidate manifest",
                "rejected_final": (
                    "final manifest exactly equals frozen baseline plus "
                    "compact_source_proof.rs and public_exact_output.rs"
                ),
            },
            "adverse_review": {
                "canonical": "adverse-review.json",
                "schema": REVIEW_SCHEMA,
                "hashes": ["comparison_sha256", "guard_analysis_sha256"],
                "groups": list(REVIEW_SOURCES.values()),
                "row_rule": "each source array is matched one-for-one by original row",
            },
        },
        "decision_rule": (
            "accepted iff main, guard, cap, conditional profile, final quality, "
            "and adverse-review adoption gates pass; otherwise rejected only "
            "after restored final source custody passes"
        ),
        "rejection_public_children": {
            "compact_source_proof": COMPACT_CHILD,
            "public_exact_output": EXACT_CHILD,
        },
        "no_claims": [
            "No decision is emitted while any required artifact is missing.",
            "Allocator elapsed time is not promoted to latency evidence.",
            "ODF remains deferred and iWork is excluded.",
        ],
    }


def write_identical(path: Path, value: dict[str, Any]) -> None:
    encoded = (json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n").encode()
    if path.exists():
        require(path.is_file() and not path.is_symlink(),
                f"decision output is not a regular file: {path}")
        require(path.read_bytes() == encoded,
                f"existing decision output differs; refusing replacement: {path}")
        return
    if SEAL_PATH.exists() and path.resolve().is_relative_to(HERE.resolve()):
        raise DecisionError("sealed evidence has no retained decision output")
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
    except FileExistsError:
        require(path.read_bytes() == encoded,
                f"decision output changed during exclusive create: {path}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--schema", action="store_true",
                        help="print the required artifact schema without reading evidence")
    parser.add_argument("--output", type=Path, default=DECISION_PATH,
                        help="decision output; exclusive-create/identical-replay")
    args = parser.parse_args(argv)
    if args.schema:
        print(json.dumps(schema_document(), indent=2, sort_keys=True))
        return 0
    try:
        result = evaluate()
        write_identical(args.output, result)
    except IncompleteDecision as error:
        print(f"0552 decision incomplete: {error}", file=sys.stderr)
        return 2
    except (DecisionError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"0552 decision rejected by evidence checks: {error}", file=sys.stderr)
        return 2
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
