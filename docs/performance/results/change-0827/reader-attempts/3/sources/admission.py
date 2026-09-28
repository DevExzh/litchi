"""Root-owned admission for the 0827 ordinary-save validation packet.

The artifact lane is intentionally a hard gate.  It invokes the independent
XML/OPC auditor supplied by the readers lane, replays the independent ZIP
preservation proof, binds every retained artifact file, and writes one
immutable ``artifact-admission-before.json`` or
``artifact-admission-after.json``.  Qualification is a second gate: it is
allowed only after the corresponding artifact admission and checks all twelve
real-format phase oracles before comparative capture can proceed.

Supported invocations are::

    python3 -B admission.py artifacts before
    python3 -B admission.py artifacts after
    python3 -B admission.py qualification before
    python3 -B admission.py qualification after

For compatibility with small driver wrappers, ``before``/``after`` alone
means ``artifacts``.  This module performs no Cargo invocation and does not
run a benchmark or an in-process format reader itself; the auditor and ZIP
checker are separate packet-local validation programs whose complete logs and
reports are retained.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
import time
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PLAN_PATH = P / "plan.json"

# The allocation helper is a sealed, execution-free report-schema validator.
# Admission imports it explicitly so the qualification gate has the same
# vector shape contract as the capture custody checker, while keeping the
# zero-failed-allocation trial policy here.
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))
from tools import perf_allocation_schema as allocation_schema

POLICIES = ("default", "full", "file-only", "no-sync", "stream")
PHASES = {
    "lifecycle": "open+edit+save",
    "edit": "edit",
    "atomic_publish": "save-to-path",
    "counting_publish": "serialize-to-counting-sink",
}
EDIT_DIGEST = hashlib.sha256(b"admitted").hexdigest()
CAPTURE_SCHEMA = "litchi.performance.0827.capture-receipt.v1"

# These are intentionally duplicated as immutable validation constants.  The
# exporter manifest and the packet plan are inputs to this check, not sources
# from which a changed output identity may be learned.
SOURCE_ORACLES = {
    "test-data/ooxml/docx/documentProperties.docx": {
        "bytes": 23_503,
        "sha256": "1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5",
        "published_bytes": 23_535,
        "published_sha256": "de9e163ac26e170ee3881c7d7e53ac29836efd6caea6f72db7ebd2cd31d6c774",
    },
    "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx": {
        "bytes": 8_435,
        "sha256": "d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4",
        "published_bytes": 8_521,
        "published_sha256": "0f6902152f1b8c40c47086206023887e87f54ef57f1da417bc91d8494a4d3e68",
    },
    "test-data/ooxml/pptx/shapes.pptx": {
        "bytes": 68_822,
        "sha256": "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571",
        "published_bytes": 68_284,
        "published_sha256": "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf",
    },
}

REAL_INPUT_BY_FORMAT = {
    "docx": "test-data/ooxml/docx/documentProperties.docx",
    "xlsx": "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx",
    "pptx": "test-data/ooxml/pptx/shapes.pptx",
}

AUDIT_CLOSURES = {
    "DOCX": {"word/document.xml"},
    # The workbook and target worksheet are semantic edit closure.  The
    # independent audit also permits an existing sharedStrings part because
    # a valid XLSX writer may need to rebuild that table; preservation.py
    # still requires the exact historical bytes and therefore catches any
    # unnecessary sharedStrings rewrite.
    "XLSX": {"xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/sharedStrings.xml"},
    "PPTX": {"ppt/slides/slide1.xml"},
}


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def read(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_new(path: Path, value: Any) -> None:
    assert not path.exists(), f"refusing to overwrite {path}"
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def artifact(path: Path) -> dict[str, Any]:
    assert path.is_file() and not path.is_symlink(), f"missing or symlink artifact: {path}"
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def packet_path(value: str | Path) -> Path:
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = P / candidate
    resolved = candidate.resolve()
    assert resolved == P.resolve() or P.resolve() in resolved.parents, (
        f"packet path escapes packet root: {value}"
    )
    return resolved


def verify_descriptor(value: Any, expected: Path) -> dict[str, Any]:
    assert isinstance(value, dict) and isinstance(value.get("path"), str)
    path = packet_path(value["path"])
    assert path == expected.resolve(), f"descriptor path mismatch: {path} != {expected}"
    actual = artifact(path)
    assert value == actual, f"stale descriptor: {path}"
    return actual


def verify_external_descriptor(value: Any, expected: Path) -> dict[str, Any]:
    """Verify an absolute descriptor outside the packet root.

    Build outputs live in the separately owned target directory.  They still
    need the same immutable byte binding as packet files, but ``packet_path``
    must not be used because it intentionally rejects paths outside ``P``.
    """
    assert isinstance(value, dict) and isinstance(value.get("path"), str)
    path = Path(value["path"])
    assert path.is_absolute(), f"external descriptor must be absolute: {value}"
    assert path.resolve() == expected.resolve(), (
        f"external descriptor path mismatch: {path} != {expected}"
    )
    actual = artifact(expected)
    assert value == actual, f"stale external descriptor: {expected}"
    return actual


def verify_timestamp(value: Any, label: str) -> None:
    assert type(value) in (int, float) and value >= 0, label


def plan() -> dict[str, Any]:
    value = read(PLAN_PATH)
    assert value["schema"] == "litchi.performance.0827.plan.v1"
    assert value["expected"]["artifact_cases_per_leg"] == 6
    assert value["expected"]["artifact_policy_outputs_per_case"] == 5
    assert value["expected"]["qualification_reports_per_leg"] == 12
    assert value["expected"]["qualification_samples_per_leg"] == 12
    return value


def input_oracle(path: str) -> dict[str, Any]:
    assert path in SOURCE_ORACLES, f"unknown real input oracle: {path}"
    current = ROOT / path
    expected = SOURCE_ORACLES[path]
    assert current.is_file() and not current.is_symlink()
    assert current.stat().st_size == expected["bytes"]
    assert sha(current) == expected["sha256"]
    return expected


def input_path_matches(value: Any, expected_relative: str) -> bool:
    """Accept the exporter’s absolute or packet-relative spelling only."""
    assert isinstance(value, str) and value
    candidate = Path(value)
    if not candidate.is_absolute():
        candidate = ROOT / candidate
    return candidate.resolve() == (ROOT / expected_relative).resolve()


def plan_cases(value: dict[str, Any]) -> list[dict[str, Any]]:
    cases = value["cases"]
    assert isinstance(cases, list) and len(cases) == 12
    expected: set[str] = set()
    for case in cases:
        assert set(case) == {"case", "format", "input", "phase"}
        assert case["format"] in {"docx", "xlsx", "pptx"}
        assert case["phase"] in PHASES
        assert case["case"] == f"{case['format']}_real_file_ordinary_save_{case['phase']}"
        assert case["input"] in SOURCE_ORACLES
        assert case["case"] not in expected
        expected.add(case["case"])
    assert len(expected) == 12
    return cases


def case_id_for(format_name: str, origin: str) -> str:
    prefix = {"docx": "docx", "xlsx": "xlsx", "pptx": "pptx"}[format_name.lower()]
    if origin == "generated-harness-corpus":
        return f"generated-{prefix}-medium"
    if origin == "caller-named-real-file":
        return {
            "docx": "real-000-docx",
            "xlsx": "real-001-xlsx",
            "pptx": "real-002-pptx",
        }[prefix]
    raise AssertionError(f"unsupported artifact origin: {origin!r}")


def expected_artifact_ids() -> set[str]:
    return {
        "generated-docx-medium",
        "generated-xlsx-medium",
        "generated-pptx-medium",
        "real-000-docx",
        "real-001-xlsx",
        "real-002-pptx",
    }


def artifact_output_path(case: dict[str, Any], spec: dict[str, Any], root: Path) -> Path:
    output = spec.get("output")
    assert isinstance(output, dict) and isinstance(output.get("path"), str)
    candidate = (root / output["path"]).resolve()
    root_resolved = root.resolve()
    assert root_resolved in candidate.parents, f"output escapes artifact root: {candidate}"
    assert candidate.is_file() and not candidate.is_symlink(), candidate
    actual = artifact(candidate)
    assert output == {"path": output["path"], **{k: actual[k] for k in ("bytes", "sha256")}}, (
        f"stale output descriptor: {candidate}"
    )
    return candidate


def validate_complete(leg: str, artifact_root: Path, value: dict[str, Any]) -> dict[str, Any]:
    frozen_plan = plan()
    complete_path = P / frozen_plan["paths"]["artifact_complete"][leg]
    assert complete_path.is_file() and not complete_path.is_symlink()
    complete = read(complete_path)
    assert complete.get("schema") == f"litchi.performance.0827.artifacts-{leg}.complete.v1"
    assert complete.get("status") == "pass"
    assert complete.get("leg") == leg
    assert complete.get("children") == 1
    assert complete.get("cases") == frozen_plan["expected"]["artifact_cases_per_leg"]
    assert complete.get("policy_outputs_per_case") == frozen_plan["expected"]["artifact_policy_outputs_per_case"]
    assert complete.get("reports") == 0
    assert complete.get("samples") == 0
    manifest_path = artifact_root / "manifest.json"
    assert manifest_path.is_file()
    verify_descriptor(complete["plan"], PLAN_PATH)
    build_path = P / f"build-{leg}" / "build.json"
    verify_descriptor(complete["build"], build_path)
    source_path = P / f"artifacts-{leg}-source.json"
    verify_descriptor(complete["source"], source_path)
    receipt_path = P / f"artifacts-{leg}-receipt.json"
    verify_descriptor(complete["receipt"], receipt_path)
    verify_descriptor(complete["manifest"], manifest_path)
    return artifact(complete_path)


def run_auditor(leg: str, artifact_root: Path, attempt: Path) -> tuple[Path, dict[str, Any]]:
    auditor = P / "artifact_audit.py"
    assert auditor.is_file() and not auditor.is_symlink()
    report_path = attempt / "artifact-audit.json"
    log_path = attempt / "artifact-audit.log"
    command = [
        sys.executable,
        "-B",
        str(auditor),
        "--artifacts",
        str(artifact_root),
        "--report",
        str(report_path),
    ]
    started = time.time()
    completed: subprocess.CompletedProcess[str] | None = None
    launch_error: str | None = None
    try:
        with log_path.open("x", encoding="utf-8") as stream:
            completed = subprocess.run(command, cwd=ROOT, stdout=stream, stderr=subprocess.STDOUT)
    except Exception as exc:
        launch_error = f"{type(exc).__name__}: {exc}"
    ended = time.time()
    # Keep the wrapper's process receipt beside the child's log before making
    # any admission assertion.  A nonzero child may still have emitted a
    # useful report, and both that report and the complete log must remain
    # bound for offline diagnosis.
    receipt = {
        "schema": "litchi.performance.0827.artifact-audit-attempt.v1",
        "leg": leg,
        "command": command,
        "started": started,
        "ended": ended,
        "exit_code": completed.returncode if completed is not None else None,
        "error": launch_error,
        "auditor": artifact(auditor),
        "artifact_directory": str(artifact_root),
        "report": (
            artifact(report_path)
            if report_path.is_file() and not report_path.is_symlink()
            else None
        ),
        # Opening the child log can itself fail (permissions, a stale path,
        # or a filesystem error).  Preserve the launch receipt with a null
        # log descriptor instead of losing the receipt to a second exception.
        "log": (
            artifact(log_path)
            if log_path.is_file() and not log_path.is_symlink()
            else None
        ),
    }
    write_new(attempt / "artifact-audit-receipt.json", receipt)
    assert launch_error is None, f"independent artifact audit could not start; retained {log_path}"
    assert completed is not None
    assert completed.returncode == 0, f"independent artifact audit failed; retained {log_path}"
    assert report_path.is_file() and not report_path.is_symlink(), report_path
    value = read(report_path)
    assert value.get("schema") == "litchi.performance.0827.artifact-audit.v1"
    assert value.get("ok") is True and not value.get("errors")
    rows = value.get("cases")
    assert isinstance(rows, list) and len(rows) == 6
    assert all(row.get("ok") is True for row in rows)
    return report_path, value


def run_preservation(leg: str, attempt: Path) -> Path:
    script = P / "preservation.py"
    assert script.is_file() and not script.is_symlink()
    report_path = P / f"zip-preservation-{leg}.json"

    def run_checked(action: str, command: list[str], log_path: Path) -> None:
        started = time.time()
        completed: subprocess.CompletedProcess[str] | None = None
        launch_error: str | None = None
        try:
            with log_path.open("x", encoding="utf-8") as stream:
                completed = subprocess.run(command, cwd=ROOT, stdout=stream, stderr=subprocess.STDOUT)
        except Exception as exc:
            launch_error = f"{type(exc).__name__}: {exc}"
        ended = time.time()
        # Write this before checking the child result so a failed write or
        # replay still leaves its command, timing, exit state, log, and any
        # partial preservation report available to the parent admission.
        receipt = {
            "schema": "litchi.performance.0827.preservation-attempt.v1",
            "leg": leg,
            "action": action,
            "command": command,
            "started": started,
            "ended": ended,
            "exit_code": completed.returncode if completed is not None else None,
            "error": launch_error,
            "script": artifact(script),
            "report": (
                artifact(report_path)
                if report_path.is_file() and not report_path.is_symlink()
                else None
            ),
            "log": (
                artifact(log_path)
                if log_path.is_file() and not log_path.is_symlink()
                else None
            ),
        }
        write_new(attempt / f"preservation-{action}-receipt.json", receipt)
        assert launch_error is None, f"preservation {action} could not start; retained {log_path}"
        assert completed is not None
        assert completed.returncode == 0, f"preservation {action} failed; retained {log_path}"

    if not report_path.exists():
        write_log = attempt / "preservation-write.log"
        command = [sys.executable, "-B", str(script), "--leg", leg, "--write"]
        run_checked("write", command, write_log)
    check_log = attempt / "preservation-check.log"
    command = [sys.executable, "-B", str(script), "--leg", leg, "--check"]
    run_checked("check", command, check_log)
    assert report_path.is_file() and not report_path.is_symlink()
    preservation = read(report_path)
    assert preservation.get("schema") == "litchi.performance.0827.zip-preservation.v1"
    assert preservation.get("leg") == leg and preservation.get("ok") is True
    assert len(preservation.get("cases", [])) == 6
    checks = preservation.get("checks")
    assert isinstance(checks, dict) and checks and all(value is True for value in checks.values())
    all_files = preservation.get("all_files")
    assert isinstance(all_files, list) and len(all_files) == 37
    return report_path


def validate_audit_against_manifest(audit: dict[str, Any], manifest: dict[str, Any], artifact_root: Path) -> list[dict[str, Any]]:
    manifest_cases = {str(case["case_id"]): case for case in manifest["cases"]}
    assert set(manifest_cases) == expected_artifact_ids()
    audit_cases = {str(case["case_id"]): case for case in audit["cases"]}
    assert set(audit_cases) == expected_artifact_ids()
    selectors: list[dict[str, Any]] = []
    for case_id in sorted(expected_artifact_ids()):
        case = manifest_cases[case_id]
        audited = audit_cases[case_id]
        assert audited.get("format") == case.get("format")
        assert audited.get("origin") == case.get("origin")
        assert audited.get("edit_admitted") is True
        assert audited.get("edit_outcome") == "admitted"
        allowed = audited.get("allowed_changed_members")
        changed = audited.get("changed_members")
        assert isinstance(allowed, list) and isinstance(changed, list)
        changed_names = sorted(
            item.get("name") if isinstance(item, dict) else item for item in changed
        )
        assert set(allowed) == AUDIT_CLOSURES[str(case["format"])]
        assert set(changed_names) == (
            {"word/document.xml"}
            if case["format"] == "DOCX"
            else {"xl/workbook.xml", "xl/worksheets/sheet1.xml"}
            if case["format"] == "XLSX"
            else {"ppt/slides/slide1.xml"}
        )
        assert set(changed_names) <= set(allowed)
        policy_rows = audited.get("policy_outputs")
        assert isinstance(policy_rows, list) and len(policy_rows) == 5
        assert {row.get("policy") for row in policy_rows} == set(POLICIES)
        expected_format = str(case["format"]).lower()
        input_path = (
            REAL_INPUT_BY_FORMAT[expected_format]
            if case.get("origin") == "caller-named-real-file"
            else None
        )
        if input_path is not None:
            oracle = input_oracle(input_path)
            assert case.get("source_archive_sha256") == oracle["sha256"]
            assert case.get("source_archive_bytes") == oracle["bytes"]
            published_sha = oracle["published_sha256"]
            published_bytes = oracle["published_bytes"]
        else:
            # Generated cases are checked by preservation.py against fixed
            # historical source/output identities.  The audit still must
            # expose complete source/output identity fields.
            generated = {
                "generated-docx-medium": (9_051, "2004b8046b78320f2d05f27bce1ef293c1783912e2f27c83674e8f7d4c9431d9", 9_070, "6e2a7afe4d670787c7d428032a8fc82ca866ab9e17ba50523ad108c366bf3775"),
                "generated-xlsx-medium": (4_226_429, "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036", 4_226_568, "20335c4480405e051be2f4fea0e2a5aa8c5912b729cb1df980112ea7c5056687"),
                "generated-pptx-medium": (40_788, "50ad2f81099ee29d4768d7080b7fc51ea2b5ca2aadd531031efcab65e8d5409e", 40_802, "95efb7b3b4f621b68961bb3db87618cbe59065650a11b64ead3debc8e598d2cd"),
            }
            source_bytes, source_sha, published_bytes, published_sha = generated[case_id]
            assert case.get("source_archive_bytes") == source_bytes
            assert case.get("source_archive_sha256") == source_sha
        assert case.get("published_bytes") == published_bytes
        assert case.get("published_sha256") == published_sha
        for row in policy_rows:
            assert row.get("bytes") == published_bytes
            assert row.get("sha256") == published_sha
            assert row.get("matches_default") is True
            assert row.get("matches_source") is False
            inventory = row.get("inventory")
            assert isinstance(inventory, list) and inventory
            content_types = row.get("content_types")
            relationships = row.get("relationships")
            assert isinstance(content_types, dict) and content_types.get("valid") is True
            assert isinstance(relationships, dict) and relationships.get("valid") is True
            output = row.get("path")
            assert isinstance(output, str)
            assert (artifact_root / output).resolve().is_file()
        selectors.append(
            {
                "case": f"{expected_format}_real_file_ordinary_save_lifecycle",
                "format": expected_format,
                "input": input_path,
                "source_bytes": case.get("source_archive_bytes"),
                "source_sha256": case.get("source_archive_sha256"),
                "published_bytes": published_bytes,
                "published_sha256": published_sha,
                "edit_outcome": "admitted",
            }
        )
    # Selectors are created from artifact cases for the 12 qualification
    # rows below; this function returns the six corpus identities and the
    # qualification helper expands them per phase.
    return selectors


def artifact_admission(leg: str) -> dict[str, Any]:
    value = plan()
    assert leg in {"before", "after"}
    output = P / value["paths"]["artifact_output"][leg]
    assert output.is_dir() and not output.is_symlink()
    complete_path = P / value["paths"]["artifact_complete"][leg]
    assert complete_path.is_file() and not complete_path.is_symlink()
    manifest_path = output / "manifest.json"
    manifest = read(manifest_path)
    assert set(manifest) >= {"schema_version", "kind", "generator", "cases"}
    assert manifest["schema_version"] == 1
    assert manifest["kind"] == "ordinary-save-artifact-export"
    assert manifest["generator"] == "litchi-perf-ordinary-save-artifacts-v1"
    assert isinstance(manifest["cases"], list) and len(manifest["cases"]) == 6
    assert not (P / value["paths"]["qualification_output"][leg]).exists(), (
        "artifact admission must precede qualification"
    )

    # Every artifact output is required to bind one of the fixed real inputs;
    # this catches an exporter that silently changes the corpus order or
    # substitutes a different file while leaving its own metadata coherent.
    for case in manifest["cases"]:
        origin = case.get("origin")
        fmt = str(case.get("format", "")).lower()
        assert case.get("case_id") == case_id_for(fmt, origin)
        if origin == "caller-named-real-file":
            path = case.get("input_path")
            expected_path = REAL_INPUT_BY_FORMAT[fmt]
            assert input_path_matches(path, expected_path)
            input_oracle(expected_path)
        else:
            assert case.get("input_path") is None

    complete = validate_complete(leg, output, value)
    attempt = P / f"admission-{leg}"
    retry = 0
    while attempt.exists():
        retry += 1
        attempt = P / f"admission-{leg}-retry{retry}"
    attempt.mkdir()
    audit_path, audit = run_auditor(leg, output, attempt)
    preservation_path = run_preservation(leg, attempt)
    selectors_from_audit = validate_audit_against_manifest(audit, manifest, output)
    preservation = read(preservation_path)
    assert preservation["all_files"]

    # Expand the six fixed artifact identities to the twelve planned phase
    # selectors.  The capture driver consumes these rows as its oracle; phase
    # timing itself is never borrowed from historical evidence.
    by_format: dict[str, dict[str, Any]] = {}
    for row in selectors_from_audit:
        if row["input"] is not None:
            by_format[row["format"]] = row
    selectors: list[dict[str, Any]] = []
    for case in plan_cases(value):
        row = by_format[case["format"]]
        selectors.append(
            {
                "case": case["case"],
                "format": case["format"],
                "phase": case["phase"],
                "input": case["input"],
                "source_bytes": row["source_bytes"],
                "source_sha256": row["source_sha256"],
                "published_bytes": row["published_bytes"],
                "published_sha256": row["published_sha256"],
                "edit_outcome": "admitted",
            }
        )

    checks = {
        "fresh_manifest_identity": True,
        "fresh_all_files_bound": True,
        "independent_xml_zip_audit": True,
        "independent_zip_preservation": True,
        "historical_0819_0821_byte_identity": True,
        "source_inputs_exact": True,
        "five_policy_outputs_per_case": True,
        "chosen_edit_closures_exact": True,
        "qualification_not_started": True,
    }
    attempt_files = {
        str(path.relative_to(P)): artifact(path)
        for path in sorted(attempt.rglob("*"))
        if path.is_file() and not path.is_symlink()
    }
    result = {
        "schema": "litchi.performance.0827.artifact-admission.v1",
        "accepted": True,
        "leg": leg,
        "plan_sha256": sha(PLAN_PATH),
        "artifact_complete": artifact(complete_path),
        "manifest": artifact(manifest_path),
        "audit": artifact(audit_path),
        "auditor": artifact(P / "artifact_audit.py"),
        "zip_preservation": artifact(preservation_path),
        "selectors": selectors,
        "checks": checks,
        "fresh_bound": {
            "artifact_files": preservation["all_files"],
            "attempt_files": attempt_files,
            "manifest": artifact(manifest_path),
            "complete": complete,
            "audit": artifact(audit_path),
            "preservation": artifact(preservation_path),
        },
        "historical_reference": {
            "source": "docs/performance/results/change-0819/artifacts and change-0821/artifacts",
            "identity": "exact bytes and SHA-256 for all six default outputs; no timing reuse",
        },
        "ended": time.time(),
    }
    target = P / value["paths"]["artifact_admission"][leg]
    write_new(target, result)
    print(f"0827 artifact admission {leg} PASS")
    return result


def descriptor_from_receipt(value: Any, expected: Path) -> dict[str, Any]:
    return verify_descriptor(value, expected)


def validate_build(leg: str, frozen_plan: dict[str, Any]) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    """Bind the qualification observer to the current frozen build."""
    build_path = P / f"build-{leg}" / "build.json"
    source_path = build_path.parent / "source.json"
    frozen_inputs_path = build_path.parent / "frozen-inputs.json"
    assert build_path.is_file() and not build_path.is_symlink()
    assert source_path.is_file() and not source_path.is_symlink()
    assert frozen_inputs_path.is_file() and not frozen_inputs_path.is_symlink()
    build = read(build_path)
    assert build.get("schema") == f"litchi.performance.0827.build-{leg}.v1"
    assert build.get("leg") == leg
    assert build.get("target") == frozen_plan["target"]
    assert build.get("profile") == frozen_plan["build"]
    verify_descriptor(build["source"], source_path)
    verify_descriptor(build["frozen_inputs"], frozen_inputs_path)
    assert read(source_path).get("revision") == frozen_plan["base"]
    assert read(frozen_inputs_path) == read(P / "freeze.json")

    observer_plan = frozen_plan["binaries"]["observer"]
    observer = build.get("binaries", {}).get("observer")
    assert isinstance(observer, dict)
    assert observer.get("cargo_bin") == observer_plan["cargo_bin"]
    assert observer.get("features") == observer_plan["features"]
    observer_path = Path(frozen_plan["target"]) / f"{leg}-observer"
    observer_descriptor = verify_external_descriptor(observer["artifact"], observer_path)
    return build_path, build, observer_descriptor


def admission_attempt_root(leg: str, attempt_files: Any) -> Path:
    """Resolve the one retained artifact-admission attempt directory."""
    assert isinstance(attempt_files, dict) and attempt_files
    roots: set[str] = set()
    prefix = f"admission-{leg}"
    for raw in attempt_files:
        assert isinstance(raw, str) and raw
        relative = Path(raw)
        assert not relative.is_absolute() and ".." not in relative.parts
        assert relative.parts
        root_name = relative.parts[0]
        retry_suffix = root_name[len(prefix) + len("-retry"):]
        is_retry = root_name.startswith(prefix + "-retry") and retry_suffix.isdigit()
        assert root_name == prefix or is_retry
        roots.add(root_name)
    assert len(roots) == 1
    root = P / next(iter(roots))
    assert root.is_dir() and not root.is_symlink()
    return root


def verify_tree_binding(root: Path, rows: Any) -> None:
    """Require preservation's relative file descriptors to cover the tree."""
    assert isinstance(rows, list) and rows
    expected: dict[str, dict[str, Any]] = {}
    root_resolved = root.resolve()
    for row in rows:
        assert isinstance(row, dict) and set(row) == {"path", "bytes", "sha256"}
        relative_name = row["path"]
        assert isinstance(relative_name, str) and relative_name
        relative = Path(relative_name)
        assert not relative.is_absolute() and ".." not in relative.parts
        assert "\\" not in relative_name
        path = (root / relative).resolve()
        assert root_resolved in path.parents
        assert path.is_file() and not path.is_symlink()
        observed = {
            "path": relative_name,
            "bytes": path.stat().st_size,
            "sha256": sha(path),
        }
        assert row == observed
        assert relative_name not in expected
        expected[relative_name] = observed
    actual: dict[str, dict[str, Any]] = {}
    for path in sorted(root.rglob("*")):
        assert not path.is_symlink(), f"symlink in bound artifact tree: {path}"
        if path.is_file():
            relative_name = str(path.relative_to(root))
            actual[relative_name] = {
                "path": relative_name,
                "bytes": path.stat().st_size,
                "sha256": sha(path),
            }
    assert actual == expected


def validate_artifact_admission_record(
    leg: str,
    frozen_plan: dict[str, Any],
    artifact_value: dict[str, Any],
) -> dict[str, Any]:
    """Rebind every artifact-admission descriptor before qualification."""
    assert artifact_value.get("schema") == "litchi.performance.0827.artifact-admission.v1"
    assert artifact_value.get("accepted") is True
    assert artifact_value.get("leg") == leg
    assert artifact_value.get("plan_sha256") == sha(PLAN_PATH)
    artifact_root = P / frozen_plan["paths"]["artifact_output"][leg]
    assert artifact_root.is_dir() and not artifact_root.is_symlink()
    complete_path = P / frozen_plan["paths"]["artifact_complete"][leg]
    validate_complete(leg, artifact_root, frozen_plan)
    manifest_path = artifact_root / "manifest.json"
    verify_descriptor(artifact_value["artifact_complete"], complete_path)
    verify_descriptor(artifact_value["manifest"], manifest_path)
    verify_descriptor(artifact_value["auditor"], P / "artifact_audit.py")
    audit_path = packet_path(artifact_value["audit"]["path"])
    verify_descriptor(artifact_value["audit"], audit_path)
    preservation_path = P / f"zip-preservation-{leg}.json"
    verify_descriptor(artifact_value["zip_preservation"], preservation_path)

    audit = read(audit_path)
    assert audit.get("schema") == "litchi.performance.0827.artifact-audit.v1"
    assert audit.get("ok") is True and audit.get("errors") == []
    assert isinstance(audit.get("cases"), list) and len(audit["cases"]) == 6
    assert all(case.get("ok") is True for case in audit["cases"])
    assert audit.get("artifact_directory") == str(artifact_root)
    preservation = read(preservation_path)
    assert preservation.get("schema") == "litchi.performance.0827.zip-preservation.v1"
    assert preservation.get("leg") == leg and preservation.get("ok") is True
    assert preservation.get("artifact_directory") == str(artifact_root)
    assert isinstance(preservation.get("cases"), list) and len(preservation["cases"]) == 6
    preservation_checks = preservation.get("checks")
    assert isinstance(preservation_checks, dict) and preservation_checks
    assert all(item is True for item in preservation_checks.values())
    all_files = preservation.get("all_files")
    assert isinstance(all_files, list) and len(all_files) == 37
    verify_tree_binding(artifact_root, all_files)

    fresh_bound = artifact_value.get("fresh_bound")
    assert isinstance(fresh_bound, dict)
    assert set(fresh_bound) == {
        "artifact_files", "attempt_files", "manifest", "complete", "audit", "preservation"
    }
    assert fresh_bound["artifact_files"] == all_files
    verify_descriptor(fresh_bound["manifest"], manifest_path)
    verify_descriptor(fresh_bound["complete"], complete_path)
    verify_descriptor(fresh_bound["audit"], audit_path)
    verify_descriptor(fresh_bound["preservation"], preservation_path)

    attempt_files = fresh_bound["attempt_files"]
    attempt_root = admission_attempt_root(leg, attempt_files)
    expected_attempt_files: dict[str, dict[str, Any]] = {}
    for relative_name, descriptor in attempt_files.items():
        assert isinstance(relative_name, str)
        expected_path = P / relative_name
        assert expected_path.resolve().is_relative_to(attempt_root.resolve())
        verify_descriptor(descriptor, expected_path)
        expected_attempt_files[relative_name] = artifact(expected_path)
    actual_attempt_files: dict[str, dict[str, Any]] = {}
    for path in sorted(attempt_root.rglob("*")):
        assert not path.is_symlink(), f"symlink in admission attempt: {path}"
        if path.is_file():
            actual_attempt_files[str(path.relative_to(P))] = artifact(path)
    assert actual_attempt_files == expected_attempt_files
    assert audit_path == attempt_root / "artifact-audit.json"

    checks = artifact_value.get("checks")
    required_checks = {
        "fresh_manifest_identity", "fresh_all_files_bound", "independent_xml_zip_audit",
        "independent_zip_preservation", "historical_0819_0821_byte_identity",
        "source_inputs_exact", "five_policy_outputs_per_case", "chosen_edit_closures_exact",
        "qualification_not_started",
    }
    assert isinstance(checks, dict) and required_checks <= set(checks)
    assert all(checks[key] is True for key in required_checks)

    expected_cases = {case["case"]: case for case in plan_cases(frozen_plan)}
    selectors = artifact_value.get("selectors")
    assert isinstance(selectors, list) and len(selectors) == len(expected_cases)
    seen: set[str] = set()
    for selector in selectors:
        assert isinstance(selector, dict)
        case_name = selector.get("case")
        assert case_name in expected_cases and case_name not in seen
        case = expected_cases[case_name]
        oracle = SOURCE_ORACLES[case["input"]]
        assert selector.get("format") == case["format"]
        assert selector.get("phase") == case["phase"]
        assert selector.get("input") == case["input"]
        assert selector.get("source_bytes") == oracle["bytes"]
        assert selector.get("source_sha256") == oracle["sha256"]
        assert selector.get("published_bytes") == oracle["published_bytes"]
        assert selector.get("published_sha256") == oracle["published_sha256"]
        assert selector.get("edit_outcome") == "admitted"
        seen.add(case_name)
    assert seen == set(expected_cases)
    return artifact_value


def qualification_selector(case: dict[str, Any]) -> dict[str, Any]:
    oracle = SOURCE_ORACLES[case["input"]]
    return {
        "case": case["case"],
        "format": case["format"],
        "phase": case["phase"],
        "input": case["input"],
        "source_bytes": oracle["bytes"],
        "source_sha256": oracle["sha256"],
        "published_bytes": oracle["published_bytes"],
        "published_sha256": oracle["published_sha256"],
        "edit_outcome": "admitted",
    }


def qualification_command(
    frozen_plan: dict[str, Any],
    case: dict[str, Any],
    observer_descriptor: dict[str, Any],
    report_path: Path,
    rss_path: Path,
) -> list[str]:
    lane = frozen_plan["lanes"]["qualification"]
    return [
        "/usr/bin/time", "-f", "%M", "-o", str(rss_path),
        "taskset", "-c", str(frozen_plan["cpu"]), observer_descriptor["path"],
        "--warmup", str(lane["warmup"]),
        "--samples", str(lane["samples"]),
        "--case", case["case"],
        "--json", str(report_path),
        "--filesystem-root", str(frozen_plan["scratch"]),
        "--ooxml-file", str(ROOT / case["input"]),
    ]


def validate_zero_failed_allocations(result: dict[str, Any], samples: int, label: str) -> None:
    """Apply the qualification policy after the reusable shape validator."""
    allocation = result.get("operation_metrics", {}).get("allocation")
    assert isinstance(allocation, dict), label
    failed = allocation.get("failed_allocation_calls")
    assert isinstance(failed, dict), f"{label}.failed_allocation_calls"
    values = failed.get("values")
    assert isinstance(values, list) and len(values) == samples
    assert all(item == 0 for item in values), f"{label}.failed_allocation_calls.values"


def validate_qualification_report(
    report: dict[str, Any],
    case: dict[str, Any],
    selector: dict[str, Any],
    samples: int = 1,
    warmup: int = 0,
    binary_descriptor: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Purely validate one observer report and all semantic format oracles."""
    allocation_schema.validate_report(
        report,
        expected_case=case["case"],
        expected_samples=samples,
        expected_warmup=warmup,
        mode="observer",
    )
    tool = report.get("tool")
    assert isinstance(tool, dict)
    assert tool.get("name") == "litchi-perf-baseline"
    assert tool.get("binary") == "litchi-perf-baseline-alloc"
    assert tool.get("instrumentation") == (
        "ordinary_save_procfs_and_system_allocator_operation_scoped"
    )
    assert tool.get("allocator_counter_revision") == allocation_schema.REVISION
    identity = report.get("binary_identity")
    assert isinstance(identity, dict)
    assert isinstance(identity.get("path"), str) and identity["path"]
    assert isinstance(identity.get("binary_sha256"), str) and len(identity["binary_sha256"]) == 64
    assert type(identity.get("binary_bytes")) is int and identity["binary_bytes"] > 0
    assert identity.get("profile") == "release"
    if binary_descriptor is not None:
        assert identity.get("path") == binary_descriptor["path"]
        assert identity.get("binary_sha256") == binary_descriptor["sha256"]
        assert identity.get("binary_bytes") == binary_descriptor["bytes"]

    configuration = report.get("configuration")
    assert isinstance(configuration, dict)
    assert configuration.get("samples_per_case") == samples
    assert configuration.get("warmup_iterations_per_case") == warmup
    results = report.get("results")
    assert isinstance(results, list) and len(results) == 1
    result = results[0]
    assert isinstance(result, dict) and result.get("case") == case["case"]
    validate_zero_failed_allocations(result, samples, case["case"])

    ordinary = result.get("source", {}).get("ordinary_save")
    assert isinstance(ordinary, dict)
    assert ordinary.get("format") == case["format"].upper()
    assert ordinary.get("origin") == "caller-named-real-file"
    assert ordinary.get("phase") == PHASES[case["phase"]]
    assert ordinary.get("save_durability") in (None, "default")
    assert isinstance(ordinary.get("atomic_publication_steps"), str)
    corpus = ordinary.get("corpus")
    assert isinstance(corpus, dict)
    oracle = SOURCE_ORACLES[case["input"]]
    assert corpus.get("real_file") == {
        "path": str(ROOT / case["input"]),
        "bytes": oracle["bytes"],
        "sha256": oracle["sha256"],
    }
    assert corpus.get("source_archive_bytes") == selector["source_bytes"]
    assert corpus.get("source_archive_sha256") == selector["source_sha256"]
    assert corpus.get("published_bytes") == selector["published_bytes"]
    assert corpus.get("published_sha256") == selector["published_sha256"]
    assert corpus.get("edit_admitted") is True and corpus.get("edit_outcome") == "admitted"
    assert corpus.get("repeated_cycles_identical") is True
    assert corpus.get("repeated_saves_identical") is True
    assert ordinary.get("publications_identical") is True
    assert ordinary.get("edit_outcomes_identical") is True
    assert ordinary.get("edit_outcome_sha256") == [EDIT_DIGEST]
    expected_published = [] if case["phase"] == "edit" else [selector["published_sha256"]]
    assert ordinary.get("published_sha256") == expected_published
    process_probe = ordinary.get("process_probe")
    assert isinstance(process_probe, dict) and process_probe.get("fixed_count") == 32
    if case["phase"] == "counting_publish":
        assert isinstance(ordinary.get("sample_byte_split"), dict)
    elapsed = result.get("elapsed_ns")
    assert isinstance(elapsed, dict) and isinstance(elapsed.get("samples"), list)
    assert len(elapsed["samples"]) == samples
    assert all(type(item) is int and item > 0 for item in elapsed["samples"])
    return report


def validate_qualification_complete(
    leg: str,
    frozen_plan: dict[str, Any],
    build_path: Path,
    source_path: Path,
    receipts_path: Path,
    complete_path: Path,
) -> dict[str, Any]:
    complete = read(complete_path)
    lane = frozen_plan["lanes"]["qualification"]
    assert complete.get("schema") == f"litchi.performance.0827.qualification-{leg}.complete.v1"
    assert complete.get("status") == "pass"
    assert complete.get("lane") == "qualification"
    assert complete.get("leg") == leg
    assert complete.get("blocks") == lane["blocks"]
    assert complete.get("reports") == frozen_plan["expected"]["qualification_reports_per_leg"]
    assert complete.get("samples") == frozen_plan["expected"]["qualification_samples_per_leg"]
    assert complete.get("expected_reports") == frozen_plan["expected"]["qualification_reports_per_leg"]
    assert complete.get("expected_samples") == frozen_plan["expected"]["qualification_samples_per_leg"]
    verify_descriptor(complete["plan"], PLAN_PATH)
    verify_descriptor(complete["build"], build_path)
    verify_descriptor(complete["source"], source_path)
    verify_descriptor(complete["receipts"], receipts_path)
    return complete


def validate_qualification_receipt(
    receipt: dict[str, Any],
    index: int,
    leg: str,
    frozen_plan: dict[str, Any],
    case: dict[str, Any],
    selector: dict[str, Any],
    qualification_root: Path,
    source_path: Path,
    artifact_admission_path: Path,
    observer_descriptor: dict[str, Any],
) -> tuple[Path, dict[str, Any]]:
    lane = frozen_plan["lanes"]["qualification"]
    stem = f"{index:02d}-{case['format']}-{case['phase']}"
    report_path = qualification_root / f"{stem}.json"
    rss_path = qualification_root / f"{stem}.rss"
    log_path = qualification_root / f"{stem}.log"
    assert receipt.get("schema") == CAPTURE_SCHEMA
    assert receipt.get("lane") == "qualification"
    assert receipt.get("leg") == leg
    assert receipt.get("block") == 0
    for key in ("case", "format", "input", "phase"):
        assert receipt.get(key) == case[key]
    assert receipt.get("samples") == lane["samples"]
    assert receipt.get("warmup") == lane["warmup"]
    assert receipt.get("exit_code") == 0
    if "policy" in receipt:
        assert receipt["policy"] == "default"
    verify_timestamp(receipt.get("started"), f"receipt {index} started")
    verify_timestamp(receipt.get("ended"), f"receipt {index} ended")
    assert receipt["ended"] >= receipt["started"]
    assert receipt.get("command") == qualification_command(
        frozen_plan, case, observer_descriptor, report_path, rss_path
    )
    verify_external_descriptor(receipt["binary"], Path(observer_descriptor["path"]))
    verify_descriptor(receipt["source"], source_path)
    verify_descriptor(receipt["frozen_inputs"], P / "freeze.json")
    verify_descriptor(receipt["artifact_admission"], artifact_admission_path)
    verify_descriptor(receipt["log"], log_path)
    verify_descriptor(receipt["report"], report_path)
    verify_descriptor(receipt["rss"], rss_path)
    rss_value = rss_path.read_text(encoding="utf-8").strip()
    assert rss_value.isdigit() and int(rss_value) > 0
    report = read(report_path)
    validate_qualification_report(
        report,
        case,
        selector,
        samples=lane["samples"],
        warmup=lane["warmup"],
        binary_descriptor=observer_descriptor,
    )
    return report_path, report


def qualification_preflight() -> dict[str, Any]:
    """Exercise qualification report checks on retained historical evidence.

    The function reads twelve 0819 qualification reports and the retained
    0825 failed-qualification report.  It does not create descriptors for a
    fresh run, write a result, or reuse historical timing as a measurement.
    """
    frozen_plan = plan()
    cases = plan_cases(frozen_plan)
    rows: list[dict[str, Any]] = []
    counts: dict[str, dict[str, int]] = {
        "0819/qualification": {"reports": 0, "sample_envelopes": 0},
        "0825/failed-qualification": {"reports": 0, "sample_envelopes": 0},
    }
    for origin, root in (
        ("0819", ROOT / "docs/performance/results/change-0819/qualification"),
        ("0825", ROOT / "docs/performance/results/change-0825/qualification-before"),
    ):
        selected_cases = cases if origin == "0819" else [cases[0]]
        for case in selected_cases:
            report_path = root / f"00-{case['format']}-{case['phase']}.json"
            report = read(report_path)
            validate_qualification_report(report, case, qualification_selector(case))
            rows.append({
                "origin": origin,
                "case": case["case"],
                "report": artifact(report_path),
            })
            key = "0819/qualification" if origin == "0819" else "0825/failed-qualification"
            counts[key]["reports"] += 1
            counts[key]["sample_envelopes"] += 1
    return {
        "schema": "litchi.performance.0827.qualification-preflight.v1",
        "status": "pass",
        "counts": counts,
        "reports": len(rows),
        "sample_envelopes": len(rows),
        "rows": rows,
        "new_measurements": 0,
        "performance_claim": None,
    }


def qualification_admission(leg: str) -> dict[str, Any]:
    frozen_plan = plan()
    assert leg in {"before", "after"}
    artifact_path = P / frozen_plan["paths"]["artifact_admission"][leg]
    assert artifact_path.is_file() and not artifact_path.is_symlink(), (
        "artifact admission is required before qualification"
    )
    artifact_value = validate_artifact_admission_record(leg, frozen_plan, read(artifact_path))
    build_path, build, observer_descriptor = validate_build(leg, frozen_plan)
    qualification_root = P / frozen_plan["paths"]["qualification_output"][leg]
    complete_path = P / frozen_plan["paths"]["qualification_complete"][leg]
    receipts_path = qualification_root / "receipts.json"
    source_path = qualification_root / "source.json"
    assert qualification_root.is_dir() and not qualification_root.is_symlink()
    assert complete_path.is_file() and not complete_path.is_symlink()
    assert receipts_path.is_file() and not receipts_path.is_symlink()
    assert source_path.is_file() and not source_path.is_symlink()
    assert read(source_path) == read(P / f"build-{leg}" / "source.json")
    assert read(P / "freeze.json") == read(P / f"build-{leg}" / "frozen-inputs.json")
    validate_qualification_complete(
        leg, frozen_plan, build_path, source_path, receipts_path, complete_path
    )
    receipts = read(receipts_path)
    expected_list = plan_cases(frozen_plan)
    assert isinstance(receipts, list)
    assert len(receipts) == frozen_plan["expected"]["qualification_reports_per_leg"]
    admitted_by_case = {row["case"]: row for row in artifact_value["selectors"]}
    assert set(admitted_by_case) == {case["case"] for case in expected_list}
    rows: list[dict[str, Any]] = []
    for index, (receipt, case) in enumerate(zip(receipts, expected_list, strict=True)):
        assert isinstance(receipt, dict)
        assert receipt.get("case") == case["case"]
        report_path, _ = validate_qualification_receipt(
            receipt,
            index,
            leg,
            frozen_plan,
            case,
            admitted_by_case[case["case"]],
            qualification_root,
            source_path,
            artifact_path,
            observer_descriptor,
        )
        rows.append({"case": case["case"], "report": artifact(report_path), "receipt": receipt})

    expected_reports = frozen_plan["expected"]["qualification_reports_per_leg"]
    expected_samples = frozen_plan["expected"]["qualification_samples_per_leg"]
    target = P / f"qualification-admission-{leg}.json"
    result = {
        "schema": "litchi.performance.0827.qualification-admission.v1",
        "accepted": True,
        "leg": leg,
        "plan_sha256": sha(PLAN_PATH),
        "artifact_admission": artifact(artifact_path),
        "qualification_complete": artifact(complete_path),
        "qualification_receipts": artifact(receipts_path),
        "reports": expected_reports,
        "samples": expected_samples,
        "checks": {
            "artifact_gate_consumed": True,
            "all_three_real_formats": True,
            "all_four_phases_per_format": True,
            "source_oracles": True,
            "publication_oracles": True,
            "edit_outcome_oracles": True,
            "repeated_output_oracles": True,
            "timed_sample_cardinality": True,
        },
        "validation": {
            "artifact_descriptors": True,
            "qualification_complete_descriptors": True,
            "receipt_descriptors": True,
            "observer_report_schema": True,
            "zero_failed_allocation_policy": True,
            "exact_commands": True,
            "frozen_build_binding": True,
        },
        "rows": rows,
    }
    write_new(target, result)
    print(f"0827 qualification admission {leg} PASS")
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", nargs="?", choices=("artifacts", "qualification", "before", "after"))
    parser.add_argument("leg", nargs="?", choices=("before", "after"))
    args = parser.parse_args()
    if args.mode in {"before", "after"}:
        assert args.leg is None
        mode, leg = "artifacts", args.mode
    else:
        assert args.mode in {"artifacts", "qualification"} and args.leg in {"before", "after"}
        mode, leg = args.mode, args.leg
    if mode == "artifacts":
        artifact_admission(leg)
    else:
        qualification_admission(leg)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
