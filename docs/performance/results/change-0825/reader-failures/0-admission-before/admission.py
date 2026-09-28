"""Root-owned admission for the 0825 ordinary-save validation packet.

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

POLICIES = ("default", "full", "file-only", "no-sync", "stream")
PHASES = {
    "lifecycle": "open+edit+save",
    "edit": "edit",
    "atomic_publish": "save-to-path",
    "counting_publish": "serialize-to-counting-sink",
}
EDIT_DIGEST = hashlib.sha256(b"admitted").hexdigest()

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


def plan() -> dict[str, Any]:
    value = read(PLAN_PATH)
    assert value["schema"] == "litchi.performance.0825.plan.v1"
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
    complete_path = P / plan()["paths"]["artifact_complete"][leg]
    complete = read(complete_path)
    assert complete.get("cases") == 6
    assert complete.get("reports", 0) == 0
    assert complete.get("samples", 0) == 0
    manifest_path = artifact_root / "manifest.json"
    assert manifest_path.is_file()
    if isinstance(complete.get("manifest"), dict):
        manifest_descriptor = complete["manifest"]
        manifest_value = Path(manifest_descriptor.get("path", ""))
        if not manifest_value.is_absolute():
            manifest_value = P / manifest_value
        assert manifest_value.resolve() == manifest_path.resolve()
        assert manifest_descriptor == artifact(manifest_path)
    return {"path": str(complete_path), "bytes": complete_path.stat().st_size, "sha256": sha(complete_path)}


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
        "schema": "litchi.performance.0825.artifact-audit-attempt.v1",
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
        "log": artifact(log_path),
    }
    write_new(attempt / "artifact-audit-receipt.json", receipt)
    assert launch_error is None, f"independent artifact audit could not start; retained {log_path}"
    assert completed is not None
    assert completed.returncode == 0, f"independent artifact audit failed; retained {log_path}"
    assert report_path.is_file() and not report_path.is_symlink(), report_path
    value = read(report_path)
    assert value.get("schema") == "litchi.performance.0825.artifact-audit.v1"
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
            "schema": "litchi.performance.0825.preservation-attempt.v1",
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
            "log": artifact(log_path),
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
    assert preservation.get("schema") == "litchi.performance.0825.zip-preservation.v1"
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
    assert not attempt.exists(), f"refusing to overwrite {attempt}"
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
        "schema": "litchi.performance.0825.artifact-admission.v1",
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
    print(f"0825 artifact admission {leg} PASS")
    return result


def descriptor_from_receipt(value: Any, expected: Path) -> dict[str, Any]:
    return verify_descriptor(value, expected)


def qualification_admission(leg: str) -> dict[str, Any]:
    value = plan()
    assert leg in {"before", "after"}
    artifact_path = P / value["paths"]["artifact_admission"][leg]
    assert artifact_path.is_file(), "artifact admission is required before qualification"
    artifact_value = read(artifact_path)
    assert artifact_value.get("accepted") is True and artifact_value.get("leg") == leg
    assert artifact_value.get("plan_sha256") == sha(PLAN_PATH)
    qualification_root = P / value["paths"]["qualification_output"][leg]
    complete_path = P / value["paths"]["qualification_complete"][leg]
    assert qualification_root.is_dir() and complete_path.is_file()
    complete = read(complete_path)
    assert complete.get("status") == "pass"
    assert complete.get("reports") == 12 and complete.get("samples") == 12
    receipts_path = qualification_root / "receipts.json"
    assert receipts_path.is_file()
    receipts = read(receipts_path)
    assert isinstance(receipts, list) and len(receipts) == 12
    expected_cases = {case["case"]: case for case in plan_cases(value)}
    seen: set[str] = set()
    rows: list[dict[str, Any]] = []
    admitted_by_case = {row["case"]: row for row in artifact_value["selectors"]}
    for receipt in receipts:
        assert receipt.get("exit_code") == 0
        if "policy" in receipt:
            assert receipt["policy"] == "default"
        case_name = receipt.get("case")
        assert isinstance(case_name, str) and case_name in expected_cases and case_name not in seen
        case = expected_cases[case_name]
        selector = admitted_by_case[case_name]
        report_descriptor = receipt.get("report")
        assert isinstance(report_descriptor, dict) and isinstance(report_descriptor.get("path"), str)
        report_path = Path(report_descriptor["path"])
        if not report_path.is_absolute():
            report_path = P / report_path
        report_path = report_path.resolve()
        assert qualification_root.resolve() in report_path.parents
        assert report_descriptor == artifact(report_path)
        report = read(report_path)
        config = report.get("configuration")
        assert isinstance(config, dict)
        assert config.get("samples_per_case") == 1
        assert config.get("warmup_iterations_per_case") == 0
        results = report.get("results")
        assert isinstance(results, list) and len(results) == 1
        result = results[0]
        assert result.get("case") == case_name
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
        if case["phase"] == "counting_publish":
            assert isinstance(ordinary.get("sample_byte_split"), dict)
        elapsed = result.get("elapsed_ns")
        assert isinstance(elapsed, dict) and isinstance(elapsed.get("samples"), list)
        assert len(elapsed["samples"]) == 1 and isinstance(elapsed["samples"][0], int)
        assert elapsed["samples"][0] > 0
        seen.add(case_name)
        rows.append({"case": case_name, "report": artifact(report_path), "receipt": receipt})
    assert seen == set(expected_cases)
    target = P / f"qualification-admission-{leg}.json"
    result = {
        "schema": "litchi.performance.0825.qualification-admission.v1",
        "accepted": True,
        "leg": leg,
        "plan_sha256": sha(PLAN_PATH),
        "artifact_admission": artifact(artifact_path),
        "qualification_complete": artifact(complete_path),
        "qualification_receipts": artifact(receipts_path),
        "reports": 12,
        "samples": 12,
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
        "rows": rows,
    }
    write_new(target, result)
    print(f"0825 qualification admission {leg} PASS")
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
