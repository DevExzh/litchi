#!/usr/bin/env python3
"""Run bounded corruption checks against the retained 0721 evidence.

The checks exercise the packet's Python validators with in-memory copies or
temporary JSON files.  They never alter a retained report, receipt, source
manifest, trace, binary, or analyzer output.  The intended invocation is once,
after the primary and read-control analyses have been finalized and before
the packet's terminal audit is sealed.
"""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile
from typing import Any, Callable, Iterable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
OUTPUT = HERE / "negative-checks.json"
HEX = set("0123456789abcdef")


class NegativeCheckError(AssertionError):
    """A preparation, mutation, or expected-rejection check failed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise NegativeCheckError(message)


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None,
            f"cannot load validator module: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


ANALYZE = load_module(HERE / "analyze-final.py", "change0721_negative_analyze")
READ_CONTROLS = load_module(HERE / "read-controls-analyze.py",
                            "change0721_negative_read_controls")
TRACE = load_module(HERE / "trace-analyze.py", "change0721_negative_trace")
AUDIT = load_module(HERE / "audit.py", "change0721_negative_audit")


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing input: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON input: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise NegativeCheckError(f"invalid JSON input {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(REPO))
    except ValueError:
        return str(path.resolve())


def input_binding(path: Path) -> dict[str, Any]:
    return {"path": relative(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def distinct_digest(value: str) -> str:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            "cannot mutate a non-SHA-256 digest")
    replacement = "0" if value[0] != "0" else "1"
    result = replacement + value[1:]
    require(result != value, "digest mutation did not change its value")
    return result


def report_result(path: Path) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict), f"{path.name}: report is not an object")
    results = value.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{path.name}: report result shape is not the expected single result")
    result = results[0]
    require(isinstance(result, dict), f"{path.name}: result is not an object")
    return result


def write_temp_json(value: Any, filename: str) -> tuple[tempfile.TemporaryDirectory[str], Path]:
    directory = tempfile.TemporaryDirectory(prefix="litchi-0721-negative-")
    path = Path(directory.name) / filename
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")
    return directory, path


def expect_rejection(
    name: str,
    target: Iterable[str],
    validator: Callable[[], Any],
    exception_types: tuple[type[BaseException], ...],
    message_check: Callable[[str], bool],
) -> dict[str, Any]:
    """Require a validator to reject the intended tamper for the intended reason."""

    try:
        validator()
    except exception_types as error:
        message = str(error)
        require(message_check(message),
                f"{name}: validator rejected for an unrelated reason: {message}")
        return {
            "name": name,
            "target": list(target),
            "rejected": True,
            "error_type": type(error).__name__,
            "error": message,
        }
    except BaseException as error:
        raise NegativeCheckError(
            f"{name}: unexpected exception type {type(error).__name__}: {error}"
        ) from error
    raise NegativeCheckError(f"{name}: validator accepted the tampered input")


def positive_call(name: str, validator: Callable[[], Any]) -> dict[str, Any]:
    try:
        validator()
    except BaseException as error:
        raise NegativeCheckError(
            f"{name}: retained input failed before tampering: "
            f"{type(error).__name__}: {error}"
        ) from error
    return {"name": name, "passed": True}


def with_patched_receipt(
    receipt_path: Path, value: dict[str, Any], validator: Callable[[], Any]
) -> Any:
    """Feed one temporary receipt to validate_receipt and restore the reader."""

    directory, temporary = write_temp_json(value, receipt_path.name)
    original_read = ANALYZE.read

    def read_override(path: Path) -> Any:
        if Path(path).resolve() == receipt_path.resolve():
            return read_json(temporary)
        return original_read(path)

    ANALYZE.read = read_override
    try:
        return validator()
    finally:
        ANALYZE.read = original_read
        directory.cleanup()


def run_python(command: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(command, cwd=REPO, check=False,
                          stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                          text=True)


def main() -> int:
    require(not OUTPUT.exists() and not OUTPUT.is_symlink(),
            f"refusing to replace existing output: {OUTPUT}")

    # These are the complete retained inputs read directly by this script or
    # by an imported validator.  Optional terminal custody files are included
    # when present because build/binary validation may consult them after
    # owned binaries have been cleaned up.
    fixed_inputs = [
        HERE / "negative-checks.py",
        HERE / "analyze-final.py",
        HERE / "read-controls-analyze.py",
        # These remain capture-time inputs.  Receipts intentionally bind the
        # original capture scripts even though terminal analysis uses the
        # corrected analysis-only copies above.
        HERE / "analyze.py",
        HERE / "read-controls.py",
        HERE / "trace-analyze.py",
        HERE / "audit.py",
        HERE / "pilot.py",
        HERE / "plan.json",
        HERE / "read-controls-plan.json",
        HERE / "source-baseline.json",
        HERE / "source-candidate.json",
        HERE / "build-baseline.json",
        HERE / "build-candidate.json",
        HERE / "analysis.json",
        HERE / "read-controls-analysis.json",
        HERE / "baseline-A1-native-generated-edit.json",
        HERE / "baseline-A1-native-generated-lifecycle.json",
        HERE / "baseline-A1-read-generated-medium-list-paragraphs.json",
        HERE / "trace" / "baseline" / "stderr",
        HERE / "trace" / "candidate" / "stderr",
    ]
    optional_inputs = [
        HERE / "source-final.json",
        HERE / "disposition.json",
        HERE / "cleanup.json",
        HERE / "cleanup-witness.json",
        HERE / "cleanup-receipt.json",
    ]
    input_paths = fixed_inputs + [path for path in optional_inputs if path.exists()]
    bindings_before = {relative(path): input_binding(path) for path in input_paths}

    primary_plan = ANALYZE.CAPTURE.load_plan()
    baseline_source, candidate_source, _changed = ANALYZE.CAPTURE.source_pair()
    generated = next(item for item in primary_plan["corpora"] if item["id"] == "generated")

    primary_edit_path = HERE / "baseline-A1-native-generated-edit.json"
    primary_edit_result = report_result(primary_edit_path)
    primary_edit_name = primary_edit_path.stem
    elapsed_original = copy.deepcopy(primary_edit_result["elapsed_ns"])
    elapsed_vector = copy.deepcopy(elapsed_original)
    elapsed_vector["samples"].pop()
    positive = [
        positive_call(
            "primary-elapsed-vector-retained",
            lambda: ANALYZE.LEGACY.validate_elapsed(
                elapsed_original, primary_plan["native"]["samples"], primary_edit_name
            ),
        )
    ]
    checks: list[dict[str, Any]] = []
    checks.append(expect_rejection(
        "primary-raw-elapsed-vector-length",
        [relative(primary_edit_path), "results[0].elapsed_ns.samples"],
        lambda: ANALYZE.LEGACY.validate_elapsed(
            elapsed_vector, primary_plan["native"]["samples"], primary_edit_name
        ),
        (AssertionError,),
        lambda message: "samples length changed" in message,
    ))

    stale_summary = copy.deepcopy(elapsed_original)
    stale_summary["mean"] = float(stale_summary["mean"]) + 1.0
    checks.append(expect_rejection(
        "primary-raw-elapsed-summary",
        [relative(primary_edit_path), "results[0].elapsed_ns.mean"],
        lambda: ANALYZE.LEGACY.validate_elapsed(
            stale_summary, primary_plan["native"]["samples"], primary_edit_name
        ),
        (AssertionError,),
        lambda message: "mean" in message,
    ))

    lifecycle_path = HERE / "baseline-A1-native-generated-lifecycle.json"
    lifecycle_result = report_result(lifecycle_path)
    lifecycle_result_mutated = copy.deepcopy(lifecycle_result)
    lifecycle_result_mutated["corpus"]["archive_bytes"] += 1
    positive.append(positive_call(
        "primary-semantic-identity-retained",
        lambda: ANALYZE.LEGACY.validate_ordinary_save(
            lifecycle_result, generated, "lifecycle", "native",
            primary_plan["native"]["samples"], lifecycle_path.stem,
        ),
    ))
    checks.append(expect_rejection(
        "primary-semantic-corpus-identity",
        [relative(lifecycle_path), "results[0].corpus.archive_bytes"],
        lambda: ANALYZE.LEGACY.validate_ordinary_save(
            lifecycle_result_mutated, generated, "lifecycle", "native",
            primary_plan["native"]["samples"], lifecycle_path.stem,
        ),
        (AssertionError,),
        lambda message: "archive" in message,
    ))

    read_plan = READ_CONTROLS.load_plan()
    read_baseline, read_candidate = READ_CONTROLS.load_sources(read_plan)
    control = read_plan["controls"][0]
    read_report_path = HERE / "baseline-A1-read-generated-medium-list-paragraphs.json"
    read_build = READ_CONTROLS.build_info(
        read_plan, "baseline", read_baseline, read_candidate, executable=False
    )
    positive.append(positive_call(
        "read-control-report-retained",
        lambda: READ_CONTROLS.check_report(
            read_plan, control, read_build, read_report_path,
        ),
    ))
    read_report = read_json(read_report_path)
    read_summary_mutated = copy.deepcopy(read_report)
    read_summary_mutated["results"][0]["elapsed_ns"]["p50"] += 1
    directory, read_summary_path = write_temp_json(
        read_summary_mutated, read_report_path.name,
    )
    try:
        checks.append(expect_rejection(
            "read-control-raw-elapsed-summary",
            [relative(read_report_path), "results[0].elapsed_ns.p50"],
            lambda: READ_CONTROLS.check_report(
                read_plan, control, read_build, read_summary_path,
            ),
            (RuntimeError,),
            lambda message: "summary" in message or "p50" in message,
        ))
    finally:
        directory.cleanup()

    read_identity_mutated = copy.deepcopy(read_report)
    read_identity_mutated["results"][0]["corpus"]["shape"] = "tampered"
    directory, read_identity_path = write_temp_json(
        read_identity_mutated, read_report_path.name,
    )
    try:
        checks.append(expect_rejection(
            "read-control-semantic-corpus-identity",
            [relative(read_report_path), "results[0].corpus.shape"],
            lambda: READ_CONTROLS.check_report(
                read_plan, control, read_build, read_identity_path,
            ),
            (RuntimeError,),
            lambda message: "corpus" in message,
        ))
    finally:
        directory.cleanup()

    baseline_trace = TRACE.parse_trace(HERE / "trace" / "baseline" / "stderr")
    candidate_trace = TRACE.parse_trace(HERE / "trace" / "candidate" / "stderr")
    positive.append(positive_call(
        "trace-differential-retained",
        lambda: TRACE.compare_trace_documents(
            copy.deepcopy(baseline_trace), copy.deepcopy(candidate_trace),
        ),
    ))

    mce_target: tuple[tuple[str, int], int] | None = None
    for document in candidate_trace["documents"]:
        calls = document["end"].get("active_calls", [])
        for index, call in enumerate(calls):
            offsets = call.get("input_offsets")
            if isinstance(offsets, list) and offsets:
                mce_target = (TRACE.document_key(document), index)
                break
        if mce_target is not None:
            break
    require(mce_target is not None,
            "trace input has no active MCE call with an input offset")
    mce_mutated = copy.deepcopy(candidate_trace)
    for document in mce_mutated["documents"]:
        if TRACE.document_key(document) == mce_target[0]:
            document["end"]["active_calls"][mce_target[1]]["input_offsets"][0] += 1
            break
    checks.append(expect_rejection(
        "trace-mce-input-vector",
        ["trace/candidate/stderr", "document_end.active_calls.input_offsets[0]"],
        lambda: TRACE.compare_trace_documents(
            copy.deepcopy(baseline_trace), mce_mutated,
        ),
        (TRACE.TraceError,),
        lambda message: "active_calls differs" in message,
    ))

    reader_target: tuple[str, int] | None = None
    for document in candidate_trace["documents"]:
        reads = document["end"].get("reader_reads")
        if isinstance(reads, list):
            reader_target = (document["case"] or "<uncased>", document["occurrence"])
            break
    require(reader_target is not None,
            "trace input has no reader scope to exercise the reader-proof check")
    reader_mutated = copy.deepcopy(candidate_trace)
    for document in reader_mutated["documents"]:
        key = (document["case"] or "<uncased>", document["occurrence"])
        if key == reader_target:
            document["end"].setdefault("reader_reads", []).append({
                "owner": "range", "start": 0, "end": 0,
            })
            break
    checks.append(expect_rejection(
        "trace-reader-proof-fake-range-owner",
        ["trace/candidate/stderr", "document_end.reader_reads"],
        lambda: TRACE.compare_trace_documents(
            copy.deepcopy(baseline_trace), reader_mutated,
        ),
        (TRACE.TraceError,),
        lambda message: "candidate performed" in message and "range reader" in message,
    ))

    build_candidate = read_json(HERE / "build-candidate.json")
    require(isinstance(build_candidate, list), "build-candidate.json is not a record list")
    native_record = next(
        (row for row in build_candidate
         if isinstance(row, dict) and row.get("lane") == "native"),
        None,
    )
    require(isinstance(native_record, dict), "candidate native build record is missing")
    native_binary = Path(str(native_record["binary"])).resolve()
    native_digest = native_record["binary_sha256"]
    native_bytes = native_record["binary_bytes"]
    witnesses = AUDIT.cleanup_witnesses()
    positive.append(positive_call(
        "binary-custody-retained",
        lambda: AUDIT.custody_path(
            native_binary, native_digest, native_bytes, witnesses,
            "candidate/native binary",
        ),
    ))
    checks.append(expect_rejection(
        "binary-receipt-sha256",
        [relative(HERE / "build-candidate.json"), "native.binary_sha256"],
        lambda: AUDIT.custody_path(
            native_binary, distinct_digest(native_digest), native_bytes, witnesses,
            "candidate/native binary",
        ),
        (AUDIT.AuditError,),
        lambda message: "hash changed" in message or "cleanup witness" in message,
    ))

    receipt_job = next(
        job for job in ANALYZE.child_jobs(primary_plan)
        if job["name"] == "baseline-A1-native-generated-edit"
    )
    receipt_build = ANALYZE.validate_builds("baseline", "native")
    receipt_path = HERE / f"{receipt_job['name']}.receipt.json"
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), "primary receipt is not an object")
    positive.append(positive_call(
        "primary-receipt-source-binding-retained",
        lambda: ANALYZE.validate_receipt(
            primary_plan, receipt_job, receipt_build,
            baseline_source, candidate_source,
            ANALYZE.sha(HERE / "plan.json"), ANALYZE.sha(HERE / "pilot.py"),
            ANALYZE.validate_constraints(),
        ),
    ))
    receipt_mutated = copy.deepcopy(receipt)
    receipt_mutated["build_source_manifest_sha256"] = distinct_digest(
        receipt["build_source_manifest_sha256"]
    )

    def tampered_receipt_validation() -> Any:
        return with_patched_receipt(
            receipt_path,
            receipt_mutated,
            lambda: ANALYZE.validate_receipt(
                primary_plan, receipt_job, receipt_build,
                baseline_source, candidate_source,
                ANALYZE.sha(HERE / "plan.json"), ANALYZE.sha(HERE / "pilot.py"),
                ANALYZE.validate_constraints(),
            ),
        )

    checks.append(expect_rejection(
        "primary-receipt-source-manifest-sha256",
        [relative(receipt_path), "build_source_manifest_sha256"],
        tampered_receipt_validation,
        (AssertionError,),
        lambda message: "binary source manifest" in message,
    ))

    positive_replay: list[dict[str, Any]] = []
    primary_replay = run_python([
        sys.executable, "-B", str(HERE / "analyze-final.py"), "--check",
        "--output", str(HERE / "analysis.json"),
    ])
    require(primary_replay.returncode == 0,
            f"primary analyzer exact replay failed: "
            f"{primary_replay.stderr.strip() or primary_replay.stdout.strip()}")
    positive_replay.append({
        "name": "primary-analysis-exact-replay",
        "passed": True,
        "output": relative(HERE / "analysis.json"),
    })
    with tempfile.TemporaryDirectory(prefix="litchi-0721-negative-replay-") as directory:
        read_output = Path(directory) / "read-controls-analysis.json"
        control_replay = run_python([
            sys.executable, "-B", str(HERE / "read-controls-analyze.py"), "analyze",
            "--output", str(read_output),
        ])
        require(control_replay.returncode == 0,
                f"read-control analyzer replay failed: "
                f"{control_replay.stderr.strip() or control_replay.stdout.strip()}")
        require(read_output.read_bytes()
                == (HERE / "read-controls-analysis.json").read_bytes(),
                "read-control analyzer replay bytes differ")
    positive_replay.append({
        "name": "read-control-analysis-exact-replay",
        "passed": True,
        "output": relative(HERE / "read-controls-analysis.json"),
    })

    bindings_after = {relative(path): input_binding(path) for path in input_paths}
    require(bindings_before == bindings_after,
            "retained input bytes changed while running negative checks")
    result = {
        "schema_version": 1,
        "packet": "change-0721-docx-negative-checks",
        "status": "pass",
        "inputs": bindings_before,
        "positive_replays": positive + positive_replay,
        "checks": checks,
        "retained_inputs_unchanged": True,
        "scope": {
            "native_or_profiler_invocations": 0,
            "temporary_mutation_files": True,
            "mutated_retained_artifacts": False,
            "check_count": len(checks),
        },
    }
    OUTPUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n",
                      encoding="utf-8")
    print(f"PASS: {len(checks)} bounded corruption checks rejected and retained inputs stayed unchanged")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (NegativeCheckError, AssertionError, KeyError, OSError, RuntimeError,
            TypeError, ValueError) as error:
        print(f"negative checks failed: {error}", file=sys.stderr)
        raise SystemExit(1)
