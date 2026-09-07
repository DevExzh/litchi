#!/usr/bin/env python3
"""Run bounded smoke checks for the 0464 PPTX pair-lifecycle binaries.

The script is deliberately separate from the capture harness.  It consumes
the retained normal and allocator executables, writes one report and one
published PPTX per provider/instrumentation lane below the retained bundle's
``smoke/`` directory, and records only immutable identities in
``smoke-results.json``.  Negative cases use a distinct temporary directory
below the caller-owned 0464 task root and are removed before the result is
written.
"""

from __future__ import annotations

import argparse
import copy
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import tempfile
from typing import Any


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
DEFAULT_OWNED_ROOT = Path("/tmp/litchi-goal-0464")
DEFAULT_MANIFEST = ROOT / "pair.json"
DEFAULT_BINDING = ROOT / "binding.json"
DEFAULT_RESULTS = ROOT / "smoke-results.json"
DEFAULT_SMOKE_DIR = ROOT / "smoke"
DEFAULT_NORMAL = DEFAULT_OWNED_ROOT / "normal"
DEFAULT_ALLOCATOR = DEFAULT_OWNED_ROOT / "allocator"

SCHEMA = "litchi-0464-pptx-pair-smoke-v1"
SMOKE_REPEAT = "smoke"
SMOKE_SAMPLES = 1
SMOKE_WARMUP = 0
TIMING_ALLOCATION_FIELDS = (
    "open_source_allocation_metrics",
    "open_destination_allocation_metrics",
    "plan_allocation_metrics",
    "publication_allocation_metrics",
)
PHASE_LABELS = (
    "baseline",
    "opened",
    "planned",
    "published",
    "drop_result",
    "drop_plan",
    "drop_view",
    "drop_caller_sources",
    "drop_sink",
)


class SmokeError(RuntimeError):
    """A smoke precondition or observed CLI contract failed."""


def fail(message: str) -> None:
    raise SmokeError(message)


def now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def require_regular(path: Path, label: str) -> Path:
    if path.is_symlink():
        fail(f"{label} must not be a symlink: {path}")
    if not path.is_file():
        fail(f"{label} is not a regular file: {path}")
    return path


def under(path: Path, parent: Path, label: str) -> Path:
    resolved = path.expanduser().resolve(strict=False)
    root = parent.expanduser().resolve(strict=False)
    try:
        resolved.relative_to(root)
    except ValueError as error:
        fail(f"{label} must be below {root}: {path}")
        raise AssertionError from error
    return resolved


def identity(path: Path, label: str) -> dict[str, Any]:
    require_regular(path, label)
    raw = path.read_bytes()
    return {
        "path": str(path.resolve()),
        "bytes": len(raw),
        "sha256": sha(raw),
    }


def json_identity(path: Path, label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    record = identity(path, label)
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        fail(f"{label} is not valid UTF-8 JSON: {error}")
    if not isinstance(value, dict):
        fail(f"{label} must contain a JSON object")
    return record, value


def bounded_text(raw: bytes) -> str:
    # Keep diagnostics useful without allowing a panic/backtrace to make the
    # immutable result unbounded.  The full stream identity remains recorded.
    text = raw.decode("utf-8", errors="replace")
    return text if len(text) <= 8192 else text[:8192] + "...[truncated]"


def diagnostics(result: subprocess.CompletedProcess[bytes]) -> dict[str, Any]:
    stdout = result.stdout or b""
    stderr = result.stderr or b""
    return {
        "stdout": {
            "bytes": len(stdout),
            "sha256": sha(stdout),
            "text": bounded_text(stdout),
        },
        "stderr": {
            "bytes": len(stderr),
            "sha256": sha(stderr),
            "text": bounded_text(stderr),
        },
    }


def invoke(argv: list[str], label: str) -> dict[str, Any]:
    started = now()
    try:
        result = subprocess.run(
            argv,
            cwd=REPO,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )
    except OSError as error:
        fail(f"{label} could not be started: {error}")
    return {
        "label": label,
        "argv": [str(value) for value in argv],
        "cwd": str(REPO),
        "started_utc": started,
        "finished_utc": now(),
        "exit_code": result.returncode,
        "diagnostics": diagnostics(result),
    }


def diagnostic_text(record: dict[str, Any]) -> str:
    streams = record["diagnostics"]
    return streams["stdout"]["text"] + "\n" + streams["stderr"]["text"]


def assert_failure(
    record: dict[str, Any], expected_fragments: tuple[str, ...], label: str
) -> None:
    if record["exit_code"] == 0:
        fail(f"{label} unexpectedly succeeded")
    text = diagnostic_text(record)
    if not any(fragment in text for fragment in expected_fragments):
        fail(
            f"{label} error did not contain one of {expected_fragments!r}; "
            f"diagnostics were {text!r}"
        )


def load_pair(path: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    manifest_identity, pair = json_identity(path, "pair manifest")
    for key in ("source", "destination", "operation", "limits", "source_revision"):
        if key not in pair:
            fail(f"pair manifest is missing {key!r}")
    for side in ("source", "destination"):
        value = pair[side]
        if not isinstance(value, dict):
            fail(f"pair manifest {side} must be an object")
        for key in ("path", "sha256", "bytes"):
            if key not in value:
                fail(f"pair manifest {side} is missing {key!r}")
        if len(value["sha256"]) != 64 or value["sha256"] != value["sha256"].lower():
            fail(f"pair manifest {side}.sha256 is not a lowercase digest")
    return manifest_identity, pair


def input_path(manifest: Path, value: str) -> Path:
    path = Path(value)
    return path.resolve() if path.is_absolute() else (manifest.parent / path).resolve()


def expect_pair_identity(pair: dict[str, Any], manifest: Path) -> dict[str, Any]:
    source = input_path(manifest, pair["source"]["path"])
    destination = input_path(manifest, pair["destination"]["path"])
    source_identity = identity(source, "pair source")
    destination_identity = identity(destination, "pair destination")
    for side, observed in (("source", source_identity), ("destination", destination_identity)):
        expected = pair[side]
        if observed["sha256"] != expected["sha256"] or observed["bytes"] != expected["bytes"]:
            fail(f"pair {side} file does not match its pinned identity")
    return {
        "source": source_identity,
        "destination": destination_identity,
    }


def binding_path(value: Any, label: str) -> Path:
    if not isinstance(value, str) or not value:
        fail(f"{label} must be a nonempty path string")
    path = Path(value).expanduser()
    return (ROOT / path).resolve(strict=False) if not path.is_absolute() else path.resolve(strict=False)


def validate_binding(
    path: Path,
    pair: dict[str, Any],
    normal: Path,
    allocator: Path,
    normal_identity: dict[str, Any],
    allocator_identity: dict[str, Any],
) -> tuple[dict[str, Any], dict[str, Any]]:
    binding_file, value = json_identity(path, "binary binding")
    if value.get("schema") != "litchi-0464-binaries-v1":
        fail("binary binding has an unexpected schema")
    if value.get("revision") != pair["source_revision"]:
        fail("binary binding revision differs from pair.json source_revision")
    binaries = value.get("binaries")
    if not isinstance(binaries, dict):
        fail("binary binding binaries is not an object")
    source_records: list[dict[str, Any]] = []
    summaries: dict[str, Any] = {}
    for name, actual_path, actual_identity in (
        ("normal", normal, normal_identity),
        ("allocator", allocator, allocator_identity),
    ):
        entry = binaries.get(name)
        if not isinstance(entry, dict):
            fail(f"binary binding is missing {name} entry")
        for key in (
            "path",
            "sha256",
            "bytes",
            "build_receipt",
            "build_receipt_sha256",
            "source",
        ):
            if key not in entry:
                fail(f"binary binding {name} is missing {key!r}")
        bound_path = binding_path(entry["path"], f"binary binding {name}.path")
        if bound_path != Path(actual_identity["path"]):
            fail(f"binary binding {name}.path differs from the retained binary")
        if entry["sha256"] != actual_identity["sha256"] or entry["bytes"] != actual_identity["bytes"]:
            fail(f"binary binding {name} identity differs from the retained binary")
        receipt_path = binding_path(entry["build_receipt"], f"binary binding {name}.build_receipt")
        receipt_identity, receipt = json_identity(receipt_path, f"{name} build receipt")
        if entry["build_receipt_sha256"] != receipt_identity["sha256"]:
            fail(f"binary binding {name} build receipt hash is stale")
        if receipt.get("status") != "pass" or receipt.get("source_unchanged") is not True:
            fail(f"{name} build receipt is not passing and source-unchanged")
        source_before = receipt.get("source_before")
        source_after = receipt.get("source_after")
        if not isinstance(source_before, dict) or source_before != source_after:
            fail(f"{name} build receipt source custody is not unchanged")
        if entry["source"] != source_before:
            fail(f"binary binding {name}.source differs from its build receipt")
        source_records.append(source_before)
        summaries[name] = {
            "path": actual_identity["path"],
            "sha256": actual_identity["sha256"],
            "bytes": actual_identity["bytes"],
            "build_receipt": receipt_identity,
            "build_receipt_sha256": entry["build_receipt_sha256"],
            "source": source_before,
        }
    if source_records[0] != source_records[1]:
        fail("normal and allocator binding source records differ")
    return binding_file, {
        "schema": value["schema"],
        "revision": value["revision"],
        "binaries": summaries,
        "source": source_records[0],
    }


def cli_argv(
    binary: Path,
    manifest: Path,
    pair: dict[str, Any],
    provider: str,
    report: Path,
    output: Path,
) -> list[str]:
    operation = pair["operation"]
    limits = pair["limits"]
    argv = [
        str(binary),
        "pptx-pair-lifecycle",
        "--manifest",
        str(manifest),
        "--provider",
        provider,
        "--samples",
        str(SMOKE_SAMPLES),
        "--warmup",
        str(SMOKE_WARMUP),
        "--repeat",
        SMOKE_REPEAT,
        "--source-revision",
        str(pair["source_revision"]),
        "--output",
        str(report),
        "--output-pptx",
        str(output),
    ]
    if provider == "range":
        argv.extend(
            [
                "--max-range",
                str(limits["max_range_bytes"]),
                "--delay-us",
                str(limits["delay_us"]),
            ]
        )
    # Keep this access explicit: a malformed manifest must fail before any
    # command is launched rather than silently selecting a different case.
    for key in ("source_slide", "destination_slide", "insertion_position"):
        if key not in operation:
            fail(f"pair operation is missing {key!r}")
    return argv


def validate_report(
    report: dict[str, Any],
    pair: dict[str, Any],
    pair_manifest_identity: dict[str, Any],
    binary_identity: dict[str, Any],
    provider: str,
    allocator: bool,
    input_identities: dict[str, Any],
) -> dict[str, Any]:
    expected_instrumentation = (
        "system_allocator_operation_scoped" if allocator else "none"
    )
    if report.get("schema") != "pptx_pair_lifecycle_v1":
        fail("positive report has an unexpected schema")
    for key, expected in (
        ("pair_id", pair["pair_id"]),
        ("provider", provider),
        ("repeat", SMOKE_REPEAT),
        ("samples", SMOKE_SAMPLES),
        ("warmup", SMOKE_WARMUP),
        ("checked_iteration_count", 1),
        ("source_revision", pair["source_revision"]),
        ("manifest_sha256", pair_manifest_identity["sha256"]),
        ("instrumentation", expected_instrumentation),
    ):
        if report.get(key) != expected:
            fail(f"positive report {key!r} differs from its smoke contract")
    if report.get("binary_sha256") != binary_identity["sha256"]:
        fail("positive report binary SHA-256 does not bind the retained binary")
    if report.get("binary_bytes") != binary_identity["bytes"]:
        fail("positive report binary byte count does not bind the retained binary")
    current_exe = report.get("current_exe")
    if not isinstance(current_exe, str) or Path(current_exe).resolve() != Path(binary_identity["path"]):
        fail("positive report current_exe does not bind the retained binary path")
    if allocator:
        if not isinstance(report.get("allocator_counter_revision"), str):
            fail("allocator report is missing allocator_counter_revision")
    elif "allocator_counter_revision" in report:
        fail("normal report unexpectedly contains allocator_counter_revision")
    for side in ("source", "destination"):
        value = report.get(side)
        expected = input_identities[side]
        if not isinstance(value, dict) or value.get("sha256") != expected["sha256"] or value.get("bytes") != expected["bytes"]:
            fail(f"positive report {side} identity differs from the pair input")
    operation = report.get("operation")
    expected_operation = {
        "source_slide": pair["operation"]["source_slide"],
        "destination_slide": pair["operation"]["destination_slide"],
        "insertion_position": pair["operation"]["insertion_position"],
    }
    if not isinstance(operation, dict) or any(
        operation.get(key) != value for key, value in expected_operation.items()
    ):
        fail("positive report operation differs from pair.json")
    phases = report.get("phases")
    if not isinstance(phases, list) or any(not isinstance(phase, dict) for phase in phases):
        fail("positive report phase descriptions are malformed")
    if [phase.get("label") for phase in phases] != list(PHASE_LABELS):
        fail("positive report phase descriptions differ from the CLI contract")
    rows = report.get("samples_raw")
    if not isinstance(rows, list) or len(rows) != 1:
        fail("positive report does not contain exactly one retained sample")
    row = rows[0]
    if not isinstance(row, dict):
        fail("positive sample row is not an object")
    timings = row.get("timings")
    if not isinstance(timings, dict):
        fail("positive sample row has no timings object")
    for field in TIMING_ALLOCATION_FIELDS:
        present = field in timings
        if allocator != present:
            fail(
                f"{('allocator' if allocator else 'normal')} report allocation field "
                f"{field!r} presence is {present}, expected {allocator}"
            )
        if allocator:
            sample = timings[field]
            if not isinstance(sample, dict) or sample.get("status") != "measured":
                fail(f"allocator report {field!r} is not measured")
    output_hash = row.get("output_sha256")
    if not isinstance(output_hash, str) or len(output_hash) != 64:
        fail("positive sample row has no valid output SHA-256")
    output_artifact = report.get("output_artifact")
    if not isinstance(output_artifact, dict) or output_artifact.get("sha256") != output_hash:
        fail("positive report output artifact does not match its sample")
    raw_row_phases = row.get("phases")
    if not isinstance(raw_row_phases, list) or any(
        not isinstance(phase, dict) for phase in raw_row_phases
    ):
        fail("positive sample phase records are malformed")
    row_phase_labels = [phase.get("label") for phase in raw_row_phases]
    if row_phase_labels != list(PHASE_LABELS):
        fail("positive sample phase records differ from the CLI contract")
    return {
        "instrumentation": report["instrumentation"],
        "allocation_fields": {
            field: field in timings for field in TIMING_ALLOCATION_FIELDS
        },
        "output_sha256": output_hash,
        "output_bytes": row.get("output_bytes"),
        "phase_labels": row_phase_labels,
    }


def run_positive(
    binary: Path,
    binary_identity: dict[str, Any],
    manifest: Path,
    pair: dict[str, Any],
    pair_manifest_identity: dict[str, Any],
    input_identities: dict[str, Any],
    provider: str,
    allocator: bool,
    smoke_dir: Path,
) -> tuple[dict[str, Any], bytes]:
    lane = f"{'allocator' if allocator else 'normal'}-{provider}"
    report_path = smoke_dir / f"{lane}.json"
    output_path = smoke_dir / f"{lane}.pptx"
    argv = cli_argv(binary, manifest, pair, provider, report_path, output_path)
    execution = invoke(argv, lane)
    if execution["exit_code"] != 0:
        fail(f"positive lane {lane} failed: {diagnostic_text(execution)!r}")
    report_file_identity, report = json_identity(report_path, f"{lane} report")
    output_file_identity = identity(output_path, f"{lane} output")
    details = validate_report(
        report,
        pair,
        pair_manifest_identity,
        binary_identity,
        provider,
        allocator,
        input_identities,
    )
    if output_file_identity["sha256"] != details["output_sha256"]:
        fail(f"{lane} output file differs from the report output hash")
    if output_file_identity["bytes"] != details["output_bytes"]:
        fail(f"{lane} output file differs from the report output byte count")
    execution.update(
        {
            "report": report_file_identity,
            "output": output_file_identity,
            "report_binding": {
                "binary_sha256": report["binary_sha256"],
                "binary_bytes": report["binary_bytes"],
                "source_revision": report["source_revision"],
                "manifest_sha256": report["manifest_sha256"],
            },
            "diagnostic_identity": execution["diagnostics"],
            "contract": details,
        }
    )
    return execution, output_path.read_bytes()


def write_manifest(path: Path, value: dict[str, Any]) -> None:
    raw = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    with path.open("xb") as output:
        output.write(raw)


def case_manifest(
    path: Path,
    pair: dict[str, Any],
    source: Path,
    destination: Path,
    mutate: Any,
) -> None:
    value = copy.deepcopy(pair)
    # A copied manifest lives in a temporary directory.  Absolute input paths
    # keep the negative case focused on its intended mutation.
    value["source"]["path"] = str(source)
    value["destination"]["path"] = str(destination)
    mutate(value)
    write_manifest(path, value)


def run_negative(
    binary: Path,
    pair: dict[str, Any],
    source: Path,
    destination: Path,
    owned_root: Path,
    name: str,
    mutate: Any,
    expected_fragments: tuple[str, ...],
    existing: str | None = None,
) -> dict[str, Any]:
    with tempfile.TemporaryDirectory(prefix=f"smoke-{name}-", dir=owned_root) as directory:
        temp = Path(directory)
        manifest = temp / "pair.json"
        case_manifest(manifest, pair, source, destination, mutate)
        report = temp / "report.json"
        output = temp / "output.pptx"
        preserved: dict[str, Any] | None = None
        if existing == "report":
            report.write_bytes(b"existing-report-sentinel\n")
            preserved = identity(report, f"{name} existing report before")
        elif existing == "output":
            output.write_bytes(b"existing-output-sentinel\n")
            preserved = identity(output, f"{name} existing output before")
        argv = cli_argv(binary, manifest, pair, "bytes", report, output)
        execution = invoke(argv, name)
        assert_failure(execution, expected_fragments, name)
        if existing == "report":
            after = identity(report, f"{name} existing report after")
            if after != preserved:
                fail(f"{name} overwrote the existing report")
        elif existing == "output":
            after = identity(output, f"{name} existing output after")
            if after != preserved:
                fail(f"{name} overwrote the existing output")
        record = {
            "name": name,
            "argv": execution["argv"],
            "exit_code": execution["exit_code"],
            "expected_error_fragments": list(expected_fragments),
            "diagnostics": execution["diagnostics"],
            "temporary_directory_removed": False,
        }
    record["temporary_directory_removed"] = not temp.exists()
    if not record["temporary_directory_removed"]:
        fail(f"{name} temporary directory was not removed")
    if preserved is not None:
        record["preserved_existing"] = preserved
    return record


def run_oversized_source_negative(
    binary: Path,
    pair: dict[str, Any],
    source: Path,
    destination: Path,
    owned_root: Path,
) -> dict[str, Any]:
    name = "oversized-source-metadata-bound"
    with tempfile.TemporaryDirectory(prefix=f"smoke-{name}-", dir=owned_root) as directory:
        temp = Path(directory)
        oversized = temp / "oversized-source.pptx"
        input_limit = max(int(pair["source"]["bytes"]), int(pair["destination"]["bytes"]))
        original = source.read_bytes()
        extra_bytes = input_limit + 1 - len(original)
        suffix = (b"oversized-source" * ((extra_bytes // 16) + 1))[:extra_bytes]
        oversized.write_bytes(original + suffix)
        oversized_identity = identity(oversized, f"{name} source")
        if oversized_identity["bytes"] <= input_limit:
            fail(f"{name} did not exceed its pinned metadata bound")
        manifest = temp / "pair.json"
        case_manifest(
            manifest,
            pair,
            oversized,
            destination,
            lambda value: value["limits"].update(input_bytes=input_limit),
        )
        report = temp / "report.json"
        output = temp / "output.pptx"
        argv = cli_argv(binary, manifest, pair, "bytes", report, output)
        execution = invoke(argv, name)
        assert_failure(
            execution,
            ("source input exceeds bounded read limit",),
            name,
        )
        record = {
            "name": name,
            "argv": execution["argv"],
            "exit_code": execution["exit_code"],
            "expected_error_fragments": ["source input exceeds bounded read limit"],
            "diagnostics": execution["diagnostics"],
            "pinned_source": {
                "bytes": pair["source"]["bytes"],
                "sha256": pair["source"]["sha256"],
            },
            "input_limit": input_limit,
            "oversized_source": oversized_identity,
            "temporary_directory_removed": False,
        }
    record["temporary_directory_removed"] = not temp.exists()
    if not record["temporary_directory_removed"]:
        fail(f"{name} temporary directory was not removed")
    return record


def receipt_identity(path: Path | None) -> dict[str, Any] | None:
    if path is None:
        return None
    record, value = json_identity(path, "build receipt")
    if value.get("status") != "pass":
        fail(f"build receipt is not passing: {path}")
    return {
        "file": record,
        "status": value.get("status"),
        "revision": value.get("revision"),
        "argv": value.get("argv"),
    }


def immutable_write(path: Path, value: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    raw = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    with path.open("xb") as output:
        output.write(raw)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--normal", type=Path, default=DEFAULT_NORMAL)
    parser.add_argument("--allocator", type=Path, default=DEFAULT_ALLOCATOR)
    parser.add_argument("--manifest", type=Path, default=DEFAULT_MANIFEST)
    parser.add_argument("--binding", type=Path, default=DEFAULT_BINDING)
    parser.add_argument("--owned-root", type=Path, default=DEFAULT_OWNED_ROOT)
    parser.add_argument("--smoke-dir", type=Path)
    parser.add_argument("--results", type=Path, default=DEFAULT_RESULTS)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    owned_root = args.owned_root.expanduser().resolve(strict=False)
    if owned_root != DEFAULT_OWNED_ROOT.resolve() and DEFAULT_OWNED_ROOT.resolve() not in owned_root.parents:
        fail("owned root must be /tmp/litchi-goal-0464 or one of its descendants")
    if owned_root.is_symlink():
        fail("owned root must not be a symlink")
    owned_root.mkdir(parents=True, exist_ok=True)
    manifest = args.manifest.expanduser().resolve()
    require_regular(manifest, "pair manifest")
    results = args.results.expanduser().resolve(strict=False)
    if results.exists():
        fail(f"smoke result already exists and is immutable: {results}")
    smoke_dir = (
        args.smoke_dir.expanduser().resolve(strict=False)
        if args.smoke_dir is not None
        else DEFAULT_SMOKE_DIR.resolve()
    )
    under(smoke_dir, ROOT, "smoke directory")
    if smoke_dir.exists():
        fail(f"smoke directory already exists and is immutable: {smoke_dir}")

    pair_manifest_identity, pair = load_pair(manifest)
    input_identities = expect_pair_identity(pair, manifest)
    normal = args.normal.expanduser().resolve()
    allocator = args.allocator.expanduser().resolve()
    normal_identity = identity(normal, "normal binary")
    allocator_identity = identity(allocator, "allocator binary")
    binding_file, binding = validate_binding(
        args.binding.expanduser().resolve(),
        pair,
        normal,
        allocator,
        normal_identity,
        allocator_identity,
    )
    # Delay the first mutation of the owned root until every immutable input
    # and optional build binding has passed validation.
    smoke_dir.mkdir(parents=True)

    started = now()
    positives: list[dict[str, Any]] = []
    output_bytes: bytes | None = None
    output_identity: dict[str, Any] | None = None
    for binary, binary_identity, allocator_mode in (
        (normal, normal_identity, False),
        (allocator, allocator_identity, True),
    ):
        for provider in ("bytes", "range"):
            record, published = run_positive(
                binary,
                binary_identity,
                manifest,
                pair,
                pair_manifest_identity,
                input_identities,
                provider,
                allocator_mode,
                smoke_dir,
            )
            if output_bytes is None:
                output_bytes = published
                output_identity = record["output"]
            elif published != output_bytes:
                fail(f"positive lane {record['label']} output is not deterministic")
            positives.append(record)
    if output_identity is None or output_bytes is None:
        fail("positive smoke produced no output")

    source = Path(input_identities["source"]["path"])
    destination = Path(input_identities["destination"]["path"])
    negatives = [
        run_oversized_source_negative(
            normal,
            pair,
            source,
            destination,
            owned_root,
        ),
        run_negative(
            normal,
            pair,
            source,
            destination,
            owned_root,
            "source-hash-mismatch",
            lambda value: value["source"].update(sha256="0" * 64),
            ("source SHA-256 differs from the pinned manifest",),
        ),
        run_negative(
            normal,
            pair,
            source,
            destination,
            owned_root,
            "out-of-range-slide-selector",
            lambda value: value["operation"].update(source_slide=999),
            ("pair operation selector is outside the input-derived presentation catalogs",),
        ),
        run_negative(
            normal,
            pair,
            source,
            destination,
            owned_root,
            "too-small-output-budget",
            lambda value: value["limits"].update(output_bytes=1),
            # The finalized CLI may reject while reserving the destination
            # execution budget (typed OutputBytes) before the bounded sink
            # receives its first write; older builds reached the sink guard.
            ("OutputBytes", "sequential sink byte budget exceeded"),
        ),
        run_negative(
            normal,
            pair,
            source,
            destination,
            owned_root,
            "tiny-memory-expansion-bound",
            lambda value: value["limits"].update(memory_bytes=1),
            ("resource: ArchiveMetadataBytes",),
        ),
        run_negative(
            normal,
            pair,
            source,
            destination,
            owned_root,
            "existing-report-refusal",
            lambda value: None,
            ("File exists",),
            existing="report",
        ),
        run_negative(
            normal,
            pair,
            source,
            destination,
            owned_root,
            "existing-output-refusal",
            lambda value: None,
            ("File exists",),
            existing="output",
        ),
    ]

    source_after = identity(source, "pair source after smoke")
    destination_after = identity(destination, "pair destination after smoke")
    if source_after != input_identities["source"] or destination_after != input_identities["destination"]:
        fail("smoke changed a retained pair input")
    try:
        revision = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=REPO, text=True
        ).strip()
    except (OSError, subprocess.CalledProcessError):
        revision = None

    result = {
        "schema": SCHEMA,
        "change": 464,
        "status": "pass",
        "started_utc": started,
        "finished_utc": now(),
        "manifest": pair_manifest_identity,
        "pair_id": pair["pair_id"],
        "source_revision": pair["source_revision"],
        "repository_revision": revision,
        "input_identities": input_identities,
        "build_binding": {
            "binding": binding_file,
            "value": binding,
        },
        "positive_lanes": positives,
        "deterministic_output": {
            "same_bytes_all_lanes": True,
            "bytes": output_identity["bytes"],
            "sha256": output_identity["sha256"],
            "lanes": [record["label"] for record in positives],
        },
        "negative_cases": negatives,
        "owned_root": str(owned_root),
        "smoke_directory": str(smoke_dir),
    }
    immutable_write(results, result)
    print(json.dumps({"status": "pass", "positive_lanes": 4, "negative_cases": 7}))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SmokeError as error:
        print(f"smoke failed: {error}", file=__import__("sys").stderr)
        raise SystemExit(1)
