#!/usr/bin/env python3
"""Capture diagnostic XLS chain routes for the 0723 worksheet checkpoint.

This driver deliberately builds a stderr-instrumented binary.  Its output is
diagnostic route evidence only; it must never be mixed with the native timing
captures.  The production sources are restored byte-for-byte in ``finally``.
"""

from __future__ import annotations

import argparse
import difflib
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
BASELINE_REF = "45cb480eaa"
TARGET = Path("/home/zhuhe/code/litchi-target-0723-trace")
BINARY_ROOT = Path("/home/zhuhe/code/litchi-0723-trace-bin")
PROBE_MANIFEST = PACKET.parent / "change-0686" / "probe"
PROBE_CARGO = PROBE_MANIFEST / "Cargo.toml"
PROBE_NAME = "xls-index-retry-probe-0686"
SEQUENCE_MANIFEST = PACKET / "sequence-probe"
SEQUENCE_CARGO = SEQUENCE_MANIFEST / "Cargo.toml"
SEQUENCE_NAME = "xls-query-chain-sequence-0723"

# The candidate is deliberately confined to these XLS files.  shared.rs is
# instrumented too, because it is where the physical link count is observed.
TRACE_PATHS = (
    "crates/litchi-cfb/src/shared.rs",
    "crates/litchi-xls/src/workbook/query_cache.rs",
    "crates/litchi-xls/src/workbook/source.rs",
)
BASELINE_ONLY_TEST_PATHS = (
    "crates/litchi-xls/tests/xls_query_chain_checkpoint.rs",
)


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha_file(path: Path) -> str:
    return sha_bytes(path.read_bytes())


def source_hashes() -> dict[str, str]:
    return {
        relative: sha_file(ROOT / relative)
        for relative in TRACE_PATHS
    }


def probe_hashes() -> dict[str, str]:
    return {
        str(path.relative_to(ROOT)): sha_file(path)
        for path in sorted(PROBE_MANIFEST.rglob("*"))
        if path.is_file()
    }


def sequence_probe_hashes() -> dict[str, str]:
    return {
        str(path.relative_to(ROOT)): sha_file(path)
        for path in sorted(SEQUENCE_MANIFEST.rglob("*"))
        if path.is_file()
    }


def git_bytes(revision: str, relative: str) -> bytes:
    return subprocess.check_output(
        ["git", "show", f"{revision}:{relative}"],
        cwd=ROOT,
    )


def baseline_head() -> str:
    return subprocess.check_output(
        ["git", "rev-parse", BASELINE_REF], cwd=ROOT, text=True
    ).strip()


def load_cases(path: Path) -> list[dict[str, object]]:
    cases = json.loads(path.read_text())
    if not isinstance(cases, list) or not cases:
        raise ValueError(f"{path} must contain a non-empty case list")
    for case in cases:
        if not isinstance(case, dict):
            raise ValueError("trace cases must be objects")
        for field in ("case", "path", "budget", "row", "column", "sheet"):
            if field not in case:
                raise ValueError(f"trace case is missing {field!r}: {case!r}")
    return cases


def function_at(lines: list[str], index: int) -> str:
    """Return the nearest Rust function name before a source line."""

    function = "unknown"
    pattern = re.compile(r"^\s*(?:pub(?:\([^)]*\))?\s+)?(?:async\s+)?fn\s+([A-Za-z_][A-Za-z0-9_]*)\b")
    for line in lines[: index + 1]:
        match = pattern.match(line)
        if match:
            function = match.group(1)
    return function


def impl_at(lines: list[str], index: int) -> str:
    """Return the nearest simple Rust impl type before a source line."""

    implementation = "unknown"
    pattern = re.compile(r"^\s*impl(?:<[^>]*>)?\s+([A-Za-z_][A-Za-z0-9_]*)")
    for line in lines[: index + 1]:
        match = pattern.match(line)
        if match:
            implementation = match.group(1)
    return implementation


def route_for(path: str, lines: list[str], index: int) -> str:
    """Assign a stable diagnostic route to a CFB cursor call site."""

    function = function_at(lines, index)
    implementation = impl_at(lines, index)
    nearby = "\n".join(lines[max(0, index - 120) : min(len(lines), index + 25)]).lower()

    if path.endswith("query_cache.rs"):
        if "checkpoint" in nearby or "worksheet" in nearby:
            return "worksheet-checkpoint-build"
        return "query-cache"
    if function == "replay_indexed_cell":
        return "worksheet-replay"
    if function == "resolve_shared_string_inner":
        return "sst-entry"
    if function == "fill" and implementation == "SharedStringResolver":
        return "sst-fill"
    if function == "new" and implementation == "WorksheetScan":
        return "worksheet-scan"
    # A future implementation may put the post-scan exact-offset seek in a
    # helper.  Keep that extra walk visible without coupling this driver to a
    # private helper name.
    if "checkpoint" in nearby or "first_target" in nearby or "target_slot" in nearby:
        return "worksheet-checkpoint-build"
    return f"source-{function}"


def instrument_cursor_call(line: str, route: str, source_line: int) -> str:
    """Put the marker in the first call argument without breaking a chain."""

    needle = ".stream_cursor_at_hinted("
    if line.count(needle) != 1:
        raise RuntimeError("cursor call line is not unique")
    opening = line.index(needle) + len(needle)
    depth = 0
    closing = None
    for index in range(opening, len(line)):
        character = line[index]
        if character in "([{":
            depth += 1
        elif character in ")]}":
            depth -= 1
        elif character == "," and depth == 0:
            closing = index
            break
    if closing is None or not line[opening:closing].strip():
        raise RuntimeError("could not isolate stream cursor first argument")
    marker = f'eprintln!("TRACE call route={route} line={source_line}");'
    argument = line[opening:closing]
    return line[:opening] + "{ " + marker + " " + argument + " }" + line[closing:]


def inject_trace(source: bytes, relative: str) -> bytes:
    """Add route markers and physical chain-link diagnostics to one Rust file."""

    text = source.decode()
    lines = text.splitlines(keepends=True)
    if relative == "crates/litchi-cfb/src/shared.rs":
        joined = "".join(lines)
        start = joined.index("fn cursor_chain_sector(")
        end = joined.index("\n}\n\n#[inline]", start) + 2
        part = joined[start:end]
        marker = "    let (mut sector, walked) = resume;\n"
        if part.count(marker) != 1:
            raise RuntimeError("cursor_chain_sector marker is no longer unique")
        replacement = (
            "    eprintln!(\n"
            '        "TRACE chain table={table_name} from={} to={} links={}",\n'
            "        resume.1,\n"
            "        ordinal,\n"
            "        ordinal.saturating_sub(resume.1),\n"
            "    );\n"
            + marker
        )
        return (joined[:start] + part.replace(marker, replacement) + joined[end:]).encode()

    # Mark every actual source-side cursor construction.  The marker lives in
    # the first call argument, so a multiline method chain remains valid Rust.
    out: list[str] = []
    call_count = 0
    for index, line in enumerate(lines):
        if "stream_cursor_at_hinted(" in line:
            route = route_for(relative, lines, index)
            line = instrument_cursor_call(line, route, index + 1)
            call_count += 1
        out.append(line)
    if relative.endswith("source.rs"):
        joined = "".join(out)
        query = "    owner: &Arc<SourceInner>,\n    sheet_index: usize,\n    row: u32,\n    column: u32,\n"
        if joined.count(query) != 1:
            raise RuntimeError("query_cell signature is no longer unique")
        # Keep the query boundary in stderr so the analyzer can distinguish a
        # build walk from the later repeated replay calls.
        brace = joined.index(") -> Result<Option<SourceBackedCell>> {", joined.index("fn query_cell("))
        marker = (
            '    eprintln!("TRACE query sheet={} row={} column={}", sheet_index, row, column);\n'
        )
        joined = joined[: brace + len(") -> Result<Option<SourceBackedCell>> {") + 1] + marker + joined[brace + len(") -> Result<Option<SourceBackedCell>> {") + 1 :]
        out = [joined]
    if call_count == 0 and relative.endswith(("source.rs", "query_cache.rs")):
        raise RuntimeError(f"no stream cursor call found in {relative}")
    return "".join(out).encode()


def make_patch(before: dict[str, bytes], after: dict[str, bytes]) -> str:
    chunks: list[str] = []
    for relative in TRACE_PATHS:
        old = before[relative].decode().splitlines(keepends=True)
        new = after[relative].decode().splitlines(keepends=True)
        chunks.extend(
            difflib.unified_diff(
                old,
                new,
                fromfile=f"a/{relative}",
                tofile=f"b/{relative}",
            )
        )
    return "".join(chunks)


def scrub_probe_report(value: object) -> object:
    """Keep semantic probe output while excluding every timing field."""

    if isinstance(value, dict):
        return {
            key: scrub_probe_report(item)
            for key, item in value.items()
            if not any(token in key.lower() for token in ("elapsed", "nanos", "timing", "seconds"))
        }
    if isinstance(value, list):
        return [scrub_probe_report(item) for item in value]
    return value


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def run_phase(phase: str, cases_path: Path, baseline_head: str, queries: int, warmups: int) -> Path:
    if phase not in {"baseline", "candidate"}:
        raise ValueError("phase must be baseline or candidate")
    if queries < 1 or warmups < 1:
        raise ValueError("queries and warmups must be positive")

    out = PACKET / "trace" / phase
    out.mkdir(parents=True, exist_ok=True)
    original = {relative: (ROOT / relative).read_bytes() for relative in TRACE_PATHS}
    original_aux = {
        relative: (ROOT / relative).read_bytes()
        for relative in BASELINE_ONLY_TEST_PATHS
        if (ROOT / relative).is_file()
    }
    effective = dict(original)
    baseline_paths: list[str] = []
    manifest_path = out / "manifest.json"
    manifest: dict[str, object] | None = None

    try:
        if phase == "baseline":
            # The baseline is a real git source snapshot.  In particular this
            # prevents a candidate's private query-cache/source edits from
            # leaking into the before binary.  shared.rs is included if a
            # future candidate changes the CFB helper as well.
            for relative in TRACE_PATHS:
                baseline = git_bytes(baseline_head, relative)
                if effective[relative] != baseline:
                    effective[relative] = baseline
                    (ROOT / relative).write_bytes(baseline)
                    baseline_paths.append(relative)
            # The sequence-probe correctness test is candidate-only.  Keep it
            # out of the baseline build census and restore its exact bytes in
            # the same finally block as the instrumented production sources.
            for relative in BASELINE_ONLY_TEST_PATHS:
                path = ROOT / relative
                if path.is_file():
                    path.unlink()

        original_effective = dict(effective)
        instrumented = {
            relative: inject_trace(effective[relative], relative)
            for relative in TRACE_PATHS
        }
        for relative, value in instrumented.items():
            (ROOT / relative).write_bytes(value)
        (out / "instrumentation.patch").write_text(make_patch(original_effective, instrumented))

        source_before_build = source_hashes()
        command = [
            "cargo",
            "build",
            "--manifest-path",
            str(PROBE_CARGO.relative_to(ROOT)),
            "--release",
            "--locked",
            "--offline",
        ]
        build_log = out / "build.log"
        started = time.monotonic()
        with build_log.open("w") as log:
            result = subprocess.run(
                command,
                cwd=ROOT,
                env=dict(os.environ, CARGO_BUILD_JOBS="1", CARGO_TARGET_DIR=str(TARGET)),
                stdout=log,
                stderr=subprocess.STDOUT,
            )
        if result.returncode != 0:
            raise RuntimeError(f"instrumented probe build failed: {result.returncode}")
        if source_hashes() != source_before_build:
            raise RuntimeError("production source changed while trace binary was building")

        binary_dir = BINARY_ROOT / phase
        binary_dir.mkdir(parents=True, exist_ok=True)
        built_binary = TARGET / "release" / PROBE_NAME
        if not built_binary.is_file():
            raise RuntimeError(f"missing trace binary: {built_binary}")
        binary = binary_dir / PROBE_NAME
        shutil.copy2(built_binary, binary)

        commands: list[dict[str, object]] = [
            {
                "command": command,
                "cwd": str(ROOT),
                "target": str(TARGET),
                "exit_code": result.returncode,
                "seconds": round(time.monotonic() - started, 3),
                "binary": str(binary),
                "binary_sha256": sha_file(binary),
            }
        ]
        sequence_command = [
            "cargo",
            "build",
            "--manifest-path",
            str(SEQUENCE_CARGO.relative_to(ROOT)),
            "--release",
            "--locked",
            "--offline",
        ]
        sequence_build_log = out / "sequence-build.log"
        started = time.monotonic()
        with sequence_build_log.open("w") as log:
            sequence_result = subprocess.run(
                sequence_command,
                cwd=ROOT,
                env=dict(os.environ, CARGO_BUILD_JOBS="1", CARGO_TARGET_DIR=str(TARGET)),
                stdout=log,
                stderr=subprocess.STDOUT,
            )
        if sequence_result.returncode != 0:
            raise RuntimeError(f"sequence probe build failed: {sequence_result.returncode}")
        if source_hashes() != source_before_build:
            raise RuntimeError("production source changed while sequence probe was building")
        sequence_built = TARGET / "release" / SEQUENCE_NAME
        if not sequence_built.is_file():
            raise RuntimeError(f"missing sequence trace binary: {sequence_built}")
        sequence_binary = binary_dir / SEQUENCE_NAME
        shutil.copy2(sequence_built, sequence_binary)
        commands.append(
            {
                "command": sequence_command,
                "cwd": str(ROOT),
                "target": str(TARGET),
                "exit_code": sequence_result.returncode,
                "seconds": round(time.monotonic() - started, 3),
                "binary": str(sequence_binary),
                "binary_sha256": sha_file(sequence_binary),
            }
        )
        case_outputs: list[dict[str, object]] = []
        cases = load_cases(cases_path)
        for case in cases:
            name = str(case["case"])
            input_path = ROOT / str(case["path"])
            if not input_path.is_file():
                raise RuntimeError(f"missing trace input {input_path}")
            command = [
                str(binary),
                "--input",
                str(input_path.relative_to(ROOT)),
                "--budget",
                str(case["budget"]),
                "--mode",
                "owned",
                "--worksheet",
                str(case["sheet"]),
                "--row",
                str(case["row"]),
                "--column",
                str(case["column"]),
                "--queries",
                str(queries),
                "--samples",
                "1",
                "--warmups",
                str(warmups),
            ]
            started = time.monotonic()
            completed = subprocess.run(command, cwd=ROOT, capture_output=True, text=True)
            trace_path = out / f"{name}.trace"
            trace_path.write_text(completed.stderr)
            semantic_path = out / f"{name}.semantic.json"
            try:
                semantic = scrub_probe_report(json.loads(completed.stdout))
            except (json.JSONDecodeError, UnicodeDecodeError) as error:
                raise RuntimeError(f"probe did not emit JSON for {name}: {error}") from error
            write_json(semantic_path, semantic)
            if completed.returncode != 0:
                raise RuntimeError(
                    f"trace probe failed for {name}: {completed.returncode}\n{completed.stderr}"
                )
            commands.append(
                {
                    "case": name,
                    "command": command,
                    "cwd": str(ROOT),
                    "exit_code": completed.returncode,
                    "seconds": round(time.monotonic() - started, 3),
                    "trace": trace_path.name,
                    "semantic": semantic_path.name,
                }
            )
            case_outputs.append(
                {
                    "case": name,
                    "trace": trace_path.name,
                    "semantic": semantic_path.name,
                    "trace_sha256": sha_file(trace_path),
                    "semantic_sha256": sha_file(semantic_path),
                }
            )

        case_by_name = {str(case["case"]): case for case in cases}
        for required in ("54016-late", "54016-stored-2097152", "54016-missing-1048576"):
            if required not in case_by_name:
                raise RuntimeError(f"sequence route case is missing {required}")
        late = case_by_name["54016-late"]
        first = case_by_name["54016-stored-2097152"]
        missing = case_by_name["54016-missing-1048576"]
        sequence_command = [
            str(sequence_binary),
            "--input",
            str(ROOT / str(late["path"])),
            "--budget",
            str(late["budget"]),
            "--worksheet",
            str(late["sheet"]),
            "--late-row",
            str(late["row"]),
            "--late-column",
            str(late["column"]),
            "--first-row",
            str(first["row"]),
            "--first-column",
            str(first["column"]),
            "--missing-row",
            str(missing["row"]),
            "--missing-column",
            str(missing["column"]),
        ]
        started = time.monotonic()
        sequence_result = subprocess.run(
            sequence_command,
            cwd=ROOT,
            capture_output=True,
            text=True,
        )
        sequence_trace = out / "sequence.trace"
        sequence_semantic = out / "sequence.semantic.json"
        sequence_trace.write_text(sequence_result.stderr)
        if sequence_result.returncode != 0:
            raise RuntimeError(
                f"sequence trace probe failed: {sequence_result.returncode}\n{sequence_result.stderr}"
            )
        try:
            sequence_report = scrub_probe_report(json.loads(sequence_result.stdout))
        except (json.JSONDecodeError, UnicodeDecodeError) as error:
            raise RuntimeError(f"sequence probe did not emit JSON: {error}") from error
        write_json(sequence_semantic, sequence_report)
        commands.append(
            {
                "command": sequence_command,
                "cwd": str(ROOT),
                "exit_code": sequence_result.returncode,
                "seconds": round(time.monotonic() - started, 3),
                "trace": sequence_trace.name,
                "semantic": sequence_semantic.name,
            }
        )

        instrumented_hashes = source_hashes()
        manifest = {
            "schema_version": 1,
            "phase": phase,
            "baseline_head": baseline_head,
            "baseline_checkout_paths": baseline_paths,
            "baseline_excluded_test_paths": list(original_aux),
            "baseline_excluded_test_sha256": {
                relative: sha_bytes(value) for relative, value in original_aux.items()
            },
            "source_paths": list(TRACE_PATHS),
            "original_source_sha256": {
                relative: sha_bytes(original[relative]) for relative in TRACE_PATHS
            },
            "effective_source_sha256": {
                relative: sha_bytes(original_effective[relative]) for relative in TRACE_PATHS
            },
            "instrumented_source_sha256": {
                relative: instrumented_hashes[relative] for relative in TRACE_PATHS
            },
            "source_sha256_before_instrumentation": source_before_build,
            "probe_sha256": probe_hashes(),
            "sequence_probe_sha256": sequence_probe_hashes(),
            "cases": str(cases_path.relative_to(ROOT)),
            "cases_sha256": sha_file(cases_path),
            "target": str(TARGET),
            "binary": str(binary),
            "binary_sha256": sha_file(binary),
            "commands": commands,
            "outputs": case_outputs,
            "sequence_output": {
                "trace": sequence_trace.name,
                "semantic": sequence_semantic.name,
                "trace_sha256": sha_file(sequence_trace),
                "semantic_sha256": sha_file(sequence_semantic),
            },
            "sequence_binary": str(sequence_binary),
            "sequence_binary_sha256": sha_file(sequence_binary),
            "scope": "Diagnostic stderr routes only; instrumented timings are excluded from evidence.",
        }
        write_json(manifest_path, manifest)
    finally:
        for relative, value in original.items():
            (ROOT / relative).write_bytes(value)
        for relative, value in original_aux.items():
            (ROOT / relative).write_bytes(value)
        for relative in BASELINE_ONLY_TEST_PATHS:
            if relative not in original_aux:
                path = ROOT / relative
                if path.exists():
                    path.unlink()
        if any((ROOT / relative).read_bytes() != value for relative, value in original.items()):
            raise RuntimeError("trace driver failed to restore a production source byte")
        if any((ROOT / relative).read_bytes() != value for relative, value in original_aux.items()):
            raise RuntimeError("trace driver failed to restore a candidate test byte")
        if manifest is not None:
            manifest["restored_source_sha256"] = {
                relative: sha_file(ROOT / relative) for relative in TRACE_PATHS
            }
            manifest["restored"] = True
            write_json(manifest_path, manifest)

    return manifest_path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("phase", choices=("baseline", "candidate"))
    parser.add_argument("--baseline-head", default=baseline_head())
    parser.add_argument("--cases", type=Path, default=PACKET / "cases.json")
    parser.add_argument("--queries", type=int, default=4)
    parser.add_argument("--warmups", type=int, default=1)
    arguments = parser.parse_args()
    path = run_phase(
        arguments.phase,
        arguments.cases.resolve(),
        arguments.baseline_head,
        arguments.queries,
        arguments.warmups,
    )
    print(path)
    return 0


if __name__ == "__main__":
    sys.exit(main())
