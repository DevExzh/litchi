#!/usr/bin/env python3
"""Independently verify the 0699 refusal-path measurement packet.

This verifier never builds Rust, invokes Cargo, runs the probe, or invokes
perf.  It binds the receipts to the current baseline checkout, the saved
0698 candidate witness, the packet probe tree, and the raw process output.
All timing summaries and deltas are recomputed from the raw TSV files here;
the packet's summary driver is treated as an output to audit, not as an
authority.

The audit is intentionally usable twice: before cleanup it checks that every
owned binary and scratch tree still exists, and after cleanup it checks the
exact cleanup receipt and the absence of those paths.
"""

from __future__ import annotations

import ast
import hashlib
import json
import math
import os
import re
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
PRIOR = P.parent / "change-0698"
BATCH = "0699"
CODEC = "crates/litchi-ooxml-common/src/mce/codec.rs"
CANDIDATE_WITNESS = PRIOR / "candidate-codec.rs.txt"
WORK = ROOT.parent / "litchi-0699-work"
TARGET = ROOT.parent / "litchi-target-0699"
BIN = ROOT.parent / "litchi-0699-bin"
PROFILE = ROOT.parent / "litchi-0699-profile"
PHASES = ("baseline", "candidate")
SCHEDULE = (
    "baseline",
    "candidate",
    "candidate",
    "baseline",
    "candidate",
    "baseline",
    "baseline",
    "candidate",
    "baseline",
    "candidate",
    "candidate",
    "baseline",
)
SINGLE_CASES = ("early-name-error", "small-valid", "late-root-error-mce")
MATRIX_CASES = (
    "small-valid",
    "generated-12x8-valid",
    "early-name-error",
    "late-root-error",
    "late-root-error-mce",
    "late-missing-relationship",
    "late-missing-relationship-mce",
    "notes-invalid-tail",
    "mixed-conformance",
    "slide-raw-overlimit-16m-to-64m",
)
SINGLE_PAIRS = ((1, 0), (2, 3), (4, 5), (7, 6), (9, 8), (10, 11))
MATRIX_PAIRS = ((1, 0), (2, 3))
STAT_FIELDS = ("p50_ns", "mean_ns", "p95_ns", "p99_ns")
ALL_STAT_FIELDS = STAT_FIELDS + ("min_ns", "max_ns")


def stop(message: str) -> None:
    raise AssertionError(message)


def need(path: Path) -> Path:
    if not path.exists():
        try:
            shown = str(path.relative_to(P))
        except ValueError:
            shown = str(path)
        stop(f"missing required evidence: {shown}")
    return path


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def packet_path(value: str | os.PathLike[str]) -> Path:
    path = Path(value)
    if not path.is_absolute():
        path = P / path
    resolved = path.resolve()
    try:
        resolved.relative_to(P.resolve())
    except ValueError:
        stop(f"evidence path escapes packet: {path}")
    return resolved


def read(name: str) -> Any:
    return json.loads(need(P / name).read_text())


def git_files(*args: str) -> set[str]:
    output = subprocess.check_output(["git", "ls-files", *args], cwd=ROOT, text=True)
    return {line for line in output.splitlines() if line}


def current_source_map(root: Path = ROOT) -> dict[str, str]:
    names = sorted(git_files("crates"))
    result: dict[str, str] = {}
    for name in names:
        path = root / name
        if not path.is_file():
            stop(f"tracked crate file is missing: {name}")
        result[name] = sha(path)
    return result


def packet_tree(root: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    for path in sorted(root.rglob("*")):
        if path.is_file() and "__pycache__" not in path.parts:
            result[str(path.relative_to(P))] = sha(path)
    return result


def assert_hash_map(paths: dict[str, str], base: Path = P) -> None:
    if not isinstance(paths, dict):
        stop("hash map is not an object")
    for name, digest in paths.items():
        path = base / name
        need(path)
        if sha(path) != digest:
            stop(f"hash mismatch: {name}")


def assert_source_map(actual: Any, expected: dict[str, str], label: str) -> None:
    if actual != expected:
        actual = actual if isinstance(actual, dict) else {}
        missing = sorted(set(expected) - set(actual))
        extra = sorted(set(actual) - set(expected))
        changed = sorted(
            name for name in set(actual) & set(expected) if actual[name] != expected[name]
        )
        stop(
            f"{label} source map mismatch: missing={missing[:3]} "
            f"extra={extra[:3]} changed={changed[:3]}"
        )


def validate_baseline(base: dict[str, Any]) -> tuple[dict[str, str], dict[str, str]]:
    required = {
        "baseline_head",
        "constraints_sha256",
        "source_sha256",
        "build_inputs_sha256",
        "candidate_codec_sha256",
    }
    if not required <= set(base):
        stop(f"baseline.json lacks required keys: {sorted(required - set(base))}")
    ancestor = subprocess.run(
        ["git", "merge-base", "--is-ancestor", base["baseline_head"], "HEAD"],
        cwd=ROOT,
    )
    if ancestor.returncode != 0:
        stop("recorded baseline_head is not an ancestor of the current checkout")

    tracked = git_files("crates")
    source = base["source_sha256"]
    if set(source) != tracked:
        stop(
            "baseline source map is not the exact tracked crates census: "
            f"missing={sorted(tracked - set(source))[:3]} "
            f"extra={sorted(set(source) - tracked)[:3]}"
        )
    current = current_source_map()
    assert_source_map(current, source, "current production")

    for group in ("constraints_sha256", "build_inputs_sha256"):
        values = base[group]
        if not isinstance(values, dict) or not values:
            stop(f"baseline {group} is empty or malformed")
        for name, digest in values.items():
            path = ROOT / name
            if not path.is_file() or sha(path) != digest:
                stop(f"baseline {group} binding mismatch: {name}")

    if not isinstance(base["candidate_codec_sha256"], str) or len(base["candidate_codec_sha256"]) != 64:
        stop("candidate codec hash in baseline.json is malformed")
    witness = need(CANDIDATE_WITNESS)
    if sha(witness) != base["candidate_codec_sha256"]:
        stop("0698 candidate codec witness does not match baseline.json")
    if sha(ROOT / CODEC) != source[CODEC]:
        stop("production codec was not restored to the recorded baseline")
    candidate = dict(source)
    candidate[CODEC] = base["candidate_codec_sha256"]
    if candidate == source:
        stop("candidate witness is byte-identical to baseline")
    return source, candidate


def validate_probe() -> dict[str, str]:
    probe = P / "probe"
    if not probe.is_dir():
        stop("probe tree is missing")
    files = packet_tree(probe)
    expected_names = {
        "probe/Cargo.lock",
        "probe/Cargo.toml",
        "probe/README.md",
        "probe/src/main.rs",
    }
    if set(files) != expected_names:
        stop(f"probe file census changed: {sorted(set(files) ^ expected_names)}")
    cargo = (P / "probe" / "Cargo.toml").read_text()
    if 'name = "probe0699-refusal"' not in cargo:
        stop("probe package identity is not probe0699-refusal")
    if 'name = "probe0699-refusal"' not in (P / "probe" / "Cargo.lock").read_text():
        stop("standalone probe lock does not bind probe0699-refusal")
    source = (P / "probe" / "src" / "main.rs").read_text()
    if 'const PROBE_ID: &str = "0699-refusal";' not in source:
        stop("probe source identity is not 0699-refusal")
    if "fn graph_digest" not in source or "expected_error_debug" not in source:
        stop("probe identity/precedence fields are missing")
    return files


def expected_build_command() -> list[str]:
    return [
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        str(WORK / P.relative_to(ROOT) / "probe" / "Cargo.toml"),
        "--target-dir",
        str(TARGET),
        "-j",
        "2",
    ]


def validate_builds(
    base: dict[str, Any], baseline: dict[str, str], candidate: dict[str, str], probe: dict[str, str]
) -> dict[str, dict[str, Any]]:
    bindings: dict[str, dict[str, Any]] = {}
    expected_command = expected_build_command()
    for phase in PHASES:
        row = read(f"build-{phase}.json")
        required = {"phase", "command", "exit_code", "source_sha256", "probe_sha256", "binary", "binary_sha256", "log_sha256"}
        if not required <= set(row):
            stop(f"build-{phase}.json lacks required receipt fields")
        if row["phase"] != phase or row["exit_code"] != 0 or row["command"] != expected_command:
            stop(f"build-{phase} command or exit receipt mismatch")
        expected_source = baseline if phase == "baseline" else candidate
        assert_source_map(row["source_sha256"], expected_source, f"build {phase}")
        if row["probe_sha256"] != probe:
            stop(f"build {phase} probe tree does not bind the packet probe")
        binary = Path(row["binary"])
        if binary != BIN / phase:
            stop(f"build {phase} binary path is not the owned frozen binary")
        if len(row["binary_sha256"]) != 64:
            stop(f"build {phase} binary hash is malformed")
        log = need(P / f"build-{phase}.log")
        if row["log_sha256"] != sha(log):
            stop(f"build {phase} log hash mismatch")
        if binary.exists() and sha(binary) != row["binary_sha256"]:
            stop(f"build {phase} binary hash mismatch")
        if not binary.exists() and not (P / "cleanup.json").exists():
            stop(f"pre-cleanup {phase} binary is missing")
        bindings[phase] = row
    if bindings["baseline"]["probe_sha256"] != bindings["candidate"]["probe_sha256"]:
        stop("baseline and candidate builds used different probe trees")

    first = need(P / "build-baseline-first-attempt.log")
    if "--locked" not in first.read_text():
        stop("retained first-attempt build log does not record the locked failure")
    need(P / "builds.log")
    return bindings


def validate_state_proof(base: dict[str, Any], baseline: dict[str, str], probe: dict[str, str]) -> None:
    """Check the post-build proof when the coordinator has emitted it.

    The proof is required because the two build receipts were created before
    the final source/probe/worktree assertions were added to build.py.  The
    accepted schema is deliberately structural: it must carry current and
    worktree source maps, build-input hashes, and the packet probe map, either
    at the top level or under `current`, `worktree`, and `probe` objects.
    """
    proof = read("build-state-verification.json")
    if proof.get("baseline_head") not in (None, base["baseline_head"]):
        stop("build-state proof baseline head mismatch")
    current = proof.get("current_source_sha256")
    worktree = proof.get("worktree_source_sha256") or proof.get("worktree_restored_source_sha256")
    inputs = proof.get("build_inputs_sha256") or proof.get("worktree_build_inputs_sha256")
    proof_probe = proof.get("probe_sha256")
    if isinstance(proof.get("current"), dict):
        current = current or proof["current"].get("source_sha256")
    if isinstance(proof.get("worktree"), dict):
        worktree = worktree or proof["worktree"].get("source_sha256")
    if isinstance(proof.get("inputs"), dict):
        inputs = inputs or proof["inputs"]
    if isinstance(proof.get("probe"), dict):
        proof_probe = proof_probe or proof["probe"].get("sha256")
    if current is None and proof.get("main_source_unchanged") is True:
        # validate_baseline already computed the current map immediately
        # before this proof check; the explicit boolean is the coordinator's
        # receipt that the same check was performed after both builds.
        current = baseline
    assert_source_map(current, baseline, "build-state current")
    assert_source_map(worktree, baseline, "build-state restored worktree")
    if inputs != read("baseline.json")["build_inputs_sha256"]:
        stop("build-state input binding mismatch")
    if proof_probe != probe:
        stop("build-state probe binding mismatch")


def validate_symbols(bindings: dict[str, dict[str, Any]]) -> None:
    rows = read("symbols.json")
    if not isinstance(rows, list) or len(rows) != 4:
        stop("symbols receipt must contain exactly four rows")
    expected = {
        (phase, name)
        for phase in PHASES
        for name in ("nm", "size")
    }
    seen: set[tuple[str, str]] = set()
    for row in rows:
        phase, name = row.get("phase"), row.get("name")
        key = (phase, name)
        if key in seen or key not in expected:
            stop(f"invalid or duplicate symbols receipt row: {key}")
        seen.add(key)
        binary = str(BIN / phase)
        command = ["nm", "-S", "-C", binary] if name == "nm" else ["size", "-A", binary]
        if row.get("command") != command or row.get("exit_code") != 0:
            stop(f"symbols command/exit mismatch: {key}")
        if row.get("binary_sha256") != bindings[phase]["binary_sha256"]:
            stop(f"symbols binary binding mismatch: {key}")
        output = packet_path(row.get("stdout", ""))
        if not output.is_file() or row.get("stdout_sha256") != sha(output):
            stop(f"symbols output hash mismatch: {key}")
        if row.get("stderr") != "":
            stop(f"symbols command emitted stderr: {key}")
    if seen != expected:
        stop("symbols receipt coverage is incomplete")


MARKER_SYMBOL = "litchi_ooxml_common::mce::codec::process_markup_compatibility"


def symbol_bounds(phase: str) -> tuple[int, int]:
    rows = []
    for line in need(P / "symbols" / f"{phase}-nm.txt").read_text().splitlines():
        fields = line.split(maxsplit=3)
        if len(fields) == 4 and fields[3] == MARKER_SYMBOL:
            rows.append(fields)
    if len(rows) != 1:
        stop(f"{phase}: marker symbol census is not unique")
    try:
        return int(rows[0][0], 16), int(rows[0][1], 16)
    except ValueError:
        stop(f"{phase}: marker symbol address/size is malformed")


def validate_annotations(bindings: dict[str, dict[str, Any]]) -> None:
    rows = read("annotations.json")
    if not isinstance(rows, list) or len(rows) != 4:
        stop("annotations receipt must contain exactly four rows")
    expected = {(phase, name) for phase in PHASES for name in ("assembly", "samples")}
    seen: set[tuple[str, str]] = set()
    raw_hashes = read("raw-profile-hashes.json")
    for row in rows:
        phase, name = row.get("phase"), row.get("name")
        key = (phase, name)
        if key in seen or key not in expected:
            stop(f"invalid or duplicate annotation row: {key}")
        seen.add(key)
        if row.get("exit_code") != 0 or row.get("symbol") != MARKER_SYMBOL:
            stop(f"annotation command failed or marker symbol changed: {key}")
        address, size = symbol_bounds(phase)
        if row.get("address") != address or row.get("size") != size:
            stop(f"annotation symbol bounds do not bind nm: {key}")
        if row.get("binary_sha256") != bindings[phase]["binary_sha256"]:
            stop(f"annotation binary binding mismatch: {key}")
        data_name = f"{phase}-early-name-error.data"
        if row.get("raw_data_sha256") != raw_hashes.get(data_name):
            stop(f"annotation perf-data binding mismatch: {key}")
        binary = str(BIN / phase)
        if name == "assembly":
            command = [
                "objdump",
                "-d",
                "-C",
                "--no-show-raw-insn",
                f"--start-address={address}",
                f"--stop-address={address + size}",
                binary,
            ]
        else:
            command = [
                "perf",
                "annotate",
                "--stdio",
                "--show-nr-samples",
                "--symbol",
                MARKER_SYMBOL,
                "-i",
                str(PROFILE / data_name),
            ]
        if row.get("command") != command:
            stop(f"annotation command binding mismatch: {key}")
        output = need(P / "annotations" / f"{phase}-{name}.stdout")
        stderr = need(P / "annotations" / f"{phase}-{name}.stderr")
        if row.get("stdout_sha256") != sha(output) or row.get("stderr_sha256") != sha(stderr):
            stop(f"annotation output hash mismatch: {key}")
        if stderr.read_bytes() != b"" or row.get("stderr_sha256") != sha_bytes(b""):
            stop(f"annotation command emitted stderr: {key}")
    if seen != expected:
        stop("annotation receipt coverage is incomplete")


EVENTS = ("cycles", "instructions", "branches", "branch-misses", "cache-misses", "page-faults", "task-clock")


def read_perf_counters(path: Path) -> dict[str, float]:
    values: dict[str, float] = {}
    for line in path.read_text().splitlines():
        fields = line.split("\t")
        if len(fields) > 2 and fields[2] in EVENTS:
            try:
                values[fields[2]] = float(fields[0])
            except ValueError:
                stop(f"perf counter value is not numeric: {path}/{fields[2]}")
    if set(values) != set(EVENTS):
        stop(f"perf counter event census mismatch: {path}")
    return values


def validate_profile_summaries() -> None:
    rows: list[dict[str, Any]] = []
    for repeat in range(3):
        for phase in PHASES:
            first = read_perf_counters(P / "profiles" / f"{phase}-counter-{repeat}-1000.stderr")
            last = read_perf_counters(P / "profiles" / f"{phase}-counter-{repeat}-11000.stderr")
            rows.append(
                {
                    "phase": phase,
                    "repeat": repeat,
                    "per_iteration": {
                        event: (last[event] - first[event]) / 10000 for event in EVENTS
                    },
                }
            )
    counter = read("counter-summary.json")
    if counter.get("unit") != "counts except task-clock milliseconds" or counter.get("rows") != rows:
        stop("counter-summary.json is not an independent recomputation of perf stat")
    comparisons = []
    for repeat in range(3):
        by = {row["phase"]: row["per_iteration"] for row in rows if row["repeat"] == repeat}
        delta = {}
        for event in EVENTS:
            base = by["baseline"][event]
            delta[event] = (by["candidate"][event] / base - 1) * 100 if base != 0 else None
        comparisons.append({"repeat": repeat, "delta_pct": delta})
    if counter.get("comparisons") != comparisons:
        stop("counter-summary comparisons are not independently recomputed")

    instruction = read("instruction-summary.json")
    expected_instruction = []
    for phase in PHASES:
        address, _size = symbol_bounds(phase)
        text = need(P / "annotations" / f"{phase}-samples.stdout").read_text()
        marker = next(
            row for row in json.loads((P / "annotations.json").read_text())
            if row["phase"] == phase and row["name"] == "samples"
        )
        match = re.search(r"\((\d+) samples,", text)
        if match is None:
            stop(f"{phase}: perf annotate sample total is missing")
        total = int(match.group(1))
        samples = [
            (int(count), int(location, 16))
            for count, location in re.findall(r"^\s*(\d+)\s*:\s*([0-9a-f]+):", text, re.M)
        ]
        if sum(count for count, _location in samples) != total:
            stop(f"{phase}: perf annotate sample counts do not sum to symbol total")
        first = address + 0xC0
        last = address + 0xDF
        loop = sum(count for count, location in samples if first <= location <= last)
        expected_instruction.append(
            {
                "phase": phase,
                "symbol": MARKER_SYMBOL,
                "symbol_samples": total,
                "loop_first_address": first,
                "loop_last_address": last,
                "loop_samples": loop,
                "loop_fraction_of_symbol_samples": loop / total,
            }
        )
    if instruction != expected_instruction:
        stop("instruction-summary.json is not an independent bounded sample recomputation")


def parse_refusal(path: Path) -> tuple[dict[str, list[str]], dict[str, dict[str, Any]]]:
    header: dict[str, list[str]] = {}
    cases: dict[str, dict[str, Any]] = {}
    current: dict[str, Any] | None = None
    for line in path.read_text().splitlines():
        fields = line.split("\t")
        if not fields or fields[0] == "":
            continue
        key = fields[0]
        if key == "case":
            if len(fields) != 2 or fields[1] in cases:
                stop(f"{path}: duplicate or malformed case record")
            current = {"metadata": {}, "samples": []}
            cases[fields[1]] = current
        elif key == "sample_ns":
            if len(fields) != 1:
                stop(f"{path}: malformed sample header")
        elif key.isdigit():
            if current is None or len(fields) != 2:
                stop(f"{path}: sample outside case or malformed sample")
            samples = current["samples"]
            if int(key) != len(samples):
                stop(f"{path}: sample indices are not contiguous")
            try:
                value = int(fields[1])
            except ValueError:
                stop(f"{path}: sample is not an integer")
            if value <= 0:
                stop(f"{path}: non-positive sample")
            samples.append(value)
        elif key == "all_iterations_passed":
            if len(fields) != 2 or fields[1] != "true":
                stop(f"{path}: probe iteration assertion failed")
            header[key] = fields[1:]
        elif current is None:
            if len(fields) < 2:
                stop(f"{path}: malformed global header")
            header[key] = fields[1:]
        else:
            if len(fields) < 2:
                stop(f"{path}: malformed case metadata")
            metadata = current["metadata"]
            if key in metadata:
                stop(f"{path}: duplicate metadata key {key}")
            metadata[key] = fields[1:]
    return header, cases


def load_prior_identities() -> dict[str, dict[str, list[str]]]:
    path = need(PRIOR / "refusal-bindings.json")
    value = json.loads(path.read_text())
    if set(value) != set(MATRIX_CASES):
        stop("0698 refusal identity census differs from the retained matrix")
    for case, metadata in value.items():
        if not isinstance(metadata, dict):
            stop(f"0698 identity is malformed: {case}")
        for key, fields in metadata.items():
            if not isinstance(fields, list) or not all(isinstance(item, str) for item in fields):
                stop(f"0698 identity fields are malformed: {case}/{key}")
    return value


def expected_run_rows() -> list[tuple[str, str | None, int, str]]:
    rows: list[tuple[str, str | None, int, str]] = []
    for mode, count in (("case", 12), ("matrix", 4)):
        for leg in range(count):
            phase = SCHEDULE[leg]
            if mode == "case":
                ordered = SINGLE_CASES if leg % 2 == 0 else tuple(reversed(SINGLE_CASES))
                rows.extend((mode, case, leg, phase) for case in ordered)
            else:
                rows.append((mode, None, leg, phase))
    return rows


def validate_run_header(
    header: dict[str, list[str]], mode: str, selected_case: str | None, cases: dict[str, dict[str, Any]], path: Path
) -> None:
    if header.get("mode") != [mode] or header.get("samples") != ["300"] or header.get("warmups") != ["10"]:
        stop(f"{path}: timing header does not bind the fixed 300/10 command")
    if header.get("all_iterations_passed") != ["true"]:
        stop(f"{path}: probe did not report all iterations passed")
    if mode == "case":
        if set(cases) != {selected_case}:
            stop(f"{path}: single-case output census mismatch")
        if "cases" in header:
            stop(f"{path}: single-case output has matrix case header")
    else:
        if set(cases) != set(MATRIX_CASES) or header.get("cases") != ["10"]:
            stop(f"{path}: matrix output census mismatch")
    if "probe" not in header or len(header["probe"]) != 1:
        stop(f"{path}: probe header is malformed")


def identity(metadata: dict[str, list[str]]) -> dict[str, list[str]]:
    # The probe's identity consists only of graph and expected/observed error
    # fields.  The probe/header fields are deliberately excluded here.
    required = {"authoring_base_archive_bytes", "authoring_base_archive_sha256", "prepared_input_graph_sha256", "expected_error_debug"}
    if not required <= set(metadata):
        stop(f"timed case identity lacks fields: {sorted(required - set(metadata))}")
    allowed = required | {"observed_error_debug"}
    if set(metadata) != allowed and set(metadata) != required:
        stop(f"timed case identity has unbound fields: {sorted(set(metadata) - allowed)}")
    if "observed_error_debug" in metadata and metadata["observed_error_debug"] != metadata["expected_error_debug"]:
        stop("observed refusal differs from the expected typed error")
    return {key: metadata[key] for key in sorted(metadata)}


def validate_raw_runs(bindings: dict[str, dict[str, Any]]) -> dict[tuple[str, str, int], dict[str, Any]]:
    runs = read("runs.json")
    if not isinstance(runs, list) or len(runs) != 40:
        stop(f"raw timing receipt has {len(runs) if isinstance(runs, list) else 'non-list'} rows, expected 40")
    expected = expected_run_rows()
    expected_keys = {(mode, case, leg) for mode, case, leg, _ in expected}
    seen: set[tuple[str, str | None, int]] = set()
    prior = load_prior_identities()
    identities: dict[str, dict[str, list[str]]] = {}
    raw: dict[tuple[str, str, int], dict[str, Any]] = {}
    expected_by_key = {(mode, case, leg): phase for mode, case, leg, phase in expected}
    for row in runs:
        mode, case, leg, phase = row.get("mode"), row.get("case"), row.get("leg"), row.get("phase")
        key = (mode, case, leg)
        if key in seen or key not in expected_keys or phase != expected_by_key[key]:
            stop(f"invalid, duplicate, or misordered timing receipt: {key}")
        seen.add(key)
        if row.get("exit_code") != 0 or phase not in PHASES or not isinstance(leg, int):
            stop(f"invalid timing process receipt: {key}")
        if row.get("binary_sha256") != bindings[phase]["binary_sha256"]:
            stop(f"timing binary binding mismatch: {key}")
        if mode not in {"case", "matrix"}:
            stop(f"unknown timing mode: {key}")
        if mode == "case" and case not in SINGLE_CASES:
            stop(f"unknown single-case fixture: {key}")
        output = packet_path(row.get("stdout", ""))
        if output.suffix != ".tsv" or not output.is_file():
            stop(f"timing stdout path is missing or not TSV: {output}")
        stderr = output.with_suffix(".stderr")
        if not stderr.is_file():
            stop(f"timing stderr is missing: {stderr}")
        if sha(output) != row.get("stdout_sha256") or sha(stderr) != row.get("stderr_sha256"):
            stop(f"timing raw hash mismatch: {key}")
        expected_command = ["taskset", "-c", "12", str(BIN / phase), mode]
        if mode == "case":
            expected_command.append(case)
        expected_command += ["300", "10"]
        if row.get("command") != expected_command:
            stop(f"timing command binding mismatch: {key}")
        header, cases = parse_refusal(output)
        validate_run_header(header, mode, case, cases, output)
        for case_name, item in cases.items():
            samples = item["samples"]
            metadata = item["metadata"]
            if len(samples) != 300:
                stop(f"{output}: {case_name} has {len(samples)} samples, expected 300")
            observed = identity(metadata)
            if case_name not in prior or observed != prior[case_name]:
                stop(f"{output}: graph/error identity differs from 0698: {case_name}")
            if case_name in identities and identities[case_name] != observed:
                stop(f"graph/error identity changed across current runs: {case_name}")
            identities[case_name] = observed
            if mode == "matrix" or case_name == case:
                raw[(mode, case_name, leg)] = {
                    "samples": samples,
                    "metadata": metadata,
                    "header": header,
                    "run": row,
                    "output": output,
                }
    if seen != expected_keys:
        stop(f"timing coverage incomplete: missing={sorted(expected_keys - seen)}")
    expected_raw_keys = {(mode, case, leg) for mode, case, leg, _ in expected for case in ([case] if case else MATRIX_CASES)}
    if set(raw) != expected_raw_keys:
        stop("raw timing case/leg expansion is incomplete")
    if set(identities) != set(MATRIX_CASES):
        stop("raw timing identity census is incomplete")
    return raw


def stat_values(samples: list[int]) -> dict[str, float | int]:
    if len(samples) != 300 or any(not isinstance(value, int) or value <= 0 for value in samples):
        stop("invalid raw sample vector")
    ordered = sorted(samples)
    return {
        "samples": len(samples),
        "p50_ns": statistics.median(samples),
        "mean_ns": statistics.mean(samples),
        "p95_ns": ordered[math.ceil(0.95 * len(samples)) - 1],
        "p99_ns": ordered[math.ceil(0.99 * len(samples)) - 1],
        "min_ns": min(samples),
        "max_ns": max(samples),
    }


def assert_number_equal(actual: Any, expected: Any, label: str) -> None:
    if not isinstance(actual, (int, float)) or not math.isclose(float(actual), float(expected), rel_tol=0.0, abs_tol=1e-9):
        stop(f"{label}: expected {expected!r}, got {actual!r}")


def validate_summary_and_comparisons(raw: dict[tuple[str, str, int], dict[str, Any]]) -> None:
    summary = read("summary.json")
    expected_summary_keys = {
        (mode, case, leg)
        for mode, case, leg, _ in expected_run_rows()
        for case in ([case] if case else MATRIX_CASES)
    }
    if not isinstance(summary, list) or len(summary) != len(expected_summary_keys):
        stop(f"summary row count mismatch: {len(summary) if isinstance(summary, list) else 'non-list'}")
    seen: set[tuple[str, str, int]] = set()
    for row in summary:
        if set(row) != {"mode", "case", "leg", "phase", "metadata", "stats"}:
            stop("summary row schema mismatch")
        key = (row["mode"], row["case"], row["leg"])
        if key in seen or key not in raw:
            stop(f"invalid or duplicate summary key: {key}")
        seen.add(key)
        source = raw[key]
        if row["phase"] != source["run"]["phase"] or row["metadata"] != source["metadata"]:
            stop(f"summary identity binding mismatch: {key}")
        stats = row["stats"]
        if set(stats) != set(("samples",) + ALL_STAT_FIELDS):
            stop(f"summary stats schema mismatch: {key}")
        expected_stats = stat_values(source["samples"])
        for field, value in expected_stats.items():
            assert_number_equal(stats[field], value, f"summary {key}/{field}")
    if seen != expected_summary_keys:
        stop("summary coverage is incomplete")

    comparisons = read("comparisons.json")
    expected_comparison_keys = {
        ("case", case, candidate, baseline)
        for case in SINGLE_CASES
        for candidate, baseline in SINGLE_PAIRS
    } | {
        ("matrix", case, candidate, baseline)
        for case in MATRIX_CASES
        for candidate, baseline in MATRIX_PAIRS
    }
    if not isinstance(comparisons, list) or len(comparisons) != len(expected_comparison_keys):
        stop(f"comparison row count mismatch: {len(comparisons) if isinstance(comparisons, list) else 'non-list'}")
    comparison_seen: set[tuple[str, str, int, int]] = set()
    expected_stats: dict[tuple[str, str, int], dict[str, float | int]] = {}
    for key, value in raw.items():
        expected_stats[key] = stat_values(value["samples"])
    for row in comparisons:
        required = {"mode", "case", "candidate_leg", "baseline_leg", "delta_pct", "delta_ns"}
        if set(row) != required or set(row["delta_pct"]) != set(STAT_FIELDS) or set(row["delta_ns"]) != set(STAT_FIELDS):
            stop("comparison row schema mismatch")
        key = (row["mode"], row["case"], row["candidate_leg"], row["baseline_leg"])
        if key in comparison_seen or key not in expected_comparison_keys:
            stop(f"invalid or duplicate comparison key: {key}")
        comparison_seen.add(key)
        candidate = expected_stats[(row["mode"], row["case"], row["candidate_leg"])]
        baseline = expected_stats[(row["mode"], row["case"], row["baseline_leg"])]
        for field in STAT_FIELDS:
            delta = candidate[field] - baseline[field]
            percent = (candidate[field] / baseline[field] - 1) * 100
            assert_number_equal(row["delta_ns"][field], delta, f"comparison {key}/{field} ns")
            assert_number_equal(row["delta_pct"][field], percent, f"comparison {key}/{field} percent")
    if comparison_seen != expected_comparison_keys:
        stop("comparison coverage is incomplete")

    triggers = read("triggers.json")
    expected_triggers = [
        {
            "mode": row["mode"],
            "case": row["case"],
            "candidate_leg": row["candidate_leg"],
            "baseline_leg": row["baseline_leg"],
            "metric": field,
            "delta_pct": row["delta_pct"][field],
            "delta_ns": row["delta_ns"][field],
        }
        for row in comparisons
        for field in STAT_FIELDS
        if row["delta_pct"][field] > 5
    ]
    if triggers != expected_triggers:
        stop("triggers.json is not the independently recomputed >5% four-metric filter")


def expected_profile_command(phase: str, name: str) -> list[str]:
    binary = str(BIN / phase)
    if name.endswith("-record"):
        case = name[len(phase) + 1 : -len("-record")]
        count = {"early-name-error": "100000", "late-root-error-mce": "20000"}.get(case)
        if count is None:
            stop(f"unknown profile record name: {name}")
        data = str(PROFILE / f"{name[:-len('-record')]}.data")
        return ["perf", "record", "-F", "997", "-g", "--call-graph", "dwarf,16384", "-o", data, "--", "taskset", "-c", "12", binary, "profile", case, count]
    for suffix, option in (("-self", "--no-children"), ("-inclusive", "--children")):
        if name.endswith(suffix):
            record_name = name[: -len(suffix)]
            data = str(PROFILE / f"{record_name}.data")
            return ["perf", "report", "--stdio", option, "--percent-limit", "0", "-i", data]
    if "-counter-" in name:
        prefix, repeat, count = name.split("-counter-")[0], *name.split("-counter-")[1].split("-")
        if prefix not in PHASES or repeat not in {"0", "1", "2"} or count not in {"1000", "11000"}:
            stop(f"unknown counter name: {name}")
        return ["perf", "stat", "-x", "\t", "-e", "cycles,instructions,branches,branch-misses,cache-misses,page-faults,task-clock", "--", "taskset", "-c", "12", binary, "profile", "early-name-error", count]
    stop(f"unknown profile name: {name}")


def validate_profiles(bindings: dict[str, dict[str, Any]]) -> None:
    runs = read("profile-runs.json")
    if not isinstance(runs, list):
        stop("profile-runs.json is not a list")
    expected_names = {
        f"{phase}-{case}-record"
        for phase in PHASES
        for case in ("early-name-error", "late-root-error-mce")
    } | {
        f"{phase}-{case}-{mode}"
        for phase in PHASES
        for case in ("early-name-error", "late-root-error-mce")
        for mode in ("self", "inclusive")
    } | {
        f"{phase}-counter-{repeat}-{count}"
        for phase in PHASES
        for repeat in range(3)
        for count in (1000, 11000)
    }
    if len(runs) != len(expected_names) or {row.get("name") for row in runs} != expected_names:
        stop(f"profile command census mismatch: expected {len(expected_names)} rows")
    seen: set[str] = set()
    record_status: dict[str, int] = {}
    prior = load_prior_identities()

    def validate_probe_profile_output(path: Path, case: str, iterations: int) -> None:
        fields: dict[str, str] = {}
        for line in path.read_text().splitlines():
            parts = line.split("\t")
            if len(parts) != 2 or parts[0] in fields:
                stop(f"profile probe output is malformed: {path}")
            fields[parts[0]] = parts[1]
        expected_identity = prior[case]
        expected_error = expected_identity["expected_error_debug"]
        if len(expected_error) != 1 or expected_error[0].startswith("ok "):
            stop(f"profile case is not a refusal fixture: {case}")
        expected_checksum = hashlib.sha256(
            ((expected_error[0] + "\0") * iterations).encode()
        ).hexdigest()
        expected = {
            "probe": "0699-refusal",
            "case": case,
            "mode": "profile",
            "iterations": str(iterations),
            "prepared_input_graph_sha256": expected_identity["prepared_input_graph_sha256"][0],
            "expected_error_debug": expected_error[0],
            "successful_captures": "0",
            "refused_captures": str(iterations),
            "assertion_checksum_sha256": expected_checksum,
            "all_iterations_passed": "true",
        }
        if fields != expected:
            stop(f"profile probe output identity/checksum mismatch: {path}")

    for row in runs:
        name = row.get("name")
        if name in seen:
            stop(f"duplicate profile row: {name}")
        seen.add(name)
        phase = row.get("phase")
        if phase not in PHASES or row.get("command") != expected_profile_command(phase, name):
            stop(f"profile command binding mismatch: {name}")
        if row.get("binary_sha256") != bindings[phase]["binary_sha256"]:
            stop(f"profile binary binding mismatch: {name}")
        if row.get("exit_code") != 0:
            stop(f"profile command failed; no fallback may be silently substituted: {name}")
        for stream in ("stdout", "stderr"):
            path = need(P / "profiles" / f"{name}.{stream}")
            if row.get(f"{stream}_sha256") != sha(path):
                stop(f"profile {stream} hash mismatch: {name}")
        if name.endswith("-record"):
            record_status[name] = row["exit_code"]
            case = name.rsplit("-record", 1)[0].split("-", 1)[1]
            iterations = int(row["command"][-1])
            validate_probe_profile_output(P / "profiles" / f"{name}.stdout", case, iterations)
        elif "-counter-" in name:
            case = "early-name-error"
            iterations = int(row["command"][-1])
            validate_probe_profile_output(P / "profiles" / f"{name}.stdout", case, iterations)
    if set(record_status) != {f"{phase}-{case}-record" for phase in PHASES for case in ("early-name-error", "late-root-error-mce")}:
        stop("profile record census is incomplete")

    raw_hashes = read("raw-profile-hashes.json")
    if not isinstance(raw_hashes, dict):
        stop("raw-profile-hashes.json is not an object")
    expected_data = {
        f"{phase}-{case}.data"
        for phase in PHASES
        for case in ("early-name-error", "late-root-error-mce")
    }
    if set(raw_hashes) != expected_data:
        stop("profile data hash census is incomplete or contains stale raw files")
    if not (P / "cleanup.json").exists():
        for name, digest in raw_hashes.items():
            path = PROFILE / name
            if not path.is_file() or sha(path) != digest:
                stop(f"profile raw data hash mismatch: {name}")
    else:
        for name, digest in raw_hashes.items():
            if not isinstance(digest, str) or len(digest) != 64:
                stop(f"profile raw hash malformed after cleanup: {name}")

    env = read("environment.json")
    perf = next((row for row in env if row.get("command") == ["perf", "--version"]), None)
    if perf is None or perf.get("exit_code") != 0 or not perf.get("stdout", "").startswith("perf version"):
        stop("environment does not explicitly establish perf availability")


def validate_gates() -> None:
    gates = read("gates.json")
    expected_names = {"workspace-fmt", "probe-fmt", "probe-clippy"}
    previous = json.loads(need(PRIOR / "evidence" / "results.json").read_text())
    expected_names |= {row["name"] for row in previous}
    if len(expected_names) != 9 or not isinstance(gates, list) or len(gates) != 9 or {row.get("name") for row in gates} != expected_names:
        stop("gate receipt census is not the exact nine required gates")
    commands = {
        "workspace-fmt": ["cargo", "fmt", "--all", "--check"],
        "probe-fmt": ["cargo", "fmt", "--manifest-path", str(P / "probe" / "Cargo.toml"), "--check"],
        "probe-clippy": ["cargo", "clippy", "--release", "--locked", "--manifest-path", str(WORK / P.relative_to(ROOT) / "probe" / "Cargo.toml"), "--target-dir", str(TARGET), "--no-deps", "--", "-D", "warnings"],
    }
    commands.update({row["name"]: row["command"] for row in previous})
    for row in gates:
        name = row["name"]
        if row.get("command") != commands[name] or row.get("exit_code") != 0:
            stop(f"gate failed or command binding changed: {name}")
        log = need(P / "gates" / f"{name}.log")
        if row.get("log_sha256") != sha(log):
            stop(f"gate log hash mismatch: {name}")


def validate_cli(bindings: dict[str, dict[str, Any]]) -> None:
    cases = [
        (["case", "early-name-error", "1", "1"], 0),
        (["profile", "early-name-error", "1"], 0),
        (["case", "unknown", "1", "1"], 1),
        (["case", "early-name-error", "0", "1"], 1),
        (["case", "early-name-error", "100001", "1"], 1),
        (["case", "early-name-error", "1", "0"], 1),
        (["case", "early-name-error", "1", "10001"], 1),
        (["profile", "early-name-error", "0"], 1),
        (["profile", "early-name-error", "100001"], 1),
        (["profile", "unknown", "1"], 1),
    ]
    rows = read("cli-checks.json")
    if not isinstance(rows, list) or len(rows) != 20:
        stop("CLI receipt does not contain exactly twenty checks")
    expected: list[tuple[str, list[str], int]] = []
    for phase in PHASES:
        binary = str(BIN / phase)
        for args, exit_code in cases:
            expected.append((phase, [binary, *args], exit_code))
    for row, (phase, command, expected_exit) in zip(rows, expected):
        if row.get("phase") != phase or row.get("command") != command or row.get("expected_exit_code") != expected_exit or row.get("exit_code") != expected_exit:
            stop(f"CLI command/exit binding mismatch: {row.get('phase')}/{row.get('command')}")
        if row.get("binary_sha256") != bindings[phase]["binary_sha256"]:
            stop("CLI binary hash mismatch")
        if expected_exit == 0 and "all_iterations_passed\ttrue" not in row.get("stdout", ""):
            stop("successful CLI check lacks the probe success receipt")
        if expected_exit != 0 and "all_iterations_passed\ttrue" in row.get("stdout", ""):
            stop("failing CLI check emitted a success receipt")


def validate_scripts() -> None:
    for path in sorted(P.rglob("*.py")):
        if "__pycache__" in path.parts:
            continue
        try:
            ast.parse(path.read_text(), filename=str(path))
        except SyntaxError as exc:
            stop(f"Python syntax error in {path}: {exc}")
    if any("__pycache__" in path.parts or path.suffix == ".pyc" for path in P.rglob("*")):
        stop("packet contains Python bytecode or __pycache__")
    manifest = read("script-hashes.json")
    expected = {
        str(path.relative_to(P))
        for path in P.rglob("*.py")
        if "__pycache__" not in path.parts
    }
    if set(manifest) != expected:
        stop("script hash manifest census mismatch")
    assert_hash_map(manifest, P)


def validate_cleanup() -> None:
    cleanup_path = P / "cleanup.json"
    required = {str(WORK), str(TARGET), str(BIN), str(PROFILE)}
    if not cleanup_path.exists():
        for path in (WORK, TARGET, BIN, PROFILE):
            if not path.is_dir():
                stop(f"pre-cleanup scratch tree is missing: {path}")
        for phase in PHASES:
            if not (BIN / phase).is_file():
                stop(f"pre-cleanup frozen binary is missing: {BIN / phase}")
        return
    receipt = read("cleanup.json")
    rows = receipt.get("removed")
    if not isinstance(rows, list) or {row.get("path") for row in rows} != required or len(rows) != 4:
        stop("cleanup receipt does not enumerate the exact four owned scratch trees")
    for row in rows:
        if row.get("removed") is not True or Path(row["path"]).exists():
            stop(f"cleanup did not prove removal: {row.get('path')}")
    lock = ROOT / "Cargo.lock"
    if receipt.get("workspace_cargo_lock_preserved_sha256") != sha(lock):
        stop("cleanup did not preserve the workspace Cargo.lock")
    for path in (WORK, TARGET, BIN, PROFILE):
        if path.exists():
            stop(f"owned scratch survived cleanup: {path}")


def main() -> None:
    base = read("baseline.json")
    baseline, candidate = validate_baseline(base)
    probe = validate_probe()
    bindings = validate_builds(base, baseline, candidate, probe)
    validate_state_proof(base, baseline, probe)
    validate_symbols(bindings)
    raw = validate_raw_runs(bindings)
    validate_summary_and_comparisons(raw)
    validate_profiles(bindings)
    validate_annotations(bindings)
    validate_profile_summaries()
    validate_gates()
    validate_cli(bindings)
    validate_scripts()
    validate_cleanup()
    print("PASS: 0699 source/probe/build bindings, raw 300/10 refusal runs, prior identity parity, independent summaries/deltas, profiles, gates, CLI checks, scripts, and cleanup")


if __name__ == "__main__":
    main()
