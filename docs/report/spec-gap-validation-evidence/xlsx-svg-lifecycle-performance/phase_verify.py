#!/usr/bin/env python3
"""Verify exploratory same-drawing phase-decomposition receipts."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path
from typing import Any


SCHEMA = "xlsx-svg-lifecycle-phase-profile-v1"
PICTURE_COUNTS = (16, 64, 256)
PROCESSES = (1, 2, 3)
MIN_WARMUP = 2
MIN_SAMPLES = 20
PHASES = ("open", "stages", "commit", "firstsave", "reopen_secondsave", "validation")
RECEIPT_FIELDS = {
    "schema",
    "lane",
    "picture_count",
    "input_bytes",
    "input_hash_fnv1a64",
    "input_sha256",
    "warmup",
    "sample_count",
    "expected_success",
    "phase_order",
    "allocation_note",
    "semantic_checks",
    "samples",
}
SAMPLE_FIELDS = {"semantic_ok", "output_exact", "phases"}
SEMANTIC_CHECKS = [
    "drawing_picture_inventory",
    "anchor_and_raster_relationships",
    "unique_svg_relationships",
    "svg_media_bytes",
    "reopen_byte_identity",
    "graph_closure",
    "picture_closure",
    "opaque_fragments",
]
ALLOCATION_NOTE = (
    "requested_alloc_bytes is phase-local; live_after includes objects retained for later phases; "
    "phase values must not be summed or subtracted across unlike retained live sets"
)
COMPILER_FLAGS = "RUSTFLAGS=unset CARGO_ENCODED_RUSTFLAGS=unset RUSTC_BOOTSTRAP=unset"
ALLOCATOR = "CountingAllocator (process-local GlobalAlloc observer)"
PHASE_SCOPE = ",".join(PHASES)
PHASE_FIELDS = (
    "elapsed_ns",
    "requested_alloc_bytes",
    "direct_allocated_bytes",
    "realloc_new_bytes",
    "realloc_old_bytes",
    "deallocated_bytes",
    "live_before_bytes",
    "live_after_bytes",
    "live_delta_bytes",
    "retained_live_bytes_after",
    "peak_live_delta_bytes",
    "alloc_balance_ok",
    "alloc_invalid",
    "alloc_failed",
)
SHA256 = re.compile(r"^[0-9a-f]{64}$")
COMMIT = re.compile(r"^[0-9a-f]{40}$")
RECEIPT = re.compile(r"^phase_(?P<pictures>[0-9]+)-p(?P<process>[1-9][0-9]*)\.json$")
U64_MAX = 2**64 - 1


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def integer(
    value: Any,
    field: str,
    path: Path,
    *,
    nonnegative: bool = False,
    minimum: int | None = None,
    maximum: int | None = U64_MAX,
) -> int:
    require(type(value) is int, f"{field} must be an integer: {path}")
    if nonnegative:
        require(value >= 0, f"{field} must be nonnegative: {path}")
    if minimum is not None:
        require(value >= minimum, f"{field} is below its minimum: {path}")
    if maximum is not None:
        require(value <= maximum, f"{field} exceeds its maximum: {path}")
    return value


def read_rss(path: Path) -> int:
    marker = "Maximum resident set size (kbytes):"
    values = [
        integer(int(line.split(":", 1)[1].strip()), "RSS", path, nonnegative=True)
        for line in path.read_text().splitlines()
        if line.lstrip().startswith(marker)
    ]
    require(len(values) == 1, f"RSS marker must occur once: {path}")
    require(values[0] >= 0, f"RSS must be nonnegative: {path}")
    return values[0]


def verify_process_output(path: Path) -> int:
    stderr = path.with_suffix(".stderr.log")
    timing = path.with_suffix(".time.txt")
    require(stderr.is_file(), f"stderr receipt missing: {stderr}")
    require(stderr.read_bytes() == b"", f"stderr was not empty: {stderr}")
    require(timing.is_file(), f"time receipt missing: {timing}")
    statuses = [
        line.strip()
        for line in timing.read_text().splitlines()
        if line.strip().startswith("Exit status:")
    ]
    require(statuses == ["Exit status: 0"], f"process did not exit cleanly: {timing}")
    return read_rss(timing)


def verify_phase(value: Any, path: Path, phase: str) -> None:
    require(isinstance(value, dict), f"phase is not an object: {path} {phase}")
    require(set(value) == set(PHASE_FIELDS), f"phase fields changed: {path} {phase}")
    nonnegative_fields = {
        "elapsed_ns",
        "requested_alloc_bytes",
        "direct_allocated_bytes",
        "realloc_new_bytes",
        "realloc_old_bytes",
        "deallocated_bytes",
        "live_before_bytes",
        "live_after_bytes",
        "retained_live_bytes_after",
        "peak_live_delta_bytes",
        "alloc_failed",
    }
    for field in PHASE_FIELDS:
        if field in ("alloc_balance_ok", "alloc_invalid"):
            require(type(value[field]) is bool, f"{field} must be boolean: {path} {phase}")
        else:
            integer(
                value[field],
                field,
                path,
                nonnegative=field in nonnegative_fields,
                minimum=-U64_MAX if field == "live_delta_bytes" else None,
            )
    require(value["requested_alloc_bytes"] == value["direct_allocated_bytes"] + value["realloc_new_bytes"], f"allocation equation failed: {path} {phase}")
    require(value["live_delta_bytes"] == value["live_after_bytes"] - value["live_before_bytes"], f"live delta failed: {path} {phase}")
    require(value["retained_live_bytes_after"] == value["live_after_bytes"], f"retained live boundary failed: {path} {phase}")
    require(value["alloc_balance_ok"] is True, f"allocator balance failed: {path} {phase}")
    require(value["alloc_invalid"] is False, f"allocator underflow failed: {path} {phase}")
    require(value["alloc_failed"] == 0, f"allocator failure failed: {path} {phase}")
    expected_live = (
        value["live_before_bytes"]
        + value["direct_allocated_bytes"]
        + value["realloc_new_bytes"]
        - value["realloc_old_bytes"]
        - value["deallocated_bytes"]
    )
    require(value["live_after_bytes"] == expected_live, f"live equation failed: {path} {phase}")
    require(
        value["peak_live_delta_bytes"] >= max(value["live_delta_bytes"], 0),
        f"peak live delta is below the retained live increase: {path} {phase}",
    )


def receipt_paths(results: Path, pictures: int) -> list[Path]:
    paths = sorted(results.glob(f"phase_{pictures}-p*.json"), key=lambda path: path.name)
    expected = [results / f"phase_{pictures}-p{process}.json" for process in PROCESSES]
    require(paths == expected, f"expected exactly phase_{pictures}-p1..p3 receipts")
    return paths


def verify_receipt(path: Path, pictures: int, process: int) -> tuple[dict[str, Any], int]:
    payload = json.loads(path.read_text())
    require(set(payload) == RECEIPT_FIELDS, f"receipt fields changed: {path}")
    require(payload.get("schema") == SCHEMA, f"schema mismatch: {path}")
    require(payload.get("lane") == f"multi_picture_same_drawing_{pictures}", f"lane mismatch: {path}")
    require(type(payload.get("picture_count")) is int and payload["picture_count"] == pictures, f"picture count mismatch: {path}")
    require(payload.get("expected_success") is True, f"phase lane unexpectedly refused: {path}")
    integer(payload.get("input_bytes"), "input_bytes", path, nonnegative=True)
    integer(payload.get("input_hash_fnv1a64"), "input_hash_fnv1a64", path, nonnegative=True)
    input_sha256 = payload.get("input_sha256")
    require(isinstance(input_sha256, str) and SHA256.fullmatch(input_sha256), f"input SHA-256 malformed: {path}")
    warmup = integer(payload.get("warmup"), "warmup", path, nonnegative=True)
    require(warmup >= MIN_WARMUP, f"warmup below {MIN_WARMUP}: {path}")
    samples = payload.get("samples")
    require(isinstance(samples, list), f"samples missing: {path}")
    sample_count = integer(payload.get("sample_count"), "sample_count", path, nonnegative=True)
    require(sample_count == len(samples) and sample_count >= MIN_SAMPLES, f"sample count invalid: {path}")
    require(payload.get("phase_order") == list(PHASES), f"phase order changed: {path}")
    require(payload.get("semantic_checks") == SEMANTIC_CHECKS, f"semantic checks changed: {path}")
    require(payload.get("allocation_note") == ALLOCATION_NOTE, f"allocation note changed: {path}")
    for sample in samples:
        require(isinstance(sample, dict), f"sample malformed: {path}")
        require(set(sample) == SAMPLE_FIELDS, f"sample fields changed: {path}")
        require(sample.get("semantic_ok") is True, f"semantic gate failed: {path}")
        require(sample.get("output_exact") is True, f"output gate failed: {path}")
        phases = sample.get("phases")
        require(isinstance(phases, dict) and list(phases) == list(PHASES), f"sample phase set changed: {path}")
        for phase in PHASES:
            verify_phase(phases[phase], path, phase)
    rss = verify_process_output(path)
    return payload, rss


def _single_line_value(lines: list[str], prefix: str, path: Path) -> str:
    values = [line[len(prefix) :] for line in lines if line.startswith(prefix)]
    require(len(values) == 1, f"provenance field must occur once: {path} {prefix}")
    value = values[0].strip()
    require(value and value != "unavailable", f"provenance field is unavailable: {path} {prefix}")
    return value


def verify_build_provenance(results: Path) -> tuple[str, str, str]:
    path = results / "build-provenance.txt"
    require(path.is_file(), f"build provenance missing: {path}")
    lines = path.read_text().splitlines()
    binary = _single_line_value(lines, "binary=", path)
    _single_line_value(lines, "cargo=", path)
    _single_line_value(lines, "target=", path)
    cargo_incremental = _single_line_value(lines, "cargo_incremental=", path)
    require(cargo_incremental == "0", f"cargo incremental setting changed: {path}")
    compiler_flags = _single_line_value(lines, "compiler_flags=", path)
    require(compiler_flags == COMPILER_FLAGS, f"compiler flags changed: {path}")
    allocator = _single_line_value(lines, "allocator=", path)
    require(allocator == ALLOCATOR, f"allocator provenance changed: {path}")
    phase_scope = _single_line_value(lines, "phase_scope=", path)
    require(phase_scope == PHASE_SCOPE, f"phase scope changed: {path}")
    git_head = _single_line_value(lines, "git_head=", path)
    source_pin = _single_line_value(lines, "approved_source_pin=", path)
    for prefix in ("os=", "cpu_model=", "core_count=", "memory_total=", "storage=", "environment="):
        _single_line_value(lines, prefix, path)
    require("rustc -vV:" in lines, f"rustc provenance missing: {path}")
    require(any(line.startswith("rustc ") for line in lines), f"rustc version missing: {path}")
    require(COMMIT.fullmatch(git_head), f"Git head malformed: {path}")
    require(COMMIT.fullmatch(source_pin), f"source pin malformed: {path}")
    return git_head, source_pin, binary


def verify_binary_receipts(results: Path, binary: str) -> None:
    before_path = results / "binary.sha256"
    after_path = results / "binary-after.sha256"
    require(before_path.is_file() and after_path.is_file(), f"binary identity receipts missing: {results}")
    before = before_path.read_text()
    after = after_path.read_text()
    require(before == after, "binary identity changed")
    fields = before.strip().split(maxsplit=1)
    require(len(fields) == 2 and SHA256.fullmatch(fields[0]) and fields[1], f"binary SHA-256 receipt malformed: {results}")
    require(fields[1] == binary, f"binary path differs from build provenance: {results}")
    provenance = (results / "build-provenance.txt").read_text().splitlines()
    digest_lines = []
    for line in provenance:
        if not line.strip():
            continue
        candidate = line.strip().split(maxsplit=1)
        if len(candidate) == 2 and SHA256.fullmatch(candidate[0]):
            digest_lines.append(candidate)
    require(digest_lines and digest_lines[0] == fields, f"build binary digest differs from receipt: {results}")


def verify_metadata(results: Path) -> None:
    for name in ("metadata-before.json", "metadata-after.json"):
        path = results / name
        require(path.is_file(), f"Cargo metadata missing: {path}")
        value = json.loads(path.read_text())
        require(isinstance(value, dict), f"Cargo metadata is not an object: {path}")
        require(isinstance(value.get("packages"), list), f"Cargo metadata packages missing: {path}")


def verify_source_manifests(results: Path) -> str:
    manifests = []
    for name in ("source-manifest-before.txt", "source-manifest-after.txt"):
        path = results / name
        require(path.is_file(), f"source manifest missing: {path}")
        lines = path.read_text().splitlines()
        require(lines and lines[0] == "format=xlsx-svg-lifecycle-profile-build-source-v2", f"source manifest format changed: {path}")
        commits = [line.split("=", 1)[1] for line in lines if line.startswith("git_commit=")]
        metadata = [line.split("=", 1)[1] for line in lines if line.startswith("metadata_sha256=")]
        require(len(commits) == 1 and COMMIT.fullmatch(commits[0]), f"source manifest Git commit malformed: {path}")
        require(len(metadata) == 1 and SHA256.fullmatch(metadata[0]), f"source manifest metadata digest malformed: {path}")
        require(any(line.startswith("package=") for line in lines), f"source package manifest missing: {path}")
        require(any(line.startswith("file=") for line in lines), f"source file manifest missing: {path}")
        manifests.append((path, commits[0], metadata[0]))
    require(manifests[0][1:] == manifests[1][1:], "source manifest identity changed")
    for path, _, metadata_digest in manifests:
        metadata_name = "metadata-before.json" if path.name.endswith("before.txt") else "metadata-after.json"
        require(hashlib.sha256((results / metadata_name).read_bytes()).hexdigest() == metadata_digest, f"source manifest metadata hash mismatch: {path}")
    return manifests[0][1]


def verify_source_provenance(results: Path, git_head: str, source_pin: str) -> None:
    path = results / "source-provenance.txt"
    require(path.is_file(), f"source provenance missing: {path}")
    lines = path.read_text().splitlines()
    before = _single_line_value(lines, "source_manifest_before_sha256=", path)
    after = _single_line_value(lines, "source_manifest_after_sha256=", path)
    require(SHA256.fullmatch(before) and SHA256.fullmatch(after), f"source manifest provenance digest malformed: {path}")
    require(before == hashlib.sha256((results / "source-manifest-before.txt").read_bytes()).hexdigest(), f"source-before provenance mismatch: {path}")
    require(after == hashlib.sha256((results / "source-manifest-after.txt").read_bytes()).hexdigest(), f"source-after provenance mismatch: {path}")
    provenance_head = _single_line_value(lines, "git_head=", path)
    provenance_pin = _single_line_value(lines, "approved_source_pin=", path)
    require(COMMIT.fullmatch(provenance_head) and provenance_head == git_head, f"source provenance Git head mismatch: {path}")
    require(COMMIT.fullmatch(provenance_pin) and provenance_pin == source_pin, f"source provenance pin mismatch: {path}")


def verify_commands(results: Path, binary: str) -> None:
    path = results / "commands.txt"
    require(path.is_file(), f"command provenance missing: {path}")
    lines = path.read_text().splitlines()
    expected_shapes: set[tuple[int, int]] = set()
    for pictures in PICTURE_COUNTS:
        for process in PROCESSES:
            receipt = results / f"phase_{pictures}-p{process}.json"
            payload = json.loads(receipt.read_text())
            expected_shapes.add((payload["warmup"], payload["sample_count"]))
    require(len(expected_shapes) == 1, f"receipt command shape changed across processes: {path}")
    warmup, samples = expected_shapes.pop()
    expected = [
        f"run=/usr/bin/time -v {binary} --pictures {pictures} --warmup {warmup} --samples {samples} (fresh_process={process})"
        for pictures in PICTURE_COUNTS
        for process in PROCESSES
    ]
    require(lines == expected, f"command flags or process invocation changed: {path}")


def verify_results(results: Path, report: Path | None = None) -> dict[str, Any]:
    require(results.is_dir(), f"results directory missing: {results}")
    for path in results.glob("phase_*-p*.json"):
        match = RECEIPT.fullmatch(path.name)
        require(match is not None, f"malformed phase receipt name: {path}")
        require(int(match["pictures"]) in PICTURE_COUNTS, f"unknown phase picture count: {path}")
        require(int(match["process"]) in PROCESSES, f"unknown phase process identity: {path}")
    verified: list[dict[str, Any]] = []
    for pictures in PICTURE_COUNTS:
        paths = receipt_paths(results, pictures)
        identities = set()
        counts = set()
        for process, path in zip(PROCESSES, paths):
            payload, rss = verify_receipt(path, pictures, process)
            identities.add(
                (
                    payload["input_bytes"],
                    payload["input_hash_fnv1a64"],
                    payload["input_sha256"],
                )
            )
            counts.add(payload["sample_count"])
            verified.append({"pictures": pictures, "process": process, "payload": payload, "rss_kib": rss})
        require(len(identities) == 1, f"input identity changed across picture-{pictures} processes")
        require(len(counts) == 1, f"sample count changed across picture-{pictures} processes")
    verify_metadata(results)
    manifest_git_head = verify_source_manifests(results)
    require(
        (results / "source-manifest-before.txt").read_bytes()
        == (results / "source-manifest-after.txt").read_bytes(),
        "source manifest changed",
    )
    require((results / "build.log").is_file(), f"build log missing: {results / 'build.log'}")
    git_head, source_pin, binary = verify_build_provenance(results)
    require(manifest_git_head == git_head, "source manifest Git head differs from build provenance")
    verify_binary_receipts(results, binary)
    verify_source_provenance(results, git_head, source_pin)
    verify_commands(results, binary)
    before = (results / "source-manifest-before.txt").read_bytes()
    if report is not None:
        require(report.is_file(), f"phase report missing: {report}")
        from phase_summarize import summarize

        require(report.read_text() == summarize(results), f"phase report is not derived from raw receipts: {report}")
    return {
        "schema": "xlsx-svg-lifecycle-phase-verification-v1",
        "picture_counts": list(PICTURE_COUNTS),
        "processes_per_picture_count": len(PROCESSES),
        "minimum_warmup": MIN_WARMUP,
        "minimum_samples_per_process": MIN_SAMPLES,
        "receipts": len(verified),
        "source_manifest_sha256": hashlib.sha256(before).hexdigest(),
        "binary_identity_stable": True,
        "semantic_checks_passed": True,
        "raw_receipts_unchanged": True,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = verify_results(args.results.resolve(), args.report.resolve())
    args.output.write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
