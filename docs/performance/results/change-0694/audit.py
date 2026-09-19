#!/usr/bin/env python3
"""Verify the frozen 0694 evidence packet without rebuilding or measuring.

The packet deliberately keeps raw process output as the evidence boundary.
This audit checks the input/source/binary bindings in every receipt, parses the
raw native, allocation, refusal and oracle records, reruns only deterministic
summarizers, and verifies the final gate and cleanup receipts.
"""

from __future__ import annotations

import ast
import hashlib
import json
import math
import os
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OWNERS = ("litchi-pptx", "litchi-ooxml-common", "litchi-opc")
PHASES = {"baseline", "candidate"}
CASES = {row["case"]: row["workflow"] for row in json.loads((P / "cases.json").read_text())}
NATIVE_LEGS = {"a0", "a1", "a2", "a3", "b0", "b1"}
REFUSAL_CASES = {
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
}
ALLOC_PHASES = ("capture", "clone", "settext", "commit", "apply", "total")
ALLOC_METRICS = (
    "alloc_calls",
    "requested_bytes",
    "baseline_live_bytes",
    "peak_live_bytes",
    "current_live_bytes",
    "realloc_calls",
    "realloc_requested_bytes",
)


def stop(message: str) -> None:
    raise AssertionError(message)


def need(path: Path) -> Path:
    if not path.exists():
        stop(f"missing required packet file: {path.relative_to(P)}")
    return path


def read(name: str) -> Any:
    return json.loads(need(P / name).read_text())


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


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


def source_map() -> dict[str, str]:
    return {
        str(path.relative_to(ROOT)): sha(path)
        for owner in OWNERS
        for path in (ROOT / "crates" / owner).rglob("*.rs")
    }


def packet_tree(root: Path, *, names: set[str] | None = None) -> dict[str, str]:
    result = {}
    for path in root.rglob("*"):
        if not path.is_file():
            continue
        if names is not None and path.name not in names:
            continue
        result[str(path.relative_to(P))] = sha(path)
    return result


def assert_hash_map(paths: dict[str, str], base: Path = P) -> None:
    for name, digest in paths.items():
        path = base / name
        need(path)
        if sha(path) != digest:
            stop(f"hash mismatch for {name}")


def assert_source_map(actual: dict[str, str], expected: dict[str, str], label: str) -> None:
    if actual != expected:
        missing = sorted(set(expected) - set(actual))
        extra = sorted(set(actual) - set(expected))
        changed = sorted(name for name in set(actual) & set(expected) if actual[name] != expected[name])
        stop(f"{label} source map mismatch: missing={missing[:3]} extra={extra[:3]} changed={changed[:3]}")


def parse_tsv(path: Path) -> tuple[list[str], list[dict[str, int]], dict[str, list[str]]]:
    lines = path.read_text().splitlines()
    header_line = next((line for line in lines if line.startswith("sample\t")), None)
    if header_line is None:
        stop(f"{path}: missing sample header")
    header = header_line.split("\t")
    rows: list[dict[str, int]] = []
    metadata: dict[str, list[str]] = {}
    for line in lines:
        fields = line.split("\t")
        if not fields:
            continue
        if fields[0].isdigit():
            if len(fields) != len(header):
                stop(f"{path}: sample/header width mismatch")
            try:
                values = [int(value) for value in fields]
            except ValueError as exc:
                stop(f"{path}: non-integer sample value: {exc}")
            rows.append(dict(zip(header, values)))
        elif fields[0] not in {"sample", "sample_ns"}:
            metadata[fields[0]] = fields[1:]
    if not rows:
        stop(f"{path}: no sample rows")
    if [row["sample"] for row in rows] != list(range(len(rows))):
        stop(f"{path}: sample indices are not contiguous")
    return header[1:], rows, metadata


def target(metadata: dict[str, list[str]], key: str) -> tuple[int, int]:
    values = metadata.get(key)
    if values is None or len(values) != 2:
        stop(f"malformed {key} metadata: {values!r}")
    if not values[0].startswith("slide:") or not values[1].startswith("shape:"):
        stop(f"malformed {key} coordinates: {values!r}")
    return int(values[0][6:]), int(values[1][6:])


def validate_workflow(metadata: dict[str, list[str]], workflow: str, path: Path) -> None:
    if metadata.get("workflow", [workflow]) != [workflow]:
        stop(f"{path}: workflow metadata mismatch")
    if workflow == "one":
        target(metadata, "target")
        required = {
            "before_revision_sha256",
            "after_revision_sha256",
            "before_semantic_sha256",
            "after_semantic_sha256",
            "candidate_archive_sha256",
            "reopened_target_text_sha256",
            "correctness_target_text",
        }
        if missing := required - set(metadata):
            stop(f"{path}: one-edit metadata missing {sorted(missing)}")
        if metadata["before_revision_sha256"] == metadata["after_revision_sha256"]:
            stop(f"{path}: one-edit revision did not change")
        if metadata["before_semantic_sha256"] == metadata["after_semantic_sha256"]:
            stop(f"{path}: one-edit semantic digest did not change")
        if "target1" in metadata or "target2" in metadata:
            stop(f"{path}: one-edit has multi-target metadata")
    elif workflow == "noop":
        target(metadata, "target1")
        for key in ("commit_is_changed", "revision_identical", "output_identical"):
            expected = {"commit_is_changed": "false", "revision_identical": "true", "output_identical": "true"}[key]
            if metadata.get(key) != [expected]:
                stop(f"{path}: no-op {key} is {metadata.get(key)!r}")
        for key in ("before_revision_sha256", "after_revision_sha256", "before_archive_sha256", "after_archive_sha256", "before_semantic_sha256", "after_semantic_sha256", "correctness_target_text_sha256"):
            if key not in metadata:
                stop(f"{path}: no-op metadata missing {key}")
        for before, after in (("before_revision_sha256", "after_revision_sha256"), ("before_archive_sha256", "after_archive_sha256"), ("before_semantic_sha256", "after_semantic_sha256")):
            if metadata[before] != metadata[after]:
                stop(f"{path}: no-op {before}/{after} differ")
    elif workflow == "two":
        first = target(metadata, "target1")
        second = target(metadata, "target2")
        if first[0] == second[0]:
            stop(f"{path}: two-edit targets share a slide")
        required = {
            "before_revision_sha256",
            "after_revision_sha256",
            "before_semantic_sha256",
            "after_semantic_sha256",
            "candidate_archive_sha256",
            "reopened_target1_text_sha256",
            "reopened_target2_text_sha256",
            "correctness_target_text",
        }
        if missing := required - set(metadata):
            stop(f"{path}: two-edit metadata missing {sorted(missing)}")
        if metadata["before_revision_sha256"] == metadata["after_revision_sha256"]:
            stop(f"{path}: two-edit revision did not change")
        if metadata["before_semantic_sha256"] == metadata["after_semantic_sha256"]:
            stop(f"{path}: two-edit semantic digest did not change")
        if metadata["reopened_target1_text_sha256"] != metadata["reopened_target2_text_sha256"]:
            stop(f"{path}: two-edit reopened markers differ")
        if "target" in metadata:
            stop(f"{path}: two-edit has one-target metadata")
    else:
        stop(f"{path}: unknown workflow {workflow}")


def validate_control() -> None:
    manifest = read("control-manifest.json")
    source = ROOT / manifest["source"]
    control = P / manifest["control"]
    if sha(source) != manifest["source_sha256"]:
        stop("control archive hash mismatch")
    # cleanup.py intentionally removes the generated marker archive.  The
    # retained manifest/hash census is sufficient for a post-cleanup audit;
    # when the archive is still present, perform the stronger ZIP check below.
    control_retained = control.exists()
    if control_retained and sha(control) != manifest["control_sha256"]:
        stop("control archive hash mismatch")
    members = manifest["members"]
    if len(members) != 103 or sum(row["replacements"] for row in members) != 43:
        stop("control member/replacement census mismatch")
    old = b"http://schemas.openxmlformats.org/markup-compatibility/2006"
    new = b"http://schemas.openxmlformats.org/markup-kompatibility/2006"
    if control_retained:
        with zipfile.ZipFile(source) as before, zipfile.ZipFile(control) as after:
            if before.namelist() != after.namelist():
                stop("control ZIP member order differs")
            for row in members:
                before_bytes = before.read(row["member"])
                after_bytes = after.read(row["member"])
                if len(before_bytes) != len(after_bytes) or before_bytes.replace(old, new) != after_bytes:
                    stop(f"control replacement mismatch: {row['member']}")
                if before_bytes.count(old) != row["replacements"]:
                    stop(f"control replacement count mismatch: {row['member']}")
                if sha_bytes(before_bytes) != row["before_sha256"] or sha_bytes(after_bytes) != row["after_sha256"]:
                    stop(f"control member hash mismatch: {row['member']}")
    metadata = read("control-metadata-check.json")
    if metadata.get("archive_comment_equal") is not True:
        stop("control ZIP archive comment changed")
    if metadata.get("fields") != ["date_time", "compress_type", "create_system", "external_attr", "internal_attr", "extra", "comment", "create_version", "extract_version", "flag_bits"]:
        stop("control metadata field census changed")


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def validate_builds(base: dict[str, Any]) -> dict[tuple[str, str], dict[str, Any]]:
    bindings: dict[tuple[str, str], dict[str, Any]] = {}
    for phase in sorted(PHASES):
        rows = read(f"builds-{phase}.json")
        if not isinstance(rows, list) or {row.get("label") for row in rows} != {"native", "allocations"}:
            stop(f"invalid builds-{phase}.json labels")
        for row in rows:
            label = row["label"]
            if row.get("phase") != phase or row.get("exit_code") != 0 or row.get("rustflags") != "-D warnings":
                stop(f"invalid {phase}/{label} build receipt")
            if row.get("build_inputs_sha256") != base["build_inputs_sha256"]:
                stop(f"{phase}/{label}: build-input hash binding changed")
            assert_hash_map(row["probe_sha256"])
            if phase == "baseline":
                assert_source_map(row["source_sha256"], base["source_sha256"], f"baseline {label}")
            else:
                assert_source_map(row["source_sha256"], source_map(), f"candidate {label}")
            binary = Path(row["binary"])
            if binary.exists() and sha(binary) != row["binary_sha256"]:
                stop(f"{phase}/{label}: binary hash changed")
            bindings[phase, label] = row
    if bindings["baseline", "native"]["probe_sha256"] != bindings["candidate", "native"]["probe_sha256"]:
        stop("native probe changed between phases")
    if bindings["baseline", "allocations"]["probe_sha256"] != bindings["candidate", "allocations"]["probe_sha256"]:
        stop("allocation probe changed between phases")
    return bindings


def validate_refusal_builds(base: dict[str, Any], bindings: dict[tuple[str, str], dict[str, Any]]) -> None:
    for phase in sorted(PHASES):
        row = read(f"build-refusal-{phase}.json")
        if row.get("phase") != phase or row.get("exit_code") != 0:
            stop(f"invalid refusal build receipt for {phase}")
        expected_source = base["source_sha256"] if phase == "baseline" else source_map()
        assert_source_map(row["source_sha256"], expected_source, f"refusal {phase}")
        assert_hash_map(row["probe_sha256"])
        binary = Path(row["binary"])
        if binary.exists() and sha(binary) != row["binary_sha256"]:
            stop(f"refusal {phase}: binary hash changed")
    if read("build-refusal-baseline.json")["probe_sha256"] != read("build-refusal-candidate.json")["probe_sha256"]:
        stop("refusal probe changed between phases")


def validate_native(bindings: dict[tuple[str, str], dict[str, Any]]) -> None:
    records = read("native-runs-baseline.json") + read("native-runs-compare.json")
    if len(records) != len(CASES) * len(NATIVE_LEGS):
        stop(f"native manifest has {len(records)} rows, expected {len(CASES) * len(NATIVE_LEGS)}")
    seen: set[tuple[str, str]] = set()
    controls = read("control-manifest.json")
    for row in records:
        key = (row.get("case"), row.get("leg"))
        if key in seen or row["case"] not in CASES or row["leg"] not in NATIVE_LEGS:
            stop(f"invalid or duplicate native row {key}")
        seen.add(key)
        phase = "baseline" if row["leg"].startswith("a") else "candidate"
        if row.get("phase") != phase or row.get("workflow") != CASES[row["case"]] or row.get("exit_code") != 0:
            stop(f"native phase/workflow binding mismatch: {row}")
        if row.get("binary_sha256") != bindings[phase, "native"]["binary_sha256"]:
            stop(f"native binary binding mismatch: {key}")
        output = packet_path(row["output"])
        if sha(output) != row["output_sha256"]:
            stop(f"native output hash mismatch: {output}")
        stderr = output.with_suffix(".stderr")
        if sha(stderr) != row["stderr_sha256"]:
            stop(f"native stderr hash mismatch: {stderr}")
        command = row.get("command", [])
        if len(command) < 9 or command[4] != "phases" or command[-1] != CASES[row["case"]]:
            stop(f"native command binding mismatch: {key}")
        source_arg = command[5]
        if row.get("source_sha256") is None:
            if not str(source_arg).startswith("generated:"):
                stop(f"non-generated native case lacks source hash: {key}")
        else:
            source = Path(source_arg)
            control_removed = not (P / "marker-control.pptx").exists() and (P / "cleanup.json").exists()
            if not source.exists() and not (control_removed and row["case"].endswith("-control")):
                stop(f"native input archive missing: {key}")
            if source.exists() and sha(source) != row["source_sha256"]:
                stop(f"native input archive mismatch: {key}")
            if row["case"].endswith("-control") and row["source_sha256"] != controls["control_sha256"]:
                stop(f"control native input binding mismatch: {key}")
        columns, samples, metadata = parse_tsv(output)
        if len(samples) != 100 or metadata.get("samples") != ["100"] or metadata.get("warmups") != ["5"]:
            stop(f"native sample count mismatch: {output}")
        if metadata.get("probe") != ["0694"]:
            stop(f"native probe identifier mismatch: {output}")
        for required in ("capture_ns", "clone_ns", "settext_ns", "commit_ns", "apply_ns", "total_ns"):
            if required not in columns:
                stop(f"native timing column missing {required}: {output}")
        validate_workflow(metadata, CASES[row["case"]], output)
    if seen != {(case, leg) for case in CASES for leg in NATIVE_LEGS}:
        stop("native matrix coverage is incomplete")


def validate_allocations(bindings: dict[tuple[str, str], dict[str, Any]]) -> None:
    seen: set[tuple[str, str]] = set()
    records: list[dict[str, Any]] = []
    for phase in sorted(PHASES):
        rows = read(f"allocation-runs-{phase}.json")
        if len(rows) != len(CASES):
            stop(f"allocation {phase} manifest has {len(rows)} rows")
        for row in rows:
            key = (row.get("case"), phase)
            if key in seen or row["case"] not in CASES:
                stop(f"invalid or duplicate allocation row {key}")
            seen.add(key)
            if row.get("phase") != phase or row.get("workflow") != CASES[row["case"]] or row.get("exit_code") != 0:
                stop(f"allocation phase/workflow mismatch: {key}")
            if row.get("binary_sha256") != bindings[phase, "allocations"]["binary_sha256"]:
                stop(f"allocation binary binding mismatch: {key}")
            output = packet_path(row["output"])
            if sha(output) != row["output_sha256"]:
                stop(f"allocation output hash mismatch: {output}")
            if "stderr_sha256" in row and sha(output.with_suffix(".stderr")) != row["stderr_sha256"]:
                stop(f"allocation stderr hash mismatch: {output}")
            columns, samples, metadata = parse_tsv(output)
            if len(samples) != 3 or metadata.get("samples") != ["3"] or metadata.get("warmups") != ["2"]:
                stop(f"allocation sample count mismatch: {output}")
            validate_workflow(metadata, CASES[row["case"]], output)
            for stage in ALLOC_PHASES:
                for metric in ALLOC_METRICS:
                    if f"{stage}_{metric}" not in columns:
                        stop(f"allocation metric missing {stage}_{metric}: {output}")
            records.append(row)
    if seen != {(case, phase) for case in CASES for phase in PHASES}:
        stop("allocation matrix coverage is incomplete")


def parse_refusal(path: Path) -> tuple[dict[str, str], dict[str, dict[str, list[str] | list[int]]]]:
    header: dict[str, str] = {}
    cases: dict[str, dict[str, list[str] | list[int]]] = {}
    current: dict[str, list[str] | list[int]] | None = None
    for line in path.read_text().splitlines():
        fields = line.split("\t")
        key = fields[0]
        if key == "case":
            if len(fields) != 2 or fields[1] in cases:
                stop(f"{path}: duplicate/malformed refusal case")
            current = {"metadata": {}, "samples": []}
            cases[fields[1]] = current
        elif key == "sample_ns":
            continue
        elif key.isdigit():
            if current is None or len(fields) != 2:
                stop(f"{path}: sample outside refusal case")
            samples = current["samples"]
            assert isinstance(samples, list)
            if int(key) != len(samples):
                stop(f"{path}: refusal sample index is not contiguous")
            samples.append(int(fields[1]))
        elif key == "all_iterations_passed":
            if len(fields) != 2 or fields[1] != "true":
                stop(f"{path}: refusal iteration failed")
            header[key] = fields[1]
        elif current is None:
            if len(fields) != 2:
                stop(f"{path}: malformed refusal header")
            header[key] = fields[1]
        else:
            metadata = current["metadata"]
            assert isinstance(metadata, dict)
            metadata[key] = fields[1:]
    return header, cases


def validate_refusal() -> None:
    records = read("refusal-runs-baseline.json") + read("refusal-runs-compare.json")
    if len(records) != 6 or {row["leg"] for row in records} != NATIVE_LEGS:
        stop("refusal run matrix is incomplete")
    bindings: dict[str, dict[str, list[str]]] = {}
    for row in records:
        phase = "baseline" if row["leg"].startswith("a") else "candidate"
        if row.get("phase") != phase or row.get("exit_code") != 0:
            stop(f"invalid refusal phase row: {row}")
        build = read(f"build-refusal-{phase}.json")
        if row.get("binary_sha256") != build["binary_sha256"]:
            stop(f"refusal binary binding mismatch: {row['leg']}")
        output = packet_path(row["output"])
        if sha(output) != row["output_sha256"] or sha(output.with_suffix(".stderr")) != row["stderr_sha256"]:
            stop(f"refusal output hash mismatch: {output}")
        header, cases = parse_refusal(output)
        if header.get("probe") != "0693-refusal" or header.get("mode") != "matrix" or header.get("samples") != "100" or header.get("warmups") != "5" or header.get("all_iterations_passed") != "true":
            stop(f"refusal header mismatch: {output}")
        if set(cases) != REFUSAL_CASES or header.get("cases") != "10":
            stop(f"refusal case census mismatch: {output}")
        for name, data in cases.items():
            samples = data["samples"]
            metadata = data["metadata"]
            assert isinstance(samples, list) and isinstance(metadata, dict)
            if len(samples) != 100 or any(not isinstance(value, int) or value < 0 for value in samples):
                stop(f"refusal sample census mismatch: {output}/{name}")
            if "expected_error_debug" not in metadata:
                stop(f"refusal expected error missing: {output}/{name}")
            if "observed_error_debug" in metadata and metadata["observed_error_debug"] != metadata["expected_error_debug"]:
                stop(f"refusal outcome mismatch: {output}/{name}")
            if name in bindings and bindings[name] != metadata:
                stop(f"refusal metadata changed across legs: {name}")
            bindings.setdefault(name, metadata)


def validate_profiles_and_sizes(bindings: dict[tuple[str, str], dict[str, Any]]) -> None:
    for phase in sorted(PHASES):
        profile = read(f"profile/{phase}/binding.json")
        if profile.get("binary_sha256") != bindings[phase, "native"]["binary_sha256"]:
            stop(f"profile binary binding mismatch: {phase}")
        source = ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"
        if profile.get("source_sha256") != sha(source):
            stop(f"profile source binding mismatch: {phase}")
        for name, digest in profile.get("files", {}).items():
            path = P / "profile" / phase / name
            if sha(path) != digest:
                stop(f"profile file hash mismatch: {path}")
        if profile.get("raw_data_sha256") is not None and not profile.get("raw_data_sha256"):
            stop(f"invalid profile raw-data hash: {phase}")
    sizes = read("binary-sizes.json")
    for phase in sorted(PHASES):
        for label in ("native", "allocations", "refusal"):
            row = sizes[phase][label]
            if row.get("exit_code") != 0 or len(row.get("sha256", "")) != 64 or row.get("bytes", 0) <= 0:
                stop(f"invalid binary size receipt: {phase}/{label}")
            binary = ROOT.parent / "litchi-0694-bin" / f"{phase}-{label}"
            if binary.exists() and sha(binary) != row["sha256"]:
                stop(f"binary size hash mismatch: {phase}/{label}")
    metrics = read("profile-summary.json")
    expected_counters = {"cycles", "instructions", "branches", "branch-misses", "cache-misses", "page-faults", "task-clock"}
    if set(metrics) != PHASES:
        stop("profile summary phase census mismatch")
    for phase, summary in metrics.items():
        counters = summary.get("counters", {})
        # JSON object keys are strings in the retained file.
        if set(counters) != {"10", "210"}:
            stop(f"profile counter census mismatch: {phase}")
        ten = counters.get("10", counters.get(10, {}))
        two_ten = counters.get("210", counters.get(210, {}))
        if set(ten) != expected_counters or set(two_ten) != expected_counters:
            stop(f"profile counter names mismatch: {phase}")
        for name in expected_counters:
            if not all(math.isfinite(float(value)) for value in (ten[name], two_ten[name])):
                stop(f"non-finite profile counter: {phase}/{name}")
        per_open = summary.get("per_open_capture", {})
        for name in expected_counters:
            expected = (float(two_ten[name]) - float(ten[name])) / 200.0
            if not math.isclose(float(per_open[name]), expected, rel_tol=0.0, abs_tol=1e-9):
                stop(f"profile per-open formula mismatch: {phase}/{name}")
        if int(summary.get("whole_child_peak_rss_kib", 0)) <= 0:
            stop(f"profile RSS receipt is missing: {phase}")


def validate_gates() -> None:
    current = source_map()
    integration = read("integration/results.json")
    if len(integration) != 7 or {row.get("name") for row in integration} != {"fmt", "check", "clippy", "tests-default", "tests", "facade", "rustdoc"}:
        stop("integration gate census mismatch")
    for row in integration:
        if row.get("exit_code") != 0:
            stop(f"integration gate failed: {row.get('name')}")
        # Integration intentionally records the broader five-crate source
        # census (including tracked non-Rust files and generated Rust under
        # fuzz targets).  Check every recorded digest and require the frozen
        # 602-file evidence source map as a subset.
        for name, digest in row.get("source_sha256", {}).items():
            path = ROOT / name
            if not path.is_file() or sha(path) != digest:
                stop(f"integration source binding mismatch: {row['name']}/{name}")
        for name, digest in current.items():
            if row.get("source_sha256", {}).get(name) != digest:
                stop(f"integration source census omits/changes evidence source: {row['name']}/{name}")
        log = need(P / "integration" / (row["name"] + ".log"))
        if "log_sha256" in row and sha(log) != row["log_sha256"]:
            stop(f"integration log hash mismatch: {row['name']}")
    quality = read("quality-summary.json")
    if len(quality) != 7:
        stop("quality summary gate census mismatch")
    for row in quality:
        if row.get("exit_code") != 0 or row.get("test_totals", {}).get("failed") != 0:
            stop(f"quality gate failed: {row.get('name')}")
        log = need(P / "integration" / (row["name"] + ".log"))
        if sha(log) != row.get("log_sha256"):
            stop(f"quality log hash mismatch: {row['name']}")
    evidence = read("evidence/results.json")
    expected = {"crate-boundaries", "claims", "claims-structural", "report", "coverage", "non-iwork"}
    if len(evidence) != 6 or {row.get("name") for row in evidence} != expected:
        stop("evidence gate census mismatch")
    for row in evidence:
        if row.get("exit_code") != 0:
            stop(f"evidence gate failed: {row.get('name')}")
        for name, digest in row.get("source_sha256", {}).items():
            path = ROOT / name
            if not path.is_file() or sha(path) != digest:
                stop(f"evidence source binding mismatch: {row['name']}/{name}")
        for name, digest in current.items():
            if row.get("source_sha256", {}).get(name) != digest:
                stop(f"evidence source census omits/changes evidence source: {row['name']}/{name}")
        log = need(P / "evidence" / (row["name"] + ".log"))
        if sha(log) != row.get("log_sha256"):
            stop(f"evidence log hash mismatch: {row['name']}")


def validate_oracle_builds(base: dict[str, Any]) -> None:
    probe = packet_tree(P / "oracle", names={"Cargo.toml", "Cargo.lock", "main.rs"})
    for phase in sorted(PHASES):
        row = read(f"build-oracle-{phase}.json")
        if row.get("phase") != phase or row.get("exit_code") != 0 or row.get("restored_exact") is not True:
            stop(f"oracle {phase} build/restoration receipt failed")
        expected_source = base["source_sha256"] if phase == "baseline" else source_map()
        assert_source_map(row.get("source_sha256", {}), expected_source, f"oracle {phase}")
        if row.get("probe_sha256") != probe:
            stop(f"oracle {phase} probe binding mismatch")
        expected_restored = {
            "crates/litchi-ooxml-common/src/mce/codec.rs": sha(ROOT / "crates/litchi-ooxml-common/src/mce/codec.rs"),
            "crates/litchi-ooxml-common/src/mce/tests.rs": sha(ROOT / "crates/litchi-ooxml-common/src/mce/tests.rs"),
        }
        if row.get("restored_sha256") != expected_restored:
            stop(f"oracle {phase} restoration hash mismatch")
        binary = Path(row.get("binary", ""))
        if binary.exists() and sha(binary) != row.get("binary_sha256"):
            stop(f"oracle {phase} binary hash mismatch")
    if read("build-oracle-baseline.json").get("restore_required") is not True or read("build-oracle-candidate.json").get("restore_required") is not False:
        stop("oracle restore-required flags are wrong")
    archive = P / "oracle-baseline-restoration"
    for name in ("codec.rs", "tests.rs"):
        if sha(archive / name) != sha(ROOT / "crates/litchi-ooxml-common/src/mce" / name):
            stop(f"oracle baseline restoration archive mismatch: {name}")
    initial = P / "oracle-lock-initial" / "build-oracle-baseline.json"
    if initial.exists():
        receipt = json.loads(initial.read_text())
        if receipt.get("phase") != "baseline" or receipt.get("exit_code") == 0 or receipt.get("restored_exact") is not True:
            stop("superseded oracle-lock initial receipt is not a recorded failed/restored attempt")
        assert_source_map(receipt.get("source_sha256", {}), base["source_sha256"], "superseded oracle baseline")
        initial_probe = dict(probe)
        initial_probe["oracle/Cargo.lock"] = sha(P / "oracle-lock-initial" / "Cargo.lock")
        if receipt.get("probe_sha256") != initial_probe:
            stop("superseded oracle-lock probe binding changed")
        if receipt.get("restored_sha256") != read("build-oracle-baseline.json").get("restored_sha256"):
            stop("superseded oracle-lock restoration hash changed")


def validate_oracle_runs() -> None:
    """Validate corpus.py's exact-output differential packet."""
    output = need(P / "oracle-results")
    corpus = json.loads(need(output / "corpus.json").read_text())
    if corpus.get("probe") != "0694-mce-oracle" or corpus.get("schema") != 1:
        stop("oracle corpus identity/schema mismatch")
    profiles = ("baseline", "opaque", "opaque-small", "opaque-large", "opaque-many")
    if tuple(corpus.get("profiles", ())) != profiles:
        stop("oracle corpus profile census mismatch")
    cases = corpus.get("cases", [])
    if corpus.get("case_count") != len(cases) or len(cases) < 10:
        stop("oracle corpus case census is too small or inconsistent")
    by_id: dict[str, dict[str, Any]] = {}
    for ordinal, case in enumerate(cases):
        case_id = case.get("id")
        if not isinstance(case_id, str) or case_id in by_id or case.get("ordinal") != ordinal:
            stop(f"oracle corpus case ordering/identity mismatch: {case_id}")
        case_path = packet_path(output / case["path"])
        data = case_path.read_bytes()
        if len(data) != case.get("input_len") or sha(case_path) != case.get("input_sha256"):
            stop(f"oracle corpus input hash/length mismatch: {case_id}")
        origin = case.get("origin", {})
        if origin.get("kind") != "synthetic":
            archive = ROOT / "test-data" / origin["archive"]
            if not archive.is_file() or sha(archive) != origin.get("archive_sha256"):
                stop(f"oracle corpus archive binding mismatch: {case_id}")
        mutation = case.get("mutation", {})
        if mutation.get("parent"):
            parent = by_id.get(mutation["parent"])
            if parent is None or mutation.get("parent_sha256") != parent.get("input_sha256"):
                stop(f"oracle corpus mutation parent mismatch: {case_id}")
        by_id[case_id] = case
    canonical_rows = [
        {
            "id": case["id"],
            "path": case["path"],
            "input_len": case["input_len"],
            "input_sha256": case["input_sha256"],
            "origin": case["origin"],
            "mutation": case["mutation"],
        }
        for case in cases
    ]
    canonical = json.dumps(canonical_rows, ensure_ascii=True, sort_keys=True, separators=(",", ":")).encode()
    if sha_bytes(canonical) != corpus.get("corpus_sha256"):
        stop("oracle corpus aggregate hash mismatch")

    invocation_path = need(output / "invocations.jsonl")
    invocations = [json.loads(line) for line in invocation_path.read_text().splitlines() if line]
    expected_count = len(cases) * len(profiles) * 2
    if len(invocations) != expected_count:
        stop(f"oracle invocation count {len(invocations)} != {expected_count}")
    grouped: dict[tuple[str, str, str], dict[str, Any]] = {}
    oracle_builds = {phase: read(f"build-oracle-{phase}.json") for phase in sorted(PHASES)}
    for row in invocations:
        side = row.get("side")
        if side not in PHASES or row.get("probe") != "0694-mce-oracle" or row.get("status") != "completed" or row.get("returncode") != 0:
            stop(f"invalid oracle invocation receipt: {row}")
        case_id = row.get("case")
        profile = row.get("profile")
        if case_id not in by_id or profile not in profiles:
            stop(f"oracle invocation references unknown case/profile: {case_id}/{profile}")
        key = (side, case_id, profile)
        if key in grouped:
            stop(f"duplicate oracle invocation: {key}")
        grouped[key] = row
        if row.get("case_sha256") != by_id[case_id]["input_sha256"]:
            stop(f"oracle invocation case hash mismatch: {key}")
        if row.get("binary_sha256") != oracle_builds[side]["binary_sha256"]:
            stop(f"oracle invocation binary binding mismatch: {key}")
        stdout = row.get("stdout", "").encode()
        stderr = row.get("stderr", "").encode()
        if sha_bytes(stdout) != row.get("stdout_sha256") or sha_bytes(stderr) != row.get("stderr_sha256"):
            stop(f"oracle invocation stream hash mismatch: {key}")
        lines = stdout.decode("utf-8", "replace").splitlines()
        if len(lines) != 1 or not lines[0].startswith(("OK\t", "ERR\t")) or "probe=0694-mce-oracle" not in lines[0]:
            stop(f"oracle invocation did not emit one typed record: {key}")
        command = row.get("command")
        if command != ["run", profile, by_id[case_id]["path"]]:
            stop(f"oracle invocation command mismatch: {key}")
    expected_keys = {(side, case_id, profile) for side in PHASES for case_id in by_id for profile in profiles}
    if set(grouped) != expected_keys:
        stop("oracle invocation coverage is incomplete")
    for case_id in by_id:
        for profile in profiles:
            baseline = grouped["baseline", case_id, profile]["stdout"]
            candidate = grouped["candidate", case_id, profile]["stdout"]
            if baseline != candidate:
                stop(f"oracle exact-output mismatch: {case_id}/{profile}")

    result = json.loads(need(output / "results.json").read_text())
    if result.get("probe") != "0694-mce-oracle" or result.get("schema") != 1:
        stop("oracle result identity/schema mismatch")
    if result.get("baseline_binary_sha256") != oracle_builds["baseline"]["binary_sha256"] or result.get("candidate_binary_sha256") != oracle_builds["candidate"]["binary_sha256"]:
        stop("oracle result binary binding mismatch")
    if result.get("case_count") != len(cases) or result.get("profile_count") != len(profiles) or result.get("invocation_count") != expected_count:
        stop("oracle result count mismatch")
    if result.get("corpus_sha256") != corpus["corpus_sha256"] or result.get("mismatch_count") != 0 or result.get("mismatches") != []:
        stop("oracle differential reports a mismatch")
    timings = output / "timings.json"
    if timings.exists():
        timing_data = json.loads(timings.read_text())
        if timing_data.get("probe") != "0694-mce-oracle" or not timing_data.get("records"):
            stop("oracle timing receipt is malformed")
        for row in timing_data["records"]:
            if row.get("side") not in PHASES or row.get("profile") not in {"baseline", "opaque-small", "opaque-large"} or row.get("status") != "completed" or row.get("returncode") != 0:
                stop("oracle timing receipt failed")


def _validate_oracle_control_matrix(rows_name: str, comparisons_name: str, cases: tuple[str, ...]) -> None:
    rows = read(rows_name)
    expected_row_count = len(cases) * 3 * len(NATIVE_LEGS)
    if len(rows) != expected_row_count:
        stop(f"{rows_name} row count {len(rows)} != {expected_row_count}")
    expected = {
        (case, profile, leg)
        for case in cases
        for profile in ("baseline", "opaque", "opaque-many")
        for leg in NATIVE_LEGS
    }
    seen: set[tuple[str, str, str]] = set()
    identities: dict[tuple[str, str], str] = {}
    raw_stats: dict[tuple[str, str, str], dict[str, float]] = {}
    for row in rows:
        key = (row.get("case"), row.get("profile"), row.get("leg"))
        if key in seen or key not in expected:
            stop(f"invalid/duplicate oracle control row: {key}")
        seen.add(key)
        phase = "baseline" if row["leg"].startswith("a") else "candidate"
        if row.get("phase") != phase or row.get("exit_code") != 0:
            stop(f"oracle control phase/exit mismatch: {key}")
        if row.get("binary_sha256") != read(f"build-oracle-{phase}.json")["binary_sha256"]:
            stop(f"oracle control binary mismatch: {key}")
        source = packet_path(row["source"])
        if sha(source) != row.get("source_sha256"):
            stop(f"oracle control source mismatch: {key}")
        output = packet_path(row["output"])
        if sha(output) != row.get("output_sha256") or sha(output.with_suffix(".stderr")) != row.get("stderr_sha256"):
            stop(f"oracle control raw hash mismatch: {key}")
        lines = output.read_text().splitlines()
        if not lines or not lines[0].startswith("TIMING\t") or sum(line.startswith("SAMPLE\t") for line in lines) != 300:
            stop(f"oracle control timing shape mismatch: {key}")
        timing = dict(field.split("=", 1) for field in lines[0].split("\t")[1:] if "=" in field)
        if timing.get("probe") != "0694-mce-oracle" or timing.get("profile") != row["profile"] or timing.get("warmups") != "10" or timing.get("samples") != "300":
            stop(f"oracle control timing header mismatch: {key}")
        samples: list[int] = []
        for line in lines[1:]:
            fields = line.split("\t")
            if len(fields) != 3 or not fields[0].startswith("SAMPLE"):
                stop(f"oracle control sample record malformed: {key}")
            try:
                index = int(fields[1].split("=", 1)[1])
                elapsed = int(fields[2].split("=", 1)[1])
            except (IndexError, ValueError) as exc:
                stop(f"oracle control sample value malformed: {key}: {exc}")
            if index != len(samples) or elapsed <= 0:
                stop(f"oracle control sample sequence malformed: {key}")
            samples.append(elapsed)
        if len(samples) != 300:
            stop(f"oracle control sample count mismatch: {key}")
        ordered = sorted(samples)
        median = (ordered[149] + ordered[150]) / 2
        expected_stats = {
            "p50_ns": median,
            "mean_ns": sum(samples) / len(samples),
            "p95_ns": ordered[284],
            "p99_ns": ordered[296],
        }
        for name, expected_value in expected_stats.items():
            if not math.isclose(float(row[name]), float(expected_value), rel_tol=0.0, abs_tol=1e-9):
                stop(f"oracle control raw {name} mismatch: {key}")
        raw_stats[key] = expected_stats
        identity = row.get("identity", "")
        if not identity.startswith("OK\tprobe=0694-mce-oracle"):
            stop(f"oracle control identity record mismatch: {key}")
        identity_key = (row["case"], row["profile"])
        if identity_key in identities and identities[identity_key] != identity:
            stop(f"oracle control identity changed across legs: {identity_key}")
        identities.setdefault(identity_key, identity)
    if seen != expected:
        stop(f"{rows_name} matrix coverage is incomplete")
    comparisons = read(comparisons_name)
    expected_comparison_count = len(cases) * 3 * 4
    if len(comparisons) != expected_comparison_count:
        stop(f"{comparisons_name} comparison count mismatch")
    pairs = {"a1/a0", "b0/a2", "b1/a3", "a3/a2"}
    expected_comparisons = {
        (case, profile, pair)
        for case in cases
        for profile in ("baseline", "opaque", "opaque-many")
        for pair in pairs
    }
    comparison_keys = {(row.get("case"), row.get("profile"), row.get("pair")) for row in comparisons}
    if comparison_keys != expected_comparisons:
        stop(f"{comparisons_name} case/profile/pair census mismatch")
    for row in comparisons:
        case, profile, pair = row["case"], row["profile"], row["pair"]
        numerator, denominator = pair.split("/", 1)
        if numerator not in NATIVE_LEGS or denominator not in NATIVE_LEGS:
            stop(f"oracle control comparison legs malformed: {case}/{profile}/{pair}")
        expected_delta = {
            metric: (raw_stats[case, profile, numerator][metric] / raw_stats[case, profile, denominator][metric] - 1) * 100
            for metric in ("p50_ns", "mean_ns", "p95_ns", "p99_ns")
        }
        if set(row.get("delta_pct", {})) != set(expected_delta):
            stop(f"oracle control comparison metrics malformed: {case}/{profile}/{pair}")
        for metric, expected_value in expected_delta.items():
            if not math.isclose(float(row["delta_pct"][metric]), expected_value, rel_tol=0.0, abs_tol=1e-9):
                stop(f"oracle control delta mismatch: {case}/{profile}/{pair}/{metric}")


def validate_oracle_controls() -> None:
    _validate_oracle_control_matrix(
        "oracle-controls.json", "oracle-control-comparisons.json", ("ordinary", "opaque")
    )
    _validate_oracle_control_matrix(
        "oracle-real-controls.json",
        "oracle-real-control-comparisons.json",
        ("docx", "xlsx", "pptx"),
    )
    selected = read("oracle-real-cases.json")
    if len(selected) != 3 or {row.get("family") for row in selected} != {"docx", "xlsx", "pptx"}:
        stop("oracle real control source-family census mismatch")
    corpus = read("oracle-results/corpus.json")
    corpus_cases = {row["id"]: row for row in corpus["cases"]}
    for row in selected:
        family = row["family"]
        case = row.get("case", {})
        if case.get("id") not in corpus_cases or case.get("origin", {}).get("kind") != family or case.get("mutation", {}).get("kind") != "identity":
            stop(f"oracle real control selection is not a corpus identity case: {family}")
        source = P / "oracle-real-controls" / f"{family}.xml"
        source_case = corpus_cases[case["id"]]
        if not source.is_file() or sha(source) != source_case["input_sha256"]:
            stop(f"oracle real control source bytes mismatch: {family}")
        if b"http://schemas.openxmlformats.org/markup-compatibility/2006" not in source.read_bytes():
            stop(f"oracle real control source lacks MCE marker: {family}")


def rerun_deterministic() -> None:
    scripts = [
        "summarize.py",
        "summarize-allocations.py",
        "summarize-refusal.py",
        "report-metrics.py",
    ]
    if (P / "marker-control.pptx").exists():
        scripts.insert(0, "check-control.py")
    if (P / "quality-summary.py").exists() and (P / "integration" / "results.json").exists():
        scripts.append("quality-summary.py")
    for path in sorted(P.glob("summarize-oracle*.py")):
        scripts.append(path.name)
    outputs = {
        path: path.read_bytes()
        for path in P.iterdir()
        if path.is_file() and path.name in {
            "control-metadata-check.json",
            "native-summary.json",
            "native-comparisons.json",
            "semantic-bindings.json",
            "native-leg-metadata.json",
            "native-review-triggers.json",
            "baseline-noise.json",
            "tables.md",
            "allocation-summary.json",
            "allocation-comparisons.json",
            "refusal-summary.json",
            "refusal-bindings.json",
            "refusal-comparisons.json",
            "refusal-review-triggers.json",
            "refusal-tables.md",
            "profile-summary.json",
            "quality-summary.json",
        }
    }
    for script in scripts:
        result = subprocess.run([sys.executable, str(P / script)], cwd=ROOT, capture_output=True, text=True)
        if result.returncode:
            stop(f"deterministic script failed: {script}: {result.stderr[-1000:]}")
    for path, before in outputs.items():
        if not path.exists() or path.read_bytes() != before:
            stop(f"deterministic output changed after rerun: {path.name}")


def validate_script_syntax() -> None:
    for path in sorted(P.rglob("*.py")):
        if "__pycache__" in path.parts:
            continue
        try:
            ast.parse(path.read_text(), filename=str(path))
        except SyntaxError as exc:
            stop(f"script syntax error: {path.name}: {exc}")
    # A final script hash manifest is optional for older packet tooling, but if
    # present it must bind every listed driver to the retained bytes.
    manifest = P / "script-hashes.json"
    if manifest.exists():
        data = json.loads(manifest.read_text())
        assert_hash_map(data, P)


def validate_cleanup() -> None:
    cleanup_path = P / "cleanup.json"
    if not cleanup_path.exists():
        # The first audit is intentionally run before cleanup.py.  It must
        # prove that all eight phase binaries and the generated control input
        # are still available before the destructive, receipt-producing step.
        for phase in sorted(PHASES):
            for label in ("native", "allocations", "refusal", "oracle"):
                binary = ROOT.parent / "litchi-0694-bin" / f"{phase}-{label}"
                if not binary.is_file():
                    stop(f"pre-cleanup binary is missing: {binary}")
        need(P / "marker-control.pptx")
        return
    cleanup = read("cleanup.json")
    rows = cleanup.get("removed")
    if not isinstance(rows, list) or not rows:
        stop("cleanup receipt lacks removed paths")
    paths: set[str] = set()
    for row in rows:
        path_text = row.get("path")
        if not isinstance(path_text, str) or path_text in paths:
            stop("cleanup receipt has duplicate/malformed path")
        paths.add(path_text)
        if row.get("removed") is not True or Path(path_text).exists():
            stop(f"cleanup receipt does not prove removal: {path_text}")
    required = {
        str(ROOT.parent / "litchi-target-0694"),
        str(ROOT.parent / "litchi-0694-bin"),
        str(ROOT.parent / "litchi-0694-profile"),
        str(P / "marker-control.pptx"),
    }
    if paths != required:
        stop(f"cleanup receipt path set mismatch: extra={sorted(paths - required)} missing={sorted(required - paths)}")
    workspace_lock = ROOT / "Cargo.lock"
    if cleanup.get("workspace_cargo_lock_preserved_sha256") != sha(workspace_lock):
        stop("cleanup changed or failed to bind workspace Cargo.lock")


def main() -> None:
    base = read("baseline.json")
    for name, digest in base["constraints_sha256"].items():
        if sha(ROOT / name) != digest:
            stop(f"constraint hash changed: {name}")
    for name, digest in base["build_inputs_sha256"].items():
        if sha(ROOT / name) != digest:
            stop(f"build-input hash changed: {name}")
    current = source_map()
    if len(current) != len(base["source_sha256"]):
        stop(f"source file census changed: {len(current)} vs {len(base['source_sha256'])}")
    validate_script_syntax()
    validate_control()
    bindings = validate_builds(base)
    validate_refusal_builds(base, bindings)
    validate_oracle_builds(base)
    validate_native(bindings)
    validate_allocations(bindings)
    validate_refusal()
    validate_profiles_and_sizes(bindings)
    validate_oracle_runs()
    validate_oracle_controls()
    validate_gates()
    rerun_deterministic()
    validate_cleanup()
    print("PASS: 0694 constraints, source/probe/build bindings, raw matrices, oracle parity, profiles, gates, deterministic summaries and cleanup")


if __name__ == "__main__":
    main()
