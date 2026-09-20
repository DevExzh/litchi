#!/usr/bin/env python3
"""Verify the frozen 0702 evidence packet without rebuilding or measuring.

The packet deliberately keeps raw process output as the evidence boundary.
This audit checks the input/source/binary bindings in every receipt, parses the
raw native, allocation, refusal, oracle and marker-control records, reruns only
deterministic summarizers, and verifies the final gate and cleanup receipts.
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
import tempfile
import zipfile
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OWNERS = ("litchi-pptx", "litchi-ooxml-common", "litchi-opc")
PHASES = {"baseline", "candidate"}
BATCH = "0702"
ORACLE_PROBE = f"{BATCH}-mce-oracle"
BIN_DIR = f"litchi-{BATCH}-bin"
TARGET_DIR = f"litchi-target-{BATCH}"
PROFILE_DIR = f"litchi-{BATCH}-profile"
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
STAT_FIELDS = ("p50_ns", "mean_ns", "p95_ns", "p99_ns", "min_ns", "max_ns")
# The tester supplies the final focused `mce::` census before the terminal
# audit. Keep this explicit rather than accepting an unconstrained count.
FOCUSED_PASS_COUNT = 104


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


def assert_0701_helper_shape(text: str, label: str) -> None:
    """Require the already-measured first-byte helper shape in both phases."""
    if "memmem" in text:
        stop(f"{label} unexpectedly contains memmem")
    if ".windows(NAMESPACE.len())" in text:
        stop(f"{label} still contains the namespace windows scan")
    if not re.search(r"#\s*\[\s*inline\s*\(\s*never\s*\)\s*\]", text):
        stop(f"{label} lacks the inline-never namespace helper")
    helper = re.search(r"\bfn\s+contains_mce_namespace\s*\(", text)
    if helper is None:
        stop(f"{label} lacks contains_mce_namespace")
    body = text[helper.start():]
    if not re.search(r"\bneedle\s*=\s*NAMESPACE\.as_bytes\(\)", body):
        stop(f"{label} helper does not bind the MCE URI needle")
    if not re.search(r"\bfirst\s*=\s*needle\s*\[\s*0\s*\]", body):
        stop(f"{label} helper does not bind the URI first byte")
    if "starts_with(needle)" not in body and "starts_with(NAMESPACE.as_bytes())" not in body:
        stop(f"{label} helper lacks the exact URI starts_with check")
    if not re.search(r"(?:memchr::)?memchr\s*\(", body):
        stop(f"{label} helper lacks the first-byte memchr search")
    if not re.search(r"\bvalid_starts\s*=\s*remaining\.len\(\)\s*-\s*needle\.len\(\)\s*\+\s*1", body):
        stop(f"{label} helper lacks the valid-start bound")
    if not re.search(r"memchr\s*\(\s*first\s*,\s*&?remaining\[\.\.valid_starts\]", body):
        stop(f"{label} helper does not bound memchr to valid starts")
    # The loop must continue after a failed candidate byte, rather than
    # repeatedly testing the same position or skipping overlapping matches.
    if not re.search(r"(?:\+\s*1|checked_add\(\s*1\s*\))", body):
        stop(f"{label} helper lacks one-byte advancement")
    if re.search(r"if\s*!\s*contains_mce_namespace\s*\(\s*xml\s*\)", text) is None:
        stop(f"{label} does not use the helper at the MCE early predicate")


def helper_source(text: str) -> str:
    match = re.search(
        r"(?s)#\s*\[\s*inline\s*\(\s*never\s*\)\s*\]\s*"
        r"fn\s+contains_mce_namespace\s*\(.*?\n\}\n",
        text,
    )
    if match is None:
        stop("MCE helper source is missing")
    return match.group(0)


def assert_borrowed_inherited_shape(text: str) -> None:
    """Require the 0698 borrowed namespace-scope representation."""
    inherited = re.search(
        r"struct\s+Inherited\s*<\s*'a\s*>\s*\{(?P<body>.*?)\n\}",
        text,
        re.DOTALL,
    )
    if inherited is None:
        stop("candidate codec does not use the borrowed Inherited scope")
    body = inherited.group("body")
    if not re.search(r"ns:\s*Option\s*<\s*&\s*'a\s+Arc<NamespaceLayer>\s*>", body):
        stop("candidate Inherited.ns is not an Option<&'a Arc<NamespaceLayer>>")
    if not re.search(r"emitted:\s*Option\s*<\s*&\s*'a\s+Arc<NamespaceLayer>\s*>", body):
        stop("candidate Inherited.emitted is not an Option<&'a Arc<NamespaceLayer>>")
    impl = re.search(r"impl\s+Inherited\s*<\s*'_\s*>\s*\{(?P<body>.*?)\n\}", text, re.DOTALL)
    if impl is None:
        stop("candidate Inherited impl does not retain its lifetime")
    impl_body = impl.group("body")
    if "hoists(self.ns, self.emitted)" not in impl_body:
        stop("candidate Inherited hoists call does not borrow its fields")
    if "self.emitted.cloned()" not in impl_body:
        stop("candidate Inherited::after does not clone the borrowed emitted scope")
    if "ns: parent.and_then(|f| f.ctx.ns.head.as_ref())" not in text:
        stop("candidate start path does not borrow the parent namespace head")
    if "emitted: parent.and_then(|f| f.emitted_ns.as_ref())" not in text:
        stop("candidate start path does not borrow the parent emitted namespace head")
    if "Namespaces::for_each_hoisted(inherited.ns, inherited.emitted" not in text:
        stop("candidate writer does not pass borrowed inherited scopes")


def validate_source_change(base: dict[str, Any], current: dict[str, str]) -> dict[str, str]:
    """Bind the current checkout to the exact measured candidate diff."""
    codec_name = "crates/litchi-ooxml-common/src/mce/codec.rs"
    tests_name = "crates/litchi-ooxml-common/src/mce/tests.rs"
    expected_keys = set(base["source_sha256"])
    expected_changed = {codec_name}
    if len(current) != len(expected_keys) or set(current) != expected_keys:
        stop("candidate source map census mismatch")
    changed = {name for name in expected_keys if current[name] != base["source_sha256"][name]}
    if changed != expected_changed:
        stop(f"candidate source change census mismatch: {sorted(changed)}")

    baseline_codec = need(P / "baseline-codec.rs.txt")
    baseline_tests = need(P / "baseline-tests.rs.txt")
    if sha(baseline_codec) != base["source_sha256"][codec_name]:
        stop("baseline codec source witness hash mismatch")
    if sha(baseline_tests) != base["source_sha256"][tests_name]:
        stop("baseline tests source witness hash mismatch")
    candidate_text = (ROOT / codec_name).read_text()
    baseline_text = baseline_codec.read_text()
    assert_0701_helper_shape(baseline_text, "baseline codec")
    assert_0701_helper_shape(candidate_text, "candidate codec")
    if helper_source(candidate_text) != helper_source(baseline_text):
        stop("candidate changed the 0701 MCE helper source")
    if sha(baseline_tests) != base["source_sha256"][tests_name]:
        stop("baseline tests source witness hash mismatch")
    if sha(ROOT / tests_name) != base["source_sha256"][tests_name]:
        stop("candidate tests source changed")
    assert_borrowed_inherited_shape(candidate_text)

    receipt = read("source-diff.json")
    patch = need(P / "source-diff.patch")
    if receipt.get("baseline_head") != base["baseline_head"]:
        stop("source diff baseline head mismatch")
    if receipt.get("paths") != [codec_name]:
        stop("source diff path census mismatch")
    command = [
        "git", "diff", "--no-ext-diff", "--binary", "--full-index",
        base["baseline_head"], "--", codec_name,
    ]
    name_command = ["git", "diff", "--name-only", base["baseline_head"], "--", codec_name]
    if receipt.get("name_command") != name_command or receipt.get("diff_command") != command:
        stop("source diff command mismatch")
    if receipt.get("patch") != "source-diff.patch":
        stop("source diff patch path mismatch")
    if receipt.get("patch_bytes") != patch.stat().st_size or receipt.get("patch_sha256") != sha(patch):
        stop("source diff patch receipt mismatch")

    with tempfile.TemporaryDirectory(prefix="litchi-0702-source-") as temporary:
        root = Path(temporary)
        for name, witness in ((codec_name, baseline_codec),):
            target = root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(witness.read_bytes())
        checked = subprocess.run(["git", "apply", "--check", str(patch)], cwd=root, capture_output=True, text=True)
        if checked.returncode != 0:
            stop(f"candidate source diff cannot apply: {checked.stderr[-500:]}")
        applied = subprocess.run(["git", "apply", str(patch)], cwd=root, capture_output=True, text=True)
        if applied.returncode != 0:
            stop(f"candidate source diff application failed: {applied.stderr[-500:]}")
        if sha(root / codec_name) != current[codec_name]:
            stop("source diff does not produce the current candidate witnesses")
    return current


def validate_rejection_source_change(
    base: dict[str, Any],
    current: dict[str, str],
    rejection: dict[str, Any],
) -> tuple[dict[str, str], dict[str, str]]:
    """Bind the measured candidate, restored checkout, and rejection patch."""
    codec_name = "crates/litchi-ooxml-common/src/mce/codec.rs"
    tests_name = "crates/litchi-ooxml-common/src/mce/tests.rs"
    expected_keys = set(base["source_sha256"])
    expected_receipt_keys = {
        "disposition",
        "baseline_head",
        "reason",
        "codec_path",
        "tests_path",
        "candidate_witness",
        "baseline_codec_sha256",
        "candidate_codec_sha256",
        "final_codec_sha256",
        "final_tests_sha256",
        "source_diff_sha256",
        "candidate_source_sha256",
        "final_source_sha256",
        "retained_source_changes",
    }
    if set(rejection) != expected_receipt_keys:
        stop("rejection receipt schema mismatch")
    if rejection.get("disposition") != "rejected" or not rejection.get("reason"):
        stop("rejection disposition/reason is invalid")
    if rejection.get("baseline_head") != base["baseline_head"]:
        stop("rejection baseline head mismatch")
    if rejection.get("codec_path") != codec_name or rejection.get("tests_path") != tests_name:
        stop("rejection source path mismatch")
    if rejection.get("candidate_witness") != "candidate-codec.rs.txt":
        stop("rejection candidate witness path mismatch")
    if rejection.get("retained_source_changes") != []:
        stop("rejection retained source census mismatch")

    candidate_source = rejection.get("candidate_source_sha256")
    final_source = rejection.get("final_source_sha256")
    if not isinstance(candidate_source, dict) or not isinstance(final_source, dict):
        stop("rejection source maps are malformed")
    if set(candidate_source) != expected_keys or set(final_source) != expected_keys:
        stop("rejection source map census mismatch")
    if len(candidate_source) != 602 or len(final_source) != 602:
        stop("rejection source map file count mismatch")
    if rejection.get("candidate_codec_sha256") != candidate_source[codec_name]:
        stop("rejection candidate codec hash mismatch")
    if rejection.get("final_codec_sha256") != final_source[codec_name]:
        stop("rejection final codec hash mismatch")
    if rejection.get("final_tests_sha256") != final_source[tests_name]:
        stop("rejection final tests hash mismatch")
    if rejection.get("baseline_codec_sha256") != base["source_sha256"][codec_name]:
        stop("rejection baseline codec hash mismatch")
    changed_candidate = {name for name in expected_keys if candidate_source[name] != base["source_sha256"][name]}
    changed_final = {name for name in expected_keys if final_source[name] != base["source_sha256"][name]}
    if changed_candidate != {codec_name} or changed_final:
        stop(f"rejection source transition mismatch: candidate={sorted(changed_candidate)} final={sorted(changed_final)}")
    if candidate_source[tests_name] != final_source[tests_name]:
        stop("candidate and retained tests differ")
    assert_source_map(current, final_source, "restored final")

    baseline_codec = need(P / "baseline-codec.rs.txt")
    baseline_tests = need(P / "baseline-tests.rs.txt")
    candidate_witness = need(P / rejection["candidate_witness"])
    if sha(baseline_codec) != base["source_sha256"][codec_name]:
        stop("baseline codec witness hash mismatch")
    if sha(baseline_tests) != base["source_sha256"][tests_name]:
        stop("baseline tests witness hash mismatch")
    if sha(candidate_witness) != rejection["candidate_codec_sha256"]:
        stop("candidate codec witness hash mismatch")
    candidate_text = candidate_witness.read_text()
    assert_0701_helper_shape(candidate_text, "candidate codec witness")
    assert_borrowed_inherited_shape(candidate_text)

    receipt = read("source-diff.json")
    patch = need(P / "source-diff.patch")
    if receipt.get("baseline_head") != base["baseline_head"] or receipt.get("paths") != [codec_name]:
        stop("source diff baseline/path census mismatch")
    command = [
        "git", "diff", "--no-ext-diff", "--binary", "--full-index",
        base["baseline_head"], "--", codec_name,
    ]
    name_command = ["git", "diff", "--name-only", base["baseline_head"], "--", codec_name]
    if receipt.get("name_command") != name_command or receipt.get("diff_command") != command:
        stop("source diff command mismatch")
    if receipt.get("patch") != "source-diff.patch" or receipt.get("patch_bytes") != patch.stat().st_size:
        stop("source diff patch receipt mismatch")
    if receipt.get("patch_sha256") != sha(patch) or rejection.get("source_diff_sha256") != sha(patch):
        stop("source diff patch hash mismatch")

    with tempfile.TemporaryDirectory(prefix="litchi-0702-rejection-") as temporary:
        root = Path(temporary)
        for name, witness in ((codec_name, baseline_codec),):
            target = root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(witness.read_bytes())
        checked = subprocess.run(["git", "apply", "--check", str(patch)], cwd=root, capture_output=True, text=True)
        if checked.returncode:
            stop(f"candidate source diff cannot apply: {checked.stderr[-500:]}")
        applied = subprocess.run(["git", "apply", str(patch)], cwd=root, capture_output=True, text=True)
        if applied.returncode:
            stop(f"candidate source diff application failed: {applied.stderr[-500:]}")
        if sha(root / codec_name) != candidate_source[codec_name]:
            stop("source diff replay does not produce the candidate witnesses")
    return candidate_source, final_source

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


def validate_builds(
    base: dict[str, Any],
    candidate_source: dict[str, str],
) -> dict[tuple[str, str], dict[str, Any]]:
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
                assert_source_map(row["source_sha256"], candidate_source, f"candidate {label}")
            binary = Path(row["binary"])
            if binary.exists() and sha(binary) != row["binary_sha256"]:
                stop(f"{phase}/{label}: binary hash changed")
            bindings[phase, label] = row
    if bindings["baseline", "native"]["probe_sha256"] != bindings["candidate", "native"]["probe_sha256"]:
        stop("native probe changed between phases")
    if bindings["baseline", "allocations"]["probe_sha256"] != bindings["candidate", "allocations"]["probe_sha256"]:
        stop("allocation probe changed between phases")
    return bindings


def validate_refusal_builds(
    base: dict[str, Any],
    bindings: dict[tuple[str, str], dict[str, Any]],
    candidate_source: dict[str, str],
) -> None:
    for phase in sorted(PHASES):
        row = read(f"build-refusal-{phase}.json")
        if row.get("phase") != phase or row.get("exit_code") != 0:
            stop(f"invalid refusal build receipt for {phase}")
        expected_source = base["source_sha256"] if phase == "baseline" else candidate_source
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
        if metadata.get("probe") != [BATCH]:
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


def validate_refusal() -> dict[str, dict[str, list[str]]]:
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
    return bindings

def identity_metadata(metadata: dict[str, list[str]]) -> str:
    canonical = json.dumps(metadata, ensure_ascii=True, sort_keys=True, separators=(",", ":"))
    return sha_bytes(canonical.encode())


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
            binary = ROOT.parent / BIN_DIR / f"{phase}-{label}"
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


def validate_assembly(bindings: dict[tuple[str, str], dict[str, Any]]) -> None:
    """Verify the retained disassembly is tied to each frozen native binary."""
    folder = need(P / "assembly")
    for phase in sorted(PHASES):
        receipt_path = need(folder / f"{phase}.json")
        receipt = json.loads(receipt_path.read_text())
        if receipt.get("exit_code") != 0:
            stop(f"assembly command failed: {phase}")
        if receipt.get("binary_sha256") != bindings[phase, "native"]["binary_sha256"]:
            stop(f"assembly binary binding mismatch: {phase}")
        symbol = receipt.get("symbol")
        command = receipt.get("command")
        if not isinstance(symbol, str) or not symbol or not all(
            part in symbol for part in ("mce", "codec", "start")
        ):
            stop(f"assembly symbol is not the shared MCE start function: {phase}")
        if not isinstance(command, list) or command[:3] != ["objdump", "-d", "-l"]:
            stop(f"assembly command shape mismatch: {phase}")
        if command[3:] != [f"--disassemble={symbol}", str(ROOT.parent / BIN_DIR / f"{phase}-native")]:
            stop(f"assembly command binding mismatch: {phase}")
        if receipt.get("stderr") != "":
            stop(f"assembly command emitted stderr: {phase}")
        output = need(folder / f"{phase}-start.txt")
        if not output.read_text().strip():
            stop(f"assembly output is empty: {phase}")
        if sha(output) != receipt.get("output_sha256"):
            stop(f"assembly output hash mismatch: {phase}")


def validate_processor_assembly(bindings: dict[tuple[str, str], dict[str, Any]]) -> None:
    """Verify address-bounded processor disassembly and its command receipts."""
    folder = need(P / "processor-assembly")
    expected_symbol = "litchi_ooxml_common::mce::codec::process_markup_compatibility"
    helper_sizes: dict[str, int] = {}
    for phase in sorted(PHASES):
        receipt = json.loads(need(folder / f"{phase}.json").read_text())
        binary = str(ROOT.parent / BIN_DIR / f"{phase}-native")
        if receipt.get("phase") != phase or receipt.get("binary_sha256") != bindings[phase, "native"]["binary_sha256"]:
            stop(f"processor assembly phase/binary binding mismatch: {phase}")
        if receipt.get("symbol") != expected_symbol:
            stop(f"processor assembly symbol mismatch: {phase}")
        address, size = receipt.get("address"), receipt.get("size")
        if not isinstance(address, int) or address < 0 or not isinstance(size, int) or size <= 0:
            stop(f"processor assembly bounds are invalid: {phase}")
        commands = {row.get("name"): row for row in receipt.get("commands", [])}
        expected_commands = {"symbols", "processor", "marker-search"}
        if set(commands) != expected_commands:
            stop(f"processor assembly command census mismatch: {phase}")
        symbols = commands["symbols"]
        if symbols.get("command") != ["nm", "-S", "-C", binary] or symbols.get("exit_code") != 0:
            stop(f"processor nm command mismatch: {phase}")
        symbol_text = (folder / f"{phase}-symbols.stdout").read_text()
        symbol_bounds: dict[str, tuple[int, int]] = {}
        for line in symbol_text.splitlines():
            fields = line.split(maxsplit=3)
            if len(fields) != 4:
                continue
            try:
                symbol_bounds[fields[3]] = (int(fields[0], 16), int(fields[1], 16))
            except ValueError:
                continue
        if symbol_bounds.get(expected_symbol) != (address, size):
            stop(f"processor symbol bounds do not match nm: {phase}")
        processor = commands["processor"]
        expected = [
            "objdump", "-d", "-C", "--no-show-raw-insn",
            f"--start-address={address}", f"--stop-address={address + size}", binary,
        ]
        if processor.get("command") != expected or processor.get("exit_code") != 0:
            stop(f"processor disassembly command/bounds mismatch: {phase}")
        for name, row in commands.items():
            stdout = need(folder / f"{phase}-{name}.stdout")
            stderr = need(folder / f"{phase}-{name}.stderr")
            if sha(stdout) != row.get("stdout_sha256") or sha(stderr) != row.get("stderr_sha256"):
                stop(f"processor assembly output hash mismatch: {phase}/{name}")
            if stderr.read_bytes() != b"":
                stop(f"processor assembly emitted stderr: {phase}/{name}")
        if f"<{expected_symbol}>:" not in (folder / f"{phase}-processor.stdout").read_text():
            stop(f"processor disassembly lacks symbol heading: {phase}")
        processor_text = (folder / f"{phase}-processor.stdout").read_text()
        helper_symbol = "litchi_ooxml_common::mce::codec::contains_mce_namespace"
        helper_address = receipt.get("helper_address")
        helper_size = receipt.get("helper_size")
        if receipt.get("helper_symbol") != helper_symbol:
            stop(f"{phase} processor assembly helper symbol mismatch")
        if not isinstance(helper_address, int) or helper_address < 0:
            stop(f"{phase} processor assembly helper address is invalid")
        if not isinstance(helper_size, int) or helper_size <= 0:
            stop(f"{phase} processor assembly helper size is invalid")
        if symbol_bounds.get(helper_symbol) != (helper_address, helper_size):
            stop(f"{phase} helper bounds do not match nm")
        helper_sizes[phase] = helper_size
        marker = commands["marker-search"]
        expected_marker = [
            "objdump", "-d", "-C", "--no-show-raw-insn",
            f"--start-address={helper_address}",
            f"--stop-address={helper_address + helper_size}",
            binary,
        ]
        if marker.get("command") != expected_marker or marker.get("exit_code") != 0:
            stop(f"{phase} marker-search disassembly command/bounds mismatch")
        marker_text = (folder / f"{phase}-marker-search.stdout").read_text()
        if f"<{helper_symbol}>:" not in marker_text:
            stop(f"{phase} marker-search disassembly lacks helper heading")
        if "FinderBuilder" in processor_text or "memmem" in processor_text:
            stop(f"{phase} processor disassembly retains the memmem finder")
        if "FinderBuilder" in marker_text or "memmem" in marker_text:
            stop(f"{phase} helper disassembly retains the memmem finder")
        if re.search(r"\bcall\w*\b[^\n]*contains_mce_namespace", processor_text) is None:
            stop(f"{phase} processor disassembly lacks an outlined helper call")
    if helper_sizes.get("baseline") != helper_sizes.get("candidate"):
        stop("0701 helper code size changed between baseline and candidate")


def validate_marker_controls() -> None:
    """Run the independent marker-control verifier when its packet is present."""
    verifier = P / "audit_marker_controls.py"
    if not verifier.exists():
        stop("marker-control audit script is missing")
    result = subprocess.run(
        [sys.executable, str(verifier)],
        cwd=ROOT,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
        capture_output=True,
        text=True,
    )
    if result.returncode:
        stop(f"marker-control audit failed: {result.stderr[-1000:]}{result.stdout[-1000:]}")


def validate_followup_packet() -> None:
    """Require the independent native/refusal follow-up audit."""
    verifier = P / "audit_followup.py"
    if not verifier.exists():
        stop("follow-up audit script is missing")
    result = subprocess.run(
        [sys.executable, str(verifier)],
        cwd=ROOT,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
        capture_output=True,
        text=True,
    )
    if result.returncode:
        stop(f"follow-up audit failed: {result.stderr[-1000:]}{result.stdout[-1000:]}")

def validate_focused_tests(
    base: dict[str, Any],
    candidate_source: dict[str, str],
    final_source: dict[str, str] | None = None,
) -> None:
    """Require the fresh MCE focused test run on both frozen source phases."""
    codec_name = "crates/litchi-ooxml-common/src/mce/codec.rs"
    tests_name = "crates/litchi-ooxml-common/src/mce/tests.rs"
    names = {codec_name, tests_name}
    expected_command = [
        "cargo",
        "test",
        "-p",
        "litchi-ooxml-common",
        "--lib",
        "--locked",
        "mce::",
        "--",
        "--test-threads=1",
    ]
    receipts = {}
    for phase in sorted(PHASES):
        receipt = read(f"focused-{phase}.json")
        if receipt.get("phase") != phase or receipt.get("exit_code") != 0:
            stop(f"focused MCE test receipt failed: {phase}")
        if receipt.get("command") != expected_command:
            stop(f"focused MCE command mismatch: {phase}")
        if set(receipt.get("source_sha256", {})) != names:
            stop(f"focused MCE source census mismatch: {phase}")
        if phase == "baseline":
            if receipt["source_sha256"][codec_name] != base["source_sha256"][codec_name]:
                stop("focused baseline codec source binding mismatch")
        elif receipt["source_sha256"] != {name: candidate_source[name] for name in sorted(names)}:
            stop("focused candidate source binding mismatch")
        log = need(P / receipt.get("log", ""))
        if sha(log) != receipt.get("log_sha256"):
            stop(f"focused MCE log hash mismatch: {phase}")
        lines = [line.strip() for line in log.read_text().splitlines() if line.strip()]
        result_lines = [line for line in lines if line.startswith("test result:")]
        if len(result_lines) != 1 or lines[-1] != result_lines[0]:
            stop(f"focused MCE log lacks one final test result: {phase}")
        match = re.fullmatch(
            rf"test result: ok\. ({FOCUSED_PASS_COUNT}) passed; (\d+) failed; (\d+) ignored; (\d+) measured; (\d+) filtered out; finished in [0-9.]+s",
            result_lines[0],
        )
        if match is None or int(match.group(1)) <= 0 or int(match.group(2)) != 0:
            stop(f"focused MCE result census mismatch: {phase}")
        receipts[phase] = receipt
    candidate = receipts["candidate"]["source_sha256"]
    if candidate != {name: candidate_source[name] for name in sorted(names)}:
        stop("focused candidate source is stale")
    if receipts["baseline"]["source_sha256"][tests_name] != candidate[tests_name]:
        stop("focused runs used different final test files")
    if final_source is None:
        return

    retained = read("focused-retained.json")
    if retained.get("phase") != "retained" or retained.get("exit_code") != 0:
        stop("focused retained test receipt failed")
    if retained.get("command") != expected_command:
        stop("focused retained command mismatch")
    if retained.get("source_sha256") != {name: final_source[name] for name in sorted(names)}:
        stop("focused retained source binding mismatch")
    retained_log = need(P / retained.get("log", ""))
    if sha(retained_log) != retained.get("log_sha256"):
        stop("focused retained MCE log hash mismatch")
    retained_lines = [line.strip() for line in retained_log.read_text().splitlines() if line.strip()]
    retained_results = [line for line in retained_lines if line.startswith("test result:")]
    if len(retained_results) != 1 or retained_lines[-1] != retained_results[0]:
        stop("focused retained log lacks one final test result")
    retained_match = re.fullmatch(
        rf"test result: ok\. ({FOCUSED_PASS_COUNT}) passed; (\d+) failed; (\d+) ignored; (\d+) measured; (\d+) filtered out; finished in [0-9.]+s",
        retained_results[0],
    )
    if retained_match is None or int(retained_match.group(1)) != FOCUSED_PASS_COUNT or int(retained_match.group(2)) != 0:
        stop("focused retained result census mismatch")


def validate_gates(candidate_source: dict[str, str]) -> None:
    integration = read("integration/results.json")
    if len(integration) != 7 or {row.get("name") for row in integration} != {"fmt", "check", "clippy", "tests-default", "tests", "facade", "rustdoc"}:
        stop("integration gate census mismatch")
    for row in integration:
        if row.get("exit_code") != 0:
            stop(f"integration gate failed: {row.get('name')}")
        # Integration intentionally records the broader five-crate source
        # census (including tracked non-Rust files and generated Rust under
        # fuzz targets). Check every recorded digest and require the frozen
        # candidate 602-file evidence source map as a subset.
        for name, digest in row.get("source_sha256", {}).items():
            path = ROOT / name
            expected = candidate_source.get(name)
            if expected is not None:
                if digest != expected:
                    stop(f"integration source binding mismatch: {row['name']}/{name}")
            elif not path.is_file() or sha(path) != digest:
                stop(f"integration source binding mismatch: {row['name']}/{name}")
        for name, digest in candidate_source.items():
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
            expected = candidate_source.get(name)
            if expected is not None:
                if digest != expected:
                    stop(f"evidence source binding mismatch: {row['name']}/{name}")
            elif not path.is_file() or sha(path) != digest:
                stop(f"evidence source binding mismatch: {row['name']}/{name}")
        for name, digest in candidate_source.items():
            if row.get("source_sha256", {}).get(name) != digest:
                stop(f"evidence source census omits/changes evidence source: {row['name']}/{name}")
        log = need(P / "evidence" / (row["name"] + ".log"))
        if sha(log) != row.get("log_sha256"):
            stop(f"evidence log hash mismatch: {row['name']}")


def validate_oracle_builds(
    base: dict[str, Any],
    candidate_source: dict[str, str],
) -> None:
    # The direct 0702 oracle builder binds every retained oracle input,
    # including its corpus driver and README, so the receipt cannot silently
    # swap any part of the standalone probe tree.
    probe = packet_tree(P / "oracle")
    for phase in sorted(PHASES):
        row = read(f"build-oracle-{phase}.json")
        if row.get("phase") != phase or row.get("exit_code") != 0:
            stop(f"oracle {phase} build receipt failed")
        if "restore_required" in row or "restored_exact" in row or "restored_sha256" in row:
            stop(f"oracle {phase} receipt contains obsolete restoration fields")
        expected_source = base["source_sha256"] if phase == "baseline" else candidate_source
        assert_source_map(row.get("source_sha256", {}), expected_source, f"oracle {phase}")
        if row.get("probe_sha256") != probe:
            stop(f"oracle {phase} probe binding mismatch")
        binary = Path(row.get("binary", ""))
        if binary.exists() and sha(binary) != row.get("binary_sha256"):
            stop(f"oracle {phase} binary hash mismatch")


def validate_oracle_runs() -> None:
    """Validate corpus.py's exact-output differential packet."""
    output = need(P / "oracle-results")
    corpus = json.loads(need(output / "corpus.json").read_text())
    if corpus.get("probe") != ORACLE_PROBE or corpus.get("schema") != 1:
        stop("oracle corpus identity/schema mismatch")
    profiles = ("baseline", "opaque", "opaque-small", "opaque-large", "opaque-many")
    if tuple(corpus.get("profiles", ())) != profiles:
        stop("oracle corpus profile census mismatch")
    cases = corpus.get("cases", [])
    if corpus.get("case_count") != 192 or len(cases) != 192:
        stop("oracle corpus case census is not exactly 192 inputs")
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
        if side not in PHASES or row.get("probe") != ORACLE_PROBE or row.get("status") != "completed" or row.get("returncode") != 0:
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
        if len(lines) != 1 or not lines[0].startswith(("OK\t", "ERR\t")) or f"probe={ORACLE_PROBE}" not in lines[0]:
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
    if result.get("probe") != ORACLE_PROBE or result.get("schema") != 1:
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
        if timing_data.get("probe") != ORACLE_PROBE or not timing_data.get("records"):
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
        if timing.get("probe") != ORACLE_PROBE or timing.get("profile") != row["profile"] or timing.get("warmups") != "10" or timing.get("samples") != "300":
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
        if not identity.startswith(f"OK\tprobe={ORACLE_PROBE}"):
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
    _validate_oracle_control_matrix(
        "oracle-declaration-controls.json",
        "oracle-declaration-control-comparisons.json",
        ("declared", "mixed"),
    )
    validate_oracle_declaration_sources()
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


def validate_oracle_declaration_sources() -> None:
    """Bind declaration-control inputs to their deterministic generator."""
    mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
    drawingml = "http://schemas.openxmlformats.org/drawingml/2006/main"
    head = f'<r xmlns:mc="{mc}">'
    expected = {
        "declared": head + (f'<a:n xmlns:a="{drawingml}" a:x="1"/>' * 1000) + "</r>",
        "mixed": head + (f'<a:n xmlns:a="{drawingml}" a:x="1"><a:c a:y="2"/></a:n>' * 1000) + "</r>",
    }
    for case, text in expected.items():
        source = need(P / "oracle-declaration-controls" / f"{case}.xml")
        if source.read_text() != text:
            stop(f"oracle declaration-control source differs from generator: {case}")


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
    manifest = need(P / "script-hashes.json")
    data = json.loads(manifest.read_text())
    expected = {str(path.relative_to(P)) for path in P.rglob("*.py")
                if "__pycache__" not in path.parts}
    if set(data) != expected:
        stop("script hash manifest census mismatch")
    assert_hash_map(data, P)


def validate_cleanup() -> None:
    cleanup_path = P / "cleanup.json"
    if not cleanup_path.exists():
        # The first audit is intentionally run before cleanup.py.  It must
        # prove that all eight phase binaries and the generated control input
        # are still available before the destructive, receipt-producing step.
        for phase in sorted(PHASES):
            for label in ("native", "allocations", "refusal", "oracle"):
                binary = ROOT.parent / BIN_DIR / f"{phase}-{label}"
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
        str(ROOT.parent / TARGET_DIR),
        str(ROOT.parent / BIN_DIR),
        str(ROOT.parent / PROFILE_DIR),
        str(P / "marker-control.pptx"),
    }
    if paths != required:
        stop(f"cleanup receipt path set mismatch: extra={sorted(paths - required)} missing={sorted(required - paths)}")
    workspace_lock = ROOT / "Cargo.lock"
    if cleanup.get("workspace_cargo_lock_preserved_sha256") != sha(workspace_lock):
        stop("cleanup changed or failed to bind workspace Cargo.lock")


def main() -> None:
    base = read("baseline.json")
    rejection = read("rejection.json") if (P / "rejection.json").exists() else None
    for name, digest in base["constraints_sha256"].items():
        if sha(ROOT / name) != digest:
            stop(f"constraint hash changed: {name}")
    for name, digest in base["build_inputs_sha256"].items():
        if sha(ROOT / name) != digest:
            stop(f"build-input hash changed: {name}")
    current = source_map()
    if len(current) != len(base["source_sha256"]):
        stop(f"source file census changed: {len(current)} vs {len(base['source_sha256'])}")
    if rejection is None:
        candidate_source = validate_source_change(base, current)
        final_source = None
    else:
        candidate_source, final_source = validate_rejection_source_change(base, current, rejection)
    validate_script_syntax()
    validate_control()
    bindings = validate_builds(base, candidate_source)
    validate_refusal_builds(base, bindings, candidate_source)
    validate_oracle_builds(base, candidate_source)
    validate_native(bindings)
    validate_allocations(bindings)
    validate_refusal()
    validate_followup_packet()
    validate_profiles_and_sizes(bindings)
    validate_assembly(bindings)
    validate_processor_assembly(bindings)
    validate_marker_controls()
    validate_focused_tests(base, candidate_source, final_source)
    validate_oracle_runs()
    validate_oracle_controls()
    validate_gates(candidate_source)
    rerun_deterministic()
    validate_cleanup()
    disposition = "rejected candidate with retained tests" if rejection is not None else "candidate evidence"
    print(f"PASS: 0702 {disposition}, constraints, source/probe/build bindings, raw matrices, oracle parity, profiles, assembly, focused tests, gates, deterministic summaries and cleanup")


if __name__ == "__main__":
    main()
