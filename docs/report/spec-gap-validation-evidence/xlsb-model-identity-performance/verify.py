#!/usr/bin/env python3
"""Verify bounded smoke receipts without claiming sealed measurements."""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
from pathlib import Path


LANES = {
    "neutral_open_tiny",
    "host_open_tiny",
    "neutral_open_relationship",
    "host_stage_noop_tiny",
    "host_stage_rename_relationship",
    "host_commit_rename_relationship",
    "host_save_reopen_relationship",
    "host_inverse_relationship",
    "host_exact_cap_relationship",
    "host_refusal_opaque",
    "host_refusal_limit",
}
REFUSALS = {"host_refusal_opaque", "host_refusal_limit"}
NEUTRAL = {"neutral_open_tiny", "neutral_open_relationship"}
EXACT_PRESERVATION = {
    "host_open_tiny",
    "host_stage_noop_tiny",
    "host_inverse_relationship",
    "host_refusal_opaque",
    "host_refusal_limit",
}
OPAQUE_ERROR = (
    "Invalid format: XLDM outer identity proof failed: "
    "MS-XLDM storage must contain at least three complete 4096-byte pages"
)
LIMIT_ERROR = "Invalid format: Data Model rewritten workbook bytes limit exceeded"
SOURCE_MANIFEST_FORMAT = "xlsb-model-identity-cargo-source-closure-v3"


def fail(message: str) -> None:
    raise SystemExit(f"smoke verification failed: {message}")


def load(path: Path) -> dict:
    try:
        value = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        fail(f"{path}: {error}")
    if not isinstance(value, dict):
        fail(f"{path}: receipt is not an object")
    return value


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def resolve(root: Path, shown: str) -> Path:
    path = Path(shown)
    return path if path.is_absolute() else root / path


def committed_blob_sha256(root: Path, commit: str, shown: str) -> str:
    try:
        committed = subprocess.check_output(
            [
                "git",
                "--no-replace-objects",
                "-C",
                str(root),
                "cat-file",
                "blob",
                f"{commit}:{shown}",
            ]
        )
    except subprocess.CalledProcessError as error:
        raise AssertionError(f"cannot read committed source input blob: {shown}") from error
    return hashlib.sha256(committed).hexdigest()


def verify_git_snapshot(paths: list[Path], root: Path, commit: str) -> None:
    """Verify source bytes against the pinned commit, independent of Git status."""

    root = root.resolve()
    relative: list[str] = []
    for path in sorted({path.resolve() for path in paths}, key=str):
        if not path.is_file():
            fail(f"retained Git input missing: {path}")
        try:
            relative.append(path.relative_to(root).as_posix())
        except ValueError as error:
            raise AssertionError(
                f"non-Git local input has no retained source snapshot: {path}"
            ) from error
    if not relative:
        return
    current = subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), "rev-parse", "HEAD"],
        text=True,
    ).strip()
    if current != commit:
        raise AssertionError(f"Git snapshot commit changed: {commit} -> {current}")
    missing: list[str] = []
    for shown in relative:
        tracked = subprocess.run(
            [
                "git",
                "--no-replace-objects",
                "-C",
                str(root),
                "ls-files",
                "--error-unmatch",
                "--",
                shown,
            ],
            capture_output=True,
            text=True,
            check=False,
        )
        if tracked.returncode != 0:
            missing.append(shown)
    if missing:
        raise AssertionError(
            "local Git inputs are not tracked: " + ", ".join(sorted(missing))
        )
    changed = [
        shown
        for shown in relative
        if sha256(root / shown) != committed_blob_sha256(root, commit, shown)
    ]
    if changed:
        raise AssertionError(
            "local Git inputs differ from committed snapshot: "
            + ", ".join(sorted(set(changed)))
        )


def verify_source_manifest(
    path: Path, root: Path, *, require_transitive: bool = True
) -> tuple[int, str, str]:
    lines = path.read_text().splitlines()
    if not lines or lines[0] != f"format={SOURCE_MANIFEST_FORMAT}":
        fail("source manifest format changed")
    git_commit = None
    git_head = None
    metadata_hash = None
    packages = {}
    files = {}
    extras = []
    for line in lines[1:]:
        parts = line.split("\t")
        if line.startswith("git_commit="):
            if git_commit is not None:
                fail("source manifest Git commit is duplicated")
            git_commit = line.split("=", 1)[1]
            if len(git_commit) != 40 or any(
                character not in "0123456789abcdef" for character in git_commit
            ):
                fail("source manifest Git commit is malformed")
        elif line.startswith("git_head="):
            if git_head is not None:
                fail("source manifest Git head is duplicated")
            git_head = line.split("=", 1)[1]
            if len(git_head) != 40 or any(
                character not in "0123456789abcdef" for character in git_head
            ):
                fail("source manifest Git head is malformed")
        elif line.startswith("metadata_sha256="):
            if metadata_hash is not None:
                fail("source manifest metadata hash is duplicated")
            metadata_hash = line.split("=", 1)[1]
            if len(metadata_hash) != 64:
                fail("source manifest metadata hash is malformed")
        elif line.startswith("package="):
            if len(parts) != 7:
                fail(f"malformed source package line: {line}")
            key = (parts[0][len("package=") :], parts[1], parts[3])
            if key in packages:
                fail(f"duplicate source package line: {key}")
            try:
                count = int(parts[5])
            except ValueError:
                fail(f"source package file count is malformed: {line}")
            if count <= 0 or len(parts[4]) != 64 or len(parts[6]) != 64:
                fail(f"source package digest/count is malformed: {line}")
            packages[key] = (parts[2], count, parts[6])
        elif line.startswith("file="):
            if len(parts) != 5:
                fail(f"malformed source file line: {line}")
            key = (parts[0][len("file=") :], parts[1], parts[2])
            files.setdefault(key, []).append((parts[3], parts[4]))
        elif line.startswith("extra=\t"):
            if len(parts) != 3 or len(parts[2]) != 64:
                fail(f"malformed source extra line: {line}")
            extras.append((parts[1], parts[2]))
        else:
            fail(f"unknown source manifest line: {line}")
    if git_commit is None or git_head is None or metadata_hash is None:
        fail("source manifest provenance is incomplete")
    if git_head != git_commit:
        fail("source manifest Git head differs from pinned commit")
    if not packages or set(packages) != set(files):
        fail("source package/file manifest sets differ")
    if require_transitive and len(packages) < 100:
        fail(f"source manifest is not transitively complete: only {len(packages)} packages")
    checked = 0
    local_inputs = []
    for key, (source, expected_count, expected_tree) in packages.items():
        entries = files[key]
        if len(entries) != expected_count:
            fail(f"source package file count changed: {key}")
        tree_payload = "\n".join(f"{shown}\t{digest}" for shown, digest in entries)
        if hashlib.sha256(tree_payload.encode()).hexdigest() != expected_tree:
            fail(f"source package tree changed: {key}")
        for shown, expected in entries:
            current = resolve(root, shown)
            if not current.is_file() or sha256(current) != expected:
                fail(f"source manifest input changed or disappeared: {shown}")
            if source == "path":
                local_inputs.append(current)
            checked += 1
    for shown, expected in extras:
        current = resolve(root, shown)
        if not current.is_file() or sha256(current) != expected:
            fail(f"source manifest extra changed or disappeared: {shown}")
        local_inputs.append(current)
        checked += 1
    if require_transitive and checked < 1000:
        fail(f"source manifest checked too few files: {checked}")
    try:
        verify_git_snapshot(local_inputs, root, git_commit)
    except (AssertionError, OSError) as error:
        fail(str(error))
    return checked, git_commit, metadata_hash


def verify_allocator(sample: dict, lane: str) -> None:
    required = (
        "live_before",
        "live_after",
        "direct_allocated_bytes",
        "realloc_old_bytes",
        "realloc_new_bytes",
        "deallocated_bytes",
        "requested_alloc_bytes",
        "peak_live_delta",
        "alloc_balance_ok",
        "alloc_invalid",
        "allocation_failed",
    )
    for key in required:
        if key not in sample:
            fail(f"{lane}: missing allocator field {key}")
    expected_requested = sample["direct_allocated_bytes"] + sample["realloc_new_bytes"]
    if sample["requested_alloc_bytes"] != expected_requested:
        fail(f"{lane}: requested allocator bytes equation failed")
    expected_live = (
        sample["live_before"]
        + sample["direct_allocated_bytes"]
        + sample["realloc_new_bytes"]
        - sample["realloc_old_bytes"]
        - sample["deallocated_bytes"]
    )
    if sample["live_after"] != expected_live:
        fail(f"{lane}: live allocator equation failed")
    if not sample["alloc_balance_ok"] or sample["alloc_invalid"]:
        fail(f"{lane}: allocator balance/validity failed")
    if sample["allocation_failed"] != 0:
        fail(f"{lane}: allocator reported a failed allocation")


def verify_time_file(path: Path, lane: str) -> None:
    if not path.exists():
        fail(f"{lane}: missing time sidecar")
    text = path.read_text()
    if "Maximum resident set size (kbytes):" not in text:
        fail(f"{lane}: time sidecar has no RSS field")
    if "Exit status: 0" not in text:
        fail(f"{lane}: process exit status was not zero")


def verify_binary_receipts(results: Path) -> str:
    before_lines = (results / "binary.sha256").read_text().splitlines()
    after_lines = (results / "binary-after.sha256").read_text().splitlines()
    if len(before_lines) != 1 or before_lines != after_lines:
        fail("profile executable changed during smoke")
    fields = before_lines[0].split()
    if len(fields) != 2:
        fail("binary digest receipt is malformed")
    digest, shown = fields
    if len(digest) != 64 or any(char not in "0123456789abcdef" for char in digest):
        fail("binary digest is malformed")
    binary = Path(shown)
    if not binary.is_absolute() or binary.name != "xlsb-model-identity-profile":
        fail("binary path is malformed")
    build = (results / "build-provenance.txt").read_text()
    if f"binary={shown}" not in build or f"{digest}  {shown}" not in build:
        fail("build provenance does not bind the binary digest")
    if binary.is_file() and sha256(binary) != digest:
        fail("live smoke binary hash changed")
    return digest


def verify_provenance(
    results: Path,
    root: Path,
    manifest: dict,
    manifest_commit: str,
    metadata_hash: str,
    source_manifest_hash: str,
    binary_digest: str,
) -> None:
    provenance = (results / "provenance.txt").read_text().splitlines()
    values = {}
    for line in provenance:
        if "=" in line:
            key, value = line.split("=", 1)
            values[key] = value
    current = subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), "rev-parse", "HEAD"],
        text=True,
    ).strip()
    if values.get("git_head") != current or values.get("git_head") != manifest_commit:
        fail("provenance Git head does not match the committed source manifest")
    expected_host = manifest.get("host_feature_commit")
    if expected_host and values.get("git_baseline") != expected_host:
        fail("provenance host implementation pin does not match corpus manifest")
    if values.get("neutral_baseline") != manifest.get("neutral_baseline_commit"):
        fail("provenance neutral baseline does not match corpus manifest")
    if values.get("git_status_relevant") != "":
        fail("relevant source tree was dirty or provenance was incomplete")
    metadata_before = results / "metadata-before.json"
    metadata_after = results / "metadata-after.json"
    if metadata_before.read_bytes() != metadata_after.read_bytes():
        fail("Cargo metadata changed during smoke")
    if sha256(metadata_before) != metadata_hash:
        fail("source manifest metadata hash does not bind metadata-before.json")
    if values.get("metadata_before_sha256") != metadata_hash or values.get(
        "metadata_after_sha256"
    ) != metadata_hash:
        fail("provenance metadata hashes do not match source manifest")
    if values.get("source_manifest_before_sha256") != source_manifest_hash or values.get(
        "source_manifest_after_sha256"
    ) != source_manifest_hash:
        fail("provenance source-manifest hashes do not match receipts")
    if values.get("binary_sha256") != binary_digest or values.get(
        "binary_after_sha256"
    ) != binary_digest:
        fail("provenance binary hashes do not match receipts")


def verify_phases(sample: dict, lane: str) -> None:
    phases = sample.get("phases")
    if not isinstance(phases, dict):
        fail(f"{lane}: phase receipt is missing")
    expected = {
        "open_ns",
        "stage_ns",
        "commit_ns",
        "save_ns",
        "reopen_ns",
        "inverse_ns",
        "validation_ns",
    }
    if set(phases) != expected:
        fail(f"{lane}: phase receipt fields are not explicit: {sorted(phases)}")
    for name, value in phases.items():
        if value is not None and (type(value) is not int or value <= 0):
            fail(f"{lane}: phase {name} is not a positive integer or null")
    required = {
        "neutral_open_tiny": {"open_ns"},
        "host_open_tiny": {"open_ns", "validation_ns"},
        "neutral_open_relationship": {"open_ns"},
        "host_stage_noop_tiny": {"stage_ns", "commit_ns", "validation_ns"},
        "host_stage_rename_relationship": {"stage_ns", "commit_ns", "validation_ns"},
        "host_commit_rename_relationship": {"commit_ns", "validation_ns"},
        "host_save_reopen_relationship": {"save_ns", "reopen_ns", "validation_ns"},
        "host_inverse_relationship": {"inverse_ns", "save_ns", "validation_ns"},
        "host_exact_cap_relationship": {"commit_ns", "validation_ns"},
        "host_refusal_opaque": {"stage_ns", "validation_ns"},
        "host_refusal_limit": {"commit_ns", "validation_ns"},
    }[lane]
    for name in required:
        if phases[name] is None:
            fail(f"{lane}: required phase {name} was not measured")


def verify_preservation(sample: dict, lane: str) -> None:
    preservation = sample.get("preservation")
    if lane in NEUTRAL:
        if preservation is not None:
            fail(f"{lane}: neutral lane unexpectedly claims host preservation")
        return
    if not isinstance(preservation, dict):
        fail(f"{lane}: preservation manifest was not reported")
    booleans = (
        "all_parts_equal",
        "unchanged_parts_equal",
        "relationships_equal",
        "content_types_equal",
        "inner_all_equal",
        "inner_unchanged_equal",
    )
    for field in booleans:
        if type(preservation.get(field)) is not bool:
            fail(f"{lane}: preservation field {field} is not boolean")
    for field in ("relationship_owner_count", "content_types_bytes"):
        value = preservation.get(field)
        if type(value) is not int or value <= 0:
            fail(f"{lane}: preservation count {field} is not positive")
    inner_count = preservation.get("inner_member_count")
    if type(inner_count) is not int or (lane not in REFUSALS and inner_count <= 0):
        fail(f"{lane}: preservation count inner_member_count is invalid")
    if lane in EXACT_PRESERVATION:
        for field in ("all_parts_equal", "relationships_equal", "content_types_equal", "inner_all_equal"):
            if preservation[field] is not True:
                fail(f"{lane}: exact preservation field {field} failed")
    else:
        for field in (
            "unchanged_parts_equal",
            "relationships_equal",
            "content_types_equal",
            "inner_unchanged_equal",
        ):
            if preservation[field] is not True:
                fail(f"{lane}: changed preservation field {field} failed")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--source-before", type=Path, required=True)
    parser.add_argument("--source-after", type=Path, required=True)
    args = parser.parse_args()

    manifest = load(args.manifest)
    if manifest.get("fixture_kind") != "synthetic_complete_xldm140":
        fail("manifest fixture kind is not complete synthetic XLDM 140")
    if manifest.get("native_acceptance_claim") is not False:
        fail("manifest makes a native acceptance claim")
    if args.source_before.read_bytes() != args.source_after.read_bytes():
        fail("source manifest changed during the run")
    manifest_checked, manifest_commit, metadata_hash = verify_source_manifest(
        args.source_before, args.root.resolve()
    )
    source_manifest_hash = sha256(args.source_before)
    binary_digest = verify_binary_receipts(args.results)
    verify_provenance(
        args.results,
        args.root.resolve(),
        manifest,
        manifest_commit,
        metadata_hash,
        source_manifest_hash,
        binary_digest,
    )

    receipts = {}
    for lane in sorted(LANES):
        path = args.results / f"{lane}.json"
        receipt = load(path)
        receipts[lane] = receipt
        if receipt.get("schema") != "xlsb-model-identity-profile-v1-smoke":
            fail(f"{lane}: wrong schema")
        if receipt.get("fixture_kind") != "synthetic_complete_xldm140":
            fail(f"{lane}: wrong fixture kind")
        if receipt.get("source_backed_api") is not True:
            fail(f"{lane}: source-backed API flag missing")
        if receipt.get("native_acceptance_claim") is not False:
            fail(f"{lane}: native acceptance flag is not false")
        if receipt.get("sample_count") != 1 or receipt.get("warmup") != 0:
            fail(f"{lane}: smoke process does not have one sample and zero warmups")
        if receipt.get("table_count") != 1:
            fail(f"{lane}: smoke fixture is not one table")
        if receipt.get("relationship_count") not in (0, 1):
            fail(f"{lane}: smoke relationship count is outside the contract")
        samples = receipt.get("samples")
        if not isinstance(samples, list) or len(samples) != 1:
            fail(f"{lane}: sample list shape is invalid")
        sample = samples[0]
        if not isinstance(sample, dict):
            fail(f"{lane}: sample is not an object")
        verify_allocator(sample, lane)
        verify_phases(sample, lane)
        verify_preservation(sample, lane)
        verify_time_file(args.results / f"{lane}.time.txt", lane)
        if (args.results / f"{lane}.stderr.log").read_bytes() != b"":
            fail(f"{lane}: process stderr was not empty")
        if sample.get("source_unchanged") is not True:
            fail(f"{lane}: source changed or source gate was unavailable")
        opaque_ok = sample.get("opaque_ok")
        if opaque_ok is False:
            fail(f"{lane}: opaque preservation gate failed")
        if lane not in {"neutral_open_tiny", "neutral_open_relationship"} and opaque_ok is not True:
            fail(f"{lane}: opaque preservation gate was unavailable")
        expected_success = lane not in REFUSALS
        if receipt.get("expected_success") is not expected_success:
            fail(f"{lane}: expected-success field is incorrect")
        if sample.get("actual_success") is not expected_success:
            fail(f"{lane}: actual-success field is incorrect")
        if sample.get("semantic_ok") is not True:
            fail(f"{lane}: semantic gate failed")
        error = sample.get("error")
        if expected_success:
            if error is not None:
                fail(f"{lane}: successful lane has an error")
        else:
            if not isinstance(error, dict) or error.get("typed_match") is not True:
                fail(f"{lane}: refusal has no typed error receipt")
            if sample.get("candidate_bytes") is not None:
                fail(f"{lane}: refusal has candidate bytes")
            if error.get("class") != "invalid_format":
                fail(f"{lane}: refusal class is not exact invalid_format")
            expected_message = OPAQUE_ERROR if lane == "host_refusal_opaque" else LIMIT_ERROR
            if error.get("message") != expected_message:
                fail(f"{lane}: refusal resource/message changed: {error.get('message')!r}")
        if lane == "host_exact_cap_relationship":
            if sample.get("exact_cap_ok") is not True:
                fail(f"{lane}: exact cap case did not succeed")
            if sample.get("one_under_cap_refused") is not True:
                fail(f"{lane}: one-byte-under cap was accepted")

    if receipts["neutral_open_tiny"]["input_sha256"] != receipts["host_open_tiny"]["input_sha256"]:
        fail("neutral and host tiny fixtures differ")
    if receipts["neutral_open_relationship"]["input_sha256"] != receipts[
        "host_stage_rename_relationship"
    ]["input_sha256"]:
        fail("neutral and host relationship fixtures differ")
    print(
        f"verified {len(receipts)} synthetic identity smoke lanes "
        f"({manifest_checked} source inputs, binary {binary_digest})"
    )


if __name__ == "__main__":
    main()
