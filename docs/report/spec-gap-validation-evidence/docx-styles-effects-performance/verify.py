#!/usr/bin/env python3
"""Fail-closed verifier for the bounded stylesWithEffects smoke receipts."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
import zipfile
from pathlib import Path


SOURCE_COMMIT = "d1f299d00e0dd5cc5cd8ddf9811c4b1ad21d1119"
# The historical 52-lane smoke remains bound to SOURCE_COMMIT.  Profile
# receipts use the separately approved attribution baseline so retaining the
# smoke does not rewrite its source identity.
PROFILE_SOURCE_COMMIT = "d000d977b99e03f8542c7dae74acf767a91b1feb"
MANIFEST_FORMAT = "docx-styles-effects-cargo-source-closure-v1"
RECEIPT_SCHEMA = "docx-styles-effects-smoke-v1"
U64_MAX = (1 << 64) - 1
PHASE_FIELDS = (
    "capture_ns",
    "snapshot_ns",
    "stage_ns",
    "commit_ns",
    "publish_ns",
    "reopen_ns",
    "inverse_reopen_ns",
    "inverse_ns",
    "projection_ns",
    "opaque_ns",
    "graph_ns",
    "readback_ns",
    "validation_ns",
)
COUNTER_FIELDS = (
    "elapsed_ns",
    "output_bytes",
    "direct_allocated_bytes",
    "realloc_old_bytes",
    "realloc_new_bytes",
    "deallocated_bytes",
    "requested_alloc_bytes",
    "live_before",
    "live_after",
    "peak_live_delta",
    "allocation_calls",
    "reallocation_calls",
    "deallocation_calls",
    "allocation_failed",
)
BOOL_FIELDS = (
    "actual_success",
    "ingress_refusal",
    "no_output_ok",
    "semantic_ok",
    "opaque_ok",
    "exact_inverse_ok",
    "alloc_balance_ok",
    "alloc_invalid",
)
METRIC_FIELDS = (
    "parts",
    "total_part_bytes",
    "total_relationships",
    "relationship_parts",
    "relationship_graph_nodes",
    "relationship_xml_bytes",
    "relationship_xml_events",
)
EXPECTED_REFUSALS = {
    "add_glossary_missing",
    "signed_changed",
    "stale_patch_main",
    "malformed_duplicate_owner",
    "malformed_third_orphan",
    "malformed_external",
    "malformed_wrong_content_type",
    "malformed_outbound",
    "malformed_shared_inbound",
    "malformed_root",
    "malformed_namespace",
    "malformed_opaque_xml",
    "malformed_unbound_descendant",
    "malformed_invalid_qname",
    "malformed_raw_attribute",
    "malformed_raw_text",
    "malformed_control",
    "malformed_invalid_char_ref",
    "malformed_empty_prefix",
    "malformed_reserved_xml_uri",
    "malformed_xml_version",
    "malformed_xml_events",
    "malformed_xml_depth",
}
EXPECTED_VARIANTS = {
    "add_glossary_missing": "DocxError::PartNotFound",
    "signed_changed": "DocxError::UnsafeEdit",
    "stale_patch_main": "DocxError::InvalidFormat",
    "malformed_opaque_xml": "DocxError::Opc::SourceBackedOverlayUnavailable",
    "malformed_xml_events": "DocxError::Opc::ReadLimit",
    "malformed_xml_depth": "DocxError::Opc::ReadLimit",
}
EXPECTED_RESOURCES = {
    "malformed_xml_events": "XmlEvents",
    "malformed_xml_depth": "XmlDepth",
}
SYNTHETIC_NATIVE_LANES = {
    "malformed_duplicate_owner",
    "malformed_third_orphan",
    "malformed_external",
    "malformed_wrong_content_type",
    "malformed_outbound",
    "malformed_shared_inbound",
    "malformed_root",
    "malformed_namespace",
    "malformed_opaque_xml",
    "malformed_unbound_descendant",
    "malformed_invalid_qname",
    "malformed_raw_attribute",
    "malformed_raw_text",
    "malformed_control",
    "malformed_invalid_char_ref",
    "malformed_empty_prefix",
    "malformed_reserved_xml_uri",
    "malformed_xml_version",
}
CAP_RESOURCES = {
    "cap_parts": "Parts",
    "cap_total_part_bytes": "TotalPartBytes",
    "cap_total_relationships": "TotalRelationships",
    "cap_total_relationship_xml_events": "TotalRelationshipXmlEvents",
    "cap_total_relationship_xml_bytes": "TotalRelationshipXmlBytes",
    "cap_relationship_parts": "RelationshipParts",
    "cap_relationship_graph_nodes": "RelationshipGraphNodes",
}
NATIVE_FIXTURES = {
    "Bug54849.docx": {
        "bytes": 27566,
        "sha256": "f54182713ea5ce5d77b9593d3d9d24e645460043cec0b40ef59c932385f084d3",
    },
    "ms-office-2010-signed.docx": {
        "bytes": 16142,
        "sha256": "bc55c0362722818823a6dd95f8e0ca9869e179ace972a0915241feb4677bde5f",
    },
    "ComplexNumberedLists.docx": {
        "bytes": 14458,
        "sha256": "297a085a7d433af2eeee7661e8db21539452cb585096484774a1e9f5f258b0b6",
    },
    "testGlossary.docx": {
        "bytes": 25741,
        "sha256": "8ccd581d8f0ae102b220228ad26b3974821a7ce8e3ff7df4b78f7da8a0d06ed9",
    },
}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def required(mapping: dict[str, object], key: str, context: str) -> object:
    require(key in mapping, f"{context} missing {key}")
    return mapping[key]


def unsigned(value: object, label: str, *, positive: bool = False) -> int:
    require(type(value) is int, f"{label} must be an integer")
    require(0 <= value <= U64_MAX, f"{label} is outside u64")
    require(not positive or value > 0, f"{label} must be positive")
    return value


def boolean(value: object, label: str) -> bool:
    require(type(value) is bool, f"{label} must be a boolean")
    return value


def text(value: object, label: str) -> str:
    require(type(value) is str, f"{label} must be text")
    return value


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def git_blob_sha256(root: Path, commit: str, shown: str) -> str | None:
    result = subprocess.run(
        ["git", "--no-replace-objects", "-C", str(root), "cat-file", "-e", f"{commit}:{shown}"],
        check=False,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    if result.returncode != 0:
        return None
    data = subprocess.check_output(
        ["git", "--no-replace-objects", "-C", str(root), "cat-file", "blob", f"{commit}:{shown}"]
    )
    return sha256_bytes(data)


def retained_input(root: Path, shown: str, evidence_root: str, retained_evidence: Path) -> Path:
    staged = (root / shown).resolve()
    if staged.is_file():
        return staged
    prefix = evidence_root + "/"
    if shown.startswith(prefix):
        retained = (retained_evidence / shown[len(prefix) :]).resolve()
        if retained.is_file():
            return retained
    return staged


def is_evidence_path(shown: str, evidence_root: str) -> bool:
    return shown.startswith(evidence_root + "/")


def manifest_blob_commit(
    shown: str,
    evidence_root: str,
    manifest_head: str,
    source_commit: str | None = None,
) -> str:
    if source_commit is None:
        source_commit = SOURCE_COMMIT
    return manifest_head if is_evidence_path(shown, evidence_root) else source_commit


def verify_manifest_path(
    root: Path,
    shown: str,
    expected: str,
    evidence_root: str,
    manifest_head: str,
    retained_evidence: Path,
    context: str,
    source_commit: str | None = None,
) -> None:
    if source_commit is None:
        source_commit = SOURCE_COMMIT
    require(
        not Path(shown).is_absolute()
        and all(part not in ("", ".", "..") for part in Path(shown).parts),
        f"unsafe source path: {shown}",
    )
    current = retained_input(root, shown, evidence_root, retained_evidence)
    require(current.is_file(), f"{context} missing: {shown}")
    require(sha256(current) == expected, f"{context} changed: {shown}")
    commit = manifest_blob_commit(shown, evidence_root, manifest_head, source_commit)
    committed = git_blob_sha256(root, commit, shown)
    require(committed is not None, f"{context} is absent at bound Git head: {shown}")
    require(committed == expected, f"{context} Git blob changed: {shown}")


def verify_manifest(
    path: Path,
    root: Path,
    metadata: Path,
    retained_evidence: Path,
    *,
    source_commit: str | None = None,
) -> dict[str, object]:
    if source_commit is None:
        source_commit = SOURCE_COMMIT
    lines = path.read_text().splitlines()
    require(lines and lines[0] == f"format={MANIFEST_FORMAT}", "source manifest format changed")
    fields: dict[str, str] = {}
    packages: list[list[str]] = []
    files: list[list[str]] = []
    extras: list[list[str]] = []
    for line in lines[1:]:
        if line.startswith("package="):
            packages.append(line[len("package=") :].split("\t"))
        elif line.startswith("file="):
            files.append(line[len("file=") :].split("\t"))
        elif line.startswith("extra="):
            extras.append(line[len("extra=") :].split("\t"))
        elif "=" in line:
            key, value = line.split("=", 1)
            fields[key] = value
    require(fields.get("source_commit") == source_commit, "source commit changed")
    manifest_head = fields.get("git_head")
    require(
        manifest_head is not None
        and len(manifest_head) == 40
        and all(character in "0123456789abcdef" for character in manifest_head),
        "manifest Git head is missing or malformed",
    )
    evidence_root = fields.get("evidence_root")
    try:
        expected_evidence_root = retained_evidence.resolve().relative_to(root.resolve()).as_posix()
    except ValueError as error:
        raise AssertionError("retained evidence is outside the source checkout") from error
    require(evidence_root == expected_evidence_root, "staged evidence root changed")
    require(fields.get("metadata_sha256") == sha256(metadata), "metadata receipt hash changed")
    require(files, "source manifest has no package files")
    require(packages, "source manifest has no local packages")

    seen: set[str] = set()
    grouped: dict[tuple[str, str, str], list[tuple[str, str]]] = {}
    for record in files:
        require(len(record) == 5, f"source file record shape changed: {record}")
        package_name, version, manifest, shown, expected = record
        require(len(expected) == 64 and all(c in "0123456789abcdef" for c in expected), f"bad source hash: {shown}")
        require(shown not in seen, f"source path listed twice: {shown}")
        seen.add(shown)
        verify_manifest_path(
            root,
            shown,
            expected,
            evidence_root,
            manifest_head,
            retained_evidence,
            "source input",
            source_commit,
        )
        grouped.setdefault((package_name, version, manifest), []).append((shown, expected))

    for record in extras:
        require(len(record) == 3, f"source extra record shape changed: {record}")
        _, shown, expected = record
        require(shown not in seen, f"source path listed twice: {shown}")
        seen.add(shown)
        verify_manifest_path(
            root,
            shown,
            expected,
            evidence_root,
            manifest_head,
            retained_evidence,
            "source extra",
            source_commit,
        )

    package_keys: set[tuple[str, str, str]] = set()
    for record in packages:
        require(len(record) == 7, f"source package record shape changed: {record}")
        name, version, source, manifest, manifest_sha, count_text, tree_expected = record
        key = (name, version, manifest)
        require(key not in package_keys, f"source package listed twice: {key}")
        package_keys.add(key)
        require(source == "path", f"unexpected source kind: {source}")
        unsigned(int(count_text), f"source file count for {name}")
        verify_manifest_path(
            root,
            manifest,
            manifest_sha,
            evidence_root,
            manifest_head,
            retained_evidence,
            "package manifest",
            source_commit,
        )
        package_files = sorted(grouped.get(key, []))
        require(package_files, f"package has no file records: {name}")
        file_lines = [f"{shown}\t{digest}" for shown, digest in package_files]
        tree_actual = sha256_bytes("\n".join(file_lines).encode())
        require(tree_actual == tree_expected, f"package source tree changed: {name}")
        require(len(package_files) == int(count_text), f"package source count changed: {name}")

    required_paths = {
        f"{evidence_root}/harness/Cargo.toml",
        f"{evidence_root}/harness/Cargo.lock",
        f"{evidence_root}/harness/adapter.rs",
        f"{evidence_root}/harness/main.rs",
        f"{evidence_root}/harness/support.rs",
        f"{evidence_root}/source_manifest.py",
        f"{evidence_root}/test_source_snapshot.py",
        f"{evidence_root}/test_verify.py",
        f"{evidence_root}/verify.py",
        f"{evidence_root}/requirements.md",
        f"{evidence_root}/corpus-manifest.json",
        "crates/litchi-docx/src/styles/effects.rs",
        "crates/litchi-docx/src/package/package/styles_with_effects.rs",
        "crates/litchi-docx/tests/styles_with_effects.rs",
        "crates/litchi-opc/src/phys_pkg.rs",
        "crates/litchi-opc/src/limits.rs",
    }
    require(required_paths <= seen, "source manifest omits a required production or harness input")
    return {"packages": len(packages), "files": len(files), "extras": len(extras)}


def verify_corpus(path: Path, native_root: Path, evidence: Path) -> dict[str, object]:
    corpus = json.loads(path.read_text())
    require(type(corpus) is dict, "corpus manifest must be an object")
    require(corpus.get("schema") == "docx-styles-effects-corpus-v1", "corpus schema changed")
    require(corpus.get("source_commit") == SOURCE_COMMIT, "corpus source commit changed")
    fixtures = corpus.get("fixtures")
    require(type(fixtures) is list and fixtures, "corpus fixture list is empty")
    names: set[str] = set()
    for fixture in fixtures:
        require(type(fixture) is dict, "corpus fixture record is not an object")
        name = text(required(fixture, "name", "fixture"), "fixture.name")
        require(name not in names, f"duplicate fixture: {name}")
        names.add(name)
        source = native_root / text(required(fixture, "source_path", name), f"{name}.source_path")
        copy = evidence / text(required(fixture, "copy_path", name), f"{name}.copy_path")
        expected_bytes = unsigned(required(fixture, "bytes", name), f"{name}.bytes")
        expected_hash = text(required(fixture, "sha256", name), f"{name}.sha256")
        require(source.is_file() and copy.is_file(), f"fixture missing: {name}")
        require(source.stat().st_size == expected_bytes and copy.stat().st_size == expected_bytes, f"fixture size changed: {name}")
        source_hash = sha256(source)
        copy_hash = sha256(copy)
        require(source_hash == expected_hash and copy_hash == expected_hash, f"fixture hash changed: {name} source={source_hash} copy={copy_hash} expected={expected_hash}")
        members = fixture.get("members")
        require(type(members) is list and members, f"fixture member list is empty: {name}")
        with zipfile.ZipFile(copy) as archive:
            for member in members:
                member_name = text(required(member, "name", name), f"{name}.member.name")
                data = archive.read(member_name)
                member_bytes = unsigned(required(member, "bytes", member_name), f"{name}.{member_name}.bytes")
                member_hash = text(required(member, "sha256", member_name), f"{name}.{member_name}.sha256")
                require(len(data) == member_bytes and sha256_bytes(data) == member_hash, f"fixture member changed: {name}:{member_name}")
    lanes = corpus.get("lanes")
    require(type(lanes) is list and lanes and all(type(lane) is str for lane in lanes), "corpus lane list changed")
    require(len(set(lanes)) == len(lanes), "corpus lane list contains duplicates")
    refusals = corpus.get("expected_refusals")
    require(type(refusals) is list and set(refusals) == EXPECTED_REFUSALS, "corpus refusal list changed")
    return {"fixtures": len(fixtures), "lanes": lanes}


def verify_time_sidecar(path: Path, stderr: Path) -> None:
    require(path.is_file(), f"RSS sidecar missing: {path}")
    require(stderr.is_file() and stderr.read_bytes() == b"", f"unexpected lane stderr: {stderr}")
    content = path.read_text()
    statuses = [line.strip() for line in content.splitlines() if line.strip().startswith("Exit status:")]
    require(statuses == ["Exit status: 0"], f"process status is not zero: {path}")
    marker = "Maximum resident set size (kbytes):"
    values = [line.split(":", 1)[1].strip() for line in content.splitlines() if line.lstrip().startswith(marker)]
    require(len(values) == 1, f"RSS marker missing or duplicated: {path}")
    value = values[0]
    require(value.isascii() and value.isdecimal() and int(value) > 0, f"RSS is not positive decimal: {path}")


def receipt_key_values(path: Path, label: str, *, skip_hash_line: bool = False) -> dict[str, str]:
    require(path.is_file(), f"{label} receipt is missing: {path}")
    values: dict[str, str] = {}
    for line in path.read_text().splitlines():
        if not line.strip():
            continue
        if skip_hash_line and re.fullmatch(r"[0-9a-f]{64}\s+\S+", line):
            continue
        require("=" in line, f"{label} receipt has an unkeyed line: {path}")
        key, value = line.split("=", 1)
        require(key and key not in values, f"{label} receipt has duplicate key: {path}:{key}")
        values[key] = value
    require(values, f"{label} receipt is empty: {path}")
    return values


def hash_receipt(path: Path, label: str) -> tuple[str, str]:
    require(path.is_file(), f"{label} receipt is missing: {path}")
    lines = [line.strip() for line in path.read_text().splitlines() if line.strip()]
    require(len(lines) == 1, f"{label} receipt must contain one hash line: {path}")
    fields = lines[0].split()
    require(len(fields) == 2, f"{label} receipt shape changed: {path}")
    digest, shown = fields
    require(len(digest) == 64 and all(character in "0123456789abcdef" for character in digest), f"{label} hash malformed: {path}")
    require(Path(shown).is_absolute(), f"{label} path is not absolute: {path}")
    return digest, shown


def verify_build_receipts(
    results: Path,
    current_head: str,
    manifest_before: Path,
    manifest_after: Path,
    metadata_before: Path,
    metadata_after: Path,
    lanes: list[str],
) -> None:
    build_log = results / "smoke-build.log"
    require(build_log.is_file() and build_log.read_text().strip(), f"build log is missing or empty: {build_log}")
    require("Finished " in build_log.read_text(), f"build log has no successful cargo completion marker: {build_log}")

    binary_digest, binary_path = hash_receipt(results / "smoke-binary.sha256", "binary")
    after_digest, after_path = hash_receipt(results / "smoke-binary-after.sha256", "post-build binary")
    require((binary_digest, binary_path) == (after_digest, after_path), "binary changed after lane execution")
    binary = Path(binary_path)

    provenance_path = results / "smoke-build-provenance.txt"
    provenance_lines = provenance_path.read_text().splitlines() if provenance_path.is_file() else []
    raw_hashes = [line.split() for line in provenance_lines if re.fullmatch(r"[0-9a-f]{64}\s+\S+", line)]
    require(len(raw_hashes) == 1 and raw_hashes[0] == [binary_digest, binary_path], f"build provenance binary hash changed: {provenance_path}")
    provenance = receipt_key_values(provenance_path, "build provenance", skip_hash_line=True)
    required_provenance = {
        "binary",
        "source_commit",
        "git_head",
        "git_status",
        "rustc",
        "cargo",
        "target",
        "allocator",
        "rss",
        "mode",
    }
    require(set(provenance) == required_provenance, f"build provenance keys changed: {provenance_path}")
    require(provenance["binary"] == binary_path, f"build provenance binary path changed: {provenance_path}")
    require(provenance["source_commit"] == SOURCE_COMMIT, f"build provenance source changed: {provenance_path}")
    require(provenance["git_head"] == current_head, f"build provenance Git head changed: {provenance_path}")
    require(provenance["git_status"] == "clean", f"build provenance checkout was not clean: {provenance_path}")
    require(provenance["rustc"] and provenance["cargo"] and provenance["target"], f"build tool provenance is incomplete: {provenance_path}")
    target = Path(provenance["target"])
    require(target.is_absolute(), f"build target path is not absolute: {provenance_path}")
    if binary.is_file():
        require(sha256(binary) == binary_digest, f"built binary hash changed: {binary}")
    else:
        require(not target.exists(), f"built binary is missing while its disposable target exists: {binary}")
    require("CountingAllocator" in provenance["allocator"], f"allocator provenance is missing: {provenance_path}")
    require("/usr/bin/time -v" in provenance["rss"], f"RSS provenance is missing: {provenance_path}")
    require(provenance["mode"] == "smoke processes=1 warmup=0 samples=1", f"build mode changed: {provenance_path}")

    source_provenance = receipt_key_values(results / "smoke-source-provenance.txt", "source provenance")
    require(
        set(source_provenance) == {"source_manifest_sha256", "metadata_before_sha256", "metadata_after_sha256", "git_head"},
        "source provenance keys changed",
    )
    require(source_provenance["source_manifest_sha256"] == sha256(manifest_before), "source manifest provenance hash changed")
    require(source_provenance["metadata_before_sha256"] == sha256(metadata_before), "metadata-before provenance hash changed")
    require(source_provenance["metadata_after_sha256"] == sha256(metadata_after), "metadata-after provenance hash changed")
    require(source_provenance["git_head"] == current_head, "source provenance Git head changed")

    commands_path = results / "smoke-commands.txt"
    require(commands_path.is_file(), f"command receipt is missing: {commands_path}")
    command_lines = commands_path.read_text().splitlines()
    require(any(line.startswith("command=cargo metadata ") for line in command_lines), "cargo metadata command is missing")
    require(any(line.startswith("command=python3 ") for line in command_lines), "source manifest command is missing")
    require(any(line.startswith("command=cargo build ") for line in command_lines), "cargo build command is missing")
    environments = [line for line in command_lines if line.startswith("environment=")]
    require(len(environments) == 1, "environment receipt is missing or duplicated")
    environment = environments[0]
    for marker in ("CARGO_TARGET_DIR=", "CARGO_INCREMENTAL=0", "LC_ALL=C", "RUSTFLAGS=unset", "CARGO_ENCODED_RUSTFLAGS=unset", "RUSTC_BOOTSTRAP=unset", "RUSTDOCFLAGS=unset"):
        require(marker in environment, f"environment receipt is missing {marker}")
    run_lines = [line for line in command_lines if line.startswith("run=")]
    require(len(run_lines) == len(lanes), "lane command receipt count changed")
    observed: list[str] = []
    for line in run_lines:
        match = re.search(r"(?:^|\s)--lane\s+([A-Za-z0-9_]+)(?:\s|$)", line)
        require(match is not None, f"lane command has no lane name: {line}")
        observed.append(match.group(1))
        require("--warmup 0 --samples 1" in line and "fresh_process=1" in line, f"lane command settings changed: {line}")
    require(set(observed) == set(lanes) and len(set(observed)) == len(lanes), "lane command set changed")


def verify_error(value: object, path: str, expected_variant: str | None = None, resource: str | None = None) -> None:
    require(type(value) is dict, f"typed error missing: {path}")
    error = value
    require(set(error) == {"class", "variant", "message", "typed_match", "resource", "actual", "maximum"}, f"typed error schema changed: {path}")
    variant = text(required(error, "variant", path), f"{path}.variant")
    error_class = text(required(error, "class", path), f"{path}.class")
    require(error_class == variant and variant.startswith("DocxError::"), f"typed error variant is not concrete: {path}")
    if expected_variant is not None:
        require(variant == expected_variant, f"wrong typed error variant at {path}: {variant}")
    require(text(required(error, "message", path), f"{path}.message"), f"empty typed error message: {path}")
    require(boolean(required(error, "typed_match", path), f"{path}.typed_match") is True, f"typed error did not match: {path}")
    actual = required(error, "actual", path)
    maximum = required(error, "maximum", path)
    actual_value = None if actual is None else unsigned(actual, f"{path}.actual")
    maximum_value = None if maximum is None else unsigned(maximum, f"{path}.maximum")
    if resource is None:
        require(required(error, "resource", path) is None and actual_value is None and maximum_value is None, f"unexpected ReadLimit fields at {path}")
    else:
        require(text(required(error, "resource", path), f"{path}.resource") == resource, f"wrong ReadLimit resource at {path}")
        require(actual_value is not None and maximum_value is not None and actual_value > maximum_value, f"invalid ReadLimit boundary at {path}")


def verify_metrics(value: object, path: str) -> None:
    require(type(value) is dict, f"metrics must be an object: {path}")
    require(set(value) == set(METRIC_FIELDS), f"metric keys changed: {path}")
    for key in METRIC_FIELDS:
        unsigned(required(value, key, path), f"{path}.{key}")


def verify_cap_evidence(value: object, path: str, lane: str) -> None:
    require(type(value) is dict, f"existing-owner cap evidence must be an object: {path}")
    evidence = value
    require(
        set(evidence)
        == {
            "applicable",
            "source_metrics",
            "projected_metrics",
            "exact_fit_ok",
            "exact_opaque_ok",
            "source_unrelated_member_digest",
            "exact_unrelated_member_digest",
            "under_refused_ok",
            "commit_stage_checked",
            "refusal",
            "commit_refusal",
        },
        f"existing-owner cap schema changed: {path}",
    )
    applicable = boolean(required(evidence, "applicable", path), f"{path}.applicable")
    verify_metrics(required(evidence, "source_metrics", path), f"{path}.source_metrics")
    verify_metrics(required(evidence, "projected_metrics", path), f"{path}.projected_metrics")
    exact = boolean(required(evidence, "exact_fit_ok", path), f"{path}.exact_fit_ok")
    exact_opaque = boolean(required(evidence, "exact_opaque_ok", path), f"{path}.exact_opaque_ok")
    source_unrelated_digest = required(evidence, "source_unrelated_member_digest", path)
    exact_unrelated_digest = required(evidence, "exact_unrelated_member_digest", path)
    under = boolean(required(evidence, "under_refused_ok", path), f"{path}.under_refused_ok")
    commit_stage_checked = boolean(
        required(evidence, "commit_stage_checked", path),
        f"{path}.commit_stage_checked",
    )
    refusal = required(evidence, "refusal", path)
    commit_refusal = required(evidence, "commit_refusal", path)
    if applicable:
        require(exact and under, f"existing-owner cap boundary failed: {path}")
        require(exact_opaque, f"existing-owner exact-fit opaque members changed: {path}")
        source_digest = text(source_unrelated_digest, f"{path}.source_unrelated_member_digest")
        exact_digest = text(exact_unrelated_digest, f"{path}.exact_unrelated_member_digest")
        require(
            len(source_digest) == 64
            and all(character in "0123456789abcdef" for character in source_digest),
            f"existing-owner source unrelated digest malformed: {path}",
        )
        require(
            len(exact_digest) == 64
            and all(character in "0123456789abcdef" for character in exact_digest),
            f"existing-owner exact unrelated digest malformed: {path}",
        )
        require(source_digest == exact_digest, f"existing-owner unrelated members changed: {path}")
        require(commit_stage_checked, f"existing-owner cap lacked a commit-stage refusal: {path}")
        verify_error(refusal, f"{path}.refusal", "DocxError::Opc::ReadLimit", CAP_RESOURCES[lane])
        if commit_stage_checked:
            verify_error(commit_refusal, f"{path}.commit_refusal", "DocxError::Opc::ReadLimit", CAP_RESOURCES[lane])
        else:
            require(commit_refusal is None, f"unobserved commit refusal has a receipt: {path}")
        projected = evidence["projected_metrics"]
        source = evidence["source_metrics"]
        require(projected != source, f"existing-owner cap did not grow: {path}")
    else:
        require(not exact and not exact_opaque and not under, f"inapplicable existing-owner cap was marked successful: {path}")
        require(source_unrelated_digest is None and exact_unrelated_digest is None, f"inapplicable cap has exact opaque digests: {path}")
        require(refusal is None and commit_refusal is None, f"inapplicable cap has refusal receipts: {path}")


def verify_sample(sample: object, lane: str, expected_success: bool, path: str) -> None:
    require(type(sample) is dict, f"sample must be an object: {path}")
    value = sample
    counters = {key: unsigned(required(value, key, "sample"), f"{path}:{key}") for key in COUNTER_FIELDS}
    for key in BOOL_FIELDS:
        boolean(required(value, key, "sample"), f"{path}:{key}")
    require(counters["elapsed_ns"] > 0, f"elapsed_ns must be positive: {path}")
    require(required(value, "actual_success", "sample") is expected_success, f"success status mismatch: {path}")
    require(required(value, "no_output_ok", "sample") is True, f"no-output policy failed: {path}")
    require(required(value, "semantic_ok", "sample") is True, f"semantic gate failed: {path}")
    require(required(value, "opaque_ok", "sample") is True, f"opaque gate failed: {path}")
    if expected_success:
        require(required(value, "ingress_refusal", "sample") is False, f"successful lane marked as ingress refusal: {path}")
        require(required(value, "exact_inverse_ok", "sample") is True, f"inverse gate failed: {path}")
    else:
        require(required(value, "exact_inverse_ok", "sample") is False, f"refusal inverse field changed: {path}")
    phases = required(value, "phases", "sample")
    require(type(phases) is dict and set(phases) == set(PHASE_FIELDS), f"phase schema changed: {path}")
    phase_values = {key: unsigned(required(phases, key, "phases"), f"{path}.phases.{key}") for key in PHASE_FIELDS}
    require(all(item <= counters["elapsed_ns"] for item in phase_values.values()), f"phase exceeds elapsed time: {path}")
    require(sum(phase_values.values()) <= counters["elapsed_ns"], f"phase sum exceeds elapsed time: {path}")

    direct = counters["direct_allocated_bytes"]
    realloc_new = counters["realloc_new_bytes"]
    realloc_old = counters["realloc_old_bytes"]
    freed = counters["deallocated_bytes"]
    live_before = counters["live_before"]
    live_after = counters["live_after"]
    require(direct + realloc_new <= U64_MAX, f"allocation equation overflows: {path}")
    expected_live = live_before + direct + realloc_new - realloc_old - freed
    require(expected_live >= 0 and expected_live <= U64_MAX, f"live equation underflows: {path}")
    require(counters["requested_alloc_bytes"] == direct + realloc_new, f"requested allocation equation failed: {path}")
    require(live_after == expected_live, f"live allocation equation failed: {path}")
    require(counters["peak_live_delta"] >= max(0, live_after - live_before), f"peak live delta is impossible: {path}")
    require(required(value, "alloc_balance_ok", "sample") is True, f"allocator balance failed: {path}")
    require(required(value, "alloc_invalid", "sample") is False, f"allocator invalid flag set: {path}")
    require(counters["allocation_failed"] == 0, f"allocator failure flag set: {path}")

    verify_metrics(required(value, "input_metrics", "sample"), f"{path}.input_metrics")
    output_metrics = required(value, "output_metrics", "sample")
    input_sha = text(required(value, "input_sha256", "sample"), f"{path}.input_sha256")
    require(len(input_sha) == 64 and all(c in "0123456789abcdef" for c in input_sha), f"input SHA malformed: {path}")
    output_sha = required(value, "output_sha256", "sample")
    physical = required(value, "source_readback_physical_ok", "sample")
    metadata = required(value, "source_readback_metadata_ok", "sample")
    for name, item in (("source_readback_physical_ok", physical), ("source_readback_metadata_ok", metadata)):
        require(item is None or type(item) is bool, f"{name} must be null or boolean: {path}")
    cap_error = required(value, "cap_refusal", "sample")
    commit_cap_error = required(value, "cap_commit_refusal", "sample")
    cap_source_metrics = required(value, "cap_source_metrics", "sample")
    cap_projected_metrics = required(value, "cap_projected_metrics", "sample")
    existing_cap = required(value, "cap_existing", "sample")
    if expected_success:
        require(required(value, "error", "sample") is None, f"unexpected primary refusal: {path}")
        require(counters["output_bytes"] > 0, f"successful lane emitted no output: {path}")
        require(type(output_sha) is str and len(output_sha) == 64 and all(c in "0123456789abcdef" for c in output_sha), f"output SHA malformed: {path}")
        require(type(output_metrics) is dict, f"successful lane has no output metrics: {path}")
        verify_metrics(output_metrics, f"{path}.output_metrics")
        if lane.startswith("cap_"):
            require(required(value, "cap_exact_fit_ok", "sample") is True, f"exact cap control failed: {path}")
            require(required(value, "cap_under_refused_ok", "sample") is True, f"under cap refusal failed: {path}")
            verify_error(cap_error, f"{path}.cap_refusal", "DocxError::Opc::ReadLimit", CAP_RESOURCES[lane])
            verify_error(commit_cap_error, f"{path}.cap_commit_refusal", "DocxError::Opc::ReadLimit", CAP_RESOURCES[lane])
            require(physical is True and metadata is True, f"cap source readback failed: {path}")
            verify_metrics(cap_source_metrics, f"{path}.cap_source_metrics")
            verify_metrics(cap_projected_metrics, f"{path}.cap_projected_metrics")
            require(cap_source_metrics == required(value, "input_metrics", "sample"), f"cap source metrics mismatch: {path}")
            require(cap_projected_metrics == output_metrics, f"cap projected metrics mismatch: {path}")
            verify_cap_evidence(existing_cap, f"{path}.cap_existing", lane)
        else:
            require(cap_error is None and commit_cap_error is None, f"unexpected cap receipt: {path}")
            require(cap_source_metrics is None and cap_projected_metrics is None, f"unexpected cap metrics: {path}")
            require(existing_cap is None, f"unexpected existing-owner cap receipt: {path}")
            require(physical is None and metadata is None, f"unexpected successful refusal readback: {path}")
    else:
        require(lane in EXPECTED_REFUSALS, f"unclassified refusal lane: {lane}")
        ingress = required(value, "ingress_refusal", "sample")
        require(ingress is False or lane.startswith("malformed_") or lane == "add_glossary_missing", f"unexpected ingress status: {path}")
        verify_error(
            required(value, "error", "sample"),
            f"{path}.error",
            EXPECTED_VARIANTS.get(lane),
            EXPECTED_RESOURCES.get(lane),
        )
        require(cap_error is None and commit_cap_error is None, f"unexpected cap receipt on refusal: {path}")
        require(cap_source_metrics is None and cap_projected_metrics is None, f"unexpected cap metrics on refusal: {path}")
        require(existing_cap is None, f"unexpected existing-owner cap receipt on refusal: {path}")
        require(counters["output_bytes"] == 0 and output_metrics is None and output_sha is None, f"refusal emitted output: {path}")
        if lane.startswith("malformed_") and ingress:
            require(physical is None and metadata is None, f"ingress refusal readback shape changed: {path}")
        elif lane.startswith("malformed_"):
            require(physical is True and metadata is True, f"post-ingress malformed readback failed: {path}")
        else:
            require(physical is True and metadata is True, f"refusal source readback failed: {path}")


def expected_fixture(lane: str) -> str:
    if "bug" in lane or lane in {
        "source_noop_main", "projection_main", "replace_main", "remove_main", "inverse_replace_main",
        "inverse_remove_main", "stale_patch_main", "independent_main",
    } or lane.startswith("malformed_"):
        return "Bug54849.docx"
    if "signed" in lane:
        return "ms-office-2010-signed.docx"
    if "complex" in lane:
        return "ComplexNumberedLists.docx"
    if lane == "add_glossary_missing" or lane.startswith("cap_") or lane == "add_main_absent":
        return "ComplexNumberedLists.docx" if lane == "add_glossary_missing" else "ComplexNumberedLists.docx:main-effects-absent"
    if "glossary" in lane:
        return "testGlossary.docx"
    raise AssertionError(f"no fixture mapping for {lane}")


def verify_receipts(results: Path, lanes: list[str]) -> dict[str, int]:
    total_samples = 0
    for lane in lanes:
        expected_success = lane not in EXPECTED_REFUSALS
        path = results / f"smoke-{lane}-p1.json"
        require(path.is_file(), f"missing smoke receipt: {lane}")
        value = json.loads(path.read_text())
        require(type(value) is dict, f"receipt must be an object: {path}")
        require(text(required(value, "schema", "receipt"), f"{path}.schema") == RECEIPT_SCHEMA, f"receipt schema changed: {path}")
        require(text(required(value, "source_commit", "receipt"), f"{path}.source_commit") == SOURCE_COMMIT, f"receipt source changed: {path}")
        require(text(required(value, "opc_source_label", "receipt"), f"{path}.opc_source_label") == "styles-effects-source-committed", f"OPC source label changed: {path}")
        require(text(required(value, "lane", "receipt"), f"{path}.lane") == lane, f"receipt lane changed: {path}")
        require(boolean(required(value, "expected_success", "receipt"), f"{path}.expected_success") is expected_success, f"receipt expected status changed: {path}")
        require(unsigned(required(value, "warmup", "receipt"), f"{path}.warmup") == 0, f"smoke warmup changed: {path}")
        require(unsigned(required(value, "sample_count", "receipt"), f"{path}.sample_count") == 1, f"smoke sample count changed: {path}")
        require(boolean(required(value, "source_backed_api", "receipt"), f"{path}.source_backed_api") is True, f"source-backed API gate failed: {path}")
        fixture_name = text(required(value, "fixture", "receipt"), f"{path}.fixture")
        require(fixture_name == expected_fixture(lane), f"fixture mapping changed: {path}")
        for key in ("fixture_native", "fixture_signed", "fixture_main_present", "fixture_glossary_present"):
            boolean(required(value, key, "receipt"), f"{path}.{key}")
        expected_package = required(value, "fixture_expected_package_sha256", "receipt")
        if fixture_name in NATIVE_FIXTURES:
            require(expected_package == NATIVE_FIXTURES[fixture_name]["sha256"], f"native package hash changed: {path}")
        samples = required(value, "samples", "receipt")
        require(type(samples) is list and len(samples) == 1, f"sample array changed: {path}")
        if boolean(required(value, "fixture_native", "receipt"), "fixture_native") and lane not in SYNTHETIC_NATIVE_LANES:
            require(
                samples[0].get("input_sha256") == expected_package,
                f"native input package hash changed: {path}",
            )
        verify_sample(samples[0], lane, expected_success, str(path))
        verify_time_sidecar(results / f"smoke-{lane}-p1.time.txt", results / f"smoke-{lane}-p1.stderr.log")
        total_samples += 1
    return {"lanes": len(lanes), "samples": total_samples}


def verify_clean_status(root: Path) -> None:
    status = subprocess.check_output(
        ["git", "-C", str(root), "status", "--porcelain", "--untracked-files=all"],
        text=True,
    )
    require(status == "", "source checkout is dirty during receipt verification")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--evidence", type=Path, required=True)
    parser.add_argument("--results", type=Path, required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--manifest-after", type=Path, required=True)
    parser.add_argument("--metadata-before", type=Path, required=True)
    parser.add_argument("--metadata-after", type=Path, required=True)
    parser.add_argument("--corpus", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    root = args.root.resolve()
    evidence = args.evidence.resolve()
    results = args.results.resolve()
    require(results.is_dir(), "results directory is missing")
    current_head = subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip()
    ancestry = subprocess.run(
        ["git", "-C", str(root), "merge-base", "--is-ancestor", SOURCE_COMMIT, current_head],
        check=False,
    )
    require(ancestry.returncode == 0, "checkout does not descend from the approved source commit")
    verify_clean_status(root)
    manifest_summary = verify_manifest(args.manifest, root, args.metadata_before, evidence)
    manifest_fields = {
        line.split("=", 1)[0]: line.split("=", 1)[1]
        for line in args.manifest.read_text().splitlines()
        if "=" in line and not line.startswith(("file=", "package=", "extra="))
    }
    require(manifest_fields.get("git_head") == current_head, "manifest Git head differs from checkout")
    require(args.manifest_after.read_text() == args.manifest.read_text(), "source closure changed after smoke")
    require(sha256(args.metadata_before) == sha256(args.metadata_after), "Cargo metadata changed after smoke")
    corpus = verify_corpus(args.corpus, root, evidence)
    lanes = [str(lane) for lane in corpus["lanes"]]
    verify_build_receipts(
        results,
        current_head,
        args.manifest,
        args.manifest_after,
        args.metadata_before,
        args.metadata_after,
        lanes,
    )
    receipt_summary = verify_receipts(results, lanes)
    output = {
        "schema": "docx-styles-effects-smoke-verification-v1",
        "passed": True,
        "source_commit": SOURCE_COMMIT,
        "manifest": manifest_summary,
        "corpus": {"fixtures": corpus["fixtures"], "lanes": len(lanes)},
        "receipts": receipt_summary,
    }
    args.output.write_text(json.dumps(output, indent=2, sort_keys=True) + "\n")


if __name__ == "__main__":
    main()
