#!/usr/bin/env python3
"""Verify bounded PPTX InkAction profile provenance and receipts.

The verifier is intentionally independent of the Rust adapter.  It checks the
committed matrix, the full source closure, process/RSS companions, typed
refusal evidence, and every phase allocation equation before accepting a
profile bundle.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import subprocess
from pathlib import Path


SEMANTIC_OWNER_COMMIT = "cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd"
PRODUCTION_SOURCE_BASELINE_COMMIT = "2a2ffa1cae4e6b7070082768ce84483e5d411dc8"
SEMANTIC_OWNER_DESIGN_PATH = "docs/report/spec-gap-validation-evidence/pptx-ink-actions-design.md"
SEMANTIC_OWNER_DESIGN_SHA256 = "30b78cca84c4ca24ae44f3d3694c3097f54b5e5a1f2004af9ce5007bcaf4173d"
SEMANTIC_OWNER_DESIGN_GIT_BLOB = "597400950b1027c47cd6e4cbbedd23915bc0980e"
CAPTURE_CONTEXT_EXTRAS = {
    "docs/adr/0001-priorities-and-api-layers.md",
    "docs/adr/0003-snapshots-edits-and-patches.md",
    "docs/adr/0005-io-memory-and-performance.md",
    "docs/adr/0006-validation-security-and-compatibility.md",
}
# Compatibility name for receipts written before the dual-pin fields were
# introduced.  It always denotes the semantic owner, never the production
# source baseline.
SOURCE_COMMIT = SEMANTIC_OWNER_COMMIT
HELPER_SHA256 = "bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e"
SCHEMA = "pptx-ink-actions-performance-v1"
SOURCE_FORMAT = "pptx-ink-actions-profile-build-source-v1"
PRODUCTION_WORKSPACE_EXTRAS = {
    "Cargo.toml",
    "rust-toolchain.toml",
    ".cargo/config.toml",
    "rustfmt.toml",
    "clippy.toml",
    "deny.toml",
}
PRODUCTION_PACKAGE_EXCLUDES = {"litchi-pptx-ink-actions-performance"}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def resolve(root: Path, shown: str) -> Path:
    path = Path(shown)
    return path if path.is_absolute() else root / path


def blob_sha256(root: Path, commit: str, shown: str) -> str:
    value = subprocess.check_output(
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
    return hashlib.sha256(value).hexdigest()


def blob_id(root: Path, commit: str, shown: str) -> str:
    return subprocess.check_output(
        [
            "git",
            "--no-replace-objects",
            "-C",
            str(root),
            "rev-parse",
            f"{commit}:{shown}",
        ],
        text=True,
    ).strip()


def verify_manifest(path: Path, root: Path) -> tuple[str, str, str]:
    lines = path.read_text().splitlines()
    require(lines and lines[0] == f"format={SOURCE_FORMAT}", f"source manifest format: {path}")
    commit: str | None = None
    semantic_owner_commit: str | None = None
    production_source_baseline_commit: str | None = None
    source_commit_alias: str | None = None
    owner_commit: str | None = None
    packages: dict[tuple[str, str, str], tuple[str, int, str]] = {}
    files: dict[tuple[str, str, str], list[tuple[str, str]]] = {}
    extras: list[tuple[str, str]] = []
    context_extras: list[tuple[str, str]] = []
    production_extras: list[tuple[str, str]] = []
    semantic_extras: list[tuple[str, str, str]] = []
    production_excludes: set[str] = set()
    for line in lines[1:]:
        if line.startswith("git_commit="):
            require(commit is None, "duplicate source manifest git_commit")
            commit = line.split("=", 1)[1]
        elif line.startswith("semantic_owner_commit="):
            require(semantic_owner_commit is None, "duplicate source manifest semantic owner pin")
            semantic_owner_commit = line.split("=", 1)[1]
        elif line.startswith("production_source_baseline_commit="):
            require(
                production_source_baseline_commit is None,
                "duplicate source manifest production baseline pin",
            )
            production_source_baseline_commit = line.split("=", 1)[1]
        elif line.startswith("source_commit="):
            require(source_commit_alias is None, "duplicate source manifest source_commit alias")
            source_commit_alias = line.split("=", 1)[1]
        elif line.startswith("owner_commit="):
            require(owner_commit is None, "duplicate source manifest owner_commit")
            owner_commit = line.split("=", 1)[1]
        elif line.startswith("production_exclude_package="):
            production_excludes.add(line.split("=", 1)[1])
        elif line.startswith("metadata_sha256="):
            digest = line.split("=", 1)[1]
            require(len(digest) == 64 and all(c in "0123456789abcdef" for c in digest), "bad metadata hash")
        elif line.startswith("package="):
            fields = line.split("\t")
            require(len(fields) == 7, f"malformed package line: {line}")
            key = (fields[0][len("package=") :], fields[1], fields[3])
            require(key not in packages, f"duplicate package: {key}")
            packages[key] = (fields[2], int(fields[5]), fields[6])
        elif line.startswith("file="):
            fields = line.split("\t")
            require(len(fields) == 5, f"malformed file line: {line}")
            key = (fields[0][len("file=") :], fields[1], fields[2])
            files.setdefault(key, []).append((fields[3], fields[4]))
        elif line.startswith("extra=\t"):
            fields = line.split("\t")
            require(len(fields) == 3, f"malformed extra line: {line}")
            extras.append((fields[1], fields[2]))
        elif line.startswith("context_extra=\t"):
            fields = line.split("\t")
            require(len(fields) == 3, f"malformed capture context extra line: {line}")
            context_extras.append((fields[1], fields[2]))
        elif line.startswith("semantic_extra=\t"):
            fields = line.split("\t")
            require(len(fields) == 4, f"malformed semantic owner extra line: {line}")
            semantic_extras.append((fields[1], fields[2], fields[3]))
        elif line.startswith("production_extra=\t") or line.startswith("owner_extra=\t"):
            fields = line.split("\t")
            require(len(fields) == 3, f"malformed owner extra line: {line}")
            production_extras.append((fields[1], fields[2]))
        else:
            raise AssertionError(f"unknown source manifest line: {line}")
    require(
        commit is not None
        and len(commit) == 40
        and all(char in "0123456789abcdef" for char in commit),
        "source manifest capture HEAD is malformed",
    )
    require(semantic_owner_commit == SEMANTIC_OWNER_COMMIT, "source manifest semantic owner pin changed")
    require(
        production_source_baseline_commit == PRODUCTION_SOURCE_BASELINE_COMMIT,
        "source manifest production baseline pin changed",
    )
    require(source_commit_alias == semantic_owner_commit, "source manifest source_commit alias changed")
    require(owner_commit == semantic_owner_commit, "source manifest owner alias changed")
    require(
        {shown for shown, _ in context_extras} == CAPTURE_CONTEXT_EXTRAS,
        "source manifest capture context inputs changed",
    )
    require(
        semantic_extras == [
            (
                SEMANTIC_OWNER_DESIGN_PATH,
                SEMANTIC_OWNER_DESIGN_SHA256,
                SEMANTIC_OWNER_DESIGN_GIT_BLOB,
            )
        ],
        "source manifest semantic owner design input changed",
    )
    require(
        production_excludes == PRODUCTION_PACKAGE_EXCLUDES,
        f"source manifest production package exclusions changed: {sorted(production_excludes)}",
    )
    require(set(packages) == set(files), "source package/file keys differ")
    local_paths: list[Path] = []
    production_paths: list[Path] = []
    checked = 0
    for key, (source, count, tree) in packages.items():
        entries = files[key]
        require(len(entries) == count, f"source package file count changed: {key}")
        payload = "\n".join(f"{name}\t{digest}" for name, digest in entries)
        require(hashlib.sha256(payload.encode()).hexdigest() == tree, f"source package tree changed: {key}")
        package_name = key[0]
        for shown, digest in entries:
            path_value = resolve(root, shown)
            require(path_value.is_file(), f"source file missing: {shown}")
            require(sha256(path_value) == digest, f"source file changed: {shown}")
            if source == "path":
                local_paths.append(path_value)
                if package_name not in production_excludes:
                    production_paths.append(path_value)
            checked += 1
    for shown, digest in extras:
        path_value = resolve(root, shown)
        require(path_value.is_file(), f"source extra missing: {shown}")
        require(sha256(path_value) == digest, f"source extra changed: {shown}")
        local_paths.append(path_value)
        checked += 1
    for shown, digest in context_extras:
        path_value = resolve(root, shown)
        require(path_value.is_file(), f"capture context input missing: {shown}")
        require(sha256(path_value) == digest, f"capture context input changed: {shown}")
        local_paths.append(path_value)
        checked += 1
    for shown, digest, expected_blob in semantic_extras:
        path_value = resolve(root, shown)
        require(path_value.is_file(), f"semantic owner design input missing: {shown}")
        require(sha256(path_value) == digest, f"semantic owner design input changed: {shown}")
        local_paths.append(path_value)
        checked += 1
    production_extra_names = {shown for shown, _ in production_extras}
    require(
        production_extra_names == PRODUCTION_WORKSPACE_EXTRAS,
        f"workspace production extras are incomplete: {sorted(production_extra_names)}",
    )
    for shown, digest in production_extras:
        path_value = resolve(root, shown)
        require(path_value.is_file(), f"production source extra missing: {shown}")
        require(sha256(path_value) == digest, f"production source extra changed: {shown}")
        local_paths.append(path_value)
        checked += 1
    require(checked > 20, f"source manifest is too small: {checked}")
    head = subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip()
    require(head == commit, f"source HEAD changed: {commit} -> {head}")
    for path_value in sorted(set(local_paths), key=str):
        relative = path_value.resolve().relative_to(root.resolve()).as_posix()
        tracked = subprocess.run(
            ["git", "-C", str(root), "ls-files", "--error-unmatch", "--", relative],
            capture_output=True,
            text=True,
            check=False,
        )
        require(tracked.returncode == 0, f"local source input is untracked: {relative}")
        require(
            sha256(path_value) == blob_sha256(root, commit, relative),
            f"local source input differs from committed tree: {relative}",
        )
    for shown, digest in production_extras:
        path_value = resolve(root, shown)
        require(
            sha256(path_value) == blob_sha256(root, PRODUCTION_SOURCE_BASELINE_COMMIT, shown),
            f"production workspace input differs from approved baseline: {shown}",
        )
    for path_value in sorted(set(production_paths), key=str):
        relative = path_value.resolve().relative_to(root.resolve()).as_posix()
        require(
            sha256(path_value) == blob_sha256(root, PRODUCTION_SOURCE_BASELINE_COMMIT, relative),
            f"production path-package input differs from approved baseline: {relative}",
        )
    for shown, digest, expected_blob in semantic_extras:
        require(
            digest == SEMANTIC_OWNER_DESIGN_SHA256,
            f"semantic owner design SHA-256 changed: {shown}",
        )
        require(
            blob_id(root, SEMANTIC_OWNER_COMMIT, shown) == expected_blob == SEMANTIC_OWNER_DESIGN_GIT_BLOB,
            f"semantic owner design Git blob changed: {shown}",
        )
        require(
            sha256(resolve(root, shown)) == blob_sha256(root, SEMANTIC_OWNER_COMMIT, shown),
            f"semantic owner design differs from semantic owner commit: {shown}",
        )
    require(
        subprocess.run(
            [
                "git",
                "-C",
                str(root),
                "merge-base",
                "--is-ancestor",
                SEMANTIC_OWNER_COMMIT,
                commit,
            ],
            check=False,
        ).returncode
        == 0,
        "capture HEAD does not descend from semantic owner pin",
    )
    require(
        subprocess.run(
            [
                "git",
                "-C",
                str(root),
                "merge-base",
                "--is-ancestor",
                PRODUCTION_SOURCE_BASELINE_COMMIT,
                commit,
            ],
            check=False,
        ).returncode
        == 0,
        "capture HEAD does not descend from production baseline pin",
    )
    return commit, semantic_owner_commit, production_source_baseline_commit


def verify_corpus(path: Path) -> tuple[dict[str, dict[str, object]], dict[str, dict[str, object]], str]:
    corpus = json.loads(path.read_text())
    require(corpus["semantic_owner_commit"] == SEMANTIC_OWNER_COMMIT, "corpus semantic owner pin changed")
    require(
        corpus["production_source_baseline_commit"] == PRODUCTION_SOURCE_BASELINE_COMMIT,
        "corpus production baseline pin changed",
    )
    require(corpus["owner_commit"] == corpus["semantic_owner_commit"], "corpus owner alias changed")
    require(corpus["source_commit"] == corpus["semantic_owner_commit"], "corpus source_commit alias changed")
    design = corpus["semantic_owner_design"]
    require(design["commit"] == SEMANTIC_OWNER_COMMIT, "corpus semantic design commit changed")
    require(design["path"] == SEMANTIC_OWNER_DESIGN_PATH, "corpus semantic design path changed")
    require(design["sha256"] == SEMANTIC_OWNER_DESIGN_SHA256, "corpus semantic design SHA-256 changed")
    require(design["git_blob"] == SEMANTIC_OWNER_DESIGN_GIT_BLOB, "corpus semantic design Git blob changed")
    root = path.parent.parent.parent.parent.parent
    fixture = corpus["fixture_authority"]
    fixture_path = root / str(fixture["path"])
    require(fixture_path.is_file(), "fixture authority is missing")
    require(sha256(fixture_path) == fixture["sha256"], "fixture authority hash changed")
    fixture_blob = subprocess.check_output(
        ["git", "-C", str(root), "rev-parse", f"{SEMANTIC_OWNER_COMMIT}:{fixture['path']}"],
        text=True,
    ).strip()
    require(fixture_blob == fixture["git_blob"], "fixture authority Git blob changed")
    design_path = root / SEMANTIC_OWNER_DESIGN_PATH
    require(design_path.is_file(), "semantic owner design input is missing")
    require(sha256(design_path) == SEMANTIC_OWNER_DESIGN_SHA256, "semantic owner design input hash changed")
    require(
        blob_id(root, SEMANTIC_OWNER_COMMIT, SEMANTIC_OWNER_DESIGN_PATH) == SEMANTIC_OWNER_DESIGN_GIT_BLOB,
        "semantic owner design authority Git blob changed",
    )
    generator = corpus["retained_opc_generator"]
    require(generator["path"] == "harness/adapter.rs", "generator path changed")
    require(
        isinstance(generator["sha256"], str)
        and len(generator["sha256"]) == 64,
        "generator hash missing",
    )
    generator_path = path.parent / str(generator["path"])
    require(generator_path.is_file(), "retained generator is missing")
    require(sha256(generator_path) == generator["sha256"], "retained generator hash changed")
    git_blob = subprocess.check_output(
        ["git", "-C", str(root), "hash-object", str(generator_path)],
        text=True,
    ).strip()
    require(git_blob == generator["git_blob"], "retained generator Git blob changed")
    lockfile = path.parent / "harness" / "Cargo.lock"
    require(lockfile.is_file(), "isolated harness lockfile is missing")
    require(
        hashlib.sha256(lockfile.read_bytes()).hexdigest() == corpus["isolated_lockfile_sha256"],
        "isolated harness lockfile hash changed",
    )
    recipes = {item["id"]: item for item in corpus["recipes"]}
    lanes = {item["id"]: item for item in corpus["lanes"]}
    require(len(recipes) == 23 and len(lanes) == 42, "bounded matrix counts changed")
    require(len(recipes) == len(corpus["recipes"]), "recipe IDs are not unique")
    require(len(lanes) == len(corpus["lanes"]), "lane IDs are not unique")
    used = {item["recipe"] for item in corpus["lanes"]}
    require(used == set(recipes), "recipe usage has orphaned or unresolved entries")
    for lane in corpus["lanes"]:
        require(lane["recipe"] in recipes, f"unresolved lane recipe: {lane['id']}")
    gate = corpus["measurement_gate"]
    require(gate["fresh_processes"] == 3, "fresh process count changed")
    require(gate["warmups_per_process"] == 2, "warmup count changed")
    require(gate["samples_per_process"] == 20, "sample count changed")
    return recipes, lanes, str(generator["sha256"])


def rss(path: Path) -> int:
    values = [
        line.split(":", 1)[1].strip()
        for line in path.read_text().splitlines()
        if line.startswith("Maximum resident set size (kbytes):")
    ]
    require(len(values) == 1, f"timing file must contain one RSS line: {path}")
    require(values[0].isdigit() and int(values[0]) > 0, f"malformed RSS: {path}")
    status = [
        line.split(":", 1)[1].strip()
        for line in path.read_text().splitlines()
        if line.startswith("Exit status:")
    ]
    require(status == ["0"], f"timed command did not exit zero: {path}")
    return int(values[0])


def timing_fields(path: Path) -> dict[str, str]:
    lines = path.read_text().splitlines()
    prefixes = {
        "user": "User time (seconds):",
        "system": "System time (seconds):",
        "elapsed": "Elapsed (wall clock) time (h:mm:ss or m:ss):",
    }
    values: dict[str, str] = {}
    for name, prefix in prefixes.items():
        matches = [line.split(":", 1)[1].strip() for line in lines if line.startswith(prefix)]
        require(len(matches) == 1 and matches[0], f"timing file missing {name}: {path}")
        values[name] = matches[0]
    for name in ("user", "system"):
        try:
            require(float(values[name]) >= 0, f"negative {name} time: {path}")
        except ValueError as error:
            raise AssertionError(f"malformed {name} time: {path}") from error
    parts = values["elapsed"].split(":")
    require(len(parts) in (2, 3), f"malformed elapsed time: {path}")
    try:
        elapsed_seconds = sum(
            float(part) * (60 ** (len(parts) - index - 1))
            for index, part in enumerate(parts)
        )
    except ValueError as error:
        raise AssertionError(f"malformed elapsed time: {path}") from error
    require(elapsed_seconds >= 0, f"negative elapsed time: {path}")
    return values


def verify_phase(phase: dict[str, object], path: Path) -> None:
    fields = (
        "elapsed_ns",
        "requested_alloc_bytes",
        "direct_allocated_bytes",
        "realloc_new_bytes",
        "realloc_old_bytes",
        "deallocated_bytes",
        "live_before_bytes",
        "live_after_bytes",
        "peak_live_during_bytes",
        "peak_live_delta_bytes",
        "alloc_failed",
        "alloc_calls",
        "realloc_calls",
        "dealloc_calls",
    )
    for field in fields:
        value = phase.get(field)
        require(type(value) is int and value >= 0, f"bad phase field {field}: {path}")
    live_delta = phase.get("live_delta_bytes")
    require(type(live_delta) is int, f"bad live_delta_bytes: {path}")
    require(phase.get("alloc_balance_ok") is True, f"allocation equation flag failed: {path}")
    require(phase.get("alloc_invalid") is False, f"allocator invalid flag: {path}")
    require(phase["alloc_failed"] == 0, f"allocation failure count: {path}")
    require(
        phase["requested_alloc_bytes"]
        == phase["direct_allocated_bytes"] + phase["realloc_new_bytes"],
        f"requested allocation equation failed: {path}",
    )
    require(
        phase["live_before_bytes"]
        + phase["direct_allocated_bytes"]
        + phase["realloc_new_bytes"]
        == phase["live_after_bytes"]
        + phase["realloc_old_bytes"]
        + phase["deallocated_bytes"],
        f"live allocation equation failed: {path}",
    )
    require(
        phase["live_delta_bytes"] == phase["live_after_bytes"] - phase["live_before_bytes"],
        f"live delta equation failed: {path}",
    )
    require(
        phase["peak_live_delta_bytes"]
        == phase["peak_live_during_bytes"] - phase["live_before_bytes"],
        f"phase-local peak equation failed: {path}",
    )
    require(
        phase["peak_live_delta_bytes"]
        >= max(0, phase["live_after_bytes"] - phase["live_before_bytes"]),
        f"peak retained live bytes is too small: {path}",
    )


def error_matches(expected: str, sample: dict[str, object]) -> bool:
    if expected == "success":
        return False
    expected_kind = expected.split(" {", 1)[0] if " {" in expected else expected
    type_ok = sample.get("actual_error_type") == expected_kind
    if "resource:" in expected:
        wanted = expected.split("resource:", 1)[1].split(",", 1)[0].strip()
        type_ok &= sample.get("actual_error_resource") == wanted
    if "limit:" in expected:
        wanted = expected.split("limit:", 1)[1].split("}", 1)[0].strip()
        type_ok &= sample.get("actual_error_limit") == int(wanted)
    return type_ok


def verify_receipt(
    path: Path,
    timing: Path,
    stderr: Path,
    recipe: dict[str, object],
    lane: dict[str, object],
    generator_sha256: str,
    capture_head: str,
) -> tuple[int, list[int], list[int], list[int], dict[str, str]]:
    receipt = json.loads(path.read_text())
    require(receipt["schema"] == SCHEMA, f"receipt schema changed: {path}")
    require(receipt["lane"] == lane["id"], f"receipt lane mismatch: {path}")
    require(receipt["recipe_id"] == recipe["id"], f"receipt recipe mismatch: {path}")
    require(receipt["source_commit"] == SEMANTIC_OWNER_COMMIT, f"receipt semantic owner pin changed: {path}")
    require(
        receipt.get("semantic_owner_commit") == SEMANTIC_OWNER_COMMIT,
        f"receipt semantic owner pin changed: {path}",
    )
    require(
        receipt.get("production_source_baseline_commit") == PRODUCTION_SOURCE_BASELINE_COMMIT,
        f"receipt production baseline pin changed: {path}",
    )
    require(receipt.get("capture_head") == capture_head, f"receipt capture HEAD changed: {path}")
    require(receipt["helper_sha256"] == HELPER_SHA256, f"receipt helper hash changed: {path}")
    require(receipt.get("generator_sha256") == generator_sha256, f"generator hash changed: {path}")
    require(receipt.get("generator_source_sha256") == generator_sha256, f"generator hash changed: {path}")
    require(receipt["warmup"] == 2 and receipt["sample_count"] == 20, f"receipt sample gate changed: {path}")
    require(len(receipt["samples"]) == 20, f"sample count changed: {path}")
    require(receipt["source_sha256"] == receipt["samples"][0]["source_sha256"], f"source hash mismatch: {path}")
    require(receipt["source_fnv1a64"] == receipt["samples"][0]["source_fnv1a64"], f"source FNV mismatch: {path}")
    require(
        receipt["source_manifest_sha256"] == receipt["samples"][0]["source_manifest_sha256"],
        f"source manifest hash mismatch: {path}",
    )
    require(
        receipt["package_before_manifest_sha256"]
        == receipt["samples"][0]["package_before_manifest_sha256"],
        f"package manifest hash mismatch: {path}",
    )
    phases = ("setup", "operation", "validation", "drop", "postdrop")
    expected_error = lane.get("expected")
    rss_value = rss(timing)
    timing_values = timing_fields(timing)
    require(stderr.read_text() == "", f"timed lane wrote stderr: {stderr}")
    values: list[int] = []
    latencies: list[int] = []
    operation_allocations: list[int] = []
    process_ids: set[int] = set()
    for index, sample in enumerate(receipt["samples"]):
        sample_path = Path(f"{path}#sample-{index}")
        require(
            type(sample.get("process_id")) is int and sample["process_id"] > 0,
            f"process ID missing: {sample_path}",
        )
        process_ids.add(int(sample["process_id"]))
        require(type(sample.get("semantic_ok")) is bool, f"semantic flag malformed: {sample_path}")
        require(type(sample.get("preservation_ok")) is bool, f"preservation flag malformed: {sample_path}")
        require(sample.get("inverse_ok") is True, f"inverse flag failed: {sample_path}")
        require(
            sample.get("package_manifest_preserved") is True,
            f"package manifest preservation failed: {sample_path}",
        )
        require(type(sample.get("source_sha256")) is str and len(sample["source_sha256"]) == 64, f"source digest malformed: {sample_path}")
        for field in ("source_manifest_sha256", "package_before_manifest_sha256"):
            require(
                type(sample.get(field)) is str and len(sample[field]) == 64,
                f"{field} malformed: {sample_path}",
            )
        for field in ("baseline_unique_target_bytes", "expected_unique_target_bytes"):
            require(
                type(sample.get(field)) is int and sample[field] >= 0,
                f"{field} malformed: {sample_path}",
            )
        if sample.get("output_manifest_sha256") is not None:
            require(
                type(sample["output_manifest_sha256"]) is str
                and len(sample["output_manifest_sha256"]) == 64,
                f"output manifest digest malformed: {sample_path}",
            )
        for field in (
            "retained_baseline_live_bytes",
            "after_drop_live_bytes",
            "after_postdrop_live_bytes",
            "baseline_release_bytes",
        ):
            require(
                type(sample.get(field)) is int and sample[field] >= 0,
                f"{field} malformed: {sample_path}",
            )
        require(sample.get("baseline_reopenable") is True, f"baseline reopenability failed: {sample_path}")
        require(
            sample.get("retained_baseline_balance_ok") is True
            and sample["after_drop_live_bytes"] == sample["retained_baseline_live_bytes"],
            f"retained baseline balance failed: {sample_path}",
        )
        require(
            sample["after_drop_live_bytes"] >= sample["after_postdrop_live_bytes"]
            and sample["baseline_release_bytes"]
            == sample["after_drop_live_bytes"] - sample["after_postdrop_live_bytes"],
            f"retained-baseline release equation failed: {sample_path}",
        )
        for field in (
            "opaque_choice_preserved",
            "opaque_fallback_preserved",
            "opaque_payload_preserved",
            "opaque_default_namespace_preserved",
            "opaque_prefix_preserved",
            "opaque_unknown_requires_preserved",
        ):
            require(type(sample.get(field)) is bool, f"{field} malformed: {sample_path}")
        for field in ("owner_xml_sha256", "profile_source_sha256"):
            require(
                type(sample.get(field)) is str and len(sample[field]) == 64,
                f"{field} malformed: {sample_path}",
            )
        for phase in phases:
            value = sample.get("phases", {}).get(phase)
            require(isinstance(value, dict), f"missing phase {phase}: {sample_path}")
            verify_phase(value, sample_path)
        require(
            sample["phases"]["operation"]["elapsed_ns"] > 0,
            f"operation elapsed time is zero: {sample_path}",
        )
        latencies.append(sample["phases"]["operation"]["elapsed_ns"])
        operation_allocations.append(sample["phases"]["operation"]["requested_alloc_bytes"])
        if expected_error and expected_error != "success":
            require(sample["semantic_ok"] is True, f"expected refusal not accepted: {sample_path}")
            require(sample["preservation_ok"] is True, f"refusal preservation failed: {sample_path}")
            require(sample["source_unchanged_on_refusal"] is True, f"refusal mutated source: {sample_path}")
            require(
                isinstance(sample.get("actual_error_type"), str)
                and (sample.get("actual_error_resource") is None or isinstance(sample.get("actual_error_resource"), str))
                and (sample.get("actual_error_limit") is None or isinstance(sample.get("actual_error_limit"), int))
                and isinstance(sample.get("actual_error_debug"), str)
                and isinstance(sample.get("actual_error_display"), str),
                f"typed error receipt is incomplete: {sample_path}",
            )
            require(error_matches(str(expected_error), sample), f"wrong typed error: {sample_path}")
        else:
            require(sample["semantic_ok"] is True, f"semantic validation failed: {sample_path}")
            require(sample["preservation_ok"] is True, f"preservation validation failed: {sample_path}")
            require(sample["anchors"] == recipe["anchors"], f"anchor count changed: {sample_path}")
            require(sample["unique_targets"] == recipe["unique_targets"], f"target count changed: {sample_path}")
            require(sample["inbound_edges"] == recipe["edges"], f"inbound graph count changed: {sample_path}")
            require(
                sample["outbound_edges"] == recipe.get("outbound_edges", 0),
                f"outbound graph count changed: {sample_path}",
            )
            require(
                sample["baseline_unique_target_bytes"]
                == recipe["target_bytes"] * recipe["unique_targets"],
                f"baseline unique target bytes changed: {sample_path}",
            )
            require(
                sample["unique_target_bytes"] == sample["expected_unique_target_bytes"],
                f"published unique target bytes changed unexpectedly: {sample_path}",
            )
            expected_pointer = (
                "distinct"
                if recipe["topology"] == "distinct"
                else "shared"
                if recipe["topology"] in ("shared", "case_equivalent_shared")
                else sample["shared_pointer_observation"]
            )
            require(
                sample["shared_pointer_observation"] == expected_pointer,
                f"pointer sharing observation changed: {sample_path}",
            )
            if recipe.get("unknown_internal_outbound", False):
                require(sample["unknown_internal_outbound_preserved"] is True, f"internal outbound dropped: {sample_path}")
            if recipe.get("unknown_external_outbound", False):
                require(sample["unknown_external_outbound_preserved"] is True, f"external outbound dropped: {sample_path}")
            if recipe.get("opaque_mce", False):
                require(sample["opaque_choice_preserved"] is True, f"opaque MCE choice dropped: {sample_path}")
                require(sample["opaque_fallback_preserved"] is True, f"opaque MCE fallback dropped: {sample_path}")
                require(sample["opaque_payload_preserved"] is True, f"opaque payload dropped: {sample_path}")
                require(sample["opaque_default_namespace_preserved"] is True, f"opaque default namespace dropped: {sample_path}")
                require(sample["opaque_prefix_preserved"] is True, f"opaque prefix dropped: {sample_path}")
                require(sample["opaque_unknown_requires_preserved"] is True, f"opaque unknown Requires dropped: {sample_path}")
        values.append(rss_value)
    require(len(process_ids) == 1, f"samples in one process receipt have mixed process IDs: {path}")
    return next(iter(process_ids)), values, latencies, operation_allocations, timing_values


def verify_host_probe(path: Path, capture_head: str) -> None:
    probe = json.loads(path.read_text())
    require(probe["schema"] == "pptx-ink-actions-host-probe-v1", "host probe schema changed")
    require(probe["source_commit"] == SEMANTIC_OWNER_COMMIT, "host probe semantic owner changed")
    require(probe["semantic_owner_commit"] == SEMANTIC_OWNER_COMMIT, "host probe semantic owner changed")
    require(
        probe["production_source_baseline_commit"] == PRODUCTION_SOURCE_BASELINE_COMMIT,
        "host probe production baseline changed",
    )
    require(probe["capture_head"] == capture_head, "host probe capture HEAD changed")
    require(probe["helper_sha256"] == HELPER_SHA256, "host probe helper changed")
    require(probe["native_powerpoint_claim"] is False, "host probe made a native PowerPoint claim")
    require(probe["synthetic_complete_opc"] is True, "host probe fixture authority changed")
    require(probe["package_anchors"] == 1 and probe["presentation_anchors"] == 1, "host probe routes failed")
    for field in (
        "opaque_choice_preserved",
        "opaque_fallback_preserved",
        "opaque_payload_preserved",
        "opaque_default_namespace_preserved",
        "opaque_prefix_preserved",
        "opaque_unknown_requires_preserved",
        "opaque_internal_outbound_preserved",
        "opaque_external_outbound_preserved",
        "opaque_manifest_preserved",
        "opaque_patch_publication_preserved",
    ):
        require(probe.get(field) is True, f"opaque host probe failed: {field}")


def verify_host_metadata(path: Path, capture_head: str) -> None:
    lines = path.read_text().splitlines()
    required = (
        "source_commit=",
        "semantic_owner_commit=",
        "production_source_baseline_commit=",
        "semantic_owner_design_path=",
        "semantic_owner_design_sha256=",
        "semantic_owner_design_git_blob=",
        "capture_head=",
        "git_head=",
        "processes=",
        "warmup=",
        "samples=",
        "rustc=",
        "cargo=",
        "uname=",
        "hostname=",
        "loadavg=",
        "cpu_model=",
        "memory=",
    )
    for prefix in required:
        matches = [line for line in lines if line.startswith(prefix)]
        require(len(matches) == 1 and matches[0].split("=", 1)[1].strip(), f"host metadata missing {prefix}")
    values = read_host_metadata(path)
    require(values["source_commit"] == SEMANTIC_OWNER_COMMIT, "host metadata semantic owner changed")
    require(values["semantic_owner_commit"] == SEMANTIC_OWNER_COMMIT, "host metadata semantic owner changed")
    require(
        values["production_source_baseline_commit"] == PRODUCTION_SOURCE_BASELINE_COMMIT,
        "host metadata production baseline changed",
    )
    require(values["semantic_owner_design_path"] == SEMANTIC_OWNER_DESIGN_PATH, "host metadata semantic design path changed")
    require(values["semantic_owner_design_sha256"] == SEMANTIC_OWNER_DESIGN_SHA256, "host metadata semantic design SHA-256 changed")
    require(values["semantic_owner_design_git_blob"] == SEMANTIC_OWNER_DESIGN_GIT_BLOB, "host metadata semantic design Git blob changed")
    require(values["capture_head"] == capture_head and values["git_head"] == capture_head, "host metadata HEAD changed")


def read_host_metadata(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    for line in path.read_text().splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key] = value
    return values


def verify_provenance_file(path: Path, capture_head: str) -> None:
    values = read_host_metadata(path)
    require(values.get("semantic_owner_commit") == SEMANTIC_OWNER_COMMIT, f"provenance semantic owner changed: {path}")
    require(
        values.get("production_source_baseline_commit") == PRODUCTION_SOURCE_BASELINE_COMMIT,
        f"provenance production baseline changed: {path}",
    )
    require(values.get("semantic_owner_design_path") == SEMANTIC_OWNER_DESIGN_PATH, f"provenance semantic design path changed: {path}")
    require(values.get("semantic_owner_design_sha256") == SEMANTIC_OWNER_DESIGN_SHA256, f"provenance semantic design SHA-256 changed: {path}")
    require(values.get("semantic_owner_design_git_blob") == SEMANTIC_OWNER_DESIGN_GIT_BLOB, f"provenance semantic design Git blob changed: {path}")
    require(values.get("capture_head") == capture_head, f"provenance capture HEAD changed: {path}")


def verify_guard_output(path: Path, capture_head: str) -> None:
    values = read_host_metadata(path)
    require(values.get("semantic_owner_commit") == SEMANTIC_OWNER_COMMIT, "guard semantic owner changed")
    require(values.get("production_source_baseline_commit") == PRODUCTION_SOURCE_BASELINE_COMMIT, "guard production baseline changed")
    require(values.get("source_commit") == SEMANTIC_OWNER_COMMIT, "guard source_commit alias changed")
    require(values.get("capture_head") == capture_head, "guard capture HEAD changed")
    require(
        values.get("semantic_owner_extra")
        == f"{SEMANTIC_OWNER_DESIGN_PATH}\t{SEMANTIC_OWNER_DESIGN_SHA256}\t{SEMANTIC_OWNER_DESIGN_GIT_BLOB}",
        "guard semantic owner design authority changed",
    )


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
    recipes, lanes, generator_sha256 = verify_corpus(args.corpus.resolve())
    before_manifest = verify_manifest(args.manifest.resolve(), root)
    after_manifest = verify_manifest(args.manifest_after.resolve(), root)
    before_commit, before_semantic_owner, before_production_baseline = before_manifest
    after_commit, after_semantic_owner, after_production_baseline = after_manifest
    require(before_commit == after_commit, "source manifests use different commits")
    require(before_semantic_owner == after_semantic_owner == SEMANTIC_OWNER_COMMIT, "source manifests use different semantic owners")
    require(
        before_production_baseline
        == after_production_baseline
        == PRODUCTION_SOURCE_BASELINE_COMMIT,
        "source manifests use different production baselines",
    )
    require(
        subprocess.run(
            ["git", "-C", str(root), "merge-base", "--is-ancestor", SEMANTIC_OWNER_COMMIT, before_commit],
            check=False,
        ).returncode == 0,
        "capture HEAD does not descend from approved semantic owner pin",
    )
    require(
        subprocess.run(
            [
                "git",
                "-C",
                str(root),
                "merge-base",
                "--is-ancestor",
                PRODUCTION_SOURCE_BASELINE_COMMIT,
                before_commit,
            ],
            check=False,
        ).returncode
        == 0,
        "capture HEAD does not descend from approved production baseline pin",
    )
    require(args.metadata_before.read_bytes() == args.metadata_after.read_bytes(), "Cargo metadata changed")
    verify_host_probe(results / "host-probe.json", before_commit)
    verify_host_metadata(results / "host.txt", before_commit)
    host_metadata = read_host_metadata(results / "host.txt")
    require((results / "host-probe.stderr").read_text() == "", "host probe wrote stderr")
    require((results / "matrix-correctness.json").is_file(), "correctness matrix receipt missing")
    matrix = json.loads((results / "matrix-correctness.json").read_text())
    require(matrix.get("timings_collected") is False, "matrix correctness receipt claims timing")
    require(matrix.get("source_commit") == SEMANTIC_OWNER_COMMIT, "matrix semantic owner changed")
    require(matrix.get("semantic_owner_commit") == SEMANTIC_OWNER_COMMIT, "matrix semantic owner changed")
    require(
        matrix.get("production_source_baseline_commit") == PRODUCTION_SOURCE_BASELINE_COMMIT,
        "matrix production baseline changed",
    )
    require(matrix.get("capture_head") == before_commit, "matrix capture HEAD changed")
    require(matrix.get("recipe_count") == 23 and matrix.get("lane_count") == 42, "matrix count receipt changed")
    require(matrix.get("accepted_lanes") == 42, "correctness matrix did not accept every lane")
    require(matrix.get("failed_lanes") == 0, "correctness matrix reports failed lanes")
    require((results / "commands.jsonl").is_file(), "command metadata missing")
    require((results / "build-provenance.txt").is_file(), "build provenance missing")
    require((results / "host.txt").is_file(), "host metadata missing")
    require((results / "source-provenance.txt").is_file(), "source provenance missing")
    require((results / "committed-inputs-guard.stdout").is_file(), "committed input guard output missing")
    verify_provenance_file(results / "build-provenance.txt", before_commit)
    verify_provenance_file(results / "source-provenance.txt", before_commit)
    verify_guard_output(results / "committed-inputs-guard.stdout", before_commit)
    command_records = [
        json.loads(line)
        for line in (results / "commands.jsonl").read_text().splitlines()
        if line.strip()
    ]
    starts: dict[str, dict[str, object]] = {}
    exits: dict[str, dict[str, object]] = {}
    for record in command_records:
        label = record.get("label")
        require(isinstance(label, str) and label, "command record label is missing")
        event = record.get("event")
        require(event in {"start", "exit"}, f"command event is malformed: {label}")
        if event == "start":
            require(label not in starts, f"duplicate command start: {label}")
            starts[label] = record
        else:
            require(label not in exits, f"duplicate command exit: {label}")
            exits[label] = record
    require(set(starts) == set(exits), "command start/exit records are incomplete")
    labels = set(starts)
    require("cargo-metadata-before" in labels, "metadata command provenance missing")
    require("committed-inputs-guard" in labels, "owner input guard provenance missing")
    require("cargo-build-release" in labels, "build command provenance missing")
    require("host-probe" in labels and "correctness-matrix" in labels, "host preflight provenance missing")
    require("cargo-metadata-after" in labels and "verify" in labels, "postflight command provenance missing")
    lane_labels = {
        f"lane-{lane_id}-p{process}"
        for lane_id in lanes
        for process in (1, 2, 3)
    }
    require(lane_labels <= labels, "lane command metadata is incomplete")
    for label, record in starts.items():
        if label.startswith("lane-"):
            argv = record.get("argv")
            require(isinstance(argv, list), f"lane argv metadata is malformed: {label}")
            require("/usr/bin/time" in argv and "-v" in argv, f"lane is not timed with /usr/bin/time -v: {label}")
            require("--warmup" in argv and "--samples" in argv, f"lane sample gate missing: {label}")
    for label, record in starts.items():
        require(record.get("cwd") == str(root), "command cwd is not the isolated root")
        environment = record.get("environment", {})
        require(environment.get("CARGO_INCREMENTAL") == "0", "command metadata missed CARGO_INCREMENTAL=0")
        require(environment.get("LC_ALL") == "C", "command metadata missed LC_ALL=C")
        require(
            environment.get("PPTX_INK_ACTIONS_CAPTURE_HEAD") == before_commit,
            "command metadata missed the exact capture HEAD",
        )
        require(
            isinstance(environment.get("CARGO_TARGET_DIR"), str)
            and environment["CARGO_TARGET_DIR"],
            "command metadata missed CARGO_TARGET_DIR",
        )
        require(
            all(environment.get(name) is None for name in (
                "RUSTFLAGS",
                "CARGO_ENCODED_RUSTFLAGS",
                "RUSTC_BOOTSTRAP",
                "RUSTDOCFLAGS",
            )),
            "command metadata records inherited Rust flags",
        )
        exit_record = exits[label]
        require(exit_record.get("cwd") == str(root), f"command exit cwd changed: {label}")
        require(exit_record.get("pid") == record.get("pid"), f"command PID provenance changed: {label}")
        require(exit_record.get("exit_code") == 0, f"command did not exit zero: {label}")
        outputs = record.get("outputs")
        exit_outputs = exit_record.get("outputs")
        require(isinstance(outputs, dict) and outputs == exit_outputs, f"command outputs are incomplete: {label}")
        for stream in ("stdout", "stderr"):
            output = outputs.get(stream)
            if output is not None:
                require(isinstance(output, str) and Path(output).is_file(), f"command {stream} output is missing: {label}")
        require(
            isinstance(record.get("started_unix_ns"), int)
            and isinstance(exit_record.get("finished_unix_ns"), int)
            and exit_record["finished_unix_ns"] >= record["started_unix_ns"],
            f"command timestamps are malformed: {label}",
        )
    require((results / "matrix-correctness.stderr").read_text() == "", "matrix preflight wrote stderr")
    binary_before = results / "binary.sha256"
    binary_after = results / "binary-after.sha256"
    require(binary_before.read_text() == binary_after.read_text(), "binary changed during capture")
    require((results / "git-status-before.txt").read_text() == "", "production checkout dirty before capture")
    require((results / "git-status-after.txt").read_text() == "", "production checkout mutated during capture")

    command_provenance = []
    for label in sorted(starts):
        start = starts[label]
        exit_record = exits[label]
        command_provenance.append(
            {
                "label": label,
                "cwd": start["cwd"],
                "pid": start["pid"],
                "argv": start["argv"],
                "environment": start["environment"],
                "started_unix_ns": start["started_unix_ns"],
                "finished_unix_ns": exit_record["finished_unix_ns"],
                "exit_code": exit_record["exit_code"],
                "outputs": start["outputs"],
            }
        )

    all_rss: list[int] = []
    all_latency: list[int] = []
    all_operation_allocations: list[int] = []
    all_time_fields: list[dict[str, object]] = []
    missing: list[str] = []
    lane_process_ids: dict[str, set[int]] = {}
    for lane_id, lane in lanes.items():
        recipe = recipes[lane["recipe"]]
        for process in (1, 2, 3):
            receipt = results / f"{lane_id}-p{process}.json"
            timing = results / f"{lane_id}-p{process}.time.txt"
            stderr = results / f"{lane_id}-p{process}.stderr.log"
            if not receipt.exists() or not timing.exists() or not stderr.exists():
                missing.append(f"{lane_id}-p{process}")
                continue
            process_id, rss_values, latency_values, allocation_values, timing_values = verify_receipt(
                receipt, timing, stderr, recipe, lane, generator_sha256, before_commit
            )
            lane_process_ids.setdefault(lane_id, set()).add(process_id)
            all_rss.extend(rss_values)
            all_latency.extend(latency_values)
            all_operation_allocations.extend(allocation_values)
            all_time_fields.append(
                {
                    "lane": lane_id,
                    "process": process,
                    "process_id": process_id,
                    **timing_values,
                }
            )
    require(not missing, "missing lane/process receipts: " + ", ".join(missing))
    require(
        all(len(process_ids) == 3 for process_ids in lane_process_ids.values()),
        "each lane must have three fresh process IDs",
    )
    report = {
        "schema": "pptx-ink-actions-performance-verification-v1",
        "source_commit": SEMANTIC_OWNER_COMMIT,
        "semantic_owner_commit": SEMANTIC_OWNER_COMMIT,
        "production_source_baseline_commit": PRODUCTION_SOURCE_BASELINE_COMMIT,
        "semantic_owner_design": {
            "path": SEMANTIC_OWNER_DESIGN_PATH,
            "sha256": SEMANTIC_OWNER_DESIGN_SHA256,
            "git_blob": SEMANTIC_OWNER_DESIGN_GIT_BLOB,
        },
        "capture_head": before_commit,
        "helper_sha256": HELPER_SHA256,
        "recipes": len(recipes),
        "lanes": len(lanes),
        "fresh_processes": 3,
        "warmups": 2,
        "samples_per_process": 20,
        "receipt_count": len(lanes) * 3,
        "host": host_metadata,
        "command_provenance": command_provenance,
        "rss_kib": {
            "min": min(all_rss),
            "p50": sorted(all_rss)[math.ceil(len(all_rss) * 0.50) - 1],
            "p95": sorted(all_rss)[math.ceil(len(all_rss) * 0.95) - 1],
            "p99": sorted(all_rss)[math.ceil(len(all_rss) * 0.99) - 1],
            "max": max(all_rss),
        },
        "operation_elapsed_ns": {
            "p50": sorted(all_latency)[math.ceil(len(all_latency) * 0.50) - 1],
            "p95": sorted(all_latency)[math.ceil(len(all_latency) * 0.95) - 1],
            "p99": sorted(all_latency)[math.ceil(len(all_latency) * 0.99) - 1],
        },
        "operation_requested_alloc_bytes": {
            "p50": sorted(all_operation_allocations)[math.ceil(len(all_operation_allocations) * 0.50) - 1],
            "p95": sorted(all_operation_allocations)[math.ceil(len(all_operation_allocations) * 0.95) - 1],
            "p99": sorted(all_operation_allocations)[math.ceil(len(all_operation_allocations) * 0.99) - 1],
        },
        "process_time_fields": all_time_fields,
        "timing_claim": "descriptive absolute latency and allocation receipts only; no native PowerPoint or speedup claim",
        "evidence_root": str(evidence),
    }
    args.output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n")
    print(json.dumps(report, sort_keys=True))


if __name__ == "__main__":
    main()
