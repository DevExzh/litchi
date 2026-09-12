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


SOURCE_COMMIT = "cf6fdb8e91dd232d7d762596763d2e9d8a5b9dbd"
HELPER_SHA256 = "bec6baafcf735d778216fb54fe6299312707e912dcaed5f89de988f7112bb58e"
SCHEMA = "pptx-ink-actions-performance-v1"
SOURCE_FORMAT = "pptx-ink-actions-profile-build-source-v1"
OWNER_WORKSPACE_EXTRAS = {
    "Cargo.toml",
    "rust-toolchain.toml",
    ".cargo/config.toml",
    "rustfmt.toml",
    "clippy.toml",
    "deny.toml",
}


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


def verify_manifest(path: Path, root: Path) -> str:
    lines = path.read_text().splitlines()
    require(lines and lines[0] == f"format={SOURCE_FORMAT}", f"source manifest format: {path}")
    commit: str | None = None
    owner_commit: str | None = None
    packages: dict[tuple[str, str, str], tuple[str, int, str]] = {}
    files: dict[tuple[str, str, str], list[tuple[str, str]]] = {}
    extras: list[tuple[str, str]] = []
    owner_extras: list[tuple[str, str]] = []
    for line in lines[1:]:
        if line.startswith("git_commit="):
            require(commit is None, "duplicate source manifest git_commit")
            commit = line.split("=", 1)[1]
        elif line.startswith("owner_commit="):
            require(owner_commit is None, "duplicate source manifest owner_commit")
            owner_commit = line.split("=", 1)[1]
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
        elif line.startswith("owner_extra=\t"):
            fields = line.split("\t")
            require(len(fields) == 3, f"malformed owner extra line: {line}")
            owner_extras.append((fields[1], fields[2]))
        else:
            raise AssertionError(f"unknown source manifest line: {line}")
    require(commit is not None and len(commit) == 40, "source manifest commit is malformed")
    require(owner_commit == SOURCE_COMMIT, "source manifest owner pin changed")
    require(set(packages) == set(files), "source package/file keys differ")
    local_paths: list[Path] = []
    checked = 0
    for key, (source, count, tree) in packages.items():
        entries = files[key]
        require(len(entries) == count, f"source package file count changed: {key}")
        payload = "\n".join(f"{name}\t{digest}" for name, digest in entries)
        require(hashlib.sha256(payload.encode()).hexdigest() == tree, f"source package tree changed: {key}")
        for shown, digest in entries:
            path_value = resolve(root, shown)
            require(path_value.is_file(), f"source file missing: {shown}")
            require(sha256(path_value) == digest, f"source file changed: {shown}")
            if source == "path":
                local_paths.append(path_value)
            checked += 1
    for shown, digest in extras:
        path_value = resolve(root, shown)
        require(path_value.is_file(), f"source extra missing: {shown}")
        require(sha256(path_value) == digest, f"source extra changed: {shown}")
        local_paths.append(path_value)
        checked += 1
    owner_extra_names = {shown for shown, _ in owner_extras}
    require(
        owner_extra_names == OWNER_WORKSPACE_EXTRAS,
        f"workspace owner extras are incomplete: {sorted(owner_extra_names)}",
    )
    for shown, digest in owner_extras:
        path_value = resolve(root, shown)
        require(path_value.is_file(), f"owner source extra missing: {shown}")
        require(sha256(path_value) == digest, f"owner source extra changed: {shown}")
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
    for shown, digest in owner_extras:
        path_value = resolve(root, shown)
        require(
            sha256(path_value) == blob_sha256(root, owner_commit, shown),
            f"owner workspace input differs from approved tree: {shown}",
        )
    return commit


def verify_corpus(path: Path) -> tuple[dict[str, dict[str, object]], dict[str, dict[str, object]], str]:
    corpus = json.loads(path.read_text())
    require(corpus["owner_commit"] == SOURCE_COMMIT, "corpus owner pin changed")
    root = path.parent.parent.parent.parent.parent
    fixture = corpus["fixture_authority"]
    fixture_path = root / str(fixture["path"])
    require(fixture_path.is_file(), "fixture authority is missing")
    require(sha256(fixture_path) == fixture["sha256"], "fixture authority hash changed")
    fixture_blob = subprocess.check_output(
        ["git", "-C", str(root), "rev-parse", f"{SOURCE_COMMIT}:{fixture['path']}"],
        text=True,
    ).strip()
    require(fixture_blob == fixture["git_blob"], "fixture authority Git blob changed")
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
) -> tuple[int, list[int], list[int], list[int], dict[str, str]]:
    receipt = json.loads(path.read_text())
    require(receipt["schema"] == SCHEMA, f"receipt schema changed: {path}")
    require(receipt["lane"] == lane["id"], f"receipt lane mismatch: {path}")
    require(receipt["recipe_id"] == recipe["id"], f"receipt recipe mismatch: {path}")
    require(receipt["source_commit"] == SOURCE_COMMIT, f"receipt owner pin changed: {path}")
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


def verify_host_probe(path: Path) -> None:
    probe = json.loads(path.read_text())
    require(probe["schema"] == "pptx-ink-actions-host-probe-v1", "host probe schema changed")
    require(probe["source_commit"] == SOURCE_COMMIT, "host probe source changed")
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


def verify_host_metadata(path: Path) -> None:
    lines = path.read_text().splitlines()
    required = (
        "source_commit=",
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


def read_host_metadata(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    for line in path.read_text().splitlines():
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key] = value
    return values


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
    before_commit = verify_manifest(args.manifest.resolve(), root)
    after_commit = verify_manifest(args.manifest_after.resolve(), root)
    require(before_commit == after_commit, "source manifests use different commits")
    require(
        subprocess.run(
            ["git", "-C", str(root), "merge-base", "--is-ancestor", SOURCE_COMMIT, before_commit],
            check=False,
        ).returncode == 0,
        "full source manifest commit does not descend from approved owner pin",
    )
    require(args.metadata_before.read_bytes() == args.metadata_after.read_bytes(), "Cargo metadata changed")
    verify_host_probe(results / "host-probe.json")
    verify_host_metadata(results / "host.txt")
    host_metadata = read_host_metadata(results / "host.txt")
    require((results / "host-probe.stderr").read_text() == "", "host probe wrote stderr")
    require((results / "matrix-correctness.json").is_file(), "correctness matrix receipt missing")
    matrix = json.loads((results / "matrix-correctness.json").read_text())
    require(matrix.get("timings_collected") is False, "matrix correctness receipt claims timing")
    require(matrix.get("recipe_count") == 23 and matrix.get("lane_count") == 42, "matrix count receipt changed")
    require(matrix.get("accepted_lanes") == 42, "correctness matrix did not accept every lane")
    require(matrix.get("failed_lanes") == 0, "correctness matrix reports failed lanes")
    require((results / "commands.jsonl").is_file(), "command metadata missing")
    require((results / "build-provenance.txt").is_file(), "build provenance missing")
    require((results / "host.txt").is_file(), "host metadata missing")
    require((results / "source-provenance.txt").is_file(), "source provenance missing")
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
                receipt, timing, stderr, recipe, lane, generator_sha256
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
        "source_commit": SOURCE_COMMIT,
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
