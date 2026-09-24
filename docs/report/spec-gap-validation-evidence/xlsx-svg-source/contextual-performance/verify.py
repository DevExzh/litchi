#!/usr/bin/env python3
"""Verify bounded XLSX source-profile receipts and generated report."""

from __future__ import annotations

import hashlib
import json
import argparse
from pathlib import Path

from summarize import LANES, rows


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def verify_manifest(path: Path, root: Path) -> int:
    lines = path.read_text().splitlines()
    require(lines and lines[0] == "format=xlsx-svg-contextual-profile-build-source-v1", "manifest format changed")
    packages: dict[tuple[str, str, str], tuple[int, str]] = {}
    files: dict[tuple[str, str, str], list[tuple[str, str]]] = {}
    extras: list[tuple[str, str]] = []
    for line in lines[1:]:
        parts = line.split("\t")
        if line.startswith("metadata_sha256="):
            continue
        if line.startswith("package="):
            require(len(parts) == 7, f"malformed package line: {line}")
            key = (parts[0][len("package=") :], parts[1], parts[3])
            require(key not in packages, f"duplicate package line: {key}")
            packages[key] = (int(parts[5]), parts[6])
        elif line.startswith("file="):
            require(len(parts) == 5, f"malformed file line: {line}")
            key = (parts[0][len("file=") :], parts[1], parts[2])
            files.setdefault(key, []).append((parts[3], parts[4]))
        elif line.startswith("extra=\t"):
            require(len(parts) == 3, f"malformed extra line: {line}")
            extras.append((parts[1], parts[2]))
        else:
            raise AssertionError(f"unknown manifest line: {line}")
    require(set(packages) == set(files), "package/file manifest sets differ")
    checked = 0
    for key, (expected_count, expected_tree) in packages.items():
        entries = files[key]
        require(len(entries) == expected_count, f"package file count changed: {key}")
        tree_payload = "\n".join(f"{shown}\t{digest}" for shown, digest in entries)
        require(hashlib.sha256(tree_payload.encode()).hexdigest() == expected_tree, f"package tree changed: {key}")
        for shown, expected in entries:
            current = Path(shown) if Path(shown).is_absolute() else root / shown
            require(current.is_file(), f"retained manifest path missing: {shown}")
            require(sha256(current) == expected, f"retained manifest hash changed: {shown}")
            checked += 1
    for shown, expected in extras:
        current = Path(shown) if Path(shown).is_absolute() else root / shown
        require(current.is_file(), f"retained extra path missing: {shown}")
        require(sha256(current) == expected, f"retained extra hash changed: {shown}")
        checked += 1
    require(checked >= 8, f"too few retained manifest inputs checked: {checked}")
    return checked


def verify_sample(sample: dict[str, object], expected_success: bool, path: Path) -> None:
    require(sample["expected_success"] is expected_success, f"sample expected status mismatch: {path}")
    require(sample["actual_success"] is expected_success, f"sample actual status mismatch: {path}")
    require(sample["semantic_ok"] is True, f"semantic gate failed: {path}")
    require(sample["output_exact"] is True, f"output gate failed: {path}")
    direct = int(sample["direct_allocated_bytes"])
    new = int(sample["realloc_new_bytes"])
    old = int(sample["realloc_old_bytes"])
    freed = int(sample["deallocated_bytes"])
    expected_live = int(sample["live_before"]) + direct + new - old - freed
    require(int(sample["requested_alloc_bytes"]) == direct + new, f"allocation equation failed: {path}")
    require(int(sample["live_after"]) == expected_live, f"live equation failed: {path}")
    require(sample["alloc_balance_ok"] is True, f"allocator balance failed: {path}")
    require(sample["alloc_invalid"] is False, f"allocator underflow failed: {path}")
    require(int(sample["alloc_failed"]) == 0, f"allocation failure failed: {path}")
    if expected_success:
        require(sample["error"] is None, f"unexpected profile error: {path}")
    else:
        require(int(sample["refusal_count"]) > 0, f"refusal lane did not refuse: {path}")


def verify_lanes(results: Path, snapshot: str) -> None:
    after = snapshot.startswith("after-")
    before = snapshot.startswith("before-")
    require(after or before, f"unknown snapshot label: {snapshot}")
    for lane in LANES:
        paths = sorted(results.glob(f"{lane}-p*.json"))
        require(len(paths) == 3, f"fresh process count mismatch for {lane}")
        expected = not lane.startswith("small_cap_refusal_")
        input_hashes: set[int] = set()
        input_sha_hashes: set[str] = set()
        for path in paths:
            value = json.loads(path.read_text())
            stderr = path.with_suffix(".stderr.txt")
            require(stderr.is_file(), f"lane stderr receipt missing: {stderr}")
            require(stderr.read_text() == "", f"lane stderr was not empty: {stderr}")
            require(value["schema"] == "xlsx-svg-contextual-profile-v1", f"schema mismatch: {path}")
            require(value["lane"] == lane, f"lane mismatch: {path}")
            require(int(value["warmup"]) >= 2, f"warm-up count below minimum: {path}")
            require(int(value["sample_count"]) == len(value["samples"]) >= 20, f"sample count below minimum: {path}")
            require(value["expected_success"] is expected, f"expected status mismatch: {path}")
            corpus_hash = str(value["input_hash_sha256"])
            require(len(corpus_hash) == 64 and all(character in "0123456789abcdef" for character in corpus_hash), f"corpus SHA-256 malformed: {path}")
            input_hashes.add(int(value["input_hash_fnv1a64"]))
            input_sha_hashes.add(corpus_hash)
            for sample in value["samples"]:
                verify_sample(sample, expected, path)
                pictures = int(value["picture_count"])
                if after:
                    require(int(sample["source_none_count"]) == pictures, f"after source retention was not lazy: {path}")
                    require(int(sample["context_present_count"]) == pictures, f"after context count mismatch: {path}")
                    require(sample["shared_context_identity"] is True, f"after context storage was not shared: {path}")
                    require(int(sample["context_distinct_count"]) == 1, f"after context storage count changed: {path}")
                elif before:
                    require(int(sample["source_none_count"]) == 0, f"before source projection unexpectedly dropped source: {path}")
                    require(int(sample["context_present_count"]) == 0, f"before context projection unexpectedly present: {path}")
                    require(int(sample["context_distinct_count"]) == 0, f"before context storage count changed: {path}")
                    require(sample["shared_context_identity"] is False, f"before context identity unexpectedly shared: {path}")
                if int(sample["context_present_count"]) > 0:
                    require(sample["shared_context_identity"] is True, f"context storage was not shared: {path}")
                    require(int(sample["context_distinct_count"]) == 1, f"context storage count changed: {path}")
                if lane.startswith(("standalone_export_", "scalar_reference_edit_")):
                    require(int(sample["readback_count"]) == int(value["picture_count"]), f"readback count mismatch: {path}")
                    require(sample["qname_preserved"] is True, f"opaque QName was lost: {path}")
                    require(sample["exact_embedded_reference"] is True, f"embedded reference was not exact: {path}")
                    require(int(sample["linked_reference_count"]) == 0, f"linked reference appeared: {path}")
            if lane.startswith("small_cap_refusal_"):
                for sample in value["samples"]:
                    require(int(sample["refusal_count"]) == int(value["picture_count"]), f"partial refusal: {path}")
        require(len(input_hashes) == 1, f"fixture hash changed across processes for {lane}")
        require(len(input_sha_hashes) == 1, f"corpus SHA-256 changed across processes for {lane}")


def verify_corpus_hashes(results: Path) -> None:
    lines = results.joinpath("corpus-sha256.tsv").read_text().splitlines()
    require(lines and lines[0] == "lane\tinput_bytes\tpictures\tnamespace_declarations\tsha256", "corpus hash header changed")
    parsed = {}
    for line in lines[1:]:
        fields = line.split("\t")
        require(len(fields) == 5, f"malformed corpus hash row: {line}")
        parsed[fields[0]] = fields[1:]
    require(set(parsed) == set(LANES), "corpus hash lane set differs")
    for lane in LANES:
        receipt = json.loads(next(results.glob(f"{lane}-p*.json")).read_text())
        expected = [
            str(receipt["input_bytes"]),
            str(receipt["picture_count"]),
            str(receipt["namespace_declarations"]),
            str(receipt["input_hash_sha256"]),
        ]
        require(parsed[lane] == expected, f"corpus hash row differs: {lane}")


def verify_provenance(results: Path, snapshot: str, expected_commit: str) -> None:
    provenance = results.joinpath("build-provenance.txt").read_text().splitlines()
    require(f"snapshot_label={snapshot}" in provenance, "build snapshot label missing")
    require(f"git_head={expected_commit}" in provenance, "build commit provenance mismatch")
    recorded = results.joinpath("binary.sha256").read_text().split()
    require(len(recorded) == 2 and len(recorded[0]) == 64, "binary hash receipt malformed")
    receipt = json.loads(results.joinpath("provenance-verification.json").read_text())
    require(receipt["passed"] is True, "binary provenance gate did not pass")
    require(receipt["snapshot_label"] == snapshot, "binary provenance label changed")
    require(receipt["expected_commit"] == expected_commit, "binary provenance commit changed")
    require(receipt["fresh_binary_validated_before_target_cleanup"] is True, "fresh binary validation receipt missing")


def verify_report(path: Path, recomputed: list[dict[str, object]]) -> None:
    parsed: dict[str, dict[str, object]] = {}
    for line in path.read_text().splitlines():
        if not line.startswith("| ") or line.startswith("|---") or line.startswith("| lane"):
            continue
        fields = [field.strip() for field in line.strip().strip("|").split("|")]
        require(len(fields) == 8, f"malformed report row: {line}")
        parsed[fields[0]] = {
            "processes": int(fields[1]), "samples": int(fields[2]), "input_bytes": int(fields[3]),
            "elapsed": tuple(int(item) for item in fields[4].split("/")),
            "alloc": tuple(int(item) for item in fields[5].split("/")),
            "peak": tuple(int(item) for item in fields[6].split("/")),
            "rss": tuple(int(item) for item in fields[7].split("/")),
        }
    require(set(parsed) == set(LANES), "report lane set differs")
    for row in recomputed:
        actual = parsed[str(row["lane"])]
        for field in ("processes", "samples", "input_bytes", "elapsed", "alloc", "peak", "rss"):
            require(actual[field] == row[field], f"report {field} differs: {row['lane']}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--snapshot", required=True)
    parser.add_argument("--expected-commit", required=True)
    arguments = parser.parse_args()
    here = Path(__file__).resolve().parent
    root = next(path for path in here.parents if (path / "crates").is_dir())
    results = here / "results"
    before = results / "source-manifest-before.txt"
    after = results / "source-manifest-after.txt"
    require(before.read_bytes() == after.read_bytes(), "source manifest changed during profile")
    checked = verify_manifest(before, root)
    verify_provenance(results, arguments.snapshot, arguments.expected_commit)
    verify_lanes(results, arguments.snapshot)
    verify_corpus_hashes(results)
    recomputed = rows(results)
    verify_report(here / "report.md", recomputed)
    result = {
        "passed": True,
        "lanes": len(LANES),
        "processes_per_lane": 3,
        "warmup_minimum": 2,
        "samples_minimum_per_process": 20,
        "source_manifest_sha256": sha256(before),
        "manifest_inputs_checked": checked,
        "status": "exploratory source-bound evidence; no lifecycle acceptance claim",
        "snapshot": arguments.snapshot,
        "expected_commit": arguments.expected_commit,
    }
    (here / "verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
