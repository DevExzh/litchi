#!/usr/bin/env python3
"""Verify sealed XLSX profile receipts and recompute the report."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path

from summarize import LANES, rows


REFUSAL_LANES = {
    "namespace_limit_refusal": "namespace_limit",
    "limit_small": "caller_limit",
    "limit_large": "caller_limit",
    "malformed_duplicate_owner": "duplicate_owner",
    "malformed_mce_owner": "mce_ancestry",
    "malformed_linked_owner": "linked_owner",
    "mixed_caps_rejection": "mixed_limit",
}

# The synthetic namespace fixture declares these seven root bindings before
# adding the generated declarations; keep aligned with its Rust constructor.
FIXED_ROOT_NAMESPACE_BINDINGS = 7


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


def verify_manifest(path: Path, root: Path) -> int:
    lines = path.read_text().splitlines()
    require(
        lines and lines[0] == "format=xlsx-svg-lifecycle-profile-build-source-v1",
        "manifest format changed",
    )
    packages: dict[tuple[str, str, str], tuple[int, str]] = {}
    files: dict[tuple[str, str, str], list[tuple[str, str]]] = {}
    extras: list[tuple[str, str]] = []
    for line in lines[1:]:
        if line.startswith("metadata_sha256="):
            continue
        parts = line.split("\t")
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
        require(
            hashlib.sha256(tree_payload.encode()).hexdigest() == expected_tree,
            f"package tree changed: {key}",
        )
        for shown, expected in entries:
            current = resolve(root, shown)
            require(current.is_file(), f"retained manifest path missing: {shown}")
            require(sha256(current) == expected, f"retained manifest hash changed: {shown}")
            checked += 1
    for shown, expected in extras:
        current = resolve(root, shown)
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
        error = sample["error"]
        require(isinstance(error, dict), f"refusal reason missing: {path}")
        require(error.get("class") == REFUSAL_LANES[path.stem.rsplit("-p", 1)[0]], f"refusal class mismatch: {path}")


def verify_namespace_boundary(value: dict[str, object], lane: str, path: Path) -> tuple[int, int, int] | None:
    fields = ("namespace_generated_bindings", "namespace_active_bindings", "namespace_active_limit")
    values = tuple(value.get(field) for field in fields)
    if lane != "namespace_limit_refusal":
        require(all(item is None for item in values), f"unexpected namespace boundary: {path}")
        return None
    require(
        all(type(item) is int and item > 0 for item in values),
        f"namespace boundary missing or malformed: {path}",
    )
    generated, active, limit = values
    require(
        active == generated + FIXED_ROOT_NAMESPACE_BINDINGS,
        f"namespace boundary does not include the fixture's inherited bindings: {path}",
    )
    require(active == limit + 1, f"namespace boundary is not first refused active count: {path}")
    return generated, active, limit


def verify_lanes(results: Path) -> None:
    for lane in LANES:
        paths = sorted(results.glob(f"{lane}-p*.json"))
        require(
            [path.name for path in paths] == [f"{lane}-p{n}.json" for n in range(1, 4)],
            f"fresh process identities mismatch for {lane}",
        )
        expected = lane not in REFUSAL_LANES
        input_hashes = set()
        input_sha256s = set()
        input_sizes = set()
        namespace_boundaries = set()
        for path in paths:
            verify_process_output(path)
            value = json.loads(path.read_text())
            require(value["schema"] == "xlsx-svg-lifecycle-profile-v1", f"schema mismatch: {path}")
            require(value["lane"] == lane, f"lane mismatch: {path}")
            namespace_boundaries.add(verify_namespace_boundary(value, lane, path))
            require(int(value["warmup"]) >= 2, f"warm-up count below minimum: {path}")
            require(
                int(value["sample_count"]) == len(value["samples"]) >= 20,
                f"sample count below minimum: {path}",
            )
            require(value["expected_success"] is expected, f"expected status mismatch: {path}")
            input_hashes.add(int(value["input_hash_fnv1a64"]))
            input_sizes.add(int(value["input_bytes"]))
            digest = value.get("input_sha256")
            require(
                isinstance(digest, str) and len(digest) == 64 and all(c in "0123456789abcdef" for c in digest),
                f"input SHA-256 missing or malformed: {path}",
            )
            input_sha256s.add(digest)
            for sample in value["samples"]:
                verify_sample(sample, expected, path)
        require(len(input_hashes) == 1, f"fixture hash changed across processes for {lane}")
        require(len(input_sha256s) == 1, f"fixture SHA-256 changed across processes for {lane}")
        require(len(input_sizes) == 1, f"fixture size changed across processes for {lane}")
        require(len(namespace_boundaries) == 1, f"namespace boundary changed across processes for {lane}")


def verify_process_output(path: Path) -> None:
    stderr = path.with_suffix(".stderr.log")
    require(stderr.is_file(), f"process stderr receipt missing: {path}")
    require(stderr.read_bytes() == b"", f"unexpected process stderr: {path}")
    timing = path.with_suffix(".time.txt")
    require(timing.is_file(), f"process timing receipt missing: {path}")
    statuses = [
        line.strip() for line in timing.read_text().splitlines()
        if line.strip().startswith("Exit status:")
    ]
    require(statuses == ["Exit status: 0"], f"process exit status is not successful: {path}")


def verify_binary_receipts(results: Path) -> str:
    before = (results / "binary.sha256").read_text()
    after = (results / "binary-after.sha256").read_text()
    require(before == after, "profile executable changed during measurement")
    fields = before.strip().split(maxsplit=1)
    require(len(fields) == 2, "malformed executable digest receipt")
    digest = fields[0]
    require(
        len(digest) == 64 and all(c in "0123456789abcdef" for c in digest),
        "malformed executable SHA-256",
    )
    # Targets are disposable; retain the measured identity without claiming
    # that the cleaned executable can be rehashed by a later verifier.
    return digest


def verify_native_identity(results: Path, corpus: dict[str, object], root: Path) -> None:
    fixture = corpus["source_fixture"]
    require(isinstance(fixture, dict), "native fixture identity missing from corpus")
    native = resolve(root, str(fixture["path"]))
    expected = str(fixture["sha256"])
    require(sha256(native) == expected, "native fixture differs from documented producer input")
    for number in range(1, 4):
        path = results / f"capture_native_fixture-p{number}.json"
        value = json.loads(path.read_text())
        require(value["input_sha256"] == expected, f"native receipt uses a different fixture: {path}")
        require(int(value["input_bytes"]) == native.stat().st_size, f"native receipt size mismatch: {path}")


def run_shape(results: Path) -> tuple[int, int]:
    first = json.loads(next(results.glob(f"{LANES[0]}-p*.json")).read_text())
    warmup = int(first["warmup"])
    samples = int(first["sample_count"])
    for lane in LANES:
        for path in results.glob(f"{lane}-p*.json"):
            value = json.loads(path.read_text())
            require(int(value["warmup"]) == warmup, f"warm-up differs across lanes: {path}")
            require(int(value["sample_count"]) == samples, f"sample count differs across lanes: {path}")
    return warmup, samples


def verify_report(path: Path, recomputed: list[dict[str, object]]) -> None:
    parsed = {}
    for line in path.read_text().splitlines():
        if not line.startswith("| ") or line.startswith("|---") or line.startswith("| lane"):
            continue
        fields = [field.strip() for field in line.strip().strip("|").split("|")]
        require(len(fields) == 8, f"malformed report row: {line}")
        elapsed = tuple(int(item) for item in fields[4].split("/"))
        alloc = tuple(int(item) for item in fields[5].split("/"))
        peak = tuple(int(item) for item in fields[6].split("/"))
        rss = tuple(int(item) for item in fields[7].replace("–", "-").split("-"))
        parsed[fields[0]] = {
            "processes": int(fields[1]),
            "samples": int(fields[2]),
            "input_bytes": int(fields[3]),
            "elapsed": elapsed,
            "alloc": alloc,
            "peak": peak,
            "rss": rss,
        }
    require(set(parsed) == set(LANES), "report lane set differs")
    for row in recomputed:
        actual = parsed[str(row["lane"])]
        for field in ("processes", "samples", "input_bytes", "elapsed", "alloc", "peak", "rss"):
            require(actual[field] == row[field], f"report {field} differs: {row['lane']}")


def main() -> None:
    here = Path(__file__).resolve().parent
    root = next(path for path in here.parents if (path / "crates").is_dir())
    results = here / "results"
    before = results / "source-manifest-before.txt"
    after = results / "source-manifest-after.txt"
    require(before.read_bytes() == after.read_bytes(), "source manifest changed during profile")
    manifest_sha = sha256(before)
    checked = verify_manifest(before, root)
    binary_sha = verify_binary_receipts(results)
    verify_lanes(results)
    verify_native_identity(results, json.loads((here / "corpus-manifest.json").read_text()), root)
    warmup, samples = run_shape(results)
    recomputed = rows(results)
    verify_report(here / "report.md", recomputed)
    result = {
        "passed": True,
        "lanes": len(LANES),
        "processes_per_lane": 3,
        "warmup_per_process": warmup,
        "samples_per_process": samples,
        "source_manifest_sha256": manifest_sha,
        "manifest_inputs_checked": checked,
        "binary_sha256": binary_sha,
        "binary_identity_stable_during_measurement": True,
        "expected_refusals": sorted(REFUSAL_LANES),
        "report_recomputed_from_samples": True,
    }
    (here / "verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
