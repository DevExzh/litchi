#!/usr/bin/env python3
"""Prepare and attest the source inputs for a legacy native qualification.

This command is intentionally source-only.  It does not invoke Cargo, build a
binary, or run a native probe.  The external qualification setup may apply a
reviewed consumer adapter to a baseline checkout; this script records the
resulting source state and rejects dirty paths outside the lane's explicit
allow-list.  The runner and the receipt validator bind the resulting receipt
to every later build/run event.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys
from typing import Any


SCHEMA = "litchi.legacy-native-preparation.v1"
ROOT = Path(__file__).resolve().parents[2]

LANES: dict[str, dict[str, Any]] = {
    "legacy-rtf": {
        "baseline": "86d708878fa537eab42fd4bd5ffbc0f79c606191",
        "probe": "crates/litchi-rtf/examples/refusal_contract_bench.rs",
        "probe_sha256": {
            "before": "c5deeaff97e028134d378dccbdd423963eeaceec45f93b7c83285b9b799b6bea",
            "after": "c5deeaff97e028134d378dccbdd423963eeaceec45f93b7c83285b9b799b6bea",
        },
        "lock": "Cargo.lock",
        "fixtures": [],
        "allowed_dirty": [
            "Cargo.lock",
            "crates/litchi-rtf/examples/refusal_contract_bench.rs",
        ],
    },
    "legacy-ole-properties": {
        "baseline": "604822b44107999aaddb0bffb96b705e9464f546",
        "probe": "tools/perf-baseline/examples/ole_w5_probe.rs",
        "probe_sha256": {
            "before": "3b7cde72762a74796250df572704206baf9e438cca3d9ed11119af783fc8488f",
            "after": "3b7cde72762a74796250df572704206baf9e438cca3d9ed11119af783fc8488f",
        },
        "lock": "tools/perf-baseline/Cargo.lock",
        "fixtures": [],
        "allowed_dirty": [
            "tools/perf-baseline/Cargo.toml",
            "tools/perf-baseline/Cargo.lock",
            "tools/perf-baseline/examples/ole_w5_probe.rs",
        ],
    },
    "legacy-xls-properties": {
        "baseline": "509bf28e7531ad75f278eaec3a732714a243915a",
        "probe": "crates/litchi-ops/examples/ole_metadata_bench.rs",
        "probe_sha256": {
            "before": "b840acb620c09644ca4b89cbe42927485da9d457f3bf93b578eff90fa1008173",
            "after": "53b26136e915fa0bf88518fd1c05e494d55ce5e8f91012be04ac7adf476c1f10",
        },
        "lock": "Cargo.lock",
        "fixtures": [
            "test-data/ole/doc/FancyFoot.doc",
            "test-data/ole/xls/Simple.xls",
            "test-data/ole/ppt/SampleShow.ppt",
        ],
        "allowed_dirty": [
            "Cargo.lock",
            "crates/litchi-ops/examples/ole_metadata_bench.rs",
        ],
    },
}


class PreparationError(ValueError):
    """The source checkout does not match the reviewed preparation contract."""


def _expect(condition: bool, message: str) -> None:
    if not condition:
        raise PreparationError(message)


def _sha256(path: Path) -> tuple[str, int]:
    _expect(not path.is_symlink() and path.is_file(), f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise PreparationError(f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}")
    return result.stdout


def _identity_file(root: Path, relative: str) -> dict[str, Any]:
    raw_path = root / relative
    _expect(not raw_path.is_symlink(), f"identity input may not be a symlink: {relative}")
    path = raw_path.resolve()
    _expect(path.parent == (root / relative).parent.resolve(), f"identity path escapes source root: {relative}")
    digest, size = _sha256(path)
    return {"path": str(path), "relative": relative, "bytes": size, "sha256": digest}


def _source_identity(root: Path, lane: str, leg: str, spec: dict[str, Any]) -> dict[str, Any]:
    root = root.resolve()
    _expect(root.is_dir(), f"source root is not a directory: {root}")
    top = Path(_git(root, "rev-parse", "--show-toplevel").strip()).resolve()
    _expect(top == root, f"source root is not the worktree top-level: {root}")
    head = _git(root, "rev-parse", "HEAD").strip()
    tree = _git(root, "rev-parse", "HEAD^{tree}").strip()
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    allowed = set(spec["allowed_dirty"])
    status_paths: list[str] = []
    for line in status.splitlines():
        _expect(len(line) >= 4, f"malformed git status entry in {root}: {line!r}")
        value = line[3:]
        if " -> " in value:
            value = value.split(" -> ", 1)[1]
        value = value.strip()
        status_paths.append(value)
        _expect(value in allowed, f"unreviewed dirty path in {root}: {value}")
    relative_files = [spec["lock"], spec["probe"], *spec["fixtures"]]
    files = [_identity_file(root, relative) for relative in relative_files]
    probe_record = next(record for record in files if record["relative"] == spec["probe"])
    _expect(probe_record["sha256"] == spec["probe_sha256"][leg], f"{lane}/{leg} probe is not the reviewed generator")
    file_digest = hashlib.sha256()
    for record in sorted(files, key=lambda item: (str(Path(item["path"]).parent), Path(item["path"]).name)):
        relative = Path(record["path"]).name
        file_digest.update(relative.encode("utf-8"))
        file_digest.update(b"\0")
        file_digest.update(str(record["bytes"]).encode("ascii"))
        file_digest.update(b"\0")
        file_digest.update(bytes.fromhex(record["sha256"]))
    return {
        "source_root": str(root),
        "head": head,
        "tree": tree,
        "clean": not bool(status),
        "status": status,
        "status_paths": status_paths,
        "status_sha256": hashlib.sha256(status.encode("utf-8")).hexdigest(),
        "diff_sha256": hashlib.sha256(_git(root, "diff", "--no-ext-diff", "--binary").encode("utf-8")).hexdigest(),
        "source_tree_sha256": hashlib.sha256(
            (tree + "\0" + status + "\0" + _git(root, "diff", "--no-ext-diff", "--binary")).encode("utf-8")
        ).hexdigest(),
        "lock_paths": [str((root / spec["lock"]).resolve())],
        "probe_paths": [str((root / spec["probe"]).resolve())],
        "fixture_paths": [str((root / value).resolve()) for value in spec["fixtures"]],
        "identity_sha256": file_digest.hexdigest(),
        "identity_files": files,
    }


def prepare(lane: str, manifest: Path, before: Path, after: Path, output: Path) -> None:
    _expect(lane in LANES, f"unknown legacy lane: {lane}")
    spec = LANES[lane]
    manifest = manifest.resolve()
    manifest_digest, manifest_bytes = _sha256(manifest)
    try:
        manifest_data = json.loads(manifest.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise PreparationError(f"cannot read qualification manifest: {error}") from error
    _expect(isinstance(manifest_data, dict) and manifest_data.get("lane") == lane, "manifest lane does not match preparation lane")
    _expect(manifest_data.get("schema") == "litchi.native-qualification.v1", "manifest schema is not the runner schema")
    before_identity = _source_identity(before, lane, "before", spec)
    after_identity = _source_identity(after, lane, "after", spec)
    _expect(before_identity["head"] == spec["baseline"], "before source is not the reviewed historical baseline")
    _expect(after_identity["head"] != spec["baseline"], "after source must be a distinct candidate")
    preparer = Path(__file__).resolve()
    preparer_digest, _ = _sha256(preparer)
    receipt = {
        "schema": SCHEMA,
        "lane": lane,
        "manifest": {"path": str(manifest), "sha256": manifest_digest, "bytes": manifest_bytes},
        "preparer": {"path": str(preparer), "sha256": preparer_digest},
        "reviewed_inputs": {
            "probe_relative_path": spec["probe"],
            "probe_sha256": dict(spec["probe_sha256"]),
            "lock_relative_path": spec["lock"],
            "fixture_relative_paths": list(spec["fixtures"]),
            "allowed_dirty_paths": list(spec["allowed_dirty"]),
        },
        "identities": {"before": before_identity, "after": after_identity},
    }
    output = output.resolve()
    _expect(not output.exists() and not output.is_symlink(), f"refusing to overwrite preparation receipt: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as stream:
        json.dump(receipt, stream, indent=2, sort_keys=True)
        stream.write("\n")


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--lane", choices=sorted(LANES), required=True)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--before", type=Path, required=True)
    parser.add_argument("--after", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        prepare(args.lane, args.manifest, args.before, args.after, args.output)
    except (OSError, PreparationError) as error:
        print(f"legacy native preparation refused: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
