#!/usr/bin/env python3
"""Overlay the sealed 0890 MCE candidate onto the retained diagnostic source."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import shutil
import stat


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
SOURCE = Path("/tmp/litchi-0890-correctness-source")
FREEZE = PACKET / "candidate-freeze-v2.json"
PATCH = PACKET / "candidate-v2.patch"
SOURCE_REVIEW = PACKET / "source-review-v2.json"
BASELINE_RECEIPT = PACKET / "attempts/baseline-f790c2228c-retry1/baseline-receipt.json"
OUTPUT = PACKET / "candidate-diagnostic-overlay.json"
BASE_MARKER = SOURCE / ".owner-0890-source"
CANDIDATE_PATHS = (
    "crates/litchi-ooxml-common/src/mce/codec.rs",
    "crates/litchi-ooxml-common/src/mce/tests.rs",
)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def identity(path: Path) -> dict[str, object]:
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "mode": oct(stat.S_IMODE(path.stat().st_mode)),
        "sha256": sha(path),
    }


def read_json(path: Path) -> dict[str, object]:
    value = json.loads(path.read_text())
    if not isinstance(value, dict):
        raise RuntimeError(f"expected JSON object: {path}")
    return value


def main() -> int:
    if OUTPUT.exists():
        raise SystemExit(f"create-only overlay receipt already exists: {OUTPUT}")
    if not SOURCE.is_dir() or not BASE_MARKER.is_file():
        raise SystemExit(f"owned diagnostic source is missing: {SOURCE}")
    freeze = read_json(FREEZE)
    if freeze.get("base_commit") != "f790c2228cbd9e89c2892ba3ffaeb49802fac43b":
        raise SystemExit("candidate freeze base does not match diagnostic baseline")
    frozen = {entry["path"]: entry for entry in freeze.get("files", [])}
    if set(frozen) != set(CANDIDATE_PATHS):
        raise SystemExit(f"candidate path set mismatch: {sorted(frozen)}")
    review = read_json(SOURCE_REVIEW)
    patch = review.get("patch")
    if not isinstance(patch, dict) or patch.get("paths") != list(CANDIDATE_PATHS):
        raise SystemExit("source review does not bind the expected candidate paths")
    if patch.get("sha256") != sha(PATCH) or patch.get("bytes") != PATCH.stat().st_size:
        raise SystemExit("candidate patch identity differs from source review")
    baseline = read_json(BASELINE_RECEIPT)
    if baseline.get("status") != "pass":
        raise SystemExit("baseline receipt is not passing")
    base_files = {entry["path"]: entry for entry in baseline["source"]["files"]}
    before: dict[str, dict[str, object]] = {}
    candidate_sources: dict[str, dict[str, object]] = {}
    for path in CANDIDATE_PATHS:
        destination = SOURCE / path
        source = ROOT / path
        if identity(destination)["sha256"] != base_files[path]["sha256"]:
            raise SystemExit(f"diagnostic source was changed before overlay: {path}")
        frozen_identity = frozen[path]
        source_identity = identity(source)
        if source_identity["bytes"] != frozen_identity["bytes"] or source_identity["sha256"] != frozen_identity["sha256"]:
            raise SystemExit(f"root candidate differs from freeze receipt: {path}")
        before[path] = identity(destination)
        candidate_sources[path] = source_identity

    for path in CANDIDATE_PATHS:
        destination = SOURCE / path
        source = ROOT / path
        shutil.copyfile(source, destination)
        destination.chmod(source.stat().st_mode & 0o7777)

    after = {path: identity(SOURCE / path) for path in CANDIDATE_PATHS}
    for path in CANDIDATE_PATHS:
        frozen_identity = frozen[path]
        if after[path]["bytes"] != frozen_identity["bytes"] or after[path]["sha256"] != frozen_identity["sha256"]:
            raise SystemExit(f"candidate overlay differs from freeze receipt: {path}")

    unchanged = True
    for path, entry in base_files.items():
        current = identity(SOURCE / path)
        if path in CANDIDATE_PATHS:
            continue
        if current["bytes"] != entry["bytes"] or current["sha256"] != entry["sha256"]:
            unchanged = False
            raise SystemExit(f"non-candidate diagnostic source changed: {path}")
    result = {
        "schema": "litchi.mce-0890-diagnostic-overlay.v1",
        "status": "pass",
        "base_commit": freeze["base_commit"],
        "source_root": str(SOURCE),
        "baseline_receipt": {"path": str(BASELINE_RECEIPT.relative_to(PACKET)), "sha256": sha(BASELINE_RECEIPT)},
        "candidate_freeze": {"path": str(FREEZE.relative_to(PACKET)), "sha256": sha(FREEZE)},
        "candidate_patch": {"path": str(PATCH.relative_to(PACKET)), "bytes": PATCH.stat().st_size, "sha256": sha(PATCH)},
        "candidate_paths": list(CANDIDATE_PATHS),
        "before": before,
        "root_candidate": candidate_sources,
        "after": after,
        "unchanged_non_candidate_sources": unchanged,
        "source_marker": identity(BASE_MARKER),
    }
    OUTPUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
