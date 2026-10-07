#!/usr/bin/env python3
"""Bind the reviewed 0948 XLSX protocol after source and builds are ready.

This helper only writes ``plan-v2.json``.  It never launches Cargo or a workload.
Every source, helper, binary, and build receipt is bound by packet-relative
descriptor so the frozen plan can be copied with its external evidence.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent.parent
TOOLS = PACKET / "tools"
TEMPLATE = PACKET / "plan-v1-draft.json"
PLAN = PACKET / "plan-v2.json"
BASELINE_MANIFEST = PACKET / "baseline-source-manifest-v1.json"
CANDIDATE_MANIFEST = PACKET / "candidate-build-manifest-v3.json"
CANDIDATE_ROOT = "candidate-build-v3"
CANDIDATE_FREEZE_RECEIPT = PACKET / "receipts" / "candidate-freeze-v3.json"
BASELINE_REUSE_RECEIPT = PACKET / "receipts" / "baseline-reuse-v2.json"
BUILD_REVISIONS = {"baseline": "v1", "candidate": "v3"}
BUILD_BINDING_IDS = {"baseline": "baseline1", "candidate": "candidate2"}


class FreezeError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise FreezeError(message)


def strict_json(path: Path) -> Any:
    def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            require(key not in result, f"duplicate JSON key {key!r} in {path}")
            result[key] = value
        return result

    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicates)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise FreezeError(f"cannot read JSON {path}: {error}") from error


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def relative_path(path: Path) -> str:
    path = path.resolve()
    try:
        return path.relative_to(PACKET.resolve()).as_posix()
    except ValueError as error:
        raise FreezeError(f"bound path is outside packet: {path}") from error


def descriptor(path: Path) -> dict[str, Any]:
    path = Path(path)
    require(not path.is_symlink() and path.is_file(), f"bound artifact is not a regular file: {path}")
    path = path.resolve()
    require(path.is_relative_to(PACKET.resolve()), f"bound artifact escaped packet: {path}")
    mode = path.stat().st_mode & 0o777
    return {"path": relative_path(path), "bytes": path.stat().st_size,
            "mode": mode, "executable": bool(mode & 0o111), "sha256": sha256(path)}


def file_map(manifest_path: Path) -> dict[str, dict[str, Any]]:
    raw = strict_json(manifest_path)
    files = raw.get("files") if isinstance(raw, dict) else raw
    require(isinstance(files, list), f"{manifest_path} files must be a list")
    output: dict[str, dict[str, Any]] = {}
    for item in files:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                f"{manifest_path} has malformed file entry")
        name = item["path"]
        require(name not in output, f"duplicate manifest path {name}")
        path = Path(name)
        require(name and not path.is_absolute() and path.as_posix() == name
                and ".." not in path.parts, f"unsafe manifest path {name!r}")
        require(isinstance(item.get("bytes"), int) and item["bytes"] >= 0
                and isinstance(item.get("sha256"), str) and len(item["sha256"]) == 64,
                f"malformed manifest digest {name}")
        output[name] = item
    require(output, f"empty manifest {manifest_path}")
    return output


def source_binding(candidate_manifest: Path = CANDIDATE_MANIFEST,
                   candidate_root: str = CANDIDATE_ROOT,
                   candidate_freeze_receipt: Path | None = CANDIDATE_FREEZE_RECEIPT) -> dict[str, Any]:
    require(BASELINE_MANIFEST.is_file(), f"missing {BASELINE_MANIFEST}")
    require(BASELINE_REUSE_RECEIPT.is_file() and not BASELINE_REUSE_RECEIPT.is_symlink(),
            f"missing {BASELINE_REUSE_RECEIPT}")
    require(candidate_manifest.is_file(), f"missing {candidate_manifest}")
    root_value = Path(candidate_root)
    require(candidate_root and not root_value.is_absolute()
            and root_value.as_posix() == candidate_root and ".." not in root_value.parts,
            "candidate source root must be a normalized packet-relative path")
    baseline_raw = strict_json(BASELINE_MANIFEST)
    candidate_raw = strict_json(candidate_manifest)
    baseline = file_map(BASELINE_MANIFEST)
    candidate = file_map(candidate_manifest)
    require(set(baseline) == set(candidate), "source manifest path sets differ")
    changes = candidate_raw.get("changes")
    require(isinstance(changes, list) and changes, "candidate manifest has no frozen changes")
    changed = set()
    for item in changes:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                "candidate manifest change is malformed")
        changed.add(item["path"])
        name = item["path"]
        require(name in baseline and name in candidate
                and item.get("baseline_sha256") == baseline[name]["sha256"]
                and item.get("candidate_sha256") == candidate[name]["sha256"],
                f"candidate manifest change does not match file manifests: {name}")
    actual_changed = {name for name in baseline
                      if (baseline[name]["bytes"], baseline[name]["sha256"]) !=
                      (candidate[name]["bytes"], candidate[name]["sha256"])}
    require(actual_changed == changed and changed,
            f"candidate source delta does not match the candidate manifest: {sorted(actual_changed)}")
    source_export = strict_json(PACKET / "source-export-v1.json")
    require(isinstance(source_export, dict)
            and isinstance(source_export.get("source_commit"), str),
            "source-export-v1.json lacks source_commit")
    # The baseline manifest is the historical file-list export and therefore
    # carries no top-level census metadata.  Its commit, root, and file count
    # are bound by source-export-v1.json; the candidate build manifest carries
    # the corresponding metadata directly.
    require(candidate_raw.get("file_count") == len(candidate)
            and source_export.get("file_count") == len(baseline)
            and source_export.get("source_commit")
            and candidate_raw.get("base_commit") == source_export["source_commit"],
            "source manifest identity changed")
    if candidate_raw.get("source_root") is not None:
        require(candidate_raw["source_root"] == candidate_root,
                "candidate source root does not match the selected manifest")
    return {
        "baseline_manifest": descriptor(BASELINE_MANIFEST),
        "candidate_manifest": descriptor(candidate_manifest),
        "baseline_root": "baseline-source",
        "candidate_root": candidate_root,
        "candidate_files": sorted(changed),
        "file_count": len(candidate),
        "root_worktree": ".",
        "source_commit": source_export["source_commit"],
        "source_export": descriptor(PACKET / "source-export-v1.json"),
        "baseline_reuse_receipt": descriptor(BASELINE_REUSE_RECEIPT),
        **({"candidate_freeze_receipt": descriptor(candidate_freeze_receipt)}
           if candidate_freeze_receipt is not None else {}),
    }


def role_bindings() -> dict[str, dict[str, dict[str, Any]]]:
    """Load explicit role bindings; never silently reuse an older packet."""
    binding_file = PACKET / "role-bindings.json"
    if binding_file.is_file():
        raw = strict_json(binding_file)
        require(isinstance(raw, dict) and set(raw) == {"baseline", "candidate"},
                "role-bindings.json must contain exactly baseline and candidate")
        bindings = raw
    else:
        raise FreezeError(f"missing explicit role bindings: {binding_file}")
    output: dict[str, dict[str, dict[str, Any]]] = {}
    for role in ("baseline", "candidate"):
        require(isinstance(bindings.get(role), dict), f"missing {role} role bindings")
        output[role] = {}
        for kind in ("native", "allocator"):
            value = bindings[role].get(kind)
            require(isinstance(value, dict) and isinstance(value.get("path"), str)
                    and isinstance(value.get("build_receipt"), str)
                    and value.get("build_revision") == BUILD_REVISIONS[role]
                    and value.get("build_binding_id") == BUILD_BINDING_IDS[role],
                    f"missing {role}.{kind} binding")
            binary = PACKET / value["path"]
            receipt = PACKET / value["build_receipt"]
            require(binary.is_file() and not binary.is_symlink(),
                    f"missing {role}/{kind} binary: {binary}")
            require(os.access(binary, os.X_OK), f"{role}/{kind} binary is not executable")
            require(receipt.is_file() and not receipt.is_symlink(),
                    f"missing {role}/{kind} build receipt: {receipt}")
            binary_descriptor = descriptor(binary)
            binary_descriptor["build_receipt"] = descriptor(receipt)
            binary_descriptor["build_revision"] = BUILD_REVISIONS[role]
            binary_descriptor["build_binding_id"] = BUILD_BINDING_IDS[role]
            output[role][kind] = binary_descriptor
    return output


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--candidate-manifest", type=Path, default=CANDIDATE_MANIFEST,
                        help="packet-relative or absolute final candidate source manifest")
    parser.add_argument("--candidate-root", default=CANDIDATE_ROOT,
                        help="normalized packet-relative final candidate source root")
    parser.add_argument("--candidate-freeze-receipt", type=Path, default=CANDIDATE_FREEZE_RECEIPT,
                        help="optional packet-relative candidate freeze receipt to bind")
    args = parser.parse_args()
    require(not PLAN.exists(), f"refusing to overwrite frozen plan {PLAN}")
    plan = strict_json(TEMPLATE)
    require(plan.get("schema") == "litchi.paired-audit-capture-0948.v1"
            and plan.get("status") == "draft", "plan template is not the reviewed 0948 draft")
    require(plan.get("build_revisions") == BUILD_REVISIONS
            and plan.get("build_binding_ids") == BUILD_BINDING_IDS,
            "plan template build revision bindings changed")
    candidate_manifest = args.candidate_manifest
    if not candidate_manifest.is_absolute():
        candidate_manifest = PACKET / candidate_manifest
    candidate_freeze = args.candidate_freeze_receipt
    if candidate_freeze is not None and not candidate_freeze.is_absolute():
        candidate_freeze = PACKET / candidate_freeze
    plan["source"] = source_binding(candidate_manifest, args.candidate_root,
                                     candidate_freeze)
    plan["build_revisions"] = dict(BUILD_REVISIONS)
    plan["build_binding_ids"] = dict(BUILD_BINDING_IDS)
    plan["roles"] = role_bindings()
    plan["semantic_validator"] = descriptor(TOOLS / "semantic_0948_v1.py")
    plan["helpers"] = {
        "capture": descriptor(TOOLS / "capture_0948_v1.py"),
        "reducer": descriptor(TOOLS / "reduce_0948_v1.py"),
        "semantic": descriptor(TOOLS / "semantic_0948_v1.py"),
        "freeze": descriptor(Path(__file__).resolve()),
        "method": descriptor(PACKET / "METHOD-v2-draft.md"),
        "reuse_validator": descriptor(TOOLS / "reuse_0948_v1.py"),
    }
    plan["status"] = "frozen"
    encoded = (json.dumps(plan, indent=2, sort_keys=True, ensure_ascii=False,
                          allow_nan=False) + "\n").encode()
    try:
        with PLAN.open("xb") as output:
            output.write(encoded)
            output.flush()
            os.fsync(output.fileno())
    except FileExistsError as error:
        raise FreezeError(f"refusing to overwrite frozen plan {PLAN}") from error
    directory = os.open(PLAN.parent, os.O_RDONLY | os.O_DIRECTORY)
    try:
        os.fsync(directory)
    finally:
        os.close(directory)
    print(json.dumps({"status": "frozen", "plan": str(PLAN),
                      "sha256": sha256(PLAN), "candidate_files": plan["source"]["candidate_files"]},
                     sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except (FreezeError, OSError, ValueError) as error:
        print(f"0948 freeze failed: {error}")
        raise SystemExit(1)
