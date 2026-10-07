#!/usr/bin/env python3
"""Archive and verify the 0945 performance evidence root.

The archive, extraction, and cleanup actions are deliberately gated twice:
the caller must pass ``--go`` and set ``LITCHI_PERF_0945_ARCHIVE_GO=1``.
``--plan`` is read-only.  The selector excludes only the explicitly named
build and quality target roots below.  It does not filter arbitrary path
components beginning with ``target-``; those names can occur in source and
fuzz corpora and must remain in the custody inventory.

The external binary descriptor is intentionally left pending until all final
native and allocator builds produce their bound release hashes.  Archive and
verify modes refuse to run until all five expected binary entries are bound.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from datetime import datetime, timezone
from pathlib import Path
import shlex
import shutil
import subprocess
import sys
from typing import Any, Iterable, Iterator


SOURCE_ROOT = Path("/home/zhuhe/litchi-performance-0945").resolve()
WORK_ROOT = Path("/home/zhuhe/litchi-performance-0945-archive").resolve()
SCRIPT = Path(__file__).resolve()
ARCHIVE = WORK_ROOT / "litchi-performance-0945-v1.tar.zst"
MANIFEST = WORK_ROOT / "litchi-performance-0945-v1.files.jsonl"
MEMBERS = WORK_ROOT / "litchi-performance-0945-v1.members.list"
ARCHIVE_RECEIPT = WORK_ROOT / "litchi-performance-0945-v1.archive-receipt.json"
VERIFY_RECEIPT = WORK_ROOT / "litchi-performance-0945-v1.verify-receipt.json"
CLEANUP_RECEIPT = WORK_ROOT / "litchi-performance-0945-v1.cleanup-receipt.json"
EXTRACT_ROOT = WORK_ROOT / "litchi-performance-0945-v1.extract"
BINARY_DESCRIPTOR = WORK_ROOT / "binary-descriptor.json"
ARCHIVE_STDOUT = WORK_ROOT / "litchi-performance-0945-v1.archive.stdout.log"
ARCHIVE_STDERR = WORK_ROOT / "litchi-performance-0945-v1.archive.stderr.log"
VERIFY_STDOUT = WORK_ROOT / "litchi-performance-0945-v1.verify.stdout.log"
VERIFY_STDERR = WORK_ROOT / "litchi-performance-0945-v1.verify.stderr.log"
TAR = Path("/usr/bin/tar").resolve()
GO_ENV = "LITCHI_PERF_0945_ARCHIVE_GO"
GO_VALUE = "1"
ROOT_NAME = SOURCE_ROOT.name


# These are exact root-relative cache directories.  No other directory is
# treated as a target cache, even if one of its components starts with
# ``target-``.
BUILD_TARGET_CACHE_DIRS = frozenset({
    Path("target-baseline-allocator-v1"),
    Path("target-candidate-native-v2"),
    Path("target-candidate-allocator-v2"),
})
QUALITY_TARGET_DIRS = frozenset({
    Path("quality/target-baseline"),
    Path("quality/target-candidate"),
    Path("quality-v2/target-baseline"),
    Path("quality-v2/target-candidate"),
})
EXCLUDED_DIRECTORIES = BUILD_TARGET_CACHE_DIRS | QUALITY_TARGET_DIRS

ALLOWLISTED_TARGET_FILES = frozenset({
    Path("target-baseline-allocator-v1/release/litchi-perf-baseline-alloc"),
    Path("target-candidate-native-v2/release/litchi-perf-baseline"),
    Path("target-candidate-allocator-v2/release/litchi-perf-baseline-alloc"),
})

# The final build owner writes this descriptor in the external archive
# workspace after all five bound binaries exist.  Archive and verify modes
# refuse to run until it has exact positive sizes and SHA-256 values for every
# required path.  The descriptor itself is hash-bound in both receipts.
BINARY_DESCRIPTOR_SCHEMA = "litchi.performance.0945.binary-descriptor.v1"
BOUND_BINARY_PATHS = frozenset({
    *ALLOWLISTED_TARGET_FILES,
    Path("binaries/baseline-native"),
    Path("binaries/baseline-allocator"),
})

# These files are frozen custody controls already produced by the 0945
# preparation/quality lanes.  They are included in the source inventory and
# are checked by exact size and SHA-256 before and after selection.
FROZEN_CONTROL_FILES: dict[Path, dict[str, Any]] = {
    Path("baseline-source-manifest.json"): {
        "bytes": 2_993_755,
        "sha256": "c276582f412ac9c670a8b0fc3cc85ef12dc8a8705c201dd78e5584efa3a5a100",
    },
    Path("candidate-source-manifest-v1.json"): {
        "bytes": 3_048_238,
        "sha256": "39e813810818fdfd45f8c7862f22a71c3c9e89b16c84588ea9affa0a0199b353",
    },
    Path("candidate-source-manifest-v2.json"): {
        "bytes": 3_048_241,
        "sha256": "da7128f699b4bf7a9a4745e99b99606830df1aa192ce389a0b7fe3cceb8a0994",
    },
    Path("quality/source-freeze-v1.json"): {
        "bytes": 5_107_822,
        "sha256": "3664eb3f28b8e9587b05de5ef638a41589f0b85f638965e4663856a2dae426fe",
    },
    Path("quality/boundary-inputs-v1.json"): {
        "bytes": 2_947,
        "sha256": "754acc95258300180b2ee212f6653e405868cb9e0f1cf575a9531d18654bc798",
    },
    Path("quality/candidate-binding-v1.json"): {
        "bytes": 4_342,
        "sha256": "81a0e11346f6fc9b75d0570cfc3b81afcebdaa57f86bd40c062b3d5181f5cbc6",
    },
    Path("quality/candidate-binding-v2.json"): {
        "bytes": 5_025,
        "sha256": "b46bc421fc057631565d678988a96b5f1b0184795c374a8c5afc7c40700921e7",
    },
    Path("quality/candidate-binding-v3.json"): {
        "bytes": 6_344,
        "sha256": "8738ce7405438e9b0137f8cc7ec3deef86e1325d5852b2b965105eb69ed59af4",
    },
    Path("quality-v2/source-freeze-v2.json"): {
        "bytes": 10_463,
        "sha256": "eb87c42b6811c497f1e4de7453d75ecee05ae8ac19ae1386b6cdba2bd82bc331",
    },
    Path("quality-v2/quality-inputs-v2.json"): {
        "bytes": 8_912,
        "sha256": "49abd99e8f46de5529448f0d10a58a1d085abf6c4e2ddb0c4f219bb2aef54272",
    },
    Path("quality-v2/candidate-binding-v4.json"): {
        "bytes": 2_300,
        "sha256": "f223080b29e3ba2c59bdcb5a366e87436d077a536f0c048ea4da85c34df4f90e",
    },
    Path("receipts/candidate-freeze-v1.json"): {
        "bytes": 901,
        "sha256": "94f2e75ad659214ed5310017bd0670df0f11d821c54263b1e6f0f8397fea2af3",
    },
    Path("receipts/candidate-freeze-v2.json"): {
        "bytes": 1_269,
        "sha256": "f624f161ad7e04499f6ae3a5f5079533ecc489c3690a2f46fdf947f3492dce32",
    },
}

SOURCE_MANIFESTS = (
    (Path("baseline-source-manifest.json"), Path("baseline-source")),
    (Path("candidate-source-manifest-v1.json"), Path("candidate-source")),
    (Path("candidate-source-manifest-v2.json"), Path("candidate-source-v2")),
)
QUALITY_SOURCE_MANIFEST_COPIES = (
    (Path("baseline-source-manifest.json"), Path("quality/baseline-source")),
    (Path("candidate-source-manifest-v1.json"), Path("quality/candidate-source")),
    (Path("baseline-source-manifest.json"), Path("quality-v2/baseline-source")),
    (Path("candidate-source-manifest-v2.json"), Path("quality-v2/candidate-source")),
)
QUALITY_INPUT_REGISTRY = Path("quality-v2/quality-inputs-v2.json")
QUALITY_FREEZE = Path("quality-v2/source-freeze-v2.json")
QUALITY_FREEZE_V1 = Path("quality/source-freeze-v1.json")
BOUNDARY_INPUTS = Path("quality/boundary-inputs-v1.json")
QUALITY_BINDING_V3 = Path("quality/candidate-binding-v3.json")
TAR_ENVIRONMENT = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
}


class ArchiveError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ArchiveError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def sha256_path(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_descriptor(path: Path) -> dict[str, Any]:
    path = Path(path)
    require(not path.is_symlink(), f"symlink is not a custody file: {path}")
    require(path.is_file(), f"expected regular file: {path}")
    path = path.resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_path(path)}


def write_json_exclusive(path: Path, value: Any) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise ArchiveError(f"refusing to replace existing artifact: {path}") from error


def write_bytes_exclusive(path: Path, data: bytes) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise ArchiveError(f"refusing to replace existing artifact: {path}") from error


def relative_to_source(path: Path) -> Path:
    path = path.resolve()
    try:
        return path.relative_to(SOURCE_ROOT)
    except ValueError as error:
        raise ArchiveError(f"path escaped source root: {path}") from error


def is_under(relative: Path, directory: Path) -> bool:
    return relative == directory or directory in relative.parents


def is_excluded(relative: Path) -> bool:
    return any(is_under(relative, directory) for directory in EXCLUDED_DIRECTORIES)


def has_allowlisted_descendant(relative: Path) -> bool:
    return any(
        allowed.parts[: len(relative.parts)] == relative.parts
        for allowed in ALLOWLISTED_TARGET_FILES
    )


def may_descend(relative_directory: Path) -> bool:
    if not is_excluded(relative_directory):
        return True
    return has_allowlisted_descendant(relative_directory)


def selected_file(relative: Path) -> bool:
    if not is_excluded(relative):
        return True
    return relative in ALLOWLISTED_TARGET_FILES


def iter_selected_files() -> Iterator[tuple[Path, Path]]:
    require(SOURCE_ROOT.is_dir() and not SOURCE_ROOT.is_symlink(), f"missing source root: {SOURCE_ROOT}")
    for current, directories, filenames in os.walk(SOURCE_ROOT, topdown=True, followlinks=False):
        current_path = Path(current).resolve()
        current_relative = relative_to_source(current_path)
        kept_directories: list[str] = []
        for name in sorted(directories):
            child = current_path / name
            child_relative = current_relative / name
            if child.is_symlink():
                raise ArchiveError(f"symlink encountered under source root: {child}")
            if may_descend(child_relative):
                kept_directories.append(name)
        directories[:] = kept_directories
        for name in sorted(filenames):
            path = current_path / name
            relative = current_relative / name
            if not selected_file(relative):
                continue
            require(not path.is_symlink(), f"symlink encountered under source root: {path}")
            require(path.is_file(), f"non-regular selected artifact: {path}")
            yield relative, path


def expected_binary_descriptors(*, allow_pending: bool = False) -> dict[str, Any]:
    if not BINARY_DESCRIPTOR.exists() or BINARY_DESCRIPTOR.is_symlink():
        require(allow_pending, f"bound binary descriptor is pending: {BINARY_DESCRIPTOR}")
        return {
            "descriptor": {"path": str(BINARY_DESCRIPTOR), "present": False},
            "required_paths": sorted(path.as_posix() for path in BOUND_BINARY_PATHS),
        }
    data = load_json(BINARY_DESCRIPTOR)
    require(data.get("schema") == BINARY_DESCRIPTOR_SCHEMA, "binary descriptor schema changed")
    require(data.get("source_root") == str(SOURCE_ROOT), "binary descriptor source root changed")
    entries = data.get("binaries")
    require(isinstance(entries, list) and entries, "binary descriptor has no binaries")
    actual: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(entries):
        require(isinstance(item, dict), f"binary descriptor entry {index} is not an object")
        relative = Path(validate_relative_value(item.get("path"), label=f"binary descriptor entry {index}"))
        require(relative in BOUND_BINARY_PATHS, f"binary descriptor path is not bound: {relative}")
        require(relative.as_posix() not in actual, f"duplicate bound binary: {relative}")
        expected_bytes = item.get("bytes")
        expected_sha = item.get("sha256")
        require(type(expected_bytes) is int and expected_bytes > 0,
                f"invalid bound binary size: {relative}")
        require(isinstance(expected_sha, str) and len(expected_sha) == 64 and
                all(c in "0123456789abcdef" for c in expected_sha),
                f"invalid bound binary SHA: {relative}")
        path = SOURCE_ROOT / relative
        observed = file_descriptor(path)
        require(observed["bytes"] == expected_bytes, f"bound binary size changed: {relative}")
        require(observed["sha256"] == expected_sha, f"bound binary SHA changed: {relative}")
        actual[relative.as_posix()] = observed
    require(set(actual) == {path.as_posix() for path in BOUND_BINARY_PATHS},
            "binary descriptor does not bind exactly the three target and two inventory binaries")
    return {
        "descriptor": file_descriptor(BINARY_DESCRIPTOR),
        "binaries": actual,
        "build_receipts": data.get("build_receipts"),
    }


def verify_frozen_control_files() -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for relative, expected in sorted(FROZEN_CONTROL_FILES.items(), key=lambda item: item[0].as_posix()):
        actual = file_descriptor(SOURCE_ROOT / relative)
        require(actual["bytes"] == expected["bytes"], f"frozen control size changed: {relative}")
        require(actual["sha256"] == expected["sha256"], f"frozen control hash changed: {relative}")
        result[relative.as_posix()] = actual
    return result


def load_json(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing JSON control: {path}")
    with path.open("r", encoding="utf-8") as stream:
        value = json.load(stream)
    require(isinstance(value, dict), f"JSON control is not an object: {path}")
    return value


def validate_relative_value(value: Any, *, label: str) -> str:
    require(isinstance(value, str) and value, f"{label} is missing")
    relative = Path(value)
    require(not relative.is_absolute() and ".." not in relative.parts,
            f"{label} escapes its root: {value}")
    return relative.as_posix()


def verify_source_manifests(record_map: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    checked: dict[str, dict[str, Any]] = {}
    manifest_copies = SOURCE_MANIFESTS + QUALITY_SOURCE_MANIFEST_COPIES
    for manifest_relative, source_relative in manifest_copies:
        manifest_path = SOURCE_ROOT / manifest_relative
        data = load_json(manifest_path)
        rows = data.get("files")
        require(isinstance(rows, list), f"source manifest files is not a list: {manifest_relative}")
        require(data.get("file_count") == len(rows), f"source manifest count changed: {manifest_relative}")
        require(data.get("bytes") == sum(row.get("bytes", -1) for row in rows),
                f"source manifest byte total changed: {manifest_relative}")
        previous = ""
        for row in rows:
            require(isinstance(row, dict), f"source manifest row is not an object: {manifest_relative}")
            path_value = validate_relative_value(row.get("path"), label=f"{manifest_relative} row path")
            require(path_value > previous, f"source manifest rows are not sorted: {manifest_relative}")
            previous = path_value
            expected_bytes = row.get("bytes")
            expected_sha = row.get("sha256")
            require(type(expected_bytes) is int and expected_bytes >= 0,
                    f"source manifest row size invalid: {manifest_relative}:{path_value}")
            require(isinstance(expected_sha, str) and len(expected_sha) == 64,
                    f"source manifest row hash invalid: {manifest_relative}:{path_value}")
            archive_path = (source_relative / path_value).as_posix()
            actual = record_map.get(archive_path)
            require(actual is not None, f"source manifest path missing from selection: {archive_path}")
            require(actual["bytes"] == expected_bytes and actual["sha256"] == expected_sha,
                    f"source manifest row differs from selected source: {archive_path}")
        checked[f"{manifest_relative.as_posix()}::{source_relative.as_posix()}"] = {
            "source_root": source_relative.as_posix(),
            "file_count": len(rows),
            "bytes": data["bytes"],
        }
    return checked


def load_quality_input_registry() -> list[dict[str, Any]]:
    registry_path = SOURCE_ROOT / QUALITY_INPUT_REGISTRY
    data = load_json(registry_path)
    require(data.get("status") == "pass", "quality input registry is not passing")
    inputs = data.get("inputs")
    require(isinstance(inputs, list), "quality input registry inputs is not a list")
    require(all(isinstance(item, dict) for item in inputs), "quality input registry contains a non-object")
    return inputs


def verify_standard_quality_inputs(
    inputs: list[dict[str, Any]],
    record_map: dict[str, dict[str, Any]] | None,
    *,
    label: str,
    quality_root: Path,
) -> dict[str, Any]:
    checked: list[dict[str, Any]] = []
    for index, item in enumerate(inputs):
        require(isinstance(item, dict), f"{label} input {index} is not an object")
        origin = item.get("origin")
        require(isinstance(origin, dict), f"{label} input {index} has no origin")
        origin_path = Path(origin.get("path", ""))
        actual_origin = file_descriptor(origin_path)
        require(actual_origin["bytes"] == origin.get("bytes") and actual_origin["sha256"] == origin.get("sha256"),
                f"{label} origin changed: {origin_path}")
        copies = item.get("copies")
        require(isinstance(copies, list) and copies, f"{label} input {index} has no copies")
        copy_records: list[dict[str, Any]] = []
        for copy in copies:
            require(isinstance(copy, dict), f"{label} input {index} copy is not an object")
            role = copy.get("role")
            require(role in {"baseline", "candidate"}, f"unknown {label} role: {role}")
            destination = copy.get("destination")
            require(isinstance(destination, dict), f"{label} input {index} destination is missing")
            relative_value = validate_relative_value(destination.get("path"), label=f"{label} destination")
            archive_path = (quality_root / f"{role}-source" / relative_value).as_posix()
            direct_path = SOURCE_ROOT / archive_path
            direct = file_descriptor(direct_path)
            require(direct["bytes"] == destination.get("bytes") and direct["sha256"] == destination.get("sha256"),
                    f"{label} copy changed: {archive_path}")
            if record_map is not None:
                actual = record_map.get(archive_path)
                require(actual is not None, f"{label} copy missing from selection: {archive_path}")
                require(actual["bytes"] == destination.get("bytes") and actual["sha256"] == destination.get("sha256"),
                        f"{label} selected copy changed: {archive_path}")
            copy_records.append({
                "role": role,
                "path": archive_path,
                "bytes": direct["bytes"],
                "sha256": direct["sha256"],
            })
        checked.append({
            "kind": item.get("kind"),
            "relative_destination": item.get("relative_destination"),
            "origin": actual_origin,
            "copies": copy_records,
        })
    return {"registry": label, "inputs": checked}


def verify_boundary_inputs(record_map: dict[str, dict[str, Any]] | None) -> dict[str, Any]:
    data = load_json(SOURCE_ROOT / BOUNDARY_INPUTS)
    inputs = data.get("inputs")
    require(isinstance(inputs, list), "boundary input registry inputs is not a list")
    checked: list[dict[str, Any]] = []
    for index, item in enumerate(inputs):
        require(isinstance(item, dict), f"boundary input {index} is not an object")
        checkout = item.get("checkout")
        require(isinstance(checkout, dict), f"boundary input {index} checkout is missing")
        checkout_path = Path(checkout.get("path", ""))
        actual_checkout = file_descriptor(checkout_path)
        require(actual_checkout["bytes"] == checkout.get("bytes") and actual_checkout["sha256"] == checkout.get("sha256"),
                f"boundary checkout input changed: {checkout_path}")
        copies: list[dict[str, Any]] = []
        for quality_copy in item.get("quality_copies", []):
            require(isinstance(quality_copy, dict), f"boundary input {index} copy is not an object")
            role = quality_copy.get("role")
            copy = quality_copy.get("copy")
            require(role in {"baseline", "candidate"} and isinstance(copy, dict),
                    f"boundary input {index} copy is malformed")
            copy_path = Path(copy.get("path", ""))
            actual_copy = file_descriptor(copy_path)
            require(actual_copy["bytes"] == copy.get("bytes") and actual_copy["sha256"] == copy.get("sha256"),
                    f"boundary quality copy changed: {copy_path}")
            relative = relative_to_source(copy_path).as_posix()
            if record_map is not None:
                actual_selected = record_map.get(relative)
                require(actual_selected is not None, f"boundary quality copy missing from selection: {relative}")
                require(actual_selected["bytes"] == copy.get("bytes") and actual_selected["sha256"] == copy.get("sha256"),
                        f"boundary selected copy changed: {relative}")
            copies.append({"role": role, "path": relative, "bytes": actual_copy["bytes"], "sha256": actual_copy["sha256"]})
        checked.append({"relative_path": item.get("relative_path"), "checkout": actual_checkout, "copies": copies})
    return {"registry": BOUNDARY_INPUTS.as_posix(), "inputs": checked}


def verify_candidate_binding_quality_inputs(record_map: dict[str, dict[str, Any]] | None) -> dict[str, Any]:
    data = load_json(SOURCE_ROOT / QUALITY_BINDING_V3)
    verification = data.get("copied_source_verification")
    require(isinstance(verification, dict), "candidate binding quality verification is missing")
    declared = verification.get("quality_only_inputs")
    require(isinstance(declared, dict), "candidate binding quality inputs are not a map")
    checked: list[dict[str, Any]] = []
    for relative_value, descriptor in sorted(declared.items()):
        relative = validate_relative_value(relative_value, label="candidate binding quality input")
        require(isinstance(descriptor, dict), f"candidate binding descriptor is malformed: {relative}")
        path = Path(descriptor.get("path", ""))
        actual = file_descriptor(path)
        require(actual["bytes"] == descriptor.get("bytes") and actual["sha256"] == descriptor.get("sha256"),
                f"candidate binding quality input changed: {path}")
        selected_relative = relative_to_source(path).as_posix()
        require(selected_relative == f"quality/candidate-source/{relative}",
                f"candidate binding quality input path changed: {path}")
        if record_map is not None:
            selected = record_map.get(selected_relative)
            require(selected is not None, f"candidate binding quality input missing from selection: {selected_relative}")
            require(selected["bytes"] == descriptor.get("bytes") and selected["sha256"] == descriptor.get("sha256"),
                    f"candidate binding selected quality input changed: {selected_relative}")
        checked.append({"path": selected_relative, "bytes": actual["bytes"], "sha256": actual["sha256"]})
    return {"registry": QUALITY_BINDING_V3.as_posix(), "inputs": checked}


def verify_quality_freeze_v1_manifest(record_map: dict[str, dict[str, Any]] | None) -> dict[str, Any]:
    data = load_json(SOURCE_ROOT / QUALITY_FREEZE_V1)
    roles = data.get("roles")
    require(isinstance(roles, dict), "quality v1 freeze roles are missing")
    checked: dict[str, Any] = {}
    for role in ("baseline", "candidate"):
        value = roles.get(role)
        require(isinstance(value, dict), f"quality v1 freeze role is missing: {role}")
        quality_root = relative_to_source(Path(value.get("quality_root", "")))
        expected_root = Path("quality") / f"{role}-source"
        require(quality_root == expected_root, f"quality v1 freeze root changed: {quality_root}")
        rows = value.get("manifest")
        require(isinstance(rows, list), f"quality v1 freeze manifest is not a list: {role}")
        require(value.get("file_count") == len(rows), f"quality v1 freeze count changed: {role}")
        require(value.get("total_bytes") == sum(row.get("bytes", -1) for row in rows),
                f"quality v1 freeze byte total changed: {role}")
        previous = ""
        for row in rows:
            require(isinstance(row, dict), f"quality v1 freeze row is not an object: {role}")
            relative = validate_relative_value(row.get("path"), label=f"quality v1 freeze {role} path")
            require(relative > previous, f"quality v1 freeze rows are not sorted: {role}")
            previous = relative
            expected_bytes = row.get("bytes")
            expected_sha = row.get("sha256")
            require(type(expected_bytes) is int and expected_bytes >= 0,
                    f"quality v1 freeze row size invalid: {role}:{relative}")
            require(isinstance(expected_sha, str) and len(expected_sha) == 64,
                    f"quality v1 freeze row hash invalid: {role}:{relative}")
            archive_path = (quality_root / relative).as_posix()
            direct = file_descriptor(SOURCE_ROOT / archive_path)
            require(direct["bytes"] == expected_bytes and direct["sha256"] == expected_sha,
                    f"quality v1 freeze copy changed: {archive_path}")
            if record_map is not None:
                actual = record_map.get(archive_path)
                require(actual is not None, f"quality v1 freeze path missing from selection: {archive_path}")
                require(actual["bytes"] == expected_bytes and actual["sha256"] == expected_sha,
                        f"quality v1 freeze selected copy changed: {archive_path}")
        checked[role] = {
            "root": quality_root.as_posix(),
            "file_count": len(rows),
            "bytes": value["total_bytes"],
        }
    return {"registry": QUALITY_FREEZE_V1.as_posix(), "roles": checked}


def verify_quality_input_origins() -> dict[str, Any]:
    registry = verify_standard_quality_inputs(
        load_quality_input_registry(), None, label=QUALITY_INPUT_REGISTRY.as_posix(), quality_root=Path("quality-v2"))
    freeze = load_json(SOURCE_ROOT / QUALITY_FREEZE)
    freeze_inputs = freeze.get("quality_only_inputs")
    require(isinstance(freeze_inputs, list), "quality freeze inputs is not a list")
    freeze_checked = verify_standard_quality_inputs(
        freeze_inputs, None, label=QUALITY_FREEZE.as_posix(), quality_root=Path("quality-v2"))
    quality_v1 = verify_quality_freeze_v1_manifest(None)
    boundary = verify_boundary_inputs(None)
    binding = verify_candidate_binding_quality_inputs(None)
    return {
        "registry": QUALITY_INPUT_REGISTRY.as_posix(),
        "origins": registry["inputs"],
        "freeze": freeze_checked,
        "quality_v1_freeze": quality_v1,
        "boundary": boundary,
        "candidate_binding": binding,
    }


def verify_quality_inputs(record_map: dict[str, dict[str, Any]]) -> dict[str, Any]:
    registry = verify_standard_quality_inputs(
        load_quality_input_registry(), record_map,
        label=QUALITY_INPUT_REGISTRY.as_posix(), quality_root=Path("quality-v2"))
    freeze = load_json(SOURCE_ROOT / QUALITY_FREEZE)
    freeze_inputs = freeze.get("quality_only_inputs")
    require(isinstance(freeze_inputs, list), "quality freeze inputs is not a list")
    freeze_checked = verify_standard_quality_inputs(
        freeze_inputs, record_map, label=QUALITY_FREEZE.as_posix(), quality_root=Path("quality-v2"))
    quality_v1 = verify_quality_freeze_v1_manifest(record_map)
    boundary = verify_boundary_inputs(record_map)
    binding = verify_candidate_binding_quality_inputs(record_map)
    return {
        "registry": registry,
        "freeze": freeze_checked,
        "quality_v1_freeze": quality_v1,
        "boundary": boundary,
        "candidate_binding": binding,
    }


def verify_selected_custody(records: list[dict[str, Any]]) -> dict[str, Any]:
    record_map = {record["path"]: record for record in records}
    require(len(record_map) == len(records), "duplicate selected archive path")
    source_manifests = verify_source_manifests(record_map)
    quality_inputs = verify_quality_inputs(record_map)
    return {"source_manifests": source_manifests, "quality_inputs": quality_inputs}


def collect_records() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    verify_frozen_control_files()
    bound_binaries = expected_binary_descriptors()
    records: list[dict[str, Any]] = []
    seen: set[str] = set()
    for relative, path in iter_selected_files():
        key = relative.as_posix()
        require(key not in seen, f"duplicate selected path: {key}")
        seen.add(key)
        records.append({"path": key, "bytes": path.stat().st_size, "sha256": sha256_path(path)})
    records.sort(key=lambda item: item["path"])
    verify_selected_custody(records)
    header = {
        "schema": "litchi.performance.0945.archive-file-manifest.v1",
        "source_root": str(SOURCE_ROOT),
        "root_name": ROOT_NAME,
        "generated_utc": utc_now(),
        "files": len(records),
        "bytes": sum(item["bytes"] for item in records),
        "excluded": {
            "exact_directories": sorted(path.as_posix() for path in EXCLUDED_DIRECTORIES),
            "target_cache_policy": "exclude only the seven named directories; preserve every other target-* path",
            "target_binary_allowlist": sorted(path.as_posix() for path in ALLOWLISTED_TARGET_FILES),
        },
        "custody": {
            "frozen_control_files": sorted(path.as_posix() for path in FROZEN_CONTROL_FILES),
            "source_manifests": sorted({path.as_posix() for path, _root in SOURCE_MANIFESTS}),
            "quality_source_manifest_copies": sorted(
                f"{path.as_posix()}::{root.as_posix()}"
                for path, root in QUALITY_SOURCE_MANIFEST_COPIES
            ),
            "quality_input_registry": QUALITY_INPUT_REGISTRY.as_posix(),
        },
        "bound_binaries": bound_binaries,
    }
    return header, records


def write_manifest(header: dict[str, Any], records: list[dict[str, Any]]) -> None:
    MANIFEST.parent.mkdir(parents=True, exist_ok=True)
    try:
        with MANIFEST.open("x", encoding="utf-8") as stream:
            stream.write(json.dumps(header, sort_keys=True) + "\n")
            for record in records:
                stream.write(json.dumps(record, sort_keys=True) + "\n")
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise ArchiveError(f"refusing to replace existing artifact: {MANIFEST}") from error


def write_member_list(records: list[dict[str, Any]]) -> None:
    encoded = b"".join((ROOT_NAME + "/" + record["path"]).encode("utf-8") + b"\0" for record in records)
    write_bytes_exclusive(MEMBERS, encoded)


def validate_manifest_record(value: Any) -> str:
    relative = validate_relative_value(value, label="archive manifest path")
    path = Path(relative)
    require(not is_excluded(path) or path in ALLOWLISTED_TARGET_FILES,
            f"archive manifest includes excluded path: {relative}")
    return relative


def read_manifest() -> tuple[dict[str, Any], list[dict[str, Any]]]:
    require(MANIFEST.is_file() and not MANIFEST.is_symlink(), f"missing archive manifest: {MANIFEST}")
    records: list[dict[str, Any]] = []
    with MANIFEST.open("r", encoding="utf-8") as stream:
        first = stream.readline()
        require(first, "archive manifest is empty")
        header = json.loads(first)
        require(isinstance(header, dict), "archive manifest header is not an object")
        require(header.get("schema") == "litchi.performance.0945.archive-file-manifest.v1",
                "archive manifest schema changed")
        require(header.get("source_root") == str(SOURCE_ROOT), "archive manifest source root changed")
        require(header.get("root_name") == ROOT_NAME, "archive manifest root name changed")
        previous = ""
        for line_number, line in enumerate(stream, start=2):
            record = json.loads(line)
            require(isinstance(record, dict), f"archive manifest line {line_number} is not an object")
            relative = validate_manifest_record(record.get("path"))
            require(relative > previous, f"archive manifest paths are not strictly sorted at line {line_number}")
            previous = relative
            require(type(record.get("bytes")) is int and record["bytes"] >= 0,
                    f"archive manifest size invalid at line {line_number}")
            digest = record.get("sha256")
            require(isinstance(digest, str) and len(digest) == 64 and all(c in "0123456789abcdef" for c in digest),
                    f"archive manifest digest invalid at line {line_number}")
            records.append({"path": relative, "bytes": record["bytes"], "sha256": digest})
    require(header.get("files") == len(records), "archive manifest file count mismatch")
    require(header.get("bytes") == sum(item["bytes"] for item in records),
            "archive manifest byte count mismatch")
    require(header.get("excluded", {}).get("exact_directories") ==
            sorted(path.as_posix() for path in EXCLUDED_DIRECTORIES),
            "archive manifest exclusions changed")
    return header, records


def tar_archive_command() -> list[str]:
    return [
        str(TAR), "--create", "--zstd", "--format=posix", "--numeric-owner", "--owner=0", "--group=0",
        "--mtime=@0", "--no-recursion", "--null", "--verbatim-files-from",
        "--directory", str(SOURCE_ROOT.parent), "--files-from", str(MEMBERS), "--file", str(ARCHIVE),
    ]


def tar_extract_command() -> list[str]:
    return [
        str(TAR), "--zstd", "--extract", "--file", str(ARCHIVE), "--directory", str(EXTRACT_ROOT),
        "--no-same-owner", "--no-same-permissions", "--no-overwrite-dir",
    ]


def run_logged(command: list[str], stdout_path: Path, stderr_path: Path, *, cwd: Path) -> int:
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            completed = subprocess.run(
                command, cwd=str(cwd), env=dict(TAR_ENVIRONMENT), stdin=subprocess.DEVNULL,
                stdout=stdout, stderr=stderr, check=False,
            )
    except FileExistsError as error:
        raise ArchiveError(f"refusing to replace command log: {error}") from error
    return completed.returncode


def script_command(*actions: str) -> str:
    return f"{GO_ENV}=1 /usr/bin/python3 {shlex.quote(str(SCRIPT))} --go " + " ".join(actions)


_GO_FLAG = False


def require_go() -> None:
    require(os.environ.get(GO_ENV) == GO_VALUE and _GO_FLAG,
            f"archive/verify/cleanup is gated; require --go and {GO_ENV}=1")


def require_absent(paths: Iterable[Path]) -> None:
    for path in paths:
        require(not path.exists() and not path.is_symlink(), f"refusing to replace existing artifact: {path}")


def custody_snapshot() -> dict[str, Any]:
    return {
        "controls": verify_frozen_control_files(),
        "bound_binaries": expected_binary_descriptors(),
        "quality_input_origins": verify_quality_input_origins(),
    }


def current_matches_archive(records: list[dict[str, Any]]) -> None:
    _header, current = collect_records()
    require(current == records, "selected evidence changed after archive selection")


def archive_mode() -> dict[str, Any]:
    require_go()
    require(TAR.is_file() and os.access(TAR, os.X_OK), f"tar with zstd support is unavailable: {TAR}")
    require_absent((ARCHIVE, MANIFEST, MEMBERS, ARCHIVE_RECEIPT, VERIFY_RECEIPT, CLEANUP_RECEIPT,
                    EXTRACT_ROOT, ARCHIVE_STDOUT, ARCHIVE_STDERR, VERIFY_STDOUT, VERIFY_STDERR))
    custody_before = custody_snapshot()
    header, records = collect_records()
    write_manifest(header, records)
    write_member_list(records)
    command = tar_archive_command()
    return_code = run_logged(command, ARCHIVE_STDOUT, ARCHIVE_STDERR, cwd=SOURCE_ROOT.parent)
    require(return_code == 0, f"tar archive command failed with exit code {return_code}")
    require(ARCHIVE.is_file() and ARCHIVE.stat().st_size > 0, "tar archive was not produced")
    current_matches_archive(records)
    custody_after = custody_snapshot()
    require(custody_after == custody_before, "custody controls changed during archive creation")
    receipt = {
        "schema": "litchi.performance.0945.archive-receipt.v1",
        "status": "archived",
        "created_utc": utc_now(),
        "script": file_descriptor(SCRIPT),
        "source_root": str(SOURCE_ROOT),
        "manifest": file_descriptor(MANIFEST),
        "member_list": file_descriptor(MEMBERS),
        "archive": file_descriptor(ARCHIVE),
        "archive_stdout": file_descriptor(ARCHIVE_STDOUT),
        "archive_stderr": file_descriptor(ARCHIVE_STDERR),
        "files": header["files"],
        "bytes": header["bytes"],
        "custody_before": custody_before,
        "custody_after": custody_after,
        "bound_binaries": header["bound_binaries"],
        "target_binary_allowlist": sorted(path.as_posix() for path in ALLOWLISTED_TARGET_FILES),
        "tar_command": command,
        "tar_command_shell": shlex.join(command),
        "regeneration_commands": {
            "archive": script_command("--archive"),
            "verify": script_command("--verify"),
            "cleanup_after_verification": script_command("--cleanup"),
        },
        "excluded_rules": header["excluded"],
        "cleanup_policy": {
            "preserve_bound_release_binaries": True,
            "remove_only_named_excluded_directories": True,
            "remove_verification_extraction": str(EXTRACT_ROOT),
        },
    }
    write_json_exclusive(ARCHIVE_RECEIPT, receipt)
    return {"status": receipt["status"], "files": header["files"], "bytes": header["bytes"],
            "archive": receipt["archive"], "archive_receipt": str(ARCHIVE_RECEIPT)}


def extracted_file_records() -> list[dict[str, Any]]:
    root = EXTRACT_ROOT / ROOT_NAME
    require(not root.is_symlink() and root.is_dir(), f"archive extraction root missing: {root}")
    actual: list[dict[str, Any]] = []
    for current, directories, filenames in os.walk(root, topdown=True, followlinks=False):
        current_path = Path(current)
        for name in directories:
            require(not (current_path / name).is_symlink(), f"symlink in archive extraction: {current_path / name}")
        for name in sorted(filenames):
            path = current_path / name
            require(not path.is_symlink() and path.is_file(), f"non-regular extraction member: {path}")
            relative = path.relative_to(root).as_posix()
            actual.append({"path": ROOT_NAME + "/" + relative, "bytes": path.stat().st_size,
                           "sha256": sha256_path(path)})
    actual.sort(key=lambda item: item["path"])
    return actual


def load_receipt(path: Path, schema: str, status: str) -> dict[str, Any]:
    value = load_json(path)
    require(value.get("schema") == schema and value.get("status") == status,
            f"receipt is not {status}: {path}")
    return value


def verify_mode() -> dict[str, Any]:
    require_go()
    archive_receipt = load_receipt(ARCHIVE_RECEIPT, "litchi.performance.0945.archive-receipt.v1", "archived")
    header, records = read_manifest()
    require(custody_snapshot() == archive_receipt.get("custody_before"),
            "current custody controls differ from archive receipt")
    require(header.get("bound_binaries") == archive_receipt.get("bound_binaries"),
            "manifest binary custody differs from archive receipt")
    verify_selected_custody(records)
    require(file_descriptor(ARCHIVE) == archive_receipt["archive"], "archive changed since archive receipt")
    require(file_descriptor(MANIFEST) == archive_receipt["manifest"], "manifest changed since archive receipt")
    current_matches_archive(records)
    require_absent((EXTRACT_ROOT, VERIFY_RECEIPT, VERIFY_STDOUT, VERIFY_STDERR))
    EXTRACT_ROOT.mkdir(parents=True, exist_ok=False)
    command = tar_extract_command()
    return_code = run_logged(command, VERIFY_STDOUT, VERIFY_STDERR, cwd=SOURCE_ROOT.parent)
    require(return_code == 0, f"tar extraction command failed with exit code {return_code}")
    actual = extracted_file_records()
    expected = [{"path": ROOT_NAME + "/" + item["path"], "bytes": item["bytes"], "sha256": item["sha256"]}
                for item in records]
    require(actual == expected, "archive extraction file set, size, or digest differs from manifest")
    require(len(actual) == header["files"], "archive extraction file count differs from manifest")
    require(sum(item["bytes"] for item in actual) == header["bytes"],
            "archive extraction byte total differs from manifest")
    receipt = {
        "schema": "litchi.performance.0945.archive-verify-receipt.v1",
        "status": "passed",
        "verified_utc": utc_now(),
        "script": file_descriptor(SCRIPT),
        "archive": file_descriptor(ARCHIVE),
        "archive_receipt": file_descriptor(ARCHIVE_RECEIPT),
        "manifest": file_descriptor(MANIFEST),
        "extraction_root": str(EXTRACT_ROOT),
        "files": header["files"],
        "bytes": header["bytes"],
        "tar_command": command,
        "tar_command_shell": shlex.join(command),
        "verify_stdout": file_descriptor(VERIFY_STDOUT),
        "verify_stderr": file_descriptor(VERIFY_STDERR),
        "every_file_rehashed": True,
        "file_set_exact": True,
        "source_rechecked_before_extraction": True,
        "cleanup_authorized_by_this_receipt": False,
    }
    write_json_exclusive(VERIFY_RECEIPT, receipt)
    return {"status": receipt["status"], "files": receipt["files"], "bytes": receipt["bytes"],
            "verify_receipt": str(VERIFY_RECEIPT)}


def remove_path(path: Path) -> None:
    require(SOURCE_ROOT in path.parents, f"refusing removal outside source root: {path}")
    if path.is_symlink() or path.is_file():
        path.unlink()
    elif path.is_dir():
        shutil.rmtree(path)
    else:
        raise ArchiveError(f"cannot remove unknown path type: {path}")


def cleanup_named_directories() -> list[str]:
    removed: list[str] = []
    for directory in sorted(EXCLUDED_DIRECTORIES, key=Path.as_posix, reverse=True):
        root = SOURCE_ROOT / directory
        if not root.exists() and not root.is_symlink():
            continue
        require(root.is_dir() and not root.is_symlink(), f"excluded path is not a directory: {root}")
        for current, directories, filenames in os.walk(root, topdown=False, followlinks=False):
            current_path = Path(current)
            for name in sorted(filenames):
                path = current_path / name
                relative = path.relative_to(SOURCE_ROOT)
                if relative in ALLOWLISTED_TARGET_FILES:
                    continue
                remove_path(path)
                removed.append(relative.as_posix())
            for name in sorted(directories):
                path = current_path / name
                relative = path.relative_to(SOURCE_ROOT)
                if has_allowlisted_descendant(relative):
                    continue
                if path.exists() or path.is_symlink():
                    remove_path(path)
                    removed.append(relative.as_posix())
        if root.exists() and not has_allowlisted_descendant(directory):
            remove_path(root)
            removed.append(directory.as_posix())
    return removed


def assert_post_cleanup() -> dict[str, dict[str, Any]]:
    preserved = expected_binary_descriptors()
    for directory in EXCLUDED_DIRECTORIES:
        root = SOURCE_ROOT / directory
        if not root.exists():
            continue
        require(has_allowlisted_descendant(directory), f"unallowlisted excluded directory remains: {root}")
        for current, directories, filenames in os.walk(root, topdown=True, followlinks=False):
            current_path = Path(current)
            for name in directories:
                require(not (current_path / name).is_symlink(), f"symlink remains after cleanup: {current_path / name}")
            for name in filenames:
                relative = (current_path / name).relative_to(SOURCE_ROOT)
                require(relative in ALLOWLISTED_TARGET_FILES, f"unallowlisted target file remains: {relative}")
    return preserved


def cleanup_mode() -> dict[str, Any]:
    require_go()
    archive_receipt = load_receipt(ARCHIVE_RECEIPT, "litchi.performance.0945.archive-receipt.v1", "archived")
    verify_receipt = load_receipt(VERIFY_RECEIPT, "litchi.performance.0945.archive-verify-receipt.v1", "passed")
    header, records = read_manifest()
    require(file_descriptor(ARCHIVE) == archive_receipt["archive"], "archive changed before cleanup")
    require(file_descriptor(MANIFEST) == archive_receipt["manifest"], "manifest changed before cleanup")
    require(header.get("bound_binaries") == archive_receipt.get("bound_binaries"),
            "manifest binary custody differs before cleanup")
    require(verify_receipt["archive"] == archive_receipt["archive"], "verify receipt archive binding changed")
    require(EXTRACT_ROOT.is_dir() and not EXTRACT_ROOT.is_symlink(),
            "verified extraction is missing; cleanup requires completed verification")
    current_matches_archive(records)
    removed = cleanup_named_directories()
    shutil.rmtree(EXTRACT_ROOT)
    preserved = assert_post_cleanup()
    receipt = {
        "schema": "litchi.performance.0945.archive-cleanup-receipt.v1",
        "status": "completed",
        "cleaned_utc": utc_now(),
        "script": file_descriptor(SCRIPT),
        "source_root": str(SOURCE_ROOT),
        "archive": archive_receipt["archive"],
        "manifest": archive_receipt["manifest"],
        "verify_receipt": file_descriptor(VERIFY_RECEIPT),
        "files": header["files"],
        "bytes": header["bytes"],
        "removed_named_excluded_paths": removed,
        "removed_verification_extraction": str(EXTRACT_ROOT),
        "preserved_bound_release_binaries": preserved,
        "checkout_touched": False,
        "other_performance_roots_touched": [],
    }
    write_json_exclusive(CLEANUP_RECEIPT, receipt)
    return {"status": receipt["status"], "removed_paths": len(removed),
            "cleanup_receipt": str(CLEANUP_RECEIPT)}


def plan_mode() -> dict[str, Any]:
    controls = verify_frozen_control_files()
    binaries = expected_binary_descriptors(allow_pending=True)
    return {
        "schema": "litchi.performance.0945.archive-plan.v1",
        "status": "prepared_not_started",
        "script": str(SCRIPT),
        "source_root": str(SOURCE_ROOT),
        "workspace": str(WORK_ROOT),
        "archive": str(ARCHIVE),
        "manifest": str(MANIFEST),
        "member_list": str(MEMBERS),
        "verification_extraction": str(EXTRACT_ROOT),
        "include": "all regular files under 0945, including source/quality copies, receipts, helpers, snapshots, raw outputs, and binaries",
        "exclude_exact_directories": sorted(path.as_posix() for path in EXCLUDED_DIRECTORIES),
        "preserve_exact_target_files": sorted(path.as_posix() for path in ALLOWLISTED_TARGET_FILES),
        "preserve_nested_target_names_elsewhere": True,
        "binary_descriptor": str(BINARY_DESCRIPTOR),
        "bound_binary_status": binaries,
        "frozen_control_files": controls,
        "quality_input_registry": str(SOURCE_ROOT / QUALITY_INPUT_REGISTRY),
        "go_gate": {"flag": "--go", "environment": f"{GO_ENV}=1"},
        "exact_regeneration_commands": {
            "archive": script_command("--archive"),
            "verify": script_command("--verify"),
            "cleanup": script_command("--cleanup"),
        },
    }


def main(argv: list[str] | None = None) -> int:
    global _GO_FLAG
    parser = argparse.ArgumentParser(description=__doc__)
    modes = parser.add_mutually_exclusive_group(required=True)
    modes.add_argument("--plan", action="store_true")
    modes.add_argument("--archive", action="store_true")
    modes.add_argument("--verify", action="store_true")
    modes.add_argument("--cleanup", action="store_true")
    parser.add_argument("--go", action="store_true", help="required with archive, verify, or cleanup")
    args = parser.parse_args(argv)
    _GO_FLAG = bool(args.go)
    try:
        if args.plan:
            result = plan_mode()
        elif args.archive:
            result = archive_mode()
        elif args.verify:
            result = verify_mode()
        else:
            result = cleanup_mode()
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (ArchiveError, OSError, json.JSONDecodeError, subprocess.SubprocessError) as error:
        print(f"0945 archive operation refused: {type(error).__name__}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
