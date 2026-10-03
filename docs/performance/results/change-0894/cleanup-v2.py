#!/usr/bin/env python3
"""Seal and remove only the marked 0894 workspace roots.

The default action is a read-only plan.  ``apply`` is the explicit cleanup
action; it creates a start receipt, checks ownership markers and ``/proc``
command/cwd references, verifies source trees against retained packet
snapshots, relocates probe receipts/logs/binaries into the packet, and then
removes the owned roots.  Every packet output and relocated file is
create-only.  A failed apply leaves its start/failure receipts and any
partial relocation in place for review; it never retries or overwrites them.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
TARGET = ROOT / "target"
SOURCE_MANIFEST = PACKET / "source.json"
QUALITY_SOURCE_RECEIPT = PACKET / "quality-source.json"
CANDIDATE_SOURCE_BEFORE = PACKET / "candidate-source-before.json"
CANDIDATE_SOURCE_AFTER = PACKET / "candidate-source-after.json"
CANDIDATE_BUILD = PACKET / "candidate-build.json"
BINDINGS = PACKET / "bindings.json"
CLEANUP_START = PACKET / "cleanup-v2-start.json"
CLEANUP_RECEIPT = PACKET / "cleanup.json"
CLEANUP_FAILURE = PACKET / "cleanup-v2-failure.json"
PROBE_RETAINED = PACKET / "probe" / "retained"
EARLY_PROBE_CLEANUP = PACKET / "early-probe-target-cleanup.json"
HELPER = Path(__file__).resolve()

ROOT_SPECS = (
    {
        "name": "candidate-source",
        "kind": "source",
        "role": "candidate",
        "path": ROOT / "target/litchi-0894-candidate-src",
        "marker": ".owner-litchi-0894-candidate",
        "marker_text": "litchi 0894 frozen candidate source\n",
    },
    {
        "name": "candidate-build",
        "kind": "build",
        "role": "candidate",
        "path": ROOT / "target/litchi-0894-candidate-build",
        "marker": ".owner-litchi-0894-candidate",
        "marker_text": "litchi 0894 frozen candidate build target\n",
    },
    {
        "name": "quality-source",
        "kind": "source",
        "role": "quality",
        "path": ROOT / "target/litchi-0894-quality-src",
        "marker": ".owner-litchi-0894-quality-src",
        "marker_text": "litchi 0894 isolated quality source\n",
    },
    {
        "name": "quality-build",
        "kind": "build",
        "role": "quality",
        "path": ROOT / "target/litchi-0894-quality",
        "marker": ".owner-0894-quality",
        "marker_text": "litchi 0894 selected-metadata quality target\n",
    },
    {
        "name": "probe-baseline",
        "kind": "probe",
        "role": "baseline",
        "path": ROOT / "target/litchi-0894-probe-baseline",
        "marker": ".owner-litchi-0894-probe-baseline",
        "marker_text": "litchi 0894 observable probe baseline\n",
    },
    {
        "name": "probe-candidate",
        "kind": "probe",
        "role": "candidate",
        "path": ROOT / "target/litchi-0894-probe-candidate",
        "marker": ".owner-litchi-0894-probe-candidate",
        "marker_text": "litchi 0894 observable probe candidate\n",
    },
)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    require(not path.is_symlink(), f"expected a non-symlink regular file: {path}")
    resolved = path.resolve()
    require(resolved.is_file(), f"expected a regular file: {path}")
    info = resolved.stat()
    require(stat.S_ISREG(info.st_mode), f"expected a regular file: {path}")
    return {
        "path": str(resolved),
        "bytes": info.st_size,
        "sha256": sha256(resolved),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def write_once(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"create-only output already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"JSON receipt is missing: {path}")
    with path.open(encoding="utf-8") as stream:
        return json.load(stream)


def relative_file_record(path: Path, root: Path) -> dict[str, Any]:
    item = descriptor(path)
    item["path"] = path.relative_to(root).as_posix()
    return item


def tree_snapshot(root: Path, marker: str) -> dict[str, dict[str, Any]]:
    """Hash a marked source tree, excluding its ownership marker."""
    result: dict[str, dict[str, Any]] = {}
    for path in root.rglob("*"):
        require(not path.is_symlink(), f"owned source tree contains a symlink: {path}")
        if path.is_dir():
            continue
        require(path.is_file(), f"owned source tree contains a non-regular entry: {path}")
        relative = path.relative_to(root).as_posix()
        if relative == marker:
            continue
        result[relative] = relative_file_record(path, root)
    return dict(sorted(result.items()))


def census_digest(records: dict[str, dict[str, Any]]) -> str:
    normalized = {
        name: {
            key: records[name].get(key)
            for key in ("bytes", "sha256", "mode", "executable")
            if key in records[name]
        }
        for name in sorted(records)
    }
    payload = json.dumps(normalized, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(payload).hexdigest()


def summary(records: dict[str, dict[str, Any]]) -> dict[str, Any]:
    return {
        "file_count": len(records),
        "logical_bytes": sum(int(value.get("bytes", 0)) for value in records.values()),
        "census_sha256": census_digest(records),
    }


def normalize_expected(records: dict[str, Any]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for name, raw in records.items():
        require(isinstance(name, str) and isinstance(raw, dict), f"invalid retained source record: {name!r}")
        result[name] = {key: raw[key] for key in ("bytes", "sha256", "mode", "executable") if key in raw}
    return dict(sorted(result.items()))


def compare_census(root: Path, marker: str, expected: dict[str, dict[str, Any]], references: list[Path], label: str) -> dict[str, Any]:
    if not root.exists():
        return {
            "label": label,
            "status": "absent",
            "verified": False,
            "root": str(root),
            "references": [descriptor(path) for path in references if path.is_file()],
            "reason": "owned root was already removed before this cleanup helper ran",
        }
    actual = tree_snapshot(root, marker)
    mismatches: list[dict[str, Any]] = []
    for name in sorted(set(expected) | set(actual)):
        if name not in expected:
            mismatches.append({"path": name, "reason": "unexpected_file", "actual": actual[name]})
            continue
        if name not in actual:
            mismatches.append({"path": name, "reason": "missing_file", "expected": expected[name]})
            continue
        for key in ("bytes", "sha256", "mode", "executable"):
            if key in expected[name] and actual[name].get(key) != expected[name].get(key):
                mismatches.append({"path": name, "reason": f"{key}_mismatch", "expected": expected[name], "actual": actual[name]})
                break
    result = {
        "label": label,
        "status": "pass" if not mismatches else "fail",
        "verified": not mismatches,
        "root": str(root),
        "references": [descriptor(path) for path in references if path.is_file()],
        "expected": summary(expected),
        "actual": summary(actual),
        "mismatch_count": len(mismatches),
        "mismatches": mismatches[:64],
    }
    require(not mismatches, f"{label} source census differs from retained snapshots ({len(mismatches)} mismatches)")
    return result


def root_status(spec: dict[str, Any]) -> dict[str, Any]:
    root = spec["path"]
    require(root.parent == TARGET, f"owned root escaped target directory: {root}")
    if not root.exists():
        return {**spec, "status": "absent", "path": str(root)}
    require(not root.is_symlink() and root.is_dir(), f"owned root is not a regular directory: {root}")
    marker = root / spec["marker"]
    require(not marker.is_symlink() and marker.is_file(), f"owned root marker is missing: {marker}")
    require(marker.read_text(encoding="utf-8") == spec["marker_text"], f"owned root marker text mismatch: {marker}")
    return {
        **spec,
        "status": "present",
        "path": str(root),
        "marker_descriptor": descriptor(marker),
        "marker_text": spec["marker_text"],
    }


def proc_matches(roots: list[Path]) -> list[dict[str, Any]]:
    matches: list[dict[str, Any]] = []
    current = os.getpid()
    for entry in sorted(Path("/proc").glob("[0-9]*"), key=lambda item: int(item.name)):
        try:
            pid = int(entry.name)
        except ValueError:
            continue
        if pid == current:
            continue
        try:
            cwd = os.readlink(entry / "cwd")
            raw_cmdline = (entry / "cmdline").read_bytes()
        except (FileNotFoundError, PermissionError, ProcessLookupError, OSError):
            continue
        command = raw_cmdline.replace(b"\0", b" ").decode("utf-8", "replace").strip()
        matched: list[dict[str, str]] = []
        for root in roots:
            root_text = str(root)
            cwd_match = cwd == root_text or cwd.startswith(root_text + os.sep)
            command_match = root_text in command
            if cwd_match or command_match:
                matched.append({"root": root_text, "cwd_match": str(cwd_match), "command_match": str(command_match)})
        if matched:
            matches.append({"pid": pid, "cwd": cwd, "command": command, "roots": matched})
    return matches


def source_receipt(path: Path) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict), f"source receipt is not an object: {path}")
    files = value.get("files")
    require(isinstance(files, dict), f"source receipt has no full file census: {path}")
    return value


def candidate_expected(require_final: bool = False) -> tuple[dict[str, dict[str, Any]], list[Path]] | None:
    available = [path for path in (CANDIDATE_SOURCE_BEFORE, CANDIDATE_SOURCE_AFTER) if path.is_file()]
    if require_final:
        require(CANDIDATE_SOURCE_BEFORE.is_file() and CANDIDATE_SOURCE_AFTER.is_file(), "final candidate source before/after receipts are required")
        require(CANDIDATE_BUILD.is_file(), "final candidate build receipt is required")
        build = read_json(CANDIDATE_BUILD)
        require(build.get("status") == "pass", "final candidate build receipt does not pass")
        require(build.get("schema") in ("litchi.xlsx-selected-metadata-candidate-build-0894.v1", "litchi.xlsx-candidate-build-0894.v1"), "final candidate build receipt schema mismatch")
        available = [CANDIDATE_SOURCE_BEFORE, CANDIDATE_SOURCE_AFTER]
    if not available:
        return None
    receipts = [source_receipt(path) for path in available]
    maps = [normalize_expected(receipt["files"]) for receipt in receipts]
    if len(maps) == 2:
        require(maps[0] == maps[1], "candidate source before/after retained censuses differ")
    return maps[-1], available


def path_relative_to(path: str, root: Path) -> str:
    value = Path(path).resolve()
    try:
        return value.relative_to(root.resolve()).as_posix()
    except ValueError as error:
        raise RuntimeError(f"retained descriptor escapes source root: {path}") from error


def bound_quality_receipt() -> Path | None:
    """Resolve candidate quality only through the packet binding manifest."""
    if not BINDINGS.is_file():
        return None
    bindings = read_json(BINDINGS)
    item = bindings.get("candidate_quality") if isinstance(bindings, dict) else None
    if not isinstance(item, dict) or not isinstance(item.get("path"), str):
        return None
    path = Path(item["path"])
    if not path.is_absolute():
        path = PACKET / path
    return path.resolve()


def quality_expected(require_binding: bool = False) -> tuple[dict[str, dict[str, Any]], list[Path]]:
    manifest = read_json(SOURCE_MANIFEST)
    require(manifest.get("status") == "pass" and isinstance(manifest.get("files"), dict), "baseline source manifest is not a full passing census")
    quality = read_json(QUALITY_SOURCE_RECEIPT)
    require(quality.get("status") == "pass", "quality source receipt is not passing")
    root = ROOT / "target/litchi-0894-quality-src"
    expected = normalize_expected(manifest["files"])
    references = [SOURCE_MANIFEST, QUALITY_SOURCE_RECEIPT]

    for section_name in ("support", "overlays"):
        section = quality.get(section_name, {})
        require(isinstance(section, dict), f"quality source {section_name} inventory is invalid")
        for name, item in section.items():
            require(isinstance(item, dict) and isinstance(item.get("path"), str), f"quality source {section_name} descriptor is invalid: {name}")
            relative = path_relative_to(item["path"], root)
            require(relative == name, f"quality source {section_name} path mismatch: {name}")
            expected[relative] = normalize_expected({relative: item})[relative]

    lock = quality.get("lock")
    require(isinstance(lock, dict) and isinstance(lock.get("path"), str), "quality source generated lock descriptor is missing")
    lock_relative = path_relative_to(lock["path"], root)
    require(lock_relative == "Cargo.lock", "quality source generated lock is not root Cargo.lock")
    expected[lock_relative] = normalize_expected({lock_relative: lock})[lock_relative]

    bound = bound_quality_receipt()
    if require_binding:
        require(bound is not None and bound.is_file(), "candidate quality binding receipt is missing")
    if bound is not None and bound.is_file():
        receipt = read_json(bound)
        if receipt.get("status") != "pass":
            require(not require_binding, "bound candidate quality receipt does not pass")
            references.extend((BINDINGS, bound))
            return expected, references
        for label in ("candidate_sources_after", "candidate_sources_before"):
            records = receipt.get(label)
            if not isinstance(records, list):
                continue
            for item in records:
                if not isinstance(item, dict) or not isinstance(item.get("relative_path"), str):
                    continue
                source = item.get("source")
                if isinstance(source, dict):
                    expected[item["relative_path"]] = normalize_expected({item["relative_path"]: source})[item["relative_path"]]
            references.extend((BINDINGS, bound))
            break
    return expected, references


def source_censuses(statuses: list[dict[str, Any]], require_final_candidate: bool = False, require_quality_binding: bool = False) -> list[dict[str, Any]]:
    results: list[dict[str, Any]] = []
    for status in statuses:
        if status["kind"] != "source":
            continue
        root = Path(status["path"])
        if status["name"] == "candidate-source":
            expected = candidate_expected(require_final=require_final_candidate)
            if expected is None:
                results.append({"label": status["name"], "status": "unavailable", "verified": False, "root": str(root), "reason": "candidate source snapshots are not retained"})
            else:
                records, references = expected
                results.append(compare_census(root, status["marker"], records, references, status["name"]))
        elif status["name"] == "quality-source":
            records, references = quality_expected(require_binding=require_quality_binding)
            results.append(compare_census(root, status["marker"], records, references, status["name"]))
    return results


def copy_create(source: Path, destination: Path) -> dict[str, Any]:
    before = descriptor(source)
    require(not destination.exists() and not destination.is_symlink(), f"retained probe output already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    with source.open("rb") as input_stream, destination.open("xb") as output_stream:
        shutil.copyfileobj(input_stream, output_stream, length=1024 * 1024)
        output_stream.flush()
        os.fsync(output_stream.fileno())
    destination.chmod(before["mode"])
    after = descriptor(destination)
    require(after["bytes"] == before["bytes"] and after["sha256"] == before["sha256"], f"retained probe copy changed: {source}")
    return {"source": before, "retained": after}


def retain_probe(status: dict[str, Any]) -> dict[str, Any]:
    root = Path(status["path"])
    role = status["role"]
    if status["status"] == "absent":
        return {"role": role, "status": "absent", "root": str(root), "retained": None, "relocations": []}
    destination = PROBE_RETAINED / role
    require(not destination.exists() and not destination.is_symlink(), f"probe retained root already exists: {destination}")
    selected: list[Path] = []
    marker = root / status["marker"]
    selected.append(marker)
    for path in sorted(root.iterdir()):
        if path.is_file() and (path.name in {"build-start.json", "build.json", "build-failure.json", "Cargo.toml", "Cargo.lock", "probe.rs", "xlsx0894-probe"} or path.suffix == ".txt"):
            selected.append(path)
    cargo_binary = root / "cargo-target/debug/xlsx0894-probe"
    if cargo_binary.is_file() and not cargo_binary.is_symlink():
        selected.append(cargo_binary)
    require(selected, f"probe root has no retainable receipts/logs/binaries: {root}")
    relocations: list[dict[str, Any]] = []
    for source in selected:
        relative = source.relative_to(root)
        retained = destination / relative
        copied = copy_create(source, retained)
        relocations.append({
            "old_path": str(source.resolve()),
            "new_path": str(retained.resolve()),
            "relative_path": relative.as_posix(),
            **copied,
        })
    return {
        "role": role,
        "status": "pass",
        "root": str(root),
        "retained": str(destination.resolve()),
        "file_count": len(relocations),
        "relocations": relocations,
    }


def remove_owned(status: dict[str, Any]) -> dict[str, Any]:
    root = Path(status["path"])
    if status["status"] == "absent":
        return {"name": status["name"], "path": str(root), "status": "absent"}
    require(not root.is_symlink() and root.is_dir(), f"owned root disappeared or changed before removal: {root}")
    shutil.rmtree(root)
    require(not root.exists(), f"owned root remains after removal: {root}")
    return {"name": status["name"], "path": str(root), "status": "removed", "marker": status["marker_descriptor"]}


def plan() -> dict[str, Any]:
    statuses = [root_status(spec) for spec in ROOT_SPECS]
    roots = [Path(item["path"]) for item in statuses if item["status"] == "present"]
    processes = proc_matches(roots)
    census = source_censuses(statuses)
    quality_binding = bound_quality_receipt()
    quality_binding_status = None
    if quality_binding is not None and quality_binding.is_file():
        quality_binding_status = read_json(quality_binding).get("status")
    return {
        "schema": "litchi.xlsx-0894-cleanup-plan.v1",
        "status": "blocked" if processes or quality_binding_status != "pass" or any(item.get("status") in ("fail", "unavailable") for item in census) else "ready",
        "created_utc": now(),
        "helper": descriptor(HELPER),
        "roots": statuses,
        "active_processes": processes,
        "source_census": census,
        "candidate_quality_binding": {"path": str(quality_binding), "status": quality_binding_status} if quality_binding is not None else None,
        "early_probe_target_cleanup": descriptor(EARLY_PROBE_CLEANUP) if EARLY_PROBE_CLEANUP.is_file() else None,
        "probe_retention_root": str(PROBE_RETAINED),
        "performance_claim": None,
    }


def apply_cleanup() -> int:
    for path in (CLEANUP_START, CLEANUP_RECEIPT, CLEANUP_FAILURE):
        require(not path.exists() and not path.is_symlink(), f"cleanup receipt already exists: {path}")
    initial = plan()
    write_once(CLEANUP_START, {**initial, "schema": "litchi.xlsx-0894-cleanup-start.v1", "status": "building"})
    try:
        require(not initial["active_processes"], "owned root has an active process; cleanup is refused")
        strict_census = source_censuses(
            [root_status(spec) for spec in ROOT_SPECS],
            require_final_candidate=True,
            require_quality_binding=True,
        )
        early_quality_path = PACKET / "early-quality-root-cleanup.json"
        early_quality = read_json(early_quality_path)
        require(early_quality.get("status") == "pass", "early quality cleanup is not passing")
        require(early_quality.get("source_census", {}).get("status") == "pass", "early quality source census was not verified")
        retained_lock = early_quality["private_lock_retention"]["retained"]
        require(descriptor(Path(retained_lock["path"])) == retained_lock, "retained quality lock changed")
        require(early_quality["quality"] == descriptor(bound_quality_receipt()), "early quality receipt differs from frozen binding")
        for item in strict_census:
            if item.get("label") == "quality-source" and item.get("status") == "absent":
                require(item["root"] == early_quality["source_census"]["root"], "absent quality root is not the verified early-cleanup root")
                require(any(record["path"] == item["root"] and record["status"] == "removed" for record in early_quality["removed"]), "quality root removal was not recorded")
            else:
                require(item.get("status") == "pass", "final source census verification failed")
        statuses = [root_status(spec) for spec in ROOT_SPECS]
        probe_retention = [retain_probe(item) for item in statuses if item["kind"] == "probe"]
        roots = [Path(item["path"]) for item in statuses if item["status"] == "present"]
        active_after_retention = proc_matches(roots)
        require(not active_after_retention, "owned root gained an active process before removal")
        removed = [remove_owned(item) for item in statuses]
        receipt = {
            "schema": "litchi.xlsx-0894-cleanup.v1",
            "status": "pass",
            "created_utc": now(),
            "helper": descriptor(HELPER),
            "start": descriptor(CLEANUP_START),
            "roots": statuses,
            "source_census": strict_census,
            "early_quality_root_cleanup": descriptor(early_quality_path),
            "recovery_from": descriptor(PACKET / "cleanup-failure.json"),
            "probe_retention": probe_retention,
            "early_probe_target_cleanup": descriptor(EARLY_PROBE_CLEANUP) if EARLY_PROBE_CLEANUP.is_file() else None,
            "active_processes_before": initial["active_processes"],
            "active_processes_after_retention": active_after_retention,
            "removed": removed,
            "worktrees_created": [],
            "performance_claim": None,
        }
        write_once(CLEANUP_RECEIPT, receipt)
        print(json.dumps(receipt, indent=2, sort_keys=True))
        return 0
    except Exception as error:
        failure = {
            "schema": "litchi.xlsx-0894-cleanup-failure.v1",
            "status": "fail",
            "created_utc": now(),
            "helper": descriptor(HELPER),
            "start": descriptor(CLEANUP_START),
            "error": str(error),
            "performance_claim": None,
        }
        write_once(CLEANUP_FAILURE, failure)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("plan", "apply"), nargs="?", default="plan")
    args = parser.parse_args()
    if args.action == "plan":
        print(json.dumps(plan(), indent=2, sort_keys=True))
        return 0
    return apply_cleanup()


if __name__ == "__main__":
    raise SystemExit(main())
