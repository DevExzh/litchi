"""Create-only plan and guarded cleanup for the 0905 temporary roots.

The default action is read-only: it prints a JSON plan and never writes or
removes anything.  ``--remove`` requires the frozen plan, terminal build and
quality receipts, canonical and micro reductions, a final decision, a passing
replay receipt, source custody, and a process census before it creates a
start receipt and removes only the two marked roots.  Fixture symlinks are
unlinked by name without following them.  Retained packet binaries are only
stat'ed and are never cleanup targets.
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
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
REPO = PACKET.parents[3]
SOURCE_ROOT = Path("/tmp/litchi-0905-src")
BUILD_ROOT = REPO / "target/litchi-0905-build"
MARKER_NAME = ".litchi-0905-owned.json"
SOURCE_MARKER = {
    "owner": "change-0905",
    "source": str(SOURCE_ROOT),
    "target": str(BUILD_ROOT),
}

CLEANUP_START = PACKET / "cleanup-start.json"
CLEANUP_RECEIPT = PACKET / "cleanup.json"
CLEANUP_FAILURE = PACKET / "cleanup-failure.json"

PLAN = PACKET / "plan.json"
QUALITY = PACKET / "quality.json"
ANALYSIS = PACKET / "analysis.json"
MICRO_ANALYSIS = PACKET / "micro-analysis.json"
DECISION = PACKET / "decision.json"
REPLAY_NAMES = ("post-cleanup-replay.json", "replay.json", "raw-replay.json")

SCHEMA = "litchi.xml-audit-cleanup-0905.v1"
CANDIDATE_FILE = "crates/xml-minifier/src/audit.rs"


class CleanupError(RuntimeError):
    """A custody, terminal-receipt, or cleanup precondition failed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CleanupError(message)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON receipt: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise CleanupError(f"cannot read JSON receipt {path}: {error}") from error


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        raise CleanupError(f"cannot hash {path}: {error}") from error
    return digest.hexdigest()


def file_descriptor(path: Path, *, hash_bytes: bool = True) -> dict[str, Any]:
    require(not path.is_symlink(), f"file must not be a symlink: {path}")
    require(path.is_file(), f"file is missing or not regular: {path}")
    mode = stat.S_IMODE(path.stat().st_mode)
    result: dict[str, Any] = {
        "path": str(path.resolve()),
        "bytes": path.stat().st_size,
        "mode": mode,
        "executable": bool(mode & 0o111),
    }
    if hash_bytes:
        result["sha256"] = sha256_file(path)
    return result


def verify_descriptor(value: Any, label: str, *, hash_bytes: bool = True) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("path"), str),
            f"{label} is not a file descriptor")
    path = Path(value["path"])
    if not path.is_absolute():
        path = (PACKET / path).resolve()
    actual = file_descriptor(path, hash_bytes=hash_bytes)
    for key in ("bytes", "mode", "executable"):
        if key in value:
            require(actual[key] == value[key], f"{label}.{key} differs from {path}")
    if hash_bytes and "sha256" in value:
        require(actual["sha256"] == value["sha256"], f"{label}.sha256 differs from {path}")
    return actual


def write_once(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(),
            f"refusing to overwrite create-only receipt: {path}")
    encoded = json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False) + "\n"
    try:
        with path.open("x", encoding="utf-8") as stream:
            stream.write(encoded)
    except (FileExistsError, OSError) as error:
        raise CleanupError(f"could not create receipt {path}: {error}") from error


def resolve_descriptor(value: Any, label: str) -> Path:
    if isinstance(value, dict):
        raw_path = value.get("path")
    else:
        raw_path = value
    require(isinstance(raw_path, str) and raw_path,
            f"{label} lacks a path descriptor")
    path = Path(raw_path)
    return path.resolve() if path.is_absolute() else (PACKET / path).resolve()


def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "litchi.paired-audit-capture-0905.v1",
            "plan schema is not the frozen 0905 capture schema")
    require(plan.get("status") in {"frozen", "capture-ready", "ready"},
            f"plan is not terminal/frozen: {plan.get('status')!r}")
    require(plan.get("plan_id") == "0905" and plan.get("seed") == 905009,
            "plan identity or seed changed")
    source = plan.get("source")
    require(isinstance(source, dict), "plan.source is missing")
    require(source.get("file_count") == 10773, "plan source file count is not 10,773")
    require(source.get("candidate_files") == [CANDIDATE_FILE],
            "plan does not select exactly audit.rs")
    return plan


def load_manifest(value: Any, label: str) -> tuple[Path, dict[str, dict[str, Any]]]:
    path = resolve_descriptor(value, label)
    verify_descriptor(value, label)
    raw = read_json(path)
    files = raw.get("files") if isinstance(raw, dict) else raw
    require(isinstance(files, dict) and files, f"{label} has no files map")
    normalized: dict[str, dict[str, Any]] = {}
    for name, item in files.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"{label} has an unsafe path {name!r}")
        relative = Path(name)
        require(".." not in relative.parts and relative.as_posix() == name,
                f"{label} has a non-normalized path {name!r}")
        require(isinstance(item, dict) and isinstance(item.get("bytes"), int)
                and isinstance(item.get("sha256"), str)
                and len(item["sha256"]) == 64,
                f"{label} entry {name!r} is malformed")
        normalized[name] = item
    return path, normalized


def source_custody(plan: dict[str, Any]) -> dict[str, Any]:
    source = plan["source"]
    baseline_path, baseline = load_manifest(source["baseline_manifest"], "baseline source manifest")
    candidate_path, candidate = load_manifest(source["candidate_manifest"], "candidate source manifest")
    require(len(candidate) == source["file_count"] and set(candidate) == set(baseline),
            "source manifests do not have the frozen common file set")
    changed = {
        name for name in candidate
        if (baseline[name].get("bytes"), baseline[name].get("sha256"))
        != (candidate[name].get("bytes"), candidate[name].get("sha256"))
    }
    require(changed == {CANDIDATE_FILE},
            f"source manifest delta is not exactly {CANDIDATE_FILE}: {sorted(changed)}")

    root_worktree = Path(source.get("root_worktree", str(REPO))).resolve()
    source_root = Path(source["source_root"]).resolve()
    require(root_worktree == REPO.resolve(), "plan root_worktree is not the shared repository")
    require(source_root == SOURCE_ROOT.resolve(), "plan source_root is not the owned source root")
    selected = candidate[CANDIDATE_FILE]
    selected_actual: dict[str, Any] | None = None
    for name, expected in candidate.items():
        actual = file_descriptor(source_root / name)
        require(actual["bytes"] == expected["bytes"]
                and actual["sha256"] == expected["sha256"],
                f"owned candidate source does not match the candidate manifest: {name}")
        if name == CANDIDATE_FILE:
            selected_actual = actual
    require(selected_actual is not None,
            "candidate manifest omitted the selected candidate source")

    packet_candidate = PACKET / "candidate-audit.rs"
    if packet_candidate.exists():
        packet_actual = file_descriptor(packet_candidate)
        require(packet_actual["bytes"] == selected["bytes"]
                and packet_actual["sha256"] == selected["sha256"],
                "retained candidate-audit.rs does not match the selected manifest")
    return {
        "baseline_manifest": file_descriptor(baseline_path),
        "candidate_manifest": file_descriptor(candidate_path),
        "candidate_file": CANDIDATE_FILE,
        "candidate_file_descriptor": selected_actual,
        "hashed_entries": len(candidate),
        "file_count": len(candidate),
        "changed_files": sorted(changed),
        "root_worktree": str(root_worktree),
        "source_root": str(source_root),
    }


def marker_record(root: Path) -> dict[str, Any]:
    require(root == root.resolve(), f"owned root path is not canonical: {root}")
    require(root.is_dir() and not root.is_symlink(), f"owned root is missing or unsafe: {root}")
    marker = root / MARKER_NAME
    require(marker.is_file() and not marker.is_symlink(), f"ownership marker is missing: {marker}")
    value = read_json(marker)
    require(value == SOURCE_MARKER, f"ownership marker mismatch: {marker}")
    return {"path": str(root), "marker": file_descriptor(marker), "marker_value": value}


def inventory(root: Path) -> dict[str, Any]:
    marker_record(root)
    files = 0
    logical_bytes = 0
    symlinks: list[dict[str, str]] = []
    special: list[str] = []
    for current, directories, names in os.walk(root, topdown=True, followlinks=False):
        current_path = Path(current)
        for name in list(directories):
            path = current_path / name
            if path.is_symlink():
                symlinks.append({"path": str(path), "target": os.readlink(path)})
                directories.remove(name)
        for name in names:
            path = current_path / name
            if path.is_symlink():
                symlinks.append({"path": str(path), "target": os.readlink(path)})
                continue
            mode = os.lstat(path).st_mode
            if stat.S_ISREG(mode):
                files += 1
                logical_bytes += os.lstat(path).st_size
            else:
                special.append(str(path))
    require(not special, f"owned root contains special entries: {special[:8]}")
    return {
        **marker_record(root),
        "files": files,
        "logical_bytes": logical_bytes,
        "symlinks": sorted(symlinks, key=lambda item: item["path"]),
    }


def verify_fixture_links(source_inventory: dict[str, Any]) -> list[dict[str, str]]:
    allowed = {
        str(SOURCE_ROOT / "docs"): str(REPO / "docs"),
        str(SOURCE_ROOT / "3rdparty"): str(REPO / "3rdparty"),
    }
    links = source_inventory["symlinks"]
    for item in links:
        require(item["path"] in allowed and item["target"] == allowed[item["path"]],
                f"unexpected owned-root symlink: {item}")
    return links


def process_matches(roots: tuple[Path, ...]) -> list[dict[str, Any]]:
    matches: list[dict[str, Any]] = []
    current_pid = os.getpid()

    def match_target(value: str, root: Path) -> bool:
        return value == str(root) or value.startswith(str(root) + os.sep)

    for entry in sorted(Path("/proc").glob("[0-9]*"), key=lambda item: int(item.name)):
        try:
            pid = int(entry.name)
        except ValueError:
            continue
        if pid == current_pid:
            continue
        try:
            cwd = os.readlink(entry / "cwd")
            command = (entry / "cmdline").read_bytes().replace(b"\0", b" ").decode(
                "utf-8", "replace"
            ).strip()
        except (FileNotFoundError, PermissionError, ProcessLookupError, OSError):
            continue
        evidence: list[dict[str, str]] = []
        for root in roots:
            if match_target(cwd, root) or str(root) in command:
                evidence.append({"root": str(root), "source": "cwd-or-command"})
        try:
            for descriptor in (entry / "fd").iterdir():
                target = os.readlink(descriptor)
                for root in roots:
                    if match_target(target, root):
                        evidence.append({
                            "root": str(root),
                            "source": f"fd/{descriptor.name}",
                            "target": target,
                        })
        except (FileNotFoundError, PermissionError, ProcessLookupError, OSError):
            pass
        if evidence:
            matches.append({"pid": pid, "cwd": cwd, "command": command, "evidence": evidence})
    return matches


def terminal_builds(plan: dict[str, Any]) -> list[dict[str, Any]]:
    roles = plan.get("roles")
    require(isinstance(roles, dict), "plan.roles is missing")
    source = plan["source"]
    _, baseline = load_manifest(source["baseline_manifest"], "baseline source manifest")
    _, candidate = load_manifest(source["candidate_manifest"], "candidate source manifest")
    expected = {"baseline": baseline[CANDIDATE_FILE], "candidate": candidate[CANDIDATE_FILE]}
    records: list[dict[str, Any]] = []
    for role in ("baseline", "candidate"):
        role_value = roles.get(role)
        require(isinstance(role_value, dict), f"plan role is missing: {role}")
        for kind in ("native", "allocator"):
            entry = role_value.get(kind)
            require(isinstance(entry, dict), f"plan role is missing: {role}/{kind}")
            receipt_ref = entry.get("build_receipt")
            receipt_path = resolve_descriptor(receipt_ref, f"{role}/{kind} build receipt")
            receipt_descriptor = verify_descriptor(receipt_ref, f"{role}/{kind} build receipt")
            receipt = read_json(receipt_path)
            require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
                    f"{role}/{kind} build did not terminate successfully")
            require(receipt.get("inputs_unchanged") is True
                    and isinstance(receipt.get("inputs_before"), list)
                    and receipt.get("inputs_before") == receipt.get("inputs_after"),
                    f"{role}/{kind} build input census is not terminal and unchanged")
            selected = [item for item in receipt["inputs_after"]
                        if isinstance(item, dict) and item.get("path", "").endswith(CANDIDATE_FILE)]
            require(len(selected) == 1, f"{role}/{kind} build does not bind one selected source file")
            require(selected[0].get("sha256") == expected[role]["sha256"],
                    f"{role}/{kind} build binds the wrong selected source manifest")
            binary = entry.get("path")
            binary_path = resolve_descriptor(binary, f"{role}/{kind} binary")
            binary_actual = file_descriptor(binary_path, hash_bytes=False)
            require(binary_actual["executable"], f"{role}/{kind} binary is not executable")
            records.append({
                "role": role,
                "kind": kind,
                "receipt": receipt_descriptor,
                "binary": binary_actual,
                "selected_source_sha256": selected[0]["sha256"],
            })
    return records


def terminal_quality() -> dict[str, Any]:
    quality = read_json(QUALITY)
    require(isinstance(quality, dict) and quality.get("status") == "pass",
            "quality.json is not a passing terminal receipt")
    gates = quality.get("gates")
    require(isinstance(gates, list) and gates, "quality.json has no gates")
    receipts: list[dict[str, Any]] = []
    for gate in gates:
        require(isinstance(gate, dict) and isinstance(gate.get("receipt"), dict),
                "quality gate has no receipt descriptor")
        receipt_path = resolve_descriptor(gate["receipt"], f"quality/{gate.get('label')}")
        descriptor = verify_descriptor(gate["receipt"], f"quality/{gate.get('label')}")
        value = read_json(receipt_path)
        require(isinstance(value, dict) and value.get("exit_code") == 0,
                f"quality gate did not terminate successfully: {gate.get('label')}")
        receipts.append({"label": gate.get("label"), "receipt": descriptor})
    return {"status": quality["status"], "gates": receipts}


def terminal_review_artifacts() -> dict[str, Any]:
    analysis = read_json(ANALYSIS)
    require(isinstance(analysis, dict) and analysis.get("status") == "pass"
            and analysis.get("stage") == "all",
            "canonical analysis is not a passing all-stage reduction")
    micro = read_json(MICRO_ANALYSIS)
    require(isinstance(micro, dict) and micro.get("status") == "pass",
            "micro analysis is not a passing reduction")
    decision = read_json(DECISION)
    require(isinstance(decision, dict)
            and decision.get("status") in {"accepted", "rejected"}
            and isinstance(decision.get("production_retained"), bool),
            "decision is not a terminal accepted/rejected receipt")
    replay_path = next((PACKET / name for name in REPLAY_NAMES if (PACKET / name).is_file()), None)
    require(replay_path is not None, "no canonical replay receipt is present")
    replay = read_json(replay_path)
    require(isinstance(replay, dict) and replay.get("status") == "pass",
            f"replay receipt is not passing: {replay_path}")
    return {
        "analysis": file_descriptor(ANALYSIS),
        "micro_analysis": file_descriptor(MICRO_ANALYSIS),
        "decision": file_descriptor(DECISION),
        "decision_status": decision["status"],
        "production_retained": decision["production_retained"],
        "replay": file_descriptor(replay_path),
    }


def retained_binaries(plan: dict[str, Any]) -> list[dict[str, Any]]:
    roles = plan.get("roles")
    require(isinstance(roles, dict), "plan.roles is missing")
    records: list[dict[str, Any]] = []
    for role in ("baseline", "candidate"):
        for kind in ("native", "allocator"):
            entry = roles[role][kind]
            path = resolve_descriptor(entry["path"], f"retained {role}/{kind} binary")
            require(path.parent == PACKET / "binaries",
                    f"retained binary escaped packet binaries directory: {path}")
            actual = file_descriptor(path, hash_bytes=False)
            require(actual["executable"], f"retained binary is not executable: {path}")
            records.append({"role": role, "kind": kind, **actual})
    return records


def preflight() -> dict[str, Any]:
    plan = load_plan()
    source = source_custody(plan)
    source_inventory = inventory(SOURCE_ROOT)
    build_inventory = inventory(BUILD_ROOT)
    links = verify_fixture_links(source_inventory)
    processes = process_matches((SOURCE_ROOT, BUILD_ROOT))
    require(not processes, f"active process references owned roots: {processes}")
    builds = terminal_builds(plan)
    quality = terminal_quality()
    review = terminal_review_artifacts()
    binaries = retained_binaries(plan)
    require(not CLEANUP_START.exists() and not CLEANUP_RECEIPT.exists(),
            "cleanup already has a retained start/final receipt")
    require(not CLEANUP_FAILURE.exists(), "a prior cleanup failure receipt requires review")
    return {
        "schema": SCHEMA,
        "status": "ready",
        "checked_utc": utc_now(),
        "source": source,
        "roots": [source_inventory, build_inventory],
        "fixture_symlinks": links,
        "process_matches": processes,
        "terminal_builds": builds,
        "terminal_quality": quality,
        "terminal_review": review,
        "retained_binaries": binaries,
        "worktrees_created": 0,
        "worktrees_removed": 0,
    }


def unlink_fixture_links(links: list[dict[str, str]]) -> list[dict[str, Any]]:
    removed: list[dict[str, Any]] = []
    for item in links:
        path = Path(item["path"])
        require(path.is_symlink(), f"fixture link changed before unlink: {path}")
        target_exists_before = Path(item["target"]).exists()
        path.unlink()
        require(not path.exists() and not path.is_symlink(), f"fixture link remains: {path}")
        removed.append({**item, "target_exists_before": target_exists_before,
                        "unlinked_without_following": True})
    return removed


def remove_roots(plan: dict[str, Any], preflight_result: dict[str, Any]) -> dict[str, Any]:
    require(SOURCE_ROOT == SOURCE_ROOT.resolve() and BUILD_ROOT == BUILD_ROOT.resolve(),
            "cleanup roots are not canonical")
    before = preflight_result["roots"]
    links = preflight_result["fixture_symlinks"]
    removed_links = unlink_fixture_links(links)
    for root in (SOURCE_ROOT, BUILD_ROOT):
        require(root.is_dir() and not root.is_symlink(), f"owned root changed before removal: {root}")
        shutil.rmtree(root)
        require(not root.exists() and not root.is_symlink(), f"owned root remains after removal: {root}")
    after = [{"path": str(root), "removed": not root.exists(), "files": 0,
              "logical_bytes": 0, "symlinks": []}
             for root in (SOURCE_ROOT, BUILD_ROOT)]
    return {
        "schema": SCHEMA,
        "status": "pass",
        "finished_utc": utc_now(),
        "preflight": {key: value for key, value in preflight_result.items()
                       if key not in {"roots", "fixture_symlinks", "process_matches"}},
        "roots_before": before,
        "roots_after": after,
        "fixture_symlinks_unlinked_without_following": removed_links,
        "process_matches_before_remove": preflight_result["process_matches"],
        "worktrees_created": 0,
        "worktrees_removed": 0,
        "retained_results_binaries": preflight_result["retained_binaries"],
    }


def plan_only() -> int:
    try:
        result = preflight()
    except CleanupError as error:
        print(json.dumps({"schema": SCHEMA, "status": "blocked", "reason": str(error)}, indent=2))
        return 0
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


def remove() -> int:
    result = preflight()
    write_once(CLEANUP_START, {
        "schema": SCHEMA,
        "status": "started",
        "started_utc": utc_now(),
        "preflight": result,
    })
    try:
        final = remove_roots(load_plan(), result)
    except Exception as error:
        if not CLEANUP_FAILURE.exists() and not CLEANUP_FAILURE.is_symlink():
            write_once(CLEANUP_FAILURE, {
                "schema": SCHEMA,
                "status": "failed",
                "failed_utc": utc_now(),
                "error": str(error),
                "cleanup_start": file_descriptor(CLEANUP_START),
            })
        raise
    write_once(CLEANUP_RECEIPT, final)
    print(json.dumps({"status": final["status"], "roots_removed": 2}, sort_keys=True))
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--remove", action="store_true",
                        help="perform the guarded destructive cleanup")
    args = parser.parse_args()
    return remove() if args.remove else plan_only()


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (CleanupError, OSError, ValueError) as error:
        print(f"0905 cleanup failed: {error}", file=sys.stderr)
        raise SystemExit(1)
