#!/usr/bin/env python3
"""Validation rules for the signed native qualification protocol.

This module is intentionally independent of the Rust probe.  It validates
the probe's correctness receipts, endpoint provenance, and the before/after
capability boundary before a guarded timing run is allowed to start.
"""
from __future__ import annotations

import ctypes
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import math
import re
import signal
import sys
import threading
import time
import tomllib
from typing import Any

SCHEMA = "litchi_signed_native_probe.v1"
ENDPOINT_SCHEMA = "litchi_signed_native_endpoint.v1"
FULL_COMMIT = re.compile(r"^[0-9a-f]{40}$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")

TARGET_NS = 500_000_000
MIN_SAMPLE_NS = 150_000_000
MAX_CALLS = 2_000_000
MIN_ABBA_ROUNDS = 5
CHILD_TIMEOUT_S = 300.0
CHILD_OUTPUT_LIMIT = 4 * 1024 * 1024
CHILD_RSS_LIMIT = 20 * 1024 * 1024 * 1024
MEMORY_OUTPUT_LIMIT = 4 * 1024 * 1024
MEMORY_RSS_LIMIT = 20 * 1024 * 1024 * 1024
BASELINE_SOURCE_COMMIT = "f2bbbc8d07ab5da26c180f1ff74813d39f632d54"
CANDIDATE_SOURCE_COMMIT = "472b6705475b073e7fe08f6e11ee17fb1e29f838"
REVIEWED_PARSER_SHA256 = "e80c9706189f942ccce645685dafea8612e5e0dc6ea5e659bd1490ffc3256cc4"

ROUTES = (
    "signed_leaf/noop_commit",
    "signed_facade/noop_commit",
    "signed_leaf/noop_write",
    "signed_facade/noop_write",
    "signed_leaf/noop_replay",
    "signed_facade/noop_replay",
    "signed_leaf/refusal",
    "signed_facade/refusal",
    "signed_leaf/unsign_write",
    "signed_facade/unsign_write",
    "signed_leaf/unsigned_negative",
    "signed_facade/unsigned_negative",
    "signed_leaf/source_freshness",
    "signed_facade/source_freshness",
    "signed_leaf/budget",
    "signed_facade/budget",
)

CORRECTNESS_ONLY = frozenset(
    {
        "signed_leaf/refusal",
        "signed_facade/refusal",
        "signed_leaf/unsigned_negative",
        "signed_facade/unsigned_negative",
        "signed_leaf/source_freshness",
        "signed_facade/source_freshness",
        "signed_leaf/budget",
        "signed_facade/budget",
    }
)
TIMED = tuple(route for route in ROUTES if route not in CORRECTNESS_ONLY)
FINAL_ONLY = frozenset(
    {
        "signed_leaf/noop_commit",
        "signed_facade/noop_commit",
        "signed_leaf/noop_write",
        "signed_facade/noop_write",
        "signed_leaf/noop_replay",
        "signed_facade/noop_replay",
        "signed_leaf/unsign_write",
        "signed_facade/unsign_write",
    }
)
BASELINE_NOT_APPLICABLE = frozenset(
    {
        "signed_leaf/unsigned_negative",
        "signed_facade/unsigned_negative",
        "signed_leaf/source_freshness",
        "signed_facade/source_freshness",
        "signed_leaf/budget",
        "signed_facade/budget",
    }
)

ORIGINS = {"Opc", "Cfb"}
CRYPTO_STATUS = "not_performed"
NATIVE_FORMATS = frozenset({"DOCX", "PPTX", "XLSX", "XLSB", "DOC", "PPT", "XLS"})
REPOSITORY = Path(__file__).resolve().parents[4]
REVIEWED_W8_POLICY = REPOSITORY / "tools" / "w8_projection_policy.py"
REVIEWED_W8_POLICY_JSON = REPOSITORY / "docs/unified-ops/results/0362yh/original-to-final-policy.json"
REVIEWED_W8_BEFORE_PROJECTION = (
    REPOSITORY / "docs/unified-ops/results/0362yh/images/before/projection.json"
)
_REVIEWED_PARSER_MODULE: Any | None = None
_SUBREAPER_ENABLED = False
_ACTIVE_PROCESS_TREE: Any | None = None
REVIEWED_W8_AFTER_PROJECTION = (
    REPOSITORY / "docs/unified-ops/results/0362yh/images/after/projection.json"
)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def _proc_records() -> dict[int, tuple[int, int, int, str]]:
    """Return pid -> (parent, process-group, start-time, state) without /proc races."""
    records: dict[int, tuple[int, int, int, str]] = {}
    for entry in Path("/proc").iterdir():
        if not entry.name.isdigit():
            continue
        try:
            raw = (entry / "stat").read_text()
            close = raw.rfind(")")
            fields = raw[close + 2 :].split()
            records[int(entry.name)] = (
                int(fields[1]),
                int(fields[2]),
                int(fields[19]),
                fields[0],
            )
        except (FileNotFoundError, PermissionError, ValueError, IndexError):
            continue
    return records


def enable_child_subreaper() -> None:
    """Make this runner adopt orphaned descendants so it can kill and reap them."""
    global _SUBREAPER_ENABLED
    if _SUBREAPER_ENABLED:
        return
    try:
        libc = ctypes.CDLL(None, use_errno=True)
        prctl = libc.prctl
        prctl.argtypes = [ctypes.c_int, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong, ctypes.c_ulong]
        prctl.restype = ctypes.c_int
        result = prctl(36, 1, 0, 0, 0)  # Linux PR_SET_CHILD_SUBREAPER
    except (AttributeError, OSError) as error:
        raise RuntimeError(f"child subreaper is unavailable: {error}") from error
    if result != 0:
        error_number = ctypes.get_errno()
        raise RuntimeError(f"child subreaper setup failed: errno {error_number}")
    _SUBREAPER_ENABLED = True


class OwnedProcessTree:
    """Track one spawned process and every descendant, including adopted ones.

    This helper is intentionally scoped to the standalone qualification
    runners: one controller thread owns one Popen subtree at a time, and the
    controller does not spawn unrelated children while the tracker is live.
    New direct children after the baseline are therefore owned descendants,
    not a generic embedding guarantee.
    """

    def __init__(self) -> None:
        global _ACTIVE_PROCESS_TREE
        if threading.active_count() != 1:
            raise RuntimeError("process-tree ownership requires a single controller thread")
        if _ACTIVE_PROCESS_TREE is not None:
            raise RuntimeError("process-tree ownership is already active")
        enable_child_subreaper()
        _ACTIVE_PROCESS_TREE = self
        self._released = False
        self.owner_pid = os.getpid()
        records = _proc_records()
        self._baseline_children = {
            pid: record[2]
            for pid, record in records.items()
            if record[0] == self.owner_pid
        }
        self._root: tuple[int, int | None] | None = None
        self._seen: dict[int, int] = {}
        self._active: dict[int, tuple[int, int, int, str]] = {}

    def attach(self, pid: int) -> None:
        record = _proc_records().get(pid)
        if record is None:
            # A short-lived child can exit between Popen and the first /proc
            # read.  Its descendants are still discoverable through the
            # subreaper's direct-child scan below.
            self._root = (pid, None)
            self._seen = {}
            self._active = {}
            return
        self._root = (pid, record[2])
        self._seen = {pid: record[2]}
        self._active = {pid: record}

    def refresh(self) -> dict[int, tuple[int, int, int, str]]:
        if self._root is None:
            raise RuntimeError("process tree has no root")
        records = _proc_records()
        known = dict(self._seen)
        active: dict[int, tuple[int, int, int, str]] = {}
        root_pid, root_start = self._root
        root_record = records.get(root_pid)
        if root_start is not None and root_record is not None and root_record[2] == root_start:
            active[root_pid] = root_record
        changed = True
        while changed:
            changed = False
            for pid, record in records.items():
                parent, _group, start, _state = record
                previous_start = known.get(pid)
                if previous_start is not None and previous_start != start:
                    continue
                if pid in active:
                    continue
                if parent in active:
                    active[pid] = record
                    known[pid] = start
                    changed = True
        # A double-fork can reparent a descendant directly to this subreaper
        # before the next poll.  New direct children cannot be unrelated runner
        # work: the baseline was captured before the target was spawned.
        for pid, record in records.items():
            parent, _group, start, _state = record
            if (
                parent != self.owner_pid
                or self._baseline_children.get(pid) == start
            ):
                continue
            previous_start = known.get(pid)
            if previous_start is not None and previous_start != start:
                continue
            active[pid] = record
            known[pid] = start
        self._seen = known
        self._active = active
        return dict(active)

    def pids(self) -> set[int]:
        return set(self.refresh())

    def descendant_pids(self) -> set[int]:
        root_pid = self._root[0] if self._root is not None else None
        return self.pids() - ({root_pid} if root_pid is not None else set())

    def escaped_pids(self, process_group: int) -> set[int]:
        active = self.refresh()
        root_pid = self._root[0] if self._root is not None else None
        return {
            pid
            for pid, (_parent, group, _start, _state) in active.items()
            if pid != root_pid and group != process_group
        }

    def kill(self) -> None:
        for pid in sorted(self.pids(), reverse=True):
            if pid == self.owner_pid:
                continue
            try:
                os.kill(pid, signal.SIGKILL)
            except (ProcessLookupError, PermissionError):
                continue

    def reap(self, timeout: float = 5.0) -> bool:
        deadline = time.monotonic() + timeout
        while True:
            active = self.refresh()
            for pid, (_parent, _group, _start, _state) in active.items():
                if pid == self.owner_pid or pid == (self._root[0] if self._root else None):
                    continue
                try:
                    os.waitpid(pid, os.WNOHANG)
                except (ChildProcessError, ProcessLookupError, PermissionError):
                    continue
            active = self.refresh()
            if not active or time.monotonic() >= deadline:
                return not active
            time.sleep(0.01)

    def empty(self) -> bool:
        return not self.pids()

    def release(self) -> None:
        global _ACTIVE_PROCESS_TREE
        if not self._released and _ACTIVE_PROCESS_TREE is self:
            _ACTIVE_PROCESS_TREE = None
        self._released = True


def reject_symlink_path(path: Path, label: str) -> None:
    """Reject a path or any existing parent symlink before canonicalizing it."""
    raw = Path(path)
    lexical = raw if raw.is_absolute() else Path.cwd() / raw
    for parent in (lexical, *lexical.parents):
        try:
            if parent.is_symlink():
                raise ValueError(f"{label} may not traverse a symlink: {parent}")
        except OSError as error:
            raise ValueError(f"{label} cannot inspect {parent}: {error}") from error


def artifact(path: Path) -> dict[str, object]:
    path = Path(path)
    reject_symlink_path(path, "artifact")
    path = path.resolve()
    if not path.is_file():
        raise ValueError(f"artifact is not a file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def canonical(value: object) -> bytes:
    return (
        json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))
        + "\n"
    ).encode()


def load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ValueError(f"{label} is invalid JSON: {error}") from error
    if not isinstance(value, dict):
        raise ValueError(f"{label} is not a JSON object")
    return value


def recompute_heap_metrics(print_path: Path, histogram_path: Path) -> dict[str, Any]:
    global _REVIEWED_PARSER_MODULE
    parser_path = REPOSITORY / "tools/perf_resource_profile.py"
    if _REVIEWED_PARSER_MODULE is None:
        spec = importlib.util.spec_from_file_location("reviewed_perf_resource_profile", parser_path)
        if spec is None or spec.loader is None:
            raise ValueError("reviewed memory parser cannot be loaded")
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        _REVIEWED_PARSER_MODULE = module
    module = _REVIEWED_PARSER_MODULE
    parsed = module.parse_heaptrack_print(print_path, retained=False)
    parsed["allocated_bytes"] = module.parse_heaptrack_histogram(histogram_path)
    return parsed


def checked_hash(value: object, label: str) -> str:
    if not isinstance(value, str) or SHA256.fullmatch(value) is None:
        raise ValueError(f"{label} is not a SHA-256")
    return value


def checked_artifact(row: object, label: str) -> dict[str, object]:
    if not isinstance(row, dict):
        raise ValueError(f"{label} is missing")
    path = row.get("path")
    if not isinstance(path, str) or not path:
        raise ValueError(f"{label}.path is missing")
    raw_path = Path(path)
    if raw_path.is_symlink():
        raise ValueError(f"{label}.path may not be a symlink")
    expected = {
        "path": str(Path(path).resolve()),
        "bytes": row.get("bytes"),
        "sha256": checked_hash(row.get("sha256"), f"{label}.sha256"),
    }
    actual = artifact(Path(path))
    if actual != expected:
        raise ValueError(f"{label} changed: expected {expected}, found {actual}")
    return actual


def require_external(path: Path, label: str) -> None:
    path = Path(path)
    reject_symlink_path(path, label)
    resolved = path.resolve()
    if resolved == REPOSITORY or REPOSITORY in resolved.parents:
        raise ValueError(f"{label} must be outside this checkout")


def checked_relative_path(value: object, label: str) -> str:
    """Validate a fixture path before either consumer joins it to its root."""
    if not isinstance(value, str) or not value:
        raise ValueError(f"{label} is missing")
    path = Path(value)
    if path.is_absolute() or "\\" in value or "3rdparty" in path.parts:
        raise ValueError(f"{label} must be a portable relative path")
    if any(part in {"", ".", ".."} for part in path.parts):
        raise ValueError(f"{label} contains traversal or an unnormalised component")
    return value


def memory_row_label(fixture: str, route: str, calls: int) -> str:
    return f"{fixture.replace('/', '_')}-{route.replace('/', '_')}-{calls}"


def load_fixture_manifest(value: dict[str, Any], label: str = "fixtures") -> dict[str, Any]:
    if value.get("schema") != "litchi_signed_fixture_manifest.v1":
        raise ValueError(f"{label} schema is invalid")
    fixtures = value.get("fixtures")
    if not isinstance(fixtures, list) or not fixtures:
        raise ValueError(f"{label}.fixtures is empty")
    labels: set[str] = set()
    formats: set[str] = set()
    for index, fixture in enumerate(fixtures):
        if not isinstance(fixture, dict):
            raise ValueError(f"{label}.fixtures[{index}] is not an object")
        name = fixture.get("name")
        format_name = fixture.get("format")
        origin = fixture.get("origin")
        mode = fixture.get("fixture_mode")
        if not isinstance(name, str) or not name or name in labels:
            raise ValueError(f"{label}.fixtures[{index}] has a duplicate name")
        if not isinstance(format_name, str) or not format_name:
            raise ValueError(f"{label}.fixtures[{index}] has no format")
        checked_relative_path(fixture.get("relative_path"), f"{label}.fixtures[{index}].relative_path")
        if origin not in ORIGINS or mode not in {"producer_signed", "synthetic_presence"}:
            raise ValueError(f"{label}.fixtures[{index}] has an invalid origin/mode")
        if fixture.get("signature_presence") != "detected":
            raise ValueError(f"{label}.fixtures[{index}] does not prove detection")
        if fixture.get("cryptographic_verification") != CRYPTO_STATUS:
            raise ValueError(f"{label}.fixtures[{index}] makes a crypto claim")
        checked_hash(fixture.get("input_sha256"), f"{label}.fixtures[{index}].input_sha256")
        if mode == "producer_signed":
            checked_hash(
                fixture.get("signed_sha256"),
                f"{label}.fixtures[{index}].signed_sha256",
            )
        elif fixture.get("signed_sha256") is not None:
            checked_hash(
                fixture.get("signed_sha256"),
                f"{label}.fixtures[{index}].signed_sha256",
            )
        elif fixture.get("signed_hash_mode") != "probe_generated":
            raise ValueError(
                f"{label}.fixtures[{index}] lacks a deterministic signed hash policy"
            )
        target = fixture.get("target")
        property_name = fixture.get("property")
        if (
            not isinstance(target, str)
            or not target
            or not isinstance(property_name, str)
            or not property_name
        ):
            raise ValueError(f"{label}.fixtures[{index}] native target is incomplete")
        labels.add(name)
        formats.add(format_name)
    if formats != set(NATIVE_FORMATS):
        raise ValueError(
            f"{label} formats are {sorted(formats)}, expected {sorted(NATIVE_FORMATS)}"
        )
    return value


def _reviewed_w8_projection_pair() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    """Load the fixed W8 graph rule; callers cannot substitute a policy file."""
    if not REVIEWED_W8_POLICY.is_file():
        raise ValueError(f"reviewed W8 projection policy is missing: {REVIEWED_W8_POLICY}")
    spec = importlib.util.spec_from_file_location("reviewed_w8_projection_policy", REVIEWED_W8_POLICY)
    if spec is None or spec.loader is None:
        raise ValueError("reviewed W8 projection policy cannot be loaded")
    module = importlib.util.module_from_spec(spec)
    inserted = False
    if str(REPOSITORY) not in sys.path:
        sys.path.insert(0, str(REPOSITORY))
        inserted = True
    try:
        spec.loader.exec_module(module)
        try:
            module.load_reviewed_policy(REVIEWED_W8_POLICY_JSON)
        except Exception as error:
            raise ValueError(f"reviewed W8 policy bytes failed: {error}") from error
        before = load_json(REVIEWED_W8_BEFORE_PROJECTION, "reviewed W8 before projection")
        after = load_json(REVIEWED_W8_AFTER_PROJECTION, "reviewed W8 after projection")
        try:
            evidence = module.validate_projection_pair(before, after)
        except Exception as error:
            raise ValueError(f"reviewed W8 projection pair failed: {error}") from error
        if not isinstance(evidence, dict):
            raise ValueError("reviewed W8 projection policy returned no evidence")
        return before, after, evidence
    finally:
        if inserted:
            sys.path.remove(str(REPOSITORY))


def _projection_closure(value: dict[str, Any], label: str) -> dict[str, Any]:
    lock = value.get("lock")
    packages = lock.get("packages") if isinstance(lock, dict) else None
    if not isinstance(packages, list) or not packages:
        raise ValueError(f"{label} reviewed projection has no lock closure")
    registry: list[dict[str, Any]] = []
    local: list[dict[str, Any]] = []
    for index, package in enumerate(packages):
        if (
            not isinstance(package, dict)
            or not isinstance(package.get("name"), str)
            or not isinstance(package.get("version"), str)
            or not isinstance(package.get("dependencies"), list)
            or not all(isinstance(item, str) for item in package["dependencies"])
        ):
            raise ValueError(f"{label} reviewed package {index} is incomplete")
        source = package.get("source")
        identity = {
            "name": package["name"],
            "version": package["version"],
            "source": source,
            "checksum": package.get("checksum"),
            "dependencies": sorted(package["dependencies"]),
        }
        if isinstance(source, str) and source.startswith("registry+"):
            if SHA256.fullmatch(str(package.get("checksum", ""))) is None:
                raise ValueError(f"{label} reviewed registry checksum is invalid")
            registry.append(identity)
        elif source is None:
            local.append(identity)
        else:
            raise ValueError(f"{label} reviewed package has an unsupported source")
    return {
        "registry": sorted(registry, key=lambda item: (item["name"], item["version"], item["source"])),
        "local": sorted(local, key=lambda item: (item["name"], item["version"])),
    }


def _registry_closure(endpoint: dict[str, Any], label: str) -> dict[str, Any]:
    lock_path = Path(endpoint["lock"]["path"])
    try:
        lock_value = tomllib.loads(lock_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, tomllib.TOMLDecodeError) as error:
        raise ValueError(f"{label}.lock is not valid TOML") from error
    packages = lock_value.get("package")
    if not isinstance(packages, list) or not packages:
        raise ValueError(f"{label}.lock has no package closure")
    registry = []
    local = []
    identities: set[tuple[str, str, object, object]] = set()
    for package in packages:
        if not isinstance(package, dict) or not isinstance(package.get("name"), str) or not isinstance(package.get("version"), str):
            raise ValueError(f"{label}.lock package identity is incomplete")
        source = package.get("source")
        dependencies = package.get("dependencies", [])
        if not isinstance(dependencies, list) or not all(isinstance(item, str) for item in dependencies):
            raise ValueError(f"{label}.lock package dependencies are incomplete")
        identity = {
            "name": package["name"],
            "version": package["version"],
            "source": source,
            "checksum": package.get("checksum"),
            "dependencies": sorted(dependencies),
        }
        identity_key = (
            identity["name"],
            identity["version"],
            identity["source"],
            identity["checksum"],
        )
        if identity_key in identities:
            raise ValueError(f"{label}.lock contains a duplicate package identity")
        identities.add(identity_key)
        if isinstance(source, str) and source.startswith("registry+"):
            if not isinstance(package.get("checksum"), str) or SHA256.fullmatch(package["checksum"]) is None:
                raise ValueError(f"{label}.lock registry checksum is invalid")
            registry.append(identity)
        else:
            local.append(identity)
    registry.sort(key=lambda item: (item["name"], item["version"], item["source"]))
    local.sort(key=lambda item: (item["name"], item["version"], str(item["source"])))
    return {"registry": registry, "local": local}


def _shared_registry(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    left = _registry_closure(before, "before")
    right = _registry_closure(after, "after")
    reviewed_before, reviewed_after, policy_evidence = _reviewed_w8_projection_pair()
    reviewed_left = _projection_closure(reviewed_before, "before")
    reviewed_right = _projection_closure(reviewed_after, "after")
    if left["registry"] != right["registry"]:
        raise ValueError("before/after registry identities, checksums or dependency edges differ")
    reviewed_registry_left = {
        (item["name"], item["version"], item["source"], item["checksum"]): item
        for item in reviewed_left["registry"]
    }
    reviewed_registry_right = {
        (item["name"], item["version"], item["source"], item["checksum"]): item
        for item in reviewed_right["registry"]
    }
    for side, closure, reviewed in (
        ("before", left, reviewed_registry_left),
        ("after", right, reviewed_registry_right),
    ):
        for package in closure["registry"]:
            key = (package["name"], package["version"], package["source"], package["checksum"])
            if key not in reviewed or package["dependencies"] != reviewed[key]["dependencies"]:
                raise ValueError(f"generated {side} consumer lock is outside the reviewed registry closure")
    allowed_local = {"signed-native-consumer"}
    left_consumer = [item for item in left["local"] if item["name"] in allowed_local]
    right_consumer = [item for item in right["local"] if item["name"] in allowed_local]
    if len(left_consumer) != 1 or len(right_consumer) != 1:
        raise ValueError("generated consumer lock must contain exactly one signed-native-consumer package")
    if (
        left_consumer[0]["version"] != right_consumer[0]["version"]
        or left_consumer[0]["dependencies"] != right_consumer[0]["dependencies"]
    ):
        raise ValueError("consumer adapter dependency edge differs between role locks")
    left_local = {
        (item["name"], item["version"]): item
        for item in left["local"]
        if item["name"] not in allowed_local
    }
    right_local = {
        (item["name"], item["version"]): item
        for item in right["local"]
        if item["name"] not in allowed_local
    }
    expected_left = {(item["name"], item["version"]): item for item in reviewed_left["local"]}
    expected_right = {(item["name"], item["version"]): item for item in reviewed_right["local"]}
    if set(left_local) != set(right_local):
        raise ValueError("generated consumer local selected closure differs between roles")
    for side, closure, expected in (
        ("before", left_local, expected_left),
        ("after", right_local, expected_right),
    ):
        for key, package in closure.items():
            if key not in expected or package != expected[key]:
                raise ValueError(f"generated {side} consumer local closure is outside the reviewed role projection")
    expected_deltas = {
        (
            tuple(item["package"]),
            tuple(sorted(item["before"])),
            tuple(sorted(item["after"])),
        )
        for item in policy_evidence.get("local_edge_differences", [])
    }
    actual_deltas = {
        (
            (key[0], key[1]),
            tuple(left_local[key]["dependencies"]),
            tuple(right_local[key]["dependencies"]),
        )
        for key in left_local
        if left_local[key]["dependencies"] != right_local.get(key, {}).get("dependencies")
    }
    if not actual_deltas.issubset(expected_deltas):
        raise ValueError("generated consumer local closure changed outside the reviewed W8 edge delta")
    return {
        "registry_package_count": len(left["registry"]),
        "reviewed_registry_package_count": len(reviewed_left["registry"]),
        "registry_identities_checksums_dependencies_equal": True,
        "reviewed_policy": "tools/w8_projection_policy.py fixed original-to-final pair",
        "reviewed_local_edge_differences": policy_evidence.get("local_edge_differences", []),
        "allowed_local_deltas": sorted(allowed_local),
    }


def _capability_receipt_path(value: object, label: str) -> dict[str, object]:
    if not isinstance(value, dict):
        raise ValueError(f"{label} must be a hash-bound artifact")
    return checked_artifact(value, label)


def _validate_capabilities_receipt(
    endpoint: dict[str, Any], capabilities_artifact: dict[str, object], label: str
) -> None:
    require_external(Path(str(capabilities_artifact["path"])), f"{label} capability receipt")
    capabilities_path = Path(str(capabilities_artifact["path"]))
    capabilities = load_json(capabilities_path, f"{label}.capabilities_receipt")
    if canonical(capabilities) != capabilities_path.read_bytes():
        raise ValueError(f"{label}.capabilities_receipt is not canonical JSON")
    if capabilities.get("schema") != "litchi_signed_native_capabilities.v1":
        raise ValueError(f"{label}.capabilities_receipt schema is invalid")
    if capabilities.get("source_commit") != endpoint["source_commit"]:
        raise ValueError(f"{label}.capabilities_receipt source commit differs")
    if capabilities.get("binary") != endpoint["binary"]:
        raise ValueError(f"{label}.capabilities_receipt binary differs")
    if capabilities.get("routes") != endpoint["capabilities"]:
        raise ValueError(f"{label}.capabilities_receipt route statuses differ")
    evidence = capabilities.get("evidence")
    if not isinstance(evidence, dict) or set(evidence) != set(ROUTES):
        raise ValueError(f"{label}.capabilities_receipt evidence is incomplete")
    cpu = capabilities.get("cpu")
    if not isinstance(cpu, int):
        raise ValueError(f"{label}.capabilities_receipt CPU binding is missing")
    endpoint_kind = "before" if endpoint["source_commit"] == BASELINE_SOURCE_COMMIT else "final"
    for route in ROUTES:
        child_artifact = _capability_receipt_path(evidence[route], f"{label}.{route}.receipt")
        require_external(Path(str(child_artifact["path"])), f"{label}.{route}.receipt")
        child_path = Path(str(child_artifact["path"]))
        result = load_json(child_path, f"{label}.{route}.receipt")
        stdout = result.get("stdout")
        if not isinstance(stdout, dict):
            raise ValueError(f"{label}.{route}: capability stdout binding is missing")
        stdout_path = Path(str(checked_artifact(stdout, f"{label}.{route}.stdout")["path"]))
        child_value = load_json(stdout_path, f"{label}.{route}.stdout")
        child_value["_receipt"] = str(child_path)
        child_value["_receipt_artifact"] = child_artifact
        validate_child_receipt(
            child_value,
            expected_route=route,
            expected_mode="validate",
            expected_calls=None,
            expected_binary=endpoint["binary"],
            expected_fixture_root=endpoint["fixture_root"],
            expected_fixture_manifest=endpoint["fixtures"],
            expected_cpu=cpu,
            output_path=child_path.parent,
            expected_endpoint=endpoint_kind,
        )


def validate_endpoint_manifest(path: Path, label: str) -> dict[str, Any]:
    path = Path(path)
    reject_symlink_path(path, label)
    if not path.is_file():
        raise ValueError(f"{label} is not a regular endpoint manifest")
    path = path.resolve()
    require_external(path, label)
    value = load_json(path, label)
    if canonical(value) != path.read_bytes():
        raise ValueError(f"{label} is not canonical JSON")
    if value.get("schema") != ENDPOINT_SCHEMA:
        raise ValueError(f"{label} schema is invalid")
    commit = value.get("source_commit")
    if not isinstance(commit, str) or FULL_COMMIT.fullmatch(commit) is None:
        raise ValueError(f"{label}.source_commit is not a full commit")
    for name in (
        "binary",
        "source",
        "source_archive",
        "lock",
        "probe",
        "fixtures",
        "dependency_projection",
        "build_receipt",
        "capabilities_receipt",
    ):
        checked = checked_artifact(value.get(name), f"{label}.{name}")
        if name != "fixtures":
            require_external(Path(str(checked["path"])), f"{label}.{name}")
    if not os.access(Path(value["binary"]["path"]), os.X_OK):
        raise ValueError(f"{label}.binary is not executable")
    identity = value.get("source_identity")
    if not isinstance(identity, dict):
        raise ValueError(f"{label}.source_identity is missing")
    if identity.get("commit") != commit:
        raise ValueError(f"{label}.source_identity.commit does not match source_commit")
    if not isinstance(identity.get("source_ref"), str) or not identity["source_ref"]:
        raise ValueError(f"{label}.source_identity.source_ref is missing")
    if not isinstance(identity.get("source_capture"), str) or not identity["source_capture"]:
        raise ValueError(f"{label}.source_identity.source_capture is missing")
    source_root = identity.get("source_root")
    if (
        not isinstance(source_root, str)
        or not Path(source_root).is_absolute()
        or not Path(source_root).is_dir()
        or Path(source_root).is_symlink()
    ):
        raise ValueError(f"{label}.source_identity.source_root must be an existing absolute directory")
    require_external(Path(source_root), f"{label}.source_identity.source_root")
    checked_hash(identity.get("tree_sha256"), f"{label}.source_identity.tree_sha256")
    archive = checked_artifact(identity.get("archive"), f"{label}.source_identity.archive")
    if archive != value["source_archive"]:
        raise ValueError(f"{label}.source_identity.archive is not the endpoint source archive")
    adapter = checked_artifact(identity.get("adapter"), f"{label}.source_identity.adapter")
    if not adapter["path"].endswith(".rs"):
        raise ValueError(f"{label}.source_identity.adapter is not a Rust consumer")
    source_path = Path(value["source"]["path"])
    source_record = load_json(source_path, f"{label}.source")
    if canonical(source_record) != source_path.read_bytes():
        raise ValueError(f"{label}.source is not canonical JSON")
    source_capture_path = Path(identity["source_capture"])
    reject_symlink_path(source_capture_path, f"{label}.source_identity.source_capture")
    source_artifact_path = Path(value["source"]["path"])
    reject_symlink_path(source_artifact_path, f"{label}.source")
    if source_capture_path.resolve() != source_artifact_path.resolve():
        raise ValueError(f"{label}.source_identity.source_capture is not the source artifact")
    for field in ("commit", "source_ref", "source_capture", "source_root", "tree_sha256", "archive", "adapter"):
        source_field = "source_archive" if field == "archive" else field
        if source_record.get(source_field) != identity.get(field):
            raise ValueError(f"{label}.source does not bind source_identity.{field}")
    build_path = Path(value["build_receipt"]["path"])
    build_receipt = load_json(build_path, f"{label}.build_receipt")
    if canonical(build_receipt) != build_path.read_bytes():
        raise ValueError(f"{label}.build_receipt is not canonical JSON")
    if build_receipt.get("source_commit") != commit:
        raise ValueError(f"{label}.build_receipt source commit differs")
    if build_receipt.get("binary_sha256") != value["binary"]["sha256"]:
        raise ValueError(f"{label}.build_receipt binary differs")
    if build_receipt.get("source_archive_sha256") != value["source_archive"]["sha256"]:
        raise ValueError(f"{label}.build_receipt source archive differs")
    if build_receipt.get("adapter_sha256") != adapter["sha256"]:
        raise ValueError(f"{label}.build_receipt adapter differs")
    if build_receipt.get("lock_sha256") != value["lock"]["sha256"]:
        raise ValueError(f"{label}.build_receipt lock differs")
    for field in ("source_archive", "adapter", "lock", "binary", "probe", "fixtures", "dependency_projection"):
        if build_receipt.get(field) != value[field]:
            raise ValueError(f"{label}.build_receipt {field} differs")
    projection_value = load_json(Path(value["dependency_projection"]["path"]), f"{label}.dependency_projection")
    if canonical(projection_value) != Path(value["dependency_projection"]["path"]).read_bytes():
        raise ValueError(f"{label}.dependency_projection is not canonical JSON")
    if not isinstance(projection_value.get("packages"), list) or not projection_value["packages"]:
        raise ValueError(f"{label}.dependency_projection has no generated package closure")
    projection_identities: set[tuple[str, str, str]] = set()
    for index, package in enumerate(projection_value["packages"]):
        if (
            not isinstance(package, dict)
            or not isinstance(package.get("name"), str)
            or not isinstance(package.get("version"), str)
            or not isinstance(package.get("dependencies"), list)
        ):
            raise ValueError(f"{label}.dependency_projection package {index} is incomplete")
        source = package.get("source")
        source_kind = "registry" if isinstance(source, str) and source.startswith("registry+") else "local"
        identity = (package["name"], package["version"], source_kind)
        if identity in projection_identities:
            raise ValueError(f"{label}.dependency_projection has a duplicate package identity")
        projection_identities.add(identity)
    lock_value = tomllib.loads(Path(value["lock"]["path"]).read_text(encoding="utf-8"))
    lock_packages = lock_value.get("package")
    if not isinstance(lock_packages, list):
        raise ValueError(f"{label}.lock has no package list")
    lock_identities = {
        (
            package.get("name"),
            package.get("version"),
            "registry" if isinstance(package.get("source"), str) and package["source"].startswith("registry+") else "local",
        )
        for package in lock_packages
        if isinstance(package, dict)
    }
    if projection_identities != lock_identities:
        raise ValueError(f"{label}.dependency_projection is not the generated lock closure")
    fixtures_path = Path(value["fixtures"]["path"])
    fixture_value = load_json(fixtures_path, f"{label}.fixtures")
    if canonical(fixture_value) != fixtures_path.read_bytes():
        raise ValueError(f"{label}.fixtures is not canonical JSON")
    load_fixture_manifest(fixture_value, f"{label}.fixtures")
    capabilities = value.get("capabilities")
    if not isinstance(capabilities, dict):
        raise ValueError(f"{label}.capabilities is missing")
    if set(capabilities) != set(ROUTES):
        raise ValueError(f"{label}.capabilities does not enumerate every signed route")
    for route, status in capabilities.items():
        if status not in {
            "available",
            "capability_missing",
            "expected_refusal",
            "not_applicable",
        }:
            raise ValueError(f"{label}.capabilities[{route}] has invalid status")
        if status == "capability_missing" and route not in FINAL_ONLY:
            raise ValueError(f"{label}.capabilities[{route}] cannot be missing")
        if status == "not_applicable" and route not in BASELINE_NOT_APPLICABLE:
            raise ValueError(f"{label}.capabilities[{route}] cannot be not_applicable")
        if status == "expected_refusal" and not route.endswith("/refusal"):
            raise ValueError(f"{label}.capabilities[{route}] cannot be expected_refusal")
    root = value.get("fixture_root")
    if not isinstance(root, str) or not Path(root).is_absolute() or not Path(root).is_dir():
        raise ValueError(f"{label}.fixture_root must be an existing absolute directory")
    root_path = Path(root)
    repository = REPOSITORY
    reject_symlink_path(root_path, f"{label}.fixture_root")
    root_resolved = root_path.resolve()
    if root_resolved == repository or repository in root_resolved.parents:
        raise ValueError(f"{label}.fixture_root must be external to this checkout")
    expected_manifest = root_path / "docs/unified-ops/results/0362ul24/fixtures.json"
    reject_symlink_path(fixtures_path, f"{label}.fixtures")
    reject_symlink_path(expected_manifest, f"{label}.expected_fixture_manifest")
    if fixtures_path.resolve() != expected_manifest.resolve():
        raise ValueError(f"{label}.fixtures is not the manifest the Rust probe loads")
    for index, fixture in enumerate(fixture_value["fixtures"]):
        relative = checked_relative_path(fixture.get("relative_path"), f"{label}.fixtures[{index}].relative_path")
        fixture_path = root_path / relative
        reject_symlink_path(fixture_path, f"{label}.fixtures[{index}]")
        if not fixture_path.is_file():
            raise ValueError(f"{label}.fixtures[{index}] is not a regular staged file")
        resolved = fixture_path.resolve()
        if not resolved.is_relative_to(root_resolved):
            raise ValueError(f"{label}.fixtures[{index}] escapes fixture_root")
        if sha256(resolved) != fixture["input_sha256"]:
            raise ValueError(f"{label}.fixtures[{index}] input hash differs")
    _validate_capabilities_receipt(value, checked_artifact(value["capabilities_receipt"], f"{label}.capabilities_receipt"), label)
    return value


def shared_identity(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    fields = ("probe", "fixtures")
    for field in fields:
        left = before[field]["sha256"]
        right = after[field]["sha256"]
        if left != right:
            raise ValueError(f"endpoint {field} differs: {left} != {right}")
    registry = _shared_registry(before, after)
    return {
        **{f"{field}_sha256": before[field]["sha256"] for field in fields},
        "before_lock_sha256": before["lock"]["sha256"],
        "after_lock_sha256": after["lock"]["sha256"],
        "before_dependency_projection_sha256": before["dependency_projection"]["sha256"],
        "after_dependency_projection_sha256": after["dependency_projection"]["sha256"],
        "registry": registry,
    }


def validate_pair(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    if before["source_commit"] != BASELINE_SOURCE_COMMIT:
        raise ValueError(
            f"before endpoint is not the pinned f2 baseline: {before['source_commit']}"
        )
    if after["source_commit"] != CANDIDATE_SOURCE_COMMIT:
        raise ValueError(
            f"after endpoint is not the pinned 472b capture: {after['source_commit']}"
        )
    if before["binary"]["path"] == after["binary"]["path"]:
        raise ValueError("before and after endpoint binaries must have distinct paths")
    if before["binary"]["sha256"] == after["binary"]["sha256"]:
        raise ValueError("before and after endpoint binaries must differ")
    if before["source_archive"]["sha256"] == after["source_archive"]["sha256"]:
        raise ValueError("before and after source archives must differ")
    if before["source_identity"]["tree_sha256"] == after["source_identity"]["tree_sha256"]:
        raise ValueError("before and after source trees must differ")
    shared = shared_identity(before, after)
    for route in ROUTES:
        final_status = after["capabilities"][route]
        before_status = before["capabilities"][route]
        if final_status == "capability_missing":
            raise ValueError(f"final endpoint lacks required route {route}")
        if route in BASELINE_NOT_APPLICABLE and before_status != "not_applicable":
            raise ValueError(f"baseline route {route} must be explicitly not_applicable")
        if route.endswith("/refusal") and final_status != "expected_refusal":
            raise ValueError(f"final refusal route {route} is not marked expected_refusal")
        if route not in CORRECTNESS_ONLY and final_status != "available":
            raise ValueError(f"final timed route {route} is not available")
    return {
        "shared": shared,
        "baseline_capability_gaps": [
            route for route in ROUTES if before["capabilities"][route] != "available"
        ],
    }


def _fixture_rows(value: dict[str, Any], route: str) -> list[dict[str, Any]]:
    if value.get("schema") != SCHEMA or value.get("route") != route:
        raise ValueError(f"{route}: probe schema or route identity is invalid")
    rows = value.get("fixtures")
    if not isinstance(rows, list) or not rows:
        raise ValueError(f"{route}: fixture results are empty")
    return rows


def _common_fixture_checks(row: dict[str, Any], route: str) -> None:
    for key in ("format", "name", "origin"):
        if not isinstance(row.get(key), str) or not row[key]:
            raise ValueError(f"{route}: fixture {key} is missing")
    if row.get("origin") not in ORIGINS:
        raise ValueError(f"{route}: unknown signature origin")
    if row.get("signature_presence") != "detected":
        raise ValueError(f"{route}: signature presence was not detected")
    if row.get("cryptographic_verification") != CRYPTO_STATUS:
        raise ValueError(f"{route}: cryptographic status is not explicit")
    checked_hash(row.get("source_sha256"), f"{route}: source")


def _refusal_publication_checks(row: dict[str, Any], route: str) -> None:
    if row.get("output_published") is not False or row.get("publication_attempts") != 0:
        raise ValueError(f"{route}: refusal crossed a publication boundary")
    if row.get("source_before_sha256") != row.get("source_sha256"):
        raise ValueError(f"{route}: refusal source-before hash differs")


def validate_route(
    value: dict[str, Any], route: str, *, endpoint: str = "final"
) -> dict[str, Any]:
    rows = _fixture_rows(value, route)
    capability = value.get("capability")
    if capability not in {
        "available",
        "capability_missing",
        "expected_refusal",
        "not_applicable",
    }:
        raise ValueError(f"{route}: invalid capability status")
    if capability == "not_applicable":
        if endpoint != "before" or route not in BASELINE_NOT_APPLICABLE:
            raise ValueError(f"{route}: not_applicable is only a baseline disposition")
        for row in rows:
            _common_fixture_checks(row, route)
            if (
                row.get("status") != "not_applicable"
                or row.get("disposition") != "requires_signed_noop_admission"
                or row.get("reason")
                != "the historical endpoint cannot construct a signed no-op base"
                or row.get("signed_before") is not True
            ):
                raise ValueError(f"{route}: baseline not_applicable disposition is incomplete")
        return {"route": route, "status": "not_applicable", "endpoint": endpoint}
    if capability == "capability_missing":
        if endpoint == "final" or route not in FINAL_ONLY:
            raise ValueError(f"{route}: required capability missing on {endpoint}")
        for row in rows:
            _common_fixture_checks(row, route)
            if (
                row.get("status") != "capability_missing"
                or row.get("refusal_kind") != "ProtectedDocument"
                or row.get("refusal_code") != "document_protected"
                or row.get("signed_before") is not True
            ):
                raise ValueError(f"{route}: baseline gap lacks its typed refusal proof")
        return {"route": route, "status": "capability_missing", "endpoint": endpoint}
    native = value.get("native_preflight")
    if not isinstance(native, list) or not native:
        raise ValueError(f"{route}: direct native preflight is missing")
    native_formats: set[str] = set()
    for native_row in native:
        if not isinstance(native_row, dict):
            raise ValueError(f"{route}: native preflight row is not an object")
        format_name = native_row.get("format")
        status = native_row.get("status")
        if not isinstance(format_name, str) or status not in {
            "exact_noop",
            "unsupported_signed_source",
        }:
            raise ValueError(f"{route}: invalid native preflight disposition")
        if format_name in native_formats:
            raise ValueError(f"{route}: duplicate native preflight format {format_name}")
        native_formats.add(format_name)
        if format_name in {"DOC", "PPT"} and status != "unsupported_signed_source":
            raise ValueError(f"{route}: {format_name} signed source should refuse direct source edit")
        if format_name not in {"DOC", "PPT"} and status != "exact_noop":
            raise ValueError(f"{route}: {format_name} direct native no-op was not proven")
    if native_formats != NATIVE_FORMATS:
        raise ValueError(
            f"{route}: direct native preflight covers {sorted(native_formats)}, "
            f"expected {sorted(NATIVE_FORMATS)}"
        )
    for row in rows:
        if not isinstance(row, dict):
            raise ValueError(f"{route}: fixture row is not an object")
        _common_fixture_checks(row, route)
        if route.endswith("/refusal"):
            if row.get("status") != "expected_refusal":
                raise ValueError(f"{route}: changed signed edit did not refuse")
            if (
                row.get("refusal_kind") != "ProtectedDocument"
                or row.get("code") != "document_protected"
            ):
                raise ValueError(
                    f"{route}: refusal classification is not protected/document_protected"
                )
            if row.get("source_after_sha256") != row.get("source_sha256"):
                raise ValueError(f"{route}: refusal changed source bytes")
            _refusal_publication_checks(row, route)
            continue
        if route.endswith("/unsigned_negative"):
            if (
                row.get("status") != "expected_refusal"
                or row.get("refusal_kind") != "PatchConflict"
                or row.get("code") != "invalid_input"
            ):
                raise ValueError(
                    f"{route}: unsigned negative is not a typed patch conflict"
                )
            if row.get("re_marked_signed") is not False:
                raise ValueError(f"{route}: unsigned negative re-marked a signature")
            if row.get("recipient_signed_after") is not True:
                raise ValueError(f"{route}: patch conflict did not preserve the signed recipient")
            _refusal_publication_checks(row, route)
            continue
        if route.endswith("/source_freshness"):
            if (
                row.get("status") != "expected_refusal"
                or row.get("refusal_kind") != "SourceChanged"
                or row.get("code") != "io_error"
            ):
                raise ValueError(f"{route}: source fence is not SourceChanged/io_error")
            checked_hash(row.get("source_mutated_sha256"), f"{route}: mutated source")
            if row.get("source_mutated_sha256") == row.get("source_sha256"):
                raise ValueError(f"{route}: source fence did not mutate the source")
            if row.get("source_after_sha256") != row.get("source_mutated_sha256"):
                raise ValueError(f"{route}: source fence lacks its post-mutation source hash")
            _refusal_publication_checks(row, route)
            continue
        if route.endswith("/budget"):
            if (
                row.get("status") != "expected_refusal"
                or row.get("refusal_kind") != "LimitExceeded"
                or row.get("code") != "limit_exceeded"
            ):
                raise ValueError(f"{route}: budget refusal is not typed")
            if row.get("source_after_sha256") != row.get("source_sha256"):
                raise ValueError(f"{route}: budget refusal changed source bytes")
            if row.get("limit_kind") != "operations":
                raise ValueError(f"{route}: budget refusal lacks its typed limit kind")
            _refusal_publication_checks(row, route)
            continue
        if row.get("status") != "ok":
            raise ValueError(f"{route}: available route has non-ok fixture result")
        if row.get("signed_before") is not True:
            raise ValueError(f"{route}: input is not signed")
        if route.endswith("/noop_commit"):
            if row.get("diagnostics_changed") is not False or row.get("patch_noop") is not True:
                raise ValueError(f"{route}: no-op diagnostics/patch proof failed")
            if (
                row.get("signed_after") is not True
                or row.get("output_sha256") != row.get("source_sha256")
            ):
                raise ValueError(f"{route}: signed no-op changed bytes or signature state")
            if row.get("output_equals_source") is not True:
                raise ValueError(f"{route}: signed no-op lacks exact byte proof")
            if row.get("source_after_sha256") != row.get("source_sha256"):
                raise ValueError(f"{route}: signed no-op changed its source endpoint")
        elif route.endswith("/noop_write"):
            if (
                row.get("diagnostics_changed") is not False
                or row.get("patch_noop") is not True
                or row.get("signed_after") is not True
                or row.get("output_sha256") != row.get("source_sha256")
            ):
                raise ValueError(f"{route}: signed publication changed no-op/state proof")
            if (
                row.get("stream_bounded") is not True
                or not isinstance(row.get("max_write_bytes"), int)
                or row["max_write_bytes"] <= 0
            ):
                raise ValueError(f"{route}: publication is not bounded")
            if row.get("output_equals_source") is not True:
                raise ValueError(f"{route}: signed publication lacks exact byte proof")
            if row.get("source_after_sha256") != row.get("source_sha256"):
                raise ValueError(f"{route}: signed publication changed its source endpoint")
        elif route.endswith("/noop_replay"):
            if row.get("patch_noop") is not True or row.get("replay_base_match") is not True:
                raise ValueError(f"{route}: replay lacks content-bound no-op proof")
            if (
                row.get("signed_after") is not True
                or row.get("output_sha256") != row.get("source_sha256")
            ):
                raise ValueError(f"{route}: replay changed signed publication")
            if row.get("output_equals_source") is not True:
                raise ValueError(f"{route}: replay lacks exact byte proof")
        elif route.endswith("/unsign_write"):
            if row.get("signed_after") is not False or row.get("signature_parts_absent") is not True:
                raise ValueError(f"{route}: unsign output retains signature state/parts")
            if row.get("semantic_changed") is not True or row.get("output_equals_source") is not False:
                raise ValueError(f"{route}: unsign route did not publish a real change")
            if row.get("target_changed") is not True:
                raise ValueError(f"{route}: unsign route did not change the native target")
            checked_hash(row.get("output_sha256"), f"{route}: output")
            if row.get("stream_sha256") != row.get("output_sha256"):
                raise ValueError(f"{route}: bounded stream digest differs from output")
            if row.get("stream_bounded") is not True:
                raise ValueError(f"{route}: unsign publication is not bounded")
            if row.get("signed_clone_still_refuses") is not True:
                raise ValueError(f"{route}: signed clone lost protection")
    return {"route": route, "status": "ok", "endpoint": endpoint, "fixtures": len(rows)}


def validate_correctness_receipt(
    value: dict[str, Any], route: str, *, endpoint: str = "final"
) -> dict[str, Any]:
    """Public wrapper used by tests and the guarded runner."""
    return validate_route(value, route, endpoint=endpoint)


def validate_measure(
    value: dict[str, Any], route: str, calls: int, *, expected_route: str | None = None
) -> None:
    actual_route = expected_route or route
    if value.get("schema") != SCHEMA or value.get("route") != actual_route:
        raise ValueError(f"{route}: timed output identity is invalid")
    if value.get("mode") != "measure" or value.get("calls") != calls:
        raise ValueError(f"{route}: timed call count/mode mismatch")
    elapsed = value.get("elapsed_ns")
    if not isinstance(elapsed, int) or elapsed < MIN_SAMPLE_NS:
        raise ValueError(f"{route}: sample is below the {MIN_SAMPLE_NS} ns floor")
    if calls <= 0 or calls > MAX_CALLS:
        raise ValueError(f"{route}: calls exceed the explicit ceiling")
    fixture_count = value.get("fixture_count")
    if not isinstance(fixture_count, int) or fixture_count <= 0:
        raise ValueError(f"{route}: fixture count is not reported")
    if not isinstance(value.get("operations"), int) or value["operations"] != calls * fixture_count:
        raise ValueError(f"{route}: operation count is not exact")


def validate_calibration(value: dict[str, Any], route: str) -> int:
    if (
        value.get("schema") != SCHEMA
        or value.get("route") != route
        or value.get("mode") != "calibrate"
    ):
        raise ValueError(f"{route}: malformed calibration")
    calls = value.get("common_calls")
    elapsed = value.get("elapsed_ns")
    if not isinstance(calls, int) or calls <= 0 or calls > MAX_CALLS:
        raise ValueError(f"{route}: calibration call count is outside the ceiling")
    if not isinstance(elapsed, int) or elapsed < TARGET_NS:
        raise ValueError(f"{route}: calibration did not reach the {TARGET_NS} ns target")
    fixture_count = value.get("fixture_count")
    if not isinstance(fixture_count, int) or fixture_count <= 0:
        raise ValueError(f"{route}: calibration fixture count is not reported")
    if not isinstance(value.get("operations"), int) or value["operations"] != calls * fixture_count:
        raise ValueError(f"{route}: calibration operation count is not exact")
    if value.get("target_ns") != TARGET_NS:
        raise ValueError(f"{route}: calibration target differs from protocol")
    return calls


def validate_child_receipt(
    value: object,
    *,
    expected_route: str,
    expected_mode: str,
    expected_calls: int | None,
    expected_binary: dict[str, object],
    expected_fixture_root: str,
    expected_fixture_manifest: dict[str, object],
    expected_cpu: int,
    output_path: Path,
    expected_endpoint: str = "final",
) -> dict[str, Any]:
    if (
        not isinstance(value, dict)
        or not isinstance(value.get("_receipt"), str)
        or not isinstance(value.get("_receipt_artifact"), dict)
    ):
        raise ValueError(f"{expected_route}: child receipt binding is missing")
    if expected_mode not in {"validate", "measure", "calibrate"}:
        raise ValueError(f"{expected_route}: unsupported timed child mode")
    receipt_raw_path = Path(value["_receipt"])
    reject_symlink_path(receipt_raw_path, f"{expected_route} child receipt")
    if not receipt_raw_path.is_file():
        raise ValueError(f"{expected_route}: child receipt is not a regular file")
    receipt_path = receipt_raw_path.resolve()
    receipt_artifact = checked_artifact(
        value["_receipt_artifact"], f"{expected_route}.receipt"
    )
    if receipt_artifact["path"] != str(receipt_path):
        raise ValueError(f"{expected_route}: child receipt path/hash binding differs")
    output_path = Path(output_path)
    reject_symlink_path(output_path, f"{expected_route} qualification output")
    output_path = output_path.resolve()
    if not receipt_path.is_relative_to(output_path):
        raise ValueError(f"{expected_route}: child receipt escapes qualification output")
    receipt = load_json(receipt_path, f"{expected_route} child receipt")
    if canonical(receipt) != receipt_path.read_bytes():
        raise ValueError(f"{expected_route}: child receipt is not canonical JSON")
    if receipt.get("schema") != "litchi_signed_native_child.v1":
        raise ValueError(f"{expected_route}: child receipt schema is invalid")
    if receipt.get("mode") != expected_mode or receipt.get("route") != expected_route:
        raise ValueError(f"{expected_route}: child receipt identity differs")
    if expected_calls is not None and receipt.get("calls") != expected_calls:
        raise ValueError(f"{expected_route}: child receipt call count differs")
    if expected_calls is None and expected_mode in {"validate", "calibrate"} and receipt.get("calls") is not None:
        raise ValueError(f"{expected_route}: non-measure child unexpectedly has a measure call count")
    if receipt.get("binary") != expected_binary:
        raise ValueError(f"{expected_route}: child receipt binary differs")
    if receipt.get("fixture_root") != expected_fixture_root:
        raise ValueError(f"{expected_route}: child fixture root differs")
    if receipt.get("fixture_manifest") != expected_fixture_manifest:
        raise ValueError(f"{expected_route}: child fixture manifest differs")
    command = receipt.get("command")
    expected_command = [
        "taskset",
        "-c",
        str(expected_cpu),
        str(expected_binary["path"]),
        "--mode",
        expected_mode,
        "--route",
        expected_route,
        "--blocks",
        "1",
    ]
    if expected_mode == "measure":
        if expected_calls is None:
            raise ValueError(f"{expected_route}: measure child call count is missing")
        expected_command.extend(("--calls", str(expected_calls)))
    if (
        not isinstance(command, list)
        or command != expected_command
    ):
        raise ValueError(f"{expected_route}: child argv is not bound to the expected route/mode")
    if receipt.get("exit_code") != 0 or receipt.get("timed_out") is not False:
        raise ValueError(f"{expected_route}: child did not finish successfully")
    if receipt.get("guard_error") is not None or receipt.get("after_guard_error") is not None:
        raise ValueError(f"{expected_route}: child guard reported an error")
    if (
        not isinstance(receipt.get("process_group"), int)
        or receipt["process_group"] <= 0
        or not isinstance(receipt.get("rss_max_bytes"), int)
        or receipt["rss_max_bytes"] < 0
        or receipt["rss_max_bytes"] > CHILD_RSS_LIMIT
        or not isinstance(receipt.get("aggregate_rss_max_bytes"), int)
        or receipt["aggregate_rss_max_bytes"] < 0
        or receipt["aggregate_rss_max_bytes"] > CHILD_RSS_LIMIT
        or receipt["rss_max_bytes"] != receipt["aggregate_rss_max_bytes"]
        or receipt.get("rss_limit_bytes") != CHILD_RSS_LIMIT
        or receipt.get("output_limit_bytes") != CHILD_OUTPUT_LIMIT
        or receipt.get("timeout_seconds") != CHILD_TIMEOUT_S
    ):
        raise ValueError(f"{expected_route}: aggregate process guard is missing")
    failed_after = receipt.get("failed_after")
    if (
        not isinstance(failed_after, dict)
        or failed_after.get("process_group_empty") is not True
        or failed_after.get("guard_error") is not None
        or failed_after.get("after_guard_error") is not None
        or failed_after.get("timed_out") is not False
        or not isinstance(failed_after.get("rss_max_bytes"), int)
        or failed_after["rss_max_bytes"] != receipt.get("rss_max_bytes")
        or failed_after["rss_max_bytes"] > CHILD_RSS_LIMIT
        or failed_after.get("aggregate_rss_max_bytes") != receipt.get("aggregate_rss_max_bytes")
        or failed_after.get("rss_limit_bytes") != CHILD_RSS_LIMIT
        or failed_after.get("output_limit_bytes") != CHILD_OUTPUT_LIMIT
    ):
        raise ValueError(f"{expected_route}: failed-after process-group receipt is incomplete")
    snapshots = receipt.get("host_snapshots")
    if not isinstance(snapshots, list) or len(snapshots) < 2:
        raise ValueError(f"{expected_route}: child host snapshots are incomplete")
    if (
        not isinstance(receipt.get("during_snapshots"), int)
        or receipt["during_snapshots"] < 1
        or len(snapshots) != receipt["during_snapshots"] + 2
    ):
        raise ValueError(f"{expected_route}: child has no during snapshot")
    for index, snapshot in enumerate(snapshots):
        snapshot_artifact = checked_artifact(snapshot, f"{expected_route}.host[{index}]")
        if not Path(str(snapshot_artifact["path"])).is_relative_to(output_path):
            raise ValueError(f"{expected_route}: host snapshot escapes qualification output")
        host = load_json(Path(snapshot_artifact["path"]), f"{expected_route}.host[{index}]")
        if host.get("blockers") != []:
            raise ValueError(f"{expected_route}: host snapshot has blockers")
    outputs: dict[str, dict[str, object]] = {}
    for name in ("stdout", "stderr"):
        child_output = checked_artifact(receipt.get(name), f"{expected_route}.{name}")
        if not Path(str(child_output["path"])).is_relative_to(output_path):
            raise ValueError(f"{expected_route}: child output escapes qualification output")
        outputs[name] = child_output
    failed_after = receipt["failed_after"]
    for key in ("stdout_bytes", "stderr_bytes"):
        output_name = key.removesuffix("_bytes")
        if (
            not isinstance(failed_after.get(key), int)
            or failed_after[key] > CHILD_OUTPUT_LIMIT
            or failed_after[key] != outputs[output_name]["bytes"]
        ):
            raise ValueError(f"{expected_route}: failed-after output cap is missing")
    if failed_after.get("exit_code") != receipt.get("exit_code"):
        raise ValueError(f"{expected_route}: failed-after exit code differs")
    if failed_after.get("aggregate_rss_max_bytes") != receipt.get("aggregate_rss_max_bytes"):
        raise ValueError(f"{expected_route}: failed-after RSS receipt differs")
    stdout_value = load_json(Path(str(outputs["stdout"]["path"])), f"{expected_route}.stdout")
    expected_value = dict(value)
    expected_value.pop("_receipt", None)
    expected_value.pop("_receipt_artifact", None)
    if stdout_value != expected_value:
        raise ValueError(f"{expected_route}: child stdout does not match the bound route result")
    if expected_mode == "measure":
        validate_measure(stdout_value, expected_route, expected_calls or 0)
    elif expected_mode == "calibrate":
        validate_calibration(stdout_value, expected_route)
    else:
        validate_correctness_receipt(stdout_value, expected_route, endpoint=expected_endpoint)
    if receipt_path.stat().st_size <= 0:
        raise ValueError(f"{expected_route}: child receipt is empty")
    return receipt


def validate_abba(manifest: dict[str, Any]) -> None:
    if manifest.get("schema") != "litchi_signed_native_qualification.v1":
        raise ValueError("qualification manifest schema is invalid")
    if manifest.get("status") != "latency_complete":
        raise ValueError("qualification manifest is not a complete latency receipt")
    rounds = manifest.get("rounds")
    if not isinstance(rounds, int) or rounds < MIN_ABBA_ROUNDS:
        raise ValueError("qualification needs at least five ABBA rounds")
    if manifest.get("target_ns") != TARGET_NS or manifest.get("minimum_sample_ns") != MIN_SAMPLE_NS:
        raise ValueError("qualification timing constants differ from the protocol")
    output = manifest.get("output")
    if not isinstance(output, str) or not Path(output).is_absolute() or not Path(output).is_dir():
        raise ValueError("qualification output directory is not bound")
    repository = REPOSITORY
    output_raw_path = Path(output)
    reject_symlink_path(output_raw_path, "qualification output")
    output_path = output_raw_path.resolve()
    if output_path == repository or repository in output_path.parents:
        raise ValueError("qualification output must be external to this checkout")
    provenance = manifest.get("provenance")
    if not isinstance(provenance, dict):
        raise ValueError("qualification provenance is missing")
    pair = provenance.get("pair")
    if not isinstance(pair, dict):
        raise ValueError("qualification endpoint pair provenance is missing")
    before_manifest = checked_artifact(provenance.get("before_manifest"), "qualification.before_manifest")
    after_manifest = checked_artifact(provenance.get("after_manifest"), "qualification.after_manifest")
    require_external(Path(str(before_manifest["path"])), "qualification.before_manifest")
    require_external(Path(str(after_manifest["path"])), "qualification.after_manifest")
    before = validate_endpoint_manifest(Path(before_manifest["path"]), "qualification.before")
    after = validate_endpoint_manifest(Path(after_manifest["path"]), "qualification.after")
    expected_pair = validate_pair(before, after)
    if pair != expected_pair:
        raise ValueError("qualification pair provenance differs from endpoint manifests")
    if provenance.get("before_binary") != before["binary"] or provenance.get("after_binary") != after["binary"]:
        raise ValueError("qualification binary provenance differs from endpoint manifests")
    if provenance.get("before_source_commit") != before["source_commit"] or provenance.get("after_source_commit") != after["source_commit"]:
        raise ValueError("qualification source provenance differs from endpoint manifests")
    if provenance.get("cpu") != manifest.get("cpu") or not isinstance(manifest.get("cpu"), int):
        raise ValueError("qualification CPU provenance is missing")
    if manifest.get("guard") != "reviewed W8 full /proc guard before/during/after every child":
        raise ValueError("qualification guard policy is not the reviewed guard")
    rows = manifest.get("routes")
    if not isinstance(rows, list):
        raise ValueError("qualification route rows are missing")
    if len(rows) != len(TIMED) or {row.get("route") for row in rows if isinstance(row, dict)} != set(TIMED):
        raise ValueError("qualification must contain each timed route exactly once")
    for row in rows:
        if not isinstance(row, dict) or row.get("route") not in TIMED:
            raise ValueError("qualification contains an unknown timed route")
        calls = row.get("calls")
        if not isinstance(calls, int) or calls <= 0 or calls > MAX_CALLS:
            raise ValueError(f"{row.get('route')}: invalid shared call count")
        blocks = row.get("blocks")
        if not isinstance(blocks, list) or len(blocks) < rounds:
            raise ValueError(f"{row.get('route')}: incomplete ABBA blocks")
        route = row["route"]
        if row.get("baseline") == "leaf_facade":
            if route not in FINAL_ONLY:
                raise ValueError(f"{route}: leaf/facade comparison is final-only")
            paired = row.get("paired_route")
            expected_pair = _paired_route(route)
            if paired != expected_pair:
                raise ValueError(f"{route}: paired route is incomplete or incorrect")
            if row.get("comparison") != "final_leaf_vs_facade":
                raise ValueError(f"{route}: final comparison label is missing")
            calibration = row.get("calibration")
            if not isinstance(calibration, dict):
                raise ValueError(f"{route}: leaf/facade calibration is missing")
            for side, side_route in (("leaf", _leaf_route(route)), ("facade", _facade_route(route))):
                value = calibration.get(side)
                if not isinstance(value, dict):
                    raise ValueError(f"{route}: {side} calibration is missing")
                calibrated_calls = validate_calibration(value, side_route)
                if calibrated_calls > calls:
                    raise ValueError(f"{route}: shared calls undercut {side} calibration")
                validate_child_receipt(
                    value,
                    expected_route=side_route,
                    expected_mode="calibrate",
                    expected_calls=None,
                    expected_binary=after["binary"],
                    expected_fixture_root=after["fixture_root"],
                    expected_fixture_manifest=after["fixtures"],
                    expected_cpu=manifest["cpu"],
                    output_path=output_path,
                )
            for block in blocks:
                legs = block.get("legs") if isinstance(block, dict) else None
                if (
                    not isinstance(legs, list)
                    or [leg.get("side") for leg in legs]
                    != ["leaf", "facade", "facade", "leaf"]
                ):
                    raise ValueError(f"{route}: leaf/facade ABBA order is invalid")
                for leg in legs:
                    side = leg.get("side")
                    expected = _leaf_route(route) if side == "leaf" else _facade_route(route)
                    value = leg.get("value")
                    validate_measure(value, route, calls, expected_route=expected)
                    validate_child_receipt(
                        value,
                        expected_route=expected,
                        expected_mode="measure",
                        expected_calls=calls,
                        expected_binary=after["binary"],
                        expected_fixture_root=after["fixture_root"],
                        expected_fixture_manifest=after["fixtures"],
                        expected_cpu=manifest["cpu"],
                        output_path=output_path,
                    )
            leaf = _median_per_call(blocks, "leaf")
            facade = _median_per_call(blocks, "facade")
            _finite_ratio(row.get("leaf_ns_per_call"), leaf, route)
            _finite_ratio(row.get("facade_ns_per_call"), facade, route)
            expected_ratio = facade / leaf
            _finite_ratio(row.get("facade_leaf_ratio"), expected_ratio, route)
            if "after_before_ratio" in row or "historical_ratio" in row:
                raise ValueError(f"{route}: final-only row contains a historical ratio")
            continue
        if row.get("baseline") != "matched":
            raise ValueError(f"{route}: timing baseline is not a truthful comparison")
        if route in FINAL_ONLY:
            raise ValueError(f"{route}: final-only route cannot claim matched baseline")
        calibration = row.get("calibration")
        if not isinstance(calibration, dict):
            raise ValueError(f"{route}: matched calibration is missing")
        for side, side_route, endpoint_value in (
            ("before", route, before),
            ("after", route, after),
        ):
            calibration_value = calibration.get(side)
            if not isinstance(calibration_value, dict):
                raise ValueError(f"{route}: {side} calibration is missing")
            validate_calibration(calibration_value, side_route)
            validate_child_receipt(
                calibration_value,
                expected_route=side_route,
                expected_mode="calibrate",
                expected_calls=None,
                expected_binary=endpoint_value["binary"],
                expected_fixture_root=endpoint_value["fixture_root"],
                expected_fixture_manifest=endpoint_value["fixtures"],
                expected_cpu=manifest["cpu"],
                output_path=output_path,
            )
        for block in blocks:
            legs = block.get("legs") if isinstance(block, dict) else None
            if (
                not isinstance(legs, list)
                or [leg.get("side") for leg in legs]
                != ["before", "after", "after", "before"]
            ):
                raise ValueError(f"{route}: ABBA order is invalid")
            for leg in legs:
                value = leg.get("value")
                validate_measure(value, route, calls)
                expected_binary = before["binary"] if leg.get("side") == "before" else after["binary"]
                validate_child_receipt(
                    value,
                    expected_route=route,
                    expected_mode="measure",
                    expected_calls=calls,
                    expected_binary=expected_binary,
                    expected_fixture_root=before["fixture_root"] if leg.get("side") == "before" else after["fixture_root"],
                    expected_fixture_manifest=before["fixtures"] if leg.get("side") == "before" else after["fixtures"],
                    expected_cpu=manifest["cpu"],
                    output_path=output_path,
                )
        before = _median_per_call(blocks, "before")
        after = _median_per_call(blocks, "after")
        _finite_ratio(row.get("before_ns_per_call"), before, route)
        _finite_ratio(row.get("after_ns_per_call"), after, route)
        _finite_ratio(row.get("after_before_ratio"), after / before, route)


def _paired_route(route: str) -> str:
    if route.startswith("signed_leaf/"):
        return "signed_facade/" + route.split("/", 1)[1]
    if route.startswith("signed_facade/"):
        return "signed_leaf/" + route.split("/", 1)[1]
    raise ValueError(f"invalid signed route {route}")


def _leaf_route(route: str) -> str:
    return route if route.startswith("signed_leaf/") else _paired_route(route)


def _facade_route(route: str) -> str:
    return route if route.startswith("signed_facade/") else _paired_route(route)


def _median_per_call(blocks: list[dict[str, Any]], side: str) -> float:
    values = [
        leg["value"]["elapsed_ns"] / leg["value"]["calls"]
        for block in blocks
        for leg in block["legs"]
        if leg.get("side") == side
    ]
    if not values:
        raise ValueError(f"no {side} measurements")
    return float(__import__("statistics").median(values))


def _finite_ratio(value: object, expected: float, route: str) -> None:
    if not isinstance(value, (int, float)) or not math.isfinite(float(value)):
        raise ValueError(f"{route}: ratio/median is not finite")
    if expected < 0 or not math.isclose(float(value), expected, rel_tol=1e-9, abs_tol=1e-9):
        raise ValueError(f"{route}: ratio/median does not match ABBA measurements")


def validate_memory_summary(
    value: dict[str, Any],
    endpoint: dict[str, Any],
    fixture_names: set[str],
    routes: tuple[str, ...],
    expected_manifest_artifact: dict[str, object] | None = None,
) -> None:
    if value.get("schema") != "litchi_signed_native_memory.v1":
        raise ValueError("memory summary schema is invalid")
    if value.get("status") != "complete":
        raise ValueError("memory summary is not complete")
    if value.get("call_counts") != [1, 8, 128]:
        raise ValueError("memory summary must contain the paired 1/8/128 call counts")
    if value.get("scope") != "whole_process_including_fixture_open_and_setup; profiler overhead included":
        raise ValueError("memory summary scope is not whole-process")
    if set(routes) != set(FINAL_ONLY):
        raise ValueError("memory summary must cover every final signed timed route")
    raw_root = value.get("raw_root")
    if not isinstance(raw_root, str) or not Path(raw_root).is_absolute():
        raise ValueError("memory raw_root must be absolute")
    repository = REPOSITORY
    raw_root_path = Path(raw_root)
    reject_symlink_path(raw_root_path, "memory raw artifacts")
    raw_path = raw_root_path.resolve()
    if not raw_path.is_dir() or raw_path == repository or repository in raw_path.parents:
        raise ValueError("memory raw artifacts must be outside the repository")
    rows = value.get("rows")
    if not isinstance(rows, list):
        raise ValueError("memory rows are missing")
    source = value.get("source")
    if not isinstance(source, dict) or source.get("shared") is None:
        raise ValueError("memory source pair provenance is missing")
    endpoint_value = value.get("endpoint")
    if not isinstance(endpoint_value, dict):
        raise ValueError("memory endpoint provenance is missing")
    manifest_artifact = checked_artifact(endpoint_value.get("manifest_artifact"), "memory endpoint manifest")
    require_external(Path(str(manifest_artifact["path"])), "memory endpoint manifest")
    if expected_manifest_artifact is None:
        raise ValueError("memory endpoint manifest must be bound to the validated endpoint path")
    expected_manifest = checked_artifact(expected_manifest_artifact, "validated endpoint manifest")
    if manifest_artifact != expected_manifest:
        raise ValueError("memory endpoint manifest differs from the validated endpoint")
    if endpoint_value.get("manifest") != expected_manifest["path"]:
        raise ValueError("memory endpoint manifest path differs from its artifact")
    if endpoint_value.get("source_commit") != CANDIDATE_SOURCE_COMMIT:
        raise ValueError("memory endpoint source commit is not the final capture")
    if endpoint_value.get("binary") != endpoint["binary"]:
        raise ValueError("memory endpoint binary differs")
    if endpoint_value.get("fixture_manifest") != endpoint["fixtures"]:
        raise ValueError("memory endpoint fixture manifest differs")
    if endpoint_value.get("build_receipt") != endpoint["build_receipt"]:
        raise ValueError("memory endpoint build receipt differs")
    tooling = value.get("tooling")
    if not isinstance(tooling, dict) or not isinstance(tooling.get("heaptrack"), str) or not isinstance(tooling.get("heaptrack_print"), str):
        raise ValueError("memory profiler tool provenance is missing")
    parser_artifact = checked_artifact(tooling.get("parser"), "memory parser")
    reviewed_parser = (REPOSITORY / "tools/perf_resource_profile.py").resolve()
    if (
        Path(str(parser_artifact["path"])) != reviewed_parser
        or parser_artifact["sha256"] != REVIEWED_PARSER_SHA256
    ):
        raise ValueError("memory parser is not the pinned reviewed resource parser")
    expected = {(name, route, calls) for name in fixture_names for route in routes for calls in (1, 8, 128)}
    observed: set[tuple[str, str, int]] = set()
    retained_paths: set[str] = set()
    for index, row in enumerate(rows):
        if not isinstance(row, dict):
            raise ValueError(f"memory row {index} is not an object")
        fixture = row.get("fixture")
        route = row.get("route")
        calls = row.get("calls")
        if not isinstance(fixture, str) or not isinstance(route, str) or not isinstance(calls, int):
            raise ValueError(f"memory row {index} has invalid identity fields")
        key = (fixture, route, calls)
        if fixture not in fixture_names or route not in routes or calls not in (1, 8, 128) or key in observed:
            raise ValueError(f"memory row {index} is not an exact fixture/route/call tuple")
        observed.add(key)
        label = row.get("label")
        expected_label = memory_row_label(fixture, route, calls)
        if label != expected_label:
            raise ValueError(f"memory row {index} label is not bound to its fixture/route/call")
        work_path = raw_path / expected_label
        reject_symlink_path(work_path, f"memory row {index} work directory")
        if not work_path.is_dir() or not work_path.resolve().is_relative_to(raw_path):
            raise ValueError(f"memory row {index} work directory is not bound to raw_root")
        if row.get("scope") != value["scope"]:
            raise ValueError(f"memory row {index} has a different scope")
        if row.get("binary") != endpoint["binary"]:
            raise ValueError(f"memory row {index} is bound to a different binary")
        if (
            not isinstance(row.get("rss_max_bytes"), int)
            or row["rss_max_bytes"] < 0
            or row["rss_max_bytes"] > MEMORY_RSS_LIMIT
            or row.get("rss_limit_bytes") != MEMORY_RSS_LIMIT
        ):
            raise ValueError(f"memory row {index} RSS cap receipt is invalid")
        if (
            not isinstance(row.get("process_group"), int)
            or row["process_group"] <= 0
            or not isinstance(row.get("aggregate_rss_max_bytes"), int)
            or row["aggregate_rss_max_bytes"] < 0
            or row["aggregate_rss_max_bytes"] > MEMORY_RSS_LIMIT
            or row["aggregate_rss_max_bytes"] != row["rss_max_bytes"]
        ):
            raise ValueError(f"memory row {index} process guard is incomplete")
        checked_outputs: dict[str, dict[str, object]] = {}
        for artifact_name in ("capture", "print", "histogram", "stdout", "stderr"):
            checked = checked_artifact(row.get(artifact_name), f"memory row {index}.{artifact_name}")
            if not Path(str(checked["path"])).is_relative_to(raw_path):
                raise ValueError(f"memory row {index}.{artifact_name} escapes raw_root")
            artifact_path = Path(str(checked["path"]))
            if artifact_name == "capture":
                if artifact_path.parent != work_path.resolve() or not artifact_path.name.startswith("heaptrack-profile"):
                    raise ValueError(f"memory row {index}.capture is not in its bound work directory")
            else:
                expected_output = work_path.resolve() / {
                    "print": "heaptrack-print.txt",
                    "histogram": "heaptrack-histogram.tsv",
                    "stdout": "probe.stdout",
                    "stderr": "probe.stderr",
                }[artifact_name]
                if artifact_path != expected_output:
                    raise ValueError(f"memory row {index}.{artifact_name} is not bound to its label")
            if str(artifact_path) in retained_paths:
                raise ValueError(f"memory row {index}.{artifact_name} reuses a retained artifact")
            retained_paths.add(str(artifact_path))
            checked_outputs[artifact_name] = checked
        command = row.get("command")
        cpu = value.get("cpu")
        capture_path = Path(str(checked_outputs["capture"]["path"]))
        expected_prefix = capture_path.parent / "heaptrack-profile"
        expected_command = [
            "/usr/bin/heaptrack",
            "--record-only",
            "-o",
            str(expected_prefix),
            "taskset",
            "-c",
            str(cpu),
            str(endpoint["binary"]["path"]),
            "--mode",
            "memory",
            "--route",
            route,
            "--calls",
            str(calls),
            "--blocks",
            "1",
        ]
        if not isinstance(cpu, int) or command != expected_command:
            raise ValueError(f"memory row {index} CPU/profiler command is not bound")
        probe = load_json(Path(str(checked_outputs["stdout"]["path"])), f"memory row {index}.stdout")
        probe_rows = probe.get("fixtures")
        if (
            probe.get("schema") != SCHEMA
            or probe.get("route") != route
            or probe.get("mode") != "memory"
            or probe.get("calls") != calls
            or probe.get("fixture_count") != 1
            or probe.get("operations") != calls
            or probe.get("scope") != "whole_process_including_fixture_open_and_setup"
            or not isinstance(probe_rows, list)
            or len(probe_rows) != 1
            or not isinstance(probe_rows[0], dict)
            or probe_rows[0].get("name") != fixture
        ):
            raise ValueError(f"memory row {index} probe counter/scope proof is invalid")
        snapshots = row.get("host_snapshots")
        if (
            not isinstance(snapshots, list)
            or len(snapshots) < 2
            or not isinstance(row.get("during_snapshots"), int)
            or row["during_snapshots"] < 1
            or len(snapshots) != row["during_snapshots"] + 2
        ):
            raise ValueError(f"memory row {index} host snapshots are incomplete")
        for snapshot_index, snapshot in enumerate(snapshots):
            snapshot_artifact = checked_artifact(snapshot, f"memory row {index}.host")
            snapshot_path = Path(str(snapshot_artifact["path"]))
            if not snapshot_path.is_relative_to(raw_path):
                raise ValueError(f"memory row {index} host snapshot escapes raw_root")
            if snapshot_index == 0:
                expected_snapshot = raw_path / f"{expected_label}.before.host.json"
            elif snapshot_index == len(snapshots) - 1:
                expected_snapshot = raw_path / f"{expected_label}.after.host.json"
            else:
                expected_snapshot = raw_path / f"{expected_label}.during-{snapshot_index:05}.host.json"
            if snapshot_path != expected_snapshot.resolve():
                raise ValueError(f"memory row {index} host snapshot is not bound to its label")
            if str(snapshot_path) in retained_paths:
                raise ValueError(f"memory row {index} host snapshot reuses a retained artifact")
            retained_paths.add(str(snapshot_path))
            snapshot_value = load_json(snapshot_path, f"memory row {index} host snapshot")
            if snapshot_value.get("blockers") != []:
                raise ValueError(f"memory row {index} host snapshot has blockers")
        failed_after = row.get("failed_after")
        if (
            not isinstance(failed_after, dict)
            or failed_after.get("process_group_empty") is not True
            or failed_after.get("exit_code") != 0
            or failed_after.get("guard_error") is not None
            or failed_after.get("after_guard_error") is not None
            or failed_after.get("timed_out") is not False
            or failed_after.get("rss_max_bytes") != row["rss_max_bytes"]
            or failed_after.get("rss_limit_bytes") != MEMORY_RSS_LIMIT
            or failed_after.get("output_limit_bytes") != MEMORY_OUTPUT_LIMIT
            or not isinstance(failed_after.get("stdout_bytes"), int)
            or not isinstance(failed_after.get("stderr_bytes"), int)
            or failed_after["stdout_bytes"] > MEMORY_OUTPUT_LIMIT
            or failed_after["stderr_bytes"] > MEMORY_OUTPUT_LIMIT
            or failed_after["stdout_bytes"] != checked_outputs["stdout"]["bytes"]
            or failed_after["stderr_bytes"] != checked_outputs["stderr"]["bytes"]
            or failed_after.get("aggregate_rss_max_bytes") != row["aggregate_rss_max_bytes"]
        ):
            raise ValueError(f"memory row {index} failed-after guard receipt is incomplete")
        parsed = row.get("parsed")
        if not isinstance(parsed, dict) or parsed.get("status") != "ok":
            raise ValueError(f"memory row {index} has no parsed heap totals")
        recomputed = recompute_heap_metrics(
            Path(str(checked_outputs["print"]["path"])),
            Path(str(checked_outputs["histogram"]["path"])),
        )
        if recomputed.get("status") != "ok":
            raise ValueError(f"memory row {index} retained heaptrack output is not parseable")
        for metric in ("allocation_calls", "temporary_allocations", "peak_heap_bytes", "peak_rss_bytes", "allocated_bytes"):
            metric_value = parsed.get(metric)
            if not isinstance(metric_value, (int, float)) or not math.isfinite(float(metric_value)) or metric_value < 0:
                raise ValueError(f"memory row {index}.{metric} is missing or non-finite")
            recomputed_value = recomputed.get(metric)
            if recomputed_value != metric_value:
                raise ValueError(f"memory row {index}.{metric} does not match retained output")
    if observed != expected:
        raise ValueError("memory summary is missing a fixture/route/call tuple")
