#!/usr/bin/env python3
"""Run the source-fenced paired DrawingML endpoint qualification protocol.

The raw families and four owner transitions do not share one historical
parent.  This coordinator manifest therefore describes immutable endpoint
pairs, each with its own baseline worktree and preparation receipt, and one
assembled candidate.  Each endpoint is materialized into the already-reviewed generic native runner, so
Cargo path validation, 20 GB target/address-space limits, process quietness,
fresh child processes, output handoff, and ABBA ordering remain shared with
the established W4b/W5 lanes.

This command is orchestration only.  The checked-in probe contains real
source-backed adapters; a source-pair preparation or endpoint failure is
fatal and cannot be converted into an acceptance claim.  The validator still
understands ``unavailable`` control reports for static protocol tests.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import threading
import time
from typing import Any, Mapping


TOOLS_ROOT = Path(__file__).resolve().parents[1]
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

import native_qualification_runner as generic  # noqa: E402
from native_qualification import prepare_drawingml_raw_inputs as preparation  # noqa: E402


SCHEMA = "litchi.drawingml.raw-qualification.manifest.v1"
RECEIPT_SCHEMA = "litchi.drawingml.raw-qualification.receipt.v1"
PLAN_PATH = "docs/unified-ops/results/0362dvie/qualification-plan.json"
PLAN_SHA256 = "e0cb370e0399c072de9fe4b141ceff682b87d7936bec5332cedc9a181fc13505"
EXTERNAL_RESOURCE_SCHEMA = "litchi.drawingml.external-process-resource.v1"
EXTERNAL_RESOURCE_FILENAME = "external-process-resource.jsonl"
EXTERNAL_RESOURCE_EXCLUDED_EVENTS = frozenset(
    {
        "validator-drawingml-raw-identity",
        "control-drawingml-fixture-preflight-before",
        "control-drawingml-fixture-preflight-after",
    }
)
LANES = (
    "drawingml-chart-raw",
    "drawingml-diagram-raw",
    "drawingml-text-media-raw",
    "drawingml-theme-ink",
    "drawingml-geometry-color-formatcode2-transform",
)
ABBA = ("baseline", "candidate", "candidate", "baseline")
OWNER_ROUTES = {
    "drawingml-chart": "drawingml-chart-raw",
    "drawingml-theme-family": "drawingml-theme-ink",
    "drawingml-ink-actions": "drawingml-theme-ink",
    "drawingml-transform": "drawingml-geometry-color-formatcode2-transform",
}
RAW_SCOPES = {"bulk", "family"}
DRAWINGML_FEATURES = {
    "drawingml-owner-current",
    "drawingml-theme-family-current",
    "drawingml-raw-current",
    "drawingml-ink-limits-current",
    "drawingml-theme-family-owner-limits-current",
    "drawingml-transform-limits-current",
}
DRAWINGML_PROBE_FEATURE = "drawingml-probe"
ADAPTER_PATHS = tuple(preparation.ADAPTER_PATHS)
PROBE_PATHS = tuple(path for path in ADAPTER_PATHS if not path.endswith("Cargo.lock"))
TOKEN = re.compile(r"\$\{([A-Z][A-Z0-9_]*)\}")


class RunnerError(RuntimeError):
    """A fail-closed manifest, preparation, or child-runner error."""


def _argv_sha256(argv: list[str]) -> str:
    """Match the probe's NUL-joined command identity, independently."""

    return hashlib.sha256("\0".join(argv).encode("utf-8")).hexdigest()


def _boot_id() -> str:
    try:
        value = Path("/proc/sys/kernel/random/boot_id").read_text(encoding="ascii").strip()
    except (OSError, UnicodeError) as error:
        raise RunnerError(f"cannot read kernel boot identity: {error}") from error
    if not value:
        raise RunnerError("kernel boot identity is empty")
    return value


def _process_start_identity(pid: int, boot_id: str) -> dict[str, Any] | None:
    """Read a PID identity that remains meaningful if the PID is recycled."""

    try:
        stat = (Path("/proc") / str(pid) / "stat").read_text(
            encoding="utf-8", errors="replace"
        )
    except (FileNotFoundError, PermissionError, OSError):
        return None
    right = stat.rsplit(") ", 1)
    if len(right) != 2:
        return None
    fields = right[1].split()
    try:
        # /proc/<pid>/stat field 22 is the twentieth value after the
        # comm/state prefix: starttime in clock ticks since boot.
        starttime_ticks = int(fields[19])
    except (IndexError, ValueError):
        return None
    return {"pid": pid, "boot_id": boot_id, "starttime_ticks": starttime_ticks}


def _status_bytes(pid: int, key: str) -> int | None:
    try:
        status = (Path("/proc") / str(pid) / "status").read_text(
            encoding="ascii", errors="replace"
        )
    except (FileNotFoundError, PermissionError, OSError):
        return None
    prefix = f"{key}:"
    for line in status.splitlines():
        if not line.startswith(prefix):
            continue
        fields = line[len(prefix) :].split()
        if not fields:
            return None
        try:
            value = int(fields[0])
        except ValueError:
            return None
        unit = fields[1] if len(fields) > 1 else "bytes"
        multiplier = {"kB": 1024, "MB": 1024 * 1024, "GB": 1024 * 1024 * 1024}.get(
            unit, 1
        )
        return value * multiplier
    return None


class _ResourceToken:
    """Poll one actual measurement child until its process exits."""

    def __init__(
        self,
        *,
        event_id: str,
        leg: str,
        argv: list[str],
        cwd: Path,
        identity: Mapping[str, Any],
        process: Any,
        boot_id: str,
    ) -> None:
        self.event_id = event_id
        self.leg = leg
        self.argv = argv
        self.cwd = str(cwd)
        self.source_ref = {
            "source_root": str(identity["source_root"]),
            "source_revision": str(identity["head"]),
            "source_identity_sha256": str(identity["identity_sha256"]),
            "leg": leg,
        }
        self.pid = int(process.pid)
        self.start_identity = _process_start_identity(self.pid, boot_id)
        self.argv_sha256 = _argv_sha256(argv)
        self._stop = threading.Event()
        self._thread: threading.Thread | None = None
        self.sample_count = 0
        self.process_rss_bytes = 0
        self.peak_rss_bytes = 0
        self.peak_vm_hwm_bytes = 0
        self.error: str | None = (
            None
            if self.start_identity is not None
            else f"measurement child {self.pid} disappeared before start identity capture"
        )
        if self.start_identity is not None:
            self._thread = threading.Thread(
                target=self._poll,
                name=f"drawingml-resource-{self.pid}",
                daemon=True,
            )
            self._thread.start()

    def _sample(self) -> bool:
        current = _process_start_identity(self.pid, self.start_identity["boot_id"])
        if current is None:
            return False
        if current["starttime_ticks"] != self.start_identity["starttime_ticks"]:
            self.error = "measurement PID was recycled while sampling"
            return False
        rss = _status_bytes(self.pid, "VmRSS")
        hwm = _status_bytes(self.pid, "VmHWM")
        if rss is None or hwm is None:
            self.error = "measurement child has no readable VmRSS/VmHWM"
            return False
        self.sample_count += 1
        self.process_rss_bytes = rss
        self.peak_rss_bytes = max(self.peak_rss_bytes, rss)
        self.peak_vm_hwm_bytes = max(self.peak_vm_hwm_bytes, hwm)
        return True

    def _poll(self) -> None:
        self._sample()
        while not self._stop.wait(0.01):
            if not self._sample():
                return

    def finish(self, returncode: int | None) -> dict[str, Any]:
        self._stop.set()
        if self._thread is not None:
            self._thread.join(timeout=2)
            if self._thread.is_alive() and self.error is None:
                self.error = "resource sampler did not stop"
        captured = (
            self.error is None
            and self.sample_count > 0
            and self.process_rss_bytes > 0
            and self.peak_vm_hwm_bytes > 0
        )
        return {
            "schema": EXTERNAL_RESOURCE_SCHEMA,
            "event_id": self.event_id,
            "leg": self.leg,
            "status": "captured" if captured else "unavailable",
            "reason": None if captured else (self.error or "resource sampler observed no usable memory"),
            "pid": self.pid,
            "start_identity": self.start_identity,
            "argv_sha256": self.argv_sha256,
            "cwd": self.cwd,
            "source_ref": self.source_ref,
            "returncode": returncode,
            "sample_count": self.sample_count,
            "process_rss_bytes": self.process_rss_bytes,
            "peak_rss_bytes": self.peak_rss_bytes,
            "peak_vm_hwm_bytes": self.peak_vm_hwm_bytes,
        }


class ExternalResourceObserver:
    """Capture only the direct child exposed by SafeRunner's observer hook."""

    def __init__(self, output_dir: Path) -> None:
        self.path = output_dir / EXTERNAL_RESOURCE_FILENAME
        if self.path.is_symlink() or self.path.exists():
            raise RunnerError(f"external resource journal is not fresh: {self.path}")
        self._boot_id = _boot_id()
        self._lock = threading.Lock()
        self._records = 0

    def process_started(self, **kwargs: Any) -> _ResourceToken | None:
        # Explicit metadata from the raw adapter takes precedence over an
        # event-name fallback.  Controls are journaled generically but are
        # never measured in the eight-child resource sidecar.
        if (
            kwargs.get("event_kind") == "control"
            or kwargs.get("measured") is False
            or kwargs.get("event_id") in EXTERNAL_RESOURCE_EXCLUDED_EVENTS
        ):
            return None
        return _ResourceToken(boot_id=self._boot_id, **kwargs)

    def process_finished(self, token: _ResourceToken, *, returncode: int | None) -> dict[str, Any]:
        record = token.finish(returncode)
        line = json.dumps(record, sort_keys=True, separators=(",", ":")) + "\n"
        with self._lock:
            if self.path.is_symlink():
                raise RunnerError(f"external resource journal became a symlink: {self.path}")
            mode = "x" if not self.path.exists() else "a"
            with self.path.open(mode, encoding="utf-8") as stream:
                stream.write(line)
                stream.flush()
                os.fsync(stream.fileno())
            self._records += 1
        line_hash = _sha256_bytes(line.encode("utf-8"))
        return {
            "schema": EXTERNAL_RESOURCE_SCHEMA,
            "path": str(self.path),
            "event_id": record["event_id"],
            "record_sha256": line_hash,
            "record_bytes": len(line.encode("utf-8")),
        }

    def receipt(self) -> dict[str, Any]:
        if self.path.is_symlink() or not self.path.is_file():
            raise RunnerError(f"external resource journal is missing: {self.path}")
        digest, size = _sha256_file(self.path)
        return {
            "schema": EXTERNAL_RESOURCE_SCHEMA,
            "path": str(self.path),
            "sha256": digest,
            "bytes": size,
            "records": self._records,
        }


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _expand(value: Any) -> Any:
    if isinstance(value, str):
        expanded = os.path.expandvars(value)
        if "$" in expanded:
            raise RunnerError(f"unexpanded environment variable in manifest value: {value!r}")
        return expanded
    if isinstance(value, list):
        return [_expand(item) for item in value]
    if isinstance(value, dict):
        return {key: _expand(item) for key, item in value.items()}
    return value


def _path(value: str, *, base: Path | None = None) -> Path:
    try:
        return generic._resolve_path(value, base=base)
    except generic.RunnerError as error:
        raise RunnerError(str(error)) from error


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _disjoint(paths: Mapping[str, Path]) -> None:
    values = list(paths.items())
    for index, (left_name, left) in enumerate(values):
        for right_name, right in values[index + 1 :]:
            if left == right or _under(left, right) or _under(right, left):
                raise RunnerError(f"paths must be disjoint: {left_name}={left}, {right_name}={right}")


def _load(path: Path) -> tuple[dict[str, Any], str]:
    try:
        raw = path.read_bytes()
        value = json.loads(raw)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"cannot read manifest {path}: {error}") from error
    if not isinstance(value, dict):
        raise RunnerError("manifest root must be an object")
    return _expand(value), _sha256_bytes(raw)


def _full_sha(value: Any, label: str) -> str:
    if not isinstance(value, str) or len(value) != 40 or value != value.lower() or any(char not in "0123456789abcdef" for char in value):
        raise RunnerError(f"{label} must be a lowercase 40-character commit SHA")
    return value


def _full_hash(value: Any, label: str) -> str:
    if not isinstance(value, str) or len(value) != 64 or value != value.lower() or any(char not in "0123456789abcdef" for char in value):
        raise RunnerError(f"{label} must be a lowercase 64-character SHA-256")
    return value


def _validate_master(data: Mapping[str, Any]) -> dict[str, Any]:
    if data.get("schema") != SCHEMA:
        raise RunnerError(f"manifest schema must be {SCHEMA!r}")
    if data.get("plan_path") != PLAN_PATH:
        raise RunnerError("external DrawingML qualification plan path differs")
    if data.get("plan_sha256") != PLAN_SHA256:
        raise RunnerError("external DrawingML qualification plan hash differs")
    if tuple(profile.get("id") for profile in data.get("profiles", []) if isinstance(profile, dict)) != LANES:
        raise RunnerError("manifest profile order or count differs from the five reviewed raw lanes")
    candidate_root = _path(str(data.get("candidate_root", "")))
    candidate_revision = _full_sha(data.get("candidate_revision"), "candidate_revision")
    baseline_workspace_lock_sha256 = _full_hash(
        data.get("baseline_workspace_lock_sha256"),
        "baseline_workspace_lock_sha256",
    )
    if baseline_workspace_lock_sha256 != preparation.HISTORICAL_WORKSPACE_LOCK_SHA256:
        raise RunnerError("baseline workspace lock is not the reviewed historical 52ed identity")
    baseline_workspace_lock_bytes = data.get("baseline_workspace_lock_bytes")
    if baseline_workspace_lock_bytes != preparation.HISTORICAL_WORKSPACE_LOCK_BYTES:
        raise RunnerError("baseline workspace lock byte count differs from the reviewed 52ed identity")
    candidate_workspace_lock_sha256 = _full_hash(
        data.get("candidate_workspace_lock_sha256"),
        "candidate_workspace_lock_sha256",
    )
    candidate_workspace_lock_bytes = data.get("candidate_workspace_lock_bytes")
    if not isinstance(candidate_workspace_lock_bytes, int) or candidate_workspace_lock_bytes <= 0:
        raise RunnerError("candidate_workspace_lock_bytes must be a positive integer")
    output = _path(str(data.get("output_dir", "")))
    target_root = _path(str(data.get("target_root", "")))
    if output.exists() and output.is_symlink():
        raise RunnerError("output_dir may not be a symlink")
    if target_root.exists() and target_root.is_symlink():
        raise RunnerError("target_root may not be a symlink")
    if not candidate_root.is_dir():
        raise RunnerError(f"candidate root is not a directory: {candidate_root}")
    if _under(output, candidate_root) or _under(target_root, candidate_root):
        raise RunnerError("output and target roots must be external to candidate source")
    fixture = _path("tools/native_qualification/fixtures/drawingml_raw_fixtures.json", base=candidate_root)
    if fixture.is_symlink() or not fixture.is_file():
        raise RunnerError(f"tracked DrawingML fixture descriptor is missing: {fixture}")
    fixture_hash, fixture_bytes = _sha256_file(fixture)
    if data.get("fixture_sha256") != fixture_hash or data.get("fixture_bytes") != fixture_bytes:
        raise RunnerError("manifest fixture identity differs from the tracked descriptor")
    profiles: list[dict[str, Any]] = []
    profile_specs = list(data["profiles"])
    for key in ("route_profiles", "owner_profiles"):
        extra = data.get(key, [])
        if not isinstance(extra, list):
            raise RunnerError(f"{key} must be a list")
        profile_specs.extend(extra)
    source_paths = {"candidate": candidate_root, "output": output, "target": target_root}
    profile_ids: set[str] = set()
    for profile in profile_specs:
        if not isinstance(profile, dict):
            raise RunnerError("profile must be an object")
        profile_id = profile["id"]
        if not isinstance(profile_id, str) or not profile_id or profile_id in profile_ids:
            raise RunnerError(f"duplicate or invalid endpoint profile id: {profile_id!r}")
        profile_ids.add(profile_id)
        lane = profile.get("lane", profile_id)
        endpoint = profile.get("endpoint", "raw")
        owner_route = profile.get("owner_route")
        raw_scope = profile.get("raw_scope")
        if lane not in LANES:
            raise RunnerError(f"{profile_id} lane is not one of the reviewed raw lanes")
        if endpoint not in {"raw", "owner"}:
            raise RunnerError(f"{profile_id} endpoint is not raw or owner")
        if endpoint == "owner":
            if owner_route not in OWNER_ROUTES:
                raise RunnerError(f"{profile_id} owner route is not reviewed: {owner_route!r}")
            if OWNER_ROUTES[owner_route] != lane:
                raise RunnerError(f"{profile_id} owner route does not belong to its lane")
        if endpoint == "raw" and owner_route is not None:
            raise RunnerError(f"{profile_id} raw endpoint carries an owner route")
        if endpoint == "owner" and raw_scope is not None:
            raise RunnerError(f"{profile_id} owner endpoint carries a raw scope")
        if raw_scope is not None and raw_scope not in RAW_SCOPES:
            raise RunnerError(f"{profile_id} has an unknown raw scope")
        baseline_features = profile.get("baseline_features", [])
        candidate_features = profile.get("candidate_features", ["drawingml-owner-current"])
        for label, features in (
            ("baseline_features", baseline_features),
            ("candidate_features", candidate_features),
        ):
            if not isinstance(features, list) or any(
                not isinstance(feature, str) or not feature for feature in features
            ):
                raise RunnerError(f"{profile_id}.{label} must be a list of non-empty feature names")
            unknown_features = set(features) - DRAWINGML_FEATURES
            if unknown_features:
                raise RunnerError(
                    f"{profile_id}.{label} has unknown DrawingML feature(s): {sorted(unknown_features)}"
                )
        if profile_id == "drawingml-theme-family-raw":
            if "drawingml-owner-current" not in baseline_features:
                raise RunnerError(
                    "drawingml-theme-family-raw must compile its 11728 baseline with the current Package/Edit owner API"
                )
            if "drawingml-theme-family-current" in baseline_features:
                raise RunnerError(
                    "drawingml-theme-family-raw must keep the 11728 part caller-limit API on its legacy branch"
                )
            if "drawingml-theme-family-current" not in candidate_features:
                raise RunnerError(
                    "drawingml-theme-family-raw candidate must select the current caller-limit API"
                )
        baseline_root = _path(str(profile.get("baseline_root", "")))
        baseline_revision = _full_sha(profile.get("baseline_revision"), f"{profile_id}.baseline_revision")
        perf_lock_sha256 = _full_hash(profile.get("perf_lock_sha256"), f"{profile_id}.perf_lock_sha256")
        receipt = _path(str(profile.get("preparation_receipt", "")))
        approved = profile.get("approved_baseline_revisions")
        if not isinstance(approved, list) or not approved:
            raise RunnerError(f"{profile_id} baseline revision is not one of its reviewed source parents")
        for index, revision in enumerate(approved):
            _full_sha(revision, f"{profile_id}.approved_baseline_revisions[{index}]")
        if baseline_revision not in approved:
            raise RunnerError(f"{profile_id} baseline revision is not one of its reviewed source parents")
        route_baselines = profile.get("route_baselines")
        if route_baselines is not None:
            if not isinstance(route_baselines, dict):
                raise RunnerError(f"{profile_id}.route_baselines must be an object")
            if lane == "drawingml-theme-ink" and set(route_baselines) != {"theme", "ink"}:
                raise RunnerError("drawingml-theme-ink route_baselines must name only theme and ink")
            for route, revisions in route_baselines.items():
                values = [revisions] if isinstance(revisions, str) else revisions
                if not isinstance(values, list) or not values:
                    raise RunnerError(f"{profile_id}.{route} has no exact baseline parent")
                for revision in values:
                    _full_sha(revision, f"{profile_id}.{route}.baseline")
            theme_parent = route_baselines.get("theme")
            if lane == "drawingml-theme-ink" and not isinstance(theme_parent, str):
                raise RunnerError("drawingml-theme-ink needs one exact theme parent; Ink routes are separate")
            if isinstance(theme_parent, str) and baseline_revision != theme_parent:
                raise RunnerError(f"{profile_id} baseline must be the exact theme parent; Ink routes need a separate endpoint")
        if not baseline_root.is_dir():
            raise RunnerError(f"{profile_id} baseline root is not a directory: {baseline_root}")
        if not receipt.is_file() or receipt.is_symlink():
            raise RunnerError(f"{profile_id} preparation receipt is missing: {receipt}")
        # Several lanes intentionally share one immutable parent worktree
        # (chart/diagram/geometry currently use the same owner parent).  That
        # is safe because each child target/output is separate; reject only
        # overlap with candidate/output/targets and receipt paths.
        for name, other in source_paths.items():
            # The same immutable parent may serve chart/diagram/geometry, but
            # a baseline must never contain or be contained by a receipt from
            # another profile.  Otherwise a later source preparation could
            # mutate a path that the manifest treats as an external receipt.
            if name.startswith("baseline-"):
                continue
            if baseline_root == other or _under(baseline_root, other) or _under(other, baseline_root):
                raise RunnerError(f"{profile_id} baseline overlaps {name}: {baseline_root}")
        for name, other in source_paths.items():
            if receipt == other or _under(receipt, other) or _under(other, receipt):
                raise RunnerError(f"{profile_id} preparation receipt overlaps {name}: {receipt}")
        source_paths[f"baseline-{profile_id}"] = baseline_root
        source_paths[f"receipt-{profile_id}"] = receipt
        for key in ("source_fence", "record"):
            if not isinstance(profile.get(key), str) or not profile[key]:
                raise RunnerError(f"{profile_id}.{key} is required")
        profiles.append({
            "id": profile_id,
            "lane": lane,
            "endpoint": endpoint,
            "owner_route": owner_route,
            "raw_scope": raw_scope,
            "baseline_root": baseline_root,
            "baseline_revision": baseline_revision,
            "perf_lock_sha256": perf_lock_sha256,
            "preparation_receipt": receipt,
            "source_fence": profile["source_fence"],
            "record": profile["record"],
            "approved_baseline_revisions": tuple(approved),
            "route_baselines": route_baselines,
            "baseline_features": tuple(baseline_features),
            "candidate_features": tuple(candidate_features),
        })
    # Baseline worktrees may be shared between profiles at the same parent;
    # all of them remain disjoint from the candidate, outputs, targets, and
    # external preparation receipts by the checks above.
    return {
        "schema": SCHEMA,
        "plan_path": PLAN_PATH,
        "plan_sha256": PLAN_SHA256,
        "candidate_root": candidate_root,
        "candidate_revision": candidate_revision,
        "baseline_workspace_lock_sha256": baseline_workspace_lock_sha256,
        "baseline_workspace_lock_bytes": baseline_workspace_lock_bytes,
        "candidate_workspace_lock_sha256": candidate_workspace_lock_sha256,
        "candidate_workspace_lock_bytes": candidate_workspace_lock_bytes,
        "output_dir": output,
        "target_root": target_root,
        "fixture": fixture,
        "fixture_sha256": fixture_hash,
        "fixture_bytes": fixture_bytes,
        "profiles": profiles,
        "manifest_path": data.get("_manifest_path"),
    }


def _profile_manifest(master: Mapping[str, Any], profile: Mapping[str, Any], output: Path) -> dict[str, Any]:
    profile_id = str(profile["id"])
    lane = str(profile.get("lane", profile_id))
    endpoint = str(profile.get("endpoint", "raw"))
    owner_route = profile.get("owner_route")
    raw_scope = profile.get("raw_scope")
    before = Path(profile["baseline_root"])
    after = Path(master["candidate_root"])
    before_revision = str(profile["baseline_revision"])
    after_revision = str(master["candidate_revision"])
    lane_output = output / "lanes" / profile_id
    target_before = Path(master["target_root"]) / profile_id / "baseline"
    target_after = Path(master["target_root"]) / profile_id / "candidate"
    adapter_paths = list(PROBE_PATHS)

    def features_for(role: str) -> list[str]:
        features = list(
            profile.get(
                "baseline_features" if role == "baseline" else "candidate_features",
                (),
            )
        )
        if DRAWINGML_PROBE_FEATURE not in features:
            features.append(DRAWINGML_PROBE_FEATURE)
        return features

    def build(role: str) -> list[dict[str, Any]]:
        features = features_for(role)
        feature_args = ["--features", ",".join(features)] if features else []
        no_default = ["--no-default-features"] if role == "baseline" else []
        cargo = [
            "cargo",
            "build",
            "--release",
            "--locked",
            "--manifest-path",
            "tools/perf-baseline/Cargo.toml",
            *no_default,
            *feature_args,
        ]
        return [
            {
                "id": "drawingml_raw_native",
                "argv": [*cargo, "--bin", "drawingml-raw-native"],
                "artifacts": ["{target_dir}/release/drawingml-raw-native"],
            },
            {
                "id": "drawingml_raw_allocator",
                "argv": [*cargo, "--bin", "drawingml-raw-allocator"],
                "artifacts": ["{target_dir}/release/drawingml-raw-allocator"],
            },
        ]

    def programs(role: str, leg: str) -> list[dict[str, Any]]:
        revision = before_revision if role == "baseline" else after_revision
        result = []
        for program_id, mode, binary in (
            ("raw-native", "normal", "drawingml-raw-native"),
            ("raw-allocator", "allocator", "drawingml-raw-allocator"),
        ):
            prefix = "native" if mode == "normal" else "allocator"
            output_name = f"{{output_dir}}/dml-{prefix}-{{leg}}-{{index}}.json"
            result.append({
                "id": program_id,
                "binary": f"{{target_dir}}/release/{binary}",
                "argv": [
                    "--lane", lane, "--endpoint", endpoint,
                    *( ["--owner-route", str(owner_route)] if owner_route else [] ),
                    *( ["--raw-scope", str(raw_scope)] if raw_scope else [] ),
                    "--role", role, "--mode", mode,
                    "--fixture", "{fixture}", "--source-revision", revision,
                    "--iterations", "1", "--warmups", "3", "--blocks", "5",
                    "--samples", "20", "--target-ms", "500", "--minimum-ms", "150",
                    "--json", "{output}",
                ],
                "outputs": [output_name],
                "stdout": output_name.removesuffix(".json") + ".stdout",
                "stderr": output_name.removesuffix(".json") + ".stderr",
                "require_stdout": False,
            })
        return result

    def fixture_preflight(role: str, leg: str) -> dict[str, Any]:
        """Run the untimed fixture-identity handoff for one endpoint."""

        revision = before_revision if role == "baseline" else after_revision
        output_name = f"{{output_dir}}/controls/dml-preflight-{leg}.json"
        return {
            "id": "fixture-preflight",
            "binary": "{target_dir}/release/drawingml-raw-native",
            "argv": [
                "--lane", lane, "--endpoint", endpoint,
                *( ["--owner-route", str(owner_route)] if owner_route else [] ),
                *( ["--raw-scope", str(raw_scope)] if raw_scope else [] ),
                "--role", role, "--mode", "normal",
                "--fixture", "{fixture}", "--source-revision", revision,
                "--iterations", "1", "--warmups", "3", "--blocks", "5",
                "--samples", "20", "--target-ms", "500", "--minimum-ms", "150",
                "--preflight", "--json", "{output}",
            ],
            "outputs": [output_name],
            "stdout": output_name.removesuffix(".json") + ".stdout",
            "stderr": output_name.removesuffix(".json") + ".stderr",
            "require_stdout": False,
        }

    return {
        "schema": generic.SCHEMA,
        "lane": "drawingml-raw",
        "drawingml_profile": profile_id,
        "drawingml_lane": lane,
        "drawingml_endpoint": endpoint,
        "drawingml_owner_route": owner_route,
        "drawingml_raw_scope": raw_scope,
        "baseline_workspace_lock_sha256": str(master["baseline_workspace_lock_sha256"]),
        "baseline_workspace_lock_bytes": int(master["baseline_workspace_lock_bytes"]),
        "candidate_workspace_lock_sha256": str(master["candidate_workspace_lock_sha256"]),
        "candidate_workspace_lock_bytes": int(master["candidate_workspace_lock_bytes"]),
        "drawingml_baseline_features": features_for("baseline"),
        "drawingml_candidate_features": features_for("candidate"),
        "source_fence": str(profile["source_fence"]),
        "rounds": 31,
        "abba_order": list(generic.ABBA_ORDER),
        "output_dir": str(lane_output),
        "shared": {
            "probe_paths": adapter_paths,
            "fixture_paths": ["tools/native_qualification/fixtures/drawingml_raw_fixtures.json"],
            "preparation_receipt": str(profile["preparation_receipt"]),
        },
        "legs": {
            "before": {
                "source_root": str(before),
                "source_revision": before_revision,
                "target_dir": str(target_before),
                "allow_dirty": True,
                "allowed_dirty_paths": list(preparation.ADAPTER_PATHS),
                "lock_paths": ["Cargo.lock", "tools/perf-baseline/Cargo.lock"],
                "build": build("baseline"),
                "programs": programs("baseline", "before"),
            },
            "after": {
                "source_root": str(after),
                "source_revision": after_revision,
                "target_dir": str(target_after),
                "allow_dirty": True,
                "allowed_dirty_paths": list(preparation.ADAPTER_PATHS),
                "lock_paths": ["Cargo.lock", "tools/perf-baseline/Cargo.lock"],
                "build": build("candidate"),
                "programs": programs("candidate", "after"),
            },
        },
        "controls": [{
            "id": "drawingml-fixture-preflight",
            "compare": [
                "fixture_sha256",
                "fixture_bytes",
                "fixture_lane_identity_sha256",
                "generated_fixture_sha256",
                "generated_fixture_parts",
            ],
            "programs": {
                "before": fixture_preflight("baseline", "before"),
                "after": fixture_preflight("candidate", "after"),
            },
        }],
        "validators": [{
            "id": "drawingml-raw-identity",
            "argv": [
                "python3",
                "{source_root}/tools/native_qualification/drawingml_raw_validator.py",
                "--output-dir", "{output_dir}", "--lane", lane, "--endpoint", endpoint, "--fixture", "{fixture}",
                "--baseline-revision", before_revision,
                "--candidate-revision", after_revision,
                "--source-fence", str(profile["source_fence"]),
            ],
            "script": "{source_root}/tools/native_qualification/drawingml_raw_validator.py",
            "outputs": ["{output_dir}/summary.json"],
            "cwd": "{output_dir}",
            "require_stdout": False,
        }],
    }


def _prepare_profile(master: Mapping[str, Any], profile: Mapping[str, Any]) -> None:
    try:
        preparation.validate_preparation_receipt(
            Path(profile["preparation_receipt"]),
            baseline_root=Path(profile["baseline_root"]),
            candidate_root=Path(master["candidate_root"]),
            baseline_revision=str(profile["baseline_revision"]),
            candidate_revision=str(master["candidate_revision"]),
            perf_lock_sha256=str(profile["perf_lock_sha256"]),
            candidate_workspace_lock_sha256=str(master["candidate_workspace_lock_sha256"]),
            candidate_workspace_lock_bytes=int(master["candidate_workspace_lock_bytes"]),
        )
    except (OSError, ValueError) as error:
        raise RunnerError(f"{profile['id']} preparation receipt validation failed: {error}") from error


def _write_json(path: Path, value: Mapping[str, Any]) -> None:
    if path.exists() or path.is_symlink():
        raise RunnerError(f"refusing to overwrite protocol receipt: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def _run_profile(master: Mapping[str, Any], profile: Mapping[str, Any], phase: str, cpu: int | None, profile_path: Path) -> dict[str, Any]:
    data = _profile_manifest(master, profile, Path(master["output_dir"]))
    rendered, _ = _load(profile_path)
    if rendered != data:
        raise RunnerError(f"materialized {profile['id']} profile manifest changed")
    observer = ExternalResourceObserver(Path(data["output_dir"])) if phase == "abba" else None
    try:
        loaded = generic.load_manifest(profile_path)
        runner = generic.SafeRunner(
            loaded,
            profile_path,
            phase,
            cpu,
            process_observer=observer,
        )
        if phase == "build":
            runner.build()
        else:
            runner.abba()
    except generic.RunnerError as error:
        raise RunnerError(f"{profile['id']} generic runner refused: {error}") from error
    receipt = Path(data["output_dir"]) / f"{phase}-receipt.json"
    if not receipt.is_file() or receipt.is_symlink():
        raise RunnerError(f"{profile['id']} did not produce its {phase} receipt")
    value, _ = _load(receipt)
    if observer is not None:
        value["external_resource_capture"] = observer.receipt()
    return value


def run(manifest_path: Path, phase: str, cpu: int | None) -> None:
    raw, manifest_hash = _load(manifest_path.resolve(strict=True))
    raw["_manifest_path"] = str(manifest_path.resolve())
    master = _validate_master(raw)
    output = Path(master["output_dir"])
    runner_path = Path(__file__).resolve()
    generic_runner_path = Path(generic.__file__).resolve()
    runner_sha256, runner_bytes = _sha256_file(runner_path)
    generic_runner_sha256, generic_runner_bytes = _sha256_file(generic_runner_path)
    top_receipt = output / f"{phase}-receipt.json"
    if phase == "build":
        if output.exists() and any(output.iterdir()):
            raise RunnerError(f"build output must be a fresh empty directory: {output}")
        output.mkdir(parents=True, exist_ok=True)
    else:
        build_receipt_path = output / "build-receipt.json"
        if not build_receipt_path.is_file() or build_receipt_path.is_symlink():
            raise RunnerError("ABBA requires the master build receipt")
        build_receipt, _ = _load(build_receipt_path)
        if build_receipt.get("schema") != RECEIPT_SCHEMA or build_receipt.get("phase") != "build" or build_receipt.get("status") != "passed":
            raise RunnerError("master build receipt is not a passed build handoff")
        if (
            build_receipt.get("plan_path") != PLAN_PATH
            or build_receipt.get("plan_sha256") != PLAN_SHA256
        ):
            raise RunnerError("master build receipt plan identity differs")
        if build_receipt.get("manifest_sha256") != manifest_hash:
            raise RunnerError("master manifest differs from the build receipt")
        if build_receipt.get("runner") != {
            "path": str(runner_path),
            "sha256": runner_sha256,
            "bytes": runner_bytes,
            "generic_path": str(generic_runner_path),
            "generic_sha256": generic_runner_sha256,
            "generic_bytes": generic_runner_bytes,
        }:
            raise RunnerError("paired runner or generic guard changed since build")
        if top_receipt.exists() or top_receipt.is_symlink():
            raise RunnerError(f"ABBA receipt already exists: {top_receipt}")
    profile_receipts: dict[str, Any] = {}
    profile_files: list[dict[str, str]] = []
    try:
        for profile in master["profiles"]:
            _prepare_profile(master, profile)
            profile_path = output / "profiles" / f"{profile['id']}.json"
            if phase == "build":
                _write_json(profile_path, _profile_manifest(master, profile, output))
            else:
                if not profile_path.is_file() or profile_path.is_symlink():
                    raise RunnerError(f"{profile['id']} materialized profile is missing")
                profile_hash, _ = _sha256_file(profile_path)
                profile_files.append({
                    "profile": profile["id"],
                    "lane": profile["lane"],
                    "endpoint": profile["endpoint"],
                    "owner_route": profile["owner_route"],
                    "raw_scope": profile["raw_scope"],
                    "baseline_workspace_lock_sha256": master["baseline_workspace_lock_sha256"],
                    "baseline_workspace_lock_bytes": master["baseline_workspace_lock_bytes"],
                    "candidate_workspace_lock_sha256": master["candidate_workspace_lock_sha256"],
                    "candidate_workspace_lock_bytes": master["candidate_workspace_lock_bytes"],
                    "path": str(profile_path),
                    "sha256": profile_hash,
                    "baseline_revision": profile["baseline_revision"],
                    "perf_lock_sha256": profile["perf_lock_sha256"],
                    "source_fence": profile["source_fence"],
                })
            profile_receipts[profile["id"]] = _run_profile(master, profile, phase, cpu, profile_path)
        receipt = {
            "schema": RECEIPT_SCHEMA,
            "phase": phase,
            "status": "passed",
            "plan_path": PLAN_PATH,
            "manifest": {"path": str(manifest_path.resolve()), "sha256": manifest_hash},
            "runner": {
                "path": str(runner_path),
                "sha256": runner_sha256,
                "bytes": runner_bytes,
                "generic_path": str(generic_runner_path),
                "generic_sha256": generic_runner_sha256,
                "generic_bytes": generic_runner_bytes,
            },
            "plan_sha256": PLAN_SHA256,
            "candidate_root": str(master["candidate_root"]),
            "candidate_revision": master["candidate_revision"],
            "baseline_workspace_lock_sha256": master["baseline_workspace_lock_sha256"],
            "baseline_workspace_lock_bytes": master["baseline_workspace_lock_bytes"],
            "candidate_workspace_lock_sha256": master["candidate_workspace_lock_sha256"],
            "candidate_workspace_lock_bytes": master["candidate_workspace_lock_bytes"],
            "abba_order": list(ABBA),
            "profiles": profile_files or [
                {
                    "profile": profile["id"],
                    "lane": profile["lane"],
                    "endpoint": profile["endpoint"],
                    "owner_route": profile["owner_route"],
                    "raw_scope": profile["raw_scope"],
                    "baseline_workspace_lock_sha256": master["baseline_workspace_lock_sha256"],
                    "baseline_workspace_lock_bytes": master["baseline_workspace_lock_bytes"],
                    "candidate_workspace_lock_sha256": master["candidate_workspace_lock_sha256"],
                    "candidate_workspace_lock_bytes": master["candidate_workspace_lock_bytes"],
                    "path": str(output / "profiles" / f"{profile['id']}.json"),
                    "sha256": _sha256_file(output / "profiles" / f"{profile['id']}.json")[0],
                    "baseline_revision": profile["baseline_revision"],
                    "perf_lock_sha256": profile["perf_lock_sha256"],
                    "source_fence": profile["source_fence"],
                }
                for profile in master["profiles"]
            ],
            "children": profile_receipts,
            "claim_scope": "paired DrawingML raw and detached-owner endpoints only; no absolute-profile or whole-library claim",
        }
        if phase == "build":
            _write_json(output / "build-receipt.json", receipt)
        else:
            _write_json(output / "abba-receipt.json", receipt)
    except BaseException as error:
        failure = {
            "schema": RECEIPT_SCHEMA,
            "phase": phase,
            "status": "failed",
            "plan_path": PLAN_PATH,
            "manifest": {"path": str(manifest_path.resolve()), "sha256": manifest_hash},
            "runner": {
                "path": str(runner_path),
                "sha256": runner_sha256,
                "bytes": runner_bytes,
                "generic_path": str(generic_runner_path),
                "generic_sha256": generic_runner_sha256,
                "generic_bytes": generic_runner_bytes,
            },
            "plan_sha256": PLAN_SHA256,
            "baseline_workspace_lock_sha256": master["baseline_workspace_lock_sha256"],
            "baseline_workspace_lock_bytes": master["baseline_workspace_lock_bytes"],
            "candidate_workspace_lock_sha256": master["candidate_workspace_lock_sha256"],
            "candidate_workspace_lock_bytes": master["candidate_workspace_lock_bytes"],
            "children": profile_receipts,
            "error": str(error),
        }
        if not top_receipt.exists() and not top_receipt.is_symlink():
            _write_json(top_receipt, failure)
        raise


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--phase", choices=("build", "abba"), required=True)
    parser.add_argument("--cpu", type=int)
    arguments = parser.parse_args(argv)
    try:
        run(arguments.manifest, arguments.phase, arguments.cpu)
    except (OSError, RunnerError, ValueError) as error:
        print(f"drawingml paired qualification refused: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
