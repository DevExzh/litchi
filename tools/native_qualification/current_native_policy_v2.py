#!/usr/bin/env python3
"""The versioned trust boundary for current-source native qualifications.

The older PPT auxiliary gate and the legacy native preparer intentionally keep
their historical contracts.  This module is the opt-in v2 boundary for a
future current-source run.  In particular, it does not accept a digest from a
caller: the source, lock, toolchain, probe, fixture, and route pins live in
the reviewed policy below and are checked again when a receipt is consumed.

The candidate source pins are deliberately incomplete at this revision.  A
source-only preparation can therefore record a blocked diagnostic receipt,
but every execution/validation entry point fails closed until a later, review
ed commit fills in the exact candidate pins.  The abbreviated font references
are provenance hints only; they are never treated as identities.
"""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import tarfile
from typing import Any, Iterable, Mapping, Sequence


POLICY_SCHEMA = "litchi.current-native-policy.v2"
PREPARATION_SCHEMA = "litchi.current-native-preparation.v2"
QUALIFICATION_SCHEMA = "litchi.current-native-qualification.v2"
PROTOCOL_SCHEMA = "litchi.native-evidence-protocol.v2"
POLICY_VERSION = 2
EXECUTION_RECEIPT_RELATIVE = "execution-receipt.json"

BASELINE_COMMIT = "f2bbbc8d07ab5da26c180f1ff74813d39f632d54"
BASELINE_GIT_TREE = "2e0d0cdf7f6a345961ceca361d53202200d00970"
BASELINE_TREE_SHA256 = "ce92652c439d810849b74b80e650821bc21de9029ea861cd6bbc13e38d809589"
BASELINE_LOCK_SHA256 = "1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c"
BASELINE_LOCK_BYTES = 139_441
BASELINE_ARCHIVE_SHA256 = "de625729ec8e6f8055e619443ad31bf1a91719cd31465975f951d54dbc1c2280"
BASELINE_ARCHIVE_BYTES = 6_614_528_000
TOOLCHAIN_SHA256 = "e3a213e0d222e94d213cafbc20932eb3f76c643b4dd63756acf95192df2aa310"
TOOLCHAIN_BYTES = 86
CARGO_CONFIG_SHA256 = "50b044c83828f2c448ee59c362afaa79de2dd47fbd390f0cd2888391f1d1119d"
CARGO_CONFIG_BYTES = 470

# These are intentionally not hashes.  They identify the three source
# transitions discussed by the current-font candidate work and must not be
# accepted as source identities.
CANDIDATE_HINTS = {
    "root": "efff87...",
    "perf": "de1f8...",
    "native": "989da...",
    "vendor_files_and_manifests": "verified-source-only; not an identity",
}
SOURCE_TRANSITIONS = (
    "font-root -> font-perf",
    "font-perf -> font-native",
    "font-native -> reviewed-current-source",
)

TARGET = "x86_64-unknown-linux-gnu"
PROFILE = "release"
FEATURES = ("all-features",)
ABBA_ORDER = ("before", "after", "after", "before")

FULL_SHA = re.compile(r"^[0-9a-f]{40}$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")


class PolicyError(ValueError):
    """The fixed current-source policy cannot admit this operation."""


class PolicyIncomplete(PolicyError):
    """The reviewed candidate pins have not been frozen yet."""


class PolicyMismatch(PolicyError):
    """An observed source, artifact, or route differs from the policy."""


def _full_sha(value: str | None, label: str) -> str:
    if not isinstance(value, str) or FULL_SHA.fullmatch(value) is None:
        raise PolicyError(f"{label} is not a full lowercase commit/tree SHA")
    return value


def _sha256(value: str | None, label: str) -> str:
    if not isinstance(value, str) or SHA256.fullmatch(value) is None:
        raise PolicyError(f"{label} is not a full lowercase SHA-256")
    return value


def _canonical(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=True) + "\n").encode("utf-8")


@dataclass(frozen=True)
class ArtifactPin:
    """A fixed regular-file identity, relative to the source endpoint."""

    relative_path: str
    sha256: str | None
    bytes: int | None

    def complete(self) -> bool:
        return (
            bool(self.relative_path)
            and not Path(self.relative_path).is_absolute()
            and ".." not in Path(self.relative_path).parts
            and isinstance(self.sha256, str)
            and SHA256.fullmatch(self.sha256) is not None
            and isinstance(self.bytes, int)
            and self.bytes >= 0
        )

    def json(self) -> dict[str, Any]:
        return {
            "relative_path": self.relative_path,
            "sha256": self.sha256,
            "bytes": self.bytes,
        }


@dataclass(frozen=True)
class RoutePin:
    """The immutable route identity used by both the preparer and validator."""

    route: str
    kind: str
    program: str
    output: str
    role: str = "paired"
    status: str = "ok"

    def json(self) -> dict[str, str]:
        return {
            "route": self.route,
            "kind": self.kind,
            "program": self.program,
            "output": self.output,
            "role": self.role,
            "status": self.status,
        }


@dataclass(frozen=True)
class SourcePin:
    """All source-side fields which affect a native qualification."""

    commit: str | None
    git_tree: str | None
    tree_sha256: str | None
    lock: ArtifactPin
    toolchain: ArtifactPin
    cargo_config: ArtifactPin
    probe: ArtifactPin
    fixtures: tuple[ArtifactPin, ...]
    archive: ArtifactPin
    target: str
    profile: str
    features: tuple[str, ...]

    def complete(self) -> bool:
        return (
            isinstance(self.commit, str)
            and FULL_SHA.fullmatch(self.commit) is not None
            and isinstance(self.git_tree, str)
            and FULL_SHA.fullmatch(self.git_tree) is not None
            and isinstance(self.tree_sha256, str)
            and SHA256.fullmatch(self.tree_sha256) is not None
            and self.lock.complete()
            and self.toolchain.complete()
            and self.cargo_config.complete()
            and self.probe.complete()
            and bool(self.fixtures)
            and all(item.complete() for item in self.fixtures)
            and self.archive.complete()
            and self.target == TARGET
            and self.profile == PROFILE
            and self.features == FEATURES
        )

    def json(self) -> dict[str, Any]:
        return {
            "commit": self.commit,
            "git_tree": self.git_tree,
            "tree_sha256": self.tree_sha256,
            "lock": self.lock.json(),
            "toolchain": self.toolchain.json(),
            "cargo_config": self.cargo_config.json(),
            "probe": self.probe.json(),
            "fixtures": [item.json() for item in self.fixtures],
            "archive": self.archive.json(),
            "target": self.target,
            "profile": self.profile,
            "features": list(self.features),
        }


PPT_ROUTES = (
    RoutePin("ppt-scoped-build", "build", "cargo", f"target/{TARGET}/release/examples/scoped_patch_probe"),
    RoutePin("ppt-scoped-correctness-suite", "correctness", "cargo", "state/ppt-correctness-suite.json"),
    RoutePin("ppt-scoped-correctness", "probe", "scoped_patch_correctness", "state/ppt-scoped-correctness.json"),
    RoutePin("ppt-scoped-correctness-validate", "validate", "python3", "state/ppt-scoped-correctness.json"),
)

LEGACY_CURRENT_ROUTES = (
    RoutePin("legacy-current/open_1k", "probe", "legacy-current", "state/open_1k.json"),
    RoutePin("legacy-current/noop_1k", "probe", "legacy-current", "state/noop_1k.json"),
    RoutePin("legacy-current/text_commit_1k", "probe", "legacy-current", "state/text_commit_1k.json"),
    RoutePin("legacy-current/text_commit_save_1k", "probe", "legacy-current", "state/text_commit_save_1k.json"),
    RoutePin("legacy-current/replay_1k", "probe", "legacy-current", "state/replay_1k.json"),
    RoutePin("legacy-current/tail_commit_1k", "probe", "legacy-current", "state/tail_commit_1k.json"),
    RoutePin("legacy-current/unsupported_content", "refusal", "legacy-current", "state/unsupported_content.json", status="refused"),
    RoutePin("legacy-current/protected_commit", "refusal", "legacy-current", "state/protected_commit.json", status="refused"),
)


def _artifact(path: str, digest: str | None, size: int | None) -> ArtifactPin:
    return ArtifactPin(path, digest, size)


_BASELINE_PIN = SourcePin(
    commit=BASELINE_COMMIT,
    git_tree=BASELINE_GIT_TREE,
    tree_sha256=BASELINE_TREE_SHA256,
    lock=_artifact("Cargo.lock", BASELINE_LOCK_SHA256, BASELINE_LOCK_BYTES),
    toolchain=_artifact("rust-toolchain.toml", TOOLCHAIN_SHA256, TOOLCHAIN_BYTES),
    cargo_config=_artifact(".cargo/config.toml", CARGO_CONFIG_SHA256, CARGO_CONFIG_BYTES),
    probe=_artifact("crates/litchi-ppt/examples/scoped_patch_probe.rs", "0350fc1696688e576613bc8b2c2ccb3d39c12b4b8ad890c5dcd14e3e7b4138ca", 5158),
    fixtures=(
        _artifact("test-data/unified-ops/0294/video.ppt", "f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a", 43_520),
    ),
    archive=_artifact("source-archive.tar.zst", BASELINE_ARCHIVE_SHA256, BASELINE_ARCHIVE_BYTES),
    target=TARGET,
    profile=PROFILE,
    features=FEATURES,
)

# The candidate is intentionally not represented by a guessed hash.  These
# None values are an executable part of the policy: ``require_execution_ready``
# must continue to reject the run until a reviewed commit replaces this
# constant with exact pins and updates the record.
_CANDIDATE_PIN = SourcePin(
    commit=None,
    git_tree=None,
    tree_sha256=None,
    lock=_artifact("Cargo.lock", None, None),
    toolchain=_artifact("rust-toolchain.toml", TOOLCHAIN_SHA256, TOOLCHAIN_BYTES),
    cargo_config=_artifact(".cargo/config.toml", CARGO_CONFIG_SHA256, CARGO_CONFIG_BYTES),
    probe=_artifact("crates/litchi-ppt/examples/scoped_patch_probe.rs", None, None),
    fixtures=(),
    archive=_artifact("source-archive.tar.zst", None, None),
    target=TARGET,
    profile=PROFILE,
    features=FEATURES,
)


@dataclass(frozen=True)
class CurrentNativePolicyV2:
    """Fixed policy shared by current-source preparation and validation."""

    baseline: SourcePin
    candidate: SourcePin
    ppt_routes: tuple[RoutePin, ...]
    legacy_routes: tuple[RoutePin, ...]

    @property
    def status(self) -> str:
        return "ready" if self.baseline.complete() and self.candidate.complete() else "incomplete"

    @property
    def routes(self) -> dict[str, tuple[RoutePin, ...]]:
        return {
            "ppt-auxiliary": self.ppt_routes,
            "legacy-current": self.legacy_routes,
        }

    def json(self) -> dict[str, Any]:
        return {
            "schema": POLICY_SCHEMA,
            "version": POLICY_VERSION,
            "status": self.status,
            "protocol": {
                "schema": PROTOCOL_SCHEMA,
                "evidence_catalog": "gate-evidence-catalog.v2",
                "signed_native": "signed-native-protocol.v2",
                "abba_order": list(ABBA_ORDER),
            },
            "source_transitions": list(SOURCE_TRANSITIONS),
            "source_transition_count": len(SOURCE_TRANSITIONS),
            "candidate_hints": dict(CANDIDATE_HINTS),
            "baseline": self.baseline.json(),
            "candidate": self.candidate.json(),
            "routes": {
                lane: [route.json() for route in routes]
                for lane, routes in self.routes.items()
            },
        }

    def digest(self) -> str:
        return hashlib.sha256(_canonical(self.json())).hexdigest()

    def require_execution_ready(self) -> None:
        if self.status != "ready":
            raise PolicyIncomplete(
                "current-source policy v2 is incomplete: exact candidate source/tree/lock/probe/fixture pins are not frozen"
            )

    def routes_for(self, lane: str) -> tuple[RoutePin, ...]:
        try:
            return self.routes[lane]
        except KeyError as error:
            raise PolicyError(f"lane is not admitted by current-source policy v2: {lane}") from error


TRUSTED_POLICY = CurrentNativePolicyV2(
    baseline=_BASELINE_PIN,
    candidate=_CANDIDATE_PIN,
    ppt_routes=PPT_ROUTES,
    legacy_routes=LEGACY_CURRENT_ROUTES,
)


def policy_document() -> dict[str, Any]:
    """Return a detached copy for diagnostics; execution uses TRUSTED_POLICY."""

    return json.loads(json.dumps(TRUSTED_POLICY.json(), sort_keys=True))


def _reject_symlink_components(path: Path, label: str) -> None:
    raw = Path(path)
    cursor = Path(raw.anchor) if raw.is_absolute() else Path()
    components = raw.parts[1:] if raw.is_absolute() else raw.parts
    for component in components:
        cursor /= component
        if cursor.is_symlink():
            raise PolicyMismatch(f"{label} contains a symlink component: {cursor}")


def _safe_source(path: Path, label: str) -> Path:
    lexical = Path(path).absolute()
    _reject_symlink_components(lexical, label)
    if lexical.is_symlink() or not lexical.is_dir():
        raise PolicyMismatch(f"{label} is not a regular source directory: {path}")
    resolved = lexical.resolve()
    if not (resolved / "Cargo.toml").is_file():
        raise PolicyMismatch(f"{label} lacks Cargo.toml: {resolved}")
    return resolved


def _hash_file(path: Path, label: str) -> tuple[str, int]:
    _reject_symlink_components(path, label)
    if path.is_symlink() or not path.is_file():
        raise PolicyMismatch(f"{label} is not a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while True:
            block = stream.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _pin_file(source: Path, pin: ArtifactPin, label: str) -> dict[str, Any]:
    if not pin.complete():
        raise PolicyIncomplete(f"{label} has no exact reviewed pin")
    relative = Path(pin.relative_path)
    if relative.is_absolute() or ".." in relative.parts:
        raise PolicyMismatch(f"{label} has an unsafe relative path")
    path = source / relative
    digest, size = _hash_file(path, label)
    if digest != pin.sha256 or size != pin.bytes:
        raise PolicyMismatch(
            f"{label} differs from the trusted policy: expected {pin.sha256}/{pin.bytes}, observed {digest}/{size}"
        )
    return {"relative_path": pin.relative_path, "path": str(path), "sha256": digest, "bytes": size}


def _git(source: Path, *args: str) -> str:
    try:
        completed = subprocess.run(
            ["git", *args], cwd=source, check=True, text=True,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        raise PolicyMismatch(f"cannot inspect source git identity: {source}: {error}") from error
    return completed.stdout.strip()


def _git_tree_entries(source: Path) -> tuple[dict[str, str], ...]:
    """Return every tracked blob in HEAD, including sparse/omitted paths."""

    raw = _git_tree_payload(source)
    entries: list[dict[str, str]] = []
    for record in raw.split(b"\0"):
        if not record:
            continue
        try:
            header, path_bytes = record.split(b"\t", 1)
            mode, kind, object_id = header.decode("ascii").split(" ", 2)
            relative = path_bytes.decode("utf-8")
        except (ValueError, UnicodeDecodeError) as error:
            raise PolicyMismatch("Git tree contains an unreadable path") from error
        if kind != "blob" or mode not in {"100644", "100755"}:
            raise PolicyMismatch(f"source tree contains a non-regular tracked entry: {relative}")
        if not relative or Path(relative).is_absolute() or ".." in Path(relative).parts:
            raise PolicyMismatch(f"source tree contains an unsafe path: {relative!r}")
        entries.append({"path": relative, "mode": mode, "object": object_id})
    if not entries:
        raise PolicyMismatch("source tree is empty")
    entries.sort(key=lambda item: item["path"])
    return tuple(entries)


def _git_tree_payload(source: Path) -> bytes:
    try:
        raw = subprocess.run(
            ["git", "ls-tree", "-r", "--full-tree", "-z", "HEAD"],
            cwd=source,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        ).stdout
    except (OSError, subprocess.CalledProcessError) as error:
        raise PolicyMismatch(f"cannot inspect the complete source tree: {source}: {error}") from error
    return raw


def _git_blob_sha1(payload: bytes) -> str:
    return hashlib.sha1(f"blob {len(payload)}\0".encode("ascii") + payload).hexdigest()


def _tracked_materialization(
    source: Path,
    entries: Sequence[Mapping[str, str]],
    allowed_overlays: Sequence[ArtifactPin] = (),
) -> dict[str, Any]:
    """Verify every tracked blob is materialized exactly once.

    Git's normal status intentionally hides ignored files and sparse checkouts
    can retain the HEAD/tree identity while omitting blobs.  The native trust
    boundary cannot allow either condition: every tracked path is read from
    disk and checked against its Git blob, while any dirty, untracked, ignored,
    or symlinked worktree entry is rejected.
    """

    status = _git(source, "status", "--porcelain=v1", "--ignored=matching", "--untracked-files=all")
    allowed_by_path = {
        item.relative_path: item
        for item in allowed_overlays
        if item.relative_path == "Cargo.lock" and item.complete()
    }
    if status:
        for line in status.splitlines():
            if len(line) < 4:
                raise PolicyMismatch(f"source checkout has a malformed status entry: {line!r}")
            relative = line[3:]
            if relative.startswith('"') and relative.endswith('"'):
                relative = relative[1:-1]
            pin = allowed_by_path.get(relative.rstrip("/"))
            if pin is None:
                raise PolicyMismatch(f"source checkout has a dirty, ignored, or untracked overlay: {line}")
            digest, size = _hash_file(source / relative, "pinned root Cargo.lock")
            if digest != pin.sha256 or size != pin.bytes:
                raise PolicyMismatch("ignored root Cargo.lock differs from its explicit policy pin")
    skip_worktree = _git(source, "ls-files", "-v", "-z")
    # ``git ls-files -v`` marks skip-worktree entries with a lower-case status
    # letter.  The text mode is safe for the repository's UTF-8 tracked paths;
    # malformed records are rejected rather than silently omitted.
    for record in skip_worktree.split("\0"):
        if not record:
            continue
        if len(record) < 2 or not record[1:]:
            raise PolicyMismatch("git index contains a malformed tracked-path record")
        if record[0].islower():
            raise PolicyMismatch(f"source checkout omits a skip-worktree path: {record[1:]}")
    expected_paths = {str(item["path"]) for item in entries}
    observed_paths: set[str] = set()
    total_bytes = 0
    digest = hashlib.sha256()
    for item in entries:
        relative = str(item["path"])
        path = source / relative
        _reject_symlink_components(path, f"tracked source path {relative}")
        if path.is_symlink() or not path.is_file():
            raise PolicyMismatch(f"tracked source blob is not materialized as a regular file: {relative}")
        try:
            payload = path.read_bytes()
        except OSError as error:
            raise PolicyMismatch(f"cannot read tracked source blob: {relative}: {error}") from error
        if _git_blob_sha1(payload) != item["object"]:
            raise PolicyMismatch(f"tracked source blob differs on disk: {relative}")
        observed_paths.add(relative)
        total_bytes += len(payload)
        digest.update(f"{item['mode']} {item['object']} {relative}\0".encode("utf-8"))
    if observed_paths != expected_paths:
        raise PolicyMismatch("materialized source path set differs from the complete Git tree")
    return {"files": len(entries), "bytes": total_bytes, "sha256": digest.hexdigest()}


def _archive_observation(path: Path, label: str, tree_sha256: str) -> dict[str, Any]:
    digest, size = _hash_file(path, label)
    return {
        "path": str(path.resolve()),
        "sha256": digest,
        "bytes": size,
        "tree_sha256": tree_sha256,
        "tree_verified": False,
    }


def _archive_tree_binding(source: Path, archive: Path, entries: Sequence[Mapping[str, str]], tree_sha256: str) -> dict[str, Any]:
    """Verify a tar archive's regular members against the tracked Git blobs."""

    expected = {str(item["path"]): item for item in entries}
    expected_dirs = {""}
    for relative in expected:
        parent = Path(relative).parent
        while str(parent) not in {"", "."}:
            expected_dirs.add(str(parent).replace("\\", "/"))
            parent = parent.parent
    seen: set[str] = set()
    files: set[str] = set()
    prefix: str | None = None
    pending_directories: list[str] = []
    try:
        with tarfile.open(archive, "r|*") as stream:
            for member in stream:
                name = member.name.replace("\\", "/")
                if name.startswith("/") or ".." in Path(name).parts:
                    raise PolicyMismatch(f"source archive contains an unsafe member: {name}")
                name = "/".join(part for part in name.split("/") if part not in {"", "."})
                if name in seen:
                    raise PolicyMismatch(f"source archive contains a duplicate member: {name}")
                seen.add(name)
                if member.issym() or member.islnk() or not (member.isdir() or member.isfile()):
                    raise PolicyMismatch(f"source archive contains a non-regular member: {name}")
                if prefix is None:
                    if member.isdir():
                        pending_directories.append(name)
                        continue
                    if name in expected:
                        prefix = ""
                    else:
                        candidate, separator, relative = name.partition("/")
                        if not separator or relative not in expected:
                            raise PolicyMismatch("source archive has no reviewed root prefix")
                        prefix = candidate
                    for directory in pending_directories:
                        if prefix:
                            marker = prefix + "/"
                            if directory == prefix:
                                relative_directory = ""
                            elif directory.startswith(marker):
                                relative_directory = directory[len(marker):]
                            else:
                                raise PolicyMismatch(f"source archive directory is outside its root prefix: {directory}")
                        else:
                            relative_directory = directory
                        if relative_directory and relative_directory not in expected_dirs:
                            raise PolicyMismatch(f"source archive contains an unexpected directory: {relative_directory}")
                    pending_directories.clear()
                if prefix:
                    marker = prefix + "/"
                    if name == prefix:
                        relative = ""
                    elif name.startswith(marker):
                        relative = name[len(marker):]
                    else:
                        raise PolicyMismatch(f"source archive member is outside its root prefix: {name}")
                else:
                    relative = name
                if not relative:
                    if not member.isdir():
                        raise PolicyMismatch(f"source archive root member is not a directory: {name}")
                    continue
                if member.isdir():
                    if relative not in expected_dirs:
                        raise PolicyMismatch(f"source archive contains an unexpected directory: {relative}")
                    continue
                if relative not in expected:
                    raise PolicyMismatch(f"source archive contains an unexpected file: {relative}")
                expected_item = expected[relative]
                if bool(member.mode & 0o111) != (expected_item["mode"] == "100755"):
                    raise PolicyMismatch(f"source archive executable mode differs: {relative}")
                handle = stream.extractfile(member)
                if handle is None:
                    raise PolicyMismatch(f"source archive member is unreadable: {relative}")
                hasher = hashlib.sha1(f"blob {member.size}\0".encode("ascii"))
                for block in iter(lambda: handle.read(1024 * 1024), b""):
                    hasher.update(block)
                if hasher.hexdigest() != expected_item["object"]:
                    raise PolicyMismatch(f"source archive blob differs: {relative}")
                files.add(relative)
    except (OSError, tarfile.TarError) as error:
        raise PolicyMismatch(f"cannot inspect source archive {archive}: {error}") from error
    if prefix is None or files != set(expected):
        raise PolicyMismatch("source archive file set differs from the complete Git tree")
    archive_identity = _archive_observation(archive, "source archive", tree_sha256)
    archive_identity["tree_verified"] = True
    archive_identity["files"] = len(expected)
    return archive_identity


def observe_source(source: Path, side: str, archive: Path | None = None) -> dict[str, Any]:
    """Observe a source without allowing a pending policy to be bypassed.

    This function is the source-only path.  It records exact observed hashes
    even for a pending candidate so that a later review can diagnose why the
    policy remains blocked.  It never turns those observations into trusted
    pins.
    """

    if side not in {"baseline", "candidate"}:
        raise PolicyError(f"unknown policy side: {side}")
    root = _safe_source(source, f"{side} source")
    head = _git(root, "rev-parse", "HEAD")
    tree = _git(root, "rev-parse", "HEAD^{tree}")
    entries = _git_tree_entries(root)
    listing = _git_tree_payload(root)
    pin = TRUSTED_POLICY.baseline if side == "baseline" else TRUSTED_POLICY.candidate
    tracked_tree = _tracked_materialization(root, entries, (pin.lock,))
    artifacts: dict[str, Any] = {}
    # The candidate may have no frozen probe/fixture pins.  We still report
    # the paths, but do not call _pin_file with a missing expected identity.
    for name, item in (
        ("lock", pin.lock),
        ("toolchain", pin.toolchain),
        ("cargo_config", pin.cargo_config),
        ("probe", pin.probe),
    ):
        path = root / item.relative_path
        if path.is_file() and not path.is_symlink():
            digest, size = _hash_file(path, f"{side} {name}")
            artifacts[name] = {
                "relative_path": item.relative_path,
                "path": str(path),
                "sha256": digest,
                "bytes": size,
            }
        else:
            artifacts[name] = {"relative_path": item.relative_path, "path": str(path), "missing": True}
    fixtures: list[dict[str, Any]] = []
    for item in pin.fixtures:
        path = root / item.relative_path
        if path.is_file() and not path.is_symlink():
            digest, size = _hash_file(path, f"{side} fixture")
            fixtures.append({"relative_path": item.relative_path, "path": str(path), "sha256": digest, "bytes": size})
        else:
            fixtures.append({"relative_path": item.relative_path, "path": str(path), "missing": True})
    archive_identity = None
    if archive is not None:
        archive_path = Path(archive).absolute()
        _reject_symlink_components(archive_path, f"{side} source archive")
        if archive_path.is_symlink() or not archive_path.is_file():
            raise PolicyMismatch(f"{side} source archive is not a regular file: {archive_path}")
        archive_root = archive_path.resolve()
        if archive_root == root or root in archive_root.parents or archive_root in root.parents:
            raise PolicyMismatch(f"{side} source archive overlaps its source root")
        archive_identity = _archive_observation(archive_path, f"{side} source archive", hashlib.sha256(listing).hexdigest())
    return {
        "side": side,
        "source_root": str(root),
        "commit": head,
        "git_tree": tree,
        "tree_sha256": hashlib.sha256(listing).hexdigest(),
        "tracked_tree": tracked_tree,
        "clean": True,
        "status": "",
        "artifacts": artifacts,
        "fixtures": fixtures,
        "archive": archive_identity,
    }


def validate_source(source: Path, side: str, archive: Path | None = None) -> dict[str, Any]:
    """Validate one endpoint against its fixed policy pins."""

    policy_pin = TRUSTED_POLICY.baseline if side == "baseline" else TRUSTED_POLICY.candidate
    if not policy_pin.complete():
        raise PolicyIncomplete(f"{side} source pins are not complete in current-source policy v2")
    observed = observe_source(source, side, archive)
    if observed["commit"] != policy_pin.commit:
        raise PolicyMismatch(f"{side} source commit differs from the trusted policy")
    if observed["git_tree"] != policy_pin.git_tree or observed["tree_sha256"] != policy_pin.tree_sha256:
        raise PolicyMismatch(f"{side} source tree differs from the trusted policy")
    if not observed["clean"]:
        raise PolicyMismatch(f"{side} source checkout is dirty")
    for name, item in (
        ("lock", policy_pin.lock),
        ("toolchain", policy_pin.toolchain),
        ("cargo_config", policy_pin.cargo_config),
        ("probe", policy_pin.probe),
    ):
        _pin_file(Path(observed["source_root"]), item, f"{side} {name}")
    for item in policy_pin.fixtures:
        _pin_file(Path(observed["source_root"]), item, f"{side} fixture {item.relative_path}")
    if archive is None:
        raise PolicyMismatch(f"{side} source archive is required by the trusted policy")
    archive_path = Path(archive).absolute()
    source_root = Path(observed["source_root"]).resolve()
    archive_root = archive_path.resolve(strict=False)
    if archive_root == source_root or source_root in archive_root.parents or archive_root in source_root.parents:
        raise PolicyMismatch(f"{side} source archive overlaps its source root")
    expected_archive = policy_pin.archive
    if not expected_archive.complete():
        raise PolicyIncomplete(f"{side} source archive pin is not complete")
    archive_identity = _archive_tree_binding(
        Path(observed["source_root"]),
        archive_path,
        _git_tree_entries(Path(observed["source_root"])),
        observed["tree_sha256"],
    )
    if archive_identity["sha256"] != expected_archive.sha256 or archive_identity["bytes"] != expected_archive.bytes:
        raise PolicyMismatch(f"{side} source archive differs from the trusted policy")
    observed["archive"] = archive_identity
    return observed


def validate_lock(path: Path, side: str = "candidate") -> dict[str, Any]:
    """Validate a lock using only the fixed policy (never a caller digest)."""

    if side not in {"baseline", "candidate"}:
        raise PolicyError(f"unknown lock side: {side}")
    TRUSTED_POLICY.require_execution_ready()
    pin = TRUSTED_POLICY.baseline.lock if side == "baseline" else TRUSTED_POLICY.candidate.lock
    digest, size = _hash_file(Path(path), f"{side} Cargo.lock")
    if digest != pin.sha256 or size != pin.bytes:
        raise PolicyMismatch(f"{side} Cargo.lock differs from the trusted current-source policy")
    return {"path": str(Path(path).resolve()), "sha256": digest, "bytes": size, "side": side}


def validate_routes(lane: str, routes: Sequence[Mapping[str, Any]]) -> None:
    """Check exact ordered route identities, including output destinations."""

    expected = TRUSTED_POLICY.routes_for(lane)
    if len(routes) != len(expected):
        raise PolicyMismatch(f"{lane} route count differs from the trusted policy")
    for index, (actual, wanted) in enumerate(zip(routes, expected)):
        if not isinstance(actual, Mapping):
            raise PolicyMismatch(f"{lane} route {index} is not an object")
        for field in ("route", "kind", "program", "output", "role", "status"):
            if actual.get(field) != getattr(wanted, field):
                raise PolicyMismatch(f"{lane} route {index} {field} differs from the trusted policy")


def _check_plan_paths(root: Path, state_dir: Path, target: Path, source_archive: Path | None) -> None:
    paths = [("source", root), ("state output", state_dir), ("Cargo target output", target)]
    if source_archive is not None:
        paths.append(("source archive", source_archive))
    resolved: list[tuple[str, Path]] = []
    for label, raw in paths:
        lexical = Path(raw).absolute()
        _reject_symlink_components(lexical, label)
        if label == "source" and (lexical.is_symlink() or not lexical.is_dir()):
            raise PolicyMismatch(f"source is not a regular directory: {lexical}")
        resolved.append((label, lexical.resolve(strict=False)))
    for index, (left_label, left) in enumerate(resolved):
        for right_label, right in resolved[index + 1 :]:
            if left == right or left in right.parents or right in left.parents:
                raise PolicyMismatch(f"PPT v2 source/output/archive paths overlap: {left_label} and {right_label}")


def make_ppt_steps_v2(
    root: Path,
    state_dir: Path,
    target: Path,
    *,
    source_archive: Path | None = None,
) -> list[dict[str, Any]]:
    """Build the opt-in PPT auxiliary plan after the fixed policy is ready.

    The old ``make_ppt_steps`` is deliberately not imported or changed.  The
    lock guard below has no expected-digest argument; it reads the reviewed
    policy itself and therefore cannot be redirected by a manifest caller.
    """

    TRUSTED_POLICY.require_execution_ready()
    root = Path(root)
    state_dir = Path(state_dir)
    target = Path(target)
    _check_plan_paths(root, state_dir, target, source_archive)
    if source_archive is None:
        raise PolicyMismatch("PPT v2 execution requires the verified candidate source archive")
    # Do the full candidate check at plan construction time as well as at
    # receipt consumption.  A plan must never be able to pair a later binary
    # with a merely named source root or archive.
    validate_source(root, "candidate", source_archive)
    lock = root / "Cargo.lock"
    guard = Path(__file__).resolve()
    host_guard = root / "docs/unified-ops/results/0362g/ppt_host_process_guard.py"
    fixture = root / "test-data/unified-ops/0294/video.ppt"
    binary = target / TARGET / "release/examples/scoped_patch_correctness"
    receipt = state_dir / "ppt-scoped-correctness.json"
    guarded = [sys.executable, str(guard), "guard-lock", str(lock), "--"]
    correctness = (
        "set -eu; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-ppt --all-features --lib scoped::tests -- --test-threads=1; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-ppt --all-features --lib -- --test-threads=1; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-cfb --all-features --lib -- --test-threads=1; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-ppt --all-features --tests -- --test-threads=1; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-cfb --all-features --tests -- --test-threads=1; "
        f"cargo clippy --offline --locked --target {TARGET} --release -p litchi-ppt --all-features --all-targets -- -D warnings; "
        f"cargo clippy --offline --locked --target {TARGET} --release -p litchi-cfb --all-features --all-targets -- -D warnings"
    )

    def direct(name: str, argv: list[str], *, env: dict[str, str] | None = None) -> dict[str, Any]:
        return {"name": name, "kind": "argv", "argv": argv, "isolated": True, "env": env or {}}

    return [
        direct(
            "ppt-current-scoped-build",
            [*guarded, str(host_guard), "--", "cargo", "build", "--offline", "--locked", "--target", TARGET, "--release", "--all-features", "-p", "litchi-ppt", "--example", "scoped_patch_probe", "--example", "scoped_patch_correctness"],
        ),
        direct("ppt-current-correctness-suite", [*guarded, "sh", "-c", correctness], env={"RUST_TEST_THREADS": "1"}),
        direct("ppt-current-correctness", [sys.executable, str(host_guard), "--", str(binary), str(fixture), str(receipt)]),
        direct("ppt-current-correctness-validate", [sys.executable, str(root / "docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py"), str(receipt), str(fixture)]),
    ]


def _cli_guard_lock(path: Path, command: Sequence[str]) -> int:
    """CLI implementation for the new PPT guard; no expected hash input."""

    validate_lock(path, "candidate")
    if not command:
        print(TRUSTED_POLICY.candidate.lock.sha256)
        return 0
    completed = subprocess.run(list(command), check=False)
    return int(completed.returncode)


def main(argv: Sequence[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] != "guard-lock":
        print("usage: current_native_policy_v2.py guard-lock CARGO_LOCK [-- COMMAND ...]", file=sys.stderr)
        return 2
    try:
        separator = args.index("--")
    except ValueError:
        separator = len(args)
    if len(args) < 2 or separator != 2:
        print("usage: current_native_policy_v2.py guard-lock CARGO_LOCK [-- COMMAND ...]", file=sys.stderr)
        return 2
    try:
        return _cli_guard_lock(Path(args[1]), args[separator + 1 :])
    except (OSError, PolicyError) as error:
        print(f"current-source policy guard refused: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
