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
import shlex
import subprocess
import sys
import tarfile
from types import MappingProxyType
from typing import Any, Callable, Iterable, Mapping, Sequence


POLICY_SCHEMA = "litchi.current-native-policy.v2"
PREPARATION_SCHEMA = "litchi.current-native-preparation.v2"
QUALIFICATION_SCHEMA = "litchi.current-native-qualification.v2"
PROTOCOL_SCHEMA = "litchi.native-evidence-protocol.v2"
EXECUTION_SCHEMA = "litchi.current-native-execution.v2"
RUNNER_KIND = "current-native-v2"
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

# A current-source run has three independent dependency consumers.  Keeping
# their locks in one ordered tuple makes the policy immutable while still
# allowing every consumer to select its own role without accepting a path or
# digest from a receipt.
ROLE_NAMES = ("workspace", "perf", "native")
ROLE_LOCK_PATHS = MappingProxyType({
    "workspace": "Cargo.lock",
    "perf": "tools/perf-baseline/Cargo.lock",
    "native": "tools/native-resave/Cargo.lock",
})
SOURCE_MANIFEST_PATHS = (
    "Cargo.toml",
    "tools/perf-baseline/Cargo.toml",
    "tools/native-resave/Cargo.toml",
)
PPT_PROBE_PATH = "crates/litchi-ppt/examples/scoped_patch_probe.rs"
PPT_CORRECTNESS_PATH = "crates/litchi-ppt/examples/scoped_patch_correctness.rs"
LEGACY_PROBE_PATH = "tools/perf-baseline/examples/refusal_contract_bench.rs"

# The frozen f2 source predates both the standalone RTF refusal consumer and
# the outside-clock PPT correctness companion.  They are reviewed consumer
# adapters, not production-source overlays.  Their endpoint paths are fixed
# here so a preparation caller cannot substitute a source file with the same
# basename from a later candidate checkout.
ADAPTER_SOURCE_PATHS = MappingProxyType({
    "tools/perf-baseline/Cargo.toml": "tools/native_qualification/adapters/rtf-baseline-Cargo.toml",
    LEGACY_PROBE_PATH: "tools/native_qualification/adapters/rtf-refusal_contract_bench.rs",
    PPT_CORRECTNESS_PATH: "tools/native_qualification/adapters/ppt-scoped_patch_correctness-f2.rs",
})

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
# The auxiliary PPT workload has one four-leg ABBA sequence.  The historical
# RTF consumer contract is deliberately stronger: it has five independent
# blocks, each with the reviewed four-leg order (31 samples per route in the
# native child).  Keep the block count in this protocol rather than silently
# collapsing the legacy workload to the generic one-block shape.
ABBA_BLOCKS = MappingProxyType({"ppt-auxiliary": 1, "legacy-current": 5})
ADAPTER_FEATURES = MappingProxyType({
    "ppt-auxiliary": MappingProxyType({"baseline": ("all-features",), "candidate": ("all-features",)}),
    "legacy-current": MappingProxyType({"baseline": ("rtf-legacy-api",), "candidate": ("rtf-current-api",)}),
})

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
    role_locks: tuple[ArtifactPin, ...]
    source_manifests: tuple[ArtifactPin, ...]
    toolchain: ArtifactPin
    cargo_config: ArtifactPin
    probes: tuple[ArtifactPin, ...]
    fixtures: tuple[ArtifactPin, ...]
    archive: ArtifactPin
    target: str
    profile: str
    features: tuple[str, ...]
    adapter_overlays: tuple[ArtifactPin, ...] = ()

    def complete(self) -> bool:
        return (
            isinstance(self.commit, str)
            and FULL_SHA.fullmatch(self.commit) is not None
            and isinstance(self.git_tree, str)
            and FULL_SHA.fullmatch(self.git_tree) is not None
            and isinstance(self.tree_sha256, str)
            and SHA256.fullmatch(self.tree_sha256) is not None
            and len(self.role_locks) == len(ROLE_NAMES)
            and all(
                item.complete() and item.relative_path == ROLE_LOCK_PATHS[role]
                for role, item in zip(ROLE_NAMES, self.role_locks)
            )
            and len(self.source_manifests) == len(SOURCE_MANIFEST_PATHS)
            and all(
                item.complete() and item.relative_path == path
                for item, path in zip(self.source_manifests, SOURCE_MANIFEST_PATHS)
            )
            and self.toolchain.complete()
            and self.cargo_config.complete()
            and {item.relative_path for item in self.probes} >= {
                PPT_PROBE_PATH,
                LEGACY_PROBE_PATH,
                PPT_CORRECTNESS_PATH,
            }
            and all(item.complete() for item in self.probes)
            and bool(self.fixtures)
            and all(item.complete() for item in self.fixtures)
            and self.archive.complete()
            and self.target == TARGET
            and self.profile == PROFILE
            and self.features == FEATURES
            and all(item.complete() for item in self.adapter_overlays)
            and len({item.relative_path for item in self.adapter_overlays}) == len(self.adapter_overlays)
            and all(item.relative_path in ADAPTER_SOURCE_PATHS for item in self.adapter_overlays)
        )

    def json(self) -> dict[str, Any]:
        return {
            "commit": self.commit,
            "git_tree": self.git_tree,
            "tree_sha256": self.tree_sha256,
            "role_locks": {
                role: item.json()
                for role, item in zip(ROLE_NAMES, self.role_locks)
            },
            "source_manifests": [item.json() for item in self.source_manifests],
            "toolchain": self.toolchain.json(),
            "cargo_config": self.cargo_config.json(),
            "probes": [item.json() for item in self.probes],
            "fixtures": [item.json() for item in self.fixtures],
            "archive": self.archive.json(),
            "target": self.target,
            "profile": self.profile,
            "features": list(self.features),
            "adapter_overlays": [item.json() for item in self.adapter_overlays],
        }

    def probe_for(self, relative_path: str) -> ArtifactPin:
        for item in self.probes:
            if item.relative_path == relative_path:
                return item
        raise PolicyIncomplete(f"source pin has no probe for {relative_path}")

    def lock_for_role(self, role: str) -> ArtifactPin:
        try:
            index = ROLE_NAMES.index(role)
        except ValueError as error:
            raise PolicyError(f"unknown source role: {role}") from error
        try:
            return self.role_locks[index]
        except IndexError as error:
            raise PolicyIncomplete(f"{role} lock pin is not present") from error

    def adapter_for(self, relative_path: str) -> ArtifactPin | None:
        for item in self.adapter_overlays:
            if item.relative_path == relative_path:
                return item
        return None


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
    role_locks=(
        _artifact("Cargo.lock", BASELINE_LOCK_SHA256, BASELINE_LOCK_BYTES),
        _artifact("tools/perf-baseline/Cargo.lock", "57c2a19ffeff395fdea95c33f295c85b08400b00e2fd9cec94d66cff56d1583f", 57_028),
        _artifact("tools/native-resave/Cargo.lock", "6aaa0e3800a4db72afbe8783aab382accdb4a6355d06bb97fb1c12dd78839408", 56_929),
    ),
    source_manifests=(
        _artifact("Cargo.toml", "66a8f4d2b0ffb6b9ba4eb87d84fee8a3049e8ef7e4380b6ee48acaaf372dd7b5", 10_608),
        _artifact("tools/perf-baseline/Cargo.toml", "9e9d64678dafa5e96034f9f6b5cdb1fca28a65cc7a30adb83c42d6d4ad8b34cd", 3_478),
        _artifact("tools/native-resave/Cargo.toml", "14a3b0fbf2da037af795d4dc820ed5d91ccbc2d4f5f009f7bb2182c25849fa8b", 1_425),
    ),
    toolchain=_artifact("rust-toolchain.toml", TOOLCHAIN_SHA256, TOOLCHAIN_BYTES),
    cargo_config=_artifact(".cargo/config.toml", CARGO_CONFIG_SHA256, CARGO_CONFIG_BYTES),
    probes=(
        _artifact("crates/litchi-ppt/examples/scoped_patch_probe.rs", "0350fc1696688e576613bc8b2c2ccb3d39c12b4b8ad890c5dcd14e3e7b4138ca", 5_158),
        _artifact(LEGACY_PROBE_PATH, "0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81", 12_782),
        _artifact(PPT_CORRECTNESS_PATH, "032e74e081d2b442df49ce5bd7fb2edc66b89a4a1d5368872dd7bfeb94c8ee7a", 3_168),
    ),
    fixtures=(
        _artifact("test-data/unified-ops/0294/video.ppt", "f32f0614958a177a5195cd7537cdfd2efef52c57de4decc13cf65eadc66cdb8a", 43_520),
    ),
    archive=_artifact("source-archive.tar.zst", BASELINE_ARCHIVE_SHA256, BASELINE_ARCHIVE_BYTES),
    target=TARGET,
    profile=PROFILE,
    features=FEATURES,
    adapter_overlays=(
        _artifact("tools/perf-baseline/Cargo.toml", "af876fbf715edb643d2c7abb8643efb252413c640824890e90b1849207889676", 3_638),
        _artifact(LEGACY_PROBE_PATH, "0613f279db17121711b140c273137465c23670033a4dda5392444703b01b1f81", 12_782),
        _artifact(PPT_CORRECTNESS_PATH, "032e74e081d2b442df49ce5bd7fb2edc66b89a4a1d5368872dd7bfeb94c8ee7a", 3_168),
    ),
)

# The candidate is intentionally not represented by a guessed hash.  These
# None values are an executable part of the policy: ``require_execution_ready``
# must continue to reject the run until a reviewed commit replaces this
# constant with exact pins and updates the record.
_CANDIDATE_PIN = SourcePin(
    commit=None,
    git_tree=None,
    tree_sha256=None,
    role_locks=tuple(_artifact(ROLE_LOCK_PATHS[role], None, None) for role in ROLE_NAMES),
    source_manifests=tuple(_artifact(path, None, None) for path in SOURCE_MANIFEST_PATHS),
    toolchain=_artifact("rust-toolchain.toml", TOOLCHAIN_SHA256, TOOLCHAIN_BYTES),
    cargo_config=_artifact(".cargo/config.toml", CARGO_CONFIG_SHA256, CARGO_CONFIG_BYTES),
    probes=(
        _artifact(PPT_PROBE_PATH, None, None),
        _artifact(LEGACY_PROBE_PATH, None, None),
        _artifact(PPT_CORRECTNESS_PATH, None, None),
    ),
    # The path is reviewed even while the payload identity is pending.  This
    # lets manifests and diagnostics name the only admitted fixture without
    # making the candidate executable.
    fixtures=(_artifact("test-data/unified-ops/0294/video.ppt", None, None),),
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
                "abba_blocks": dict(ABBA_BLOCKS),
                "adapter_features": {
                    lane: {
                        ("before" if side == "baseline" else "after"): list(features)
                        for side, features in sides.items()
                    }
                    for lane, sides in ADAPTER_FEATURES.items()
                }
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

    def manifest_binding(self) -> dict[str, Any]:
        """Return the fixed manifest-side identity projection.

        The manifests intentionally carry no candidate source identity while
        the freeze is pending.  This projection is derived from the trusted
        code object, not from manifest or command-line fields, and therefore
        makes a copied or caller-rewritten route manifest fail closed.
        """

        baseline_locks = {
            role: {
                "baseline": {
                    "bytes": self.baseline.lock_for_role(role).bytes,
                    "sha256": self.baseline.lock_for_role(role).sha256,
                },
                "candidate": (
                    {
                        "bytes": self.candidate.lock_for_role(role).bytes,
                        "sha256": self.candidate.lock_for_role(role).sha256,
                    }
                    if self.candidate.lock_for_role(role).complete()
                    else None
                ),
                "path": ROLE_LOCK_PATHS[role],
            }
            for role in ROLE_NAMES
        }
        return {
            "candidate_status": "bound" if self.status == "ready" else "pending_final_source_freeze",
            "fixture_paths": [item.relative_path for item in self.candidate.fixtures],
            "role_locks": baseline_locks,
            "probe_paths": [item.relative_path for item in self.candidate.probes],
            "source": {
                "archive": self.candidate.archive.json() if self.candidate.archive.complete() else None,
                "commit": self.candidate.commit if self.candidate.complete() else None,
                "git_tree": self.candidate.git_tree if self.candidate.complete() else None,
                "tree_sha256": self.candidate.tree_sha256 if self.candidate.complete() else None,
            },
            "source_manifests": list(SOURCE_MANIFEST_PATHS),
            "source_adapters": {
                "baseline": [item.json() for item in self.baseline.adapter_overlays],
                "candidate": [item.json() for item in self.candidate.adapter_overlays],
            },
            "toolchain": {
                "features": list(self.candidate.features),
                "adapter_features": {
                    lane: {
                        ("before" if side == "baseline" else "after"): list(features)
                        for side, features in sides.items()
                    }
                    for lane, sides in ADAPTER_FEATURES.items()
                },
                "paths": [self.candidate.toolchain.relative_path, self.candidate.cargo_config.relative_path],
                "profile": self.candidate.profile,
                "target": self.candidate.target,
            },
            "abba_blocks": dict(ABBA_BLOCKS),
        }

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


def _reject_protocol_source(path: Path, label: str) -> None:
    """Keep executable source endpoints outside this protocol checkout."""

    source = Path(path).absolute().resolve(strict=False)
    protocol = Path(__file__).resolve().parents[2]
    if source == protocol or source in protocol.parents or protocol in source.parents:
        raise PolicyMismatch(f"{label} overlaps the protocol checkout: {source}")


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


def _git_blob_payload(source: Path, relative: str) -> bytes:
    """Read one immutable HEAD blob without trusting its worktree copy."""

    path = Path(relative)
    if path.is_absolute() or ".." in path.parts or not relative:
        raise PolicyMismatch(f"unsafe Git blob path: {relative}")
    try:
        completed = subprocess.run(
            ["git", "show", f"HEAD:{relative}"], cwd=source, check=True,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        raise PolicyMismatch(f"source HEAD lacks required tracked blob: {relative}") from error
    return completed.stdout


def _validate_git_pin(source: Path, pin: ArtifactPin, label: str) -> dict[str, Any]:
    """Validate the pristine Git blob behind an adapter-modified path."""

    if not pin.complete():
        raise PolicyIncomplete(f"{label} has no exact original-source pin")
    payload = _git_blob_payload(source, pin.relative_path)
    digest = hashlib.sha256(payload).hexdigest()
    if digest != pin.sha256 or len(payload) != pin.bytes:
        raise PolicyMismatch(
            f"{label} original Git blob differs from the trusted source: "
            f"expected {pin.sha256}/{pin.bytes}, observed {digest}/{len(payload)}"
        )
    return {
        "relative_path": pin.relative_path,
        "sha256": digest,
        "bytes": len(payload),
        "git_blob": _git_blob_sha1(payload),
    }


def _validate_adapter_source(adapter: ArtifactPin, label: str) -> dict[str, Any]:
    """Bind an endpoint overlay to the reviewed protocol adapter source."""

    try:
        relative = ADAPTER_SOURCE_PATHS[adapter.relative_path]
    except KeyError as error:
        raise PolicyMismatch(f"{label} is not an admitted adapter path") from error
    protocol_root = Path(__file__).resolve().parents[2]
    source = protocol_root / relative
    digest, size = _hash_file(source, f"{label} source")
    if digest != adapter.sha256 or size != adapter.bytes:
        raise PolicyMismatch(
            f"{label} protocol adapter differs from its fixed identity: "
            f"expected {adapter.sha256}/{adapter.bytes}, observed {digest}/{size}"
        )
    return {
        "relative_path": adapter.relative_path,
        "source_path": str(source.resolve()),
        "sha256": digest,
        "bytes": size,
    }


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
    *,
    require_overlays: bool = False,
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
        if item.complete()
    }
    overlay_paths: set[str] = set()
    status_rows: list[dict[str, Any]] = []
    if status:
        for line in status.splitlines():
            if len(line) < 4:
                raise PolicyMismatch(f"source checkout has a malformed status entry: {line!r}")
            relative = line[3:]
            if relative.startswith('"') and relative.endswith('"'):
                relative = relative[1:-1]
            relative = relative.rstrip("/")
            pin = allowed_by_path.get(relative)
            if pin is None:
                raise PolicyMismatch(f"source checkout has a dirty, ignored, or untracked overlay: {line}")
            digest, size = _hash_file(source / relative, f"pinned overlay {relative}")
            if digest != pin.sha256 or size != pin.bytes:
                raise PolicyMismatch(f"pinned overlay differs from its explicit policy pin: {relative}")
            overlay_paths.add(relative)
            status_rows.append({"path": relative, "status": line[:2], "sha256": digest, "bytes": size})
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
        overlay = allowed_by_path.get(relative)
        if overlay is not None:
            observed_digest = hashlib.sha256(payload).hexdigest()
            if observed_digest != overlay.sha256 or len(payload) != overlay.bytes:
                raise PolicyMismatch(f"tracked adapter overlay differs on disk: {relative}")
            overlay_paths.add(relative)
        elif _git_blob_sha1(payload) != item["object"]:
            raise PolicyMismatch(f"tracked source blob differs on disk: {relative}")
        observed_paths.add(relative)
        total_bytes += len(payload)
        digest.update(f"{item['mode']} {item['object']} {relative}\0".encode("utf-8"))
    if observed_paths != expected_paths:
        raise PolicyMismatch("materialized source path set differs from the complete Git tree")
    expected_overlays = set(allowed_by_path)
    # An allowed root Cargo.lock is ignored and may be absent from status when
    # a caller has not materialized it yet; the later fixed-pin check reports
    # that missing required input.  Adapter overlays, however, are required
    # to be present before an execution receipt can be admitted.
    missing_overlays = sorted(
        path for path in expected_overlays
        if path != "Cargo.lock" and not (source / path).is_file()
    )
    if missing_overlays and require_overlays:
        raise PolicyMismatch(f"required adapter overlays are missing: {', '.join(missing_overlays)}")
    return {
        "files": len(entries),
        "bytes": total_bytes,
        "sha256": digest.hexdigest(),
        "overlays": status_rows,
    }


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
    # The ignored workspace lock and the baseline's reviewed consumer adapters
    # are the only admitted endpoint overlays.  They are fixed policy objects,
    # never caller-selected paths or hashes.
    workspace_lock = pin.lock_for_role("workspace")
    tracked_tree = _tracked_materialization(root, entries, (workspace_lock, *pin.adapter_overlays))
    artifacts: dict[str, Any] = {}
    # The candidate may have no frozen identities.  We still report every
    # reviewed path, but never turn an observation into a trusted pin.
    role_locks: dict[str, Any] = {}
    for role in ROLE_NAMES:
        item = pin.lock_for_role(role)
        path = root / item.relative_path
        if path.is_file() and not path.is_symlink():
            digest, size = _hash_file(path, f"{side} {role} lock")
            role_locks[role] = {
                "relative_path": item.relative_path,
                "path": str(path),
                "sha256": digest,
                "bytes": size,
            }
        else:
            role_locks[role] = {"relative_path": item.relative_path, "path": str(path), "missing": True}
    artifacts["role_locks"] = role_locks
    # Preserve a workspace-lock view in the observation for consumers of the
    # earlier diagnostic receipt shape; it is derived from the role map and is
    # never used as a caller-selected identity.
    artifacts["lock"] = role_locks["workspace"]
    for name, item in (("toolchain", pin.toolchain), ("cargo_config", pin.cargo_config)):
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
    manifests: list[dict[str, Any]] = []
    for item in pin.source_manifests:
        path = root / item.relative_path
        if path.is_file() and not path.is_symlink():
            digest, size = _hash_file(path, f"{side} source manifest")
            manifests.append({"relative_path": item.relative_path, "path": str(path), "sha256": digest, "bytes": size})
        else:
            manifests.append({"relative_path": item.relative_path, "path": str(path), "missing": True})
    artifacts["source_manifests"] = manifests
    probes: list[dict[str, Any]] = []
    for item in pin.probes:
        path = root / item.relative_path
        if path.is_file() and not path.is_symlink():
            digest, size = _hash_file(path, f"{side} probe")
            probes.append({"relative_path": item.relative_path, "path": str(path), "sha256": digest, "bytes": size})
        else:
            probes.append({"relative_path": item.relative_path, "path": str(path), "missing": True})
    artifacts["probes"] = probes
    artifacts["adapter_overlays"] = [
        {
            "relative_path": item.relative_path,
            "path": str(root / item.relative_path),
            "sha256": item.sha256,
            "bytes": item.bytes,
            "source_path": ADAPTER_SOURCE_PATHS[item.relative_path],
        }
        for item in pin.adapter_overlays
    ]
    # Keep the old singular diagnostic key only as a projection of the first
    # fixed probe.  All admission checks below iterate the complete tuple.
    if probes:
        artifacts["probe"] = probes[0]
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
        "status": "adapter-overlays" if tracked_tree["overlays"] else "",
        "artifacts": artifacts,
        "fixtures": fixtures,
        "archive": archive_identity,
    }


def _validate_observed_source(observed: Mapping[str, Any], side: str, policy_pin: SourcePin) -> Path:
    """Validate source/tree and every in-tree artifact, excluding archives."""

    if observed["commit"] != policy_pin.commit:
        raise PolicyMismatch(f"{side} source commit differs from the trusted policy")
    if observed["git_tree"] != policy_pin.git_tree or observed["tree_sha256"] != policy_pin.tree_sha256:
        raise PolicyMismatch(f"{side} source tree differs from the trusted policy")
    if not observed["clean"]:
        raise PolicyMismatch(f"{side} source checkout is dirty")
    source_root = Path(observed["source_root"])
    for role in ROLE_NAMES:
        _pin_file(source_root, policy_pin.lock_for_role(role), f"{side} {role} lock")
    for item in policy_pin.source_manifests:
        overlay = policy_pin.adapter_for(item.relative_path)
        if overlay is not None:
            # The worktree manifest is intentionally different, but the
            # immutable source identity remains the f2 Git blob.  Both sides
            # of that boundary are checked before a build can use the file.
            _validate_git_pin(source_root, item, f"{side} original source manifest {item.relative_path}")
            _pin_file(source_root, overlay, f"{side} adapter source manifest {item.relative_path}")
        else:
            _pin_file(source_root, item, f"{side} source manifest {item.relative_path}")
    for name, item in (("toolchain", policy_pin.toolchain), ("cargo_config", policy_pin.cargo_config)):
        _pin_file(source_root, item, f"{side} {name}")
    for item in policy_pin.probes:
        overlay = policy_pin.adapter_for(item.relative_path)
        if overlay is not None:
            _pin_file(source_root, overlay, f"{side} adapter probe {item.relative_path}")
        else:
            _pin_file(source_root, item, f"{side} probe {item.relative_path}")
    for item in policy_pin.fixtures:
        _pin_file(source_root, item, f"{side} fixture {item.relative_path}")
    for item in policy_pin.adapter_overlays:
        _validate_adapter_source(item, f"{side} adapter {item.relative_path}")
    return source_root


def validate_source(source: Path, side: str, archive: Path | None = None) -> dict[str, Any]:
    """Validate one endpoint against its fixed policy pins."""

    policy_pin = TRUSTED_POLICY.baseline if side == "baseline" else TRUSTED_POLICY.candidate
    if not policy_pin.complete():
        raise PolicyIncomplete(f"{side} source pins are not complete in current-source policy v2")
    _reject_protocol_source(Path(source), f"{side} source")
    observed = observe_source(source, side, archive)
    _validate_observed_source(observed, side, policy_pin)
    source_root = Path(observed["source_root"])
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


def validate_lock(source_root: Path, side: str = "candidate", role: str = "workspace") -> dict[str, Any]:
    """Validate a role lock at its fixed path under a trusted source root.

    The first argument is deliberately a source checkout, never a copied
    lockfile.  Its commit/tree/materialization and all fixed in-tree inputs are
    checked before the role-specific lock is read.
    """

    if side not in {"baseline", "candidate"}:
        raise PolicyError(f"unknown lock side: {side}")
    TRUSTED_POLICY.require_execution_ready()
    if role not in ROLE_NAMES:
        raise PolicyError(f"unknown lock role: {role}")
    policy_pin = TRUSTED_POLICY.baseline if side == "baseline" else TRUSTED_POLICY.candidate
    _reject_protocol_source(Path(source_root), f"{side} source root")
    root = _safe_source(Path(source_root), f"{side} source root")
    observed = observe_source(root, side)
    _validate_observed_source(observed, side, policy_pin)
    pin = policy_pin.lock_for_role(role)
    lock_path = root / ROLE_LOCK_PATHS[role]
    if pin.relative_path != ROLE_LOCK_PATHS[role]:
        raise PolicyMismatch(f"{side} {role} lock policy path is not fixed")
    digest, size = _hash_file(lock_path, f"{side} {role} Cargo.lock")
    if digest != pin.sha256 or size != pin.bytes:
        raise PolicyMismatch(f"{side} {role} lock differs from the trusted current-source policy")
    return {"path": str(lock_path.resolve()), "sha256": digest, "bytes": size, "side": side, "role": role, "source_root": str(root)}


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
    protocol_root = Path(__file__).resolve().parents[2]
    paths = [("source", root), ("state output", state_dir), ("Cargo target output", target), ("protocol checkout", protocol_root)]
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


def _protocol_tool_path(relative: str) -> Path:
    """Resolve a fixed protocol helper, never a file inside an endpoint."""

    protocol_root = Path(__file__).resolve().parents[2]
    path = protocol_root / relative
    _reject_symlink_components(path, f"protocol helper {relative}")
    if path.is_symlink() or not path.is_file():
        raise PolicyMismatch(f"reviewed protocol helper is missing: {path}")
    return path


def _policy_pin(side: str) -> SourcePin:
    if side in {"baseline", "before"}:
        return TRUSTED_POLICY.baseline
    if side in {"candidate", "after"}:
        return TRUSTED_POLICY.candidate
    raise PolicyError(f"unknown policy side: {side}")


def _policy_side(side: str) -> str:
    if side in {"baseline", "before"}:
        return "baseline"
    if side in {"candidate", "after"}:
        return "candidate"
    raise PolicyError(f"unknown policy side: {side}")


def _adapter_features(lane: str, side: str) -> tuple[str, ...]:
    try:
        return tuple(ADAPTER_FEATURES[lane][_policy_side(side)])
    except KeyError as error:
        raise PolicyError(f"no reviewed adapter features for {lane}/{side}") from error


def _role_lock_path(root: Path, role: str, side: str = "candidate") -> Path:
    """Resolve a role lock only from the reviewed source pin."""

    pin = _policy_pin(side).lock_for_role(role)
    if not pin.complete() or pin.relative_path != ROLE_LOCK_PATHS[role]:
        raise PolicyIncomplete(f"{side} {role} lock pin is not complete")
    return root / pin.relative_path


def _guarded_lock(root: Path, role: str, guard: Path, side: str = "candidate") -> list[str]:
    return [sys.executable, str(guard), "guard-role-lock", role, str(root), "--side", side, "--"]


def make_ppt_steps_v2(
    root: Path,
    state_dir: Path,
    target: Path,
    *,
    source_archive: Path | None = None,
    side: str = "candidate",
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
        raise PolicyMismatch(f"PPT v2 {side} execution requires the verified source archive")
    # Do the full candidate check at plan construction time as well as at
    # receipt consumption.  A plan must never be able to pair a later binary
    # with a merely named source root or archive.
    policy_side = _policy_side(side)
    validate_source(root, policy_side, source_archive)
    guard = Path(__file__).resolve()
    host_guard = _protocol_tool_path("docs/unified-ops/results/0362g/ppt_host_process_guard.py")
    ppt_validator = _protocol_tool_path("docs/unified-ops/results/0362g/validate_ppt_correctness_receipt.py")
    fixture_pins = _policy_pin(policy_side).fixtures
    if not fixture_pins:
        raise PolicyIncomplete("candidate fixture pins are not complete")
    fixture = root / fixture_pins[0].relative_path
    binary_paths = [
        target / TARGET / "release/examples/scoped_patch_probe",
        target / TARGET / "release/examples/scoped_patch_correctness",
    ]
    binary = binary_paths[1]
    receipt = state_dir / "ppt-scoped-correctness.json"
    suite_receipt = state_dir / "ppt-correctness-suite.json"
    guarded = _guarded_lock(root, "workspace", guard, policy_side)
    adapter_features = _adapter_features("ppt-auxiliary", side)
    feature_flags = ["--all-features"] if adapter_features == ("all-features",) else [
        "--no-default-features", "--features", ",".join(adapter_features)
    ]
    correctness = (
        "set -eu; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-ppt {' '.join(feature_flags)} --lib scoped::tests -- --test-threads=1; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-ppt {' '.join(feature_flags)} --lib -- --test-threads=1; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-cfb {' '.join(feature_flags)} --lib -- --test-threads=1; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-ppt {' '.join(feature_flags)} --tests -- --test-threads=1; "
        f"cargo test --offline --locked --target {TARGET} --release -p litchi-cfb {' '.join(feature_flags)} --tests -- --test-threads=1; "
        f"cargo clippy --offline --locked --target {TARGET} --release -p litchi-ppt {' '.join(feature_flags)} --all-targets -- -D warnings; "
        f"cargo clippy --offline --locked --target {TARGET} --release -p litchi-cfb {' '.join(feature_flags)} --all-targets -- -D warnings; "
        f"printf '%s\\n' {shlex.quote(json.dumps({'route': 'ppt-scoped-correctness-suite', 'status': 'ok'}, separators=(',', ':')))} > {shlex.quote(str(suite_receipt))}"
    )

    def direct(name: str, argv: list[str], *, route: RoutePin | None = None, phase: str = "route", env: dict[str, str] | None = None, artifact_path: Path | None = None) -> dict[str, Any]:
        step = {
            "name": name, "kind": "argv", "argv": argv, "isolated": True,
            "env": env or {}, "cwd": str(root), "phase": phase, "side": side,
            "adapter_features": list(adapter_features),
            "binary_paths": [str(path) for path in binary_paths],
            "binary": str(binary),
        }
        if artifact_path is not None:
            step["artifact_path"] = str(artifact_path)
        if phase == "build" and name.endswith("-build"):
            step["produces_binaries"] = True
        if route is not None:
            step.update({"route": route.route, "output": route.output, "status": route.status})
            step["fixture"] = str(fixture)
        return step

    return [
        direct(
            "ppt-scoped-build",
            [*guarded, str(host_guard), "--", "cargo", "build", "--offline", "--locked", "--target", TARGET, "--release", *feature_flags, "-p", "litchi-ppt", "--example", "scoped_patch_probe", "--example", "scoped_patch_correctness"],
            route=PPT_ROUTES[0], phase="build",
            artifact_path=binary_paths[0],
        ),
        direct("ppt-scoped-correctness-suite", [*guarded, "sh", "-c", correctness], route=PPT_ROUTES[1], phase="build", env={"RUST_TEST_THREADS": "1"}, artifact_path=suite_receipt),
        direct("ppt-scoped-correctness", [sys.executable, str(host_guard), "--", str(binary), str(fixture), str(receipt)], route=PPT_ROUTES[2], artifact_path=receipt),
        direct("ppt-scoped-correctness-validate", [sys.executable, str(ppt_validator), str(receipt), str(fixture)], route=PPT_ROUTES[3], artifact_path=receipt),
    ]


def make_legacy_steps_v2(
    root: Path,
    state_dir: Path,
    target: Path,
    *,
    source_archive: Path | None = None,
    side: str = "candidate",
) -> list[dict[str, Any]]:
    """Build the current legacy consumer plan after fixed-source admission.

    The perf-baseline example is compiled with its explicit current adapter;
    route output names remain the eight fixed entries in the legacy manifest.
    The source archive is required even though the command itself only reads
    the source tree: it keeps the plan and its eventual receipt tied to the
    same archive proof as the PPT lane.
    """

    TRUSTED_POLICY.require_execution_ready()
    root = Path(root)
    state_dir = Path(state_dir)
    target = Path(target)
    _check_plan_paths(root, state_dir, target, source_archive)
    if source_archive is None:
        raise PolicyMismatch(f"legacy v2 {side} execution requires the verified source archive")
    policy_side = _policy_side(side)
    validate_source(root, policy_side, source_archive)
    guard = Path(__file__).resolve()
    host_guard = _protocol_tool_path("docs/unified-ops/results/0362g/ppt_host_process_guard.py")
    binary_paths = [target / TARGET / "release/examples/refusal_contract_bench"]
    binary = binary_paths[0]
    guarded = _guarded_lock(root, "perf", guard, policy_side)
    adapter_features = _adapter_features("legacy-current", side)
    build = [
        *guarded,
        str(host_guard),
        "--",
        "cargo", "build", "--offline", "--locked", "--target", TARGET, "--release",
        "--manifest-path", "tools/perf-baseline/Cargo.toml", "--no-default-features",
        "--features", *adapter_features, "--example", "refusal_contract_bench",
    ]

    def direct(name: str, argv: list[str], *, route: RoutePin | None = None, env: dict[str, str] | None = None, artifact_path: Path | None = None) -> dict[str, Any]:
        step = {
            "name": name, "kind": "argv", "argv": argv, "isolated": True,
            "env": env or {}, "cwd": str(root), "phase": "route" if route is not None else "build",
            "side": side, "adapter_features": list(adapter_features),
            "binary_paths": [str(path) for path in binary_paths], "binary": str(binary),
        }
        if artifact_path is not None:
            step["artifact_path"] = str(artifact_path)
        if route is None:
            step["produces_binaries"] = True
        if route is not None:
            step.update({"route": route.route, "output": route.output, "status": route.status, "phase": "route"})
        return step

    steps = [direct("legacy-current-build", build, artifact_path=binary)]
    for route in LEGACY_CURRENT_ROUTES:
        steps.append(
            direct(
                route.route,
                [sys.executable, str(host_guard), "--", str(binary)],
                route=route,
                env={"RUST_TEST_THREADS": "1"},
                artifact_path=state_dir / Path(route.output).name,
            )
        )
    return steps


def _default_v2_executor(step: Mapping[str, Any]) -> dict[str, Any]:
    """Run one reviewed argv step and return bounded receipt material.

    Tests and a host-specific runner may inject an executor instead.  The
    default deliberately does no interpretation of stdout; route output is
    admitted only when the command creates the reviewed output path or emits
    bytes which the caller can bind to that route.
    """

    argv = step.get("argv")
    if not isinstance(argv, Sequence) or isinstance(argv, (str, bytes)) or not argv:
        raise PolicyError("v2 execution step has no argv")
    if not all(isinstance(item, str) for item in argv):
        raise PolicyError("v2 execution argv contains a non-string argument")
    env = os.environ.copy()
    supplied_env = step.get("env", {})
    if not isinstance(supplied_env, Mapping):
        raise PolicyError("v2 execution step environment is malformed")
    env.update({str(key): str(value) for key, value in supplied_env.items()})
    cwd = step.get("cwd")
    completed = subprocess.run(
        list(argv),
        cwd=str(cwd) if isinstance(cwd, str) else None,
        env=env,
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    result: dict[str, Any] = {
        "returncode": int(completed.returncode),
        "stdout": completed.stdout,
        "stderr": completed.stderr,
        "cwd": str(cwd) if isinstance(cwd, str) else None,
        "argv": list(argv),
    }
    artifact_path = step.get("artifact_path")
    if isinstance(artifact_path, str) and Path(artifact_path).is_file() and not Path(artifact_path).is_symlink():
        result["output_path"] = artifact_path
    return result


def _v2_result_payload(
    result: Mapping[str, Any],
    route: RoutePin,
    output_root: Path,
    *,
    expected_fixture: Path | None = None,
) -> bytes:
    """Extract one reviewed, typed route stream without accepting raw bytes.

    The native RTF child emits a JSON-lines report containing all eight fixed
    scenarios.  The PPT correctness child emits its reviewed receipt, while
    the suite wrapper emits a tiny typed completion receipt.  Every accepted
    path is parsed and checked against the route; arbitrary stdout or an
    untagged ``output_bytes`` value is never promoted into evidence.
    """

    def objects(payload: bytes, label: str) -> list[tuple[dict[str, Any], bytes]]:
        if not payload:
            raise PolicyMismatch(f"{label} produced an empty output stream")
        try:
            value = json.loads(payload.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError):
            value = None
        if isinstance(value, Mapping):
            return [(dict(value), payload)]
        parsed: list[tuple[dict[str, Any], bytes]] = []
        for line in payload.splitlines(keepends=True):
            if not line.strip():
                continue
            try:
                item = json.loads(line.decode("utf-8"))
            except (UnicodeDecodeError, json.JSONDecodeError) as error:
                raise PolicyMismatch(f"{label} contains a non-JSON route record") from error
            if not isinstance(item, Mapping):
                raise PolicyMismatch(f"{label} contains a non-object route record")
            parsed.append((dict(item), line))
        if not parsed:
            raise PolicyMismatch(f"{label} produced no JSON route records")
        return parsed

    def matching(payload: bytes, label: str) -> bytes:
        rows = objects(payload, label)
        suffix = route.route.rsplit("/", 1)[-1]
        if route.program == "legacy-current":
            selected = []
            for value, raw in rows:
                if value.get("scenario") != suffix:
                    continue
                if value.get("workload") != "rtf-refusal-contract-v1" or value.get("outcome") not in {"ok", "refused"}:
                    raise PolicyMismatch(f"{route.route} has an untyped legacy workload record")
                median = value.get("median_us")
                if not isinstance(median, (int, float)) or isinstance(median, bool) or median <= 0:
                    raise PolicyMismatch(f"{route.route} has an invalid median")
                if value.get("fidelity") not in {"exact", "unchanged"}:
                    raise PolicyMismatch(f"{route.route} has an invalid fidelity marker")
                if route.status == "refused":
                    expected_kind = "protected" if suffix == "protected_commit" else "unsupported_content"
                    expected_reason = (
                        "document protection refuses changed publication"
                        if suffix == "protected_commit"
                        else "underline edits require non-empty text within one paragraph"
                    )
                    if value.get("outcome") != "refused" or value.get("refusal_kind") != expected_kind or value.get("operation") != "edit" or value.get("reason") != expected_reason:
                        raise PolicyMismatch(f"{route.route} refusal fields differ from the fixed contract")
                elif value.get("outcome") != "ok" or value.get("fidelity") != "exact":
                    raise PolicyMismatch(f"{route.route} success fields differ from the fixed contract")
                selected.append(raw)
            if len(selected) != 1:
                raise PolicyMismatch(f"{route.route} did not produce exactly one fixed scenario record")
            return selected[0]
        selected: list[bytes] = []

        def digest(value: Any, field: str) -> str:
            if not isinstance(value, str) or SHA256.fullmatch(value) is None:
                raise PolicyMismatch(f"{route.route} {field} is not a SHA-256 digest")
            return value

        def correctness(value: Mapping[str, Any]) -> None:
            required = {
                "mode", "fixture", "media_id", "requested_external_media_path",
                "external_media_path", "source_len", "target_len", "source_sha256",
                "target_sha256", "inverse_source_sha256", "inverse_matches_source",
            }
            missing = sorted(required - value.keys())
            if missing:
                raise PolicyMismatch(f"{route.route} correctness record misses: {', '.join(missing)}")
            if value.get("mode") != "scoped_patch_correctness":
                raise PolicyMismatch(f"{route.route} has the wrong correctness mode")
            fixture_value = value.get("fixture")
            if not isinstance(fixture_value, str):
                raise PolicyMismatch(f"{route.route} correctness fixture is not a path")
            fixture_path = Path(fixture_value).resolve()
            if expected_fixture is not None and fixture_path != Path(expected_fixture).resolve():
                raise PolicyMismatch(f"{route.route} correctness fixture differs from the fixed fixture")
            if value.get("media_id") != 1 or type(value.get("media_id")) is not int:
                raise PolicyMismatch(f"{route.route} correctness media ID is not 1")
            if value.get("requested_external_media_path") != "clip-0000.avi" or value.get("external_media_path") != "clip-0000.avi":
                raise PolicyMismatch(f"{route.route} correctness path differs from clip-0000.avi")
            source_len = value.get("source_len")
            target_len = value.get("target_len")
            if type(source_len) is not int or type(target_len) is not int or target_len <= 0:
                raise PolicyMismatch(f"{route.route} correctness extents are invalid")
            source_hash = digest(value.get("source_sha256"), "source_sha256")
            target_hash = digest(value.get("target_sha256"), "target_sha256")
            inverse_hash = digest(value.get("inverse_source_sha256"), "inverse_source_sha256")
            if source_hash == target_hash or inverse_hash != source_hash or value.get("inverse_matches_source") is not True:
                raise PolicyMismatch(f"{route.route} correctness replay/inverse relation is invalid")
            if expected_fixture is not None:
                fixture_bytes = Path(expected_fixture).read_bytes()
                if source_len != len(fixture_bytes) or source_hash != hashlib.sha256(fixture_bytes).hexdigest():
                    raise PolicyMismatch(f"{route.route} correctness source does not match the fixed fixture")
            replay_fields = {"replay_target_sha256", "replay_matches_target"} & value.keys()
            if replay_fields:
                if replay_fields != {"replay_target_sha256", "replay_matches_target"}:
                    raise PolicyMismatch(f"{route.route} replay fields must be supplied together")
                if digest(value.get("replay_target_sha256"), "replay_target_sha256") != target_hash or value.get("replay_matches_target") is not True:
                    raise PolicyMismatch(f"{route.route} forward replay relation is invalid")

        for value, raw in rows:
            if value.get("route") == route.route:
                if route.kind == "validate":
                    if value.get("status") != "passed" or value.get("external_media_path") != "clip-0000.avi" or value.get("inverse_matches_source") is not True:
                        raise PolicyMismatch(f"{route.route} validator summary is not the reviewed correctness result")
                selected.append(raw)
                continue
            if route.kind == "correctness" and value.get("status") == "ok":
                selected.append(raw)
            elif route.kind == "probe" and value.get("mode") == "scoped_patch_correctness":
                correctness(value)
                selected.append(raw)
            elif route.kind == "validate" and value.get("status") == "passed":
                if value.get("external_media_path") != "clip-0000.avi" or value.get("inverse_matches_source") is not True:
                    raise PolicyMismatch(f"{route.route} validator summary is not the reviewed correctness result")
                selected.append(raw)
        if len(selected) != 1:
            raise PolicyMismatch(f"{route.route} did not produce exactly one typed route record")
        return selected[0]

    output_bytes = result.get("output_bytes")
    if isinstance(output_bytes, bytes):
        return matching(output_bytes, f"{route.route} output_bytes")
    if isinstance(output_bytes, str):
        return matching(output_bytes.encode("utf-8"), f"{route.route} output_bytes")
    output_path = result.get("output_path")
    if isinstance(output_path, str):
        path = Path(output_path).absolute()
        if output_root.absolute() not in path.parents:
            raise PolicyMismatch(f"{route.route} executor output is outside the prepared output root")
        _reject_symlink_components(path, f"{route.route} executor output")
        if path.is_symlink() or not path.is_file():
            raise PolicyMismatch(f"{route.route} executor output is not a regular file")
        payload = path.read_bytes()
        # The validator route's child writes the correctness receipt and emits
        # a separate typed ``status=passed`` summary.  Prefer that summary for
        # its leg when present; the receipt itself remains the bound artifact.
        if route.kind == "validate" and result.get("stdout"):
            return matching(result["stdout"].encode() if isinstance(result["stdout"], str) else result["stdout"], f"{route.route} validator stdout")
        return matching(payload, f"{route.route} output file")
    stdout = result.get("stdout")
    if isinstance(stdout, str):
        stdout = stdout.encode("utf-8")
    if isinstance(stdout, bytes) and stdout:
        return matching(stdout, f"{route.route} stdout")
    expected = output_root / route.output
    if expected.is_file() and not expected.is_symlink():
        return matching(expected.read_bytes(), f"{route.route} policy output")
    raise PolicyMismatch(f"{route.route} executor produced no bound output stream")


def _v2_output_path(output_root: Path, relative: str) -> Path:
    path = Path(relative)
    if path.is_absolute() or ".." in path.parts or not relative:
        raise PolicyMismatch(f"v2 execution output path is unsafe: {relative}")
    expected = output_root / path
    _reject_symlink_components(expected, f"v2 execution output {relative}")
    try:
        expected.absolute().relative_to(output_root.absolute())
    except ValueError as error:
        raise PolicyMismatch(f"v2 execution output escapes its prepared root: {relative}") from error
    expected.parent.mkdir(parents=True, exist_ok=True)
    return expected


def _v2_output_identity(path: Path, relative: str) -> dict[str, Any]:
    digest, size = _hash_file(path, f"v2 execution output {relative}")
    return {"path": str(path.resolve()), "sha256": digest, "bytes": size, "relative_path": relative}


def _v2_file_identity(path: Path, label: str) -> dict[str, Any]:
    digest, size = _hash_file(path, label)
    return {"path": str(path.resolve()), "sha256": digest, "bytes": size}


def _binary_identities(step: Mapping[str, Any], result: Mapping[str, Any]) -> list[dict[str, Any]]:
    """Bind every binary produced by one side's reviewed build step."""

    paths = step.get("binary_paths")
    if not isinstance(paths, Sequence) or isinstance(paths, (str, bytes)) or not paths:
        raise PolicyMismatch(f"{step.get('name')} has no fixed binary output paths")
    expected_paths = [str(Path(value).absolute()) for value in paths if isinstance(value, str)]
    if len(expected_paths) != len(paths):
        raise PolicyMismatch(f"{step.get('name')} binary output paths are malformed")
    supplied = result.get("binary_identities")
    if supplied is not None:
        if not isinstance(supplied, Sequence) or isinstance(supplied, (str, bytes)) or len(supplied) != len(expected_paths):
            raise PolicyMismatch(f"{step.get('name')} binary identity count differs from its fixed plan")
        identities: list[dict[str, Any]] = []
        for expected, value in zip(expected_paths, supplied):
            if not isinstance(value, Mapping) or value.get("path") != str(Path(expected).resolve()):
                raise PolicyMismatch(f"{step.get('name')} returned a binary outside its fixed target")
            digest = value.get("sha256")
            size = value.get("bytes")
            if not isinstance(digest, str) or SHA256.fullmatch(digest) is None or type(size) is not int or size < 0:
                raise PolicyMismatch(f"{step.get('name')} returned a malformed binary identity")
            path = Path(expected)
            if not path.is_file() or path.is_symlink():
                raise PolicyMismatch(f"{step.get('name')} binary identity has no regular file: {path}")
            if _v2_file_identity(path, f"{step.get('name')} binary") != {
                "path": str(path.resolve()), "sha256": digest, "bytes": size
            }:
                raise PolicyMismatch(f"{step.get('name')} binary changed after executor reported it")
            identities.append({"path": str(path.resolve()), "sha256": digest, "bytes": size})
        return identities
    identities = []
    for expected in expected_paths:
        path = Path(expected)
        if not path.is_file() or path.is_symlink():
            raise PolicyMismatch(f"{step.get('name')} did not produce required binary: {path}")
        identities.append(_v2_file_identity(path, f"{step.get('name')} binary"))
    return identities


def _ensure_execution_paths(
    output_root: Path,
    qualification: Path,
    source_roots: Mapping[str, Path],
    archives: Mapping[str, Path],
    targets: Mapping[str, Path],
    states: Mapping[str, Path],
) -> None:
    """Reject every disallowed overlap before a child or receipt is written."""

    paths: list[tuple[str, Path]] = [("qualification output", qualification), ("prepared output root", output_root)]
    paths.extend((f"{side} source", path) for side, path in source_roots.items())
    paths.extend((f"{side} source archive", path) for side, path in archives.items())
    paths.extend((f"{side} target", path) for side, path in targets.items())
    paths.extend((f"{side} state", path) for side, path in states.items())
    resolved: list[tuple[str, Path]] = []
    for label, raw in paths:
        lexical = Path(raw).absolute()
        _reject_symlink_components(lexical, label)
        resolved.append((label, lexical.resolve(strict=False)))
    for index, (left_label, left) in enumerate(resolved):
        for right_label, right in resolved[index + 1 :]:
            # The target/state directories are intentionally inside the fresh
            # prepared output root; all other roots must remain disjoint from
            # one another and from the qualification receipt.
            if left_label == "prepared output root" and right_label in {"before target", "after target", "before state", "after state"}:
                continue
            if right_label == "prepared output root" and left_label in {"before target", "after target", "before state", "after state"}:
                continue
            if left == right or left in right.parents or right in left.parents:
                raise PolicyMismatch(f"execution paths overlap: {left_label} and {right_label}")


def execute_v2(
    preparation: Path,
    qualification: Path,
    *,
    target: Path | None = None,
    state_dir: Path | None = None,
    source_archive: Path | None = None,
    executor: Callable[[Mapping[str, Any]], Mapping[str, Any]] | None = None,
) -> dict[str, Any]:
    """Execute a prepared v2 lane and consume it through the v2 validator.

    The injected executor is intentionally narrow: it receives only fixed
    policy steps and returns a return code plus a bound output stream.  This
    makes source-only tests exercise the same receipt/route path as a native
    host runner without starting Cargo or a native child.  The default uses
    ``subprocess.run`` for a later quiet-host invocation.
    """

    from native_qualification import validate_current_native_v2 as validator

    prepared = validator.validate_preparation(Path(preparation))
    output_root = prepared.get("output_root")
    if not isinstance(output_root, Path):
        raise PolicyMismatch("v2 execution requires a prepared output root")
    output_root = output_root.resolve()
    base_target = Path(target) if target is not None else output_root / "target"
    base_state = Path(state_dir) if state_dir is not None else output_root / "state"
    lane = prepared["lane"]
    if lane not in TRUSTED_POLICY.routes:
        raise PolicyError(f"v2 execution lane is not admitted: {lane}")
    identities = prepared["receipt"].get("identities")
    if not isinstance(identities, Mapping):
        raise PolicyMismatch("v2 preparation has no paired source identities")
    source_roots: dict[str, Path] = {}
    archives: dict[str, Path] = {}
    for side in ("before", "after"):
        identity = identities.get(side)
        if not isinstance(identity, Mapping) or not isinstance(identity.get("source_root"), str):
            raise PolicyMismatch(f"v2 preparation identity is missing {side} source root")
        archive = identity.get("archive")
        archive_path = archive.get("path") if isinstance(archive, Mapping) else None
        if not isinstance(archive_path, str):
            raise PolicyMismatch(f"v2 preparation identity is missing {side} source archive")
        source_roots[side] = Path(str(identity["source_root"])).resolve()
        archives[side] = Path(archive_path).resolve()
    if source_archive is not None and Path(source_archive).resolve() != archives["after"]:
        raise PolicyMismatch("caller source archive differs from the preparation-bound candidate archive")
    targets = {"before": base_target / "before", "after": base_target / "after"}
    states = {"before": base_state / "before", "after": base_state / "after"}
    qualification_path = Path(qualification).absolute()
    _reject_protocol_source(qualification_path.parent, "v2 qualification output")
    if qualification_path.exists() or qualification_path.is_symlink():
        raise PolicyMismatch("v2 qualification receipt path is not fresh")
    _ensure_execution_paths(output_root, qualification_path, source_roots, archives, targets, states)
    for side in ("before", "after"):
        states[side].mkdir(parents=True, exist_ok=True)

    plan_builder = make_ppt_steps_v2 if lane == "ppt-auxiliary" else make_legacy_steps_v2
    plans: dict[str, list[dict[str, Any]]] = {}
    for side in ("before", "after"):
        plans[side] = plan_builder(
            source_roots[side], states[side], targets[side],
            source_archive=archives[side], side=side,
        )
    run_step = executor or _default_v2_executor
    step_receipts: list[dict[str, Any]] = []
    route_steps: dict[str, dict[str, Mapping[str, Any]]] = {"before": {}, "after": {}}
    preflight_results: dict[str, dict[str, Mapping[str, Any]]] = {"before": {}, "after": {}}
    side_binaries: dict[str, list[dict[str, Any]]] = {}
    for side in ("before", "after"):
        for step in plans[side]:
            route = step.get("route")
            if isinstance(route, str):
                route_steps[side][route] = step
            if step.get("phase") == "route":
                continue
            result = run_step(step)
            if not isinstance(result, Mapping):
                raise PolicyError(f"executor returned a non-object for {step.get('name')}")
            returncode = int(result.get("returncode", 0))
            if returncode != 0:
                raise PolicyMismatch(f"v2 {side} build step failed: {step.get('name')}")
            binary_identities = _binary_identities(step, result) if step.get("produces_binaries") else []
            if step.get("produces_binaries"):
                side_binaries[side] = binary_identities
            if isinstance(route, str):
                preflight_results[side][route] = result
            step_receipts.append({
                "name": step.get("name"), "side": side, "phase": step.get("phase"),
                "returncode": returncode, "cwd": step.get("cwd"),
                "argv": list(step.get("argv", [])), "adapter_features": step.get("adapter_features"),
                "binary_identities": binary_identities,
            })
        if side not in side_binaries:
            raise PolicyMismatch(f"v2 {side} build did not produce binary provenance")
    expected_routes = TRUSTED_POLICY.routes_for(lane)
    expected_route_names = {route.route for route in expected_routes}
    for side in ("before", "after"):
        if set(route_steps[side]) != expected_route_names:
            raise PolicyMismatch(f"v2 {side} plan route mapping differs from the fixed {lane} routes")
    source_bindings = {
        side: hashlib.sha256(_canonical(identities[side])).hexdigest()
        for side in ("before", "after")
    }
    blocks = ABBA_BLOCKS[lane]
    leg_records: dict[str, list[dict[str, Any]]] = {route.route: [] for route in expected_routes}
    final_payloads: dict[str, bytes] = {}
    for route in expected_routes:
        for block in range(1, blocks + 1):
            for index, side in enumerate(ABBA_ORDER):
                planned = route_steps[side][route.route]
                invocation = dict(planned)
                binary_index = 1 if lane == "ppt-auxiliary" and route.program == "scoped_patch_correctness" else 0
                binary_identity = side_binaries[side][binary_index]
                invocation.update({
                    "abba_block": block,
                    "abba_index": index,
                    "abba_role": side,
                    "source_root": str(source_roots[side]),
                    "source_binding": source_bindings[side],
                    "binary": binary_identity["path"],
                    "binary_identity": binary_identity,
                })
                if planned.get("phase") == "route":
                    result = run_step(invocation)
                else:
                    result = preflight_results[side][route.route]
                if not isinstance(result, Mapping):
                    raise PolicyError(f"executor returned a non-object for {route.route}/{side}")
                returncode = int(result.get("returncode", 0))
                if returncode != 0:
                    raise PolicyMismatch(f"v2 route failed: {route.route}/{side}")
                if route.kind == "build":
                    payload = _canonical({"route": route.route, "status": "ok", "binary": binary_identity})
                    actual_for_final = Path(binary_identity["path"]).read_bytes() if Path(binary_identity["path"]).is_file() else payload
                else:
                    payload = _v2_result_payload(
                        result,
                        route,
                        output_root,
                        expected_fixture=Path(planned["fixture"]) if isinstance(planned.get("fixture"), str) else None,
                    )
                    artifact_path = result.get("output_path") or planned.get("artifact_path")
                    actual_for_final = Path(str(artifact_path)).read_bytes() if isinstance(artifact_path, str) and Path(artifact_path).is_file() else payload
                relative_leg = f".abba/{re.sub(r'[^A-Za-z0-9_.-]+', '_', route.route)}/b{block}-{index}-{side}.stream"
                leg_path = _v2_output_path(output_root, relative_leg)
                leg_path.write_bytes(payload)
                leg_records[route.route].append({
                    "route": route.route, "role": side, "index": index, "block": block,
                    "returncode": returncode, "source_binding": source_bindings[side],
                    "cwd": invocation.get("cwd"), "argv": invocation.get("argv"),
                    "binary": binary_identity,
                    "output_artifact": _v2_output_identity(leg_path, relative_leg),
                })
                if side == "after":
                    final_payloads[route.route] = actual_for_final
                if planned.get("phase") == "route":
                    step_receipts.append({
                        "name": route.route, "side": side, "phase": "route", "block": block,
                        "abba_index": index, "role": side, "returncode": returncode,
                        "cwd": invocation.get("cwd"), "argv": invocation.get("argv"),
                        "binary": binary_identity,
                    })
    written_outputs: set[str] = set()
    for route in expected_routes:
        final_path = _v2_output_path(output_root, route.output)
        if route.output in written_outputs:
            continue
        if final_path.exists() or final_path.is_symlink():
            raise PolicyMismatch(f"v2 route output is not fresh: {route.output}")
        payload = final_payloads.get(route.route)
        if payload is None:
            raise PolicyMismatch(f"v2 route has no actual output artifact: {route.route}")
        final_path.write_bytes(payload)
        written_outputs.add(route.output)
    rows: list[dict[str, Any]] = []
    for route in expected_routes:
        final_path = _v2_output_path(output_root, route.output)
        evidence = {
            "source_bindings": source_bindings,
            "binary_provenance": side_binaries,
            "output_artifact": _v2_output_identity(final_path, route.output),
            "abba": leg_records[route.route],
        }
        rows.append({**route.json(), "evidence": evidence})
    runner = _v2_file_identity(Path(__file__).resolve(), "current-native v2 runner")
    execution = {
        "schema": EXECUTION_SCHEMA,
        "status": "completed",
        "runner_kind": RUNNER_KIND,
        "policy_sha256": TRUSTED_POLICY.digest(),
        "preparation_receipt_sha256": prepared["receipt_sha256"],
        "lane": lane,
        "output_root": str(output_root),
        "routes": [route.json() for route in expected_routes],
        "source_artifacts": prepared["receipt"]["identities"],
        "tooling": prepared["receipt"]["tooling"],
        "runner": runner,
        "abba_order": list(ABBA_ORDER),
        "abba_blocks": blocks,
        "targets": {side: str(path.resolve()) for side, path in targets.items()},
        "states": {side: str(path.resolve()) for side, path in states.items()},
        "adapter_features": {side: list(_adapter_features(lane, side)) for side in ("before", "after")},
        "binary_provenance": side_binaries,
        "step_receipts": step_receipts,
        "route_receipts": rows,
    }
    execution_path = output_root / EXECUTION_RECEIPT_RELATIVE
    if execution_path.exists() or execution_path.is_symlink():
        raise PolicyMismatch("v2 execution receipt path is not fresh")
    execution_path.write_bytes(_canonical(execution))
    qualification_path = Path(qualification).absolute()
    _reject_symlink_components(qualification_path, "v2 qualification receipt")
    _reject_protocol_source(qualification_path.parent, "v2 qualification output")
    if qualification_path.exists() or qualification_path.is_symlink():
        raise PolicyMismatch("v2 qualification receipt path is not fresh")
    if qualification_path == output_root or output_root in qualification_path.parents:
        raise PolicyMismatch("v2 qualification receipt must be outside the output root")
    qualification_path.parent.mkdir(parents=True, exist_ok=True)
    execution_identity = _v2_output_identity(execution_path, EXECUTION_RECEIPT_RELATIVE)
    qualification_value = {
        "schema": QUALIFICATION_SCHEMA,
        "policy_schema": POLICY_SCHEMA,
        "policy_sha256": TRUSTED_POLICY.digest(),
        "preparation_receipt_sha256": prepared["receipt_sha256"],
        "lane": lane,
        "abba_order": list(ABBA_ORDER),
        "abba_blocks": blocks,
        "output_root": str(output_root),
        "source_artifacts": prepared["receipt"]["identities"],
        "tooling": prepared["receipt"]["tooling"],
        "binary_provenance": side_binaries,
        "routes": [route.json() for route in expected_routes],
        "results": rows,
        "execution_receipt": execution_identity,
    }
    qualification_path.write_bytes(_canonical(qualification_value))
    return validator.validate_qualification(qualification_path, Path(preparation))


def _cli_guard_lock(source_root: Path, command: Sequence[str], role: str = "workspace", side: str = "candidate") -> int:
    """CLI implementation for a role guard; no expected hash input."""

    validate_lock(source_root, side, role)
    if not command:
        print(_policy_pin(side).lock_for_role(role).sha256)
        return 0
    completed = subprocess.run(list(command), check=False)
    return int(completed.returncode)


def main(argv: Sequence[str] | None = None) -> int:
    args = list(sys.argv[1:] if argv is None else argv)
    if not args or args[0] not in {"guard-lock", "guard-role-lock"}:
        print("usage: current_native_policy_v2.py guard-lock SOURCE_ROOT [-- COMMAND ...]", file=sys.stderr)
        print("       current_native_policy_v2.py guard-role-lock ROLE SOURCE_ROOT [-- COMMAND ...]", file=sys.stderr)
        return 2
    try:
        separator = args.index("--")
    except ValueError:
        separator = len(args)
    if args[0] == "guard-role-lock":
        if len(args) < 3 or args[1] not in ROLE_NAMES:
            print("usage: current_native_policy_v2.py guard-role-lock ROLE SOURCE_ROOT [--side SIDE] [-- COMMAND ...]", file=sys.stderr)
            return 2
        role = args[1]
        path_index = 2
        side = "candidate"
        pre_separator = args[3:separator]
        if pre_separator:
            if len(pre_separator) != 2 or pre_separator[0] != "--side" or pre_separator[1] not in {"baseline", "candidate"}:
                print("usage: current_native_policy_v2.py guard-role-lock ROLE SOURCE_ROOT [--side SIDE] [-- COMMAND ...]", file=sys.stderr)
                return 2
            side = pre_separator[1]
    else:
        if len(args) < 2 or separator != 2:
            print("usage: current_native_policy_v2.py guard-lock SOURCE_ROOT [-- COMMAND ...]", file=sys.stderr)
            return 2
        role = "workspace"
        path_index = 1
        side = "candidate"
    try:
        return _cli_guard_lock(Path(args[path_index]), args[separator + 1 :], role, side)
    except (OSError, PolicyError) as error:
        print(f"current-source policy guard refused: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
