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
from typing import Any, Iterable, Mapping, Sequence


POLICY_SCHEMA = "litchi.current-native-policy.v2"
PREPARATION_SCHEMA = "litchi.current-native-preparation.v2"
QUALIFICATION_SCHEMA = "litchi.current-native-qualification.v2"
PROTOCOL_SCHEMA = "litchi.native-evidence-protocol.v2"
POLICY_VERSION = 2

BASELINE_COMMIT = "f2bbbc8d07ab5da26c180f1ff74813d39f632d54"
BASELINE_GIT_TREE = "2e0d0cdf7f6a345961ceca361d53202200d00970"
BASELINE_TREE_SHA256 = "ce92652c439d810849b74b80e650821bc21de9029ea861cd6bbc13e38d809589"
BASELINE_LOCK_SHA256 = "1b9b10ae7a9fd65db3b8b29a5bedad1aa560473c10b462f8a7d654c169335b2c"
BASELINE_LOCK_BYTES = 139_441
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

    def json(self) -> dict[str, str]:
        return {
            "route": self.route,
            "kind": self.kind,
            "program": self.program,
            "output": self.output,
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
            "target": self.target,
            "profile": self.profile,
            "features": list(self.features),
        }


PPT_ROUTES = (
    RoutePin("ppt-scoped-build", "build", "cargo", "target/release/examples/scoped_patch_probe"),
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
    RoutePin("legacy-current/unsupported_content", "refusal", "legacy-current", "state/unsupported_content.json"),
    RoutePin("legacy-current/protected_commit", "refusal", "legacy-current", "state/protected_commit.json"),
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


def observe_source(source: Path, side: str) -> dict[str, Any]:
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
    listing = subprocess.run(
        ["git", "ls-tree", "-r", "-z", "HEAD"], cwd=root, check=True,
        stdout=subprocess.PIPE, stderr=subprocess.PIPE,
    ).stdout
    status = _git(root, "status", "--porcelain=v1")
    pin = TRUSTED_POLICY.baseline if side == "baseline" else TRUSTED_POLICY.candidate
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
    return {
        "side": side,
        "source_root": str(root),
        "commit": head,
        "git_tree": tree,
        "tree_sha256": hashlib.sha256(listing).hexdigest(),
        "clean": not bool(status),
        "status": status,
        "artifacts": artifacts,
        "fixtures": fixtures,
    }


def validate_source(source: Path, side: str) -> dict[str, Any]:
    """Validate one endpoint against its fixed policy pins."""

    policy_pin = TRUSTED_POLICY.baseline if side == "baseline" else TRUSTED_POLICY.candidate
    if not policy_pin.complete():
        raise PolicyIncomplete(f"{side} source pins are not complete in current-source policy v2")
    observed = observe_source(source, side)
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
        for field in ("route", "kind", "program", "output"):
            if actual.get(field) != getattr(wanted, field):
                raise PolicyMismatch(f"{lane} route {index} {field} differs from the trusted policy")


def make_ppt_steps_v2(root: Path, state_dir: Path, target: Path) -> list[dict[str, Any]]:
    """Build the opt-in PPT auxiliary plan after the fixed policy is ready.

    The old ``make_ppt_steps`` is deliberately not imported or changed.  The
    lock guard below has no expected-digest argument; it reads the reviewed
    policy itself and therefore cannot be redirected by a manifest caller.
    """

    TRUSTED_POLICY.require_execution_ready()
    root = Path(root)
    state_dir = Path(state_dir)
    target = Path(target)
    lock = root / "Cargo.lock"
    guard = Path(__file__).resolve()
    host_guard = root / "docs/unified-ops/results/0362g/ppt_host_process_guard.py"
    fixture = root / "test-data/unified-ops/0294/video.ppt"
    binary = target / "release/examples/scoped_patch_correctness"
    receipt = state_dir / "ppt-scoped-correctness.json"
    guarded = [sys.executable, str(guard), "guard-lock", str(lock), "--"]
    correctness = (
        "set -eu; "
        "cargo test --offline --locked -p litchi-ppt --all-features --lib scoped::tests -- --test-threads=1; "
        "cargo test --offline --locked -p litchi-ppt --all-features --lib -- --test-threads=1; "
        "cargo test --offline --locked -p litchi-cfb --all-features --lib -- --test-threads=1; "
        "cargo test --offline --locked -p litchi-ppt --all-features --tests -- --test-threads=1; "
        "cargo test --offline --locked -p litchi-cfb --all-features --tests -- --test-threads=1; "
        "cargo clippy --offline --locked -p litchi-ppt --all-features --all-targets -- -D warnings; "
        "cargo clippy --offline --locked -p litchi-cfb --all-features --all-targets -- -D warnings"
    )

    def direct(name: str, argv: list[str], *, env: dict[str, str] | None = None) -> dict[str, Any]:
        return {"name": name, "kind": "argv", "argv": argv, "isolated": True, "env": env or {}}

    return [
        direct(
            "ppt-current-scoped-build",
            [*guarded, str(host_guard), "--", "cargo", "build", "--offline", "--locked", "--release", "-p", "litchi-ppt", "--example", "scoped_patch_probe", "--example", "scoped_patch_correctness"],
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
