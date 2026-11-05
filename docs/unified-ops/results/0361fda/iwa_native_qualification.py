#!/usr/bin/env python3
"""Build and qualify matched native IWA constructor and owner probes.

The build phase snapshots both commits and records every source, lockfile,
toolchain, dependency and binary identity outside the checkout.  Validation
must complete for both native and counting binaries before the native phase is
allowed to start.  The native phase uses calibrated, five-block ABBA runs and
fails closed on any build/profiler/benchmark overlap.  No tolerant timing is
emitted.

Examples (all generated files are external to the checkout)::

  python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py build \
      constructor /outside/0361fda --base BASE --candidate CANDIDATE \
      --lock Cargo.lock
  python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py validate \
      constructor /outside/0361fda
  python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py run \
      constructor /outside/0361fda --cpu 29
  python3 docs/unified-ops/results/0361fda/iwa_native_qualification.py allocation \
      constructor /outside/0361fda --cpu 29 --attempt allocation-a1

The owner build, allocation, and run take ``--fixtures``.  The fixture directory is an
external copy of the tracked test-data fixtures; this harness never reaches
into an untracked third-party checkout.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
from pathlib import Path
import re
import shlex
import shutil
import signal
import subprocess
import sys
import tarfile
import tempfile
import time
from typing import Any, Iterable, Mapping

try:
    import resource
except ImportError:  # pragma: no cover - the qualification host is POSIX.
    resource = None  # type: ignore[assignment]


ROOT = Path(__file__).resolve().parents[4]
HERE = Path(__file__).resolve().parent
PROBES = {
    "constructor": HERE / "constructor-probe.rs",
    "owner": HERE / "owner_probe.rs",
}
ROUTES = {
    "constructor": (
        "field_path",
        "field_info",
        "message_info",
        "archive_info",
        "archive_object",
        "archive_object_reject",
        "parse",
        "validate",
        "encode",
        "retained_parse",
        "retained_validate",
        "retained_encoded_len",
        "retained_encode",
        "retained_validate_tight",
        "retained_encode_tight",
    ),
    "owner": (
        "numbers_headers",
        "keynote_media",
        "pages_footnote_insert",
        "pages_section_name",
    ),
}
START_COUNTS = {"constructor": 1_000, "owner": 10}
ALLOCATION_ITERATIONS = 100
ALLOCATION_ORDER = "ABBA"
ALLOCATION_METRICS = ("peak", "retained", "allocated", "allocations")
TARGET_NS = 500_000_000
MINIMUM_NS = 150_000_000
MAX_ITERATIONS = 1_000_000_000
FIXED_BLOCKS = 5
FIXED_WARMUPS = 2
FIXED_GROWTH = 1.2
FIXED_TIMEOUT_S = 120.0
COMMAND_TIMEOUT_S = 900.0
# This first metadata pass is intentionally unlocked only so Cargo can add the
# external probe package and source-local edges to the copied seed lock.  It
# must not be replaced with generate-lockfile: lock_delta_manifest rejects any
# registry refresh before a release build is allowed to start.
LOCK_PREPARATION_COMMAND = ("cargo", "metadata", "--offline", "--format-version=1")
# The generator has its own 900-second Cargo child bound.  Keep the outer
# wrapper bound slightly longer so its normal timeout handler can reap that
# child before the runner has to tear down both process groups.
FIXTURE_GENERATOR_TIMEOUT_S = 900.0
FIXTURE_GENERATOR_OUTER_GRACE_S = 30.0
FIXTURE_GENERATOR_OUTER_TIMEOUT_S = (
    FIXTURE_GENERATOR_TIMEOUT_S + FIXTURE_GENERATOR_OUTER_GRACE_S
)
SCHEMA = "litchi.iwa.native-qualification.v4"
COMPARISON = "integrated-route"
SOURCE_DELTA_SCHEMA = "litchi.iwa.source-delta.v1"
LOCK_DELTA_SCHEMA = "litchi.iwa.lock-delta.v4"
REGISTRY_SOURCE = "registry+https://github.com/rust-lang/crates.io-index"
REVIEWED_REGISTRY_ADDITIONS = {
    "constructor": [
        {
            "name": "smallvec",
            "version": "1.16.3",
            "source": REGISTRY_SOURCE,
            "checksum": "5b3dc8af474f516a851ff4bd12db780f948b9250ad37211e4eec0bccea54e01b",
        }
    ],
    # The owner source adds the edge, but its baseline closure already carries
    # this exact identity through another selected local package.
    "owner": [],
}
OWNER_FIXTURES = (
    "basic.numbers",
    "media-replacement-native.key",
    "tone.wav",
    "footnote-fixture.pages",
    "basic.pages",
)
FIXTURE_PROVENANCE = HERE / "fixture-provenance.json"
FIXTURE_GENERATOR = HERE / "generate_footnote_fixture.py"


class QualificationError(RuntimeError):
    """A fail-closed qualification precondition or receipt failure."""


def _reject_constant(value: str) -> None:
    raise QualificationError(f"non-finite JSON value {value!r}")


def _reject_duplicate(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise QualificationError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def load_json(path: Path) -> Any:
    try:
        text = path.read_text(encoding="utf-8")
        return json.loads(
            text, object_pairs_hook=_reject_duplicate, parse_constant=_reject_constant
        )
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise QualificationError(f"cannot read strict JSON {path}: {error}") from error


def save_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps(value, indent=2, sort_keys=True, ensure_ascii=True) + "\n",
        encoding="utf-8",
    )


def artifact(path: Path) -> dict[str, Any]:
    """Return a content identity without loading the whole artifact."""

    try:
        resolved = path.resolve(strict=True)
        stat = resolved.stat()
    except OSError as error:
        raise QualificationError(f"cannot inspect artifact {path}: {error}") from error
    if not resolved.is_file():
        raise QualificationError(f"artifact is not a regular file: {resolved}")
    digest = hashlib.sha256()
    with resolved.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return {"path": str(resolved), "bytes": stat.st_size, "sha256": digest.hexdigest()}


def verify_artifact(row: Mapping[str, Any]) -> None:
    if not isinstance(row, Mapping) or not isinstance(row.get("path"), str):
        raise QualificationError("malformed artifact identity")
    actual = artifact(Path(row["path"]))
    expected = {key: row.get(key) for key in ("path", "bytes", "sha256")}
    if actual != expected:
        raise QualificationError(f"artifact changed: {row['path']}")


def require_family(family: str) -> str:
    if family not in ROUTES:
        raise QualificationError(f"unknown family {family!r}")
    return family


def outside_checkout(path: Path) -> Path:
    resolved = path.resolve()
    if resolved.is_relative_to(ROOT):
        raise QualificationError(
            f"generated source, binary, timing, or evidence path is inside checkout: {resolved}"
        )
    return resolved


def _tracked_identity(spec: Mapping[str, Any], label: str) -> dict[str, Any]:
    relative = spec.get("path")
    if not isinstance(relative, str) or Path(relative).is_absolute():
        raise QualificationError(f"tracked fixture provenance has no relative path: {label}")
    path = (ROOT / relative).resolve()
    if not path.is_relative_to(ROOT) or "3rdparty" in path.parts:
        raise QualificationError(f"tracked fixture provenance escapes the repository: {label}")
    actual = artifact(path)
    expected = {
        "path": str(path),
        "bytes": spec.get("bytes"),
        "sha256": spec.get("sha256"),
    }
    if actual != expected:
        raise QualificationError(f"tracked fixture provenance changed: {label}")
    return actual


def _fixture_provenance() -> dict[str, Any]:
    provenance = load_json(FIXTURE_PROVENANCE)
    if provenance.get("schema") != "litchi.iwa.owner-fixtures.provenance.v1":
        raise QualificationError("tracked owner fixture provenance has the wrong schema")
    rows = provenance.get("files")
    if not isinstance(rows, dict) or set(rows) != set(OWNER_FIXTURES):
        raise QualificationError("tracked owner fixture provenance has the wrong workload set")
    for name in OWNER_FIXTURES:
        spec = rows[name]
        if not isinstance(spec, Mapping):
            raise QualificationError(f"tracked owner fixture provenance is malformed: {name}")
        if name == "footnote-fixture.pages":
            generator = spec.get("generator")
            sources = spec.get("sources")
            if not isinstance(generator, Mapping) or not isinstance(sources, list) or not sources:
                raise QualificationError("footnote fixture provenance lacks generator sources")
            generator_path = (ROOT / str(generator.get("path", ""))).resolve()
            if generator_path != FIXTURE_GENERATOR.resolve():
                raise QualificationError("footnote fixture uses an unapproved generator")
            generator_identity = artifact(generator_path)
            expected_generator = {
                "path": str(generator_path),
                "bytes": generator.get("bytes"),
                "sha256": generator.get("sha256"),
            }
            if generator_identity != expected_generator:
                raise QualificationError("footnote fixture generator changed")
            for source in sources:
                if not isinstance(source, Mapping):
                    raise QualificationError("footnote fixture generator source is malformed")
                _tracked_identity(source, name)
        else:
            _tracked_identity(spec, name)
    return provenance


def fixture_identity(path: Path, *, regenerate: bool = False) -> dict[str, Any]:
    """Verify an external owner fixture set and its immutable manifest."""

    directory = outside_checkout(path)
    if not directory.is_dir():
        raise QualificationError(f"owner fixtures are not a directory: {directory}")
    manifest_path = directory / "manifest.json"
    manifest = load_json(manifest_path)
    if manifest.get("schema") != "litchi.iwa.owner-fixtures.v1":
        raise QualificationError("owner fixture manifest has the wrong schema")
    provenance_identity = artifact(FIXTURE_PROVENANCE)
    if manifest.get("provenance") != provenance_identity:
        raise QualificationError("owner fixture manifest is not bound to tracked provenance")
    provenance = _fixture_provenance()
    declared = manifest.get("files")
    if not isinstance(declared, dict) or set(declared) != set(OWNER_FIXTURES):
        raise QualificationError("owner fixture manifest does not declare exactly the five workload files")
    identities = {}
    for name in OWNER_FIXTURES:
        row = declared[name]
        if not isinstance(row, Mapping):
            raise QualificationError(f"owner fixture identity is malformed: {name}")
        expected_path = directory / name
        if Path(row.get("path", "")).resolve() != expected_path.resolve():
            raise QualificationError(f"owner fixture path is not rooted in the fixture directory: {name}")
        verify_artifact(row)
        spec = provenance["files"][name]
        if name == "footnote-fixture.pages":
            generator = spec["generator"]
            sources = spec["sources"]
            expected_generated = {
                "generator": artifact(FIXTURE_GENERATOR),
                "sources": [
                    _tracked_identity(source, name)
                    for source in sources
                ],
                "command": spec.get("command"),
                "method": spec.get("method"),
            }
            if row.get("generated") != expected_generated:
                raise QualificationError("footnote fixture output is not tied to its generator")
            if generator.get("path") != str(FIXTURE_GENERATOR.relative_to(ROOT)):
                raise QualificationError("footnote fixture generator provenance is malformed")
            if regenerate:
                with tempfile.TemporaryDirectory(prefix="litchi-iwa-footnote-") as temporary:
                    generated = Path(temporary) / name
                    _run_fixture_generator(generated)
                    generated_identity = artifact(generated)
                    # The replay is deliberately written to a fresh temporary
                    # path.  Compare content identity only; the manifest keeps
                    # the original fixture path and generated provenance.
                    if {
                        key: generated_identity[key] for key in ("bytes", "sha256")
                    } != {key: row.get(key) for key in ("bytes", "sha256")}:
                        raise QualificationError("generated footnote fixture bytes changed")
        else:
            expected_source = _tracked_identity(spec, name)
            if row.get("source") != expected_source:
                raise QualificationError(f"fixture {name} is not tied to its tracked source")
        identities[name] = row
    return {
        "manifest": artifact(manifest_path),
        "provenance": provenance_identity,
        "files": identities,
    }


def prepare_fixtures(path: Path) -> None:
    """Materialize and manifest the external owner workload set."""

    directory = new_external_directory(path)
    provenance = _fixture_provenance()
    tracked = provenance["files"]
    sources = {
        "basic.numbers": ROOT / tracked["basic.numbers"]["path"],
        "media-replacement-native.key": ROOT / tracked["media-replacement-native.key"]["path"],
        "tone.wav": ROOT / tracked["tone.wav"]["path"],
        "basic.pages": ROOT / tracked["basic.pages"]["path"],
    }
    for name, source in sources.items():
        shutil.copyfile(source, directory / name)
    _run_fixture_generator(directory / "footnote-fixture.pages")
    files: dict[str, Any] = {}
    for name in OWNER_FIXTURES:
        row: dict[str, Any] = artifact(directory / name)
        if name == "footnote-fixture.pages":
            row["generated"] = {
                "generator": artifact(FIXTURE_GENERATOR),
                "sources": [
                    _tracked_identity(source, name)
                    for source in tracked[name]["sources"]
                ],
                "command": tracked[name]["command"],
                "method": tracked[name]["method"],
            }
        else:
            row["source"] = artifact(sources[name])
        files[name] = row
    save_json(
        directory / "manifest.json",
        {
            "schema": "litchi.iwa.owner-fixtures.v1",
            "provenance": artifact(FIXTURE_PROVENANCE),
            "files": files,
        },
    )
    fixture_identity(directory)


def new_external_directory(path: Path) -> Path:
    path = outside_checkout(path)
    try:
        path.mkdir(parents=True, exist_ok=False)
    except OSError as error:
        raise QualificationError(f"cannot create fresh external directory {path}: {error}") from error
    return path


def _generator_process_groups(child: subprocess.Popen[bytes]) -> set[int]:
    """Find the wrapper and its descendants' process groups.

    ``generate_footnote_fixture.py`` puts Cargo in a new session so that its
    own timeout can cleanly stop Cargo's compiler tree.  An outer timeout must
    therefore collect both the wrapper's group and that separate Cargo group.
    The process snapshot is taken before signalling, while parent links are
    still available; the group ids remain useful if the wrapper exits while a
    compiler child is being reparented.
    """

    rows = process_rows()
    descendants = _descendants(rows, child.pid)
    groups: set[int] = set()
    for pid in descendants:
        try:
            groups.add(os.getpgid(pid))
        except ProcessLookupError:
            continue
    # The child is a session leader in normal operation.  Keeping its pid as
    # a fallback also covers the short race where /proc has not listed it yet.
    groups.add(child.pid)
    return groups


def _signal_process_groups(groups: Iterable[int], sig: signal.Signals) -> None:
    for pgid in sorted(set(groups)):
        try:
            os.killpg(pgid, sig)
        except ProcessLookupError:
            continue


def _terminate_generator(child: subprocess.Popen[bytes]) -> int:
    """Stop a timed generator and every Cargo descendant it started."""

    groups = _generator_process_groups(child)
    _signal_process_groups(groups, signal.SIGTERM)
    try:
        status = child.wait(timeout=2.0)
    except subprocess.TimeoutExpired:
        # Refresh before the hard kill: a compiler may have appeared after the
        # first snapshot, while the original groups still cover reparented
        # children from the timed-out Cargo invocation.
        groups.update(_generator_process_groups(child))
        _signal_process_groups(groups, signal.SIGKILL)
        try:
            status = child.wait(timeout=2.0)
        except subprocess.TimeoutExpired:
            child.kill()
            status = child.wait(timeout=2.0)
    # The wrapper can exit after SIGTERM while its separate-session Cargo
    # child is still unwinding.  A final hard signal prevents that child from
    # outliving the failed qualification process.
    _signal_process_groups(groups, signal.SIGKILL)
    return status


def _run_fixture_generator(output: Path) -> None:
    """Run the fixture wrapper with coordinated outer timeout cleanup."""

    command = [sys.executable, str(FIXTURE_GENERATOR), "--output", str(output)]
    try:
        child = subprocess.Popen(command, cwd=ROOT, start_new_session=True)
    except OSError as error:
        raise QualificationError("footnote fixture generator failed to start") from error
    try:
        status = child.wait(timeout=FIXTURE_GENERATOR_OUTER_TIMEOUT_S)
    except subprocess.TimeoutExpired as error:
        _terminate_generator(child)
        raise QualificationError(
            f"footnote fixture generator exceeded {FIXTURE_GENERATOR_OUTER_TIMEOUT_S:g}s"
        ) from error
    if status:
        raise QualificationError(f"footnote fixture generator failed with status {status}")
    if not output.is_file():
        raise QualificationError(f"footnote fixture generator did not produce {output}")


def run_command(
    directory: Path,
    label: str,
    argv: list[str],
    cwd: Path,
    env: Mapping[str, str] | None = None,
    timeout_s: float = COMMAND_TIMEOUT_S,
) -> dict[str, Any]:
    """Run one non-shell command and retain a receipt before checking status."""

    if not math.isfinite(timeout_s) or timeout_s <= 0:
        raise QualificationError("command timeout must be finite and positive")
    stdout = directory / f"{label}.stdout"
    stderr = directory / f"{label}.stderr"
    started = time.time_ns()
    with stdout.open("wb") as out, stderr.open("wb") as err:
        child = subprocess.Popen(
            argv,
            cwd=cwd,
            env=env,
            stdout=out,
            stderr=err,
            start_new_session=True,
            preexec_fn=_child_limits(timeout_s, file_bytes=None),
        )
        timed_out = False
        try:
            status = child.wait(timeout=timeout_s)
        except subprocess.TimeoutExpired:
            timed_out = True
            status = _terminate_child(child)
    receipt = {
        "command": argv,
        "cwd": str(cwd.resolve()),
        "started_ns": started,
        "finished_ns": time.time_ns(),
        "exit_code": status,
        "timed_out": timed_out,
        "timeout_s": timeout_s,
        "stdout": artifact(stdout),
        "stderr": artifact(stderr),
    }
    receipt_path = directory / f"{label}.receipt.json"
    save_json(receipt_path, receipt)
    receipt["receipt"] = artifact(receipt_path)
    if status:
        raise QualificationError(f"{label} failed; see {stderr}")
    return receipt


def verify_command_receipt(row: Mapping[str, Any]) -> None:
    """Verify a retained command's streams, receipt, and successful status."""

    if row.get("exit_code") != 0:
        raise QualificationError("retained command receipt is not successful")
    for key in ("stdout", "stderr", "receipt"):
        verify_artifact(row[key])
    saved = load_json(Path(row["receipt"]["path"]))
    if saved.get("exit_code") != row["exit_code"]:
        raise QualificationError("retained command receipt status changed")
    for key in (
        "command",
        "cwd",
        "started_ns",
        "finished_ns",
        "stdout",
        "stderr",
        "timed_out",
        "timeout_s",
    ):
        if saved.get(key) != row.get(key):
            raise QualificationError(f"retained command receipt field changed: {key}")


def verify_lock_preparation_command(commands: Any) -> None:
    """Require seed-preserving metadata preparation before toolchain checks."""

    if not isinstance(commands, list) or not commands:
        raise QualificationError("missing endpoint lock-preparation receipt")
    first = commands[0]
    if not isinstance(first, Mapping) or first.get("command") != list(LOCK_PREPARATION_COMMAND):
        raise QualificationError("endpoint lock was not prepared with seed-preserving metadata")


def git_revision(ref: str) -> str:
    try:
        value = subprocess.check_output(
            ["git", "rev-parse", "--verify", f"{ref}^{{commit}}"],
            cwd=ROOT,
            text=True,
            stderr=subprocess.STDOUT,
            timeout=COMMAND_TIMEOUT_S,
        ).strip()
    except (subprocess.CalledProcessError, subprocess.TimeoutExpired) as error:
        raise QualificationError(f"cannot resolve source revision {ref!r}: {error.output}") from error
    if not re.fullmatch(r"[0-9a-f]{40}", value):
        raise QualificationError(f"git returned a non-commit revision for {ref!r}")
    return value


def _source_files(root: Path) -> list[dict[str, Any]]:
    rows = []
    for path in sorted(item for item in root.rglob("*") if item.is_file()):
        row = artifact(path)
        row["path"] = str(path.relative_to(root))
        rows.append(row)
    return rows


def _source_content(row: Mapping[str, Any]) -> dict[str, Any]:
    """Return the path-independent identity retained in a source delta."""

    bytes_count = row.get("bytes")
    digest = row.get("sha256")
    if (
        not isinstance(bytes_count, int)
        or bytes_count < 0
        or not isinstance(digest, str)
        or re.fullmatch(r"[0-9a-f]{64}", digest) is None
    ):
        raise QualificationError("source delta contains an invalid content identity")
    return {"bytes": bytes_count, "sha256": digest}


def _validate_source_delta_path(path: str) -> None:
    if not isinstance(path, str) or not path or Path(path).is_absolute() or ".." in Path(path).parts:
        raise QualificationError("source delta has an invalid path")


def _source_delta_row(path: str, content: Mapping[str, Any], side: str) -> dict[str, Any]:
    _validate_source_delta_path(path)
    return {"path": path, side: _source_content(content)}


def source_delta(
    base: Mapping[str, Mapping[str, Any]],
    candidate: Mapping[str, Mapping[str, Any]],
    base_commit: str,
    candidate_commit: str,
) -> dict[str, Any]:
    """Describe every changed, added, or removed source file.

    The comparison is intentionally over the complete archived source tree,
    rather than a format or shared-crate prefix.  This permits an integrated
    route to compare the original baseline with the reviewed candidate while
    retaining an immutable, content-bound explanation of every source delta.
    """

    if re.fullmatch(r"[0-9a-f]{40}", base_commit) is None or re.fullmatch(
        r"[0-9a-f]{40}", candidate_commit
    ) is None:
        raise QualificationError("source delta refs must be full commit ids")
    base_only: list[dict[str, Any]] = []
    candidate_only: list[dict[str, Any]] = []
    changed: list[dict[str, Any]] = []
    for path in sorted(set(base) | set(candidate)):
        _validate_source_delta_path(path)
        before = base.get(path)
        after = candidate.get(path)
        if before is None:
            if after is None:
                raise QualificationError("source delta has an inconsistent file map")
            candidate_only.append(_source_delta_row(path, after, "candidate"))
        elif after is None:
            base_only.append(_source_delta_row(path, before, "base"))
        elif _source_content(before) != _source_content(after):
            changed.append(
                {
                    "path": path,
                    "base": _source_content(before),
                    "candidate": _source_content(after),
                }
            )
    return {
        "schema": SOURCE_DELTA_SCHEMA,
        "comparison": COMPARISON,
        "base_commit": base_commit,
        "candidate_commit": candidate_commit,
        "base_only": base_only,
        "candidate_only": candidate_only,
        "changed": changed,
    }


def verify_source_delta(
    value: Mapping[str, Any],
    source_maps: Mapping[str, Mapping[str, Mapping[str, Any]]],
    base_commit: str,
    candidate_commit: str,
) -> None:
    """Verify an integrated source delta against both retained source maps."""

    if not isinstance(value, Mapping):
        raise QualificationError("missing integrated source delta")
    expected = source_delta(
        source_maps["base"], source_maps["candidate"], base_commit, candidate_commit
    )
    if dict(value) != expected:
        raise QualificationError("integrated source delta does not match retained source files")


def _toml_file(path: Path, description: str) -> dict[str, Any]:
    try:
        import tomllib

        value = tomllib.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise QualificationError(f"cannot inspect {description} {path}: {error}") from error
    if not isinstance(value, dict):
        raise QualificationError(f"{description} is not a TOML table: {path}")
    return value


def _json_value(value: Any) -> Any:
    if isinstance(value, Mapping):
        return {str(key): _json_value(item) for key, item in sorted(value.items())}
    if isinstance(value, list):
        return [_json_value(item) for item in value]
    if isinstance(value, (str, int, float, bool)) or value is None:
        return value
    raise QualificationError(f"unsupported Cargo manifest value {value!r}")


def _source_dependency_entries(root: Path) -> tuple[dict[str, dict[str, Any]], set[str]]:
    """Read normal/build dependency declarations from an archived workspace."""

    entries: dict[str, dict[str, Any]] = {}
    packages: set[str] = set()
    manifests = sorted(path for path in root.rglob("Cargo.toml") if path.is_file())
    if not manifests:
        raise QualificationError(f"source snapshot has no Cargo manifests: {root}")
    workspace_dependencies: dict[str, Mapping[str, Any]] = {}
    for path in manifests:
        document = _toml_file(path, "Cargo manifest")
        workspace = document.get("workspace")
        if not isinstance(workspace, Mapping):
            continue
        dependencies = workspace.get("dependencies", {})
        if dependencies and not isinstance(dependencies, Mapping):
            raise QualificationError(f"{path} has malformed workspace dependencies")
        for dependency, specification in dependencies.items():
            if not isinstance(specification, (str, Mapping)):
                raise QualificationError(
                    f"{path} has malformed workspace dependency {dependency}"
                )
            workspace_dependencies[dependency] = specification
    for path in manifests:
        document = _toml_file(path, "Cargo manifest")
        package = document.get("package")
        if not isinstance(package, Mapping) or not isinstance(package.get("name"), str):
            continue
        owner = package["name"]
        packages.add(owner)
        sections: list[tuple[str, Mapping[str, Any]]] = []
        for section in ("dependencies", "build-dependencies", "dev-dependencies"):
            value = document.get(section, {})
            if value and not isinstance(value, Mapping):
                raise QualificationError(f"{path} has malformed {section}")
            sections.append((section, value))
        targets = document.get("target", {})
        if targets and not isinstance(targets, Mapping):
            raise QualificationError(f"{path} has malformed target dependencies")
        for target_name, target in sorted(targets.items()):
            if not isinstance(target, Mapping):
                raise QualificationError(f"{path} has malformed target {target_name}")
            for section in ("dependencies", "build-dependencies", "dev-dependencies"):
                value = target.get(section, {})
                if value and not isinstance(value, Mapping):
                    raise QualificationError(f"{path} has malformed target {section}")
                sections.append((f"target:{target_name}:{section}", value))
        for scope, dependencies in sections:
            for dependency, specification in sorted(dependencies.items()):
                if isinstance(specification, str):
                    normalized: Any = {"version": specification}
                    package_name = dependency
                    features: list[str] = []
                elif isinstance(specification, Mapping):
                    normalized = _json_value(specification)
                    workspace_specification = None
                    if specification.get("workspace") is True:
                        workspace_specification = workspace_dependencies.get(dependency)
                        if workspace_specification is None:
                            raise QualificationError(
                                f"{path} references missing workspace dependency {dependency}"
                            )
                        normalized["workspace_resolved"] = _json_value(workspace_specification)
                        if (
                            isinstance(workspace_specification, Mapping)
                            and "optional" in workspace_specification
                            and "optional" not in normalized
                        ):
                            normalized["optional"] = workspace_specification["optional"]
                    effective = (
                        {"version": workspace_specification}
                        if isinstance(workspace_specification, str)
                        else workspace_specification or specification
                    )
                    package_name = effective.get("package", dependency)
                    workspace_features = (
                        workspace_specification.get("features", [])
                        if isinstance(workspace_specification, Mapping)
                        else []
                    )
                    raw_features = list(workspace_features) + list(specification.get("features", []))
                    if not isinstance(package_name, str) or not isinstance(raw_features, list):
                        raise QualificationError(
                            f"{path} has malformed dependency declaration {dependency}"
                        )
                    if not all(isinstance(feature, str) for feature in raw_features):
                        raise QualificationError(
                            f"{path} has malformed dependency features {dependency}"
                        )
                    features = sorted(raw_features)
                else:
                    raise QualificationError(
                        f"{path} has malformed dependency declaration {dependency}"
                    )
                if not isinstance(package_name, str):
                    raise QualificationError(f"{path} has a non-string package name for {dependency}")
                key = json.dumps([owner, scope, dependency], separators=(",", ":"))
                entries[key] = {
                    "owner": owner,
                    "scope": scope,
                    "dependency": dependency,
                    "package": package_name,
                    "features": features,
                    "spec": normalized,
                }
    return entries, packages


def _source_dependency_active(entry: Mapping[str, Any]) -> bool:
    """Keep lock checks faithful to the host's active target dependencies."""

    specification = entry.get("spec")
    if isinstance(specification, Mapping) and specification.get("optional") is True:
        return False
    scope = entry.get("scope")
    if not isinstance(scope, str) or not scope.startswith("target:"):
        if scope == "dev-dependencies":
            return False
        return True
    parts = scope.split(":", 2)
    target = parts[1]
    if len(parts) == 3 and parts[2] == "dev-dependencies":
        return False
    if target == "cfg(unix)":
        return os.name == "posix"
    if target == "cfg(windows)":
        return os.name == "nt"
    # Unknown target expressions are not proven active by this host-only
    # harness; accepting them would permit a lock edge that Cargo will not
    # resolve for the measured target.
    return False


def _source_dependency_delta(
    base_source: Path, candidate_source: Path, family: str
) -> dict[str, Any]:
    base_entries, base_packages = _source_dependency_entries(base_source)
    candidate_entries, candidate_packages = _source_dependency_entries(candidate_source)
    local_packages = base_packages | candidate_packages
    roots = {
        "constructor": {"litchi-iwa-core"},
        "owner": {"litchi-numbers", "litchi-keynote", "litchi-pages"},
    }[family]
    if not roots <= local_packages:
        raise QualificationError(f"source snapshot lacks {family} probe dependency roots")
    closure = set(roots)
    while True:
        discovered = {
            entry["package"]
            for entries in (base_entries, candidate_entries)
            for entry in entries.values()
            if (
                entry["owner"] in closure
                and entry["package"] in local_packages
                and _source_dependency_active(entry)
            )
        }
        expanded = closure | discovered
        if expanded == closure:
            break
        closure = expanded

    def dependency_edge_map(
        entries: Mapping[str, Mapping[str, Any]], *, active: bool
    ) -> dict[str, list[str]]:
        result: dict[str, set[str]] = {owner: set() for owner in sorted(closure)}
        for entry in entries.values():
            if entry["owner"] in closure and (not active or _source_dependency_active(entry)):
                result[entry["owner"]].add(entry["package"])
        return {owner: sorted(values) for owner, values in result.items()}

    def delta_row(entry: Mapping[str, Any]) -> dict[str, Any]:
        return {
            "owner": entry["owner"],
            "scope": entry["scope"],
            "dependency": entry["dependency"],
            "package": entry["package"],
            "features": list(entry["features"]),
            "spec": entry["spec"],
        }

    added: list[dict[str, Any]] = []
    removed: list[dict[str, Any]] = []
    changed: list[dict[str, Any]] = []
    for key in sorted(set(base_entries) | set(candidate_entries)):
        base_entry = base_entries.get(key)
        candidate_entry = candidate_entries.get(key)
        owner = (base_entry or candidate_entry)["owner"]
        if owner not in closure:
            continue
        if base_entry is None:
            added.append(delta_row(candidate_entry))
        elif candidate_entry is None:
            removed.append(delta_row(base_entry))
        elif delta_row(base_entry) != delta_row(candidate_entry):
            changed.append({"base": delta_row(base_entry), "candidate": delta_row(candidate_entry)})

    def registry_names(rows: Iterable[Mapping[str, Any]]) -> list[str]:
        return sorted(
            {
                row["package"]
                for row in rows
                if row["package"] not in local_packages and _source_dependency_active(row)
            }
        )

    base_dependency_edges = dependency_edge_map(base_entries, active=True)
    candidate_dependency_edges = dependency_edge_map(candidate_entries, active=True)
    base_declared_dependency_edges = dependency_edge_map(base_entries, active=False)
    candidate_declared_dependency_edges = dependency_edge_map(candidate_entries, active=False)

    return {
        "schema": "litchi.iwa.dependency-delta.v2",
        "family": family,
        "packages": sorted(closure),
        "added": added,
        "removed": removed,
        "changed": changed,
        "base_dependency_edges": base_dependency_edges,
        "candidate_dependency_edges": candidate_dependency_edges,
        "base_declared_dependency_edges": base_declared_dependency_edges,
        "candidate_declared_dependency_edges": candidate_declared_dependency_edges,
        "base_local_edges": {
            owner: sorted(set(values) & local_packages)
            for owner, values in base_dependency_edges.items()
        },
        "candidate_local_edges": {
            owner: sorted(set(values) & local_packages)
            for owner, values in candidate_dependency_edges.items()
        },
        "registry_added": registry_names(added),
        "registry_removed": registry_names(removed),
        "registry_changed": sorted(
            {
                row["base"]["package"]
                for row in changed
                if row["base"]["package"] not in local_packages
            }
            | {
                row["candidate"]["package"]
                for row in changed
                if row["candidate"]["package"] not in local_packages
            }
        ),
    }


def _lock_rows(path: Path) -> tuple[int, list[dict[str, Any]]]:
    data = _toml_file(path, "Cargo.lock")
    version = data.get("version")
    packages = data.get("package")
    if not isinstance(version, int) or not isinstance(packages, list):
        raise QualificationError(f"Cargo.lock has no strict version/package table: {path}")
    rows = []
    for package in packages:
        if not isinstance(package, Mapping) or not isinstance(package.get("name"), str):
            raise QualificationError(f"Cargo.lock has a malformed package row: {path}")
        if not isinstance(package.get("version"), str):
            raise QualificationError(f"Cargo.lock has a malformed package version: {path}")
        dependencies = package.get("dependencies", [])
        if dependencies and not isinstance(dependencies, list):
            raise QualificationError(f"Cargo.lock has malformed dependencies: {path}")
        if not all(isinstance(dependency, str) for dependency in dependencies):
            raise QualificationError(f"Cargo.lock has non-string dependencies: {path}")
        rows.append(
            {
                "name": package["name"],
                "version": package.get("version"),
                "source": package.get("source"),
                "checksum": package.get("checksum"),
                "dependencies": list(dependencies),
            }
        )
    return version, rows


def _registry_rows(rows: Iterable[Mapping[str, Any]]) -> list[dict[str, Any]]:
    result = []
    seen: dict[tuple[str, str, str], str] = {}
    for row in rows:
        source = row.get("source")
        if not isinstance(source, str) or not source.startswith("registry+"):
            continue
        checksum = row.get("checksum")
        if not isinstance(checksum, str):
            raise QualificationError(f"registry package has no checksum: {row['name']}")
        identity = (row["name"], row["version"], source)
        previous = seen.get(identity)
        if previous is not None:
            if previous != checksum:
                raise QualificationError(
                    f"Cargo.lock has conflicting checksums for {row['name']} {row['version']}"
                )
            raise QualificationError(f"Cargo.lock has a duplicate registry package: {row['name']}")
        seen[identity] = checksum
        result.append(
            {
                "name": row["name"],
                "version": row["version"],
                "source": source,
                "checksum": checksum,
            }
        )
    return sorted(result, key=lambda row: (row["name"], row["version"], row["source"]))


def _dependency_edges(rows: Iterable[Mapping[str, Any]], package_names: set[str]) -> dict[str, list[str]]:
    result: dict[str, list[str]] = {}
    for row in rows:
        if row["name"] not in package_names or row.get("source") is not None:
            continue
        if row["name"] in result:
            raise QualificationError(f"Cargo.lock has duplicate local package {row['name']}")
        dependencies = {
            dependency.split(" ", 1)[0]
            for dependency in row["dependencies"]
        }
        result[row["name"]] = sorted(dependencies)
    if set(result) != package_names:
        missing = sorted(package_names - set(result))
        raise QualificationError(f"Cargo.lock lacks selected source packages: {missing}")
    return result


def _registry_comparison(
    base: list[dict[str, Any]], candidate: list[dict[str, Any]]
) -> dict[str, Any]:
    def by_identity(
        rows: Iterable[Mapping[str, Any]],
    ) -> dict[tuple[str, str, str], dict[str, Any]]:
        result: dict[tuple[str, str, str], dict[str, Any]] = {}
        for row in rows:
            result[(row["name"], row["version"], row["source"])] = dict(row)
        return result

    base_by_identity = by_identity(base)
    candidate_by_identity = by_identity(candidate)
    added = [
        candidate_by_identity[key]
        for key in sorted(set(candidate_by_identity) - set(base_by_identity))
    ]
    removed = [
        base_by_identity[key]
        for key in sorted(set(base_by_identity) - set(candidate_by_identity))
    ]
    changed = [
        {
            "name": key[0],
            "version": key[1],
            "source": key[2],
            "base": base_by_identity[key],
            "candidate": candidate_by_identity[key],
        }
        for key in sorted(set(base_by_identity) & set(candidate_by_identity))
        if base_by_identity[key] != candidate_by_identity[key]
    ]
    common = [
        base_by_identity[key]
        for key in sorted(set(base_by_identity) & set(candidate_by_identity))
    ]
    return {
        "base": base,
        "candidate": candidate,
        "common": common,
        "added": added,
        "removed": removed,
        "changed": changed,
    }


def _registry_identity(row: Mapping[str, Any]) -> tuple[str, str, str, str]:
    """Return the complete immutable identity of a registry package row."""

    return (row["name"], row["version"], row["source"], row["checksum"])


def lock_delta_manifest(
    family: str,
    base_source: Path,
    candidate_source: Path,
    base_lock: Path,
    candidate_lock: Path,
    seed_lock: Path,
) -> dict[str, Any]:
    """Bind exact endpoint locks to the source-declared dependency delta."""

    source_dependencies = _source_dependency_delta(base_source, candidate_source, family)
    base_version, base_rows = _lock_rows(base_lock)
    candidate_version, candidate_rows = _lock_rows(candidate_lock)
    seed_version, seed_rows = _lock_rows(seed_lock)
    base_registry = _registry_rows(base_rows)
    candidate_registry = _registry_rows(candidate_rows)
    seed_registry = _registry_rows(seed_rows)
    registry_delta = _registry_comparison(base_registry, candidate_registry)
    base_by_name: dict[str, set[tuple[str, str, str, str]]] = {}
    candidate_by_name: dict[str, set[tuple[str, str, str, str]]] = {}
    for row in base_registry:
        base_by_name.setdefault(row["name"], set()).add(_registry_identity(row))
    for row in candidate_registry:
        candidate_by_name.setdefault(row["name"], set()).add(_registry_identity(row))
    expected_added = {
        name
        for name in source_dependencies["registry_added"]
        if candidate_by_name.get(name, set()) - base_by_name.get(name, set())
    }
    expected_removed = {
        name
        for name in source_dependencies["registry_removed"]
        if base_by_name.get(name, set()) - candidate_by_name.get(name, set())
    }
    actual_added = {row["name"] for row in registry_delta["added"]}
    actual_removed = {row["name"] for row in registry_delta["removed"]}
    if actual_added != expected_added or actual_removed != expected_removed:
        raise QualificationError(
            "Cargo.lock registry additions/removals do not match source-declared dependencies"
        )
    if registry_delta["changed"]:
        raise QualificationError("common Cargo.lock registry identities changed")
    reviewed_additions = REVIEWED_REGISTRY_ADDITIONS[family]
    if registry_delta["added"] != reviewed_additions:
        raise QualificationError("Cargo.lock additions do not match reviewed registry identities")
    seed_identities = {_registry_identity(row) for row in seed_registry}
    for row in reviewed_additions:
        if _registry_identity(row) not in seed_identities:
            raise QualificationError("reviewed registry addition is absent from the seed lock")
    missing_common_seed = [
        row for row in registry_delta["common"] if _registry_identity(row) not in seed_identities
    ]
    if missing_common_seed:
        raise QualificationError(
            "common Cargo.lock registry identities are absent from the pinned seed lock"
        )
    package_names = set(source_dependencies["packages"])
    base_edges = _dependency_edges(base_rows, package_names)
    candidate_edges = _dependency_edges(candidate_rows, package_names)
    for label, actual, expected, declared in (
        (
            "baseline",
            base_edges,
            source_dependencies["base_dependency_edges"],
            source_dependencies["base_declared_dependency_edges"],
        ),
        (
            "candidate",
            candidate_edges,
            source_dependencies["candidate_dependency_edges"],
            source_dependencies["candidate_declared_dependency_edges"],
        ),
    ):
        missing_edges = {
            owner: sorted(set(expected[owner]) - set(actual[owner]))
            for owner in expected
            if set(expected[owner]) - set(actual[owner])
        }
        if missing_edges:
            raise QualificationError(
                f"{label} Cargo.lock lacks source-declared dependency edges: {missing_edges}"
            )
        undeclared_edges = {
            owner: sorted(set(actual[owner]) - set(declared[owner]))
            for owner in actual
            if set(actual[owner]) - set(declared[owner])
        }
        if undeclared_edges:
            raise QualificationError(
                f"{label} Cargo.lock has undeclared dependency edges: {undeclared_edges}"
            )
    required_seed_rows: dict[str, list[dict[str, Any]]] = {}
    for label, edges, rows in (
        ("baseline", base_edges, base_rows),
        ("candidate", candidate_edges, candidate_rows),
    ):
        available_registry = {row["name"] for row in _registry_rows(rows)}
        required = {dependency for values in edges.values() for dependency in values}
        required_registry = sorted(required - package_names)
        missing = sorted(
            dependency for dependency in required_registry if dependency not in available_registry
        )
        if missing:
            raise QualificationError(f"{label} lock lacks required dependency packages: {missing}")
        endpoint_required_rows = sorted(
            (row for row in _registry_rows(rows) if row["name"] in required_registry),
            key=lambda row: (row["name"], row["version"], row["source"]),
        )
        missing_seed = [
            row for row in endpoint_required_rows if _registry_identity(row) not in seed_identities
        ]
        if missing_seed:
            raise QualificationError(
                f"{label} required registry identities are absent from the pinned seed lock"
            )
        required_seed_rows[label] = endpoint_required_rows
    return {
        "schema": LOCK_DELTA_SCHEMA,
        "family": family,
        "source_dependencies": source_dependencies,
        "reviewed_registry_additions": reviewed_additions,
        "locks": {
            "base": artifact(base_lock),
            "candidate": artifact(candidate_lock),
            "seed": artifact(seed_lock),
            "base_version": base_version,
            "candidate_version": candidate_version,
            "seed_version": seed_version,
        },
        "registry": registry_delta,
        "seed_registry": {
            "common": registry_delta["common"],
            "required": required_seed_rows,
        },
        "local_edges": {"base": base_edges, "candidate": candidate_edges},
        "declared_edges": {
            "base": source_dependencies["base_declared_dependency_edges"],
            "candidate": source_dependencies["candidate_declared_dependency_edges"],
        },
    }


def verify_lock_delta(
    value: Mapping[str, Any],
    family: str,
    base_source: Path,
    candidate_source: Path,
    base_lock: Path,
    candidate_lock: Path,
    seed_lock: Path,
) -> None:
    if not isinstance(value, Mapping):
        raise QualificationError("missing endpoint lock comparison")
    expected = lock_delta_manifest(
        family, base_source, candidate_source, base_lock, candidate_lock, seed_lock
    )
    if dict(value) != expected:
        raise QualificationError("endpoint lock comparison does not match retained artifacts")


def snapshot_source(output: Path, label: str, revision: str, lock: Path) -> tuple[Path, dict[str, Any]]:
    commit = git_revision(revision)
    source = output / f"source-{label}"
    archive = output / f"source-{label}.tar"
    source.mkdir()
    with archive.open("wb") as stream:
        try:
            subprocess.run(
                ["git", "archive", "--format=tar", commit, "Cargo.toml", "rust-toolchain.toml", "crates"],
                cwd=ROOT,
                stdout=stream,
                stderr=subprocess.PIPE,
                check=True,
                timeout=COMMAND_TIMEOUT_S,
            )
        except (OSError, subprocess.SubprocessError) as error:
            raise QualificationError(f"cannot snapshot {label} source") from error
    with tarfile.open(archive) as stream:
        stream.extractall(source, filter="data")
    shutil.copyfile(lock, source / "Cargo.lock")
    manifest = output / f"source-{label}.json"
    save_json(manifest, {"commit": commit, "archive": artifact(archive), "files": _source_files(source)})
    return source, {"commit": commit, "manifest": artifact(manifest)}


def registry(lock: Path) -> list[dict[str, Any]]:
    _version, rows = _lock_rows(lock)
    return _registry_rows(rows)


def _project_manifest(family: str, source: Path) -> str:
    if family == "constructor":
        dependency = f'litchi-iwa-core = {{ path = {json.dumps(str(source / "crates/litchi-iwa-core"))} }}'
        name = "iwa-constructor-probe"
    else:
        dependency = "\n".join(
            f'{name} = {{ path = {json.dumps(str(source / "crates" / name))} }}'
            for name in ("litchi-numbers", "litchi-keynote", "litchi-pages")
        )
        name = "iwa-owner-probe"
    return f"""[package]
name = {json.dumps(name)}
version = "0.0.1"
edition = "2024"
publish = false

[features]
default = []
before = []
allocations = []

[dependencies]
{dependency}
sha2 = "0.11"

[profile.release]
lto = true
codegen-units = 1
panic = "abort"

[workspace]
"""


def prepare_endpoint_lock(
    receipts: Path, project: Path, env: Mapping[str, str]
) -> dict[str, Any]:
    """Resolve only the probe/local closure while retaining the copied seed pins."""

    return run_command(
        receipts,
        "metadata",
        list(LOCK_PREPARATION_COMMAND),
        project,
        env,
    )


def build(args: argparse.Namespace) -> None:
    family = require_family(args.family)
    output = new_external_directory(args.output)
    lock = args.lock.resolve(strict=True)
    probe = PROBES[family]
    if not probe.is_file():
        raise QualificationError(f"missing qualification probe {probe}")
    manifest: dict[str, Any] = {
        "schema": SCHEMA,
        "family": family,
        "comparison": COMPARISON,
        "requested_refs": {"base": args.base, "candidate": args.candidate},
        "status": "building",
        "routes": list(ROUTES[family]),
        "builder": artifact(Path(__file__)),
        "probe": artifact(probe),
        "input_lock": artifact(lock),
        "endpoints": {},
    }
    save_json(output / "build.json", manifest)
    snapshots: dict[str, tuple[Path, dict[str, Any]]] = {}
    source_maps: dict[str, dict[str, dict[str, Any]]] = {}
    for label, revision in (("base", args.base), ("candidate", args.candidate)):
        source, provenance = snapshot_source(output, label, revision, lock)
        snapshots[label] = (source, provenance)
        source_manifest = load_json(Path(provenance["manifest"]["path"]))
        source_maps[label] = {
            row["path"]: _source_content(row) for row in source_manifest["files"]
        }
    resolved_refs = {
        label: snapshots[label][1]["commit"] for label in ("base", "candidate")
    }
    manifest["resolved_refs"] = resolved_refs
    manifest["source_delta"] = source_delta(
        source_maps["base"], source_maps["candidate"], resolved_refs["base"], resolved_refs["candidate"]
    )
    verify_source_delta(manifest["source_delta"], source_maps, resolved_refs["base"], resolved_refs["candidate"])
    # The complete source comparison is admitted before the first Cargo child
    # starts.  Subsequent consumers independently recompute it from the
    # retained source manifests.
    save_json(output / "build.json", manifest)
    projects: dict[str, tuple[Path, Path, list[dict[str, Any]], Mapping[str, str]]] = {}
    for label in ("base", "candidate"):
        source, provenance = snapshots[label]
        project = output / f"project-{label}"
        (project / "src").mkdir(parents=True)
        shutil.copyfile(probe, project / "src/main.rs")
        shutil.copyfile(source / "rust-toolchain.toml", project / "rust-toolchain.toml")
        (project / "Cargo.toml").write_text(_project_manifest(family, source), encoding="utf-8")
        receipts = output / f"build-{label}"
        receipts.mkdir()
        env = dict(os.environ, CARGO_INCREMENTAL="0", CARGO_BUILD_JOBS="1")
        shutil.copyfile(lock, project / "Cargo.lock")
        rows = [prepare_endpoint_lock(receipts, project, env)]
        rows.append(run_command(receipts, "rustc", ["rustc", "-Vv"], project, env))
        rows.append(run_command(receipts, "cargo", ["cargo", "--version"], project, env))
        endpoint: dict[str, Any] = {
            "source": provenance,
            "project_manifest": artifact(project / "Cargo.toml"),
            "project_source": artifact(project / "src/main.rs"),
            "lock": artifact(project / "Cargo.lock"),
            "registry": registry(project / "Cargo.lock"),
            "rustc": (receipts / "rustc.stdout").read_text(encoding="utf-8"),
            "cargo": (receipts / "cargo.stdout").read_text(encoding="utf-8"),
            "commands": rows,
            "binaries": {},
        }
        manifest["endpoints"][label] = endpoint
        projects[label] = (project, receipts, rows, env)
        save_json(output / "build.json", manifest)
    base = manifest["endpoints"]["base"]
    candidate = manifest["endpoints"]["candidate"]
    for key in ("rustc", "cargo"):
        if base[key] != candidate[key]:
            raise QualificationError(f"matched endpoints disagree on {key}")
    manifest["lock_delta"] = lock_delta_manifest(
        family,
        snapshots["base"][0],
        snapshots["candidate"][0],
        Path(base["lock"]["path"]),
        Path(candidate["lock"]["path"]),
        lock,
    )
    verify_lock_delta(
        manifest["lock_delta"],
        family,
        snapshots["base"][0],
        snapshots["candidate"][0],
        Path(base["lock"]["path"]),
        Path(candidate["lock"]["path"]),
        lock,
    )
    # Endpoint lockfiles are now independently pinned and admitted.  Only
    # after the source/lock/toolchain comparison passes may release binaries
    # be built.
    save_json(output / "build.json", manifest)
    package_name = "iwa-constructor-probe" if family == "constructor" else "iwa-owner-probe"
    for label in ("base", "candidate"):
        project, receipts, rows, env = projects[label]
        endpoint = manifest["endpoints"][label]
        for variant in ("native", "allocation"):
            features: list[str] = []
            if family == "constructor" and label == "base":
                features.append("before")
            if variant == "allocation":
                features.append("allocations")
            command = [
                "cargo",
                "build",
                "--offline",
                "--locked",
                "--release",
                "--message-format=json-render-diagnostics",
            ]
            if features:
                command += ["--features", ",".join(features)]
            target = output / f"target-{label}-{variant}"
            rows.append(run_command(receipts, variant, command, project, dict(env, CARGO_TARGET_DIR=str(target))))
            built = target / "release" / package_name
            if not built.is_file():
                raise QualificationError(f"cargo did not produce expected probe {built}")
            protected = receipts / f"{family}-{variant}"
            shutil.copy2(built, protected)
            protected.chmod(0o555)
            endpoint["binaries"][variant] = artifact(protected)
            save_json(output / "build.json", manifest)
    manifest["status"] = "complete"
    save_json(output / "build.json", manifest)


def load_build(output: Path, family: str) -> dict[str, Any]:
    build_path = output / "build.json"
    data = load_json(build_path)
    if data.get("schema") != SCHEMA or data.get("family") != family or data.get("status") != "complete":
        raise QualificationError("build manifest is not a complete matching qualification build")
    for key in ("builder", "probe", "input_lock"):
        verify_artifact(data[key])
    if data["probe"]["path"] != str(PROBES[family].resolve()):
        raise QualificationError("build was made with a different probe source")
    if data.get("comparison") != COMPARISON:
        raise QualificationError("build manifest is not an integrated-route comparison")
    resolved_refs = data.get("resolved_refs")
    if not isinstance(resolved_refs, Mapping) or set(resolved_refs) != {"base", "candidate"}:
        raise QualificationError("build manifest lacks resolved baseline/candidate refs")
    source_maps: dict[str, dict[str, dict[str, Any]]] = {}
    for label in ("base", "candidate"):
        endpoint = data["endpoints"].get(label)
        if not isinstance(endpoint, Mapping):
            raise QualificationError(f"missing {label} endpoint")
        verify_artifact(endpoint["source"]["manifest"])
        source_manifest = load_json(Path(endpoint["source"]["manifest"]["path"]))
        verify_artifact(source_manifest["archive"])
        source_dir = output / f"source-{label}"
        source_maps[label] = {}
        for row in source_manifest["files"]:
            actual = artifact(source_dir / row["path"])
            if (actual["bytes"], actual["sha256"]) != (row["bytes"], row["sha256"]):
                raise QualificationError(f"source snapshot changed: {label}/{row['path']}")
            source_maps[label][row["path"]] = _source_content(row)
        if endpoint["source"].get("commit") != resolved_refs[label]:
            raise QualificationError(f"{label} source ref does not match the integrated comparison")
        for key in ("project_manifest", "project_source", "lock"):
            verify_artifact(endpoint[key])
        lock_path = Path(endpoint["lock"]["path"])
        if endpoint.get("registry") != registry(lock_path):
            raise QualificationError(f"retained {label} registry identity does not match its lock")
        commands = endpoint.get("commands")
        verify_lock_preparation_command(commands)
        for command in commands:
            verify_command_receipt(command)
        for key in ("native", "allocation"):
            verify_artifact(endpoint["binaries"][key])
    base = data["endpoints"]["base"]
    candidate = data["endpoints"]["candidate"]
    for key in ("rustc", "cargo"):
        if base[key] != candidate[key]:
            raise QualificationError(f"matched endpoints disagree on {key}")
    verify_source_delta(
        data.get("source_delta"),
        source_maps,
        resolved_refs["base"],
        resolved_refs["candidate"],
    )
    verify_lock_delta(
        data.get("lock_delta"),
        family,
        output / "source-base",
        output / "source-candidate",
        Path(base["lock"]["path"]),
        Path(candidate["lock"]["path"]),
        Path(data["input_lock"]["path"]),
    )
    return data


def _validation_command(family: str, binary: Path, fixtures: Path | None, destination: Path) -> list[str]:
    if family == "owner":
        if fixtures is None:
            raise QualificationError("owner validation requires --fixtures")
        return [str(binary), "validate_outputs", str(fixtures), str(destination)]
    return [str(binary), "validate_outputs", str(destination)]


def _validate_report_value(value: Any, family: str) -> None:
    expected_schema = f"litchi.iwa.{family}.validation.v1"
    if not isinstance(value, dict) or value.get("schema") != expected_schema or value.get("status") != "ok":
        raise QualificationError(f"invalid {family} validation report schema")
    routes = value.get("routes")
    if routes != list(ROUTES[family]):
        raise QualificationError(f"{family} validation report does not cover every route")
    if family == "constructor":
        for key in (
            "source_sha256",
            "encoded_sha256",
            "retained_source_sha256",
            "retained_encoded_sha256",
        ):
            if re.fullmatch(r"[0-9a-f]{64}", str(value.get(key))) is None:
                raise QualificationError(f"constructor validation lacks {key}")
        for key in ("source_bytes", "encoded_bytes"):
            if not isinstance(value.get(key), int) or value[key] <= 0:
                raise QualificationError(f"constructor validation lacks {key}")
        if value.get("roundtrip_equal") is not True or value.get("retained_roundtrip_equal") is not True:
            raise QualificationError("constructor validation did not prove round trips")
        boundary = value.get("header_boundary")
        below = boundary.get("below") if isinstance(boundary, Mapping) else None
        if (
            not isinstance(boundary, Mapping)
            or boundary.get("kind") != "HeaderBytes"
            or not isinstance(boundary.get("accepted"), int)
            or boundary["accepted"] <= 1
            or not isinstance(below, Mapping)
            or not isinstance(below.get("observed"), int)
            or not isinstance(below.get("maximum"), int)
            or below["observed"] <= below["maximum"]
            or below["maximum"] != boundary["accepted"] - 1
        ):
            raise QualificationError("constructor validation lacks the tight header boundary proof")
        retained_boundary = value.get("retained_header_boundary")
        retained_below = (
            retained_boundary.get("below") if isinstance(retained_boundary, Mapping) else None
        )
        if (
            not isinstance(retained_boundary, Mapping)
            or retained_boundary.get("kind") != "HeaderFields"
            or not isinstance(retained_boundary.get("accepted"), int)
            or retained_boundary["accepted"] <= 1
            or not isinstance(retained_below, Mapping)
            or not isinstance(retained_below.get("observed"), int)
            or not isinstance(retained_below.get("maximum"), int)
            or retained_below["observed"] <= retained_below["maximum"]
            or retained_below["maximum"] != retained_boundary["accepted"] - 1
        ):
            raise QualificationError("constructor validation lacks the retained header-field boundary proof")
    checks = value.get("route_checks")
    if not isinstance(checks, list) or len(checks) != len(ROUTES[family]):
        raise QualificationError(f"{family} validation report has incomplete route checks")
    seen: list[str] = []
    for check in checks:
        if not isinstance(check, Mapping) or not isinstance(check.get("route"), str):
            raise QualificationError(f"{family} validation route check is malformed")
        seen.append(check["route"])
        if family == "constructor":
            if check.get("status") != "ok" or not isinstance(check.get("output_bytes"), int) or check["output_bytes"] <= 0:
                raise QualificationError("constructor validation route fields are incomplete")
        else:
            correctness = check.get("correctness")
            if not isinstance(correctness, Mapping):
                raise QualificationError("owner validation route lacks correctness fields")
            for key in (
                "source_sha256",
                "changed_sha256",
                "replay_sha256",
                "inverse_sha256",
                "noop_sha256",
            ):
                if re.fullmatch(r"[0-9a-f]{64}", str(correctness.get(key))) is None:
                    raise QualificationError(f"owner validation route lacks {key}")
            for key in ("source_bytes", "changed_bytes", "noop_bytes"):
                if not isinstance(correctness.get(key), int) or correctness[key] <= 0:
                    raise QualificationError(f"owner validation route lacks {key}")
            for key in (
                "changed_differs",
                "replay_matches_changed",
                "inverse_matches_source",
                "noop_is_noop",
                "noop_matches_expected",
                "semantic_ok",
            ):
                if correctness.get(key) is not True:
                    raise QualificationError(f"owner validation route lacks successful {key}")
    if seen != list(ROUTES[family]):
        raise QualificationError(f"{family} validation route checks are duplicated or reordered")


def validate(args: argparse.Namespace) -> None:
    family = require_family(args.family)
    output = outside_checkout(args.output)
    build = load_build(output, family)
    fixtures = outside_checkout(args.fixtures) if args.fixtures is not None else None
    fixture_manifest = None
    if family == "owner" and fixtures is None:
        raise QualificationError("owner validation requires --fixtures")
    if fixtures is not None:
        fixture_manifest = fixture_identity(fixtures, regenerate=True)
    rows: dict[str, Any] = {}
    validation_root = output / "validation"
    validation_root.mkdir(exist_ok=False)
    for label in ("base", "candidate"):
        rows[label] = {}
        for variant in ("native", "allocation"):
            endpoint = build["endpoints"][label]
            destination = validation_root / f"{label}-{variant}"
            destination.mkdir()
            receipt = run_command(
                destination,
                "probe",
                _validation_command(family, Path(endpoint["binaries"][variant]["path"]), fixtures, destination),
                output,
            )
            report = destination / "validation.json"
            value = load_json(report)
            _validate_report_value(value, family)
            rows[label][variant] = {
                "report": artifact(report),
                "receipt": receipt["receipt"],
                "value": value,
            }
    for variant in ("native", "allocation"):
        if rows["base"][variant]["value"] != rows["candidate"][variant]["value"]:
            raise QualificationError(f"baseline/candidate {variant} correctness differs")
    result = {
        "schema": SCHEMA,
        "family": family,
        "status": "complete",
        "build": artifact(output / "build.json"),
        "comparison": build["comparison"],
        "resolved_refs": build["resolved_refs"],
        "source_delta": build["source_delta"],
        "lock_delta": build["lock_delta"],
        "fixtures": fixture_manifest,
        "endpoints": rows,
    }
    save_json(output / "validation.json", result)


def load_validation(output: Path, family: str, build: Mapping[str, Any]) -> dict[str, Any]:
    path = output / "validation.json"
    data = load_json(path)
    if data.get("schema") != SCHEMA or data.get("family") != family or data.get("status") != "complete":
        raise QualificationError("correctness validation is missing or incomplete")
    if data.get("build", {}).get("sha256") != artifact(output / "build.json")["sha256"]:
        raise QualificationError("correctness validation does not match build manifest")
    for key in ("comparison", "resolved_refs", "source_delta", "lock_delta"):
        if data.get(key) != build.get(key):
            raise QualificationError(f"correctness validation does not retain build {key}")
    for label in ("base", "candidate"):
        for variant in ("native", "allocation"):
            row = data["endpoints"][label][variant]
            verify_artifact(row["report"])
            verify_command_receipt(
                {
                    **load_json(Path(row["receipt"]["path"])),
                    "receipt": row["receipt"],
                }
            )
            _validate_report_value(row["value"], family)
    if data["endpoints"]["base"]["native"]["value"] != data["endpoints"]["candidate"]["native"]["value"]:
        raise QualificationError("baseline/candidate native correctness differs")
    if data["endpoints"]["base"]["allocation"]["value"] != data["endpoints"]["candidate"]["allocation"]["value"]:
        raise QualificationError("baseline/candidate allocation correctness differs")
    return data


def collect(args: argparse.Namespace) -> None:
    """Verify a completed build/validation/run packet for handoff.

    Collection is deliberately read-only with respect to source and binary
    provenance.  ``load_build`` recomputes the complete integrated source
    delta, so this consumer cannot silently fall back to the removed
    shared-prefix equality rule.
    """

    family = require_family(args.family)
    output = outside_checkout(args.output)
    build = load_build(output, family)
    validation = load_validation(output, family, build)
    run_value: dict[str, Any] | None = None
    if args.attempt is not None:
        attempt = output / args.attempt
        run_manifest = load_json(attempt / "manifest.json")
        if (
            run_manifest.get("schema") != SCHEMA
            or run_manifest.get("family") != family
            or run_manifest.get("status") != "complete"
        ):
            raise QualificationError("attempt is not a complete matching run")
        if run_manifest.get("build") != artifact(output / "build.json"):
            raise QualificationError("native attempt does not match the build manifest")
        if run_manifest.get("validation") != artifact(output / "validation.json"):
            raise QualificationError("native attempt does not match the validation manifest")
        for key in ("comparison", "resolved_refs", "source_delta", "lock_delta"):
            if run_manifest.get(key) != build.get(key):
                raise QualificationError(f"native attempt does not retain build {key}")
        for label in ("base", "candidate"):
            if run_manifest.get("sources", {}).get(label) != build["endpoints"][label]["source"]:
                raise QualificationError(f"native attempt source identity changed: {label}")
        phase = run_manifest.get("phase", "native")
        if phase not in ("native", "allocation"):
            raise QualificationError(f"unknown qualification attempt phase: {phase}")
        if phase == "allocation":
            protocol = run_manifest.get("protocol")
            if (
                not isinstance(protocol, Mapping)
                or protocol.get("iterations") != ALLOCATION_ITERATIONS
                or protocol.get("order") != ALLOCATION_ORDER
                or protocol.get("routes") != list(ROUTES[family])
                or protocol.get("timeout_s") != FIXED_TIMEOUT_S
            ):
                raise QualificationError("allocation attempt does not retain the fixed protocol")
            cases = run_manifest.get("cases")
            if not isinstance(cases, list) or len(cases) != len(ROUTES[family]):
                raise QualificationError("allocation attempt does not cover every route")
            for route, case in zip(ROUTES[family], cases):
                if (
                    not isinstance(case, Mapping)
                    or case.get("route") != route
                    or case.get("iterations") != ALLOCATION_ITERATIONS
                    or case.get("order") != ALLOCATION_ORDER
                ):
                    raise QualificationError("allocation attempt route schedule is malformed")
                legs = case.get("legs")
                if not isinstance(legs, list) or len(legs) != len(ALLOCATION_ORDER):
                    raise QualificationError("allocation attempt is missing an ABBA leg")
                values = [leg.get("value") for leg in legs if isinstance(leg, Mapping)]
                if len(values) != len(ALLOCATION_ORDER):
                    raise QualificationError("allocation attempt contains a malformed ABBA leg")
                _allocation_abba_equal(values)
                for value in values:
                    _verify_allocation_receipt(value)
            if run_manifest.get("qualified") is not True:
                raise QualificationError("allocation attempt is not qualified")
        run_value = {
            "manifest": artifact(attempt / "manifest.json"),
            "phase": phase,
            "qualified": run_manifest.get("qualified"),
        }
    result = {
        "schema": SCHEMA,
        "family": family,
        "status": "complete",
        "build": artifact(output / "build.json"),
        "validation": artifact(output / "validation.json"),
        "comparison": build["comparison"],
        "resolved_refs": build["resolved_refs"],
        "source_delta": build["source_delta"],
        "lock_delta": build["lock_delta"],
        "run": run_value,
    }
    destination = outside_checkout(args.out) if args.out is not None else output / "collection.json"
    save_json(destination, result)


def _program(arguments: str) -> str:
    try:
        words = shlex.split(arguments)
    except ValueError:
        words = arguments.split()
    if not words:
        return ""
    return os.path.basename(words[0])


FORBIDDEN_PROCESS = re.compile(
    r"(?:^(?:cargo|rustc|rustdoc|rustfmt|clippy-driver|heaptrack|perf|valgrind|"
    r"callgrind|clang(?:\+\+)?|gcc|g\+\+|cc|cmake|ninja|make|ld|lld|mold|"
    r"cc1(?:plus)?|nextest|pytest)$|"
    r"(?:run-abba|run-matched|counting-allocator|constructor|owner-probe|"
    r"iwa-native|litchi-perf-baseline|file_w8_.*probe|xlsx_.*probe|ole_.*probe|"
    r"scoped_patch_probe|test-adblock|peculiar))",
    re.IGNORECASE,
)


def process_rows() -> list[dict[str, Any]]:
    rows = []
    for entry in Path("/proc").iterdir():
        if not entry.name.isdigit():
            continue
        try:
            stat = (entry / "stat").read_text()
            parent = int(stat[stat.rfind(")") + 2 :].split()[1])
            arguments = (entry / "cmdline").read_bytes().replace(b"\0", b" ").decode(errors="replace")
            executable = os.readlink(entry / "exe")
            rows.append({"pid": int(entry.name), "parent": parent, "exe": executable, "args": arguments})
        except (FileNotFoundError, ProcessLookupError, PermissionError, ValueError):
            continue
    return rows


def _descendants(rows: Iterable[Mapping[str, Any]], root: int) -> set[int]:
    result = {root}
    while True:
        expanded = result | {int(row["pid"]) for row in rows if int(row["parent"]) in result}
        if expanded == result:
            return result
        result = expanded


def host_snapshot(output: Path, label: str, child: int | None = None) -> None:
    rows = process_rows()
    parents = {int(row["pid"]): int(row["parent"]) for row in rows}
    exempt = _descendants(rows, os.getpid())
    if child is not None:
        exempt |= _descendants(rows, child)
    blockers = []
    for row in rows:
        if int(row["pid"]) in exempt:
            continue
        executable = os.path.basename(str(row["exe"]))
        program = _program(str(row["args"]))
        arguments = str(row["args"])
        if (
            FORBIDDEN_PROCESS.search(executable)
            or FORBIDDEN_PROCESS.search(program)
            or FORBIDDEN_PROCESS.search(arguments)
        ):
            blockers.append(row)
    snapshot = {
        "time_ns": time.time_ns(),
        "load": Path("/proc/loadavg").read_text().split()[:3],
        "exempt": sorted(exempt),
        "blockers": blockers,
        "processes": rows,
    }
    save_json(output / f"{label}.host.json", snapshot)
    if blockers:
        raise QualificationError(f"host overlap at {label}; timing is unqualified")


def _strict_probe_value(
    value: Any,
    family: str,
    route: str,
    count: int,
    *,
    require_allocations: bool = False,
) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise QualificationError("probe did not emit one JSON object")
    expected_schema = f"litchi.iwa.{family}.probe.v2" if family == "owner" else "litchi.iwa.constructor.probe.v3"
    if value.get("schema") != expected_schema:
        raise QualificationError("probe schema does not match the selected family")
    if value.get("route") != route or value.get("iterations") != count:
        raise QualificationError(f"probe identity mismatch for {route}/{count}")
    if not isinstance(value.get("ns"), int) or value["ns"] <= 0:
        raise QualificationError("probe did not report a positive integer duration")
    if not isinstance(value.get("checksum"), int):
        raise QualificationError("probe did not report a semantic checksum")
    if not isinstance(value.get("output_bytes"), int) or value["output_bytes"] <= 0:
        raise QualificationError("probe did not report a positive output byte count")
    for key in ("source_sha256", "output_sha256"):
        if not isinstance(value.get(key), str) or re.fullmatch(r"[0-9a-f]{64}", value[key]) is None:
            raise QualificationError(f"probe has an invalid {key}")
    correctness = value.get("correctness")
    if not isinstance(correctness, dict) or correctness.get("status") != "ok":
        raise QualificationError("probe correctness precondition is not ok")
    if family == "owner":
        if value.get("output_differs") is not True:
            raise QualificationError("owner probe did not prove a changed output")
        for key in (
            "replay_matches_changed",
            "inverse_matches_source",
            "noop_is_noop",
            "noop_matches_expected",
            "semantic_ok",
        ):
            if correctness.get(key) is not True:
                raise QualificationError(f"owner probe correctness field {key} is false")
        if correctness.get("changed_sha256") != value["output_sha256"]:
            raise QualificationError("owner probe output digest disagrees with correctness digest")
        for key in ("changed_sha256", "replay_sha256", "inverse_sha256", "noop_sha256"):
            if re.fullmatch(r"[0-9a-f]{64}", str(correctness.get(key))) is None:
                raise QualificationError(f"owner probe correctness field {key} is not a digest")
        for key in ("source_bytes", "changed_bytes", "noop_bytes"):
            if not isinstance(correctness.get(key), int) or correctness[key] <= 0:
                raise QualificationError(f"owner probe correctness field {key} is not positive")
        if correctness["changed_bytes"] != value["output_bytes"]:
            raise QualificationError("owner probe output bytes disagree with correctness bytes")
    else:
        if correctness.get("roundtrip") is not True or value.get("output_matches_source") is not True:
            raise QualificationError("constructor probe did not prove exact roundtrip")
    if require_allocations:
        for key in ALLOCATION_METRICS:
            metric = value.get(key)
            if type(metric) is not int or metric < 0:
                raise QualificationError(f"probe has an invalid allocation metric {key}")
        if value["peak"] < value["retained"]:
            raise QualificationError("probe allocation peak is below retained bytes")
        if family == "owner" and (value["allocated"] <= 0 or value["allocations"] <= 0):
            raise QualificationError("owner allocation probe did not instrument the edit")
    return value


def _child_limits(timeout_s: float, *, file_bytes: int | None = 64 * 1024 * 1024):
    if resource is None:
        return None

    def configure() -> None:
        cpu_seconds = max(1, math.ceil(timeout_s))
        resource.setrlimit(resource.RLIMIT_CPU, (cpu_seconds, cpu_seconds + 1))
        if file_bytes is not None:
            resource.setrlimit(resource.RLIMIT_FSIZE, (file_bytes, file_bytes))

    return configure


def _terminate_child(child: subprocess.Popen[bytes]) -> int:
    """Terminate a timed child process group and leave no unbounded wait."""

    try:
        os.killpg(child.pid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    try:
        return child.wait(timeout=2.0)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(child.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        return child.wait(timeout=2.0)


def invoke(
    output: Path,
    family: str,
    label: str,
    binary: Path,
    route: str,
    count: int,
    cpu: int,
    fixtures: Path | None,
    expected_binary: Mapping[str, Any],
    timeout_s: float = FIXED_TIMEOUT_S,
    *,
    allocation: bool = False,
) -> dict[str, Any]:
    if not math.isfinite(timeout_s) or timeout_s <= 0:
        raise QualificationError("timed child timeout must be finite and positive")
    verify_artifact(expected_binary)
    host_snapshot(output, f"{label}-before")
    stdout = output / f"{label}.stdout"
    stderr = output / f"{label}.stderr"
    command = ["taskset", "-c", str(cpu), str(binary), route, str(count)]
    if family == "owner":
        if fixtures is None:
            raise QualificationError("owner timing requires --fixtures")
        command.append(str(fixtures))
    started = time.time_ns()
    overlap: str | None = None
    timed_out = False
    deadline = time.monotonic() + timeout_s
    with stdout.open("wb") as out, stderr.open("wb") as err:
        child = subprocess.Popen(
            command,
            stdout=out,
            stderr=err,
            start_new_session=True,
            preexec_fn=_child_limits(timeout_s),
        )
        checks = 0
        while True:
            try:
                remaining = deadline - time.monotonic()
                if remaining <= 0:
                    timed_out = True
                    overlap = f"timed child exceeded {timeout_s:g}s"
                    status = _terminate_child(child)
                    break
                status = child.wait(timeout=min(0.1, remaining))
                break
            except subprocess.TimeoutExpired:
                if time.monotonic() >= deadline:
                    timed_out = True
                    overlap = f"timed child exceeded {timeout_s:g}s"
                    status = _terminate_child(child)
                    break
                try:
                    host_snapshot(output, f"{label}-during-{checks:04d}", child.pid)
                except QualificationError as error:
                    overlap = str(error)
                    status = _terminate_child(child)
                    break
                checks += 1
        if checks == 0:
            try:
                host_snapshot(output, f"{label}-during-0000", child.pid)
            except QualificationError as error:
                overlap = str(error)
    receipt = {
        "command": command,
        "started_ns": started,
        "finished_ns": time.time_ns(),
        "exit_code": status,
        "overlap": overlap,
        "timed_out": timed_out,
        "timeout_s": timeout_s,
        "binary": artifact(binary),
        "stdout": artifact(stdout),
        "stderr": artifact(stderr),
    }
    receipt_path = output / f"{label}.receipt.json"
    save_json(receipt_path, receipt)
    try:
        host_snapshot(output, f"{label}-after")
    except QualificationError as error:
        overlap = overlap or str(error)
        receipt["overlap"] = overlap
        save_json(receipt_path, receipt)
    if status or overlap:
        raise QualificationError(f"invalid timing attempt {label}; see {output}")
    try:
        value = load_json(stdout)
    except QualificationError as error:
        raise QualificationError(f"probe stdout is not strict JSON for {label}") from error
    value = _strict_probe_value(value, family, route, count, require_allocations=allocation)
    value["_receipt"] = {
        "receipt": artifact(receipt_path),
        "binary": receipt["binary"],
        "stdout": receipt["stdout"],
        "stderr": receipt["stderr"],
        "exit_code": receipt["exit_code"],
        "overlap": receipt["overlap"],
        "timed_out": receipt["timed_out"],
    }
    return value


def _ratios_equal(values: list[dict[str, Any]]) -> None:
    checksums = {value["checksum"] for value in values}
    outputs = {value["output_sha256"] for value in values}
    sources = {value["source_sha256"] for value in values}
    if len(checksums) != 1 or len(outputs) != 1 or len(sources) != 1:
        raise QualificationError("ABBA semantic or output identities disagree")


def _allocation_abba_equal(values: list[dict[str, Any]]) -> None:
    if len(values) != len(ALLOCATION_ORDER):
        raise QualificationError("allocation packet does not contain one ABBA value per leg")
    for value in values:
        if not isinstance(value, Mapping):
            raise QualificationError("allocation packet contains a malformed probe value")
        for key in ALLOCATION_METRICS:
            metric = value.get(key)
            if type(metric) is not int or metric < 0:
                raise QualificationError(f"allocation packet has an invalid metric {key}")
        if value["peak"] < value["retained"]:
            raise QualificationError("allocation packet peak is below retained bytes")
    _ratios_equal(values)
    for key in ALLOCATION_METRICS:
        if values[0][key] != values[3][key] or values[1][key] != values[2][key]:
            raise QualificationError(f"allocation metric {key} changed across its ABBA pair")


def _protocol(args: argparse.Namespace, family: str, routes: tuple[str, ...]) -> dict[str, Any]:
    return {
        "start_count": args.start_count,
        "blocks": args.blocks,
        "warmups": args.warmups,
        "target_ns": args.target_ns,
        "minimum_ns": args.minimum_ns,
        "growth": args.growth,
        "max_iterations": args.max_iterations,
        "timeout_s": args.timeout_s,
        "routes": list(routes),
        "family": family,
    }


def _fixed_protocol(protocol: Mapping[str, Any], family: str) -> bool:
    expected = {
        "start_count": START_COUNTS[family],
        "blocks": FIXED_BLOCKS,
        "warmups": FIXED_WARMUPS,
        "target_ns": TARGET_NS,
        "minimum_ns": MINIMUM_NS,
        "growth": FIXED_GROWTH,
        "max_iterations": MAX_ITERATIONS,
        "timeout_s": FIXED_TIMEOUT_S,
        "routes": list(ROUTES[family]),
        "family": family,
    }
    return dict(protocol) == expected


def _verify_timing_receipt(value: Mapping[str, Any]) -> None:
    receipt = value.get("_receipt")
    if not isinstance(receipt, Mapping):
        raise QualificationError("timing value did not retain its command receipt")
    verify_command_receipt(
        {
            **load_json(Path(receipt["receipt"]["path"])),
            "receipt": receipt["receipt"],
        }
    )
    verify_artifact(receipt["binary"])
    if receipt.get("exit_code") != 0 or receipt.get("overlap") is not None:
        raise QualificationError("retained timing receipt is not successful")
    if receipt.get("timed_out") is not False:
        raise QualificationError("retained timing receipt exceeded its bound")


def _verify_allocation_receipt(value: Mapping[str, Any]) -> None:
    _verify_timing_receipt(value)
    for key in ALLOCATION_METRICS:
        metric = value.get(key)
        if type(metric) is not int or metric < 0:
            raise QualificationError(f"retained allocation value has an invalid metric {key}")
    if value["peak"] < value["retained"]:
        raise QualificationError("retained allocation peak is below retained bytes")


def run_allocations(args: argparse.Namespace) -> None:
    """Collect operation allocation metrics in a guarded, non-timing ABBA packet."""

    family = require_family(args.family)
    output = outside_checkout(args.output)
    build = load_build(output, family)
    validation = load_validation(output, family, build)
    fixtures = outside_checkout(args.fixtures) if args.fixtures is not None else None
    fixture_manifest = None
    if family == "owner" and fixtures is None:
        raise QualificationError("owner allocation measurements require --fixtures")
    if fixtures is not None:
        fixture_manifest = fixture_identity(fixtures, regenerate=True)
        if validation.get("fixtures") != fixture_manifest:
            raise QualificationError(
                "allocation fixtures do not match correctness-validation fixtures"
            )
    try:
        affinity = os.sched_getaffinity(0)
    except AttributeError as error:
        raise QualificationError("cannot prove CPU affinity on this host") from error
    if args.cpu not in affinity:
        raise QualificationError(f"CPU {args.cpu} is outside the runner affinity {sorted(affinity)}")
    if args.iterations != ALLOCATION_ITERATIONS or args.timeout_s != FIXED_TIMEOUT_S:
        raise QualificationError(
            f"allocation protocol requires {ALLOCATION_ITERATIONS} iterations and a {FIXED_TIMEOUT_S:g}s timeout"
        )
    allocation = {
        "A": build["endpoints"]["base"]["binaries"]["allocation"],
        "B": build["endpoints"]["candidate"]["binaries"]["allocation"],
    }
    for row in allocation.values():
        verify_artifact(row)
    attempt = new_external_directory(output / args.attempt)
    manifest: dict[str, Any] = {
        "schema": SCHEMA,
        "family": family,
        "phase": "allocation",
        "status": "running",
        "qualified": False,
        "acceptance_eligible": True,
        "build": artifact(output / "build.json"),
        "validation": artifact(output / "validation.json"),
        "runner": artifact(Path(__file__)),
        "probe": artifact(PROBES[family]),
        "comparison": build["comparison"],
        "resolved_refs": build["resolved_refs"],
        "source_delta": build["source_delta"],
        "lock_delta": build["lock_delta"],
        "sources": {
            label: build["endpoints"][label]["source"] for label in ("base", "candidate")
        },
        "binaries": allocation,
        "fixtures": fixture_manifest,
        "protocol": {
            "iterations": args.iterations,
            "order": ALLOCATION_ORDER,
            "routes": list(ROUTES[family]),
            "timeout_s": args.timeout_s,
        },
        "cpu": args.cpu,
        "cases": [],
    }
    save_json(attempt / "manifest.json", manifest)
    try:
        for route in ROUTES[family]:
            values = []
            for index, side in enumerate(ALLOCATION_ORDER):
                values.append(
                    invoke(
                        attempt,
                        family,
                        f"{route}-alloc-{index}-{side}",
                        Path(allocation[side]["path"]),
                        route,
                        args.iterations,
                        args.cpu,
                        fixtures,
                        allocation[side],
                        args.timeout_s,
                        allocation=True,
                    )
                )
            _allocation_abba_equal(values)
            manifest["cases"].append(
                {
                    "route": route,
                    "iterations": args.iterations,
                    "order": ALLOCATION_ORDER,
                    "legs": [
                        {"index": index, "side": side, "value": value}
                        for index, (side, value) in enumerate(zip(ALLOCATION_ORDER, values))
                    ],
                }
            )
            save_json(attempt / "manifest.json", manifest)
        host_snapshot(attempt, "final")
        final_build = load_build(output, family)
        final_validation = load_validation(output, family, final_build)
        if artifact(Path(__file__)) != manifest["runner"]:
            raise QualificationError("allocation runner changed during collection")
        if artifact(PROBES[family]) != manifest["probe"]:
            raise QualificationError("allocation probe changed during collection")
        if artifact(output / "build.json") != manifest["build"]:
            raise QualificationError("build manifest changed during allocation collection")
        if artifact(output / "validation.json") != manifest["validation"]:
            raise QualificationError("validation manifest changed during allocation collection")
        for label in ("base", "candidate"):
            if final_build["endpoints"][label]["source"] != manifest["sources"][label]:
                raise QualificationError(f"{label} source identity changed during allocation collection")
        for key in ("comparison", "resolved_refs", "source_delta", "lock_delta"):
            if final_build.get(key) != manifest[key]:
                raise QualificationError(f"{key} changed during allocation collection")
        if family == "owner":
            if fixtures is None or fixture_identity(fixtures, regenerate=True) != manifest["fixtures"]:
                raise QualificationError("owner fixture identity changed during allocation collection")
            if final_validation.get("fixtures") != manifest["fixtures"]:
                raise QualificationError("validation fixture identity changed during allocation collection")
        for case in manifest["cases"]:
            _allocation_abba_equal([leg["value"] for leg in case["legs"]])
            for leg in case["legs"]:
                _verify_allocation_receipt(leg["value"])
        for side, row in allocation.items():
            if artifact(Path(row["path"])) != row:
                raise QualificationError(f"protected {side} allocation binary changed during collection")
        manifest["qualified"] = True
        manifest["status"] = "complete"
        save_json(attempt / "manifest.json", manifest)
    except Exception as error:
        manifest["status"] = "aborted"
        manifest["error"] = str(error)
        save_json(attempt / "manifest.json", manifest)
        raise


def run_native(args: argparse.Namespace) -> None:
    family = require_family(args.family)
    output = outside_checkout(args.output)
    build = load_build(output, family)
    validation = load_validation(output, family, build)
    fixtures = outside_checkout(args.fixtures) if args.fixtures is not None else None
    fixture_manifest = None
    if family == "owner" and fixtures is None:
        raise QualificationError("owner timing requires --fixtures")
    if fixtures is not None:
        fixture_manifest = fixture_identity(fixtures, regenerate=True)
        validation_fixtures = validation.get("fixtures")
        if validation_fixtures != fixture_manifest:
            raise QualificationError("timing fixtures do not match correctness-validation fixtures")
    try:
        affinity = os.sched_getaffinity(0)
    except AttributeError as error:
        raise QualificationError("cannot prove CPU affinity on this host") from error
    if args.cpu not in affinity:
        raise QualificationError(f"CPU {args.cpu} is outside the runner affinity {sorted(affinity)}")
    if (
        args.blocks <= 0
        or args.warmups < 0
        or args.target_ns <= 0
        or args.minimum_ns <= 0
        or args.start_count <= 0
        or args.max_iterations <= 0
        or not math.isfinite(args.growth)
        or not math.isfinite(args.timeout_s)
        or args.timeout_s <= 0
    ):
        raise QualificationError("timing parameters must be positive")
    if args.minimum_ns > args.target_ns or args.growth <= 1.0:
        raise QualificationError("minimum leg cannot exceed calibration target")
    native = {
        "A": build["endpoints"]["base"]["binaries"]["native"],
        "B": build["endpoints"]["candidate"]["binaries"]["native"],
    }
    for row in native.values():
        verify_artifact(row)
    if args.routes:
        routes = tuple(args.routes)
        unknown = set(routes) - set(ROUTES[family])
        if unknown or len(set(routes)) != len(routes):
            raise QualificationError(f"invalid or duplicate route selection: {sorted(unknown)}")
    else:
        routes = ROUTES[family]
    protocol = _protocol(args, family, routes)
    diagnostic = bool(getattr(args, "diagnostic", False))
    protocol_exact = _fixed_protocol(protocol, family)
    if not protocol_exact and not diagnostic:
        raise QualificationError(
            "non-fixed blocks/warmups/calibration/minimum protocol or route subset requires --diagnostic"
        )
    attempt = new_external_directory(output / args.attempt)
    manifest: dict[str, Any] = {
        "schema": SCHEMA,
        "family": family,
        "status": "running",
        "qualified": False,
        "acceptance_eligible": protocol_exact and not diagnostic,
        "diagnostic": diagnostic,
        "build": artifact(output / "build.json"),
        "validation": artifact(output / "validation.json"),
        "runner": artifact(Path(__file__)),
        "probe": artifact(PROBES[family]),
        "comparison": build["comparison"],
        "resolved_refs": build["resolved_refs"],
        "source_delta": build["source_delta"],
        "lock_delta": build["lock_delta"],
        "sources": {
            label: build["endpoints"][label]["source"] for label in ("base", "candidate")
        },
        "binaries": native,
        "fixtures": fixture_manifest,
        "protocol": protocol,
        "cpu": args.cpu,
        "target_ns": args.target_ns,
        "minimum_ns": args.minimum_ns,
        "blocks_requested": args.blocks,
        "warmups": args.warmups,
        "routes": list(routes),
        "cases": [],
    }
    save_json(attempt / "manifest.json", manifest)
    try:
        for route in routes:
            count = max(args.start_count, 1)
            calibrations = []
            while True:
                values = []
                for side in "AB":
                    value = invoke(
                        attempt,
                        family,
                        f"{route}-cal-{count}-{side}",
                        Path(native[side]["path"]),
                        route,
                        count,
                        args.cpu,
                        fixtures,
                        native[side],
                        args.timeout_s,
                    )
                    values.append(value)
                _ratios_equal(values)
                calibrations.append({"count": count, "values": values})
                fastest = min(value["ns"] for value in values)
                if fastest >= args.target_ns:
                    break
                next_count = max(count + 1, math.ceil(count * args.target_ns / fastest * args.growth))
                if next_count > args.max_iterations:
                    raise QualificationError(f"calibration exceeds iteration ceiling for {route}")
                count = next_count
            case: dict[str, Any] = {
                "route": route,
                "count": count,
                "calibrations": calibrations,
                "warmups": [],
                "blocks": [],
            }
            manifest["cases"].append(case)
            save_json(attempt / "manifest.json", manifest)
            for warmup in range(args.warmups):
                values = []
                for side in "AB":
                    value = invoke(
                        attempt,
                        family,
                        f"{route}-warm-{warmup}-{side}",
                        Path(native[side]["path"]),
                        route,
                        count,
                        args.cpu,
                        fixtures,
                        native[side],
                        args.timeout_s,
                    )
                    values.append(value)
                _ratios_equal(values)
                case["warmups"].append({"index": warmup, "values": values})
                save_json(attempt / "manifest.json", manifest)
            for block in range(args.blocks):
                legs = []
                for index, side in enumerate("ABBA"):
                    value = invoke(
                        attempt,
                        family,
                        f"{route}-b{block}-{index}-{side}",
                        Path(native[side]["path"]),
                        route,
                        count,
                        args.cpu,
                        fixtures,
                        native[side],
                        args.timeout_s,
                    )
                    if value["ns"] < args.minimum_ns:
                        raise QualificationError(f"sample below minimum for {route} block {block}")
                    legs.append({"side": side, "value": value})
                values = [leg["value"] for leg in legs]
                _ratios_equal(values)
                ratio = (values[1]["ns"] + values[2]["ns"]) / (values[0]["ns"] + values[3]["ns"])
                case["blocks"].append({"legs": legs, "ratio": ratio})
                save_json(attempt / "manifest.json", manifest)
            case["median_ratio"] = sorted(block["ratio"] for block in case["blocks"])[len(case["blocks"]) // 2]
            save_json(attempt / "manifest.json", manifest)
        host_snapshot(attempt, "final")
        # Re-open every retained identity after timing.  A qualified result
        # must still point at the exact runner, probe, source snapshots,
        # validation receipts, and fixture manifest that were admitted before
        # the first leg; checking only protected binaries is insufficient.
        final_build = load_build(output, family)
        final_validation = load_validation(output, family, final_build)
        if artifact(Path(__file__)) != manifest["runner"]:
            raise QualificationError("qualification runner changed during timing")
        if artifact(PROBES[family]) != manifest["probe"]:
            raise QualificationError("qualification probe changed during timing")
        if artifact(output / "build.json") != manifest["build"]:
            raise QualificationError("build manifest changed during timing")
        if artifact(output / "validation.json") != manifest["validation"]:
            raise QualificationError("validation manifest changed during timing")
        for label in ("base", "candidate"):
            if final_build["endpoints"][label]["source"] != manifest["sources"][label]:
                raise QualificationError(f"{label} source identity changed during timing")
        for key in ("comparison", "resolved_refs", "source_delta", "lock_delta"):
            if final_build.get(key) != manifest[key]:
                raise QualificationError(f"{key} changed during timing")
        if family == "owner":
            if fixtures is None or fixture_identity(fixtures, regenerate=True) != manifest["fixtures"]:
                raise QualificationError("owner fixture identity changed during timing")
            if final_validation.get("fixtures") != manifest["fixtures"]:
                raise QualificationError("validation fixture identity changed during timing")
        for case in manifest["cases"]:
            for calibration in case["calibrations"]:
                for value in calibration["values"]:
                    _verify_timing_receipt(value)
            for warmup in case["warmups"]:
                for value in warmup["values"]:
                    _verify_timing_receipt(value)
            for block in case["blocks"]:
                for leg in block["legs"]:
                    _verify_timing_receipt(leg["value"])
        for side, row in native.items():
            if artifact(Path(row["path"])) != row:
                raise QualificationError(f"protected {side} binary changed during timing")
        manifest["qualified"] = bool(manifest["acceptance_eligible"] and protocol_exact)
        manifest["status"] = "complete"
        save_json(attempt / "manifest.json", manifest)
    except Exception as error:
        manifest["status"] = "aborted"
        manifest["error"] = str(error)
        save_json(attempt / "manifest.json", manifest)
        raise


def parser() -> argparse.ArgumentParser:
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("family", choices=tuple(ROUTES))
    common.add_argument("output", type=Path)
    command = argparse.ArgumentParser(description=__doc__)
    sub = command.add_subparsers(dest="phase", required=True)
    build_parser = sub.add_parser("build", parents=[common])
    build_parser.add_argument("--base", required=True)
    build_parser.add_argument("--candidate", required=True)
    build_parser.add_argument("--lock", type=Path, default=ROOT / "Cargo.lock")
    validate_parser = sub.add_parser("validate", parents=[common])
    validate_parser.add_argument("--fixtures", type=Path)
    run_parser = sub.add_parser("run", parents=[common])
    run_parser.add_argument("--fixtures", type=Path)
    run_parser.add_argument("--attempt", default="native-a1")
    run_parser.add_argument("--cpu", type=int, default=29)
    run_parser.add_argument("--blocks", type=int, default=5)
    run_parser.add_argument("--warmups", type=int, default=2)
    run_parser.add_argument("--target-ns", type=int, default=TARGET_NS)
    run_parser.add_argument("--minimum-ns", type=int, default=MINIMUM_NS)
    run_parser.add_argument("--start-count", type=int)
    run_parser.add_argument("--max-iterations", type=int, default=MAX_ITERATIONS)
    run_parser.add_argument("--growth", type=float, default=FIXED_GROWTH)
    run_parser.add_argument("--timeout-s", type=float, default=FIXED_TIMEOUT_S)
    run_parser.add_argument(
        "--diagnostic",
        action="store_true",
        help="run a non-fixed route/protocol subset without producing an acceptance result",
    )
    run_parser.add_argument("--routes", nargs="*")
    allocation_parser = sub.add_parser("allocation", parents=[common])
    allocation_parser.add_argument("--fixtures", type=Path)
    allocation_parser.add_argument("--attempt", default="allocation-a1")
    allocation_parser.add_argument("--cpu", type=int, default=29)
    allocation_parser.add_argument("--iterations", type=int, default=ALLOCATION_ITERATIONS)
    allocation_parser.add_argument("--timeout-s", type=float, default=FIXED_TIMEOUT_S)
    collect_parser = sub.add_parser("collect", parents=[common])
    collect_parser.add_argument("--attempt")
    collect_parser.add_argument("--out", type=Path)
    prepare_parser = sub.add_parser("prepare-fixtures")
    prepare_parser.add_argument("output", type=Path)
    return command


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        if args.phase == "build":
            build(args)
        elif args.phase == "validate":
            validate(args)
        elif args.phase == "collect":
            collect(args)
        elif args.phase == "prepare-fixtures":
            prepare_fixtures(args.output)
        elif args.phase == "allocation":
            run_allocations(args)
        else:
            if args.start_count is None:
                args.start_count = START_COUNTS[args.family]
            run_native(args)
    except QualificationError as error:
        print(f"qualification refused: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
