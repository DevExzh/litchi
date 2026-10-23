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

The owner build and run take ``--fixtures``.  The fixture directory is an
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
import subprocess
import sys
import tarfile
import time
from typing import Any, Iterable, Mapping


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
TARGET_NS = 500_000_000
MINIMUM_NS = 150_000_000
MAX_ITERATIONS = 1_000_000_000
SCHEMA = "litchi.iwa.native-qualification.v1"
OWNER_FIXTURES = (
    "basic.numbers",
    "media-replacement-native.key",
    "tone.wav",
    "footnote-fixture.pages",
    "basic.pages",
)


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


def fixture_identity(path: Path) -> dict[str, Any]:
    """Verify an external owner fixture set and its immutable manifest."""

    directory = outside_checkout(path)
    if not directory.is_dir():
        raise QualificationError(f"owner fixtures are not a directory: {directory}")
    manifest_path = directory / "manifest.json"
    manifest = load_json(manifest_path)
    if manifest.get("schema") != "litchi.iwa.owner-fixtures.v1":
        raise QualificationError("owner fixture manifest has the wrong schema")
    declared = manifest.get("files")
    if not isinstance(declared, dict) or set(declared) != set(OWNER_FIXTURES):
        raise QualificationError("owner fixture manifest does not declare exactly the four workloads")
    identities = {}
    for name in OWNER_FIXTURES:
        row = declared[name]
        if not isinstance(row, Mapping):
            raise QualificationError(f"owner fixture identity is malformed: {name}")
        expected_path = directory / name
        if Path(row.get("path", "")).resolve() != expected_path.resolve():
            raise QualificationError(f"owner fixture path is not rooted in the fixture directory: {name}")
        verify_artifact(row)
        identities[name] = row
    return {"manifest": artifact(manifest_path), "files": identities}


def new_external_directory(path: Path) -> Path:
    path = outside_checkout(path)
    try:
        path.mkdir(parents=True, exist_ok=False)
    except OSError as error:
        raise QualificationError(f"cannot create fresh external directory {path}: {error}") from error
    return path


def run_command(
    directory: Path,
    label: str,
    argv: list[str],
    cwd: Path,
    env: Mapping[str, str] | None = None,
) -> dict[str, Any]:
    """Run one non-shell command and retain a receipt before checking status."""

    stdout = directory / f"{label}.stdout"
    stderr = directory / f"{label}.stderr"
    started = time.time_ns()
    with stdout.open("wb") as out, stderr.open("wb") as err:
        completed = subprocess.run(argv, cwd=cwd, env=env, stdout=out, stderr=err, check=False)
    receipt = {
        "command": argv,
        "cwd": str(cwd.resolve()),
        "started_ns": started,
        "finished_ns": time.time_ns(),
        "exit_code": completed.returncode,
        "stdout": artifact(stdout),
        "stderr": artifact(stderr),
    }
    save_json(directory / f"{label}.receipt.json", receipt)
    if completed.returncode:
        raise QualificationError(f"{label} failed; see {stderr}")
    return receipt


def git_revision(ref: str) -> str:
    try:
        value = subprocess.check_output(
            ["git", "rev-parse", "--verify", f"{ref}^{{commit}}"],
            cwd=ROOT,
            text=True,
            stderr=subprocess.STDOUT,
        ).strip()
    except subprocess.CalledProcessError as error:
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


def snapshot_source(output: Path, label: str, revision: str, lock: Path) -> tuple[Path, dict[str, Any]]:
    commit = git_revision(revision)
    source = output / f"source-{label}"
    archive = output / f"source-{label}.tar"
    source.mkdir()
    with archive.open("wb") as stream:
        subprocess.run(
            ["git", "archive", "--format=tar", commit, "Cargo.toml", "rust-toolchain.toml", "crates"],
            cwd=ROOT,
            stdout=stream,
            stderr=subprocess.PIPE,
            check=True,
        )
    with tarfile.open(archive) as stream:
        stream.extractall(source, filter="data")
    shutil.copyfile(lock, source / "Cargo.lock")
    manifest = output / f"source-{label}.json"
    save_json(manifest, {"commit": commit, "archive": artifact(archive), "files": _source_files(source)})
    return source, {"commit": commit, "manifest": artifact(manifest)}


def registry(lock: Path) -> list[dict[str, Any]]:
    try:
        import tomllib

        data = tomllib.loads(lock.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        raise QualificationError(f"cannot inspect Cargo.lock {lock}: {error}") from error
    return sorted(
        [
            {
                "name": package["name"],
                "version": package["version"],
                "source": package.get("source"),
                "checksum": package.get("checksum"),
            }
            for package in data.get("package", [])
            if str(package.get("source", "")).startswith("registry+")
        ],
        key=lambda row: (row["name"], row["version"], row["source"] or ""),
    )


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
        "status": "building",
        "routes": list(ROUTES[family]),
        "builder": artifact(Path(__file__)),
        "probe": artifact(probe),
        "input_lock": artifact(lock),
        "endpoints": {},
    }
    save_json(output / "build.json", manifest)
    common_lock: Path | None = None
    for label, revision in (("base", args.base), ("candidate", args.candidate)):
        source, provenance = snapshot_source(output, label, revision, lock)
        project = output / f"project-{label}"
        (project / "src").mkdir(parents=True)
        shutil.copyfile(probe, project / "src/main.rs")
        shutil.copyfile(source / "rust-toolchain.toml", project / "rust-toolchain.toml")
        (project / "Cargo.toml").write_text(_project_manifest(family, source), encoding="utf-8")
        receipts = output / f"build-{label}"
        receipts.mkdir()
        env = dict(os.environ, CARGO_INCREMENTAL="0", CARGO_BUILD_JOBS="1")
        rows = []
        if common_lock is None:
            shutil.copyfile(lock, project / "Cargo.lock")
            rows.append(run_command(receipts, "generate-lock", ["cargo", "generate-lockfile", "--offline"], project, env))
            common_lock = project / "Cargo.lock"
        else:
            shutil.copyfile(common_lock, project / "Cargo.lock")
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
        save_json(output / "build.json", manifest)
        package_name = "iwa-constructor-probe" if family == "constructor" else "iwa-owner-probe"
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
    base = manifest["endpoints"]["base"]
    candidate = manifest["endpoints"]["candidate"]
    for key in ("rustc", "cargo", "registry"):
        if base[key] != candidate[key]:
            raise QualificationError(f"matched endpoints disagree on {key}")
    if base["lock"]["sha256"] != candidate["lock"]["sha256"]:
        raise QualificationError("matched endpoints use different Cargo.lock files")
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
    for label in ("base", "candidate"):
        endpoint = data["endpoints"].get(label)
        if not isinstance(endpoint, Mapping):
            raise QualificationError(f"missing {label} endpoint")
        verify_artifact(endpoint["source"]["manifest"])
        source_manifest = load_json(Path(endpoint["source"]["manifest"]["path"]))
        verify_artifact(source_manifest["archive"])
        source_dir = output / f"source-{label}"
        for row in source_manifest["files"]:
            actual = artifact(source_dir / row["path"])
            if (actual["bytes"], actual["sha256"]) != (row["bytes"], row["sha256"]):
                raise QualificationError(f"source snapshot changed: {label}/{row['path']}")
        for key in ("project_manifest", "project_source", "lock"):
            verify_artifact(endpoint[key])
        for key in ("native", "allocation"):
            verify_artifact(endpoint["binaries"][key])
    base = data["endpoints"]["base"]
    candidate = data["endpoints"]["candidate"]
    for key in ("rustc", "cargo", "registry"):
        if base[key] != candidate[key]:
            raise QualificationError(f"matched endpoints disagree on {key}")
    if base["lock"]["sha256"] != candidate["lock"]["sha256"]:
        raise QualificationError("matched endpoints use different Cargo.lock files")
    base_files = {row["path"]: row["sha256"] for row in load_json(Path(base["source"]["manifest"]["path"]))["files"]}
    candidate_files = {row["path"]: row["sha256"] for row in load_json(Path(candidate["source"]["manifest"]["path"]))["files"]}
    excluded = {
        "constructor": ("crates/litchi-iwa-core/",),
        "owner": (
            "crates/litchi-numbers/",
            "crates/litchi-keynote/",
            "crates/litchi-pages/",
        ),
    }[family]
    for path, digest in base_files.items():
        if path not in candidate_files:
            raise QualificationError(f"candidate source is missing {path}")
        if not path.startswith(excluded) and candidate_files[path] != digest:
            raise QualificationError(f"unmatched shared source changed: {path}")
    return data


def _validation_command(family: str, binary: Path, fixtures: Path | None, destination: Path) -> list[str]:
    if family == "owner":
        if fixtures is None:
            raise QualificationError("owner validation requires --fixtures")
        return [str(binary), "validate_outputs", str(fixtures), str(destination)]
    return [str(binary), "validate_outputs", str(destination)]


def validate(args: argparse.Namespace) -> None:
    family = require_family(args.family)
    output = outside_checkout(args.output)
    build = load_build(output, family)
    fixtures = outside_checkout(args.fixtures) if args.fixtures is not None else None
    fixture_manifest = None
    if family == "owner" and fixtures is None:
        raise QualificationError("owner validation requires --fixtures")
    if fixtures is not None:
        fixture_manifest = fixture_identity(fixtures)
    rows: dict[str, Any] = {}
    validation_root = output / "validation"
    validation_root.mkdir(exist_ok=False)
    for label in ("base", "candidate"):
        rows[label] = {}
        for variant in ("native", "allocation"):
            endpoint = build["endpoints"][label]
            destination = validation_root / f"{label}-{variant}"
            destination.mkdir()
            run_command(
                destination,
                "probe",
                _validation_command(family, Path(endpoint["binaries"][variant]["path"]), fixtures, destination),
                output,
            )
            report = destination / "validation.json"
            value = load_json(report)
            expected_schema = f"litchi.iwa.{family}.validation.v1"
            if value.get("schema") != expected_schema or value.get("status") != "ok":
                raise QualificationError(f"invalid {label}/{variant} validation report")
            rows[label][variant] = {"report": artifact(report), "value": value}
    for variant in ("native", "allocation"):
        if rows["base"][variant]["value"] != rows["candidate"][variant]["value"]:
            raise QualificationError(f"baseline/candidate {variant} correctness differs")
    result = {
        "schema": SCHEMA,
        "family": family,
        "status": "complete",
        "build": artifact(output / "build.json"),
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
    for label in ("base", "candidate"):
        for variant in ("native", "allocation"):
            row = data["endpoints"][label][variant]
            verify_artifact(row["report"])
            expected = f"litchi.iwa.{family}.validation.v1"
            if row["value"].get("schema") != expected or row["value"].get("status") != "ok":
                raise QualificationError("malformed correctness validation value")
    if data["endpoints"]["base"]["native"]["value"] != data["endpoints"]["candidate"]["native"]["value"]:
        raise QualificationError("baseline/candidate native correctness differs")
    if data["endpoints"]["base"]["allocation"]["value"] != data["endpoints"]["candidate"]["allocation"]["value"]:
        raise QualificationError("baseline/candidate allocation correctness differs")
    return data


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


def _strict_probe_value(value: Any, family: str, route: str, count: int) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise QualificationError("probe did not emit one JSON object")
    if value.get("route") != route or value.get("iterations") != count:
        raise QualificationError(f"probe identity mismatch for {route}/{count}")
    if not isinstance(value.get("ns"), int) or value["ns"] <= 0:
        raise QualificationError("probe did not report a positive integer duration")
    if not isinstance(value.get("checksum"), int):
        raise QualificationError("probe did not report a semantic checksum")
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
            "noop_matches_source",
        ):
            if correctness.get(key) is not True:
                raise QualificationError(f"owner probe correctness field {key} is false")
    else:
        if correctness.get("roundtrip") is not True or value.get("output_matches_source") is not True:
            raise QualificationError("constructor probe did not prove exact roundtrip")
    return value


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
) -> dict[str, Any]:
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
    with stdout.open("wb") as out, stderr.open("wb") as err:
        child = subprocess.Popen(command, stdout=out, stderr=err)
        checks = 0
        while True:
            try:
                status = child.wait(timeout=0.1)
                break
            except subprocess.TimeoutExpired:
                try:
                    host_snapshot(output, f"{label}-during-{checks:04d}", child.pid)
                except QualificationError as error:
                    overlap = str(error)
                    child.terminate()
                    status = child.wait()
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
        "binary": artifact(binary),
        "stdout": artifact(stdout),
        "stderr": artifact(stderr),
    }
    save_json(output / f"{label}.receipt.json", receipt)
    try:
        host_snapshot(output, f"{label}-after")
    except QualificationError as error:
        overlap = overlap or str(error)
        receipt["overlap"] = overlap
        save_json(output / f"{label}.receipt.json", receipt)
    if status or overlap:
        raise QualificationError(f"invalid timing attempt {label}; see {output}")
    try:
        value = load_json(stdout)
    except QualificationError as error:
        raise QualificationError(f"probe stdout is not strict JSON for {label}") from error
    return _strict_probe_value(value, family, route, count)


def _ratios_equal(values: list[dict[str, Any]]) -> None:
    checksums = {value["checksum"] for value in values}
    outputs = {value["output_sha256"] for value in values}
    sources = {value["source_sha256"] for value in values}
    if len(checksums) != 1 or len(outputs) != 1 or len(sources) != 1:
        raise QualificationError("ABBA semantic or output identities disagree")


def run_native(args: argparse.Namespace) -> None:
    family = require_family(args.family)
    output = outside_checkout(args.output)
    build = load_build(output, family)
    load_validation(output, family, build)
    fixtures = outside_checkout(args.fixtures) if args.fixtures is not None else None
    fixture_manifest = None
    if family == "owner" and fixtures is None:
        raise QualificationError("owner timing requires --fixtures")
    if fixtures is not None:
        fixture_manifest = fixture_identity(fixtures)
        validation_fixtures = load_validation(output, family, build).get("fixtures")
        if validation_fixtures != fixture_manifest:
            raise QualificationError("timing fixtures do not match correctness-validation fixtures")
    try:
        affinity = os.sched_getaffinity(0)
    except AttributeError as error:
        raise QualificationError("cannot prove CPU affinity on this host") from error
    if args.cpu not in affinity:
        raise QualificationError(f"CPU {args.cpu} is outside the runner affinity {sorted(affinity)}")
    if args.blocks <= 0 or args.warmups < 0 or args.target_ns <= 0 or args.minimum_ns <= 0:
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
    attempt = new_external_directory(output / args.attempt)
    manifest: dict[str, Any] = {
        "schema": SCHEMA,
        "family": family,
        "status": "running",
        "qualified": False,
        "build": artifact(output / "build.json"),
        "validation": artifact(output / "validation.json"),
        "runner": artifact(Path(__file__)),
        "probe": artifact(PROBES[family]),
        "binaries": native,
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
                    )
                    values.append(value)
                calibrations.append({"count": count, "values": values})
                fastest = min(value["ns"] for value in values)
                if fastest >= args.target_ns:
                    break
                next_count = max(count + 1, math.ceil(count * args.target_ns / fastest * args.growth))
                if next_count > args.max_iterations:
                    raise QualificationError(f"calibration exceeds iteration ceiling for {route}")
                count = next_count
            for warmup in range(args.warmups):
                for side in "AB":
                    invoke(
                        attempt,
                        family,
                        f"{route}-warm-{warmup}-{side}",
                        Path(native[side]["path"]),
                        route,
                        count,
                        args.cpu,
                        fixtures,
                        native[side],
                    )
            case: dict[str, Any] = {"route": route, "count": count, "calibrations": calibrations, "blocks": []}
            manifest["cases"].append(case)
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
        for side, row in native.items():
            if artifact(Path(row["path"])) != row:
                raise QualificationError(f"protected {side} binary changed during timing")
        manifest["qualified"] = set(routes) == set(ROUTES[family])
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
    build_parser.add_argument("--fixtures", type=Path)
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
    run_parser.add_argument("--growth", type=float, default=1.2)
    run_parser.add_argument("--routes", nargs="*")
    return command


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    try:
        if args.phase == "build":
            build(args)
        elif args.phase == "validate":
            validate(args)
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
