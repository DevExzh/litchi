#!/usr/bin/env python3
"""Capture the final all-features rustdoc inventory outside the checkout.

The command is intentionally sequential.  It uses one Cargo target cache,
records the exact source/lock/toolchain/command identity for every JSON file,
and only reuses a JSON artifact when that identity and its size/SHA-256 match
the current frozen source.  A target cache may span source commits: Cargo's
fingerprints decide what needs rebuilding.  The output directory is expected
to be outside the repository.

The normal invocation is made after sourcing a per-worker ``env.sh`` that
puts the memory-capped Cargo shim on PATH.  This module never starts Cargo at
import time; its unit tests use a fake command runner.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path
import shutil
import subprocess
import sys
from typing import Callable, Iterable, Mapping, Sequence
import tomllib


MIN_TARGET_BYTES = 19_200_000_000
MAX_TARGET_BYTES = 20_000_000_000
MEMORY_CAP_BYTES = 20 * 1024**3
DEFAULT_PACKAGE_COUNT = 68
REQUIRED_RUSTC_PREFIX = "rustc 1.95.0"
CAPTURE_ENV = {
    "CARGO_BUILD_JOBS": "1",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_DEV_DEBUG": "0",
    "CARGO_PROFILE_TEST_DEBUG": "0",
    "CARGO_PROFILE_RELEASE_DEBUG": "0",
    "CARGO_NET_OFFLINE": "true",
    "RUSTC_BOOTSTRAP": "1",
    "LITCHI_MEM_MAX": "20G",
    "LC_ALL": "C",
}
IDENTITY_ENV_KEYS = tuple(CAPTURE_ENV) + (
    "LITCHI_CAPPED", "RUSTFLAGS", "RUSTDOCFLAGS", "RUSTUP_TOOLCHAIN"
)
RECEIPT_SCHEMA = 1


class InventoryError(RuntimeError):
    """The capture cannot safely produce a complete inventory."""


@dataclass(frozen=True)
class Package:
    name: str
    lib_name: str
    manifest: str
    manifest_sha256: str


@dataclass(frozen=True)
class SourceContext:
    root: str
    source_commit: str
    lock_sha256: str
    rustc: str
    cargo: str
    rust_toolchain_sha256: str
    cargo_path: str
    rustc_path: str


@dataclass(frozen=True)
class CaptureResult:
    package: str
    json_path: str
    receipt_path: str
    reused: bool
    json_bytes: int
    json_sha256: str


CommandRunner = Callable[[Sequence[str], Path, Mapping[str, str], Path, Path, Path], int]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def run_text(command: Sequence[str], *, cwd: Path, env: Mapping[str, str] | None = None) -> str:
    result = subprocess.run(command, cwd=cwd, env=dict(env) if env else None,
                            check=True, capture_output=True, text=True)
    return result.stdout.strip()


def enumerate_packages(root: Path) -> list[Package]:
    """Return every direct ``crates/*`` library package in manifest order."""
    manifests = sorted((root / "crates").glob("*/Cargo.toml"))
    packages: list[Package] = []
    seen: set[str] = set()
    for manifest in manifests:
        try:
            document = tomllib.loads(manifest.read_text(encoding="utf-8"))
        except (OSError, tomllib.TOMLDecodeError) as error:
            raise InventoryError(f"cannot read {manifest}: {error}") from error
        package = document.get("package")
        if not isinstance(package, dict) or not isinstance(package.get("name"), str):
            continue
        name = package["name"]
        lib = document.get("lib")
        if isinstance(lib, dict):
            lib_name = lib.get("name", name.replace("-", "_"))
            lib_path = manifest.parent / lib.get("path", "src/lib.rs")
        else:
            lib_name = name.replace("-", "_")
            lib_path = manifest.parent / "src/lib.rs"
        if not isinstance(lib_name, str) or not lib_path.is_file():
            continue
        if name in seen:
            raise InventoryError(f"duplicate package name: {name}")
        seen.add(name)
        packages.append(Package(
            name=name,
            lib_name=lib_name,
            manifest=str(manifest.relative_to(root)),
            manifest_sha256=sha256_file(manifest),
        ))
    return packages


def source_context(root: Path, *, cargo: str = "cargo", rustc: str = "rustc") -> SourceContext:
    """Read the source identity without invoking a build."""
    git_env = {**os.environ, "GIT_OPTIONAL_LOCKS": "0"}
    commit = run_text(["git", "rev-parse", "--verify", "HEAD^{commit}"], cwd=root, env=git_env)
    if len(commit) != 40 or any(char not in "0123456789abcdef" for char in commit):
        raise InventoryError(f"git did not return a full commit SHA: {commit!r}")
    dirty = run_text(["git", "status", "--porcelain"], cwd=root, env=git_env)
    if dirty:
        raise InventoryError("source checkout is dirty; freeze a clean candidate first")
    lockfile = root / "Cargo.lock"
    if not lockfile.is_file():
        raise InventoryError(f"missing lockfile: {lockfile}")
    toolchain = next((root / name for name in ("rust-toolchain.toml", "rust-toolchain")
                      if (root / name).is_file()), None)
    if toolchain is None:
        raise InventoryError("missing rust-toolchain.toml or rust-toolchain")
    cargo_path = shutil.which(cargo) or cargo
    rustc_path = shutil.which(rustc) or rustc
    rustc_version = run_text([rustc, "-Vv"], cwd=root)
    if not rustc_version.startswith(REQUIRED_RUSTC_PREFIX):
        first_line = rustc_version.splitlines()[0] if rustc_version else "<empty>"
        raise InventoryError(f"expected {REQUIRED_RUSTC_PREFIX}, got {first_line!r}")
    # Use the exact capture environment for tool identity probing too.  In
    # particular, an inherited marker must not be able to make a capped Cargo
    # shim take its uncapped path before captures begin.
    identity_env = capture_env()
    return SourceContext(
        root=str(root),
        source_commit=commit,
        lock_sha256=sha256_file(lockfile),
        rustc=rustc_version,
        cargo=run_text([cargo, "-V"], cwd=root, env=identity_env),
        rust_toolchain_sha256=sha256_file(toolchain),
        cargo_path=cargo_path,
        rustc_path=rustc_path,
    )


def source_stable(context: SourceContext) -> None:
    """Reject a source/lock change before accepting another artifact."""
    root = Path(context.root)
    git_env = {**os.environ, "GIT_OPTIONAL_LOCKS": "0"}
    commit = run_text(["git", "rev-parse", "--verify", "HEAD^{commit}"], cwd=root, env=git_env)
    dirty = run_text(["git", "status", "--porcelain"], cwd=root, env=git_env)
    lock_hash = sha256_file(root / "Cargo.lock")
    toolchain = next((root / name for name in ("rust-toolchain.toml", "rust-toolchain")
                      if (root / name).is_file()), None)
    toolchain_hash = sha256_file(toolchain) if toolchain is not None else ""
    if commit != context.source_commit:
        raise InventoryError(f"source commit changed during capture: {context.source_commit} -> {commit}")
    if dirty:
        raise InventoryError("source checkout became dirty during capture")
    if lock_hash != context.lock_sha256:
        raise InventoryError("Cargo.lock changed during capture")
    if toolchain_hash != context.rust_toolchain_sha256:
        raise InventoryError("rust-toolchain changed during capture")


def target_bytes(target: Path) -> int:
    """Use apparent bytes, matching the coordinator's ``du -sb`` check."""
    if not target.is_dir():
        return 0
    output = subprocess.check_output(["du", "-sb", str(target)], text=True)
    return int(output.split()[0])


def check_target(target: Path, limit: int, *, label: str,
                 size_fn: Callable[[Path], int] = target_bytes) -> int:
    size = size_fn(target)
    if size >= limit:
        raise InventoryError(f"{label}: target is {size} bytes, limit is {limit}")
    return size


def rustdoc_command(package: Package, target: Path, cargo: str = "cargo") -> list[str]:
    return [
        cargo, "rustdoc", "--package", package.name, "--lib", "--all-features",
        "--offline", "--locked", "--target-dir", str(target), "--",
        "-Zunstable-options", "--output-format", "json", "--document-hidden-items",
    ]


def capture_env(base: Mapping[str, str] | None = None) -> dict[str, str]:
    environment = dict(base or os.environ)
    environment.update(CAPTURE_ENV)
    # The coordinator shim deliberately bypasses its cgroup when this inherited
    # marker is nonempty.  Never allow a caller's marker to leak into a run.
    environment.pop("LITCHI_CAPPED", None)
    environment["LITCHI_CAPPED"] = ""
    return environment


def verify_cargo_shim(cargo_path: str, environment: Mapping[str, str]) -> dict[str, object]:
    """Prove that Cargo resolves to the capped shim and cannot be bypassed."""
    if environment.get("LITCHI_CAPPED"):
        raise InventoryError("LITCHI_CAPPED must be unset before starting Cargo")
    if environment.get("LITCHI_MEM_MAX") != "20G":
        raise InventoryError("LITCHI_MEM_MAX must be exactly 20G")
    systemd_run = shutil.which("systemd-run")
    if not systemd_run:
        raise InventoryError("systemd-run is unavailable; the Cargo cap cannot be verified")
    shim = Path(cargo_path)
    if not shim.is_file() or not os.access(shim, os.X_OK):
        raise InventoryError(f"resolved Cargo is not an executable shim: {shim}")
    try:
        source = shim.read_text(encoding="utf-8")
    except (OSError, UnicodeDecodeError) as error:
        raise InventoryError(f"cannot inspect Cargo shim {shim}: {error}") from error
    if "systemd-run" not in source or "LITCHI_CAPPED" not in source:
        raise InventoryError(f"Cargo path is not the coordinator cap shim: {shim}")
    if not re.search(r"MemoryMax=\$\{LITCHI_MEM_MAX:-[^}]+\}", source):
        raise InventoryError(f"Cargo shim has no configurable MemoryMax: {shim}")
    return {
        "path": str(shim),
        "sha256": sha256_file(shim),
        "systemd_run_path": systemd_run,
        "memory_max": environment["LITCHI_MEM_MAX"],
        "memory_max_bytes": MEMORY_CAP_BYTES,
        "litchi_capped_unset": True,
    }


def default_runner(command: Sequence[str], cwd: Path, env: Mapping[str, str],
                   stdout: Path, stderr: Path, resource_log: Path) -> int:
    time_command = "/usr/bin/time"
    if not Path(time_command).is_file():
        raise InventoryError(f"resource measurement tool is missing: {time_command}")
    timed_command = [time_command, "-v", "-o", str(resource_log), *command]
    with stdout.open("wb") as out, stderr.open("wb") as err:
        return subprocess.run(timed_command, cwd=cwd, env=dict(env), stdout=out, stderr=err,
                              check=False).returncode


def peak_rss(resource_log: Path) -> dict[str, object]:
    """Parse GNU time's per-command maximum resident set size."""
    try:
        text = resource_log.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise InventoryError(f"missing resource receipt {resource_log}: {error}") from error
    match = re.search(r"Maximum resident set size \(kbytes\):\s*(\d+)", text)
    if match is None:
        raise InventoryError(f"resource receipt has no peak RSS: {resource_log}")
    kib = int(match.group(1))
    return {
        "tool": "/usr/bin/time -v",
        "log": str(resource_log),
        "peak_rss_kib": kib,
        "peak_rss_bytes": kib * 1024,
        "memory_cap": "20G",
        "memory_cap_bytes": MEMORY_CAP_BYTES,
    }


def validate_json(path: Path, lib_name: str) -> tuple[int, str]:
    try:
        document = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise InventoryError(f"invalid rustdoc JSON for {lib_name}: {path}: {error}") from error
    if (not isinstance(document, dict) or not isinstance(document.get("index"), dict)
            or not isinstance(document.get("paths"), dict)):
        raise InventoryError(f"rustdoc JSON lacks index/paths: {path}")
    root = str(document.get("root"))
    root_item = document["index"].get(root)
    if (not isinstance(root_item, dict) or root_item.get("name") != lib_name
            or not isinstance(root_item.get("inner"), dict)
            or "module" not in root_item["inner"]):
        raise InventoryError(f"rustdoc JSON root is not library {lib_name}: {path}")
    if document.get("includes_private") is not False:
        raise InventoryError(f"rustdoc JSON must exclude private items: {path}")
    return path.stat().st_size, sha256_file(path)


def _identity(context: SourceContext, package: Package, command: Sequence[str],
              environment: Mapping[str, str],
              shim_identity: Mapping[str, object] | None) -> dict[str, object]:
    return {
        "schema": RECEIPT_SCHEMA,
        "package": package.name,
        "lib_name": package.lib_name,
        "manifest": package.manifest,
        "manifest_sha256": package.manifest_sha256,
        "source_commit": context.source_commit,
        "cargo_lock_sha256": context.lock_sha256,
        "rustc": context.rustc,
        "cargo": context.cargo,
        "rust_toolchain_sha256": context.rust_toolchain_sha256,
        "cargo_path": context.cargo_path,
        "rustc_path": context.rustc_path,
        "command": list(command),
        "capture_env": {key: environment.get(key, "") for key in IDENTITY_ENV_KEYS},
        "cargo_shim": dict(shim_identity or {}),
    }


def _read_json(path: Path) -> object | None:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None


def _reusable(json_path: Path, receipt_path: Path, identity: dict[str, object]) -> CaptureResult | None:
    receipt = _read_json(receipt_path)
    if not isinstance(receipt, dict) or receipt.get("status") != "ok":
        return None
    if receipt.get("identity") != identity or not json_path.is_file():
        return None
    size, digest = validate_json(json_path, str(identity["lib_name"]))
    artifact = receipt.get("artifact")
    if (not isinstance(artifact, dict) or artifact.get("bytes") != size
            or artifact.get("sha256") != digest):
        return None
    resource = receipt.get("resource")
    if (not isinstance(resource, dict) or not isinstance(resource.get("peak_rss_bytes"), int)
            or resource.get("peak_rss_bytes", 0) <= 0):
        return None
    return CaptureResult(identity["package"], str(json_path), str(receipt_path), True, size, digest)


def _write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def capture_package(package: Package, context: SourceContext, output: Path, target: Path,
                    *, cargo: str = "cargo", reuse: bool = True,
                    runner: CommandRunner = default_runner,
                    base_env: Mapping[str, str] | None = None,
                    shim_identity: Mapping[str, object] | None = None) -> CaptureResult:
    output.mkdir(parents=True, exist_ok=True)
    inventory_dir = output / "inventory"
    receipt_dir = output / "receipts"
    log_dir = output / "logs"
    inventory_dir.mkdir(parents=True, exist_ok=True)
    receipt_dir.mkdir(parents=True, exist_ok=True)
    log_dir.mkdir(parents=True, exist_ok=True)
    slug = package.name.replace("-", "_")
    json_path = inventory_dir / f"{slug}.json"
    receipt_path = receipt_dir / f"{slug}.json"
    environment = capture_env(base_env)
    command = rustdoc_command(package, target, cargo)
    identity = _identity(context, package, command, environment, shim_identity)
    if reuse:
        cached = _reusable(json_path, receipt_path, identity)
        if cached is not None:
            return cached

    target.mkdir(parents=True, exist_ok=True)
    pre_bytes = check_target(target, MIN_TARGET_BYTES, label="target preflight")
    stdout = log_dir / f"{slug}.stdout.log"
    stderr = log_dir / f"{slug}.stderr.log"
    resource_log = log_dir / f"{slug}.time.log"
    started = utc_now()
    returncode = runner(command, Path(context.root), environment, stdout, stderr, resource_log)
    post_bytes = check_target(target, MAX_TARGET_BYTES, label="target hard cap")
    resource = peak_rss(resource_log)
    if returncode != 0:
        _write_json(receipt_path, {
            "schema": RECEIPT_SCHEMA, "status": "failed", "identity": identity,
            "started": started, "finished": utc_now(), "returncode": returncode,
            "target_pre_bytes": pre_bytes, "target_post_bytes": post_bytes,
            "resource": resource,
            "stdout": str(stdout), "stderr": str(stderr),
        })
        raise InventoryError(f"rustdoc failed for {package.name} with exit {returncode}")
    if int(resource["peak_rss_bytes"]) >= MEMORY_CAP_BYTES:
        raise InventoryError(f"{package.name} exceeded the 20G RSS cap")
    generated = target / "doc" / f"{package.lib_name}.json"
    size, digest = validate_json(generated, package.lib_name)
    shutil.copyfile(generated, json_path)
    copied_size, copied_digest = validate_json(json_path, package.lib_name)
    if (size, digest) != (copied_size, copied_digest):
        raise InventoryError(f"rustdoc copy changed {package.name} JSON")
    receipt = {
        "schema": RECEIPT_SCHEMA,
        "status": "ok",
        "identity": identity,
        "started": started,
        "finished": utc_now(),
        "target_pre_bytes": pre_bytes,
        "target_post_bytes": post_bytes,
        "resource": resource,
        "artifact": {"path": str(json_path), "bytes": size, "sha256": digest},
        "stdout": str(stdout),
        "stderr": str(stderr),
    }
    _write_json(receipt_path, receipt)
    return CaptureResult(package.name, str(json_path), str(receipt_path), False, size, digest)


def write_package_manifest(output: Path, context: SourceContext, packages: Iterable[Package]) -> Path:
    path = output / "packages.json"
    _write_json(path, {
        "schema": RECEIPT_SCHEMA,
        "source_commit": context.source_commit,
        "cargo_lock_sha256": context.lock_sha256,
        "rustc": context.rustc,
        "cargo": context.cargo,
        "packages": [asdict(package) for package in packages],
    })
    return path


def package_set_sha256(packages: Sequence[Package]) -> str:
    payload = json.dumps([asdict(package) for package in packages],
                         sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def write_run_state(path: Path, context: SourceContext, packages: Sequence[Package],
                    *, status: str, **extra: object) -> None:
    _write_json(path, {
        "schema": RECEIPT_SCHEMA,
        "status": status,
        "source_commit": context.source_commit,
        "cargo_lock_sha256": context.lock_sha256,
        "package_count": len(packages),
        "package_set_sha256": package_set_sha256(packages),
        "updated": utc_now(),
        **extra,
    })


def run_inventory(root: Path, output: Path, captures: Sequence[CaptureResult],
                  *, python: str = sys.executable,
                  inventory_script: Path | None = None) -> tuple[Path, Path]:
    if not captures:
        raise InventoryError("cannot run inventory with no captured JSON")
    script = inventory_script or root / "tools" / "api_debt_inventory.py"
    json_paths = [Path(item.json_path) for item in sorted(captures, key=lambda x: x.package)]
    inventory_path = output / "inventory.tsv"
    log_path = output / "inventory.stderr.log"
    with inventory_path.open("w", encoding="utf-8", newline="") as stream, log_path.open("wb") as error:
        result = subprocess.run([python, str(script), *map(str, json_paths)], cwd=root,
                                stdout=stream, stderr=error, check=False)
    if result.returncode != 0:
        raise InventoryError(f"api_debt_inventory.py failed with exit {result.returncode}")
    disposition = write_disposition_scaffold(inventory_path, output / "disposition.tsv")
    return inventory_path, disposition


def write_disposition_scaffold(inventory: Path, destination: Path) -> Path:
    """Add explicit, unresolved decisions to every inventory row."""
    with inventory.open(newline="", encoding="utf-8") as source:
        rows = list(csv.DictReader(source, delimiter="\t"))
    fields = ["crate", "category", "path", "detail", "target",
              "decision", "evidence", "owner"]
    with destination.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=fields, delimiter="\t", lineterminator="\n")
        writer.writeheader()
        for row in rows:
            category = row.get("category", "")
            path = row.get("path", "")
            if category == "mut_draft" and "builder" in path.split("::"):
                decision = "DRAFT_CANDIDATE_REVIEW"
            elif row.get("crate") in {"litchi-cfb", "soapberry-zip"}:
                decision = "SUBSTRATE_SCOPE_REVIEW"
            else:
                decision = "UNASSIGNED_REQUIRED_REVIEW"
            writer.writerow({**row, "decision": decision, "evidence": "", "owner": ""})
    return destination


def capture_all(root: Path, output: Path, target: Path, *,
                cargo: str = "cargo", reuse: bool = True, run_final_inventory: bool = True,
                runner: CommandRunner = default_runner) -> list[CaptureResult]:
    for label, path in (("output", output), ("target", target)):
        try:
            path.relative_to(root)
        except ValueError:
            continue
        raise InventoryError(f"{label} must be outside the source checkout: {path}")
    context = source_context(root, cargo=cargo)
    packages = enumerate_packages(root)
    if len(packages) != DEFAULT_PACKAGE_COUNT:
        raise InventoryError(f"expected {DEFAULT_PACKAGE_COUNT} direct library packages, found {len(packages)}")
    shim_identity = verify_cargo_shim(context.cargo_path, capture_env())
    output.mkdir(parents=True, exist_ok=True)
    marker = output / "run.json"
    previous = _read_json(marker)
    package_digest = package_set_sha256(packages)
    if isinstance(previous, dict):
        if previous.get("source_commit") != context.source_commit:
            raise InventoryError("output directory belongs to a different source freeze")
        if (previous.get("package_count") != DEFAULT_PACKAGE_COUNT
                or previous.get("package_set_sha256") != package_digest):
            raise InventoryError("output directory belongs to a different package set")
    write_package_manifest(output, context, packages)
    write_run_state(marker, context, packages, status="running", phase="captures")
    captures: list[CaptureResult] = []
    try:
        for package in packages:
            source_stable(context)
            captures.append(capture_package(package, context, output, target, cargo=cargo,
                                            reuse=reuse, runner=runner,
                                            shim_identity=shim_identity))
            source_stable(context)
        if not run_final_inventory:
            raise InventoryError("inventory was skipped; run is explicitly incomplete")
        source_stable(context)
        inventory_path, disposition_path = run_inventory(root, output, captures)
        source_stable(context)
        write_run_state(
            marker, context, packages, status="complete", phase="complete",
            inventory=str(inventory_path), disposition=str(disposition_path),
            reused=sum(item.reused for item in captures),
            captured=sum(not item.reused for item in captures),
        )
        return captures
    except Exception as error:
        write_run_state(marker, context, packages, status="incomplete", phase="failed",
                        captured=len(captures), error=str(error))
        raise


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path.cwd())
    parser.add_argument("--output", type=Path, required=True,
                        help="external output directory, never a repository path")
    parser.add_argument("--target-dir", type=Path, required=True,
                        help="one shared Cargo target cache, preferably an existing external cache")
    parser.add_argument("--cargo", default="cargo")
    parser.add_argument("--no-reuse", action="store_true")
    parser.add_argument("--skip-inventory", action="store_true")
    args = parser.parse_args(argv)
    try:
        capture_all(args.root.resolve(), args.output.resolve(), args.target_dir.resolve(),
                    cargo=args.cargo,
                    reuse=not args.no_reuse, run_final_inventory=not args.skip_inventory)
    except (InventoryError, OSError, subprocess.CalledProcessError, ValueError) as error:
        print(f"final API inventory: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
