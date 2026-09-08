#!/usr/bin/env python3
"""Capture one authenticated lane of the frozen 0470 targeted guard.

The caller must serialize this command with the experiment CPU lock.  The
driver refuses an existing output lane, checks the bound binary and complete
source inventory, switches the shared absolute checkout to the bound revision,
and records the exact command and protocol hash before starting the workload.
"""

from __future__ import annotations

import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
from typing import Any, Mapping


ROOT = Path(__file__).resolve().parent
PROTOCOL_PATH = ROOT / "guard-protocol.json"
TEMP = Path("/tmp/litchi-goal-0470")
SHA256_HEX = 64


class CaptureError(RuntimeError):
    """Raised when the frozen capture custody check fails."""


def _fail(message: str) -> None:
    raise CaptureError(message)


def _sha(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        _fail(f"not a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
                size += len(block)
    except OSError as error:
        _fail(f"cannot hash {path}: {error}")
    return digest.hexdigest(), size


def _json(path: Path, label: str) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        _fail(f"{label}: missing or non-regular file")
    try:
        with path.open(encoding="utf-8") as stream:
            value = json.load(stream)
    except (OSError, UnicodeError, ValueError) as error:
        _fail(f"{label}: invalid JSON ({error})")
    if not isinstance(value, dict):
        _fail(f"{label}: expected JSON object")
    return value


def _text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        _fail(f"{label}: expected non-empty text")
    return value


def _relative(root: Path, value: Any, label: str) -> Path:
    raw = _text(value, label)
    path = Path(raw)
    if path.is_absolute() or path.as_posix() != raw or ".." in path.parts:
        _fail(f"{label}: expected a relative traversal-free path")
    resolved = (root / path).resolve()
    try:
        resolved.relative_to(root.resolve())
    except ValueError:
        _fail(f"{label}: path escapes the evidence bundle")
    return resolved


def _run(command: list[str], *, cwd: Path, env: Mapping[str, str], text: bool = True) -> str:
    try:
        return subprocess.check_output(command, cwd=cwd, env=env, text=text, stderr=subprocess.STDOUT)
    except (OSError, subprocess.CalledProcessError) as error:
        _fail(f"command failed: {' '.join(command)} ({error})")
    raise AssertionError("unreachable")


def _utc() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _protocol() -> tuple[dict[str, Any], str]:
    protocol = _json(PROTOCOL_PATH, "guard-protocol.json")
    protocol_sha, _ = _sha(PROTOCOL_PATH)
    if protocol.get("schema") != "litchi-0470-targeted-guard-protocol-v1":
        _fail("guard protocol schema differs")
    if protocol.get("order") != ["guard-A1", "guard-B1", "guard-B2", "guard-A2"]:
        _fail("guard protocol order differs")
    if protocol.get("cases") != [
        "cfb_shared_read_one", "opc_noop_save", "cfb_read_one", "opc_source_open", "ppt_fresh_write_to",
    ]:
        _fail("guard protocol case list differs")
    if protocol.get("shapes") != ["many-small", "few-large", "wide-root"]:
        _fail("guard protocol shape list differs")
    if protocol.get("payload_kinds") != ["compressible", "incompressible"]:
        _fail("guard protocol payload list differs")
    if protocol.get("writer_shapes") != ["payload-heavy"]:
        _fail("guard protocol writer-shape list differs")
    if protocol.get("samples") != 100 or protocol.get("warmups") != 5:
        _fail("guard protocol sample configuration differs")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        _fail("guard protocol CPU/worker identity differs")
    if protocol.get("expected_rows") != 25:
        _fail("guard protocol row cardinality differs")
    if protocol.get("performance_claim") != "none":
        _fail("guard protocol permits a performance claim")
    if not isinstance(protocol.get("capture_driver_sha256"), str) or len(protocol["capture_driver_sha256"]) != SHA256_HEX:
        _fail("guard protocol capture driver hash is missing")
    if protocol["capture_driver_sha256"] != _sha(Path(__file__))[0]:
        _fail("guard protocol capture driver hash differs")
    return protocol, protocol_sha


def _role_for_lane(lane: str, protocol: Mapping[str, Any]) -> str:
    roles = protocol.get("roles")
    if not isinstance(roles, dict) or lane not in roles:
        _fail(f"unknown guard lane: {lane}")
    role = roles[lane]
    if role not in ("control", "candidate"):
        _fail(f"{lane}: invalid role")
    return role


def _authenticate_role(role: str, protocol: Mapping[str, Any]) -> tuple[dict[str, Any], Path, Path, dict[str, str], dict[str, str]]:
    binding = _json(ROOT / f"{role}-binding.json", f"{role}-binding.json")
    if binding.get("role") != role or binding.get("schema") != "litchi-0470-role-binding-v1":
        _fail(f"{role}: role binding schema differs")
    revision = _text(binding.get("revision"), f"{role}.revision")
    binary = Path(_text(binding.get("binary_path"), f"{role}.binary_path"))
    worktree = Path(_text(binding.get("build_path"), f"{role}.build_path"))
    if not binary.is_absolute() or not worktree.is_absolute():
        _fail(f"{role}: binary/build paths must be absolute")
    if str(worktree) != protocol.get("build_path"):
        _fail(f"{role}: build path differs from frozen protocol")
    binary_paths = protocol.get("binary_paths")
    if not isinstance(binary_paths, dict) or str(binary) != binary_paths.get(role):
        _fail(f"{role}: binary path differs from frozen protocol")
    bound_binary_sha, bound_binary_bytes = _sha(binary)
    if bound_binary_sha != binding.get("binary_sha256") or bound_binary_bytes != binding.get("bytes"):
        _fail(f"{role}: binary identity differs from role binding")
    manifest_path = _relative(ROOT, binding.get("source_manifest"), f"{role}.source_manifest")
    manifest_sha, _ = _sha(manifest_path)
    if manifest_sha != binding.get("source_manifest_sha256"):
        _fail(f"{role}: source manifest hash differs")
    manifest = _json(manifest_path, f"{role} source manifest")
    if not manifest:
        _fail(f"{role}: source manifest is empty")
    source_inventory: dict[str, str] = {}
    for name, expected in sorted(manifest.items()):
        if not isinstance(name, str) or not isinstance(expected, str) or len(expected) != SHA256_HEX:
            _fail(f"{role}: malformed source manifest entry")
        # The shared checkout may still contain the other role at this point.
        # Authenticate the expected inventory now and hash its files only
        # after _clean_role_checkout has switched to this role.
        _relative(worktree, name, f"{role}.sources.{name}")
        source_inventory[name] = expected
    fixtures = binding.get("included_fixtures")
    if not isinstance(fixtures, dict) or not fixtures:
        _fail(f"{role}: fixture inventory is missing")
    fixture_inventory: dict[str, str] = {}
    for name, expected in sorted(fixtures.items()):
        _relative(worktree, name, f"{role}.fixtures.{name}")
        if not isinstance(expected, str) or len(expected) != SHA256_HEX:
            _fail(f"{role}: malformed fixture inventory entry")
        fixture_inventory[name] = expected
    return binding, binary, worktree, source_inventory, fixture_inventory


def _clean_role_checkout(worktree: Path, revision: str, env: Mapping[str, str]) -> None:
    before = _run(["git", "status", "--porcelain"], cwd=worktree, env=env)
    if before.strip():
        _fail(f"source checkout is dirty before role checkout: {worktree}")
    _run(["git", "checkout", "--detach", revision], cwd=worktree, env=env)
    actual = _run(["git", "rev-parse", "HEAD"], cwd=worktree, env=env).strip()
    if actual != revision:
        _fail("role checkout revision differs")
    after = _run(["git", "status", "--porcelain"], cwd=worktree, env=env)
    if after.strip():
        _fail(f"source checkout is dirty after role checkout: {worktree}")


def _verify_report_identity(path: Path, binding: Mapping[str, Any], revision: str) -> bool:
    report = _json(path, str(path))
    environment = report.get("environment")
    binary_identity = report.get("binary_identity")
    if not isinstance(environment, dict) or not isinstance(binary_identity, dict):
        return False
    return (
        environment.get("git_revision") == revision
        and environment.get("git_worktree_dirty") is False
        and binary_identity.get("binary_sha256") == binding.get("binary_sha256")
        and binary_identity.get("binary_bytes") == binding.get("bytes")
    )


def main(argv: list[str] | None = None) -> int:
    lane = (argv if argv is not None else sys.argv[1:])
    if len(lane) != 1:
        raise CaptureError("usage: guard_capture.py guard-A1|guard-B1|guard-B2|guard-A2")
    lane = lane[0]
    protocol, protocol_sha = _protocol()
    role = _role_for_lane(lane, protocol)
    binding, binary, worktree, source_inventory, fixture_inventory = _authenticate_role(role, protocol)
    output = ROOT / lane
    if output.exists():
        _fail(f"refusing existing guard output lane: {output}")

    environment_values = protocol.get("environment")
    if not isinstance(environment_values, dict):
        _fail("guard protocol environment is missing")
    env = dict(os.environ)
    env.update({str(key): str(value) for key, value in environment_values.items()})
    env["PYTHONDONTWRITEBYTECODE"] = "1"
    _clean_role_checkout(worktree, str(binding["revision"]), env)
    for name, expected in source_inventory.items():
        actual, _ = _sha(_relative(worktree, name, f"{role}.sources.{name}"))
        if actual != expected:
            _fail(f"{role}: source digest changed after checkout for {name}")
    for name, expected in fixture_inventory.items():
        actual, _ = _sha(_relative(worktree, name, f"{role}.fixtures.{name}"))
        if actual != expected:
            _fail(f"{role}: fixture digest changed after checkout for {name}")

    output.mkdir()
    samples = int(protocol["samples"])
    warmups = int(protocol["warmups"])
    cases = protocol["cases"]
    shapes = protocol["shapes"]
    writer_shapes = protocol["writer_shapes"]
    command = [
        "taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o", str(output / "resource.log"),
        str(binary), "--workers", str(protocol["workers"]), "--warmup", str(warmups),
        "--samples", str(samples), "--case", ",".join(cases), "--shape", ",".join(shapes),
        "--writer-shape", ",".join(writer_shapes), "--json", str(output / "report.json"),
        "--corpus-manifest", str(output / "corpus-catalog.json"),
    ]
    captured_env = {key: env[key] for key in (
        "RUSTUP_TOOLCHAIN", "RUSTFLAGS", "CARGO_PROFILE_RELEASE_DEBUG", "CARGO_INCREMENTAL",
        "CARGO_BUILD_JOBS", "DEBUGINFOD_URLS", "LC_ALL",
    )}
    receipt = {
        "schema": "litchi-0470-targeted-guard-capture-v1",
        "lane": lane,
        "role": role,
        "revision": binding["revision"],
        "binary_path": str(binary),
        "binary_sha256": _sha(binary)[0],
        "binary_bytes": _sha(binary)[1],
        "binding_sha256": _sha(ROOT / f"{role}-binding.json")[0],
        "source_manifest": binding["source_manifest"],
        "source_manifest_sha256": binding["source_manifest_sha256"],
        "source_inventory_count": len(source_inventory),
        "source_inventory_sha256": _sha(_relative(ROOT, binding["source_manifest"], f"{role}.source_manifest"))[0],
        "fixture_inventory": fixture_inventory,
        "capture_driver_sha256": _sha(Path(__file__))[0],
        "protocol_sha256": protocol_sha,
        "argv": command,
        "cwd": str(worktree),
        "samples": samples,
        "warmups": warmups,
        "environment": captured_env,
        "started_utc": _utc(),
        "clean_before": True,
        "source_inventory_verified_before": True,
    }
    (output / "started.json").write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    with (output / "stdout.log").open("x", encoding="utf-8") as stdout, (output / "stderr.log").open("x", encoding="utf-8") as stderr:
        result = subprocess.run(command, cwd=worktree, env=env, stdout=stdout, stderr=stderr)
    _clean_role_checkout(worktree, str(binding["revision"]), env)
    binary_sha, binary_bytes = _sha(binary)
    report_ok = result.returncode == 0 and _verify_report_identity(output / "report.json", binding, str(binding["revision"]))
    receipt.update({
        "exit_code": result.returncode,
        "clean_after": True,
        "binary_unchanged": binary_sha == binding.get("binary_sha256") and binary_bytes == binding.get("bytes"),
        "report_metadata_matches_clean_role": report_ok,
        "finished_utc": _utc(),
        "artifacts": {
            path.name: {"sha256": _sha(path)[0], "bytes": _sha(path)[1]}
            for path in sorted(output.iterdir())
            if path.is_file() and path.name != "receipt.json"
        },
    })
    (output / "receipt.json").write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    if result.returncode != 0 or not receipt["binary_unchanged"] or not report_ok:
        return result.returncode or 1
    print(json.dumps({"lane": lane, "role": role, "exit_code": result.returncode}, sort_keys=True), flush=True)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (CaptureError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"guard capture failed: {error}", file=sys.stderr)
        raise SystemExit(1)
