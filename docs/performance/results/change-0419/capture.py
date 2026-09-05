#!/usr/bin/env python3
"""Capture matched 0419 resource diagnostics in process-isolated ABBA order.

The build driver owns clean source and copied ``/tmp`` binary identities.  This
driver binds those identities to the frozen measurement protocol, runs one
selector in one fresh process per ABBA leg, and records the report, catalog,
GNU ``time -v`` output, stdout, stderr, and a small per-run journal.  It does
not make a latency or allocation claim; the normal and allocator lanes have
different sample contracts and are kept separate.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


CHANGE = 419
LEGS = ("A1", "B1", "B2", "A2")
ROLES = ("control", "candidate", "candidate", "control")
ROLE_FOR_LEG = dict(zip(LEGS, ROLES))
SELECTORS = (
    "pptx_cross_copy_media_rich_lifecycle",
    "pptx_cross_copy_plain_lifecycle",
)
MODES = ("normal", "allocator")
RUNTIME_TOOLCHAIN = "1.98.1"
SAFE_NAME = re.compile(r"^[A-Za-z0-9_.-]+$")
FORBIDDEN_FLAGS = {
    "--case", "--json", "--corpus-manifest", "--samples", "--warmup",
}


class CaptureError(RuntimeError):
    """An input, identity, output, or subprocess failure."""


def fail(message: str) -> None:
    raise CaptureError(message)


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def strict_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON number {value!r}")


def load_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=strict_pairs,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"cannot load strict JSON {path}: {error}")
    raise AssertionError("unreachable")


def write_json(path: Path, value: Any) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    try:
        payload = json.dumps(
            value, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False
        ) + "\n"
        temporary.write_text(payload, encoding="utf-8")
        temporary.replace(path)
    except (OSError, TypeError, ValueError, OverflowError) as error:
        try:
            temporary.unlink()
        except OSError:
            pass
        fail(f"cannot write {path}: {error}")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def digest_json(value: Any) -> str:
    try:
        payload = json.dumps(
            value, sort_keys=True, separators=(",", ":"),
            ensure_ascii=False, allow_nan=False,
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        fail(f"cannot canonicalize identity: {error}")
    return hashlib.sha256(payload).hexdigest()


def object_value(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(f"{label} must be an object")
    return value


def string_value(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{label} must be a non-empty string")
    return value


def integer_value(value: Any, label: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(f"{label} must be an integer >= {minimum}")
    return value


def relative_path(root: Path, raw: Any, label: str) -> Path:
    value = string_value(raw, label)
    path = Path(value)
    if path.is_absolute() or ".." in path.parts:
        fail(f"{label} must be a relative path without '..'")
    resolved = (root / path).resolve()
    try:
        resolved.relative_to(root.resolve())
    except ValueError:
        fail(f"{label} escapes the capture root")
    return resolved


def git_state(worktree: Path) -> dict[str, Any]:
    def run(arguments: list[str]) -> str:
        try:
            result = subprocess.run(
                ["git", *arguments], cwd=worktree, capture_output=True,
                text=True, check=False,
            )
        except OSError as error:
            fail(f"cannot run git in {worktree}: {error}")
        if result.returncode != 0:
            fail(
                f"git {' '.join(arguments)} failed in {worktree}: "
                f"{(result.stderr or result.stdout).strip()}"
            )
        return result.stdout

    revision = run(["rev-parse", "HEAD"]).strip()
    status = run(["status", "--porcelain=v1", "--untracked-files=all"])
    return {"worktree": str(worktree), "revision": revision,
            "clean": status == "", "git_status_porcelain": status}


def validate_clean_source(source: Any, label: str) -> dict[str, Any]:
    value = object_value(source, label)
    worktree = Path(string_value(value.get("worktree"), f"{label}.worktree"))
    if not worktree.is_absolute() or not worktree.is_dir():
        fail(f"{label}.worktree must be an existing absolute directory")
    if value.get("clean") is not True or value.get("git_status_porcelain") != "":
        fail(f"{label} is not clean")
    revision = string_value(value.get("revision"), f"{label}.revision")
    if len(revision) != 40 or revision != revision.lower() or any(
        character not in "0123456789abcdef" for character in revision
    ):
        fail(f"{label}.revision must be a lowercase 40-character git revision")
    return value


def source_binding(build: dict[str, Any], role: str) -> dict[str, Any]:
    before = validate_clean_source(build.get("source_before"), f"{role}.source_before")
    after = validate_clean_source(build.get("source_after"), f"{role}.source_after")
    if before != after:
        fail(f"{role} source_before and source_after differ")
    worktree = Path(before["worktree"]).resolve()
    current = git_state(worktree)
    if current["revision"] != before["revision"] or not current["clean"]:
        fail(f"{role} source worktree changed or is dirty before capture")
    return {
        "worktree": str(worktree),
        "revision": before["revision"],
        "clean": True,
        "identity_sha256": digest_json(before),
    }


def binary_binding(entry: Any, role: str, mode: str) -> dict[str, Any]:
    value = object_value(entry, f"{role}.binaries.{mode}")
    path = Path(string_value(value.get("path"), f"{role}.binaries.{mode}.path"))
    if not path.is_absolute() or not str(path).startswith("/tmp/"):
        fail(f"{role}/{mode} binary must be an absolute copied /tmp binary")
    if not path.is_file() or not os.access(path, os.X_OK):
        fail(f"{role}/{mode} binary is missing or not executable: {path}")
    digest = sha256_file(path)
    size = path.stat().st_size
    expected_digest = string_value(value.get("sha256"), f"{role}.binaries.{mode}.sha256")
    binary_digest = value.get("binary_sha256", expected_digest)
    if binary_digest != expected_digest or digest != expected_digest:
        fail(f"{role}/{mode} binary hash does not match build identity")
    expected_bytes = value.get("bytes", value.get("binary_bytes"))
    if expected_bytes is not None and integer_value(
        expected_bytes, f"{role}.binaries.{mode}.bytes", minimum=1
    ) != size:
        fail(f"{role}/{mode} binary size does not match build identity")
    return {
        "path": str(path), "sha256": digest, "bytes": size,
        "label": value.get("label"),
    }


def protocol_binding(
    root: Path,
    role: str,
    build_protocol_sha256: Any,
    common: dict[str, Any],
) -> dict[str, Any]:
    """Bind a build to current protocol, or the recorded control amendment."""
    build_hash = string_value(
        build_protocol_sha256, f"{role}.protocol_sha256"
    )
    current_hash = common["sha256"]
    if build_hash == current_hash:
        return {
            "kind": "current",
            "build_protocol_sha256": build_hash,
            "current_protocol_sha256": current_hash,
        }
    if role != "control":
        fail(f"{role} build is not bound to the current common protocol")
    original_path = root / "protocol-build-start.json"
    original = object_value(load_json(original_path), str(original_path))
    original_hash = sha256_file(original_path)
    if build_hash != original_hash:
        fail(f"{role} build protocol hash is neither current nor the recorded original")
    current_path = Path(common["absolute_path"])
    comparable_original = dict(original)
    comparable_current = object_value(load_json(current_path), str(current_path))
    original_trace = comparable_original.pop("trace", None)
    current_trace = comparable_current.pop("trace", None)
    if comparable_original != comparable_current:
        fail("protocol amendment changed fields other than trace")
    if not isinstance(original_trace, list) or not isinstance(current_trace, list):
        fail("protocol amendment must retain trace arrays")
    if len(original_trace) != len(current_trace):
        fail("protocol amendment changed trace length")
    replacements: list[dict[str, Any]] = []
    for index, (old, new) in enumerate(zip(original_trace, current_trace)):
        if old == new:
            continue
        if (old, new) == ("--cases", "--case"):
            replacements.append({"index": index, "from": old, "to": new})
        elif (old, new) == ("--output", "--json"):
            replacements.append({"index": index, "from": old, "to": new})
        else:
            fail(f"protocol amendment has an unapproved trace change at index {index}")
    if replacements != [
        {"index": 5, "from": "--cases", "to": "--case"},
        {"index": 11, "from": "--output", "to": "--json"},
    ]:
        fail("protocol amendment trace replacements do not match the recorded correction")
    return {
        "kind": "control_build_amended_trace_only",
        "build_protocol_sha256": build_hash,
        "original_protocol_path": str(original_path.relative_to(root.resolve())),
        "original_protocol_sha256": original_hash,
        "current_protocol_path": str(current_path.relative_to(root.resolve())),
        "current_protocol_sha256": current_hash,
        "trace_replacements": replacements,
    }


def load_build(
    root: Path, role: str, common: dict[str, Any]
) -> dict[str, Any]:
    path = root / f"build-{role}.json"
    value = object_value(load_json(path), str(path))
    if value.get("change") != CHANGE or value.get("role") != role:
        fail(f"{path} has the wrong change or role")
    if value.get("status") != "pass" or value.get("exit_code") != 0:
        fail(f"{path} is not a successful build record")
    binding = protocol_binding(root, role, value.get("protocol_sha256"), common)
    source = source_binding(value, role)
    binaries = {
        mode: binary_binding(
            object_value(value.get("binaries"), f"{role}.binaries").get(mode),
            role, mode,
        )
        for mode in MODES
    }
    return {
        "path": str(path),
        "sha256": sha256_file(path),
        "role": role,
        "source": source,
        "protocol_binding": binding,
        "binaries": binaries,
    }


def load_common_protocol(root: Path, measurement: dict[str, Any]) -> tuple[dict[str, Any], Path]:
    reference = measurement.get("protocol")
    expected_digest = measurement.get(
        "protocol_sha256", measurement.get("common_flags_sha256")
    )
    if isinstance(reference, dict):
        expected_digest = reference.get("sha256", expected_digest)
        reference = reference.get("path", "protocol.json")
    elif reference is None:
        reference = measurement.get(
            "common_flags_source", measurement.get("protocol_path", "protocol.json")
        )
    path = relative_path(root, reference, "measurement-protocol.protocol.path")
    protocol = object_value(load_json(path), str(path))
    if protocol.get("change") != CHANGE:
        fail(f"{path} is not the 0419 common protocol")
    actual_digest = sha256_file(path)
    if expected_digest is not None and expected_digest != actual_digest:
        fail("measurement protocol's common protocol hash does not match protocol.json")
    flags = protocol.get("common_flags")
    if not isinstance(flags, list) or not flags or any(
        not isinstance(flag, str) or not flag for flag in flags
    ):
        fail(f"{path}.common_flags must be a non-empty string list")
    if any(flag in FORBIDDEN_FLAGS for flag in flags):
        fail(f"{path}.common_flags contains a capture-owned flag")
    return {
        "path": str(path.relative_to(root.resolve())),
        "absolute_path": str(path),
        "sha256": actual_digest,
        "common_flags": flags,
    }, path


def mode_contract(measurement: dict[str, Any], mode: str) -> dict[str, int]:
    block: Any = measurement.get(mode)
    if block is None:
        modes = measurement.get("modes")
        if isinstance(modes, dict):
            block = modes.get(mode)
    if block is None:
        fail(f"measurement-protocol.json is missing the {mode} contract")
    block = object_value(block, f"measurement-protocol.json.{mode}")
    return {
        "samples": integer_value(block.get("samples"), f"{mode}.samples", minimum=1),
        "warmups": integer_value(block.get("warmups"), f"{mode}.warmups", minimum=0),
    }


def load_measurement_protocol(root: Path) -> dict[str, Any]:
    path = root / "measurement-protocol.json"
    measurement = object_value(load_json(path), str(path))
    if measurement.get("change") != CHANGE:
        fail(f"{path} is not the 0419 measurement protocol")
    order = measurement.get("order", measurement.get("abba_order", list(LEGS)))
    if order != list(LEGS):
        fail(f"{path} ABBA order must be {list(LEGS)!r}")
    roles_value = measurement.get("roles", list(ROLES))
    if isinstance(roles_value, dict):
        if set(roles_value) != set(LEGS):
            fail(f"{path}.roles must map exactly the four ABBA legs")
        roles = [roles_value.get(leg) for leg in LEGS]
    else:
        roles = roles_value
    if roles != list(ROLES):
        fail(f"{path} ABBA roles must be {list(ROLES)!r}")
    cpu = integer_value(measurement.get("cpu"), f"{path}.cpu", minimum=0)
    if cpu != 2:
        fail(f"{path}.cpu must be the predeclared CPU 2")
    raw_selectors = measurement.get("selectors")
    if not isinstance(raw_selectors, list):
        fail(f"{path}.selectors must be a list")
    selectors: list[str] = []
    for index, item in enumerate(raw_selectors):
        if isinstance(item, str):
            selector = item
        elif isinstance(item, dict):
            selector = item.get("selector", item.get("name"))
        else:
            fail(f"{path}.selectors[{index}] must be a selector or object")
        selector = string_value(selector, f"{path}.selectors[{index}]")
        if not SAFE_NAME.fullmatch(selector):
            fail(f"{path}.selectors[{index}] contains unsafe characters")
        selectors.append(selector)
    if tuple(selectors) != SELECTORS:
        fail(f"{path}.selectors must be {list(SELECTORS)!r}")
    common, common_path = load_common_protocol(root, measurement)
    modes = {mode: mode_contract(measurement, mode) for mode in MODES}
    if modes != {
        "normal": {"samples": 100, "warmups": 10},
        "allocator": {"samples": 30, "warmups": 3},
    }:
        fail(f"{path} must declare normal 100/10 and allocator 30/3 contracts")
    return {
        "path": str(path),
        "sha256": sha256_file(path),
        "change": CHANGE,
        "abba_order": list(LEGS),
        "roles": list(ROLES),
        "cpu": cpu,
        "selectors": selectors,
        "modes": modes,
        "common": common,
        "common_path": str(common_path),
    }


def artifact_record(
    path: Path, root: Path, *, allow_empty: bool = False
) -> dict[str, Any]:
    if not path.is_file():
        fail(f"expected capture artifact is missing: {path}")
    size = path.stat().st_size
    if size == 0 and not allow_empty:
        fail(f"expected capture artifact is empty: {path}")
    return {
        "path": str(path.relative_to(root.resolve())),
        "bytes": size,
        "sha256": sha256_file(path),
    }


def source_after_check(binding: dict[str, Any]) -> dict[str, Any]:
    state = git_state(Path(binding["worktree"]))
    if state["revision"] != binding["revision"] or not state["clean"]:
        fail(f"source changed or became dirty during capture: {binding['worktree']}")
    return {
        "worktree": binding["worktree"],
        "revision": state["revision"],
        "clean": state["clean"],
    }


def binary_after_check(binding: dict[str, Any], role: str, mode: str) -> None:
    if sha256_file(Path(binding["path"])) != binding["sha256"]:
        fail(f"{role}/{mode} copied binary changed during capture")


def run_one(
    *,
    root: Path,
    output: Path,
    measurement: dict[str, Any],
    builds: dict[str, dict[str, Any]],
    mode: str,
    leg: str,
    selector: str,
    index: int,
    total: int,
) -> dict[str, Any]:
    role = ROLE_FOR_LEG[leg]
    build = builds[role]
    binary = build["binaries"][mode]
    source = build["source"]
    contract = measurement["modes"][mode]
    folder = output / mode / leg / selector
    if folder.exists():
        fail(f"refusing to overwrite existing run directory: {folder}")
    folder.mkdir(parents=True)
    report = folder / "report.json"
    catalog = folder / "catalog.json"
    time_report = folder / "time.txt"
    stdout = folder / "stdout.txt"
    stderr = folder / "stderr.txt"
    journal = folder / "journal.json"
    command = [
        "taskset", "-c", str(measurement["cpu"]), "/usr/bin/time", "-v",
        "-o", str(time_report), binary["path"], "--case", selector,
        *measurement["common"]["common_flags"],
        "--samples", str(contract["samples"]), "--warmup", str(contract["warmups"]),
        "--json", str(report), "--corpus-manifest", str(catalog),
    ]
    record: dict[str, Any] = {
        "schema_version": 1,
        "change": CHANGE,
        "status": "running",
        "mode": mode,
        "leg": leg,
        "role": role,
        "selector": selector,
        "index": index,
        "source": source,
        "source_identity_sha256": source["identity_sha256"],
        "source_revision": source["revision"],
        "build": {
            "path": str(Path(build["path"]).relative_to(root.resolve())),
            "sha256": build["sha256"],
            "source_identity_sha256": source["identity_sha256"],
        },
        "build_sha256": build["sha256"],
        "binary": binary,
        "binary_sha256": binary["sha256"],
        "binary_bytes": binary["bytes"],
        "protocol": {
            "measurement_path": str(Path(measurement["path"]).relative_to(root.resolve())),
            "measurement_sha256": measurement["sha256"],
            "common_path": str(Path(measurement["common_path"]).relative_to(root.resolve())),
            "common_sha256": measurement["common"]["sha256"],
            "build_binding": build["protocol_binding"],
        },
        "measurement_protocol_sha256": measurement["sha256"],
        "common_protocol_sha256": measurement["common"]["sha256"],
        "contract": contract,
        "environment": {"RUSTUP_TOOLCHAIN": RUNTIME_TOOLCHAIN},
        "argv": command,
        "artifacts": {
            "report": str(report.relative_to(root.resolve())),
            "catalog": str(catalog.relative_to(root.resolve())),
            "time_v": str(time_report.relative_to(root.resolve())),
            "stdout": str(stdout.relative_to(root.resolve())),
            "stderr": str(stderr.relative_to(root.resolve())),
        },
        "started_utc": utc_now(),
    }
    write_json(journal, record)
    print(f"[{index}/{total}] {mode}/{leg}/{selector}", flush=True)
    env = os.environ.copy()
    env["RUSTUP_TOOLCHAIN"] = RUNTIME_TOOLCHAIN
    try:
        with stdout.open("wb") as stdout_stream, stderr.open("wb") as stderr_stream:
            result = subprocess.run(
                command,
                cwd=source["worktree"],
                env=env,
                stdout=stdout_stream,
                stderr=stderr_stream,
                check=False,
            )
    except (OSError, subprocess.SubprocessError) as error:
        record.update({
            "status": "failed", "error": str(error), "finished_utc": utc_now(),
        })
        write_json(journal, record)
        raise CaptureError(f"{mode}/{leg}/{selector} failed to start: {error}") from error
    record["exit_code"] = result.returncode
    record["finished_utc"] = utc_now()
    try:
        record["source_after"] = source_after_check(source)
        if result.returncode != 0:
            fail(
                f"{mode}/{leg}/{selector} exited {result.returncode}; "
                f"see {stderr.relative_to(root.resolve())}"
            )
        record["artifacts"] = {
            name: artifact_record(path, root, allow_empty=name in ("stdout", "stderr"))
            for name, path in (
                ("report", report), ("catalog", catalog), ("time_v", time_report),
                ("stdout", stdout), ("stderr", stderr),
            )
        }
    except CaptureError as error:
        record["status"] = "failed"
        record["error"] = str(error)
        write_json(journal, record)
        raise
    record["status"] = "pass"
    write_json(journal, record)
    print(f"[{index}/{total}] pass", flush=True)
    return record


def capture(root: Path) -> None:
    root = root.resolve()
    if not root.is_dir():
        fail(f"capture root is not a directory: {root}")
    capture_path = root / "capture.json"
    output = root / "runs"
    if capture_path.exists() or output.exists():
        fail(f"refusing to overwrite existing 0419 capture output: {capture_path} or {output}")
    measurement = load_measurement_protocol(root)
    builds = {
        role: load_build(root, role, measurement["common"])
        for role in ("control", "candidate")
    }
    output.mkdir()
    total = len(MODES) * len(LEGS) * len(SELECTORS)
    manifest: dict[str, Any] = {
        "schema_version": 1,
        "change": CHANGE,
        "status": "running",
        "started_utc": utc_now(),
        "protocol": {
            "measurement_path": str(Path(measurement["path"]).relative_to(root)),
            "measurement_sha256": measurement["sha256"],
            "common_path": str(Path(measurement["common_path"]).relative_to(root)),
            "common_sha256": measurement["common"]["sha256"],
            "abba_order": measurement["abba_order"],
            "roles": measurement["roles"],
            "cpu": measurement["cpu"],
            "selectors": measurement["selectors"],
            "modes": measurement["modes"],
            "common_flags": measurement["common"]["common_flags"],
            "runtime_toolchain": RUNTIME_TOOLCHAIN,
            "build_protocol_bindings": {
                role: builds[role]["protocol_binding"] for role in ("control", "candidate")
            },
        },
        "identities": builds,
        "runs": [],
        "expected_run_count": total,
    }
    write_json(capture_path, manifest)
    try:
        index = 0
        for mode in MODES:
            for leg in LEGS:
                for selector in SELECTORS:
                    index += 1
                    run = run_one(
                        root=root, output=output, measurement=measurement,
                        builds=builds, mode=mode, leg=leg, selector=selector,
                        index=index, total=total,
                    )
                    manifest["runs"].append(run)
                    write_json(capture_path, manifest)
        for role, build in builds.items():
            source_after_check(build["source"])
            for mode in MODES:
                binary_after_check(build["binaries"][mode], role, mode)
        manifest["status"] = "pass"
        manifest["finished_utc"] = utc_now()
        write_json(capture_path, manifest)
    except CaptureError as error:
        manifest["status"] = "failed"
        manifest["finished_utc"] = utc_now()
        manifest["error"] = str(error)
        write_json(capture_path, manifest)
        raise


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--root", type=Path, default=Path(__file__).resolve().parent,
        help="0419 evidence root containing measurement-protocol.json and builds",
    )
    args = parser.parse_args()
    try:
        capture(args.root)
    except CaptureError as error:
        print(f"0419 capture failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps({
        "status": "pass", "change": CHANGE,
        "runs": len(MODES) * len(LEGS) * len(SELECTORS),
    }, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
