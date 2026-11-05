#!/usr/bin/env python3
"""Fail-closed, calibrated ABBA runner for the W8 File probe.

The probe performs leaf/facade ABBA internally. This runner performs the
outer before/after ABBA and keeps the four legs of every block in separate
files. Every process snapshot uses /proc cmdlines (rather than the truncated
``comm`` column) and is taken before, after, and every 100 ms while a child
tree is alive.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
import time
import tomllib
from typing import Any

MIN_SAMPLE_NS = 150_000_000
TARGET_NS = 500_000_000
HEADROOM = 1.25
MAX_CALLS = 2_000_000
SHA256 = re.compile(r"^[0-9a-f]{64}$")
FULL_COMMIT = re.compile(r"^[0-9a-f]{40}$")
EXPECTED_ROUTE_COUNT = 109
REVIEWED_PROBE_SHA256 = "21c8e3833c96411a612a2631f98093bd70cf0a5c28b2952c3d15873f464e54b0"
PACKAGE_NAME = "litchi-ops"
PACKAGE_MANIFEST_RELATIVE = Path("crates/litchi-ops/Cargo.toml")
PROBE_RELATIVE = Path("crates/litchi-ops/examples/file_w8_qualified_probe.rs")
EXPECTED_FAMILIES = (
    "docx", "odt", "doc", "rtf", "xlsx", "xlsb", "xls", "ods",
    "pptx", "ppt", "odp",
)
BASE_ROUTES = (
    "open_get", "save", "repeated_save", "set_commit", "set_publication",
    "batch_commit", "batch_publication", "merge_commit", "merge_publication",
)
WORKBOOK_FAMILIES = ("xlsx", "xlsb", "xls", "ods")
ROUTE_SUFFIXES = {
    "open_get", "save", "repeated_save", "set_commit", "set_publication",
    "batch_commit", "batch_publication", "import_commit", "import_publication",
    "merge_commit", "merge_publication", "refresh_commit", "refresh_publication",
}
FIXTURE_RECIPE = {
    "docx": {"constructor": "docx_refresh", "count": 1000, "batch": 1000, "import_rows": 0, "marker": "/body/p[1]"},
    "odt": {"constructor": "unified_words", "format": "Odt", "count": 1000, "batch": 1000, "import_rows": 0, "marker": "/body/p[1]"},
    "doc": {"constructor": "doc", "count": 1000, "batch": 256, "import_rows": 0, "marker": "/body/p[1]"},
    "rtf": {"constructor": "rtf", "count": 1000, "batch": 256, "import_rows": 0, "marker": "/body/p[1]"},
    "xlsx": {"constructor": "workbook", "format": "Xlsx", "count": 100, "batch": 1000, "import_rows": 1000, "marker": "/sheet[1]/L1"},
    "xlsb": {"constructor": "workbook", "format": "Xlsb", "count": 100, "batch": 1000, "import_rows": 100, "marker": "/sheet[1]/L1"},
    "xls": {"constructor": "workbook", "format": "Xls", "count": 100, "batch": 1000, "import_rows": 100, "marker": "/sheet[1]/L1"},
    "ods": {"constructor": "workbook", "format": "Ods", "count": 100, "batch": 1000, "import_rows": 1000, "marker": "/sheet[1]/L1"},
    "pptx": {"constructor": "pptx", "count": 100, "batch": 1000, "import_rows": 0, "marker": "/slide[1]/shape[1]"},
    "ppt": {"constructor": "ppt", "count": 100, "batch": 1000, "import_rows": 0, "marker": "/slide[1]/shape[1]"},
    "odp": {"constructor": "odp", "count": 100, "batch": 100, "import_rows": 0, "marker": "/slide[1]/shape[1]"},
}
FIXTURE_CONSTANTS = {"target_ns": 500_000_000, "max_calls": 2_000_000, "paragraphs": 1_000, "shapes": 100}
FIXTURE_API = {
    "commit_extraction": "into_file",
    "route_list": "--mode list --route '*'",
    "route_validate": "--mode validate --route FAMILY/ROUTE",
}
SERVICE_EXEMPT_EXE = "/usr/local/bin/peculiar_shell"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def artifact(path: Path) -> dict[str, object]:
    path = path.resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def save(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n")


def canonical(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode()


def full_commit(value: object, label: str) -> str:
    if not isinstance(value, str) or not FULL_COMMIT.fullmatch(value):
        raise RuntimeError(f"{label} is not a full source commit")
    return value


def load_object(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text())
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"{label} is invalid JSON: {error}") from error
    if not isinstance(value, dict):
        raise RuntimeError(f"{label} is not a JSON object")
    return value


def reviewed_routes(value: dict[str, Any], label: str) -> list[str]:
    if value.get("schema") != "litchi_w8_fixture_descriptor.v1":
        raise RuntimeError(f"{label} fixture schema is invalid")
    probe = value.get("probe")
    if (
        not isinstance(probe, dict)
        or probe.get("name") != "file_w8_qualified_probe.rs"
        or probe.get("sha256") != REVIEWED_PROBE_SHA256
    ):
        raise RuntimeError(f"{label} fixture is not bound to the reviewed probe")
    families = EXPECTED_FAMILIES
    routes = value.get("routes")
    if value.get("families") != list(families) or not isinstance(routes, dict):
        raise RuntimeError(f"{label} fixture family/route declaration is incomplete")
    result: list[str] = []
    for family in families:
        entries = routes.get(family)
        if not isinstance(entries, list) or not entries or not all(isinstance(item, str) for item in entries):
            raise RuntimeError(f"{label} fixture routes for {family} are incomplete")
        expected = list(BASE_ROUTES)
        if family in WORKBOOK_FAMILIES:
            expected.extend(("import_commit", "import_publication"))
        if family == "docx":
            expected.extend(("refresh_commit", "refresh_publication"))
        if entries != expected:
            raise RuntimeError(f"{label} fixture route recipe for {family} differs from the reviewed probe")
        result.extend(f"{family}/{item}" for item in entries)
    if len(result) != EXPECTED_ROUTE_COUNT or len(set(result)) != EXPECTED_ROUTE_COUNT or value.get("route_count") != EXPECTED_ROUTE_COUNT:
        raise RuntimeError(f"{label} fixture does not describe exactly {EXPECTED_ROUTE_COUNT} routes")
    if set(value.get("route_suffix_vocabulary", [])) != ROUTE_SUFFIXES:
        raise RuntimeError(f"{label} fixture route vocabulary is incomplete")
    if value.get("recipe") != FIXTURE_RECIPE or value.get("constants") != FIXTURE_CONSTANTS or value.get("api") != FIXTURE_API:
        raise RuntimeError(f"{label} fixture recipe/API markers are not the reviewed values")
    return result


def reviewed_projection(
    value: dict[str, Any],
    fields: dict[str, Any],
    target: object,
    profile: object,
    label: str,
) -> None:
    """Reject a projection that only carries a schema and one convenient hash."""
    if value.get("schema") != "litchi_w8_dependency_projection.v1":
        raise RuntimeError(f"{label} dependency projection schema is invalid")
    workspace = value.get("workspace")
    if (
        not isinstance(workspace, dict)
        or workspace.get("manifest_sha256") != fields["workspace_manifest"]["sha256"]
        or workspace.get("package_manifest_sha256") != fields["package_manifest"]["sha256"]
        or workspace.get("package") != PACKAGE_NAME
        or not isinstance(workspace.get("package_version"), str)
        or not workspace["package_version"]
    ):
        raise RuntimeError(f"{label} workspace projection is incomplete")
    lock = value.get("lock")
    packages = lock.get("packages") if isinstance(lock, dict) else None
    if (
        not isinstance(lock, dict)
        or lock.get("bytes") != fields["lock"]["bytes"]
        or lock.get("sha256") != fields["lock"]["sha256"]
        or not isinstance(packages, list)
        or not packages
    ):
        raise RuntimeError(f"{label} lock projection is incomplete")
    for index, package in enumerate(packages):
        if (
            not isinstance(package, dict)
            or not isinstance(package.get("name"), str)
            or not package["name"]
            or not isinstance(package.get("version"), str)
            or not package["version"]
            or "dependencies" not in package
            or not isinstance(package["dependencies"], list)
        ):
            raise RuntimeError(f"{label} lock package {index} is incomplete")
    if not any(
        package["name"] == PACKAGE_NAME and package["version"] == workspace["package_version"]
        for package in packages
    ):
        raise RuntimeError(f"{label} lock projection has no {PACKAGE_NAME} package")
    features = value.get("features")
    if (
        not isinstance(features, dict)
        or not isinstance(features.get("package"), str)
        or not features["package"]
        or not isinstance(features.get("selected"), list)
        or not features["selected"]
        or not all(isinstance(item, str) and item for item in features["selected"])
        or not isinstance(features.get("all_features"), bool)
    ):
        raise RuntimeError(f"{label} feature projection is incomplete")
    build = value.get("build")
    toolchain = build.get("toolchain") if isinstance(build, dict) else None
    if (
        not isinstance(build, dict)
        or build.get("target") != target
        or build.get("profile") != profile
        or not isinstance(toolchain, dict)
        or toolchain.get("sha256") != fields["toolchain"]["sha256"]
        or not isinstance(toolchain.get("channel"), str)
        or not toolchain["channel"]
        or not isinstance(toolchain.get("components"), list)
        or not all(isinstance(item, str) for item in toolchain["components"])
        or "profile" not in toolchain
    ):
        raise RuntimeError(f"{label} build/toolchain projection is incomplete")
    workload = value.get("workload")
    if (
        not isinstance(workload, dict)
        or workload.get("probe_sha256") != fields["probe"]["sha256"]
        or workload.get("probe_sha256") != REVIEWED_PROBE_SHA256
        or workload.get("fixture_sha256") != fields["fixture"]["sha256"]
        or workload.get("fixture_schema") != "litchi_w8_fixture_descriptor.v1"
        or workload.get("route_count") != EXPECTED_ROUTE_COUNT
    ):
        raise RuntimeError(f"{label} workload projection is incomplete")
    declarations = value.get("package_declarations")
    if (
        not isinstance(declarations, dict)
        or not isinstance(declarations.get("dependencies"), dict)
        or not isinstance(declarations.get("features"), dict)
        or not isinstance(declarations.get("workspace_dependencies"), dict)
    ):
        raise RuntimeError(f"{label} package declarations are incomplete")


def command_line(pid: int) -> tuple[str, str, str]:
    """Return comm, full cmdline, and executable identity."""
    try:
        comm = Path(f"/proc/{pid}/comm").read_text().strip()
        raw = Path(f"/proc/{pid}/cmdline").read_bytes()
        executable = os.readlink(f"/proc/{pid}/exe")
    except (FileNotFoundError, PermissionError):
        return "", "", ""
    return comm, raw.replace(b"\0", b" ").decode(errors="replace").strip(), executable


def process_rows() -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []
    for entry in Path("/proc").iterdir():
        if not entry.name.isdigit():
            continue
        pid = int(entry.name)
        try:
            stat = (entry / "stat").read_text()
            close = stat.rfind(")")
            if close < 0:
                continue
            # After the comm field, field 3 is state and field 4 is ppid.
            parent = int(stat[close + 2 :].split()[1])
        except (FileNotFoundError, PermissionError, ValueError, IndexError):
            continue
        comm, args, executable = command_line(pid)
        rows.append({"pid": pid, "ppid": parent, "comm": comm, "args": args, "exe": executable})
    return rows


def descendants(rows: list[dict[str, object]], roots: set[int]) -> set[int]:
    children: dict[int, list[int]] = {}
    for row in rows:
        children.setdefault(int(row["ppid"]), []).append(int(row["pid"]))
    result = set(roots)
    pending = list(roots)
    while pending:
        parent = pending.pop()
        for child in children.get(parent, []):
            if child not in result:
                result.add(child)
                pending.append(child)
    return result


FORBIDDEN_NAMES = re.compile(
    r"^(?:cargo|rustc|rustdoc|rustfmt|clippy-driver|heaptrack|perf|"
    r"valgrind|callgrind|clang\+\+|ld|lld|ld\.lld|script_program)$"
)
FORBIDDEN_COMMANDS = re.compile(
    r"(?:peculiar(?:[_-]|$)|test-adblock|scoped_patch_probe|"
    r"file_w8_(?:probe|qualified_probe)|run-abba|callgrind|heaptrack|"
    r"(?:pytest|nextest|criterion|hyperfine|strace|dtrace|gdb|lldb|massif)|"
    r"(?:cargo|rustc|rustdoc|rustfmt|clippy)\s+(?:test|bench|run|build|doc|fmt)|"
    r"(?:^|[ /])(?:cargo|rustc|rustdoc|rustfmt|clippy-driver|perf|"
    r"heaptrack|valgrind|callgrind)(?:\s|$))",
    re.IGNORECASE,
)


def forbidden(row: dict[str, object]) -> bool:
    executable = str(row.get("exe", ""))
    args = str(row.get("args", ""))
    comm = str(row.get("comm", ""))
    # This exact installed service is long-lived and unrelated to builds. A
    # build-tree copy or wrapper remains a blocker.
    if executable == SERVICE_EXEMPT_EXE and Path(executable).name == "peculiar_shell":
        return False
    return (
        FORBIDDEN_NAMES.fullmatch(comm) is not None
        or FORBIDDEN_NAMES.fullmatch(Path(executable).name) is not None
        or FORBIDDEN_COMMANDS.search(args) is not None
    )


def guard(out: Path, label: str, allowed_roots: set[int] | None = None) -> str:
    rows = process_rows()
    allowed = descendants(rows, {os.getpid()})
    allowed |= descendants(rows, allowed_roots or set())
    blockers = [
        row
        for row in rows
        if int(row["pid"]) not in allowed
        and (
            forbidden(row)
        )
    ]
    save(
        out / f"{label}.host.json",
        {"time_ns": time.time_ns(), "allowed": sorted(allowed), "blockers": blockers, "processes": rows},
    )
    if blockers:
        raise RuntimeError(f"host overlap at {label}: {blockers}")
    return str((out / f"{label}.host.json").resolve())


def run_probe(
    out: Path,
    label: str,
    binary: Path,
    route: str,
    mode: str,
    calls: int | None,
    blocks: int,
    cpu: int,
) -> dict:
    stdout = out / f"{label}.stdout"
    stderr = out / f"{label}.stderr"
    command = [
        "taskset", "-c", str(cpu), str(binary), "--mode", mode,
        "--route", route, "--blocks", str(blocks),
    ]
    env = os.environ.copy()
    env["LITCHI_FILE_TARGET_NS"] = str(TARGET_NS)
    env["LITCHI_FILE_HEADROOM"] = str(HEADROOM)
    if calls is not None:
        env["LITCHI_FILE_CALLS"] = str(calls)
    before_host = guard(out, f"{label}.before")
    started = time.time_ns()
    with stdout.open("w") as output, stderr.open("w") as errors:
        child = subprocess.Popen(
            command,
            stdout=output,
            stderr=errors,
            env=env,
            start_new_session=True,
        )
        overlap: str | None = None
        during_hosts: list[str] = []
        checks = 0
        while True:
            try:
                status = child.wait(timeout=0.1)
                break
            except subprocess.TimeoutExpired:
                checks += 1
                if overlap is None:
                    try:
                        during_hosts.append(guard(out, f"{label}.during-{checks:05}", {child.pid}))
                    except RuntimeError as error:
                        overlap = str(error)
                        try:
                            os.killpg(child.pid, signal.SIGTERM)
                        except ProcessLookupError:
                            pass
        finished = time.time_ns()
    after_host: str | None = None
    after_error: str | None = None
    try:
        after_host = guard(out, f"{label}.after")
    except RuntimeError as error:
        after_error = str(error)
    host_paths = [before_host, *during_hosts, after_host] if after_host else [before_host, *during_hosts]
    host_artifacts = [artifact(Path(path)) for path in host_paths]
    receipt = {
        "schema": "litchi_file_probe_receipt.v2",
        "label": label,
        "binary": artifact(binary),
        "route": route,
        "mode": mode,
        "blocks": blocks,
        "calls": calls,
        "command": command,
        "environment": {
            "LITCHI_FILE_CALLS": env.get("LITCHI_FILE_CALLS"),
            "target_ns": TARGET_NS,
            "headroom": HEADROOM,
        },
        "started_ns": started,
        "finished_ns": finished,
        "exit_code": status,
        "overlap": overlap,
        "guard_error": after_error,
        "host_snapshots": host_artifacts,
        "stdout": artifact(stdout),
        "stderr": artifact(stderr),
    }
    save(out / f"{label}.receipt.json", receipt)
    if after_error:
        raise RuntimeError(after_error)
    if overlap:
        raise RuntimeError(overlap)
    if status != 0:
        raise RuntimeError(f"probe {label} exited {status}; see {stderr}")
    try:
        value = json.loads(stdout.read_text())
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"invalid probe output {label}: {error}") from error
    if not isinstance(value, dict):
        raise RuntimeError(f"probe {label} did not return an object")
    value["_receipt"] = str((out / f"{label}.receipt.json").resolve())
    value["_host_snapshots"] = host_artifacts
    return value


def routes(binary: Path, out: Path, cpu: int, suffix: str) -> tuple[list[str], dict]:
    row = run_probe(out, f"route-list-{suffix}", binary, "*", "list", None, 1, cpu)
    values = row.get("routes")
    if not isinstance(values, list) or not values or not all(isinstance(item, str) for item in values):
        raise RuntimeError("probe route list is empty or malformed")
    return values, row


def endpoint_manifest(path: Path, binary: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text())
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RuntimeError(f"{label} endpoint manifest is invalid: {error}") from error
    if not isinstance(value, dict) or value.get("schema") != "litchi_file_endpoint.v1":
        raise RuntimeError(f"{label} endpoint manifest has an invalid schema")

    def checked_artifact(name: str) -> dict[str, Any]:
        row = value.get(name)
        if (
            not isinstance(row, dict)
            or not isinstance(row.get("path"), str)
            or not isinstance(row.get("sha256"), str)
            or not SHA256.fullmatch(row["sha256"])
        ):
            raise RuntimeError(f"{label} endpoint manifest lacks a valid {name} artifact")
        actual = artifact(Path(row["path"]))
        if actual["sha256"] != row["sha256"] or actual["bytes"] != row.get("bytes"):
            raise RuntimeError(f"{label} {name} artifact hash does not match its file")
        return actual

    binary_row = checked_artifact("binary")
    actual_binary = artifact(binary)
    if binary_row["sha256"] != actual_binary["sha256"] or binary_row["bytes"] != actual_binary["bytes"]:
        raise RuntimeError(f"{label} endpoint binary hash does not match the selected executable")
    if not os.access(binary, os.X_OK):
        raise RuntimeError(f"{label} selected binary is not executable")
    source_commit = full_commit(value.get("source_commit"), f"{label}.source_commit")
    dependency = value.get("dependency_projection_sha256")
    if not isinstance(source_commit, str) or not source_commit:
        raise RuntimeError(f"{label} endpoint manifest lacks source_commit")
    if not isinstance(dependency, str) or not SHA256.fullmatch(dependency):
        raise RuntimeError(f"{label} endpoint manifest lacks dependency_projection_sha256")
    fields = {
        "manifest": artifact(path),
        "source_commit": source_commit,
        "dependency_projection_sha256": dependency,
        "binary": binary_row,
        "source": checked_artifact("source"),
        "lock": checked_artifact("lock"),
        "probe": checked_artifact("probe"),
        "fixture": checked_artifact("fixture"),
    }
    return verify_endpoint_details(value, fields, label, binary)


def verify_endpoint_details(value: dict[str, Any], fields: dict[str, Any], label: str, binary: Path) -> dict[str, Any]:
    """Verify cross-input provenance before any route calibration starts."""
    def endpoint_artifact(name: str) -> dict[str, Any]:
        row = value.get(name)
        if not isinstance(row, dict) or not isinstance(row.get("path"), str) or not SHA256.fullmatch(str(row.get("sha256", ""))):
            raise RuntimeError(f"{label} endpoint manifest lacks {name}")
        actual = artifact(Path(row["path"]))
        if actual["sha256"] != row["sha256"] or actual["bytes"] != row.get("bytes"):
            raise RuntimeError(f"{label} {name} artifact identity is stale")
        return actual

    def receipt_artifact(actual: dict[str, Any], row: object, name: str) -> None:
        if not isinstance(row, dict) or row.get("bytes") != actual["bytes"] or row.get("sha256") != actual["sha256"]:
            raise RuntimeError(f"{label} {name} does not identify the endpoint artifact")
        if not isinstance(row.get("path"), str):
            raise RuntimeError(f"{label} {name} receipt path is missing")
        receipt_actual = artifact(Path(row["path"]))
        if receipt_actual["bytes"] != row["bytes"] or receipt_actual["sha256"] != row["sha256"]:
            raise RuntimeError(f"{label} {name} receipt artifact is stale")

    projection = endpoint_artifact("dependency_projection")
    projection_value = load_object(Path(projection["path"]), f"{label}.dependency_projection")
    if projection_value.get("schema") != "litchi_w8_dependency_projection.v1" or canonical(projection_value) != Path(projection["path"]).read_bytes():
        raise RuntimeError(f"{label} dependency projection is invalid or noncanonical")
    if projection["sha256"] != fields["dependency_projection_sha256"]:
        raise RuntimeError(f"{label} dependency projection hash differs from endpoint")
    workspace = endpoint_artifact("workspace_manifest")
    package = endpoint_artifact("package_manifest")
    toolchain = endpoint_artifact("toolchain")
    fields["workspace_manifest"] = workspace
    fields["package_manifest"] = package
    fields["toolchain"] = toolchain
    workspace_path = Path(workspace["path"]).resolve()
    if workspace_path.name != "Cargo.toml":
        raise RuntimeError(f"{label} workspace manifest is not Cargo.toml")
    expected_package = (workspace_path.parent / PACKAGE_MANIFEST_RELATIVE).resolve()
    expected_probe = (workspace_path.parent / PROBE_RELATIVE).resolve()
    if package["path"] != str(expected_package):
        raise RuntimeError(f"{label} package manifest is not crates/litchi-ops/Cargo.toml")
    if fields["probe"]["path"] != str(expected_probe):
        raise RuntimeError(f"{label} probe is not crates/litchi-ops/examples/file_w8_qualified_probe.rs")
    try:
        workspace_data = tomllib.loads(workspace_path.read_text(encoding="utf-8"))
        package_data = tomllib.loads(Path(package["path"]).read_text(encoding="utf-8"))
    except (OSError, UnicodeError, tomllib.TOMLDecodeError) as error:
        raise RuntimeError(f"{label} workspace/package manifest is invalid: {error}") from error
    package_table = package_data.get("package")
    if not isinstance(package_table, dict) or package_table.get("name") != PACKAGE_NAME:
        raise RuntimeError(f"{label} package manifest does not declare {PACKAGE_NAME}")
    package_version = package_table.get("version")
    if package_version is True or (
        isinstance(package_version, dict) and package_version.get("workspace") is True
    ):
        workspace_table = workspace_data.get("workspace")
        workspace_package = workspace_table.get("package") if isinstance(workspace_table, dict) else None
        package_version = workspace_package.get("version") if isinstance(workspace_package, dict) else None
    if not isinstance(package_version, str) or not package_version:
        raise RuntimeError(f"{label} package version is not resolved")
    if projection_value.get("workspace", {}).get("package_version") != package_version:
        raise RuntimeError(f"{label} projection package version differs from the manifests")
    features = value.get("features")
    if not isinstance(features, dict) or projection_value.get("features") != features:
        raise RuntimeError(f"{label} feature projection is missing or stale")
    reviewed_projection(projection_value, fields, value.get("target"), value.get("profile"), label)

    fixture_value = load_object(Path(fields["fixture"]["path"]), f"{label}.fixture")
    if canonical(fixture_value) != Path(fields["fixture"]["path"]).read_bytes():
        raise RuntimeError(f"{label}.fixture is not canonical JSON")
    expected_routes = reviewed_routes(fixture_value, f"{label}.fixture")
    if fixture_value.get("probe", {}).get("sha256") != fields["probe"]["sha256"] or projection_value.get("workload", {}).get("probe_sha256") != fields["probe"]["sha256"] or projection_value.get("workload", {}).get("fixture_sha256") != fields["fixture"]["sha256"]:
        raise RuntimeError(f"{label} workload projection is stale")

    assembly = endpoint_artifact("source_assembly")
    assembly_value = load_object(Path(assembly["path"]), f"{label}.source_assembly")
    if (
        assembly_value.get("schema") != "litchi_w8_source_assembly.v1"
        or canonical(assembly_value) != Path(assembly["path"]).read_bytes()
        or full_commit(assembly_value.get("resolved_commit"), f"{label}.source_assembly.resolved_commit") != fields["source_commit"]
        or full_commit(assembly_value.get("source_commit"), f"{label}.source_assembly.source_commit") != fields["source_commit"]
    ):
        raise RuntimeError(f"{label} source assembly is stale")
    receipt_artifact(fields["source"], assembly_value.get("source_archive"), f"{label}.source_assembly.source_archive")
    overlay = assembly_value.get("overlay")
    if not isinstance(overlay, dict) or not isinstance(overlay.get("description"), str) or not overlay["description"].strip():
        raise RuntimeError(f"{label} source assembly has no explicit overlay description")
    if overlay.get("kind") == "compatibility":
        full_commit(overlay.get("base_commit"), f"{label}.source_assembly.overlay.base_commit")
        patch = overlay.get("patch")
        if not isinstance(patch, dict):
            raise RuntimeError(f"{label} compatibility overlay patch is missing")
        receipt_artifact(artifact(Path(patch.get("path", ""))), patch, f"{label}.source_assembly.overlay.patch")
    elif overlay.get("kind") != "none":
        raise RuntimeError(f"{label} source assembly overlay kind is invalid")

    validation = endpoint_artifact("validation")
    validation_value = load_object(Path(validation["path"]), f"{label}.validation")
    if (
        validation_value.get("schema") != "litchi_file_untimed_validation.v1"
        or full_commit(validation_value.get("source_commit"), f"{label}.validation.source_commit") != fields["source_commit"]
    ):
        raise RuntimeError(f"{label} validation receipt source/schema is invalid")
    outcomes = validation_value.get("outcomes")
    if validation_value.get("route_count") != EXPECTED_ROUTE_COUNT or not isinstance(outcomes, list) or len(outcomes) != EXPECTED_ROUTE_COUNT:
        raise RuntimeError(f"{label} validation receipt is not complete")
    if [item.get("route") for item in outcomes if isinstance(item, dict)] != expected_routes:
        raise RuntimeError(f"{label} validation receipt routes differ from fixture")
    for index, outcome in enumerate(outcomes):
        if not isinstance(outcome, dict) or outcome.get("index") != index:
            raise RuntimeError(f"{label} validation outcome indexes are incomplete")
        if any(outcome.get(key) is not True for key in ("semantic_equal", "semantic_state_valid", "source_state_preserved", "source_state_exact", "exact_noop")):
            raise RuntimeError(f"{label} validation route {outcome.get('route')} is not valid")
    validation_artifacts = validation_value.get("artifacts")
    if not isinstance(validation_artifacts, dict):
        raise RuntimeError(f"{label} validation artifacts are missing")
    for name in ("source", "binary", "lock", "probe"):
        receipt_artifact(fields[name] if name != "binary" else fields["binary"], validation_artifacts.get(name), f"{label}.validation.{name}")
    routes_artifact = validation_artifacts.get("routes")
    if not isinstance(routes_artifact, dict):
        raise RuntimeError(f"{label} validation raw routes artifact is missing")
    receipt_artifact(artifact(Path(routes_artifact.get("path", ""))), routes_artifact, f"{label}.validation.routes")

    build = endpoint_artifact("build_receipt")
    build_value = load_object(Path(build["path"]), f"{label}.build_receipt")
    if (
        build_value.get("schema") != "litchi_w8_build_receipt.v1"
        or canonical(build_value) != Path(build["path"]).read_bytes()
        or build_value.get("status") != "complete"
        or full_commit(build_value.get("source_commit"), f"{label}.build_receipt.source_commit") != fields["source_commit"]
    ):
        raise RuntimeError(f"{label} build receipt is incomplete or stale")
    for name, receipt_name, actual in (
        ("source", "source_archive", fields["source"]),
        ("binary", "binary", fields["binary"]),
        ("lock", "lock", fields["lock"]),
        ("probe", "probe", fields["probe"]),
        ("fixture", "fixture", fields["fixture"]),
        ("source_assembly", "source_assembly", assembly),
    ):
        receipt_artifact(actual, build_value.get(receipt_name), f"{label}.build_receipt.{receipt_name}")
    if build_value.get("dependency_projection_sha256") != fields["dependency_projection_sha256"] or build_value.get("workspace_manifest_sha256") != workspace["sha256"] or build_value.get("package_manifest_sha256") != package["sha256"] or build_value.get("toolchain_sha256") != toolchain["sha256"] or build_value.get("features") != features or build_value.get("target") != value.get("target") or build_value.get("profile") != value.get("profile"):
        raise RuntimeError(f"{label} build receipt provenance differs from endpoint")
    if not isinstance(build_value.get("build_command"), str) or not build_value["build_command"].strip():
        raise RuntimeError(f"{label} build receipt command is missing")
    compiler_artifacts = build_value.get("compiler_artifacts")
    if not isinstance(compiler_artifacts, list) or not compiler_artifacts:
        raise RuntimeError(f"{label} compiler artifact hashes are missing")
    for index, compiler_artifact in enumerate(compiler_artifacts):
        if not isinstance(compiler_artifact, dict) or not compiler_artifact.get("name"):
            raise RuntimeError(f"{label} compiler artifact {index} is incomplete")
        receipt_artifact(artifact(Path(compiler_artifact.get("path", ""))), compiler_artifact, f"{label}.compiler_artifact[{index}]")
    fields["routes"] = expected_routes
    return fields


def shared_identity(before: dict[str, Any], after: dict[str, Any]) -> dict[str, str]:
    fields = ("dependency_projection_sha256", "lock", "probe", "fixture")
    for field in fields:
        if field == "dependency_projection_sha256":
            left, right = before[field], after[field]
        else:
            left, right = before[field]["sha256"], after[field]["sha256"]
        if left != right:
            raise RuntimeError(f"endpoint {field} identities differ: {left} != {right}")
    return {
        "dependency_projection_sha256": before["dependency_projection_sha256"],
        "lock_sha256": before["lock"]["sha256"],
        "probe_sha256": before["probe"]["sha256"],
        "fixture_sha256": before["fixture"]["sha256"],
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("before", type=Path)
    parser.add_argument("after", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("rounds", type=int, nargs="?", default=5)
    parser.add_argument("cpu", type=int, nargs="?", default=29)
    parser.add_argument("--before-manifest", required=True, type=Path)
    parser.add_argument("--after-manifest", required=True, type=Path)
    parser.add_argument("--original-to-final-policy", type=Path)
    parser.add_argument("--original-to-final-receipt", type=Path)
    args = parser.parse_args()
    if args.rounds < 5:
        parser.error("at least five ABBA blocks are required")
    if args.cpu not in os.sched_getaffinity(0):
        parser.error(f"CPU {args.cpu} is unavailable")
    before, after, out = (p.resolve() for p in (args.before, args.after, args.output))
    if before == after:
        parser.error("before and after must be distinct protected binaries")
    if not before.is_file() or not os.access(before, os.X_OK) or not after.is_file() or not os.access(after, os.X_OK):
        parser.error("before and after must be executable protected binaries")
    if (args.original_to_final_policy is None) != (args.original_to_final_receipt is None):
        parser.error("--original-to-final-policy and --original-to-final-receipt must be supplied together")
    if out.exists() and any(out.iterdir()):
        parser.error(f"refusing to overwrite nonempty output directory {out}")
    out.mkdir(parents=True, exist_ok=True)
    pair_policy = None
    if args.original_to_final_policy is not None:
        repo_root = Path(__file__).resolve().parents[4]
        if str(repo_root) not in sys.path:
            sys.path.insert(0, str(repo_root))
        from tools import w8_projection_policy

        pair_policy = w8_projection_policy
    if pair_policy is None:
        before_endpoint = endpoint_manifest(args.before_manifest.resolve(), before, "before")
        after_endpoint = endpoint_manifest(args.after_manifest.resolve(), after, "after")
        # The ordinary protocol remains strict shared identity.  The explicit
        # original-to-final policy is opt-in and cannot loosen this path.
        shared = shared_identity(before_endpoint, after_endpoint)
    else:
        before_endpoint = pair_policy.validate_endpoint_manifest(args.before_manifest.resolve(), before, "before")
        after_endpoint = pair_policy.validate_endpoint_manifest(args.after_manifest.resolve(), after, "after")
        receipt_path = args.original_to_final_receipt.resolve()
        pair = pair_policy.verify_pair_receipt(
            args.original_to_final_policy.resolve(), receipt_path, before_endpoint, after_endpoint,
        )
        shared = {
            "mode": "original_to_final",
            "policy": pair["policy"],
            "receipt": artifact(receipt_path),
            "projection": pair["projection"],
            "workload": pair["workload"],
            "build": pair["build"],
        }
    before_provenance = pair_policy.endpoint_public_fields(before_endpoint) if pair_policy is not None else before_endpoint
    after_provenance = pair_policy.endpoint_public_fields(after_endpoint) if pair_policy is not None else after_endpoint
    guard(out, "initial")
    before_routes, before_route_receipt = routes(before, out, args.cpu, "before")
    after_routes, after_route_receipt = routes(after, out, args.cpu, "after")
    # Pair validation exposes an immutable endpoint capability, whose recursive
    # freeze stores route sequences as tuples.  The route probe returns its
    # canonical JSON-shaped list, so normalize the validated value at this
    # boundary before enforcing the exact reviewed order.
    expected_routes = list(before_endpoint["routes"])
    if before_routes != expected_routes or after_routes != expected_routes:
        raise RuntimeError("endpoint validation does not cover the complete 109-route list")
    if before_routes != after_routes:
        raise RuntimeError("before/after route lists differ")
    manifest = {
        "schema": "litchi_file_qualification.v3",
        "before": artifact(before),
        "after": artifact(after),
        "routes": before_routes,
        "route_lists": {"before": before_route_receipt, "after": after_route_receipt},
        "rounds": args.rounds,
        "cpu": args.cpu,
        "target_ns": TARGET_NS,
        "minimum_sample_ns": MIN_SAMPLE_NS,
        "headroom": HEADROOM,
        "guard": "full /proc comm+cmdline+exe snapshots before/after and every 100ms; only this runner and the child subtree are exempt",
        "provenance": {
            "before": before_provenance,
            "after": after_provenance,
            "shared": shared,
        },
        "status": "running",
        "calibration": [],
        "abba": [],
    }
    save(out / "manifest.json", manifest)
    for index, route in enumerate(before_routes):
        safe = re.sub(r"[^A-Za-z0-9_.-]", "_", route)
        calibrations = []
        for side, binary in (("before", before), ("after", after)):
            calibration = run_probe(out, f"{safe}-{side}-calibration", binary, route, "calibrate", None, 1, args.cpu)
            if calibration.get("route") != route or not isinstance(calibration.get("common_calls"), int):
                raise RuntimeError(f"malformed calibration for {route}/{side}")
            common = int(calibration["common_calls"])
            if common <= 0 or common > MAX_CALLS or calibration.get("minimum_elapsed_ns", 0) < MIN_SAMPLE_NS:
                raise RuntimeError(f"calibration cannot meet target for {route}/{side}: {calibration}")
            calibrations.append(calibration)
        common_calls = max(int(item["common_calls"]) for item in calibrations)
        if common_calls > MAX_CALLS:
            raise RuntimeError(f"shared calibration exceeds ceiling for {route}: {common_calls}")
        manifest["calibration"].append({"route": route, "rows": calibrations, "common_calls": common_calls})
        save(out / "manifest.json", manifest)
        blocks = []
        route_row = {"route": route, "calls": common_calls, "blocks": blocks}
        manifest["abba"].append(route_row)
        save(out / "manifest.json", manifest)
        for block_index in range(args.rounds):
            legs = []
            for leg_index, side in enumerate(("before", "after", "after", "before")):
                binary = before if side == "before" else after
                value = run_probe(
                    out,
                    f"{safe}-{index:02}-block-{block_index:02}-leg-{leg_index}-{side}",
                    binary,
                    route,
                    "measure",
                    common_calls,
                    1,
                    args.cpu,
                )
                if value.get("route") != route or value.get("calls") != common_calls:
                    raise RuntimeError(f"shared call count mismatch for {route}/{side}")
                legs.append({"side": side, "value": value})
            blocks.append({"block": block_index, "legs": legs})
            save(out / "manifest.json", manifest)
        save(out / "manifest.json", manifest)
    guard(out, "finished")
    manifest["status"] = "complete; analyze.py review pending"
    save(out / "manifest.json", manifest)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"0362h qualification failed: {error}", file=sys.stderr)
        raise
