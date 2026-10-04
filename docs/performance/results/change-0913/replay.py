"""Replay the retained 0913 custody after its temporary roots are gone.

This is an offline verifier.  It reconstructs the frozen 0910 source into a
temporary directory, applies the 0913 overlay patch, and compares both full
inventories with their retained manifests.  It then checks the retained
binary and packet descriptors and reruns the 0913 reducer into a temporary
output.  It never builds, starts a benchmark, invokes a profiler, or writes a
packet in this directory (the requested output is the one new result file).
"""

from __future__ import annotations

import argparse
import hashlib
import io
import json
from pathlib import Path
import subprocess
import sys
import tarfile
import tempfile
import re
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OLD = P.parent / "change-0910"

BASELINE_PARENT = "468fe2931d6532c32f4c6608590816b8eaa88e31"
CANDIDATE_PARENT = "66b944b07fcf4c6fa37f10e7ed83a02417455f57"
OLD_MARKER = ".owned-0910"
OLD_MARKER_CONTENT = BASELINE_PARENT + "\n"
CANDIDATE_MARKER = ".owned-0913"
CANDIDATE_MARKER_CONTENT = "owned-direct-read-0913\n"
EXPECTED_OVERLAYS = (
    "crates/soapberry-zip/src/office.rs",
    "crates/litchi-opc/src/payload.rs",
    "crates/soapberry-zip/tests/owned_direct_read.rs",
    "crates/litchi-opc/tests/owned_direct_concurrency.rs",
)
EXPECTED_RUNTIME_FIXTURES = {
    "3rdparty/Open-XML-SDK/test/DocumentFormat.OpenXml.Tests.Assets/assets/TestFiles/svg.docx": "docs/performance/results/change-0913/quality-inputs/svg.docx",
    "3rdparty/libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf164835_nonDummyLineHeight.docx": "docs/performance/results/change-0913/quality-inputs/tdf164835_nonDummyLineHeight.docx",
    "3rdparty/libreoffice-core/sd/qa/unit/data/pptx/tdf169496_hidden_graphic.pptx": "docs/performance/results/change-0913/quality-inputs/tdf169496_hidden_graphic.pptx",
    "3rdparty/libreoffice-core/sd/qa/unit/data/pptx/tdf163852.pptx": "docs/performance/results/change-0913/quality-inputs/tdf163852.pptx",
    "3rdparty/libreoffice-core/sd/qa/unit/data/pptx/tdf164622.pptx": "docs/performance/results/change-0913/quality-inputs/tdf164622.pptx",
}

EPHEMERAL_ROOTS = (
    Path("/tmp/litchi-0910-src"),
    Path("/tmp/litchi-0913-src"),
    ROOT / "target/litchi-0910-build",
    ROOT / "target/litchi-0913-build",
)


class ReplayError(ValueError):
    """Retained evidence is missing, changed, or internally inconsistent."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReplayError(message)


def load(path: Path, label: str | None = None) -> Any:
    label = label or str(path)
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {label}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ReplayError(f"invalid JSON {label}: {error}") from error


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path, base: Path = ROOT) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    try:
        relative = str(path.relative_to(base.resolve()))
    except ValueError:
        relative = str(path)
    return {"path": relative, "bytes": path.stat().st_size, "sha256": sha256(path)}


def descriptor_shape(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: descriptor is malformed")
    path = value.get("path")
    require(isinstance(path, str) and path, f"{label}: descriptor path is missing")
    size = value.get("bytes")
    require(type(size) is int and size >= 0, f"{label}: descriptor byte count is invalid")
    digest = value.get("sha256")
    require(
        isinstance(digest, str)
        and len(digest) == 64
        and all(char in "0123456789abcdef" for char in digest),
        f"{label}: descriptor SHA-256 is invalid",
    )
    return {"path": path, "bytes": size, "sha256": digest}


def resolve_path(raw: str) -> Path:
    path = Path(raw)
    if path.is_absolute():
        return path.resolve()
    # Every retained run.py descriptor is workspace-relative.  The packet
    # fallback also makes this verifier tolerant of old packet-local paths.
    workspace = (ROOT / path).resolve()
    packet = (P / path).resolve()
    if packet.exists() or packet.is_symlink():
        return packet
    # Keep a missing workspace-relative path anchored at the workspace.  This
    # is needed for original build binaries whose target root was removed by
    # cleanup; choosing the packet fallback would hide the intended ephemeral
    # path and make the custody check fail for the wrong reason.
    return workspace


def is_ephemeral(path: Path) -> bool:
    path = path.resolve()
    return any(path == root.resolve() or root.resolve() in path.parents for root in EPHEMERAL_ROOTS)


def check_descriptor(value: Any, label: str, *, allow_removed: bool = False) -> dict[str, Any]:
    expected = descriptor_shape(value, label)
    path = resolve_path(expected["path"])
    if not path.exists():
        require(allow_removed and is_ephemeral(path), f"{label}: described file is missing: {path}")
        return expected
    actual = descriptor(path)
    require(actual == expected, f"{label}: descriptor changed")
    return actual


def check_identity(left: Any, right: Any, label: str) -> None:
    left = descriptor_shape(left, f"{label}.left")
    right = descriptor_shape(right, f"{label}.right")
    require(
        left["bytes"] == right["bytes"] and left["sha256"] == right["sha256"],
        f"{label}: byte identity differs",
    )


def check_descriptor_tree(value: Any, label: str) -> int:
    """Check every descriptor-shaped object nested in a retained receipt."""
    count = 0
    if isinstance(value, dict):
        if {"path", "bytes", "sha256"}.issubset(value):
            check_descriptor(value, label)
            return 1
        for key, child in value.items():
            count += check_descriptor_tree(child, f"{label}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            count += check_descriptor_tree(child, f"{label}[{index}]")
    return count


def descriptor_occurs(value: Any, expected: dict[str, Any]) -> bool:
    """Return whether a nested value carries the exact retained descriptor."""
    if isinstance(value, dict):
        if {"path", "bytes", "sha256"}.issubset(value):
            return descriptor_shape(value, "nested descriptor") == expected
        return any(descriptor_occurs(child, expected) for child in value.values())
    if isinstance(value, list):
        return any(descriptor_occurs(child, expected) for child in value)
    return False


def retained_candidate_sources() -> dict[str, dict[str, Any]]:
    raw = load(P / "candidate-retention.json")
    require(raw.get("status") == "pass", "candidate retention did not pass")
    rows = raw.get("files", [])
    expected = set(EXPECTED_OVERLAYS) | {"crates/soapberry-zip/tests/owned_direct_adversarial.rs"}
    require(len(rows) == 5 and {row.get("logical_path") for row in rows} == expected, "candidate retained source set differs")
    result = {}
    for row in rows:
        name = row["logical_path"]
        original = descriptor_shape(row["original"], f"retained candidate {name}")
        retained = check_descriptor(row["retained"], f"retained candidate {name}")
        require(original["path"] == name and retained["path"] == str((P / "candidate-source" / name).relative_to(ROOT)), "candidate retained source path differs")
        require((original["bytes"], original["sha256"]) == (retained["bytes"], retained["sha256"]), "candidate retained source identity differs")
        result[name] = original
    return result


def check_supplemental_revision() -> dict[str, Any]:
    """Validate the retained v1 source and the corrected supplemental test."""
    path = P / "supplemental-owned-direct-adversarial-v2.json"
    raw = load(path, "0913 supplemental-owned-direct-adversarial-v2.json")
    require(isinstance(raw, dict) and raw.get("status") == "pass", "0913 supplemental revision did not pass")
    require(raw.get("kind") == "supplemental-test-revision" and raw.get("revision") == 2, "0913 supplemental revision identity differs")
    source_path = raw.get("path")
    require(source_path == "crates/soapberry-zip/tests/owned_direct_adversarial.rs", "0913 supplemental revision source path differs")
    previous = check_descriptor(raw.get("previous"), "0913 supplemental revision.previous")
    current = descriptor_shape(raw.get("current"), "0913 supplemental revision.current")
    require(current == retained_candidate_sources()[source_path], "0913 supplemental retained current source differs")
    require(previous["path"] == "docs/performance/results/change-0913/quality-inputs/owned_direct_adversarial-v1.rs", "0913 supplemental v1 retention path differs")
    require(current["path"] == source_path, "0913 supplemental v2 source binding differs")
    require(previous["bytes"] != current["bytes"] or previous["sha256"] != current["sha256"], "0913 supplemental revision did not change the test")
    require(isinstance(raw.get("change"), str) and raw["change"], "0913 supplemental revision description is missing")
    return descriptor(path)


def inventory(root: Path) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for path in sorted(root.rglob("*")):
        require(not path.is_symlink(), f"source reconstruction contains symlink: {path}")
        if path.is_file():
            rows.append(
                {
                    "path": str(path.relative_to(root)),
                    "bytes": path.stat().st_size,
                    "sha256": sha256(path),
                }
            )
    return rows


def validate_inventory(actual: list[dict[str, Any]], expected: Any, label: str) -> None:
    require(isinstance(expected, list), f"{label}: inventory is missing")
    require(len(actual) == len(expected), f"{label}: file count differs")
    require(actual == expected, f"{label}: full file inventory differs")


def write_lock_overlays(target: Path, base: dict[str, Any]) -> None:
    overlays = base.get("lock_overlays")
    require(isinstance(overlays, dict), "source-base: lock overlays are missing")
    for name, row in overlays.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute() and ".." not in Path(name).parts, "source-base: unsafe lock overlay path")
        require(isinstance(row, dict), f"source-base: malformed lock overlay {name}")
        artifact = row.get("artifact")
        require(isinstance(artifact, str), f"source-base: lock overlay artifact missing: {name}")
        data_path = resolve_path(artifact)
        check_descriptor({"path": artifact, "bytes": data_path.stat().st_size, "sha256": row.get("sha256")}, f"source-base.lock_overlays.{name}")
        destination = target / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        destination.write_bytes(data_path.read_bytes())


def archive_source(target: Path, base: dict[str, Any]) -> None:
    commit = base.get("commit")
    paths = base.get("archive_paths")
    require(commit == BASELINE_PARENT, "source-base: baseline commit differs")
    require(isinstance(paths, list) and all(isinstance(path, str) for path in paths), "source-base: archive paths are malformed")
    archive = subprocess.check_output(["git", "archive", commit, "--", *paths], cwd=ROOT)
    require(sha256_bytes(archive) == base.get("archive_sha256"), "source-base: archive digest differs")
    with tarfile.open(fileobj=io.BytesIO(archive)) as stream:
        stream.extractall(target, filter="data")


def validate_old_source(old_manifest: dict[str, Any], base: dict[str, Any], target: Path) -> list[dict[str, Any]]:
    require(old_manifest.get("parent") == BASELINE_PARENT, "0910 source manifest parent differs")
    expected = old_manifest.get("files")
    require(isinstance(expected, list) and len(expected) == 9355, "0910 source manifest must contain 9355 files")
    archive_source(target, base)
    write_lock_overlays(target, base)
    marker = target / OLD_MARKER
    marker.write_text(OLD_MARKER_CONTENT, encoding="utf-8")
    patch_row = old_manifest.get("overlay_patch")
    patch = resolve_path(descriptor_shape(patch_row, "0910 overlay patch")["path"])
    check_descriptor(patch_row, "0910 overlay patch")
    subprocess.run(["git", "apply", str(patch)], cwd=target, check=True)
    actual = inventory(target)
    validate_inventory(actual, expected, "0910 reconstructed source")
    require((target / OLD_MARKER).read_text(encoding="utf-8") == OLD_MARKER_CONTENT, "0910 marker differs")
    return actual


def validate_candidate_source(
    manifest: dict[str, Any], old_manifest: dict[str, Any], target: Path
) -> list[dict[str, Any]]:
    require(manifest.get("parent") == CANDIDATE_PARENT, "0913 source manifest parent differs")
    require(manifest.get("overlay_paths") == list(EXPECTED_OVERLAYS), "0913 overlay set differs")
    overlays = manifest.get("overlays")
    require(isinstance(overlays, list) and len(overlays) == len(EXPECTED_OVERLAYS), "0913 overlay trace is incomplete")
    by_path = {row.get("path"): row for row in overlays if isinstance(row, dict)}
    require(set(by_path) == set(EXPECTED_OVERLAYS), "0913 overlay trace differs")

    baseline_binding = manifest.get("baseline_source_manifest")
    require(check_descriptor(baseline_binding, "0913 baseline source manifest") == descriptor(OLD / "source-v1.json"), "0913 baseline manifest binding differs")
    patch_row = manifest.get("overlay_patch")
    patch_path = resolve_path(descriptor_shape(patch_row, "0913 overlay patch")["path"])
    check_descriptor(patch_row, "0913 overlay patch")
    subprocess.run(["git", "apply", str(patch_path)], cwd=target, check=True)
    marker = target / CANDIDATE_MARKER
    require(not marker.exists(), "0913 marker was already present before replay")
    marker.write_text(CANDIDATE_MARKER_CONTENT, encoding="utf-8")

    actual = inventory(target)
    validate_inventory(actual, manifest.get("files"), "0913 reconstructed source")
    require((target / OLD_MARKER).read_text(encoding="utf-8") == OLD_MARKER_CONTENT, "0910 marker was not preserved")
    require((target / CANDIDATE_MARKER).read_text(encoding="utf-8") == CANDIDATE_MARKER_CONTENT, "0913 marker differs")

    rows = {row["path"]: row for row in actual}
    for name in EXPECTED_OVERLAYS:
        row = by_path[name]
        require(isinstance(row.get("workspace"), dict) and isinstance(row.get("source"), dict), f"0913 overlay custody missing: {name}")
        workspace = descriptor_shape(row["workspace"], f"0913 overlay workspace {name}")
        source = descriptor_shape(row["source"], f"0913 overlay source {name}")
        require(workspace == retained_candidate_sources()[name], f"0913 retained overlay differs: {name}")
        require(workspace == source, f"0913 overlay workspace/source differs: {name}")
        require(rows[name] == source, f"0913 overlay inventory differs: {name}")
    return actual


def reconstruct_sources(manifest: dict[str, Any]) -> dict[str, Any]:
    old_base_path = OLD / "source-base.json"
    old_manifest_path = OLD / "source-v1.json"
    old_base = load(old_base_path, "0910 source-base.json")
    old_manifest = load(old_manifest_path, "0910 source-v1.json")
    base_row = old_manifest.get("base")
    require(check_descriptor(base_row, "0910 source base") == descriptor(old_base_path), "0910 source base binding differs")
    with tempfile.TemporaryDirectory(prefix="litchi-0913-source-replay-") as directory:
        target = Path(directory)
        old_actual = validate_old_source(old_manifest, old_base, target)
        candidate_actual = validate_candidate_source(manifest, old_manifest, target)
        # The two inventories are deliberately retained in the result only as
        # counts; all bytes have already been compared before this directory
        # is removed by TemporaryDirectory.
        return {
            "baseline_files": len(old_actual),
            "candidate_files": len(candidate_actual),
            "baseline_manifest": descriptor(old_manifest_path),
        }


def check_logs(receipt: dict[str, Any], label: str) -> int:
    logs = receipt.get("logs")
    require(isinstance(logs, list) and len(logs) == 2, f"{label}: stdout/stderr receipts missing")
    for index, row in enumerate(logs):
        check_descriptor(row, f"{label}.logs[{index}]")
    return len(logs)


def check_start(folder: Path, receipt: dict[str, Any], label: str) -> None:
    start = load(folder / "start.json", f"{label}.start.json")
    require(isinstance(start, dict), f"{label}.start: malformed")
    require(start.get("argv") == receipt.get("argv"), f"{label}: start/command argv differs")
    require(start.get("start_utc") == receipt.get("start_utc"), f"{label}: start/command timestamp differs")


def baseline_binary() -> dict[str, Any]:
    metadata = load(OLD / "binary.json", "0910 binary.json")
    retained = check_descriptor(metadata.get("retained"), "0910 retained binary")
    original = descriptor_shape(metadata.get("original"), "0910 original binary")
    require(original["bytes"] == retained["bytes"] and original["sha256"] == retained["sha256"], "0910 binary relocation identity differs")
    # The retained 0910 binary is the control identity used by every 0913
    # child.  Its original build path may have been deleted during cleanup.
    check_descriptor(metadata.get("original"), "0910 original binary", allow_removed=True)
    return retained


def check_old_build_receipt(control: dict[str, Any]) -> int:
    path = OLD / "build-v1" / "command.json"
    receipt = load(path, "0910 build-v1/command.json")
    require(receipt.get("exit_code") == 0, "0910 build receipt failed")
    check_start(path.parent, receipt, "0910 build")
    check_logs(receipt, "0910 build")
    check_descriptor(receipt.get("source"), "0910 build.source")
    check_descriptor(receipt.get("helper"), "0910 build.helper")
    binary = descriptor_shape(receipt.get("binary"), "0910 build.binary")
    require(binary["bytes"] == control["bytes"] and binary["sha256"] == control["sha256"], "0910 build binary identity differs")
    check_descriptor(receipt.get("binary"), "0910 build.binary", allow_removed=True)
    return 1


def check_candidate_build(manifest: dict[str, Any], plan: dict[str, Any], control: dict[str, Any], attempt: str) -> dict[str, Any]:
    folder = P / f"build-{attempt}"
    receipt = load(folder / "command.json", f"build-{attempt}/command.json")
    require(receipt.get("exit_code") == 0, f"build-{attempt}: build failed")
    check_start(folder, receipt, f"build-{attempt}")
    argv = receipt.get("argv")
    require(isinstance(argv, list) and all(isinstance(item, str) for item in argv), f"build-{attempt}: argv is malformed")
    for flag in ("cargo", "build", "--locked", "--offline", "--release", "--example", "pptx_owned_benchmark", "-j", "2"):
        require(flag in argv, f"build-{attempt}: missing build flag {flag}")
    source = descriptor(manifest_path := P / f"source-{attempt}.json")
    plan_descriptor = descriptor(P / "capture-plan.json")
    require(receipt.get("source") == source, f"build-{attempt}: source binding differs")
    require(receipt.get("helper") == descriptor(P / "run.py"), f"build-{attempt}: helper binding differs")
    require(receipt.get("plan") == plan_descriptor, f"build-{attempt}: plan binding differs")
    require(isinstance(receipt.get("toolchain"), dict), f"build-{attempt}: toolchain receipt missing")
    for name in ("cargo", "rustc"):
        row = receipt["toolchain"].get(name)
        require(isinstance(row, dict) and row.get("launcher") and row.get("version"), f"build-{attempt}: {name} identity missing")
    check_logs(receipt, f"build-{attempt}")

    candidate_original = descriptor_shape(receipt.get("binary"), f"build-{attempt}.binary")
    retained = check_descriptor(receipt.get("retained_binary"), f"build-{attempt}.retained_binary")
    require(candidate_original["bytes"] == retained["bytes"] and candidate_original["sha256"] == retained["sha256"], f"build-{attempt}: retained binary identity differs")
    check_descriptor(receipt.get("binary"), f"build-{attempt}.binary", allow_removed=True)
    require(receipt.get("baseline_binary") == control, f"build-{attempt}: baseline binary binding differs")

    binary_meta_path = folder / "binary.json"
    binary_meta = load(binary_meta_path, f"build-{attempt}/binary.json")
    check_descriptor(binary_meta.get("retained"), f"build-{attempt}.binary.json.retained")
    check_descriptor(binary_meta.get("original"), f"build-{attempt}.binary.json.original", allow_removed=True)
    check_identity(binary_meta.get("original"), binary_meta.get("retained"), f"build-{attempt}.binary.json")
    require(binary_meta.get("retained") == retained, f"build-{attempt}: sidecar retained binary differs")
    return {"receipt": receipt, "candidate": retained, "baseline": control, "plan": plan, "manifest": source}


def check_capture_packets(plan: dict[str, Any], build: dict[str, Any], attempt: str) -> dict[str, int]:
    """Hash every retained capture packet before the reducer is replayed.

    analyze.py repeats these checks while validating the report semantics.  We
    keep this small custody pass here as well so the replay result records the
    actual packet cardinalities, including any future auxiliary artifact.
    """
    legs = plan.get("legs")
    require(isinstance(legs, list) and legs, "0913 capture plan has no legs")
    source = build["manifest"]
    plan_descriptor = descriptor(P / "capture-plan.json")
    candidate = build["candidate"]
    control = build["baseline"]
    commands = logs = artifacts = process_metrics = 0
    labels: set[str] = set()
    for leg in legs:
        require(isinstance(leg, dict), "0913 capture plan contains a malformed leg")
        label = leg.get("label")
        require(isinstance(label, str) and label not in labels, "0913 capture labels are not unique")
        labels.add(label)
        folder = P / "capture" / label
        receipt = load(folder / "command.json", f"capture/{label}/command.json")
        require(receipt.get("exit_code") == 0, f"capture {label}: child did not pass")
        check_start(folder, receipt, f"capture {label}")
        require(receipt.get("leg") == leg and receipt.get("variant") == leg.get("variant"), f"capture {label}: leg binding differs")
        require(receipt.get("source") == source, f"capture {label}: source binding differs")
        require(receipt.get("helper") == descriptor(P / "capture.py"), f"capture {label}: helper binding differs")
        require(receipt.get("build_helper") == descriptor(P / "run.py"), f"capture {label}: build helper binding differs")
        require(check_descriptor(receipt.get("host_environment"), f"capture {label}.host_environment") == descriptor(P / "environment-after-recovery.json"), f"capture {label}: host environment binding differs")
        require(check_descriptor(receipt.get("process_tools"), f"capture {label}.process_tools") == descriptor(P / "process-tools.json"), f"capture {label}: process tool metadata binding differs")
        require(receipt.get("plan") == plan_descriptor, f"capture {label}: plan binding differs")
        expected = control if leg.get("variant") == "baseline" else candidate
        require(receipt.get("binary") == expected, f"capture {label}: executable binding differs")
        require(receipt.get("baseline_binary") == control and receipt.get("candidate_binary") == candidate, f"capture {label}: binary custody differs")
        logs += check_logs(receipt, f"capture {label}")
        rows = receipt.get("artifacts")
        require(isinstance(rows, list) and rows, f"capture {label}: artifacts are missing")
        artifact_root = (folder / "artifacts").resolve()
        seen: set[Path] = set()
        for index, row in enumerate(rows):
            path = resolve_path(descriptor_shape(row, f"capture {label}.artifacts[{index}]")["path"])
            require(path.parent == artifact_root and path not in seen, f"capture {label}: artifact escaped or duplicated")
            seen.add(path)
            check_descriptor(row, f"capture {label}.artifacts[{index}]")
            artifacts += 1
        process = receipt.get("process_metrics")
        require(process is not None, f"capture {label}: process metrics are missing")
        check_descriptor(process, f"capture {label}.process_metrics")
        process_metrics += 1
        commands += 1
    require(commands == len(legs), "0913 capture command count differs from plan")
    return {"commands": commands, "logs": logs, "artifacts": artifacts, "process_metrics": process_metrics}


def check_quality_receipts(manifest: dict[str, Any]) -> dict[str, int | str]:
    """Check every retained quality attempt while admitting failed attempts.

    The create-only quality helpers retain an aborted or failing attempt for
    audit.  Only the numerically latest quality-vN attempt is an admission
    gate; this lets a later corrected attempt supersede an earlier failure.
    """
    source_expected = descriptor(P / "source-v1.json")
    fixtures_path = P / "fixtures.json"
    fixtures_expected = descriptor(fixtures_path) if fixtures_path.is_file() else None
    extra_fixtures_path = P / "extra-fixtures.json"
    extra_fixtures_expected = descriptor(extra_fixtures_path) if extra_fixtures_path.is_file() else None
    more_fixtures_path = P / "more-fixtures.json"
    more_fixtures_expected = descriptor(more_fixtures_path) if more_fixtures_path.is_file() else None
    runtime_fixtures_path = P / "runtime-fixtures.json"
    runtime_fixtures_expected = descriptor(runtime_fixtures_path) if runtime_fixtures_path.is_file() else None
    supplemental_revision_expected = check_supplemental_revision()
    attempts: list[tuple[int, Path]] = []
    for path in sorted(P.iterdir()):
        if not path.is_dir():
            continue
        if path.name == "quality":
            attempts.append((0, path))
        else:
            match = re.fullmatch(r"quality-v(\d+)", path.name)
            if match:
                attempts.append((int(match.group(1)), path))
    require(attempts, "0913 quality attempts are missing")
    command_count = 0
    failed_count = 0
    latest_number, latest = max(attempts, key=lambda item: item[0])
    latest_by_kind: dict[str, dict[str, Any]] = {}
    quality_kinds = ("fmt", "tests", "clippy", "docs")
    for number, folder in attempts:
        commands = sorted(folder.glob("*/command.json"))
        require(commands, f"{folder.name}: no quality command receipts")
        seen_kinds: set[str] = set()
        for path in commands:
            receipt = load(path, str(path.relative_to(P)))
            label = str(path.relative_to(P))
            kind = path.parent.name
            require(kind in quality_kinds and kind not in seen_kinds, f"{label}: unknown or duplicate quality command kind")
            seen_kinds.add(kind)
            require(isinstance(receipt, dict) and isinstance(receipt.get("exit_code"), int), f"{label}: malformed quality receipt")
            descriptor_count = check_descriptor_tree(receipt, label)
            require(descriptor_count >= 3, f"{label}: retained descriptor bindings are incomplete")
            if "source" in receipt:
                require(check_descriptor(receipt["source"], f"{label}.source") == source_expected, f"{label}: source binding differs")
            if "fixtures" in receipt and fixtures_expected is not None:
                require(check_descriptor(receipt["fixtures"], f"{label}.fixtures") == fixtures_expected, f"{label}: fixture binding differs")
            if "extra_fixtures" in receipt and extra_fixtures_expected is not None:
                require(check_descriptor(receipt["extra_fixtures"], f"{label}.extra_fixtures") == extra_fixtures_expected, f"{label}: extra fixture binding differs")
            if "more_fixtures" in receipt and more_fixtures_expected is not None:
                require(check_descriptor(receipt["more_fixtures"], f"{label}.more_fixtures") == more_fixtures_expected, f"{label}: more fixture binding differs")
            if "runtime_fixtures" in receipt and runtime_fixtures_expected is not None:
                require(check_descriptor(receipt["runtime_fixtures"], f"{label}.runtime_fixtures") == runtime_fixtures_expected, f"{label}: runtime fixture binding differs")
            command_count += 1
            if receipt["exit_code"] != 0:
                failed_count += 1
            if folder == latest:
                latest_by_kind[kind] = receipt
    require(set(latest_by_kind) == set(quality_kinds), f"{latest.name}: final quality attempt is incomplete")
    latest_commands = [latest_by_kind[kind] for kind in quality_kinds]
    require(all(receipt["exit_code"] == 0 for receipt in latest_commands), f"{latest.name}: final quality attempt did not pass")
    for key, expected in (("source", source_expected), ("fixtures", fixtures_expected), ("extra_fixtures", extra_fixtures_expected), ("more_fixtures", more_fixtures_expected), ("runtime_fixtures", runtime_fixtures_expected)):
        if expected is None:
            continue
        for receipt in latest_commands:
            require(key in receipt and check_descriptor(receipt[key], f"{latest.name}.{key}") == expected, f"{latest.name}: {key} binding is missing or differs")
    for receipt in latest_commands:
        require(descriptor_occurs(receipt, supplemental_revision_expected), f"{latest.name}: supplemental revision binding is missing or differs")
    final_record_path = P / "quality.json"
    final_record = load(final_record_path, "0913 quality.json")
    require(isinstance(final_record, dict) and final_record.get("status") == "pass", "0913 final quality receipt did not pass")
    declared_commands = final_record.get("commands")
    require(declared_commands == [receipt.get("argv") for receipt in latest_commands], "0913 final quality receipt commands differ")
    return {"attempts": len(attempts), "commands": command_count, "failed_attempts": failed_count, "final": latest.name}


def check_census() -> int:
    path = P / "census-runs.json"
    raw = load(path, "0913 census-runs.json")
    require(isinstance(raw, dict) and raw.get("status") == "pass", "0913 census did not pass")
    check_descriptor(raw.get("helper"), "0913 census helper")
    runs = raw.get("runs")
    require(isinstance(runs, list) and len(runs) == 4, "0913 census must cover four variants")
    expected_labels = {"materialized-baseline", "materialized-candidate", "owned-baseline", "owned-candidate"}
    changed = sorted(
        f"ppt/slides/slide{number + 1}.xml"
        for number in (0, 28, 57, 85, 114, 142, 171, 199)
    )
    labels: set[str] = set()
    run_rows: list[dict[str, Any]] = []
    for index, run in enumerate(runs):
        require(isinstance(run, dict), f"census run {index}: malformed")
        label = run.get("label")
        require(label in expected_labels and label not in labels, f"census run {index}: label differs")
        labels.add(label)
        require(run.get("exit_code") == 0 and run.get("physical_order_equal") is True, f"census run {label}: failed guard")
        require(run.get("changed_members") == changed and run.get("unchanged_members") == 437, f"census run {label}: changed-member set differs")
        check_descriptor(run.get("helper"), f"census run {label}.helper")
        report_path = resolve_path(descriptor_shape(run.get("report"), f"census run {label}.report")["path"])
        check_descriptor(run.get("report"), f"census run {label}.report")
        report = load(report_path, f"census report {label}")
        require(report.get("status") == "pass" and not report.get("violations"), f"census report {label}: failed")
        require(len(report.get("members", [])) == 445 and len(report.get("unchanged_members", [])) == 437, f"census report {label}: member count differs")
        require(sorted(report.get("changed_members", {}).get("names", [])) == changed, f"census report {label}: changed names differ")
        require(report.get("ordered_filenames", {}).get("equal") is True, f"census report {label}: physical order differs")
        require(report.get("archive", {}).get("comment_comparison", {}).get("equal") is True, f"census report {label}: archive comment differs")
        run_rows.append(run)
    require(labels == expected_labels, "0913 census variant set is incomplete")
    # Re-run only the retained Python ZIP census helper.  Its report is the
    # independent raw custody check; the PPTX producer and both binaries stay
    # completely out of this replay.
    with tempfile.TemporaryDirectory(prefix="litchi-0913-census-replay-") as directory:
        temporary = Path(directory)
        for index, run in enumerate(run_rows):
            argv = run.get("argv")
            require(isinstance(argv, list) and all(isinstance(item, str) for item in argv), f"census run {index}: argv is malformed")
            require("--report" in argv, f"census run {index}: report option is missing")
            original = resolve_path(descriptor_shape(run["report"], f"census run {index}.report")["path"])
            replay_report = temporary / f"{index}.json"
            replay_argv = list(argv)
            replay_argv[replay_argv.index("--report") + 1] = str(replay_report)
            subprocess.run(replay_argv, cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
            require(replay_report.read_bytes() == original.read_bytes(), f"census run {run['label']}: replay report differs byte-for-byte")
    return len(runs)


def check_fixtures() -> int:
    fixtures = P / "fixtures.json"
    if not fixtures.exists():
        return 0
    raw = load(fixtures, "0913 fixtures.json")
    require(isinstance(raw, dict) and raw.get("status") == "pass", "0913 fixture admission did not pass")
    require(raw.get("benchmark_production_files_changed") is False, "0913 fixture admission changed production files")
    require(raw.get("parent") == CANDIDATE_PARENT and raw.get("archive_paths") == ["test-data"], "0913 fixture parent/scope differs")

    def archive_rows(record: dict[str, Any], label: str) -> dict[str, tuple[int, str]]:
        paths = record.get("archive_paths")
        require(isinstance(paths, list) and paths and all(isinstance(path, str) for path in paths), f"{label}: archive paths are malformed")
        archive = subprocess.check_output(["git", "archive", CANDIDATE_PARENT, "--", *paths], cwd=ROOT)
        require(sha256_bytes(archive) == record.get("archive_sha256"), f"{label}: archive digest differs")
        result: dict[str, tuple[int, str]] = {}
        with tarfile.open(fileobj=io.BytesIO(archive)) as stream:
            for member in stream.getmembers():
                if not member.isfile():
                    continue
                extracted = stream.extractfile(member)
                require(extracted is not None, f"{label}: archive member cannot be read: {member.name}")
                data = extracted.read()
                result[member.name] = (len(data), sha256_bytes(data))
        return result

    archived = archive_rows(raw, "0913 fixtures")
    added = raw.get("added")
    require(isinstance(added, list) and added, "0913 fixture admission has no additions")
    supplemental = "crates/soapberry-zip/tests/owned_direct_adversarial.rs"
    rows = {row.get("path"): row for row in added if isinstance(row, dict)}
    require(len(rows) == len(added), "0913 fixture admission contains duplicate or malformed paths")
    for name, row in rows.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute() and ".." not in Path(name).parts, "0913 fixture admission path is unsafe")
        require(type(row.get("bytes")) is int and row["bytes"] >= 0, f"0913 fixture admission byte count is invalid: {name}")
        digest = row.get("sha256")
        require(isinstance(digest, str) and len(digest) == 64 and all(char in "0123456789abcdef" for char in digest), f"0913 fixture admission digest is invalid: {name}")
        if name != supplemental:
            require(archived.get(name) == (row["bytes"], row["sha256"]), f"0913 archived fixture bytes differ: {name}")
    require(supplemental in rows, "0913 supplemental fixture is not retained")
    row = rows[supplemental]
    revision = load(P / "supplemental-owned-direct-adversarial-v2.json")
    previous = check_descriptor(revision["previous"], "0913 original supplemental fixture")
    require((previous["bytes"], previous["sha256"]) == (row["bytes"], row["sha256"]), "0913 original supplemental fixture binding differs")

    extra_path = P / "extra-fixtures.json"
    extra = load(extra_path, "0913 extra-fixtures.json")
    require(isinstance(extra, dict) and extra.get("status") == "pass", "0913 extra fixture admission did not pass")
    require(extra.get("parent") == CANDIDATE_PARENT and extra.get("archive_paths") == ["docs/performance/results/change-0416/corpus"], "0913 extra fixture parent/scope differs")
    extra_archived = archive_rows(extra, "0913 extra fixtures")
    extra_added = extra.get("added")
    require(isinstance(extra_added, list) and len(extra_added) == 13, "0913 extra fixture count differs")
    extra_rows = {row.get("path"): row for row in extra_added if isinstance(row, dict)}
    require(len(extra_rows) == len(extra_added), "0913 extra fixture admission contains duplicate or malformed paths")
    for name, row in extra_rows.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute() and ".." not in Path(name).parts, "0913 extra fixture path is unsafe")
        require(extra_archived.get(name) == (row.get("bytes"), row.get("sha256")), f"0913 extra archived fixture bytes differ: {name}")

    more_path = P / "more-fixtures.json"
    more = load(more_path, "0913 more-fixtures.json")
    require(isinstance(more, dict) and more.get("status") == "pass" and more.get("parent") == CANDIDATE_PARENT, "0913 more fixture admission did not pass")
    more_added = more.get("added")
    require(isinstance(more_added, list) and len(more_added) == 2, "0913 more fixture count differs")
    more_rows = {row.get("path"): row for row in more_added if isinstance(row, dict)}
    require(set(more_rows) == {
        "docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/native-corpus.json",
        "3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx",
    }, "0913 more fixture paths differ")
    for name, row in more_rows.items():
        require(type(row.get("bytes")) is int and row["bytes"] >= 0, f"0913 more fixture byte count is invalid: {name}")
        digest = row.get("sha256")
        require(isinstance(digest, str) and len(digest) == 64 and all(char in "0123456789abcdef" for char in digest), f"0913 more fixture digest is invalid: {name}")
    git_name = "docs/report/spec-gap-validation-evidence/xlsx-form-control-properties/native-corpus.json"
    git_data = subprocess.check_output(["git", "show", f"{CANDIDATE_PARENT}:{git_name}"], cwd=ROOT)
    require((len(git_data), sha256_bytes(git_data)) == (more_rows[git_name]["bytes"], more_rows[git_name]["sha256"]), "0913 git more fixture bytes differ")
    local_name = "3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx"
    retained_raw = more_rows[local_name].get("retained")
    require(isinstance(retained_raw, str), "0913 local more fixture retention path is missing")
    retained_path = resolve_path(retained_raw)
    require(check_descriptor({"path": retained_raw, "bytes": more_rows[local_name]["bytes"], "sha256": more_rows[local_name]["sha256"]}, "0913 local more fixture") == descriptor(retained_path), "0913 local more fixture bytes differ")

    runtime_path = P / "runtime-fixtures.json"
    runtime = load(runtime_path, "0913 runtime-fixtures.json")
    require(isinstance(runtime, dict) and runtime.get("status") == "pass", "0913 runtime fixture admission did not pass")
    runtime_added = runtime.get("added")
    require(isinstance(runtime_added, list) and len(runtime_added) == 5, "0913 runtime fixture count differs")
    runtime_rows = {row.get("path"): row for row in runtime_added if isinstance(row, dict)}
    require(len(runtime_rows) == len(runtime_added), "0913 runtime fixture admission contains duplicate or malformed paths")
    require(set(runtime_rows) == set(EXPECTED_RUNTIME_FIXTURES), "0913 runtime fixture paths differ")
    for name, row in runtime_rows.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute() and ".." not in Path(name).parts, "0913 runtime fixture path is unsafe")
        require(type(row.get("bytes")) is int and row["bytes"] >= 0, f"0913 runtime fixture byte count is invalid: {name}")
        digest = row.get("sha256")
        require(isinstance(digest, str) and len(digest) == 64 and all(char in "0123456789abcdef" for char in digest), f"0913 runtime fixture digest is invalid: {name}")
        retained_raw = row.get("retained")
        require(isinstance(retained_raw, str) and retained_raw, f"0913 runtime fixture retention path is missing: {name}")
        require(retained_raw == EXPECTED_RUNTIME_FIXTURES[name], f"0913 runtime fixture retention path differs: {name}")
        retained_path = resolve_path(retained_raw)
        expected = {"path": retained_raw, "bytes": row["bytes"], "sha256": row["sha256"]}
        require(check_descriptor(expected, f"0913 runtime fixture {name}") == descriptor(retained_path), f"0913 runtime fixture bytes differ: {name}")
        source_path = resolve_path(name)
        if source_path.exists():
            require(descriptor(source_path) == {"path": name, "bytes": row["bytes"], "sha256": row["sha256"]}, f"0913 runtime source fixture bytes differ: {name}")

    removed_path = P / "fixtures-removed.json"
    removed = load(removed_path, "0913 fixtures-removed.json")
    require(removed.get("status") == "pass" and removed.get("removed") == len(added) and removed.get("extra_removed") == len(extra_added) and removed.get("more_removed") == len(more_added) and removed.get("runtime_removed") == len(runtime_added) and removed.get("benchmark_inventory_restored") is True, "0913 fixture removal receipt differs")
    return len(added) + len(extra_added) + len(more_added) + len(runtime_added)


def rerun_analysis(attempt: str) -> dict[str, Any]:
    analyzer = P / "analyze.py"
    expected_path = P / "analysis.json"
    require(analyzer.is_file() and expected_path.is_file(), "0913 analysis packet is incomplete")
    expected = expected_path.read_bytes()
    with tempfile.TemporaryDirectory(prefix="litchi-0913-analysis-replay-") as directory:
        output = Path(directory) / "analysis.json"
        subprocess.run(
            [sys.executable, "-B", str(analyzer), "--output", str(output)],
            cwd=ROOT,
            check=True,
            stdout=subprocess.DEVNULL,
        )
        actual = output.read_bytes()
        require(actual == expected, "0913 analysis replay differs byte-for-byte")
        parsed = json.loads(actual)
    require(parsed.get("status") == "qualified_candidate_control", "0913 analysis status differs")
    checks = parsed.get("checks")
    require(isinstance(checks, dict) and checks and all(value is True for value in checks.values()), "0913 analysis has a failed check")
    legs = parsed.get("legs")
    require(isinstance(legs, list) and len(legs) == 48, "0913 analysis does not retain all 48 legs")
    return {"bytes": len(expected), "sha256": sha256_bytes(expected), "legs": len(legs), "checks": len(checks)}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", default="v1")
    parser.add_argument("--output", type=Path, required=True, help="new replay result JSON")
    args = parser.parse_args()
    output = args.output.resolve()
    try:
        require(args.attempt.replace("-", "").isalnum(), "invalid attempt")
        require(not output.exists() and not output.is_symlink(), f"refusing to overwrite replay output: {output}")
        manifest_path = P / f"source-{args.attempt}.json"
        manifest = load(manifest_path, f"source-{args.attempt}.json")
        source = reconstruct_sources(manifest)
        control = baseline_binary()
        old_build_receipts = check_old_build_receipt(control)
        plan = load(P / "capture-plan.json", "0913 capture-plan.json")
        build = check_candidate_build(manifest, plan, control, args.attempt)
        capture = check_capture_packets(plan, build, args.attempt)
        quality = check_quality_receipts(manifest)
        census_runs = check_census()
        fixtures = check_fixtures()
        analysis = rerun_analysis(args.attempt)
        recovery = load(P / "source-recovery.json")
        require(recovery.get("status") == "pass" and recovery.get("baseline_files") == 9355 and recovery.get("candidate_files") == 9358 and recovery.get("quality_files") == 10803 and recovery.get("production_changes") is False, "source recovery scope differs")
        check_descriptor_tree(recovery, "source recovery")
        require(recovery["source"] == descriptor(manifest_path) and recovery["candidate_binary"] == build["candidate"], "source recovery identity differs")
        with tempfile.TemporaryDirectory(prefix="litchi-0913-stat-replay-") as directory:
            fresh = Path(directory) / "independent-stats.json"
            subprocess.run([sys.executable, "-B", str(P / "independent-stats.py"), "--output", str(fresh)], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
            require(fresh.read_bytes() == (P / "independent-stats.json").read_bytes(), "independent statistical replay differs")
            for helper_name, report_name in (("iwa-replay.py", "iwa-replay.json"), ("span-eligibility.py", "span-eligibility.json")):
                report = Path(directory) / report_name
                subprocess.run([sys.executable, "-B", str(P / helper_name), "--output", str(report)], cwd=ROOT, check=True, stdout=subprocess.DEVNULL)
                require(report.read_bytes() == (P / report_name).read_bytes(), f"{report_name}: offline replay differs")
        result = {
            "status": "pass",
            "source_file_count": source["candidate_files"],
            "baseline_source_file_count": source["baseline_files"],
            "source_manifest": f"source-{args.attempt}.json",
            "baseline_source_manifest": str((OLD / "source-v1.json").relative_to(ROOT)),
            "candidate_overlay_count": len(EXPECTED_OVERLAYS),
            "candidate_overlay_paths": list(EXPECTED_OVERLAYS),
            "old_build_receipts_checked": old_build_receipts,
            "candidate_build_receipts_checked": 1,
            "baseline_binary": control,
            "candidate_binary": build["candidate"],
            "capture_command_receipts_checked": capture["commands"],
            "capture_stdout_stderr_receipts_checked": capture["logs"],
            "capture_artifacts_checked": capture["artifacts"],
            "capture_process_metrics_checked": capture["process_metrics"],
            "quality_attempts_checked": quality["attempts"],
            "quality_command_receipts_checked": quality["commands"],
            "quality_failed_attempts_retained": quality["failed_attempts"],
            "quality_final_attempt": quality["final"],
            "census_runs_checked": census_runs,
            "census_reports_byte_exact": True,
            "temporary_fixture_files_checked": fixtures,
            "runtime_fixture_files_checked": 5,
            "source_recovery_checked": True,
            "independent_statistics_byte_exact": True,
            "iwa_commit_replay_byte_exact": True,
            "span_eligibility_replay_byte_exact": True,
            "analysis_byte_exact": True,
            "analysis": {
                "bytes": analysis["bytes"],
                "sha256": analysis["sha256"],
                "checks": analysis["checks"],
            },
            "deleted_source_root_used": False,
            "deleted_build_root_used": False,
            "build_or_workload_or_profiler_run": False,
        }
        output.parent.mkdir(parents=True, exist_ok=True)
        with output.open("x", encoding="utf-8") as stream:
            json.dump(result, stream, indent=2)
            stream.write("\n")
    except (ReplayError, OSError, subprocess.CalledProcessError, json.JSONDecodeError) as error:
        print(f"replay.py: error: {error}", file=sys.stderr)
        return 2
    print("0913 offline source/custody/analysis replay passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
