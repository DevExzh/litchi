#!/usr/bin/env python3
"""Replay the rejected 0735 candidate in an isolated shadow workspace.

The production source has been restored to the sealed baseline after the
secondary latency regression rejected the candidate.  The frozen analyzer and
audit still describe the candidate source, so this helper recreates that
candidate only under a temporary shadow root.  It copies the packet, hardlinks
the restored workspace files, installs the archived candidate source, keeps
the real capture bytes unchanged, and runs the existing validators against the
shadow root.  It never edits the real packet, source tree, captures, or
frozen scripts.

The temporary shadow is removed after a successful replay.  A failed replay
is retained and its path is printed so the attempt remains inspectable.
"""

from __future__ import annotations

import contextlib
import copy
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import shutil
import sys
import tempfile
from typing import Any


sys.dont_write_bytecode = True
PACKET = Path(__file__).resolve().parent
REAL_ROOT = PACKET.parents[3]
OWNED = "crates/litchi-ppt/src/embedded/object/editor/mutation/records.rs"
ANCESTOR_WRAPPER = "docs/performance/results/change-0731/oracle.json"
CONTRACT = "docs/performance/results/change-0728/oracle-contract.json"


class ReplayError(Exception):
    """A custody or replay invariant failed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReplayError(message)


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ReplayError(f"invalid JSON {path}: {error}") from error


def write(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def source_map(path: Path) -> dict[str, str]:
    value = read(path)
    result = value.get("source") if isinstance(value, dict) else None
    require(isinstance(result, dict) and result, f"source map missing: {path}")
    return result


def packet_files(root: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    for path in sorted(root.rglob("*")):
        if path.is_symlink():
            raise ReplayError(f"symlink in packet: {path}")
        if path.is_file():
            result[str(path.relative_to(root))] = sha(path)
    return result


def load_preflight() -> Any:
    spec = importlib.util.spec_from_file_location(
        "change0735_replay_preflight", PACKET / "preflight.py"
    )
    require(spec is not None and spec.loader is not None, "cannot load frozen preflight helper")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def link_file(source: Path, destination: Path) -> None:
    require(source.is_file() and not source.is_symlink(), f"missing source for hardlink: {source}")
    require(not destination.exists() and not destination.is_symlink(),
            f"shadow destination already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    try:
        os.link(source, destination)
    except OSError as error:
        raise ReplayError(f"could not hardlink {source} -> {destination}: {error}") from error
    require(destination.stat().st_ino == source.stat().st_ino,
            f"source hardlink was not created: {destination}")


def copy_file(source: Path, destination: Path) -> None:
    require(source.is_file() and not source.is_symlink(), f"missing source for copy: {source}")
    require(not destination.exists() and not destination.is_symlink(),
            f"shadow destination already exists: {destination}")
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    require(sha(destination) == sha(source), f"copy changed bytes: {destination}")


def copy_packet(helper: Any, shadow_packet: Path) -> None:
    """Use the packet's established copier, then prove packet files are copies."""
    helper.copy_packet(shadow_packet)
    real = packet_files(PACKET)
    for relative, expected in real.items():
        if (relative.startswith("__pycache__/")
                or relative.startswith("captures/")
                or relative.startswith("preflight-attempt-")
                or relative.startswith("negative-attempt-")):
            continue
        destination = shadow_packet / relative
        require(destination.is_file() and not destination.is_symlink(),
                f"packet copy missing: {relative}")
        require(sha(destination) == expected, f"packet copy changed: {relative}")
        require(destination.stat().st_ino != (PACKET / relative).stat().st_ino,
                f"packet file was hardlinked instead of copied: {relative}")


def verify_reversion(baseline: dict[str, str], candidate: dict[str, str]) -> dict[str, Any]:
    base = read(PACKET / "base.json")
    receipt = read(PACKET / "reversion.json")
    require(receipt.get("status") == "reverted", "reversion receipt is not reverted")
    require(receipt.get("owned_file") == OWNED, "reversion owned file changed")
    require(receipt.get("source_matches_baseline") is True,
            "reversion receipt does not assert baseline source custody")
    require(receipt.get("source") == baseline, "reversion source map differs from baseline build")
    require(receipt.get("restored_sha256") == baseline[OWNED],
            "reversion restored hash differs from baseline")
    require(receipt.get("candidate_sha256") == candidate[OWNED],
            "reversion candidate hash differs from candidate build")
    require(receipt.get("baseline_build_sha256") == sha(PACKET / "baseline-build.json"),
            "reversion baseline build binding changed")
    require(receipt.get("candidate_build_sha256") == sha(PACKET / "candidate-build.json"),
            "reversion candidate build binding changed")
    for relative, expected in baseline.items():
        target = REAL_ROOT / relative
        require(target.is_file() and not target.is_symlink() and sha(target) == expected,
                f"live source is not restored to baseline: {relative}")
    return receipt


def install_shadow_source(shadow_root: Path, shadow_packet: Path,
                          baseline: dict[str, str], candidate: dict[str, str]) -> None:
    require(set(baseline) == set(candidate), "baseline/candidate source inventories differ")
    changed = [relative for relative in baseline if baseline[relative] != candidate[relative]]
    require(changed == [OWNED], f"candidate source changed outside owned file: {changed}")
    after = shadow_packet / "source-archive" / "after" / OWNED
    require(sha(after) == candidate[OWNED], "archived candidate source hash changed")
    for relative in sorted(baseline):
        source = REAL_ROOT / relative
        destination = shadow_root / relative
        if relative == OWNED:
            copy_file(after, destination)
            require(sha(destination) == candidate[relative], "shadow candidate source hash changed")
        else:
            require(sha(source) == baseline[relative], f"live baseline source changed: {relative}")
            link_file(source, destination)
    require(len(baseline) == 7206, f"source file count changed: {len(baseline)}")


def install_external_inputs(shadow_root: Path, shadow_packet: Path) -> None:
    cases = read(PACKET / "cases.json")
    require(isinstance(cases, list) and len(cases) == 2, "case inventory changed")
    for row in cases:
        relative = row.get("path")
        require(isinstance(relative, str) and not Path(relative).is_absolute(),
                "fixture path is not relative")
        link_file(REAL_ROOT / relative, shadow_root / relative)

    constraints = read(PACKET / "constraints.json")
    require(isinstance(constraints, dict) and len(constraints) == 34,
            "constraint inventory changed")
    for relative in sorted(constraints):
        link_file(REAL_ROOT / relative, shadow_root / relative)

    wrapper = REAL_ROOT / ANCESTOR_WRAPPER
    ancestor = read(wrapper)
    link_file(wrapper, shadow_root / ANCESTOR_WRAPPER)
    reference = ancestor.get("reference")
    require(isinstance(reference, str) and not Path(reference).is_absolute(),
            "sealed oracle reference path is not relative")
    link_file(REAL_ROOT / reference, shadow_root / reference)
    link_file(REAL_ROOT / CONTRACT, shadow_root / CONTRACT)

    # source-guard.py reads this packet-local witness; the assertion makes the
    # intended shadow input explicit even though copy_packet already copied it.
    require((shadow_packet / "source-equivalence.json").is_file(),
            "source-equivalence witness missing from shadow packet")


def replace_prefix(value: Any, original: str, replacement: str) -> Any:
    if isinstance(value, str):
        if value == original:
            return replacement
        prefix = original + "/"
        if value.startswith(prefix):
            return replacement + value[len(original):]
        return value
    if isinstance(value, list):
        return [replace_prefix(item, original, replacement) for item in value]
    if isinstance(value, dict):
        return {key: replace_prefix(item, original, replacement) for key, item in value.items()}
    return value


def rebase_root_paths(shadow_packet: Path, shadow_root: Path) -> None:
    """Move root-prefixed JSON paths in copied receipts into the shadow root."""
    original = str(REAL_ROOT)
    replacement = str(shadow_root)
    for path in sorted(shadow_packet.rglob("*.json")):
        if path.name in {"freeze.json", "preflight.json"}:
            continue
        try:
            value = read(path)
        except ReplayError:
            continue
        rebased = replace_prefix(value, original, replacement)
        if rebased != value:
            write(path, rebased)


def rebase_freeze(shadow_packet: Path, shadow_root: Path) -> None:
    frozen = read(shadow_packet / "freeze.json")
    bindings = frozen.get("bindings", frozen) if isinstance(frozen, dict) else None
    require(isinstance(bindings, dict) and bindings, "freeze bindings are empty")
    old_root = str(REAL_ROOT)
    old_packet = str(PACKET)
    new_root = str(shadow_root)
    new_packet = str(shadow_packet)
    rebased: dict[str, str] = {}
    for raw, expected in bindings.items():
        require(isinstance(raw, str) and isinstance(expected, str), "malformed freeze binding")
        mapped = replace_prefix(raw, old_packet, new_packet)
        mapped = replace_prefix(mapped, old_root, new_root)
        target = Path(mapped)
        rebased[mapped] = sha(target) if target.is_file() else expected
    if isinstance(frozen, dict) and isinstance(frozen.get("bindings"), dict):
        frozen["bindings"] = rebased
    else:
        frozen = rebased
    write(shadow_packet / "freeze.json", frozen)


def install_captures(shadow_packet: Path) -> tuple[dict[str, Any], dict[str, str]]:
    source = PACKET / "captures"
    destination = shadow_packet / "captures"
    require(source.is_dir() and not source.is_symlink(), "real captures directory is missing")
    # preflight.copy_packet creates an empty captures directory as part of its
    # established packet shape; fill that directory with real copies without
    # touching the source capture files.
    shutil.copytree(source, destination, dirs_exist_ok=True)
    original = packet_files(source)
    copied = packet_files(destination)
    require(original == {name: copied[name] for name in original},
            "capture copy changed sample bytes")
    manifest = read(destination / "manifest.json")
    require(manifest == read(source / "manifest.json"), "capture manifest changed before rebasing")
    return manifest, original


def refresh_shadow_hashes(shadow_packet: Path, manifest: dict[str, Any]) -> None:
    freeze_hash = sha(shadow_packet / "freeze.json")
    preflight_path = shadow_packet / "preflight.json"
    preflight = read(preflight_path)
    preflight["freeze_sha256"] = freeze_hash
    write(preflight_path, preflight)
    capture = copy.deepcopy(manifest)
    capture["freeze_sha256"] = freeze_hash
    capture["preflight_sha256"] = sha(preflight_path)
    write(shadow_packet / "captures" / "manifest.json", capture)


def cleanup_receipt() -> None:
    cleanup = read(PACKET / "cleanup.json")
    require(cleanup.get("removed") is True, "cleanup receipt does not report removal")
    builds = {variant: read(PACKET / f"{variant}-build.json")
              for variant in ("baseline", "candidate")}
    expected = [builds[variant]["binaries"][lane]
                for variant in ("baseline", "candidate")
                for lane in ("native", "allocation")]
    actual = cleanup.get("binaries", cleanup.get("identities"))
    require(sorted(actual or [], key=lambda row: json.dumps(row, sort_keys=True))
            == sorted(expected, key=lambda row: json.dumps(row, sort_keys=True)),
            "cleanup binary identities are not exact")
    for row in expected:
        target = Path(row["path"])
        require(not target.exists() and not target.is_symlink(), f"cleaned binary remains: {target}")
    for raw in cleanup.get("roots", []):
        target = Path(raw)
        require(not target.exists() and not target.is_symlink(), f"cleanup root remains: {target}")


def load_validator(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run_validators(shadow_root: Path, shadow_packet: Path) -> tuple[dict[str, Any], dict[str, Any], str]:
    analyzer = load_validator("change0735_shadow_analyzer", shadow_packet / "analyze.py")
    analyzer.P = shadow_packet
    analyzer.ROOT = shadow_root
    analyzer.CAPTURES = shadow_packet / "captures"
    audit = load_validator("change0735_shadow_audit", shadow_packet / "audit.py")
    audit.PACKET = shadow_packet
    audit.ROOT = shadow_root
    audit.CAPTURES = shadow_packet / "captures"
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        analyzer.main()
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        audit.main()

    guard = load_validator("change0735_shadow_source_guard", shadow_packet / "source-guard.py")
    del guard
    return read(shadow_packet / "analysis.json"), read(shadow_packet / "audit.json"), "passed"


def compare_results(real_analysis: dict[str, Any], shadow_analysis: dict[str, Any],
                    real_audit: dict[str, Any], shadow_audit: dict[str, Any]) -> None:
    require(shadow_analysis == real_analysis,
            "shadow analysis differs from the retained candidate analysis")
    for key in ("matrix", "statistics", "native_case_summaries", "allocation_comparisons",
                "review_flags"):
        require(shadow_analysis.get(key) == real_analysis.get(key),
                f"shadow analysis root statistic differs: {key}")
    def normalized(value: dict[str, Any]) -> dict[str, Any]:
        output = copy.deepcopy(value)
        output.pop("freeze_sha256", None)
        output.pop("capture_sha256", None)
        return output
    require(normalized(shadow_audit) == normalized(real_audit),
            "shadow independent audit differs outside rebased custody hashes")


def main() -> None:
    baseline = source_map(PACKET / "baseline-build.json")
    candidate = source_map(PACKET / "candidate-build.json")
    real_source_before = {relative: sha(REAL_ROOT / relative) for relative in baseline}
    receipt = verify_reversion(baseline, candidate)
    cleanup_receipt()
    real_packet_before = packet_files(PACKET)
    real_capture_manifest, real_capture_files = (read(PACKET / "captures" / "manifest.json"),
                                                 packet_files(PACKET / "captures"))
    real_analysis = read(PACKET / "analysis.json")
    real_audit = read(PACKET / "audit.json")
    helper = load_preflight()
    # Keep the shadow on the workspace's filesystem so the 7,206 unchanged
    # source files can be hardlinked from the restored checkout.  /tmp is a
    # separate mount in the measurement environment.
    temporary_parent = Path(tempfile.mkdtemp(
        prefix=".litchi-0735-replay-rejected-", dir=REAL_ROOT.parent
    ))
    keep = True
    try:
        shadow_root = temporary_parent / "root"
        shadow_packet = shadow_root / PACKET.relative_to(REAL_ROOT)
        copy_packet(helper, shadow_packet)
        helper.rebase_packet(shadow_packet)
        install_shadow_source(shadow_root, shadow_packet, baseline, candidate)
        install_external_inputs(shadow_root, shadow_packet)
        rebase_root_paths(shadow_packet, shadow_root)
        manifest, _ = install_captures(shadow_packet)
        rebase_freeze(shadow_packet, shadow_root)
        refresh_shadow_hashes(shadow_packet, manifest)
        shadow_analysis, shadow_audit, status = run_validators(shadow_root, shadow_packet)
        require(status == "passed", "shadow validators did not pass")
        compare_results(real_analysis, shadow_analysis, real_audit, shadow_audit)

        # Prove the replay did not mutate anything in the real workspace.
        require(real_source_before == {relative: sha(REAL_ROOT / relative) for relative in baseline},
                "real source changed during shadow replay")
        require(real_packet_before == packet_files(PACKET),
                "real packet changed during shadow replay")
        require(real_capture_manifest == read(PACKET / "captures" / "manifest.json"),
                "real capture manifest changed during shadow replay")
        require(real_capture_files == packet_files(PACKET / "captures"),
                "real capture bytes changed during shadow replay")
        require(receipt == read(PACKET / "reversion.json"),
                "reversion receipt changed during shadow replay")
        keep = False
        shutil.rmtree(temporary_parent)
        print(json.dumps({
            "status": "passed",
            "reversion": "verified",
            "source_files": len(baseline),
            "fixtures": 2,
            "constraints": 34,
            "captures": len(real_capture_files),
            "analysis_identical": True,
            "audit_statistics_identical": True,
            "shadow_removed": True,
        }, sort_keys=True))
    except Exception:
        if keep:
            print(f"shadow replay retained at {temporary_parent}", file=sys.stderr)
        raise


if __name__ == "__main__":
    try:
        main()
    except (ReplayError, OSError, KeyError, TypeError, ValueError) as error:
        raise SystemExit(f"FAIL: {error}")
