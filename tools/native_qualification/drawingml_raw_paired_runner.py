#!/usr/bin/env python3
"""Run the source-fenced paired DrawingML raw qualification protocol.

The five raw families do not share one historical parent.  This coordinator
manifest therefore describes five profile pairs, each with its own immutable
baseline worktree and preparation receipt, and one assembled candidate.  Each
profile is materialized into the already-reviewed generic native runner, so
Cargo path validation, 20 GB target/address-space limits, process quietness,
fresh child processes, output handoff, and ABBA ordering remain shared with
the established W4b/W5 lanes.

This command is orchestration only.  The checked-in probe contains real
source-backed adapters; a source-pair preparation or endpoint failure is
fatal and cannot be converted into an acceptance claim.  The validator still
understands ``unavailable`` control reports for static protocol tests.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import sys
from typing import Any, Mapping


TOOLS_ROOT = Path(__file__).resolve().parents[1]
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

import native_qualification_runner as generic  # noqa: E402
from native_qualification import prepare_drawingml_raw_inputs as preparation  # noqa: E402


SCHEMA = "litchi.drawingml.raw-qualification.manifest.v1"
RECEIPT_SCHEMA = "litchi.drawingml.raw-qualification.receipt.v1"
PLAN_SHA256 = "d671c2dbd793d6b1b68c0aae25def384d6ceb44169be93b7daf848d6f13044f9"
LANES = (
    "drawingml-chart-raw",
    "drawingml-diagram-raw",
    "drawingml-text-media-raw",
    "drawingml-theme-ink",
    "drawingml-geometry-color-formatcode2-transform",
)
ABBA = ("baseline", "candidate", "candidate", "baseline")
ADAPTER_PATHS = tuple(preparation.ADAPTER_PATHS)
PROBE_PATHS = tuple(path for path in ADAPTER_PATHS if not path.endswith("Cargo.lock"))
TOKEN = re.compile(r"\$\{([A-Z][A-Z0-9_]*)\}")


class RunnerError(RuntimeError):
    """A fail-closed manifest, preparation, or child-runner error."""


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _expand(value: Any) -> Any:
    if isinstance(value, str):
        expanded = os.path.expandvars(value)
        if "$" in expanded:
            raise RunnerError(f"unexpanded environment variable in manifest value: {value!r}")
        return expanded
    if isinstance(value, list):
        return [_expand(item) for item in value]
    if isinstance(value, dict):
        return {key: _expand(item) for key, item in value.items()}
    return value


def _path(value: str, *, base: Path | None = None) -> Path:
    try:
        return generic._resolve_path(value, base=base)
    except generic.RunnerError as error:
        raise RunnerError(str(error)) from error


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _disjoint(paths: Mapping[str, Path]) -> None:
    values = list(paths.items())
    for index, (left_name, left) in enumerate(values):
        for right_name, right in values[index + 1 :]:
            if left == right or _under(left, right) or _under(right, left):
                raise RunnerError(f"paths must be disjoint: {left_name}={left}, {right_name}={right}")


def _load(path: Path) -> tuple[dict[str, Any], str]:
    try:
        raw = path.read_bytes()
        value = json.loads(raw)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"cannot read manifest {path}: {error}") from error
    if not isinstance(value, dict):
        raise RunnerError("manifest root must be an object")
    return _expand(value), _sha256_bytes(raw)


def _full_sha(value: Any, label: str) -> str:
    if not isinstance(value, str) or len(value) != 40 or value != value.lower() or any(char not in "0123456789abcdef" for char in value):
        raise RunnerError(f"{label} must be a lowercase 40-character commit SHA")
    return value


def _full_hash(value: Any, label: str) -> str:
    if not isinstance(value, str) or len(value) != 64 or value != value.lower() or any(char not in "0123456789abcdef" for char in value):
        raise RunnerError(f"{label} must be a lowercase 64-character SHA-256")
    return value


def _validate_master(data: Mapping[str, Any]) -> dict[str, Any]:
    if data.get("schema") != SCHEMA:
        raise RunnerError(f"manifest schema must be {SCHEMA!r}")
    if data.get("plan_sha256") != PLAN_SHA256:
        raise RunnerError("external DrawingML qualification plan hash differs")
    if tuple(profile.get("id") for profile in data.get("profiles", []) if isinstance(profile, dict)) != LANES:
        raise RunnerError("manifest profile order or count differs from the five reviewed raw lanes")
    candidate_root = _path(str(data.get("candidate_root", "")))
    candidate_revision = _full_sha(data.get("candidate_revision"), "candidate_revision")
    output = _path(str(data.get("output_dir", "")))
    target_root = _path(str(data.get("target_root", "")))
    if output.exists() and output.is_symlink():
        raise RunnerError("output_dir may not be a symlink")
    if target_root.exists() and target_root.is_symlink():
        raise RunnerError("target_root may not be a symlink")
    if not candidate_root.is_dir():
        raise RunnerError(f"candidate root is not a directory: {candidate_root}")
    if _under(output, candidate_root) or _under(target_root, candidate_root):
        raise RunnerError("output and target roots must be external to candidate source")
    fixture = _path("tools/native_qualification/fixtures/drawingml_raw_fixtures.json", base=candidate_root)
    if fixture.is_symlink() or not fixture.is_file():
        raise RunnerError(f"tracked DrawingML fixture descriptor is missing: {fixture}")
    fixture_hash, fixture_bytes = _sha256_file(fixture)
    if data.get("fixture_sha256") != fixture_hash or data.get("fixture_bytes") != fixture_bytes:
        raise RunnerError("manifest fixture identity differs from the tracked descriptor")
    profiles: list[dict[str, Any]] = []
    source_paths = {"candidate": candidate_root, "output": output, "target": target_root}
    for profile in data["profiles"]:
        if not isinstance(profile, dict):
            raise RunnerError("profile must be an object")
        lane = profile["id"]
        baseline_root = _path(str(profile.get("baseline_root", "")))
        baseline_revision = _full_sha(profile.get("baseline_revision"), f"{lane}.baseline_revision")
        perf_lock_sha256 = _full_hash(profile.get("perf_lock_sha256"), f"{lane}.perf_lock_sha256")
        receipt = _path(str(profile.get("preparation_receipt", "")))
        approved = profile.get("approved_baseline_revisions")
        if not isinstance(approved, list) or not approved:
            raise RunnerError(f"{lane} baseline revision is not one of its reviewed source parents")
        for index, revision in enumerate(approved):
            _full_sha(revision, f"{lane}.approved_baseline_revisions[{index}]")
        if baseline_revision not in approved:
            raise RunnerError(f"{lane} baseline revision is not one of its reviewed source parents")
        route_baselines = profile.get("route_baselines")
        if lane == "drawingml-theme-ink" and route_baselines is None:
            raise RunnerError("drawingml-theme-ink needs explicit theme and Ink route parents")
        if route_baselines is not None:
            if not isinstance(route_baselines, dict):
                raise RunnerError(f"{lane}.route_baselines must be an object")
            if lane == "drawingml-theme-ink" and set(route_baselines) != {"theme", "ink"}:
                raise RunnerError("drawingml-theme-ink route_baselines must name only theme and ink")
            for route, revisions in route_baselines.items():
                values = [revisions] if isinstance(revisions, str) else revisions
                if not isinstance(values, list) or not values:
                    raise RunnerError(f"{lane}.{route} has no exact baseline parent")
                for revision in values:
                    _full_sha(revision, f"{lane}.{route}.baseline")
            theme_parent = route_baselines.get("theme")
            if lane == "drawingml-theme-ink" and not isinstance(theme_parent, str):
                raise RunnerError("drawingml-theme-ink needs one exact theme parent; Ink routes are separate")
            if isinstance(theme_parent, str) and baseline_revision != theme_parent:
                raise RunnerError(f"{lane} baseline must be the exact theme parent; Ink routes need a separate endpoint")
        if not baseline_root.is_dir():
            raise RunnerError(f"{lane} baseline root is not a directory: {baseline_root}")
        if not receipt.is_file() or receipt.is_symlink():
            raise RunnerError(f"{lane} preparation receipt is missing: {receipt}")
        # Several lanes intentionally share one immutable parent worktree
        # (chart/diagram/geometry currently use the same owner parent).  That
        # is safe because each child target/output is separate; reject only
        # overlap with candidate/output/targets and receipt paths.
        for name, other in source_paths.items():
            # The same immutable parent may serve chart/diagram/geometry, but
            # a baseline must never contain or be contained by a receipt from
            # another profile.  Otherwise a later source preparation could
            # mutate a path that the manifest treats as an external receipt.
            if name.startswith("baseline-"):
                continue
            if baseline_root == other or _under(baseline_root, other) or _under(other, baseline_root):
                raise RunnerError(f"{lane} baseline overlaps {name}: {baseline_root}")
        for name, other in source_paths.items():
            if receipt == other or _under(receipt, other) or _under(other, receipt):
                raise RunnerError(f"{lane} preparation receipt overlaps {name}: {receipt}")
        source_paths[f"baseline-{lane}"] = baseline_root
        source_paths[f"receipt-{lane}"] = receipt
        for key in ("source_fence", "record"):
            if not isinstance(profile.get(key), str) or not profile[key]:
                raise RunnerError(f"{lane}.{key} is required")
        profiles.append({
            "id": lane,
            "baseline_root": baseline_root,
            "baseline_revision": baseline_revision,
            "perf_lock_sha256": perf_lock_sha256,
            "preparation_receipt": receipt,
            "source_fence": profile["source_fence"],
            "record": profile["record"],
            "approved_baseline_revisions": tuple(approved),
            "route_baselines": route_baselines,
        })
    # Baseline worktrees may be shared between profiles at the same parent;
    # all of them remain disjoint from the candidate, outputs, targets, and
    # external preparation receipts by the checks above.
    return {
        "schema": SCHEMA,
        "plan_sha256": PLAN_SHA256,
        "candidate_root": candidate_root,
        "candidate_revision": candidate_revision,
        "output_dir": output,
        "target_root": target_root,
        "fixture": fixture,
        "fixture_sha256": fixture_hash,
        "fixture_bytes": fixture_bytes,
        "profiles": profiles,
        "manifest_path": data.get("_manifest_path"),
    }


def _profile_manifest(master: Mapping[str, Any], profile: Mapping[str, Any], output: Path) -> dict[str, Any]:
    lane = str(profile["id"])
    before = Path(profile["baseline_root"])
    after = Path(master["candidate_root"])
    before_revision = str(profile["baseline_revision"])
    after_revision = str(master["candidate_revision"])
    lane_output = output / "lanes" / lane
    target_before = Path(master["target_root"]) / lane / "baseline"
    target_after = Path(master["target_root"]) / lane / "candidate"
    adapter_paths = list(PROBE_PATHS)
    build = [
        {
            "id": "drawingml_raw_native",
            "argv": ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin", "drawingml-raw-native"],
            "artifacts": ["{target_dir}/release/drawingml-raw-native"],
        },
        {
            "id": "drawingml_raw_allocator",
            "argv": ["cargo", "build", "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin", "drawingml-raw-allocator"],
            "artifacts": ["{target_dir}/release/drawingml-raw-allocator"],
        },
    ]

    def programs(role: str, leg: str) -> list[dict[str, Any]]:
        revision = before_revision if role == "baseline" else after_revision
        result = []
        for program_id, mode, binary in (
            ("raw-native", "normal", "drawingml-raw-native"),
            ("raw-allocator", "allocator", "drawingml-raw-allocator"),
        ):
            prefix = "native" if mode == "normal" else "allocator"
            output_name = f"{{output_dir}}/dml-{prefix}-{{leg}}-{{index}}.json"
            result.append({
                "id": program_id,
                "binary": f"{{target_dir}}/release/{binary}",
                "argv": [
                    "--lane", lane, "--role", role, "--mode", mode,
                    "--fixture", "{fixture}", "--source-revision", revision,
                    "--iterations", "1", "--warmups", "3", "--blocks", "5",
                    "--samples", "20", "--target-ms", "500", "--minimum-ms", "150",
                    "--json", "{output}",
                ],
                "outputs": [output_name],
                "stdout": output_name.removesuffix(".json") + ".stdout",
                "stderr": output_name.removesuffix(".json") + ".stderr",
                "require_stdout": False,
            })
        return result

    return {
        "schema": generic.SCHEMA,
        "lane": "drawingml-raw",
        "drawingml_profile": lane,
        "source_fence": str(profile["source_fence"]),
        "rounds": 31,
        "abba_order": list(generic.ABBA_ORDER),
        "output_dir": str(lane_output),
        "shared": {
            "probe_paths": adapter_paths,
            "fixture_paths": ["tools/native_qualification/fixtures/drawingml_raw_fixtures.json"],
            "preparation_receipt": str(profile["preparation_receipt"]),
        },
        "legs": {
            "before": {
                "source_root": str(before),
                "source_revision": before_revision,
                "target_dir": str(target_before),
                "lock_paths": ["Cargo.lock", "tools/perf-baseline/Cargo.lock"],
                "build": build,
                "programs": programs("baseline", "before"),
            },
            "after": {
                "source_root": str(after),
                "source_revision": after_revision,
                "target_dir": str(target_after),
                "lock_paths": ["Cargo.lock", "tools/perf-baseline/Cargo.lock"],
                "build": build,
                "programs": programs("candidate", "after"),
            },
        },
        "validators": [{
            "id": "drawingml-raw-identity",
            "argv": [
                "python3",
                "{source_root}/tools/native_qualification/drawingml_raw_validator.py",
                "--output-dir", "{output_dir}", "--lane", lane, "--fixture", "{fixture}",
                "--baseline-revision", before_revision,
                "--candidate-revision", after_revision,
                "--source-fence", str(profile["source_fence"]),
            ],
            "script": "{source_root}/tools/native_qualification/drawingml_raw_validator.py",
            "outputs": ["{output_dir}/summary.json"],
            "cwd": "{output_dir}",
            "require_stdout": False,
        }],
    }


def _prepare_profile(master: Mapping[str, Any], profile: Mapping[str, Any]) -> None:
    try:
        preparation.validate_preparation_receipt(
            Path(profile["preparation_receipt"]),
            baseline_root=Path(profile["baseline_root"]),
            candidate_root=Path(master["candidate_root"]),
            baseline_revision=str(profile["baseline_revision"]),
            candidate_revision=str(master["candidate_revision"]),
            perf_lock_sha256=str(profile["perf_lock_sha256"]),
        )
    except (OSError, ValueError) as error:
        raise RunnerError(f"{profile['id']} preparation receipt validation failed: {error}") from error


def _write_json(path: Path, value: Mapping[str, Any]) -> None:
    if path.exists() or path.is_symlink():
        raise RunnerError(f"refusing to overwrite protocol receipt: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def _run_profile(master: Mapping[str, Any], profile: Mapping[str, Any], phase: str, cpu: int | None, profile_path: Path) -> dict[str, Any]:
    data = _profile_manifest(master, profile, Path(master["output_dir"]))
    rendered, _ = _load(profile_path)
    if rendered != data:
        raise RunnerError(f"materialized {profile['id']} profile manifest changed")
    try:
        loaded = generic.load_manifest(profile_path)
        runner = generic.SafeRunner(loaded, profile_path, phase, cpu)
        if phase == "build":
            runner.build()
        else:
            runner.abba()
    except generic.RunnerError as error:
        raise RunnerError(f"{profile['id']} generic runner refused: {error}") from error
    receipt = Path(data["output_dir"]) / f"{phase}-receipt.json"
    if not receipt.is_file() or receipt.is_symlink():
        raise RunnerError(f"{profile['id']} did not produce its {phase} receipt")
    value, _ = _load(receipt)
    return value


def run(manifest_path: Path, phase: str, cpu: int | None) -> None:
    raw, manifest_hash = _load(manifest_path.resolve(strict=True))
    raw["_manifest_path"] = str(manifest_path.resolve())
    master = _validate_master(raw)
    output = Path(master["output_dir"])
    runner_path = Path(__file__).resolve()
    generic_runner_path = Path(generic.__file__).resolve()
    runner_sha256, runner_bytes = _sha256_file(runner_path)
    generic_runner_sha256, generic_runner_bytes = _sha256_file(generic_runner_path)
    top_receipt = output / f"{phase}-receipt.json"
    if phase == "build":
        if output.exists() and any(output.iterdir()):
            raise RunnerError(f"build output must be a fresh empty directory: {output}")
        output.mkdir(parents=True, exist_ok=True)
    else:
        build_receipt_path = output / "build-receipt.json"
        if not build_receipt_path.is_file() or build_receipt_path.is_symlink():
            raise RunnerError("ABBA requires the master build receipt")
        build_receipt, _ = _load(build_receipt_path)
        if build_receipt.get("schema") != RECEIPT_SCHEMA or build_receipt.get("phase") != "build" or build_receipt.get("status") != "passed":
            raise RunnerError("master build receipt is not a passed build handoff")
        if build_receipt.get("plan_sha256") != PLAN_SHA256:
            raise RunnerError("master build receipt plan hash differs")
        if build_receipt.get("manifest_sha256") != manifest_hash:
            raise RunnerError("master manifest differs from the build receipt")
        if build_receipt.get("runner") != {
            "path": str(runner_path),
            "sha256": runner_sha256,
            "bytes": runner_bytes,
            "generic_path": str(generic_runner_path),
            "generic_sha256": generic_runner_sha256,
            "generic_bytes": generic_runner_bytes,
        }:
            raise RunnerError("paired runner or generic guard changed since build")
        if top_receipt.exists() or top_receipt.is_symlink():
            raise RunnerError(f"ABBA receipt already exists: {top_receipt}")
    profile_receipts: dict[str, Any] = {}
    profile_files: list[dict[str, str]] = []
    try:
        for profile in master["profiles"]:
            _prepare_profile(master, profile)
            profile_path = output / "profiles" / f"{profile['id']}.json"
            if phase == "build":
                _write_json(profile_path, _profile_manifest(master, profile, output))
            else:
                if not profile_path.is_file() or profile_path.is_symlink():
                    raise RunnerError(f"{profile['id']} materialized profile is missing")
                profile_hash, _ = _sha256_file(profile_path)
                profile_files.append({
                    "lane": profile["id"],
                    "path": str(profile_path),
                    "sha256": profile_hash,
                    "baseline_revision": profile["baseline_revision"],
                    "perf_lock_sha256": profile["perf_lock_sha256"],
                    "source_fence": profile["source_fence"],
                })
            profile_receipts[profile["id"]] = _run_profile(master, profile, phase, cpu, profile_path)
        receipt = {
            "schema": RECEIPT_SCHEMA,
            "phase": phase,
            "status": "passed",
            "manifest": {"path": str(manifest_path.resolve()), "sha256": manifest_hash},
            "runner": {
                "path": str(runner_path),
                "sha256": runner_sha256,
                "bytes": runner_bytes,
                "generic_path": str(generic_runner_path),
                "generic_sha256": generic_runner_sha256,
                "generic_bytes": generic_runner_bytes,
            },
            "plan_sha256": PLAN_SHA256,
            "candidate_root": str(master["candidate_root"]),
            "candidate_revision": master["candidate_revision"],
            "abba_order": list(ABBA),
            "profiles": profile_files or [
                {
                    "lane": profile["id"],
                    "path": str(output / "profiles" / f"{profile['id']}.json"),
                    "sha256": _sha256_file(output / "profiles" / f"{profile['id']}.json")[0],
                    "baseline_revision": profile["baseline_revision"],
                    "perf_lock_sha256": profile["perf_lock_sha256"],
                    "source_fence": profile["source_fence"],
                }
                for profile in master["profiles"]
            ],
            "children": profile_receipts,
            "claim_scope": "paired raw DrawingML protocol only; no absolute-profile or whole-library claim",
        }
        if phase == "build":
            _write_json(output / "build-receipt.json", receipt)
        else:
            _write_json(output / "abba-receipt.json", receipt)
    except BaseException as error:
        failure = {
            "schema": RECEIPT_SCHEMA,
            "phase": phase,
            "status": "failed",
            "manifest": {"path": str(manifest_path.resolve()), "sha256": manifest_hash},
            "runner": {
                "path": str(runner_path),
                "sha256": runner_sha256,
                "bytes": runner_bytes,
                "generic_path": str(generic_runner_path),
                "generic_sha256": generic_runner_sha256,
                "generic_bytes": generic_runner_bytes,
            },
            "plan_sha256": PLAN_SHA256,
            "children": profile_receipts,
            "error": str(error),
        }
        if not top_receipt.exists() and not top_receipt.is_symlink():
            _write_json(top_receipt, failure)
        raise


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--phase", choices=("build", "abba"), required=True)
    parser.add_argument("--cpu", type=int)
    arguments = parser.parse_args(argv)
    try:
        run(arguments.manifest, arguments.phase, arguments.cpu)
    except (OSError, RunnerError, ValueError) as error:
        print(f"drawingml paired qualification refused: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
