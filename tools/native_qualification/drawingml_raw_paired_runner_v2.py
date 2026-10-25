#!/usr/bin/env python3
"""Run the current-source DrawingML raw qualification protocol (v2).

The v1 coordinator remains frozen because its receipts describe a different
consumer adapter.  This module reuses only the reviewed process/resource
orchestration from that coordinator; source preparation, receipt validation,
manifest identity, and the report validator are all explicitly v2.

No candidate hash is accepted from a caller.  The checked-in preparation
policy supplies the candidate binding and currently leaves it incomplete, so
this command fails closed until the coordinator records a reviewed source.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
from typing import Any, Mapping

from native_qualification import drawingml_raw_paired_runner as legacy
from native_qualification import prepare_drawingml_raw_inputs_v2 as preparation


# The generic runner and process observer are shared implementation, not
# source-policy identity.  Keeping this import one-way lets v1 receipts keep
# their original import and schema.
generic = legacy.generic
ExternalResourceObserver = legacy.ExternalResourceObserver
RunnerError = legacy.RunnerError
_argv_sha256 = legacy._argv_sha256
_full_hash = legacy._full_hash
_full_sha = legacy._full_sha
_load = legacy._load
_path = legacy._path
_sha256_file = legacy._sha256_file
_under = legacy._under
_write_json = legacy._write_json


SCHEMA = "litchi.drawingml.raw-qualification.manifest.v2"
RECEIPT_SCHEMA = "litchi.drawingml.raw-qualification.receipt.v2"
PLAN_PATH = legacy.PLAN_PATH
PLAN_SHA256 = legacy.PLAN_SHA256
LANES = legacy.LANES
ABBA = legacy.ABBA
OWNER_ROUTES = legacy.OWNER_ROUTES
RAW_SCOPES = legacy.RAW_SCOPES
DRAWINGML_FEATURES = legacy.DRAWINGML_FEATURES
DRAWINGML_PROBE_FEATURE = legacy.DRAWINGML_PROBE_FEATURE
ADAPTER_PATHS = tuple(preparation.PROBE_COPY_PATHS)
PROBE_PATHS = tuple(ADAPTER_PATHS)
ROLE_LOCK_PATHS = tuple(preparation.ROLE_LOCK_PATHS)


def _require_policy_binding(data: Mapping[str, Any]) -> dict[str, Any]:
    """Require both the checked-in binding and its manifest projection."""

    # Check the manifest state before resolving any caller-controlled path.
    # The checked-in policy is intentionally pending, and that diagnostic must
    # not be hidden behind an unrelated missing environment variable.
    if data.get("status") != "ready":
        raise RunnerError(
            "v2 manifest is pending final candidate binding; refusing to execute"
        )
    try:
        binding = preparation._require_candidate_binding()
    except (OSError, ValueError) as error:
        raise RunnerError(f"v2 candidate policy is not execution-ready: {error}") from error
    candidate = data.get("candidate")
    if not isinstance(candidate, dict):
        raise RunnerError("v2 manifest candidate policy is missing")
    if candidate.get("status") != "bound":
        raise RunnerError("v2 manifest candidate policy is not bound")
    if candidate.get("revision") != binding.get("revision"):
        raise RunnerError("v2 manifest candidate revision differs from checked-in binding")
    if candidate.get("manifests") != binding.get("manifests"):
        raise RunnerError("v2 manifest candidate manifests differ from checked-in binding")
    if candidate.get("vendor_tree") != binding.get("vendor_tree"):
        raise RunnerError("v2 manifest candidate vendor tree differs from checked-in binding")
    return binding


def _validate_master(data: Mapping[str, Any]) -> dict[str, Any]:
    binding = _require_policy_binding(data)
    if data.get("schema") != SCHEMA:
        raise RunnerError(f"manifest schema must be {SCHEMA!r}")
    if data.get("plan_path") != PLAN_PATH or data.get("plan_sha256") != PLAN_SHA256:
        raise RunnerError("external DrawingML qualification plan identity differs")
    profile_values = data.get("profiles")
    if not isinstance(profile_values, list):
        raise RunnerError("manifest profiles must be a list")
    if tuple(profile.get("id") for profile in profile_values if isinstance(profile, dict)) != LANES:
        raise RunnerError("manifest profile order or count differs from the five reviewed raw lanes")

    role_locks = data.get("role_locks")
    if role_locks != preparation.ROLE_LOCKS:
        raise RunnerError("v2 role-lock policy differs from the checked-in three-role policy")
    for prefix, role in (("baseline", "baseline"), ("candidate", "candidate")):
        for key, spec in preparation.ROLE_LOCKS.items():
            field = f"{prefix}_{key}_lock_sha256"
            expected = spec[role]
            if data.get(field) != expected["sha256"]:
                raise RunnerError(f"{field} differs from the checked-in role lock")
            bytes_field = f"{prefix}_{key}_lock_bytes"
            if data.get(bytes_field) != expected["bytes"]:
                raise RunnerError(f"{bytes_field} differs from the checked-in role lock")

    candidate_root = _path(str(data.get("candidate_root", "")))
    candidate_revision = _full_sha(data.get("candidate_revision"), "candidate_revision")
    if candidate_revision != binding["revision"]:
        raise RunnerError("candidate revision is not the checked-in v2 binding")
    trusted_root = _path(str(data.get("trusted_root", "")))
    output = _path(str(data.get("output_dir", "")))
    target_root = _path(str(data.get("target_root", "")))
    for label, path in (("candidate_root", candidate_root), ("trusted_root", trusted_root)):
        if not path.is_dir():
            raise RunnerError(f"{label} is not a directory: {path}")
    if output.exists() and output.is_symlink():
        raise RunnerError("output_dir may not be a symlink")
    if target_root.exists() and target_root.is_symlink():
        raise RunnerError("target_root may not be a symlink")
    _disjoint = {
        "candidate": candidate_root,
        "trusted": trusted_root,
        "output": output,
        "target": target_root,
    }
    names = list(_disjoint.items())
    for index, (left_name, left) in enumerate(names):
        for right_name, right in names[index + 1 :]:
            if left == right or _under(left, right) or _under(right, left):
                raise RunnerError(f"v2 source/output paths overlap: {left_name} and {right_name}")
    fixture = _path(
        "tools/native_qualification/fixtures/drawingml_raw_fixtures.json",
        base=candidate_root,
    )
    if fixture.is_symlink() or not fixture.is_file():
        raise RunnerError(f"tracked DrawingML fixture descriptor is missing: {fixture}")
    fixture_hash, fixture_bytes = _sha256_file(fixture)
    if data.get("fixture_sha256") != fixture_hash or data.get("fixture_bytes") != fixture_bytes:
        raise RunnerError("manifest fixture identity differs from the tracked descriptor")

    profiles: list[dict[str, Any]] = []
    profile_specs = list(profile_values)
    for key in ("route_profiles", "owner_profiles"):
        extra = data.get(key, [])
        if not isinstance(extra, list):
            raise RunnerError(f"{key} must be a list")
        profile_specs.extend(extra)
    profile_ids: set[str] = set()
    source_paths = {"candidate": candidate_root, "trusted": trusted_root, "output": output, "target": target_root}
    for profile in profile_specs:
        if not isinstance(profile, dict):
            raise RunnerError("profile must be an object")
        profile_id = profile.get("id")
        if not isinstance(profile_id, str) or not profile_id or profile_id in profile_ids:
            raise RunnerError(f"duplicate or invalid endpoint profile id: {profile_id!r}")
        profile_ids.add(profile_id)
        lane = profile.get("lane", profile_id)
        endpoint = profile.get("endpoint", "raw")
        owner_route = profile.get("owner_route")
        raw_scope = profile.get("raw_scope")
        if lane not in LANES:
            raise RunnerError(f"{profile_id} lane is not one of the reviewed raw lanes")
        if endpoint not in {"raw", "owner"}:
            raise RunnerError(f"{profile_id} endpoint is not raw or owner")
        if endpoint == "owner":
            if owner_route not in OWNER_ROUTES or OWNER_ROUTES[owner_route] != lane:
                raise RunnerError(f"{profile_id} owner route is not reviewed for its lane")
            if raw_scope is not None:
                raise RunnerError(f"{profile_id} owner endpoint carries a raw scope")
        elif owner_route is not None:
            raise RunnerError(f"{profile_id} raw endpoint carries an owner route")
        if raw_scope is not None and raw_scope not in RAW_SCOPES:
            raise RunnerError(f"{profile_id} has an unknown raw scope")
        for label in ("baseline_features", "candidate_features"):
            features = profile.get(label, ["drawingml-owner-current"] if label == "candidate_features" else [])
            if not isinstance(features, list) or any(not isinstance(item, str) or not item for item in features):
                raise RunnerError(f"{profile_id}.{label} must be a list of non-empty feature names")
            if set(features) - DRAWINGML_FEATURES:
                raise RunnerError(f"{profile_id}.{label} has an unknown DrawingML feature")
        baseline_root = _path(str(profile.get("baseline_root", "")))
        baseline_revision = _full_sha(profile.get("baseline_revision"), f"{profile_id}.baseline_revision")
        if baseline_revision not in preparation.APPROVED_BASELINE_REVISIONS:
            raise RunnerError(f"{profile_id} baseline revision is not approved")
        if profile.get("perf_lock_sha256") != preparation.ROLE_LOCKS["perf"]["baseline"]["sha256"]:
            raise RunnerError(f"{profile_id} perf lock differs from the v2 baseline role lock")
        receipt = _path(str(profile.get("preparation_receipt", "")))
        if not baseline_root.is_dir():
            raise RunnerError(f"{profile_id} baseline root is not a directory: {baseline_root}")
        if not receipt.is_file() or receipt.is_symlink():
            raise RunnerError(f"{profile_id} preparation receipt is missing: {receipt}")
        for name, other in source_paths.items():
            if baseline_root == other or _under(baseline_root, other) or _under(other, baseline_root):
                raise RunnerError(f"{profile_id} baseline overlaps {name}")
            if receipt == other or _under(receipt, other) or _under(other, receipt):
                raise RunnerError(f"{profile_id} preparation receipt overlaps {name}")
        source_paths[f"baseline-{profile_id}"] = baseline_root
        source_paths[f"receipt-{profile_id}"] = receipt
        for key in ("source_fence", "record"):
            if not isinstance(profile.get(key), str) or not profile[key]:
                raise RunnerError(f"{profile_id}.{key} is required")
        profiles.append({
            "id": profile_id,
            "lane": lane,
            "endpoint": endpoint,
            "owner_route": owner_route,
            "raw_scope": raw_scope,
            "baseline_root": baseline_root,
            "baseline_revision": baseline_revision,
            "perf_lock_sha256": profile["perf_lock_sha256"],
            "preparation_receipt": receipt,
            "source_fence": profile["source_fence"],
            "record": profile["record"],
            "approved_baseline_revisions": tuple(profile.get("approved_baseline_revisions", ())),
            "route_baselines": profile.get("route_baselines"),
            "baseline_features": tuple(profile.get("baseline_features", ())),
            "candidate_features": tuple(profile.get("candidate_features", ("drawingml-owner-current",))),
        })
    return {
        "schema": SCHEMA,
        "plan_path": PLAN_PATH,
        "plan_sha256": PLAN_SHA256,
        "candidate_root": candidate_root,
        "candidate_revision": candidate_revision,
        "trusted_root": trusted_root,
        "baseline_workspace_lock_sha256": data["baseline_workspace_lock_sha256"],
        "baseline_workspace_lock_bytes": data["baseline_workspace_lock_bytes"],
        "candidate_workspace_lock_sha256": data["candidate_workspace_lock_sha256"],
        "candidate_workspace_lock_bytes": data["candidate_workspace_lock_bytes"],
        "output_dir": output,
        "target_root": target_root,
        "fixture": fixture,
        "fixture_sha256": fixture_hash,
        "fixture_bytes": fixture_bytes,
        "profiles": profiles,
        "manifest_path": data.get("_manifest_path"),
    }


def _profile_manifest(master: Mapping[str, Any], profile: Mapping[str, Any], output: Path) -> dict[str, Any]:
    """Materialize one generic v1 process manifest under the v2 source fence."""

    value = legacy._profile_manifest(master, profile, output)
    value["drawingml_protocol_version"] = 2
    value["drawingml_trusted_root"] = str(master["trusted_root"])
    value["shared"]["probe_paths"] = list(PROBE_PATHS)
    value["shared"]["preparation_receipt"] = str(profile["preparation_receipt"])
    for leg in ("before", "after"):
        value["legs"][leg]["allowed_dirty_paths"] = list(ADAPTER_PATHS)
        value["legs"][leg]["lock_paths"] = list(ROLE_LOCK_PATHS)
    validator = value["validators"][0]
    script = "{source_root}/tools/native_qualification/drawingml_raw_validator_v2.py"
    validator["argv"][1] = script
    validator["script"] = script
    validator["argv"].extend([
        "--preparation-receipt",
        str(profile["preparation_receipt"]),
        "--baseline-root",
        str(profile["baseline_root"]),
        "--candidate-root",
        str(master["candidate_root"]),
        "--trusted-root",
        str(master["trusted_root"]),
    ])
    return value


def _prepare_profile(master: Mapping[str, Any], profile: Mapping[str, Any]) -> None:
    try:
        preparation.validate_preparation_receipt(
            Path(profile["preparation_receipt"]),
            baseline_root=Path(profile["baseline_root"]),
            candidate_root=Path(master["candidate_root"]),
            baseline_revision=str(profile["baseline_revision"]),
            candidate_revision=str(master["candidate_revision"]),
            trusted_root=Path(master["trusted_root"]),
        )
    except (OSError, ValueError) as error:
        raise RunnerError(f"{profile['id']} v2 preparation receipt validation failed: {error}") from error


def _run_profile(master: Mapping[str, Any], profile: Mapping[str, Any], phase: str, cpu: int | None, profile_path: Path) -> dict[str, Any]:
    data = _profile_manifest(master, profile, Path(master["output_dir"]))
    rendered, _ = _load(profile_path)
    if rendered != data:
        raise RunnerError(f"materialized {profile['id']} v2 profile manifest changed")
    observer = ExternalResourceObserver(Path(data["output_dir"])) if phase == "abba" else None
    try:
        loaded = generic.load_manifest(profile_path)
        runner = generic.SafeRunner(loaded, profile_path, phase, cpu, process_observer=observer)
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
    if observer is not None:
        value["external_resource_capture"] = observer.receipt()
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
            raise RunnerError("v2 ABBA requires the master build receipt")
        build_receipt, _ = _load(build_receipt_path)
        if build_receipt.get("schema") != RECEIPT_SCHEMA or build_receipt.get("phase") != "build" or build_receipt.get("status") != "passed":
            raise RunnerError("v2 master build receipt is not a passed build handoff")
        if build_receipt.get("manifest", {}).get("sha256") != manifest_hash:
            raise RunnerError("v2 manifest differs from the build receipt")
        if build_receipt.get("runner") != {
            "path": str(runner_path),
            "sha256": runner_sha256,
            "bytes": runner_bytes,
            "generic_path": str(generic_runner_path),
            "generic_sha256": generic_runner_sha256,
            "generic_bytes": generic_runner_bytes,
        }:
            raise RunnerError("v2 paired runner or generic guard changed since build")
        if top_receipt.exists() or top_receipt.is_symlink():
            raise RunnerError(f"v2 ABBA receipt already exists: {top_receipt}")
    profile_receipts: dict[str, Any] = {}
    profile_files: list[dict[str, Any]] = []
    try:
        for profile in master["profiles"]:
            _prepare_profile(master, profile)
            profile_path = output / "profiles" / f"{profile['id']}.json"
            if phase == "build":
                _write_json(profile_path, _profile_manifest(master, profile, output))
            else:
                if not profile_path.is_file() or profile_path.is_symlink():
                    raise RunnerError(f"{profile['id']} materialized v2 profile manifest is missing")
                profile_hash, _ = _sha256_file(profile_path)
                profile_files.append({
                    "profile": profile["id"],
                    "lane": profile["lane"],
                    "endpoint": profile["endpoint"],
                    "owner_route": profile["owner_route"],
                    "raw_scope": profile["raw_scope"],
                    "path": str(profile_path),
                    "sha256": profile_hash,
                    "baseline_revision": profile["baseline_revision"],
                    "source_fence": profile["source_fence"],
                })
            profile_receipts[profile["id"]] = _run_profile(master, profile, phase, cpu, profile_path)
        receipt = {
            "schema": RECEIPT_SCHEMA,
            "protocol_version": 2,
            "phase": phase,
            "status": "passed",
            "plan_path": PLAN_PATH,
            "plan_sha256": PLAN_SHA256,
            "manifest": {"path": str(manifest_path.resolve()), "sha256": manifest_hash},
            "runner": {
                "path": str(runner_path), "sha256": runner_sha256, "bytes": runner_bytes,
                "generic_path": str(generic_runner_path), "generic_sha256": generic_runner_sha256,
                "generic_bytes": generic_runner_bytes,
            },
            "candidate_root": str(master["candidate_root"]),
            "candidate_revision": master["candidate_revision"],
            "trusted_root": str(master["trusted_root"]),
            "role_locks": preparation.ROLE_LOCKS,
            "probe_copy_paths": list(ADAPTER_PATHS),
            "abba_order": list(ABBA),
            "profiles": profile_files or [
                {
                    "profile": profile["id"], "lane": profile["lane"], "endpoint": profile["endpoint"],
                    "owner_route": profile["owner_route"], "raw_scope": profile["raw_scope"],
                    "path": str(output / "profiles" / f"{profile['id']}.json"),
                    "sha256": _sha256_file(output / "profiles" / f"{profile['id']}.json")[0],
                    "baseline_revision": profile["baseline_revision"], "source_fence": profile["source_fence"],
                }
                for profile in master["profiles"]
            ],
            "children": profile_receipts,
            "claim_scope": "paired DrawingML raw and detached-owner endpoints only; no absolute-profile or whole-library claim",
        }
        _write_json(output / ("build-receipt.json" if phase == "build" else "abba-receipt.json"), receipt)
    except BaseException as error:
        failure = {
            "schema": RECEIPT_SCHEMA,
            "protocol_version": 2,
            "phase": phase,
            "status": "failed",
            "plan_path": PLAN_PATH,
            "plan_sha256": PLAN_SHA256,
            "manifest": {"path": str(manifest_path.resolve()), "sha256": manifest_hash},
            "runner": {
                "path": str(runner_path), "sha256": runner_sha256, "bytes": runner_bytes,
                "generic_path": str(generic_runner_path), "generic_sha256": generic_runner_sha256,
                "generic_bytes": generic_runner_bytes,
            },
            "candidate_root": str(master["candidate_root"]),
            "candidate_revision": master["candidate_revision"],
            "trusted_root": str(master["trusted_root"]),
            "role_locks": preparation.ROLE_LOCKS,
            "children": profile_receipts,
            "error": str(error),
        }
        if not top_receipt.exists() and not top_receipt.is_symlink():
            _write_json(top_receipt, failure)
        raise


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--phase", choices=("build", "abba"), required=True)
    parser.add_argument("--cpu", type=int)
    arguments = parser.parse_args(argv)
    try:
        run(arguments.manifest, arguments.phase, arguments.cpu)
    except (OSError, RunnerError, ValueError) as error:
        print(f"drawingml v2 paired qualification refused: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
