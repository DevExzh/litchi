#!/usr/bin/env python3
"""Bind the PPT build receipt through its canonical input projection.

The runner requires a receipt to hash the complete qualification manifest
while the manifest requires the receipt hash.  This adapter removes only
``build.receipt.bytes`` and ``build.receipt.sha256`` from a canonical
projection, keeps the receipt path and every other input, and asks the public
runner lifecycle to validate that projection.  It independently pins the
historical runner, checks the reviewed baseline overlay status/content and the
accepted 472b preflight/build-log identities.  It never builds or invokes a
probe.
"""

from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import subprocess
import sys
from typing import Any, Mapping, Sequence


SCHEMA = "litchi.ppt.lazy-native.build-adapter.v1"
RUNNER_RELATIVE = Path("tools/ppt_lazy_native_qualification.py")
FROZEN_RUNNER_BYTES = 73_596
FROZEN_RUNNER_SHA256 = "19d80d7f6e9450b60862929c4d864c0caae213aaa9e0682fca1c5fcc67d28c4f"
# Updated after the public lifecycle integration is committed.  Keeping this
# identity fixed prevents an adapter from silently loading a different runner.
PUBLIC_RUNNER_BYTES = 86_075
PUBLIC_RUNNER_SHA256 = "88f10be52bf02a8d5210a831d38af4358738ea35aea956341d7d4e8af1342977"
PREFLIGHT_BYTES = 7_646
PREFLIGHT_SHA256 = "fbe59be89acb3a588b659c217f05ddf3c6ac49cdbeee2dce03c821e469714db8"
NORMAL_LOG = (39_797, "ca7297042ecfc6cd212126ba8fe92eff78cd023cd4d4f5d658ebb0cc72fdd223")
ALLOCATOR_LOG = (39_797, "5bca2dad0e894e48f8a55d7263f8d6f99147c6b686fd0f4e8f1b351e1754cea0")
OVERLAY_STATUS = (
    "M tools/perf-baseline/src/lib.rs\n"
    "?? test-data/unified-ops/ppt-lazy-native-fixture.ppt\n"
    "?? tools/perf-baseline/src/bin/ppt_lazy_native.rs\n"
    "?? tools/perf-baseline/src/bin/ppt_lazy_native_alloc.rs\n"
    "?? tools/perf-baseline/src/ppt_lazy_native.rs"
)


class AdapterError(RuntimeError):
    """The source-only build packet is not the reviewed pair."""


def identity(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if path.is_symlink() or not path.is_file():
        raise AdapterError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
            size += len(block)
    return {"path": str(path), "bytes": size, "sha256": digest.hexdigest()}


def strict_json(path: Path, label: str) -> Any:
    raw = path.read_bytes()

    def unique(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise AdapterError(f"{label} contains duplicate key {key!r}")
            result[key] = value
        return result

    def reject(value: str) -> Any:
        raise AdapterError(f"{label} contains non-finite JSON constant {value!r}")

    try:
        value = json.loads(raw.decode("utf-8"), object_pairs_hook=unique, parse_constant=reject)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise AdapterError(f"{label} is not strict JSON: {error}") from error
    def finite(item: Any) -> None:
        if isinstance(item, float) and not math.isfinite(item):
            raise AdapterError(f"{label} contains a non-finite JSON number")
        if isinstance(item, dict):
            for child in item.values():
                finite(child)
        elif isinstance(item, list):
            for child in item:
                finite(child)
    finite(value)
    return value


def canonical(value: Mapping[str, Any]) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def projection_value(manifest: Mapping[str, Any]) -> dict[str, Any]:
    value = copy.deepcopy(dict(manifest))
    try:
        receipt = value["build"]["receipt"]
    except (KeyError, TypeError) as error:
        raise AdapterError("manifest.build.receipt is missing") from error
    if set(receipt) != {"path", "bytes", "sha256"}:
        raise AdapterError("manifest.build.receipt identity fields differ")
    del receipt["bytes"]
    del receipt["sha256"]
    return value


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), "--no-optional-locks", *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        timeout=30,
    )
    if result.returncode:
        raise AdapterError(f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}")
    return result.stdout.strip()


def _runner() -> tuple[Any, dict[str, Any]]:
    path = Path(__file__).with_name("ppt_lazy_native_qualification.py").resolve()
    tools = path.parent
    sys.path.insert(0, str(tools))
    actual = identity(path)
    if actual["bytes"] != PUBLIC_RUNNER_BYTES or actual["sha256"] != PUBLIC_RUNNER_SHA256:
        raise AdapterError("public PPT qualification runner identity differs")
    spec = importlib.util.spec_from_file_location("ppt_lazy_native_qualification_adapter", path)
    if spec is None or spec.loader is None:
        raise AdapterError("cannot load public PPT qualification runner")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module, actual


def _frozen_runner(candidate_root: Path) -> dict[str, Any]:
    path = (candidate_root / RUNNER_RELATIVE).resolve()
    actual = identity(path)
    if actual["bytes"] != FROZEN_RUNNER_BYTES or actual["sha256"] != FROZEN_RUNNER_SHA256:
        raise AdapterError("frozen PPT qualification runner identity differs")
    return actual


def _check_expected(actual: Mapping[str, Any], expected: tuple[int, str], label: str) -> dict[str, Any]:
    if actual["bytes"] != expected[0] or actual["sha256"] != expected[1]:
        raise AdapterError(f"{label} identity differs")
    return dict(actual)


def _check_baseline_overlay(runner: Any, manifest: Mapping[str, Any], receipt: Mapping[str, Any], prep_path: Path) -> dict[str, Any]:
    baseline = manifest["baseline"]
    root = Path(baseline["root"]).resolve()
    if _git(root, "rev-parse", "HEAD") != baseline["ref"] or _git(root, "rev-parse", "HEAD^{tree}") != baseline["tree"]:
        raise AdapterError("baseline HEAD/tree differs from the reviewed pre-lazy parent")
    status = _git(root, "status", "--porcelain", "--untracked-files=all")
    if status != OVERLAY_STATUS:
        raise AdapterError("baseline overlay changed paths outside the reviewed closure")
    prep = strict_json(prep_path, "PPT preparation receipt")
    after = prep.get("baseline", {}).get("after_overlay", {})
    if after.get("status") != OVERLAY_STATUS:
        raise AdapterError("preparation receipt does not bind the exact baseline overlay status")
    expected_overlay_paths = sorted(
        [
            manifest["baseline_adapter"]["fixture"]["path"],
            manifest["baseline_adapter"]["registration"]["path"],
            *manifest["baseline_adapter"]["overlay_paths"],
            *[item["path"] for item in manifest["baseline_adapter"]["compile_fixtures"]],
        ]
    )
    if prep.get("baseline", {}).get("overlay_paths") != expected_overlay_paths:
        raise AdapterError("preparation receipt overlay closure differs from the reviewed paths")
    registration = receipt["baseline_adapter"]["registration"]
    actual_registration = runner._identity(root / registration["path"])
    expected_registration = {
        "path": registration["before"]["path"],
        "bytes": registration["before"]["bytes"],
        "sha256": registration["before"]["sha256"],
    }
    if actual_registration != {"path": str(root / registration["path"]), "bytes": expected_registration["bytes"], "sha256": expected_registration["sha256"]}:
        raise AdapterError("baseline registration overlay bytes differ")
    fixture = receipt["baseline_adapter"]["fixture"]
    actual_fixture = runner._identity(root / fixture["path"])
    if actual_fixture != {"path": str(root / fixture["path"]), "bytes": fixture["bytes"], "sha256": fixture["sha256"]}:
        raise AdapterError("baseline fixture overlay bytes differ")
    for item in receipt["baseline_adapter"]["overlay_paths"]:
        relative = item["path"]
        actual = runner._identity(root / relative)
        candidate = runner._identity(Path(manifest["candidate"]["root"]) / relative)
        if actual != {"path": str(root / relative), "bytes": item["bytes"], "sha256": item["sha256"]}:
            raise AdapterError(f"baseline overlay bytes differ for {relative}")
        if candidate["bytes"] != item["bytes"] or candidate["sha256"] != item["sha256"]:
            raise AdapterError(f"baseline overlay bytes differ for {relative}")
    for item in receipt["baseline_adapter"]["compile_fixtures"]:
        relative = item["path"]
        actual = runner._identity(root / relative)
        candidate = runner._identity(Path(manifest["candidate"]["root"]) / relative)
        expected = {"path": str(root / relative), "bytes": item["bytes"], "sha256": item["sha256"]}
        if actual != expected or candidate["bytes"] != item["bytes"] or candidate["sha256"] != item["sha256"]:
            raise AdapterError(f"baseline compile fixture bytes differ for {relative}")
    candidate_fixture = runner._identity(Path(manifest["candidate"]["root"]) / fixture["path"])
    if candidate_fixture["bytes"] != fixture["bytes"] or candidate_fixture["sha256"] != fixture["sha256"]:
        raise AdapterError("candidate fixture bytes differ")
    return {"root": str(root), "status": status, "overlay_paths": expected_overlay_paths}


def validate(manifest_path: Path, projection_path: Path, preflight_path: Path, logs_root: Path) -> dict[str, Any]:
    manifest_path = manifest_path.resolve()
    projection_path = projection_path.resolve()
    manifest = strict_json(manifest_path, "final PPT manifest")
    if not isinstance(manifest, dict):
        raise AdapterError("final PPT manifest must be an object")
    projection = strict_json(projection_path, "PPT build-input projection")
    if canonical(projection) != canonical(projection_value(manifest)):
        raise AdapterError("build-input projection differs outside receipt identity fields")
    projection_identity = identity(projection_path)
    candidate_root = Path(manifest["candidate"]["root"]).resolve()
    runner, public_runner_identity = _runner()
    frozen_runner_identity = _frozen_runner(candidate_root)
    lifecycle = runner.validate_inputs(manifest_path, projection_path)
    if lifecycle["build_input_projection"] != projection_identity:
        raise AdapterError("public runner returned a different canonical projection identity")
    loaded = strict_json(manifest_path, "final PPT manifest")
    receipt = strict_json(Path(loaded["build"]["receipt"]["path"]), "PPT build receipt")
    build_result = lifecycle["build_receipt"]
    prep_path = Path(build_result["preparation"]["path"])
    overlay = _check_baseline_overlay(runner, loaded, receipt, prep_path)
    preflight_identity = _check_expected(identity(preflight_path), (PREFLIGHT_BYTES, PREFLIGHT_SHA256), "corrected qualification receipt")
    logs = {
        "candidate_normal": _check_expected(identity(logs_root / "candidate-472b-normal-build.log"), NORMAL_LOG, "normal build log"),
        "candidate_allocator": _check_expected(identity(logs_root / "candidate-472b-allocator-build.log"), ALLOCATOR_LOG, "allocator build log"),
    }
    return {
        "schema": SCHEMA,
        "status": "structural_pass",
        "manifest": identity(manifest_path),
        "build_input_projection": projection_identity,
        "build_receipt": {"path": build_result["path"], "bytes": build_result["bytes"], "sha256": build_result["sha256"]},
        "qualification_receipt": preflight_identity,
        "public_runner": public_runner_identity,
        "frozen_runner": frozen_runner_identity,
        "preparation": build_result["preparation"],
        "overlay": overlay,
        "build_logs": logs,
        "source_and_binary_validation": "public runner validate_inputs lifecycle plus exact baseline overlay path/content checks",
        "claim_scope": "source-only structural binding of existing builds and correctness/allocation preflight; no rebuild or native timing claim",
        "known_frozen_runner_boundary": "the public runner receives the full manifest and its explicit canonical build-input projection; the frozen historical runner identity is independently pinned",
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--projection", type=Path, required=True)
    parser.add_argument("--preflight-receipt", type=Path, required=True)
    parser.add_argument("--logs-root", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        result = validate(args.manifest, args.projection, args.preflight_receipt, args.logs_root)
        if args.output.exists() or args.output.is_symlink():
            raise AdapterError(f"refusing to overwrite output: {args.output}")
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_bytes(canonical(result))
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (OSError, AdapterError, subprocess.TimeoutExpired) as error:
        print(f"PPT build receipt adapter failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
