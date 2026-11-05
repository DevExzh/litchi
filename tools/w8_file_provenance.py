#!/usr/bin/env python3
"""Create fail-closed provenance inputs for the W8 File qualification.

The native runner in ``docs/unified-ops/results/0362h`` deliberately accepts
only already-built binaries.  This tool produces the small, deterministic
inputs that let its endpoint manifests prove that the two binaries used the
same workload and dependency resolution:

* ``fixture`` scans the checked-in corrected probe, records its fixture recipe
  and route set, and writes a descriptor whose identity is the probe source
  hash plus the parsed recipe;
* ``dependency`` projects the supplied lockfile, package manifests, feature
  selection, target and toolchain into a canonical JSON object; and
* ``build-receipt`` hashes the completed binary, compiler and cargo shim and
  writes the complete post-build receipt consumed by ``endpoint``; and
* ``endpoint`` writes the runner's endpoint manifest from real artifact
  hashes.

No Cargo command is run.  Generated descriptors and manifests are intended
for an external scratch directory; the tool rejects the repository's large
evidence and third-party trees as output or input locations.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import sys
import tomllib
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
SHA256 = re.compile(r"^[0-9a-f]{64}$")
FULL_COMMIT = re.compile(r"^[0-9a-f]{40}$")
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


class ProvenanceError(RuntimeError):
    """An input is missing, stale, ambiguous or not bound to the probe."""


def fail(message: str) -> ProvenanceError:
    return ProvenanceError(message)


def digest(path: Path) -> str:
    value = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                value.update(block)
    except OSError as error:
        raise fail(f"cannot read {path}: {error}") from error
    return value.hexdigest()


def artifact(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if not path.is_file():
        raise fail(f"artifact is not a regular file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": digest(path)}


def same_artifact(actual: dict[str, Any], expected: object, label: str) -> None:
    if not isinstance(expected, dict):
        raise fail(f"{label} artifact is missing")
    for key in ("bytes", "sha256"):
        if expected.get(key) != actual[key]:
            raise fail(f"{label} {key} does not match its supplied file")


def checked_commit(value: object, label: str) -> str:
    if not isinstance(value, str) or not FULL_COMMIT.fullmatch(value):
        raise fail(f"{label} must be a full 40-character commit id")
    return value


def canonical(value: object) -> bytes:
    return (json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")) + "\n").encode()


def safe_location(path: Path, label: str, *, output: bool = False) -> Path:
    path = path.resolve()
    pieces = set(path.parts)
    upstream_tree = "3rd" + "party"
    if upstream_tree in pieces:
        raise fail(f"{label} may not use an upstream checkout path: {path}")
    if output:
        for forbidden in (ROOT / "docs" / "performance" / "results", ROOT / "docs" / "report"):
            try:
                path.relative_to(forbidden.resolve())
            except ValueError:
                continue
            raise fail(f"{label} must stay outside {forbidden}")
    return path


def write_new(path: Path, value: object) -> dict[str, Any]:
    path = safe_location(path, "output", output=True)
    if path.exists():
        raise fail(f"refusing to overwrite existing output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = canonical(value)
    path.write_bytes(payload)
    return {"path": str(path), "bytes": len(payload), "sha256": hashlib.sha256(payload).hexdigest()}


def read_text(path: Path, label: str) -> str:
    path = safe_location(path, label)
    if not path.is_file():
        raise fail(f"{label} is not a regular file: {path}")
    try:
        return path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        raise fail(f"cannot read {label} {path}: {error}") from error


def function_segment(source: str, name: str) -> str:
    match = re.search(rf"\bfn\s+{re.escape(name)}\s*\(", source)
    if not match:
        raise fail(f"probe has no function {name}()")
    # The checked-in probe uses one four-space item per top-level helper.  A
    # segment boundary is sufficient here and, unlike a brace parser, does
    # not mistake braces in raw Rust strings for syntax.
    end = re.search(r"\n    (?:fn|struct|enum|impl|#\[)", source[match.start() + 1 :])
    stop = match.start() + 1 + end.start() if end else len(source)
    return source[match.start() : stop]


def integer_constant(source: str, name: str) -> int:
    match = re.search(
        rf"\bconst\s+{re.escape(name)}\s*:\s*[^=;]+?=\s*([0-9][0-9_]*)\s*;",
        source,
    )
    if not match:
        raise fail(f"probe has no parseable {name} constant")
    return int(match.group(1).replace("_", ""))


def route_recipe(source: str) -> tuple[list[str], dict[str, list[str]]]:
    suffix = function_segment(source, "suffix")
    suffixes = dict(re.findall(r"Self::([A-Za-z0-9_]+)\s*=>\s*\"([^\"]+)\"", suffix))
    if set(suffixes.values()) != {
        "open_get", "save", "repeated_save", "set_commit", "set_publication",
        "batch_commit", "batch_publication", "import_commit", "import_publication",
        "merge_commit", "merge_publication", "refresh_commit", "refresh_publication",
    }:
        raise fail("probe Route::suffix() does not expose the reviewed route vocabulary")
    routes_for = function_segment(source, "routes_for")
    initial = re.search(r"let\s+mut\s+routes\s*=\s*vec!\s*\[(.*?)\];", routes_for, re.S)
    if not initial:
        raise fail("probe routes_for() has no explicit base route list")
    names = [suffixes[name] for name in re.findall(r"Route::([A-Za-z0-9_]+)", initial.group(1))]
    if tuple(names) != BASE_ROUTES:
        raise fail(f"probe base routes changed: {names!r}")
    workbook_match = re.search(r"matches!\s*\(\s*name\s*,\s*((?:\"[a-z]+\"\s*\|\s*)+\"[a-z]+\")\s*\)", routes_for)
    workbook = tuple(re.findall(r'\"([a-z]+)\"', workbook_match.group(1))) if workbook_match else ()
    if workbook != WORKBOOK_FAMILIES:
        raise fail(f"probe workbook import family set changed: {workbook!r}")
    if 'name == "docx"' not in routes_for:
        raise fail("probe DOCX refresh route guard is missing")
    routes: dict[str, list[str]] = {}
    for family in EXPECTED_FAMILIES:
        values = list(names)
        if family in WORKBOOK_FAMILIES:
            values.extend(("import_commit", "import_publication"))
        if family == "docx":
            values.extend(("refresh_commit", "refresh_publication"))
        routes[family] = values
    if sum(len(values) for values in routes.values()) != 109:
        raise fail("probe route recipe does not contain exactly 109 routes")
    return list(suffixes.values()), routes


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


def fixture_descriptor(probe: Path) -> dict[str, Any]:
    probe = safe_location(probe, "probe")
    source = read_text(probe, "probe")
    probe_hash = digest(probe)
    if probe_hash != REVIEWED_PROBE_SHA256:
        raise fail(
            "probe is not the reviewed 109-route source; expected "
            f"{REVIEWED_PROBE_SHA256}, got {probe_hash}"
        )
    families = function_segment(source, "family_names")
    parsed_families = re.findall(r'"([a-z]+)"', families)
    if tuple(parsed_families) != EXPECTED_FAMILIES:
        raise fail(f"probe family order changed: {parsed_families!r}")
    for name in ("line", "unified_words", "doc", "docx_refresh", "docx_refresh_value", "rtf", "workbook", "pptx", "ppt", "odp", "csv", "family", "routes_for"):
        function_segment(source, name)
    # These call sites are the workload contract, not decorative labels in a
    # descriptor.  If a constructor, format, row count or batch bound moves,
    # descriptor generation fails until the recipe is reviewed alongside the
    # probe change.
    required_recipe_fragments = (
        'docx_refresh(PARAGRAPHS)',
        'word("odt", unified_words(FileFormat::Odt, PARAGRAPHS), 1_000)',
        'word("doc", doc(PARAGRAPHS), 256)',
        'word("rtf", rtf(PARAGRAPHS), 256)',
        'sheet("xlsx", FileFormat::Xlsx, 1_000)',
        'sheet("xlsb", FileFormat::Xlsb, 100)',
        'sheet("xls", FileFormat::Xls, 100)',
        'sheet("ods", FileFormat::Ods, 1_000)',
        'slides("pptx", pptx(SHAPES))',
        'slides("ppt", ppt(SHAPES))',
        'bytes: odp(SHAPES)',
        'batch: SHAPES',
    )
    for fragment in required_recipe_fragments:
        if fragment not in source:
            raise fail(f"probe fixture recipe changed; missing source marker: {fragment}")
    if "into_file()" not in source:
        raise fail("probe does not use the current Commit::into_file() extraction API")
    if '"--mode"' not in source or '"list"' not in source or '"validate"' not in source:
        raise fail("probe does not expose the guarded list/validate modes")
    constants = {
        "target_ns": integer_constant(source, "TARGET_NS"),
        "max_calls": integer_constant(source, "MAX_CALLS"),
        "paragraphs": integer_constant(source, "PARAGRAPHS"),
        "shapes": integer_constant(source, "SHAPES"),
    }
    reviewed_constants = {"target_ns": 500_000_000, "max_calls": 2_000_000, "paragraphs": 1_000, "shapes": 100}
    if constants != reviewed_constants:
        raise fail(f"probe calibration/workload constants changed: {constants!r}")
    route_names, routes = route_recipe(source)
    probe_bytes = probe.stat().st_size
    return {
        "schema": "litchi_w8_fixture_descriptor.v1",
        "generator": "tools/w8_file_provenance.py",
        "probe": {"name": probe.name, "bytes": probe_bytes, "sha256": probe_hash},
        "constants": constants,
        "families": list(EXPECTED_FAMILIES),
        "recipe": FIXTURE_RECIPE,
        "routes": routes,
        "route_count": sum(len(values) for values in routes.values()),
        "route_suffix_vocabulary": sorted(route_names),
        "api": {
            "commit_extraction": "into_file",
            "route_list": "--mode list --route '*'",
            "route_validate": "--mode validate --route FAMILY/ROUTE",
        },
        "source_fingerprints": {
            name: hashlib.sha256(function_segment(source, name).encode()).hexdigest()
            for name in ("family_names", "routes_for", "family", "docx_refresh", "workbook", "odp")
        },
    }


def load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(read_text(path, label))
    except json.JSONDecodeError as error:
        raise fail(f"{label} is not valid JSON: {error}") from error
    if not isinstance(value, dict):
        raise fail(f"{label} must contain a JSON object")
    return value


def fixture_routes(value: dict[str, Any], label: str = "fixture descriptor") -> list[str]:
    if value.get("schema") != "litchi_w8_fixture_descriptor.v1":
        raise fail(f"{label} schema is not supported")
    if value.get("probe", {}).get("sha256") != REVIEWED_PROBE_SHA256:
        raise fail(f"{label} is not bound to the reviewed probe")
    families = value.get("families")
    routes = value.get("routes")
    if families != list(EXPECTED_FAMILIES) or not isinstance(routes, dict):
        raise fail(f"{label} family or route declaration is incomplete")
    expected: list[str] = []
    for family in EXPECTED_FAMILIES:
        entries = routes.get(family)
        if not isinstance(entries, list) or not entries or not all(isinstance(route, str) for route in entries):
            raise fail(f"{label} route declaration for {family} is incomplete")
        expected.extend(f"{family}/{route}" for route in entries)
    if len(expected) != 109 or len(set(expected)) != 109 or value.get("route_count") != 109:
        raise fail(f"{label} does not describe the complete 109-route workload")
    if set(value.get("route_suffix_vocabulary", [])) != {
        "open_get", "save", "repeated_save", "set_commit", "set_publication",
        "batch_commit", "batch_publication", "import_commit", "import_publication",
        "merge_commit", "merge_publication", "refresh_commit", "refresh_publication",
    }:
        raise fail(f"{label} route vocabulary is incomplete")
    recipe = value.get("recipe")
    if recipe != FIXTURE_RECIPE:
        raise fail(f"{label} fixture recipe differs from the reviewed recipe")
    constants = value.get("constants")
    if constants != {"target_ns": 500_000_000, "max_calls": 2_000_000, "paragraphs": 1_000, "shapes": 100}:
        raise fail(f"{label} calibration/workload constants differ from the reviewed probe")
    if value.get("api") != {
        "commit_extraction": "into_file",
        "route_list": "--mode list --route '*'",
        "route_validate": "--mode validate --route FAMILY/ROUTE",
    }:
        raise fail(f"{label} API markers are incomplete")
    return expected


def verified_fixture(probe: Path, fixture: Path) -> tuple[dict[str, Any], list[str]]:
    probe = safe_location(probe, "probe")
    fixture = safe_location(fixture, "fixture descriptor")
    expected = fixture_descriptor(probe)
    value = load_json(fixture, "fixture descriptor")
    if canonical(value) != fixture.read_bytes():
        raise fail("fixture descriptor is not canonical JSON")
    if value != expected:
        raise fail("fixture descriptor does not match the reviewed probe recipe")
    return value, fixture_routes(value)


def dependency_projection(args: argparse.Namespace) -> dict[str, Any]:
    root = safe_location(args.workspace, "workspace")
    lock = safe_location(args.lock, "lockfile")
    package_manifest = safe_location(args.package_manifest, "package manifest")
    workspace_manifest = root / "Cargo.toml"
    if not workspace_manifest.is_file():
        raise fail(f"workspace Cargo.toml is missing: {workspace_manifest}")
    probe = safe_location(args.probe, "probe")
    if package_manifest != root / PACKAGE_MANIFEST_RELATIVE:
        raise fail(f"package manifest must be {PACKAGE_MANIFEST_RELATIVE} within the workspace")
    if probe != root / PROBE_RELATIVE:
        raise fail(f"probe must be {PROBE_RELATIVE} within the workspace")
    fixture = safe_location(args.fixture, "fixture descriptor")
    toolchain = safe_location(args.toolchain, "toolchain")
    for path, label in ((lock, "lockfile"), (package_manifest, "package manifest"), (toolchain, "toolchain")):
        if not path.is_file():
            raise fail(f"{label} is missing: {path}")
    try:
        lock_data = tomllib.loads(read_text(lock, "lockfile"))
        package_data = tomllib.loads(read_text(package_manifest, "package manifest"))
        workspace_data = tomllib.loads(read_text(workspace_manifest, "workspace manifest"))
    except tomllib.TOMLDecodeError as error:
        raise fail(f"invalid TOML input: {error}") from error
    packages = lock_data.get("package")
    if not isinstance(packages, list) or not packages:
        raise fail("lockfile has no package entries")
    lock_packages = []
    for item in packages:
        if not isinstance(item, dict) or not isinstance(item.get("name"), str) or not isinstance(item.get("version"), str):
            raise fail("lockfile contains an incomplete package entry")
        lock_packages.append({
            "name": item["name"],
            "version": item["version"],
            "source": item.get("source"),
            "checksum": item.get("checksum"),
            "dependencies": sorted(item.get("dependencies", [])) if isinstance(item.get("dependencies", []), list) else [],
        })
    lock_packages.sort(key=lambda item: (item["name"], item["version"], item.get("source") or ""))
    package = package_data.get("package")
    if not isinstance(package, dict) or package.get("name") != PACKAGE_NAME:
        raise fail(f"package manifest must declare package.name = {PACKAGE_NAME!r}")
    features = package_data.get("features", {})
    if not isinstance(features, dict):
        raise fail("package manifest features table is malformed")
    selected = sorted({name for value in args.features for name in value.split(",") if name})
    if args.all_features:
        selected = sorted(features)
    if not selected:
        raise fail("dependency projection requires --features or --all-features")
    missing = [name for name in selected if name not in features]
    if missing:
        raise fail(f"requested features are not declared by {package['name']}: {missing}")
    package_version = package.get("version")
    if package_version is True or (
        isinstance(package_version, dict) and package_version.get("workspace") is True
    ):
        workspace_table = workspace_data.get("workspace")
        workspace_package = workspace_table.get("package") if isinstance(workspace_table, dict) else None
        package_version = workspace_package.get("version") if isinstance(workspace_package, dict) else None
    if not isinstance(package_version, str) or not package_version:
        raise fail(f"could not resolve {package['name']} package version")
    if not any(
        item["name"] == PACKAGE_NAME and item["version"] == package_version
        for item in lock_packages
    ):
        raise fail(f"lockfile has no {PACKAGE_NAME} package at version {package_version}")
    toolchain_text = read_text(toolchain, "toolchain")
    try:
        toolchain_data = tomllib.loads(toolchain_text)
    except tomllib.TOMLDecodeError as error:
        raise fail(f"toolchain is not valid TOML: {error}") from error
    fixture_data, routes = verified_fixture(probe, fixture)
    probe_artifact = artifact(probe)
    fixture_probe = fixture_data.get("probe")
    if not isinstance(fixture_probe, dict) or fixture_probe.get("sha256") != probe_artifact["sha256"]:
        raise fail("fixture descriptor is not bound to the supplied probe")
    return {
        "schema": "litchi_w8_dependency_projection.v1",
        "generator": "tools/w8_file_provenance.py",
        "workspace": {
            "manifest_sha256": digest(workspace_manifest),
            "package_manifest_sha256": digest(package_manifest),
            "package": package["name"],
            "package_version": package_version,
        },
        "lock": {"bytes": lock.stat().st_size, "sha256": digest(lock), "packages": lock_packages},
        "features": {
            "package": package["name"],
            "selected": selected,
            "all_features": bool(args.all_features),
        },
        "build": {
            "target": args.target,
            "profile": args.profile,
            "toolchain": {
                "sha256": digest(toolchain),
                "channel": toolchain_data.get("toolchain", {}).get("channel"),
                "components": sorted(toolchain_data.get("toolchain", {}).get("components", [])),
                "profile": toolchain_data.get("toolchain", {}).get("profile"),
            },
        },
        "workload": {
            "probe_sha256": probe_artifact["sha256"],
            "fixture_sha256": digest(fixture),
            "fixture_schema": fixture_data["schema"],
            "route_count": len(routes),
        },
        # Keep the direct package declaration in the projection: lockfiles
        # alone do not record feature declarations or optional dependency
        # intent, and a changed manifest must invalidate the projection.
        "package_declarations": {
            "dependencies": package_data.get("dependencies", {}),
            "features": features,
            "workspace_dependencies": workspace_data.get("workspace", {}).get("dependencies", {}),
        },
    }


def source_assembly(args: argparse.Namespace) -> dict[str, Any]:
    source = artifact(safe_location(args.source, "source archive"))
    commit = checked_commit(args.source_commit, "source commit")
    description = args.description.strip()
    if not description:
        raise fail("source assembly requires a nonempty overlay description")
    overlay: dict[str, Any] = {"kind": args.overlay_kind, "description": description}
    if args.overlay_kind == "compatibility":
        if not args.overlay_base_commit or not args.overlay_patch:
            raise fail("compatibility source assembly requires base commit and overlay patch")
        overlay["base_commit"] = checked_commit(args.overlay_base_commit, "overlay base commit")
        overlay["patch"] = artifact(safe_location(args.overlay_patch, "overlay patch"))
    elif args.overlay_base_commit or args.overlay_patch:
        raise fail("a no-overlay source assembly cannot carry base commit or patch")
    return {
        "schema": "litchi_w8_source_assembly.v1",
        "resolved_commit": commit,
        "source_commit": commit,
        "source_archive": source,
        "overlay": overlay,
    }


def verified_source_assembly(path: Path, source: dict[str, Any], commit: str) -> dict[str, Any]:
    value = load_json(path, "source assembly")
    if canonical(value) != path.read_bytes():
        raise fail("source assembly is not canonical JSON")
    if value.get("schema") != "litchi_w8_source_assembly.v1":
        raise fail("source assembly schema is not supported")
    if checked_commit(value.get("resolved_commit"), "source assembly resolved commit") != commit:
        raise fail("source assembly commit differs from endpoint source commit")
    if checked_commit(value.get("source_commit"), "source assembly source commit") != commit:
        raise fail("source assembly source commit differs from endpoint source commit")
    same_artifact(source, value.get("source_archive"), "source assembly archive")
    overlay = value.get("overlay")
    if not isinstance(overlay, dict) or not isinstance(overlay.get("description"), str) or not overlay["description"].strip():
        raise fail("source assembly overlay description is missing")
    kind = overlay.get("kind")
    if kind == "compatibility":
        checked_commit(overlay.get("base_commit"), "source assembly overlay base commit")
        patch = overlay.get("patch")
        if not isinstance(patch, dict):
            raise fail("compatibility source assembly lacks its patch artifact")
        same_artifact(artifact(Path(patch.get("path", ""))), patch, "source assembly overlay patch")
    elif kind != "none":
        raise fail("source assembly overlay kind must be none or compatibility")
    return value


def verified_projection(path: Path, args: argparse.Namespace) -> tuple[dict[str, Any], dict[str, Any]]:
    path = safe_location(path, "dependency projection")
    projection_artifact = artifact(path)
    value = load_json(path, "dependency projection")
    if value.get("schema") != "litchi_w8_dependency_projection.v1":
        raise fail("dependency projection schema is not supported")
    if canonical(value) != path.read_bytes():
        raise fail("dependency projection is not canonical JSON")
    expected = dependency_projection(args)
    if value != expected:
        raise fail("dependency projection does not match supplied lock, manifests, features and workload")
    return value, projection_artifact


def verified_validation(
    path: Path,
    source_commit: str,
    source: dict[str, Any],
    binary: dict[str, Any],
    lock: dict[str, Any],
    probe: dict[str, Any],
    routes: list[str],
) -> dict[str, Any]:
    value = load_json(path, "untimed validation receipt")
    if value.get("schema") != "litchi_file_untimed_validation.v1":
        raise fail("untimed validation receipt schema is not supported")
    if checked_commit(value.get("source_commit"), "untimed validation source commit") != source_commit:
        raise fail("untimed validation receipt source commit differs from endpoint source")
    if value.get("route_count") != 109:
        raise fail("untimed validation receipt is not the complete 109-route matrix")
    outcomes = value.get("outcomes")
    if not isinstance(outcomes, list) or len(outcomes) != 109:
        raise fail("untimed validation receipt has no complete outcome list")
    outcome_routes = [item.get("route") for item in outcomes if isinstance(item, dict)]
    if outcome_routes != routes or len(set(outcome_routes)) != 109:
        raise fail("untimed validation receipt route order differs from the reviewed fixture")
    for index, outcome in enumerate(outcomes):
        if not isinstance(outcome, dict) or outcome.get("index") != index:
            raise fail("untimed validation receipt outcome indexes are incomplete")
        for key in ("semantic_equal", "semantic_state_valid", "source_state_preserved", "source_state_exact", "exact_noop"):
            if outcome.get(key) is not True:
                raise fail(f"untimed validation receipt route {outcome.get('route')} is not fully validated")
    artifacts = value.get("artifacts")
    if not isinstance(artifacts, dict):
        raise fail("untimed validation receipt artifacts are missing")
    for name, expected in (("source", source), ("binary", binary), ("lock", lock), ("probe", probe)):
        receipt_artifact = artifacts.get(name)
        if not isinstance(receipt_artifact, dict):
            raise fail(f"untimed validation receipt lacks {name} artifact")
        same_artifact(expected, receipt_artifact, f"untimed validation {name}")
        actual_receipt_artifact = artifact(Path(receipt_artifact.get("path", "")))
        same_artifact(actual_receipt_artifact, receipt_artifact, f"untimed validation {name} receipt")
    routes_artifact = artifacts.get("routes")
    if not isinstance(routes_artifact, dict):
        raise fail("untimed validation receipt lacks raw routes artifact")
    same_artifact(artifact(Path(routes_artifact.get("path", ""))), routes_artifact, "untimed validation routes")
    return value


def verified_build_receipt(
    path: Path,
    source_commit: str,
    source: dict[str, Any],
    source_assembly: dict[str, Any],
    binary: dict[str, Any],
    lock: dict[str, Any],
    probe: dict[str, Any],
    fixture: dict[str, Any],
    projection_sha256: str,
    workspace_manifest: dict[str, Any],
    package_manifest: dict[str, Any],
    toolchain: dict[str, Any],
    features: dict[str, Any],
    target: str,
    profile: str,
) -> dict[str, Any]:
    value = load_json(path, "build receipt")
    if value.get("schema") != "litchi_w8_build_receipt.v1" or value.get("status") != "complete":
        raise fail("build receipt is not a complete W8 build receipt")
    if checked_commit(value.get("source_commit"), "build receipt source commit") != source_commit:
        raise fail("build receipt source commit differs from endpoint source")
    for name, expected in (
        ("source_archive", source), ("source_assembly", source_assembly), ("binary", binary),
        ("lock", lock), ("probe", probe), ("fixture", fixture),
    ):
        receipt_artifact = value.get(name)
        if not isinstance(receipt_artifact, dict):
            raise fail(f"build receipt lacks {name}")
        same_artifact(expected, receipt_artifact, f"build receipt {name}")
        same_artifact(artifact(Path(receipt_artifact.get("path", ""))), receipt_artifact, f"build receipt {name} file")
    if not isinstance(value.get("dependency_projection_sha256"), str) or not SHA256.fullmatch(value["dependency_projection_sha256"]):
        raise fail("build receipt dependency projection hash is missing")
    if value["dependency_projection_sha256"] != projection_sha256:
        raise fail("build receipt dependency projection differs from endpoint projection")
    if value.get("workspace_manifest_sha256") != workspace_manifest["sha256"] or value.get("package_manifest_sha256") != package_manifest["sha256"]:
        raise fail("build receipt manifest identities differ from endpoint manifests")
    if value.get("toolchain_sha256") != toolchain["sha256"]:
        raise fail("build receipt toolchain differs from endpoint toolchain")
    if value.get("features") != features or value.get("target") != target or value.get("profile") != profile:
        raise fail("build receipt build selection differs from dependency projection")
    build_command = value.get("build_command")
    if not isinstance(build_command, str) or not build_command.strip():
        raise fail("build receipt has no build command")
    # Older reviewed receipts predate explicit shim binding.  New receipts
    # produced by ``build-receipt`` carry the cargo artifact and therefore get
    # the stronger command/identity check here as well.
    cargo_value = value.get("cargo")
    if cargo_value is not None:
        if not isinstance(cargo_value, dict) or not isinstance(cargo_value.get("path"), str):
            raise fail("build receipt cargo shim artifact is incomplete")
        cargo_path = safe_location(Path(cargo_value["path"]), "build receipt cargo shim")
        actual_cargo = artifact(cargo_path)
        same_artifact(actual_cargo, cargo_value, "build receipt cargo shim")
        if not os.access(cargo_path, os.X_OK):
            raise fail("build receipt cargo shim is not executable")
        _check_build_command(build_command, target, profile, cargo_path)
    compiler_artifacts = value.get("compiler_artifacts")
    if not isinstance(compiler_artifacts, list) or not compiler_artifacts:
        raise fail("build receipt has no compiler artifact hashes")
    for index, compiler_artifact in enumerate(compiler_artifacts):
        if not isinstance(compiler_artifact, dict) or not compiler_artifact.get("name"):
            raise fail(f"build receipt compiler artifact {index} is incomplete")
        same_artifact(artifact(Path(compiler_artifact.get("path", ""))), compiler_artifact, f"build receipt compiler artifact {index}")
    return value


def _check_build_command(command: str, target: str, profile: str, cargo: Path) -> str:
    command = command.strip()
    if not command:
        raise fail("build receipt requires a nonempty build command")
    required = ("CARGO_BUILD_JOBS=1", "CARGO_TARGET_DIR=", "--offline", "--locked", f"--target {target}")
    missing = [marker for marker in required if marker not in command]
    if profile == "release":
        if "--release" not in command:
            missing.append("--release")
    elif f"--profile {profile}" not in command:
        missing.append(f"--profile {profile}")
    if str(cargo.resolve()) not in command:
        missing.append(str(cargo.resolve()))
    if missing:
        raise fail(f"build receipt command is missing reviewed flags or cargo shim: {missing!r}")
    return command


def write_build_receipt(args: argparse.Namespace) -> dict[str, Any]:
    """Create and immediately validate a complete post-build receipt.

    This is deliberately separate from ``endpoint``: the build stage creates
    the receipt immediately after the binary exists, while the endpoint stage
    later binds it to the untimed 109-route validation receipt.  Every source,
    lock, manifest, projection, compiler and cargo-shim artifact is hashed
    here; no caller-supplied metadata is accepted as an identity.
    """
    workspace = safe_location(args.workspace, "workspace")
    workspace_manifest_path = safe_location(args.workspace_manifest, "workspace manifest")
    package_manifest_path = safe_location(args.package_manifest, "package manifest")
    if workspace_manifest_path != workspace / "Cargo.toml":
        raise fail("workspace manifest is not the supplied workspace Cargo.toml")
    if package_manifest_path != workspace / PACKAGE_MANIFEST_RELATIVE:
        raise fail(f"package manifest must be {PACKAGE_MANIFEST_RELATIVE} within the workspace")
    probe_path = safe_location(args.probe, "probe")
    if probe_path != workspace / PROBE_RELATIVE:
        raise fail(f"probe must be {PROBE_RELATIVE} within the workspace")
    fixture_path = safe_location(args.fixture, "fixture descriptor")
    _fixture_value, _routes = verified_fixture(probe_path, fixture_path)
    fixture_artifact = artifact(fixture_path)
    probe_artifact = artifact(probe_path)
    source_artifact = artifact(safe_location(args.source, "source archive"))
    source_commit = checked_commit(args.source_commit, "source commit")
    source_assembly_path = safe_location(args.source_assembly, "source assembly")
    source_assembly_artifact = artifact(source_assembly_path)
    verified_source_assembly(source_assembly_path, source_artifact, source_commit)
    lock_artifact = artifact(safe_location(args.lock, "lockfile"))
    workspace_manifest = artifact(workspace_manifest_path)
    package_manifest = artifact(package_manifest_path)
    toolchain = artifact(safe_location(args.toolchain, "toolchain"))
    binary_path = safe_location(args.binary, "binary")
    binary_artifact = artifact(binary_path)
    if not os.access(binary_path, os.X_OK):
        raise fail(f"binary is not executable: {binary_path}")
    projection_args = argparse.Namespace(
        workspace=workspace,
        package_manifest=package_manifest_path,
        lock=args.lock,
        toolchain=args.toolchain,
        probe=probe_path,
        fixture=fixture_path,
        features=args.features,
        all_features=args.all_features,
        target=args.target,
        profile=args.profile,
    )
    projection, projection_artifact = verified_projection(args.dependency_projection, projection_args)
    features = {
        "package": projection["features"]["package"],
        "selected": projection["features"]["selected"],
        "all_features": projection["features"]["all_features"],
    }
    cargo_path = safe_location(args.cargo, "cargo shim")
    cargo_artifact = artifact(cargo_path)
    if not os.access(cargo_path, os.X_OK):
        raise fail(f"cargo shim is not executable: {cargo_path}")
    compiler_artifacts: list[dict[str, Any]] = []
    compiler_names: set[str] = set()
    for declaration in args.compiler_artifacts:
        if "=" not in declaration:
            raise fail("compiler artifact must use NAME=PATH")
        name, raw_path = declaration.split("=", 1)
        name = name.strip()
        if not name or not raw_path:
            raise fail("compiler artifact NAME=PATH is incomplete")
        if name in compiler_names:
            raise fail(f"compiler artifact is duplicated: {name}")
        compiler_names.add(name)
        compiler_path = safe_location(Path(raw_path), f"compiler artifact {name}")
        compiler_artifact = artifact(compiler_path)
        if not os.access(compiler_path, os.X_OK):
            raise fail(f"compiler artifact is not executable: {compiler_path}")
        compiler_artifacts.append({"name": name, **compiler_artifact})
    if not compiler_artifacts:
        raise fail("build receipt requires at least one compiler artifact")
    command = _check_build_command(args.build_command, args.target, args.profile, cargo_path)
    value = {
        "schema": "litchi_w8_build_receipt.v1",
        "status": "complete",
        "source_commit": source_commit,
        "source_archive": source_artifact,
        "source_assembly": source_assembly_artifact,
        "binary": binary_artifact,
        "lock": lock_artifact,
        "probe": probe_artifact,
        "fixture": fixture_artifact,
        "dependency_projection_sha256": projection_artifact["sha256"],
        "workspace_manifest_sha256": workspace_manifest["sha256"],
        "package_manifest_sha256": package_manifest["sha256"],
        "toolchain_sha256": toolchain["sha256"],
        "features": features,
        "target": args.target,
        "profile": args.profile,
        "cargo": cargo_artifact,
        "build_command": command,
        "compiler_artifacts": compiler_artifacts,
    }
    output = safe_location(args.output, "build receipt", output=True)
    if output.exists() or output.is_symlink():
        raise fail(f"refusing to overwrite existing output: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.{os.getpid()}.tmp")
    if temporary.exists() or temporary.is_symlink():
        raise fail(f"refusing to reuse temporary output: {temporary}")
    try:
        temporary.write_bytes(canonical(value))
        verified_build_receipt(
            temporary, source_commit, source_artifact, source_assembly_artifact,
            binary_artifact, lock_artifact, probe_artifact, fixture_artifact,
            projection_artifact["sha256"], workspace_manifest, package_manifest,
            toolchain, features, args.target, args.profile,
        )
        os.replace(temporary, output)
    except Exception:
        try:
            temporary.unlink()
        except FileNotFoundError:
            pass
        raise
    return artifact(output)


def endpoint_manifest(args: argparse.Namespace) -> dict[str, Any]:
    workspace = safe_location(args.workspace, "workspace")
    workspace_manifest_path = safe_location(args.workspace_manifest, "workspace manifest")
    package_manifest_path = safe_location(args.package_manifest, "package manifest")
    toolchain_path = safe_location(args.toolchain, "toolchain")
    if workspace_manifest_path != workspace / "Cargo.toml":
        raise fail("workspace manifest is not the supplied workspace Cargo.toml")
    if package_manifest_path != workspace / PACKAGE_MANIFEST_RELATIVE:
        raise fail(f"package manifest must be {PACKAGE_MANIFEST_RELATIVE} within the workspace")
    probe_path = safe_location(args.probe, "probe")
    if probe_path != workspace / PROBE_RELATIVE:
        raise fail(f"probe must be {PROBE_RELATIVE} within the workspace")
    workspace_manifest = artifact(workspace_manifest_path)
    package_manifest = artifact(package_manifest_path)
    toolchain = artifact(toolchain_path)
    lock = artifact(safe_location(args.lock, "lockfile"))
    probe = artifact(probe_path)
    fixture_path = safe_location(args.fixture, "fixture descriptor")
    fixture, routes = verified_fixture(args.probe, fixture_path)
    fixture_artifact = artifact(fixture_path)
    source = artifact(safe_location(args.source, "source archive"))
    binary_path = safe_location(args.binary, "binary")
    binary = artifact(binary_path)
    if not os.access(binary_path, os.X_OK):
        raise fail(f"binary is not executable: {binary_path}")
    commit = checked_commit(args.source_commit, "source commit")
    source_assembly_path = safe_location(args.source_assembly, "source assembly")
    source_assembly_artifact = artifact(source_assembly_path)
    verified_source_assembly(source_assembly_path, source, commit)
    projection_args = argparse.Namespace(
        workspace=workspace,
        package_manifest=package_manifest_path,
        lock=args.lock,
        toolchain=toolchain_path,
        probe=args.probe,
        fixture=args.fixture,
        features=args.features,
        all_features=args.all_features,
        target=args.target,
        profile=args.profile,
    )
    projection, projection_artifact = verified_projection(args.dependency_projection, projection_args)
    features = {
        "package": projection["features"]["package"],
        "selected": projection["features"]["selected"],
        "all_features": projection["features"]["all_features"],
    }
    build_receipt_path = safe_location(args.build_receipt, "build receipt")
    build_receipt_artifact = artifact(build_receipt_path)
    build_receipt = verified_build_receipt(
        build_receipt_path, commit, source, source_assembly_artifact, binary, lock, probe,
        fixture_artifact, projection_artifact["sha256"], workspace_manifest, package_manifest, toolchain,
        features, args.target, args.profile,
    )
    # Keep the receipt's own identity in the endpoint: runtime verification
    # must not trust only a filename supplied by the caller.
    del build_receipt
    validation_path = safe_location(args.validation_receipt, "untimed validation receipt")
    validation_artifact = artifact(validation_path)
    verified_validation(validation_path, commit, source, binary, lock, probe, routes)
    return {
        "schema": "litchi_file_endpoint.v1",
        "source_commit": commit,
        "dependency_projection_sha256": projection_artifact["sha256"],
        "binary": binary,
        "source": source,
        "lock": lock,
        "probe": probe,
        "fixture": fixture_artifact,
        "dependency_projection": projection_artifact,
        "workspace_manifest": workspace_manifest,
        "package_manifest": package_manifest,
        "toolchain": toolchain,
        "source_assembly": source_assembly_artifact,
        "build_receipt": build_receipt_artifact,
        "validation": validation_artifact,
        "features": features,
        "target": args.target,
        "profile": args.profile,
    }


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__.split("\n\n", 1)[0])
    commands = root.add_subparsers(dest="command", required=True)
    fixture = commands.add_parser("fixture", help="write the deterministic W8 fixture descriptor")
    fixture.add_argument("--probe", required=True, type=Path)
    fixture.add_argument("--output", required=True, type=Path)
    dependency = commands.add_parser("dependency", help="write the lock/feature/toolchain projection")
    dependency.add_argument("--workspace", required=True, type=Path)
    dependency.add_argument("--package-manifest", required=True, type=Path)
    dependency.add_argument("--lock", required=True, type=Path)
    dependency.add_argument("--toolchain", required=True, type=Path)
    dependency.add_argument("--probe", required=True, type=Path)
    dependency.add_argument("--fixture", required=True, type=Path)
    dependency.add_argument("--feature", dest="features", action="append", default=[])
    dependency.add_argument("--all-features", action="store_true")
    dependency.add_argument("--target", required=True)
    dependency.add_argument("--profile", default="release")
    dependency.add_argument("--output", required=True, type=Path)
    assembly = commands.add_parser("source", help="write explicit source/compatibility assembly metadata")
    assembly.add_argument("--source", required=True, type=Path)
    assembly.add_argument("--source-commit", required=True)
    assembly.add_argument("--overlay-kind", choices=("none", "compatibility"), required=True)
    assembly.add_argument("--overlay-base-commit")
    assembly.add_argument("--overlay-patch", type=Path)
    assembly.add_argument("--description", required=True)
    assembly.add_argument("--output", required=True, type=Path)
    build_receipt = commands.add_parser("build-receipt", help="write and validate a complete post-build receipt")
    for name in (
        "source", "source-assembly", "binary", "lock", "probe", "fixture",
        "dependency-projection", "workspace", "workspace-manifest",
        "package-manifest", "toolchain", "cargo", "output",
    ):
        build_receipt.add_argument(f"--{name}", required=True, type=Path)
    build_receipt.add_argument("--source-commit", required=True)
    build_receipt.add_argument("--compiler-artifact", dest="compiler_artifacts", action="append", required=True)
    build_receipt.add_argument("--build-command", required=True)
    build_receipt.add_argument("--feature", dest="features", action="append", default=[])
    build_receipt.add_argument("--all-features", action="store_true")
    build_receipt.add_argument("--target", required=True)
    build_receipt.add_argument("--profile", default="release")
    endpoint = commands.add_parser("endpoint", help="write a W8 endpoint manifest")
    for name in (
        "binary", "source", "lock", "probe", "fixture", "dependency-projection",
        "workspace", "workspace-manifest", "package-manifest", "toolchain",
        "source-assembly", "build-receipt", "validation-receipt", "output",
    ):
        endpoint.add_argument(f"--{name}", required=True, type=Path)
    endpoint.add_argument("--source-commit", required=True)
    endpoint.add_argument("--feature", dest="features", action="append", default=[])
    endpoint.add_argument("--all-features", action="store_true")
    endpoint.add_argument("--target", required=True)
    endpoint.add_argument("--profile", default="release")
    return root


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    if args.command == "fixture":
        result = write_new(args.output, fixture_descriptor(args.probe))
    elif args.command == "dependency":
        result = write_new(args.output, dependency_projection(args))
    elif args.command == "source":
        result = write_new(args.output, source_assembly(args))
    elif args.command == "build-receipt":
        result = write_build_receipt(args)
    else:
        result = write_new(args.output, endpoint_manifest(args))
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ProvenanceError as error:
        print(f"w8 provenance: {error}", file=sys.stderr)
        raise SystemExit(2)
