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
from pathlib import Path
import re
import sys
import tomllib
from typing import Any


ROOT = Path(__file__).resolve().parents[1]
SHA256 = re.compile(r"^[0-9a-f]{64}$")
COMMIT = re.compile(r"^[0-9a-f]{7,64}$")
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
        "probe": {"name": probe.name, "bytes": probe_bytes, "sha256": digest(probe)},
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


def dependency_projection(args: argparse.Namespace) -> dict[str, Any]:
    root = safe_location(args.workspace, "workspace")
    lock = safe_location(args.lock, "lockfile")
    package_manifest = safe_location(args.package_manifest, "package manifest")
    workspace_manifest = root / "Cargo.toml"
    if not workspace_manifest.is_file():
        raise fail(f"workspace Cargo.toml is missing: {workspace_manifest}")
    probe = safe_location(args.probe, "probe")
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
    if not isinstance(package, dict) or not isinstance(package.get("name"), str):
        raise fail("package manifest has no package.name")
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
    if package_version is True:
        package_version = workspace_data.get("workspace", {}).get("package", {}).get("version")
    if not isinstance(package_version, str) or not package_version:
        raise fail(f"could not resolve {package['name']} package version")
    toolchain_text = read_text(toolchain, "toolchain")
    try:
        toolchain_data = tomllib.loads(toolchain_text)
    except tomllib.TOMLDecodeError as error:
        raise fail(f"toolchain is not valid TOML: {error}") from error
    fixture_data = load_json(fixture, "fixture descriptor")
    if fixture_data.get("schema") != "litchi_w8_fixture_descriptor.v1":
        raise fail("fixture descriptor schema is not supported")
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
            "route_count": fixture_data.get("route_count"),
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


def source_commit(value: str) -> str:
    if not COMMIT.fullmatch(value):
        raise fail(f"source commit is not a hexadecimal commit id: {value!r}")
    return value


def endpoint_manifest(args: argparse.Namespace) -> dict[str, Any]:
    projection = safe_location(args.dependency_projection, "dependency projection")
    projection_artifact = artifact(projection)
    return {
        "schema": "litchi_file_endpoint.v1",
        "source_commit": source_commit(args.source_commit),
        "dependency_projection_sha256": projection_artifact["sha256"],
        "binary": artifact(safe_location(args.binary, "binary")),
        "source": artifact(safe_location(args.source, "source archive")),
        "lock": artifact(safe_location(args.lock, "lockfile")),
        "probe": artifact(safe_location(args.probe, "probe")),
        "fixture": artifact(safe_location(args.fixture, "fixture descriptor")),
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
    endpoint = commands.add_parser("endpoint", help="write a W8 endpoint manifest")
    for name in ("binary", "source", "lock", "probe", "fixture", "dependency-projection", "output"):
        endpoint.add_argument(f"--{name}", required=True, type=Path)
    endpoint.add_argument("--source-commit", required=True)
    return root


def main(argv: list[str] | None = None) -> int:
    args = parser().parse_args(argv)
    if args.command == "fixture":
        result = write_new(args.output, fixture_descriptor(args.probe))
    elif args.command == "dependency":
        result = write_new(args.output, dependency_projection(args))
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
