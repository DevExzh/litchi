#!/usr/bin/env python3
"""Run the non-iWork release gate without selecting the iWork workspace.

The workspace deliberately keeps the Apple iWork packages in the same Cargo
workspace as the Office and ODF packages.  This gate is the cheap, explicit
non-iWork slice: it derives the package and facade-feature sets from Cargo
metadata, checks every selected dependency tree, and then runs the requested
Cargo operation with the 45 ordinary packages plus the safe facade feature
closure.  The library-test mode serializes package roots and cleans each
successful root's artifacts to keep its test target from retaining every
package's test binaries at once.

All Cargo invocations are argv lists.  In particular, feature names and
package exclusions never pass through a shell.  ``verify`` does not build
targets; it performs metadata and ``cargo tree`` checks and may resolve the
workspace or update the ignored root lockfile.  Most other modes run one bulk
workspace command and one facade command using an isolated target directory by
default.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import shlex
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parents[1]
FACADE_PACKAGE = "litchi"
PY_FACADE_PACKAGE = "litchi-py"
IWORK_LEAF_PACKAGES = frozenset(
    {
        "litchi-keynote",
        "litchi-numbers",
        "litchi-numbers-wire",
        "litchi-pages",
    }
)
IWORK_PACKAGE_PREFIX = "litchi-iwa"
EXPECTED_IWORK_PACKAGES = 17
EXPECTED_EXCLUDED_PACKAGES = 18
EXPECTED_BULK_PACKAGES = 45

# The release gate is intentionally audited against the current workspace
# topology, not merely package counts.  A rename, swap, or newly added member
# must update this inventory and its review before it can enter the gate.
EXPECTED_IWORK_PACKAGE_NAMES = frozenset(
    {
        "litchi-iwa",
        "litchi-iwa-archive",
        "litchi-iwa-cache",
        "litchi-iwa-common",
        "litchi-iwa-core",
        "litchi-iwa-detect",
        "litchi-iwa-graph",
        "litchi-iwa-index",
        "litchi-iwa-package",
        "litchi-iwa-protos",
        "litchi-iwa-structured",
        "litchi-iwa-text",
        "litchi-iwa-text-wire",
        "litchi-keynote",
        "litchi-numbers",
        "litchi-numbers-wire",
        "litchi-pages",
    }
)
EXPECTED_EXCLUDED_PACKAGE_NAMES = EXPECTED_IWORK_PACKAGE_NAMES | {PY_FACADE_PACKAGE}
EXPECTED_BULK_PACKAGE_NAMES = frozenset(
    {
        "litchi-biff",
        "litchi-cfb",
        "litchi-codepage",
        "litchi-core",
        "litchi-crypto",
        "litchi-doc",
        "litchi-docx",
        "litchi-drawingml",
        "litchi-eval",
        "litchi-fonts",
        "litchi-formula",
        "litchi-imgconv",
        "litchi-markdown",
        "litchi-odb",
        "litchi-odc",
        "litchi-odf",
        "litchi-odf-common",
        "litchi-odf-formula",
        "litchi-odg",
        "litchi-odi",
        "litchi-odm",
        "litchi-odp",
        "litchi-odraw",
        "litchi-ods",
        "litchi-odt",
        "litchi-ograph",
        "litchi-ole-common",
        "litchi-ooxml-common",
        "litchi-opc",
        "litchi-oth",
        "litchi-ppt",
        "litchi-pptx",
        "litchi-rtf",
        "litchi-sheet",
        "litchi-sign",
        "litchi-slide",
        "litchi-spreadsheet-drawing",
        "litchi-vba",
        "litchi-word",
        "litchi-xls",
        "litchi-xlsb",
        "litchi-xlsx",
        "soapberry-zip",
        "xml-minifier",
        "xml-minifier-macros",
    }
)
EXPECTED_WORKSPACE_PACKAGE_NAMES = (
    EXPECTED_BULK_PACKAGE_NAMES
    | EXPECTED_EXCLUDED_PACKAGE_NAMES
    | {FACADE_PACKAGE}
)

# ``prost`` is the actual wire-format implementation used by the iWork
# protobuf crate.  Keep the family check broad so a newly split prost helper
# cannot silently enter a purportedly protobuf-free tree.
PROTOBUF_PACKAGE_RE = re.compile(r"^(?:prost(?:[-_]|$)|protobuf(?:[-_]|$))")
TREE_PACKAGE_RE = re.compile(r"^([A-Za-z0-9][A-Za-z0-9_-]*) v(?:[0-9]|\*)")

MODES = (
    "print",
    "verify",
    "check",
    "clippy",
    "doc",
    "lib-tests",
    "doc-tests",
    "deprecated",
)


class GateError(RuntimeError):
    """Raised for a topology, metadata, or Cargo gate failure."""


@dataclass(frozen=True)
class Dependency:
    """The workspace-relevant part of one Cargo dependency declaration."""

    name: str
    features: frozenset[str]


@dataclass(frozen=True)
class Package:
    """A workspace package projected from ``cargo metadata``."""

    name: str
    features: Mapping[str, tuple[str, ...]]
    dependencies: tuple[Dependency, ...]


@dataclass(frozen=True)
class WorkspacePlan:
    """The derived package and facade feature closure for the gate."""

    packages: Mapping[str, Package]
    facade_features: Mapping[str, tuple[str, ...]]
    iwork_packages: frozenset[str]
    unsafe_facade_dependents: frozenset[str]
    excluded_packages: frozenset[str]
    bulk_packages: frozenset[str]
    safe_facade_features: frozenset[str]
    unsafe_facade_features: frozenset[str]


@dataclass(frozen=True)
class CommandSpec:
    """One Cargo command and its non-shell environment additions."""

    scope: str
    argv: tuple[str, ...]
    env: tuple[tuple[str, str], ...] = ()


def _metadata_packages(metadata: Mapping[str, Any]) -> dict[str, Package]:
    """Project workspace members from a Cargo metadata document."""

    raw_packages = metadata.get("packages")
    workspace_members = metadata.get("workspace_members")
    if not isinstance(raw_packages, list) or not isinstance(workspace_members, list):
        raise GateError("cargo metadata lacks packages or workspace_members")

    by_id: dict[str, Mapping[str, Any]] = {}
    for raw_package in raw_packages:
        if not isinstance(raw_package, Mapping):
            raise GateError("cargo metadata contains a malformed package")
        package_id = raw_package.get("id")
        if not isinstance(package_id, str):
            raise GateError("cargo metadata package has no string id")
        by_id[package_id] = raw_package

    packages: dict[str, Package] = {}
    for package_id in workspace_members:
        if not isinstance(package_id, str) or package_id not in by_id:
            raise GateError(f"workspace member {package_id!r} is missing from metadata")
        raw_package = by_id[package_id]
        name = raw_package.get("name")
        raw_features = raw_package.get("features", {})
        raw_dependencies = raw_package.get("dependencies", [])
        if not isinstance(name, str) or not isinstance(raw_features, Mapping):
            raise GateError(f"package {name!r} has malformed feature metadata")
        if not isinstance(raw_dependencies, list):
            raise GateError(f"package {name!r} has malformed dependency metadata")

        features: dict[str, tuple[str, ...]] = {}
        for feature, values in raw_features.items():
            if not isinstance(feature, str) or not isinstance(values, list):
                raise GateError(f"package {name!r} has malformed feature {feature!r}")
            if not all(isinstance(value, str) for value in values):
                raise GateError(f"package {name!r} feature {feature!r} is not string-valued")
            features[feature] = tuple(values)

        dependencies: list[Dependency] = []
        for raw_dependency in raw_dependencies:
            if not isinstance(raw_dependency, Mapping):
                raise GateError(f"package {name!r} has a malformed dependency")
            dependency_name = raw_dependency.get("name")
            dependency_features = raw_dependency.get("features", [])
            if not isinstance(dependency_name, str) or not isinstance(
                dependency_features, list
            ):
                raise GateError(f"package {name!r} has malformed dependency metadata")
            if not all(isinstance(value, str) for value in dependency_features):
                raise GateError(
                    f"package {name!r} dependency {dependency_name!r} has invalid features"
                )
            dependencies.append(
                Dependency(dependency_name, frozenset(dependency_features))
            )
        if name in packages:
            raise GateError(f"duplicate workspace package {name!r}")
        packages[name] = Package(name, features, tuple(dependencies))
    return packages


def _iwork_packages(package_names: Iterable[str]) -> frozenset[str]:
    """Derive the iWork family from the accepted package naming topology."""

    return frozenset(
        name
        for name in package_names
        if name == IWORK_PACKAGE_PREFIX
        or name.startswith(f"{IWORK_PACKAGE_PREFIX}-")
        or name in IWORK_LEAF_PACKAGES
    )


def _require_exact_package_set(
    actual: Iterable[str], expected: frozenset[str], description: str
) -> None:
    """Reject an unreviewed package rename, swap, addition, or removal."""

    actual_set = frozenset(actual)
    if actual_set == expected:
        return
    unexpected = sorted(actual_set - expected)
    missing = sorted(expected - actual_set)
    details: list[str] = []
    if unexpected:
        details.append("unexpected: " + ", ".join(unexpected))
    if missing:
        details.append("missing: " + ", ".join(missing))
    raise GateError(f"{description} mismatch (" + "; ".join(details) + ")")


def _feature_references(
    feature_values: Mapping[str, Sequence[str]],
) -> tuple[dict[str, frozenset[str]], dict[str, frozenset[str]]]:
    """Split feature values into aliases and optional dependency references."""

    aliases: dict[str, set[str]] = {feature: set() for feature in feature_values}
    dependency_refs: dict[str, set[str]] = {feature: set() for feature in feature_values}
    feature_names = set(feature_values)
    for feature, values in feature_values.items():
        for value in values:
            if value in feature_names:
                aliases[feature].add(value)
                continue
            if value.startswith("dep:"):
                dependency_refs[feature].add(value.removeprefix("dep:"))
                continue
            # A value such as ``litchi-pages/internal-iwork-source`` is a
            # feature forwarded to an optional dependency.  Its package name
            # is the part before the slash.
            if value.startswith("litchi-"):
                dependency_refs[feature].add(value.split("/", 1)[0])
    return (
        {feature: frozenset(values) for feature, values in aliases.items()},
        {feature: frozenset(values) for feature, values in dependency_refs.items()},
    )


def _unsafe_features(
    feature_values: Mapping[str, Sequence[str]],
    iwork_packages: frozenset[str],
) -> frozenset[str]:
    """Return every facade feature whose expansion reaches an iWork package."""

    aliases, dependency_refs = _feature_references(feature_values)
    memo: dict[str, bool] = {}

    def reaches_iwork(feature: str, visiting: frozenset[str] = frozenset()) -> bool:
        if feature in memo:
            return memo[feature]
        if feature in visiting:
            raise GateError(f"cycle in litchi feature aliases at {feature!r}")
        result = bool(dependency_refs[feature] & iwork_packages)
        if not result:
            result = any(
                reaches_iwork(alias, visiting | {feature}) for alias in aliases[feature]
            )
        memo[feature] = result
        return result

    return frozenset(
        feature for feature in feature_values if reaches_iwork(feature)
    )


def derive_plan(metadata: Mapping[str, Any]) -> WorkspacePlan:
    """Derive and validate the exact non-iWork package/feature plan."""

    packages = _metadata_packages(metadata)
    if FACADE_PACKAGE not in packages:
        raise GateError("workspace does not contain the litchi facade")
    _require_exact_package_set(
        packages,
        EXPECTED_WORKSPACE_PACKAGE_NAMES,
        "workspace package inventory",
    )
    facade = packages[FACADE_PACKAGE]
    iwork_packages = _iwork_packages(packages)
    _require_exact_package_set(
        iwork_packages,
        EXPECTED_IWORK_PACKAGE_NAMES,
        "iWork package inventory",
    )
    if len(iwork_packages) != EXPECTED_IWORK_PACKAGES:
        raise GateError(
            "expected "
            f"{EXPECTED_IWORK_PACKAGES} iWork packages, found {len(iwork_packages)}: "
            + ", ".join(sorted(iwork_packages))
        )

    unsafe_features = _unsafe_features(facade.features, iwork_packages)
    safe_features = frozenset(set(facade.features) - unsafe_features)
    if safe_features | unsafe_features != frozenset(facade.features) or (
        safe_features & unsafe_features
    ):
        raise GateError("facade feature closure is not a disjoint partition")
    if not unsafe_features:
        raise GateError("facade feature closure does not identify any iWork feature")

    unsafe_dependents: set[str] = set()
    for package in packages.values():
        for dependency in package.dependencies:
            if dependency.name != FACADE_PACKAGE:
                continue
            unknown = dependency.features - set(facade.features)
            if unknown:
                raise GateError(
                    f"{package.name} requests unknown litchi features: "
                    + ", ".join(sorted(unknown))
                )
            if dependency.features & unsafe_features:
                unsafe_dependents.add(package.name)

    _require_exact_package_set(
        unsafe_dependents,
        frozenset({PY_FACADE_PACKAGE}),
        "unsafe facade dependent inventory",
    )
    excluded_packages = frozenset(iwork_packages | unsafe_dependents)
    _require_exact_package_set(
        excluded_packages,
        EXPECTED_EXCLUDED_PACKAGE_NAMES,
        "excluded package inventory",
    )
    if len(excluded_packages) != EXPECTED_EXCLUDED_PACKAGES:
        raise GateError(
            "expected "
            f"{EXPECTED_EXCLUDED_PACKAGES} excluded packages, found "
            f"{len(excluded_packages)}: "
            + ", ".join(sorted(excluded_packages))
        )
    if FACADE_PACKAGE in excluded_packages:
        raise GateError("the litchi facade must be gated separately")

    bulk_packages = frozenset(set(packages) - set(excluded_packages) - {FACADE_PACKAGE})
    _require_exact_package_set(
        bulk_packages,
        EXPECTED_BULK_PACKAGE_NAMES,
        "bulk package inventory",
    )
    if len(bulk_packages) != EXPECTED_BULK_PACKAGES:
        raise GateError(
            "expected "
            f"{EXPECTED_BULK_PACKAGES} bulk packages, found {len(bulk_packages)}"
        )
    return WorkspacePlan(
        packages=packages,
        facade_features=facade.features,
        iwork_packages=iwork_packages,
        unsafe_facade_dependents=frozenset(unsafe_dependents),
        excluded_packages=excluded_packages,
        bulk_packages=bulk_packages,
        safe_facade_features=safe_features,
        unsafe_facade_features=unsafe_features,
    )


def _cargo_metadata(cargo: str) -> Mapping[str, Any]:
    """Read workspace metadata through Cargo without invoking a shell."""

    result = subprocess.run(
        [cargo, "metadata", "--format-version", "1", "--no-deps"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode:
        details = result.stderr.strip() or result.stdout.strip()
        raise GateError(f"cargo metadata failed ({result.returncode}): {details}")
    try:
        value = json.loads(result.stdout)
    except json.JSONDecodeError as error:
        raise GateError(f"cargo metadata returned invalid JSON: {error}") from error
    if not isinstance(value, Mapping):
        raise GateError("cargo metadata returned a non-object JSON value")
    return value


def _workspace_exclusions(plan: WorkspacePlan) -> list[str]:
    """Return Cargo's workspace selection and all 19 non-bulk exclusions."""

    exclusions = [FACADE_PACKAGE, *sorted(plan.excluded_packages)]
    return [item for name in exclusions for item in ("--exclude", name)]


def _facade_feature_args(plan: WorkspacePlan) -> list[str]:
    """Return one deterministic combined safe feature argument."""

    # The default feature is empty in this workspace and is represented by
    # --no-default-features.  All other safe features are compiled together so
    # each operation mode has one bounded facade invocation.
    features = sorted(plan.safe_facade_features - {"default"})
    return ["--features", ",".join(features)] if features else []


def _facade_tree_selection(
    plan: WorkspacePlan, feature: str | None = None
) -> list[str]:
    """Build one individual or combined safe facade tree selection."""

    selection = ["--package", FACADE_PACKAGE, "--no-default-features"]
    if feature is None:
        selection.extend(_facade_feature_args(plan))
    elif feature != "default":
        selection.extend(["--features", feature])
    return selection


def command_specs(cargo: str, plan: WorkspacePlan, mode: str) -> tuple[CommandSpec, ...]:
    """Generate the bulk and facade argv lists for one run mode."""

    if mode not in MODES or mode in {"print", "verify"}:
        raise GateError(f"{mode!r} does not generate Cargo run commands")
    if mode == "lib-tests":
        # A single workspace test invocation keeps every test binary alive
        # until the final link and can consume several GiB. Serialize each
        # exact bulk root, then remove only that package's test artifacts
        # before proceeding. Dependencies remain reusable across roots.
        specs: list[CommandSpec] = []
        for package in sorted(plan.bulk_packages):
            specs.append(
                CommandSpec(
                    f"bulk-test/{package}",
                    (
                        cargo,
                        "test",
                        "--package",
                        package,
                        "--all-features",
                        "--lib",
                        "--tests",
                    ),
                )
            )
            specs.append(
                CommandSpec(
                    f"bulk-clean/{package}",
                    (cargo, "clean", "--package", package),
                )
            )
        facade = (
            cargo,
            "test",
            "--package",
            FACADE_PACKAGE,
            "--no-default-features",
            *_facade_feature_args(plan),
            "--lib",
            "--tests",
        )
        specs.append(CommandSpec("facade-safe-features", facade))
        specs.append(
            CommandSpec(
                "facade-clean",
                (cargo, "clean", "--package", FACADE_PACKAGE),
            )
        )
        return tuple(specs)
    # All features are intentional for bulk roots: each of the 45 packages
    # must prove that its non-default feature surface remains non-iWork.  The
    # facade is selected separately with the exact safe closure below.
    bulk_selection = ["--workspace", "--all-features", *_workspace_exclusions(plan)]
    facade_selection = [
        "--package",
        FACADE_PACKAGE,
        "--no-default-features",
        *_facade_feature_args(plan),
    ]
    if mode == "check":
        subcommand, common = "check", ["--lib", "--tests"]
    elif mode == "clippy":
        subcommand, common = "clippy", [
            "--lib",
            "--no-deps",
            "--",
            "-D",
            "warnings",
        ]
    elif mode == "doc":
        subcommand, common = "doc", ["--no-deps"]
    elif mode == "doc-tests":
        subcommand, common = "test", ["--doc"]
    else:
        subcommand, common = "check", ["--all-targets"]
    bulk = (cargo, subcommand, *bulk_selection, *common)
    facade = (cargo, subcommand, *facade_selection, *common)
    updates: tuple[tuple[str, str], ...] = ()
    if mode in {"doc", "doc-tests"}:
        updates = (("RUSTDOCFLAGS", "-D warnings"),)
    elif mode == "deprecated":
        updates = (("RUSTFLAGS", "-D deprecated"),)
    return (
        CommandSpec("bulk", bulk, updates),
        CommandSpec("facade-safe-features", facade, updates),
    )


def parse_tree_package_names(output: str) -> frozenset[str]:
    """Parse package names from ``cargo tree --prefix none --format {p}`."""

    names: set[str] = set()
    for line in output.splitlines():
        match = TREE_PACKAGE_RE.match(line.strip())
        if match:
            names.add(match.group(1))
    return frozenset(names)


def _protobuf_packages(names: Iterable[str]) -> frozenset[str]:
    return frozenset(name for name in names if PROTOBUF_PACKAGE_RE.match(name))


def _tree_command(cargo: str, selection: Sequence[str]) -> list[str]:
    return [
        cargo,
        "tree",
        *selection,
        # Include normal, build, and dev edges so an accidental protobuf build
        # dependency cannot hide behind the default normal-only tree.
        "--edges",
        "all",
        "--prefix",
        "none",
        "--format",
        "{p}",
    ]


def _capture_tree(cargo: str, selection: Sequence[str]) -> frozenset[str]:
    argv = _tree_command(cargo, selection)
    result = subprocess.run(
        argv,
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode:
        details = result.stderr.strip() or result.stdout.strip()
        raise GateError(f"{shlex.join(argv)} failed ({result.returncode}): {details}")
    return parse_tree_package_names(result.stdout)


def _require_tree_roots(
    names: Iterable[str], required: Iterable[str], description: str
) -> None:
    """Reject successful but empty or partial Cargo tree output."""

    present = frozenset(names)
    missing = frozenset(required) - present
    if missing:
        raise GateError(
            f"{description} is missing required package roots: "
            + ", ".join(sorted(missing))
        )


def verify_dependency_trees(cargo: str, plan: WorkspacePlan) -> tuple[int, int]:
    """Verify the bulk all-features tree and the safe facade closure."""

    bulk_selection = ["--workspace", "--all-features", *_workspace_exclusions(plan)]
    bulk_names = _capture_tree(cargo, bulk_selection)
    _require_tree_roots(bulk_names, plan.bulk_packages, "bulk dependency tree")
    forbidden_bulk = (
        set(plan.excluded_packages)
        | {FACADE_PACKAGE}
        | set(_protobuf_packages(bulk_names))
    )
    if forbidden_bulk & bulk_names:
        raise GateError(
            "bulk dependency tree pulls forbidden packages: "
            + ", ".join(sorted(forbidden_bulk & bulk_names))
        )

    facade_trees = 0
    for feature in sorted(plan.safe_facade_features):
        selection = _facade_tree_selection(plan, feature)
        names = _capture_tree(cargo, selection)
        _require_tree_roots(
            names, {FACADE_PACKAGE}, f"safe facade feature {feature!r} tree"
        )
        forbidden = (set(plan.excluded_packages) | set(_protobuf_packages(names))) & names
        if forbidden:
            raise GateError(
                f"safe facade feature {feature!r} pulls forbidden packages: "
                + ", ".join(sorted(forbidden))
            )
        facade_trees += 1
    combined_names = _capture_tree(cargo, _facade_tree_selection(plan))
    _require_tree_roots(
        combined_names, {FACADE_PACKAGE}, "combined safe facade feature tree"
    )
    combined_forbidden = (
        set(plan.excluded_packages) | set(_protobuf_packages(combined_names))
    ) & combined_names
    if combined_forbidden:
        raise GateError(
            "combined safe facade feature tree pulls forbidden packages: "
            + ", ".join(sorted(combined_forbidden))
        )
    return len(plan.bulk_packages), facade_trees


def _environment(target_dir: str | None) -> dict[str, str]:
    """Build a bounded Cargo environment for local and CI gate runs."""

    environment = os.environ.copy()
    if target_dir:
        candidate = Path(target_dir)
        if not candidate.is_absolute():
            candidate = ROOT / candidate
        environment["CARGO_TARGET_DIR"] = str(candidate)
    else:
        environment.setdefault("CARGO_TARGET_DIR", str(ROOT / "target/non-iwork-gate"))
    environment.setdefault("CARGO_BUILD_JOBS", "1")
    # All six CI modes share one isolated target sequentially. Disable
    # incremental directories so repeated modes do not accumulate a second
    # per-crate object graph alongside the reusable non-incremental artifacts.
    environment.setdefault("CARGO_INCREMENTAL", "0")
    # Keep debug/test artifacts small enough for the serialized six-mode job;
    # this changes symbol retention only, not which targets are checked.
    environment.setdefault("CARGO_PROFILE_DEV_DEBUG", "0")
    environment.setdefault("CARGO_PROFILE_TEST_DEBUG", "0")
    return environment


def _updated_environment(
    base: Mapping[str, str], updates: Sequence[tuple[str, str]]
) -> dict[str, str]:
    environment = dict(base)
    for key, value in updates:
        existing = environment.get(key)
        if existing and value not in existing:
            environment[key] = f"{existing} {value}"
        else:
            environment[key] = value
    return environment


def run_mode(cargo: str, plan: WorkspacePlan, mode: str, environment: Mapping[str, str]) -> None:
    """Run one generated mode, stopping at the first failed Cargo command."""

    for spec in command_specs(cargo, plan, mode):
        command_environment = _updated_environment(environment, spec.env)
        print(f"[{mode}/{spec.scope}] $ {shlex.join(spec.argv)}", flush=True)
        result = subprocess.run(
            list(spec.argv),
            cwd=ROOT,
            env=command_environment,
            check=False,
        )
        if result.returncode:
            raise GateError(
                f"{mode}/{spec.scope} failed with exit status {result.returncode}"
            )


def print_plan(plan: WorkspacePlan) -> None:
    """Print stable machine-readable counts followed by human-readable sets."""

    print(f"workspace_packages={len(plan.packages)}")
    print(f"excluded_packages={len(plan.excluded_packages)}")
    print(f"bulk_packages={len(plan.bulk_packages)}")
    print(f"facade_safe_features={len(plan.safe_facade_features)}")
    print(f"facade_unsafe_features={len(plan.unsafe_facade_features)}")
    print("excluded=" + ",".join(sorted(plan.excluded_packages)))
    print("bulk=" + ",".join(sorted(plan.bulk_packages)))
    print("facade-safe=" + ",".join(sorted(plan.safe_facade_features)))
    print("facade-unsafe=" + ",".join(sorted(plan.unsafe_facade_features)))
    print(
        "facade-safe-feature-argument="
        + ",".join(sorted(plan.safe_facade_features - {"default"}))
    )


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("mode", choices=MODES)
    parser.add_argument("--cargo", default="cargo", help="Cargo executable")
    parser.add_argument(
        "--target-dir",
        help="isolated Cargo target directory (default: target/non-iwork-gate)",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        plan = derive_plan(_cargo_metadata(args.cargo))
        if args.mode == "print":
            print_plan(plan)
            return 0
        if args.mode == "verify":
            print_plan(plan)
            bulk_count, facade_count = verify_dependency_trees(args.cargo, plan)
            print(f"verified_bulk_tree_roots={bulk_count}")
            print(f"verified_facade_safe_trees={facade_count}")
            print("verified_facade_combined_tree=1")
            return 0
        run_mode(args.cargo, plan, args.mode, _environment(args.target_dir))
        return 0
    except GateError as error:
        print(f"non-iWork release gate: {error}", file=sys.stderr)
        return 1
    except OSError as error:
        print(f"non-iWork release gate: cannot execute Cargo: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":  # pragma: no cover - exercised by the CLI
    raise SystemExit(main())
