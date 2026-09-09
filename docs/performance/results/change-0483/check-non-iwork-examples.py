#!/usr/bin/env python3
"""Check workspace example targets with the iWork package family scoped out.

The repository-wide checker remains authoritative for the full workspace.  This
read-only companion reuses its manifest and target resolvers, then permits
collisions whose records all belong to the explicitly recorded iWork family.
Collisions involving any retained package still fail closed.
"""

from __future__ import annotations

from collections import defaultdict
import json
from pathlib import Path
import sys
from typing import Any


BUNDLE_ROOT = Path(__file__).resolve().parent
REPO = BUNDLE_ROOT.parents[3]
EXCLUSION_PATH = BUNDLE_ROOT / "workspace-iwork-exclusion.json"
EXPECTED_EXCLUSION_FIELDS = {
    "excluded_packages",
    "full_workspace_failed_gate",
    "failure",
    "reason",
    "replacement_gate",
}

sys.path.insert(0, str(REPO / "tools"))
import check_example_targets as checker  # noqa: E402


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def _read_exclusion_inventory() -> tuple[str, ...]:
    try:
        value = json.loads(
            EXCLUSION_PATH.read_text(encoding="utf-8"),
            object_pairs_hook=_pairs,
            parse_constant=lambda constant: (_ for _ in ()).throw(
                ValueError(constant)
            ),
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        raise checker.CheckFailure(
            f"{EXCLUSION_PATH}: malformed iWork exclusion metadata: {error}"
        ) from error
    if not isinstance(value, dict) or set(value) != EXPECTED_EXCLUSION_FIELDS:
        raise checker.CheckFailure(
            f"{EXCLUSION_PATH}: exclusion metadata fields are malformed"
        )
    for field in (
        "full_workspace_failed_gate",
        "failure",
        "reason",
        "replacement_gate",
    ):
        if not isinstance(value[field], str) or not value[field].strip():
            raise checker.CheckFailure(
                f"{EXCLUSION_PATH}: {field} must be a non-empty string"
            )
    packages = value["excluded_packages"]
    if not isinstance(packages, list) or not all(
        isinstance(package, str) and package for package in packages
    ):
        raise checker.CheckFailure(
            f"{EXCLUSION_PATH}: excluded_packages must be an array of names"
        )
    try:
        names = tuple(
            checker._validate_name(package, "excluded_packages item")
            for package in packages
        )
    except checker.CheckFailure as error:
        raise checker.CheckFailure(
            f"{EXCLUSION_PATH}: malformed excluded package name: {error}"
        ) from error
    if len(set(names)) != len(names) or tuple(sorted(names)) != names:
        raise checker.CheckFailure(
            f"{EXCLUSION_PATH}: excluded_packages must be unique and sorted"
        )
    return names


def _is_iwork_package(name: str) -> bool:
    return any(
        name == family or name.startswith(f"{family}-")
        for family in (
            "litchi-iwa",
            "litchi-keynote",
            "litchi-numbers",
            "litchi-pages",
        )
    )


def _workspace_packages() -> list[tuple[checker.Package, dict[str, Any]]]:
    member_dirs, workspace_package = checker._workspace_members(REPO)
    packages: list[tuple[checker.Package, dict[str, Any]]] = []
    seen_ids: dict[str, Path] = {}
    for directory in member_dirs:
        manifest = checker._canonical_existing(
            directory / "Cargo.toml",
            REPO,
            "workspace member Cargo.toml",
            kind="file",
        )
        package, document = checker._parse_package(
            manifest, REPO, workspace_package
        )
        previous = seen_ids.get(package.id)
        if previous is not None:
            raise checker.CheckFailure(
                f"duplicate workspace package ID {package.id!r}: "
                f"{checker._context_path(previous, REPO)}, "
                f"{checker._context_path(manifest, REPO)}"
            )
        seen_ids[package.id] = manifest
        packages.append((package, document))
    return sorted(packages, key=lambda item: item[0].id)


def _collision_diagnostics(
    targets: list[checker.ExampleTarget],
) -> list[str]:
    by_name: defaultdict[str, list[checker.ExampleTarget]] = defaultdict(list)
    for target in targets:
        by_name[target.name].append(target)
    diagnostics: list[str] = []
    for name in sorted(by_name):
        records = sorted(by_name[name], key=checker.ExampleTarget.sort_key)
        package_ids = {record.package.id for record in records}
        if len(package_ids) < 2:
            continue
        has_iwork = any(_is_iwork_package(record.package.name) for record in records)
        has_retained = any(
            not _is_iwork_package(record.package.name) for record in records
        )
        if not has_retained:
            continue
        details = "; ".join(
            f"{record.package.id}: {checker._context_path(record.source, REPO)}"
            for record in records
        )
        kind = "mixed" if has_iwork else "cross-package"
        diagnostics.append(f"{kind} duplicate example target {name!r}: {details}")
    return diagnostics


def check_non_iwork_workspace() -> tuple[tuple[checker.Package, int], ...]:
    inventory = set(_read_exclusion_inventory())
    packages = _workspace_packages()
    actual_iwork = {
        package.name for package, _document in packages if _is_iwork_package(package.name)
    }
    if actual_iwork != inventory:
        missing = sorted(actual_iwork - inventory)
        stale = sorted(inventory - actual_iwork)
        diagnostics = []
        if missing:
            diagnostics.append(f"iWork packages missing from exclusion inventory: {missing}")
        if stale:
            diagnostics.append(f"excluded packages are not workspace iWork packages: {stale}")
        raise checker.CheckFailure(diagnostics)

    all_targets: list[checker.ExampleTarget] = []
    retained_counts: list[tuple[checker.Package, int]] = []
    for package, document in packages:
        # Resolve every manifest, including excluded packages, so malformed
        # metadata and a collision crossing the scope boundary cannot hide.
        targets = checker._package_examples(package, document, REPO)
        all_targets.extend(targets)
        if package.name not in inventory:
            retained_counts.append((package, len(targets)))

    diagnostics = _collision_diagnostics(all_targets)
    if diagnostics:
        raise checker.CheckFailure(diagnostics)
    return tuple(sorted(retained_counts, key=lambda item: item[0].id))


def main() -> int:
    try:
        retained = check_non_iwork_workspace()
    except checker.CheckFailure as error:
        for diagnostic in error.diagnostics:
            print(f"non-iwork-example-check: {diagnostic}", file=sys.stderr)
        return 1

    inventory = _read_exclusion_inventory()
    print("non-iwork example-target check: passed")
    print(f"excluded package count: {len(inventory)}")
    print(f"excluded packages: {', '.join(inventory)}")
    print(f"kept package count: {len(retained)}")
    print(f"kept example target count: {sum(count for _package, count in retained)}")
    for package, count in retained:
        print(f"kept targets {package.id}: {count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
