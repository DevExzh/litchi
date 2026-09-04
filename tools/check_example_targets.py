#!/usr/bin/env python3
"""Reject ambiguous Cargo workspace example targets without invoking Cargo."""

from __future__ import annotations

import glob
import os
import re
import sys
import tomllib
from collections import defaultdict
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path, PureWindowsPath
from typing import Any


ROOT = Path(__file__).resolve().parents[1]


class CheckFailure(Exception):
    """A deterministic, fail-closed manifest validation failure."""

    def __init__(self, diagnostics: str | list[str] | tuple[str, ...]) -> None:
        if isinstance(diagnostics, str):
            diagnostics = (diagnostics,)
        self.diagnostics = tuple(sorted(set(diagnostics)))
        super().__init__("\n".join(self.diagnostics))


@dataclass(frozen=True)
class Package:
    """A workspace package identified by its canonical manifest path and name."""

    id: str
    name: str
    root: Path
    manifest: Path


@dataclass(frozen=True)
class ExampleTarget:
    """One resolved example target and its canonical source path."""

    package: Package
    name: str
    source: Path
    origin: str

    def sort_key(self) -> tuple[str, str, str, str]:
        return (
            self.name,
            self.package.id,
            self.source.as_posix(),
            self.origin,
        )


def _context_path(path: Path, root: Path) -> str:
    try:
        return path.relative_to(root).as_posix() or "."
    except ValueError:
        return path.as_posix()


def _is_relative_path(value: str) -> bool:
    return not Path(value).is_absolute() and not PureWindowsPath(value).is_absolute()


def _require_mapping(value: Any, field: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise CheckFailure(f"{field} must be a TOML table")
    return value


def _require_nonempty_string(value: Any, field: str) -> str:
    if not isinstance(value, str) or not value or "\x00" in value:
        raise CheckFailure(f"{field} must be a non-empty string")
    return value


def _require_string_list(value: Any, field: str) -> list[str]:
    if not isinstance(value, list):
        raise CheckFailure(f"{field} must be an array of strings")
    result: list[str] = []
    for index, item in enumerate(value):
        result.append(_require_nonempty_string(item, f"{field}[{index}]"))
    return result


def _validate_name(value: Any, field: str) -> str:
    name = _require_nonempty_string(value, field)
    if (
        "/" in name
        or "\\" in name
        or any(character.isspace() or ord(character) < 0x20 for character in name)
    ):
        raise CheckFailure(f"{field} must be a Cargo target/package name")
    return name


def _parse_manifest(path: Path, root: Path) -> dict[str, Any]:
    display = _context_path(path, root)
    try:
        with path.open("rb") as stream:
            value = tomllib.load(stream)
    except (OSError, UnicodeError, tomllib.TOMLDecodeError) as error:
        raise CheckFailure(f"{display}: cannot parse Cargo.toml: {error}") from error
    if not isinstance(value, dict):
        raise CheckFailure(f"{display}: Cargo.toml root must be a TOML table")
    return value


def _canonical_existing(
    path: Path,
    root: Path,
    field: str,
    *,
    kind: str | None = None,
) -> Path:
    try:
        resolved = path.resolve(strict=False)
    except (OSError, RuntimeError) as error:
        raise CheckFailure(f"{field} cannot be canonicalized: {error}") from error
    try:
        resolved.relative_to(root)
    except ValueError as error:
        raise CheckFailure(
            f"{field} escapes the workspace: {_context_path(resolved, root)}"
        ) from error
    if not resolved.exists():
        raise CheckFailure(f"{field} does not exist: {_context_path(resolved, root)}")
    if kind == "directory" and not resolved.is_dir():
        raise CheckFailure(
            f"{field} must be a directory: {_context_path(resolved, root)}"
        )
    if kind == "file" and not resolved.is_file():
        raise CheckFailure(f"{field} must be a file: {_context_path(resolved, root)}")
    return resolved


def _validate_pattern(pattern: Any, field: str) -> str:
    value = _require_nonempty_string(pattern, field)
    if not _is_relative_path(value):
        raise CheckFailure(f"{field} must be a relative workspace glob")
    bracket_depth = 0
    for character in value:
        if character == "[":
            bracket_depth += 1
        elif character == "]":
            if bracket_depth == 0:
                raise CheckFailure(f"{field} has an unmatched closing bracket")
            bracket_depth -= 1
    if bracket_depth:
        raise CheckFailure(f"{field} has an unmatched opening bracket")
    return value


def _expand_workspace_pattern(pattern: Any, root: Path, field: str) -> list[Path]:
    value = _validate_pattern(pattern, field)
    try:
        matches = glob.glob(value, root_dir=os.fspath(root), recursive=True)
    except (OSError, re.error) as error:
        raise CheckFailure(f"{field} is not a valid workspace glob: {error}") from error
    if not matches:
        raise CheckFailure(f"{field} matched no paths: {value!r}")
    result: list[Path] = []
    seen: set[Path] = set()
    for relative in sorted(matches):
        path = _canonical_existing(root / relative, root, field)
        if path in seen:
            raise CheckFailure(
                f"{field} resolves the same canonical path more than once: "
                f"{_context_path(path, root)}"
            )
        seen.add(path)
        result.append(path)
    return result


def _validate_workspace_package(workspace_package: Any) -> Mapping[str, Any]:
    if workspace_package is None:
        return {}
    table = _require_mapping(workspace_package, "workspace.package")
    for field in ("version", "edition"):
        if field in table:
            _require_nonempty_string(table[field], f"workspace.package.{field}")
    return table


def _validate_inherited_string(
    value: Any,
    field: str,
    workspace_package: Mapping[str, Any],
    key: str,
) -> str:
    if isinstance(value, str):
        return _require_nonempty_string(value, field)
    if not isinstance(value, Mapping):
        raise CheckFailure(f"{field} must be a string or {{ workspace = true }}")
    if set(value) != {"workspace"} or value.get("workspace") is not True:
        raise CheckFailure(f"{field} must be a string or {{ workspace = true }}")
    inherited = workspace_package.get(key)
    if not isinstance(inherited, str) or not inherited:
        raise CheckFailure(
            f"{field} inherits from missing or invalid workspace.package.{field}"
        )
    return inherited


def _parse_package(
    manifest: Path,
    root: Path,
    workspace_package: Mapping[str, Any],
) -> tuple[Package, dict[str, Any]]:
    document = _parse_manifest(manifest, root)
    context = _context_path(manifest, root)
    package_table = _require_mapping(document.get("package"), f"{context} [package]")
    name = _validate_name(package_table.get("name"), f"{context} package.name")
    _validate_inherited_string(
        package_table.get("version"),
        f"{context} package.version",
        workspace_package,
        "version",
    )
    if "edition" in package_table:
        _validate_inherited_string(
            package_table["edition"],
            f"{context} package.edition",
            workspace_package,
            "edition",
        )
    if "autoexamples" in package_table and not isinstance(
        package_table["autoexamples"], bool
    ):
        raise CheckFailure(f"{context} package.autoexamples must be a boolean")
    package_workspace = package_table.get("workspace")
    if package_workspace is not None:
        value = _require_nonempty_string(
            package_workspace, f"{context} package.workspace"
        )
        if not _is_relative_path(value):
            raise CheckFailure(f"{context} package.workspace must be a relative path")
        workspace_root = _canonical_existing(
            manifest.parent / value,
            root,
            f"{context} package.workspace",
            kind="directory",
        )
        if workspace_root != root:
            raise CheckFailure(
                f"{context} package.workspace does not point to the workspace root"
            )
    package_root = manifest.parent
    package_id = f"{name}@{_context_path(manifest, root)}"
    return Package(package_id, name, package_root, manifest), document


def _workspace_members(root: Path) -> tuple[list[Path], Mapping[str, Any]]:
    root_manifest = _canonical_existing(
        root / "Cargo.toml", root, "root Cargo.toml", kind="file"
    )
    document = _parse_manifest(root_manifest, root)
    workspace = _require_mapping(document.get("workspace"), "workspace")
    members = _require_string_list(workspace.get("members"), "workspace.members")
    excludes = _require_string_list(workspace.get("exclude", []), "workspace.exclude")
    workspace_package = _validate_workspace_package(workspace.get("package"))

    member_dirs: list[Path] = []
    seen_members: dict[Path, str] = {}
    for index, pattern in enumerate(members):
        field = f"workspace.members[{index}]"
        for directory in _expand_workspace_pattern(pattern, root, field):
            if not directory.is_dir():
                raise CheckFailure(
                    f"{field} must match package directories: "
                    f"{_context_path(directory, root)}"
                )
            previous = seen_members.get(directory)
            if previous is not None:
                raise CheckFailure(
                    f"workspace member path is listed more than once: "
                    f"{_context_path(directory, root)} ({previous}, {field})"
                )
            seen_members[directory] = field
            manifest = _canonical_existing(
                directory / "Cargo.toml",
                root,
                f"{field} Cargo.toml",
                kind="file",
            )
            if manifest.parent != directory:
                raise CheckFailure(
                    f"{field} Cargo.toml canonical path escapes its member directory"
                )
            member_dirs.append(directory)

    excluded_dirs: set[Path] = set()
    seen_excludes: dict[Path, str] = {}
    for index, pattern in enumerate(excludes):
        field = f"workspace.exclude[{index}]"
        for directory in _expand_workspace_pattern(pattern, root, field):
            if not directory.is_dir():
                raise CheckFailure(
                    f"{field} must match package directories: "
                    f"{_context_path(directory, root)}"
                )
            previous = seen_excludes.get(directory)
            if previous is not None:
                raise CheckFailure(
                    f"workspace exclude path is listed more than once: "
                    f"{_context_path(directory, root)} ({previous}, {field})"
                )
            seen_excludes[directory] = field
            excluded_dirs.add(directory)

    member_dirs = [
        directory for directory in member_dirs if directory not in excluded_dirs
    ]
    if not member_dirs:
        raise CheckFailure("workspace has no package members after excludes")
    return sorted(member_dirs), workspace_package


def _explicit_examples(
    package: Package,
    document: Mapping[str, Any],
    root: Path,
) -> list[ExampleTarget]:
    value = document.get("example", [])
    if not isinstance(value, list):
        raise CheckFailure(
            f"{_context_path(package.manifest, root)} example targets must be an array"
        )
    result: list[ExampleTarget] = []
    for index, target_value in enumerate(value):
        field = f"{_context_path(package.manifest, root)} example[{index}]"
        target = _require_mapping(target_value, field)
        name = _validate_name(target.get("name"), f"{field}.name")
        raw_path = target.get("path")
        if raw_path is not None:
            raw_path = _require_nonempty_string(raw_path, f"{field}.path")
            if not _is_relative_path(raw_path):
                raise CheckFailure(f"{field}.path must be a relative path")
            source = _canonical_existing(
                package.root / raw_path, root, f"{field}.path", kind="file"
            )
        else:
            candidates = [
                package.root / "examples" / f"{name}.rs",
                package.root / "examples" / name / "main.rs",
            ]
            existing = [candidate for candidate in candidates if candidate.exists()]
            if not existing:
                raise CheckFailure(
                    f"{field}.path resolves to no source file for example {name!r}"
                )
            if len(existing) > 1:
                paths = ", ".join(_context_path(path, root) for path in existing)
                raise CheckFailure(
                    f"{field} has ambiguous default source paths: {paths}"
                )
            source = _canonical_existing(
                existing[0], root, f"{field}.path", kind="file"
            )
        if "required-features" in target:
            _require_string_list(
                target["required-features"], f"{field}.required-features"
            )
        result.append(ExampleTarget(package, name, source, "explicit"))
    return result


def _auto_examples(
    package: Package,
    document: Mapping[str, Any],
    root: Path,
) -> list[ExampleTarget]:
    package_table = _require_mapping(
        document.get("package"), f"{_context_path(package.manifest, root)} [package]"
    )
    if package_table.get("autoexamples", True) is False:
        return []
    examples_dir = package.root / "examples"
    if not examples_dir.exists() and not examples_dir.is_symlink():
        return []
    examples_dir = _canonical_existing(
        examples_dir,
        root,
        f"{_context_path(package.manifest, root)} examples",
        kind="directory",
    )
    result: list[ExampleTarget] = []
    top_level = sorted(glob.glob("*.rs", root_dir=os.fspath(examples_dir)))
    nested = sorted(glob.glob("*/main.rs", root_dir=os.fspath(examples_dir)))
    for relative in top_level:
        source = _canonical_existing(
            examples_dir / relative, root, "auto example source", kind="file"
        )
        result.append(ExampleTarget(package, Path(relative).stem, source, "auto"))
    for relative in nested:
        source = _canonical_existing(
            examples_dir / relative, root, "auto example source", kind="file"
        )
        result.append(
            ExampleTarget(package, Path(relative).parent.name, source, "auto")
        )
    return result


def _package_examples(
    package: Package,
    document: Mapping[str, Any],
    root: Path,
) -> list[ExampleTarget]:
    explicit = _explicit_examples(package, document, root)
    automatic = _auto_examples(package, document, root)
    by_name: defaultdict[str, list[ExampleTarget]] = defaultdict(list)
    for target in [*explicit, *automatic]:
        by_name[target.name].append(target)

    diagnostics: list[str] = []
    merged: list[ExampleTarget] = []
    for name in sorted(by_name):
        records = sorted(by_name[name], key=ExampleTarget.sort_key)
        explicit_records = [record for record in records if record.origin == "explicit"]
        automatic_records = [record for record in records if record.origin == "auto"]
        if len(explicit_records) > 1:
            details = "; ".join(
                f"{record.origin}:{_context_path(record.source, root)}"
                for record in explicit_records
            )
            diagnostics.append(
                f"same-package duplicate example target {name!r} in "
                f"{package.id}: {details}"
            )
            continue
        if explicit_records:
            # Cargo's explicit target declaration owns the name and suppresses
            # an auto-discovered target with the same name, even when its path
            # differs.
            merged.append(explicit_records[0])
            continue
        if len(automatic_records) > 1:
            details = "; ".join(
                f"{record.origin}:{_context_path(record.source, root)}"
                for record in automatic_records
            )
            diagnostics.append(
                f"same-package duplicate example target {name!r} in "
                f"{package.id}: {details}"
            )
            continue
        merged.append(automatic_records[0])
    if diagnostics:
        raise CheckFailure(diagnostics)
    return sorted(merged, key=ExampleTarget.sort_key)


def _duplicate_workspace_targets(
    targets: list[ExampleTarget],
    root: Path,
) -> list[str]:
    by_name: defaultdict[str, list[ExampleTarget]] = defaultdict(list)
    for target in targets:
        by_name[target.name].append(target)
    diagnostics: list[str] = []
    for name in sorted(by_name):
        records = sorted(by_name[name], key=ExampleTarget.sort_key)
        package_ids = {record.package.id for record in records}
        if len(package_ids) < 2:
            continue
        details = "; ".join(
            f"{record.package.id}: {_context_path(record.source, root)}"
            for record in records
        )
        diagnostics.append(
            f"cross-package duplicate example target {name!r}: {details}"
        )
    return diagnostics


def check_workspace(root: Path = ROOT) -> tuple[ExampleTarget, ...]:
    """Validate manifests and return all resolved example targets."""
    try:
        root = root.resolve(strict=True)
    except (OSError, RuntimeError) as error:
        raise CheckFailure(
            f"workspace root cannot be canonicalized: {error}"
        ) from error
    if not root.is_dir():
        raise CheckFailure(f"workspace root must be a directory: {root}")
    member_dirs, workspace_package = _workspace_members(root)
    packages: list[tuple[Package, Mapping[str, Any]]] = []
    seen_ids: dict[str, Path] = {}
    for directory in member_dirs:
        manifest = _canonical_existing(
            directory / "Cargo.toml", root, "workspace member Cargo.toml", kind="file"
        )
        package, document = _parse_package(manifest, root, workspace_package)
        previous = seen_ids.get(package.id)
        if previous is not None:
            raise CheckFailure(
                f"duplicate workspace package ID {package.id!r}: "
                f"{_context_path(previous, root)}, {_context_path(manifest, root)}"
            )
        seen_ids[package.id] = manifest
        packages.append((package, document))

    targets: list[ExampleTarget] = []
    for package, document in sorted(packages, key=lambda item: item[0].id):
        targets.extend(_package_examples(package, document, root))
    diagnostics = _duplicate_workspace_targets(targets, root)
    if diagnostics:
        raise CheckFailure(diagnostics)
    return tuple(sorted(targets, key=ExampleTarget.sort_key))


def main() -> int:
    try:
        check_workspace(ROOT)
    except CheckFailure as error:
        for diagnostic in error.diagnostics:
            print(f"example-target-check: {diagnostic}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
