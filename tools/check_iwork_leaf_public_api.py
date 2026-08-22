#!/usr/bin/env python3
"""Check the archive-free document surfaces of the iWork leaf crates.

The concrete iWork crates intentionally have a physical ``Package`` API.  A
whole-crate dependency scan would therefore reject the package boundary (for
example its caller-selected archive limits) and would also classify semantic
names such as ``identifier`` as native IDs.  This gate starts at each crate's
public ``document`` module (or crate-root re-exports when that implementation
module is private), follows the semantic types reachable from that module,
and checks the resulting rustdoc surface instead.

Rustdoc JSON is used rather than source text so resolved aliases, re-exports,
public fields, and dependency types are checked consistently.  The gate does
not claim that the physical package adapters are archive-free; those adapters
are deliberately outside this narrow semantic-reader contract.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
from collections.abc import Iterable, Mapping
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[1]


LEAF_PACKAGES = (
    "litchi-numbers",
    "litchi-pages",
    "litchi-keynote",
)


LEAF_CRATE_NAMES = {package: package.replace("-", "_") for package in LEAF_PACKAGES}


DEFAULT_JSON = {
    crate_name: ROOT / "target" / "doc" / f"{crate_name}.json"
    for crate_name in LEAF_CRATE_NAMES.values()
}


# The semantic readers may publish these focused value crates.  In particular,
# litchi-iwa-text::Storage is a semantic text value even though the package
# adapters below it understand native storage records.
ALLOWED_EXTERNAL_CRATES = frozenset(
    {
        "alloc",
        "core",
        "litchi_core",
        "litchi_iwa_text",
        "serde",
        "std",
    }
)


# These dependencies contain physical archives, native object graphs, wire
# views, generated schemas, or protobuf runtimes.  A resolved public reference
# to one of them is a leak even when a facade gives the type a semantic alias.
FORBIDDEN_CRATES = frozenset(
    {
        "buffa",
        "litchi_iwa",
        "litchi_iwa_archive",
        "litchi_iwa_common",
        "litchi_iwa_core",
        "litchi_iwa_detect",
        "litchi_iwa_index",
        "litchi_iwa_package",
        "litchi_iwa_protos",
        "litchi_iwa_structured",
        "litchi_iwa_text_wire",
        "litchi_numbers_wire",
        "prost",
        "prost_types",
    }
)


RAW_ARGUMENT_NAMES = frozenset(
    {
        "id",
        "ids",
        "native_id",
        "native_ids",
        "object_id",
        "object_ids",
        "raw_id",
        "raw_ids",
        "source_bytes",
        "raw_bytes",
    }
)


# These are physical/generated names that should not be recreated locally and
# then hidden behind a facade alias.  Deliberately do not reject every name
# containing ``identifier`` or every ``*Id``: semantic identifiers are valid
# reader values and methods (for example Build::identifier) are not native
# object handles by themselves.
RAW_TYPE_NAMES = frozenset(
    {
        "Archive",
        "ArchiveObject",
        "ArchiveView",
        "ComponentId",
        "ComponentCatalog",
        "Generated",
        "IWorkPackage",
        "MessageId",
        "MessageInfo",
        "NativeId",
        "ObjectId",
        "RawMessage",
        "RawId",
        "SourceCatalog",
        "Wire",
        "WireDescent",
        "WireFieldView",
        "WireLimits",
        "WireView",
    }
)


NAME_TOKEN = re.compile(r"[A-Z]+(?=[A-Z][a-z]|$)|[A-Z]?[a-z]+|[0-9]+")


def rustdoc_commands() -> tuple[tuple[str, ...], ...]:
    """Return one deterministic rustdoc invocation for each semantic leaf."""
    return tuple(
        (
            "cargo",
            "rustdoc",
            "--package",
            package,
            "--no-default-features",
            "--lib",
            "--",
            "-Zunstable-options",
            "--output-format",
            "json",
        )
        for package in LEAF_PACKAGES
    )


def rustdoc_command(package: str) -> tuple[str, ...]:
    """Return the rustdoc invocation for one named leaf package."""
    try:
        index = LEAF_PACKAGES.index(package)
    except ValueError as error:
        raise ValueError(f"unsupported iWork leaf package: {package}") from error
    return rustdoc_commands()[index]


def environment(source: Mapping[str, str] | None = None) -> dict[str, str]:
    """Enable rustdoc JSON without mutating the caller's environment."""
    result = dict(os.environ if source is None else source)
    result["RUSTC_BOOTSTRAP"] = "1"
    return result


def _identifier(value: Any) -> str:
    return str(value)


def _path_entry_path(entry: Any) -> tuple[str, ...]:
    if not isinstance(entry, dict):
        return ()
    path = entry.get("path")
    if not isinstance(path, list) or not all(isinstance(part, str) for part in path):
        return ()
    return tuple(path)


def _referenced_ids(value: Any) -> Iterable[str]:
    """Yield rustdoc item IDs referenced by a JSON value."""
    if isinstance(value, dict):
        for key, nested in value.items():
            if key == "id" and isinstance(nested, (int, str)):
                yield _identifier(nested)
            elif key in {
                "fields",
                "impls",
                "items",
                "tuple",
                "variants",
            } and isinstance(nested, list):
                for item_id in nested:
                    if isinstance(item_id, (int, str)):
                        yield _identifier(item_id)
                    else:
                        # Rustdoc normally stores field/variant/tuple
                        # references as scalar IDs, but newer schemas can
                        # inline a descriptor at any of these positions.
                        # Keep walking the descriptor so a nested tuple type
                        # cannot hide a physical dependency.
                        yield from _referenced_ids(item_id)
            else:
                yield from _referenced_ids(nested)
    elif isinstance(value, list):
        for nested in value:
            yield from _referenced_ids(nested)


def _argument_names(item: Mapping[str, Any]) -> Iterable[str]:
    inner = item.get("inner")
    if not isinstance(inner, dict):
        return
    function = inner.get("function")
    if not isinstance(function, dict):
        return
    declaration = function.get("sig") or function.get("decl")
    if not isinstance(declaration, dict):
        return
    inputs = declaration.get("inputs", [])
    if not isinstance(inputs, list):
        return
    for input_value in inputs:
        if (
            isinstance(input_value, list)
            and input_value
            and isinstance(input_value[0], str)
        ):
            yield input_value[0]


def _use_target_id(item: Mapping[str, Any]) -> str | None:
    inner = item.get("inner")
    if not isinstance(inner, dict):
        return None
    use = inner.get("use")
    if not isinstance(use, dict):
        return None
    value = use.get("id")
    return _identifier(value) if isinstance(value, (int, str)) else None


def _is_blanket_impl(item: Any) -> bool:
    """Ignore dependency blanket implementations synthesized by rustdoc."""
    if not isinstance(item, dict):
        return False
    inner = item.get("inner")
    if not isinstance(inner, dict):
        return False
    implementation = inner.get("impl")
    return (
        isinstance(implementation, dict)
        and implementation.get("blanket_impl") is not None
    )


def _is_keynote_selector_error_bridge(
    item: Any, paths: Mapping[str, Any], crate_name: str
) -> bool:
    """Ignore Keynote's generated selector-to-package error bridge.

    ``EditError::Selector(#[from])`` generates an ``automatically_derived``
    ``From<SlideSelectorError> for EditError`` implementation.  Rustdoc emits
    that implementation without a path entry, so a semantic graph walk that
    follows all local implementation IDs can reach the physical package error
    even though no document method has that error in its signature.  Keep this
    exception exact: only this generated implementation in the Keynote crate
    is outside the semantic-reader graph.  A direct ``EditError`` reference
    still goes through normal path validation.
    """
    if crate_name != "litchi_keynote" or not isinstance(item, dict):
        return False
    attrs = item.get("attrs")
    if not isinstance(attrs, list) or "automatically_derived" not in attrs:
        return False
    inner = item.get("inner")
    if not isinstance(inner, dict):
        return False
    implementation = inner.get("impl")
    if not isinstance(implementation, dict):
        return False

    trait = implementation.get("trait")
    if not isinstance(trait, dict) or trait.get("path") != "From":
        return False
    trait_args = trait.get("args")
    if not isinstance(trait_args, dict):
        return False
    angle_bracketed = trait_args.get("angle_bracketed")
    if not isinstance(angle_bracketed, dict):
        return False
    args = angle_bracketed.get("args")
    if not isinstance(args, list) or len(args) != 1:
        return False
    argument = args[0]
    if not isinstance(argument, dict):
        return False
    argument_type = argument.get("type")
    if not isinstance(argument_type, dict):
        return False
    selector = argument_type.get("resolved_path")
    target = implementation.get("for")
    if not isinstance(selector, dict) or not isinstance(target, dict):
        return False
    selector_id = selector.get("id")
    target_type = target.get("resolved_path")
    target_id = target_type.get("id") if isinstance(target_type, dict) else None
    selector_path = _path_entry_path(paths.get(_identifier(selector_id)))
    target_path = _path_entry_path(paths.get(_identifier(target_id)))
    return selector_path == (
        crate_name,
        "selector",
        "SlideSelectorError",
    ) and target_path == (
        crate_name,
        "package",
        "edit",
        "EditError",
    )


def _name_violation(identifier: str, *, argument: bool = False) -> str | None:
    """Classify physical names without rejecting semantic identifiers."""
    if argument and identifier in RAW_ARGUMENT_NAMES:
        if identifier in {"source_bytes", "raw_bytes"}:
            return "raw source bytes"
        return "raw identifier"
    if identifier in RAW_TYPE_NAMES:
        return "implementation type"

    words: list[str] = []
    for part in re.split(r"[^A-Za-z0-9]+", identifier):
        words.extend(word.lower() for word in NAME_TOKEN.findall(part))

    # Keep this contextual: `DocumentIdentifier`, `NodeId`, and the
    # `identifier()` accessor are semantic vocabulary, while native/object/raw
    # combinations are physical handles.
    if any(
        words[index] in {"native", "object", "raw"}
        and words[index + 1] in {"id", "ids", "identifier", "identifiers"}
        for index in range(len(words) - 1)
    ):
        return "raw identifier"
    if "generated" in words:
        return "generated type"
    if "prost" in words or "protobuf" in words:
        return "protobuf type"
    if any(word in {"wire", "archive"} for word in words):
        return "implementation type"
    return None


def _crate_path(path: tuple[str, ...]) -> str:
    return path[0].replace("-", "_") if path else ""


def _semantic_path_violation(
    path: tuple[str, ...], *, crate_name: str, kind: str | None
) -> str | None:
    """Classify one resolved type in a leaf's semantic graph."""
    if not path:
        return None
    referenced_crate = _crate_path(path)
    if referenced_crate in FORBIDDEN_CRATES:
        return f"forbidden type `{'::'.join(path)}`"
    if referenced_crate == crate_name:
        if "package" in path[1:-1]:
            return f"physical package type `{'::'.join(path)}`"
        reason = _name_violation(path[-1])
        if reason is not None:
            return f"{reason} `{'::'.join(path)}`"
        return None
    if referenced_crate not in ALLOWED_EXTERNAL_CRATES and kind != "primitive":
        return f"type from non-allowlisted crate `{'::'.join(path)}`"
    reason = _name_violation(path[-1])
    if reason is not None and referenced_crate not in {"core", "std", "alloc"}:
        return f"{reason} `{'::'.join(path)}`"
    return None


def _document_ids(
    index: Mapping[str, Any], paths: Mapping[str, Any], crate_name: str
) -> tuple[set[str], str | None]:
    """Return the public document module and its root re-export targets."""
    module_path = (crate_name, "document")
    roots = {
        item_id
        for item_id, entry in paths.items()
        if _path_entry_path(entry) == module_path
    }

    # A root `pub use document::{...}` has a root path but points at the
    # document item through `use.id`. Include the alias so a local generated
    # name cannot be hidden solely by re-exporting it. Rustdoc JSON omits a
    # path entry for these root aliases when the source module is private (as
    # litchi-pages::document intentionally is), so identify them by the
    # parent crate module rather than requiring an alias path entry.
    crate_root_id = next(
        (
            item_id
            for item_id, entry in paths.items()
            if _path_entry_path(entry) == (crate_name,)
        ),
        None,
    )
    crate_root_items: set[str] = set()
    if crate_root_id is not None:
        root_item = index.get(crate_root_id)
        if isinstance(root_item, dict):
            inner = root_item.get("inner")
            module = inner.get("module") if isinstance(inner, dict) else None
            if isinstance(module, dict):
                for item_id in module.get("items", []):
                    if isinstance(item_id, (int, str)):
                        crate_root_items.add(_identifier(item_id))

    for item_id, item in index.items():
        target = _use_target_id(item) if isinstance(item, dict) else None
        if target is None:
            continue
        target_path = _path_entry_path(paths.get(target))
        item_path = _path_entry_path(paths.get(item_id))
        is_root_alias = item_id in crate_root_items or (
            crate_root_id is None and item_path[:1] == (crate_name,)
        )
        if target_path[:2] == module_path and is_root_alias:
            roots.add(item_id)

    if not roots:
        return set(), f"missing public module `{crate_name}::document`"
    return roots, None


def violations(document: Mapping[str, Any], crate_name: str) -> list[str]:
    """Return deterministic semantic-reader violations for one leaf crate."""
    index_value = document.get("index")
    paths_value = document.get("paths")
    if not isinstance(index_value, dict) or not isinstance(paths_value, dict):
        return ["invalid rustdoc JSON: expected object-valued `index` and `paths`"]

    index = {_identifier(key): value for key, value in index_value.items()}
    paths = {_identifier(key): value for key, value in paths_value.items()}
    crate_name = crate_name.replace("-", "_")
    if crate_name not in LEAF_CRATE_NAMES.values():
        return [f"unsupported iWork leaf crate `{crate_name}`"]

    pending, missing = _document_ids(index, paths, crate_name)
    if missing is not None:
        return [missing]

    root_id = _identifier(document.get("root", ""))
    root_item = index.get(root_id)
    root_crate_id = root_item.get("crate_id") if isinstance(root_item, dict) else None

    failures: set[str] = set()
    visited: set[str] = set()
    while pending:
        item_id = pending.pop()
        if item_id in visited or _is_blanket_impl(index.get(item_id)):
            continue
        if _is_keynote_selector_error_bridge(index.get(item_id), paths, crate_name):
            continue
        visited.add(item_id)
        item = index.get(item_id)
        if not isinstance(item, dict):
            continue

        item_path = _path_entry_path(paths.get(item_id))
        display = "::".join(item_path) if item_path else f"rustdoc item {item_id}"
        name = item.get("name")
        if isinstance(name, str):
            reason = _name_violation(name)
            if reason is not None:
                failures.add(f"{display} exposes {reason} as `{name}`")
        for argument in _argument_names(item):
            reason = _name_violation(argument, argument=True)
            if reason is not None:
                failures.add(f"{display} exposes {reason} as `{argument}`")

        for referenced_id in _referenced_ids(item):
            referenced_path = _path_entry_path(paths.get(referenced_id))
            if referenced_path:
                referenced_entry = paths.get(referenced_id)
                kind = (
                    referenced_entry.get("kind")
                    if isinstance(referenced_entry, dict)
                    else None
                )
                reason = _semantic_path_violation(
                    referenced_path, crate_name=crate_name, kind=kind
                )
                if reason is not None:
                    failures.add(f"{display} exposes {reason}")

                # Follow only local semantic items.  Physical package paths
                # have already been diagnosed and are deliberately a leaf.
                if (
                    referenced_path[0] == crate_name
                    and "package" not in referenced_path[1:-1]
                ):
                    pending.add(referenced_id)
            elif referenced_id in index:
                referenced_item = index[referenced_id]
                if (
                    isinstance(referenced_item, dict)
                    and referenced_item.get("crate_id") == root_crate_id
                    and not _is_blanket_impl(referenced_item)
                ):
                    pending.add(referenced_id)

    return sorted(failures)


def load_document(json_path: Path) -> Mapping[str, Any]:
    """Read one rustdoc JSON document or terminate with a stable diagnostic."""
    try:
        return json.loads(json_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        print(f"failed to read rustdoc JSON {json_path}: {error}", file=sys.stderr)
        raise SystemExit(2) from error


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--json",
        type=Path,
        help="inspect one existing rustdoc JSON document (requires --crate)",
    )
    parser.add_argument(
        "--crate",
        choices=tuple(LEAF_CRATE_NAMES),
        help="leaf package to inspect; defaults to all leaves when building",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    if args.json is not None:
        if args.crate is None:
            print("--crate is required with --json", file=sys.stderr)
            return 2
        failures = violations(load_document(args.json), args.crate)
        if failures:
            print(f"{args.crate} semantic public API violations:", file=sys.stderr)
            for failure in failures:
                print(f"- {failure}", file=sys.stderr)
            return 1
        print(f"{args.crate} semantic public API contains no physical leaks")
        return 0

    packages = (args.crate,) if args.crate is not None else LEAF_PACKAGES
    for package in packages:
        completed = subprocess.run(
            rustdoc_command(package),
            cwd=ROOT,
            env=environment(),
            check=False,
        )
        if completed.returncode != 0:
            return completed.returncode
        crate_name = LEAF_CRATE_NAMES[package]
        failures = violations(load_document(DEFAULT_JSON[crate_name]), crate_name)
        if failures:
            print(f"{package} semantic public API violations:", file=sys.stderr)
            for failure in failures:
                print(f"- {failure}", file=sys.stderr)
            return 1

    print(
        "litchi-numbers, litchi-pages, and litchi-keynote semantic public APIs "
        "contain no raw IDs or physical archive/wire/protobuf types"
    )
    return 0


if __name__ == "__main__":
    sys.exit(main())
