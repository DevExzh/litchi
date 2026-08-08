#!/usr/bin/env python3
"""Reject implementation details from the public ``litchi::iwork`` API.

The check uses rustdoc JSON rather than grepping Rust source.  That makes
aliases, re-exports, associated items, and resolved dependency types visible
to the gate while ignoring private implementation details.
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
DEFAULT_JSON = ROOT / "target" / "doc" / "litchi.json"
IWORK_PATH = ("litchi", "iwork")
ISOLATION_FEATURES = ("pages", "keynote", "numbers")
FORBIDDEN_CRATES = frozenset(
    {
        "buffa",
        "litchi_iwa",
        "litchi_iwa_archive",
        "litchi_iwa_common",
        "litchi_iwa_core",
        "litchi_iwa_detect",
        "litchi_iwa_protos",
        "litchi_iwa_structured",
        "litchi_iwa_text",
        "litchi_iwa_text_wire",
        "litchi_keynote",
        "litchi_numbers",
        "litchi_numbers_wire",
        "litchi_pages",
        "prost",
        "prost_types",
    }
)
ALLOWED_EXTERNAL_CRATES = frozenset({"alloc", "core", "std"})
RAW_ID = re.compile(r"^(?:id|ids|[a-z][a-z0-9_]*_(?:id|ids))$", re.IGNORECASE)
CAMEL_RAW_ID = re.compile(r"^(?:Id|Ids|[A-Za-z][A-Za-z0-9]*Ids?)$")
NAME_TOKEN = re.compile(r"[A-Z]+(?=[A-Z][a-z]|$)|[A-Z]?[a-z]+|[0-9]+")
CAPABILITY_TOKENS = frozenset({"archive", "catalog", "component", "components", "prepared", "raw"})


def rustdoc_command() -> tuple[str, ...]:
    """Return the root-facade rustdoc invocation used by the gate.

    This library workspace intentionally does not version ``Cargo.lock``, so
    ``--locked`` would make the gate fail in every clean checkout.
    """
    return (
        "cargo",
        "rustdoc",
        "--package",
        "litchi",
        "--no-default-features",
        "--features",
        "iwork",
        "--",
        "-Zunstable-options",
        "--output-format",
        "json",
    )


def isolation_rustdoc_command() -> tuple[str, ...]:
    """Compile all leaf facades without enabling the aggregate facade."""
    return (
        "cargo",
        "rustdoc",
        "--package",
        "litchi",
        "--no-default-features",
        "--features",
        ",".join(ISOLATION_FEATURES),
        "--",
        "-Zunstable-options",
        "--output-format",
        "json",
    )


def environment(source: Mapping[str, str] | None = None) -> dict[str, str]:
    """Enable rustdoc JSON without mutating the caller's environment."""
    result = dict(os.environ if source is None else source)
    result["RUSTC_BOOTSTRAP"] = "1"
    return result


def _identifier(value: Any) -> str:
    return str(value)


def _referenced_ids(value: Any) -> Iterable[str]:
    """Yield rustdoc item IDs referenced by a JSON value.

    Rustdoc uses ``id`` for resolved paths and relation arrays for module
    children, fields, variants, associated items, and implementations.
    """
    if isinstance(value, dict):
        for key, nested in value.items():
            if key == "id" and isinstance(nested, (int, str)):
                yield _identifier(nested)
            elif key in {"fields", "impls", "items", "variants"} and isinstance(
                nested, list
            ):
                for item_id in nested:
                    if isinstance(item_id, (int, str)):
                        yield _identifier(item_id)
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


def _path_entry_path(entry: Any) -> tuple[str, ...]:
    if not isinstance(entry, dict):
        return ()
    path = entry.get("path")
    if not isinstance(path, list) or not all(isinstance(part, str) for part in path):
        return ()
    return tuple(path)


def _name_violation(identifier: str) -> str | None:
    """Classify public names that expose physical-package capabilities."""
    if RAW_ID.fullmatch(identifier) or CAMEL_RAW_ID.fullmatch(identifier):
        return "raw identifier"

    words: list[str] = []
    for part in re.split(r"[^A-Za-z0-9]+", identifier):
        words.extend(word.lower() for word in NAME_TOKEN.findall(part))
    if "identifier" in words or "identifiers" in words:
        return "raw identifier"
    if any(
        words[index] == "native" and words[index + 1] in {"id", "ids"}
        for index in range(len(words) - 1)
    ):
        return "raw identifier"
    capability = next((word for word in words if word in CAPABILITY_TOKENS), None)
    if capability is not None:
        return f"implementation capability `{capability}`"
    if any(
        words[index : index + 2] == ["source", "bytes"]
        for index in range(len(words) - 1)
    ):
        return "implementation capability `source_bytes`"
    return None


def _is_blanket_impl(item: Any) -> bool:
    """Return whether rustdoc synthesized an implementation from a blanket.

    Rustdoc attaches every dependency blanket implementation to each public
    facade type. Those implementations are not authored API and recursively
    traversing them reports unrelated crates such as allocator internals.
    Explicit facade-owned trait implementations remain in scope.
    """
    if not isinstance(item, dict):
        return False
    inner = item.get("inner")
    if not isinstance(inner, dict):
        return False
    implementation = inner.get("impl")
    return isinstance(implementation, dict) and implementation.get("blanket_impl") is not None


def violations(document: Mapping[str, Any]) -> list[str]:
    """Return deterministic public-surface violations from rustdoc JSON."""
    index_value = document.get("index")
    paths_value = document.get("paths")
    if not isinstance(index_value, dict) or not isinstance(paths_value, dict):
        return ["invalid rustdoc JSON: expected object-valued `index` and `paths`"]

    index = {_identifier(key): value for key, value in index_value.items()}
    paths = {_identifier(key): value for key, value in paths_value.items()}
    root_id = _identifier(document.get("root", ""))
    root_item = index.get(root_id)
    root_crate_id = root_item.get("crate_id") if isinstance(root_item, dict) else None
    iwork_ids = sorted(
        item_id
        for item_id, entry in paths.items()
        if _path_entry_path(entry) == IWORK_PATH
    )
    if not iwork_ids:
        return ["missing public module `litchi::iwork`"]

    failures: set[str] = set()
    pending = list(iwork_ids)
    visited: set[str] = set()
    while pending:
        item_id = pending.pop()
        if item_id in visited:
            continue
        if _is_blanket_impl(index.get(item_id)):
            continue
        visited.add(item_id)
        item = index.get(item_id)
        if not isinstance(item, dict):
            continue

        item_path = _path_entry_path(paths.get(item_id))
        display = "::".join(item_path) if item_path else f"rustdoc item {item_id}"
        names = []
        name = item.get("name")
        if isinstance(name, str):
            names.append(name)
        names.extend(_argument_names(item))
        for identifier in names:
            reason = _name_violation(identifier)
            if reason is not None:
                failures.add(f"{display} exposes {reason} as `{identifier}`")

        for referenced_id in _referenced_ids(item):
            referenced_path = _path_entry_path(paths.get(referenced_id))
            if referenced_path:
                crate_name = referenced_path[0].replace("-", "_")
                if crate_name in FORBIDDEN_CRATES:
                    failures.add(
                        f"{display} exposes forbidden type `{'::'.join(referenced_path)}`"
                    )
                elif crate_name == IWORK_PATH[0] and referenced_path[
                    : len(IWORK_PATH)
                ] != IWORK_PATH:
                    failures.add(
                        f"{display} exposes type outside the iWork facade "
                        f"`{'::'.join(referenced_path)}`"
                    )
                elif (
                    crate_name != IWORK_PATH[0]
                    and crate_name not in ALLOWED_EXTERNAL_CRATES
                    and paths[referenced_id].get("kind") != "primitive"
                ):
                    failures.add(
                        f"{display} exposes type from non-allowlisted crate "
                        f"`{'::'.join(referenced_path)}`"
                    )
                referenced_name = referenced_path[-1]
                reason = _name_violation(referenced_name)
                if reason is not None:
                    failures.add(
                        f"{display} exposes {reason} type `{'::'.join(referenced_path)}`"
                    )
                if referenced_path[: len(IWORK_PATH)] == IWORK_PATH:
                    pending.append(referenced_id)
            elif referenced_id in index:
                # Module children and associated items do not always receive a
                # path-table entry. Only locally-authored items belong to this
                # facade traversal; dependency trait internals and blanket
                # implementations are implementation noise.
                referenced_item = index[referenced_id]
                if (
                    isinstance(referenced_item, dict)
                    and referenced_item.get("crate_id") == root_crate_id
                    and not _is_blanket_impl(referenced_item)
                ):
                    pending.append(referenced_id)

    return sorted(failures)


def isolation_violations(document: Mapping[str, Any]) -> list[str]:
    """Reject the aggregate module when only concrete leaf features are on."""
    paths_value = document.get("paths")
    if not isinstance(paths_value, dict):
        return ["invalid rustdoc JSON: expected object-valued `paths`"]
    if any(_path_entry_path(entry) == IWORK_PATH for entry in paths_value.values()):
        return [
            "public module `litchi::iwork` is available without the `iwork` feature"
        ]
    return []


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
        help="inspect existing rustdoc JSON instead of building litchi",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    json_path = args.json
    if json_path is None:
        completed = subprocess.run(
            rustdoc_command(),
            cwd=ROOT,
            env=environment(),
            check=False,
        )
        if completed.returncode != 0:
            return completed.returncode
        json_path = DEFAULT_JSON

    document = load_document(json_path)

    failures = violations(document)
    if failures:
        print("litchi::iwork public API violations:", file=sys.stderr)
        for failure in failures:
            print(f"- {failure}", file=sys.stderr)
        return 1

    checked_isolation = args.json is None
    if checked_isolation:
        completed = subprocess.run(
            isolation_rustdoc_command(),
            cwd=ROOT,
            env=environment(),
            check=False,
        )
        if completed.returncode != 0:
            return completed.returncode
        isolation_failures = isolation_violations(load_document(DEFAULT_JSON))
        if isolation_failures:
            print("litchi::iwork feature-isolation violations:", file=sys.stderr)
            for failure in isolation_failures:
                print(f"- {failure}", file=sys.stderr)
            return 1

    message = "litchi::iwork public API contains no forbidden implementation types or raw IDs"
    if checked_isolation:
        message += "; leaf features do not publish the aggregate facade"
    print(message)
    return 0


if __name__ == "__main__":
    sys.exit(main())
