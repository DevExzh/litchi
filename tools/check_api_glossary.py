#!/usr/bin/env python3
"""Check registered ADR 0034 vocabulary waves against Rust's public API.

Rules live in api_glossary.json; each migration explicitly registers its bans
and completed transaction families.
Each profile builds its own feature-specific rustdoc JSON; --json/--profile
checks an existing document instead. Names are matched exactly under a public
path prefix, never by substring. A ban with recursive=false matches direct
children of its scope, preserving unrelated nested domain vocabulary. Family
checks assert six exported type names,
not transaction semantics (which remain Rust/CRUD test obligations).
"""
from __future__ import annotations

import argparse
from contextlib import nullcontext
import json
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
POLICY = ROOT / "tools/api_glossary.json"
FAMILY_NAMES = frozenset({"Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind"})
TYPE_KINDS = frozenset({"struct", "enum", "union", "type_alias"})
IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z_0-9]*\Z")


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def checked_policy(value: Any) -> list[dict[str, Any]]:
    """Reject typos/empty scopes instead of silently disabling enforcement."""
    if not isinstance(value, dict) or set(value) != {"version", "profiles"}:
        raise ValueError("policy requires exactly version and profiles")
    if type(value["version"]) is not int or value["version"] != 1:
        raise ValueError("unsupported policy version")
    if not isinstance(value["profiles"], list):
        raise ValueError("profiles must be an array")
    seen = set()
    for profile in value["profiles"]:
        fields = {"id", "package", "default_features", "features", "banned_names", "families"}
        optional = {"delegated_globs"}
        if (not isinstance(profile, dict) or not fields <= set(profile)
                or set(profile) - (fields | optional)):
            raise ValueError("profile requires exactly " + ", ".join(sorted(fields))
                             + ", with optional delegated_globs")
        delegated = profile.get("delegated_globs", [])
        if not isinstance(delegated, list) or any(
            not isinstance(c, str) or not IDENTIFIER.fullmatch(c) for c in delegated
        ) or len(set(delegated)) != len(delegated):
            raise ValueError("delegated_globs must be unique crate names")
        for key in ("id", "package"):
            if not isinstance(profile[key], str) or not re.fullmatch(r"[a-z][a-z0-9_-]*", profile[key]):
                raise ValueError(f"invalid profile {key}")
        if profile["id"] in seen:
            raise ValueError(f"duplicate profile {profile['id']}")
        seen.add(profile["id"])
        if type(profile["default_features"]) is not bool:
            raise ValueError("default_features must be boolean")
        features = profile["features"]
        if not isinstance(features, list) or any(
            not isinstance(f, str) or not re.fullmatch(r"[A-Za-z0-9_/-]+", f) for f in features
        ) or len(set(features)) != len(features):
            raise ValueError("features must be unique Cargo feature names")
        crate = profile["package"].replace("-", "_")

        def check_path(path: Any) -> None:
            if not isinstance(path, str) or not all(IDENTIFIER.fullmatch(p) for p in path.split("::")):
                raise ValueError("scope/family must be a fully qualified Rust path")
            if path.split("::")[0] != crate:
                raise ValueError(f"scope/family must start with {crate}")

        bans = profile["banned_names"]
        if not isinstance(bans, list):
            raise ValueError("banned_names must be an array")
        keys = set()
        for ban in bans:
            required = {"name", "scope", "replacement"}
            if (not isinstance(ban, dict) or not required <= set(ban)
                    or set(ban) - (required | {"recursive"})):
                raise ValueError("ban requires name, scope and replacement, with optional recursive")
            if "recursive" in ban and type(ban["recursive"]) is not bool:
                raise ValueError("ban recursive must be boolean")
            if not isinstance(ban["name"], str) or not IDENTIFIER.fullmatch(ban["name"]):
                raise ValueError("banned name must be one Rust identifier")
            if not isinstance(ban["replacement"], str) or not ban["replacement"].strip():
                raise ValueError("ban requires a nonempty replacement explanation")
            check_path(ban["scope"])
            key = (ban["scope"], ban["name"])
            if key in keys:
                raise ValueError("duplicate banned name and scope")
            keys.add(key)
        families = profile["families"]
        if not isinstance(families, list):
            raise ValueError("families must be an array")
        for family in families:
            check_path(family)
        if len(set(families)) != len(families):
            raise ValueError("duplicate family")
    return value["profiles"]


def referenced_types(value: Any):
    """Type paths in signatures, fields, bounds and aliases (not item lists)."""
    if isinstance(value, dict):
        for key, child in value.items():
            if key == "id" and isinstance(child, (str, int)):
                # Even opaque types omitted from index/paths carry a spelling
                # at their use site. Check that name without inventing members.
                spelling = value.get("path")
                name = spelling.rsplit("::", 1)[-1] if isinstance(spelling, str) else None
                yield str(child), name
            else:
                yield from referenced_types(child)
    elif isinstance(value, list):
        for child in value:
            yield from referenced_types(child)


def field_ids(value: Any):
    if isinstance(value, dict):
        for key, child in value.items():
            if key in {"fields", "tuple"} and isinstance(child, list):
                yield from (str(i) for i in child if isinstance(i, (str, int)))
            else:
                yield from field_ids(child)


class Surface:
    """Exported paths plus reachable types; no private-module name scanning."""

    def __init__(self, document: Any, crate: str, delegated_globs: frozenset[str] = frozenset()):
        if (
            not isinstance(document, dict)
            or not isinstance(document.get("index"), dict)
            or not isinstance(document.get("paths"), dict)
        ):
            raise ValueError("invalid rustdoc JSON: expected index and paths objects")
        self.delegated_globs = delegated_globs
        self.index = {str(k): v for k, v in document["index"].items()}
        self.paths = {str(k): v for k, v in document["paths"].items()}
        for entry in self.paths.values():
            if (
                not isinstance(entry, dict)
                or not isinstance(entry.get("path"), list)
                or not entry["path"]
                or not all(isinstance(p, str) and p for p in entry["path"])
                or not isinstance(entry.get("kind"), str)
                or type(entry.get("crate_id")) is not int
            ):
                raise ValueError("invalid rustdoc path entry")
        root = str(document.get("root"))
        item = self.index.get(root, {})
        if item.get("name") != crate or "module" not in item.get("inner", {}):
            raise ValueError(f"rustdoc root must be crate {crate}")
        if self.paths.get(root, {}).get("path") != [crate]:
            raise ValueError("missing or invalid rustdoc root path")
        if type(item.get("crate_id")) is not int:
            raise ValueError("invalid rustdoc root crate_id")
        if document.get("includes_private") is not False:
            raise ValueError("rustdoc must not include private items")
        self.crate_id = item["crate_id"]
        self.exports: dict[tuple[str, ...], str] = {}
        self.names: set[tuple[tuple[str, ...], str]] = set()
        self.references: set[str] = set()
        self.walk(root, (crate,), frozenset(), exported=True)
        # Follow nominal types in signatures even if their defining module is
        # private. Do not traverse unrelated items in that private module.
        visited = set()
        while self.references - visited:
            item_id = min(self.references - visited)
            visited.add(item_id)
            path = self.paths.get(item_id, {}).get("path")
            referenced = self.index.get(item_id)
            # Rustdoc can omit generated/opaque signature types entirely.
            # A known-local path without its item is different: it would hide
            # ordinary methods/fields and must not be accepted as complete.
            if referenced is None and path and self.paths[item_id]["crate_id"] == self.crate_id:
                raise ValueError(f"missing local rustdoc type {item_id}")
            if path:
                self.names.add((tuple(path), path[-1]))
            if referenced and referenced.get("crate_id") == self.crate_id:
                name = referenced.get("name")
                if name:
                    self.walk(item_id, tuple(path or (crate, name)), frozenset(), exported=False)

    def walk(self, item_id: str, path: tuple[str, ...], ancestors: frozenset[str], *, exported: bool) -> None:
        if item_id in ancestors:
            return
        item = self.index.get(item_id)
        if not isinstance(item, dict):
            entry = self.paths.get(item_id)
            if not isinstance(entry, dict):
                raise ValueError(f"missing rustdoc item {item_id}")
            if entry["crate_id"] == self.crate_id:
                raise ValueError(f"missing local rustdoc item {item_id}")
            if exported:
                self.exports[path] = entry["kind"]
            self.names.add((path, path[-1]))
            return
        inner = item.get("inner", {})
        if len(inner) != 1:
            raise ValueError(f"invalid rustdoc item {item_id}")
        kind, data = next(iter(inner.items()))
        if kind in {
            "module", "struct", "union", "enum", "trait", "trait_alias",
            "type_alias", "function", "impl", "use", "variant", "struct_field",
            "assoc_type", "assoc_const", "constant", "static",
        } and not isinstance(data, dict):
            raise ValueError(f"invalid rustdoc {kind} payload for {item_id}")
        ancestors = ancestors | {item_id}
        if kind == "use":
            target = data.get("id")
            if target is None:
                raise ValueError(f"unresolved public re-export {'::'.join(path)}")
            if data["is_glob"]:
                # A glob exports children into its parent's namespace.
                target_item = self.index.get(str(target), {})
                module = target_item.get("inner", {}).get("module")
                if module is None:
                    # A glob over another crate a profile explicitly
                    # delegates (the facade's `pub use litchi_pages::*`) is
                    # that crate's own profile's to check; any other glob
                    # without module data is refused, not skipped.
                    entry = self.paths.get(str(target))
                    if (isinstance(entry, dict) and entry.get("crate_id") != self.crate_id
                            and entry.get("path") and entry["path"][0] in self.delegated_globs):
                        return
                    raise ValueError("glob target requires module rustdoc data: "
                                     + "::".join((entry or {}).get("path") or [str(target)]))
                self.children(module["items"], path[:-1], ancestors, exported=exported)
            else:
                alias = path[:-1] + (data["name"],)
                self.walk(str(target), alias, ancestors, exported=exported)
            return
        if kind == "impl":
            if data.get("blanket_impl") is not None or data.get("is_synthetic") or item.get("span") is None:
                return
            self.children(
                data["items"], path, ancestors, exported=exported,
                defaults=data.get("trait") is not None,
            )
            return
        if exported:
            self.exports[path] = kind
        self.names.add((path, path[-1]))
        # Also diagnose nominal old types behind re-export aliases.
        if kind in TYPE_KINDS and item.get("name"):
            self.names.add((path, item["name"]))
        if kind == "module":
            self.children(data["items"], path, ancestors, exported=exported)
            return
        if kind == "type_alias":
            # Inherent methods are callable through a type alias as well as
            # through its nominal type. Preserve the caller-visible scope.
            target = data.get("type", {}).get("resolved_path", {}).get("id")
            if target is not None:
                self.walk(str(target), path, ancestors, exported=exported)
        if isinstance(data, dict):
            for referenced_id, spelling in referenced_types(data):
                self.references.add(referenced_id)
                reference_path = self.paths.get(referenced_id, {}).get("path")
                reference_name = (
                    reference_path[-1] if reference_path
                    else self.index.get(referenced_id, {}).get("name") or spelling
                )
                if reference_name:
                    # Attribute a signature's vocabulary to the public
                    # exporter, including dependency types without local
                    # rustdoc items (e.g. Package::metadata -> common::Props).
                    self.names.add((path, reference_name))
            if kind == "trait":
                self.children(data["items"], path, ancestors, exported=exported, defaults=True)
            if kind == "enum":
                self.children(data["variants"], path, ancestors, exported=exported, defaults=True)
            self.children(
                list(field_ids(data)), path, ancestors, exported=exported,
                defaults=kind == "variant",
            )
            for impl_id in data.get("impls", []):
                self.walk(str(impl_id), path, ancestors, exported=exported)

    def children(self, ids, parent, ancestors, *, exported, defaults=False):
        for child_id in ids:
            child = self.index.get(str(child_id))
            if not isinstance(child, dict):
                raise ValueError(f"missing rustdoc child {child_id}")
            visibility = child.get("visibility")
            if visibility != "public" and not (defaults and visibility == "default"):
                continue
            name = child.get("name") or child.get("inner", {}).get("use", {}).get("name")
            if not name:
                raise ValueError(f"unnamed public rustdoc child {child_id}")
            self.walk(str(child_id), parent + (name,), ancestors, exported=exported)


def violations(document: Any, profile: dict[str, Any]) -> list[str]:
    try:
        surface = Surface(document, profile["package"].replace("-", "_"),
                          frozenset(profile.get("delegated_globs", [])))
    except (KeyError, TypeError, AttributeError) as error:
        raise ValueError(f"invalid rustdoc graph structure: {error}") from error
    failures = set()
    for ban in profile["banned_names"]:
        scope = tuple(ban["scope"].split("::"))
        if not any(path[:len(scope)] == scope for path, _ in surface.names):
            failures.add(f"missing API scope {ban['scope']} for banned name {ban['name']}")
            continue
        for path, name in surface.names:
            if (path[:len(scope)] == scope and name == ban["name"]
                    and (ban.get("recursive", True) or len(path) == len(scope) + 1)):
                failures.add(f"{'::'.join(path)} exposes banned name {name}; {ban['replacement']}")
    for family in profile["families"]:
        path = tuple(family.split("::"))
        if surface.exports.get(path) != "module":
            failures.add(f"missing public transaction module {family}")
            continue
        for name in sorted(FAMILY_NAMES):
            if surface.exports.get(path + (name,)) not in TYPE_KINDS:
                failures.add(f"{family} requires exported type {name}")
    return sorted(failures)


def rustdoc_command(profile: dict[str, Any], target: Path) -> list[str]:
    command = ["cargo", "rustdoc", "--package", profile["package"], "--lib", "--target-dir", str(target)]
    if not profile["default_features"]:
        command.append("--no-default-features")
    if profile["features"]:
        command += ["--features", ",".join(profile["features"])]
    return command + ["--", "-Zunstable-options", "--output-format", "json", "--document-hidden-items"]


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--policy", type=Path, default=POLICY)
    parser.add_argument("--profile")
    parser.add_argument("--json", type=Path)
    parser.add_argument(
        "--target-dir", type=Path,
        help="retain build cache here; otherwise use a temporary disk-backed directory",
    )
    args = parser.parse_args(argv)
    try:
        profiles = checked_policy(read_json(args.policy))
        if args.profile:
            profiles = [p for p in profiles if p["id"] == args.profile]
            if not profiles:
                raise ValueError(f"unknown profile {args.profile}")
        if args.json and (not args.profile or len(profiles) != 1):
            raise ValueError("--json requires --profile")
        scratch_context = (
            nullcontext(args.target_dir or ROOT) if args.target_dir or args.json else
            tempfile.TemporaryDirectory(prefix="litchi-api-glossary-", dir=ROOT.parent)
        )
        with scratch_context as scratch:
            target = (args.target_dir or Path(scratch)).resolve()
            failures = []
            for profile in profiles:
                if args.json:
                    source = args.json
                else:
                    # Each profile must be checked immediately after its build:
                    # another feature profile overwrites the same rustdoc file.
                    env = {**os.environ, "RUSTC_BOOTSTRAP": "1"}
                    subprocess.run(rustdoc_command(profile, target), cwd=ROOT, env=env, check=True)
                    source = target / "doc" / (profile["package"].replace("-", "_") + ".json")
                failures.extend(f"{profile['id']}: {v}" for v in violations(read_json(source), profile))
            if failures:
                print("\n".join(failures), file=sys.stderr)
                return 1
        bans = sum(len(p["banned_names"]) for p in profiles)
        families = sum(len(p["families"]) for p in profiles)
        print(
            f"API glossary: {len(profiles)} profiles, {bans} banned names, "
            f"{families} transaction families checked"
        )
        return 0
    except (ValueError, OSError, subprocess.CalledProcessError) as error:
        print(f"API glossary: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
