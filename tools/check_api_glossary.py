#!/usr/bin/env python3
"""Check registered ADR 0034 vocabulary waves against Rust's public API.

Rules live in api_glossary.json; each migration explicitly registers its bans
and completed transaction families. A ban matches a name exactly, or as a
prefix/suffix when it says so ("match"). Record 0351 adds shape rules:
"immutable_owners" lists owner types (a format's Package and projections)
whose public methods must not take `&mut self`, and "family_methods" makes
every registered family expose the commit-triad accessors and return its
Commit from Edit::commit. "result_methods" pins a method's Result payload to a
specific Commit (including nested types and method type parameters); owner
replay must publish that owner. Local aliases are resolved by identity, not
spelling. "shared_types" explicitly pins nominated public aliases/re-exports to
the nominal litchi_core::edit::Commit or Diagnostics item; generic arguments
and richer format-owned diagnostics remain outside that rule. Direct shared
core Commit types use their canonical dependency path.
Each profile builds its own feature-specific rustdoc JSON; --json/--profile
checks an existing document instead. Names are matched exactly under a public
path prefix, never by substring. A ban with recursive=false matches direct
children of its scope, preserving unrelated nested domain vocabulary. Family
checks verify exported types and these method shapes. Transaction semantics
remain Rust/CRUD test obligations.
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

try:
    from .rustdoc_return_types import ReturnTypes
except ImportError:  # Direct command-line execution.
    from rustdoc_return_types import ReturnTypes

ROOT = Path(__file__).resolve().parents[1]
POLICY = ROOT / "tools/api_glossary.json"
FAMILY_NAMES = frozenset({"Edit", "Patch", "Commit", "Diagnostics", "Error", "LimitKind"})
BAN_MATCHES = frozenset({"exact", "prefix", "suffix"})
# Record 0351 family_methods: what each family type must (or must not) expose.
# Rustdoc does not list a dependency's inherent methods, so a member that is a
# re-export or alias of another crate's type (litchi_core::edit::Commit, a
# shared ODF metadata family) is checked where it is defined: by that crate's
# own profile, or for litchi_core::edit by its unit tests.
FAMILY_METHODS = {
    "Commit": frozenset({"package", "into_package", "patch", "diagnostics", "into_parts"}),
    "Diagnostics": frozenset({"changed"}),
    "Edit": frozenset({"commit"}),
    "Patch": frozenset({"is_noop", "inverse"}),
}
FAMILY_FORBIDDEN = {"Edit": frozenset({"rollback", "is_changed", "snapshot"}),
                    "Commit": frozenset({"snapshot", "into_snapshot", "is_changed"})}
# An edit stages changes; only its committed package exposes artifact output.
# Keep this shared with the debt inventory so an unregistered family is still
# reported. These names identify output, not ordinary read projections or
# preparation methods such as `prepare`/`plan`.
EDIT_OUTPUT_NAMES = frozenset({
    "to_bytes", "into_bytes", "to_plain_bytes", "to_encrypted_bytes",
    "to_reencrypted_bytes", "write_to", "save", "save_plain", "save_encrypted",
    "save_reencrypted", "save_with_durability",
})
TYPE_KINDS = frozenset({"struct", "enum", "union", "type_alias"})
SHARED_TYPE_TARGETS = frozenset({
    "litchi_core::edit::Commit",
    "litchi_core::edit::Diagnostics",
})
IDENTIFIER = re.compile(r"[A-Za-z_][A-Za-z_0-9]*\Z")


def is_edit_output(name: str) -> bool:
    """Known artifact-output verbs that bypass Edit::commit (record 0351)."""
    return name in EDIT_OUTPUT_NAMES or name.endswith(("_with_rendered", "_to_stream"))


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
        optional = {
            "delegated_globs", "immutable_owners", "family_methods", "result_methods",
            "shared_types",
        }
        if (not isinstance(profile, dict) or not fields <= set(profile)
                or set(profile) - (fields | optional)):
            raise ValueError("profile requires exactly " + ", ".join(sorted(fields))
                             + ", with optional " + ", ".join(sorted(optional)))
        if "family_methods" in profile and type(profile["family_methods"]) is not bool:
            raise ValueError("family_methods must be boolean")
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

        def check_path(path: Any, *, local=True) -> None:
            if (not isinstance(path, str) or not all(IDENTIFIER.fullmatch(p) for p in path.split("::"))
                    or (not local and len(path.split("::")) < 2)):
                raise ValueError("scope/family must be a fully qualified Rust path")
            if local and path.split("::")[0] != crate:
                raise ValueError(f"scope/family must start with {crate}")

        def check_type(spec: Any) -> None:
            if isinstance(spec, str):
                check_path(spec, local=False)
            elif isinstance(spec, dict) and set(spec) == {"generic"}:
                name = spec["generic"]
                if not isinstance(name, str) or not IDENTIFIER.fullmatch(name) or name == "Self":
                    raise ValueError("type generic must name a method type parameter")
            elif isinstance(spec, dict) and {"path"} <= set(spec) <= {"path", "type_args"}:
                check_path(spec["path"], local=False)
                if not isinstance(spec.get("type_args", []), list):
                    raise ValueError("nested type_args must be an array")
                for arg in spec.get("type_args", []):
                    check_type(arg)
            else:
                raise ValueError("type argument must be a path, parameter or nested nominal type")

        bans = profile["banned_names"]
        if not isinstance(bans, list):
            raise ValueError("banned_names must be an array")
        keys = set()
        for ban in bans:
            required = {"name", "scope", "replacement"}
            if (not isinstance(ban, dict) or not required <= set(ban)
                    or set(ban) - (required | {"recursive", "match"})):
                raise ValueError("ban requires name, scope and replacement, with optional recursive and match")
            if "recursive" in ban and type(ban["recursive"]) is not bool:
                raise ValueError("ban recursive must be boolean")
            if "match" in ban and ban["match"] not in BAN_MATCHES:
                raise ValueError("ban match must be exact, prefix or suffix")
            if not isinstance(ban["name"], str) or not IDENTIFIER.fullmatch(ban["name"]):
                raise ValueError("banned name must be one Rust identifier")
            if not isinstance(ban["replacement"], str) or not ban["replacement"].strip():
                raise ValueError("ban requires a nonempty replacement explanation")
            check_path(ban["scope"])
            key = (ban["scope"], ban["name"], ban.get("match", "exact"))
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
        owners = profile.get("immutable_owners", [])
        if not isinstance(owners, list):
            raise ValueError("immutable_owners must be an array")
        for owner in owners:
            check_path(owner)
        if len(set(owners)) != len(owners):
            raise ValueError("duplicate immutable owner")
        shared_types = profile.get("shared_types", [])
        if not isinstance(shared_types, list):
            raise ValueError("shared_types must be an array")
        seen_shared_paths = set()
        for shared in shared_types:
            if (not isinstance(shared, dict) or set(shared) != {"path", "target"}):
                raise ValueError("shared type requires exactly path and target")
            check_path(shared["path"])
            if (not isinstance(shared["target"], str)
                    or shared["target"] not in SHARED_TYPE_TARGETS):
                raise ValueError("shared type target must be a canonical litchi_core::edit type")
            if shared["path"] in seen_shared_paths:
                raise ValueError("duplicate shared type path")
            seen_shared_paths.add(shared["path"])
        results = profile.get("result_methods", [])
        if not isinstance(results, list):
            raise ValueError("result_methods must be an array")
        seen_methods = set()
        for result in results:
            if (not isinstance(result, dict) or not {"method", "ok"} <= set(result)
                    or set(result) - {"method", "ok", "type_args"}):
                raise ValueError("result method requires method and ok, with optional type_args")
            check_path(result["method"])
            check_path(result["ok"], local=False)
            if result["ok"].split("::")[-1] not in {"Commit", "BatchCommit"}:
                raise ValueError("result method ok must name Commit or BatchCommit")
            if result["method"] in seen_methods:
                raise ValueError("duplicate result method")
            seen_methods.add(result["method"])
            if not isinstance(result.get("type_args", []), list):
                raise ValueError("result method type_args must be an array")
            for arg in result.get("type_args", []):
                check_type(arg)
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


def self_receiver(function: Any) -> str | None:
    """The receiver of a rustdoc function item, or None for an associated fn."""
    inputs = function.get("sig", {}).get("inputs") if isinstance(function, dict) else None
    if not isinstance(inputs, list) or not inputs:
        return None
    first = inputs[0]
    if not isinstance(first, list) or len(first) != 2 or first[0] != "self":
        return None
    reference = first[1].get("borrowed_ref") if isinstance(first[1], dict) else None
    if isinstance(reference, dict):
        return "&mut self" if reference.get("is_mutable") else "&self"
    return "self"


def ban_matches(ban: dict[str, Any], name: str) -> bool:
    mode = ban.get("match", "exact")
    if mode == "prefix":
        return name.startswith(ban["name"]) and name != ban["name"]
    if mode == "suffix":
        return name.endswith(ban["name"]) and name != ban["name"]
    return name == ban["name"]


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
        # Preserve aliases at their public spelling: traversing the nominal
        # target must not discard the alias's generic substitutions.
        self.item_ids: dict[tuple[str, ...], str] = {}
        self.names: set[tuple[tuple[str, ...], str]] = set()
        # Method receivers ("&mut self", "&self", "self") by caller-visible
        # path, and dependency paths behind re-exports and type aliases.
        self.receivers: dict[tuple[str, ...], str] = {}
        self.foreign: dict[tuple[str, ...], tuple[str, ...]] = {}
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
                self.item_ids.setdefault(path, item_id)
            self.names.add((path, path[-1]))
            self.foreign[path] = tuple(entry["path"])
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
            self.item_ids.setdefault(path, item_id)
        self.names.add((path, path[-1]))
        if kind == "function":
            receiver = self_receiver(data)
            if receiver:
                self.receivers[path] = receiver
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
                entry = self.paths.get(str(target))
                if isinstance(entry, dict) and entry["crate_id"] != self.crate_id:
                    self.foreign[path] = tuple(entry["path"])
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
    return_types = ReturnTypes(surface)
    for ban in profile["banned_names"]:
        scope = tuple(ban["scope"].split("::"))
        if not any(path[:len(scope)] == scope for path, _ in surface.names):
            failures.add(f"missing API scope {ban['scope']} for banned name {ban['name']}")
            continue
        for path, name in surface.names:
            if (path[:len(scope)] == scope and ban_matches(ban, name)
                    and (ban.get("recursive", True) or len(path) == len(scope) + 1)):
                failures.add(f"{'::'.join(path)} exposes banned name {name}; {ban['replacement']}")
    for owner in profile.get("immutable_owners", []):
        path = tuple(owner.split("::"))
        if surface.exports.get(path) not in TYPE_KINDS:
            failures.add(f"missing immutable owner type {owner}")
            continue
        for method, receiver in surface.receivers.items():
            if method[:-1] == path and receiver == "&mut self":
                failures.add(f"{'::'.join(method)} takes &mut self on an immutable owner; "
                             "stage it in an Edit that commits a new Package (record 0351)")
    methods: dict[tuple[str, ...], set[str]] = {}
    for method in surface.receivers:
        methods.setdefault(method[:-1], set()).add(method[-1])
    for path, kind in surface.exports.items():
        # Associated functions have no receiver but still count as methods.
        if kind == "function" and len(path) > 1:
            methods.setdefault(path[:-1], set()).add(path[-1])
    for rule in profile.get("shared_types", []):
        public = tuple(rule["path"].split("::"))
        target = tuple(rule["target"].split("::"))
        if not return_types.matches_shared(public, target):
            failures.add(
                f"{rule['path']} must resolve nominally to {rule['target']} (record 0351)"
            )
    for family in profile["families"]:
        path = tuple(family.split("::"))
        if surface.exports.get(path) != "module":
            failures.add(f"missing public transaction module {family}")
            continue
        for name in sorted(FAMILY_NAMES):
            if surface.exports.get(path + (name,)) not in TYPE_KINDS:
                failures.add(f"{family} requires exported type {name}")
        if not profile.get("family_methods", False):
            continue
        for name, required in sorted(FAMILY_METHODS.items()):
            member = path + (name,)
            if surface.exports.get(member) not in TYPE_KINDS:
                continue
            if member in surface.foreign:
                continue
            present = methods.get(member, set())
            for method in sorted(required - present):
                failures.add(f"{family}::{name} requires method {method} (record 0351)")
            forbidden = FAMILY_FORBIDDEN.get(name, frozenset()) & present
            if name == "Edit":
                forbidden = forbidden | {method for method in present if is_edit_output(method)}
            for method in sorted(forbidden):
                failures.add(f"{family}::{name} exposes forbidden method {method} (record 0351)")
        commit_method = path + ("Edit", "commit")
        if (commit_method in surface.item_ids and path + ("Edit",) not in surface.foreign
                and surface.exports.get(path + ("Commit",)) in TYPE_KINDS
                and (not return_types.is_commit(path + ("Commit",))
                     or not return_types.matches_result(commit_method, path + ("Commit",)))):
            failures.add(f"{family}::Edit::commit must return Result<{family}::Commit> (record 0351)")
    for rule in profile.get("result_methods", []):
        method = tuple(rule["method"].split("::"))
        expected = tuple(rule["ok"].split("::"))
        args = [return_types.policy_type(arg, method)
                for arg in rule.get("type_args", [])]
        if (surface.exports.get(method) != "function" or not return_types.is_commit(expected)
                or any(arg is None for arg in args)
                or not return_types.matches_result(method, expected, args)
                or ((method[-1] == "apply" or method[-1].startswith("apply_"))
                    and not return_types.matches_owner(method, expected, args))):
            failures.add(f"{rule['method']} must return Result<{rule['ok']}> (record 0351)")
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
        owners = sum(len(p.get("immutable_owners", [])) for p in profiles)
        strict = sum(1 for p in profiles if p.get("family_methods"))
        results = sum(len(p.get("result_methods", [])) for p in profiles)
        print(
            f"API glossary: {len(profiles)} profiles, {bans} banned names, "
            f"{families} transaction families checked, {owners} immutable owners, "
            f"family methods in {strict} profiles, {results} explicit result contracts"
        )
        return 0
    except (ValueError, OSError, subprocess.CalledProcessError) as error:
        print(f"API glossary: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
