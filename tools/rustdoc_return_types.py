"""Resolve rustdoc return types by identity, including local generic aliases.

This checks API shape, not lifetime equivalence or trait solving. Lifetimes are
therefore erased; unsupported/opaque types fail closed. External nominal types
retain their rustdoc identity. An external alias without its definition cannot
be unfolded merely because it happens to be named Result or Commit.
"""
from __future__ import annotations

from typing import Any


NOMINAL_KINDS = frozenset({"struct", "enum", "union", "type_alias"})

# The unified facade has a deliberately different owner accessor from the
# format packages.  Keep this set explicit: a format-owned Commit that happens
# to grow ``file``/``into_file`` must not silently satisfy the facade contract.
FACADE_COMMIT_PATHS = frozenset({
    ("litchi_ops", "BatchCommit"),
    ("litchi_ops", "batch", "BatchCommit"),
    ("litchi_ops", "transaction", "Commit"),
})

# ODF metadata Commit is a generic type defined by the common crate and
# re-exported through each format crate as a local alias.  Its defining crate
# is audited separately, so the return resolver must prove this exact nominal
# identity before extracting the alias's owner argument.
ODF_METADATA_COMMIT = ("litchi_odf_common", "metadata", "Commit")

# rustdoc JSON for a dependency may omit a public re-export while retaining the
# defining item's path. These are explicit, audited bindings from a
# caller-visible path to its defining path; they are not basename heuristics and
# must be extended only with reviewed re-export provenance.
CANONICAL_REEXPORTS = {
    ("litchi_opc", "SourceLimits"): ("litchi_opc", "limits", "SourceLimits"),
    ("litchi_ooxml_common", "mce", "Limits"):
        ("litchi_ooxml_common", "mce", "model", "Limits"),
}


class ReturnTypes:
    def __init__(self, surface):
        self.surface = surface

    def type_id(self, path: tuple[str, ...]):
        item_id = self.surface.item_ids.get(path)
        if item_id is not None:
            return item_id
        # Direct dependency types need not have a re-export in the caller.
        # Only canonical external nominal paths qualify, with unique identity.
        candidates = [key for key, value in self.surface.paths.items()
                      if tuple(value["path"]) == path
                      and value["crate_id"] != self.surface.crate_id
                      and value["kind"] in {"struct", "enum", "union", "type_alias"}]
        return candidates[0] if len(candidates) == 1 else None

    def type_at(self, path: tuple[str, ...], args=()):
        item_id = self.type_id(path)
        if item_id is None:
            return None
        return self.resolve({"resolved_path": {
            "id": item_id,
            "args": {"angle_bracketed": {
                "args": [{"type": arg} for arg in args], "constraints": [],
            }},
        }})

    def raw_type_at(self, path: tuple[str, ...]):
        item_id = self.type_id(path)
        return {"resolved_path": {"id": item_id, "args": None}} if item_id is not None else None

    def _resolve_nominal_identity(self, item_id):
        """Resolve local type aliases to one rustdoc nominal identity.

        The identity includes the dependency crate id and rustdoc item id. A
        missing alias target, non-nominal target, or alias cycle is unknown and
        therefore fails closed. Names and the final component of a path are
        deliberately never consulted.
        """
        visited = set()
        while True:
            item_id = str(item_id)
            if item_id in visited:
                return None
            visited.add(item_id)
            item = self.surface.index.get(item_id)
            entry = self.surface.paths.get(item_id)
            if item is None:
                if not isinstance(entry, dict) or entry.get("kind") not in NOMINAL_KINDS:
                    return None
                crate_id = entry.get("crate_id")
                return (crate_id, item_id) if type(crate_id) is int else None
            inner = item.get("inner", {})
            if not isinstance(inner, dict) or len(inner) != 1:
                return None
            kind, data = next(iter(inner.items()))
            if kind == "type_alias":
                if not isinstance(data, dict):
                    return None
                resolved = data.get("type", {}).get("resolved_path", {})
                if not isinstance(resolved, dict) or resolved.get("id") is None:
                    return None
                item_id = resolved["id"]
                continue
            if kind not in NOMINAL_KINDS or not isinstance(entry, dict):
                return None
            crate_id = entry.get("crate_id")
            return (crate_id, item_id) if type(crate_id) is int else None

    def nominal_identity(self, path: tuple[str, ...], *, canonical=False):
        """Return a public type's nominal identity, or None if ambiguous."""
        if canonical:
            candidates = [
                item_id for item_id, entry in self.surface.paths.items()
                if tuple(entry.get("path", ())) == path
                and entry.get("crate_id") != self.surface.crate_id
                and entry.get("kind") in NOMINAL_KINDS
            ]
            if not candidates:
                defining_path = CANONICAL_REEXPORTS.get(path)
                if defining_path is None:
                    return None
                candidates = [
                    item_id for item_id, entry in self.surface.paths.items()
                    if tuple(entry.get("path", ())) == defining_path
                    if entry.get("crate_id") != self.surface.crate_id
                    and entry.get("kind") in NOMINAL_KINDS
                ]
            # Multiple rustdoc IDs for one canonical path are ambiguous, even
            # when their spelling happens to match. The same fail-closed rule
            # applies to every explicit re-export binding above.
            if len(candidates) != 1:
                return None
            item_id = candidates[0]
        else:
            item_id = self.surface.item_ids.get(path)
            if item_id is None or self.surface.exports.get(path) not in NOMINAL_KINDS:
                return None
        return self._resolve_nominal_identity(item_id)

    def matches_shared(self, public_path: tuple[str, ...], canonical_path: tuple[str, ...]):
        """Match a nominated alias/re-export by nominal ID, never by name."""
        actual = self.nominal_identity(public_path)
        expected = self.nominal_identity(canonical_path, canonical=True)
        return actual is not None and actual == expected

    def policy_type(self, spec, method):
        """Build a checked type from paths, nested types and method binders."""
        if isinstance(spec, str):
            return self.raw_type_at(tuple(spec.split("::")))
        if not isinstance(spec, dict):
            return None
        if set(spec) == {"generic"}:
            function = self.surface.index.get(self.surface.item_ids.get(method), {}).get("inner", {}).get("function", {})
            parameters = {p["name"] for p in function.get("generics", {}).get("params", [])
                          if "type" in p["kind"]}
            name = spec["generic"]
            return {"generic": name} if name != "Self" and name in parameters else None
        if not {"path"} <= set(spec) <= {"path", "type_args"}:
            return None
        value = self.raw_type_at(tuple(spec["path"].split("::")))
        args = [self.policy_type(arg, method) for arg in spec.get("type_args", [])]
        if value is None or any(arg is None for arg in args):
            return None
        value["resolved_path"]["args"] = {"angle_bracketed": {
            "args": [{"type": arg} for arg in args], "constraints": [],
        }}
        return value

    def is_commit(self, path: tuple[str, ...]):
        if path == ("litchi_core", "edit", "Commit") and self.type_id(path) is not None:
            # This shared contract is enforced in the core crate. Arbitrary
            # foreign types merely named Commit are not trusted here.
            return True
        foreign = self.surface.foreign.get(path)
        if foreign:
            # Dependency members are checked in their defining crate, just as
            # family_methods does. A renamed unrelated external type is not
            # thereby a commit (for example `type Commit = Duration`).
            return foreign[-1] in {"Commit", "BatchCommit"}
        members = {p[-1] for p, kind in self.surface.exports.items()
                   if kind == "function" and p[:-1] == path}
        owner_getters = ({"file", "into_file"}
                         if path in FACADE_COMMIT_PATHS
                         else {"package", "into_package"})
        required = owner_getters | {"patch", "diagnostics", "into_parts"}
        if path in FACADE_COMMIT_PATHS:
            required |= {"value", "into_value"}
        return required <= members

    def commit_package(self, path: tuple[str, ...], args=()):
        """Resolve the published owner without trusting a Commit's spelling."""
        target = self.type_at(path, args)
        if target is None or target[0] != "nominal":
            return None
        if self.surface.paths.get(target[1], {}).get("path") == ["litchi_core", "edit", "Commit"]:
            return target[2][0] if len(target[2]) in (2, 3) else None
        target_path = tuple(self.surface.paths.get(target[1], {}).get("path", ()))
        if target_path == ODF_METADATA_COMMIT:
            # The path spelling is only the first filter.  Canonical identity
            # rejects an unrelated crate using the same name and rejects an
            # ambiguous rustdoc graph with duplicate external identities.
            expected = self.nominal_identity(ODF_METADATA_COMMIT, canonical=True)
            actual = self._resolve_nominal_identity(target[1])
            if (expected is not None and actual == expected
                    and len(target[2]) == 1):
                return target[2][0]
        item = self.surface.index.get(target[1], {})
        definition = next(iter(item.get("inner", {}).values()), {})
        params = [p for p in definition.get("generics", {}).get("params", [])
                  if "type" in p["kind"]]
        bindings = {p["name"]: value for p, value in zip(params, target[2])}
        getters = ("into_file",) if path in FACADE_COMMIT_PATHS else ("into_package",)
        for getter in getters:
            item_id = self.surface.item_ids.get(path + (getter,))
            function = self.surface.index.get(item_id, {}).get("inner", {}).get("function")
            if function is not None:
                return self.resolve(function.get("sig", {}).get("output"), bindings, target)
        return None

    def matches_owner(self, method, expected, args=()):
        """A registered owner replay publishes its source package type.

        Durable wire patches can differ from the returned family's Patch.
        Their input semantics remain Rust test obligations; this return-shape
        check verifies the exact family Commit and its published owner.  A
        ``Patch::apply`` method is defined on the patch but publishes the
        package supplied as its second argument, so it must not be compared
        with the patch receiver.
        """
        owner = self.type_at(method[:-1])
        if len(method) >= 2 and method[-1] == "apply" and method[-2] == "Patch":
            item_id = self.surface.item_ids.get(method)
            function = self.surface.index.get(item_id, {}).get("inner", {}).get("function")
            inputs = function.get("sig", {}).get("inputs", []) if isinstance(function, dict) else []
            if (not isinstance(inputs, list) or len(inputs) < 2
                    or any(not isinstance(item, list) or len(item) != 2 for item in inputs)):
                return False
            owner = self.resolve(inputs[1][1])
            while isinstance(owner, tuple) and owner and owner[0] == "ref":
                owner = owner[2]
        return owner is not None and self.commit_package(expected, args) == owner

    def resolve(self, value: Any, bindings=None, self_type=None, visiting=frozenset()):
        if not isinstance(value, dict) or len(value) != 1:
            return None
        bindings = bindings or {}
        kind, data = next(iter(value.items()))
        if kind == "generic":
            if data == "Self":
                return self_type
            return bindings.get(data, ("generic", data))
        if kind == "primitive":
            return ("primitive", data)
        if kind == "tuple":
            children = tuple(self.resolve(v, bindings, self_type, visiting) for v in data)
            return ("tuple", children) if all(v is not None for v in children) else None
        if kind == "borrowed_ref":
            child = self.resolve(data.get("type"), bindings, self_type, visiting)
            return ("ref", data.get("is_mutable", False), child) if child is not None else None
        if kind != "resolved_path" or not isinstance(data, dict):
            return None
        item_id = str(data.get("id"))
        if item_id in visiting:
            return None
        item = self.surface.index.get(item_id, {})
        inner = item.get("inner", {})
        if not inner and item_id not in self.surface.paths:
            return None
        definition = next(iter(inner.values()), {})
        raw_args = data.get("args")
        args = []
        if raw_args is not None:
            bracketed = raw_args.get("angle_bracketed")
            if not isinstance(bracketed, dict) or bracketed.get("constraints"):
                return None
            for arg in bracketed.get("args", []):
                if "lifetime" in arg:
                    continue
                if "type" not in arg:
                    return None
                resolved = self.resolve(arg["type"], bindings, self_type, visiting)
                if resolved is None:
                    return None
                args.append(resolved)
        params = definition.get("generics", {}).get("params", [])
        local_bindings = dict(bindings)
        type_params = []
        for param in params:
            if "lifetime" in param["kind"]:
                continue
            if "type" not in param["kind"]:
                return None
            type_params.append(param)
        if type_params:
            if len(args) > len(type_params):
                return None
            for i, param in enumerate(type_params):
                name = param["name"]
                if i == len(args):
                    default = param["kind"]["type"].get("default")
                    arg = (self.resolve(default, local_bindings, self_type, visiting | {item_id})
                           if default is not None else ("generic", name))
                    if arg is None:
                        return None
                    args.append(arg)
                local_bindings[name] = args[i]
        elif "generics" in definition and args:
            return None
        if "type_alias" in inner:
            return self.resolve(definition.get("type"), local_bindings, self_type,
                                visiting | {item_id})
        return ("nominal", item_id, tuple(args))

    def result_payload(self, normalized):
        if normalized is None or normalized[0] != "nominal":
            return None
        path = self.surface.paths.get(normalized[1], {}).get("path")
        if path not in (["core", "result", "Result"], ["std", "result", "Result"]):
            return None
        args = normalized[2]
        return args[0] if len(args) == 2 else None

    def method_output(self, path: tuple[str, ...]):
        item_id = self.surface.item_ids.get(path)
        function = self.surface.index.get(item_id, {}).get("inner", {}).get("function")
        if function is None:
            return None
        owner = self.type_at(path[:-1])
        return self.resolve(function.get("sig", {}).get("output"), self_type=owner)

    def matches_result(self, method: tuple[str, ...], expected: tuple[str, ...], args=()):
        target = self.type_at(expected, args)
        actual = self.result_payload(self.method_output(method))
        return target is not None and actual is not None and actual == target
