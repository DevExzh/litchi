#!/usr/bin/env python3
"""Propose, but never accept, owner replay result rules.

The final inventory is captured from feature-specific rustdoc JSON.  This
small helper turns the already checked public graph into policy-row candidates:
only a direct ``apply``/``apply_*`` method of a nominated immutable owner is
considered, and a candidate is emitted only after the existing identity-aware
``Surface``/``ReturnTypes`` checks prove both the ``Result`` payload and the
published owner.  Raw/record/Patch/History helpers are outside the nominated
owner paths and therefore never enter this proposal set.

The output is review input, not a policy writer.  Existing explicit rules are
reported separately and are never replaced.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any

try:
    from . import check_api_glossary as gate
except ImportError:  # Direct command-line execution.
    import check_api_glossary as gate


COMMIT_NAMES = frozenset({"Commit", "BatchCommit"})


def _is_apply_name(name: str) -> bool:
    return name == "apply" or name.startswith("apply_")


def _path(value: tuple[str, ...]) -> str:
    return "::".join(value)


def _method_generics(surface: gate.Surface, method: tuple[str, ...]) -> set[str]:
    item_id = surface.item_ids.get(method)
    function = surface.index.get(item_id, {}).get("inner", {}).get("function", {})
    parameters = function.get("generics", {}).get("params", [])
    return {
        parameter["name"]
        for parameter in parameters
        if isinstance(parameter, dict)
        and isinstance(parameter.get("kind"), dict)
        and "type" in parameter["kind"]
        and isinstance(parameter.get("name"), str)
    }


def _public_nominal_paths(surface: gate.Surface, item_id: str) -> list[tuple[str, ...]]:
    return sorted(
        path
        for path, exported_id in surface.item_ids.items()
        if exported_id == item_id
        and surface.exports.get(path) in gate.TYPE_KINDS
    )


def _type_spec(
    normalized: Any,
    *,
    surface: gate.Surface,
    method: tuple[str, ...],
    return_types: gate.ReturnTypes,
    visiting: frozenset[str] = frozenset(),
) -> tuple[Any | None, str | None]:
    """Render a normalized rustdoc type as a checked policy type spec."""
    if not isinstance(normalized, tuple) or not normalized:
        return None, "missing-identity"
    kind = normalized[0]
    if kind == "generic":
        name = normalized[1] if len(normalized) > 1 else None
        if not isinstance(name, str) or name not in _method_generics(surface, method):
            return None, "unsupported-generic-argument"
        return {"generic": name}, None
    if kind != "nominal" or len(normalized) < 3:
        return None, "unsupported-generic-argument"
    item_id = str(normalized[1])
    if item_id in visiting:
        return None, "missing-identity"
    paths = _public_nominal_paths(surface, item_id)
    if not paths:
        entry = surface.paths.get(item_id)
        if not isinstance(entry, dict) or entry.get("crate_id") == surface.crate_id:
            return None, "missing-identity"
        canonical = tuple(entry.get("path", ()))
        if not canonical or return_types.type_id(canonical) != item_id:
            return None, "ambiguous-type-identity"
        paths = [canonical]
    # A policy type spec cannot encode arbitrary aliases.  All selected paths
    # below are identity checked again by ReturnTypes.policy_type.
    nested_specs: list[Any] = []
    for argument in normalized[2]:
        spec, reason = _type_spec(
            argument,
            surface=surface,
            method=method,
            return_types=return_types,
            visiting=visiting | {item_id},
        )
        if reason:
            return None, reason
        nested_specs.append(spec)
    path = min(paths, key=lambda value: (len(value), value))
    spec: Any = {"path": _path(path), "type_args": nested_specs} if nested_specs else _path(path)
    return spec, None


def _commit_paths(
    payload: Any,
    *,
    surface: gate.Surface,
    return_types: gate.ReturnTypes,
) -> list[tuple[str, ...]]:
    if not isinstance(payload, tuple) or len(payload) < 2 or payload[0] != "nominal":
        return []
    item_id = str(payload[1])
    identity = return_types._resolve_nominal_identity(item_id)
    if identity is None:
        return []
    paths = [
        path
        for path, kind in surface.exports.items()
        if kind in gate.TYPE_KINDS
        and path[-1] in COMMIT_NAMES
        and return_types.nominal_identity(path) == identity
        and return_types.is_commit(path)
    ]
    canonical = ("litchi_core", "edit", "Commit")
    if (return_types.nominal_identity(canonical, canonical=True) == identity
            and return_types.is_commit(canonical)):
        paths.append(canonical)
    return sorted(set(paths))


def _generic_parameters(surface: gate.Surface, path: tuple[str, ...]) -> list[dict[str, Any]]:
    item_id = surface.item_ids.get(path)
    item = surface.index.get(item_id, {})
    inner = item.get("inner", {})
    definition = next(iter(inner.values()), {})
    parameters = definition.get("generics", {}).get("params", [])
    return [
        parameter
        for parameter in parameters
        if isinstance(parameter, dict)
        and isinstance(parameter.get("name"), str)
        and isinstance(parameter.get("kind"), dict)
        and "type" in parameter["kind"]
    ]


def _raw_type_arguments(value: Any) -> list[Any] | None:
    if not isinstance(value, dict) or len(value) != 1:
        return None
    resolved = value.get("resolved_path")
    if not isinstance(resolved, dict):
        return [] if value.get("generic") is not None else None
    arguments = resolved.get("args")
    if arguments is None:
        return []
    bracketed = arguments.get("angle_bracketed")
    if not isinstance(bracketed, dict) or bracketed.get("constraints"):
        return None
    result = []
    for argument in bracketed.get("args", []):
        if not isinstance(argument, dict):
            return None
        if "lifetime" in argument:
            continue
        if "type" not in argument:
            return None
        result.append(argument["type"])
    return result


def _bind_alias_type(
    raw: Any,
    actual: Any,
    *,
    parameters: set[str],
    bindings: dict[str, Any],
    return_types: gate.ReturnTypes,
) -> bool:
    """Bind alias parameters by the rustdoc alias target, never by position."""
    if not isinstance(raw, dict) or len(raw) != 1:
        return False
    generic = raw.get("generic")
    if generic is not None:
        if generic not in parameters:
            return False
        previous = bindings.get(generic)
        if previous is not None and previous != actual:
            return False
        bindings[generic] = actual
        return True
    resolved = raw.get("resolved_path")
    if not isinstance(resolved, dict) or not isinstance(actual, tuple) or actual[0] != "nominal":
        return False
    target_identity = return_types._resolve_nominal_identity(str(resolved.get("id")))
    actual_identity = return_types._resolve_nominal_identity(str(actual[1]))
    if target_identity is None or target_identity != actual_identity:
        return False
    raw_arguments = _raw_type_arguments(raw)
    if raw_arguments is None or len(raw_arguments) != len(actual[2]):
        return False
    return all(
        _bind_alias_type(
            argument,
            value,
            parameters=parameters,
            bindings=bindings,
            return_types=return_types,
        )
        for argument, value in zip(raw_arguments, actual[2])
    )


def _generic_specs(
    expected: tuple[str, ...],
    payload: Any,
    *,
    surface: gate.Surface,
    method: tuple[str, ...],
    return_types: gate.ReturnTypes,
) -> tuple[list[Any] | None, str | None]:
    """Render only arguments justified by the expected type's own generics."""
    parameters = _generic_parameters(surface, expected)
    if not parameters:
        # A canonical dependency path has no local definition. Its arguments
        # are already ordered by the dependency's resolved signature; this is
        # the only case where the normalized payload supplies positional args.
        if surface.item_ids.get(expected) is None and isinstance(payload, tuple):
            specs = []
            for argument in payload[2]:
                spec, reason = _type_spec(
                    argument,
                    surface=surface,
                    method=method,
                    return_types=return_types,
                )
                if reason:
                    return None, reason
                specs.append(spec)
            return specs, None
        return [], None

    item_id = surface.item_ids.get(expected)
    item = surface.index.get(item_id, {})
    inner = item.get("inner", {})
    definition = next(iter(inner.values()), {})
    bindings: dict[str, Any] = {}
    if "type_alias" in inner:
        alias_target = definition.get("type")
        if not _bind_alias_type(
            alias_target,
            payload,
            parameters={parameter["name"] for parameter in parameters},
            bindings=bindings,
            return_types=return_types,
        ):
            return None, "unsupported-generic-mapping"
    elif (not isinstance(payload, tuple) or payload[0] != "nominal"
          or return_types._resolve_nominal_identity(str(payload[1]))
          != return_types.nominal_identity(expected)):
        return None, "unsupported-generic-mapping"
    else:
        for parameter, argument in zip(parameters, payload[2]):
            bindings[parameter["name"]] = argument

    last_bound = max(
        (index for index, parameter in enumerate(parameters)
         if parameter["name"] in bindings),
        default=-1,
    )
    specs: list[Any] = []
    for parameter in parameters[:last_bound + 1]:
        name = parameter["name"]
        argument = bindings.get(name)
        if argument is None:
            default = parameter["kind"]["type"].get("default")
            argument = return_types.resolve(default) if default is not None else None
        if argument is None:
            return None, "unsupported-generic-mapping"
        spec, reason = _type_spec(
            argument,
            surface=surface,
            method=method,
            return_types=return_types,
        )
        if reason:
            return None, reason
        specs.append(spec)
    return specs, None


def _candidate_specs(
    expected: tuple[str, ...],
    payload: Any,
    *,
    surface: gate.Surface,
    method: tuple[str, ...],
    return_types: gate.ReturnTypes,
) -> list[tuple[list[Any], str | None]]:
    """Try Rust's defaults first, then an identity-derived generic mapping."""
    candidates: list[tuple[list[Any], str | None]] = [([], None)]
    specs, reason = _generic_specs(
        expected,
        payload,
        surface=surface,
        method=method,
        return_types=return_types,
    )
    if specs:
        candidates.append((specs, reason))
    elif reason:
        candidates.append(([], reason))
    return candidates


def _owner_method_paths(surface: gate.Surface, owner: str) -> list[tuple[str, ...]]:
    prefix = tuple(owner.split("::"))
    return sorted(
        path
        for path, kind in surface.exports.items()
        if kind == "function"
        and len(path) == len(prefix) + 1
        and path[:-1] == prefix
        and _is_apply_name(path[-1])
    )


def propose(document: dict[str, Any], profile: dict[str, Any]) -> dict[str, Any]:
    """Return deterministic candidate policy rows and unresolved findings."""
    crate = profile["package"].replace("-", "_")
    surface = gate.Surface(document, crate, frozenset(profile.get("delegated_globs", [])))
    return_types = gate.ReturnTypes(surface)
    rules = profile.get("result_methods", [])
    existing = {
        str(rule["method"]): rule
        for rule in rules
        if isinstance(rule, dict) and isinstance(rule.get("method"), str)
    }
    family_commit_paths = {
        tuple(family.split("::")) + ("Commit",)
        for family in profile.get("families", [])
    }
    candidates: list[dict[str, Any]] = []
    unresolved: list[dict[str, Any]] = []
    nominated = profile.get("immutable_owners", [])
    for owner in nominated:
        owner_path = tuple(owner.split("::"))
        if surface.exports.get(owner_path) not in gate.TYPE_KINDS:
            unresolved.append({"owner": owner, "reason": "missing-owner"})
            continue
        for method in _owner_method_paths(surface, owner):
            method_name = _path(method)
            if method_name in existing:
                continue
            item_id = surface.item_ids.get(method)
            item = surface.index.get(item_id, {})
            function = item.get("inner", {}).get("function", {})
            receiver = gate.self_receiver(function)
            if receiver != "&self":
                unresolved.append({
                    "method": method_name,
                    "reason": "unsupported-receiver",
                    "detail": receiver or "missing-identity",
                })
                continue
            output = return_types.method_output(method)
            payload = return_types.result_payload(output)
            if output is None:
                unresolved.append({"method": method_name, "reason": "missing-identity"})
                continue
            if payload is None:
                unresolved.append({"method": method_name, "reason": "unsupported-result"})
                continue
            commit_paths = _commit_paths(payload, surface=surface, return_types=return_types)
            if not commit_paths:
                unresolved.append({"method": method_name, "reason": "non-commit-result"})
                continue
            proven: list[tuple[tuple[str, ...], list[Any]]] = []
            unsupported_reason: str | None = None
            result_proved = False
            owner_proved = False
            for expected in commit_paths:
                for specs, reason in _candidate_specs(
                    expected,
                    payload,
                    surface=surface,
                    method=method,
                    return_types=return_types,
                ):
                    if reason:
                        unsupported_reason = reason
                        continue
                    checked_args = [return_types.policy_type(spec, method) for spec in specs]
                    if any(argument is None for argument in checked_args):
                        unsupported_reason = "missing-identity"
                        continue
                    result_ok = return_types.matches_result(method, expected, checked_args)
                    owner_ok = result_ok and return_types.matches_owner(
                        method, expected, checked_args
                    )
                    result_proved = result_proved or result_ok
                    owner_proved = owner_proved or owner_ok
                    if owner_ok:
                        proven.append((expected, specs))
                        break
            if not proven:
                if owner_proved:
                    # Defensive: this branch is unreachable while proving rows,
                    # but keeps a stable unresolved shape if a future checker
                    # changes its identity model.
                    reason = "ambiguous-commit-path"
                elif result_proved:
                    reason = "wrong-owner"
                else:
                    reason = unsupported_reason or "unverified-return"
                unresolved.append({"method": method_name, "reason": reason})
                continue
            proven.sort(key=lambda row: (
                0 if row[0] in family_commit_paths else 1,
                len(row[0]),
                row[0],
            ))
            expected, specs = proven[0]
            row: dict[str, Any] = {
                "method": method_name,
                "ok": _path(expected),
            }
            if specs:
                row["type_args"] = specs
            alternatives = [_path(path) for path, _ in proven[1:]]
            if alternatives:
                row["alternatives"] = alternatives
            candidates.append(row)
    return {
        "profile": profile["id"],
        "package": profile["package"],
        "candidates": candidates,
        "unresolved": unresolved,
        "existing": sorted(existing),
    }


def _load_profile(policy: Path, profile_id: str) -> dict[str, Any]:
    profiles = gate.checked_policy(gate.read_json(policy))
    for profile in profiles:
        if profile["id"] == profile_id:
            return profile
    raise ValueError(f"unknown profile {profile_id}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--json", required=True, type=Path, help="one exact rustdoc JSON capture")
    parser.add_argument("--profile", required=True, help="one profile id from --policy")
    parser.add_argument("--policy", type=Path, default=gate.POLICY)
    args = parser.parse_args(argv)
    try:
        profile = _load_profile(args.policy, args.profile)
        result = propose(gate.read_json(args.json), profile)
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as error:
        print(f"propose_api_result_rules: {error}", file=sys.stderr)
        return 2
    json.dump(result, sys.stdout, indent=2, sort_keys=True)
    sys.stdout.write("\n")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
