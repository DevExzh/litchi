#!/usr/bin/env python3
"""Fail-closed validation for a schema-2 corpus catalog/report binding.

The performance harness keeps the report at schema 1 and emits an additive
schema-2 catalog.  This checker is intentionally standard-library-only so CI
can independently verify the sidecar's canonical hashes and that every timed
report result is represented by the catalog it references.

This is deliberately a bounded validator rather than a general JSON Schema
implementation.  It checks the complete shape of the fields which participate
in the run identity and the small set of metadata objects needed to make those
checks type-safe.  The Rust manifest validator remains the owner of the full
schema's serialization and migration rules.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path
from typing import Any


SHA256 = re.compile(r"^[0-9a-f]{64}$")
CATALOG_KEYS = {
    "manifest_version",
    "manifest_kind",
    "catalog_id",
    "canonicalization",
    "catalog_sha256",
    "content_set_sha256",
    "build",
    "corpora",
    "case_bindings",
}
REFERENCE_KEYS = {
    "manifest_version",
    "catalog_id",
    "catalog_sha256",
    "content_set_sha256",
}
LEGACY_REQUIRED_KEYS = {
    "name",
    "generator",
    "package_format",
    "shape",
    "payload_kind",
    "compression",
    "entry_count",
    "archive_member_count",
    "entry_bytes",
    "uncompressed_payload_bytes",
    "archive_bytes",
    "archive_sha256",
    "target_entry",
    "target_payload_bytes",
    "target_payload_sha256",
    "xlsx",
}
LEGACY_OPTIONAL_KEYS = {"rtf_variant"}
XLSX_KEYS = {
    "sheet_count",
    "rows_per_sheet",
    "columns_per_sheet",
    "one_percent_update_count",
    "source_members",
}
XLSX_SOURCE_MEMBER_KEYS = {"workbook", "worksheets", "shared_strings", "styles"}
BUILD_KEYS = {"tool", "tool_version", "git_revision", "git_worktree_dirty", "source_files"}
SOURCE_FILE_KEYS = {"path", "sha256"}
CANONICALIZATION_KEYS = {"algorithm", "hash"}
CORPUS_KEYS = {
    "id",
    "name",
    "legacy_v1",
    "format",
    "size_class",
    "categories",
    "generator",
    "provenance",
    "bytes",
    "shape_parameters",
    "relationships",
    "security",
    "input",
    "limits",
    "members",
    "targets",
    "coverage",
}
GENERATOR_KEYS = {
    "id",
    "kind",
    "revision",
    "algorithm_id",
    "seed_spec",
    "parameters",
}
PROVENANCE_KEYS = {
    "source_kind",
    "source_path",
    "producer",
    "producer_version",
    "source_sha256",
    "license_spdx",
    "license_evidence",
    "redistributable",
}
BYTES_KEYS = {
    "archive_bytes",
    "archive_sha256",
    "logical_payload_bytes",
    "text_bytes",
    "media_bytes",
    "metadata_bytes",
}
RELATIONSHIP_KEYS = {
    "status",
    "relationship_count",
    "dependency_closure_nodes",
    "dependency_closure_edges",
    "max_depth",
    "max_out_degree",
}
SECURITY_KEYS = {"encryption", "signature", "protection", "macros", "external_links"}
SECURITY_FEATURE_KEYS = {"state", "kind", "evidence"}
PROTECTION_KEYS = {"state", "scopes", "evidence"}
EXTERNAL_LINK_KEYS = {"state", "count", "targets_sha256", "evidence"}
INPUT_KEYS = {"validity", "malformation_kind", "expected_behavior", "within_limits", "evidence"}
LIMITS_KEYS = {"profile_id", "profile_sha256", "observed"}
OBSERVED_LIMITS_KEYS = {"input_bytes", "members", "relationships", "materialized_bytes"}
MEMBERS_KEYS = {"status", "items"}
MEMBER_KEYS = {
    "ordinal",
    "name",
    "kind",
    "logical_bytes",
    "stored_bytes",
    "sha256",
    "stored_sha256",
    "role",
}
TARGET_KEYS = {"entry", "logical_bytes", "sha256"}
COVERAGE_KEYS = {"timed_cases", "guard_cases", "inventory_only"}
BINDING_KEYS = {"case", "corpus_id", "legacy_name", "legacy_archive_sha256", "role"}


class ValidationError(ValueError):
    """A malformed or internally inconsistent report/catalog pair."""


def _no_duplicate_object_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValidationError(f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _read_object(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_object_pairs,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValidationError) as error:
        raise ValidationError(f"cannot read {path}: {error}") from error
    if not isinstance(value, dict):
        raise ValidationError(f"{path} must contain a JSON object")
    return value


def _required(mapping: dict[str, Any], key: str, context: str) -> Any:
    if not isinstance(mapping, dict):
        raise ValidationError(f"{context} must be an object")
    if key not in mapping:
        raise ValidationError(f"{context} is missing {key!r}")
    return mapping[key]


def _object(value: Any, context: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValidationError(f"{context} must be an object")
    return value


def _list(value: Any, context: str) -> list[Any]:
    if not isinstance(value, list):
        raise ValidationError(f"{context} must be a list")
    return value


def _exact_keys(
    value: Any,
    required: set[str],
    context: str,
    optional: set[str] | None = None,
) -> dict[str, Any]:
    mapping = _object(value, context)
    optional = optional or set()
    allowed = required | optional
    keys = set(mapping)
    if any(not isinstance(key, str) for key in keys):
        raise ValidationError(f"{context} keys must be strings")
    missing = sorted(required - keys)
    extra = sorted(keys - allowed)
    if missing:
        raise ValidationError(f"{context} is missing {missing[0]!r}")
    if extra:
        raise ValidationError(f"{context} has unexpected field {extra[0]!r}")
    return mapping


def _string(value: Any, context: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str):
        raise ValidationError(f"{context} must be a string")
    if nonempty and not value:
        raise ValidationError(f"{context} must not be empty")
    return value


def _optional_string(value: Any, context: str, *, nonempty: bool = False) -> None | str:
    if value is None:
        return None
    return _string(value, context, nonempty=nonempty)


def _boolean(value: Any, context: str) -> bool:
    if not isinstance(value, bool):
        raise ValidationError(f"{context} must be a boolean")
    return value


def _integer(value: Any, context: str, *, minimum: int | None = None) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValidationError(f"{context} must be an integer")
    if minimum is not None and value < minimum:
        raise ValidationError(f"{context} must be at least {minimum}")
    return value


def _optional_integer(value: Any, context: str, *, minimum: int = 0) -> None | int:
    if value is None:
        return None
    return _integer(value, context, minimum=minimum)


def _hash(value: Any, context: str) -> str:
    if not isinstance(value, str) or SHA256.fullmatch(value) is None:
        raise ValidationError(f"{context} is not a lowercase SHA-256")
    return value


def _optional_hash(value: Any, context: str) -> None | str:
    if value is None:
        return None
    return _hash(value, context)


def _enum(value: Any, context: str, values: set[str]) -> str:
    value = _string(value, context)
    if value not in values:
        raise ValidationError(f"{context} has unsupported value {value!r}")
    return value


def _string_list(value: Any, context: str, *, nonempty: bool = True) -> list[str]:
    values = _list(value, context)
    return [
        _string(item, f"{context}[{index}]", nonempty=nonempty)
        for index, item in enumerate(values)
    ]


def canonical_bytes(value: Any) -> bytes:
    """Return the catalog's sorted-key, compact UTF-8 representation."""

    try:
        return json.dumps(
            value,
            ensure_ascii=False,
            allow_nan=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise ValidationError(f"value is not canonically serializable: {error}") from error


def _digest(value: Any) -> str:
    return hashlib.sha256(canonical_bytes(value)).hexdigest()


def _format_slug(value: Any) -> str:
    value = _string(value, "package_format")
    output: list[str] = []
    for character in value:
        if character.isascii() and character.isalnum():
            output.append(character.lower())
        elif not output or output[-1] != "-":
            output.append("-")
    slug = "".join(output).strip("-")
    if not slug:
        raise ValidationError("package_format must contain an ASCII letter or digit")
    return slug


def _corpus_id(corpus: dict[str, Any], context: str) -> str:
    package_format = _required(corpus, "package_format", context)
    archive_sha256 = _hash(
        _required(corpus, "archive_sha256", context),
        f"{context} archive_sha256",
    )
    return f"{_format_slug(package_format)}:sha256:{archive_sha256}"


def _validate_legacy_xlsx(value: Any, context: str) -> None:
    if value is None:
        return
    xlsx = _exact_keys(value, XLSX_KEYS, context)
    for key in (
        "sheet_count",
        "rows_per_sheet",
        "columns_per_sheet",
        "one_percent_update_count",
    ):
        _integer(xlsx[key], f"{context} {key}", minimum=0)
    source = _exact_keys(xlsx["source_members"], XLSX_SOURCE_MEMBER_KEYS, f"{context} source_members")
    _string(source["workbook"], f"{context} source_members workbook")
    _string_list(source["worksheets"], f"{context} source_members worksheets")
    _optional_string(source["shared_strings"], f"{context} source_members shared_strings")
    _optional_string(source["styles"], f"{context} source_members styles")


def _validate_legacy(value: Any, context: str) -> dict[str, Any]:
    legacy = _exact_keys(value, LEGACY_REQUIRED_KEYS, context, LEGACY_OPTIONAL_KEYS)
    for key in (
        "name",
        "generator",
        "package_format",
        "shape",
        "payload_kind",
        "compression",
        "target_entry",
    ):
        _string(legacy[key], f"{context} {key}")
    for key in (
        "entry_count",
        "archive_member_count",
        "entry_bytes",
        "uncompressed_payload_bytes",
        "archive_bytes",
        "target_payload_bytes",
    ):
        _integer(legacy[key], f"{context} {key}", minimum=0)
    _hash(legacy["archive_sha256"], f"{context} archive_sha256")
    _hash(legacy["target_payload_sha256"], f"{context} target_payload_sha256")
    if "rtf_variant" in legacy:
        _string(legacy["rtf_variant"], f"{context} rtf_variant", nonempty=False)
    _validate_legacy_xlsx(legacy["xlsx"], f"{context} xlsx")
    return legacy


def _validate_build(value: Any, context: str) -> None:
    build = _exact_keys(value, BUILD_KEYS, context)
    _string(build["tool"], f"{context} tool")
    _string(build["tool_version"], f"{context} tool_version")
    _optional_string(build["git_revision"], f"{context} git_revision")
    if build["git_worktree_dirty"] is not None:
        _boolean(build["git_worktree_dirty"], f"{context} git_worktree_dirty")
    source_files = _list(build["source_files"], f"{context} source_files")
    for index, value in enumerate(source_files):
        source = _exact_keys(value, SOURCE_FILE_KEYS, f"{context} source_files[{index}]")
        _string(source["path"], f"{context} source_files[{index}] path")
        _hash(source["sha256"], f"{context} source_files[{index}] sha256")


def _validate_generator(value: Any, context: str) -> None:
    generator = _exact_keys(value, GENERATOR_KEYS, context)
    _string(generator["id"], f"{context} id")
    _enum(generator["kind"], f"{context} kind", {"synthetic", "fixture", "transformed", "unknown"})
    _optional_string(generator["revision"], f"{context} revision")
    _optional_string(generator["algorithm_id"], f"{context} algorithm_id")
    _optional_string(generator["seed_spec"], f"{context} seed_spec")
    _object(generator["parameters"], f"{context} parameters")


def _validate_provenance(value: Any, context: str) -> None:
    provenance = _exact_keys(value, PROVENANCE_KEYS, context)
    _enum(
        provenance["source_kind"],
        f"{context} source_kind",
        {"generated", "repository-fixture", "external-fixture", "unknown"},
    )
    _optional_string(provenance["source_path"], f"{context} source_path")
    _optional_string(provenance["producer"], f"{context} producer")
    _optional_string(provenance["producer_version"], f"{context} producer_version")
    _optional_hash(provenance["source_sha256"], f"{context} source_sha256")
    _optional_string(provenance["license_spdx"], f"{context} license_spdx")
    _optional_string(provenance["license_evidence"], f"{context} license_evidence")
    if provenance["redistributable"] is not None:
        _boolean(provenance["redistributable"], f"{context} redistributable")


def _validate_bytes(value: Any, context: str) -> dict[str, Any]:
    summary = _exact_keys(value, BYTES_KEYS, context)
    _integer(summary["archive_bytes"], f"{context} archive_bytes", minimum=0)
    _hash(summary["archive_sha256"], f"{context} archive_sha256")
    _integer(summary["logical_payload_bytes"], f"{context} logical_payload_bytes", minimum=0)
    for key in ("text_bytes", "media_bytes", "metadata_bytes"):
        _optional_integer(summary[key], f"{context} {key}")
    return summary


def _validate_relationships(value: Any, context: str) -> None:
    relationships = _exact_keys(value, RELATIONSHIP_KEYS, context)
    _enum(
        relationships["status"],
        f"{context} status",
        {"complete", "partial", "not-applicable", "unknown"},
    )
    for key in (
        "relationship_count",
        "dependency_closure_nodes",
        "dependency_closure_edges",
        "max_depth",
        "max_out_degree",
    ):
        _optional_integer(relationships[key], f"{context} {key}")


def _validate_security_feature(value: Any, context: str) -> None:
    feature = _exact_keys(value, SECURITY_FEATURE_KEYS, context)
    _enum(feature["state"], f"{context} state", {"absent", "present", "synthetic", "unknown"})
    _optional_string(feature["kind"], f"{context} kind")
    _string(feature["evidence"], f"{context} evidence")


def _validate_security(value: Any, context: str) -> None:
    security = _exact_keys(value, SECURITY_KEYS, context)
    for key in ("encryption", "signature", "macros"):
        _validate_security_feature(security[key], f"{context} {key}")
    protection = _exact_keys(security["protection"], PROTECTION_KEYS, f"{context} protection")
    _enum(protection["state"], f"{context} protection state", {"absent", "present", "synthetic", "unknown"})
    _string_list(protection["scopes"], f"{context} protection scopes", nonempty=False)
    _string(protection["evidence"], f"{context} protection evidence")
    external = _exact_keys(security["external_links"], EXTERNAL_LINK_KEYS, f"{context} external_links")
    _enum(external["state"], f"{context} external_links state", {"absent", "present", "synthetic", "unknown"})
    _optional_integer(external["count"], f"{context} external_links count")
    _optional_hash(external["targets_sha256"], f"{context} external_links targets_sha256")
    _string(external["evidence"], f"{context} external_links evidence")


def _validate_input(value: Any, context: str) -> None:
    input_summary = _exact_keys(value, INPUT_KEYS, context)
    _enum(input_summary["validity"], f"{context} validity", {"well-formed", "malformed", "adversarial", "unknown"})
    _optional_string(input_summary["malformation_kind"], f"{context} malformation_kind")
    _enum(input_summary["expected_behavior"], f"{context} expected_behavior", {"accept", "refuse", "unknown"})
    if input_summary["within_limits"] is not None:
        _boolean(input_summary["within_limits"], f"{context} within_limits")
    _string(input_summary["evidence"], f"{context} evidence")


def _validate_limits(value: Any, context: str) -> None:
    limits = _exact_keys(value, LIMITS_KEYS, context)
    _optional_string(limits["profile_id"], f"{context} profile_id")
    _optional_hash(limits["profile_sha256"], f"{context} profile_sha256")
    observed = _exact_keys(limits["observed"], OBSERVED_LIMITS_KEYS, f"{context} observed")
    for key in OBSERVED_LIMITS_KEYS:
        _optional_integer(observed[key], f"{context} observed {key}")


def _validate_members(value: Any, context: str) -> list[dict[str, Any]]:
    members = _exact_keys(value, MEMBERS_KEYS, context)
    _enum(members["status"], f"{context} status", {"complete", "partial", "unavailable"})
    items = _list(members["items"], f"{context} items")
    ordinals: set[int] = set()
    names: set[str] = set()
    previous_ordinal = -1
    for index, value in enumerate(items):
        member_context = f"{context} items[{index}]"
        member = _exact_keys(value, MEMBER_KEYS, member_context)
        ordinal = _integer(member["ordinal"], f"{member_context} ordinal", minimum=0)
        if ordinal in ordinals:
            raise ValidationError(f"{member_context} ordinal is duplicated")
        if ordinal <= previous_ordinal:
            raise ValidationError(f"{context} ordinals are not strictly increasing")
        ordinals.add(ordinal)
        previous_ordinal = ordinal
        name = _string(member["name"], f"{member_context} name")
        if name in names:
            raise ValidationError(f"{member_context} name is duplicated")
        names.add(name)
        _string(member["kind"], f"{member_context} kind")
        _integer(member["logical_bytes"], f"{member_context} logical_bytes", minimum=0)
        _optional_integer(member["stored_bytes"], f"{member_context} stored_bytes")
        _hash(member["sha256"], f"{member_context} sha256")
        _optional_hash(member["stored_sha256"], f"{member_context} stored_sha256")
        _optional_string(member["role"], f"{member_context} role")
    return items


def _validate_targets(value: Any, legacy: dict[str, Any], context: str) -> None:
    targets = _list(value, context)
    if len(targets) != 1:
        raise ValidationError(f"{context} must contain exactly one target")
    target = _exact_keys(targets[0], TARGET_KEYS, f"{context}[0]")
    _string(target["entry"], f"{context}[0] entry")
    _integer(target["logical_bytes"], f"{context}[0] logical_bytes", minimum=0)
    _hash(target["sha256"], f"{context}[0] sha256")
    expected = {
        "entry": legacy["target_entry"],
        "logical_bytes": legacy["target_payload_bytes"],
        "sha256": legacy["target_payload_sha256"],
    }
    if target != expected:
        raise ValidationError(f"{context}[0] does not match legacy_v1 target")


def _validate_coverage(value: Any, context: str) -> list[str]:
    coverage = _exact_keys(value, COVERAGE_KEYS, context)
    timed = _string_list(coverage["timed_cases"], f"{context} timed_cases")
    guard = _string_list(coverage["guard_cases"], f"{context} guard_cases")
    if coverage["inventory_only"] is not None:
        _boolean(coverage["inventory_only"], f"{context} inventory_only")
    # The schema requires a boolean, including false; the explicit branch above
    # keeps the resulting error a ValidationError for null as well.
    if not isinstance(coverage["inventory_only"], bool):
        raise ValidationError(f"{context} inventory_only must be a boolean")
    if timed != sorted(set(timed)):
        raise ValidationError(f"{context} timed_cases must be sorted and unique")
    if len(guard) != len(set(guard)):
        raise ValidationError(f"{context} guard_cases must be unique")
    return timed


def _validate_corpus(value: Any, context: str) -> dict[str, Any]:
    corpus = _exact_keys(value, CORPUS_KEYS, context)
    identifier = _string(corpus["id"], f"{context} id")
    legacy = _validate_legacy(corpus["legacy_v1"], f"{context} legacy_v1")
    expected_identifier = _corpus_id(legacy, f"{context} legacy_v1")
    if identifier != expected_identifier:
        raise ValidationError(
            f"{context} id does not match package_format/archive_sha256: {identifier!r}"
        )
    _string(corpus["name"], f"{context} name")
    _string(corpus["format"], f"{context} format")
    _enum(corpus["size_class"], f"{context} size_class", {"tiny", "small", "medium", "large", "very-large", "unknown"})
    categories = _string_list(corpus["categories"], f"{context} categories")
    if categories != sorted(set(categories)):
        raise ValidationError(f"{context} categories must be sorted and unique")
    _validate_generator(corpus["generator"], f"{context} generator")
    _validate_provenance(corpus["provenance"], f"{context} provenance")
    bytes_summary = _validate_bytes(corpus["bytes"], f"{context} bytes")
    _object(corpus["shape_parameters"], f"{context} shape_parameters")
    for key, parameter in corpus["shape_parameters"].items():
        if not isinstance(key, str):
            raise ValidationError(f"{context} shape_parameters keys must be strings")
        _integer(parameter, f"{context} shape_parameters {key!r}", minimum=0)
    expected_shape_parameters = {
        "entry_count": legacy["entry_count"],
        "entry_bytes": legacy["entry_bytes"],
    }
    if legacy["xlsx"] is not None:
        for key in ("sheet_count", "rows_per_sheet", "columns_per_sheet"):
            expected_shape_parameters[key] = legacy["xlsx"][key]
    if corpus["shape_parameters"] != expected_shape_parameters:
        raise ValidationError(f"{context} shape_parameters do not match legacy_v1")
    _validate_relationships(corpus["relationships"], f"{context} relationships")
    _validate_security(corpus["security"], f"{context} security")
    _validate_input(corpus["input"], f"{context} input")
    _validate_limits(corpus["limits"], f"{context} limits")
    _validate_members(corpus["members"], f"{context} members")
    _validate_targets(corpus["targets"], legacy, f"{context} targets")
    _validate_coverage(corpus["coverage"], f"{context} coverage")

    if corpus["name"] != legacy["name"]:
        raise ValidationError(f"{context} name does not match legacy_v1")
    if corpus["format"] != legacy["package_format"]:
        raise ValidationError(f"{context} format does not match legacy_v1 package_format")
    for key in ("archive_bytes", "logical_payload_bytes"):
        legacy_key = "archive_bytes" if key == "archive_bytes" else "uncompressed_payload_bytes"
        if bytes_summary[key] != legacy[legacy_key]:
            raise ValidationError(f"{context} bytes {key} does not match legacy_v1")
    if bytes_summary["archive_sha256"] != legacy["archive_sha256"]:
        raise ValidationError(f"{context} bytes archive_sha256 does not match legacy_v1")
    return {
        "id": identifier,
        "archive_sha256": bytes_summary["archive_sha256"],
        "members": [
            {
                "ordinal": member["ordinal"],
                "name": member["name"],
                "sha256": member["sha256"],
            }
            for member in corpus["members"]["items"]
        ],
    }


def _content_set(catalog: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    corpora = _list(_required(catalog, "corpora", "catalog"), "catalog corpora")
    bindings = _list(_required(catalog, "case_bindings", "catalog"), "catalog case_bindings")
    if not corpora:
        raise ValidationError("catalog corpora must be a non-empty list")
    if not bindings:
        raise ValidationError("catalog case_bindings must be a non-empty list")

    corpus_ids: set[str] = set()
    projected_corpora: list[dict[str, Any]] = []
    corpus_by_id: dict[str, dict[str, Any]] = {}
    for index, value in enumerate(corpora):
        context = f"catalog corpus {index}"
        projected = _validate_corpus(value, context)
        identifier = projected["id"]
        if identifier in corpus_ids:
            raise ValidationError(f"catalog corpus id is duplicated: {identifier!r}")
        corpus_ids.add(identifier)
        corpus_by_id[identifier] = _object(value, context)
        projected_corpora.append(projected)

    identifiers = [corpus["id"] for corpus in projected_corpora]
    if identifiers != sorted(identifiers):
        raise ValidationError("catalog corpora are not sorted by id")

    projected_bindings: list[dict[str, Any]] = []
    binding_keys: set[tuple[str, str]] = set()
    timed_cases_by_corpus: dict[str, list[str]] = {identifier: [] for identifier in corpus_ids}
    for index, value in enumerate(bindings):
        context = f"catalog binding {index}"
        binding = _exact_keys(value, BINDING_KEYS, context)
        case = _string(binding["case"], f"{context} case")
        identifier = _string(binding["corpus_id"], f"{context} corpus_id")
        legacy_name = _string(binding["legacy_name"], f"{context} legacy_name")
        archive_sha256 = _hash(
            binding["legacy_archive_sha256"],
            f"{context} legacy_archive_sha256",
        )
        role = _enum(binding["role"], f"{context} role", {"timed", "guard", "inventory-only"})
        if identifier not in corpus_ids:
            raise ValidationError(f"{context} references unknown corpus {identifier!r}")
        corpus = corpus_by_id[identifier]
        legacy = corpus["legacy_v1"]
        if legacy_name != legacy["name"]:
            raise ValidationError(f"{context} legacy_name does not match referenced corpus")
        if archive_sha256 != legacy["archive_sha256"]:
            raise ValidationError(f"{context} legacy_archive_sha256 does not match referenced corpus")
        key = (case, identifier)
        if key in binding_keys:
            raise ValidationError(f"duplicate catalog binding {key!r}")
        binding_keys.add(key)
        if role == "timed":
            timed_cases_by_corpus[identifier].append(case)
        projected_bindings.append({"case": case, "corpus_id": identifier, "role": role})
    binding_order = [(item["case"], item["corpus_id"]) for item in projected_bindings]
    if binding_order != sorted(binding_order):
        raise ValidationError("catalog case_bindings are not sorted")

    for identifier, corpus in corpus_by_id.items():
        expected_timed = sorted(set(timed_cases_by_corpus[identifier]))
        actual_timed = corpus["coverage"]["timed_cases"]
        if actual_timed != expected_timed:
            raise ValidationError(
                f"catalog corpus {identifier} coverage timed_cases do not match timed bindings"
            )
    return projected_corpora, projected_bindings


def _validate_catalog_header(catalog: Any) -> dict[str, Any]:
    catalog = _object(catalog, "catalog")
    if set(catalog) != CATALOG_KEYS:
        raise ValidationError("catalog top-level schema does not match schema 2")
    if _integer(_required(catalog, "manifest_version", "catalog"), "catalog manifest_version") != 2:
        raise ValidationError("catalog manifest_version must be 2")
    if _string(_required(catalog, "manifest_kind", "catalog"), "catalog manifest_kind") != "corpus-catalog":
        raise ValidationError("catalog manifest_kind is unexpected")
    _string(_required(catalog, "catalog_id", "catalog"), "catalog catalog_id")
    canonicalization = _exact_keys(
        _required(catalog, "canonicalization", "catalog"),
        CANONICALIZATION_KEYS,
        "catalog canonicalization",
    )
    if _string(canonicalization["algorithm"], "catalog canonicalization algorithm") != "sorted-json-utf8-compact-v1":
        raise ValidationError("catalog canonicalization is unsupported")
    if _string(canonicalization["hash"], "catalog canonicalization hash") != "sha256":
        raise ValidationError("catalog canonicalization is unsupported")
    _validate_build(_required(catalog, "build", "catalog"), "catalog build")
    return catalog


def _validate_reference(value: Any, catalog: dict[str, Any], catalog_sha256: str, content_set_sha256: str) -> None:
    reference = _exact_keys(value, REFERENCE_KEYS, "report corpus_catalog")
    if _integer(reference["manifest_version"], "report corpus_catalog manifest_version") != catalog["manifest_version"]:
        raise ValidationError("report corpus_catalog manifest_version does not match sidecar")
    _string(reference["catalog_id"], "report corpus_catalog catalog_id")
    _hash(reference["catalog_sha256"], "report corpus_catalog catalog_sha256")
    _hash(reference["content_set_sha256"], "report corpus_catalog content_set_sha256")
    expected = {
        "manifest_version": catalog["manifest_version"],
        "catalog_id": catalog["catalog_id"],
        "catalog_sha256": catalog_sha256,
        "content_set_sha256": content_set_sha256,
    }
    if reference != expected:
        raise ValidationError("report corpus_catalog reference does not match sidecar")


def validate_binding(report: dict[str, Any], catalog: dict[str, Any]) -> tuple[int, int]:
    """Validate and return ``(corpus_count, binding_count)``.

    Every failure raises ``ValidationError``.  In particular, the report's
    result corpus objects and case bindings are compared to the catalog, not
    merely the catalog's self-reported hashes.
    """

    report = _object(report, "report")
    if _integer(_required(report, "schema_version", "report"), "report schema_version") != 1:
        raise ValidationError("report schema_version must remain 1")
    results = _list(_required(report, "results", "report"), "report results")
    if not results:
        raise ValidationError("report results must be a non-empty list")

    catalog = _validate_catalog_header(catalog)
    catalog_sha256 = _hash(
        _required(catalog, "catalog_sha256", "catalog"),
        "catalog catalog_sha256",
    )
    content_set_sha256 = _hash(
        _required(catalog, "content_set_sha256", "catalog"),
        "catalog content_set_sha256",
    )
    catalog_without_hash = dict(catalog)
    catalog_without_hash.pop("catalog_sha256")
    if _digest(catalog_without_hash) != catalog_sha256:
        raise ValidationError("catalog_sha256 does not match catalog content")

    projected_corpora, projected_bindings = _content_set(catalog)
    if _digest({"corpora": projected_corpora, "case_bindings": projected_bindings}) != content_set_sha256:
        raise ValidationError("content_set_sha256 does not match catalog content")

    catalog_corpora = _list(_required(catalog, "corpora", "catalog"), "catalog corpora")
    catalog_by_id = {
        _string(corpus["id"], f"catalog corpus {index} id"): corpus
        for index, corpus in enumerate(catalog_corpora)
    }
    expected_bindings: list[tuple[str, str, str, str, str]] = []
    for index, value in enumerate(results):
        context = f"report result {index}"
        result = _object(value, context)
        case = _string(_required(result, "case", context), f"{context} case")
        corpus = _object(_required(result, "corpus", context), f"{context} corpus")
        _validate_legacy(corpus, f"{context} corpus")
        identifier = _corpus_id(corpus, f"{context} corpus")
        catalog_corpus = catalog_by_id.get(identifier)
        if catalog_corpus is None:
            raise ValidationError(f"{context} is absent from catalog: {identifier}")
        if _required(catalog_corpus, "legacy_v1", f"catalog corpus {identifier}") != corpus:
            raise ValidationError(f"{context} corpus differs from catalog legacy_v1")
        expected_bindings.append(
            (
                case,
                identifier,
                _string(_required(corpus, "name", f"{context} corpus"), f"{context} corpus name"),
                _hash(
                    _required(corpus, "archive_sha256", f"{context} corpus"),
                    f"{context} corpus archive_sha256",
                ),
                "timed",
            )
        )
    expected_bindings.sort(key=lambda binding: (binding[0], binding[1]))
    actual_bindings = []
    for index, value in enumerate(_list(catalog["case_bindings"], "catalog case_bindings")):
        binding = _exact_keys(value, BINDING_KEYS, f"catalog binding {index}")
        actual_bindings.append(
            (
                _string(binding["case"], f"catalog binding {index} case"),
                _string(binding["corpus_id"], f"catalog binding {index} corpus_id"),
                _string(binding["legacy_name"], f"catalog binding {index} legacy_name"),
                _hash(
                    binding["legacy_archive_sha256"],
                    f"catalog binding {index} legacy_archive_sha256",
                ),
                _enum(binding["role"], f"catalog binding {index} role", {"timed", "guard", "inventory-only"}),
            )
        )
    if actual_bindings != expected_bindings:
        raise ValidationError("catalog case bindings are not bound to report results")

    _validate_reference(
        _required(report, "corpus_catalog", "report"),
        catalog,
        catalog_sha256,
        content_set_sha256,
    )
    return len(catalog_by_id), len(actual_bindings)


def validate_paths(report_path: Path, catalog_path: Path) -> tuple[int, int]:
    return validate_binding(_read_object(report_path), _read_object(catalog_path))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--catalog", type=Path, required=True)
    args = parser.parse_args()
    try:
        corpus_count, binding_count = validate_paths(args.report, args.catalog)
    except ValidationError as error:
        parser.exit(1, f"corpus catalog validation failed: {error}\n")
    print(
        f"validated schema-2 corpus catalog: {corpus_count} corpora, "
        f"{binding_count} bindings"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
