#!/usr/bin/env python3
"""Strict validation for the non-iWork CRUD coverage index.

The index is deliberately a small, repository-facing contract rather than a
second performance report.  This checker keeps the taxonomy, the checked
schema-2 catalog, and the Rust harness selector registry bound together so a
stale or invented coverage row cannot silently pass CI.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
from pathlib import Path
from typing import Any


SHA256 = re.compile(r"^[0-9a-f]{64}$")
SELECTOR = re.compile(r"^[a-z0-9]+(?:_[a-z0-9]+)*$")
FORBIDDEN_IWORK = re.compile(r"(?:iwork|iwa|keynote|numbers|pages)", re.IGNORECASE)
STATUSES = ("measured", "correctness-only", "unsupported", "not-applicable")

# These are the fifteen Phase-1 program categories.  The checklist supplies
# the auditable row vocabulary through each category's ``checklist_refs``;
# keeping this ordered list here also makes additions/removals fail closed.
EXPECTED_CATEGORIES = (
    ("content-reading-extraction", "Content reading and extraction"),
    ("structural-queries-analysis", "Structural queries and document analysis"),
    ("conversion-export", "Conversion and export"),
    ("creation-from-scratch", "Creation from scratch"),
    ("template-filling-targeted-replacement", "Template filling and targeted replacement"),
    ("append-incremental-generation", "Append-only and incremental generation"),
    ("structural-editing", "Structural editing"),
    ("content-deletion-sanitization", "Content deletion and document sanitization"),
    ("cross-document-copying-assembly", "Cross-document copying and assembly"),
    ("merging-splitting", "Merging and splitting"),
    ("comparison-patching-three-way-merge", "Comparison, patching, and three-way merge"),
    ("validation-repair-normalization", "Validation, repair, and normalization"),
    ("dynamic-content-calculation-refresh", "Dynamic content calculation and refresh"),
    ("security-protection-encryption-signing", "Security, protection, encryption, and signing"),
    (
        "low-level-package-part-relationship-stream-extension",
        "Low-level package, Part, Relationship, stream, and extension operations",
    ),
)

STATUS_SEMANTICS = {
    "measured": "default-baseline timing contract: static identity binding is not timing evidence; scheduled/manual full runs must validate actual report rows",
    "correctness-only": "selector or API evidence is retained without a baseline timing claim; emitted timing is not claimed here",
    "unsupported": "no selector-backed evidence is claimed and an explicit reason is required",
    "not-applicable": "the capability is deliberately outside this non-iWork coverage scope and an explicit reason is required",
}

COVERAGE_CLAIM = (
    "representative checklist mappings; this index is not an exhaustive certification "
    "of every checklist row"
)
CORRECTNESS_ONLY_REQUIRED_PHRASE = "no retained baseline timing claim"
IDENTITY_ARTIFACT = "docs/performance/results/perf-regression-default-manifest-v1.json"
CHECKED_CATALOG_PATH = "docs/performance/results/perf-corpus-manifest-v2.json"
RUN_REPORT_PATH = "target/perf/container-baseline.json"
MINIMUM_SAMPLES = 15
U64_MAX = (1 << 64) - 1
STATISTICS_FLOAT_REL_TOL = 1e-12
STATISTICS_FLOAT_ABS_TOL = 1e-12
STUDENT_T_METHOD = "two-sided Student's t interval for the mean"
STUDENT_T_CRITICAL_95 = (
    12.706,
    4.303,
    3.182,
    2.776,
    2.571,
    2.447,
    2.365,
    2.306,
    2.262,
    2.228,
    2.201,
    2.179,
    2.160,
    2.145,
    2.131,
    2.120,
    2.110,
    2.101,
    2.093,
    2.086,
    2.080,
    2.074,
    2.069,
    2.064,
    2.060,
    2.056,
    2.052,
    2.048,
    2.045,
    2.042,
)
GENERATED_IDENTITY = (
    "harness-generated schema-2 corpus IDs use format:sha256:<archive_sha256>"
)
GENERATED_SELECTOR_SHAPES = {
    "rtf_semantic_text_to_sink": ("large", "medium", "tiny"),
    "odt_semantic_text_to_sink": ("large", "medium", "tiny"),
    "ods_semantic_text_to_sink": ("large", "medium", "tiny"),
    "odp_semantic_text_to_sink": ("large", "medium", "tiny"),
    "xlsx_streaming_create": ("large", "medium", "tiny"),
    "rtf_streaming_create": ("large", "medium", "tiny"),
    "xlsx_eager_defined_names_edit_save": ("media-rich",),
    "xlsx_eager_row_visibility_edit_save": ("large", "medium"),
    "xlsx_eager_cell_remove_edit_save": ("dense-sparse", "medium"),
    "docx_story_hyperlink_redaction_save": ("media-rich",),
    "pptx_source_backed_cross_copy_plain": ("plain",),
    "pptx_cross_copy_media_rich": ("media-rich",),
    "xlsx_eager_merge_commit_save": ("sparse-a1-b2",),
    "rtf_semantic_split_paragraph_save": ("large", "medium", "tiny"),
    "xlsx_join_disjoint_commit_save": ("medium",),
    "xlsx_three_way_disjoint_commit_save": ("medium",),
    "odf_validation_report": ("large", "medium", "tiny"),
    "odf_mimetype_repair_plan": ("large", "medium", "tiny"),
    "xls_validation_report": ("large", "medium", "tiny"),
    "xlsx_eager_sheet_protection_edit_save": ("media-rich",),
}
APPROVED_NAVIGATION_PATHS = {
    "tools/perf-baseline/README.md",
    "docs/performance/CRUD_COVERAGE.md",
}

IDENTITY_KEYS = {
    "schema_version",
    "manifest_kind",
    "harness_source",
    "report_schema_version",
    "source_report_samples_per_case",
    "source_report_warmup_iterations_per_case",
    "result_count",
    "case_count",
    "default_cases",
    "identity_configuration",
    "case_corpora",
    "corpora",
    "canonicalization",
    "result_keys_sha256",
}
IDENTITY_CORPUS_KEYS = {
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
IDENTITY_CORPUS_OPTIONAL_KEYS = {"rtf_variant"}
ELAPSED_KEYS = {
    "unit",
    "samples",
    "sample_order",
    "min",
    "p50",
    "p95",
    "p99",
    "max",
    "mean",
    "standard_deviation",
    "confidence_interval_95",
}
CONFIDENCE_INTERVAL_KEYS = {"method", "lower", "upper"}
IDENTITY_XLSX_KEYS = {
    "sheet_count",
    "rows_per_sheet",
    "columns_per_sheet",
    "one_percent_update_count",
    "source_members",
}
IDENTITY_XLSX_SOURCE_MEMBER_KEYS = {
    "workbook",
    "worksheets",
    "shared_strings",
    "styles",
}

INDEX_KEYS = {
    "schema_version",
    "index_kind",
    "scope",
    "taxonomy",
    "selector_registry",
    "checked_catalog",
    "status_values",
    "status_semantics",
    "coverage_claim",
    "categories",
}
SCOPE_KEYS = {"included_formats", "excluded_formats"}
TAXONOMY_KEYS = {"source", "category_count", "category_ids"}
REGISTRY_KEYS = {"source", "minimum_selectable_cases"}
CATALOG_REFERENCE_KEYS = {
    "path",
    "catalog_id",
    "catalog_sha256",
    "content_set_sha256",
}
REPORT_REFERENCE_KEYS = {
    "manifest_version",
    "catalog_id",
    "catalog_sha256",
    "content_set_sha256",
}
CATEGORY_KEYS = {
    "id",
    "name",
    "checklist_refs",
    "status",
    "measurement",
    "coverage_scope",
    "scenarios",
}
FULL_RUN_MEASUREMENT_KEYS = {
    "kind",
    "identity_artifact",
    "run_report",
    "minimum_samples",
}
CHECKED_CORPUS_KEYS = {"kind", "case", "ids", "shapes"}
GENERATED_CORPUS_KEYS = {"kind", "shapes", "identity"}
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
BINDING_KEYS = {
    "case",
    "corpus_id",
    "legacy_name",
    "legacy_archive_sha256",
    "role",
}


class ValidationError(ValueError):
    """A malformed or internally inconsistent CRUD coverage index."""


def _no_duplicate_object_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValidationError(f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _read_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_object_pairs,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, TypeError, ValidationError) as error:
        raise ValidationError(f"cannot read {path}: {error}") from error
    return _object(value, str(path))


def _read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except (OSError, UnicodeError, TypeError) as error:
        raise ValidationError(f"cannot read {path}: {error}") from error


def _object(value: Any, context: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ValidationError(f"{context} must be an object")
    return value


def _list(value: Any, context: str) -> list[Any]:
    if not isinstance(value, list):
        raise ValidationError(f"{context} must be a list")
    return value


def _exact(
    value: Any,
    required: set[str],
    context: str,
    optional: set[str] | None = None,
) -> dict[str, Any]:
    mapping = _object(value, context)
    optional = optional or set()
    allowed = required | optional
    try:
        keys = set(mapping)
    except TypeError as error:
        raise ValidationError(f"{context} keys are not hashable") from error
    if any(not isinstance(key, str) for key in keys):
        raise ValidationError(f"{context} keys must be strings")
    missing = sorted(required - keys)
    extra = sorted(keys - allowed)
    if missing:
        raise ValidationError(f"{context} is missing {missing[0]!r}")
    if extra:
        raise ValidationError(f"{context} has unexpected field {extra[0]!r}")
    return mapping


def _required(mapping: dict[str, Any], key: str, context: str) -> Any:
    if key not in mapping:
        raise ValidationError(f"{context} is missing {key!r}")
    return mapping[key]


def _string(value: Any, context: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str):
        raise ValidationError(f"{context} must be a string")
    if nonempty and not value:
        raise ValidationError(f"{context} must not be empty")
    return value


def _integer(value: Any, context: str, *, minimum: int | None = None) -> int:
    if isinstance(value, bool) or not isinstance(value, int):
        raise ValidationError(f"{context} must be an integer")
    if minimum is not None and value < minimum:
        raise ValidationError(f"{context} must be at least {minimum}")
    return value


def _u64(value: Any, context: str) -> int:
    value = _integer(value, context, minimum=0)
    if value > U64_MAX:
        raise ValidationError(f"{context} must fit an unsigned 64-bit integer")
    return value


def _hash(value: Any, context: str) -> str:
    if not isinstance(value, str) or SHA256.fullmatch(value) is None:
        raise ValidationError(f"{context} must be a lowercase SHA-256")
    return value


def _enum(value: Any, context: str, values: set[str]) -> str:
    value = _string(value, context)
    if value not in values:
        raise ValidationError(f"{context} has unsupported value {value!r}")
    return value


def _string_list(value: Any, context: str, *, sort_unique: bool = False) -> list[str]:
    values = _list(value, context)
    result = [
        _string(item, f"{context}[{index}]") for index, item in enumerate(values)
    ]
    if sort_unique and result != sorted(set(result)):
        raise ValidationError(f"{context} must be sorted and unique")
    return result


def _canonical_bytes(value: Any) -> bytes:
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
    return hashlib.sha256(_canonical_bytes(value)).hexdigest()


def _format_slug(value: str) -> str:
    output: list[str] = []
    for character in value:
        if character.isascii() and character.isalnum():
            output.append(character.lower())
        elif not output or output[-1] != "-":
            output.append("-")
    result = "".join(output).strip("-")
    if not result:
        raise ValidationError("catalog package_format has no ASCII slug")
    return result


def _corpus_id(legacy: dict[str, Any], context: str) -> str:
    package_format = _string(_required(legacy, "package_format", context), f"{context} package_format")
    archive_sha256 = _hash(
        _required(legacy, "archive_sha256", context),
        f"{context} archive_sha256",
    )
    return f"{_format_slug(package_format)}:sha256:{archive_sha256}"


def _validate_catalog(
    catalog: dict[str, Any],
) -> tuple[dict[str, dict[str, Any]], dict[str, list[str]], dict[tuple[str, str], str]]:
    catalog = _exact(catalog, CATALOG_KEYS, "catalog")
    if _integer(catalog["manifest_version"], "catalog manifest_version") != 2:
        raise ValidationError("catalog manifest_version must be 2")
    if _string(catalog["manifest_kind"], "catalog manifest_kind") != "corpus-catalog":
        raise ValidationError("catalog manifest_kind is unexpected")
    catalog_id = _string(catalog["catalog_id"], "catalog catalog_id")
    canonicalization = _exact(
        catalog["canonicalization"], {"algorithm", "hash"}, "catalog canonicalization"
    )
    if canonicalization != {
        "algorithm": "sorted-json-utf8-compact-v1",
        "hash": "sha256",
    }:
        raise ValidationError("catalog canonicalization is unsupported")
    catalog_sha256 = _hash(catalog["catalog_sha256"], "catalog catalog_sha256")
    content_set_sha256 = _hash(
        catalog["content_set_sha256"], "catalog content_set_sha256"
    )
    without_hash = dict(catalog)
    without_hash.pop("catalog_sha256")
    if _digest(without_hash) != catalog_sha256:
        raise ValidationError("catalog_sha256 does not match catalog content")

    corpora = _list(catalog["corpora"], "catalog corpora")
    if not corpora:
        raise ValidationError("catalog corpora must not be empty")
    corpus_by_id: dict[str, dict[str, Any]] = {}
    projected_corpora: list[dict[str, Any]] = []
    previous_id = ""
    for index, value in enumerate(corpora):
        context = f"catalog corpus {index}"
        corpus = _object(value, context)
        identifier = _string(_required(corpus, "id", context), f"{context} id")
        if not re.fullmatch(r"[a-z0-9-]+:sha256:[0-9a-f]{64}", identifier):
            raise ValidationError(f"{context} id is malformed")
        if identifier in corpus_by_id:
            raise ValidationError(f"{context} id is duplicated")
        if identifier <= previous_id:
            raise ValidationError("catalog corpora must be sorted by id")
        previous_id = identifier
        legacy = _object(_required(corpus, "legacy_v1", context), f"{context} legacy_v1")
        for key in ("name", "shape", "package_format"):
            _string(_required(legacy, key, f"{context} legacy_v1"), f"{context} legacy_v1 {key}")
        legacy_archive = _hash(
            _required(legacy, "archive_sha256", f"{context} legacy_v1"),
            f"{context} legacy_v1 archive_sha256",
        )
        expected_id = _corpus_id(legacy, f"{context} legacy_v1")
        if identifier != expected_id:
            raise ValidationError(f"{context} id does not match format and archive hash")
        if _string(_required(corpus, "name", context), f"{context} name") != legacy["name"]:
            raise ValidationError(f"{context} name does not match legacy_v1")
        if _string(_required(corpus, "format", context), f"{context} format") != legacy["package_format"]:
            raise ValidationError(f"{context} format does not match legacy_v1")
        bytes_summary = _object(_required(corpus, "bytes", context), f"{context} bytes")
        bytes_archive = _hash(
            _required(bytes_summary, "archive_sha256", f"{context} bytes"),
            f"{context} bytes archive_sha256",
        )
        if bytes_archive != legacy_archive:
            raise ValidationError(f"{context} bytes archive_sha256 does not match legacy_v1")
        if "archive_bytes" in bytes_summary and "archive_bytes" in legacy:
            if bytes_summary["archive_bytes"] != legacy["archive_bytes"]:
                raise ValidationError(f"{context} bytes archive_bytes does not match legacy_v1")

        members = _object(_required(corpus, "members", context), f"{context} members")
        items = _list(_required(members, "items", f"{context} members"), f"{context} members items")
        member_projection: list[dict[str, Any]] = []
        previous_ordinal = -1
        member_names: set[str] = set()
        for member_index, member_value in enumerate(items):
            member_context = f"{context} members items[{member_index}]"
            member = _object(member_value, member_context)
            ordinal = _integer(
                _required(member, "ordinal", member_context),
                f"{member_context} ordinal",
                minimum=0,
            )
            if ordinal <= previous_ordinal:
                raise ValidationError(f"{context} member ordinals must be strictly increasing")
            previous_ordinal = ordinal
            name = _string(_required(member, "name", member_context), f"{member_context} name")
            if name in member_names:
                raise ValidationError(f"{member_context} name is duplicated")
            member_names.add(name)
            member_hash = _hash(
                _required(member, "sha256", member_context), f"{member_context} sha256"
            )
            member_projection.append(
                {"ordinal": ordinal, "name": name, "sha256": member_hash}
            )
        corpus_by_id[identifier] = corpus
        projected_corpora.append(
            {
                "id": identifier,
                "archive_sha256": legacy_archive,
                "members": member_projection,
            }
        )

    bindings = _list(catalog["case_bindings"], "catalog case_bindings")
    if not bindings:
        raise ValidationError("catalog case_bindings must not be empty")
    projected_bindings: list[dict[str, str]] = []
    binding_keys: set[tuple[str, str]] = set()
    binding_roles: dict[tuple[str, str], str] = {}
    cases: dict[str, list[str]] = {}
    previous_binding: tuple[str, str] | None = None
    for index, value in enumerate(bindings):
        context = f"catalog binding {index}"
        binding = _exact(value, BINDING_KEYS, context)
        case = _string(binding["case"], f"{context} case")
        identifier = _string(binding["corpus_id"], f"{context} corpus_id")
        legacy_name = _string(binding["legacy_name"], f"{context} legacy_name")
        legacy_archive = _hash(
            binding["legacy_archive_sha256"], f"{context} legacy_archive_sha256"
        )
        role = _enum(binding["role"], f"{context} role", {"timed", "guard", "inventory-only"})
        if identifier not in corpus_by_id:
            raise ValidationError(f"{context} references an unknown corpus")
        corpus = corpus_by_id[identifier]
        legacy = _object(corpus["legacy_v1"], f"{context} referenced legacy_v1")
        if legacy_name != legacy["name"] or legacy_archive != legacy["archive_sha256"]:
            raise ValidationError(f"{context} legacy identity does not match its corpus")
        key = (case, identifier)
        if key in binding_keys:
            raise ValidationError(f"{context} is duplicated")
        if previous_binding is not None and key <= previous_binding:
            raise ValidationError("catalog case_bindings must be sorted")
        previous_binding = key
        binding_keys.add(key)
        binding_roles[key] = role
        cases.setdefault(case, []).append(identifier)
        projected_bindings.append(
            {"case": case, "corpus_id": identifier, "role": role}
        )

    for case, identifiers in cases.items():
        if identifiers != sorted(identifiers):
            raise ValidationError(f"catalog bindings for {case!r} are not sorted")
    expected_content = {
        "corpora": projected_corpora,
        "case_bindings": projected_bindings,
    }
    if _digest(expected_content) != content_set_sha256:
        raise ValidationError("content_set_sha256 does not match catalog content")
    if catalog_id != "litchi-perf-corpus-v2":
        raise ValidationError("unexpected checked catalog_id")
    return corpus_by_id, cases, binding_roles


def _balanced_body(source: str, start: int, context: str) -> str:
    brace = source.find("{", start)
    if brace < 0:
        raise ValidationError(f"{context} has no body")
    depth = 0
    for index in range(brace, len(source)):
        if source[index] == "{":
            depth += 1
        elif source[index] == "}":
            depth -= 1
            if depth == 0:
                return source[brace : index + 1]
    raise ValidationError(f"{context} has an unterminated body")


def _selector_names(source: str) -> set[str]:
    enum_start = source.find("enum Case {")
    if enum_start < 0:
        raise ValidationError("selector registry has no Case enum")
    enum_end_marker = "\n}\n\nimpl Case"
    enum_end = source.find(enum_end_marker, enum_start)
    if enum_end < 0:
        raise ValidationError("selector registry Case enum boundary is missing")
    variants = {
        match.group(1)
        for match in re.finditer(
            r"^\s*([A-Z][A-Za-z0-9_]*)\s*,\s*$",
            source[enum_start + len("enum Case {") : enum_end],
            re.MULTILINE,
        )
    }
    impl_start = source.find("impl Case", enum_end)
    name_start = source.find("const fn name(self)", impl_start)
    if impl_start < 0 or name_start < 0:
        raise ValidationError("selector registry Case::name is missing")
    name_body = _balanced_body(source, name_start, "selector registry Case::name")
    pairs = re.findall(
        r"\bSelf::([A-Za-z0-9_]+)\s*=>\s*(?:\{\s*)?\"([^\"]+)\"",
        name_body,
        re.DOTALL,
    )
    name_variants = {variant for variant, _ in pairs}
    names = [name for _, name in pairs]
    if name_variants != variants or len(pairs) != len(variants):
        raise ValidationError("selector registry Case::name is not exhaustive")
    if len(set(names)) != len(names):
        raise ValidationError("selector registry contains duplicate selector names")
    if any(SELECTOR.fullmatch(name) is None for name in names):
        raise ValidationError("selector registry contains a malformed selector name")
    return set(names)


def _validate_evidence(value: Any, context: str, repo_root: Path) -> None:
    paths = _string_list(value, context)
    if not paths:
        raise ValidationError(f"{context} must not be empty")
    for path_value in paths:
        if path_value not in APPROVED_NAVIGATION_PATHS:
            raise ValidationError(
                f"{context} contains an unapproved documentation navigation path {path_value!r}"
            )
        if path_value.startswith(("/", "\\")) or "\\" in path_value:
            raise ValidationError(f"{context} contains a non-relative path")
        path = Path(path_value)
        if path.is_absolute() or ".." in path.parts:
            raise ValidationError(f"{context} contains a path traversal")
        if not (repo_root / path).is_file():
            raise ValidationError(f"{context} references missing evidence {path_value!r}")


def _validate_identity_corpus(value: Any, context: str) -> dict[str, Any]:
    corpus = _exact(
        value,
        IDENTITY_CORPUS_KEYS,
        context,
        optional=IDENTITY_CORPUS_OPTIONAL_KEYS,
    )
    for key in (
        "name",
        "generator",
        "package_format",
        "shape",
        "payload_kind",
        "compression",
        "target_entry",
    ):
        _string(corpus[key], f"{context} {key}")
    for key in (
        "entry_count",
        "archive_member_count",
        "entry_bytes",
        "uncompressed_payload_bytes",
        "archive_bytes",
        "target_payload_bytes",
    ):
        _integer(corpus[key], f"{context} {key}", minimum=0)
    _hash(corpus["archive_sha256"], f"{context} archive_sha256")
    _hash(corpus["target_payload_sha256"], f"{context} target_payload_sha256")
    if "rtf_variant" in corpus:
        _string(corpus["rtf_variant"], f"{context} rtf_variant", nonempty=False)
    xlsx = corpus["xlsx"]
    if xlsx is not None:
        xlsx = _exact(xlsx, IDENTITY_XLSX_KEYS, f"{context} xlsx")
        for key in (
            "sheet_count",
            "rows_per_sheet",
            "columns_per_sheet",
            "one_percent_update_count",
        ):
            _integer(xlsx[key], f"{context} xlsx {key}", minimum=0)
        source = _exact(
            xlsx["source_members"],
            IDENTITY_XLSX_SOURCE_MEMBER_KEYS,
            f"{context} xlsx source_members",
        )
        _string(source["workbook"], f"{context} xlsx source_members workbook")
        _string_list(source["worksheets"], f"{context} xlsx source_members worksheets")
        for key in ("shared_strings", "styles"):
            if source[key] is not None:
                _string(source[key], f"{context} xlsx source_members {key}")
    return corpus


def _result_key_digest(
    default_cases: list[str],
    case_corpora: dict[str, list[str]],
    corpora: dict[str, dict[str, Any]],
) -> str:
    keys: list[tuple[str, str]] = []
    for case in default_cases:
        for name in case_corpora[case]:
            try:
                corpus_identity = json.dumps(
                    corpora[name], sort_keys=True, separators=(",", ":"), allow_nan=False
                )
            except (TypeError, ValueError, OverflowError) as error:
                raise ValidationError(f"identity corpus {name!r} is not canonical JSON: {error}") from error
            keys.append((case, corpus_identity))
    digest = hashlib.sha256()
    for case, corpus_identity in sorted(keys):
        digest.update(case.encode("utf-8"))
        digest.update(b"\0")
        digest.update(corpus_identity.encode("utf-8"))
        digest.update(b"\n")
    return digest.hexdigest()


def _validate_identity_artifact(
    identity: dict[str, Any],
    context: str,
) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    identity = _exact(identity, IDENTITY_KEYS, context)
    if _integer(identity["schema_version"], f"{context} schema_version") != 1:
        raise ValidationError(f"{context} schema_version must be 1")
    if _string(identity["manifest_kind"], f"{context} manifest_kind") != "case-corpus-key-identity":
        raise ValidationError(f"{context} manifest_kind is unexpected")
    if _string(identity["harness_source"], f"{context} harness_source") != (
        "tools/perf-baseline/src/main.rs:Case::DEFAULT"
    ):
        raise ValidationError(f"{context} harness_source is unexpected")
    if _integer(identity["report_schema_version"], f"{context} report_schema_version") != 1:
        raise ValidationError(f"{context} report_schema_version must be 1")
    if _integer(identity["source_report_samples_per_case"], f"{context} source_report_samples_per_case") != 1:
        raise ValidationError(f"{context} source_report_samples_per_case must be 1")
    if _integer(identity["source_report_warmup_iterations_per_case"], f"{context} source_report_warmup_iterations_per_case") != 0:
        raise ValidationError(f"{context} source_report_warmup_iterations_per_case must be 0")
    default_cases = _string_list(identity["default_cases"], f"{context} default_cases")
    if not default_cases or len(default_cases) != len(set(default_cases)):
        raise ValidationError(f"{context} default_cases must be nonempty and unique")
    if _integer(identity["case_count"], f"{context} case_count") != len(default_cases):
        raise ValidationError(f"{context} case_count does not match default_cases")
    case_corpora = _object(identity["case_corpora"], f"{context} case_corpora")
    if set(case_corpora) != set(default_cases):
        raise ValidationError(f"{context} case_corpora keys do not match default_cases")
    corpora_value = _object(identity["corpora"], f"{context} corpora")
    corpora: dict[str, dict[str, Any]] = {}
    for name, value in corpora_value.items():
        if not isinstance(name, str) or not name:
            raise ValidationError(f"{context} corpus name must be a nonempty string")
        corpus = _validate_identity_corpus(value, f"{context} corpus {name!r}")
        if corpus["name"] != name:
            raise ValidationError(f"{context} corpus {name!r} name does not match its key")
        corpora[name] = corpus
    referenced_names: set[str] = set()
    result_count = 0
    case_corpora_normalized: dict[str, list[str]] = {}
    for case in default_cases:
        names = _string_list(case_corpora[case], f"{context} case_corpora {case!r}")
        if not names or len(names) != len(set(names)):
            raise ValidationError(f"{context} case_corpora {case!r} must be nonempty and unique")
        unknown = sorted(set(names) - set(corpora))
        if unknown:
            raise ValidationError(f"{context} case_corpora {case!r} references unknown corpus {unknown[0]!r}")
        referenced_names.update(names)
        result_count += len(names)
        case_corpora_normalized[case] = names
    if referenced_names != set(corpora):
        raise ValidationError(f"{context} contains an unreferenced corpus")
    if _integer(identity["result_count"], f"{context} result_count") != result_count:
        raise ValidationError(f"{context} result_count does not match case_corpora")
    canonicalization = _string(identity["canonicalization"], f"{context} canonicalization")
    if canonicalization != "sort (case, compact canonical corpus JSON) keys; SHA-256 UTF-8 case, NUL, corpus JSON, LF":
        raise ValidationError(f"{context} canonicalization is unexpected")
    expected_digest = _hash(identity["result_keys_sha256"], f"{context} result_keys_sha256")
    if _result_key_digest(default_cases, case_corpora_normalized, corpora) != expected_digest:
        raise ValidationError(f"{context} result_keys_sha256 does not match identity content")
    _object(identity["identity_configuration"], f"{context} identity_configuration")
    return {"default_cases": default_cases, "case_corpora": case_corpora_normalized}, corpora


def _validate_identity_bindings(
    identity: tuple[dict[str, Any], dict[str, dict[str, Any]]],
    measured_scenarios: list[tuple[str, list[str]]],
    checked_corpora: dict[str, dict[str, Any]],
    context: str,
) -> None:
    identity_header, identity_corpora = identity
    for selector, identifiers in measured_scenarios:
        names = identity_header["case_corpora"].get(selector)
        if names is None:
            raise ValidationError(f"{context} has no identity entry for measured selector {selector!r}")
        identity_ids: list[str] = []
        for name in names:
            legacy = identity_corpora[name]
            identifier = _corpus_id(legacy, f"{context} corpus {name!r}")
            checked = checked_corpora.get(identifier)
            if checked is None:
                raise ValidationError(f"{context} identity corpus {name!r} is absent from checked catalog")
            if _object(checked["legacy_v1"], f"{context} checked corpus {identifier}") != legacy:
                raise ValidationError(f"{context} identity corpus {name!r} differs from checked catalog")
            identity_ids.append(identifier)
        if sorted(identity_ids) != identifiers:
            raise ValidationError(f"{context} identity corpora do not match measured selector {selector!r}")


def _finite_number(value: Any, context: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValidationError(f"{context} must be a finite number")
    result = float(value)
    if not math.isfinite(result):
        raise ValidationError(f"{context} must be a finite number")
    return result


def _midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + ((left % 2 + right % 2) // 2)


def _nearest_rank(samples: list[int], percentile: int) -> int:
    index = min((percentile * len(samples) + 99) // 100 - 1, len(samples) - 1)
    return samples[index]


def _student_t_critical_95(degrees_of_freedom: int) -> float:
    if degrees_of_freedom == 0:
        return 0.0
    if degrees_of_freedom <= len(STUDENT_T_CRITICAL_95):
        return STUDENT_T_CRITICAL_95[degrees_of_freedom - 1]
    # Keep this expression in the same grouping as the Rust harness.  The
    # fixed table covers the small samples where the correction is largest;
    # this Cornish-Fisher expansion is the harness's large-sample tail.
    z = 1.959_963_984_540_054
    degrees = float(degrees_of_freedom)
    z2 = z * z
    z3 = z2 * z
    z5 = z3 * z2
    z7 = z5 * z2
    return (
        z
        + (z3 + z) / (4.0 * degrees)
        + (5.0 * z5 + 16.0 * z3 + 3.0 * z)
        / (96.0 * degrees * degrees)
        + (3.0 * z7 + 19.0 * z5 + 17.0 * z3 - 15.0 * z)
        / (384.0 * degrees * degrees * degrees)
    )


def _expected_statistics(samples: list[int]) -> dict[str, Any]:
    """Reproduce the Rust ``statistics`` Welford fold and summaries exactly."""
    mean = 0.0
    squared_deviation_sum = 0.0
    for index, sample in enumerate(samples):
        value = float(sample)
        next_count = float(index + 1)
        delta = value - mean
        next_mean = mean + delta / next_count
        squared_deviation_sum = squared_deviation_sum + delta * (value - next_mean)
        mean = next_mean
    count = len(samples)
    standard_deviation = (
        math.sqrt(squared_deviation_sum / float(count - 1))
        if count > 1
        else 0.0
    )
    margin = (
        _student_t_critical_95(count - 1)
        * standard_deviation
        / math.sqrt(float(count))
        if count > 1
        else 0.0
    )
    return {
        "min": samples[0],
        "p50": _midpoint(samples[(count - 1) // 2], samples[count // 2]),
        "p95": _nearest_rank(samples, 95),
        "p99": _nearest_rank(samples, 99),
        "max": samples[-1],
        "mean": mean,
        "standard_deviation": standard_deviation,
        "confidence_interval_95": {
            "method": STUDENT_T_METHOD,
            "lower": max(mean - margin, 0.0),
            "upper": mean + margin,
        },
    }


def _assert_float_matches(actual: float, expected: float, context: str) -> None:
    # serde_json emits the shortest round-trippable f64 representation.  This
    # tiny tolerance allows that parse round trip while rejecting fabricated
    # report statistics (and is intentionally much tighter than measurement
    # noise or a useful performance threshold).
    if not math.isclose(
        actual,
        expected,
        rel_tol=STATISTICS_FLOAT_REL_TOL,
        abs_tol=STATISTICS_FLOAT_ABS_TOL,
    ):
        raise ValidationError(
            f"{context} does not match Rust statistics ({actual!r} != {expected!r})"
        )


def _validate_elapsed(value: Any, context: str, minimum_samples: int) -> None:
    elapsed = _exact(value, ELAPSED_KEYS, context)
    if _string(elapsed["unit"], f"{context} unit") != "ns":
        raise ValidationError(f"{context} unit must be ns")
    samples = [
        _u64(item, f"{context} samples[{index}]")
        for index, item in enumerate(_list(elapsed["samples"], f"{context} samples"))
    ]
    if len(samples) < minimum_samples:
        raise ValidationError(f"{context} has fewer than {minimum_samples} samples")
    if samples != sorted(samples):
        raise ValidationError(f"{context} samples must be sorted")
    sample_order = [
        _integer(item, f"{context} sample_order[{index}]")
        for index, item in enumerate(_list(elapsed["sample_order"], f"{context} sample_order"))
    ]
    if sorted(sample_order) != list(range(len(samples))):
        raise ValidationError(f"{context} sample_order must be a permutation of sample indexes")
    for index in range(1, len(samples)):
        if samples[index] == samples[index - 1] and sample_order[index] <= sample_order[index - 1]:
            raise ValidationError(
                f"{context} sample_order must be increasing for tied sorted samples"
            )
    minimum = _u64(elapsed["min"], f"{context} min")
    p50 = _u64(elapsed["p50"], f"{context} p50")
    p95 = _u64(elapsed["p95"], f"{context} p95")
    p99 = _u64(elapsed["p99"], f"{context} p99")
    maximum = _u64(elapsed["max"], f"{context} max")
    mean = _finite_number(elapsed["mean"], f"{context} mean")
    standard_deviation = _finite_number(
        elapsed["standard_deviation"], f"{context} standard_deviation"
    )
    confidence = _exact(
        elapsed["confidence_interval_95"],
        CONFIDENCE_INTERVAL_KEYS,
        f"{context} confidence_interval_95",
    )
    method = _string(confidence["method"], f"{context} confidence_interval_95 method")
    lower = _finite_number(confidence["lower"], f"{context} confidence_interval_95 lower")
    upper = _finite_number(confidence["upper"], f"{context} confidence_interval_95 upper")
    expected = _expected_statistics(samples)
    for key, actual in (
        ("min", minimum),
        ("p50", p50),
        ("p95", p95),
        ("p99", p99),
        ("max", maximum),
    ):
        if actual != expected[key]:
            raise ValidationError(f"{context} {key} does not match Rust statistics")
    _assert_float_matches(mean, expected["mean"], f"{context} mean")
    _assert_float_matches(
        standard_deviation,
        expected["standard_deviation"],
        f"{context} standard_deviation",
    )
    if method != STUDENT_T_METHOD:
        raise ValidationError(f"{context} confidence interval method is unexpected")
    expected_confidence = expected["confidence_interval_95"]
    _assert_float_matches(
        lower,
        expected_confidence["lower"],
        f"{context} confidence_interval_95 lower",
    )
    _assert_float_matches(
        upper,
        expected_confidence["upper"],
        f"{context} confidence_interval_95 upper",
    )


def _validate_report(
    report: dict[str, Any],
    run_catalog: dict[str, Any],
    checked_catalog: dict[str, Any],
    measured_scenarios: list[tuple[str, list[str]]],
) -> None:
    report = _object(report, "report")
    if _integer(_required(report, "schema_version", "report"), "report schema_version") != 1:
        raise ValidationError("report schema_version must be 1")
    results = _list(_required(report, "results", "report"), "report results")
    if not results:
        raise ValidationError("report results must not be empty")
    run_corpora, _run_cases, run_roles = _validate_catalog(run_catalog)
    checked_corpora, _checked_cases, _checked_roles = _validate_catalog(checked_catalog)
    if run_catalog["manifest_version"] != checked_catalog["manifest_version"]:
        raise ValidationError("run catalog manifest_version does not match checked catalog")
    if run_catalog["catalog_id"] != checked_catalog["catalog_id"]:
        raise ValidationError("run catalog catalog_id does not match checked catalog")
    if run_catalog["content_set_sha256"] != checked_catalog["content_set_sha256"]:
        raise ValidationError("run catalog content_set_sha256 does not match checked catalog")
    reference = _exact(
        _required(report, "corpus_catalog", "report"),
        REPORT_REFERENCE_KEYS,
        "report corpus_catalog",
    )
    if _integer(reference["manifest_version"], "report corpus_catalog manifest_version") != run_catalog["manifest_version"]:
        raise ValidationError("report corpus_catalog manifest_version does not match run catalog")
    if _string(reference["catalog_id"], "report corpus_catalog catalog_id") != run_catalog["catalog_id"]:
        raise ValidationError("report corpus_catalog catalog_id does not match run catalog")
    if _hash(reference["catalog_sha256"], "report corpus_catalog catalog_sha256") != run_catalog["catalog_sha256"]:
        raise ValidationError("report corpus_catalog catalog_sha256 does not match run catalog")
    if _hash(reference["content_set_sha256"], "report corpus_catalog content_set_sha256") != run_catalog["content_set_sha256"]:
        raise ValidationError("report corpus_catalog content_set_sha256 does not match run catalog")

    rows: dict[tuple[str, str], dict[str, Any]] = {}
    for index, value in enumerate(results):
        context = f"report result {index}"
        result = _object(value, context)
        case = _string(_required(result, "case", context), f"{context} case")
        legacy = _validate_identity_corpus(
            _required(result, "corpus", context),
            f"{context} corpus",
        )
        identifier = _corpus_id(legacy, f"{context} corpus")
        catalog_corpus = run_corpora.get(identifier)
        if catalog_corpus is None:
            raise ValidationError(f"{context} corpus is absent from run catalog")
        catalog_legacy = _object(
            _required(catalog_corpus, "legacy_v1", f"run catalog corpus {identifier}"),
            f"run catalog corpus {identifier} legacy_v1",
        )
        if legacy != catalog_legacy:
            raise ValidationError(f"{context} corpus differs from run catalog legacy_v1")
        if run_roles.get((case, identifier)) != "timed":
            raise ValidationError(f"{context} is not a timed run-catalog binding")
        key = (case, identifier)
        if key in rows:
            raise ValidationError(f"{context} duplicates case/corpus identity")
        _validate_elapsed(
            _required(result, "elapsed_ns", context),
            f"{context} elapsed_ns",
            MINIMUM_SAMPLES,
        )
        rows[key] = result

    for selector, identifiers in measured_scenarios:
        for identifier in identifiers:
            key = (selector, identifier)
            if key not in rows:
                raise ValidationError(f"report is missing measured row {selector!r} / {identifier!r}")
            run_corpus = run_corpora.get(identifier)
            checked_corpus = checked_corpora.get(identifier)
            if run_corpus is None or checked_corpus is None:
                raise ValidationError(f"report measured row references an unknown corpus {identifier!r}")
            if run_corpus["legacy_v1"] != checked_corpus["legacy_v1"]:
                raise ValidationError(f"run catalog corpus {identifier!r} differs from checked identity")
            if run_roles.get(key) != "timed":
                raise ValidationError(f"measured binding {selector!r} / {identifier!r} must be timed")


def _validate_corpus_reference(
    value: Any,
    context: str,
    selector: str,
    selector_names: set[str],
    corpus_by_id: dict[str, dict[str, Any]],
    cases: dict[str, list[str]],
) -> None:
    corpus = _object(value, context)
    kind = _string(_required(corpus, "kind", context), f"{context} kind")
    shapes = _string_list(corpus.get("shapes"), f"{context} shapes", sort_unique=True)
    if kind == "checked-catalog":
        corpus = _exact(corpus, CHECKED_CORPUS_KEYS, context)
        case = _string(corpus["case"], f"{context} case")
        if case not in selector_names:
            raise ValidationError(f"{context} case is not a real selector")
        expected_ids = cases.get(case)
        if expected_ids is None:
            raise ValidationError(f"{context} case has no checked catalog binding")
        identifiers = _string_list(corpus["ids"], f"{context} ids", sort_unique=True)
        if identifiers != expected_ids:
            raise ValidationError(
                f"{context} ids must exactly match all checked catalog bindings for {case!r}"
            )
        for identifier in identifiers:
            referenced = _object(
                corpus_by_id[identifier], f"{context} referenced corpus {identifier}"
            )
            legacy = _object(
                _required(referenced, "legacy_v1", f"{context} referenced corpus {identifier}"),
                f"{context} referenced corpus {identifier} legacy_v1",
            )
            package_format = _string(
                _required(legacy, "package_format", f"{context} referenced corpus {identifier} legacy_v1"),
                f"{context} referenced corpus {identifier} package_format",
            )
            if FORBIDDEN_IWORK.search(package_format):
                raise ValidationError(f"{context} references an excluded iWork corpus")
        expected_shapes = sorted(
            {
                _string(
                    _required(
                        _object(corpus_by_id[identifier], f"{context} referenced corpus"),
                        "legacy_v1",
                        f"{context} referenced corpus",
                    )["shape"],
                    f"{context} referenced corpus shape",
                )
                for identifier in identifiers
            }
        )
        if shapes != expected_shapes:
            raise ValidationError(f"{context} shapes do not match checked corpus shapes")
    elif kind == "generated-per-run":
        corpus = _exact(corpus, GENERATED_CORPUS_KEYS, context)
        identity = _string(corpus["identity"], f"{context} identity")
        if identity != GENERATED_IDENTITY:
            raise ValidationError(f"{context} identity is not the harness schema-2 contract")
        expected_shapes = GENERATED_SELECTOR_SHAPES.get(selector)
        if expected_shapes is None:
            raise ValidationError(f"{context} selector has no generated shape contract")
        if shapes != list(expected_shapes):
            raise ValidationError(
                f"{context} shapes do not match the generated shape contract for {selector!r}"
            )
    else:
        raise ValidationError(f"{context} has unsupported corpus kind {kind!r}")
    if any(FORBIDDEN_IWORK.search(shape) for shape in shapes):
        raise ValidationError(f"{context} contains an excluded iWork shape")


def _validate_index(
    index: dict[str, Any],
    catalog: dict[str, Any],
    selector_source: str,
    checklist_source: str,
    repo_root: Path,
    *,
    checked_catalog: dict[str, Any] | None = None,
    report: dict[str, Any] | None = None,
) -> tuple[int, int]:
    index = _exact(index, INDEX_KEYS, "index")
    if _integer(index["schema_version"], "index schema_version") != 1:
        raise ValidationError("index schema_version must be 1")
    if _string(index["index_kind"], "index index_kind") != "crud-coverage-index":
        raise ValidationError("index index_kind is unexpected")
    scope = _exact(index["scope"], SCOPE_KEYS, "index scope")
    if _string(scope["included_formats"], "index scope included_formats") != (
        "all supported non-iWork formats"
    ):
        raise ValidationError("index scope included_formats must remain non-iWork")
    if _string_list(scope["excluded_formats"], "index scope excluded_formats") != ["iWork"]:
        raise ValidationError("index must exclude iWork exactly")

    taxonomy = _exact(index["taxonomy"], TAXONOMY_KEYS, "index taxonomy")
    taxonomy_source = _string(taxonomy["source"], "index taxonomy source")
    if taxonomy_source != "docs/CRUD_Scenario_Checklist.md":
        raise ValidationError("index taxonomy source must be the CRUD checklist")
    if not (repo_root / taxonomy_source).is_file():
        raise ValidationError("index taxonomy source is missing")
    if _integer(taxonomy["category_count"], "index taxonomy category_count") != len(EXPECTED_CATEGORIES):
        raise ValidationError("index taxonomy category_count is not fifteen")
    category_ids = _string_list(taxonomy["category_ids"], "index taxonomy category_ids")
    expected_ids = [category_id for category_id, _ in EXPECTED_CATEGORIES]
    if category_ids != expected_ids:
        raise ValidationError("index taxonomy category_ids do not match the Phase-1 taxonomy")
    checklist_labels = {
        match.group(1).strip()
        for match in re.finditer(
            r"^\|\s*\[ \]\s*([^|]+?)\s*\|", checklist_source, re.MULTILINE
        )
    }
    if not checklist_labels:
        raise ValidationError("CRUD checklist has no check rows")

    registry = _exact(index["selector_registry"], REGISTRY_KEYS, "index selector_registry")
    registry_source_path = _string(registry["source"], "index selector_registry source")
    if registry_source_path != "tools/perf-baseline/src/lib.rs":
        raise ValidationError("index selector registry source is not authoritative")
    minimum = _integer(
        registry["minimum_selectable_cases"],
        "index selector_registry minimum_selectable_cases",
        minimum=408,
    )
    selector_names = _selector_names(selector_source)
    if len(selector_names) < minimum:
        raise ValidationError(
            f"selector registry exposes {len(selector_names)} names, below required {minimum}"
        )

    checked_catalog = catalog if checked_catalog is None else checked_catalog
    corpus_by_id, cases, binding_roles = _validate_catalog(checked_catalog)
    catalog_reference = _exact(
        index["checked_catalog"], CATALOG_REFERENCE_KEYS, "index checked_catalog"
    )
    if _string(catalog_reference["path"], "index checked_catalog path") != (
        "docs/performance/results/perf-corpus-manifest-v2.json"
    ):
        raise ValidationError("index checked catalog path is unexpected")
    if _string(catalog_reference["catalog_id"], "index checked_catalog catalog_id") != checked_catalog["catalog_id"]:
        raise ValidationError("index checked catalog_id does not match catalog")
    if _hash(catalog_reference["catalog_sha256"], "index checked_catalog catalog_sha256") != checked_catalog["catalog_sha256"]:
        raise ValidationError("index checked catalog_sha256 does not match catalog")
    if _hash(catalog_reference["content_set_sha256"], "index checked_catalog content_set_sha256") != checked_catalog["content_set_sha256"]:
        raise ValidationError("index checked content_set_sha256 does not match catalog")

    statuses = _string_list(index["status_values"], "index status_values")
    if statuses != list(STATUSES):
        raise ValidationError("index status_values must enumerate the four status values")
    semantics = _exact(
        index["status_semantics"], set(STATUSES), "index status_semantics"
    )
    if semantics != STATUS_SEMANTICS:
        raise ValidationError("index status_semantics does not match the status contract")
    if _string(index["coverage_claim"], "index coverage_claim") != COVERAGE_CLAIM:
        raise ValidationError("index coverage_claim must state that mappings are representative")
    categories = _list(index["categories"], "index categories")
    if len(categories) != len(EXPECTED_CATEGORIES):
        raise ValidationError("index must contain exactly fifteen categories")
    seen_selectors: set[str] = set()
    measured_scenarios: list[tuple[str, list[str]]] = []
    for index_number, (expected_id, expected_name) in enumerate(EXPECTED_CATEGORIES):
        context = f"index category {index_number}"
        category = _exact(categories[index_number], CATEGORY_KEYS, context)
        if _string(category["id"], f"{context} id") != expected_id:
            raise ValidationError(f"{context} is out of order or has an unexpected id")
        if _string(category["name"], f"{context} name") != expected_name:
            raise ValidationError(f"{context} has an unexpected name")
        refs = _string_list(category["checklist_refs"], f"{context} checklist_refs")
        if not refs:
            raise ValidationError(f"{context} checklist_refs must not be empty")
        if len(refs) != len(set(refs)):
            raise ValidationError(f"{context} checklist_refs must be unique")
        if any(reference not in checklist_labels for reference in refs):
            raise ValidationError(f"{context} references a missing checklist row")
        category_status = _enum(category["status"], f"{context} status", set(STATUSES))
        measurement = _object(category["measurement"], f"{context} measurement")
        coverage_scope = _string(
            category["coverage_scope"], f"{context} coverage_scope"
        )
        if coverage_scope != "representative":
            raise ValidationError(f"{context} coverage_scope must be representative")
        if category_status == "measured":
            measurement = _exact(measurement, FULL_RUN_MEASUREMENT_KEYS, f"{context} measurement")
            if _string(measurement["kind"], f"{context} measurement kind") != "full-run-timing-contract":
                raise ValidationError(
                    f"{context} measured status requires a full-run timing contract"
                )
            identity_artifact = _string(
                measurement["identity_artifact"], f"{context} measurement identity_artifact"
            )
            if identity_artifact != IDENTITY_ARTIFACT:
                raise ValidationError(f"{context} identity_artifact path is unexpected")
            run_report = _string(
                measurement["run_report"], f"{context} measurement run_report"
            )
            if run_report != RUN_REPORT_PATH:
                raise ValidationError(f"{context} run_report path is unexpected")
            if _integer(
                measurement["minimum_samples"],
                f"{context} measurement minimum_samples",
                minimum=MINIMUM_SAMPLES,
            ) != MINIMUM_SAMPLES:
                raise ValidationError(
                    f"{context} measurement minimum_samples must be {MINIMUM_SAMPLES}"
                )
        elif category_status == "correctness-only":
            measurement = _exact(measurement, {"kind", "reason"}, f"{context} measurement")
            if _string(measurement["kind"], f"{context} measurement kind") != "correctness-only":
                raise ValidationError(
                    f"{context} correctness-only status requires correctness-only evidence"
                )
            reason = _string(measurement["reason"], f"{context} measurement reason")
            if CORRECTNESS_ONLY_REQUIRED_PHRASE not in reason:
                raise ValidationError(
                    f"{context} correctness-only measurement must explain absent timing evidence"
                )
        else:
            measurement = _exact(measurement, {"kind", "reason"}, f"{context} measurement")
            if _string(measurement["kind"], f"{context} measurement kind") != category_status:
                raise ValidationError(
                    f"{context} {category_status} status requires matching measurement kind"
                )
            _string(measurement["reason"], f"{context} measurement reason")
        scenarios = _list(category["scenarios"], f"{context} scenarios")
        if not scenarios:
            raise ValidationError(f"{context} scenarios must not be empty")
        for scenario_number, scenario_value in enumerate(scenarios):
            scenario_context = f"{context} scenario {scenario_number}"
            scenario = _object(scenario_value, scenario_context)
            status = _enum(
                _required(scenario, "status", scenario_context),
                f"{scenario_context} status",
                set(STATUSES),
            )
            if status != category_status:
                raise ValidationError(f"{scenario_context} status differs from category status")
            if status in {"measured", "correctness-only"}:
                scenario = _exact(
                    scenario,
                    {"selector", "status", "corpus", "evidence"},
                    scenario_context,
                )
                selector = _string(scenario["selector"], f"{scenario_context} selector")
                if SELECTOR.fullmatch(selector) is None or selector not in selector_names:
                    raise ValidationError(f"{scenario_context} selector is not real")
                if FORBIDDEN_IWORK.search(selector):
                    raise ValidationError(f"{scenario_context} selects an excluded iWork family")
                if selector in seen_selectors:
                    raise ValidationError(f"selector {selector!r} is duplicated in the index")
                seen_selectors.add(selector)
                corpus_kind = _string(
                    _object(scenario["corpus"], f"{scenario_context} corpus").get("kind"),
                    f"{scenario_context} corpus kind",
                )
                if status == "measured" and corpus_kind != "checked-catalog":
                    raise ValidationError(
                        f"{scenario_context} measured status requires a checked catalog corpus"
                    )
                if corpus_kind == "checked-catalog":
                    corpus_case = _string(
                        _object(scenario["corpus"], f"{scenario_context} corpus").get("case"),
                        f"{scenario_context} corpus case",
                    )
                    if corpus_case != selector:
                        raise ValidationError(
                            f"{scenario_context} corpus case must match its selector"
                        )
                _validate_corpus_reference(
                    scenario["corpus"],
                    f"{scenario_context} corpus",
                    selector,
                    selector_names,
                    corpus_by_id,
                    cases,
                )
                _validate_evidence(scenario["evidence"], f"{scenario_context} evidence", repo_root)
                if status == "measured":
                    identifiers = _object(
                        scenario["corpus"], f"{scenario_context} corpus"
                    ).get("ids")
                    measured_scenarios.append(
                        (
                            selector,
                            _string_list(
                                identifiers,
                                f"{scenario_context} corpus ids",
                                sort_unique=True,
                            ),
                        )
                    )
            else:
                _exact(scenario, {"status", "reason"}, scenario_context)
                reason = _string(scenario["reason"], f"{scenario_context} reason")
                if FORBIDDEN_IWORK.search(reason):
                    raise ValidationError(f"{scenario_context} reason must not broaden into iWork")
    identity = _validate_identity_artifact(
        _read_json(repo_root / IDENTITY_ARTIFACT),
        "identity artifact",
    )
    _validate_identity_bindings(
        identity,
        measured_scenarios,
        corpus_by_id,
        "identity artifact",
    )
    for selector, identifiers in measured_scenarios:
        for identifier in identifiers:
            if binding_roles.get((selector, identifier)) != "timed":
                raise ValidationError(
                    f"measured binding {selector!r} / {identifier!r} must be timed"
                )
    if report is not None:
        _validate_report(
            report,
            catalog,
            checked_catalog,
            measured_scenarios,
        )
    return len(categories), len(seen_selectors)


def validate_index(
    index: dict[str, Any],
    catalog: dict[str, Any],
    selector_source: str,
    checklist_source: str,
    repo_root: Path | None = None,
    *,
    checked_catalog: dict[str, Any] | None = None,
    report: dict[str, Any] | None = None,
) -> tuple[int, int]:
    """Validate in-memory documents and return ``(category_count, selectors)``."""

    root = (repo_root or Path(__file__).resolve().parents[1]).resolve()
    try:
        return _validate_index(
            index,
            catalog,
            selector_source,
            checklist_source,
            root,
            checked_catalog=checked_catalog,
            report=report,
        )
    except ValidationError:
        raise
    except (KeyError, TypeError, ValueError, IndexError) as error:
        raise ValidationError(f"malformed coverage index value: {error}") from error


def validate_paths(
    index_path: Path,
    catalog_path: Path,
    selector_path: Path,
    checklist_path: Path,
    repo_root: Path | None = None,
    report_path: Path | None = None,
) -> tuple[int, int]:
    """Validate repository paths and return ``(category_count, selectors)``."""
    root = (repo_root or Path(__file__).resolve().parents[1]).resolve()
    index = _read_json(index_path)
    catalog = _read_json(catalog_path)
    report = _read_json(report_path) if report_path is not None else None
    checked_catalog = None
    if report is not None and catalog_path.resolve() != (root / CHECKED_CATALOG_PATH).resolve():
        checked_catalog = _read_json(root / CHECKED_CATALOG_PATH)
    return validate_index(
        index,
        catalog,
        _read_text(selector_path),
        _read_text(checklist_path),
        repo_root=repo_root,
        checked_catalog=checked_catalog,
        report=report,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--index",
        type=Path,
        default=Path("docs/performance/crud-coverage-index-v1.json"),
    )
    parser.add_argument(
        "--catalog",
        type=Path,
        default=Path("docs/performance/results/perf-corpus-manifest-v2.json"),
    )
    parser.add_argument(
        "--selector-source",
        type=Path,
        default=Path("tools/perf-baseline/src/lib.rs"),
    )
    parser.add_argument(
        "--checklist",
        type=Path,
        default=Path("docs/CRUD_Scenario_Checklist.md"),
    )
    parser.add_argument("--repo-root", type=Path, default=Path("."))
    parser.add_argument(
        "--report",
        type=Path,
        help="optional schema-1 full-run report to validate against the timing contract",
    )
    args = parser.parse_args()
    try:
        category_count, selector_count = validate_paths(
            args.index,
            args.catalog,
            args.selector_source,
            args.checklist,
            repo_root=args.repo_root,
            report_path=args.report,
        )
    except ValidationError as error:
        parser.exit(1, f"CRUD coverage index validation failed: {error}\n")
    suffix = (
        " and bound full-run timing report"
        if args.report is not None
        else " (contract-only; no run timing report supplied)"
    )
    print(
        f"validated non-iWork CRUD coverage index: {category_count} categories, "
        f"{selector_count} mapped selectors{suffix}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
