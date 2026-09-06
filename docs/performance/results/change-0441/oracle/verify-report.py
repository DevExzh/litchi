#!/usr/bin/env python3
"""Independent report/fixture oracle for the 0439 ODP append baseline.

This verifier intentionally validates the serialized report and the corpus
catalog, not a raw archive that the report does not contain.  The Rust runner
must provide the source/output member, manifest, and semantic readback gates;
this file independently regenerates the fixture text, semantic digests, and
opaque payload hash and binds those gates to the report.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path
import re
import sys
from typing import Any, Iterable


SELECTOR = "odp_existing_append_lifecycle"
GENERATOR = "litchi-odp-existing-append-lifecycle-v1"
PACKAGE_FORMAT = "ODP/ODF/ZIP"
MIMETYPE = "application/vnd.oasis.opendocument.presentation"
OPAQUE_PATH = "Opaque/litchi-perf-odp-existing-append-opaque.bin"
OPAQUE_MEDIA_TYPE = "application/octet-stream"
OPAQUE_BYTES = 64 * 1024
OPAQUE_COMPRESSION = "Deflate"
XML_COMPRESSION = "Deflate"
MIMETYPE_COMPRESSION = "Store"
SEMANTIC_DOMAIN = b"litchi-odp-buffered-semantic-v1\0"
ORDER_DOMAIN = b"litchi-odp-existing-append-order-v1\0"
TEXT_CONTRACT = (
    "UTF-8 mixed Unicode/entities/plain text; one interior ASCII space; no CR, "
    "LF, tab, edge space, or repeated interior spaces"
)
SHAPES = {"tiny": 64, "medium": 4_096, "large": 8_192}
EXPECTED_SOURCE_SEMANTIC = {
    "tiny": "167428219fc1603b63046c349a06ef8d0033ac9f720567c5383bb79e379d10d1",
    "medium": "0af38a123018d91665024301a916e134be575c685ee396fd7706101c3d4a6aba",
    "large": "9fb3e3c5f97292e4e19980f5aa1ed00c8eba71e2aacd1bdfb27a354fe059d08a",
}
EXPECTED_MEMBER_NAMES = {
    "mimetype",
    "content.xml",
    "styles.xml",
    "meta.xml",
    OPAQUE_PATH,
    "META-INF/manifest.xml",
}
PINNED_STYLES_SHA256 = "d9881e91085516246a19c30d9e5cde39a8b10d7e42120b135f48f5ca8afef8d2"
PINNED_META_SHA256 = "c7e55a3560c73aa42da85eec4751c3e78b5cc53ff964f50acba6c5cd105e6719"


class ValidationError(ValueError):
    """A report failed a closed-world acceptance gate."""


def _fail(message: str) -> None:
    raise ValidationError(message)


def _object(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        _fail(f"{label}: expected object")
    return value


def _array(value: Any, label: str) -> list[Any]:
    if not isinstance(value, list):
        _fail(f"{label}: expected array")
    return value


def _string(value: Any, label: str, *, allow_empty: bool = False) -> str:
    if not isinstance(value, str) or (not allow_empty and not value):
        _fail(f"{label}: expected {'non-empty ' if not allow_empty else ''}string")
    return value


def _integer(value: Any, label: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        _fail(f"{label}: expected integer >= {minimum}")
    return value


def _boolean_true(value: Any, label: str) -> None:
    if value is not True:
        _fail(f"{label}: expected true")


def _sha(value: Any, label: str) -> str:
    result = _string(value, label).lower()
    if len(result) != 64 or any(character not in "0123456789abcdef" for character in result):
        _fail(f"{label}: expected lowercase SHA-256")
    return result


def _hash_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def _get(mapping: dict[str, Any], key: str, label: str) -> Any:
    if key not in mapping:
        _fail(f"{label}: missing {key}")
    return mapping[key]


def _first(mapping: dict[str, Any], names: Iterable[str], label: str) -> Any:
    for name in names:
        if name in mapping:
            return mapping[name]
    _fail(f"{label}: missing one of {', '.join(names)}")


def _summary_value(summary: dict[str, Any], side: str, name: str, label: str) -> Any:
    """Read the canonical flat field, with nested source/output aliases.

    The final Rust schema uses flat `source_*`/`output_*` fields. The
    nested aliases let the oracle diagnose a draft schema without silently
    treating an absent source or output identity as valid.
    """

    flat_names: list[str]
    if side == "source":
        flat_names = [f"source_{name}"]
    else:
        flat_names = [
            f"output_{name}",
            f"expected_output_{name}",
        ]
        if name == "archive_sha256":
            flat_names.append("expected_output_sha256")
        elif name == "archive_bytes":
            flat_names.append("expected_output_bytes")
        elif name == "content_xml_sha256":
            flat_names.append("expected_output_content_xml_sha256")
        elif name == "content_xml_bytes":
            flat_names.append("expected_output_content_xml_bytes")
    for candidate in flat_names:
        if candidate in summary:
            return summary[candidate]
    nested = summary.get(side)
    if isinstance(nested, dict):
        for candidate in (name, f"{side}_{name}"):
            if candidate in nested:
                return nested[candidate]
    _fail(f"{label}: missing {side} {name}")


def _optional_summary_value(summary: dict[str, Any], side: str, name: str) -> Any:
    names = [f"{side}_{name}"] if side == "source" else [f"output_{name}", f"expected_output_{name}"]
    if side != "source":
        if name == "archive_sha256":
            names.append("expected_output_sha256")
        elif name == "archive_bytes":
            names.append("expected_output_bytes")
        elif name == "content_xml_sha256":
            names.append("expected_output_content_xml_sha256")
        elif name == "content_xml_bytes":
            names.append("expected_output_content_xml_bytes")
    for candidate in names:
        if candidate in summary:
            return summary[candidate]
    nested = summary.get(side)
    if isinstance(nested, dict):
        for candidate in (name, f"{side}_{name}"):
            if candidate in nested:
                return nested[candidate]
    return None


def _gate(summary: dict[str, Any], names: Iterable[str], label: str) -> None:
    for name in names:
        if name in summary:
            _boolean_true(summary[name], f"{label}.{name}")
            return
    _fail(f"{label}: missing one of {', '.join(names)}")


def _summary_field(summary: dict[str, Any], names: Iterable[str], label: str) -> Any:
    for name in names:
        if name in summary:
            return summary[name]
    _fail(f"{label}: missing one of {', '.join(names)}")


def _variant_text(index: int) -> str:
    variants = (
        "plain slide",
        "Unicode café Δ 中",
        'entities <&> "quoted"',
        "mixed façade Ω <&> value",
    )
    return variants[index % 4]


def _title(index: int) -> str:
    return f"litchi-perf-odp-buffered-title-{index:05d} {_variant_text(index)}"


def _body(index: int) -> str:
    return f"litchi-perf-odp-buffered-body-{index:05d} {_variant_text(index)}"


def _semantic_digest(slide_count: int) -> str:
    digest = hashlib.sha256()
    digest.update(SEMANTIC_DOMAIN)
    digest.update(slide_count.to_bytes(8, "little"))
    for index in range(slide_count):
        title = _title(index).encode("utf-8")
        body = _body(index).encode("utf-8")
        digest.update(len(title).to_bytes(8, "little"))
        digest.update(title)
        digest.update(len(body).to_bytes(8, "little"))
        digest.update(body)
    return digest.hexdigest()


def _order_digest(slide_count: int) -> str:
    digest = hashlib.sha256()
    digest.update(ORDER_DOMAIN)
    digest.update(slide_count.to_bytes(8, "little"))
    for index in range(slide_count):
        title = _title(index).encode("utf-8")
        body = _body(index).encode("utf-8")
        digest.update(index.to_bytes(8, "little"))
        digest.update(len(title).to_bytes(8, "little"))
        digest.update(title)
        digest.update(len(body).to_bytes(8, "little"))
        digest.update(body)
    return digest.hexdigest()


def _projection_bytes(slide_count: int) -> int:
    total = 0
    for index in range(slide_count):
        total += len(_title(index).encode()) + 1 + len(_body(index).encode())
        if index + 1 < slide_count:
            total += 2
    return total


def _text_projection(slide_count: int) -> bytes:
    return b"\n\n".join(
        (_title(index) + "\n" + _body(index)).encode("utf-8")
        for index in range(slide_count)
    )


def _opaque_payload() -> bytes:
    return bytes(
        (
            ((index & 0xFF) * 37 + ((index // 256) & 0xFF) + 0x5B)
            & 0xFF
        )
        for index in range(OPAQUE_BYTES)
    )


def _validate_fixture(summary: dict[str, Any], corpus: dict[str, Any], shape: str) -> dict[str, Any]:
    source_count = SHAPES[shape]
    output_count = source_count + 1
    source_semantic = _semantic_digest(source_count)
    output_semantic = _semantic_digest(output_count)
    source_order = _order_digest(source_count)
    output_order = _order_digest(output_count)
    if source_semantic != EXPECTED_SOURCE_SEMANTIC[shape]:
        _fail(f"internal fixture source semantic oracle is inconsistent for {shape}")
    if _string(_get(summary, "role", "odp_append"), "odp_append.role") != "existing_append":
        _fail("odp_append.role must be existing_append")
    if _string(_get(summary, "corpus_generator", "odp_append"), "odp_append.corpus_generator") != GENERATOR:
        _fail("odp_append.corpus_generator differs from the fixture")
    implementation = _string(_get(summary, "implementation", "odp_append"), "odp_append.implementation").lower()
    for token in ("from_bytes", "transaction", "add", "commit"):
        if token not in implementation:
            _fail(f"odp_append.implementation omits {token}")

    generator = _string(_get(corpus, "generator", "corpus"), "corpus.generator")
    if generator != GENERATOR:
        _fail(f"corpus.generator: expected {GENERATOR!r}, got {generator!r}")
    package_format = _string(_get(corpus, "package_format", "corpus"), "corpus.package_format")
    if package_format != PACKAGE_FORMAT:
        _fail(f"corpus.package_format: expected {PACKAGE_FORMAT!r}")
    if _string(_get(corpus, "shape", "corpus"), "corpus.shape") != shape:
        _fail("corpus.shape does not match --shape")
    if _string(_get(summary, "shape", "odp_append"), "odp_append.shape") != shape:
        _fail("odp_append.shape does not match --shape")
    if _string(_get(corpus, "name", "corpus"), "corpus.name") != f"odp-existing-append-lifecycle-{shape}":
        _fail("corpus.name does not match the independent fixture")
    if _integer(_get(corpus, "entry_count", "corpus"), "corpus.entry_count") != source_count:
        _fail("corpus.entry_count does not match source fixture")
    if _integer(_get(corpus, "archive_member_count", "corpus"), "corpus.archive_member_count") != 6:
        _fail("corpus.archive_member_count must be six")
    if _string(_get(corpus, "target_entry", "corpus"), "corpus.target_entry") != "content.xml":
        _fail("corpus.target_entry must be content.xml")
    compression = _string(_get(corpus, "compression", "corpus"), "corpus.compression").lower()
    if "mimetype=stored" not in compression or "xml=deflate" not in compression:
        _fail("corpus.compression omits the required mimetype/XML framing")
    if "opaque=deflate" not in compression:
        _fail("corpus.compression must record opaque=deflate")
    payload_kind = _string(_get(corpus, "payload_kind", "corpus"), "corpus.payload_kind").lower()
    for token in ("deterministic", "plain", "titled", "opaque"):
        if token not in payload_kind:
            _fail(f"corpus.payload_kind omits {token}")
    if "vendor" in payload_kind or "marker" in payload_kind:
        _fail("corpus.payload_kind contains the retired vendor-marker fixture")
    corpus_archive_sha = _sha(_get(corpus, "archive_sha256", "corpus"), "corpus.archive_sha256")
    source_archive_sha = _sha(
        _summary_value(summary, "source", "archive_sha256", "odp_append"),
        "odp_append.source_archive_sha256",
    )
    if corpus_archive_sha != source_archive_sha:
        _fail("corpus archive SHA does not bind to source summary")
    source_archive_bytes = _integer(
        _summary_value(summary, "source", "archive_bytes", "odp_append"),
        "odp_append.source_archive_bytes",
    )
    if _integer(_get(corpus, "archive_bytes", "corpus"), "corpus.archive_bytes") != source_archive_bytes:
        _fail("corpus archive byte count does not bind to source summary")
    source_content_sha = _sha(
        _summary_value(summary, "source", "content_xml_sha256", "odp_append"),
        "odp_append.source_content_xml_sha256",
    )
    source_content_bytes = _integer(
        _summary_value(summary, "source", "content_xml_bytes", "odp_append"),
        "odp_append.source_content_xml_bytes",
    )
    if _sha(_get(corpus, "target_payload_sha256", "corpus"), "corpus.target_payload_sha256") != source_content_sha:
        _fail("corpus target SHA does not bind to source content.xml")
    if _integer(_get(corpus, "target_payload_bytes", "corpus"), "corpus.target_payload_bytes") != source_content_bytes:
        _fail("corpus target byte count does not bind to source content.xml")
    expected_entry_bytes = len(_title(0).encode()) + 1 + len(_body(0).encode())
    if _integer(_get(corpus, "entry_bytes", "corpus"), "corpus.entry_bytes") != expected_entry_bytes:
        _fail("corpus.entry_bytes does not match the independent first-slide fixture")
    logical_bytes = _integer(
        _get(corpus, "uncompressed_payload_bytes", "corpus"),
        "corpus.uncompressed_payload_bytes",
    )
    if logical_bytes != _projection_bytes(source_count):
        _fail("corpus.uncompressed_payload_bytes must be the canonical source text projection")

    source_semantic_reported = _sha(
        _summary_value(summary, "source", "semantic_sha256", "odp_append"),
        "odp_append.source_semantic_sha256",
    )
    output_semantic_reported = _sha(
        _summary_value(summary, "output", "semantic_sha256", "odp_append"),
        "odp_append.output_semantic_sha256",
    )
    if source_semantic_reported != source_semantic or output_semantic_reported != output_semantic:
        _fail("source/output semantic digest does not match the independent fixture")
    source_order_reported = _sha(
        _summary_value(summary, "source", "order_sha256", "odp_append"),
        "odp_append.source_order_sha256",
    )
    output_order_reported = _sha(
        _summary_value(summary, "output", "order_sha256", "odp_append"),
        "odp_append.output_order_sha256",
    )
    if source_order_reported != source_order or output_order_reported != output_order:
        _fail("source/output order digest does not match the independent fixture")
    source_projection_sha = _sha(
        _summary_value(summary, "source", "text_projection_sha256", "odp_append"),
        "odp_append.source_text_projection_sha256",
    )
    output_projection_sha = _sha(
        _summary_value(summary, "output", "text_projection_sha256", "odp_append"),
        "odp_append.output_text_projection_sha256",
    )
    if source_projection_sha != _hash_bytes(_text_projection(source_count)):
        _fail("source text projection SHA differs from fixture")
    if output_projection_sha != _hash_bytes(_text_projection(output_count)):
        _fail("output text projection SHA differs from fixture")
    source_projection = _integer(
        _summary_value(summary, "source", "text_projection_bytes", "odp_append"),
        "odp_append.source_text_projection_bytes",
    )
    output_projection = _integer(
        _summary_value(summary, "output", "text_projection_bytes", "odp_append"),
        "odp_append.output_text_projection_bytes",
    )
    if source_projection != _projection_bytes(source_count):
        _fail("source text projection byte count differs from fixture")
    if output_projection != _projection_bytes(output_count):
        _fail("output text projection byte count differs from fixture")

    opaque = _opaque_payload()
    opaque_sha = _hash_bytes(opaque)
    if _string(_get(summary, "opaque_member_path", "odp_append"), "odp_append.opaque_member_path") != OPAQUE_PATH:
        _fail("opaque member path differs from the independent fixture")
    opaque_bytes = _integer(
        _summary_field(summary, ("opaque_bytes", "opaque_member_decoded_bytes"), "odp_append.opaque_bytes"),
        "odp_append.opaque_bytes",
    )
    if opaque_bytes != len(opaque):
        _fail("opaque member length differs from the independent fixture")
    opaque_sha_reported = _sha(
        _summary_field(summary, ("opaque_sha256", "opaque_member_decoded_sha256"), "odp_append.opaque_sha256"),
        "odp_append.opaque_sha256",
    )
    if opaque_sha_reported != opaque_sha:
        _fail("opaque member decoded SHA differs from the independent formula")
    opaque_compressed_bytes = _integer(
        _get(summary, "opaque_member_compressed_bytes", "odp_append"),
        "odp_append.opaque_member_compressed_bytes",
    )
    if opaque_compressed_bytes <= 0:
        _fail("opaque compressed byte count must be positive")
    _sha(_get(summary, "opaque_member_compressed_sha256", "odp_append"), "odp_append.opaque_member_compressed_sha256")

    appended_title_sha = _hash_bytes(_title(source_count).encode())
    appended_body_sha = _hash_bytes(_body(source_count).encode())
    if _sha(_get(summary, "appended_title_sha256", "odp_append"), "odp_append.appended_title_sha256") != appended_title_sha:
        _fail("appended title SHA does not match fixture index N")
    if _sha(_get(summary, "appended_body_sha256", "odp_append"), "odp_append.appended_body_sha256") != appended_body_sha:
        _fail("appended body SHA does not match fixture index N")
    if _integer(_get(summary, "source_slide_count", "odp_append"), "odp_append.source_slide_count") != source_count:
        _fail("source slide count differs from shape")
    if _integer(_get(summary, "output_slide_count", "odp_append"), "odp_append.output_slide_count") != output_count:
        _fail("output slide count is not source plus one")
    if _integer(_summary_field(summary, ("append_count", "appended_slide_count"), "odp_append.append_count"), "odp_append.append_count") != 1:
        _fail("append count is not exactly one")
    if _integer(_get(summary, "manifest_entry_count", "odp_append"), "odp_append.manifest_entry_count") != 5:
        _fail("manifest entry count must be five")
    if _integer(_get(summary, "source_member_count", "odp_append"), "odp_append.source_member_count") != 6:
        _fail("source member count must be six")
    if _integer(_get(summary, "output_member_count", "odp_append"), "odp_append.output_member_count") != 6:
        _fail("output member count must be six")
    return {
        "source_count": source_count,
        "output_count": output_count,
        "source_semantic": source_semantic,
        "output_semantic": output_semantic,
        "source_order": source_order,
        "output_order": output_order,
        "source_archive_sha": source_archive_sha,
        "source_archive_bytes": source_archive_bytes,
        "source_content_sha": source_content_sha,
        "source_content_bytes": source_content_bytes,
        "opaque_sha": opaque_sha,
        "opaque_compressed_bytes": opaque_compressed_bytes,
        "projection_bytes": _projection_bytes(output_count),
    }


def _member_map(value: Any, label: str) -> dict[str, dict[str, Any]]:
    members = _array(value, label)
    result: dict[str, dict[str, Any]] = {}
    for index, raw in enumerate(members):
        member = _object(raw, f"{label}[{index}]")
        path = _string(_get(member, "path", f"{label}[{index}]"), f"{label}[{index}].path")
        if path in result:
            _fail(f"{label}: duplicate member {path}")
        _string(_get(member, "media_type", f"{label}[{index}]"), f"{label}[{index}].media_type")
        method = _string(_get(member, "compression_method", f"{label}[{index}]"), f"{label}[{index}].compression_method")
        if method.lower() not in {"store", "stored", "deflate", "deflated"}:
            _fail(f"{label}[{index}].compression_method: unsupported {method!r}")
        if not isinstance(_get(member, "data_descriptor", f"{label}[{index}]"), bool):
            _fail(f"{label}[{index}].data_descriptor: expected boolean")
        _integer(_get(member, "crc32", f"{label}[{index}]"), f"{label}[{index}].crc32")
        _integer(_get(member, "decoded_bytes", f"{label}[{index}]"), f"{label}[{index}].decoded_bytes")
        _sha(_get(member, "decoded_sha256", f"{label}[{index}]"), f"{label}[{index}].decoded_sha256")
        _integer(_get(member, "compressed_bytes", f"{label}[{index}]"), f"{label}[{index}].compressed_bytes")
        _sha(_get(member, "compressed_sha256", f"{label}[{index}]"), f"{label}[{index}].compressed_sha256")
        result[path] = member
    if set(result) != EXPECTED_MEMBER_NAMES:
        _fail(f"{label}: member set differs from the six-member fixture")
    return result


def _validate_members(summary: dict[str, Any], fixture: dict[str, Any]) -> None:
    source = _member_map(_get(summary, "source_members", "odp_append"), "odp_append.source_members")
    output = _member_map(_get(summary, "output_members", "odp_append"), "odp_append.output_members")
    if _integer(_get(summary, "source_member_count", "odp_append"), "odp_append.source_member_count") != len(source):
        _fail("source_member_count does not bind to source_members")
    if _integer(_get(summary, "output_member_count", "odp_append"), "odp_append.output_member_count") != len(output):
        _fail("output_member_count does not bind to output_members")
    opaque_source = source[OPAQUE_PATH]
    opaque_output = output[OPAQUE_PATH]
    if opaque_source["media_type"] != OPAQUE_MEDIA_TYPE or opaque_output["media_type"] != OPAQUE_MEDIA_TYPE:
        _fail("opaque member media type differs from the fixture")
    if opaque_source["compression_method"].lower() != OPAQUE_COMPRESSION.lower() or opaque_output["compression_method"].lower() != OPAQUE_COMPRESSION.lower():
        _fail("opaque member is not deflated")
    if opaque_source["decoded_bytes"] != OPAQUE_BYTES or opaque_source["decoded_sha256"] != fixture["opaque_sha"]:
        _fail("source opaque member identity differs from the independent formula")
    if opaque_output["decoded_bytes"] != OPAQUE_BYTES or opaque_output["decoded_sha256"] != fixture["opaque_sha"]:
        _fail("output opaque member identity differs from the independent formula")
    for path in ("mimetype", "styles.xml", "meta.xml", "META-INF/manifest.xml", OPAQUE_PATH):
        if source[path] != output[path]:
            _fail(f"untouched member {path} changed between source and output")
    if source["mimetype"]["compression_method"].lower() != MIMETYPE_COMPRESSION.lower():
        _fail("mimetype is not stored")
    for path in ("content.xml", "styles.xml", "meta.xml"):
        if source[path]["compression_method"].lower() != XML_COMPRESSION.lower():
            _fail(f"{path} is not deflated")
    if source["content.xml"]["decoded_sha256"] != fixture["source_content_sha"]:
        _fail("source content.xml member does not bind to source summary")
    output_content_sha = _sha(
        _summary_value(summary, "output", "content_xml_sha256", "odp_append"),
        "odp_append.output_content_xml_sha256",
    )
    output_content_bytes = _integer(
        _summary_value(summary, "output", "content_xml_bytes", "odp_append"),
        "odp_append.output_content_xml_bytes",
    )
    if output["content.xml"]["decoded_sha256"] != output_content_sha or output["content.xml"]["decoded_bytes"] != output_content_bytes:
        _fail("output content.xml member does not bind to output summary")
    if source["styles.xml"]["decoded_sha256"] != PINNED_STYLES_SHA256 or source["meta.xml"]["decoded_sha256"] != PINNED_META_SHA256:
        _fail("source styles.xml/meta.xml do not match pinned defaults")
    if output["styles.xml"]["decoded_sha256"] != PINNED_STYLES_SHA256 or output["meta.xml"]["decoded_sha256"] != PINNED_META_SHA256:
        _fail("output styles.xml/meta.xml do not match pinned defaults")
    _gate(summary, ("untouched_members_verified", "untouched_member_decoded_identity_verified"), "odp_append")
    _gate(summary, ("opaque_member_compressed_identity_verified",), "odp_append")


def _validate_catalog(report: dict[str, Any], report_path: Path, corpus: dict[str, Any]) -> None:
    reference = _object(_get(report, "corpus_catalog", "report"), "report.corpus_catalog")
    if _integer(_get(reference, "manifest_version", "report.corpus_catalog"), "report.corpus_catalog.manifest_version") != 2:
        _fail("report.corpus_catalog.manifest_version must be 2")
    if _string(_get(reference, "catalog_id", "report.corpus_catalog"), "report.corpus_catalog.catalog_id") != "litchi-perf-corpus-v2":
        _fail("report.corpus_catalog.catalog_id is unexpected")
    reference_catalog_sha = _sha(_get(reference, "catalog_sha256", "report.corpus_catalog"), "report.corpus_catalog.catalog_sha256")
    reference_content_sha = _sha(_get(reference, "content_set_sha256", "report.corpus_catalog"), "report.corpus_catalog.content_set_sha256")
    candidates = [
        report_path.with_name(report_path.stem + "-catalog.json"),
        report_path.with_name("catalog.json"),
    ]
    sidecar = next((candidate for candidate in candidates if candidate.is_file()), None)
    if sidecar is None:
        _fail("corpus catalog sidecar is missing beside report")
    try:
        catalog = json.loads(sidecar.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        _fail(f"catalog sidecar is invalid JSON: {error}")
    catalog = _object(catalog, "catalog")
    if _integer(_get(catalog, "manifest_version", "catalog"), "catalog.manifest_version") != 2:
        _fail("catalog.manifest_version must be 2")
    if _string(_get(catalog, "manifest_kind", "catalog"), "catalog.manifest_kind") != "corpus-catalog":
        _fail("catalog.manifest_kind is unexpected")
    canonicalization = _object(_get(catalog, "canonicalization", "catalog"), "catalog.canonicalization")
    if canonicalization != {"algorithm": "sorted-json-utf8-compact-v1", "hash": "sha256"}:
        _fail("catalog canonicalization is unsupported")
    if _string(_get(catalog, "catalog_id", "catalog"), "catalog.catalog_id") != reference["catalog_id"]:
        _fail("catalog ID does not bind to report reference")
    corpora = _array(_get(catalog, "corpora", "catalog"), "catalog.corpora")
    bindings = _array(_get(catalog, "case_bindings", "catalog"), "catalog.case_bindings")
    if len(corpora) == 0 or len(bindings) == 0:
        _fail("catalog has no corpora or bindings")
    ids = [
        _string(_get(_object(item, f"catalog.corpora[{index}]"), "id", f"catalog.corpora[{index}]"), f"catalog.corpora[{index}].id")
        for index, item in enumerate(corpora)
    ]
    if ids != sorted(ids) or len(set(ids)) != len(ids):
        _fail("catalog corpus IDs are not unique and sorted")
    relevant = []
    for index, raw in enumerate(corpora):
        item = _object(raw, f"catalog.corpora[{index}]")
        legacy = _object(_get(item, "legacy_v1", f"catalog.corpora[{index}]"), f"catalog.corpora[{index}].legacy_v1")
        if legacy.get("archive_sha256") == corpus.get("archive_sha256"):
            relevant.append(item)
    if len(relevant) != 1:
        _fail("catalog does not contain exactly one source corpus binding")
    item = relevant[0]
    bytes_summary = _object(_get(item, "bytes", "catalog corpus"), "catalog corpus.bytes")
    if _sha(_get(bytes_summary, "archive_sha256", "catalog corpus.bytes"), "catalog corpus.bytes.archive_sha256") != _sha(corpus["archive_sha256"], "corpus.archive_sha256"):
        _fail("catalog corpus archive SHA does not match report corpus")
    if _integer(_get(bytes_summary, "archive_bytes", "catalog corpus.bytes"), "catalog corpus.bytes.archive_bytes") != _integer(corpus["archive_bytes"], "corpus.archive_bytes"):
        _fail("catalog corpus archive byte count does not match report corpus")
    if _integer(_get(bytes_summary, "logical_payload_bytes", "catalog corpus.bytes"), "catalog corpus.bytes.logical_payload_bytes") != _integer(corpus["uncompressed_payload_bytes"], "corpus.uncompressed_payload_bytes"):
        _fail("catalog corpus logical payload bytes do not match the canonical projection")
    legacy = _object(_get(item, "legacy_v1", "catalog corpus"), "catalog corpus.legacy_v1")
    for key in ("name", "generator", "package_format", "shape", "payload_kind", "compression", "entry_count", "archive_member_count", "entry_bytes", "uncompressed_payload_bytes", "archive_bytes", "target_entry", "target_payload_bytes", "target_payload_sha256"):
        if legacy.get(key) != corpus.get(key):
            _fail(f"catalog legacy field {key} does not match report corpus")
    member_set = _object(_get(item, "members", "catalog corpus"), "catalog corpus.members")
    member_status = _string(_get(member_set, "status", "catalog corpus.members"), "catalog corpus.members.status")
    member_items = _array(_get(member_set, "items", "catalog corpus.members"), "catalog corpus.members.items")
    if member_status == "unavailable":
        if member_items:
            _fail("unavailable catalog member set must have no items")
    elif member_status == "complete":
        member_names = {_string(_get(_object(member, f"catalog member {index}"), "name", f"catalog member {index}"), f"catalog member {index}.name") for index, member in enumerate(member_items)}
        if member_names != EXPECTED_MEMBER_NAMES:
            _fail("catalog member set differs from the six-member fixture")
        opaque_catalog = next(_object(member, "catalog member") for member in member_items if member.get("name") == OPAQUE_PATH)
        if _sha(_get(opaque_catalog, "sha256", "catalog opaque member"), "catalog opaque member.sha256") != _hash_bytes(_opaque_payload()):
            _fail("catalog opaque SHA differs from independent formula")
    else:
        _fail("catalog corpus member set has an unsupported status")
    matching_bindings = []
    for index, raw in enumerate(bindings):
        binding = _object(raw, f"catalog.case_bindings[{index}]")
        if binding.get("case") == SELECTOR:
            matching_bindings.append(binding)
    if len(matching_bindings) != 1:
        _fail("catalog must have exactly one append selector binding")
    binding = matching_bindings[0]
    if binding.get("corpus_id") != item.get("id") or binding.get("role") != "timed":
        _fail("append selector catalog binding does not identify the source corpus")
    if binding.get("legacy_archive_sha256") != corpus.get("archive_sha256"):
        _fail("append selector legacy archive binding differs")
    catalog_for_hash = json.loads(json.dumps(catalog))
    catalog_for_hash.pop("catalog_sha256", None)
    computed_catalog_sha = _hash_bytes(_canonical_json(catalog_for_hash))
    if computed_catalog_sha != reference_catalog_sha or computed_catalog_sha != catalog.get("catalog_sha256"):
        _fail("catalog SHA does not match sorted canonical JSON")
    content_value = {
        "corpora": [
            {
                "id": entry.get("id"),
                "archive_sha256": _object(entry.get("bytes"), "catalog corpus.bytes").get("archive_sha256"),
                "members": [
                    {
                        "ordinal": member.get("ordinal"),
                        "name": member.get("name"),
                        "sha256": member.get("sha256"),
                    }
                    for member in _object(entry.get("members"), "catalog corpus.members").get("items", [])
                ],
            }
            for entry in corpora
        ],
        "case_bindings": [
            {"case": entry.get("case"), "corpus_id": entry.get("corpus_id"), "role": entry.get("role")}
            for entry in bindings
        ],
    }
    if _hash_bytes(_canonical_json(content_value)) != reference_content_sha or _hash_bytes(_canonical_json(content_value)) != catalog.get("content_set_sha256"):
        _fail("catalog content-set SHA does not match sorted canonical JSON")


def _canonical_json(value: Any) -> bytes:
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode("utf-8")


def _metric_vector(value: Any, label: str, samples: int) -> tuple[str, list[int] | None]:
    vector = _object(value, label)
    status = _string(_get(vector, "status", label), f"{label}.status")
    values = vector.get("values")
    if status == "measured":
        values = _array(values, f"{label}.values")
        if len(values) != samples:
            _fail(f"{label}.values has {len(values)} values, expected {samples}")
        checked = [_integer(item, f"{label}.values[{index}]") for index, item in enumerate(values)]
        return status, checked
    if values is not None:
        _fail(f"{label}.values must be absent for {status}")
    return status, None


def _walk_metric_vectors(value: Any, label: str, samples: int) -> list[tuple[str, list[int] | None, str]]:
    found: list[tuple[str, list[int] | None, str]] = []
    if isinstance(value, dict):
        if "status" in value and "scope" in value and ("values" in value or set(value).issubset({"status", "scope"})):
            status, values = _metric_vector(value, label, samples)
            found.append((label, values, status))
        for key, child in value.items():
            if key not in {"values", "status", "scope"}:
                found.extend(_walk_metric_vectors(child, f"{label}.{key}", samples))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            found.extend(_walk_metric_vectors(child, f"{label}[{index}]", samples))
    return found


def _validate_metrics(result: dict[str, Any], summary: dict[str, Any], fixture: dict[str, Any], mode: str, samples: int) -> None:
    elapsed = _object(_get(result, "elapsed_ns", "result"), "result.elapsed_ns")
    if _string(_get(elapsed, "unit", "result.elapsed_ns"), "result.elapsed_ns.unit") != "ns":
        _fail("elapsed_ns unit is not ns")
    elapsed_values = [_integer(item, f"result.elapsed_ns.samples[{index}]", minimum=1) for index, item in enumerate(_array(_get(elapsed, "samples", "result.elapsed_ns"), "result.elapsed_ns.samples"))]
    if len(elapsed_values) != samples or elapsed_values != sorted(elapsed_values):
        _fail("elapsed_ns.samples must be a sorted vector of the requested length")
    sample_order = [_integer(item, f"result.elapsed_ns.sample_order[{index}]") for index, item in enumerate(_array(_get(elapsed, "sample_order", "result.elapsed_ns"), "result.elapsed_ns.sample_order"))]
    if len(sample_order) != samples or sorted(sample_order) != list(range(samples)):
        _fail("elapsed_ns.sample_order must be a permutation of retained sample indices")
    for key in ("min", "p50", "p95", "p99", "max"):
        _integer(_get(elapsed, key, "result.elapsed_ns"), f"result.elapsed_ns.{key}", minimum=1)
    for key in ("mean", "standard_deviation"):
        value = _get(elapsed, key, "result.elapsed_ns")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)) or float(value) < 0:
            _fail(f"result.elapsed_ns.{key}: expected finite non-negative number")
    interval = _object(_get(elapsed, "confidence_interval_95", "result.elapsed_ns"), "result.elapsed_ns.confidence_interval_95")
    if not isinstance(_get(interval, "method", "confidence interval"), str):
        _fail("confidence interval method is missing")
    for key in ("lower", "upper"):
        value = _get(interval, key, "confidence interval")
        if isinstance(value, bool) or not isinstance(value, (int, float)) or not math.isfinite(float(value)) or float(value) < 0:
            _fail(f"confidence interval {key} is invalid")

    lifecycle = _array(_get(summary, "lifecycle_ns", "odp_append"), "odp_append.lifecycle_ns")
    lifecycle_values = [_integer(item, f"odp_append.lifecycle_ns[{index}]", minimum=1) for index, item in enumerate(lifecycle)]
    if len(lifecycle_values) != samples or sorted(lifecycle_values) != elapsed_values:
        _fail("odp_append.lifecycle_ns does not bind to elapsed_ns.samples")
    if [lifecycle_values[index] for index in sample_order] != elapsed_values:
        _fail("lifecycle_ns does not bind to original sample indices")
    output_vector = [_sha(item, f"odp_append.output_sha256[{index}]") for index, item in enumerate(_array(_get(summary, "output_sha256", "odp_append"), "odp_append.output_sha256"))]
    expected_output_sha = _sha(_summary_value(summary, "output", "archive_sha256", "odp_append"), "odp_append.expected_output_sha256")
    if len(output_vector) != samples or any(item != expected_output_sha for item in output_vector):
        _fail("odp_append.output_sha256 is not a stable retained output vector")
    if _sha(_get(result, "output_sha256", "result"), "result.output_sha256") != expected_output_sha:
        _fail("result.output_sha256 does not bind to expected output")

    metrics = _object(_get(result, "operation_metrics", "result"), "result.operation_metrics")
    if _integer(_get(metrics, "sample_count", "operation_metrics"), "operation_metrics.sample_count") != samples:
        _fail("operation_metrics.sample_count does not match samples")
    if _string(_get(metrics, "alignment", "operation_metrics"), "operation_metrics.alignment") != "elapsed_ns.samples_by_elapsed_then_sample_index":
        _fail("operation_metrics alignment is not the declared elapsed/sample-index order")
    metrics_indices = [_integer(item, f"operation_metrics.sample_indices[{index}]") for index, item in enumerate(_array(_get(metrics, "sample_indices", "operation_metrics"), "operation_metrics.sample_indices"))]
    if metrics_indices != sample_order:
        _fail("operation_metrics.sample_indices do not match elapsed_ns.sample_order")
    latency_claim = _string(_get(metrics, "latency_claim", "operation_metrics"), "operation_metrics.latency_claim").lower()
    if "physical" in latency_claim or "cold" in latency_claim or "source_backed" in latency_claim:
        _fail("operation_metrics latency claim overstates the owned lifecycle scope")
    vectors = _walk_metric_vectors(metrics, "operation_metrics", samples)
    if not vectors:
        _fail("operation_metrics contains no metric vectors")
    source_metrics = _object(_get(metrics, "source", "operation_metrics"), "operation_metrics.source")
    if _string(_get(source_metrics, "status", "operation_metrics.source"), "operation_metrics.source.status") != "not_applicable":
        _fail("owned append source metrics must be not_applicable")
    for section_name in ("publication", "materialization", "cfb_phases"):
        section = _object(_get(metrics, section_name, "operation_metrics"), f"operation_metrics.{section_name}")
        if _string(_get(section, "status", f"operation_metrics.{section_name}"), f"operation_metrics.{section_name}.status") != "not_applicable":
            _fail(f"operation_metrics.{section_name} must be not_applicable")
    sink_metrics = _object(_get(metrics, "sink", "operation_metrics"), "operation_metrics.sink")
    if _string(_get(sink_metrics, "write_status", "operation_metrics.sink"), "operation_metrics.sink.write_status") != "measured":
        _fail("operation_metrics sink write status must be measured")
    accepted_status, accepted = _metric_vector(_get(sink_metrics, "accepted_bytes", "operation_metrics.sink"), "operation_metrics.sink.accepted_bytes", samples)
    expected_output_bytes = _integer(_summary_value(summary, "output", "archive_bytes", "odp_append"), "odp_append.expected_output_bytes")
    if accepted_status != "measured" or accepted is None or any(value != expected_output_bytes for value in accepted):
        _fail("operation_metrics sink accepted bytes do not equal committed output bytes")
    _, write_calls = _metric_vector(_get(sink_metrics, "write_calls", "operation_metrics.sink"), "operation_metrics.sink.write_calls", samples)
    _, largest = _metric_vector(_get(sink_metrics, "largest_write", "operation_metrics.sink"), "operation_metrics.sink.largest_write", samples)
    sink = _object(_get(result, "sink", "result"), "result.sink")
    sink_accepted = _integer(_get(sink, "accepted_bytes", "result.sink"), "result.sink.accepted_bytes")
    sink_calls = _integer(_get(sink, "write_calls", "result.sink"), "result.sink.write_calls")
    sink_largest = _integer(_get(sink, "largest_write", "result.sink"), "result.sink.largest_write")
    if sink_accepted != expected_output_bytes or sink_calls <= 0 or sink_largest <= 0:
        _fail("result.sink does not describe the committed output write")
    if write_calls is None or any(value != sink_calls for value in write_calls) or largest is None or any(value != sink_largest for value in largest):
        _fail("operation_metrics sink counters do not bind to result.sink")
    buckets = _object(_get(sink, "write_size_buckets", "result.sink"), "result.sink.write_size_buckets")
    bucket_total = 0
    for key in ("bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536"):
        bucket_total += _integer(_get(buckets, key, "result.sink.write_size_buckets"), f"result.sink.write_size_buckets.{key}")
    if bucket_total != sink_calls:
        _fail("result.sink write buckets do not sum to write_calls")
    if sink.get("retained_output_bytes") != 0:
        _fail("append sink must report zero retained output bytes")
    if sink.get("retained_authoring_window_bytes") is not None:
        _fail("append sink must not report an authoring window")

    allocation = metrics.get("allocation")
    if mode == "allocator":
        if not isinstance(allocation, dict) or allocation.get("status") != "measured":
            _fail("allocator report must contain measured allocation metrics")
        allocator_vectors = {}
        for key in ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes"):
            status, values = _metric_vector(_get(allocation, key, "operation_metrics.allocation"), f"operation_metrics.allocation.{key}", samples)
            if status != "measured" or values is None:
                _fail(f"allocator vector {key} is not measured")
            allocator_vectors[key] = values
        for index in range(samples):
            values = {key: vector[index] for key, vector in allocator_vectors.items()}
            if values['failed_allocation_calls'] != 0:
                _fail('allocator failure calls must be zero')
            if values['live_bytes_before'] + values['allocated_bytes'] - values['deallocated_bytes'] != values['live_bytes_after']:
                _fail('allocator live-byte balance differs')
            if values['region_peak_live_bytes'] < max(values['live_bytes_before'], values['live_bytes_after']):
                _fail('allocator region peak is below an endpoint')
        if allocation.get("scope") != "operation_global_system_allocator":
            _fail("allocator scope is not operation_global_system_allocator")
    elif allocation is not None:
        if not isinstance(allocation, dict) or allocation.get("status") == "measured":
            _fail("normal report must not publish measured allocator metrics")


def _validate_summary_gates(summary: dict[str, Any]) -> None:
    for names in (
        ("source_manifest_bindings_verified", "manifest_bindings_verified"),
        ("output_manifest_bindings_verified", "manifest_bindings_verified"),
        ("source_semantic_reopen_verified", "semantic_reopen_verified"),
        ("output_semantic_reopen_verified", "semantic_reopen_verified"),
        ("untouched_members_verified", "untouched_member_decoded_identity_verified"),
        ("opaque_member_compressed_identity_verified",),
        ("append_exactly_one_verified",),
        ("source_unchanged_verified",),
        ("patch_replay_verified",),
        ("inverse_patch_verified",),
        ("stale_source_refusal_verified",),
        ("exact_noop_verified",),
        ("runtime_output_digest_verified",),
        ("runtime_sink_length_verified",),
    ):
        _gate(summary, names, "odp_append")
    text_contract = _string(_get(summary, "text_contract", "odp_append"), "odp_append.text_contract")
    if text_contract != TEXT_CONTRACT:
        _fail("odp_append.text_contract differs from the existing buffered fixture")
    timing = _string(_get(summary, "timing_scope", "odp_append"), "odp_append.timing_scope")
    lower_timing = timing.lower()
    for token in ("from_bytes", "transaction", "add", "commit", "write"):
        if token not in lower_timing:
            _fail(f"odp_append.timing_scope omits {token}")
    for token in ("outside", "oracle", "drop"):
        if token not in lower_timing:
            _fail(f"odp_append.timing_scope omits outside-{token} scope")
    if re.search(r"(?:outside|excluding|exclude)[^.;]{0,100}from_bytes", lower_timing):
        _fail("odp_append.timing_scope incorrectly excludes Snapshot::from_bytes")
    if re.search(r"from_bytes[^.;]{0,100}(?:outside|excluded|exclude)", lower_timing):
        _fail("odp_append.timing_scope incorrectly excludes Snapshot::from_bytes")
    includes_from_bytes = (
        re.search(r"(?:clock|timer|timed region|measured region)[^.;]{0,100}(?:includes|contains)[^.;]*from_bytes", lower_timing)
        or re.search(r"from_bytes[^.;]{0,100}(?:inside|included|timed)", lower_timing)
    )
    if not includes_from_bytes:
        _fail("odp_append.timing_scope does not place Snapshot::from_bytes inside the timer")
    if "source_backed" in lower_timing or "bounded-memory" in lower_timing or "zero-copy" in lower_timing:
        _fail("odp_append.timing_scope makes an unsupported provider claim")
    claim = _string(_get(summary, "performance_claim", "odp_append"), "odp_append.performance_claim").lower()
    for forbidden in ("source-backed", "bounded-memory", "zero-copy", "physical-i/o", "cold-file", "cancellation", "before/after"):
        if forbidden not in claim:
            continue
        explicitly_disclaimed = (
            "no source-backed, bounded-memory, or semantic-streaming-save claim" in claim
            and forbidden in {"source-backed", "bounded-memory"}
        ) or bool(re.search(rf"\b(?:no|not|without)\s+{re.escape(forbidden)}\b", claim))
        if not explicitly_disclaimed:
            _fail(f"odp_append.performance_claim contains forbidden claim {forbidden}")


def _validate_identity(report: dict[str, Any], mode: str) -> None:
    tool = _object(_get(report, 'tool', 'report'), 'tool')
    expected_binary = 'litchi-perf-baseline' + ('-alloc' if mode == 'allocator' else '')
    if tool.get('name') != 'litchi-perf-baseline' or tool.get('binary') != expected_binary or tool.get('profile') != 'release':
        _fail('tool binary/profile identity differs')
    expected_instrumentation = 'system_allocator_operation_scoped' if mode == 'allocator' else 'none'
    if tool.get('instrumentation') != expected_instrumentation:
        _fail('tool instrumentation differs from requested mode')
    if tool.get('allocator_counter_revision') != ('serialized_region_peak_v3' if mode == 'allocator' else None):
        _fail('allocator counter revision differs')
    identity = _object(_get(report, 'binary_identity', 'report'), 'binary_identity')
    _sha(identity.get('binary_sha256'), 'binary_identity.binary_sha256')
    _integer(identity.get('binary_bytes'), 'binary_identity.binary_bytes', minimum=1)
    if identity.get('profile') != 'release' or identity.get('executable') is not True:
        _fail('binary identity must describe a release executable')
    environment = _object(_get(report, 'environment', 'report'), 'environment')
    if environment.get('rustc_version') != 'rustc 1.98.1 (48a229cea 2026-09-01)' or environment.get('rustflags') != '-Cforce-frame-pointers=yes':
        _fail('toolchain or frame-pointer build flags differ')
    if environment.get('cpu_affinity') != '2' or environment.get('logical_cpus_available') != 1:
        _fail('CPU 2 single-worker affinity differs')
    expected_allocator = 'CountingSystemAllocator(std::alloc::System)' if mode == 'allocator' else 'Rust system allocator'
    if environment.get('allocator') != expected_allocator:
        _fail('allocator identity differs from mode')


def validate_report(path: str | Path, mode: str, shape: str, *, samples: int = 30, warmups: int = 3) -> dict[str, Any]:
    """Validate one serialized 0439 report and return its decoded object."""

    if mode not in {"normal", "allocator"}:
        _fail("mode must be normal or allocator")
    if shape not in SHAPES:
        _fail("shape must be tiny, medium, or large")
    if isinstance(samples, bool) or not isinstance(samples, int) or samples <= 0:
        _fail("samples must be a positive integer")
    if isinstance(warmups, bool) or not isinstance(warmups, int) or warmups < 0:
        _fail("warmups must be a non-negative integer")
    report_path = Path(path)
    try:
        report = json.loads(report_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        _fail(f"report is invalid JSON: {error}")
    report = _object(report, "report")
    _validate_identity(report, mode)
    if _integer(_get(report, "schema_version", "report"), "report.schema_version") != 1:
        _fail("report.schema_version must be 1")
    configuration = _object(_get(report, "configuration", "report"), "report.configuration")
    if _integer(_get(configuration, "samples_per_case", "configuration"), "configuration.samples_per_case") != samples:
        _fail("configuration.samples_per_case differs from --samples")
    if _integer(_get(configuration, "warmup_iterations_per_case", "configuration"), "configuration.warmup_iterations_per_case") != warmups:
        _fail("configuration.warmup_iterations_per_case differs from --warmups")
    if _get(configuration, "cases", "configuration") != [SELECTOR]:
        _fail("configuration.cases must contain only the append selector")
    if _get(configuration, "semantic_shapes", "configuration") != [shape]:
        _fail("configuration.semantic_shapes must contain only --shape")
    if _get(configuration, "execution_workers", "configuration") != [1]:
        _fail("configuration.execution_workers must be [1]")
    results = _array(_get(report, "results", "report"), "report.results")
    if len(results) != 1:
        _fail("one report invocation must contain exactly one result")
    result = _object(results[0], "report.results[0]")
    if _string(_get(result, "case", "result"), "result.case") != SELECTOR:
        _fail("result.case is not the append selector")
    corpus = _object(_get(result, "corpus", "result"), "result.corpus")
    source = _object(_get(result, "source", "result"), "result.source")
    summary = _object(_get(source, "odp_append", "result.source"), "result.source.odp_append")
    for key in ("read_calls", "read_bytes", "ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads"):
        if key in source and source[key] not in ([], None):
            _fail(f"result.source.{key} must be empty for an owned append lifecycle")
    fixture = _validate_fixture(summary, corpus, shape)
    _validate_members(summary, fixture)
    _validate_summary_gates(summary)
    _validate_catalog(report, report_path, corpus)
    _validate_metrics(result, summary, fixture, mode, samples)
    return report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--mode", choices=("normal", "allocator"), required=True)
    parser.add_argument("--shape", choices=tuple(SHAPES), required=True)
    parser.add_argument("--samples", type=int, default=30)
    parser.add_argument("--warmups", type=int, default=3)
    # Evidence capture passes this role field; 0439 has one current-revision
    # role, so it is checked for compatibility and never used to relax gates.
    parser.add_argument("--role", default="after")
    args = parser.parse_args(argv)
    if args.role != "after":
        print("INVALID: only the current after role is supported", file=sys.stderr)
        return 2
    try:
        validate_report(args.report, args.mode, args.shape, samples=args.samples, warmups=args.warmups)
    except (OSError, TypeError, ValueError, ValidationError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 2
    print("VALID")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
