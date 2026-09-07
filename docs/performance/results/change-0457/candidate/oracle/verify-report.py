#!/usr/bin/env python3
"""Independent report/fixture oracle for the 0457 ODP source-tail candidate.

The candidate uses the normal release build with a null rustflags field; this oracle
makes no frame-pointer build claim.

This verifier validates the serialized report, the frozen control corpus, and
the retained per-shape source/output receipt produced by the native ZIP/XML
oracle. The Rust runner must provide the source/output member, manifest, and
semantic readback gates; this file independently regenerates the fixture text,
semantic digests, and opaque payload hash and binds the report's candidate
archive/content/member identities to the separately verified output fixture.
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


SELECTOR = "odp_source_tail_append_lifecycle"
GENERATOR = "litchi-odp-existing-append-lifecycle-v1"
OUTPUT_BINDING_SCHEMA = "litchi-0457-odp-source-tail-candidate-output-binding-v1"
NATIVE_OUTPUT_SCHEMA = "litchi-0457-native-odp-append-oracle-v1"
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


def _load_corpus_binding(shape: str) -> dict[str, Any]:
    path = Path(__file__).resolve().parent.parent / "corpus-bindings.json"
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        _fail(f"candidate corpus binding is invalid: {error}")
    binding = _object(value, "candidate corpus binding")
    if binding.get("schema") != "litchi-0457-odp-source-tail-candidate-corpus-binding-v1":
        _fail("candidate corpus binding schema differs")
    shapes = _object(_get(binding, "shapes", "candidate corpus binding"), "candidate corpus binding.shapes")
    return _object(_get(shapes, shape, "candidate corpus binding.shapes"), f"candidate corpus binding.shapes.{shape}")


def _load_output_binding(shape: str) -> tuple[dict[str, Any], dict[str, Any]]:
    path = Path(__file__).resolve().parent.parent / "output-binding.json"
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        _fail(f"candidate output binding is invalid: {error}")
    binding = _object(value, "candidate output binding")
    if binding.get("schema") != OUTPUT_BINDING_SCHEMA:
        _fail("candidate output binding schema differs")
    # The outer protocol records the final immutable binding hash after the
    # native fixture preflight.  Keeping this check here makes a copied bundle
    # replayable without trusting a path-only receipt.
    protocol_path = path.parent / "protocol.json"
    try:
        protocol = json.loads(protocol_path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        _fail(f"candidate protocol is invalid while checking output binding: {error}")
    protocol_output = _object(_get(protocol, "output_binding", "candidate protocol"), "candidate protocol.output_binding")
    if protocol_output.get("path") != path.name:
        _fail("candidate protocol output binding path differs")
    expected_sha = protocol_output.get("sha256")
    if expected_sha not in ("pending-output-preflight", "bound-at-capture"):
        if _hash_bytes(path.read_bytes()) != _sha(expected_sha, "candidate protocol.output_binding.sha256"):
            _fail("candidate output binding differs from the protocol hash")
    binder = _object(_get(binding, "binder", "candidate output binding"), "candidate output binding.binder")
    if binder.get("path") != "bind-output.py" or protocol_output.get("binder_path") != binder.get("path"):
        _fail("candidate output binding helper path differs")
    binder_path = path.parent / binder["path"]
    if not binder_path.is_file() or binder_path.is_symlink() or _hash_bytes(binder_path.read_bytes()) != _sha(binder.get("sha256"), "candidate output binding.binder.sha256"):
        _fail("candidate output binding helper differs from its receipt")
    expected_binder_sha = protocol_output.get("binder_sha256")
    if expected_binder_sha not in ("pending-output-preflight", "bound-at-capture") and _sha(expected_binder_sha, "candidate protocol.output_binding.binder_sha256") != _hash_bytes(binder_path.read_bytes()):
        _fail("candidate output binding helper differs from protocol")
    native = _object(_get(binding, "native_oracle", "candidate output binding"), "candidate output binding.native_oracle")
    native_relative = _string(_get(native, "path", "candidate output binding.native_oracle"), "candidate output binding.native_oracle.path")
    if native_relative != "../native/verify-output.py":
        _fail("candidate output binding must use the frozen native output oracle path")
    native_path = (path.parent / native_relative).resolve()
    if not native_path.is_file() or native_path.is_symlink():
        _fail(f"native output oracle is missing: {native_path}")
    if _hash_bytes(native_path.read_bytes()) != _sha(_get(native, "sha256", "candidate output binding.native_oracle"), "candidate output binding.native_oracle.sha256"):
        _fail("native output oracle differs from the output binding")
    shapes = _object(_get(binding, "shapes", "candidate output binding"), "candidate output binding.shapes")
    return binding, _object(_get(shapes, shape, "candidate output binding.shapes"), f"candidate output binding.shapes.{shape}")


def _validate_native_output_receipt(
    summary: dict[str, Any],
    control_source: dict[str, Any],
    generated: dict[str, Any],
    shape: str,
) -> None:
    native = _object(_get(generated, "native_result", "candidate output binding shape"), "candidate output binding.native_result")
    receipt_sha = _sha(_get(generated, "native_result_sha256", "candidate output binding shape"), "candidate output binding.native_result_sha256")
    if _hash_bytes(_canonical_json(native)) != receipt_sha:
        _fail("candidate output native receipt hash is not canonical")
    if native.get("schema") != NATIVE_OUTPUT_SCHEMA or native.get("status") != "validated":
        _fail("candidate output binding does not contain a successful native ZIP/XML receipt")
    _boolean_true(native.get("source_hash_verified"), "candidate output binding.native_result.source_hash_verified")
    source = _object(_get(generated, "source", "candidate output binding shape"), "candidate output binding.source")
    output = _object(_get(generated, "output", "candidate output binding shape"), "candidate output binding.output")
    fixtures = _object(_get(generated, "fixtures", "candidate output binding shape"), "candidate output binding.fixtures")
    fixture_root = Path(__file__).resolve().parent.parent
    for side in ("source", "output"):
        fixture = _object(_get(fixtures, side, "candidate output binding.fixtures"), f"candidate output binding.fixtures.{side}")
        fixture_path = (fixture_root / _string(_get(fixture, "path", f"candidate output binding.fixtures.{side}"), f"candidate output binding.fixtures.{side}.path")).resolve()
        try:
            fixture_path.relative_to(fixture_root)
        except ValueError:
            _fail(f"candidate output binding.{side} fixture escapes the candidate bundle")
        if not fixture_path.is_file() or fixture_path.is_symlink() or fixture_path.stat().st_size != _integer(_get(fixture, "bytes", f"candidate output binding.fixtures.{side}"), f"candidate output binding.fixtures.{side}.bytes") or _hash_bytes(fixture_path.read_bytes()) != _sha(_get(fixture, "sha256", f"candidate output binding.fixtures.{side}"), f"candidate output binding.fixtures.{side}.sha256"):
            _fail(f"retained candidate {side} fixture differs from its output binding")
    native_source = _object(_get(native, "source", "candidate output binding.native_result"), "candidate output binding.native_result.source")
    native_output = _object(_get(native, "output", "candidate output binding.native_result"), "candidate output binding.native_result.output")
    if native_source.get("path") != fixtures["source"].get("path") or native_output.get("path") != fixtures["output"].get("path"):
        _fail("native output receipt paths do not bind retained fixtures")
    for side, expected in (("source", source), ("output", output)):
        native_side = native_source if side == "source" else native_output
        if native_side.get("sha256") != expected.get("archive_sha256") or native_side.get("bytes") != expected.get("archive_bytes"):
            _fail(f"native output receipt {side} archive identity differs from its binding")
    if source.get("archive_sha256") != control_source.get("archive_sha256") or source.get("archive_bytes") != control_source.get("archive_bytes"):
        _fail("native fixture source archive differs from the frozen source corpus")
    if source.get("content_xml_sha256") != control_source.get("content_xml_sha256") or source.get("content_xml_bytes") != control_source.get("content_xml_bytes"):
        _fail("native fixture source content.xml differs from the frozen source corpus")
    content = _object(_get(native, "content", "candidate output binding.native_result"), "candidate output binding.native_result.content")
    if content.get("source_content_xml_sha256") != source.get("content_xml_sha256") or content.get("source_content_xml_bytes") != source.get("content_xml_bytes"):
        _fail("native output receipt source content identity differs")
    if content.get("output_content_xml_sha256") != output.get("content_xml_sha256") or content.get("output_content_xml_bytes") != output.get("content_xml_bytes"):
        _fail("native output receipt candidate content identity differs")
    output_count = SHAPES[shape] + 1
    request = _object(_get(generated, "request", "candidate output binding shape"), "candidate output binding.request")
    expected_request = {"title": _title(SHAPES[shape]), "body": _body(SHAPES[shape]), "name": f"page{output_count}"}
    if request != expected_request or _object(_get(native, "request", "candidate output binding.native_result"), "candidate output binding.native_result.request") != expected_request:
        _fail("native output request differs from the deterministic append request")
    if content.get("source_page_count") != SHAPES[shape] or content.get("output_page_count") != output_count:
        _fail("native output receipt page counts differ from the shape")
    if content.get("source_last_page_end_byte") != _integer(_get(summary, "proof_insert_at", "odp_source_tail_append"), "odp_source_tail_append.proof_insert_at"):
        _fail("native insertion offset does not bind to the source proof")
    if content.get("output_generated_page_start_byte") != content.get("source_last_page_end_byte"):
        _fail("native generated page is not inserted at the source last-page boundary")
    for key in ("exact_source_prefix_verified", "exact_source_suffix_verified", "one_extra_direct_draw_page_verified", "source_pages_contiguous"):
        _boolean_true(content.get(key), f"candidate output binding.native_result.content.{key}")
    zip_facts = _object(_get(native, "zip", "candidate output binding.native_result"), "candidate output binding.native_result.zip")
    for key in ("central_member_order_verified", "local_member_order_verified", "archive_comment_verified", "target_local_central_self_consistent"):
        _boolean_true(zip_facts.get(key), f"candidate output binding.native_result.zip.{key}")
    if zip_facts.get("untouched_member_count") != 5 or zip_facts.get("target_member") != "content.xml":
        _fail("native output receipt does not prove the five untouched members")


def _validate_external_binding(summary: dict[str, Any], corpus: dict[str, Any], binding: dict[str, Any], generated: dict[str, Any], shape: str) -> None:
    expected_corpus = _object(_get(binding, "corpus", "candidate corpus binding shape"), "candidate binding.corpus")
    if corpus != expected_corpus:
        _fail("report corpus differs from the independently bound control corpus")
    source = _object(_get(binding, "source", "candidate corpus binding shape"), "candidate binding.source")
    output = _object(_get(binding, "expected_output", "candidate corpus binding shape"), "candidate binding.expected_output")
    source_fields = {
        "archive_sha256": "source_archive_sha256",
        "archive_bytes": "source_archive_bytes",
        "content_xml_sha256": "source_content_xml_sha256",
        "content_xml_bytes": "source_content_xml_bytes",
        "semantic_sha256": "source_semantic_sha256",
        "order_sha256": "source_order_sha256",
        "text_projection_sha256": "source_text_projection_sha256",
        "text_projection_bytes": "source_text_projection_bytes",
        "slide_count": "source_slide_count",
    }
    output_fields = {
        "semantic_sha256": "output_semantic_sha256",
        "order_sha256": "output_order_sha256",
        "text_projection_sha256": "output_text_projection_sha256",
        "text_projection_bytes": "output_text_projection_bytes",
        "slide_count": "output_slide_count",
    }
    for expected_key, report_key in source_fields.items():
        actual = _summary_value(summary, "source", expected_key, "odp_source_tail_append")
        if actual != source.get(expected_key):
            _fail(f"source {expected_key} differs from the frozen control binding")
    for expected_key, report_key in output_fields.items():
        actual = _summary_value(summary, "output", expected_key, "odp_source_tail_append")
        if actual != output.get(expected_key):
            _fail(f"output {expected_key} differs from the frozen semantic control binding")
    if summary.get("source_members") != source.get("members"):
        _fail("source ZIP member identities differ from the frozen control binding")
    generated_source = _object(_get(generated, "source", "candidate output binding shape"), "candidate output binding.source")
    generated_output = _object(_get(generated, "output", "candidate output binding shape"), "candidate output binding.output")
    if generated_source.get("members") != source.get("members"):
        _fail("native fixture source member identities differ from the frozen control binding")
    if summary.get("output_members") != generated_output.get("members"):
        _fail("output ZIP member identities differ from the independently verified candidate fixture")
    for key in ("archive_sha256", "archive_bytes", "content_xml_sha256", "content_xml_bytes"):
        actual = _summary_value(summary, "output", key, "odp_source_tail_append")
        if actual != generated_output.get(key):
            _fail(f"candidate output {key} differs from the independently verified native fixture")
    _validate_native_output_receipt(summary, source, generated, shape)
    opaque = _object(_get(binding, "opaque", "candidate corpus binding shape"), "candidate binding.opaque")
    for key, report_key in (("path", "opaque_member_path"), ("bytes", "opaque_bytes"), ("sha256", "opaque_sha256"), ("compressed_bytes", "opaque_member_compressed_bytes"), ("compressed_sha256", "opaque_member_compressed_sha256")):
        if summary.get(report_key) != opaque.get(key):
            _fail(f"opaque {key} differs from the frozen control binding")
    append = _object(_get(binding, "append", "candidate corpus binding shape"), "candidate binding.append")
    if summary.get("append_count") != append.get("count") or summary.get("appended_title_sha256") != append.get("title_sha256") or summary.get("appended_body_sha256") != append.get("body_sha256"):
        _fail("append request identity differs from the independently bound fixture")
    source_count = SHAPES[shape]
    if summary.get("appended_title") != _title(source_count) or summary.get("appended_body") != _body(source_count):
        _fail("append title/body differs from the independently bound request")
    source_archive = _sha(source.get("archive_sha256"), "candidate binding.source.archive_sha256")
    output_archive = _sha(_summary_value(summary, "output", "archive_sha256", "odp_source_tail_append"), "odp_source_tail_append.output_archive_sha256")
    if output_archive == source_archive:
        _fail("source-tail output archive is byte-identical to its source")
    _integer(_summary_value(summary, "output", "archive_bytes", "odp_source_tail_append"), "odp_source_tail_append.output_archive_bytes")
    if shape not in SHAPES:
        _fail("candidate output archive binding is malformed")


def _validate_fixture(summary: dict[str, Any], corpus: dict[str, Any], shape: str, binding: dict[str, Any], generated: dict[str, Any]) -> dict[str, Any]:
    source_count = SHAPES[shape]
    output_count = source_count + 1
    source_semantic = _semantic_digest(source_count)
    output_semantic = _semantic_digest(output_count)
    source_order = _order_digest(source_count)
    output_order = _order_digest(output_count)
    if source_semantic != EXPECTED_SOURCE_SEMANTIC[shape]:
        _fail(f"internal fixture source semantic oracle is inconsistent for {shape}")
    if _string(_get(summary, "role", "odp_source_tail_append"), "odp_source_tail_append.role") != "source_tail_append":
        _fail("odp_source_tail_append.role must be source_tail_append")
    if _string(_get(summary, "corpus_generator", "odp_source_tail_append"), "odp_source_tail_append.corpus_generator") != GENERATOR:
        _fail("odp_source_tail_append.corpus_generator differs from the fixture")
    implementation = _string(_get(summary, "implementation", "odp_source_tail_append"), "odp_source_tail_append.implementation")
    expected_implementation = "SourceBackedPackage::from_read_at + SourceBackedTailAppendEdit::plan + SourceBackedTailAppendPublicationPlan::write_to"
    if implementation != expected_implementation:
        _fail("odp_source_tail_append.implementation differs from the specialized source-tail contract")
    if any(token in implementation.lower() for token in ("snapshot", "transaction", "commit", "patch")):
        _fail("source-tail implementation claims the ordinary owned lifecycle")

    generator = _string(_get(corpus, "generator", "corpus"), "corpus.generator")
    if generator != GENERATOR:
        _fail(f"corpus.generator: expected {GENERATOR!r}, got {generator!r}")
    package_format = _string(_get(corpus, "package_format", "corpus"), "corpus.package_format")
    if package_format != PACKAGE_FORMAT:
        _fail(f"corpus.package_format: expected {PACKAGE_FORMAT!r}")
    if _string(_get(corpus, "shape", "corpus"), "corpus.shape") != shape:
        _fail("corpus.shape does not match --shape")
    if _string(_get(summary, "shape", "odp_source_tail_append"), "odp_source_tail_append.shape") != shape:
        _fail("odp_source_tail_append.shape does not match --shape")
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
        _summary_value(summary, "source", "archive_sha256", "odp_source_tail_append"),
        "odp_source_tail_append.source_archive_sha256",
    )
    if corpus_archive_sha != source_archive_sha:
        _fail("corpus archive SHA does not bind to source summary")
    source_archive_bytes = _integer(
        _summary_value(summary, "source", "archive_bytes", "odp_source_tail_append"),
        "odp_source_tail_append.source_archive_bytes",
    )
    if _integer(_get(corpus, "archive_bytes", "corpus"), "corpus.archive_bytes") != source_archive_bytes:
        _fail("corpus archive byte count does not bind to source summary")
    source_content_sha = _sha(
        _summary_value(summary, "source", "content_xml_sha256", "odp_source_tail_append"),
        "odp_source_tail_append.source_content_xml_sha256",
    )
    source_content_bytes = _integer(
        _summary_value(summary, "source", "content_xml_bytes", "odp_source_tail_append"),
        "odp_source_tail_append.source_content_xml_bytes",
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
        _summary_value(summary, "source", "semantic_sha256", "odp_source_tail_append"),
        "odp_source_tail_append.source_semantic_sha256",
    )
    output_semantic_reported = _sha(
        _summary_value(summary, "output", "semantic_sha256", "odp_source_tail_append"),
        "odp_source_tail_append.output_semantic_sha256",
    )
    if source_semantic_reported != source_semantic or output_semantic_reported != output_semantic:
        _fail("source/output semantic digest does not match the independent fixture")
    source_order_reported = _sha(
        _summary_value(summary, "source", "order_sha256", "odp_source_tail_append"),
        "odp_source_tail_append.source_order_sha256",
    )
    output_order_reported = _sha(
        _summary_value(summary, "output", "order_sha256", "odp_source_tail_append"),
        "odp_source_tail_append.output_order_sha256",
    )
    if source_order_reported != source_order or output_order_reported != output_order:
        _fail("source/output order digest does not match the independent fixture")
    source_projection_sha = _sha(
        _summary_value(summary, "source", "text_projection_sha256", "odp_source_tail_append"),
        "odp_source_tail_append.source_text_projection_sha256",
    )
    output_projection_sha = _sha(
        _summary_value(summary, "output", "text_projection_sha256", "odp_source_tail_append"),
        "odp_source_tail_append.output_text_projection_sha256",
    )
    if source_projection_sha != _hash_bytes(_text_projection(source_count)):
        _fail("source text projection SHA differs from fixture")
    if output_projection_sha != _hash_bytes(_text_projection(output_count)):
        _fail("output text projection SHA differs from fixture")
    source_projection = _integer(
        _summary_value(summary, "source", "text_projection_bytes", "odp_source_tail_append"),
        "odp_source_tail_append.source_text_projection_bytes",
    )
    output_projection = _integer(
        _summary_value(summary, "output", "text_projection_bytes", "odp_source_tail_append"),
        "odp_source_tail_append.output_text_projection_bytes",
    )
    if source_projection != _projection_bytes(source_count):
        _fail("source text projection byte count differs from fixture")
    if output_projection != _projection_bytes(output_count):
        _fail("output text projection byte count differs from fixture")

    opaque = _opaque_payload()
    opaque_sha = _hash_bytes(opaque)
    if _string(_get(summary, "opaque_member_path", "odp_source_tail_append"), "odp_source_tail_append.opaque_member_path") != OPAQUE_PATH:
        _fail("opaque member path differs from the independent fixture")
    opaque_bytes = _integer(
        _summary_field(summary, ("opaque_bytes", "opaque_member_decoded_bytes"), "odp_source_tail_append.opaque_bytes"),
        "odp_source_tail_append.opaque_bytes",
    )
    if opaque_bytes != len(opaque):
        _fail("opaque member length differs from the independent fixture")
    opaque_sha_reported = _sha(
        _summary_field(summary, ("opaque_sha256", "opaque_member_decoded_sha256"), "odp_source_tail_append.opaque_sha256"),
        "odp_source_tail_append.opaque_sha256",
    )
    if opaque_sha_reported != opaque_sha:
        _fail("opaque member decoded SHA differs from the independent formula")
    opaque_compressed_bytes = _integer(
        _get(summary, "opaque_member_compressed_bytes", "odp_source_tail_append"),
        "odp_source_tail_append.opaque_member_compressed_bytes",
    )
    if opaque_compressed_bytes <= 0:
        _fail("opaque compressed byte count must be positive")
    _sha(_get(summary, "opaque_member_compressed_sha256", "odp_source_tail_append"), "odp_source_tail_append.opaque_member_compressed_sha256")

    appended_title_sha = _hash_bytes(_title(source_count).encode())
    appended_body_sha = _hash_bytes(_body(source_count).encode())
    if _sha(_get(summary, "appended_title_sha256", "odp_source_tail_append"), "odp_source_tail_append.appended_title_sha256") != appended_title_sha:
        _fail("appended title SHA does not match fixture index N")
    if _sha(_get(summary, "appended_body_sha256", "odp_source_tail_append"), "odp_source_tail_append.appended_body_sha256") != appended_body_sha:
        _fail("appended body SHA does not match fixture index N")
    if _integer(_get(summary, "source_slide_count", "odp_source_tail_append"), "odp_source_tail_append.source_slide_count") != source_count:
        _fail("source slide count differs from shape")
    if _integer(_get(summary, "output_slide_count", "odp_source_tail_append"), "odp_source_tail_append.output_slide_count") != output_count:
        _fail("output slide count is not source plus one")
    if _integer(_summary_field(summary, ("append_count", "appended_slide_count"), "odp_source_tail_append.append_count"), "odp_source_tail_append.append_count") != 1:
        _fail("append count is not exactly one")
    if _integer(_get(summary, "manifest_entry_count", "odp_source_tail_append"), "odp_source_tail_append.manifest_entry_count") != 5:
        _fail("manifest entry count must be five")
    if _integer(_get(summary, "source_member_count", "odp_source_tail_append"), "odp_source_tail_append.source_member_count") != 6:
        _fail("source member count must be six")
    if _integer(_get(summary, "output_member_count", "odp_source_tail_append"), "odp_source_tail_append.output_member_count") != 6:
        _fail("output member count must be six")
    _validate_external_binding(summary, corpus, binding, generated, shape)
    _validate_proof(summary, shape)
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
    source = _member_map(_get(summary, "source_members", "odp_source_tail_append"), "odp_source_tail_append.source_members")
    output = _member_map(_get(summary, "output_members", "odp_source_tail_append"), "odp_source_tail_append.output_members")
    if _integer(_get(summary, "source_member_count", "odp_source_tail_append"), "odp_source_tail_append.source_member_count") != len(source):
        _fail("source_member_count does not bind to source_members")
    if _integer(_get(summary, "output_member_count", "odp_source_tail_append"), "odp_source_tail_append.output_member_count") != len(output):
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
        _summary_value(summary, "output", "content_xml_sha256", "odp_source_tail_append"),
        "odp_source_tail_append.output_content_xml_sha256",
    )
    output_content_bytes = _integer(
        _summary_value(summary, "output", "content_xml_bytes", "odp_source_tail_append"),
        "odp_source_tail_append.output_content_xml_bytes",
    )
    if output["content.xml"]["decoded_sha256"] != output_content_sha or output["content.xml"]["decoded_bytes"] != output_content_bytes:
        _fail("output content.xml member does not bind to output summary")
    if source["styles.xml"]["decoded_sha256"] != PINNED_STYLES_SHA256 or source["meta.xml"]["decoded_sha256"] != PINNED_META_SHA256:
        _fail("source styles.xml/meta.xml do not match pinned defaults")
    if output["styles.xml"]["decoded_sha256"] != PINNED_STYLES_SHA256 or output["meta.xml"]["decoded_sha256"] != PINNED_META_SHA256:
        _fail("output styles.xml/meta.xml do not match pinned defaults")
    _gate(summary, ("source_member_raw_preservation_verified",), "odp_source_tail_append")


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
        _fail("catalog must have exactly one source-tail selector binding")
    binding = matching_bindings[0]
    if binding.get("corpus_id") != item.get("id") or binding.get("role") != "timed":
        _fail("source-tail selector catalog binding does not identify the source corpus")
    if binding.get("legacy_archive_sha256") != corpus.get("archive_sha256"):
        _fail("source-tail selector legacy archive binding differs")
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


def _vector(values: Any, label: str, samples: int, *, minimum: int = 0) -> list[int]:
    vector = _array(values, label)
    if len(vector) != samples:
        _fail(f"{label} has {len(vector)} values, expected {samples}")
    return [_integer(value, f"{label}[{index}]", minimum=minimum) for index, value in enumerate(vector)]


def _validate_metrics(result: dict[str, Any], summary: dict[str, Any], fixture: dict[str, Any], mode: str, samples: int) -> None:
    elapsed = _object(_get(result, "elapsed_ns", "result"), "result.elapsed_ns")
    if _string(_get(elapsed, "unit", "result.elapsed_ns"), "result.elapsed_ns.unit") != "ns":
        _fail("elapsed_ns unit is not ns")
    elapsed_values = _vector(_get(elapsed, "samples", "result.elapsed_ns"), "result.elapsed_ns.samples", samples, minimum=1)
    if elapsed_values != sorted(elapsed_values):
        _fail("elapsed_ns.samples must be sorted")
    sample_order = _vector(_get(elapsed, "sample_order", "result.elapsed_ns"), "result.elapsed_ns.sample_order", samples)
    if sorted(sample_order) != list(range(samples)):
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

    def aligned_vector(name: str, *, minimum: int = 0) -> list[int]:
        # The Rust runner reorders every auxiliary vector with sample_order
        # before serialization, so these vectors are already in elapsed order.
        return _vector(_get(summary, name, "odp_source_tail_append"), f"odp_source_tail_append.{name}", samples, minimum=minimum)

    lifecycle = aligned_vector("lifecycle_ns", minimum=1)
    if lifecycle != elapsed_values:
        _fail("odp_source_tail_append.lifecycle_ns does not bind to elapsed_ns.samples")
    for name in ("open_ns", "append_plan_ns", "publication_ns"):
        aligned_vector(name, minimum=1)
    report_bytes = aligned_vector("publication_report_bytes")
    expected_output_bytes = _integer(_summary_value(summary, "output", "archive_bytes", "odp_source_tail_append"), "odp_source_tail_append.output_archive_bytes")
    if any(value != expected_output_bytes for value in report_bytes):
        _fail("publication_report_bytes does not bind to the candidate output length")
    report_source_ids = aligned_vector("publication_report_source_version_id", minimum=1)
    report_source_revisions = aligned_vector("publication_report_source_version_revision")
    runtime_proof_ids = aligned_vector("runtime_proof_source_version_id", minimum=1)
    runtime_proof_revisions = aligned_vector("runtime_proof_source_version_revision")
    if runtime_proof_ids != report_source_ids:
        _fail("runtime proof source IDs do not bind to the aligned publication report IDs")
    if runtime_proof_revisions != report_source_revisions:
        _fail("runtime proof source revisions do not bind to the aligned publication report revisions")
    # The untimed fixture proof is a separate source object. Retain and
    # validate its scalar identity, but never compare it with a fresh timed
    # provider or with the per-sample publication report vectors.
    _integer(_get(summary, "proof_source_version_id", "odp_source_tail_append"), "odp_source_tail_append.proof_source_version_id", minimum=1)
    _integer(_get(summary, "proof_source_version_revision", "odp_source_tail_append"), "odp_source_tail_append.proof_source_version_revision")
    read_calls = aligned_vector("source_read_calls", minimum=1)
    read_bytes = aligned_vector("source_read_bytes", minimum=1)
    output_vector = [_sha(item, f"odp_source_tail_append.output_sha256[{index}]") for index, item in enumerate(_array(_get(summary, "output_sha256", "odp_source_tail_append"), "odp_source_tail_append.output_sha256"))]
    if len(output_vector) != samples:
        _fail("odp_source_tail_append.output_sha256 has the wrong length")
    expected_output_sha = _sha(_summary_value(summary, "output", "archive_sha256", "odp_source_tail_append"), "odp_source_tail_append.output_archive_sha256")
    if any(item != expected_output_sha for item in output_vector) or _sha(_get(result, "output_sha256", "result"), "result.output_sha256") != expected_output_sha:
        _fail("candidate output digest is not stable and bound to the report")

    source = _object(_get(result, "source", "result"), "result.source")
    for key, values in (("read_calls", read_calls), ("read_bytes", read_bytes)):
        actual = _vector(_get(source, key, "result.source"), f"result.source.{key}", samples, minimum=1)
        if actual != values:
            _fail(f"result.source.{key} does not bind to source summary")
    for key in ("ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads"):
        if source.get(key) not in ([], None):
            _fail(f"result.source.{key} must be empty for this direct positional source")

    metrics = _object(_get(result, "operation_metrics", "result"), "result.operation_metrics")
    if _integer(_get(metrics, "sample_count", "operation_metrics"), "operation_metrics.sample_count") != samples:
        _fail("operation_metrics.sample_count does not match samples")
    if _string(_get(metrics, "alignment", "operation_metrics"), "operation_metrics.alignment") != "elapsed_ns.samples_by_elapsed_then_sample_index":
        _fail("operation_metrics alignment is not the declared elapsed/sample-index order")
    metrics_indices = _vector(_get(metrics, "sample_indices", "operation_metrics"), "operation_metrics.sample_indices", samples)
    if metrics_indices != sample_order:
        _fail("operation_metrics.sample_indices do not match elapsed_ns.sample_order")
    # The shared in-process source/sink envelope labels its measured vectors
    # `comparable_timed_operation`; the source-tail performance_claim above
    # still withholds any ordinary Commit/Patch comparison or speedup claim.
    if _string(_get(metrics, "latency_claim", "operation_metrics"), "operation_metrics.latency_claim") != "comparable_timed_operation":
        _fail("operation_metrics latency claim differs from the shared in-process timed envelope")
    if not _walk_metric_vectors(metrics, "operation_metrics", samples):
        _fail("operation_metrics contains no metric vectors")
    source_metrics = _object(_get(metrics, "source", "operation_metrics"), "operation_metrics.source")
    if _string(_get(source_metrics, "status", "operation_metrics.source"), "operation_metrics.source.status") != "measured":
        _fail("source-tail source metrics must be measured")
    if _get(source_metrics, "counter_scope", "operation_metrics.source") != "in_process_instrumented_source_read_at":
        _fail("source-tail source counter scope is not the positional source observer")
    measured_source = {}
    for key in ("logical_read_calls", "logical_read_returned_bytes", "max_concurrent_reads"):
        status, values = _metric_vector(_get(source_metrics, key, "operation_metrics.source"), f"operation_metrics.source.{key}", samples)
        if status != "measured" or values is None or any(value <= 0 for value in values):
            _fail(f"operation_metrics.source.{key} must be measured and positive")
        measured_source[key] = values
    if measured_source["logical_read_calls"] != read_calls or measured_source["logical_read_returned_bytes"] != read_bytes:
        _fail("operation_metrics source call/returned-byte vectors do not bind to source summary")
    if measured_source["max_concurrent_reads"] != [1] * samples:
        _fail("source-tail max concurrency must be one for its serialized lifecycle")
    for key in ("logical_read_requested_bytes", "logical_read_largest_requested_bytes", "logical_read_largest_returned_bytes", "compressed_bytes", "decompressed_bytes", "recompressed_bytes"):
        status, values = _metric_vector(_get(source_metrics, key, "operation_metrics.source"), f"operation_metrics.source.{key}", samples)
        if status != "unavailable" or values is not None:
            _fail(f"operation_metrics.source.{key} must remain explicitly unavailable")
    pattern = _object(_get(source_metrics, "logical_read_pattern", "operation_metrics.source"), "operation_metrics.source.logical_read_pattern")
    if pattern.get("status") != "unavailable" or pattern.get("values") is not None:
        _fail("source request pattern must remain unavailable")

    for section_name in ("publication", "materialization", "cfb_phases"):
        section = _object(_get(metrics, section_name, "operation_metrics"), f"operation_metrics.{section_name}")
        if _string(_get(section, "status", f"operation_metrics.{section_name}"), f"operation_metrics.{section_name}.status") != "not_applicable":
            _fail(f"operation_metrics.{section_name} must be not_applicable")
    sink_metrics = _object(_get(metrics, "sink", "operation_metrics"), "operation_metrics.sink")
    if _string(_get(sink_metrics, "write_status", "operation_metrics.sink"), "operation_metrics.sink.write_status") != "measured":
        _fail("operation_metrics sink write status must be measured")
    accepted_status, accepted = _metric_vector(_get(sink_metrics, "accepted_bytes", "operation_metrics.sink"), "operation_metrics.sink.accepted_bytes", samples)
    if accepted_status != "measured" or accepted is None or any(value != expected_output_bytes for value in accepted):
        _fail("operation_metrics sink accepted bytes do not equal candidate output bytes")
    _, write_calls = _metric_vector(_get(sink_metrics, "write_calls", "operation_metrics.sink"), "operation_metrics.sink.write_calls", samples)
    _, largest = _metric_vector(_get(sink_metrics, "largest_write", "operation_metrics.sink"), "operation_metrics.sink.largest_write", samples)
    sink = _object(_get(result, "sink", "result"), "result.sink")
    sink_accepted = _integer(_get(sink, "accepted_bytes", "result.sink"), "result.sink.accepted_bytes")
    sink_calls = _integer(_get(sink, "write_calls", "result.sink"), "result.sink.write_calls")
    sink_largest = _integer(_get(sink, "largest_write", "result.sink"), "result.sink.largest_write")
    if sink_accepted != expected_output_bytes or sink_calls <= 0 or sink_largest <= 0:
        _fail("result.sink does not describe the candidate publication")
    if write_calls is None or any(value != sink_calls for value in write_calls) or largest is None or any(value != sink_largest for value in largest):
        _fail("operation_metrics sink counters do not bind to result.sink")
    buckets = _object(_get(sink, "write_size_buckets", "result.sink"), "result.sink.write_size_buckets")
    bucket_names = ("bytes_0", "bytes_1_to_512", "bytes_513_to_4096", "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536")
    bucket_total = 0
    for key in bucket_names:
        bucket_value = _integer(_get(buckets, key, "result.sink.write_size_buckets"), f"result.sink.write_size_buckets.{key}")
        bucket_total += bucket_value
        metric_bucket = _object(_get(_get(sink_metrics, "write_size_buckets", "operation_metrics.sink"), key, "operation_metrics.sink.write_size_buckets"), f"operation_metrics.sink.write_size_buckets.{key}")
        metric_status, metric_values = _metric_vector(metric_bucket, f"operation_metrics.sink.write_size_buckets.{key}", samples)
        if metric_status != "measured" or metric_values is None or metric_values != [bucket_value] * samples:
            _fail(f"operation_metrics sink bucket {key} does not bind to result.sink")
    if bucket_total != sink_calls:
        _fail("result.sink write buckets do not sum to write_calls")
    if sink.get("retained_output_bytes") != 0 or sink.get("retained_authoring_window_bytes") is not None:
        _fail("source-tail sink must retain no output or authoring window")

    allocation = metrics.get("allocation")
    if mode == "allocator":
        if not isinstance(allocation, dict) or allocation.get("status") != "measured" or allocation.get("scope") != "operation_global_system_allocator":
            _fail("allocator report must contain operation-scoped measured allocation metrics")
        vectors = {}
        for key in ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after", "region_peak_live_bytes"):
            status, values = _metric_vector(_get(allocation, key, "operation_metrics.allocation"), f"operation_metrics.allocation.{key}", samples)
            if status != "measured" or values is None:
                _fail(f"allocator vector {key} is not measured")
            vectors[key] = values
        for index in range(samples):
            if vectors["failed_allocation_calls"][index] != 0:
                _fail("allocator failure calls must be zero")
            if vectors["live_bytes_before"][index] + vectors["allocated_bytes"][index] - vectors["deallocated_bytes"][index] != vectors["live_bytes_after"][index]:
                _fail("allocator live-byte balance differs")
            if vectors["region_peak_live_bytes"][index] < max(vectors["live_bytes_before"][index], vectors["live_bytes_after"][index]):
                _fail("allocator region peak is below an endpoint")
    elif allocation is not None and (not isinstance(allocation, dict) or allocation.get("status") == "measured"):
        _fail("normal report must not publish measured allocator metrics")


def _validate_proof(summary: dict[str, Any], shape: str) -> None:
    source_count = SHAPES[shape]
    source_content_bytes = _integer(_summary_value(summary, "source", "content_xml_bytes", "odp_source_tail_append"), "odp_source_tail_append.source_content_xml_bytes")
    if _integer(_get(summary, "proof_slide_count", "odp_source_tail_append"), "odp_source_tail_append.proof_slide_count") != source_count:
        _fail("source proof slide count differs from the deterministic corpus")
    if _integer(_get(summary, "proof_content_xml_bytes", "odp_source_tail_append"), "odp_source_tail_append.proof_content_xml_bytes") != source_content_bytes:
        _fail("source proof content length does not bind to source content.xml")
    insert_at = _integer(_get(summary, "proof_insert_at", "odp_source_tail_append"), "odp_source_tail_append.proof_insert_at")
    if insert_at <= 0 or insert_at >= source_content_bytes:
        _fail("source proof insertion offset is outside the source content.xml")
    if _string(_get(summary, "proof_page_name", "odp_source_tail_append"), "odp_source_tail_append.proof_page_name") != f"page{source_count + 1}":
        _fail("source proof page name is not the deterministic append page")
    _integer(_get(summary, "proof_source_version_id", "odp_source_tail_append"), "odp_source_tail_append.proof_source_version_id", minimum=1)
    if _integer(_get(summary, "proof_source_version_revision", "odp_source_tail_append"), "odp_source_tail_append.proof_source_version_revision") != 0:
        _fail("source proof revision must be the initial source revision")


def _validate_summary_gates(summary: dict[str, Any]) -> None:
    for name in (
        "source_manifest_bindings_verified",
        "output_manifest_bindings_verified",
        "source_member_raw_preservation_verified",
        "source_semantic_reopen_verified",
        "output_semantic_reopen_verified",
        "append_exactly_one_verified",
        "source_unchanged_verified",
        "stale_source_refusal_verified",
        "runtime_output_digest_verified",
        "runtime_sink_length_verified",
    ):
        _gate(summary, (name,), "odp_source_tail_append")
    for retired in ("untouched_members_verified", "patch_replay_verified", "inverse_patch_verified", "exact_noop_verified"):
        if retired in summary:
            _fail(f"odp_source_tail_append must not publish ordinary lifecycle gate {retired}")
    text_contract = _string(_get(summary, "text_contract", "odp_source_tail_append"), "odp_source_tail_append.text_contract")
    if text_contract != TEXT_CONTRACT:
        _fail("odp_source_tail_append.text_contract differs from the deterministic fixture")
    timing = _string(_get(summary, "timing_scope", "odp_source_tail_append"), "odp_source_tail_append.timing_scope").lower()
    for token in (
        "source archive arc",
        "positional provider",
        "append strings/options",
        "hashingdiscardsink",
        "lifecycle_ns includes",
        "sourcebackedpackage opening",
        "bounded source scan/proof",
        "insertion-plan preparation",
        "sequential publication-plan write",
    ):
        if token not in timing:
            _fail(f"odp_source_tail_append.timing_scope omits {token}")
    for token in ("outside", "digest finalization", "oracle", "report assembly", "drops"):
        if token not in timing:
            _fail(f"odp_source_tail_append.timing_scope omits outside-{token} scope")
    if "snapshot::from_bytes" in timing or "transaction" in timing:
        _fail("source-tail timing scope contains the ordinary owned lifecycle")
    claim = _string(_get(summary, "performance_claim", "odp_source_tail_append"), "odp_source_tail_append.performance_claim").lower()
    required_claims = (
        "source-backed odp advanced publication-plan lifecycle evidence only",
        "sourcecontentpublicationreport",
        "different retained-result contract",
        "no ordinary commit/patch optimization",
        "no ordinary commit/patch optimization or general crud claim",
    )
    if not all(token in claim for token in required_claims[:4]):
        _fail("odp_source_tail_append.performance_claim does not state the specialized scope")
    for forbidden in ("before/after", "speedup", "regression", "causal hotspot", "general crud claim"):
        if forbidden in claim and not re.search(rf"\b(?:no|not|without)\b[^.;]{{0,80}}{re.escape(forbidden)}", claim):
            _fail(f"odp_source_tail_append.performance_claim overstates {forbidden}")
    if "ordinary commit/patch optimization" not in claim or "no ordinary commit/patch optimization" not in claim:
        _fail("source-tail claim does not withhold an ordinary Commit/Patch optimization claim")


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
    if environment.get('rustc_version') != 'rustc 1.98.1 (48a229cea 2026-09-01)' or environment.get('rustflags') is not None:
        _fail('toolchain or unexpected build flags differ; R1 makes no frame-pointer build claim')
    if environment.get('cpu_affinity') != '2' or environment.get('logical_cpus_available') != 1:
        _fail('CPU 2 single-worker affinity differs')
    expected_allocator = 'CountingSystemAllocator(std::alloc::System)' if mode == 'allocator' else 'Rust system allocator'
    if environment.get('allocator') != expected_allocator:
        _fail('allocator identity differs from mode')


def validate_report(path: str | Path, mode: str, shape: str, *, samples: int = 30, warmups: int = 3) -> dict[str, Any]:
    """Validate one serialized 0457 source-tail candidate report."""

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
        _fail("configuration.cases must contain only the source-tail selector")
    if _get(configuration, "semantic_shapes", "configuration") != [shape]:
        _fail("configuration.semantic_shapes must contain only --shape")
    if _get(configuration, "execution_workers", "configuration") != [1]:
        _fail("configuration.execution_workers must be [1]")
    results = _array(_get(report, "results", "report"), "report.results")
    if len(results) != 1:
        _fail("one report invocation must contain exactly one result")
    result = _object(results[0], "report.results[0]")
    if _string(_get(result, "case", "result"), "result.case") != SELECTOR:
        _fail("result.case is not the source-tail selector")
    corpus = _object(_get(result, "corpus", "result"), "result.corpus")
    source = _object(_get(result, "source", "result"), "result.source")
    summary = _object(_get(source, "odp_source_tail_append", "result.source"), "result.source.odp_source_tail_append")
    for key in ("ordinary_payload_read_calls", "ordinary_payload_read_bytes", "max_in_flight_reads"):
        if key in source and source[key] not in ([], None):
            _fail(f"result.source.{key} must be empty for the source-tail lifecycle")
    binding = _load_corpus_binding(shape)
    _, generated = _load_output_binding(shape)
    fixture = _validate_fixture(summary, corpus, shape, binding, generated)
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
    # Capture passes this compatibility role; it never relaxes report gates.
    parser.add_argument("--role", default="candidate")
    args = parser.parse_args(argv)
    if args.role != "candidate":
        print("INVALID: only the candidate role is supported", file=sys.stderr)
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
