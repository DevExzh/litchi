"""Independent, streaming Python oracle for the 0484 synthetic DOCX corpus.

This reproduces the specified text and XML grammar without invoking Litchi or
trusting hashes from the measured executable. ZIP compression and raw member
preservation require separate artifact checks. Only scalar results are cached;
the generator holds at most one authored paragraph, not a document payload.
"""

from __future__ import annotations

from functools import lru_cache
import hashlib
from typing import Any
from xml.sax.saxutils import escape


WORD = b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
HEAD = b'<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="' + WORD + b'"><w:body>'
TAIL = b"</w:body></w:document>"
PARA_HEAD = b'<w:p xmlns:w="' + WORD + b'"><w:r><w:t xml:space="preserve">'
PARA_TAIL = b"</w:t></w:r></w:p>"
NEAR_BYTES = 60 * 1024
CHUNKS = {"one": 0, "fixed64": 64, "replay_window": 8192}
TEXT_MODES = ("empty", "short", "near_limit")


def _u64(value: int) -> bytes:
    return value.to_bytes(8, "little")


def _text(index: int, mode: str) -> bytes:
    if mode == "empty":
        return b""
    prefix = f"authored-{index:08d} <&>".encode("ascii")
    if mode == "short":
        return prefix
    filler = b" near<&> plain"
    remaining = NEAR_BYTES - len(prefix)
    return prefix + (filler * ((remaining + len(filler) - 1) // len(filler)))[:remaining]


class _Semantic:
    def __init__(self, count: int):
        self.count = count
        self.index = 0
        self.bytes = 0
        self.order = hashlib.sha256(b"litchi-docx-replayable-tail-order-v1\0" + _u64(count))
        self.text = hashlib.sha256(b"litchi-docx-replayable-tail-text-v1\0" + _u64(count))

    def add(self, text: bytes) -> None:
        framed = _u64(len(text)) + text
        self.order.update(_u64(self.index))
        self.order.update(framed)
        self.text.update(framed)
        self.index += 1
        self.bytes += len(text)

    def record(self) -> dict[str, Any]:
        if self.index != self.count:
            raise ValueError("semantic generator cardinality differs")
        return {
            "paragraph_count": self.count,
            "text_bytes": self.bytes,
            "order_sha256": self.order.hexdigest(),
            "text_sha256": self.text.hexdigest(),
        }


@lru_cache(maxsize=64)
def _expected(source_count: int, authored_count: int, chunk_mode: str, text_mode: str) -> dict[str, Any]:
    if type(source_count) is not int or not 1 <= source_count <= 131_072:
        raise ValueError("source count outside finite corpus domain")
    if type(authored_count) is not int or not 1 <= authored_count <= 16_384:
        raise ValueError("authored count outside finite corpus domain")
    if chunk_mode not in CHUNKS or text_mode not in TEXT_MODES:
        raise ValueError("unknown authored corpus mode")
    # The specified near-limit axis uses 64 paragraphs. Refuse an unplanned
    # billion-byte Python validation workload before generating any payload.
    if text_mode == "near_limit" and authored_count > 64:
        raise ValueError("near-limit count outside selected corpus matrix")
    source_hash = hashlib.sha256(HEAD)
    candidate_hash = hashlib.sha256(HEAD)
    source_length = len(HEAD)
    source_semantic = _Semantic(source_count)
    candidate_semantic = _Semantic(source_count + authored_count)
    for index in range(source_count):
        text = f"source-{index:08d}".encode("ascii")
        paragraph = b"<w:p><w:r><w:t>" + text + b"</w:t></w:r></w:p>"
        source_hash.update(paragraph)
        candidate_hash.update(paragraph)
        source_length += len(paragraph)
        source_semantic.add(text)
        candidate_semantic.add(text)
    insertion = source_length
    source_hash.update(TAIL)
    source_length += len(TAIL)
    event_hash = hashlib.sha256()
    encoded_hash = hashlib.sha256()
    event_count = text_bytes = encoded_bytes = entities = maximum_text = maximum_encoded = 0
    for index in range(authored_count):
        text = _text(index, text_mode)
        maximum_text = max(maximum_text, len(text))
        text_bytes += len(text)
        entities += sum(text.count(value) for value in (b"&", b"<", b">", b'"', b"'"))
        event_hash.update(b"\x00")
        event_count += 2
        if text:
            width = CHUNKS[chunk_mode] or len(text)
            for start in range(0, len(text), width):
                chunk = text[start:start + width]
                event_hash.update(b"\x01" + _u64(len(chunk)) + chunk)
                event_count += 1
        event_hash.update(b"\x02")
        encoded = escape(text.decode("ascii"), {'"': "&quot;", "'": "&apos;"}).encode("ascii")
        paragraph = PARA_HEAD + encoded + PARA_TAIL
        encoded_bytes += len(paragraph)
        maximum_encoded = max(maximum_encoded, len(paragraph))
        encoded_hash.update(paragraph)
        candidate_hash.update(paragraph)
        candidate_semantic.add(text)
    candidate_hash.update(TAIL)
    chunk_limit = max(1, min(CHUNKS[chunk_mode] or maximum_text, maximum_text))
    authored = {
        "authored_count": authored_count,
        "chunk_mode": chunk_mode,
        "text_mode": text_mode,
        "max_chunk_bytes": chunk_limit,
        "replay_window_bytes": max(65_536, 6 * chunk_limit + 1024),
        "max_encoded_paragraph_bytes": maximum_encoded,
        "xml_entity_reference_count": entities,
        "text_bytes": text_bytes,
        "encoded_xml_bytes": encoded_bytes,
        "event_count": event_count,
        "expected_event_sha256": event_hash.hexdigest(),
        "expected_encoded_sha256": encoded_hash.hexdigest(),
    }
    return {
        "source": {
            "main_xml_bytes": source_length,
            "main_xml_sha256": source_hash.hexdigest(),
            "semantic": source_semantic.record(),
        },
        "authored": authored,
        "oracle": {
            "candidate_main_xml_bytes": source_length + encoded_bytes,
            "candidate_main_xml_sha256": candidate_hash.hexdigest(),
            "candidate_semantic": candidate_semantic.record(),
        },
        "proof": {"insertion_offset": insertion, "generated_offset": insertion},
    }


def expected_case(source_count: int, authored_count: int, chunk_mode: str, text_mode: str) -> dict[str, Any]:
    """Return detached scalar expectations; callers cannot mutate the cache."""
    import copy
    if type(source_count) is not int or type(authored_count) is not int:
        raise ValueError("corpus counts must be integers")
    if not isinstance(chunk_mode, str) or not isinstance(text_mode, str):
        raise ValueError("corpus modes must be strings")
    return copy.deepcopy(_expected(source_count, authored_count, chunk_mode, text_mode))


def _same_typed_value(actual: Any, expected: Any) -> bool:
    # Python considers False == 0 and True == 1. Evidence fields must retain
    # their declared JSON types even when this oracle is used independently.
    if type(actual) is not type(expected):
        return False
    if isinstance(expected, dict):
        return actual.keys() == expected.keys() and all(
            _same_typed_value(actual[key], value) for key, value in expected.items()
        )
    return actual == expected


def validate_case(observed: dict[str, Any]) -> None:
    """Reject self-consistent report hashes that differ from the corpus spec."""
    expected = expected_case(
        observed.get("source_count"), observed.get("authored_count"),
        observed.get("chunk_mode"), observed.get("text_mode"),
    )
    for section, fields in expected.items():
        actual = observed.get(section)
        if not isinstance(actual, dict):
            raise ValueError(f"corpus oracle: missing {section}")
        for field, value in fields.items():
            if not _same_typed_value(actual.get(field), value):
                raise ValueError(f"corpus oracle: {section}.{field} differs from independent generator")
    proof_authored = observed.get("proof", {}).get("authored")
    if not isinstance(proof_authored, dict):
        raise ValueError("corpus oracle: missing proof.authored")
    for field, value in expected["authored"].items():
        if not _same_typed_value(proof_authored.get(field), value):
            raise ValueError(f"corpus oracle: proof.authored.{field} differs from independent generator")
