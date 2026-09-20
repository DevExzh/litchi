#!/usr/bin/env python3
"""Generate the bounded Unicode predicates used by ODF text reducers.

The generator intentionally uses only the Python standard library.  Unicode
17.0.0's official UCD files are downloaded into a caller supplied temporary
directory, checked against ``provenance.json``, and reduced to sorted,
non-overlapping ranges in ``evaluation/text/unicode.rs``.  The full UCD files
are not vendored in the repository.

Typical reproducible invocation::

    python3 generate.py --download-dir /tmp/ucd-17 \
      --output ../../../../../crates/litchi-ods/src/codec/formula/evaluation/text/unicode.rs \
      --verify

``--verify`` performs the complete scalar-range checks, boundary checks,
determinism check, source hash checks, and generated-output hash check when
the retained provenance has one.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
import urllib.request
from pathlib import Path
from typing import Iterable, Sequence


MAX_CODE_POINT = 0x10FFFF
SURROGATE_START = 0xD800
SURROGATE_END = 0xDFFF
UNICODE_VERSION = "17.0.0"
SCRIPT_DIR = Path(__file__).resolve().parent
PROVENANCE_PATH = SCRIPT_DIR / "provenance.json"


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def read_provenance() -> dict:
    with PROVENANCE_PATH.open("r", encoding="utf-8") as stream:
        provenance = json.load(stream)
    if provenance.get("unicode_version") != UNICODE_VERSION:
        raise ValueError("provenance Unicode version is not 17.0.0")
    return provenance


def fetch_sources(destination: Path, provenance: dict) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    for source in provenance["sources"]:
        target = destination / source["filename"]
        with urllib.request.urlopen(source["url"], timeout=60) as response:
            target.write_bytes(response.read())
        verify_hash(target, source)


def verify_hash(path: Path, source: dict) -> None:
    data = path.read_bytes()
    actual = sha256_bytes(data)
    if actual != source["sha256"]:
        raise ValueError(
            f"{path} hash mismatch: expected {source['sha256']}, got {actual}"
        )
    expected_size = source.get("bytes")
    if expected_size is not None and len(data) != expected_size:
        raise ValueError(
            f"{path} size mismatch: expected {expected_size}, got {len(data)}"
        )


def source_path(source_dir: Path, filename: str) -> Path:
    path = source_dir / filename
    if not path.is_file():
        raise FileNotFoundError(
            f"missing {path}; pass --download-dir or provide the pinned UCD files"
        )
    return path


def parse_code_point(token: str) -> int:
    value = int(token, 16)
    if not 0 <= value <= MAX_CODE_POINT:
        raise ValueError(f"code point outside Unicode scalar range: {token}")
    return value


def parse_range(token: str) -> tuple[int, int]:
    parts = token.strip().split("..")
    if len(parts) == 1:
        value = parse_code_point(parts[0])
        return value, value
    if len(parts) != 2:
        raise ValueError(f"invalid Unicode range: {token!r}")
    start = parse_code_point(parts[0])
    end = parse_code_point(parts[1])
    if start > end:
        raise ValueError(f"descending Unicode range: {token!r}")
    return start, end


def parse_unicode_data(path: Path) -> list[tuple[int, int, str]]:
    """Parse UnicodeData.txt, expanding First/Last records to intervals."""

    records: list[tuple[int, int, str]] = []
    pending: tuple[int, str] | None = None
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        if not raw_line or raw_line.startswith("#"):
            continue
        fields = raw_line.split(";")
        if len(fields) < 3:
            raise ValueError(f"malformed UnicodeData line: {raw_line!r}")
        code_point = parse_code_point(fields[0])
        name = fields[1]
        category = fields[2]
        if name.endswith(", First>"):
            if pending is not None:
                raise ValueError("nested UnicodeData First range")
            pending = code_point, category
            continue
        if name.endswith(", Last>"):
            if pending is None:
                raise ValueError("UnicodeData Last range without First")
            start, first_category = pending
            if first_category != category or start > code_point:
                raise ValueError("inconsistent UnicodeData First/Last range")
            records.append((start, code_point, category))
            pending = None
            continue
        if pending is not None:
            raise ValueError("UnicodeData First range without adjacent Last")
        records.append((code_point, code_point, category))
    if pending is not None:
        raise ValueError("unterminated UnicodeData First range")
    records.sort()
    previous_end = -1
    for start, end, _category in records:
        if start <= previous_end:
            raise ValueError("overlapping UnicodeData records")
        previous_end = end
    return records


def parse_mapping(text: str, context: str) -> tuple[int, ...]:
    values = tuple(parse_code_point(part) for part in text.split())
    if not 1 <= len(values) <= 3:
        raise ValueError(f"{context} mapping is not one to three scalars")
    if any(SURROGATE_START <= value <= SURROGATE_END for value in values):
        raise ValueError(f"{context} mapping contains a surrogate")
    return values


def parse_unicode_case_mappings(
    path: Path,
) -> tuple[dict[int, tuple[int, ...]], dict[int, tuple[int, ...]]]:
    """Read simple one-to-one upper/lower mappings from UnicodeData.txt."""

    lower: dict[int, tuple[int, ...]] = {}
    upper: dict[int, tuple[int, ...]] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        if not raw_line or raw_line.startswith("#"):
            continue
        fields = raw_line.split(";")
        if len(fields) < 15:
            raise ValueError(f"malformed UnicodeData line: {raw_line!r}")
        code_point = parse_code_point(fields[0])
        if fields[12]:
            upper[code_point] = parse_mapping(
                fields[12], f"UnicodeData U+{code_point:04X} uppercase"
            )
        if fields[13]:
            lower[code_point] = parse_mapping(
                fields[13], f"UnicodeData U+{code_point:04X} lowercase"
            )
    return lower, upper


def parse_special_casing(
    path: Path,
) -> tuple[dict[int, tuple[int, ...]], dict[int, tuple[int, ...]]]:
    """Read unconditional full lower/upper mappings from SpecialCasing.txt."""

    lower: dict[int, tuple[int, ...]] = {}
    upper: dict[int, tuple[int, ...]] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.split("#", 1)[0].strip()
        if not line:
            continue
        fields = [part.strip() for part in line.split(";")]
        while fields and not fields[-1]:
            fields.pop()
        if len(fields) < 4:
            raise ValueError(f"malformed SpecialCasing line: {raw_line!r}")
        if len(fields) > 4:
            # Contextual and language-specific mappings are selected by the
            # text algorithm; only the default unconditional rows belong here.
            continue
        code_point = parse_code_point(fields[0])
        lower[code_point] = parse_mapping(
            fields[1], f"SpecialCasing U+{code_point:04X} lowercase"
        )
        upper[code_point] = parse_mapping(
            fields[3], f"SpecialCasing U+{code_point:04X} uppercase"
        )
    return lower, upper


def parse_derived_property(path: Path, wanted: str) -> bytearray:
    values = bytearray(MAX_CODE_POINT + 1)
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.split("#", 1)[0].strip()
        if not line:
            continue
        try:
            range_text, property_name = (part.strip() for part in line.split(";", 1))
        except ValueError as error:
            raise ValueError(f"malformed DerivedCoreProperties line: {raw_line!r}") from error
        if property_name != wanted:
            continue
        start, end = parse_range(range_text)
        values[start : end + 1] = b"\x01" * (end - start + 1)
    return values


def parse_case_folding(path: Path) -> dict[int, tuple[int, ...]]:
    """Read the default full fold: common (C) and full (F), excluding T."""

    mappings: dict[int, tuple[int, ...]] = {}
    statuses: dict[int, str] = {}
    for raw_line in path.read_text(encoding="utf-8").splitlines():
        line = raw_line.split("#", 1)[0].strip()
        if not line:
            continue
        fields = [part.strip() for part in line.split(";")]
        while fields and not fields[-1]:
            fields.pop()
        if len(fields) != 3:
            raise ValueError(f"malformed CaseFolding line: {raw_line!r}")
        code_point = parse_code_point(fields[0])
        status = fields[1]
        if status not in {"C", "F"}:
            continue
        mapping = tuple(parse_code_point(part) for part in fields[2].split())
        if not 1 <= len(mapping) <= 3:
            raise ValueError(
                f"CaseFolding mapping at U+{code_point:04X} is not one to three scalars"
            )
        previous_status = statuses.get(code_point)
        if previous_status is not None:
            # The current UCD has no C/F duplicate.  If one appears in a
            # future file, full (F) is the default full-fold mapping.
            if previous_status == "F" or status == "C":
                continue
        statuses[code_point] = status
        mappings[code_point] = mapping
    return {
        code_point: mapping
        for code_point, mapping in mappings.items()
        if mapping != (code_point,)
    }


def category_array(records: Sequence[tuple[int, int, str]]) -> list[str | None]:
    categories: list[str | None] = [None] * (MAX_CODE_POINT + 1)
    for start, end, category in records:
        categories[start : end + 1] = [category] * (end - start + 1)
    return categories


def scalar_code_points() -> Iterable[int]:
    for code_point in range(MAX_CODE_POINT + 1):
        if SURROGATE_START <= code_point <= SURROGATE_END:
            continue
        yield code_point


def ranges_for(predicate: Sequence[bool] | bytearray) -> list[tuple[int, int]]:
    ranges: list[tuple[int, int]] = []
    start: int | None = None
    previous: int | None = None
    for code_point in scalar_code_points():
        if predicate[code_point]:
            if start is None:
                start = code_point
            elif previous is not None and code_point != previous + 1:
                ranges.append((start, previous))
                start = code_point
            previous = code_point
        elif start is not None:
            ranges.append((start, previous if previous is not None else start))
            start = None
            previous = None
    if start is not None:
        ranges.append((start, previous if previous is not None else start))
    return ranges


def build_ranges(source_dir: Path) -> dict[str, list[tuple[int, int]]]:
    records = parse_unicode_data(source_path(source_dir, "UnicodeData.txt"))
    categories = category_array(records)
    case_ignorable = parse_derived_property(
        source_path(source_dir, "DerivedCoreProperties.txt"), "Case_Ignorable"
    )
    cased = parse_derived_property(
        source_path(source_dir, "DerivedCoreProperties.txt"), "Cased"
    )

    clean_removed = bytearray(MAX_CODE_POINT + 1)
    letters = bytearray(MAX_CODE_POINT + 1)
    for code_point, category in enumerate(categories):
        if category is None or category == "Cc":
            clean_removed[code_point] = 1
        if category is not None and category.startswith("L"):
            letters[code_point] = 1

    ranges = {
        "CLEAN_REMOVED": ranges_for(clean_removed),
        "LETTER": ranges_for(letters),
        "CASED": ranges_for(cased),
        "CASE_IGNORABLE": ranges_for(case_ignorable),
    }
    verify_full_range(categories, cased, case_ignorable, ranges)
    return ranges


def case_fold_lookup(
    mappings: Sequence[tuple[int, tuple[int, ...]]], code_point: int
) -> tuple[int, ...]:
    low = 0
    high = len(mappings)
    while low < high:
        middle = (low + high) // 2
        source, mapping = mappings[middle]
        if code_point < source:
            high = middle
        elif code_point > source:
            low = middle + 1
        else:
            return mapping
    return (code_point,)


def build_case_folding(source_dir: Path) -> list[tuple[int, tuple[int, ...]]]:
    parsed = parse_case_folding(source_path(source_dir, "CaseFolding.txt"))
    mappings = sorted(parsed.items())
    verify_case_folding(mappings)
    return mappings


def build_case_mappings(
    source_dir: Path,
) -> tuple[list[tuple[int, tuple[int, ...]]], list[tuple[int, tuple[int, ...]]]]:
    lower, upper = parse_unicode_case_mappings(
        source_path(source_dir, "UnicodeData.txt")
    )
    special_lower, special_upper = parse_special_casing(
        source_path(source_dir, "SpecialCasing.txt")
    )
    lower.update(special_lower)
    upper.update(special_upper)
    lower_mappings = sorted(
        (code_point, mapping)
        for code_point, mapping in lower.items()
        if mapping != (code_point,)
    )
    upper_mappings = sorted(
        (code_point, mapping)
        for code_point, mapping in upper.items()
        if mapping != (code_point,)
    )
    verify_mapping_table("LOWERCASE", lower_mappings)
    verify_mapping_table("UPPERCASE", upper_mappings)
    return lower_mappings, upper_mappings


def verify_case_folding(
    mappings: Sequence[tuple[int, tuple[int, ...]]],
) -> None:
    """Compare every scalar with the pinned default full-fold mapping."""

    verify_mapping_table("CASE_FOLD", mappings)


def verify_mapping_table(
    name: str,
    mappings: Sequence[tuple[int, tuple[int, ...]]],
) -> None:
    """Check sorted mappings and their identity default over every scalar."""

    previous = -1
    for source, mapping in mappings:
        if not 0 <= source <= MAX_CODE_POINT or (
            SURROGATE_START <= source <= SURROGATE_END
        ):
            raise ValueError(f"CaseFolding source is not a Unicode scalar: U+{source:04X}")
        if source <= previous:
            raise ValueError("CaseFolding mappings are not sorted and disjoint")
        if not 1 <= len(mapping) <= 3:
            raise ValueError(f"CaseFolding mapping at U+{source:04X} has invalid length")
        for target in mapping:
            if not 0 <= target <= MAX_CODE_POINT or (
                SURROGATE_START <= target <= SURROGATE_END
            ):
                raise ValueError(
                    f"CaseFolding target is not a Unicode scalar: U+{target:04X}"
                )
        if mapping == (source,):
            raise ValueError(f"identity {name} mapping retained at U+{source:04X}")
        previous = source

    expected = dict(mappings)
    for code_point in scalar_code_points():
        actual = case_fold_lookup(mappings, code_point)
        if actual != expected.get(code_point, (code_point,)):
            raise ValueError(f"full-range {name} mismatch at U+{code_point:04X}")


def verify_ranges(ranges: dict[str, list[tuple[int, int]]]) -> None:
    for name, entries in ranges.items():
        previous_end = -1
        for start, end in entries:
            if not 0 <= start <= end <= MAX_CODE_POINT:
                raise ValueError(f"{name} contains an out-of-range interval")
            if SURROGATE_START <= start <= SURROGATE_END or (
                start < SURROGATE_START <= end
            ):
                raise ValueError(f"{name} contains a surrogate interval")
            if start <= previous_end or (start <= SURROGATE_END and end >= SURROGATE_START):
                raise ValueError(f"{name} is not sorted and disjoint")
            previous_end = end


def verify_full_range(
    categories: Sequence[str | None],
    cased: bytearray,
    case_ignorable: bytearray,
    ranges: dict[str, list[tuple[int, int]]],
) -> None:
    """Compare every scalar value with the parsed UCD source properties."""

    for code_point in scalar_code_points():
        category = categories[code_point]
        expected = {
            "CLEAN_REMOVED": category is None or category == "Cc",
            "LETTER": category is not None and category.startswith("L"),
            "CASED": bool(cased[code_point]),
            "CASE_IGNORABLE": bool(case_ignorable[code_point]),
        }
        for name, expected_value in expected.items():
            if contains(ranges[name], code_point) != expected_value:
                raise ValueError(
                    f"full-range mismatch for {name} at U+{code_point:04X}"
                )


def contains(ranges: Sequence[tuple[int, int]], code_point: int) -> bool:
    low = 0
    high = len(ranges)
    while low < high:
        middle = (low + high) // 2
        start, end = ranges[middle]
        if code_point < start:
            high = middle
        elif code_point > end:
            low = middle + 1
        else:
            return True
    return False


def verify_boundaries(ranges: dict[str, list[tuple[int, int]]]) -> None:
    cases = {
        "CLEAN_REMOVED": {
            0x0000: True,
            0x001F: True,
            0x0020: False,
            0x007F: True,
            0x009F: True,
            0x00A0: False,
            0x0378: True,
            0xE000: False,
            0x10FFFF: True,
        },
        "LETTER": {
            0x0041: True,
            0x0061: True,
            0x0301: False,
            0x1F600: False,
            0x4E00: True,
            0x10FFFF: False,
        },
        "CASED": {
            0x0041: True,
            0x0061: True,
            0x03A3: True,
            0x0030: False,
            0x0301: False,
        },
        "CASE_IGNORABLE": {
            0x0301: True,
            0x0027: True,
            0x0041: False,
            0x0030: False,
        },
    }
    for name, expected in cases.items():
        for code_point, value in expected.items():
            actual = contains(ranges[name], code_point)
            if actual != value:
                raise ValueError(
                    f"{name} boundary U+{code_point:04X}: expected {value}, got {actual}"
                )


def rust_range_table(name: str, entries: Sequence[tuple[int, int]]) -> str:
    lines = [f"const {name}: &[Range] = &[\n"]
    for start, end in entries:
        lines.append(
            f"    Range {{\n        start: 0x{start:06X},\n        end: 0x{end:06X},\n    }},\n"
        )
    lines.append("];\n")
    return "".join(lines)


def rust_mapping_table(
    name: str, entries: Sequence[tuple[int, tuple[int, ...]]]
) -> str:
    lines = [f"const {name}: &[Mapping] = &[\n"]
    for source, mapping in entries:
        target = tuple(mapping) + (0,) * (3 - len(mapping))
        lines.append(
            "    Mapping {\n"
            f"        source: 0x{source:06X},\n"
            f"        len: {len(mapping)},\n"
            "        target: ["
            + ", ".join(f"0x{value:06X}" for value in target)
            + "],\n"
            "    },\n"
        )
    lines.append("];\n")
    return "".join(lines)


def render_rust(
    ranges: dict[str, list[tuple[int, int]]],
    case_folding: Sequence[tuple[int, tuple[int, ...]]],
    lowercase: Sequence[tuple[int, tuple[int, ...]]],
    uppercase: Sequence[tuple[int, tuple[int, ...]]],
    provenance: dict,
) -> bytes:
    source_hashes = ", ".join(
        f"{source['filename']}={source['sha256']}" for source in provenance["sources"]
    )
    body = [
        "//! Unicode 17.0.0 predicates generated from the official UCD files.\n",
        "//!\n",
        "//! Do not edit manually; rerun unicode-data/generate.py after changing\n",
        "//! the pinned source receipts.  The tables contain only Unicode scalar\n",
        "//! values and are searched without allocation or ambient I/O.\n",
        f"//! Source SHA-256: {source_hashes}\n\n",
        "#[derive(Clone, Copy)]\n",
        "struct Range {\n    start: u32,\n    end: u32,\n}\n\n",
        rust_range_table("CLEAN_REMOVED", ranges["CLEAN_REMOVED"]),
        "\n",
        rust_range_table("LETTER", ranges["LETTER"]),
        "\n",
        rust_range_table("CASED", ranges["CASED"]),
        "\n",
        rust_range_table("CASE_IGNORABLE", ranges["CASE_IGNORABLE"]),
        "\n",
        "#[derive(Clone, Copy)]\n",
        "struct Mapping {\n    source: u32,\n    len: u8,\n    target: [u32; 3],\n}\n\n",
        rust_mapping_table("CASE_FOLD", case_folding),
        "\n",
        rust_mapping_table("LOWERCASE", lowercase),
        "\n",
        rust_mapping_table("UPPERCASE", uppercase),
        "\n",
        "#[inline]\n",
        "fn contains(ranges: &[Range], code_point: u32) -> bool {\n",
        "    let mut low = 0_usize;\n",
        "    let mut high = ranges.len();\n",
        "    while low < high {\n",
        "        let middle = low + (high - low) / 2;\n",
        "        let range = ranges[middle];\n",
        "        if code_point < range.start {\n",
        "            high = middle;\n",
        "        } else if code_point > range.end {\n",
        "            low = middle + 1;\n",
        "        } else {\n",
        "            return true;\n",
        "        }\n",
        "    }\n",
        "    false\n",
        "}\n\n",
        "#[inline]\n",
        "pub(super) fn is_clean_removed(value: char) -> bool {\n",
        "    contains(CLEAN_REMOVED, value as u32)\n",
        "}\n\n",
        "#[inline]\n",
        "pub(super) fn is_letter(value: char) -> bool {\n",
        "    contains(LETTER, value as u32)\n",
        "}\n\n",
        "#[inline]\n",
        "pub(super) fn is_cased(value: char) -> bool {\n",
        "    contains(CASED, value as u32)\n",
        "}\n\n",
        "#[inline]\n",
        "pub(super) fn is_case_ignorable(value: char) -> bool {\n",
        "    contains(CASE_IGNORABLE, value as u32)\n",
        "}\n",
        "\n",
        "#[inline]\n",
        "fn find_mapping(table: &[Mapping], code_point: u32) -> Option<Mapping> {\n",
        "    let mut low = 0_usize;\n",
        "    let mut high = table.len();\n",
        "    while low < high {\n",
        "        let middle = low + (high - low) / 2;\n",
        "        let mapping = table[middle];\n",
        "        if code_point < mapping.source {\n",
        "            high = middle;\n",
        "        } else if code_point > mapping.source {\n",
        "            low = middle + 1;\n",
        "        } else {\n",
        "            return Some(mapping);\n",
        "        }\n",
        "    }\n",
        "    None\n",
        "}\n\n",
        "#[derive(Clone, Copy)]\n",
        "pub(super) struct CaseMapping {\n",
        "    chars: [char; 3],\n",
        "    len: u8,\n",
        "}\n\n",
        "impl CaseMapping {\n",
        "    #[inline]\n",
        "    pub(super) fn iter(self) -> impl Iterator<Item = char> {\n",
        "        self.chars.into_iter().take(usize::from(self.len))\n",
        "    }\n",
        "}\n\n",
        "#[inline]\n",
        "fn map_character(table: &[Mapping], value: char) -> CaseMapping {\n",
        "    if let Some(mapping) = find_mapping(table, value as u32) {\n",
        "        CaseMapping {\n",
        "            chars: [\n",
        "                char::from_u32(mapping.target[0]).expect(\"generated Unicode scalar\"),\n",
        "                char::from_u32(mapping.target[1]).expect(\"generated Unicode scalar\"),\n",
        "                char::from_u32(mapping.target[2]).expect(\"generated Unicode scalar\"),\n",
        "            ],\n",
        "            len: mapping.len,\n",
        "        }\n",
        "    } else {\n",
        "        CaseMapping {\n",
        "            chars: [value, value, value],\n",
        "            len: 1,\n",
        "        }\n",
        "    }\n",
        "}\n\n",
        "#[inline]\n",
        "pub(super) fn case_fold(value: char) -> CaseMapping {\n",
        "    map_character(CASE_FOLD, value)\n",
        "}\n\n",
        "#[inline]\n",
        "pub(super) fn lowercase(value: char) -> CaseMapping {\n",
        "    map_character(LOWERCASE, value)\n",
        "}\n\n",
        "#[inline]\n",
        "pub(super) fn uppercase(value: char) -> CaseMapping {\n",
        "    map_character(UPPERCASE, value)\n",
        "}\n",
        "\n",
        "#[cfg(test)]\n",
        "mod tests {\n",
        "    use super::{\n",
        "        CASE_FOLD, CASE_IGNORABLE, CASED, CLEAN_REMOVED, LETTER, LOWERCASE, UPPERCASE, contains,\n",
        "        find_mapping,\n",
        "    };\n",
        "\n",
        "    fn assert_table_boundaries(table: &[super::Range]) {\n",
        "        let mut previous_end = None;\n",
        "        for range in table {\n",
        "            assert!(range.start <= range.end);\n",
        "            assert!(range.end < 0xD800 || range.start > 0xDFFF);\n",
        "            if let Some(end) = previous_end {\n",
        "                assert!(range.start > end);\n",
        "            }\n",
        "            assert!(contains(table, range.start));\n",
        "            assert!(contains(table, range.end));\n",
        "            previous_end = Some(range.end);\n",
        "        }\n",
        "    }\n",
        "\n",
        "    #[test]\n",
        "    fn generated_tables_are_sorted_scalar_ranges() {\n",
        "        assert_table_boundaries(CLEAN_REMOVED);\n",
        "        assert_table_boundaries(LETTER);\n",
        "        assert_table_boundaries(CASED);\n",
        "        assert_table_boundaries(CASE_IGNORABLE);\n",
        "    }\n",
        "\n",
        "    #[test]\n",
        "    fn unicode_boundary_classes_are_pinned() {\n",
        "        assert!(super::is_clean_removed('\\0'));\n",
        "        assert!(super::is_clean_removed('\\u{0378}'));\n",
        "        assert!(!super::is_clean_removed(' '));\n",
        "        assert!(!super::is_clean_removed('\\u{E000}'));\n",
        "        assert!(super::is_letter('A'));\n",
        "        assert!(super::is_letter('\\u{4E00}'));\n",
        "        assert!(!super::is_letter('\\u{0301}'));\n",
        "        assert!(super::is_cased('\\u{03A3}'));\n",
        "        assert!(!super::is_cased('0'));\n",
        "        assert!(super::is_case_ignorable('\\u{0301}'));\n",
        "        assert!(super::is_case_ignorable(char::from_u32(0x27).unwrap()));\n",
        "        assert!(!super::is_case_ignorable('A'));\n",
        "    }\n",
        "\n",
        "    fn assert_mapping_boundaries(table: &[super::Mapping]) {\n",
        "        let mut previous_source = None;\n",
        "        for mapping in table {\n",
        "            assert!((1..=3).contains(&mapping.len));\n",
        "            if let Some(source) = previous_source {\n",
        "                assert!(mapping.source > source);\n",
        "            }\n",
        "            assert_eq!(\n",
        "                find_mapping(table, mapping.source).unwrap().source,\n",
        "                mapping.source,\n",
        "            );\n",
        "            for target in mapping.target.iter().take(usize::from(mapping.len)) {\n",
        "                assert!(char::from_u32(*target).is_some());\n",
        "            }\n",
        "            previous_source = Some(mapping.source);\n",
        "        }\n",
        "    }\n",
        "\n",
        "    #[test]\n",
        "    fn generated_case_mappings_are_sorted_bounded_scalars() {\n",
        "        assert_mapping_boundaries(CASE_FOLD);\n",
        "        assert_mapping_boundaries(LOWERCASE);\n",
        "        assert_mapping_boundaries(UPPERCASE);\n",
        "    }\n",
        "\n",
        "    #[test]\n",
        "    fn default_case_mappings_are_pinned() {\n",
        "        assert_eq!(super::case_fold('A').iter().collect::<Vec<_>>(), vec!['a']);\n",
        "        assert_eq!(\n",
        "            super::case_fold('\\u{00DF}').iter().collect::<Vec<_>>(),\n",
        "            vec!['s', 's'],\n",
        "        );\n",
        "        assert_eq!(\n",
        "            super::case_fold('\\u{0130}').iter().collect::<Vec<_>>(),\n",
        "            vec!['i', '\\u{0307}'],\n",
        "        );\n",
        "        assert_eq!(super::case_fold('I').iter().collect::<Vec<_>>(), vec!['i']);\n",
        "        assert_eq!(\n",
        "            super::lowercase('\\u{0130}').iter().collect::<Vec<_>>(),\n",
        "            vec!['i', '\\u{0307}'],\n",
        "        );\n",
        "        assert_eq!(\n",
        "            super::uppercase('\\u{00DF}').iter().collect::<Vec<_>>(),\n",
        "            vec!['S', 'S'],\n",
        "        );\n",
        "        assert_eq!(\n",
        "            super::uppercase('\\u{FB03}').iter().collect::<Vec<_>>(),\n",
        "            vec!['F', 'F', 'I'],\n",
        "        );\n",
        "        assert_eq!(super::lowercase('A').iter().collect::<Vec<_>>(), vec!['a']);\n",
        "        assert_eq!(super::uppercase('a').iter().collect::<Vec<_>>(), vec!['A']);\n",
        "    }\n",
        "}\n",
    ]
    return "".join(body).encode("utf-8")


def generate(
    source_dir: Path,
) -> tuple[
    bytes,
    dict[str, list[tuple[int, int]]],
    list[tuple[int, tuple[int, ...]]],
    list[tuple[int, tuple[int, ...]]],
    list[tuple[int, tuple[int, ...]]],
]:
    ranges = build_ranges(source_dir)
    verify_ranges(ranges)
    verify_boundaries(ranges)
    case_folding = build_case_folding(source_dir)
    lowercase, uppercase = build_case_mappings(source_dir)
    provenance = read_provenance()
    output = render_rust(ranges, case_folding, lowercase, uppercase, provenance)
    return output, ranges, case_folding, lowercase, uppercase


def verify_output(output: bytes, output_path: Path, provenance: dict) -> None:
    if not output_path.is_file():
        raise ValueError(f"generated output is missing: {output_path}")
    expected = provenance.get("generated", {}).get("sha256")
    if expected and sha256_bytes(output) != expected:
        raise ValueError(
            f"generated output hash mismatch: expected {expected}, got {sha256_bytes(output)}"
        )
    if output_path.read_bytes() != output:
        raise ValueError(f"generated output differs from {output_path}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-dir", type=Path)
    parser.add_argument("--download-dir", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--verify", action="store_true")
    args = parser.parse_args()

    provenance = read_provenance()
    if args.download_dir is not None:
        fetch_sources(args.download_dir, provenance)
        source_dir = args.download_dir
    elif args.source_dir is not None:
        source_dir = args.source_dir
    else:
        parser.error("one of --source-dir or --download-dir is required")

    for source in provenance["sources"]:
        verify_hash(source_path(source_dir, source["filename"]), source)

    first_output, ranges, case_folding, lowercase, uppercase = generate(source_dir)
    (
        second_output,
        second_ranges,
        second_case_folding,
        second_lowercase,
        second_uppercase,
    ) = generate(source_dir)
    if (
        first_output != second_output
        or ranges != second_ranges
        or case_folding != second_case_folding
        or lowercase != second_lowercase
        or uppercase != second_uppercase
    ):
        raise ValueError("generator is not deterministic")
    if args.verify:
        verify_output(first_output, args.output, provenance)
    else:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_bytes(first_output)

    print(
        json.dumps(
            {
                "unicode_version": UNICODE_VERSION,
                "output_sha256": sha256_bytes(first_output),
                "ranges": {name: len(entries) for name, entries in ranges.items()},
                "mapping_counts": {
                    "CASE_FOLD": len(case_folding),
                    "LOWERCASE": len(lowercase),
                    "UPPERCASE": len(uppercase),
                },
                "verified": bool(args.verify),
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (FileNotFoundError, OSError, ValueError) as error:
        print(f"generate.py: {error}", file=sys.stderr)
        raise SystemExit(1) from error
