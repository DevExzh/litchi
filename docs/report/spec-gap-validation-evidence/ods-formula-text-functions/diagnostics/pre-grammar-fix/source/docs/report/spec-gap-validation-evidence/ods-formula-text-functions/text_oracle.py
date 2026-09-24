#!/usr/bin/env python3
"""Independent oracle for the OpenFormula 1.4 §6.20 text functions.

The generator has no dependency on the Rust evaluator.  Formula operands are
kept as typed JSON cells, text operations use Python's Unicode scalar strings,
and the width mappings are transcribed from ODF Tables 33 and 34.  The
retained observations are intentionally small and explicit: their ``formula``
field is the exact evaluator input while ``args``/``cells`` are the independent
oracle input used to derive the expected result.

Python 3.14 in the gate environment exposes Unicode 16 data.  The base profile
uses only stable mappings for ordinary case rows and pins the selected Unicode
17 additions explicitly (the U+A7CE/U+A7CF pair), so this oracle does not
silently claim that the host Python UCD is the evaluator's UCD.
"""

from __future__ import annotations

from datetime import date, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
from fractions import Fraction
import hashlib
import json
import math
from pathlib import Path
import struct
import unicodedata


HERE = Path(__file__).resolve().parent
CONTRACT = HERE / "contract.md"
GOLDENS = HERE / "text-goldens.json"

# Filled with the final frozen contract digest after the contract owner stops
# editing.  Keeping this optional during integration prevents a concurrent
# contract write from producing a false oracle mismatch.
CONTRACT_SHA256: str | None = "b0b7f6ffd8a33c93f98c9728c938eb2d55476680e11797a9b87dcf49223312e7"
ODF_ARCHIVE_SHA256 = "9867665f9702b365076c2c6557b23c8c938959b443f6f50712fdb2d0dfb8aac4"
ODF_MEMBER_SHA256 = "ace07938ef54303b57af8472e0b66b289fc6946c32390fc23b8e13fdeeb5ffa1"
UNICODE_PROFILE = "17.0.0-selected-pinned-data"

ERRORS = {
    "Value",
    "Number",
    "NotAvailable",
    "DivisionByZero",
    "Reference",
}

FUNCTIONS = (
    "ASC",
    "CHAR",
    "CLEAN",
    "CODE",
    "CONCATENATE",
    "DOLLAR",
    "EXACT",
    "FIND",
    "FIXED",
    "JIS",
    "LEFT",
    "LEN",
    "LOWER",
    "MID",
    "PROPER",
    "REPLACE",
    "REPT",
    "RIGHT",
    "SEARCH",
    "SUBSTITUTE",
    "T",
    "TEXT",
    "TRIM",
    "UNICHAR",
    "UNICODE",
    "UPPER",
)


def bits(value: float) -> str:
    return struct.pack(">d", float(value)).hex()


def number(value: float) -> dict[str, object]:
    return {"kind": "Number", "bits": bits(value)}


def text(value: str) -> dict[str, object]:
    return {"kind": "Text", "value": value}


def logical(value: bool) -> dict[str, object]:
    return {"kind": "Logical", "value": bool(value)}


def empty() -> dict[str, object]:
    return {"kind": "Empty"}


def error(value: str) -> dict[str, object]:
    if value not in ERRORS:
        raise ValueError(value)
    return {"kind": "Error", "value": value}


def missing() -> dict[str, object]:
    return {"kind": "Missing"}


def complex_value(real: float, imaginary: float) -> dict[str, object]:
    return {"kind": "Complex", "real": bits(real), "imaginary": bits(imaginary)}


def decode_number(value: dict[str, object]) -> float:
    return struct.unpack(">d", bytes.fromhex(str(value["bits"])))[0]


class OracleError(Exception):
    def __init__(self, kind: str):
        if kind not in ERRORS:
            raise ValueError(kind)
        self.kind = kind
        super().__init__(kind)


def propagate(value: dict[str, object]) -> None:
    if value.get("kind") == "Error":
        raise OracleError(str(value["value"]))


def to_text(value: dict[str, object]) -> str:
    propagate(value)
    kind = value.get("kind")
    if kind == "Text":
        return str(value["value"])
    if kind == "Empty" or kind == "Missing":
        if kind == "Missing":
            raise OracleError("Value")
        return ""
    if kind == "Logical":
        return "TRUE" if value["value"] else "FALSE"
    if kind == "Number":
        number_value = decode_number(value)
        if not math.isfinite(number_value):
            raise OracleError("Number")
        # Match the selected shortest round-trip bridge's ordinary spelling:
        # Rust's finite `Display` omits the integral `.0` suffix.  The corpus
        # keeps this independent of the evaluator and avoids locale state.
        if number_value == 0.0:
            return "-0" if math.copysign(1.0, number_value) < 0 else "0"
        if number_value.is_integer() and abs(number_value) < 1e21:
            return str(int(number_value))
        return repr(number_value)
    raise OracleError("Value")


def to_number(value: dict[str, object]) -> float:
    propagate(value)
    kind = value.get("kind")
    if kind == "Number":
        result = decode_number(value)
    elif kind == "Logical":
        result = 1.0 if value["value"] else 0.0
    elif kind == "Empty":
        result = 0.0
    elif kind == "Text":
        try:
            result = float(str(value["value"]))
        except ValueError as exc:
            raise OracleError("Value") from exc
    else:
        raise OracleError("Value")
    if not math.isfinite(result):
        raise OracleError("Number")
    return result


def to_integer(value: dict[str, object]) -> int:
    result = to_number(value)
    if not math.isfinite(result):
        raise OracleError("Number")
    return math.trunc(result)


def to_integer_floor(value: dict[str, object]) -> int:
    result = to_number(value)
    if not math.isfinite(result):
        raise OracleError("Number")
    return math.floor(result)


def scalar_count(text_value: str) -> int:
    return len(text_value)


def asc_char(char: str) -> tuple[str, str | None]:
    code = ord(char)
    if 0x30A1 <= code <= 0x30AA and code % 2 == 0:
        return chr((code - 0x30A2) // 2 + 0xFF71), None
    if 0x30A1 <= code <= 0x30AA:
        return chr((code - 0x30A1) // 2 + 0xFF67), None
    if 0x30AB <= code <= 0x30C2 and code % 2 == 0:
        return chr((code - 0x30AC) // 2 + 0xFF76), "\uFF9E"
    if 0x30AB <= code <= 0x30C2:
        return chr((code - 0x30AB) // 2 + 0xFF76), None
    if code == 0x30C3:
        return "\uFF6F", None
    if 0x30C4 <= code <= 0x30C9 and code % 2 == 1:
        return chr((code - 0x30C5) // 2 + 0xFF82), "\uFF9E"
    if 0x30C4 <= code <= 0x30C9:
        return chr((code - 0x30C4) // 2 + 0xFF82), None
    if 0x30CA <= code <= 0x30CE:
        return chr(code - 0x30CA + 0xFF85), None
    if 0x30CF <= code <= 0x30DD and code % 3 == 1:
        return chr((code - 0x30D0) // 3 + 0xFF8A), "\uFF9E"
    if 0x30CF <= code <= 0x30DD and code % 3 == 2:
        return chr((code - 0x30D1) // 3 + 0xFF8A), "\uFF9F"
    if 0x30CF <= code <= 0x30DD:
        return chr((code - 0x30CF) // 3 + 0xFF8A), None
    if 0x30DE <= code <= 0x30E2:
        return chr(code - 0x30DE + 0xFF8F), None
    if 0x30E3 <= code <= 0x30E8 and code % 2 == 0:
        return chr((code - 0x30E4) // 2 + 0xFF94), None
    if 0x30E3 <= code <= 0x30E8:
        return chr((code - 0x30E3) // 2 + 0xFF6C), None
    if 0x30E9 <= code <= 0x30ED:
        return chr(code - 0x30E9 + 0xFF97), None
    special = {0x30EF: "\uFF9C", 0x30F2: "\uFF66", 0x30F3: "\uFF9D"}
    if code in special:
        return special[code], None
    if 0xFF01 <= code <= 0xFF5E:
        return chr(code - 0xFF01 + 0x21), None
    special = {
        0x2015: "\uFF70",
        0x2018: "`",
        0x2019: "'",
        0x201D: '"',
        0x3001: "\uFF64",
        0x3002: "\uFF61",
        0x300C: "\uFF62",
        0x300D: "\uFF63",
        0x309B: "\uFF9E",
        0x309C: "\uFF9F",
        0x30FB: "\uFF65",
        0x30FC: "\uFF70",
        0xFFE5: "\\",
    }
    return special.get(code, char), None


def jis_char(char: str, next_char: str | None) -> tuple[str, bool]:
    code = ord(char)
    exceptions = {0x22: "\u201D", 0x5C: "\uFFE5", 0x60: "\u2018", 0x27: "\u2019"}
    if code in exceptions:
        return exceptions[code], False
    if 0x21 <= code <= 0x7E:
        return chr(code - 0x21 + 0xFF01), False
    if code == 0xFF66:
        return "\u30F2", False
    if 0xFF67 <= code <= 0xFF6B:
        return chr((code - 0xFF67) * 2 + 0x30A1), False
    if 0xFF6C <= code <= 0xFF6E:
        return chr((code - 0xFF6C) * 2 + 0x30E3), False
    if code == 0xFF6F:
        return "\u30C3", False
    if 0xFF71 <= code <= 0xFF75:
        return chr((code - 0xFF71) * 2 + 0x30A2), False
    if 0xFF76 <= code <= 0xFF81:
        if next_char == "\uFF9E":
            return chr((code - 0xFF76) * 2 + 0x30AC), True
        return chr((code - 0xFF76) * 2 + 0x30AB), False
    if 0xFF82 <= code <= 0xFF84:
        if next_char == "\uFF9E":
            return chr((code - 0xFF82) * 2 + 0x30C5), True
        return chr((code - 0xFF82) * 2 + 0x30C4), False
    if 0xFF85 <= code <= 0xFF89:
        return chr(code - 0xFF85 + 0x30CA), False
    if 0xFF8A <= code <= 0xFF8E:
        if next_char == "\uFF9E":
            return chr((code - 0xFF8A) * 3 + 0x30D0), True
        if next_char == "\uFF9F":
            return chr((code - 0xFF8A) * 3 + 0x30D1), True
        return chr((code - 0xFF8A) * 3 + 0x30CF), False
    if 0xFF8F <= code <= 0xFF93:
        return chr(code - 0xFF8F + 0x30DE), False
    if 0xFF94 <= code <= 0xFF96:
        return chr((code - 0xFF94) * 2 + 0x30E4), False
    if 0xFF97 <= code <= 0xFF9B:
        return chr(code - 0xFF97 + 0x30E9), False
    special = {
        0xFF9C: "\u30EF",
        0xFF9D: "\u30F3",
        0xFF9E: "\u309B",
        0xFF9F: "\u309C",
        0xFF70: "\u30FC",
        0xFF61: "\u3002",
        0xFF62: "\u300C",
        0xFF63: "\u300D",
        0xFF64: "\u3001",
        0xFF65: "\u30FB",
    }
    return special.get(code, char), False


def asc(value: str) -> str:
    output: list[str] = []
    for char in value:
        mapped, suffix = asc_char(char)
        output.append(mapped)
        if suffix:
            output.append(suffix)
    return "".join(output)


def jis(value: str) -> str:
    output: list[str] = []
    chars = list(value)
    index = 0
    while index < len(chars):
        next_char = chars[index + 1] if index + 1 < len(chars) else None
        mapped, consumed = jis_char(chars[index], next_char)
        output.append(mapped)
        index += 2 if consumed else 1
    return "".join(output)


# Unicode 17 mappings newly assigned after the Python 16 UCD.  They are kept
# in this oracle as selected official data, rather than asking Python to infer
# them from an older database.
UNICODE17_LOWER = {"\uA7CE": "\uA7CF", "\uA7D2": "\uA7D3", "\uA7D4": "\uA7D5"}
UNICODE17_UPPER = {value: key for key, value in UNICODE17_LOWER.items()}


def lower(value: str) -> str:
    # Python supplies the default context-sensitive rules (notably Greek
    # final sigma); only the selected Unicode-17 additions need an explicit
    # post-map because this host carries UCD 16.
    return "".join(UNICODE17_LOWER.get(char, char) for char in value.lower())


def upper(value: str) -> str:
    return "".join(UNICODE17_UPPER.get(char, char) for char in value.upper())


def lower_at(value: str, offset: int) -> str:
    """Lower one scalar while retaining the source's contextual casing."""
    character = value[offset]
    if character == "\u03a3":
        before = value[:offset]
        after = value[offset + 1 :]
        preceded = next((c for c in reversed(before) if not unicodedata.category(c).startswith("M")), None)
        followed = next((c for c in after if not unicodedata.category(c).startswith("M")), None)
        if preceded is not None and preceded.isalpha() and not (followed is not None and followed.isalpha()):
            return "ς"
    return "".join(UNICODE17_LOWER.get(char, char) for char in character.lower())


def proper(value: str) -> str:
    result: list[str] = []
    previous_is_letter = False
    for offset, char in enumerate(value):
        # Python's alphabetic predicate follows the UCD for ordinary cases;
        # pin the Unicode-17 additions used in this corpus as letters.
        is_letter = char.isalpha() or char in UNICODE17_LOWER or char in UNICODE17_UPPER
        if is_letter:
            result.append(lower_at(value, offset) if previous_is_letter else upper(char))
        else:
            result.append(char)
        previous_is_letter = is_letter
    return "".join(result)


def trim(value: str) -> str:
    spaces = "\t\n\r "
    start = 0
    end = len(value)
    while start < end and value[start] in spaces:
        start += 1
    while end > start and value[end - 1] in spaces:
        end -= 1
    body = value[start:end]
    out: list[str] = []
    in_space = False
    for char in body:
        if char in spaces:
            if not in_space:
                out.append(" ")
            in_space = True
        else:
            out.append(char)
            in_space = False
    return "".join(out)


def clean(value: str) -> str:
    return "".join(
        char for char in value if unicodedata.category(char) not in {"Cc", "Cn"}
    )


def find_positions(needle: str, haystack: str, start: int, *, insensitive: bool) -> int:
    if start < 1 or start > len(haystack) + 1:
        raise OracleError("Value")
    if needle == "":
        return start
    if not insensitive:
        offset = haystack.find(needle, start - 1)
        if offset < 0:
            raise OracleError("Value")
        return offset + 1

    # Full case folding can expand one source scalar (for example, ß -> ss).
    # Searching the flattened string alone would accept a match that starts or
    # ends in the middle of such an expansion.  Keep a boundary table for the
    # original scalar sequence and compare every source-bounded candidate.
    folded_parts: list[str] = []
    boundaries = [0]
    for character in haystack:
        folded_parts.append(character.casefold())
        boundaries.append(boundaries[-1] + len(folded_parts[-1]))
    folded_haystack = "".join(folded_parts)
    folded_needle = needle.casefold()
    if folded_needle == "":
        return start
    for source_start in range(start - 1, len(haystack)):
        folded_start = boundaries[source_start]
        for source_end in range(source_start + 1, len(haystack) + 1):
            folded_end = boundaries[source_end]
            if folded_haystack[folded_start:folded_end] == folded_needle:
                return source_start + 1
    raise OracleError("Value")


def round_decimal(value: float, decimals: int) -> Decimal:
    # Decimal(str(...)) deliberately rounds the represented decimal spelling
    # supplied by the fixture under the selected provider's half-away rule.
    decimal_value = Decimal(str(value))
    quantum = Decimal(1).scaleb(-decimals) if decimals >= 0 else Decimal(1).scaleb(-decimals)
    precision = max(100, abs(decimal_value.adjusted()) + abs(decimals) + 24)
    with localcontext() as context:
        context.prec = precision
        rounded = decimal_value.quantize(quantum, rounding=ROUND_HALF_UP)
    # A rounded negative value whose magnitude is below the requested unit is
    # the canonical non-negative zero in this profile.
    return rounded.copy_abs() if rounded.is_zero() else rounded


def decimal_plain(value: Decimal, decimals: int) -> str:
    if value.is_zero():
        value = value.copy_abs()
    if decimals >= 0:
        return f"{value:.{decimals}f}"
    return format(value, "f").split(".", 1)[0]


def grouped(value: str) -> str:
    sign = ""
    if value.startswith("-"):
        sign, value = "-", value[1:]
    padding = value[: len(value) - len(value.lstrip(" "))]
    value = value[len(padding) :]
    integer, dot, fraction = value.partition(".")
    chunks: list[str] = []
    while integer:
        chunks.insert(0, integer[-3:])
        integer = integer[:-3]
    result = ",".join(chunks) if chunks else "0"
    return sign + padding + result + (dot + fraction if dot else "")


def placeholder_assignments(tokens: list[str], digits: str) -> tuple[str, list[str]]:
    """Place represented digits in slots from the decimal point outward.

    Integer and exponent slots consume digits from the right.  An unfilled
    required slot emits zero, an unfilled question slot emits its alignment
    space, and an unfilled ``#`` slot emits nothing.  Keeping the token-level
    assignments here prevents a count-only implementation from moving
    required zeroes ahead of question slots.
    """
    extra_count = max(0, len(digits) - len(tokens))
    extra = digits[:extra_count]
    assigned = digits[extra_count:]
    first_assigned = len(tokens) - len(assigned)
    rendered: list[str] = []
    for index, token in enumerate(tokens):
        if index >= first_assigned:
            rendered.append(assigned[index - first_assigned])
        elif token == "0":
            rendered.append("0")
        elif token == "?":
            rendered.append(" ")
        else:
            rendered.append("")
    return extra, rendered


def render_slot_pattern(pattern: str, digits: str, *, omit_zero: bool = False) -> str:
    """Render one integer-like placeholder run with positional slots."""
    tokens = [char for char in pattern if char in "0#?"]
    if omit_zero and digits == "0" and "0" not in tokens:
        digits = ""
    extra, assignments = placeholder_assignments(tokens, digits)
    output: list[str] = [extra]
    token_index = 0
    for char in pattern:
        if char in "0#?":
            output.append(assignments[token_index])
            token_index += 1
        else:
            output.append(char)
    return "".join(output)


def currency(value: float, decimals: int) -> str:
    rounded = decimal_plain(round_decimal(value, decimals), decimals)
    if rounded.startswith("-"):
        # The production invariant formatter uses accounting-style negative
        # currency.  Keep this explicit instead of inheriting a locale's
        # negative-number convention.
        return "($" + grouped(rounded[1:]) + ")"
    return "$" + grouped(rounded)


def fixed(value: float, decimals: int, omit: bool) -> str:
    rounded = decimal_plain(round_decimal(value, decimals), decimals)
    if omit:
        return rounded
    return grouped(rounded)


def split_format_sections(source: str) -> list[str]:
    if source == "":
        raise OracleError("Value")
    sections: list[str] = []
    start = 0
    quoted = False
    index = 0
    while index < len(source):
        char = source[index]
        if char == '"':
            if quoted and index + 1 < len(source) and source[index + 1] == '"':
                index += 2
                continue
            quoted = not quoted
        elif char == ";" and not quoted:
            sections.append(source[start:index])
            start = index + 1
        index += 1
    if quoted:
        raise OracleError("Value")
    sections.append(source[start:])
    if not 1 <= len(sections) <= 4:
        raise OracleError("Value")
    return sections


def format_literal(source: str, substitution: str | None = None) -> str:
    output: list[str] = []
    quoted = False
    index = 0
    while index < len(source):
        char = source[index]
        if char == '"':
            if quoted and index + 1 < len(source) and source[index + 1] == '"':
                output.append('"')
                index += 2
            else:
                quoted = not quoted
                index += 1
        elif not quoted and char == "\\":
            if index + 1 >= len(source):
                raise OracleError("Value")
            output.append(source[index + 1])
            index += 2
        elif not quoted and char == "_":
            if index + 1 >= len(source):
                raise OracleError("Value")
            output.append(" ")
            index += 2
        elif not quoted and char == "@":
            if substitution is None:
                raise OracleError("Value")
            output.append(substitution)
            index += 1
        else:
            output.append(char)
            index += 1
    if quoted:
        raise OracleError("Value")
    return "".join(output)


def unquoted_positions(source: str, accepted: str = "") -> tuple[list[int], list[tuple[int, str]]]:
    positions: list[int] = []
    letters: list[tuple[int, str]] = []
    quoted = False
    index = 0
    while index < len(source):
        char = source[index]
        if char == '"':
            if quoted and index + 1 < len(source) and source[index + 1] == '"':
                index += 2
                continue
            quoted = not quoted
            index += 1
            continue
        if not quoted and char == "\\":
            if index + 1 >= len(source):
                raise OracleError("Value")
            index += 2
            continue
        if not quoted and char == "_":
            if index + 1 >= len(source):
                raise OracleError("Value")
            index += 2
            continue
        if not quoted:
            if char in accepted:
                positions.append(index)
            if char.isascii() and char.isalpha():
                letters.append((index, char))
        index += 1
    if quoted:
        raise OracleError("Value")
    return positions, letters


def has_unquoted_token(source: str, token: str) -> bool:
    quoted = False
    index = 0
    token_upper = token.upper()
    while index < len(source):
        char = source[index]
        if char == '"':
            if quoted and index + 1 < len(source) and source[index + 1] == '"':
                index += 2
                continue
            quoted = not quoted
            index += 1
            continue
        if quoted:
            index += 1
            continue
        if char == "\\" or char == "_":
            index += 2
            continue
        if source[index:index + len(token)].upper() == token_upper:
            return True
        index += 1
    return False


def date_is_minute(source: str, start: int, end: int) -> bool:
    before = next((char.lower() for char in reversed(source[:start]) if char.isascii() and char.isalpha()), None)
    after = next((char.lower() for char in source[end:] if char.isascii() and char.isalpha()), None)
    return before in {"h", "s"} or after in {"h", "s"}


def validate_date_format(format_code: str) -> None:
    quoted = False
    index = 0
    while index < len(format_code):
        char = format_code[index]
        if char == '"':
            if quoted and index + 1 < len(format_code) and format_code[index + 1] == '"':
                index += 2
                continue
            quoted = not quoted
            index += 1
            continue
        if quoted:
            index += 1
            continue
        if char == "\\" or char == "_":
            if index + 1 >= len(format_code):
                raise OracleError("Value")
            index += 2
            continue
        if format_code[index:index + 5].upper() == "AM/PM":
            index += 5
            continue
        if char in "yYmMdDhHsS":
            end = index + 1
            while end < len(format_code) and format_code[end].lower() == char.lower():
                end += 1
            length = end - index
            valid = (
                (char.lower() == "y" and length in {1, 2, 4})
                or (char.lower() in {"m", "d"} and 1 <= length <= 4)
                or (char.lower() in {"h", "s"} and 1 <= length <= 2)
            )
            if not valid:
                raise OracleError("Value")
            index = end
            continue
        if char not in "-/: ":
            raise OracleError("Value")
        index += 1
    if quoted:
        raise OracleError("Value")


def format_date_serial(number_value: float, format_code: str) -> str:
    if not math.isfinite(number_value):
        raise OracleError("Number")
    validate_date_format(format_code)
    whole = math.floor(number_value)
    fraction = Decimal(str(number_value)) - Decimal(whole)
    with localcontext() as context:
        context.prec = 80
        seconds = int((fraction * Decimal(86400)).quantize(Decimal(1), rounding=ROUND_HALF_UP))
    if seconds >= 86400:
        whole += 1
        seconds -= 86400
    try:
        current = date(1899, 12, 30) + timedelta(days=whole)
    except OverflowError as exc:
        raise OracleError("Number") from exc
    if current < date(1899, 1, 1) or current.year > 9999:
        raise OracleError("Number")
    hour, remainder = divmod(seconds, 3600)
    minute, second = divmod(remainder, 60)
    weekday_short = ("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    weekday_long = ("Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday")
    month_short = ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    month_long = ("January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    # Python's weekday is Monday=0; the formatter's table is Sunday=0.
    weekday = (current.weekday() + 1) % 7
    output: list[str] = []
    quoted = False
    index = 0
    while index < len(format_code):
        char = format_code[index]
        if char == '"':
            if quoted and index + 1 < len(format_code) and format_code[index + 1] == '"':
                output.append('"')
                index += 2
            else:
                quoted = not quoted
                index += 1
            continue
        if quoted:
            output.append(char)
            index += 1
            continue
        if char == "\\":
            if index + 1 >= len(format_code):
                raise OracleError("Value")
            output.append(format_code[index + 1])
            index += 2
            continue
        if char == "_":
            if index + 1 >= len(format_code):
                raise OracleError("Value")
            output.append(" ")
            index += 2
            continue
        if format_code[index:index + 5].upper() == "AM/PM":
            output.append("AM" if hour < 12 else "PM")
            index += 5
            continue
        if char not in "yYmMdDhHsS":
            output.append(char)
            index += 1
            continue
        end = index + 1
        while end < len(format_code) and format_code[end].lower() == char.lower():
            end += 1
        length = end - index
        lower = char.lower()
        if lower == "y":
            if length == 1:
                output.append(str(current.year))
            elif length == 2:
                output.append(f"{current.year % 100:02d}")
            elif length == 4:
                output.append(f"{current.year:04d}")
            else:
                raise OracleError("Value")
        elif lower == "m":
            if date_is_minute(format_code, index, end):
                output.append(str(minute) if length == 1 else f"{minute:02d}")
            elif length == 1:
                output.append(str(current.month))
            elif length == 2:
                output.append(f"{current.month:02d}")
            elif length == 3:
                output.append(month_short[current.month - 1])
            elif length == 4:
                output.append(month_long[current.month - 1])
            else:
                raise OracleError("Value")
        elif lower == "d":
            if length == 1:
                output.append(str(current.day))
            elif length == 2:
                output.append(f"{current.day:02d}")
            elif length == 3:
                output.append(weekday_short[weekday])
            elif length == 4:
                output.append(weekday_long[weekday])
            else:
                raise OracleError("Value")
        elif lower == "h":
            current_hour = hour % 12 or 12 if has_unquoted_token(format_code, "AM/PM") else hour
            output.append(str(current_hour) if length == 1 else f"{current_hour:02d}")
        elif lower == "s":
            output.append(str(second) if length == 1 else f"{second:02d}")
        index = end
    if quoted:
        raise OracleError("Value")
    return "".join(output)


def nearest_fraction(value: float, denominator_digits: int) -> tuple[int, int, int]:
    if denominator_digits < 1 or denominator_digits > 6:
        raise OracleError("Value")
    if not math.isfinite(value):
        raise OracleError("Number")
    represented = Fraction.from_float(abs(value))
    whole, remainder = divmod(represented.numerator, represented.denominator)
    fraction = Fraction(remainder, represented.denominator)
    # A d-placeholder denominator has at most d digits: 10^d itself is not
    # representable with d positions.
    limited = fraction.limit_denominator(10**denominator_digits - 1)
    numerator, denominator = limited.numerator, limited.denominator
    if numerator == denominator:
        whole += 1
        numerator = 0
        denominator = 1
    return whole, numerator, denominator


def render_numeric(value: float, pattern: str, negative: bool) -> str:
    positions, _ = unquoted_positions(pattern, "0#?,.%Ee/+- ")
    # The placeholder span is the numeric core; quoted/escaped literals are
    # retained as prefix and suffix by the caller.
    placeholder = [index for index in positions if pattern[index] in "0#?"]
    if not placeholder:
        raise OracleError("Value")
    numeric_positions = [
        index
        for index in positions
        if pattern[index] in "0#?.,%Ee/+-"
    ]
    first, last = min(numeric_positions), max(numeric_positions)
    core = pattern[first:last + 1]
    percent_count = core.count("%")
    percent = percent_count > 0
    percent_suffix = "%" * percent_count
    if any(char not in "0#?.,%Ee/+- " for char in core):
        raise OracleError("Value")
    if percent:
        core = core.replace("%", "")
    if core.count("/"):
        if core.count("/") != 1 or "." in core or "E" in core.upper():
            raise OracleError("Value")
        pre_slash, denominator_pattern = core.split("/", 1)
        denominator_digits = sum(char in "0#?" for char in denominator_pattern)
        if any(char not in "0#? /+-" for char in core):
            raise OracleError("Value")
        if denominator_digits < 1 or denominator_digits > 6:
            raise OracleError("Value")
        scaled_value = abs(value) * (100.0 if percent else 1.0)
        whole, numerator, denominator = nearest_fraction(scaled_value, denominator_digits)
        if " " in pre_slash:
            whole_pattern, numerator_pattern = pre_slash.rsplit(" ", 1)
            if not whole_pattern or not numerator_pattern:
                raise OracleError("Value")
            whole_text = render_slot_pattern(
                whole_pattern,
                str(whole) if whole else "0",
                omit_zero=whole == 0,
            )
            numerator_text = render_slot_pattern(
                numerator_pattern,
                str(numerator) if numerator else "0",
                omit_zero=numerator == 0,
            )
            result = whole_text
            if numerator:
                result += " " + numerator_text + "/" + render_slot_pattern(
                    denominator_pattern, str(denominator)
                )
        else:
            # With no separating space, the pre-slash slots hold the
            # improper numerator (whole*denominator + numerator).
            improper = whole * denominator + numerator
            result = ""
            if improper:
                result = render_slot_pattern(
                    pre_slash, str(improper), omit_zero=False
                )
                result += "/" + render_slot_pattern(
                    denominator_pattern, str(denominator)
                )
            elif "0" in pre_slash:
                result = render_slot_pattern(pre_slash, "0")
        if negative and result not in {"", "0"}:
            result = "-" + result
        return result + percent_suffix
    exponent_marker = next((char for char in "Ee" if char in core), None)
    if exponent_marker:
        mantissa_pattern, exponent_pattern = core.split(exponent_marker, 1)
        exponent_sign = "-" if "-" in exponent_pattern else "+" if "+" in exponent_pattern else ""
        exponent_tokens = [char for char in exponent_pattern if char in "0#?"]
        exponent_width = len(exponent_tokens)
        if exponent_width == 0:
            raise OracleError("Value")
        fraction_places = len(mantissa_pattern.rsplit(".", 1)[1]) if "." in mantissa_pattern else 0
        magnitude = Decimal(str(abs(value))) * (Decimal(100) if percent else Decimal(1))
        if magnitude == 0:
            power = 0
            mantissa = Decimal(0)
        else:
            power = magnitude.adjusted()
            mantissa = magnitude / (Decimal(10) ** power)
        with localcontext() as context:
            context.prec = 100
            quantum = Decimal(1).scaleb(-fraction_places)
            mantissa = mantissa.quantize(quantum, rounding=ROUND_HALF_UP)
        if mantissa >= 10:
            mantissa /= 10
            power += 1
        mantissa_text = render_numeric(float(mantissa), mantissa_pattern, negative and mantissa != 0)
        exponent_digits = str(abs(power))
        extra, assignments = placeholder_assignments(exponent_tokens, exponent_digits)
        exponent_text = extra + "".join(assignments)
        sign = "-" if power < 0 else "+" if exponent_sign == "+" else ""
        return mantissa_text + exponent_marker + sign + exponent_text + percent_suffix
    decimal_places = len(core.rsplit(".", 1)[1]) if "." in core else 0
    if core.count(".") > 1:
        raise OracleError("Value")
    integer_pattern = core.split(".", 1)[0]
    integer_required = integer_pattern.count("0")
    fraction_pattern = core.rsplit(".", 1)[1] if "." in core else ""
    if "," in fraction_pattern:
        raise OracleError("Value")
    if "," in integer_pattern:
        groups = integer_pattern.split(",")
        if (
            any(not group or len(group) > 3 for group in groups)
            or any(index > 0 and len(group) != 3 for index, group in enumerate(groups))
        ):
            raise OracleError("Value")
    fraction_required = fraction_pattern.count("0")
    fraction_total = sum(char in "0#?" for char in fraction_pattern)
    grouping = "," in integer_pattern
    decimal_value = Decimal(str(abs(value))) * (Decimal(100) if percent else Decimal(1))
    with localcontext() as context:
        context.prec = max(100, abs(decimal_value.adjusted()) + decimal_places + 24)
        rounded = decimal_value.quantize(Decimal(1).scaleb(-decimal_places), rounding=ROUND_HALF_UP)
    plain = format(rounded, f".{decimal_places}f")
    integer, _, fraction = plain.partition(".")
    integer_is_zero_optional = integer == "0" and integer_required == 0
    integer_digits = "" if integer_is_zero_optional else integer
    integer_text = render_slot_pattern(integer_pattern, integer_digits)
    integer_text = integer_text.replace(",", "")
    if grouping and integer_text.strip():
        integer_text = grouped(integer_text)
    if fraction_total:
        fraction_digits = fraction.ljust(fraction_total, "0")
        last_nonzero = max(
            (index + 1 for index, digit in enumerate(fraction_digits) if digit != "0"),
            default=0,
        )
        fraction_chars: list[str] = []
        token_index = 0
        for token in fraction_pattern:
            if token not in "0#?":
                continue
            if token_index < last_nonzero:
                fraction_chars.append(fraction_digits[token_index])
            elif token == "0":
                fraction_chars.append("0")
            elif token == "?":
                fraction_chars.append(" ")
            token_index += 1
        fraction = "".join(fraction_chars)
    else:
        fraction = ""
    result = integer_text
    decimal_required = fraction_required > 0 or "?" in fraction_pattern
    if fraction or decimal_required:
        result += "." + fraction
    if negative and rounded != 0:
        result = "-" + result
    return result + percent_suffix


def format_section(value: dict[str, object], section: str, *, auto_sign: bool) -> str:
    positions, letters = unquoted_positions(section, "0#?@.%Ee/+- ")
    placeholder = [index for index in positions if section[index] in "0#?"]
    at_positions = [index for index in positions if section[index] == "@"]
    date_letters = [char.lower() for _, char in letters if char.lower() in "ymdhs"]
    has_ampm = has_unquoted_token(section, "AM/PM")
    ampm_ranges = {
        index
        for index in range(max(0, len(section) - 4))
        if section[index:index + 5].upper() == "AM/PM"
    }
    unknown_letters = [
        (index, char)
        for index, char in letters
        if index not in {position + offset for position in ampm_ranges for offset in range(5)}
        and char.lower() not in "ymdhs"
        and not (placeholder and char in "Ee")
    ]
    if unknown_letters and not at_positions:
        raise OracleError("Value")
    if at_positions:
        if placeholder or date_letters or has_ampm:
            raise OracleError("Value")
        substitution = (
            str(value["value"])
            if value.get("kind") == "Text"
            else to_text(value)
            if value.get("kind") == "Number"
            else "TRUE"
            if value.get("kind") == "Logical" and value.get("value")
            else "FALSE"
            if value.get("kind") == "Logical"
            else None
        )
        return format_literal(section, substitution)
    if value.get("kind") == "Logical" and (placeholder or date_letters or has_ampm):
        raise OracleError("Value")
    if not placeholder and not date_letters and not has_ampm:
        return format_literal(section)
    if date_letters or has_ampm:
        if value.get("kind") != "Number":
            raise OracleError("Value")
        return format_date_serial(decode_number(value), section)
    if value.get("kind") not in {"Number", "Logical"}:
        raise OracleError("Value")
    number_value = decode_number(value) if value.get("kind") == "Number" else (1.0 if value.get("value") else 0.0)
    first = min(placeholder)
    # Percent, exponent punctuation, and fraction punctuation belong to the
    # numeric program even when they occur just outside the last placeholder.
    # Quoted/escaped characters were omitted by ``unquoted_positions``.
    numeric_positions = [
        index
        for index in positions
        if section[index] in "0#?.,%Ee/+-"
    ]
    last = max(numeric_positions) + 1
    first = min(numeric_positions)
    # Prefix/suffix are rendered separately so quoted and escaped literals do
    # not participate in numeric token parsing.
    return format_literal(section[:first]) + render_numeric(number_value, section[first:last], auto_sign and number_value < 0) + format_literal(section[last:])


def format_text(value: dict[str, object], format_code: str) -> str:
    propagate(value)
    if value.get("kind") in {"Empty", "Missing", "Complex"}:
        raise OracleError("Value")
    sections = split_format_sections(format_code)
    kind = value.get("kind")
    if kind == "Text" or kind == "Logical":
        section = sections[3] if len(sections) == 4 else sections[0]
        return format_section(value, section, auto_sign=False)
    if kind != "Number":
        raise OracleError("Value")
    number_value = decode_number(value)
    if not math.isfinite(number_value):
        raise OracleError("Number")
    if len(sections) == 1:
        index = 0
    elif number_value < 0:
        index = 1
    elif number_value == 0 and len(sections) >= 3:
        index = 2
    else:
        index = 0
    return format_section(value, sections[index], auto_sign=len(sections) == 1)


def apply(function: str, args: list[dict[str, object]]) -> dict[str, object]:
    name = function.upper()
    try:
        if name == "ASC":
            return text(asc(to_text(args[0])))
        if name == "CHAR":
            value = to_number(args[0])
            integer = math.trunc(value)
            if not (1 <= integer <= 255):
                raise OracleError("Value")
            return text(chr(integer))
        if name == "CLEAN":
            return text(clean(to_text(args[0])))
        if name == "CODE":
            value = to_text(args[0])
            if not value:
                raise OracleError("Value")
            return number(float(ord(value[0])))
        if name == "CONCATENATE":
            return text("".join(to_text(arg) for arg in args))
        if name == "DOLLAR":
            decimals = to_integer(args[1]) if len(args) > 1 else 2
            if decimals < -100 or decimals > 100:
                raise OracleError("Value")
            return text(currency(to_number(args[0]), decimals))
        if name == "EXACT":
            return logical(to_text(args[0]) == to_text(args[1]))
        if name == "FIND":
            start = to_integer(args[2]) if len(args) > 2 else 1
            return number(float(find_positions(to_text(args[0]), to_text(args[1]), start, insensitive=False)))
        if name == "FIXED":
            decimals = to_integer(args[1]) if len(args) > 1 else 2
            omit = bool(args[2]["value"]) if len(args) > 2 and args[2].get("kind") == "Logical" else (to_number(args[2]) != 0 if len(args) > 2 else False)
            if decimals < -100 or decimals > 100:
                raise OracleError("Value")
            return text(fixed(to_number(args[0]), decimals, omit))
        if name == "JIS":
            return text(jis(to_text(args[0])))
        if name == "LEFT":
            value = to_text(args[0]); count = to_integer_floor(args[1]) if len(args) > 1 else 1
            if count < 0: raise OracleError("Value")
            return text(value[:count])
        if name == "LEN":
            return number(float(scalar_count(to_text(args[0]))))
        if name == "LOWER":
            return text(lower(to_text(args[0])))
        if name == "MID":
            value = to_text(args[0]); start = to_integer_floor(args[1]); count = to_integer_floor(args[2])
            if start < 1 or count < 0: raise OracleError("Value")
            return text(value[start - 1 : start - 1 + count])
        if name == "PROPER":
            return text(proper(to_text(args[0])))
        if name == "REPLACE":
            value = to_text(args[0])
            raw_start = to_number(args[1]); raw_count = to_number(args[2])
            if raw_start < 1 or raw_count < 0:
                raise OracleError("Value")
            start = math.trunc(raw_start); count = math.trunc(raw_count); new = to_text(args[3])
            if start < 1 or count < 0: raise OracleError("Value")
            # The selected profile follows the displayed LEFT/MID equation:
            # Start = Len(T)+1 appends, while a larger Start is equivalent to
            # that append position.  Count is clamped to the suffix length.
            start_index = min(start - 1, len(value))
            count = min(count, len(value) - start_index)
            return text(value[:start_index] + new + value[start_index + count :])
        if name == "REPT":
            value = to_text(args[0]); count = to_integer(args[1])
            if count < 0: raise OracleError("Value")
            return text(value * count)
        if name == "RIGHT":
            value = to_text(args[0]); count = to_integer_floor(args[1]) if len(args) > 1 else 1
            if count < 0: raise OracleError("Value")
            return text(value[-count:] if count else "")
        if name == "SEARCH":
            start = to_integer(args[2]) if len(args) > 2 else 1
            return number(float(find_positions(to_text(args[0]), to_text(args[1]), start, insensitive=True)))
        if name == "SUBSTITUTE":
            value = to_text(args[0]); old = to_text(args[1]); new = to_text(args[2])
            if old == "": return text(value)
            if len(args) == 3: return text(value.replace(old, new))
            which = to_integer(args[3])
            if which < 1: raise OracleError("Value")
            pieces = value.split(old)
            if which >= len(pieces): return text(value)
            return text(old.join(pieces[:which]) + new + old.join(pieces[which:]))
        if name == "T":
            propagate(args[0])
            return text(str(args[0]["value"])) if args[0].get("kind") == "Text" else text("")
        if name == "TEXT":
            return text(format_text(args[0], to_text(args[1])))
        if name == "TRIM":
            return text(trim(to_text(args[0])))
        if name == "UNICHAR":
            value = to_integer(args[0])
            if not (0 <= value <= 0x10FFFF) or 0xD800 <= value <= 0xDFFF:
                raise OracleError("Value")
            return text(chr(value))
        if name == "UNICODE":
            value = to_text(args[0])
            if not value: raise OracleError("Value")
            return number(float(ord(value[0])))
        if name == "UPPER":
            return text(upper(to_text(args[0])))
        raise ValueError(f"unknown function {function}")
    except OracleError as exc:
        return error(exc.kind)


def quote(value: str) -> str:
    return '"' + value.replace('"', '""') + '"'


def obs(name: str, formula: str, args: list[dict[str, object]], *, cells: dict[str, dict[str, object]] | None = None, expected: dict[str, object] | None = None, reads: int = 0, tag: str = "direct") -> dict[str, object]:
    result = expected if expected is not None else apply(name, args)
    return {
        "id": f"{name.lower()}-{len(_OBSERVATIONS):03d}",
        "function": name,
        "formula": formula,
        "args": args,
        "cells": cells or {},
        "expected": result,
        "expected_reads": reads,
        "origin": tag,
    }


_OBSERVATIONS: list[dict[str, object]] = []


def build_observations() -> list[dict[str, object]]:
    global _OBSERVATIONS
    _OBSERVATIONS = []
    a = text("ＡＢＣ ｶﾞﾊﾟ\\￥")
    _OBSERVATIONS.append(obs("ASC", f"=ASC({quote(str(a['value']))})", [a]))
    _OBSERVATIONS.append(obs("ASC", '=ASC("ｶﾞﾊﾟ")', [text("ｶﾞﾊﾟ")]))
    _OBSERVATIONS.append(obs("ASC", '=ASC("\u2015‘’”「」")', [text("\u2015‘’”「」")]))
    _OBSERVATIONS.append(obs("ASC", '=ASC("Aあ😀")', [text("Aあ😀")]))

    for value in [1.0, 127.0, 255.0, 65.9]:
        formula = f"=CHAR({value:g})"
        _OBSERVATIONS.append(obs("CHAR", formula, [number(value)]))
    _OBSERVATIONS.append(obs("CHAR", "=CHAR(0)", [number(0.0)]))
    _OBSERVATIONS.append(obs("CHAR", "=CHAR(256)", [number(256.0)]))

    # Embedded controls are supplied through a cell so the formula parser is
    # not asked to accept a literal NUL.  The expected result still records
    # the complete Unicode scalar input independently.
    _OBSERVATIONS.append(obs("CLEAN", '=CLEAN([.A1])', [text("a\x00b\x1fc")], cells={"A1": text("a\x00b\x1fc")}, reads=1, tag="reference"))
    _OBSERVATIONS.append(obs("CLEAN", '=CLEAN("a\u0378b\u200bc\u00a0d")', [text("a\u0378b\u200bc\u00a0d")]))
    _OBSERVATIONS.append(obs("CLEAN", '=CLEAN("plain text")', [text("plain text")]))

    _OBSERVATIONS.append(obs("CODE", '=CODE("A😀")', [text("A😀")]))
    _OBSERVATIONS.append(obs("CODE", '=CODE("😀")', [text("😀")]))
    _OBSERVATIONS.append(obs("CODE", '=CODE("")', [text("")], expected=error("Value")))
    _OBSERVATIONS.append(obs("CODE", '=CODE([.A1])', [text("é")], cells={"A1": text("é")}, reads=1, tag="reference"))

    _OBSERVATIONS.append(obs("CONCATENATE", '=CONCATENATE("a";"β";2;TRUE())', [text("a"), text("β"), number(2.0), logical(True)]))
    _OBSERVATIONS.append(obs("CONCATENATE", '=CONCATENATE([.A1];"!")', [text("hello"), text("!")], cells={"A1": text("hello")}, reads=1, tag="reference"))
    _OBSERVATIONS.append(obs("CONCATENATE", '=CONCATENATE("x";#N/A;"y")', [text("x"), error("NotAvailable"), text("y")]))
    _OBSERVATIONS.append(obs("CONCATENATE", "=CONCATENATE(COMPLEX(1;2))", [complex_value(1.0, 2.0)], expected=error("Value"), tag="complex_type"))
    _OBSERVATIONS.append(obs("CONCATENATE", "=CONCATENATE()", [], expected=error("Value")))

    _OBSERVATIONS.append(obs("DOLLAR", "=DOLLAR(255)", [number(255.0)]))
    _OBSERVATIONS.append(obs("DOLLAR", "=DOLLAR(1234.567;2)", [number(1234.567), number(2.0)]))
    _OBSERVATIONS.append(obs("DOLLAR", "=DOLLAR(-1234.567;-2)", [number(-1234.567), number(-2.0)]))
    _OBSERVATIONS.append(obs("DOLLAR", "=DOLLAR(12.5;0)", [number(12.5), number(0.0)]))
    _OBSERVATIONS.append(obs("DOLLAR", "=DOLLAR(999.5;0)", [number(999.5), number(0.0)]))
    _OBSERVATIONS.append(obs("DOLLAR", "=DOLLAR(-0.5;0)", [number(-0.5), number(0.0)]))
    _OBSERVATIONS.append(obs("DOLLAR", "=DOLLAR(-0.0001;2)", [number(-0.0001), number(2.0)]))
    _OBSERVATIONS.append(obs("DOLLAR", "=DOLLAR(1e100;2)", [number(1e100), number(2.0)]))

    _OBSERVATIONS.append(obs("EXACT", '=EXACT("Text";"Text")', [text("Text"), text("Text")]))
    _OBSERVATIONS.append(obs("EXACT", '=EXACT("Text";"text")', [text("Text"), text("text")]))
    _OBSERVATIONS.append(obs("EXACT", '=EXACT("é";"e\u0301")', [text("é"), text("e\u0301")]))
    _OBSERVATIONS.append(obs("EXACT", '=EXACT("x";#N/A)', [text("x"), error("NotAvailable")]))

    _OBSERVATIONS.append(obs("FIND", '=FIND("na";"banana")', [text("na"), text("banana")]))
    _OBSERVATIONS.append(obs("FIND", '=FIND("NA";"banana")', [text("NA"), text("banana")], expected=error("Value")))
    _OBSERVATIONS.append(obs("FIND", '=FIND("a";"banana";4)', [text("a"), text("banana"), number(4.0)]))
    _OBSERVATIONS.append(obs("FIND", '=FIND("";"abc";2)', [text(""), text("abc"), number(2.0)]))

    _OBSERVATIONS.append(obs("FIXED", "=FIXED(1234.567)", [number(1234.567)]))
    _OBSERVATIONS.append(obs("FIXED", "=FIXED(1234.567;1)", [number(1234.567), number(1.0)]))
    _OBSERVATIONS.append(obs("FIXED", "=FIXED(1234.567;2;TRUE())", [number(1234.567), number(2.0), logical(True)]))
    _OBSERVATIONS.append(obs("FIXED", "=FIXED(-1234.567;-2)", [number(-1234.567), number(-2.0)]))
    _OBSERVATIONS.append(obs("FIXED", "=FIXED(999.5;0)", [number(999.5), number(0.0)]))
    _OBSERVATIONS.append(obs("FIXED", "=FIXED(-0.5;0)", [number(-0.5), number(0.0)]))
    _OBSERVATIONS.append(obs("FIXED", "=FIXED(-0.0001;2)", [number(-0.0001), number(2.0)]))
    _OBSERVATIONS.append(obs("FIXED", "=FIXED(1e100;2;TRUE())", [number(1e100), number(2.0), logical(True)]))

    _OBSERVATIONS.append(obs("JIS", '=JIS("ABC\\`\'")', [text("ABC\\`'")]))
    _OBSERVATIONS.append(obs("JIS", '=JIS("ｶﾞﾊﾟ")', [text("ｶﾞﾊﾟ")]))
    _OBSERVATIONS.append(obs("JIS", '=JIS("ｶﾟ")', [text("ｶﾟ")]))
    _OBSERVATIONS.append(obs("JIS", '=JIS("Aあ😀")', [text("Aあ😀")]))

    _OBSERVATIONS.append(obs("LEFT", '=LEFT("abcdef")', [text("abcdef")]))
    _OBSERVATIONS.append(obs("LEFT", '=LEFT("α😀β";2)', [text("α😀β"), number(2.0)]))
    _OBSERVATIONS.append(obs("LEFT", '=LEFT("abc";1.9)', [text("abc"), number(1.9)]))
    _OBSERVATIONS.append(obs("LEFT", '=LEFT("abc";0)', [text("abc"), number(0.0)]))
    _OBSERVATIONS.append(obs("LEFT", '=LEFT("abc";-0.9)', [text("abc"), number(-0.9)], expected=error("Value")))
    _OBSERVATIONS.append(obs("LEFT", '=LEFT("abc";-1)', [text("abc"), number(-1.0)], expected=error("Value")))

    _OBSERVATIONS.append(obs("LEN", '=LEN("α😀e\u0301")', [text("α😀e\u0301")]))
    _OBSERVATIONS.append(obs("LEN", "=LEN(1.5)", [number(1.5)]))
    _OBSERVATIONS.append(obs("LEN", '=LEN([.A1])', [text("hello")], cells={"A1": text("hello")}, reads=1, tag="reference"))
    _OBSERVATIONS.append(obs("LEN", '=LEN("")', [text("")]))

    _OBSERVATIONS.append(obs("LOWER", '=LOWER("ABC Greek Σ")', [text("ABC Greek Σ")]))
    _OBSERVATIONS.append(obs("LOWER", '=LOWER("ΟΣ")', [text("ΟΣ")]))
    _OBSERVATIONS.append(obs("LOWER", '=LOWER("İ ß")', [text("İ ß")]))
    _OBSERVATIONS.append(obs("LOWER", '=LOWER("\uA7CF")', [text("\uA7CF")]))

    _OBSERVATIONS.append(obs("MID", '=MID("abcdef";2;3)', [text("abcdef"), number(2.0), number(3.0)]))
    _OBSERVATIONS.append(obs("MID", '=MID("α😀β";2;1)', [text("α😀β"), number(2.0), number(1.0)]))
    _OBSERVATIONS.append(obs("MID", '=MID("abc";1.9;1.9)', [text("abc"), number(1.9), number(1.9)]))
    _OBSERVATIONS.append(obs("MID", '=MID("abc";9;2)', [text("abc"), number(9.0), number(2.0)]))
    _OBSERVATIONS.append(obs("MID", '=MID("abc";0;2)', [text("abc"), number(0.0), number(2.0)], expected=error("Value")))

    _OBSERVATIONS.append(obs("PROPER", '=PROPER("hELLO, wORLD!")', [text("hELLO, wORLD!")]))
    _OBSERVATIONS.append(obs("PROPER", '=PROPER("élan déjà")', [text("élan déjà")]))
    _OBSERVATIONS.append(obs("PROPER", '=PROPER("a\u0301BC")', [text("a\u0301BC")]))
    _OBSERVATIONS.append(obs("PROPER", '=PROPER("\uA7CFoo")', [text("\uA7CFoo")]))

    _OBSERVATIONS.append(obs("REPLACE", '=REPLACE("abcdef";2;3;"X")', [text("abcdef"), number(2.0), number(3.0), text("X")]))
    _OBSERVATIONS.append(obs("REPLACE", '=REPLACE("abc";2;0;"X")', [text("abc"), number(2.0), number(0.0), text("X")]))
    _OBSERVATIONS.append(obs("REPLACE", '=REPLACE("abc";99;2;"X")', [text("abc"), number(99.0), number(2.0), text("X")]))
    _OBSERVATIONS.append(obs("REPLACE", '=REPLACE("abc";0;1;"X")', [text("abc"), number(0.0), number(1.0), text("X")], expected=error("Value")))

    _OBSERVATIONS.append(obs("REPT", '=REPT("ab";3)', [text("ab"), number(3.0)]))
    _OBSERVATIONS.append(obs("REPT", '=REPT("ab";0)', [text("ab"), number(0.0)]))
    _OBSERVATIONS.append(obs("REPT", '=REPT("😀";2)', [text("😀"), number(2.0)]))
    _OBSERVATIONS.append(obs("REPT", '=REPT("ab";-1)', [text("ab"), number(-1.0)], expected=error("Value")))

    _OBSERVATIONS.append(obs("RIGHT", '=RIGHT("abcdef")', [text("abcdef")]))
    _OBSERVATIONS.append(obs("RIGHT", '=RIGHT("α😀β";2)', [text("α😀β"), number(2.0)]))
    _OBSERVATIONS.append(obs("RIGHT", '=RIGHT("abc";1.9)', [text("abc"), number(1.9)]))
    _OBSERVATIONS.append(obs("RIGHT", '=RIGHT("abc";0)', [text("abc"), number(0.0)]))
    _OBSERVATIONS.append(obs("RIGHT", '=RIGHT("abc";-1)', [text("abc"), number(-1.0)], expected=error("Value")))

    _OBSERVATIONS.append(obs("SEARCH", '=SEARCH("NA";"banana")', [text("NA"), text("banana")]))
    _OBSERVATIONS.append(obs("SEARCH", '=SEARCH("ss";"aßb")', [text("ss"), text("aßb")]))
    _OBSERVATIONS.append(obs("SEARCH", '=SEARCH("s";"aßb")', [text("s"), text("aßb")], expected=error("Value")))
    _OBSERVATIONS.append(obs("SEARCH", '=SEARCH("ß";"ss")', [text("ß"), text("ss")]))
    _OBSERVATIONS.append(obs("SEARCH", '=SEARCH("s";"ßs")', [text("s"), text("ßs")]))
    _OBSERVATIONS.append(obs("SEARCH", '=SEARCH("ss";"sß")', [text("ss"), text("sß")]))
    _OBSERVATIONS.append(obs("SEARCH", '=SEARCH("😀";"a😀b")', [text("😀"), text("a😀b")]))
    _OBSERVATIONS.append(obs("SEARCH", '=SEARCH("x";"abc")', [text("x"), text("abc")], expected=error("Value")))
    _OBSERVATIONS.append(obs("SEARCH", '=SEARCH("";"abc";4)', [text(""), text("abc"), number(4.0)]))

    _OBSERVATIONS.append(obs("SUBSTITUTE", '=SUBSTITUTE("a-b-a";"a";"X")', [text("a-b-a"), text("a"), text("X")]))
    _OBSERVATIONS.append(obs("SUBSTITUTE", '=SUBSTITUTE("a-b-a";"a";"X";2)', [text("a-b-a"), text("a"), text("X"), number(2.0)]))
    _OBSERVATIONS.append(obs("SUBSTITUTE", '=SUBSTITUTE("abc";"";"X")', [text("abc"), text(""), text("X")]))
    _OBSERVATIONS.append(obs("SUBSTITUTE", '=SUBSTITUTE("abc";"z";"X")', [text("abc"), text("z"), text("X")]))
    _OBSERVATIONS.append(obs("SUBSTITUTE", '=SUBSTITUTE("abc";"a";"X";0)', [text("abc"), text("a"), text("X"), number(0.0)], expected=error("Value")))

    _OBSERVATIONS.append(obs("T", '=T("text")', [text("text")]))
    _OBSERVATIONS.append(obs("T", "=T(5)", [number(5.0)]))
    _OBSERVATIONS.append(obs("T", "=T(TRUE())", [logical(True)]))
    _OBSERVATIONS.append(obs("T", "=T(COMPLEX(1;2))", [complex_value(1.0, 2.0)]))
    _OBSERVATIONS.append(obs("T", '=T(#N/A)', [error("NotAvailable")]))

    _OBSERVATIONS.append(obs("TEXT", '=TEXT(12.34567;"###.##")', [number(12.34567), text("###.##")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(12.34567;"000.00")', [number(12.34567), text("000.00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.125;"0%")', [number(0.125), text("0%")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT("raw";"@")', [text("raw"), text("@")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(999.5;"0")', [number(999.5), text("0")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(-0.5;"0")', [number(-0.5), text("0")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1.25;"0.0")', [number(1.25), text("0.0")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1.2;"???.??")', [number(1.2), text("???.??")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"???")', [number(1.0), text("???")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1.2;"0.??")', [number(1.2), text("0.??")]))
    # Mixed placeholders retain each positional slot: required zeroes do not
    # move ahead of an absent question mark, and exponent slots follow the
    # same right-aligned rule as integer slots.
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0?0")', [number(1.0), text("0?0")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"#?0")', [number(1.0), text("#?0")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0.?0")', [number(1.0), text("0.?0")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.1;"0.0E#?")', [number(0.1), text("0.0E#?")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0.##")', [number(1.0), text("0.##")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1234567.89;"#,##0.00")', [number(1234567.89), text("#,##0.00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1e100;"0")', [number(1e100), text("0")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1.999;"0.00")', [number(1.999), text("0.00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.5;"#/#")', [number(0.5), text("#/#")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1.25;"#/#")', [number(1.25), text("#/#")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1.25;"# ?/?")', [number(1.25), text("# ?/?")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.5;"# ??/??")', [number(0.5), text("# ??/??")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.333333;"#/#")', [number(0.333333), text("#/#")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.1;"# ?/?")', [number(0.1), text("# ?/?")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.09;"# ?/?")', [number(0.09), text("# ?/?")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.999;"# ?/?")', [number(0.999), text("# ?/?")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(-1.25;"# ?/?")', [number(-1.25), text("# ?/?")]))
    # These values sit near exact midpoint boundaries.  The reference uses
    # represented-binary64 Fraction distance, so a float-distance scan must
    # not silently choose a different denominator at the strict cap.
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.11805555555555555;"# ?/?")', [number(0.11805555555555555), text("# ?/?")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.13392857142857142;"# ?/?")', [number(0.13392857142857142), text("# ?/?")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.2361111111111111;"# ?/?")', [number(0.2361111111111111), text("# ?/?")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.005050505050505051;"# ??/??")', [number(0.005050505050505051), text("# ??/??")]))
    _OBSERVATIONS.append(obs(
        "TEXT",
        '=TEXT(18446744073709551616;"#/#")',
        [number(18446744073709551616.0), text("#/#")],
        expected=error("Number"),
    ))
    _OBSERVATIONS.append(obs(
        "TEXT",
        '=TEXT(18446744073709549568;"#/#")',
        [number(18446744073709549568.0), text("#/#")],
    ))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0/???????")', [number(1.0), text("0/???????")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0/")', [number(1.0), text("0/")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(12300;"0.00E+00")', [number(12300.0), text("0.00E+00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0.0123;"0.0E+00")', [number(0.0123), text("0.0E+00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(9999;"0.0E+00")', [number(9999.0), text("0.0E+00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(123;"00.0E+00")', [number(123.0), text("00.0E+00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0;"0.0E+00")', [number(0.0), text("0.0E+00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(9.9995;"0.00E+00")', [number(9.9995), text("0.00E+00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(5e-324;"0.00E+00")', [number(5e-324), text("0.00E+00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(2.2250738585072014e-308;"0.00E+00")', [number(2.2250738585072014e-308), text("0.00E+00")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1.7976931348623157e308;"0.00E+00")', [number(1.7976931348623157e308), text("0.00E+00")]))

    section_format = '0.0;"NEG"0.0;"ZERO";"TEXT:"@'
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(12.3;{quote(section_format)})", [number(12.3), text(section_format)]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(-12.3;{quote(section_format)})", [number(-12.3), text(section_format)]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(0;{quote(section_format)})", [number(0.0), text(section_format)]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(\"raw\";{quote(section_format)})", [text("raw"), text(section_format)]))
    numeric_text_sections = '@;@;@;"text:"@'
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(12;{quote(numeric_text_sections)})", [number(12.0), text(numeric_text_sections)]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(-12;{quote(numeric_text_sections)})", [number(-12.0), text(numeric_text_sections)]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(0;{quote(numeric_text_sections)})", [number(0.0), text(numeric_text_sections)]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(TRUE();{quote('"L:"@')})", [logical(True), text('"L:"@')]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(TRUE();"0")', [logical(True), text("0")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT("raw";"0")', [text("raw"), text("0")], expected=error("Value")))

    date_format = "yyyy-mm-dd hh:mm:ss"
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(45292.25;{quote(date_format)})", [number(45292.25), text(date_format)]))
    weekday_format = 'ddd", "mmm d", "yyyy'
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(45292;{quote(weekday_format)})", [number(45292.0), text(weekday_format)]))
    clock_format = "h:mm AM/PM"
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(45292.5;{quote(clock_format)})", [number(45292.5), text(clock_format)]))
    literal_ampm_format = 'h "AM/PM"'
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(0.75;{quote(literal_ampm_format)})", [number(0.75), text(literal_ampm_format)]))
    escaped_ampm_format = r"h \A\M\/\P\M"
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(0.75;{quote(escaped_ampm_format)})", [number(0.75), text(escaped_ampm_format)]))
    underscore_date_format = "yyyy_)"
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(1;{quote(underscore_date_format)})", [number(1.0), text(underscore_date_format)]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(0;"yyyy")', [number(0.0), text("yyyy")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"yyyy")', [number(1.0), text("yyyy")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(59;"yyyy-mm-dd")', [number(59.0), text("yyyy-mm-dd")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(60;"yyyy-mm-dd")', [number(60.0), text("yyyy-mm-dd")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(61;"yyyy-mm-dd")', [number(61.0), text("yyyy-mm-dd")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1.9999999;"yyyy-mm-dd hh:mm:ss")', [number(1.9999999), text("yyyy-mm-dd hh:mm:ss")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(2958465;"yyyy-mm-dd")', [number(2958465.0), text("yyyy-mm-dd")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(-1;"yyyy")', [number(-1.0), text("yyyy")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(-1.25;"yyyy-mm-dd hh:mm")', [number(-1.25), text("yyyy-mm-dd hh:mm")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(-363;"yyyy-mm-dd")', [number(-363.0), text("yyyy-mm-dd")]))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(-364;"yyyy")', [number(-364.0), text("yyyy")], expected=error("Number")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(2958466;"yyyy")', [number(2958466.0), text("yyyy")], expected=error("Number")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"yyyyy")', [number(1.0), text("yyyyy")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"yyyy!mm")', [number(1.0), text("yyyy!mm")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0Q")', [number(1.0), text("0Q")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0E")', [number(1.0), text("0E")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0:0")', [number(1.0), text("0:0")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0,0")', [number(1.0), text("0,0")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0.0.0")', [number(1.0), text("0.0.0")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"0;0;0;0;0")', [number(1.0), text("0;0;0;0;0")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;"")', [number(1.0), text("")], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(12.3;{quote('"literal"')})", [number(12.3), text('"literal"')]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(12.3;{quote('"USD "#,##0.00')})", [number(12.3), text('"USD "#,##0.00')]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(12.3;{quote(r'\$0.00')})", [number(12.3), text(r'\$0.00')]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(12.3;{quote('_ 0.0')})", [number(12.3), text("_ 0.0")]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(12.3;{quote('"He said ""x"" "0.0')})", [number(12.3), text('"He said ""x"" "0.0')]))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(1;{quote('"unterminated')})", [number(1.0), text('"unterminated')], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", f"=TEXT(1;{quote('0\\')})", [number(1.0), text('0\\')], expected=error("Value")))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(COMPLEX(1;2);"@")', [complex_value(1.0, 2.0)], expected=error("Value"), tag="complex_type"))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(#N/A;"0")', [error("NotAvailable"), text("0")], expected=error("NotAvailable"), tag="formula_error"))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT(1;#N/A)', [number(1.0), error("NotAvailable")], expected=error("NotAvailable"), tag="formula_error"))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT([.A2];"@")', [empty()], cells={"A2": empty()}, expected=error("Value"), reads=1, tag="reference"))
    _OBSERVATIONS.append(obs("TEXT", '=TEXT( ;"@")', [missing()], expected=error("Value"), reads=0, tag="missing"))

    _OBSERVATIONS.append(obs("TRIM", '=TRIM([.A1])', [text("  a\t\n b  c ")], cells={"A1": text("  a\t\n b  c ")}, reads=1, tag="reference"))
    _OBSERVATIONS.append(obs("TRIM", '=TRIM("a\u00a0 b")', [text("a\u00a0 b")]))
    _OBSERVATIONS.append(obs("TRIM", '=TRIM("   ")', [text("   ")]))
    _OBSERVATIONS.append(obs("TRIM", '=TRIM([.A1])', [text(" a  b ")], cells={"A1": text(" a  b ")}, reads=1, tag="reference"))

    _OBSERVATIONS.append(obs("UNICHAR", "=UNICHAR(0)", [number(0.0)]))
    _OBSERVATIONS.append(obs("UNICHAR", "=UNICHAR(128512)", [number(float(0x1F600))]))
    _OBSERVATIONS.append(obs("UNICHAR", "=UNICHAR(55296)", [number(float(0xD800))], expected=error("Value")))
    _OBSERVATIONS.append(obs("UNICHAR", "=UNICHAR(1114112)", [number(float(0x110000))], expected=error("Value")))
    _OBSERVATIONS.append(obs("UNICHAR", "=UNICHAR(65.9)", [number(65.9)]))

    _OBSERVATIONS.append(obs("UNICODE", '=UNICODE("😀x")', [text("😀x")]))
    _OBSERVATIONS.append(obs("UNICODE", '=UNICODE("é")', [text("é")]))
    _OBSERVATIONS.append(obs("UNICODE", '=UNICODE("")', [text("")], expected=error("Value")))
    _OBSERVATIONS.append(obs("UNICODE", '=UNICODE([.A1])', [text("Ω")], cells={"A1": text("Ω")}, reads=1, tag="reference"))

    _OBSERVATIONS.append(obs("UPPER", '=UPPER("abc Greek σ")', [text("abc Greek σ")]))
    _OBSERVATIONS.append(obs("UPPER", '=UPPER("Straße")', [text("Straße")]))
    _OBSERVATIONS.append(obs("UPPER", '=UPPER("i\u0307")', [text("i\u0307")]))
    _OBSERVATIONS.append(obs("UPPER", '=UPPER("\uA7CE")', [text("\uA7CE")]))
    _OBSERVATIONS.append(obs("UPPER", '=UPPER([.A1])', [text("mixed")], cells={"A1": text("mixed")}, reads=1, tag="reference"))

    # Structural/type/error boundary rows are retained once, separately from
    # the function cardinality rows, so the Rust target can assert zero reads.
    _OBSERVATIONS.append(obs("LEFT", '=LEFT([.A1]~[.A2])', [text("ignored")], cells={"A1": text("a"), "A2": text("b")}, expected=error("Value"), reads=0, tag="reference_list_refusal"))
    _OBSERVATIONS.append(obs("UPPER", '=UPPER(;)', [missing()], expected=error("Value"), reads=0, tag="missing"))
    _OBSERVATIONS.append(obs("FIND", '=FIND("x";#N/A;1)', [text("x"), error("NotAvailable"), number(1.0)], expected=error("NotAvailable"), tag="formula_error"))

    assert {row["function"] for row in _OBSERVATIONS} == set(FUNCTIONS)
    return _OBSERVATIONS


def build_document() -> dict[str, object]:
    contract_hash = hashlib.sha256(CONTRACT.read_bytes()).hexdigest()
    if CONTRACT_SHA256 is not None and contract_hash != CONTRACT_SHA256:
        raise SystemExit(f"contract hash changed: {contract_hash} != {CONTRACT_SHA256}")
    observations = build_observations()
    counts = {name: sum(row["function"] == name for row in observations) for name in FUNCTIONS}
    return {
        "schema": "litchi-ods-text-oracle-v1",
        "contract_sha256": contract_hash,
        "normative_source": {
            "archive_sha256": ODF_ARCHIVE_SHA256,
            "member_sha256": ODF_MEMBER_SHA256,
            "sections": ["6.20.2", "6.20.3", "6.20.4", "6.20.5", "6.20.6", "6.20.7", "6.20.8", "6.20.9", "6.20.10", "6.20.11", "6.20.12", "6.20.13", "6.20.14", "6.20.15", "6.20.16", "6.20.17", "6.20.18", "6.20.19", "6.20.20", "6.20.21", "6.20.22", "6.20.23", "6.20.24", "6.20.25", "6.20.26", "6.20.27"],
        },
        "profile": {
            "unicode": UNICODE_PROFILE,
            "character_unit": "unicode_scalar_value",
            "byte_functions": "out_of_scope_for_this_26_function_oracle",
            "normalization": "none",
            "search": "literal; FIND sensitive; SEARCH Unicode-case-insensitive",
            "format_provider": "test-en-US-dollar-comma-period-half-away",
            "currency_symbol": "$",
            "decimal_separator": ".",
            "group_separator": ",",
            "default_currency_decimals": 2,
            "rounding": "half-away-from-zero",
            "comparison": "exact UTF-8 text / exact logical / exact binary64 bits",
        },
        "function_counts": counts,
        "observations": observations,
    }


def canonical_bytes(document: dict[str, object]) -> bytes:
    return (json.dumps(document, ensure_ascii=False, indent=2, sort_keys=False) + "\n").encode("utf-8")


def main() -> None:
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    document = build_document()
    payload = canonical_bytes(document)
    if args.check:
        observed = GOLDENS.read_bytes()
        if observed != payload:
            raise SystemExit("text-goldens.json differs from independently regenerated bytes")
        print(json.dumps({"functions": len(FUNCTIONS), "observations": len(document["observations"]), "verified": True}))
    elif args.write:
        GOLDENS.write_bytes(payload)
        print(json.dumps({"functions": len(FUNCTIONS), "observations": len(document["observations"]), "written": True}))
    else:
        print(json.dumps({"functions": len(FUNCTIONS), "observations": len(document["observations"]), "contract_sha256": document["contract_sha256"]}))


if __name__ == "__main__":
    main()
