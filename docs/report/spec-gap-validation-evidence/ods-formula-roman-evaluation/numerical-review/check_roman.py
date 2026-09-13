#!/usr/bin/env python3
"""Independent bounded checks for ODF 1.4 ROMAN/ARABIC conversions.

This file deliberately does not reproduce the implementation's residue
recursion.  It uses breadth first search over all signed Roman denominations
to obtain a lower bound for format 4, and a separate state search to check
that the canonical negative-then-positive arrangement can attain that bound.
The search is integer-only and has no dependency on Cargo or the Rust crate.
"""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Iterable
from zipfile import ZipFile


ROOT = Path(__file__).resolve().parents[5]
EVIDENCE = Path(__file__).resolve().parents[1]
SYMBOLS = "IVXLCDM"
COINS = (1, 5, 10, 50, 100, 500, 1000)
MAX_N = 3999
CLASSIC_DEPTH_BOUND = 15
SUM_BOUND = CLASSIC_DEPTH_BOUND * COINS[-1]


def sha256(path: Path) -> str | None:
    if not path.is_file():
        return None
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 16), b""):
            digest.update(block)
    return digest.hexdigest()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def arabic(text: str) -> int | None:
    """Apply the ODF indirect-subtraction rule to an ASCII Roman string."""

    values: list[int] = []
    for symbol in text:
        if "a" <= symbol <= "z":
            symbol = chr(ord(symbol) - (ord("a") - ord("A")))
        try:
            values.append(COINS[SYMBOLS.index(symbol)])
        except ValueError:
            return None

    total = 0
    right_max = 0
    for value in reversed(values):
        if value < right_max:
            total -= value
        else:
            total += value
            right_max = value
    return total


def signed_distances(max_depth: int) -> dict[int, int]:
    """Shortest sums using an unconstrained +/- denomination at each step."""

    frontier = {0}
    distances = {0: 0}
    # A Roman string of at most max_depth symbols has every signed prefix in
    # [-max_depth*1000, max_depth*1000].  Keeping that full interval means
    # this is an exhaustive lower-bound search, rather than a heuristic cap.
    bound = max_depth * COINS[-1]
    for depth in range(1, max_depth + 1):
        next_frontier: set[int] = set()
        for total in frontier:
            for coin in COINS:
                for candidate in (total - coin, total + coin):
                    if -bound <= candidate <= bound and candidate not in distances:
                        distances[candidate] = depth
                        next_frontier.add(candidate)
        frontier = next_frontier
    return distances


def _append_counts(
    counts: tuple[tuple[int, ...], tuple[int, ...]],
    index: int,
    negative: bool,
) -> tuple[tuple[int, ...], tuple[int, ...]]:
    negative_counts, positive_counts = counts
    selected = list(negative_counts if negative else positive_counts)
    selected[index] += 1
    if negative:
        return tuple(selected), positive_counts
    return negative_counts, tuple(selected)


def canonical_text(
    counts: tuple[tuple[int, ...], tuple[int, ...]],
) -> str:
    """Render signed counts in the arrangement used by the proof.

    Negative symbols ascend in denomination and positive symbols descend.  A
    negative symbol therefore has a larger symbol to its right whenever the
    highest negative denomination is below the highest positive one.
    """

    negative_counts, positive_counts = counts
    negative = "".join(
        SYMBOLS[index] * amount
        for index, amount in enumerate(negative_counts)
    )
    positive = "".join(
        SYMBOLS[index] * positive_counts[index]
        for index in reversed(range(len(SYMBOLS)))
    )
    return negative + positive


def canonical_distances(
    max_depth: int,
) -> tuple[dict[int, int], dict[int, str]]:
    """Shortest signed sums whose signs can be emitted canonically.

    The state records only the sum and the highest denomination used with
    each sign.  Counts are carried for one witness per state; they are not
    used to decide reachability.  The search is independent of the
    implementation's per-ratio residue choices.
    """

    start_counts = ((0,) * len(COINS), (0,) * len(COINS))
    frontier: dict[tuple[int, int, int], tuple[tuple[int, ...], tuple[int, ...]]] = {
        (0, -1, -1): start_counts
    }
    seen = set(frontier)
    distances: dict[int, int] = {0: 0}
    witnesses: dict[int, str] = {0: ""}
    bound = max_depth * COINS[-1]

    for depth in range(1, max_depth + 1):
        next_frontier: dict[
            tuple[int, int, int], tuple[tuple[int, ...], tuple[int, ...]]
        ] = {}
        for (total, highest_negative, highest_positive), counts in frontier.items():
            for index, coin in enumerate(COINS):
                positive_state = (
                    total + coin,
                    highest_negative,
                    max(highest_positive, index),
                )
                if -bound <= positive_state[0] <= bound and positive_state not in seen:
                    next_frontier[positive_state] = _append_counts(counts, index, False)

                negative_state = (
                    total - coin,
                    max(highest_negative, index),
                    highest_positive,
                )
                if -bound <= negative_state[0] <= bound and negative_state not in seen:
                    next_frontier[negative_state] = _append_counts(counts, index, True)

        seen.update(next_frontier)
        frontier = next_frontier
        for (total, highest_negative, highest_positive), counts in frontier.items():
            # No negative symbol is already valid.  Otherwise the largest
            # negative must see a strictly larger positive symbol to its
            # right; all lower negatives then see that symbol as well.
            valid = highest_negative < 0 or highest_negative < highest_positive
            if 0 <= total <= MAX_N and valid and total not in distances:
                text = canonical_text(counts)
                if arabic(text) != total:
                    raise AssertionError((total, text, arabic(text)))
                distances[total] = depth
                witnesses[total] = text

    return distances, witnesses


def _direct_pair_allowed(format_number: int, left: int, right: int) -> bool:
    """Check the table's direct-pair restrictions for a greedy output."""

    if left >= right:
        return False
    # Powers of ten are I, X, and C.  Format 1 and 3 add V and L; format 2
    # adds L only.  Formats 0 and 1 additionally cap the ratio at ten.
    powers = left in (1, 10, 100)
    extra = (format_number in (1, 3) and left in (5, 50)) or (
        format_number == 2 and left == 50
    )
    if not (powers or extra):
        return False
    return format_number >= 2 or right <= 10 * left


def _subtractor_allowed(index: int, format_number: int) -> bool:
    return {
        0: {0, 2, 4},
        1: set(range(5)),
        2: {0, 2, 3, 4},
        3: set(range(5)),
    }[format_number].__contains__(index)


def source_like_simplified(number: int) -> str:
    """Transcribe the candidate's bounded residue construction for comparison.

    The proof of minimality does not use this function: its independent BFS
    lower bound is computed separately above.  Keeping this small transcription
    lets the receipt show that the implementation-shaped witness reaches that
    bound for the complete requested domain.
    """

    best_length: int | None = None
    best: list[int] | None = None
    for mask in range(1 << (len(COINS) - 1)):
        units = number
        coefficients: list[int] = []
        length = 0
        for index, ratio in enumerate((5, 2, 5, 2, 5, 2)):
            remainder = units % ratio
            coefficient = (
                remainder - ratio
                if remainder and (mask & (1 << index))
                else remainder
            )
            coefficients.append(coefficient)
            length += abs(coefficient)
            units = (units - coefficient) // ratio
        coefficients.append(units)
        length += units
        if best_length is None or length < best_length:
            best_length = length
            best = coefficients
    assert best is not None
    negative = "".join(
        SYMBOLS[index] * (-coefficient)
        for index, coefficient in enumerate(best)
        if coefficient < 0
    )
    positive = "".join(
        SYMBOLS[index] * best[index]
        for index in reversed(range(len(COINS)))
        if best[index] > 0
    )
    return negative + positive


def source_like_greedy(number: int, format_number: int) -> str:
    """Transcribe the finite single/pair greedy table for formats 0..3."""

    output: list[str] = []
    while number:
        best_value = 0
        best_first = 0
        best_second: int | None = None
        for index in reversed(range(len(COINS))):
            value = COINS[index]
            if value <= number and value > best_value:
                best_value = value
                best_first = index
                best_second = None
        for smaller in range(len(COINS) - 1):
            if not _subtractor_allowed(smaller, format_number):
                continue
            for larger in range(smaller + 1, len(COINS)):
                if format_number <= 1 and COINS[larger] > COINS[smaller] * 10:
                    continue
                value = COINS[larger] - COINS[smaller]
                if value <= number and value > best_value:
                    best_value = value
                    best_first = smaller
                    best_second = larger
        assert best_value > 0
        output.append(SYMBOLS[best_first])
        if best_second is not None:
            output.append(SYMBOLS[best_second])
        number -= best_value
    return "".join(output)


def greedy_shape_is_valid(text: str, format_number: int) -> bool:
    """Validate the single/pair token shape used by formats 0 through 3.

    This intentionally checks a stricter emitted-token grammar than ARABIC's
    permissive input grammar: every subtraction is a direct pair, and the
    symbol after that pair (when present) is below the subtractor.  This is
    the table's “following the larger one” condition.
    """

    if format_number not in range(4):
        raise ValueError(format_number)
    values: list[int] = []
    for symbol in text:
        if symbol not in SYMBOLS:
            return False
        values.append(COINS[SYMBOLS.index(symbol)])

    index = 0
    while index < len(values):
        value = values[index]
        if index + 1 < len(values) and value < values[index + 1]:
            if not _direct_pair_allowed(format_number, value, values[index + 1]):
                return False
            if index + 2 < len(values) and values[index + 2] >= value:
                return False
            index += 2
            continue
        if any(later > value for later in values[index + 1 :]):
            return False
        index += 1
    return True


def check_arabic_examples() -> int:
    cases = {
        "": 0,
        "I": 1,
        "iv": 4,
        "IX": 9,
        "XIX": 19,
        "IIX": 8,
        "MIM": 1999,
        "IMM": 1999,
        "IIXCMMMM": 3888,
    }
    checks = 0
    for text, expected in cases.items():
        actual = arabic(text)
        if actual != expected:
            raise AssertionError((text, expected, actual))
        checks += 1
    for text in ("A", "Ⅰ", "IV ", "-I", "0"):
        if arabic(text) is not None:
            raise AssertionError((text, arabic(text)))
        checks += 1
    return checks


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--receipt",
        type=Path,
        default=Path(__file__).with_name("receipt.json"),
    )
    args = parser.parse_args()

    unrestricted = signed_distances(CLASSIC_DEPTH_BOUND)
    canonical, witnesses = canonical_distances(CLASSIC_DEPTH_BOUND)
    missing_unrestricted = [n for n in range(MAX_N + 1) if n not in unrestricted]
    missing_canonical = [n for n in range(MAX_N + 1) if n not in canonical]
    mismatches = [
        n for n in range(MAX_N + 1) if unrestricted.get(n) != canonical.get(n)
    ]
    witness_failures = [
        n for n, text in witnesses.items() if arabic(text) != n
    ]

    shape_cases = {
        0: ("", "I", "IV", "IX", "XL", "XC", "CD", "CM"),
        1: ("VL", "LD", "VX", "LC"),
        2: ("IL", "ID", "IM", "LC"),
        3: ("IL", "VL", "ID", "VD", "IM", "VM"),
    }
    shape_checks = sum(
        1 for format_number, cases in shape_cases.items()
        for text in cases
        if greedy_shape_is_valid(text, format_number)
    )
    shape_rejections = {
        "0:IL": not greedy_shape_is_valid("IL", 0),
        "1:IL": not greedy_shape_is_valid("IL", 1),
        "2:VL": not greedy_shape_is_valid("VL", 2),
        "3:IXC": not greedy_shape_is_valid("IXC", 3),
    }
    if len(shape_rejections) != sum(shape_rejections.values()):
        raise AssertionError(shape_rejections)

    simplified_checks = []
    for number in range(MAX_N + 1):
        spelling = source_like_simplified(number) if number else ""
        simplified_checks.append(
            arabic(spelling) == number
            and len(spelling) == unrestricted[number]
        )
    if not all(simplified_checks):
        raise AssertionError("the residue transcription missed an independent BFS bound")

    greedy_checks: dict[str, dict[str, object]] = {}
    for format_number in range(4):
        spellings = [source_like_greedy(number, format_number) for number in range(MAX_N + 1)]
        failures = [
            number
            for number, spelling in enumerate(spellings)
            if arabic(spelling) != number or not greedy_shape_is_valid(spelling, format_number)
        ]
        if failures:
            raise AssertionError((format_number, failures[:8]))
        greedy_checks[str(format_number)] = {
            "roundtrip_and_shape": len(failures) == 0,
            "max_output_digits": max(map(len, spellings)),
            "sample_spellings": {
                str(number): spellings[number]
                for number in (45, 49, 99, 499, 999, 1999, 3999)
            },
        }

    if missing_unrestricted or missing_canonical or mismatches or witness_failures:
        raise AssertionError(
            {
                "missing_unrestricted": missing_unrestricted,
                "missing_canonical": missing_canonical,
                "distance_mismatches": mismatches,
                "witness_failures": witness_failures,
            }
        )

    specification_path = EVIDENCE / "specification.json"
    specification = json.loads(specification_path.read_text(encoding="utf-8"))
    normative_archive = ROOT / specification["source"]
    archive_digest = sha256(normative_archive)
    if archive_digest != specification["source_sha256"]:
        raise AssertionError(
            f"normative archive digest mismatch: {archive_digest}"
        )
    with ZipFile(normative_archive) as archive:
        normative_entry = archive.read(specification["entry"])
    entry_digest = sha256_bytes(normative_entry)
    if entry_digest != specification["entry_sha256"]:
        raise AssertionError(f"normative entry digest mismatch: {entry_digest}")

    source_paths = {
        "roman_module": ROOT / "crates/litchi-ods/src/codec/formula/evaluation/roman.rs",
        "evaluation_dispatch": ROOT / "crates/litchi-ods/src/codec/formula/evaluation.rs",
        "roman_tests": ROOT / "crates/litchi-ods/tests/ods_formula_roman_evaluation.rs",
        "specification": specification_path,
    }
    source_digests = {name: sha256(path) for name, path in source_paths.items()}
    source_digests["normative_archive"] = archive_digest
    source_digests["normative_entry"] = entry_digest
    receipt = {
        "method": "independent integer BFS and canonical-sign state BFS; no Cargo/Rust execution",
        "execution": {
            "command": ["python3", str(Path(__file__).resolve().relative_to(ROOT))],
            "status": 0,
        },
        "scope": {"min": 0, "max": MAX_N, "count": MAX_N + 1},
        "denominations": dict(zip(SYMBOLS, COINS)),
        "unrestricted_signed_bfs": {
            "max_depth": CLASSIC_DEPTH_BOUND,
            "prefix_sum_bound": SUM_BOUND,
            "reached_scope": len(unrestricted) >= MAX_N + 1,
            "max_scope_distance": max(unrestricted[n] for n in range(MAX_N + 1)),
            "distance_histogram": {
                str(depth): sum(unrestricted[n] == depth for n in range(MAX_N + 1))
                for depth in range(CLASSIC_DEPTH_BOUND + 1)
            },
        },
        "canonical_signed_bfs": {
            "reached_scope": not missing_canonical,
            "distance_matches_unrestricted": len(mismatches) == 0,
            "roundtrip_witnesses": len(witnesses) == MAX_N + 1 and not witness_failures,
            "max_scope_distance": max(canonical[n] for n in range(MAX_N + 1)),
            "sample_witnesses": {str(n): witnesses[n] for n in (0, 4, 8, 49, 499, 999, 1999, 3888, 3999)},
        },
        "source_like_simplified_checks": {
            "roundtrip_and_unrestricted_distance": all(simplified_checks),
            "checked_numbers": len(simplified_checks),
        },
        "source_like_greedy_checks": greedy_checks,
        "arabic_indirect_subtraction_checks": check_arabic_examples(),
        "greedy_format_shape_checks": {
            "accepted_examples": shape_checks,
            "rejected_examples": shape_rejections,
        },
        "source_sha256": source_digests,
        "script_sha256": sha256(Path(__file__).resolve()),
    }
    args.receipt.parent.mkdir(parents=True, exist_ok=True)
    args.receipt.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps(receipt, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
