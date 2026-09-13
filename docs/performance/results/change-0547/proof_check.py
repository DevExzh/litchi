#!/usr/bin/env python3
"""Small exhaustive model for the 0547 CFB chain-walk design.

This is a semantic model, not a Rust or allocation model.  It implements the
ordered checks in ``SectorChainScratch::collect_exact`` and compares them with
two hypothetical fast walks:

* terminal-only: omit per-step visited checks and replay the authoritative walk
  on every walk failure;
* checkpointed: use a Brent-style power-of-two sector checkpoint to request
  the same authoritative replay when a cycle is observed, or on any structural
  failure.

The script deliberately uses exact diagnostic strings for the modeled errors.
It enumerates every table over valid slots plus one out-of-range value,
ENDOFCHAIN, and representative reserved markers for table lengths 0..4.  The
model does not exercise Rust allocation failure, alias mutation, or any
hypothetical target where ``usize`` cannot represent a ``u32``; those remain
implementation-level gates.
"""

from __future__ import annotations

from dataclasses import dataclass
import itertools
import json
from typing import Optional, Sequence


MAXREGSECT = 0xFFFF_FFFA
ENDOFCHAIN = 0xFFFF_FFFE

# Representative values from the reserved marker range.  All of these take
# the same ``next >= MAXREGSECT`` branch, but retaining three positions makes
# the exhaustive input set cover the low, middle, and high reserved values.
INVALID_MARKERS = (MAXREGSECT, 0xFFFF_FFFC, 0xFFFF_FFFF)
TABLE_NAME = "model"


Outcome = tuple[str, ...]


@dataclass(frozen=True)
class Run:
    outcome: Outcome
    fast_steps: int
    replay_steps: int
    replayed: bool
    trigger: str
    # A failed Rust call resets these two logical fields after the authority
    # returns.  A terminal success leaves the table-length map logical length
    # in place, even though a terminal-only/checkpointed walk did not set its
    # bits.  That private postcondition is reviewed in design-review.md.
    sectors: tuple[int, ...]
    visited_bit_len: int
    visited_set_count: int


def error(text: str) -> Outcome:
    return ("err", text)


def ok(sectors: Sequence[int]) -> Outcome:
    return ("ok",) + tuple(str(sector) for sector in sectors)


def preflight(table: Sequence[int], start: int, count: int) -> Optional[Outcome]:
    """Return the first pre-loop error, in the current source order."""

    if count == 0:
        if start != ENDOFCHAIN:
            return error(
                f"Empty {TABLE_NAME} chain must start with ENDOFCHAIN"
            )
        return None
    if start >= MAXREGSECT:
        return error(f"Invalid start marker for {TABLE_NAME} chain")
    if count > len(table):
        return error(f"{TABLE_NAME} chain length exceeds its allocation table")
    return None


def authoritative(table: Sequence[int], start: int, count: int) -> tuple[Outcome, int]:
    """Model the current ordered bitset collector and return loop steps."""

    early = preflight(table, start, count)
    if early is not None:
        return early, 0

    sectors: list[int] = []
    seen: set[int] = set()
    sector = start
    steps = 0
    for index in range(count):
        steps += 1
        slot = sector
        # Every generated value is a nonnegative u32.  On the 32-bit and
        # 64-bit targets supported by the crate, the Rust conversion itself
        # succeeds and this is the subsequent table-bound branch.
        if slot >= len(table):
            return (
                error(f"Invalid sector index {sector} in {TABLE_NAME}"),
                steps,
            )
        if slot in seen:
            return (
                error(
                    f"Cycle detected in {TABLE_NAME} chain at sector {sector}"
                ),
                steps,
            )
        seen.add(slot)
        sectors.append(sector)
        next_sector = table[slot]
        if index + 1 == count:
            if next_sector != ENDOFCHAIN:
                return (
                    error(
                        f"{TABLE_NAME} chain exceeds its declared length"
                    ),
                    steps,
                )
        else:
            if next_sector == ENDOFCHAIN:
                return (
                    error(
                        f"{TABLE_NAME} chain ends before its declared length"
                    ),
                    steps,
                )
            if next_sector >= MAXREGSECT:
                return (
                    error(
                        f"Invalid sector marker 0x{next_sector:08X} in "
                        f"{TABLE_NAME} chain"
                    ),
                    steps,
                )
            sector = next_sector

    return ok(sectors), steps


def final_state(outcome: Outcome, table_len: int, *, visited_count: int) -> tuple[
    tuple[int, ...], int, int
]:
    if outcome[0] == "ok":
        return tuple(int(value) for value in outcome[1:]), table_len, visited_count
    # collect_exact resets both logical lengths after every error.  Capacities
    # are intentionally not represented here because they are unchanged by a
    # semantic replay.
    return (), 0, 0


def replay_after_failure(
    table: Sequence[int],
    start: int,
    count: int,
    fast_steps: int,
    trigger: str,
) -> Run:
    outcome, replay_steps = authoritative(table, start, count)
    sectors, bit_len, visited_count = final_state(
        outcome, len(table), visited_count=count if outcome[0] == "ok" else 0
    )
    return Run(
        outcome=outcome,
        fast_steps=fast_steps,
        replay_steps=replay_steps,
        replayed=True,
        trigger=trigger,
        sectors=sectors,
        visited_bit_len=bit_len,
        visited_set_count=visited_count,
    )


def terminal_replay(table: Sequence[int], start: int, count: int) -> Run:
    """Omit visited operations and replay the authority on any walk failure."""

    early = preflight(table, start, count)
    if early is not None:
        sectors, bit_len, visited_count = final_state(
            early, len(table), visited_count=0
        )
        return Run(
            early, 0, 0, False, "preflight", sectors, bit_len, visited_count
        )

    sectors: list[int] = []
    sector = start
    for index in range(count):
        fast_steps = index + 1
        slot = sector
        if slot >= len(table):
            return replay_after_failure(
                table, start, count, fast_steps, "invalid-index"
            )
        # This intentionally does not test or set a visited bit.  A repeated
        # state is only discovered if a later structural check fails.
        sectors.append(sector)
        next_sector = table[slot]
        if index + 1 == count:
            if next_sector != ENDOFCHAIN:
                return replay_after_failure(
                    table, start, count, fast_steps, "late-marker"
                )
            # There is no next valid state after a terminal successor, so no
            # checkpoint update is needed on this successful final position.
        else:
            if next_sector == ENDOFCHAIN:
                return replay_after_failure(
                    table, start, count, fast_steps, "early-marker"
                )
            if next_sector >= MAXREGSECT:
                return replay_after_failure(
                    table, start, count, fast_steps, "invalid-marker"
                )
            sector = next_sector

    outcome = ok(sectors)
    sectors_out, bit_len, visited_count = final_state(
        outcome, len(table), visited_count=0
    )
    return Run(
        outcome,
        count,
        0,
        False,
        "terminal-success",
        sectors_out,
        bit_len,
        visited_count,
    )


def brent_replay(table: Sequence[int], start: int, count: int) -> Run:
    """Model a one-checkpoint Brent walk followed by authoritative replay.

    Checkpoints are at sequence positions 0, 1, 3, 7, ... .  A current
    validated sector is compared with the previous checkpoint before its
    table entry is inspected.  Thus a checkpoint hit is a proven repeated
    valid state, but the candidate never returns a checkpoint diagnostic: it
    replays the authoritative collector to preserve the first error.
    """

    early = preflight(table, start, count)
    if early is not None:
        sectors, bit_len, visited_count = final_state(
            early, len(table), visited_count=0
        )
        return Run(
            early, 0, 0, False, "preflight", sectors, bit_len, visited_count
        )

    sectors: list[int] = []
    sector = start
    checkpoint: Optional[int] = None
    power = 1
    distance = 0

    for index in range(count):
        fast_steps = index + 1
        slot = sector
        if slot >= len(table):
            return replay_after_failure(
                table, start, count, fast_steps, "invalid-index"
            )
        # Preserve current ordering: index/bounds first, cycle observation
        # second, append/table lookup and marker checks afterward.
        if checkpoint is not None and sector == checkpoint:
            return replay_after_failure(
                table, start, count, fast_steps, "checkpoint-cycle"
            )

        sectors.append(sector)
        next_sector = table[slot]
        if index + 1 == count:
            if next_sector != ENDOFCHAIN:
                return replay_after_failure(
                    table, start, count, fast_steps, "late-marker"
                )
        else:
            if next_sector == ENDOFCHAIN:
                return replay_after_failure(
                    table, start, count, fast_steps, "early-marker"
                )
            if next_sector >= MAXREGSECT:
                return replay_after_failure(
                    table, start, count, fast_steps, "invalid-marker"
                )
            # Update only after the same intermediate-marker checks as the
            # authoritative walk.  The exhaustive model has tiny counts, so
            # this cannot overflow.  The design review requires a
            # checked/saturating Rust update with authoritative fallback at
            # the guard.
            if checkpoint is None:
                checkpoint = sector
            else:
                distance += 1
                if distance == power:
                    checkpoint = sector
                    distance = 0
                    power *= 2
            sector = next_sector

    outcome = ok(sectors)
    sectors_out, bit_len, visited_count = final_state(
        outcome, len(table), visited_count=0
    )
    return Run(
        outcome,
        count,
        0,
        False,
        "terminal-success",
        sectors_out,
        bit_len,
        visited_count,
    )


def generated_cases() -> itertools.chain:
    """Yield all requested bounded table/start/count combinations."""

    for length in range(5):
        values = tuple(range(length)) + (length, ENDOFCHAIN) + INVALID_MARKERS
        starts = tuple(range(length)) + (length, ENDOFCHAIN) + INVALID_MARKERS
        for table in itertools.product(values, repeat=length):
            for start in starts:
                for count in range(length + 2):
                    yield table, start, count


def short_cycle_demo(length: int = 32) -> dict[str, object]:
    """Show declared-N amplification for a self-cycle and Brent's bound."""

    # The first entry points to itself; the remaining entries are irrelevant
    # to this chain but make expected_count large while remaining <= table.len.
    table = (0,) + (ENDOFCHAIN,) * (length - 1)
    count = length
    reference_outcome, reference_steps = authoritative(table, 0, count)
    terminal = terminal_replay(table, 0, count)
    checkpointed = brent_replay(table, 0, count)
    assert reference_outcome[0] == "err"
    assert "Cycle detected" in reference_outcome[1]
    assert terminal.outcome == reference_outcome
    assert checkpointed.outcome == reference_outcome
    assert terminal.fast_steps == length
    assert checkpointed.fast_steps == 2
    return {
        "table_len": length,
        "reference_steps": reference_steps,
        "terminal_fast_steps": terminal.fast_steps,
        "terminal_total_steps": terminal.fast_steps + terminal.replay_steps,
        "terminal_total_over_reference": (
            terminal.fast_steps + terminal.replay_steps
        )
        / reference_steps,
        "checkpoint_fast_steps": checkpointed.fast_steps,
        "checkpoint_total_steps": (
            checkpointed.fast_steps + checkpointed.replay_steps
        ),
        "checkpoint_total_over_reference": (
            checkpointed.fast_steps + checkpointed.replay_steps
        )
        / reference_steps,
        "error": reference_outcome[1],
    }


def main() -> None:
    cases = 0
    cycle_cases = 0
    max_terminal_ratio = 0.0
    max_checkpoint_ratio = 0.0
    max_terminal_case: Optional[dict[str, object]] = None
    max_checkpoint_case: Optional[dict[str, object]] = None

    for table, start, count in generated_cases():
        cases += 1
        expected, reference_steps = authoritative(table, start, count)
        terminal = terminal_replay(table, start, count)
        checkpointed = brent_replay(table, start, count)

        assert terminal.outcome == expected, (table, start, count, expected, terminal)
        assert checkpointed.outcome == expected, (
            table,
            start,
            count,
            expected,
            checkpointed,
        )
        for name, run in (("terminal", terminal), ("checkpointed", checkpointed)):
            if expected[0] == "err":
                assert run.sectors == ()
                assert run.visited_bit_len == 0
                assert run.visited_set_count == 0
            else:
                assert run.sectors == tuple(int(value) for value in expected[1:])
                assert run.visited_bit_len == len(table)
                # Both fast walks intentionally leave the retained words clear
                # on a terminal success; no production consumer reads them.
                assert run.visited_set_count == 0

        if expected[0] == "err" and "Cycle detected" in expected[1]:
            cycle_cases += 1
            if reference_steps:
                terminal_total = terminal.fast_steps + terminal.replay_steps
                checkpoint_total = (
                    checkpointed.fast_steps + checkpointed.replay_steps
                )
                terminal_ratio = terminal_total / reference_steps
                checkpoint_ratio = checkpoint_total / reference_steps
                if terminal_ratio > max_terminal_ratio:
                    max_terminal_ratio = terminal_ratio
                    max_terminal_case = {
                        "table": table,
                        "start": start,
                        "count": count,
                        "reference_steps": reference_steps,
                        "total_steps": terminal_total,
                        "ratio": terminal_ratio,
                    }
                if checkpoint_ratio > max_checkpoint_ratio:
                    max_checkpoint_ratio = checkpoint_ratio
                    max_checkpoint_case = {
                        "table": table,
                        "start": start,
                        "count": count,
                        "reference_steps": reference_steps,
                        "total_steps": checkpoint_total,
                        "ratio": checkpoint_ratio,
                    }

    demo = short_cycle_demo()
    print(
        json.dumps(
            {
                "status": "pass",
                "cases": cases,
                "cycle_cases": cycle_cases,
                "max_terminal_cycle_ratio": max_terminal_ratio,
                "max_terminal_cycle_case": max_terminal_case,
                "max_checkpoint_cycle_ratio": max_checkpoint_ratio,
                "max_checkpoint_cycle_case": max_checkpoint_case,
                "short_cycle_demo": demo,
                "scope": (
                    "semantic ordered-error model; no Rust allocation or "
                    "aliasing equivalence claim"
                ),
            },
            sort_keys=True,
        )
    )


if __name__ == "__main__":
    main()
