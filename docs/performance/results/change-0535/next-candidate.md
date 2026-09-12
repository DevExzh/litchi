# Next bounded OLE2/OOXML investigation

The next candidate should ask one narrow code-layout question:

> Can the rare formatted error paths be kept out of the ordinary
> `SectorChainScratch::collect_exact` path without changing any collector
> contract?

This is a conditional follow-up, not an accepted implementation. The 0535
assembly maps a contiguous valid loop at `0x2f29480`–`0x2f29529` and places
the collector's bounds, cycle, terminal-marker, invalid-marker, and
allocation-error tails elsewhere in the same 1,436-byte symbol. The retained
[instruction report](instruction-analysis.json) maps 99.9325 percent of XLS
timed collector self Ir and 99.9723 percent of CFB timed collector self Ir to
that loop, with two CFB setup dumps separated. Those per-workload locations
are enough evidence for a bounded layout experiment; they do not show that
cold layout causes the native timing variation.

The candidate must target the collector's rare diagnostic construction. A
private `#[cold] #[inline(never)]` helper, or an equivalent layout-only
refactoring, could receive the copied values needed to format an error. The
successful bitset and vector operations must remain as currently generated.
Changing the out-of-line `CheckedBitSet::insert` body alone is not a useful
collector candidate: the fresh graph has no positive
`collect_exact` → `CheckedBitSet::insert` edge, even though the 223-byte
symbol exists for other callers.

## Boundaries that remain fixed

The candidate must retain the following behavior and ownership boundaries.

| Area | Required behavior |
| --- | --- |
| Entry state | Clear the retained sector length and visited logical length at the same entry boundary. |
| Empty chains | Keep the `ENDOFCHAIN` requirement and its current error precedence. |
| Preflight | Keep start-marker and table-length checks before resource growth and walking. |
| Allocation | Keep exact sector reservation before the loop, visited-map growth and zeroing, the two resource labels, and their order. |
| Visited map | Keep logical-length and word bounds checks, duplicate detection, and successful set semantics. Do not revive the 0524 fused implementation. |
| Append | Keep the direct ordered sector result, capacity proof, and vector length update. |
| Marker walk | Keep table lookup, final marker, early end, invalid marker, and declared-length checks in their current order. |
| Failure state | Every error must leave the result empty and visited logical length zero while retaining reusable storage. |
| Namespaces | Keep separate MiniFAT and regular-FAT scratch values and their capacities. |
| Ownership | Collect a complete chain before `claim_chain`; keep physical reconciliation as a required positive owner in both stages. |
| Scope | Keep provider per-read freshness and session behavior. Do not revive 0279. |

The initial sector state at `0x2f2945c`, current-slot store at `0x2f294b8`,
next-marker store at `0x2f29500`, and next-sector store at `0x2f2951f` remain
part of the control and diagnostic state. The experiment must not claim that
these three loop stores are removable latency.

## Measurement gate

If the coordinator stages this candidate, it must bind the source, binary,
plan, script, and host receipts again and inspect both `collect_exact` and
`CheckedBitSet::insert` assembly. The profile analyzer must continue to use
absolute `positions: instr` and keep `fn` self Ir separate from `cfn`, call,
and jump metadata. Positive timed constructor ancestry must stay separate from
CFB setup ancestry: XLS parts remain timed, CFB part 1 remains setup, and CFB
parts 2–6 remain timed.

The native admission gate remains the existing OLE2/OOXML gate: all primary
XLS workflows must clear the required p50 improvement threshold in both
paired repeats, while CFB tiny, many-small, and few-large remain individual
guards. The allocation lane must retain allocation-call, allocated-byte, and
incremental-region-peak behavior, with no new per-stream allocation. Review
mean, tail, RSS, correctness, and malformed-input/reset behavior together.
The exact same-build variation policy applies: the 0535 CFB few-large p50
change of +10.0959 percent is retained as an observation and cannot be
treated as a causal baseline or a candidate gain.

Assembly review must answer two concrete questions before any admission:

1. Did the valid loop remain semantically equivalent, including the bitset
   checks, direct append, marker checks, and required stack state?
2. Did the proposed layout move only rare formatting/control tails, with no
   new call or allocation on a successful path?

If either answer is uncertain, or if native/allocation/correctness gates do
not pass, restore the retained baseline and record the candidate as rejected.
No result from this document authorizes a source edit or a capture.

## Rejected directions stay closed

The 0534 paired role/FAT-prefix loop stays rejected after its native rows were
slower. The 0524 visited-bit fusion stays rejected after its admission result.
The 0279 provider freshness-session change stays rejected after its strict
ABBA drift gate. A cold-error-layout probe must not smuggle any of those
changes into a refactor of the collector.

This queue remains OLE2/OOXML work. ODF is deferred until the OLE2/OOXML
optimization goal completes, and iWork remains outside scope.

No Rust build, test, capture, runtime patch, or report rewrite was performed
for this handoff.
