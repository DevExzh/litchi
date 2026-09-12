# 0536 cold error layout candidate

Status: draft production-only candidate; no Rust source was edited, and no build,
test, or benchmark command was run for this draft; root formatted the applied
source before freezing the candidate.

## Baseline and evidence

The candidate applies to commit `8876e87b8dbcb7dca54a3416386bbfc892c68eb4`.
The baseline `crates/litchi-cfb/src/file.rs` SHA-256 is
`72fb5295dfbba5f3ed2192a9c24ecbf7d8aa2759f418b77563a4e2ff1042d03b`. The patch has one production file and no test or harness
changes.

The 0535 instruction evidence identified the successful collector loop in
`SectorChainScratch::collect_exact` and showed that its formatted diagnostic
branches are cold relative to the loop. The collector symbol was measured at
`0x2f292c0` (1,436 bytes); the normal loop occupied `0x2f29480` through
`0x2f29529`. This experiment isolates the layout change by moving only the
collector's eight formatted error constructions into private cold, out-of-line
associated helpers. It does not claim a speedup until the root-owned native,
allocation, correctness, and assembly gates pass.

## Candidate change

`SectorChainScratch` gains these private helpers, each marked
`#[cold]` and `#[inline(never)]`:

- `chain_error_empty`
- `chain_error_invalid_start`
- `chain_error_length`
- `chain_error_index`
- `chain_error_cycle`
- `chain_error_exceeds`
- `chain_error_early_end`
- `chain_error_marker`

`collect_exact` calls the corresponding helper at each existing diagnostic
branch. `CheckedBitSet`, including `insert`, is unchanged. The successful
visited check, insertion, sector push, and table lookup remain in their current
order and use the same types and operators.

The goal is to remove formatting machinery and cold error construction from the
collector's hot instruction layout. The helper arguments are the existing
`&str`, `u32`, and no additional state; helpers only construct and return the
same `OleError::CorruptedFile` value.

## Behavior and ordering contract

The exact messages remain:

| Existing branch | Exact payload |
| --- | --- |
| empty chain with a non-ENDOFCHAIN start | `Empty {table_name} chain must start with ENDOFCHAIN` |
| nonempty chain with a reserved start marker | `Invalid start marker for {table_name} chain` |
| declared count larger than the allocation table | `{table_name} chain length exceeds its allocation table` |
| conversion or table-slot bounds failure | `Invalid sector index {sector} in {table_name}` |
| repeated visited slot | `Cycle detected in {table_name} chain at sector {sector}` |
| non-ENDOFCHAIN after the declared final slot | `{table_name} chain exceeds its declared length` |
| ENDOFCHAIN before the declared final slot | `{table_name} chain ends before its declared length` |
| reserved marker in a non-final link | `Invalid sector marker 0x{next:08X} in {table_name} chain` |

The patch preserves the entry `reset`, empty-chain check, start-marker check,
allocation-table length check, exact sector reservation, visited-map preparation,
`usize::try_from` conversion, bounds check, cycle check, `insert`, sector push,
next-entry lookup, final/non-final marker checks, and error-path reset. Thus it
keeps allocation labels/order, reuse state, typed `CorruptedFile` results, and
all success/error ordering intact. It retains the host-width conversion instead
of introducing any narrower or synthetic 16-bit behavior.

No claim/reconciliation, physical-layout validation, provider freshness, public
API, OLE2/OOXML behavior, or ODF behavior is changed.

## Review and measurement gates

Before candidate measurement, root should run the existing focused collector
contract/differential tests. Because this draft leaves `CheckedBitSet` unchanged,
no separate bitset test is required by this candidate. Assembly inspection should
select `SectorChainScratch`, `chain_error_*`, and the existing
`CheckedBitSet` symbols and verify that helpers are out of line while the normal
loop preserves required state and successful checks. Any unanticipated hot call,
changed error payload/order/reset behavior, allocation change, or failed
correctness/quality gate rejects the candidate. Native and allocation admission
remain those in the 0536 plan; no performance claim is made from source layout
alone.
