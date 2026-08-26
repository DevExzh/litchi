# Change 0288: ODS semantic text to a sequential sink

Date: 2026-08-26

Status: Accepted bounded streaming evidence

Performance claim: none

## Decision

The opt-in `ods_semantic_text_to_sink` selector adds deterministic ODS
semantic-text streaming evidence for the existing generated spreadsheet
corpora. One logical worksheet row is one paragraph-equivalent output object;
the row's stored cell text is joined with tabs. A sheet with no logical rows
contributes one synthetic empty object. The harness opens and validates one
ODS `Spreadsheet`, derives the expected row text from the independent
row-major full-cell oracle, prepares `TextOutputOptions`, and times only
`Spreadsheet::write_text_to` for each sample.

After the timed call, the harness checks the exact `TextOutputReport` byte and
object counts, accepted sink bytes, and SHA-256 digest against the independent
oracle. The retained output field is `Some(0)`, and the deterministic
write-call/largest-write summary must match across measured samples. The
existing `ods_semantic_full_cell_text` case and the default matrix remain
unchanged; this selector is opt-in only.

## Scope and claim boundary

This is deterministic streaming evidence only. `performance_claim` is
`none`: the selector makes no latency, throughput, allocator, RSS,
physical-I/O, zero-copy, or whole-transaction memory-bound claim. The hashing
sink bounds retained output, not the ODS model, row/cell traversal, archive
parsing, or a complete authoring transaction. The configured text-output byte
and object ceilings are output-contract checks, not process-memory evidence.
Broader repeated-row, sparse/empty-sheet, real-producer, source-backed,
cold-cache, and resource evidence remain separate production or measurement
work.
