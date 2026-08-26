# Change 0286: ODT semantic text to a sequential sink

Date: 2026-08-26

Status: Accepted bounded streaming evidence

Performance claim: none

## Decision

The opt-in `odt_semantic_text_to_sink` selector adds deterministic ODT
semantic-text streaming evidence for the existing generated text corpora. The
harness opens and validates one ODT `Document`, prepares `TextOutputOptions`,
and then times only the `write_text_to` call for each sample. A
`HashingDiscardSink` retains no output bytes while recording the accepted byte
count, write-call count, largest write, and write-size buckets.

After the timed call, the harness checks the `TextOutputReport` byte and object
counts and compares the sink SHA-256 with the independently specified semantic
paragraph output. The retained output field is `Some(0)`, and the deterministic
write-call/largest-write summary must match across measured samples. The
existing `odt_semantic_full_text` case and the default matrix remain unchanged;
this selector is opt-in only.

## Scope and claim boundary

This is deterministic streaming evidence only. `performance_claim` is `none`:
the selector makes no latency, throughput, RSS, allocator, physical-I/O,
zero-copy, or whole-transaction memory-bound claim. The hashing sink bounds
retained output, not the ODT model or the complete authoring transaction. The
production sink parser enforces its 64 MiB aggregate decoded-text cap before
each append; broader nested-block retention and hard caps/resource evidence
covering every ODT semantic traversal remain unresolved and require separate
production and resource evidence.
