# Change 0287: ODP semantic text to a sequential sink

Date: 2026-08-26

Status: Accepted bounded streaming evidence

Performance claim: none

## Decision

The opt-in `odp_semantic_text_to_sink` selector adds deterministic ODP
semantic-text streaming evidence for the existing generated presentation
corpora. The harness opens and validates one ODP `Presentation`, builds an
independent oracle from each slide's title, a newline, and its body, and joins
slides with double newlines. `TextOutputOptions` is constructed before the
timed loop, and each sample times only `Presentation::write_text_to` into a
`HashingDiscardSink::without_authoring_window` sink.

After the timed call, the harness checks the exact `TextOutputReport` byte and
object counts, accepted sink bytes, and SHA-256 digest against the independent
oracle. The sink retains zero output bytes and requires nonzero,
deterministic write-call and largest-write evidence across measured samples.
The existing `odp_semantic_full_text` case and the default matrix remain
unchanged; this selector is opt-in only.

## Scope and claim boundary

This is deterministic streaming evidence only. `performance_claim` is `none`:
the selector makes no latency, throughput, allocation, RSS, physical-I/O,
zero-copy, or whole-transaction memory-bound claim. The hashing sink bounds
retained output, not the ODP model or complete authoring transaction. Broader
producer, resource, cold-cache, and semantic conversion performance evidence
remain separate work.
