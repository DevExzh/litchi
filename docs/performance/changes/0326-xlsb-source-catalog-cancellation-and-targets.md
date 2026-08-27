# Change 0326: XLSB source catalog cancellation, freshness fences, and targets

## Scope

This batch hardens the source-backed XLSB catalog and handle APIs. Fallible
catalog and handle operations now use phase-specific managed cancellation and
freshness fences, and constructors publish their source-backed state only after
construction has completed successfully. The catalog also rejects duplicate
sheet targets using the same case-insensitive target identity as the OPC
layer, while still allowing unique case variants. Non-worksheet catalog tabs
are refused with a typed error before any text output is produced.

Materialization now uses source-first execution with a postflight freshness
check. A cancellation-during-deferred-read regression test proves that the
operation returns typed `Cancelled` and does not publish a cache.

These changes address lifecycle and target-selection correctness. They do not
make a performance, resident-memory, or OOM claim.

## Validation evidence

- Focused `source_backed` tests: 25 passed.
- The final skip-only crate run passed 523 library tests and 114 integration
  tests, with only the unrelated
  `checked_in_unique_standard_drawing_corpus_transfers_every_anchor` test
  skipped.
- The skipped test was run independently and failed with observed `5` versus
  expected `6`.
- Strict Clippy passed after the final change.
- Rustdoc passed after the final change.
- Rustfmt was applied.
- `git diff --check` passed.

## Explicit non-claims

No performance, RSS, memory-accounting, or OOM improvement was measured or is
claimed by this batch.

## Residual gaps

- Opaque parsers still do not poll for cancellation internally.
- Semantic caches and memory accounting are unchanged.
- Dependency parity and cache publication remain unresolved.
