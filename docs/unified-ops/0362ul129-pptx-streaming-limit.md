# 0362ul129 — PPTX streaming invalid-limit consumer correction

This source-only repair starts from the frozen a8 candidate
`a8d32513b8e427c6d234ddd17fb836b0e8d2477b` (tree
`0251ddc8aeef84702604f340406eda671c9b6c54`).  The coordinator's runtime row
`pptx-full-chunk-52` reached
`invalid_global_configuration_does_not_touch_sink` with one stale assertion.
No Cargo, compiler, rustdoc, metadata, native, or performance command was run
in this leaf.

## Runtime finding

The diagnostic is retained outside the repository at
`/home/zhuhe/code/litchi-worktrees/.coord/ooxml-pptx-joint-a8d32513b8-f5-runtime-remaining43-v2/20-pptx-full-chunk-52-attempt-actual-f5-20.log`
(224,308 bytes, SHA-256
`e4698a72c7b3c5fa99c0c5ae4c1f09ae713376f1cce002a606f5cd08394fb341`).
The failing case supplies `max_output_bytes = 512 MiB + 1`, above the
underlying ZIP transport's hard `512 MiB` ceiling.  The writer correctly
rejects this caller-selected configuration in `validate_configuration` as
```rust
Error::InvalidLimit {
    kind: LimitKind::Codec(CodecLimitKind::OutputBytes),
    value: 512 MiB + 1,
    maximum: 512 MiB,
}
```
before constructing `BudgetedSink` or writing to the caller's sink.

## Repair

The test now asserts that exact typed `InvalidLimit` payload and retains the
empty-sink assertion.  Runtime exhaustion remains covered by the neighboring
`Error::Limit` cases and the exact/one-under output tests; this change does not
relax the production validation or alter streaming behavior.

## Source-only validation

`rustfmt --edition 2024 --check crates/litchi-pptx/tests/streaming_authoring.rs`
and `git diff --check` pass.  The staged tree contains only the focused test
consumer correction and this reduced record; no generated evidence, fixture,
executable, or `3rdparty/` path is included.  The coordinator owns the
focused runtime rerun.
