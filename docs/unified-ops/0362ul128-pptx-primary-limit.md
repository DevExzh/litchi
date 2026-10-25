# 0362ul128 — PPTX source text-box limit precedence

Date: 2026-10-25

## Scope

This leaf is based on frozen source commit
`a8d32513b8e427c6d234ddd17fb836b0e8d2477b`. It repairs the production
admission accounting exposed by the preserved runtime diagnostic
`pptx-full-chunk-50`.

Diagnostic log:

```text
.coord/ooxml-pptx-joint-a8d32513b8-f5-runtime-remaining43-v2/18-pptx-full-chunk-50-attempt-actual-f5-18.log
225307 bytes
SHA-256 0acc965cd9fe88eacc6a48520ac5d3f9123e5e12fca5aafde6a6a75a0d3fd739
```

The failing test was
`crates/litchi-pptx/tests/source_primary_text_box.rs:143`,
`caller_output_limits_precede_managed_memory_admission`. Its baseline
`add_text_box("retry", ...)` unexpectedly returned
`Limit { kind: Codec(MemoryBytes), observed: 379013, maximum: 102282 }`.

## Diagnosis and repair

`reserve_source_text_box` correctly calls `text_box_output_bound` before any
managed reservation, so an output-policy violation must win over memory
pressure. Its later transient-envelope check, however, reused
`max_output_bytes` as the memory ceiling. A valid small-output edit could
therefore be rejected because its scene indexes and transient buffers exceed
the serialized candidate's byte ceiling.

The function now uses the attached `ExecutionContext` budget's
`Resource::Memory` limit for the transient-envelope admission. The caller's
semantic output ceiling remains the argument to `text_box_output_bound`, and
the final `context.reserve(Resource::Memory, amount)` still enforces current
usage and hierarchical parent budgets. No limit was raised and no output
check was weakened.

The existing targeted regression covers all four physical source-limit
dimensions and managed/unmanaged owners: the ordinary baseline must succeed,
an oversized escaped value must return the output-byte limit even while the
managed memory budget is pressured, and a smaller retry must remain valid.
No test fixture or assertion was changed.

## Validation

No Cargo, build, rustdoc, metadata, or runtime command was run. With the
pinned repository toolchain and formatter configuration, these source-only
checks pass:

```text
rustfmt --edition 2024 --check crates/litchi-pptx/src/presentation/source.rs crates/litchi-pptx/tests/source_primary_text_box.rs
git diff --check
```

The diagnostic log remains the pre-fix result; this record makes no post-fix
runtime-pass claim.
