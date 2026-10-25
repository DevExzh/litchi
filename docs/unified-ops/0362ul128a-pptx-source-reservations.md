# 0362ul128a — PPTX source reservation dimensions

Date: 2026-10-25

## Scope

This follow-up is anchored to the 128 source fix
`788ee37ffdc06f59c7dc0e59426d2e68492ea5cc`. It audits every source-backed
reservation helper in `crates/litchi-pptx/src/presentation/source.rs`, not just
the text-box path exposed by the preserved runtime diagnostic
`pptx-full-chunk-50`.

## Repair

`reserve_source_shape_removal`, `reserve_source_slide_properties`,
`reserve_source_shape_properties`, `reserve_source_shape_text_chunks`, and
`reserve_source_shape_property_chunks` had reused the serialized
`max_output_bytes` value as the managed transient-memory ceiling. They now use
the attached execution context's `Resource::Memory` limit, matching the
already-corrected text-box path. A shared checked reservation helper keeps
`usize`→`u64` conversion fallible and maps the resulting charge to the typed
memory limit. Output-bound arithmetic reports `OutputBytes`; memory-envelope
arithmetic reports `MemoryBytes`.

The output policy remains independent: chunk callers perform their scalar
output admission before reservation, and the exact native output checks remain
authoritative for shape removal, slide properties, and shape properties. Those
three operations may still report managed memory first when both an exact
candidate output and the current memory budget would fail; no extra full parse
or candidate allocation was introduced merely to change that precedence.
This preserves exact no-op/shortening behavior and the existing output-before-
memory paths without weakening bounded allocation.

## Regression coverage

The source test
`source_edit_reservations_keep_output_and_memory_dimensions_independent`
exercises all six reservation helpers with a tight serialized-output ceiling
and a sufficient managed memory budget, verifies unmanaged calls do not retain
a reservation, and verifies a 1 KiB parent budget refuses a 4 MiB child
reservation through the typed execution resource-limit path. The
`source_slide_property_edit_accepts_valid_tight_output_with_managed_memory`
test exercises a real source-backed slide-property edit whose output ceiling
is only 128 bytes above the input while its managed transient envelope is
larger. The existing integration test
`caller_output_limits_precede_managed_memory_admission` continues to cover
output-first behavior for text-box authoring under managed and unmanaged
owners, including the tight retry path.

Record 0152 scopes that output-before-memory guarantee to the text-box
preflight; the adjacent exact stages retain their existing no-op/shortening
ordering and do not pre-stage a candidate merely to manufacture a global
output-first guarantee.

## Validation

No Cargo, build, rustdoc, metadata, or runtime command was run. Source-only
checks pass:

```text
rustfmt --edition 2024 --check crates/litchi-pptx/src/presentation/source.rs
git diff --check
```
