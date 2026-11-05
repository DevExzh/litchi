# 0362ul94a — shared OOXML error-contract publication

Status: implementation complete on the error-contract lane, based on
`6bf2a5f68a`. The source changes are intended for composition with the
coordinator's Web/Ribbon/VBA publication candidate and the properties lane.
No Cargo, rustdoc, glossary CLI, native, or performance command was run.

## Typed common failures

`litchi-ooxml-common::Error` now owns the shared semantic variants required by
record 0351: `Refused(Refusal)`, `PatchConflict`, `Signed`,
`SourceChanged { expected, observed }`, and `Internal`. OPC source freshness
and signature-policy errors are converted at the common boundary, while
malformed OPC, relationship, content-type, XML, and payload failures retain
their corruption variants.

The shared signature guard returns `Signed` before a changed package is
mutated. Web, Ribbon, and VBA publication paths classify stale recorded graph
bases as `PatchConflict`, valid lossless-preservation refusals as typed
`Refused`, and generated staging, admission, archive-metadata, or publication
invariants as `Internal`. VBA host/main-part mismatches use
`Refused(UnsupportedTarget)`; Word supplemental-data mismatches use
`Refused(UnsupportedContent)`. Exact no-ops and bounded source admission are
unchanged.

## Leaf and facade migration

DOCX, PPTX, XLSX, and XLSB explicitly flatten the common semantic variants
into their crate-owned errors. Their ops classifiers report refusals by
`RefusalKind`, patch conflicts as `invalid_input`, signed writes as
`document_protected`, source changes as `io_error`, and internal invariants as
`internal_error`. The PPTX embedded-VBA adapter and XLSX signature-policy
adapter preserve the same typed mapping.

A focused `litchi-ops` test covers all five shared semantic variants through
each OOXML leaf wrapper. Existing Ribbon, Custom XML, VBA, and signature tests
were migrated to the typed signed error. The common properties test error
taxonomy was extended for the new non-exhaustive variants; properties
transaction implementation remains outside this lane.

The review follow-up also migrated the remaining signed-error assertions in
custom-properties, extended-properties, and core-properties writer tests,
updated all canonical producer documentation to describe `Error::Signed`, and
retained source chains through common and leaf error wrappers.

## Source-only validation

The changed Rust paths pass `rustfmt --edition 2024 --check`, and
`git diff --check` passes. The staged repository-size and binary guards must
remain empty at commit time. Build and runtime validation are deferred to the
coordinator's composed source tree.
