# ODS data-style completion checkpoint

The checked ODS data-style model and source-qualified package owner are landed in
two validated commits: `b416a3a54` corrects and completes the pure model vocabulary,
and `8e3ad3104` adds source scanning, owner-qualified reads, and bounded source
edits. The archived [model evidence](ods-data-style-model-corrections/README.md)
and [source-integration evidence](ods-data-style-source-integration/README.md)
retain the source manifests, review records, commands, logs, and exact hashes.

## Public surface

The public [`litchi_ods::data_style` vocabulary](../../../crates/litchi-ods/src/data_style/mod.rs)
contains `Owner`, `Family`, `Selector`, typed `Number` and `Data` values, `Entry`,
`Graph`, and metadata `Patch` values. Number builders cover decimal, scientific,
and fraction bodies with the checked affix/embedded-text particles and common
transliteration metadata. Data builders cover date, time, currency, percentage,
and boolean roots.

The document boundary re-exports the source-qualified types in
[`document.rs`](../../../crates/litchi-ods/src/document.rs). `Snapshot::data_style`
and `Snapshot::data_styles` read one selected definition or a bounded catalog from
`content.xml` automatic styles, common `styles.xml` styles, or `styles.xml`
automatic styles. `Snapshot::effective_cell_data_style` resolves a cell-style
data-style reference through content automatic styles and then common styles,
with ambiguity refused. `Edit::patch_data_style` changes only common root metadata
on a selected mutable content-automatic definition; `Edit::put_extended_style_graph`
and `Edit::replace_extended_style_graph` add or replace typed number/data graph
members in that same content-automatic owner.

## Preservation and limits

The borrowed scanner keeps source-qualified ranges and inherited namespace context.
Metadata edits preserve the selected body and opaque children byte-for-byte.
Typed graph replacement skips semantically unchanged nodes, checks complete decoded
`content.xml` ranges rather than ZIP offsets, and validates the candidate before
publication. Exact semantic no-ops, inverse application, stale-source refusal,
deflated late-member replacement, aggregate output preflight, and package reopen
are covered by the source integration tests.

The two `styles.xml` owners are read-only for this slice. `text-style` and opaque or
structurally unsupported bodies are readable and retain their metadata identity, but
canonical graph replacement refuses them. Authored graph insertion/replacement is
limited to typed number/data members; there is no full data-style graph resolver,
typed text-style body authoring, calculation, or rendering claim. Native producer
acceptance is also outside this gate.

## Validation and review

The public [`data_style_vocabulary` tests](../../../crates/litchi-ods/tests/data_style_vocabulary.rs)
passed 47 tests. The combined root capture passed 2,234 tests across 127 result
groups, with one ignored test; all-target Clippy with warnings denied and rustdoc
with warnings denied also passed. The scanner review and the scoped v9 owner review
approved the source-qualified ranges, selectors, metadata edits, graph dependency
closure, failed staging, stale/inverse checks, signature policy, and facade
integration.

The transaction and preservation boundary follows [ADR 0003](../../adr/0003-snapshots-edits-and-patches.md),
the preserve-or-refuse policy follows [ADR 0006](../../adr/0006-validation-security-and-compatibility.md),
and ODS family ownership follows [ADR 0023](../../adr/0023-odf-family-crate-split.md).
Measured evidence now follows [ADR 0005](../../adr/0005-io-memory-and-performance.md).
Commit `750b8d450` retains the reviewed [source-workflow characterization](ods-data-style-source-performance/README.md).
Commit `f2daff703` removes a repeated parse during graph insertion preflight;
its [matched evidence and replay](ods-data-style-graph-preflight-performance/README.md)
show 11.16% fewer allocation calls and 8.24% fewer requested bytes for graph put
on the scale-512 synthetic fixture, with unchanged allocator-observed logical
peak memory. An independent replay matched all deterministic metrics and nine
fixture hashes. The optimization separately passed 298 library tests, 47
integration tests and strict library Clippy. These measurements do not establish
RSS reductions, durable-publication throughput, or a general latency guarantee.
