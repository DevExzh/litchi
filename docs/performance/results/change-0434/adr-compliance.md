# Batch 0434 ADR compliance

This record covers the committed ODS text-span candidate at
`65dfb1b141190271763367defce4086ef146c056`. The accepted ADR tree is
unchanged at `c950b6c8be822561b498d7bbe87c460873dcbf49`. This batch changes one
private ODS encoder path; it does not add a public API, a dependency, an
allocation strategy, or a common-crate ownership boundary.

## Change boundary

[`MAX_PLAIN_TEXT_SPAN_BYTES`](../../../../crates/litchi-ods/src/streaming.rs#L45-L46)
limits a borrowed ordinary-text span to 256 bytes. The scanner still checks
cancellation once per Unicode scalar. A span is attempted only after checked
row-window arithmetic and a cancellation check; its `Resource::Work` charge is
the encoded XML byte length, then the borrowed bytes are written directly into
the existing row window. XML references and whitespace characters retain the
previous scalar encoding path.

If a span would exceed the row window, overflow `u64`, or a finite Work limit,
the attempt returns `Fallback` without recording failure state
([`write_plain_span`](../../../../crates/litchi-ods/src/streaming.rs#L1022-L1053)).
The scalar path then reproduces the baseline refusal boundary and accepted prefix.
A cancellation or non-limit execution error is returned directly. A cancellation
while a span is pending therefore leaves its scanned prefix neither charged nor
serialized; the current row is discarded by the generated-XML producer and the
outer sink's already accepted bytes remain authoritative. Work continues to
mean encoded XML bytes, rather than source-character count.

The implementation uses `&[u8]` borrowed from the caller's text and the
existing row buffer. It does not create a second text-sized allocation. The
common generated-XML auditor, ZIP publication boundary, output accounting,
required-memory reservation, and public `stream_scalar_rows_to` contract are
unchanged.

## ADR mapping

| ADR | Compliance in this batch | Boundary |
| --- | --- | --- |
| [0001](../../../adr/0001-priorities-and-api-layers.md) | The optimization remains inside the typed ODS authoring layer and preserves explicit error and cancellation handling. | No raw archive or unrestricted XML surface is introduced. |
| [0002](../../../adr/0002-crate-topology.md) | ODS owns scalar text encoding; `litchi-odf-common` continues to own generated-XML auditing and archive publication. | No dependency direction or common ownership changes. |
| [0003](../../../adr/0003-snapshots-edits-and-patches.md) | The path still creates a fresh one-sheet stream. | It makes no existing-document append, edit, patch, or snapshot claim. |
| [0004](../../../adr/0004-semantic-api-design.md) | Existing typed scalar semantics, XML escaping, and error types are retained while only the internal write grouping changes. | No public semantic API is added. |
| [0005](../../../adr/0005-io-memory-and-performance.md) | The 256-byte span cap bounds polling time; cancellation, finite Work/row/output limits, and sequential sink progress remain active. | The span cap is not a total RSS or allocator-peak bound. Performance interpretation requires the frozen ABBA evidence. |
| [0006](../../../adr/0006-validation-security-and-compatibility.md) | The candidate preserves XML 1.0 checks, reference encoding, deterministic output, aggregate audit, and audit-before-publication. | No validation relaxation or repair is performed. |
| [0008](../../../adr/0008-migration-and-verification.md) | Baseline/candidate source custody, differential scalar-reference tests, exact limits, and release receipts are retained. | Functional receipts do not authorize a speedup or causal optimization claim. |
| [0010](../../../adr/0010-facade-archive-ownership.md) | ODS still delegates package and ZIP grammar to the existing common writer. | The candidate does not expose archive records. |
| [0023](../../../adr/0023-odf-family-crate-split.md) | The ODS family owns its scalar grammar and family-level tests. | Common-crate success is not treated as ODS or native-application compatibility. |
| [0024](../../../adr/0024-current-topology.md) | The current `litchi-odf-common`/`litchi-ods` topology is used. | No historical monolith ownership is assumed. |

## Functional custody

The retained ODS release receipt reports **454 passed, 0 failed, 0 ignored**
([`ods-tests-final.json`](checks/ods-tests-final.json)). The focused integration
receipt reports **16 passed** ([`ods-text-integration-first.json`](checks/ods-text-integration-first.json));
the full release log includes the fifth private competing-limit unit test as
well as the four tests selected by the filtered unit receipt
([`ods-text-unit-v2.json`](checks/ods-text-unit-v2.json)). Those checks cover
the independent scalar reference
([`streaming_text_unit_tests.rs`](../../../../crates/litchi-ods/src/streaming_text_unit_tests.rs#L141-L286)),
every row-window and Work offset, competing limits, cancellation, failed-span
prefix behavior, and accepted partial output
([`streaming_text_spans.rs`](../../../../crates/litchi-ods/tests/streaming_text_spans.rs#L203-L314)).

The candidate release harness build passes
([`after-build.json`](checks/after-build.json)). Scoped all-target Clippy with
the existing `large_enum_variant` exemption, documentation, format, and crate
boundary receipts pass ([`ods-strict-scoped-v2.json`](checks/ods-strict-scoped-v2.json),
[`ods-doc-final.json`](checks/ods-doc-final.json),
[`owned-format-final.json`](checks/owned-format-final.json),
[`boundaries-final.json`](checks/boundaries-final.json)). The unscoped existing
large-enum lint remains a separate workspace debt.

The test receipts retain source manifests and their recorded revision fields;
the full test receipt was captured before the candidate commit while the
release harness build receipt is at `65dfb1b1`. This record keeps those custody
facts distinct and does not relabel the historical test invocation as a new
post-commit run.

## Performance status

The formal ABBA protocol is retained in [`protocol.json`](protocol.json), with
exact archive/XML/output/semantic identity gates before comparison. All 24
reports, 720 samples and four profiles pass those gates. The [summary](summary.json)
records normal medians 12.499–16.211% lower for the tested sizes/repeats and
unchanged requested-allocation metrics. No matched >5% regression is flagged;
baseline tiny p99 repeat drift is +9.844%. These are scoped observations;
process-wide profiling, RSS and cache values do not establish operation-local
or total-memory guarantees. The overall non-iWork goal remains open.
