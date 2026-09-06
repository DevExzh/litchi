# Batch 0433 ADR compliance

This record covers the generated-XML common seam and the typed ODS scalar
stream writer. The focused common and ODS receipts pass (19 and 11 tests).
The recorded package-test, documentation, and scoped-strict gates also pass;
the unscoped strict lint still has the pre-existing `ArchiveReaderKind`
`large_enum_variant` finding. The bounded streaming implementation is now
committed as `f5bf1696192f0007db56554aa5b16719cdf1950b`; the receipts retain
their own historical revision and source-manifest provenance. It is not a full
support, compatibility, or performance claim.

The receipts identify recorded HEAD `be82ef53dc191a64bc3c3365502bf44599528577`,
but their `source_before`/`source_after` fields also point to captured source
manifests. The generated-XML and ODS implementation files were uncommitted
candidate files at capture time and are not part of that commit. Therefore the
revision and source-manifest provenance must be read together; these receipts
do not turn the candidate into committed `be82` history.

## Scope and ownership

`litchi-odf-common` owns the checked XML envelope, fragment audit, aggregate
audit limits, ZIP admission, manifest publication, and typed sequential-output
errors. `litchi-ods` owns the `Sheet1` envelope, scalar cell grammar, row
iteration, XML escaping, ODS limits, and execution-context checks. The public
entry point is a fresh one-sheet operation through
[`stream_scalar_rows_to`](../../../../crates/litchi-ods/src/streaming.rs);
it consumes rows once and accepts variable row and cell counts within the
configured limits. The fixture uses four cells in some rows, but that is not a
writer-level four-cell restriction.

The package path has a deliberate sequential partial-output boundary. Envelope
and package-limit preparation happens before ZIP admission and does not pull a
row, so those failures are recoverable before output. On the normal
`PackageWriter` path, ZIP admission writes the local header before the first
reader callback is invoked. The first callback and its fragment audit therefore
run after admission; a producer or audit error is wrapped with its source,
poisons the package, and requires the caller to discard the incomplete output.
Every later fragment is audited before its bytes are returned to the ZIP
reader, so invalid fragment bytes are never published. This is the final
post-admission callback contract; there is no separate writer prefetch step.

## Accepted ADR mapping

| ADR | Applied requirement | Evidence and boundary |
| --- | --- | --- |
| [0001](../../../adr/0001-priorities-and-api-layers.md) | Keep layers explicit, validated, typed, and panic-free; do not expose an unrestricted XML or archive escape hatch. | `GeneratedXmlEnvelope` is checked before publication, each fragment is audited, and producer/sink failures retain typed progress and sources. ODS exposes typed cells and limits rather than raw ZIP records. |
| [0002](../../../adr/0002-crate-topology.md) | Common owns neutral package/XML substrate; the family crate owns ODS grammar and authoring. | The common seam is lexical and format-neutral; `litchi-ods` supplies the fixed ODS envelope, scalar serialization, and row semantics. No dependency direction from common into a family crate is introduced. |
| [0003](../../../adr/0003-snapshots-edits-and-patches.md) | Snapshot/edit/patch semantics remain separate from a forward-only writer. | This batch creates a new package only. It makes no existing-document append, edit, patch, or snapshot claim. |
| [0004](../../../adr/0004-semantic-api-design.md) | Use focused typed semantic APIs and validate before constructing ordinary semantic output. | `StreamingCell`, `StreamingLimits`, and `stream_scalar_rows_to` carry the semantic contract; `PackageWriter` remains the lower-level common publication layer. Nonfinite numbers and illegal XML text are refused. |
| [0005](../../../adr/0005-io-memory-and-performance.md) | Enforce finite hierarchical budgets, use caller-owned sequential sinks, and keep streaming scratch bounded and explicit. | `ExecutionContext`, row/output/XML limits, `BudgetedOutput`, reusable fragment storage, cancellation checks, and typed incomplete-output progress are active. The required-memory reservation is documented separately below; it is not a physical peak measurement. |
| [0006](../../../adr/0006-validation-security-and-compatibility.md) | Validate without repair, produce deterministic output/reports, and keep security features explicit. | Scalar XML 1.0 validation, strict envelope/fragment shape checks, aggregate audit limits, deterministic `Sheet1` output, and explicit rejection of configured encryption/signing are retained. |
| [0008](../../../adr/0008-migration-and-verification.md) | Build support in dependency order and distinguish functional streaming checks from performance/support claims. | The focused receipts, semantic readback, final harness checks, and passing after-build receipt establish functional bounded creation behavior. Detailed profile interpretation remains separate, so this record does not promote the full goal or claim a measured resource improvement. |
| [0010](../../../adr/0010-facade-archive-ownership.md) | Keep ZIP grammar and physical archive ownership in the format/common owner. | ODS passes typed rows to `PackageWriter`; it does not own or expose raw ZIP grammar. |
| [0023](../../../adr/0023-odf-family-crate-split.md) | Keep one common substrate with per-family grammar and require family-level evidence. | The common generated-XML checks and ODS streaming checks are separate. Common success alone is not treated as proof of the ODS artifact or of native-application compatibility. |
| [0024](../../../adr/0024-current-topology.md) | Use the current workspace topology as authoritative. | The implementation and evidence use the current `litchi-odf-common`/`litchi-ods` package split; no old monolith ownership is assumed. |

The implementation introduces no `unsafe` requirement and no new dependency
direction. It reuses the common XML audit and ZIP publication machinery, with
the family crate supplying only its typed content producer.

## Budget and memory boundary

Before output begins, `StreamingLimits::required_memory_bytes()` reserves the
reusable row-fragment window, two copies of the fixed content XML shell, and
the common 64 KiB metadata staging reservation. Its shape is:

```text
max_row_xml_bytes + 2 * (CONTENT_PREFIX.len() + CONTENT_SUFFIX.len()) + 64 KiB
```

The reservation covers the row window, shell, and metadata staging needed by
this operation. It does not account for the full ZIP writer, XML auditor,
parser, compression, or allocator physical peak. Those allocations and any
RSS or near-limit observation require separate evidence. No memory or speedup
claim is made here.

## Focused functional evidence

The focused receipts were recorded with Rust 1.98.1 and revision field
`be82ef53dc191a64bc3c3365502bf44599528577`; each receipt retains its own
source manifest and unchanged-before/after comparison:

- [`common-generated-focused-v4.json`](checks/common-generated-focused-v4.json)
  records `cargo test --locked --release -p litchi-odf-common generated_xml -- --test-threads=1` as **19 passed, 0 failed, 0 ignored**.
- [`ods-streaming-integration-v5.json`](checks/ods-streaming-integration-v5.json)
  records `cargo test --locked --release -p litchi-ods --test streaming_creation -- --test-threads=1` as **11 passed, 0 failed, 0 ignored**.
- The ODS creation readback asserts exactly three physical members:
  `mimetype`, `content.xml`, and `META-INF/manifest.xml`. It reopens the
  generated bytes with Litchi's `Spreadsheet::from_bytes`, checks the fixed
  `Sheet1`, and verifies empty, variable, escaped, numeric, boolean, and empty
  scalar rows. Limit, cancellation, short-sink, aggregate-audit, and producer
  failure cases are included in the focused tests.

The package and production retained checks with recorded revision field
`be82ef53dc191a64bc3c3365502bf44599528577` report:

- [`common-tests-final.json`](checks/common-tests-final.json): **437 passed,
  0 failed, 1 ignored**.
- [`ods-tests-final.json`](checks/ods-tests-final.json): **444 passed, 0
  failed, 0 ignored**.
- [`production-doc-final.json`](checks/production-doc-final.json) passes the
  warning-denied documentation build.
- [`production-strict-scoped-final.json`](checks/production-strict-scoped-final.json)
  passes with only the pre-existing `clippy::large_enum_variant` category
  exempted.
- [`production-strict-final.json`](checks/production-strict-final.json) still
  fails only on the unchanged `ArchiveReaderKind` enum in
  `crates/litchi-odf-common/src/package/model.rs`; this is retained as an
  honest strict-gate debt rather than hidden by the scoped result.
- The first full harness receipt records **308 passed, 1 failed, 1 ignored**:
  the selectable-case count assertion saw 431 cases while its stale expected
  value was 430. The corrected-count retry in
  [`harness-tests-final-v2.json`](checks/harness-tests-final-v2.json) then
  records **314 passed, 0 failed, 1 ignored**. This is the broad candidate
  harness unit suite; it is distinct from the final focused ODS subset below.
- [`harness-ods-post-lint.json`](checks/harness-ods-post-lint.json) records the
  post-lint focused ODS harness subset as **12 passed, 0 failed, 0 ignored**.
  [`harness-doc-final.json`](checks/harness-doc-final.json) and
  [`owned-format-post-lint.json`](checks/owned-format-post-lint.json) also
  pass.
- [`harness-strict-final-v2.json`](checks/harness-strict-final-v2.json) is the
  pre-fix warning-denied Clippy receipt. Its one candidate helper issue was
  corrected from `ordinal % 2 == 0` to `ordinal.is_multiple_of(2)`;
  [`strict-debt-check.json`](checks/strict-debt-check.json) then reports zero
  changed harness findings and 17 unique pre-existing groups, matching the
  retained 0429 debt baseline. The unscoped strict debt remains open.
- [`after-build.json`](checks/after-build.json) records the committed
  `f5bf1696192f0007db56554aa5b16719cdf1950b` release harness build as passed.
  [`after-profiles.json`](checks/after-profiles.json) records collection of the
  four after-profile jobs as passed; the six-job profile report verification
  also passes. Detailed profile values and interpretation remain deferred to
  the performance bundle review.

These receipts are focused evidence only. They do not imply that all workspace
tests, strict checks, or performance gates are green.

## Explicit open scope

- The tested artifact is generated ODS content. No LibreOffice, Microsoft
  Office, or other native external-application fixture is exercised, so native
  application compatibility remains unverified.
- The API is a fresh one-sheet creation path. Existing-package append,
  existing-document edit, and forward-only append-to-an-existing-member
  behavior remain outside this batch.
- `required_memory_bytes()` is a bounded reservation contract, not a claim
  about total process RSS, allocator peak, ZIP/auditor allocations, or a
  bounded-memory proof for every layer.
- No latency, throughput, allocation, RSS, or optimization claim is authorized
  by these functional receipts; the broader goal remains open pending the
  independent performance interpretation and remaining evidence.
