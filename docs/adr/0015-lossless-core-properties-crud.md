# ADR 0015: Lossless, schema-typed OOXML core-properties CRUD

- Status: Accepted
- Date: 2026-08-03

## Context

ADR 0014 moved the host-neutral core-properties reader into
`litchi-ooxml-common`, but the DOCX, PPTX, and XLSX authoring facades still
created a fresh metadata value instead of retaining package absence and the
validated value they had opened. Saving could therefore invent a core part,
rewrite metadata that the caller had not edited, or discard a noncanonical
target. The writer exposed a broad `DocumentProperties` value, normalized
dates through `chrono`, modeled `revision` as an integer, flattened keywords,
and wrote XML without validating its complete package graph.

The normative OPC schema makes several of those representations too narrow.
`revision` is `xsd:string`; `created` and `modified` use the W3CDTF union of
`gYear`, `gYearMonth`, `date`, and `dateTime`; `lastPrinted` is
`xsd:dateTime`; and `keywords` is ordered mixed content with optional
`xml:lang` on both the outer element and repeated `cp:value` children. A core
part may own relationships to extension parts. The schema does not define the
legacy `cp:contentType` element.

## Decision

`litchi_ooxml_common::properties::Props` is the concise owned semantic value.
Its public fields support Rust struct-update syntax, while short builders cover
the common construction path. `Props` models revision as an unnormalized
string, dates with validated lossless `time::W3c` and `time::DateTime`
newtypes, and keywords with a plain-string shortcut plus the ordered
`Keywords`, `keyword::Item`, `keyword::Value`, and `keyword::Lang` model. The
schema-faithful semantic value does not expose or emit the non-schema
`cp:contentType` element.

The common crate owns three explicit package operations:

- `read` returns `Result<Option<Props>>`, preserving absent versus present but
  empty metadata;
- `write` consumes `Props`, retains an existing target path, relationship ID,
  dialect, and legal extension relationships, and reports whether bytes
  changed;
- `clear` is idempotent and rejects ambiguous or shared inbound ownership
  before changing the package.

Read and non-destructive update permit a differently typed relationship to
also target the core part. Clear applies the stricter sharing check because
removing such a part would otherwise dangle or silently destroy another
owner's graph. Strict and Transitional relationship and root namespaces are
retained. Writes whose parsed `Props` values compare equal are no-ops, so they
preserve exact XML bytes and signatures; actual writes or removal invalidate
signatures. All input text, XML characters, language tags, lexical dates,
cardinalities, and retained-byte budgets are validated before the relevant
graph mutation.

DOCX, PPTX, and XLSX package facades own a hidden mutation-tracked `Slot`. For
core properties, they expose only `props`, `props_mut`, `put_props`, and
`clear_props`. `put_props` moves the new value in and returns the old value; a
mutable borrow is tied to the host lifetime. An untouched slot does not reparse
or rewrite metadata on save. The umbrella facade derives its generic
`Metadata` cache from that already validated value instead of reaching through
another crate or silently discarding parse failures. No compatibility aliases
for the old methods or type remain.

The XLSX host's late save phase uses the common slot's type-bound `Stage<'a>`
guard. Staging applies an edited value to the candidate OPC graph but retains
dirty intent. The guard's mutable lifetime makes the exact originating slot
the only value its consuming `commit` can clear; the host invokes that method
only after its sink and worksheet restoration succeed. Dropping the guard after
an error leaves the edit retryable.
The host takes a structural package snapshot only when writer materialization,
an edited core-property slot, or a staged worksheet overlay makes one
necessary. Built-in immutable part payloads remain shared through `Arc`. A
late failure restores the package snapshot, including producer application
properties and original worksheet payload owners. An unchanged opened
workbook writes directly without taking this snapshot. Extended application
properties are regenerated only when the mutable writer model is materialized,
not for a metadata-only or worksheet-overlay save.

## Consequences

- Core-property create, read, update, clear, absence, and no-op behavior share
  one common implementation across the three XML Office hosts.
- Legal reduced-precision and timezone-less dates, nonnumeric revisions, and
  multilingual mixed keywords remain representable without normalization.
- “Lossless” here means retention of modeled schema lexical values and exact
  no-op bytes. A semantic edit canonicalizes the core XML and does not retain
  comments, processing instructions, formatting, or prefix spelling.
- Ordinary callers use short safe methods and never manipulate part names or
  relationship IDs; focused lower-level operations remain available in the
  common crate.
- Lossless values necessarily own their retained strings. This slice makes no
  zero-allocation or performance claim; allocation and throughput changes need
  representative measurement.
- Reduced-precision W3CDTF values remain exact in `Props` but cannot always be
  projected into the chrono-based fields of the generic `Metadata` facade.
- PPTX no longer exposes an unconditional mutable raw-OPC reference. Its
  fallible `opc` boundary rejects pending presentation, core-property, or font
  state and rejects implicit plaintext exposure for an encrypted source.
  `edit_opc` mutates a structural candidate through a non-escaping closure;
  built-in payloads retain immutable `Arc` sharing, while custom `Part`
  implementations retain their own clone and interior-mutability policy.
  Error or unwind leaves the candidate unpublished. Successful commit requires
  one internal main-document relationship with an allowed PPTX content type,
  reloads the validated property slot, rejects raw automatic-font-policy
  changes, and disables the legacy writer. This boundary does not parse the
  complete PresentationML graph. PPTX stages writer-model materialization,
  optional font publication, and the property slot in one rollback package,
  committing the slot and clean writer state only after the sink succeeds.
  XLSX still exposes mutable raw OPC and restores its in-memory package and
  dirty intent only for failures in the late publication phase; mutations
  performed by its earlier writer-model materialization are not yet one
  transaction. DOCX still flushes its slot directly before later save work.
  Closing those remaining seams is follow-up work and is not hidden by the
  concise facade.
- The XLSX dirty path structurally clones package metadata and shares built-in
  immutable payloads. Custom `Part` implementations retain their own clone
  policy. This is a correctness boundary, not evidence of lower allocation or
  latency; representative measurements remain required before a performance
  claim.

## Verification

Focused common-crate tests cover Transitional and Strict graphs, canonical and
noncanonical targets, package absence, present-empty values, no-op byte and
signature preservation, extension relationships, shared-inbound clear safety,
structured multilingual keywords, all W3CDTF granularities, timezone-less
date-times, arbitrary revision strings, malformed values, resource limits,
failure atomicity, and no-unwind rejection. All 34 focused common tests pass.
Five host integration tests create, update, clear, save, and reopen DOCX, PPTX,
and XLSX packages through their public facades. Warning-denied Clippy and
rustdoc are green for the common and host crates.

The XLSX late-save amendment adds three focused regressions. One observes
shared payload ownership during an unchanged opened save and proves that the
snapshot fast path is not entered. One injects a sink failure, checks exact
core/application payload identity restoration and retained edit intent, then
successfully retries and reopens the serialized property. One supplies invalid
property text, proves the sink was never called and the package identity was
restored, then repairs and retries. The common slot test separately proves that
dropping a staged guard retains dirty intent and that consuming it clears only
its lifetime-bound origin.

The `core_props_office` example generated the artifacts under
`target/office-core-props`; its reproducible command is
`cargo +1.89 run -p litchi-ooxml --example core_props_office --all-features`.
The exact historical invocation and application versions were not recorded.
Using Computer Use in Microsoft Word, PowerPoint, and Excel on macOS, all six
authored and cleared artifacts opened without a repair prompt. Their native
properties dialogs showed the expected values before clear and blank values
afterward, and each document, slide, or worksheet rendered its expected
content. This supports open-and-inspect compatibility for those artifacts on
the tested desktop applications; it does not certify every Office version,
extension graph, or metadata lexical form. Office-side edit/resave and
reverse-read were not performed for this slice. Per explicit user direction,
no redundant manual full-workspace run was scheduled; the repository's
mandatory pre-commit hook subsequently ran and passed the workspace lib/
integration and doctest gates.
