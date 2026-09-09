# DOCX replayable stream checkpoint

Status: correctness checkpoint for the 0484 stream design. The production contract review and integrated correctness gates are cleared with the source scopes recorded below; formal performance acceptance remains open. This note records the current API shape, proof boundaries, preservation rules, and development receipts that make the stream path a correctness enabler. It makes no timing, throughput, allocation, bounded-RSS, or constant-memory claim.

The broader work remains governed by [docs/GOAL.md](../../../GOAL.md) and the accepted ADR set in [docs/adr/README.md](../../../adr/README.md). This checkpoint does not close the full document-format goal.

## Public operation and replay lifecycle

The public source-backed path has two authored-input forms in [tail_append_stream.rs](../../../../crates/litchi-docx/src/source_backed/tail_append_stream.rs):

* `Package::tail_append_plain_paragraphs` accepts a `ReplayableParagraphSource`. Each pass opens a fresh cursor and emits the caller's paragraph events through the format-owned encoder.
* `Package::tail_append_plain_paragraphs_from_producer` accepts a `OneShotParagraphProducer` and an explicit `AuthoredReplayStore`. The producer runs once; the store retains the encoded replay needed by later validation and publication. `MemoryReplayStore` is the explicit bounded in-memory implementation.

The event grammar is deliberately small: `ParagraphStart`, borrowed `TextChunk`, and `ParagraphEnd`. An empty paragraph is a real paragraph, while an empty stream is refused. Text chunks are consumed before the next cursor call, so the borrowed chunk need only live for that part of the pass. A readable usage shape, adapted from the focused integration test, is:

```rust
let edit = package.tail_append_plain_paragraphs(source, limits)?;
let plan = edit.prepare()?;
let publication = plan.write_to_stream(&mut output)?;
```

The one-shot form is the same lifecycle with an explicit store:

```rust
let edit = package.tail_append_plain_paragraphs_from_producer(
    producer,
    MemoryReplayStore::new(limits.max_replay_bytes)?,
    limits,
)?;
let plan = edit.prepare()?;
plan.write_to_stream(&mut output)?;
```

These shapes are exercised by `source_backed_tail_append_stream.rs`; they are documentation of the existing API, not new example files.

For a replayable source, sealing authenticates one authored pass and records the event and encoded-XML proof. Preparation then uses fresh source and replay readers for the DOCX candidate scan and the OPC candidate scan. Publication opens another fresh reader. Each reader must reach EOF and call `finish` in that same pass; the returned pass proof is compared with the sealed proof before the plan or publication is accepted. A cursor or provider that changes its proof or durable reference between opens is rejected as `AuthoredReplayError::Changed`. `AuthoredReplayHandle` retains the immutable proof/reference, and `BoundReplayHandle` freezes and rechecks those values at the public binding boundary.

The deterministic ordinary route has a measurable lifecycle test: after `prepare`, the source is opened three times (sealing, DOCX candidate validation, and OPC candidate validation); after publication the total is five, with the two additional publication passes independent of the number of encoded chunks. See the five-open assertion in [source-backed-tail-append-stream tests](../../../../crates/litchi-docx/tests/source_backed_tail_append_stream.rs#L830). This is a lifecycle invariant, not a performance result.

The two replay choices have different ownership and retention semantics. The ordinary deterministic route uses a fixed encoded replay window for each live reader and does not retain a complete authored XML document in the plan. The one-shot route opts into an explicit store; `MemoryReplayStore` retains the encoded chunks under its selected replay bound and transfers one retained store reservation to the sealed handle. An external or durable store owns its own storage and is outside the package's resident-memory claim. Concurrent readers have their own live window reservation.

`CursorAccounting` accounts caller-declared cursor/provider bytes and object terms when a caller supplies them. The current public deterministic and one-shot adapters use the default zero declaration, so heap held by an unknown caller/provider is outside package accounting. The deterministic reader separately reserves its known inline state, including the concrete cursor size, before opening the cursor. The stream therefore must not be described as bounding unreported caller heap. The earlier contract-review findings about managed object leases, exact window capacity, resolved-reader context, and callback precedence are historical gates; the current review records them cleared. The formal performance evidence bundle remains unsealed.

## Source, authored, and candidate boundaries

The operation works with both ordinary `source_backed::Package::from_reader` packages and managed packages opened with `from_read_at_with_execution_context`. The package execution context is the authority for managed work, memory, and output accounting, and the source, authored, candidate, and OPC phases check cancellation/context at their phase boundaries. The stream route reads source parts through positional `ReadAt`; it does not call the document materialization route.

The limits are separate by purpose. `ParagraphStreamLimits` carries the existing source/tail-append limits and distinct ceilings for authored paragraph count, authored event count, text-chunk bytes, authored text bytes, generated XML bytes, replay retention, per-reader replay window, and compact patch bytes. The authored event count is a producer/encoder framing count. It is not the count of XML parser events in the source or candidate package. Source scanning and candidate scanning retain their own token, namespace, depth, section, and decoded-byte limits. This separation prevents a large source XML parse from being hidden inside an authored-event allowance.

The plain-text policy is also part of the public contract. The accepted XML 1.0 character ranges exclude CR, LF, TAB, and other disallowed controls in the streaming text path. The focused test `authored_text_uses_xml_space_and_rejects_streaming_writer_control_rules` covers these refusals. Arbitrary XML, fields, breaks, MCE content, or other authored markup are outside this API.

The generated paragraph is appended before the source-owned final section properties. The existing final `sectPr` remains opaque and is copied with the source section bytes. OPC publication preserves the rest of the package, including unknown ZIP members. The semantic stream proofs carry the source and candidate main-member lengths and SHA-256 digests over the exact decoded document XML bytes, before XML normalization; they are distinct from the archive proofs. The separate `ArtifactProof` values cover the complete physical ZIP archive length and SHA-256. The semantic checks separately verify paragraph counts, generated placement, namespace/depth facts, and section facts.

## Durable forward and inverse boundaries

The durable forward patch is a compact authenticated recipe. Its `StreamSourceProof` and `StreamCandidateProof` record source and candidate decoded main-member XML lengths and SHA-256 digests, source/candidate semantic facts, generated offset/length/count facts, and section facts. Their separate `ArtifactProof` fields record the complete source and candidate physical ZIP archive lengths and SHA-256 digests. The patch also carries the authored event/encoded-XML proof. It does not rely on a process-local source version or a digest alone, and it does not embed an unbounded authored XML payload. When an authored replay reference is required, it is opaque, bounded, and checked against the fresh handle returned by an explicit `AuthoredReplayResolver` during apply.

Applying a durable forward patch authenticates the current archive before asking the resolver for a replay provider. The newly resolved provider must reproduce the recorded authored proof and durable reference; the candidate is then revalidated through the same OPC replay seam. A stale current archive therefore fails before resolver output can affect the destination.

For publication with an expected candidate artifact, OPC first performs a private sequential preview into a sink and checks the expected complete archive length and fingerprint. This preview has zero external output effect; the actual publication pass writes to the caller's output once and keeps the normal partial-output semantics. See [opc-foundation.md](opc-foundation.md) and the expected-artifact path in [OPC splice](../../../../crates/litchi-opc/src/source_backed/splice.rs#L489).

Immediate inverse uses the source artifact retained by the publication. Durable inverse requires an explicit `OriginalArtifactReference` and `OriginalArtifactProvider`; the original provider is opened only after the current archive has been authenticated against the candidate artifact proof. A stale current archive is rejected before the original provider is opened or any output is written. Both current and original raw archive length/SHA-256 identities are checked before the exact OPC restoration. The inverse authorization and provider rules are exercised in [the durable stream tests](../../../../crates/litchi-docx/tests/source_backed_tail_append_stream.rs#L1120).

## Native consumer scope

The [dev56 native consumer probe](consumer/dev56/result.json) independently exported both source/candidate fixture pairs through LibreOffice and matched the complete UTF-8 text against Python ZIP/XML extraction. Its fixtures were generated by the earlier dev50 debug CLI: each source has 64 paragraphs and each candidate has 128; the near-limit candidate exports 3,933,248 text bytes. This is synthetic LibreOffice evidence, not a Microsoft Office compatibility result, and it does not exercise resource or error-policy changes made after the dev50 fixture generation.

## Development receipts

The receipts are separated into historical gates and the latest source-unchanged gates. They are correctness and validation evidence; they do not establish timing, throughput, allocation, bounded-RSS, or constant-memory results.

### Historical receipts

| Receipt | Result and scope |
| --- | --- |
| [dev51 focused DOCX tests](validation/docx-stream-tests-dev51.json) | The focused M1 stream suite ran 40 tests and the focused M2 stream suite ran 18 tests; all passed. |
| [dev42 stream unit receipt](validation/docx-stream-unit-dev42.json) | The retained unit receipt reports 8 passed and 0 failed. Its source-unchanged marker is not a final source-freeze claim. |
| [dev46 replay harness tests](validation/stream-harness-tests-dev46.json) | The replayable-tail harness reports 11 passed and 0 failed. |
| [dev50 debug CLI smoke](validation/stream-cli-smoke-dev50.json) | The source-count 64/authored-count 64, fixed-chunk 64 smoke exercised both `short` and the retained near-limit case and completed its semantic, exact XML, untouched-member, and inverse oracles. |
| [dev58 focused DOCX tests](validation/docx-stream-tests-dev58.json) | The post-fix focused receipt reports 40 existing DOCX tail tests and 21 stream tests passing. |
| [dev61 DOCX Clippy](validation/docx-clippy-dev61.json) | The historical warnings-denied Clippy gate passed. |

### Latest source-unchanged receipts

| Receipt | Result and scope |
| --- | --- |
| [dev68 all-feature DOCX tests](validation/docx-all-tests-dev68.json) | All-feature DOCX tests report 1,408 passed, 0 failed, and 31 ignored; the source manifest is unchanged for this receipt. |
| [dev69 no-default DOCX tests](validation/docx-no-default-tests-dev69.json) | No-default-feature DOCX tests report 1,389 passed, 0 failed, and 31 ignored; the source manifest is unchanged for this receipt. |
| [dev71 ASAN candidate2 fuzz smoke](validation/stream-fuzz-smoke-dev71.json) | Strict candidate2 reports 27 positive cases and two rounds of 10,000 mutations passed; its source manifest is unchanged. This is a smoke receipt, not a full fuzz-acceptance claim. |
| [dev72 DOCX Clippy](validation/docx-clippy-dev72.json) | All-feature, all-target Clippy with warnings denied passed; the source manifest is unchanged. |
| [dev73 DOCX rustdoc](validation/docx-rustdoc-dev73.json) | All-feature, dependency-free rustdoc with warnings denied passed; the source manifest is unchanged. |
| [dev75 measurement validator tests](validation/stream-measure-tests-dev75.json) | The seven focused Python measurement-validator tests passed; this validates sample and proof accounting and does not constitute a formal performance capture. |
| [dev76 scoped format check](validation/docx-stream-format-dev76.json) | The scoped rustfmt check passed after the root applied formatting-only wrapping/import-order changes. |

The runtime, ASAN, and earlier validation receipts precede the formatting-only source delta checked by dev76; their source manifests therefore do not provide an exact-byte source seal for the post-format tree. Dev76 is a formatting gate, not a rerun of those runtime or ASAN tests. The [dev77 crate-boundary check](validation/crate-boundaries-dev77.json) passed with unchanged source: 64 workspace packages, 239 internal dependency declarations, and 11 explicitly tracked iWork debt items.

The current contract review is cleared: the earlier context-binding, managed object-lease, exact-capacity, and callback-precedence findings are resolved in the live source. Those earlier findings remain historical gate context, rather than current blockers. The correctness checkpoint is ready for integration with the stated receipt and formatting scopes. The formal performance evidence bundle remains unsealed and incomplete.

The [measurement review](measurement-plan-review.md) identifies missing one-shot storage, input-adapter, sink-size, and compression arms. The new [store measurement plan](store-measurement-plan.md) defines the extension but remains a design proposal without captures. These formal performance and route arms remain required and open; dev75 only validates the measurement report checks. The successful debug smoke and native probe do not replace them.

The full docs goal remains open for the broader streaming, CRUD, preservation, and reproducible end-to-end evidence requested by [GOAL.md](../../../GOAL.md).
