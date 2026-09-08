# DOCX explicit-window tail append contract

Status: design review for change-0481. This document proposes the smallest
source-backed existing-document append capability. It is not an implementation
or a performance result.

## Decision

Add a format-owned source-backed operation for appending one plain paragraph to
an admitted DOCX body. The operation must validate the complete source by
streaming it, validate the complete candidate as a source-prefix plus a bounded
authored fragment plus a source-suffix, and only then ask the OPC owner to
replay the splice into a preserved ZIP package.

The insertion point is:

* immediately before the final direct `w:sectPr`, when the body has one; or
* immediately before the direct `</w:body>` tag, when it does not.

The section properties are treated as an opaque source span. They are checked
for placement but are not parsed, normalized, or regenerated. All existing
physical ZIP members are copied through the existing preservation path. Only
the main-document member is regenerated, using its existing compression mode.

This is deliberately one operation. It does not turn the existing paragraph
copy API into a general streaming editor and it does not combine the four
different append meanings called out in [`docs/GOAL.md`](../../../../docs/GOAL.md):
creation from scratch, logical append, adding a package part, and arbitrary
repackaging.

## Why the current path is not an explicit-window operation

[`paragraph_copy.rs`](../../../../crates/litchi-docx/src/source_backed/paragraph_copy.rs)
owns the useful source checking, plain-paragraph grammar, typed refusals,
source-checked patching, and exact immediate inverse behavior. Its snapshot
loads the complete `word/document.xml` into an `Arc<Vec<u8>>`, retains a range
for every paragraph, and creates a complete `after` XML buffer for the patch.
Its readback validation scans that complete candidate. It also currently
rejects every `w:sectPr`, so its body-end insertion point cannot preserve a
normal final section-properties element.

[`streaming.rs`](../../../../crates/litchi-docx/src/streaming.rs) has useful
bounded escaping, limits, cancellation, and sink error conventions, but it
creates a new package. It cannot splice into an existing decoded member or
preserve the existing package's physical members.

The OPC source owner already has the other half of the contract:
[`source_backed.rs`](../../../../crates/litchi-opc/src/source_backed.rs) can
verify and stream a decoded member, retain a positional source artifact, copy
untouched members through `PreservationIndex`, and copy the exact original
source for a no-op or an immediate inverse. Its existing overlay entry points
accept a complete replacement `Vec`/`Arc<Vec>`. There is no public plan for a
decoded insertion offset whose replacement is produced during replay.

The source-backed ODF insertion path is the closest existing design. Its
bounded splice, candidate pass, and replay pass in
[`source_append.rs`](../../../../crates/litchi-odp/src/package/source_append.rs)
and [`insertion.rs`](../../../../crates/litchi-odf-common/src/core/source_publication/insertion.rs)
should be reused as a model, with the physical replay owner moved to OPC for
DOCX.

## Proposed format-owned surface

Names are illustrative; the ownership and invariants are the required part.
The new code belongs beside the existing source-backed paragraph operations,
for example in `litchi-docx/src/source_backed/tail_append.rs`.

```rust
pub struct TailAppendLimits {
    pub max_source_xml_bytes: u64,
    pub max_append_text_bytes: u64,
    pub max_append_xml_bytes: u64,
    pub max_paragraphs: u64,
    pub max_events: u64,
    pub max_depth: u32,
    pub max_output_bytes: u64,
    pub max_scratch_bytes: u64,
}

pub struct TailAppendOptions<'a> {
    pub limits: TailAppendLimits,
    pub execution: &'a ExecutionContext,
    pub cancellation: Option<&'a CancellationToken>,
    pub verify_payloads: bool,
    pub scratch: Option<&'a mut ScratchLease>,
}

pub struct SourceBackedTailAppendEdit { /* source owner + one authored text */ }
pub struct TailAppendPlan { /* source proof + OPC decoded-splice plan */ }
pub struct TailAppendPublication { /* candidate proof + original SourceArtifact */ }

impl SourceBackedTailAppendEdit {
    pub fn prepare(self, options: TailAppendOptions<'_>)
        -> Result<TailAppendPlan, TailAppendError>;
}

impl TailAppendPlan {
    pub fn write_to<W: Write>(&self, sink: W)
        -> Result<TailAppendPublication, TailAppendError>;
}
```

The smallest authored input is one caller-owned UTF-8 plain-text paragraph.
The encoder emits one `w:p` containing one `w:r` and one `w:t`, escaping text
incrementally. It rejects invalid XML characters and structural markup rather
than silently manufacturing runs, breaks, tabs, or fields. A bounded
`AuthoredXmlFragment` may be used internally, but it must be charged against
`max_append_xml_bytes`; it must not become a complete source or candidate XML
buffer.

The edit may also represent no operation. A no-op plan must use
`SourceArtifact::write_to_stream` and reproduce the exact source bytes,
including compression, ZIP metadata, signatures, and opaque members admitted by
the existing topology policy. An empty text value is not implicitly a no-op: it
is a paragraph containing an empty text node if the format contract admits it.

The plan's proof should retain, at minimum:

* source version/lineage and the source artifact fingerprint;
* main-part name, decoded source length, source XML digest, and insertion offset;
* source paragraph/event/depth counts and whether a final `w:sectPr` exists;
* authored-fragment length and digest;
* candidate decoded length and digest, plus the resulting artifact proof;
* the selected source and output limits.

Proof values are compact scalars. They must not be implemented as the current
full `before`/`after` XML patch.

## DOCX admission and section placement

The source pass should preserve the existing plain-document admission rules and
refusal types from `paragraph_copy.rs`: ordinary Word main document, supported
content type, no encrypted/signature or macro infrastructure for a changed
publication, no external or unsupported relationships, no enforced protection
or tracked-revision mode, and the existing strict/transitional namespace and
plain-paragraph grammar. The package graph/topology checks should be factored
so they inspect catalog metadata and bounded settings data without first
materializing the main XML.

The main XML stream accepts only the current canonical direct structure, with
one deliberate extension:

* `w:document` contains `w:body` in the expected namespace;
* the body contains direct plain `w:p` children, with the current `w:p` →
  `w:r` → `w:t` restrictions;
* it may contain one direct `w:sectPr`, and that element must be the final
  non-whitespace body child;
* a `w:sectPr` anywhere else, more than one `w:sectPr`, a non-final
  `w:sectPr`, or any unknown body child is a typed refusal.

The scanner records the decoded byte offset at the start of the final
`w:sectPr`, or at the start of `</w:body>`. It retains no paragraph range
vector and no raw section-properties buffer. The original bytes before and
after that offset remain in the source member and are replayed verbatim. This
gives the generated paragraph the required placement while preserving every
section-property attribute, child, namespace declaration, and lexical choice.

Wrappers, tables, comments, processing instructions, CDATA, doctypes, MCE
markup, arbitrary direct body children, malformed namespace declarations, and
non-plain paragraph content remain refused. Supporting semantic edits to
section properties is a separate capability and must not be smuggled into this
append operation by weakening the scanner.

## Two-pass source and candidate proof

`prepare` is a no-output operation. It should use
`PartView::with_verified_decoded_reader` with a fixed decoded read buffer and a
streaming namespace/event scanner:

1. Check cancellation, source version, package topology, limits, and the
   caller's memory reservation before reading payload bytes.
2. Scan the complete main member. Track only the XML state stack, bounded token
   text, hashes, counters, and the insertion offset. Hash and count the source
   as it is consumed, then verify the decoded member checksum/length through
   the OPC reader.
3. Encode the one authored paragraph into a bounded fragment. Reuse the text
   escaping behavior and limit/error conventions from the streaming writer,
   without constructing a new DOCX package.
4. Reopen a verified source reader and scan a `SplicedReader` consisting of
   source prefix, fragment, and source suffix. The candidate validator must
   prove well-formedness, the same admitted grammar, exactly one new paragraph,
   the expected paragraph count, and an unchanged final `w:sectPr` span. Hash
   and count the candidate while it streams.
5. Recheck source version, cancellation, decoded lengths, and all limits before
   returning the plan. No sink write is permitted before this point.

The scanner needs decoded offsets, not compressed ZIP offsets. If the current
XML reader cannot provide reliable spans through a `BufRead`, add a small
offset-counting wrapper or event-span owner in the format layer. Do not derive
the splice offset from normalized or reserialized XML. The source pass must
finish before candidate validation and candidate validation must finish before
publication, so a source mutation is a typed stale-source failure rather than a
mixed candidate.

`write_to` performs a fresh replay pass. It rechecks the source version and
artifact proof, opens a new verified decoded reader, and streams the same
prefix/fragment/suffix while checking source and candidate digests and lengths.
The sink receives no bytes for a failed preflight. A cancellation, source
mutation, short sink write, or sink error after output has started returns the
existing typed partial-progress/output error; the sequential sink remains the
caller's atomic-publication boundary.

## OPC substrate that is missing

Add a generic OPC owner for one decoded-member splice. It should be a narrow
public abstraction such as `SourcePartSplicePlan`, not a DOCX-visible wrapper
around `PreservationIndex`, ZIP entry IDs, archive locks, or compression
internals. The plan needs:

* source package and part identity;
* decoded insertion offset and source decoded length;
* a bounded immutable authored fragment or a replay producer;
* source/candidate lengths and digests supplied by the format validator;
* source version/freshness and cancellation fences;
* output and replacement limits; and
* `write_to`/no-op replay methods with typed sink progress.

Internally it can call the existing `PreservationIndex::write_replacing_with_replay`
path. The callback should open a fresh verified source reader, stream the
decoded prefix, write the bounded fragment, stream the suffix, and verify the
proofs. The preservation publisher must raw-copy every non-selected physical
member, including non-part members, and regenerate only the selected main
member with its original Store/Deflate choice. ZIP64, data descriptors, extra
fields, central-directory ordering, and unsupported layouts must retain the
current OPC checks.

The exact no-op branch must bypass regeneration and copy the retained
`SourceArtifact`. A changed branch must refuse the existing signature policy
before output. This keeps physical preservation in OPC and XML grammar and
section placement in DOCX. The ODF insertion owner is a useful replay model,
but the DOCX API must not depend on ODF types.

## Memory and caller scratch

The operation-owned decoded XML working set can be bounded by:

* a fixed reader buffer;
* XML state proportional to maximum nesting depth;
* bounded text/token scratch;
* the authored fragment, capped by `max_append_xml_bytes`; and
* scalar hashes, counters, and proof state.

That is a bounded payload-window claim. It is not yet a complete RSS claim.
`ExecutionContext` currently supplies reservations and cancellation, but it has
no generic public caller-owned scratch lease. The streaming writer's 64-byte
stack scratch and the OPC replayer's internal ZIP buffer do not constitute a
complete-window API. `PreservationIndex` and compressor state can also consume
memory proportional to package metadata and the selected output codec.

For the strict goal of a caller-controlled complete window, add a narrow core
`ScratchLease`/`ScratchBuffer` abstraction. The caller supplies the reusable
buffer, its capacity is reserved through `ExecutionContext`, and the OPC replay
owner uses it for fixed buffers and transient compression work. Insufficient
capacity must fail before output with a typed limit/allocation error. If the
preservation index itself cannot fit the agreed window, it needs a bounded
streaming/index representation or explicitly documented package-metadata
overhead; silently calling the materialized path an explicit-window operation
would be incorrect.

Until that substrate exists, implementation may claim only “no complete main
XML or candidate XML materialization; bounded decoded splice plus package
metadata.” It must not claim that all peak memory is proportional to the
paragraph window. Any caller text or supplied fragment capacity must also be
included in the documented accounting.

## Publication and inverse

An immediate publication can retain the existing `SourceArtifact` handle and
the candidate artifact fingerprint. Its inverse must:

1. fingerprint the current candidate artifact and require an exact match with
   the artifact this publication produced;
2. recheck source context and the candidate proof; and
3. copy the original `SourceArtifact` to the sink byte-for-byte.

This restores the original compressed main member and every untouched physical
member, including bytes that a regenerated Deflate member could not reproduce.
Foreign, stale, or tampered candidates must fail before output with the
existing source/conflict refusal.

The current durable `Patch::to_bytes` format cannot be reused: it stores full
before/after XML and therefore defeats the window. A compact durable inverse
recipe is possible only if it requires an explicit replayable original source
at application time. Such a recipe may store the original artifact fingerprint,
source version/lineage, main-member digest/length, insertion offset, section
flag, authored-fragment digest/length, and published candidate fingerprint. An
inverse application must receive the original `ReadAt`/`SourceArtifact` (or an
explicit caller-owned persistent replay spool), verify both source and
candidate proofs, and then copy the original source. Missing or changed replay
source is a typed `MissingReplaySource`/`StaleSource` refusal with no output.

Without that retained source/provider, a recipe can at most remove the last
paragraph semantically and regenerate the package. It cannot be called an exact
durable inverse. A separate semantic inverse would need a different name and
contract.

## Validation and replay matrix

The implementation should add focused tests before making a performance claim:

| Area | Required cases |
| --- | --- |
| Source window | Large source-backed documents through file and range-like `ReadAt`; no complete main XML allocation; fixed read ranges and reservation failures at below/exact/above limits. |
| Placement | No `sectPr`; one final `sectPr` with unknown attributes/children and unusual lexical bytes; insertion immediately before its exact source bytes; non-final, duplicate, nested, or wrapped `sectPr` refusals. |
| Grammar | Strict/transitional admitted roots; malformed XML, namespace errors, wrappers, tables, MCE, comments, PI/CDATA/doctype, complex runs, unknown body children, and unsupported dependencies produce typed refusals. |
| Proof | Source and candidate decoded length/hash, paragraph/event/depth limits, fragment limits, source mutation during each pass, and cancellation before output. Candidate has exactly one generated paragraph and the original section span. |
| ZIP replay | Store and Deflate selected members; ZIP64/data descriptors/extra fields; non-part members; unchanged local headers, compressed payloads, central metadata, and ordering byte-for-byte; selected main member only regenerated. |
| Output | Exact no-op equals source; deterministic replay across sink chunk sizes; cancellation, short writes, and sink failure report typed partial progress; no preflight failure writes bytes. |
| Inverse | Immediate inverse restores exact source; foreign/tampered candidate is rejected; compact recipe round-trips and works only with the original replay source; missing/changed replay source is refused. |

The physical ZIP assertions should compare raw member records, not only reopened
semantic packages. The section test should compare the exact byte span of
`w:sectPr`, including whitespace and namespace spelling. Allocation tests
should distinguish decoded payload memory from package catalog/index overhead.

## Sequenced next code steps and acceptance gate

1. Factor the existing DOCX package-topology checks so metadata and bounded
   settings reads do not materialize the main XML. Preserve all current refusal
   variants and signature/macro/external-dependency policy.
2. Extract a shared bounded plain-text paragraph encoder from the streaming
   writer. Make it usable by an authored-fragment producer and charge its
   output to the append limit.
3. Implement a DOCX source/candidate scanner with an explicit final-`sectPr`
   state and decoded insertion offset. Add source freshness, cancellation, and
   proof checks before any publication API is called.
4. Add the generic OPC decoded-splice plan around the existing preservation
   replay owner. Keep source-artifact no-op and unchanged-member raw-copy paths
   exact, and expose typed source/sink/limit failures.
5. Add the caller scratch lease, or document and test the narrower bounded
   payload-window guarantee before exposing a complete-window claim. Do not
   advertise complete-window RSS until ZIP replay/index/compressor allocations
   are included in that contract.
6. Add the validation matrix above, then add measured read-range, allocation,
   peak-memory, and sequential-output results in a separate performance result.

The acceptance gate is a source-backed one-paragraph append that passes the
complete source and candidate proofs, inserts before a final opaque `w:sectPr`,
preserves untouched compressed members exactly, supports exact no-op and
immediate inverse publication, and refuses durable exact inverse without an
explicit replay source. Until all of those conditions hold, the existing
materialized paragraph-copy implementation remains the only supported path and
must not be relabeled as the explicit-window capability.

## Full-goal continuation: bounded authored streams

The one-paragraph gate above is the first substrate milestone. It is useful for
proving the decoded splice and physical replay contracts, but it does not
fulfill the full goal's very-large creation/append/window requirement. A
one-paragraph API cannot establish that memory stays bounded when an append has
64, 256, or many more authored paragraphs, and it cannot make a one-shot
producer reusable for candidate validation, publication, and a durable patch.

The next format-owned layer should accept a replayable authored event stream.
The event stream must be structured as plain paragraph input rather than
arbitrary XML so that the DOCX encoder retains grammar ownership:

```rust
pub enum PlainParagraphEvent<'a> {
    ParagraphStart,
    TextChunk(&'a str),
    ParagraphEnd,
}

pub trait ReplayableParagraphSource {
    type Error;
    fn open(&self) -> Result<Box<dyn ParagraphCursor + '_>, Self::Error>;
}

pub trait ParagraphCursor {
    fn next(
        &mut self,
        emit: &mut dyn FnMut(PlainParagraphEvent<'_>) -> Result<(), TailAppendError>,
    ) -> Result<bool, TailAppendError>;
}
```

The exact trait shape may change to avoid allocation or object-safety costs.
The invariants must not change: each `open` creates an independent pass over
the same authored sequence; text chunks are bounded borrowed input; paragraph
boundaries are explicit; and the producer cannot mutate the sequence between
passes without causing a proof mismatch. A one-shot iterator is suitable for a
fresh streaming-creation API only when that API has no later validation or
replay pass. It is insufficient for existing-document append.

The encoder should consume those events into a fixed-size chunk sink. Once a
chunk has been passed to the validator, hash, or replay store, its buffer is
released and reused for the next chunk. It must never collect all generated
`w:p` elements in a `Vec`, `String`, or `AuthoredXmlFragment`. The encoder keeps
only paragraph state, XML escape state, counters, and one bounded output chunk;
it flushes `w:p`/`w:r`/`w:t` boundaries through the same sink. A text chunk may
be split across output chunks, but a paragraph may not be silently split into
an invalid event sequence.

### One explicit replay for all required passes

Define a caller-owned `AuthoredReplay` abstraction with the following conceptual
operations:

```rust
pub trait AuthoredReplay {
    fn append_chunk(&mut self, bytes: &[u8]) -> Result<(), TailAppendError>;
    fn finish(self) -> Result<AuthoredReplayHandle, TailAppendError>;
}

pub trait AuthoredReplayHandle {
    fn proof(&self) -> &AuthoredStreamProof;
    fn open(&self) -> Result<Box<dyn Read>, TailAppendError>;
}

pub struct AuthoredStreamProof {
    pub paragraphs: u64,
    pub events: u64,
    pub text_bytes: u64,
    pub encoded_xml_bytes: u64,
    pub sha256: [u8; 32],
}
```

This is an ownership contract rather than a requirement to use these names.
The handle is the one explicit replay input shared by the candidate validator,
the OPC publication plan, and a forward durable patch. Each consumer opens a
new reader, verifies the stored proof, and releases its chunks as they are
consumed. The OPC layer should accept this replay handle through the generic
decoded-splice plan instead of retaining a complete replacement buffer.

There are two valid implementations:

* A deterministic `ReplayableParagraphSource` can be opened separately for
  every pass. The operation still records the authored proof from the first
  pass and requires all subsequent passes to produce the same counts, lengths,
  and digest.
* A non-replayable producer can write chunks into an explicit caller-supplied
  replay store during preparation. The store may be bounded memory, a caller
  owned seekable/range provider, or a durable provider. It must expose a fresh
  reader for each pass and must not hide an ambient filesystem, network, or
  process-global cache.

For an existing-document append, preparation must reject a one-shot producer
unless the caller supplies the second form. This prevents candidate validation
from consuming different content than publication. For a durable forward patch,
the replay handle or an authenticated durable replay token must remain
available after the preparation process exits; an in-memory handle alone is not
durable. The token stores the authored proof and provider identity, never just a
digest that cannot recreate the bytes.

### Counts, bytes, and scaling limits

Extend `TailAppendLimits` with separate authored-stream limits rather than
overloading the source-document limits:

```rust
pub struct TailAppendLimits {
    // Existing source and output limits remain in force.
    pub max_source_paragraphs: u64,
    pub max_source_events: u64,
    pub max_append_paragraphs: u64,
    pub max_append_events: u64,
    pub max_append_text_bytes: u64,
    pub max_append_xml_bytes: u64,
    pub max_replay_bytes: u64,
    pub max_scratch_bytes: u64,
    pub max_output_bytes: u64,
}
```

Every increment is checked before it is applied, including `u64` overflow.
The text-byte count is over authored UTF-8 input; the XML-byte count is over
the escaped generated stream; the replay-byte count covers retained encoded
chunks; and the event count bounds a producer that emits many tiny chunks. The
paragraph count includes only completed authored paragraphs. A producer that
ends with an open paragraph, emits text before `ParagraphStart`, or emits no
`ParagraphEnd` receives a typed format refusal. Limits are checked during every
pass and the proof must match across passes.

The 64-paragraph and 256-paragraph cases should be mandatory scaling vectors,
not hard-coded API ceilings. They should be tested with both small paragraphs
and paragraphs whose text approaches the byte limit. A large-stream case should
exercise an explicitly larger caller limit and a replay provider that releases
each flushed chunk. The heap working set should remain proportional to the
configured chunk/scratch window plus XML depth and package metadata; it may
scale in the explicit replay provider's storage, which is caller-owned and
separately accounted for. An in-memory replay store is a valid bounded test
fixture only when `max_replay_bytes` is charged; it is not evidence of constant
memory for an unbounded append.

The candidate pass becomes:

1. Open the source decoded reader and the authored replay reader.
2. Stream the source prefix, then consume and validate authored events while
   producing the generated paragraph chunks, then stream the source suffix.
3. Check the final section-properties span, source/candidate hashes and
   lengths, paragraph/event counts, and every configured limit.
4. Close both readers and release their buffers before returning the plan.

Publication repeats the same sequence from fresh source and replay readers.
The plan must compare the authored proof from the replay handle with the proof
captured during candidate validation before writing the first sink byte. This
also makes replay independent of sink write chunking.

### Durable patch and exact inverse at stream scale

A durable forward `TailAppendPatch` for many paragraphs must contain either the
bounded encoded authored stream inline or a reference to an explicit durable
replay provider. It may contain the compact source/candidate proofs already
described, but a digest alone is not replay data. Applying the patch must open
the original source provider and the authored replay provider, recheck source
freshness, source proof, authored counts/bytes/digest, and candidate grammar,
then invoke the same OPC decoded-splice plan. If either provider is absent or
changed, return a typed refusal before output.

The exact inverse remains independent of authored-stream size: an immediate
publication restores the retained original `SourceArtifact`, and a durable
inverse requires an explicit replayable original source/provider. A large
authored replay handle does not substitute for the original physical ZIP
bytes. The Deflate caveat and exact untouched-member requirements therefore
remain unchanged for 64, 256, and larger appends.

### Relationship to the full documentation goal

The bounded event/replay layer is the continuation needed to make existing
document append scale beyond one paragraph. It still remains distinct from a
fresh streaming-creation implementation, package-part insertion, and arbitrary
repackaging. The full `docs/GOAL.md` claim requires those append meanings to
remain separate while demonstrating very-large source and authored streams,
explicit caller limits and scratch, sequential non-seek output, source range
behavior, allocation/peak/RSS evidence, and exact reversible publication.

Consequently, the one-paragraph acceptance gate in this document must be
reported as milestone M1 for the reusable decoded-splice substrate. It cannot
be reported as completion of the very-large creation/append/window goal. M2 is
the bounded replayable authored stream with 64/256 and large-stream scaling
evidence; only after M2 and the corresponding fresh-creation work are measured
can the project make the full goal claim.
