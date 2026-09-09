# Change 0484: replayable bounded authored DOCX paragraph streams

Status: draft design, unsealed. This document defines the next implementation
slice after change 0483. It records APIs, ownership, proof, resource, and
evidence requirements; it is not an implementation or a performance result.

The target is an existing-document DOCX tail append of a caller-authored stream
of plain paragraphs. The stream must support 64, 256, and larger authored
paragraph counts while the source document size varies independently. The
operation must use one authenticated replay source for candidate validation,
OPC publication, and any durable forward patch. An immediate inverse must
restore the original ZIP artifact exactly. A durable exact inverse must require
an explicit original-source provider.

This slice remains separate from fresh streaming creation, adding a package
Part, and arbitrary repackaging. The one-paragraph API from change 0483 remains
supported and keeps its fixed-fragment path.

## Constraints and current substrate

The workspace goal requires streaming append memory to be bounded by an
explicit window rather than total output, while preserving the positional
`ReadAt` source, hierarchical execution budgets, sequential sinks, exact
no-ops, source-checked publication, and typed partial-output errors. The
relevant decisions are ADR 0001's strict public layers, ADR 0003's immutable
snapshots and source-checked reversible publication, ADR 0005's positional I/O,
explicit scratch and budgets, ADR 0006's preserve-first and fail-closed
validation, ADR 0008's migration and evidence gates, ADR 0010's archive
ownership boundary, ADR 0011's OPC ownership, and ADR 0024's current crate
topology. The full authored-stream continuation
and its distinction between replayable and one-shot producers are recorded in
[`change-0481/window-contract.md`](../change-0481/window-contract.md#full-goal-continuation-bounded-authored-streams).

The current path establishes a useful but deliberately smaller substrate:

| Owner | Current code | Limitation for this slice |
| --- | --- | --- |
| DOCX | [`source_backed/tail_append.rs`](../../../../crates/litchi-docx/src/source_backed/tail_append.rs) exposes `Package::tail_append_plain_paragraph`, `Edit`, `Plan`, and `Publication` (the API begins around lines 333–505). `prepare_edit` scans the source and candidate, encodes one paragraph, and hands a fixed fragment to OPC. | `Edit` owns one `Cow<str>` and `encode_fragment` retains one complete generated paragraph. `TailSpliceReader` accepts `&[u8]`, and the candidate proof has one generated offset/range. There is no event source or replay provider. |
| DOCX | The existing scanner (`scan_main_part`, `scan_reader`, and `validate_candidate`, around lines 2730–3526) owns Word grammar, namespace binding, final opaque `sectPr`, insertion offsets, and semantic source/candidate proof. | It must consume a bounded authored reader and prove a generated region containing many paragraphs without collecting that region. |
| OPC | [`source_backed/splice.rs`](../../../../crates/litchi-opc/src/source_backed/splice.rs) owns `SourcePartSpliceFragment`, `SourcePartSpliceProof`, `SourcePartSplicePlan`, physical preservation, compact audit, fresh replay, and exact immediate inverse. `SourcePartSplicePlan` currently stores an `Arc<Vec<u8>>` and replays one fixed slice. | The physical owner has no plan payload that opens a fresh authored reader. `SpliceAuditReader` and `stream_splice` currently take a fragment slice, so a very large authored stream would become one retained vector. |
| OPC | `SourceBackedPackage::allocate_source_part_splice_fragment` and the package-bound handoff around lines 602–710 reserve and transfer one fixed fragment owner. | The same ownership transfer must be generalized to a replay handle and a bounded per-reader window; a replay store must not be charged once per consumer or grow behind the budget. |
| OPC | [`SourceArtifact`](../../../../crates/litchi-opc/src/source_backed.rs) retains the exact positional source and can fingerprint or copy it in bounded chunks. `SourceBackedPackage::source_artifact` is an O(1) retained handle. | This is sufficient for an immediate inverse held in memory, but it is not a serializable durable provider. A durable inverse needs an explicit caller-resolved original source. |
| DOCX | [`streaming.rs`](../../../../crates/litchi-docx/src/streaming.rs) has checked paragraph/run events, escaping, fixed scratch, cancellation, and sequential-output error mapping. | `StreamingDocumentWriter` publishes a new package immediately and cannot revisit authored paragraphs, so it cannot serve as the replay source for an existing-document candidate and publication. |
| DOCX | [`source_backed/paragraph_copy.rs`](../../../../crates/litchi-docx/src/source_backed/paragraph_copy.rs) has an exact reversible model and durable byte format. | Its `Patch` stores complete `before` and `after` XML buffers (around lines 390–590); reusing it would make authored-stream memory proportional to the source or candidate. |
| ODF | [`SourceContentInsertionPlan`](../../../../crates/litchi-odf-common/src/core/source_publication/insertion.rs) demonstrates a source-prefix/fragment/suffix replay and validator callback. | `AuthoredXmlFragment` is still a complete `Vec`; it is a replay model, not the storage implementation for a large DOCX authored stream. |

The first implementation should extend these owners in place. `litchi-docx`
must not acquire a ZIP implementation dependency, and OPC must not learn
WordprocessingML event or section semantics.

## Decision

Add a format-owned authored event layer and a format-neutral OPC byte-replay
capability. The DOCX layer turns plain paragraph events into a sealed replay
handle. The handle is opened independently by the DOCX candidate scanner and
by each OPC source/candidate/publication pass. The same handle, or a durable
reference to the same provider, is retained by a forward patch.

The replay handle contains no complete candidate or source XML. It retains
only an authenticated authored proof, a provider handle, and any explicitly
charged replay-store reservation. A deterministic event source may be reopened
for every pass. A one-shot source is accepted only when its encoded chunks are
captured into an explicit caller-owned replay store. The operation never hides a
filesystem, network, temporary plaintext file, executor, or process-global
cache behind this abstraction.

The current fixed-fragment OPC methods remain as an adapter and regression
surface. The stream path gets a sibling constructor and payload variant rather
than silently changing the meaning or ownership of
`SourcePartSpliceFragment`.

## Format-owned authored API

The public DOCX API should expose plain paragraph events, not arbitrary XML.
The following names are concrete design targets; the final trait-erasure
details may change if they preserve the stated lifetimes and proof rules.

```rust,ignore
pub enum PlainParagraphEvent<'a> {
    ParagraphStart,
    TextChunk(&'a str),
    ParagraphEnd,
}

pub trait ParagraphCursor {
    type Error;

    // The returned text is borrowed from the cursor and is consumed before
    // the next call. No event transfers an owned String to the editor.
    fn next<'event>(
        &'event mut self,
    ) -> Result<Option<PlainParagraphEvent<'event>>, Self::Error>;
}

pub trait ReplayableParagraphSource {
    type Error;
    type Cursor<'source>: ParagraphCursor<Error = Self::Error>
    where
        Self: 'source;

    fn open<'source>(&'source self) -> Result<Self::Cursor<'source>, Self::Error>;

    // None means the source is replayable only for this process/handle.
    fn durable_reference(&self) -> Option<AuthoredReplayReference>;
}
```

The GAT form makes borrowed chunks explicit. An internal erased adapter may
turn it into a callback or `Read` implementation, but it must preserve the
borrowed lifetime and must not clone each text chunk. A cursor may reuse one
internal text buffer after the encoder consumes the previous event. A
`TextChunk` larger than `max_authored_chunk_bytes` is refused before the chunk
is retained; callers that own larger input must split it themselves.

Event state is strict:

* `ParagraphStart` is required before text and may not nest;
* `TextChunk` is valid only inside one paragraph and is validated as UTF-8 XML
  1.0 character data using the same `is_plain_text_character` policy as the
  current writer and one-paragraph encoder;
* `ParagraphEnd` closes exactly one open paragraph;
* the stream must finish with no open paragraph and at least one completed
  paragraph; an empty stream is a typed invalid-input refusal, not an implicit
  no-op; and
* empty text chunks may be admitted as no-output events only if they still
  count against the event limit and are represented consistently in the event
  proof. A `ParagraphStart`/`ParagraphEnd` pair produces a real empty `w:t`.

The encoder emits a self-contained paragraph for every event sequence:

```text
<w:p xmlns:w="<dialect namespace>"><w:r><w:t xml:space="preserve">
  escaped text chunks, split at the fixed output-window boundary
</w:t></w:r></w:p>
```

The exact bytes use the shared escaping behavior from
[`streaming.rs`](../../../../crates/litchi-docx/src/streaming.rs) and the
current `encode_fragment` prefixes/suffix. Each generated paragraph keeps its
own `xmlns:w` binding, so a source default namespace or a shadowed `w` prefix
cannot change the meaning of the authored bytes. `sectPr` remains source-owned
and is never emitted by the authored encoder.

## Authored replay and proof

The encoder writes to one bounded chunk sink. The sink updates a streaming hash
and forwards chunks either to a deterministic replay reader's pending buffer or
to a caller-supplied replay store. It never appends all authored paragraphs to
a `Vec`, `String`, or `AuthoredXmlFragment`.

```rust,ignore
pub struct AuthoredStreamProof {
    pub strict_namespace: bool,
    pub paragraph_count: u64,
    pub event_count: u64,
    pub text_bytes: u64,       // caller UTF-8 bytes
    pub encoded_xml_bytes: u64,
    pub event_sha256: [u8; 32],
    pub encoded_sha256: [u8; 32],
}

pub trait AuthoredReplayHandle: Send + Sync {
    fn proof(&self) -> AuthoredStreamProof;
    fn open(&self) -> Result<Box<dyn AuthoredReplayReader + '_>, AuthoredReplayError>;
    fn durable_reference(&self) -> Option<AuthoredReplayReference>;
}

pub trait AuthoredReplayReader: std::io::Read {
    fn finish(self: Box<Self>) -> Result<AuthoredPassProof, AuthoredReplayError>;
}

pub struct AuthoredReplayReference {
    // Opaque, bounded caller/provider token. It is not a path, lock, archive
    // ID, runtime handle, or digest-only substitute for replay bytes.
}
```

`AuthoredPassProof` contains the same counters and hashes as the sealed proof.
The reader returns it only after EOF and after its provider has checked its
version. Every consumer compares the pass proof with the sealed proof. The
encoded hash authenticates the exact generated bytes; the event hash and
semantic DOCX scan keep event determinism and paragraph grammar explicit. A
provider that changes between opens therefore fails as `AuthoredReplayChanged`
before output, even if the change is only an event-boundary change that happens
to encode to the same bytes.

There are two supported handle constructors:

1. A deterministic `ReplayableParagraphSource` is wrapped by a streaming
   encoder reader. Each `open` obtains a fresh cursor and emits encoded chunks
   through a fixed buffer. The first completed pass establishes the authored
   proof; later passes must match it. The source may remain caller-owned and
   does not become an unbounded library allocation.
2. A one-shot producer is run exactly once into an explicit
   `AuthoredReplayStore`. The store receives encoded chunks, has a finite
   retained-byte ceiling, and returns a sealed handle with independent readers.
   A memory store is useful for bounded tests but its retained bytes are charged
   and scale with the authored stream. A caller-owned positional/range store or
   encrypted temporary store can keep the operation's resident window bounded.

Conceptually the store surface is:

```rust,ignore
pub trait AuthoredReplayStore {
    type Handle: AuthoredReplayHandle;

    fn append(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError>;
    fn finish(self, proof: AuthoredStreamProof)
        -> Result<Self::Handle, AuthoredReplayError>;
}
```

The store is supplied explicitly to the one-shot constructor and is not
created from an ambient path. Its `append` operation must reject growth before
the chunk is copied. A handle owns or borrows one sealed store; it may not
silently reopen a different provider for publication. A durable reference is
optional and is issued only by a provider that can be resolved after the
preparation process exits.

## OPC byte-replay seam

Add a format-neutral `SourcePartSpliceReplay` capability beside the existing
fixed `SourcePartSpliceFragment` in
[`litchi-opc/src/source_backed/splice.rs`](../../../../crates/litchi-opc/src/source_backed/splice.rs).
It exposes only bytes and a compact length/hash proof:

```rust,ignore
pub struct SourcePartSpliceReplayProof {
    pub encoded_len: u64,
    pub encoded_sha256: [u8; 32],
}

pub trait SourcePartSpliceReplay: Send + Sync {
    fn proof(&self) -> SourcePartSpliceReplayProof;
    fn open(&self) -> Result<Box<dyn std::io::Read + '_>, SourcePartSpliceReplayError>;
}
```

The actual API may use an owned `SourcePartSpliceReplay` wrapper to keep the
provider error type out of `OpcError`; the ownership invariant is the same.
The wrapper must retain a bounded reader-window reservation and, for an
in-memory store, transfer one retained-storage reservation into the OPC plan.
It must not reserve the same storage again for every consumer. External or
durable storage remains caller-owned and is reported separately from the OPC
resident window.

Add a sibling such as
`SourceBackedPackage::prepare_source_part_splice_with_replay`. Its payload is
an internal enum or equivalent:

```text
Fixed(Arc<Vec<u8>>, existing reservation)    // change 0483 path
Replay(Arc<dyn SourcePartSpliceReplay>, one storage owner)
```

The fixed constructor and `SourcePartSpliceProof` remain unchanged. The replay
constructor uses a sibling proof or clearly documents that `fragment_len` means
the total encoded replay length; a new proof is preferable because it prevents
callers from assuming the plan owns all bytes. `SourcePartSpliceLimits` gains a
separate finite `max_authored_replay_window_bytes`; its existing
`max_replay_memory_bytes` continues to describe ZIP replay/compressor state.
If diagnostics use a distinct resource, add
`SpliceResource::AuthoredReplayMemoryBytes` rather than merging two owners into
one opaque number.

`SpliceAuditReader` becomes a prefix/replay/suffix reader whose replay side is
opened fresh for each pass and copied through the existing fixed adapter
buffer. `verify_candidate_proof` must drain and finish the replay reader before
accepting its encoded proof. `stream_splice` does the same during the physical
replay callback. The source and candidate hashes, lengths, and source-version
checks remain OPC-owned physical proofs. The DOCX semantic proof remains
DOCX-owned; generic XML validity is never a substitute for the DOCX candidate
scan.

The replay reader must be opened sequentially for the initial candidate pass,
OPC preparation, publication, and durable patch application. If a caller opts
into concurrent operations, every live reader window and any provider-internal
storage must be reserved separately under the same package execution context.
The initial implementation should keep the operation sequential and avoid
introducing a scheduler.

## DOCX stream plan and publication flow

Add a sibling module such as
`crates/litchi-docx/src/source_backed/tail_append_stream.rs`. It owns the event
grammar, encoder, authored proof, source/candidate semantic proof, and
format-owned patch vocabulary. It may reuse `validate_topology`,
`scan_main_part`, the settings admission, namespace helpers, and the existing
source-backed `Package`; it must not move physical ZIP logic into DOCX.

The format-facing shape is:

```rust,ignore
pub struct ParagraphStreamLimits {
    pub source: TailAppendLimits, // source, candidate, parser, OPC, output
    pub max_authored_paragraphs: u64,
    pub max_authored_events: u64,
    pub max_authored_chunk_bytes: u64,
    pub max_authored_text_bytes: u64,
    pub max_authored_xml_bytes: u64,
    pub max_replay_bytes: u64,        // retained store, if any
    pub max_replay_window_bytes: u64,
    pub max_patch_bytes: u64,
}

pub struct ParagraphStreamEdit<'package, S> { /* package + source/store */ }
pub struct ParagraphStreamPlan<'package> { /* semantic proof + OPC replay plan */ }
pub struct ParagraphStreamCommit<'package> { /* named consuming product */ }
pub struct ParagraphStreamPublication { /* candidate proof + exact inverse */ }

impl Package {
    pub fn tail_append_plain_paragraphs<S>(
        &self,
        source: S,
        limits: ParagraphStreamLimits,
    ) -> ParagraphStreamEdit<'_, S>
    where
        S: ReplayableParagraphSource;

    pub fn tail_append_plain_paragraphs_from_producer<P, R>(
        &self,
        producer: P,
        replay: R,
        limits: ParagraphStreamLimits,
    ) -> Result<ParagraphStreamEdit<'_, R::Handle>>
    where
        P: OneShotParagraphProducer,
        R: AuthoredReplayStore;
}
```

The second constructor may instead return a sealed replay handle directly and
take the producer during `prepare`; it must still run the producer once and
retain only the explicit store. `ParagraphStreamEdit::with_options` can reuse
the current additional cancellation-token scope. The source-backed package's
execution context remains the sole execution authority for semantic allocation,
OPC work, and output. No new public execution override or unmanaged fragment
lease is introduced.

Preparation is output-free and proceeds as follows:

1. Validate the finite limits and options, check the package context, reject
   changed-path encryption/signature/protection/dependency shapes through the
   existing `validate_topology`, and capture source version and main-part
   metadata. The explicit empty/no-op route remains separate; an empty authored
   stream is refused.
2. Run the existing source scanner through a verified decoded reader. Retain
   only `SourceProof`, the final-section flag/span digest, source length/hash,
   insertion offset, and bounded parser state. The source-size limit remains
   independent of authored limits.
3. Establish the authored replay handle. A deterministic source is encoded on
   its first pass into a fixed chunk reader. A one-shot producer is encoded into
   the explicit replay store. Count paragraphs, events, input UTF-8 bytes, and
   escaped generated bytes with checked increments; hash both event framing and
   encoded bytes; reject invalid event order or XML characters before the handle
   is sealed.
4. Open a fresh source reader and the same handle. Adapt the source prefix,
   encoded authored stream, and source suffix to the DOCX scanner. Replace the
   one-paragraph `(generated_offset, fragment_len)` assumption with a generated
   byte range and expected authored paragraph count. Prove the source grammar,
   exactly the authored number of additional direct paragraphs, the unchanged
   opaque `sectPr` span, the shifted insertion anchor, and the candidate
   length/hash. Finish the replay reader and compare its pass proof with the
   sealed proof.
5. Call the OPC replay preparation seam. OPC opens fresh source/replay readers
   to verify source and candidate raw proofs, compact audit policy,
   preservation metadata, ZIP limits, and replay/compressor workspace. This
   remains an independent physical check; the DOCX semantic scanner is not
   replaced by the generic audit.
6. Recheck source freshness, package execution state, authored proof, and all
   limits, then return a plan retaining the sealed replay handle and compact
   semantic/physical proofs. No caller sink has been touched.

Publication consumes the plan and opens fresh source and replay readers. The
OPC preservation owner raw-copies every untouched physical member and invokes
the replay callback only for the selected main member, preserving the selected
member's Store/Deflate mode and all permitted ZIP framing. The callback streams
source prefix, encoded authored chunks, and suffix through bounded buffers,
checks both proofs and the source version, then lets the physical owner finish
the archive. The candidate artifact fingerprint and the compact DOCX proofs
are retained in `ParagraphStreamPublication`.

The current OPC precedence rules remain mandatory: a source-version or archive
transport failure is primary over a secondary callback/provider error, a
cancellation or execution error remains typed, and any failure after accepted
output is wrapped as `IncompleteOutput` with the exact accepted count. A
provider error before the first sink byte remains a typed preflight/replay
error with an untouched sink. `Interrupted` must not create a retry loop around
a terminal cancellation/provider marker.

## Durable forward patch and exact inverse

The stream patch is a new format-owned type. It must not reuse
`paragraph_copy::Patch::to_bytes`, which serializes complete `before` and
`after` XML. A compact forward patch contains only:

* an operation/version tag and bounded selected `ParagraphStreamLimits`;
* source main-part name/profile, source decoded length/hash, insertion offset,
  source event/paragraph/depth facts, section length/hash, source archive
  fingerprint and archive length;
* the `AuthoredStreamProof`, including dialect, event and encoded hashes;
* candidate decoded length/hash and semantic candidate facts; and
* an opaque `AuthoredReplayReference`, when the provider can resolve the same
  sealed authored bytes after process exit.

The patch wire size is bounded by `max_patch_bytes`. It contains no authored
XML payload unless the caller explicitly selects a bounded inline replay store,
in which case those bytes count against `max_replay_bytes` and
`max_patch_bytes`. A digest without a replay provider is not enough to apply a
forward patch. A process-local deterministic handle may produce a live patch
but `to_bytes` must return a typed `NotDurable`/`MissingReplayProvider` refusal
unless it has a durable reference.

Applying a durable forward patch takes the current source-backed DOCX package
and an explicit resolver for `AuthoredReplayReference`. The resolver returns
the same kind of sealed handle; the application rechecks the provider token and
authored proof, rescans the source and candidate, and invokes the same OPC
replay plan. Missing, stale, or different provider data fails before output.
The patch does not trust its semantic counts without the candidate readback.

An immediate publication retains the current OPC
`SourcePartSplicePublication` and its exact `SourceArtifact`, so its inverse
delegates to the existing fingerprint-then-copy path. The inverse does not
need the authored replay stream: exact restoration is a copy of the original
physical archive, including untouched compressed members and ZIP metadata.

A durable inverse authorization must carry an explicit original-artifact
reference or receive one at application time. The format-owned API should keep
that reference out of ordinary DOCX values, for example:

```rust,ignore
pub trait OriginalArtifactProvider {
    fn open(
        &self,
        reference: &OriginalArtifactReference,
    ) -> Result<Box<dyn litchi_core::ReadAt>, OriginalArtifactError>;
}

pub fn apply_exact_inverse<W: Write>(
    current: &Package,
    inverse: &InverseAuthorization,
    original: &dyn OriginalArtifactProvider,
    sink: W,
) -> Result<()>;
```

The provider is explicit and may be a caller-owned file, range source,
encrypted store, or reopened source-backed package adapter. It must report a
stable version and length. OPC fingerprints the complete raw original archive
with a bounded buffer and compares its digest/length before emitting anything;
it then copies those exact bytes while checking both current-candidate and
original-provider freshness. A missing provider or digest mismatch is a typed
`MissingReplaySource`/`StaleSource` refusal with zero sink output. A change or
sink failure after output begins is typed `IncompleteOutput`. The main XML hash
alone never authorizes exact inverse restoration.

The immediate `Publication` path may keep the original `SourceArtifact`
privately as change 0483 does. The durable path must not serialize an archive
handle, raw lock, path, ZIP entry ID, or runtime object; it serializes only
bounded references and proofs and asks the caller to resolve providers.

## Resource and security contract

`ParagraphStreamLimits` must keep source, authored, replay, parser, candidate,
output, and patch ceilings separate. All values are finite and nonzero, and
every addition, multiplication, counter increment, and conversion is checked
before use. The relevant accounting dimensions are:

| Resource | Required accounting |
| --- | --- |
| Authored text | UTF-8 input bytes and text characters, charged as events are consumed. Caller-owned input capacity is not retained by the library; an owned capture store must report its retained capacity. |
| Authored XML | Escaped generated bytes, including every `xmlns:w`, `xml:space`, and paragraph envelope. `max_authored_xml_bytes` bounds the complete encoded replay length. |
| Authored events | Start, text-chunk, and end events, with separate authored paragraph and text-byte limits. The candidate count must equal source paragraphs plus authored paragraphs. |
| Replay storage | `max_replay_bytes` bounds a library-owned or in-memory store and is reserved before the one-shot producer runs. External/durable storage is caller-owned and must report its own retention/window for evidence; it is not silently called constant memory. |
| Replay window | One fixed encoder/provider read window per open handle, reserved through the package execution context. A concurrent reader needs another reservation. |
| DOCX scanner | Existing token, namespace, scope, depth, event, section, and hash state from `scan_reader`, charged before each parser is constructed. |
| OPC/ZIP | Existing XML audit, preservation-index, replay/compressor, source-reader, decompression, output, and archive limits from `SourcePartSpliceLimits` and the package context. Authored replay window is a distinct owner. |
| Work/objects | Checked work for each source/authored/candidate/OPC/publication pass and object charges for cursors, stores, readers, and event/paragraph state. The known authored proof permits a preflight estimate; a later budget failure remains typed and, after output, incomplete. |
| Durable patch | Reference/token bytes, proof bytes, and any explicit inline replay bytes are bounded before allocation. A token cannot cause an ambient provider or unbounded deserialization. |

The operation-owned payload window is bounded by the fixed source reader,
authored reader/encoder buffer, DOCX scanner workspace, OPC splice adapter,
hash state, and ZIP replay state. The full process working set also contains
the source package catalog, preservation index, compressor, caller sink,
caller-owned authored storage, allocator metadata, and any source cache. The
implementation may claim no complete main/candidate XML materialization and a
bounded payload/replay window only when those owners are reported separately.
It may claim a complete RSS window only after all package metadata and physical
writer owners are included in one enforced contract.

Cancellation checks are required before each pass, at bounded event/chunk
boundaries, after provider EOF, before OPC output, and at source/provider
freshness fences. The package execution context remains authoritative for full
OPC work and output. The optional DOCX cancellation token keeps the narrower
semantic/event and forward sink-callback scope established by change 0483.
There is no observer requirement hidden inside an arbitrary codec: a codec that
cannot be interrupted is fenced before and after its bounded pass, and its
known work is charged before invocation.

Security and preservation rules remain narrow:

* only plain text events are accepted; arbitrary authored XML, fields, breaks,
  MCE branches, external links, and active content remain refusals;
* generated names are locally namespace-bound and escaped with XML 1.0 rules;
* the source scanner still proves direct-body topology and final opaque
  `sectPr` placement, and the candidate scanner proves the exact appended
  region and unchanged section span;
* signature, encryption, protection, macro, external-relationship, and
  unsupported-dependency policy is reused from `validate_topology`; an empty
  authored stream cannot enter the exact no-op branch;
* a generic XML audit remains an independent physical check but cannot replace
  DOCX semantic readback;
* replay and original providers are explicit caller capabilities. Litchi does
  not create plaintext temporary files, use ambient networking, or install a
  global executor; and
* source and provider freshness errors retain typed provenance and source
  precedence. No guessed semantic removal may be used as an exact inverse.

## Smallest implementation sequence

The following order minimizes simultaneous changes and keeps the 0483 route
available for regression comparison.

1. **Freeze the contract and test fixtures.** Add no production code until the
   event state machine, proof fields, provider/reference semantics, and
   resource names above are accepted. Prepare deterministic fixtures for no
   `sectPr`, final opaque `sectPr`, Strict and Transitional/default namespace
   roots, and 64/256/larger authored streams over several source sizes.
2. **Add the OPC replay payload.** Introduce the format-neutral byte replay
   wrapper, replay proof, bounded reader-window accounting, and plan variant;
   adapt `SpliceAuditReader`, `verify_candidate_proof`, and `stream_splice` to
   open a fresh reader. Keep the fixed `Arc<Vec<u8>>` constructor unchanged.
   Prove provider ownership transfer, no double charge, source/candidate hash
   replay, typed callback errors, source-freshness precedence, and
   `IncompleteOutput` with synthetic replay providers before wiring DOCX.
3. **Add the DOCX event encoder.** Implement the plain event/cursor traits,
   fixed chunk encoder, event/encoded proofs, deterministic replay adapter, and
   one-shot store adapter. Reuse `append_character`,
   `escaped_character_len`, `is_plain_text_character`, package context, and the
   existing self-bound `w:p` envelope. Test state refusals, chunk boundaries,
   XML-character limits, escaped-size limits, proof equality, provider version
   changes, and store reservation release.
4. **Integrate semantic stream preparation.** Add the sibling DOCX stream edit
   and plan. Reuse topology/settings admission and source scanner state, then
   replace the single generated paragraph marker with an authenticated generated
   range/count proof. Scan source plus replay plus suffix through bounded
   readers, and call the OPC replay preparation only after DOCX candidate
   readback succeeds. Keep all failures before sink output.
5. **Publish and retain immediate inverse.** Have the consuming stream plan use
   the OPC replay publication, retain candidate artifact fingerprint and DOCX
   proofs, and delegate exact inverse to the retained original artifact. Add
   Store/Deflate, raw-member, short-write, cancellation, provider/source-change,
   and simultaneous-error tests.
6. **Add explicit replay and original providers for durable patches.** Define
   bounded provider references and resolvers. Serialize only compact proofs and
   bounded references; reject digest-only or process-local handles as durable.
   Re-run source/authored/candidate proof and OPC replay on forward apply. Add
   durable exact-inverse application that fingerprints and copies the explicit
   original provider, with missing/changed-provider and post-output failure
   cases.
7. **Measure only after the correctness gate.** Keep the current one-paragraph
   comparison intact, add stream route labels and pass-level accounting, and
   record the matrix below. No stream speed or memory claim is made from a
   single pilot.

## Correctness and adversarial matrix

The focused suite must cover:

* event order and replay: text before start, nested start, unmatched/endless
  paragraph, zero-event stream, empty paragraph, empty text chunk, invalid XML
  characters, oversized chunk, event/text/XML/replay-limit boundaries, and
  checked counter overflow;
* deterministic replay: three or more fresh opens see equal event and encoded
  proofs; a source that changes on the second open fails before output; a
  one-shot producer is invoked once and its sealed store is reused by candidate,
  OPC, and publication passes;
* source grammar and placement: 64/256/larger generated paragraphs, no section,
  one unusual final opaque `sectPr`, Strict/Transitional and default source
  namespaces, non-final/duplicate/nested/wrapped section properties, unknown
  body children, malformed namespace/QName/XML declaration, and unsupported
  markup; compare exact section bytes and shifted anchor;
* proof integrity: source, authored event/encoded, candidate length/hash,
  generated-range/count, section-span, source-version, and provider-reference
  mismatches all refuse with an empty sink;
* package security: encrypted, signed, macro-enabled, externally related,
  protected, tracked-revision, and unsupported-dependency sources refuse the
  changed path before payload publication; the explicit no-op behavior remains
  separately tested;
* physical replay: Store and Deflate target members, ZIP64/data descriptors,
  raw non-part members, central metadata/order, source range short reads,
  sequential short writes, sink failure, and exact source immutability;
* cancellation/error precedence: cancellation before each pass, while reading
  authored chunks, at candidate EOF, during OPC audit, during ZIP replay, and
  after accepted output; provider failure plus source mutation must retain the
  authoritative source error, and post-output failures must report exact
  `IncompleteOutput` progress; and
* durability/inverse: compact patch size independent of source and authored
  XML bytes, resolver success, missing/stale/wrong replay reference refusal,
  current-candidate fingerprint mismatch, wrong original-provider fingerprint,
  exact immediate inverse, exact durable inverse, and provider/source mutation
  after the inverse sink accepts a prefix.

The suite must also assert that no complete source or candidate XML buffer is
created in the stream path. Allocation tests should distinguish authored replay
storage from operation window and package catalog/index owners.

## Measurement plan

The 0483 corpus already provides deterministic source archives with 64, 8,192,
and 131,072 source paragraphs and an opaque binary member. Reuse those source
sizes, then cross them with authored counts of 64, 256, 4,096, and one larger
configured case. Use at least three authored text-size distributions: empty
paragraphs, short text, and text close to the configured authored-byte limit.
Vary text chunk partitioning independently (one chunk, 64-byte chunks, and
near-window chunks) and run both deterministic replay and explicit one-shot
stores. A memory store is a storage-cost arm; a caller-owned positional or
encrypted store is the bounded-resident-window arm.

Each arm should run against owned bytes, filesystem-backed positional input,
instrumented short-read/range input, and a configurable high-latency range
adapter. Use sequential sinks with several maximum write sizes. Keep Store and
Deflate selected-member cases separate. Source size and authored size must be
independent axes; a single proportional corpus cannot establish the claim.

Report, per route and per axis:

* p50/p95/p99 elapsed time and throughput in source, authored, candidate, and
  output bytes plus paragraphs per second;
* source range calls, requested/returned bytes, decoded/decompressed bytes,
  recompressed bytes, in-memory copy bytes, replay opens, replay chunks, and
  sink write-call size histograms;
* cycles, instructions, IPC, branches, cache misses, page faults, CPU
  utilization, and serial pass fractions where tools support them;
* allocation count, actual allocated/retained bytes, peak live bytes, and
  process RSS, with source storage, package metadata, replay storage, ZIP
  codec, caller input, and allocator metadata separated; and
* cancellation/refusal phase, typed error class, accepted output, and the
  exact proof fields for every retained sample.

The sealed 0483 comparison is limited to one authored paragraph: at a
131,072-paragraph source the bounded route took roughly twice the elapsed time of
the materialized route while its heap increment was 609,875 bytes versus
35,371,290 bytes (about 596 KiB versus 33.7 MiB). Those values are not a 0484
result and do not establish a stream
benefit. The stream design deliberately retains the repeated source/candidate
DOCX and OPC proof passes required by the contracts; their CPU cost must be
measured and reported by pass. A summed allocator “requested bytes” counter
must not be presented as physical allocation traffic when reallocations are
included; peak live bytes and allocation count remain separate.

The primary comparison labels should distinguish
`materialized_paragraph_copy`, `bounded_plain_text_tail_append`, and
`bounded_plain_text_stream_append`. The stream report must include whether the
authored provider was deterministic, memory-spooled, or external/durable, and
must not combine storage cost with operation-window memory. Formal claims need
isolated normal and allocator processes, fixed source manifests and hashes,
warm/cold state labels, confidence intervals, and the same source/output
oracles used by change 0483.

## Acceptance gate and open scope

M2 is accepted only when a source-backed DOCX stream can:

1. consume and validate at least the 64- and 256-paragraph authored vectors,
   with a larger explicit vector, while source size varies independently;
2. keep generated XML out of a source-sized or authored-stream-sized library
   `Vec` by using the same sealed replay handle for candidate, OPC, and
   publication passes;
3. preserve the DOCX semantic and opaque-section proofs and the OPC physical
   preservation/source-freshness/error contracts;
4. support a deterministic replay source and a one-shot producer only with an
   explicit bounded replay store;
5. publish to a sequential sink with typed preflight and partial-output errors,
   and restore an immediate publication exactly; and
6. apply a durable forward patch only through an explicit replay resolver and
   perform an exact durable inverse only through an explicit original-source
   provider.

Until this gate and its measurements pass, change 0483 remains the M1
one-paragraph bounded decoded-splice milestone. The full `docs/GOAL.md`
requirement still includes fresh streaming creation, other append meanings,
the complete Office CRUD matrix, and reproducible end-to-end evidence. This
design does not close those unrelated workstreams.
