# 0479 candidate mechanisms for the plain-paragraph copy path

This note records implementation candidates for the existing public
source-backed `plain_paragraph_copy` path. It is an audit artifact, not a
performance result and not an approval to change the path. The candidates are
conditional on the phase attribution and allocation evidence from change
0479. No mechanism below should be selected from intuition alone.

The current contract and measured-baseline boundaries are documented in
[`contract-audit.md`](contract-audit.md). The source references below use the
line numbers present during this audit. The candidates preserve the current
`Snapshot`/`Edit`/`Commit`/`Patch`/`Publication` surface unless a section
explicitly says that it is a separate future capability.

## Contract that candidates must preserve

The operation is a deliberately narrow exact-byte operation. A package must
pass package topology and dependency checks before the main story is scanned;
the main document must be the ordinary `/word/document.xml` part and its
direct body children must be plain `w:p` elements made only from direct
`w:r`/`w:t` content. The accepted parser records exact paragraph byte ranges
and the byte position immediately before `</w:body>` in a private `Layout`.
Sections, wrappers, tables, unknown markup, MCE, and other unsupported content
remain typed refusals. See
`crates/litchi-docx/src/source_backed/paragraph_copy.rs:39-59,1554-1718`
and the refusal fixtures in
`crates/litchi-docx/tests/source_backed_paragraph_copy.rs:330-473`.

`Snapshot` owns an `Arc<Vec<u8>>`, a `Layout`, the captured `SourceVersion`,
the whole-artifact SHA-256, and the caller's finite `Limits`
(`paragraph_copy.rs:218-243`). `Snapshot::edit()` creates an isolated edit by
sharing the base snapshot (`paragraph_copy.rs:248-263`). `Edit` accepts one
copy operation only; a second call returns the existing typed operations
limit (`paragraph_copy.rs:291-357`). A copy selects one source paragraph and
one source-order insertion slot. The tail slot is `before == paragraph_count`,
which inserts immediately before `</w:body>`; it does not select a streaming
or windowed mode.

`commit()` returns the projected snapshot and a reversible patch. The patch
retains exact before/after main-story bytes and the source identity
(`paragraph_copy.rs:335-424`). Applying it checks the artifact fingerprint,
the exact source XML, and the captured source version before producing a new
snapshot. A publication recaptures the current source, applies that exact
patch, and retains the current snapshot, target snapshot, original
`SourceArtifact`, published fingerprint, and inverse patch
(`paragraph_copy.rs:603-628,980-1067`). These retained values are part of the
observable lifetime and inverse-publication contract.

The source is immutable positional `ReadAt` input. Source fingerprinting reads
the complete artifact in bounded 64 KiB chunks while checking source identity;
exact restore writes that retained artifact back to the sink
(`crates/litchi-opc/src/source_backed.rs:2702-2789`). Source mutation, short
reads, output limits, cancellation, and non-atomic sink failures retain their
current typed behavior. A candidate may change ownership of immutable bytes,
but it must not weaken any of these checks or make a source change observable
only after publication has begun.

The governing accepted obligations are:

* ADR 0003 requires immutable cheap-to-share snapshots, isolated edits,
  atomic commit, a new snapshot, a reversible patch, and an unchanged source
  (`docs/adr/0003-snapshots-edits-and-patches.md:6-25`).
* ADR 0005 requires positional `ReadAt`, stable source identity, finite
  resource budgets, explicit scratch, sequential sinks, and measured evidence
  for optimization. It also keeps random-access edits separate from a
  tail-only streaming API (`docs/adr/0005-io-memory-and-performance.md:6-17,
  19-46,54-61`).
* ADR 0006 requires preservation of untouched bytes and lexical details,
  validation without mutation, and refusal before publication when the safe
  closure is not modeled (`docs/adr/0006-validation-security-and-compatibility.md:6-23,54-60`).
* ADR 0010 keeps physical archive mechanics below the format facade
  (`docs/adr/0010-facade-archive-ownership.md:14-27,58-75`), and ADR 0011
  keeps OPC ownership in `litchi-opc` without leaking ZIP types through
  ordinary public document APIs
  (`docs/adr/0011-ooxml-physical-package-ownership.md:15-35,42-52`). The
  candidates therefore use existing private metadata and the existing OPC
  shared-payload seam rather than a new format-level archive type.
* ADR 0024 records the current layer with `litchi-opc` as the physical OPC
  owner and `litchi-docx` as the concrete DOCX owner
  (`docs/adr/0024-current-topology.md:15-54`). A candidate must stay within
  those ownership boundaries.

## Current ownership and duplicate-work map

The table distinguishes a byte allocation that the current public contract
needs from a duplicate copy or reparse that a private implementation could
possibly remove. The word “candidate” does not mean that the operation is
currently safe to remove; each row has proof obligations below.

| Phase | Current owner or operation | What exists today | Candidate status |
| --- | --- | --- | --- |
| Source capture | `main.data().into_arc()` in `plain_paragraph_copy_snapshot_with_limits` (`paragraph_copy.rs:941-978`) | The complete main-story XML is retained in the snapshot and scanned into an index. | Required by `Snapshot::xml_bytes()`, exact source comparison, paragraph ranges, and patch construction. Reuse of an existing package cache allocation is a measurement-dependent ownership change, not an assumed elimination. |
| Source validation | `scan_document` (`paragraph_copy.rs:1563-1718`) | One complete bounded XML scan produces `Layout`; the package topology/dependency checks run before it. | Required once for a new source snapshot. Do not replace it with tail-only parsing while claiming the same refusal contract. |
| Edit splice | `copy_fragment` (`paragraph_copy.rs:1319-1369`) | Allocates a complete output `Vec`, copies prefix, exact source paragraph bytes, and suffix, then reparses the candidate and performs readback. | Complete output allocation is required by the current projected `Snapshot`; candidate layout derivation can remove the candidate reparse only after proof. |
| Live commit | `Edit::commit` (`paragraph_copy.rs:335-357`) | `before` and `after` `Arc` handles are cloned; their vectors are not copied. | Already sharing. Keep this distinction in any allocation report. |
| Live patch apply | `Patch::apply` (`paragraph_copy.rs:394-424`) | Checks source identity, clones all `after` bytes with `checked_clone`, then `Snapshot::with_xml` scans them again. | Shared `Arc` ownership and a validated-layout cache are candidates. The current full scan remains the safe fallback. |
| Durable decode | `Patch::from_bytes` (`paragraph_copy.rs:511-576`) | Copies wire `before` and `after` into owned `Arc<Vec<u8>>` values because the input slice may be temporary; scans both and validates the exact operation. | Input-to-owned copies are required for a durable patch's lifetime. The validation helper need not clone those same vectors again. |
| Durable validation | `validate_durable_shape` (`paragraph_copy.rs:1472-1511`) | Scans both byte strings, creates temporary snapshots, rebuilds a complete candidate with `copy_fragment`, scans that candidate, and compares its bytes to `after`. | The temporary snapshot owners currently perform extra `Arc` ownership setup but no payload copy; the rebuilt candidate allocation and scan are the substantive duplicate-work candidates. Exact segment comparison can replace the rebuilt allocation only with equivalent full-byte proof. |
| Publication target | `publish_plain_paragraph_copy_patch_to_stream` (`paragraph_copy.rs:980-1050`) | Captures current source, applies patch, then clones target XML from `Arc<Vec<u8>>` into a new `Vec` for the OPC overlay. | The clone is a direct shared-payload candidate. Current recapture and stale-source checks remain required. |
| OPC replacement handoff | `write_part_overlay_to_stream` (`crates/litchi-opc/src/source_backed.rs:7696-7703`) | Takes an owned `Vec` and wraps it in `Arc`; this transfers the vector into shared ownership. | The avoidable clone is at the DOCX call site. `write_part_overlay_shared_to_stream` already accepts `Arc<Vec<u8>>` and uses the same path (`:7705-7727`). |
| Changed publication | `write_single_part_overlay_to_stream` and `write_changed_overlays_with_appended_inner` (`source_backed.rs:7764-7829,9251-9578`) | Validates limits and replacement XML, preserves unchanged ZIP entries, and regenerates the selected member through a shared immutable payload. | Keep all source, XML, signature, cancellation, sink, and preservation steps. Sharing the target `Arc` does not permit skipping them. |
| Inverse retention | `Publication` retains snapshots, a `SourceArtifact`, and an inverse patch (`paragraph_copy.rs:603-628,1052-1067`) | The artifact is a shared source handle, not a whole-artifact memory clone. | Required for exact restore. Removing it would change the inverse contract. |
| Durable serialization | `Patch::to_bytes` (`paragraph_copy.rs:464-508`) | Allocates a canonical wire envelope containing both complete before/after payloads. | Required by the current owning durable wire API. A streaming or externalized patch format would be a new contract. |

The source artifact fingerprint is a bounded-buffer whole-source read, not a
whole-artifact clone. It is intentionally separate from main-story XML
ownership. Avoiding it would weaken the existing cross-member stale-source
check unless source identity semantics were changed and separately accepted.

## Same-contract candidate mechanisms

These mechanisms can be evaluated without changing the public snapshot or
publication types. They are hypotheses only. Implementing one requires a
phase/allocator attribution showing that the corresponding work matters and
the proof tests below.

### 1. Pass the target XML `Arc` directly to OPC publication

The changed publication branch currently does this:

1. `Patch::apply` returns a target snapshot whose `xml` is already an
   immutable `Arc<Vec<u8>>`.
2. `publish_plain_paragraph_copy_patch_to_stream` calls
   `checked_clone(target.xml_bytes(), "publication replacement")`.
3. `write_part_overlay_to_stream` wraps that newly allocated `Vec` in another
   `Arc` before building the shared replacement action.

The OPC layer already exposes
`write_part_overlay_shared_to_stream` (`source_backed.rs:7705-7727`). A
private DOCX call-site change could pass `Arc::clone(&target.xml)` to that
method. This preserves the current public API and leaves the target snapshot
alive through the publication plan. `Arc::clone` shares the vector; it does
not copy its bytes.

The shared OPC method enters the same generic helper as the owned method. It
keeps replacement limits, original-member loading, exact no-op handling,
signature refusal, XML validation, preservation planning, source checks,
cancellation checks, sequential sink behavior, and incomplete-output errors
(`source_backed.rs:7764-7829,9251-9578`). The mechanism therefore removes one
candidate payload clone without authorizing a skipped validation or a new
retention policy.

Proof obligations:

* Hold the target snapshot until the helper has finished and ensure no mutable
  access can obtain the vector through the shared `Arc`.
* Compare published bytes and reopened paragraph text with the existing first,
  middle, and tail cases. Compare every untouched ZIP member and inverse
  artifact digest with the existing preservation test.
* Exercise exact no-op, signed/refused, replacement-limit, malformed XML,
  source-version change, source-artifact change, cancellation, partial sink,
  zero-write sink, and incomplete-output cases. The shared seam must preserve
  the point at which output is first accepted.
* Attribute any observed allocation change to the replacement clone itself;
  do not count the OPC preservation index, source fingerprint, or target
  snapshot as eliminated by this mechanism.

Relevant existing tests are
`crates/litchi-docx/tests/source_backed_paragraph_copy.rs:159-192,
239-259,601-621,672-720`.

### 2. Share patch `after` bytes when applying a live or durable patch

`Patch::apply` currently verifies the exact source and then calls
`Snapshot::with_xml(checked_clone(self.after.as_slice(), ...))`
(`paragraph_copy.rs:394-424`). `Snapshot::with_xml` must scan the candidate,
but it does not need a new byte vector if the patch already owns immutable
`Arc<Vec<u8>>` bytes.

A private `Snapshot::with_shared_xml(Arc<Vec<u8>>)` helper could run the same
`scan_document` call and build a snapshot with `Arc::clone(&self.after)`. The
source version, artifact fingerprint, and limits still come from the checked
source snapshot. `Arc<Vec<u8>>` is immutable from the public surface, and the
patch retains an owner, so this does not expose a mutable alias. `Patch::apply`
would retain its current source-fingerprint, source-XML, and source-version
checks before constructing the target.

This candidate removes only the `checked_clone` allocation. It does not remove
the full candidate parse, the output-size check, or the source identity check.
It also applies to a rehydrated durable patch only after `from_bytes` has
already copied the wire payload into its owning `Arc`.

Proof obligations:

* Verify that every constructor of `Snapshot` preserves the same limits and
  identity fields; the helper must not silently inherit a patch's identity.
* Re-run live and durable forward/inverse application, exact XML equality,
  stale whole-artifact rejection, source-version rejection, clone/drop
  lifetimes, and no-op behavior. The canonical durable wire must remain byte
  identical.
* Ensure `Patch::inverse()` retains the correct shared before/after owners and
  direction. A target snapshot must not outlive the patch's data owner.
* Keep `Snapshot::with_xml(Vec<u8>)` for paths whose bytes are newly allocated
  or whose layout has not been proven. The shared helper is not a license to
  accept an unscanned candidate.

The durable and inverse assertions are in
`source_backed_paragraph_copy.rs:194-237`; position and failure-atomicity
assertions are at `:261-327`.

### 3. Retain a validated projected layout on a live patch

The projected snapshot already has a successful `Layout` after
`copy_fragment` scans and reads back the result. A live patch currently keeps
only before/after bytes and the operation; `Patch::apply` and
`Patch::effect_report` re-scan one side. A private layout cache can remove
those repeated reparses while preserving the public `Patch` shape.

One possible representation is an optional private `Layout` for each side.
`Edit::commit` supplies clones of the base and projected layouts. `Patch::apply`
uses the projected layout only after exact source bytes, artifact fingerprint,
version, and limits have passed. `Patch::inverse` swaps the private layouts
with the byte owners. `effect_report` can read the selected source range from
the cached layout instead of calling `source_layout()`.

Durable serialization must ignore the private cache. `from_bytes` can either
leave it absent and retain the current scan fallback, or populate it from the
layouts already built during durable validation. A patch loaded from bytes
must never trust a cache across a serialization boundary unless the wire
bytes were scanned in that same decode operation.

Proof obligations:

* Make the cache private and optional so every public result remains derived
  from exact bytes. On any missing or inconsistent cache, use the current
  scanner and return its existing typed error.
* Prove that the cached ranges belong to the corresponding byte owner and
  limits. A stale or mismatched layout must be treated as invalid internal
  state, not as a reason to accept a patch.
* Test repeated `effect_report`, live apply, inverse apply, durable decode,
  clone/drop, and publication. Compare all error classes and exact bytes with
  the scanner-backed path.
* Keep the parser as the authority for durable untrusted bytes. This candidate
  is a cache of successful validation, not a parser bypass for arbitrary
  input.

This is a range-index ownership optimization, not a source-tail mechanism.
`Layout.paragraphs` is already `Arc<[Range]>`; ordinary snapshot and edit
clones share that allocation (`paragraph_copy.rs:218-263`). The possible
duplicate is the repeated scanner construction, not the `Arc::clone` itself.

### 4. Derive the projected range index after the exact splice

`copy_fragment` knows all of the facts needed to construct the candidate's
paragraph ranges: the source layout, the selected source range, the insertion
offset, and the exact inserted fragment length
(`paragraph_copy.rs:1319-1369`). A private constructor could accept the newly
allocated output `Vec` and a derived `Layout` instead of calling
`Snapshot::with_xml`, avoiding a second XML parse for a candidate that was
assembled from already scanned source bytes.

For insertion slot `before`, the derived index would retain source ranges
before that slot, insert a range covering the copied fragment, shift source
ranges at or after the slot by the fragment length, and shift `body_end` by
the same length. It must use checked arithmetic and preserve the exact empty,
sole-paragraph, and tail-slot cases. The output bytes would still be fully
materialized because the current projected `Snapshot` exposes contiguous
`xml_bytes()` and is retained by `Edit`, `Commit`, `Patch`, and
`Publication`.

The current `validate_copy_readback` loop (`paragraph_copy.rs:1432-1458`)
should remain during an initial implementation. It checks the target paragraph
mapping against the source and preserves an additional exact-readback guard.
Only after a formal proof that prefix, inserted fragment, suffix, and derived
ranges are all coupled may a later change replace that loop with an equivalent
checked proof. The proof must still cover bytes outside paragraph ranges; the
durable path currently compares the complete rebuilt byte string
(`paragraph_copy.rs:1502-1510`).

Proof obligations:

* Derive only from a `Layout` produced by the accepted parser. Verify source
  and insertion ranges are in bounds, fragment bytes are exactly the selected
  source range, insertion is exactly a paragraph start or `body_end`, and all
  additions/shifts are checked.
* Preserve the parser's refusal and limit order. A complex section or hostile
  namespace must be refused while scanning the source; no derived layout may
  make an otherwise refused source appear valid.
* Test first, middle, and tail slots; zero-length and sole paragraphs; strict
  and ordinary namespaces; escaped text; output limits; paragraph/event/depth
  limits; and every existing readback/refusal fixture.
* Test durable shape validation with tampered before/after bytes. A valid
  derived layout for an in-memory edit cannot authorize arbitrary durable
  bytes.
* Keep the existing complete output length and `max_output_bytes` checks before
  allocation. A parser bypass must not turn a checked allocation into an
  unchecked one.

This candidate can remove a candidate reparse, but it cannot remove the
complete output allocation or whole-snapshot retention under the current API.

### 5. Share durable validation's already-owned before/after vectors

`Patch::from_bytes` first copies the durable wire slices into the patch's
`Arc<Vec<u8>>` owners, then calls `validate_durable_shape`. That helper scans
the slices but makes another `checked_clone` for each temporary validation
snapshot (`paragraph_copy.rs:1472-1490`). The helper can instead receive or
clone the already-owned `Arc<Vec<u8>>` values. The temporary snapshots would
share the decoded owners, and the final `Patch` would retain those same owners.

This is a private ownership change. It preserves the two scans, operation
direction, exact rebuild comparison, limits, and all durable error mapping.
It is therefore a lower-risk candidate than removing the durable parser or
the exact shape proof.

Proof obligations:

* Preserve the copy from the caller's wire slice: the returned patch must
  remain valid after the input buffer is dropped or mutated.
* Verify that temporary snapshots and the returned patch share immutable owners
  only; no mutable `Vec` can be obtained while an `Arc` is retained.
* Keep canonical wire bytes, malformed envelope rejection, limit rejection,
  operation-direction checks, and exact forward/inverse bytes unchanged.
* Exercise allocation-failure paths and ensure their resource labels remain
  the same where the existing implementation promises them.

### 6. Validate durable exact shape without rebuilding a second candidate

The durable validator already has parsed layouts for `before` and `after`, but
it calls `copy_fragment`, which allocates a complete rebuilt candidate,
reparses it, performs readback, and then compares the rebuilt bytes with the
wire `after` (`paragraph_copy.rs:1502-1510`). A private validator could prove
the same exact splice directly against the existing `after` slice:

1. Scan both wire payloads under the same limits, as today.
2. Check source and target paragraph counts and the operation bounds.
3. Compute the source insertion byte offset and selected fragment range with
   checked arithmetic.
4. Compare `after`'s prefix, inserted fragment, and suffix to the exact
   corresponding slices of `before`, including the complete byte length.
5. Apply the same target paragraph mapping/readback checks or an equivalent
   range proof before accepting the patch.

This can preserve full-byte equality without a second output `Vec`. It is
safe only if the comparison covers all bytes, not just paragraph ranges. The
existing implementation's exact byte comparison is the oracle. Error
precedence and malformed-byte behavior must be recorded before replacing it;
the candidate must not turn a malformed durable patch into a different
refusal, limit, or allocation result.

Proof obligations:

* Build a reference matrix from valid forward/inverse wires, no-op wires,
  wrong lengths, wrong source/insertion positions, altered whitespace outside
  paragraphs, altered namespace declarations, altered section/unknown markup,
  and truncated payloads.
* Compare the candidate validator's result class with `validate_durable_shape`
  for every matrix row. In particular, `InvalidDurable` must remain the
  external result for noncanonical but parseable shape mismatches.
* Preserve source-layout selection for inverse direction: an inverse patch
  applies the operation against its `after` bytes and compares with `before`.
* Keep `scan_document` as the authority for untrusted durable XML. This
  candidate removes a rebuild allocation, not validation of the wire payload.

### 7. Keep publication recapture and artifact fingerprinting intact

The current publication recaptures the package before applying a patch
(`paragraph_copy.rs:980-1003`). That work can look redundant when the caller
just committed an edit, but publication may be delayed and the package can
change between commit and publication. The source version alone is not a
complete artifact identity: the contract also rejects a package whose retained
member changed while the main XML remained equal. The `SourceArtifact` helper
therefore fingerprints the whole source with bounded `ReadAt` reads and
checks identity during the read (`source_backed.rs:2720-2758`).

No current same-contract candidate may remove these checks merely because a
phase profile attributes time to them. A cache would need a proof that its
identity is still current at the same publication boundary, including adapters
whose version behavior is not sufficient to establish whole-artifact equality.
Changing that contract belongs in an ADR/API decision. The same applies to
reusing the commit-time `Snapshot` in place of the publication-time snapshot:
it would change stale-source timing and could publish against a changed
package.

### 8. Boundaries that are already sharing, not duplicate XML clones

The following should be classified correctly before proposing a mechanism:

* `Snapshot::edit`, `Edit::commit`, `Patch::inverse`, and `Publication` clone
  `Arc<Vec<u8>>` handles. These are cheap owner-count operations, not payload
  copies (`paragraph_copy.rs:248-263,335-357,424-462,603-628`).
* `write_part_overlay_shared_to_stream` and the changed-overlay plan already
  retain shared replacement bytes through regeneration
  (`source_backed.rs:7705-7727,7812-7827,9420-9445`).
* `SourceArtifact::clone` retains an immutable source snapshot. It does not
  materialize a second complete source archive (`source_backed.rs:2702-2706`).
* The 64 KiB fingerprint/publication buffers are bounded scratch allocations,
  not a full XML copy. Their request pattern is part of the `ReadAt` and
  accounting contract.

Any 0479 phase counter should distinguish `Arc::clone`, a newly allocated
  `Vec` of the same XML length, parser range storage, and bounded I/O scratch.

## Range-index ownership and retention

`Layout` owns an `Arc<[Range]>` and a `body_end`; its `Range` entries point into
the owning snapshot's XML bytes (`paragraph_copy.rs:218-243`). This establishes
the following ownership rules:

* A layout cannot be shared across different XML owners unless the bytes are
  byte-identical and the ranges were validated against that owner. Sharing a
  range index from `before` with `after` would be invalid after insertion.
* Snapshot and edit handles may share the base `Arc<[Range]>` because they
  point to the same XML. A projected edit gets a new range array after the
  splice. The new array is a small index allocation even when the XML is
  shared.
* A live patch currently does not retain either layout. This is why
  `source_layout()` reparses the operation's source side for `effect_report`
  (`paragraph_copy.rs:433-462`) and why `apply` reparses the target through
  `with_xml`. An optional private cache can remove those reparses only after
  byte-owner coupling is proven.
* Durable decode must build a fresh layout from the decoded bytes. A layout
  derived from an earlier process or from a caller's temporary wire buffer is
  not reusable. Within one decode, the same `Arc<[Range]>` may be retained by
  validation and the resulting patch.
* `Publication` returns a target `Snapshot`, an original snapshot, and an
  inverse patch. Dropping any of their layouts would change methods such as
  `paragraph_count`, `xml_bytes`, inverse application, or effect reporting.

If range storage becomes a measured phase or allocation concern, a packed
private range representation could be considered. It would need checked
offsets, the same paragraph order, no changed limit behavior, and exact byte
slices for every existing readback. It is lower priority than removing the
known duplicate candidate XML vectors and scans, and it has no bearing on
whether a source-tail representation is authorized.

## Proof and test matrix

Every same-contract mechanism must be compared against the current behavior,
with the mechanism disabled as the reference. The minimum matrix is:

| Contract surface | Existing evidence | Required candidate proof |
| --- | --- | --- |
| Source snapshot and source order | `source_backed_paragraph_copy.rs:114-157,520-545` | Exact XML, paragraph count, first/middle/tail slot mapping, empty body, empty paragraph, and sole paragraph. |
| Publication preservation | `:159-192` | Reopened text, unchanged member bytes/order/compression where promised, and exact inverse artifact digest. |
| Durable patch | `:194-237` | Canonical forward/inverse wire, apply after input drop, foreign artifact rejection, tamper rejection, and exact bytes. |
| No-op and inverse | `:239-259` | Byte-identical source output, stale inverse writes no bytes, and retained artifact identity. |
| Position/operation/limits | `:261-327` | Failure atomicity, one-operation limit, paragraph/output/durable limits, and unchanged projected snapshot after failure. |
| Refusal boundary | `:330-473,547-570` | Dependencies, signatures, macros, hostile paths, structured paragraphs, sections, namespaces, protection, and unknown markup retain typed refusal. |
| Source mutation | `:601-621` | Version change and whole-artifact change fail before accepted publication bytes. |
| Sink behavior | `:672-720` | Partial writes complete, zero-write fails, and accepted-prefix failure reports exact incomplete output. |
| Readback | `paragraph_copy.rs:1432-1458` | Exact copied paragraph and untouched paragraph mapping; durable full-byte equality remains authoritative. |
| OPC validation | `source_backed.rs:7764-7829` | Replacement limits, no-op raw-copy path, XML audit, signature policy, cancellation, source checks, and overlay preservation remain identical. |

For each candidate, record at least the result/error class, source-read
request sequence, phase counters, allocation/copy counters, output bytes
accepted before failure, and retained object lifetime. A candidate is not
accepted because it produces the same happy-path text for one fixture.

## Separate question: an explicit-window source-tail API

The current retained snapshot contract cannot automatically become a bounded
source-window implementation. `Snapshot` exposes a contiguous complete XML
slice through `xml_bytes()`, owns a complete `Arc<Vec<u8>>`, and owns a complete
paragraph range index. `Edit` retains both base and projected snapshots.
`Commit` retains complete before/after patch payloads. `Patch::apply` requires
exact complete before/after bytes, and `Publication` retains the original
artifact and inverse material for exact restore. These requirements are
visible in `paragraph_copy.rs:230-243,291-424,603-628,980-1067`.

An implementation that reads a source prefix and a tail window, stores only an
inserted paragraph, and streams untouched bytes later would need a different
internal or public representation. If it eventually materialized a complete
`Snapshot` to satisfy `xml_bytes`, complete patch bytes, or inverse retention,
it would return to the current whole-document retention point. If it did not,
it would change the API and lifetime contract. Neither behavior can be
inferred from a tail insertion slot.

A future explicit-window capability would need its own accepted design and
proofs for at least:

1. **Representation and lifecycle.** Define whether the window object is a
   streaming writer, a tail transaction, or a deferred patch; which methods
   can inspect paragraph count/text; whether revisiting flushed content is
   forbidden; and how repeated append operations are represented. The current
   edit accepts one operation, so repeated append requires separate
   reopen/snapshot/edit/commit/publication lifecycles.
2. **Full refusal semantics.** Decide how package topology, dependencies,
   protection, signatures, namespaces, sections, and unknown markup are
   established when only a window is retained. A tail parser cannot infer the
   absence of a section or unsupported relationship from the tail alone.
3. **Source identity.** Retain stable `ReadAt` version checks and, when exact
   whole-artifact stale detection or inverse publication is promised, the
   complete artifact fingerprint or an equivalent identity protocol. Define
   the source-read ordering and short-read behavior.
4. **Publication and inverse.** Preserve raw unchanged ZIP members, XML
   replacement validation, limits, cancellation, non-atomic sink accounting,
   exact output bytes, and the ability to restore the original published
   artifact. A window representation cannot silently drop the current
   `SourceArtifact` retention requirement.
5. **Resource and privacy limits.** Bound the window, parser state, source
   ranges, inserted bytes, output, fingerprint scratch, and any explicit
   scratch capability. Do not introduce automatic plaintext temporary files.
6. **Verification and ADR review.** Add a separate public contract, refusal
   fixtures, source mutation tests, partial-sink tests, readback tests, and
   representative measurement. ADR 0005's distinction between random-access
   edits and tail-only streaming is a design constraint, not evidence that
   this existing snapshot path already has tail-window approval.

The 0479 source-first-to-tail measurement can identify whether the current
full XML candidate is a dominant phase. It cannot, by itself, approve this
new capability or establish that source-tail reads are safe for the existing
public contract.

## Decision gate for 0479

Wait for measured phase attribution and allocation/copy accounting. If a
candidate is justified, the lowest-risk order for review is:

1. pass the already-owned target XML `Arc` through the existing shared OPC
   method;
2. share patch `after` bytes through a private `Snapshot` constructor;
3. share durable validator owners and, separately, replace durable rebuild
   allocation with an exact segment proof;
4. consider derived or cached layouts only with the parser/readback matrix
   above.

Keep publication recapture, whole-artifact fingerprinting, complete projected
XML ownership, parser refusal, and inverse artifact retention in the baseline
until an explicit API/ADR decision changes them. No item in this note claims a
speedup, reduced peak memory, fewer `ReadAt` calls, or approval for a source
tail window.
