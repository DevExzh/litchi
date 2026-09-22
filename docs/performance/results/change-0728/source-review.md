# 0728 source review: current OLE2 route and rendered-candidate ownership

Status: read-only source qualification, 2026-09-22. No production source was
changed by this review. No Cargo, native, or profiler command was run. The
measurement work for this change should time the current lifecycle; this file
does not claim a performance result.

## Authority and scope

The current policy is [0663](../../0663-cfb-sector-layout-policy.md), retained
and implemented in `litchi-cfb`, `litchi-ole-common`, `litchi-doc`, and
`litchi-ppt`. It implements decision 10 of [0652](../../0652-owner-decisions-for-the-third-wave.md):
source-backed OLE2 saves default to `SectorLayoutPolicy::Reuse`, reclaim cheap
free space, and append when the source layout cannot absorb growth. It keeps
the ordinary validation and fallback path. It makes no speed claim. In
particular, [0663](../../0663-cfb-sector-layout-policy.md#what-the-measurement-means)
records the 0617 estimate as withheld, says the editor still materializes a
final `Vec<u8>`, and describes `kept_sectors` as placement retention rather
than zero output writes.

[0617](../../0617-cfb-copy-through-writer-design.md) is historical design
material. It explains why a copy-through route was considered and describes
the old unresolved policy boundary; it is not the authority for current
sector placement or retention behavior. The implementation and qualification
below use 0663's landed `Reuse`/`Rewrite` policy and its validation gates.

The constraints from [GOAL.md](../../../GOAL.md) and ADRs 0005/0006 remain
binding: preserve by default, exact no-ops remain exact, edits and publication
are failure-atomic, validation is completed before publication, source
freshness and reversible patches are checked, and resource use is bounded.
ADR 0005 also says ordinary save creates a fresh artifact and that preserve
mode may raw-copy unchanged CFB streams when possible. Those constraints make
the question here narrower than “can a `Vec<u8>` be kept?”: a retained result
must remain the exact validated result for the current candidate and must not
weaken a later freshness, preservation, validation, or budget boundary.

## Ownership and current callsites

| Route | Owning editor and mutation | Render/recapture owner | Current publication shape |
| --- | --- | --- | --- |
| Common OLE2 | `litchi_ole_common::object::Editor`; `put_stream`, `put_stream_shared`, `add_stream`, `remove_*`, and `put_streams_shared` mutate an isolated clone | `commit_candidate_with_rendered` renders, reopens CFB, captures streams, reuses equal stream allocations, and rediscovers targets | `put_streams_shared` publishes the checked candidate but discards its rendered `Vec`; `finish` renders the current package again |
| DOC body text | `body_text::Edit::replace_paragraph` → `RevisionEditor::replace_plain_text` → `RevisionEditor::commit` → common `put_streams_shared` | common editor for the internal candidate; then `body_text::Edit::commit` delegates to `RevisionEditor::finish` | one internal staged render is discarded; outer finish renders again; changed bytes are reopened by `Snapshot::open_bounded` |
| PPT slide order | `slide_order::Transaction::remove_slide` mutates the semantic document and stores a persisted-record payload | PPT's private `embedded::object::Editor`; `Transaction::commit` opens it once, updates the live Document/Pictures streams, and calls its own `finish` | one final PPT embedded-editor CFB write followed by unrelated-stream checks, full snapshot reopen, slide-order/readback checks |

There is one existing public common-editor handoff,
`put_stream_shared_with_rendered` at
[`editor.rs:328`](../../../../crates/litchi-ole-common/src/object/editor.rs#L328).
Its documentation explicitly says that it does not cache a rendering on the
editor. It returns the already checked bytes to a caller that can consume them
immediately. The repository has no production callsite for this method, and
there is no batched `put_streams_shared_with_rendered` callsite. The method
therefore proves a deliberately narrow handoff API exists; it is not evidence
that a format owner currently consumes a rendered candidate.

The PPT slide-order path should not be described as a common-editor stage
render. Its `embedded::object::Editor` is a separate PPT-specific container
editor. The common `litchi-ole-common::object::Editor` is used by DOC and by
other OLE object owners, but it is not nested in the slide-order commit traced
here.

## Common editor lifecycle

`Editor` retains `targets`, bounded `limits`, the original source bytes, a
`base_package`, the current captured `package`, discovered objects, a changed
flag, and the sector-layout policy ([`editor.rs:60`](../../../../crates/litchi-ole-common/src/object/editor.rs#L60)).
Open parses the source and captures the package streams. Cloning the editor is
therefore a cheap isolation boundary for the `Arc`-owned stream data, while
the original source and the captured baseline remain available for the 0663
source-layout decisions.

For a non-identical stream replacement, `put_stream_shared` clones the editor,
updates the candidate package, and assigns the candidate only after
`commit_candidate` succeeds ([`editor.rs:309`](../../../../crates/litchi-ole-common/src/object/editor.rs#L309)).
`put_streams_shared` applies all non-no-op replacements to one candidate and
publishes once after the batch ([`editor.rs:360`](../../../../crates/litchi-ole-common/src/object/editor.rs#L360)).
Any missing stream, size-limit failure, package check failure, render failure,
reopen failure, recapture failure, or target-discovery failure occurs before
the candidate is assigned, so the current batch is failure-atomic.

The private `commit_candidate_with_rendered` sequence is the important source
fact ([`editor.rs:657`](../../../../crates/litchi-ole-common/src/object/editor.rs#L657)):

1. Check the candidate package against its limits.
2. Under `Reuse`, attempt `Package::render_copy_through`; if it declines,
   call `render_with_layout` with the original source. Under `Rewrite`, call
   the from-scratch layout policy directly.
3. Reopen the rendered bytes as an `OleFile`, run common CFB codec opening,
   capture every stream, and reuse equal stream allocations.
4. Rediscover the selected objects, set `changed`, and return the candidate
   plus the validated rendered `Vec<u8>`.

`commit_candidate` immediately maps that pair to the candidate and drops the
`Vec` ([`editor.rs:652`](../../../../crates/litchi-ole-common/src/object/editor.rs#L652)).
Thus the current `put_streams_shared` stage pays for a complete candidate
render and recapture, but retains only the reparsed package state. The bytes
are not available to the later `finish` call.

`finish` has an independent render decision
([`editor.rs:609`](../../../../crates/litchi-ole-common/src/object/editor.rs#L609)).
For a changed candidate it tries the same-length source-backed overlay again;
if that is unavailable it invokes `render_with_layout` again. A length-changing
edit cannot satisfy `same_length_overlays`, so a length-changing stage and its
later `finish` necessarily perform two render phases. This is the route the
0728 root probe should expose with separate stage and finish timers. A
same-length edit can also pay two phases: both stages may take the validated
overlay route, with each stage materializing a `Vec` after its own checks.

The overlay itself is guarded by the source model and directory shape. It
declines length changes, topology or metadata changes, and noncanonical cases;
when admitted it reopens the composed view and reads back every stream before
allocating the output ([`codec.rs:503`](../../../../crates/litchi-ole-common/src/object/codec.rs#L503)).
The source-layout fallback retains the 0663 `Reuse` semantics for growth and
metadata edits. These checks are part of the evidence a future handoff must
preserve; returning a prior `Vec` merely because a stream edit was once
rendered would not be sufficient.

## DOC body-text paragraph route

The source-backed DOC path opens the strict common editor during
`Snapshot::open_bounded`, then opens another `RevisionEditor` for each
`Edit` ([`body_text.rs:570`](../../../../crates/litchi-doc/src/body_text.rs#L570),
[`body_text.rs:1179`](../../../../crates/litchi-doc/src/body_text.rs#L1179)).
For a main-story paragraph replacement, the current sequence is:

1. `Edit::replace_paragraph` resolves and validates the target, then calls
   `RevisionEditor::replace_plain_text`
   ([`body_text.rs:1248`](../../../../crates/litchi-doc/src/body_text.rs#L1248)).
2. `replace_plain_text` clones the revision candidate, updates WordDocument
   and table-related state, and calls its internal `commit`
   ([`package.rs:1254`](../../../../crates/litchi-doc/src/tracked_revision/package.rs#L1254)).
3. That commit builds replacement stream `Arc`s and calls common
   `put_streams_shared` ([`package.rs:2648`](../../../../crates/litchi-doc/src/tracked_revision/package.rs#L2648)).
   Common `commit_candidate_with_rendered` renders, reopens, captures, and
   rediscovers the candidate; `commit_candidate` drops the rendered bytes.
4. The outer `Edit::commit` calls `self.editor.finish()`
   ([`body_text.rs:1651`](../../../../crates/litchi-doc/src/body_text.rs#L1651)).
   For the requested length-changing paragraph case, common `finish` falls
   through to the current 0663 source-layout writer and renders the package
   again.
5. Since the output differs from the source, `Edit::commit` calls
   `Snapshot::open_bounded` on the final bytes. That performs the strict-owner
   open and the public-reader validation before the patch is constructed.

This is a duplicate render/recapture shape, with a third validation/open at
the public DOC publication boundary. The stage and finish outputs are not
currently compared byte-for-byte because the first output is discarded. The
current checks prove that each stage and the final publication are valid; they
do not prove that retaining the stage output is safe across later edits or
that the stage output can replace the final public reopen.

The distinction matters when an `Edit` performs multiple operations. Every
semantic mutator that commits a revision can replace the common package state.
An artifact rendered for operation N is stale after operation N+1, even if the
same stream path is involved, because package bytes, stream sizes, allocation
choices, directory metadata, or policy fallback can differ. A future cache
must therefore be consumed only for the exact current candidate, or invalidated
on every subsequent package/topology/metadata/policy mutation.

## PPT slide-order remove route

`Transaction::remove_slide` changes the semantic slide list and captures the
removed persisted record for the reversible patch; it does not call common
`put_streams_shared` ([`slide_order.rs:1811`](../../../../crates/litchi-ppt/src/slide_order.rs#L1811)).
At commit, the transaction first commits the semantic document, opens the
PPT-specific embedded object editor over the `working` bytes, checks that the
live persisted Document record still matches the working document, applies
the staged records and stream changes, and calls that editor's `finish`
([`slide_order.rs:1903`](../../../../crates/litchi-ppt/src/slide_order.rs#L1903)).

The PPT editor opens and retains all streams in its own state
([`open.rs:59`](../../../../crates/litchi-ppt/src/embedded/object/editor/lifecycle/open.rs#L59)).
Its changed `finish` appends the incremental records, writes an OLE package
once with `adopt_source_layout`, and reopens the Document/Current User streams
for mapping validation ([`finish.rs:10`](../../../../crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs#L10)).
The outer slide-order commit then validates unrelated streams or pictures,
reopens the full snapshot, checks semantic slide order, and compares the
persisted slide payloads. This is one final CFB render in this route, with
additional semantic/readback validation around it; it is not evidence of a
common stage-render artifact that can be handed to `finish`.

## What current proofs establish, and what they do not

The existing common tests establish the relevant safety floor:

* An identical batched replacement leaves `changed` false, retains the same
  stream `Arc`, and finishes to the exact original bytes. A failed batch leaves
  the editor unchanged and exact ([`tests/object.rs:466`](../../../../crates/litchi-ole-common/tests/object.rs#L466)).
* Same-length copy-through reopens the output and verifies changed and
  untouched streams; a noncanonical v3 source declines the overlay so the
  writer can normalize it ([`tests/object.rs:377`](../../../../crates/litchi-ole-common/tests/object.rs#L377)).
* The 0663 corpus checks stream bytes, directory metadata and CLSIDs, CFB
  validation, mutation fallback, and deterministic output. Its documented
  boundary still says that length-changing and metadata edits use the
  source-layout writer, and that final output is materialized after validation.
* The DOC publication path reopens changed output under the source limits and
  the DOC tests exercise paragraph replacement, semantic readback, and output
  reopening. The PPT slide-order path performs a source-freshness check and
  full semantic/persisted-record readback before returning a commit.

These proofs do not establish a retained rendered-artifact budget, cache
invalidation, a batched rendered handoff, or equivalence of an old stage
artifact to a later `finish` after more edits. They also do not authorize
skipping the outer DOC public validation. The source review therefore
qualifies the duplicate work for measurement but does not authorize a
production retention change.

## Conditions for possible future reuse

A future optimization should prefer a bounded, one-shot handoff or a
batched variant of `put_stream_shared_with_rendered` over an unbounded cache
on `Editor`. If a retained candidate is still chosen, these conditions are
required:

1. **Exact state identity.** Bind the rendered bytes to the exact original
   source identity/bytes, current package generation, baseline package,
   target catalog, sector-layout policy, and all relevant limits. A simple
   stream-path/value key is insufficient: directory shape, stream size class,
   CLSID/metadata, source geometry, and fallback decisions affect the output.
2. **Invalidation on every mutation.** Any stream replacement, add/remove,
   storage or metadata change, target/discovery change, or policy setter must
   advance the generation and drop a prior artifact. An editor clone must not
   accidentally carry an artifact that describes the pre-clone package. The
   artifact must be consumed at most once by `finish`, unless the same exact
   state is deliberately proven immutable.
3. **Atomic publication.** Rendered bytes, reparsed package state, discovery
   results, and the generation marker must become visible together only after
   all current `commit_candidate_with_rendered` checks succeed. Any failed
   render, reopen, capture, discovery, limit check, or later mutation must
   leave the prior editor and its prior valid cache unchanged or discard the
   cache conservatively.
4. **Preservation and freshness.** Reuse must retain the current 0663
   `Reuse`/fallback behavior, CFB and stream readback checks, directory
   metadata/CLSID rules, and deterministic output. It cannot bypass final
   source-freshness checks at a format-owner boundary. The current common
   editor owns immutable source bytes, but any future source-backed or
   external `ReadAt`/sink route must revalidate source identity/version before
   publication.
5. **Exact no-op behavior.** Equal replacements must remain `changed == false`,
   share original stream allocations where promised, and return the original
   bytes. Installing a rendered cache for an equal replacement must not turn a
   no-op into a changed save or consume a second full artifact.
6. **Explicit retained-output budget.** The current limits bound source/package
   streams and transient output checks; they do not visibly reserve a second
   persistent full rendered CFB while the editor remains live. Retaining a
   stage `Vec` would add roughly one output-sized allocation to the peak in
   addition to original bytes, captured package state, candidate temporary
   state, and the later owner validation. The design must charge that retained
   artifact against an explicit output/candidate budget and fall back to the
   current rebuild when it does not fit. A budget refusal must never be the
   only way to edit a package when the current bounded rebuild is available.
7. **Owner lifetime and consumption.** For DOC, a retained artifact is safe
   only if no later revision mutation can occur before outer `finish`, or if
   the generation invalidation is complete. A returned one-shot artifact can
   be consumed by an owner immediately; an `Editor` cache copied through
   several semantic layers is much harder to prove and risks stale output.
   For PPT slide order, the current custom editor has its own source and
   readback contract, so a common-editor cache would not remove the traced
   final writer.

The 0728 probe may therefore report the current DOC `stage` and `finish`
windows separately, alongside the public-format commit and PPT control
windows. It should retain the current output/reopen oracles and use the
current 0663 policy controls. Any future before/after claim must measure the
landed handoff or retention design with its budget and invalidation tests; it
must not subtract independent public-format and common-editor medians or
reuse the withheld 0617 estimate.

## Disposition

Source qualification passes for a baseline measurement of duplicate common
OLE2 work. The duplicate is established for the DOC paragraph route: internal
`put_streams_shared` renders and recaptures a candidate, discards the rendered
bytes, and outer `finish` renders the changed package again. For a
length-changing paragraph, the second render necessarily uses the current
0663 source-layout writer. The traced PPT slide-order remove route has no
common-editor stage render; it has one custom PPT embedded-editor finish plus
its existing reopen and semantic readbacks.

No production retention or reuse change is recommended from this review
alone. The next evidence is the root probe's current phase timing and public
format controls, followed by a separately qualified future design if the
measured duplicate justifies the added retention budget and invalidation
surface.
