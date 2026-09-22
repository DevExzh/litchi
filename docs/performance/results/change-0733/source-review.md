# 0733 PPT embedded-finish source review

This is a read-only source audit for the next PPT performance investigation. It
does not change production Rust, the probe, Cargo files, or profiler inputs, and
it makes no native latency claim. The current 0732 measurement puts the
`EmbeddedFinish` phase at 188.49–190.83 microseconds, or 16.73–16.87% of the
whole owner, with three same-round p50 observer-control flags. Those numbers
justify a bounded source investigation; they do not authorize removing a
validation boundary. The 0731 Callgrind profile used software SHA and is not
evidence for native SHA or any other native cost.

## The measured owner and the finish call count

The fixed PPT probe constructs its expected output by running one real public
format edit before the measured route
([`ppt_phase.rs:701-708`](../../../../docs/performance/results/change-0732/probe/src/ppt_phase.rs#L701)).
The ordinary split route then performs one `Snapshot::from_bytes`, one
`Snapshot::edit`, one `remove_slide`, one `Transaction::commit`, and one output
copy
([`ppt_phase.rs:177-202`](../../../../docs/performance/results/change-0732/probe/src/ppt_phase.rs#L177)).
The profiled routes use the same owner sequence and call
`commit_profiled` at the corresponding commit boundary
([`ppt_phase.rs:240-265`](../../../../docs/performance/results/change-0732/probe/src/ppt_phase.rs#L240)).

`Snapshot::edit` only runs editability checks and constructs the transaction
([`slide_order.rs:742-766`](../../../../crates/litchi-ppt/src/slide_order.rs#L742)).
`remove_slide` computes order digests, removes the projected slide, and reads
the selected persisted payload for the reversible structural change
([`slide_order.rs:1897-1921`](../../../../crates/litchi-ppt/src/slide_order.rs#L1897)).
Its helper `persisted_record` opens an editor and reads one record
([`slide_order.rs:3806-3813`](../../../../crates/litchi-ppt/src/slide_order.rs#L3806));
it does not finish or write an OLE package. Therefore neither edit
initialization nor the fixed remove operation accounts for extra embedded
finishes.

The ordinary commit has one `editor.finish()` call
([`slide_order.rs:1996-2066`](../../../../crates/litchi-ppt/src/slide_order.rs#L1996)),
and the profiled copy has one mutually exclusive call
([`slide_order.rs:2170-2225`](../../../../crates/litchi-ppt/src/slide_order.rs#L2170)).
Other formatting mutators have a different contract: setters such as slide
visibility and media changes publish their intermediate working document
through `publish_live_document`, which calls `editor.finish()`
([`slide_order.rs:3753-3775`](../../../../crates/litchi-ppt/src/slide_order.rs#L3753)).
Those additional finishes are real for workflows that compose those setters,
but they are not part of the fixed 0732 remove-slide owner and must not be
folded into its `EmbeddedFinish` fraction.

The raw 0731 Callgrind edge count of three for
`finish -> write_package -> adopt_source_layout` is an aggregate call counter,
not a per-owner count. The probe runs the expected public edit before timing
([`lib.rs:2997-3004`](../../../../docs/performance/results/change-0732/probe/src/lib.rs#L2997)),
executes a wrong-slide public edit as an oracle control
([`lib.rs:2939-2956`](../../../../docs/performance/results/change-0732/probe/src/lib.rs#L2939)),
and then runs the measured public wrapper
([`lib.rs:2421-2449`](../../../../docs/performance/results/change-0732/probe/src/lib.rs#L2421)).
Collection toggling and Callgrind's call accounting can retain calls from
outside the collected instruction interval. An independent safe-Rust call
witness in the 0733 work packet confirms that an edge can report three calls
while only one invocation is inside the collected owner. The source path above
supports one finish in the measured remove owner; no projected snapshot in
that path creates another one.

## Public commit path around the finish

For a real structural change, `Transaction::commit` first commits the document
transaction. It then captures all surviving slide payloads with
`persisted_slides`, opens a PPT embedded-record editor, checks that the live
Document record still equals the transaction's working document, stages the new
Document record, and finishes the editor
([`slide_order.rs:1997-2060`](../../../../crates/litchi-ppt/src/slide_order.rs#L1997)).
The staging call itself clones the complete editor to retain failure-atomic
mutation semantics: `Editor` derives `Clone` over its stream payloads and
`replace_persisted_record` clones before assigning the candidate
([`editor/mod.rs:27-45`](../../../../crates/litchi-ppt/src/embedded/object/editor/mod.rs#L27),
[`records.rs:6-29`](../../../../crates/litchi-ppt/src/embedded/object/editor/mutation/records.rs#L6)).
That is neighboring commit work, outside the `EmbeddedFinish` expression, and
is not safe to remove as part of a writer-only pilot.

After the embedded finish returns, the public owner independently checks all
unrelated streams, reopens the candidate as a public `Snapshot`, compares live
slide order, captures the surviving payloads again, checks the expected
payload map, validates transferred pictures, and constructs the reversible
patch
([`slide_order.rs:2061-2103`](../../../../crates/litchi-ppt/src/slide_order.rs#L2061)).
These checks overlap in bytes read, but their invariants differ: package stream
preservation, public PPT semantic reopening, and source-checked reversible
payload identity. This review does not recommend combining or deleting them.

## Embedded editor finish and writer path

The changed editor finish has the following ownership and validation sequence:

1. `finish` removes deleted mappings, projects all staged record lengths under
   the output limit, builds the appended `PowerPoint Document` stream, emits a
   PersistDirectory and UserEdit record, and updates Current User
   ([`finish.rs:10-105`](../../../../crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs#L10)).
   The `appended` vector is an intentional new incremental document stream; it
   cannot be replaced by the old document bytes.
2. `write_package` creates an `OleWriter`, selects the retained sector policy,
   adopts the original CFB layout, registers every source stream, and writes
   into a bounded output cursor
   ([`finish.rs:122-143`](../../../../crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs#L122)).
3. `OleWriter::adopt_source_layout` reparses the original CFB into bounded
   source-layout metadata. It retains the header, FAT/MiniFAT, directory and
   mini-stream images, path/entry metadata, and class identifiers so Reuse can
   preserve sector placement and metadata
   ([`core.rs:424-457`](../../../../crates/litchi-cfb/src/writer/core.rs#L424),
   [`layout.rs:390-542`](../../../../crates/litchi-cfb/src/writer/layout.rs#L390)).
4. Under the default Reuse policy, `write_to` plans the source-anchored layout,
   composes a positional read-only output view, and opens every planned stream
   through the ordinary CFB reader before the sink sees bytes
   ([`core.rs:1122-1150`](../../../../crates/litchi-cfb/src/writer/core.rs#L1122),
   [`layout.rs:814-841`](../../../../crates/litchi-cfb/src/writer/layout.rs#L814)).
   This is deliberately a pre-emission safety fence: it checks the complete
   FAT/MiniFAT/directory partition and stream readback without allocating a
   second full output artifact.
5. The validated plan emits the header, metadata images, free-sector padding,
   and stream payload sectors to the bounded sink
   ([`layout.rs:1882-1979`](../../../../crates/litchi-cfb/src/writer/layout.rs#L1882)).
   This copy is the actual final package materialization and is required.
6. `validate_rewrite` then opens the emitted bytes and rereads PowerPoint
   Document and Current User to parse the latest persist mapping
   ([`finish.rs:146-158`](../../../../crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs#L146)).
   It returns the existing output vector; it does not render a second package.
   This post-sink check remains relevant even when ReusePlan validation passed:
   the latter validates a composed pre-sink view, while this check validates the
   bytes actually emitted by the writer. The common editor also uses this path
   for object collections, so an open-record route cannot remove it globally.

The most concrete avoidable copy is at
[`finish.rs:122-135`](../../../../crates/litchi-ppt/src/embedded/object/editor/transaction/finish.rs#L122).
`Editor::streams` already owns every stream as a `Vec<u8>`, but
`write_package` passes borrowed slices to `OleWriter::create_stream`. That API
allocates a second payload vector and copies the complete stream
([`core.rs:589-599`](../../../../crates/litchi-cfb/src/writer/core.rs#L589)).
The writer already exposes `create_stream_owned`, which takes that allocation
without cloning it
([`core.rs:601-618`](../../../../crates/litchi-cfb/src/writer/core.rs#L601)).
The copy is visible in the retained 0731 instruction profile as direct
`create_stream`/memcpy work. Those are collected instruction counts under
Valgrind, not a native cost bound or speedup estimate.

There are two other real duplications that explain why the finish path should
be kept separate from its surrounding phases:

- `open_with_limit` reads PowerPoint Document and Current User once into their
  dedicated editor fields, then reads every stream again into `Editor::streams`,
  including those two streams
  ([`open.rs:70-110`](../../../../crates/litchi-ppt/src/embedded/object/editor/lifecycle/open.rs#L70)).
  This is an editor-state ownership choice and belongs to EmbeddedOpen, not
  EmbeddedFinish.
- `SourceLayout::parse` reparses the source after `Editor::open_records` has
  already parsed and materialized it. The writer needs its own bounded layout
  representation, so removing this requires a cross-crate layout handoff and
  retention decision. It is a broader follow-up than the stream-payload move.

`ReusePlan::validate`, `ReusePlan::emit`, `validate_rewrite`, the outer
unrelated-stream comparison, and the public snapshot reopen all read related
bytes, but each protects a different boundary. They are not interchangeable
copies to remove based on the 0732 fraction.

## Smallest plausible next pilot

The smallest candidate is to preserve the current writer and validation design
while transferring already-owned payload vectors into it. After source-layout
adoption succeeds, `write_package` could drain or otherwise move the
`Editor::streams` payloads, move the newly built `appended` vector for the
PowerPoint Document stream, move the updated Current User vector, and call
`create_stream_owned` for each path. The editor must retain only the fields
needed by `validate_rewrite` (document/current-user paths, collection, limits,
and the source/layout metadata needed before the move). The candidate must not
move or drop those fields earlier than the ordinary path, and it must keep
source-layout adoption before stream registration so malformed-source and
allocation error ordering remains reviewable.

The pilot must demonstrate stream-path uniqueness from the parsed topology
and all editor mutation paths before relying on a one-shot move. Existing
invariants or guards can supply that proof; a new duplicate-path check is not
required if duplicates are structurally impossible. If duplicate paths remain
reachable, their existing behavior must be preserved: the borrowed path can
copy a selected payload more than once, whereas a one-shot move cannot. The
pilot should also document the changed allocation failure surface: removing a
payload clone can remove an artificial `stream payload` allocation failure,
while all format, output-limit, writer-I/O, and mapping errors must remain
typed and ordered as before.

The retained `create_stream` inclusive work is 383,349 collected instructions
in the first 0731 profile; its direct memcpy edge accounts for 377,604.
Embedded finish has 3,087,582 inclusive instructions in that profile. Stream
ingress is therefore about 12.416% of finish Ir, while the memcpy edge alone
is about 12.230%. Neither fraction predicts native savings or proves that all
of that work can be removed. The sealed output inventory has five streams
totaling 384,906 payload bytes, which the current copying ingress materializes
again. That is a source-derived copy volume, not a measured allocation saving.
A fresh ordinary-path measurement is required before retaining the pilot.

## Required tests before any pilot is retained

The candidate should first prove exact behavioral equivalence on the existing
PPT editor and public owner tests:

- finish the 45543 document-record replacement and compare output bytes,
  latest UserEdit/PersistDirectory linkage, live persist mapping, and reopened
  record payloads with the current implementation
  ([`mapping.rs:32-88`](../../../../crates/litchi-ppt/src/embedded/object/editor/tests/mapping.rs#L32));
- run the public remove-slide success, unrelated-stream preservation,
  semantic reopen, reversible patch, inverse, exact no-op, and formatting-only
  paths; compare ordinary output/patch values byte-for-byte where the policy is
  the same;
- exercise both `SectorLayoutPolicy::Reuse` and `Rewrite`, including a Reuse
  fallback, so moving payload ownership cannot change stream topology, class
  identifiers, directory metadata, or policy selection;
- retain the output-limit and malformed-source cases
  ([`limits.rs:1-51`](../../../../crates/litchi-ppt/src/embedded/object/editor/tests/limits.rs#L1))
  and add a writer-I/O failure case if the test seam permits it, checking error
  class/message and failure atomicity;
- cover a source containing the large Document stream, the small Current User
  stream, and unchanged auxiliary streams. Verify that all stream bytes and
  paths remain exact and that no moved vector is used after writer failure;
- compare allocation ownership and peak live bytes separately from semantics.
  The expected result is fewer payload-copy allocations; no RSS or native
  speedup claim follows from that observation.

The pilot must leave `ReusePlan::validate`, final emission,
`validate_rewrite`, unrelated-stream validation, public reopen, payload
readback, and the source-checked patch intact. A cached source layout or a
skipped reopen would require a separate design and broader invariant review.

Root finalized the numeric attribution and invariant-proof wording after the
bounded source audit; the independent packet review checks the partition
interpretation. No production changes or native measurements were made here.
