# 0729 source review: public DOC phase attribution and observer equivalence

Status: read-only source qualification, 2026-09-22. The reviewed source is
the current repository at `949a4a3037` (`perf(ole): refresh current DOC and
PPT save baseline`). This review changes no production source, Cargo file,
native code, or profiler. It makes no latency, allocation, or memory claim.
The 0729 probe may use these findings to label its measurements, but the
source review does not replace the probe's semantic and artifact gates.

## Scope and authority

The question is whether the public Word 97+ DOC route can use the existing
`performance-diagnostics` APIs to attribute a complete open/edit/replace/
commit lifecycle while retaining the ordinary route's semantics. The relevant
implementation is in [`litchi-doc`'s body-text transaction](../../../../crates/litchi-doc/src/body_text.rs),
the tracked revision owner
([`package.rs`](../../../../crates/litchi-doc/src/tracked_revision/package.rs)),
and the common OLE object editor
([`editor.rs`](../../../../crates/litchi-ole-common/src/object/editor.rs)).
The 0728 review remains the source for the common render/recapture lifecycle
and its general rendered-output handoff constraints
([`change-0728/source-review.md`](../change-0728/source-review.md)).

The public route must continue to satisfy the repository's preservation and
transaction rules: strict owner validation establishes a safe editing basis;
the independent public reader validates the candidate; output publication is
failure-atomic; exact no-ops retain source identity; reversible patches are
source-checked; and package, stream, and transaction limits remain enforced.
Diagnostics may observe these boundaries, but they cannot replace any of
them.

## Actual public DOC call path

For the 0729 case, `Snapshot::open` followed by `edit`, a length-changing
`replace_paragraph`, and `Edit::commit` has this shape:

1. `Snapshot::open_bounded` first performs source-length preflight, then
   calls `RevisionEditor::open_with_ole_file`. The strict owner opens the
   CFB, captures the selected package streams, checks the DOC FIB/piece-table
   basis, and discovers its supported objects. The owner remains alive while
   `Package::validate_ole_file` independently validates the same parsed OLE
   view. Finally, the snapshot retains the exact input bytes in an `Arc`.
2. `Snapshot::edit` constructs another `RevisionEditor` from a copy of the
   exact source bytes. This is a second strict common-editor open. There is no
   profiled `Edit::new` or profiled `RevisionEditor::open` boundary in the
   current public API.
3. `Edit::replace_paragraph` resolves the public target and checks structural
   content, drawing dependencies, story length, compressed pieces, tracked
   text, replacement limits, operation limits, and character-format closure.
   A main-story length-changing replacement then enters
   `RevisionEditor::replace_plain_text`.
4. `replace_plain_text` clones the revision candidate, changes WordDocument
   and table state, rebuilds the affected CHPX/CLX/FIB data, and calls the
   candidate's internal `RevisionEditor::commit`. That method batches the
   WordDocument, selected Table, and possibly Data replacements through
   `ObjectEditor::put_streams_shared`. The common editor checks, renders,
   reopens, recaptures, reuses equal stream allocations, and rediscovers
   targets before installing the candidate package. Its validated rendered
   `Vec<u8>` is dropped by `commit_candidate`; it is not available to the
   later outer commit.
5. The public `Edit::commit` calls `RevisionEditor::finish`, which renders the
   current changed common package again under the current sector-layout policy
   (`Reuse` with its fallback, or `Rewrite`). For the 0729 length-changing
   route, the first candidate render and this final owner render are separate
   work. The outer timer therefore includes both.
6. Because the final bytes differ from the source, `Edit::commit` reopens them
   through `Snapshot::open_bounded`. This adds a strict-owner open, the
   independent public-reader validation, and exact-source retention before
   `Patch::new` constructs the source-checked reversible patch.

The route has more work than the public `Finish` phase name suggests. In
particular, replacement staging and `Edit::new` are outside the current
diagnostic phase stream, and the first staging render is not the final output
object. A whole-operation timer is therefore the only current complete
lifecycle measurement.

## What the existing observers cover

With `performance-diagnostics`, `Snapshot::open_bounded_profiled_with_cfb_observer`
keeps the high-level order of ordinary open:

```text
source-length preflight (silent)
  StrictOwnerValidation
    CfbParseEvent: Started / Finished around one top-level OleFile::open
  PublicReaderValidation
  SourceRetention
```

`DiagnosticEvent` is content-free and synchronous. It contains only a phase
and, at completion, `Success` or `Error`; it carries no bytes, offsets,
stream names, package identifiers, or timestamps. The CFB observer is a
separate content-free pair around the strict owner's top-level in-memory
`OleFile::open`. It does not cover public-reader checks, stream recapture,
target discovery, or every lower-level parse operation. Source-length
preflight returns before either observer is called.

`Edit::commit_profiled_with_cfb_observer` reports the following changed
candidate sequence:

```text
Finish
  StrictOwnerValidation
    CfbParseEvent: Started / Finished around the final candidate OleFile::open
  PublicReaderValidation
  SourceRetention
Patch
```

The CFB pair is nested inside the high-level strict-owner interval. It must
not be added as another non-overlapping phase. `Finish` is the outer owner's
render and does not include the hidden render/reopen performed earlier by
`replace_plain_text`'s common `put_streams_shared` call. The high-level
events around the final `Snapshot::open_bounded` are sequential; they are not
an additional second parse of the CFB after the strict owner has returned it.

The exact no-op commit has a different, intentional stream:

```text
Finish
ExactNoOp
Patch
```

It shares the source allocation and does not reopen a candidate, so it emits
no new CFB pair during that commit call. A same-text replacement can return
before staging; the later commit still reports the no-op decision. Ordinary
open and commit emit no events at all. This difference is part of the public
diagnostic contract, not an error in the observer.

On errors, `observe_phase` sends `Started`, runs the operation, then sends one
`Finished` event with `Error` before returning the typed error. A strict-owner
failure therefore closes `StrictOwnerValidation` and normally closes its CFB
parse pair with `Error`; a public-reader failure closes strict owner with
`Success`, then closes public-reader validation with `Error`; no source-
retention or patch event follows. A preflight failure is silent because no
phase has started. The balancing guarantee assumes callbacks return normally;
an observer panic propagates and can interrupt the event stream.

The existing format tests prove this contract for successful open, strict
failure, public-reader failure, native DOP refusal, changed commit, exact
no-op commit, and the typed finish-error helper. They also compare ordinary
and profiled semantic results, output bytes, patch state, and the expected
CFB event cardinality. The test observers use `Vec::push`, which is suitable
for event correctness but unsuitable as a low-overhead allocation/timing
recorder.

## The ordinary/profiled lifetime difference

The profiled open is semantically intended to be equivalent to ordinary open,
and the tests establish equal results and errors. Its implementation is not a
pure callback wrapper around the ordinary function, however:

* Ordinary `open_bounded` binds
  `let (_strict_editor, mut ole) = RevisionEditor::open_with_ole_file(...)`
  in the outer function. The strict editor remains live through
  `Package::validate_ole_file(&mut ole, ...)` and through construction of the
  source-retaining `Arc` before the function returns.
* Profiled open creates the same `_strict_editor` inside the
  `StrictOwnerValidation` closure and returns only `ole` from that closure.
  The strict editor is dropped when the closure returns, before
  `PublicReaderValidation` and before `SourceRetention`.

This changes ownership lifetime and destructor timing. It can change which
allocations remain live during public-reader validation and source retention,
and can change peak live memory. It does not by itself show that document
semantics differ: the current tests show equal successful snapshots and equal
error classes/messages for the exercised cases. It does show that a profiled
phase time or allocation count cannot be presented as ordinary-phase time
plus callback overhead. Any difference is a mixture of observer behavior and
the profiled implementation's lifetime boundary.

The same qualification applies to a complete changed commit: the final
profiled candidate reopen uses the profiled open implementation, while the
ordinary final reopen uses ordinary `Snapshot::open_bounded`. The profiled
route is therefore a valid public semantic route for event attribution, but
it is not a byte-for-byte control implementation for phase or allocation
comparison. This review does not recommend changing production lifetime
scopes merely to equalize the routes.

The CFB observer itself is synchronous and runs inside the phase it observes.
Its callback cost is included in that phase interval. A callback that writes a
`Vec`, formats text, reads a clock, or performs I/O would measure the
recorder as part of the phase. Callback panics also change the normal
balancing behavior. The observer contract deliberately leaves timing policy
to the harness.

## Bounded observer-overhead control

The 0729 harness should use fresh processes and the exact same source bytes,
replacement text, limits, and operation order for paired lanes:

1. The ordinary lane runs `Snapshot::open` → `edit` → replacement → ordinary
   `commit` and has no observer.
2. A profiled no-op lane runs the corresponding profiled open/commit APIs with
   callbacks that immediately return. This bounds the cost of entering and
   exiting the diagnostic wrappers while retaining the profiled lifetime
   difference described above.
3. A profiled fixed-recorder lane uses callbacks that update only preallocated
   fixed-size state: event counters, a small phase state machine, and a
   wrapping or checked `u64` event digest. It must not allocate, format, take
   timestamps, perform I/O, or retain event vectors. The semantic and CFB
   event counts/order/digest are checked after the timed operation.

The no-op-to-fixed delta is the bounded observer work. The ordinary-to-profiled
delta is the whole route difference and must be labelled as such; it cannot be
subtracted into a claimed format-phase speedup because it also contains the
strict-editor lifetime difference. Allocation measurements should use the
fixed recorder and separate fresh-process ordinary/profiled runs; a `Vec`-based
correctness observer must not be used in the allocation lane.

Every lane should include changed length, exact no-op, and at least one
reachable validation/refusal control where the route can be constructed. For
the changed case, check equal candidate bytes, semantic paragraph readback,
source/output hashes, untouched stream bytes, storage/directory metadata and
CLSID policy, and forward/inverse patch behavior. For no-op, check exact
source bytes, `changed == false`, source allocation sharing where promised,
and absence of a commit-window CFB pair. For failures, check the ordinary
error against the profiled error and verify that every started event closes
before return.

The phase timer must begin and end outside the fixed recorder's own state
updates if the harness wants a separate callback-cost bound, while the
library's phase interval still includes the synchronous callback by design.
The harness should retain both views: phase intervals as observed by the
public API, and the independent no-op/fixed observer delta. A single
`Vec::push` recorder cannot provide that bound.

## Can the complete lifecycle be timed now?

Yes, the complete public lifecycle can be timed as one outer interval on both
ordinary and profiled routes, with semantic comparison against the same source
and operation. The existing high-level events can then be used to label the
named portions that actually occur in the profiled route.

No, the current events cannot make every lifecycle phase additive or claim
complete internal attribution. The missing portions are at least:

* the second strict open performed by `Snapshot::edit`;
* public target resolution and safety checks before replacement staging;
* the clone, WordDocument/Table/FIB/PLC/CHPX/CLX work in
  `replace_plain_text`; and
* the common candidate render, CFB reopen, recapture, and discovery inside
  that staging commit.

The right current label is an explicit `edit_setup_unattributed` and
`replacement_stage_unattributed` interval inside the outer lifecycle, or one
checked `unattributed_remainder` after the named events. The outer lifecycle
must equal the sum of its measured named intervals and the checked remainder
within the timer's arithmetic policy. CFB intervals are nested evidence and
must not be counted twice. A phase sum that omits the staging render is not a
complete lifecycle attribution.

Adding events around `Edit::new` and the revision-owner staging commit could
make that attribution finer, but that would be a production instrumentation
change outside this read-only review. Even then, event timing would still
need ordinary/profiled semantic and allocation qualification because the
strict-editor lifetime difference would remain unless explicitly redesigned.

## Mandatory validation and allocation boundaries

The public candidate handoff must preserve all of the following:

* source-length preflight and strict DOC-owner validation before editing;
* independent public-reader validation of the final CFB/package/document;
* the common candidate package check, CFB reopen, recapture, allocation reuse,
  and target discovery performed by `put_streams_shared`;
* the public DOC checks for FIB, selected table, piece table, and supported
  dependency closure;
* changed-length semantic readback for the edited paragraph and preservation
  of all other stories, paragraphs, objects, streams, directory entries and
  CLSIDs under the current policy;
* exact no-op source identity and no-op patch behavior;
* forward and inverse source-checked patch validation; and
* limits and failure-atomic publication at each owner boundary.

The final public reopen is not redundant just because common candidate
validation already happened. The common editor validates an OLE candidate and
its selected targets; `Package::validate_ole_file` validates the public DOC
reader's independent invariants. A rendered handoff may avoid a duplicate
render, but it cannot turn either validation layer into an assumption.

The current route already has several full-size live allocations: the exact
source, strict/common editor state, cloned semantic candidate state, stream
replacement `Arc`s, the common candidate rendered bytes during validation,
and the final rendered bytes during owner finish/public reopen. Keeping the
first rendered `Vec` past its common validation would add an output-sized
retention interval. It may be useful for a handoff, but it must be charged to
an explicit bounded output/candidate budget. Allocation or RSS results must
be measured directly; they cannot be inferred from phase elapsed time.

## Plausible validated-render handoff

The narrowest existing seam is
`ObjectEditor::put_stream_shared_with_rendered`. It calls the same candidate
checks as an ordinary single-stream update, installs the validated candidate,
and returns the exact rendered bytes without caching them. DOC replacement is
not a single-stream operation, however: `RevisionEditor::commit` can update
WordDocument, the selected Table stream, and Data. A future DOC handoff would
therefore need a batched counterpart of this seam, conceptually at
`RevisionEditor::commit` around `self.package.put_streams_shared(...)`.

The candidate should be rendered, reopened, recaptured, allocation-reconciled,
and rediscovered once; only after all those checks succeed should the candidate
package state and a one-shot returned rendered artifact become visible to the
DOC owner. The outer `Edit::commit` could consume that exact artifact for its
next boundary, but it must still perform the final `Snapshot::open_bounded`
(or an explicitly equivalent independent strict-owner/public-reader
validation) and then build `Patch::new`. The existing single-stream handoff is
an immediate ownership transfer, not a cache and not permission to skip
public validation.

The handoff requires all of these guards:

1. **Exact identity and freshness.** Bind the artifact to the exact source
   bytes/identity, current revision and common-package generation, replacement
   paths and data, target catalog, sector-layout policy, and every relevant
   package/transaction limit. A stream path and value alone do not identify a
   rendered OLE package. The artifact must be rejected after any later edit,
   stream/topology/metadata mutation, policy change, limit change, or source
   freshness mismatch.
2. **Atomic publication.** Keep the prior owner state unchanged until render,
   reopen, codec/public common checks, recapture, allocation reconciliation,
   discovery, and all replacement validation succeed. Publish the candidate
   state and its handoff token together. A failure must not leave a token that
   describes a state the editor did not publish.
3. **One-shot consumption.** Consume the returned bytes immediately at the
   outer owner boundary or mark them consumed exactly once. If another public
   edit occurs first, discard them and use the ordinary rebuild. Do not carry
   an arbitrary rendered cache through semantic transactions.
4. **Preservation and no-op behavior.** Retain the current `Reuse`/fallback
   policy, directory and CLSID rules, deterministic rendering, untouched
   stream identity checks, and exact no-op behavior. An equal replacement must
   remain an unchanged source result and must not install a changed rendered
   artifact.
5. **Budget and fallback.** Reserve the extra full rendered output and any
   retained candidate state against a documented output/candidate limit. If
   the handoff would exceed that budget, fall back to the current validated
   render/finish route. A budget refusal cannot make an otherwise supported
   bounded edit fail when the current route would succeed.
6. **Independent final checks.** The final DOC owner must still run source
   freshness, strict-owner, public-reader, semantic readback, and patch
   checks. The handoff is a reuse of validated bytes, not a validation bypass.

This seam is specific to the common DOC path traced here. It does not imply a
common-editor integration for PPT slide-order transactions, whose embedded
editor and persisted-record checks have a separate lifecycle.

## Disposition for 0729

The current APIs support a defensible descriptive measurement: whole public
DOC lifecycle, named profiled intervals, nested CFB evidence, checked
unattributed staging, and direct semantic/error/no-op comparison. They do not
support a claim that ordinary and profiled phase timings differ only by
observer overhead, nor a claim that the named phases cover the complete
lifecycle. The ordinary/profiled strict-editor lifetime difference must be
reported with every phase or allocation result. No production optimization or
validated-render handoff should be adopted from this source review alone.
