# 0917 — current DOC many-edit retention design audit

Status: read-only design audit, 2026-10-05, against `bb19ca74bd52`. This file
owns the current-state recommendation for the non-iWork DOC follow-up. It does
not change production code, build a binary, or make a performance claim.

## Decision

The bounded validated-render handoff requested by 0729 is already implemented
and retained by 0730 (`4599c1225d`, `perf(doc): retain bounded validated render
through body commit`). The recommendation in 0906 to make that handoff the next
pilot is therefore superseded. Repeating its implementation or another
single-edit A/B would be redundant.

The smallest safe follow-up is evidence for the existing public DOC lifecycle
at 1, 2, 8, and 32 successful staged edits, using zero, default 8 MiB, and
explicit-release retention routes. The frozen protocol in
[`native-protocol.json`](native-protocol.json) deliberately repeats
`Position(0)` in one `Edit` with two different ordinary strings. That is a
useful retention-invalidation workload: every stage must change the current
paragraph, while the final text is deterministic by edit-count parity. It does
not authorize a distinct-paragraph batch or a production cache.

## Current path and the exact repeated work

The public source path is:

```text
Snapshot::open_bounded_with_policy
  -> strict RevisionEditor open
  -> independent Package::validate_ole_file
  -> exact source Arc<[u8]>
Snapshot::edit / Edit::new
  -> strict editor reopen with the transaction retention ceiling
each successful Edit mutation
  -> public target, dependency, protection, and limit checks
  -> clone-first RevisionEditor mutation
  -> ObjectEditor::put_streams_shared_with_rendered
  -> package check, Reuse/copy-through or layout render, CFB reopen,
     package recapture, equal-stream allocation reuse, and target discovery
  -> current validated Vec is optionally retained by the DOC owner
Edit::commit
  -> take the one-shot retained Vec, or render through finish when absent
  -> authorization and signed-source check
  -> strict owner + independent public-reader validation
  -> exact source-checked Patch construction
```

The relevant code is `Snapshot::open_bounded_with_policy` and
`Snapshot::editor` in
[`body_text.rs`](../../../../crates/litchi-doc/src/body_text.rs:671),
`Edit::reopen_editor`, `Edit::take_candidate`, and `Edit::commit` in the same
file (around lines 1396, 1463, and 2086),
[`RetainedRender`](../../../../crates/litchi-doc/src/tracked_revision/package.rs:903),
and `ObjectEditor::put_streams_shared_with_rendered` plus
`commit_candidate_with_rendered` in
[`object/editor.rs`](../../../../crates/litchi-ole-common/src/object/editor.rs:530).

For a successful mutation, the common candidate render and its reopen are
needed to validate that intermediate staged state before another mutation can
observe it. A clone drops the prior `RetainedRender`; assignment installs only
the newly validated state. Thus N edits do not retain N serialized documents.
There can be transient old-token/new-candidate overlap during a clone-first
mutation, which is why peak live bytes must be measured separately from the
retention ceiling.

The one repeated render that 0730 removes is the final common render when the
last validated handoff is unavailable: a zero ceiling, explicit release, or an
output whose `Vec::capacity()` exceeds the ceiling leaves no token, so commit
recomputes through `finish_unchecked`. That fallback is intentional and must
remain result-equivalent. The final strict-owner/public-reader reopen is
mandatory, and the per-stage common validation is not redundant work that this
follow-up may skip. The missing-`Data` add-stream branch in
`RevisionEditor::commit` remains a separate picture-topology path that can do
an intermediate common render; plain paragraph replacement does not exercise
it and this audit does not propose changing it.

The common layer returns a validated `Vec` to its caller and does not retain a
cache. The DOC owner keeps at most that already-made allocation in private
editor state. No global cache, source-byte key, or `Reuse` to `Rewrite` change
is justified. Later body changes for saved-selection (`Selsf`) CP updates,
`RgDofr` auxiliary ownership, protection and signed-source gates, and embedded
object routes make a byte-only cache key especially insufficient; the frozen
plain-text sequence must stay on the existing source-bound owner path.

## Fixture identity and safe edit sequence

The retained real-DOC fixtures remain the 0730 inputs:

| Fixture | Bytes | SHA-256 | `Snapshot::paragraphs(Projection::All)` source count | safe ordinary nonempty body positions observed |
| --- | ---: | --- | ---: | --- |
| `NoHeadFoot.doc` | 26,112 | `45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48` | 10 | `0, 2, 4, 7, 9` |
| `FloatingPictures.doc` | 335,360 | `41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d` | 177 | derive from the current public paragraph list; controls, table paragraphs, and picture anchors are not safe text targets |

The count 177 is the current body-text source count. The historical facade
count of 210 in 0638 is a different API and must not be substituted here.
`source_paragraphs` in `body_text.rs` (around line 4276) emits only ordinary
main-story records and skips table paragraphs; `resolve_target` indexes that
filtered list.

The frozen workload uses these plain ASCII values, which are 45 and 46 UTF-16
units respectively:

```text
FIRST  = "litchi copy-through baseline replacement text"
SECOND = "litchi retained render second replacement text"
```

For stage `j` (one-based), replace `Position(0)` with FIRST for odd `j` and
SECOND for even `j`. `NoHeadFoot` and `FloatingPictures` both have an ordinary
plain paragraph at position zero, and these values contain no structural
controls. The functional preflight must still verify that the source text is
different from FIRST, FIRST differs from SECOND, and every stage differs from
the immediately preceding staged text. This prevents a silent no-op from being
counted as a successful edit if a fixture changes.

The expected final paragraph-zero text is FIRST for odd N and SECOND for even
N; the other 9 or 176 source paragraphs remain textually unchanged and the
source/output paragraph counts remain 10 or 177. A future distinct-position
workload must derive its positions from the current filtered paragraph list and
reject empty, structural, tracked, drawing-dependent, non-uniform, table, or
field targets during an untimed preflight. `NoHeadFoot` alone cannot supply 8
or 32 distinct safe nonempty positions, so repeating position zero is the
appropriate frozen many-edit workload.

`documentProperties.doc` has a small ordinary paragraph but is protected in
the current policy. It remains a producer/protection control, not a primary
many-edit route; adding an audited authorization would change the question
being measured.

## Retention observations that must hold

`TransactionLimits::default()` permits 256 operations, 1 MiB per replacement,
4 MiB aggregate replacement units, and 8 MiB retained allocation capacity.
The sequence above is well below the text limits. A custom policy used for a
limit/refusal control must have `max_operations >= N` and enough aggregate
replacement units.

For every successful stage, record `retained_render_bytes()` as `Vec::capacity()`
and not serialized length. The route expectations are:

| Route | Before each stage | After a successful stage | After explicit release | Before commit |
| --- | --- | --- | --- | --- |
| zero | `None` | `None` | not applicable | `None` |
| default 8 MiB | previous stage's capacity, or `None` at stage 1 | `Some(capacity)` when the current handoff fits, otherwise `None` | not applicable | the last valid handoff, or `None` |
| release 8 MiB | `None` after the preceding release | `Some(capacity)` when it fits, otherwise `None` | `None` | `None` |

The capacity must never exceed the configured ceiling. A zero, released, or
over-ceiling route recomputes at commit and cannot refuse a supported edit or
change its bytes. The default route may reduce the final duplicate render but
can increase transient peak live memory when the previous token overlaps the
next candidate. The 0730 two-edit capacities (36,864 bytes for NoHeadFoot and
587,776 bytes for FloatingPictures) are historical observations, not values to
hard-code into the new capture.

## Correctness and preservation gates

The functional qualification should run before timed or allocator children and
must prove, for each fixture and N in `{1, 2, 8, 32}`:

* zero, default, and release routes produce byte-identical output, with each
  route compared to an independently run zero-retention sequence of the same
  N;
* the output reopens through the strict owner and the independent public
  reader, and the full projected paragraph list has the expected parity result
  and unchanged survivors;
* the source hash/length and case identity match the frozen fixture manifest;
* every requested stage is a real change and `Edit::changes()` contains N
  semantic changes;
* forward patch application equals the committed snapshot, inverse application
  restores the exact source bytes, and a foreign/stale source is refused;
* an exact no-op returns the exact source bytes and preserves source allocation
  sharing; a same-value operation after a changing stage preserves the current
  token; and an operation-limit refusal leaves the prior staged bytes and
  retained capacity unchanged;
* untouched CFB stream bytes, directory metadata, storage identity, and the
  existing `Reuse` preservation behavior satisfy the 0730 raw-directory oracle;
* zero and release retention state transitions are exact, default capacities
  obey the ceiling, and retained output is the allocation already produced by
  the successful common candidate rather than a copy.

These gates preserve the source-bound `Arc`, clone-first failure atomicity,
transaction-policy intersection, protection and signed-source authorization,
typed refusal behavior, and the final public semantic validation. They also
keep the retained state private to the current editor: no token may enter a
`Snapshot`, `Patch`, transfer plan, common editor, or process-wide structure.

## Measurement interpretation

The allocator lane may report allocated/deallocated bytes, successful
allocation calls, peak live bytes, and the output allocation capacity retained
at region end. `retained_bytes` at allocator-region exit is a live-byte delta,
not the `RetainedRender` capacity. Neither metric is an RSS bound. If whole
child RSS is collected, it is a separate observation and must not be charged to
the 8 MiB retained-state ceiling.

The timed lifecycle may include source clone/open, edit construction, all N
mutations, optional release calls, commit, output copy, and local owner
destruction, while excluding file I/O, preparation, correctness assertions,
and JSON serialization. Keep the fixed fixture/edit/route schedule and paired
route orders from the frozen protocol. Report N scaling and route deltas as
descriptive evidence; the two-edit 0730 result does not establish a many-edit
speedup or a universal DOC claim. No cold-storage, remote-source, concurrent,
native-Office, other-producer, or iWork conclusion follows.

## Authority

The governing constraints are [`GOAL.md`](../../../GOAL.md), the accepted ADR
index and hierarchy in [`adr/README.md`](../../../adr/README.md), especially
ADR 0003's immutable/atomic/patch boundary, ADR 0005's retained-state
amendment (lines 2344–2390), and ADR 0006's preservation and validation rules.
The 0729 recommendation is at lines 102–113 of
[`0729-doc-public-phase-attribution.md`](../../0729-doc-public-phase-attribution.md);
the retained 0730 disposition and memory limits are in lines 1–8 and 41–79 of
[`0730-doc-bounded-render-handoff.md`](../../0730-doc-bounded-render-handoff.md).
The 0906 legacy recommendation at lines 71–78 is superseded by 0730. Current
implementation review details are retained in
[`change-0730/implementation-review.md`](../change-0730/implementation-review.md).

## Initial read-only review of the 0917 probe (superseded below)

The current adaptation at [`probe/src/lib.rs`](probe/src/lib.rs) has the
required measured ownership boundary. `run_case` clones and opens the real
source, constructs one public `Edit`, performs all N replacements and any
release calls, commits, copies the published snapshot bytes into its result,
and returns only after its local snapshot, editor, and commit have been
dropped. The native clock surrounds that whole function. In the allocator
lane, `alloc_metrics::region` observes the result while its output `Vec` is
still live; the fixed stage array is on the stack, and allocator validation
checks retained bytes against `output_capacity`. Oracle, hashing, public
readback, and JSON work occur after the boundary. This is the right
interpretation of end-retained bytes: it describes the returned output
allocation, not the private retained-render token.

The current functional route checks N semantic changes, final paragraph
through `Snapshot`, independent `Package` reopening, full survivor
projections and collections through the copied 0730 oracle, cross-route exact
bytes, untouched streams, CLSIDs, and the raw-directory policy gate. The
capture validator repeats per-sample output, N-count, stage-state, and
allocator-output-capacity checks. The 0730 oracle's `check_doc_output` helper
still constructs an internal `DocNoHf` case for both DOC fixtures; both current
DOC cases take the same allowed-stream and semantic branch, so this is not a
current matrix mismatch, but the helper should remain fixture-neutral if the
matrix gains a different DOC policy.

At the initial review, two qualification gaps remained before treating a
passing capture as complete evidence. The exact no-op route at `prepare` checked source
bytes and `Arc` sharing but immediately discards the `Commit`, so it does not
assert `changed() == false`, `patch.is_noop()`, or zero patch changes. It also
exercises that no-op only with the zero-retention source policy. Retain the
commit long enough to assert those public no-op invariants and require no
retained token for each selected retention policy, outside the timed route.
The operation-limit refusal keeps the previous capacity and checks only that
paragraph zero still reads FIRST after committing; a stronger atomicity gate
should compare that committed snapshot's exact bytes and patch change count to
an independently committed one-operation control. This catches corruption in
survivor streams that paragraph-zero readback cannot see. The typed operation
refusal itself is present and occurs before the second mutation.

These were harness qualification gates, not production optimization findings;
the final review below records their closure.

## Final source review after the probe fixes

The two qualification gaps are now closed in the current probe. The no-op
control retains its `Commit` and asserts `changed() == false`,
`patch.is_noop()`, zero semantic changes, exact source bytes, and source
allocation sharing. The same-value-after-change control proves that the
second same-value request does not change the retained capacity and that its
commit is byte-identical to an independently committed one-operation result.
The operation-limit control now commits an independently prepared one-change
control and requires the refused transaction to have the same exact bytes and
patch count, while retaining the pre-refusal capacity and matching the typed
`OperationLimit` refusal.

The trimmed `oracle_0730.rs` preserves the DOC path of the accepted 0730
oracle. Its inventory, structural validation, directory shape and metadata,
normalized raw-directory comparison, changed-stream allowlist, untouched
stream comparison, CLSID checks, public `Projection::All`/Accepted/Rejected
survivor comparisons, story/table/field/revision/object collections, and
`oracle_ok` failure conjunction are unchanged for DOC. Removed code is the
PPT path, old benchmark runner, and format/container controls. The only DOC
specialization is making the former `DocNoHf | DocFloat` branches direct and
making the former `Format` raw-directory policy explicit; both are the only
call path in this probe and therefore preserve the prior behavior.

No read-only design blocker remains for the frozen 0917 matrix. The probe is
ready for the separately controlled source export/build and capture gates;
this review establishes no timing or performance result.

## Final frozen no-op scope note

After the above review, the final probe broadened the untimed exact no-op
control to all three retention policies. The frozen `lib.rs` hash is
`b932992f10ecee244eaa3bd14833a07dcd2456f0bfc2de0dce3bf64403737dc3`.
The final exported build and all 48 qualification cases use those bytes; the
earlier zero-policy-only description records the reviewed intermediate state.
