# 0732 PPT commit phase review

This is a bounded source review of the feature-gated PPT diagnostic copy and
its phase probe. It does not run Cargo, native controls, or the profiler. The
ordinary PPT `Transaction::commit` body remains the reference contract.

## Verdict

### Source verdict: pass

The implementation is behaviorally aligned on the reviewed success, no-op,
formatting-only, patch, inverse, and live-read error paths after the lifetime
corrections made during this review. One critical lifetime finding was fixed
in the frozen source: the `DocumentCommit` wrapper now keeps the consuming
`if/else` expression outside one enclosing closure and observes each branch
separately. Cargo/native qualification remains pending.

## Critical finding fixed before quality1

The pre-fix `commit_profiled` wrapped the complete document commit conditional
in one closure (`crates/litchi-ppt/src/slide_order.rs`, around
`DocumentCommit`). The success branch consumes `self.document` through
`self.document.commit()`, so closure capture moved that field into the closure
even when the rebase branch was selected. In the rebase case the unused
document transaction was then destroyed when the observer closure returned.
The ordinary `commit` leaves that partial field move in the transaction's
ordinary scope and therefore retains its destruction boundary through the
rest of the function. This changed lifetime/allocation attribution and could
change drop order. The frozen source keeps the conditional outside the
closure and observes each consuming branch at the existing expression
boundary, preserving ordinary partial moves while emitting one
`DocumentCommit` start/finish pair.

The `LiveDocumentRead` comment now states the actual combined read and source
consistency check. In addition to the synthetic mismatch case, the feature
tests now set the source persisted-record identifier to an absent value,
compare ordinary and profiled error display/classification, and require the
actual `persisted_record` failure to close with `Finished(Error)`. This covers
the stronger read-boundary failure requested by the review.

## Checked contract

The ordinary sequence is preserved: document commit; before-slide payload
capture; embedded editor open; live persisted-record read and source check;
inserted-record staging; the `document_commit.snapshot().bytes().to_vec()`
candidate; picture handling; editor finish; unrelated-stream validation;
public reopen; after-slide capture; expected payload comparison; transferred
picture validation; and the exact-source reversible patch. The profiled copy
keeps `document_commit`, `source`, and `working` in the same outer scope. The
successful live `Vec` is returned from its phase as outer `_live`; the before
and after payload vectors remain outer locals; and artifact hashes stay inline
in the patch literal. These corrections avoid closure-local destruction and
hash/snapshot evaluation reordering.

The structural branch is named `StructuralNoOp`. Its condition remains the
existing `document_commit.patch().is_empty()` condition, so it does not claim
that the entire root transaction is unchanged. The formatting-only test checks
that this branch still produces changed bytes and an equal ordinary/profiled
patch. The exact empty edit test checks source-byte pointer sharing, empty
patch behavior, and forward/inverse application. The ordinary commit has no
embedded writer work on this branch.

`observe_phase` emits synchronous content-free `Started` and `Finished`
events, and operation errors in wrapped phases close with `Error`. The live
read test now exercises the existing typed source-mismatch error and checks
the `LiveDocumentRead` error event. Errors in explicit residual work after a
read-only phase remain uninstrumented and are documented as such; the phase
contract does not claim that every commit error has a phase error. The callback
receives only copyable phase data and cannot alter the transaction, bytes,
patch, validation, or publication result. A callback's own panic is not a
format operation error and is outside the balanced-operation guarantee.

The feature-off path adds no call or observer to ordinary `commit`; the new
types, method, and tests are cfg-gated. The probe keeps ordinary opaque and
ordinary split controls beside profiled-empty and profiled-clock controls,
uses a fixed-capacity content-free event recorder, checks event sequence,
balance, outcomes, and commit-window containment, and validates each output
through the existing PPT inventory/oracle. The initial quality run exposed
only two probe test-compilation string conversions; those are now explicit
`String` conversions. No allocation, RSS, speedup, or phase-cost conclusion is
admitted until root retains the serial Cargo/native receipts and four-route
controls.

## Review scope and limits

The review is limited to the requested PPT commit copy, phase seam, no-op and
patch behavior, observer behavior, local ownership, and probe contract. It is
not a full PPT source audit. Root must re-review the post-fix source custody
map, then run the required feature-off/feature-on Cargo gates and the native
four-route schedule before treating 0732 as qualified.
