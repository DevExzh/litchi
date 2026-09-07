# Change 0457 final code review

This is a read-only review of the frozen source-backed ODP append path and its
common bounded XML/ZIP publication code.  The review used the accepted ADR
refresh at `48f11c7745f8f798a00e11a14f1f34dc1c9c7954`, `docs/GOAL.md`, the
accepted ADRs, and the retained 0457 integration notes.  No build, test, fuzz
run, or performance capture was performed for this review.

## Finding status

The previously reported retained-fragment accounting blocker is closed by the
current source.

`SourceContentInsertionPlan` retains an `AuthoredXmlFragment` in its public
plan state (`crates/litchi-odf-common/src/core/source_publication/insertion.rs:183-199`).
The fragment wrapper retains the caller's `Vec<u8>` capacity
(`crates/litchi-odf-common/src/core/xml_splice.rs:31-35,142-155`).  During
`prepare_with_validator` now measures that retained capacity and reserves it
before any source access (`insertion.rs:238-252`).  The `Option<Reservation>`
is stored in the plan (`insertion.rs:189-199,381-392`), so it remains live until
the plan is dropped.  Transient XML/index memory is reserved separately
(`insertion.rs:277-287`), and `write_to` independently reserves the retained
capacity under the supplied execution context (`insertion.rs:481-494`).

The ODP owner has the same lifetime boundary: its generated-page reservation
covers generation and semantic readback (`crates/litchi-odp/src/package/source_append.rs:431-450`),
the title/body strings are dropped (`source_append.rs:451-457`), and the
authored fragment then enters common planning (`source_append.rs:459-479`).
The common plan now owns the persistent retained-fragment lease
(`source_append.rs:481-498`).  The focused insertion tests cover refusal before
source reads and reservation release on plan drop
(`crates/litchi-odf-common/tests/source_content_insertion.rs:821-905`); the
retained-plan test run reported all 13 focused tests passing.

The early-return ordering is also correct: the fragment local is declared
after the preparation lease (`insertion.rs:244-252`), so an error drops the
backing allocation before the lease releases.  This closes the prior violation
of ADR 0005's hierarchical memory charge and the 0457 requirement that managed
`Memory` remain live through a prepared plan.

There is a deliberately short phase handoff in the ODP owner: the generation
lease ends at the block after the authored fragment is audited, and common
planning immediately acquires the retained-fragment lease.  The only work
between those leases is bounded `CandidateValidator` construction and scalar
limit setup; no source read or unbounded allocation occurs there.  I do not
classify this as a blocker.  Keeping the full generation lease through common
planning would conservatively overlap the complete readback envelope with the
common transient envelope and could reject valid operations solely because of
double charging.  If a future contract requires a continuously nonzero gauge
across phase handoffs under concurrent sibling operations, the generation lease
should be transferred or kept through the handoff rather than adding a larger
overlap.

One useful, non-blocking regression remains: explicitly prepare an oversized
fragment without an execution context, then call `write_to` with a managed
context and assert that the write reserves the fragment capacity before source
reads and releases it afterward.  The implementation already performs this
publication-local reservation; the existing tests exercise managed preparation
and managed publication separately.

## Areas with no additional verified blocker

The ZIP replay path performs target, callback measurement, framing, layout, and
archive-size checks before the first destination write
(`crates/soapberry-zip/src/preserve/replay.rs:374-475`).  It invokes the
callback once for measurement and once for emission, fixes Deflate input and
output windows, and reports non-determinism or sink failures with typed replay
progress (`replay.rs:429-505`, `1250-1385`).  The direct sink error handling is
consistent with the `Write` contract; an over-report is conservatively marked
indeterminate.  I found no separate replay atomicity or partial-output defect.

The common XML scanner and the ODP source/candidate visitors keep the source
and candidate reads bounded, reject malformed XML/control/ref/QName cases,
enforce namespace and expanded-attribute checks, and perform source/version,
encryption, ZIP-layout, hash, and length checks before and during publication.
The checked sink preserves exact accepted-byte progress and typed cancellation,
source, limit, transport, and callback causes.  I found no additional
verified XML grammar, security-boundary, or partial-output blocker in this
source-only pass.

## Specialized API contract

`SourceBackedTailAppendEdit` is a deliberately narrow ODP capability.  It
accepts one generated plain title/body slide, proves the existing positional
ODP presentation, inserts one authored `draw:page` into the decoded
`content.xml` stream, and emits a new sequential ZIP artifact through the
common source-preserving replay path.  The source must satisfy the owner’s
unencrypted, unsigned, canonical ODF MIME/framing, member-name, compression,
and trailing-byte policies; XML limits, page-count/name/position checks,
semantic readback, source version, source/candidate hashes, and cancellation
remain part of the contract.

The result is a specialized forward publication plan.  It is not an ordinary
`Snapshot`/`Edit`/`Commit`, does not produce a reversible `Patch`, does not
materialize or promise a complete candidate package, and does not provide a
general ODF or iWork append operation.  A caller-owned non-seeking sink may
receive a typed `Untouched`, prefix, complete-unflushed, complete, or
indeterminate result with the accepted-byte count when publication fails.
When prepared with a managed execution context, the plan's retained authored
fragment capacity remains charged until the plan is dropped.  An unmanaged
preparation has no retained budget lease; a later managed `write_to` acquires a
publication-local capacity reservation for the call.

## Disposition

No verified production blocker remains in the reviewed source.  The retained
integration notes and native/fuzz/performance receipts remain evidence only for
their stated scopes; this review makes no broader performance claim.
