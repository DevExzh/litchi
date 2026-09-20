# 0704 accounting review

Status: conditional approval of the accounting and ownership design, pending a
deterministic reserve-failure test. This is a bounded source review only; no
Cargo command, production edit, or commit was made by this reviewer.

## Scope and conclusion

The retained value is correctly separated from the existing MCE processing
limit and from the serialized cross-slide retention budget. The published
`RetainedMce` charge is checked, observable, and releasable. The raw allocation
key is protected against ABA by both the `(pointer, length)` lookup and the
strong raw `Arc`; final publication accepts raw owners only from the current
package's `PartDigests`. `Snapshot::rebound_to` projects both memos onto the
new package owners, so equal content in a newly allocated part does not create
a false hit.

The first implementation had quadratic capture bookkeeping. The current
append, running-charge, sort, and deduplicate path removes that concern: parent
hits use the immutable table's binary search, visits append once, and
finalization sorts once. Its dominant bookkeeping is O(P log P), where P is
the number of admitted visits, with one linear owner projection. This is
appropriate for the bounded slide catalog and avoids adding a per-capture hash
map.

## Final charge and transient storage

`RetainedMce::from_candidates` charges the final `Vec` capacity, every retained
processed vector capacity, the `RetainedMceEntry` metadata, the processed
`Arc` allowance, and the outer `Arc` allowance. All additions and the entry
capacity multiplication are checked. It trims the candidate prefix when the
actual final vector capacity would exceed the policy, so
`Snapshot::retained_mce_bytes()` cannot report a charge above the policy. The
raw payload is correctly excluded from the added charge because the snapshot's
package already owns it; retaining the raw `Arc` only supplies identity and
keeps that allocation alive.

The current `MceCapture` also keeps a checked running charge for each admitted
visit. Counting duplicate visits conservatively is safe: finalization
deduplicates by allocation key and charges the unique final table once. The
capture route checks the slide catalog against `Limits::max_parts` before it
creates the capture, and the presentation catalog has the 100,000-slide hard
bound, so pending visits and their source witnesses are finite independently of
the retained-byte setting.

That running charge is a bound for optional retained-output admission, not a
peak RSS bound. The pending `Vec<PendingMce>`, its allocator over-reservation,
the final candidate vector, and a parent table can overlap temporarily. Their
metadata is not included in `charged_bytes`; this is acceptable only if the
record continues to describe `retained_mce_bytes` as the published value's
logical charge, with transient operation storage governed by the existing
slide/part bounds. The implementation comments now state that distinction.
Do not describe the one-MiB policy as a peak-memory bound. If the policy is
intended to bound the whole operation's live memory, the pending and
candidate-vector capacities must instead be charged or explicitly reserved
under a separate operation budget.

## Ownership, lifetime, and rebound checks

The second `Part::blob()` observation remains the MCE source witness. A parent
hit requires the exact pointer and length and retains the source allocation in
the old table. At finalization, `PartDigests::owner_for` supplies an owner only
for a current package allocation that passed the existing `blob_arc` alias
gate. This closes the foreign-owner and allocator-address ABA cases without a
new `blob_arc` observation in the slide loop.

`capture_internal` computes the complete-package digest memo over the owned
package clone before publishing MCE entries. A failed owner projection or cache
reservation returns an absent/partial optional memo and leaves semantic capture
errors unchanged. Changed transactions pass the source snapshot's MCE memo;
the no-op transaction path returns that source snapshot. `rebound_to` first
projects `PartDigests`, then projects MCE entries through those verified owners,
and is reached only after the existing equal-package and equal-policy checks.
Snapshot, transaction, and commit release methods clear only their own memo
reference. The package facade retains no ambient MCE table, so ordinary facade
mutations cannot leave a memo describing an obsolete graph.

## Validation and debug oracle

Every reused parent output is still checked against the default MCE input and
output limits, then reprocessed by `debug_assert_cached_mce` in debug/test
builds and compared byte-for-byte. Root/name/notes-proof and relationship
validation continue to run over the current capture. Fresh outputs need no
second oracle because they came directly from `process_ooxml`; a projected
output is checked again on its next hit. No MCE error or typed refusal is
stored.

The current focused tests cover charge boundaries, overflow arithmetic,
foreign/non-aliasing owners, rebind invalidation, release, changed-slide
reuse, and refusal precedence. The new private
`from_candidates_with_reservation` seam deterministically forces the final
table reservation to fail with `CapacityOverflow`; the focused test verifies
that the optional table returns `None` and both raw and processed candidate
owners are released. Together with the ordinary disabled, over-ceiling, and
semantic-parity cases, this is sufficient narrow evidence for the final-table
fallback without a process-global allocator hook or a broad injection seam.

The pending `Vec` reservation still follows the same explicit boolean fallback
branch in `admit_visit`; it is covered by source inspection and the normal
over-ceiling path. A separate pending-reservation seam would add confidence,
but is not a material review blocker if the acceptance gate treats one
deterministic cache-only reservation failure plus the semantic parity matrix as
the required fallible-reserve evidence. Keep the existing documentation that
`Arc::new` remains the surrounding library's infallible allocation primitive;
the seam is for explicit `try_reserve` branches only.

## Review disposition

No source-lifetime, owner-finalization, ABA, final-ceiling, debug-oracle,
algorithmic correctness, or final-table reserve-fallback blocker remains in the
current append/sort revision. The explicit distinction between final logical
charge and transient capture workspace remains a documentation/accounting
condition; it is not a claim that the one-MiB value bounds peak RSS.
