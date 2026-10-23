# 0362wb — PPTX native snapshot route accounting

The PPTX facade has one native `Commit` shape for eager and source-backed
edits. Several older native routes initialized `snapshot_bytes` to zero and
left the proof walk to the transaction wrapper. That made the wrapper's
snapshot limit depend on the route instead of the retained native proof.

`Commit::snapshot_bytes()` now lazily obtains the cached `Proof::retained_bytes`
charge when the route did not set an eager value. The existing setter and
coalescer paths still use `measure_snapshot_bytes()` and preserve their
staged/canonical maximum for signed no-ops. The change covers structural,
ordering, removal, picture, media, comments, charts, masters, SmartArt, links,
table, notes, and slide-text routes without duplicating proof ownership.

The coalesced PPTX editor now also passes the transaction's remaining
snapshot budget into native source/eager edit options. Source-backed edits
apply that ceiling to their exact encoded retained-patch preflight; eager edits
use the opened editor's native patch limit (while reusing the admission edit
only when its default bound is no larger than the remaining budget). The old
`source_bytes + size_of::<Patch>()` check was only a lower bound and is gone;
the source-length check remains as an allocation guard, while native capture
owns the exact before/after proof. Signed canonical no-ops use the same
bounded native empty edit and retain the larger staged/canonical charge.

The regression in `crates/litchi-ops/tests/pptx_shape_remove.rs` applies a
one-byte snapshot limit to native shape removal and verifies a typed limit
refusal leaves all source shapes intact.

Source: `076ea53aad9e67205dca3323aacdcfaabad98aaf`.

Validation was limited to rustfmt, diff checks, and repository artifact guards
while the shared workspace Cargo lane was active. The coordinator must run the
focused PPTX tests and the workspace consumer gate after merging this patch.
