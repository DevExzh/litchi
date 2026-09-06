# Candidate and evidence review

The root applied the source-only coder and tester drafts and ran every build,
test and workload serially. An independent source reviewer found no blocking
issue in the applied candidate or its seven private differential tests.

The candidate groups four fixed-markup sequences into a 256-byte stack scratch
buffer. It charges the existing Work budget once and writes into the existing
preallocated fragment buffer. Aggregate local-limit or Work-budget refusal
falls back to the original pieces, preserving first-failing-piece progress and
ancestor rollback. Cancellation and other execution errors propagate; a sink
write that has begun is never retried through fallback. The root preserved
the original page-prefix/error ordering on ordinal overflow. Cooperative
cancellation may follow completion of an already charged bounded batch.

Production fragment capacity and the SlideWriter byte limit are identical;
the successful preflight fits the complete batch in that preallocated buffer.
Public caller-sink progress still belongs to the common BudgetedOutput owner.
The public sink-failure tests are included in the complete ODP release suite.

The initial outer evidence adapter incorrectly required equality of absolute
process live/high-water allocator vectors. The pilots exposed a two-byte
executable-path-length offset and differing elapsed-sort order. Before formal
capture, the root corrected the comparison to chronological operation counters,
live deltas, and region peak above entry. Raw vectors remain in every report
and summary. The original scripts, failed check, and correction are retained
in draft-history. The copied semantic oracle did not change.

Both builds were made at the same Git HEAD, with the candidate uncommitted.
Source manifests and executable hashes identify the implementations. The
verifier requires exactly the two candidate source paths to differ and both
executable hashes to differ, rather than requiring different Git revisions.
The earlier adapter version is retained. No build or capture receipt was
rewritten to manufacture a different revision.

The pre-implementation hypothesis and retention gate are immutable. These
adapter corrections do not lower the 5% median-improvement threshold, relax
the mean-confidence-interval condition, or permit higher operation allocations.

The evidence reviewer confirmed chronological operation-delta comparisons and
the exact same-HEAD source-delta proof. Root strengthened release argv/env and
custody-driver checks, source-path/digest validation, and pilot verification.
Absolute process live bytes remain descriptive and are excluded from relative
cost comparisons. All 66 matched comparisons and repeat checks have no flags.
The immutable capture protocol still lists `allocation.live_bytes_after` in
its review-metric list. This analysis correction intentionally excludes that
absolute process value from relative comparisons; it remains in every raw
report and summary. The protocol hash and predeclared acceptance rule remain
unchanged.

Lifecycle integration corrected path-keyed pilot artifact records, historical
gzip sidecar reads, and suffix-free expected-check keys. The original adapter
and failed checks/replay remain retained. The final portable checks rederive
the hypothesis, summary and rejection, validate candidate source snapshots,
and reject eight independently resealed evidence mutations.
The final artifact adapter binds paths to receipt keys even if a nested record
contains a conflicting path; the preceding successful replay driver is retained.
