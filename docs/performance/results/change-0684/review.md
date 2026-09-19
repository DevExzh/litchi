# Independent review: XLS retained occurrence index

The source/cache reviewer found no semantic blocker after reviewing the final
`COLLECT_INDEX` specialization. Publication follows `scan_worksheet` and
`finish_scan`; failed scans drop candidates, and pre/post-sort cancellation
prevents publication while preserving the already-computed result.

Replay preserves duplicate source order, typed invalid SST values, earlier
SST/formula errors, XF checks, formula STRING/CONTINUE and permitted metadata
handling, leading/trailing fences, and missing-target behavior. Ordinary scan
specialization discards collection work without bypassing structural checks.
Visitor/text paths bypass the cache.

Managed charges balance across publish, duplicate-winner loss, allocation or
admission failure, eviction, owner drop and pinned Arc lifetimes. Actual slot
and reservation-vector capacity growth is charged before retention. Opening
has no ExecutionContext: optional table storage is fallible and charged to the
local logical ceiling; managed candidate/index allocations use Resource::Memory.
The fixed 128-byte overhead is conservative, not exact allocator accounting.

Incremental pressure can evict clean entries before a later candidate fails.
This affects cache quality only. Retry after optional refusal remains possible;
local budgets too small for a full index may therefore repeatedly attempt and
abandon collection. No error result is memoized.

Root additionally repaired SST message identity, pending-formula EOF identity,
all-sheet hotness sizing, LRU touching and capacity-release ordering; strengthened
private pin/pressure/concurrent-publisher tests; and adjusted the older scan-I/O
test to explicitly disable caching. All final owner/facade quality gates passed.

Final independent evidence review recommends retention with scoped experimental
wording (`performance_claim: none`). No paired open/q1/visitor p50 regression
exceeds 5%; construction, tiny-file warm hits, repeated refusals and RSS costs
remain explicit in the change record. The first candidate's corpus parity and
quality passed, but ordinary-query regressions prompted its archived revision.
Baseline source hashes are now audited against immutable git objects; available
binary hashes are verified before temporary build cleanup. Old failed focused
and initial-quality logs are superseded by `final-verified/`, not final failures.
