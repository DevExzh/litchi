# Candidate review and design repair

The final candidate finds the first missing folded ancestor by borrowed lookup.
The private ancestor map is prefix-closed: each accepted Part adds every proper
ancestor before insertion completes. Once one ancestor is absent, all deeper
ancestors along that path must also be absent. Reserve exactly that suffix and
stage its owned strings. An entirely shared path or root name stages nothing.

The original conflict checks, names/by_folded/descendant reservation order,
final descendant try_reserve(ancestors.len()), and mutation sequence remain.
The first descendant representative stays unchanged. There is no new cache,
retained metadata, hash policy, API, source-validation bypass or unsafe code.

Three tests cover case-folded shared ancestors and first representative;
zero-capacity deep/distinct/partly-shared insertion and prefix closure; and
all four conflict forms leaving all maps unchanged before a valid insertion.
Allocator fault injection is not added; fallible staging/reservation ordering
is checked directly in the unchanged publication sequence and by review.

An archived initial draft was never applied or built. It used per-key Vec
reservation, which could add deep-unique growth allocations, and its handwritten
patch failed git apply --check. Earlier review also rejected an intermediate
proposal to reserve one map slot repeatedly before changing its length: that
would not reserve the whole staged batch and changed error precedence. The
final generated patch retains the original final batch reservation. No formal
workload or production behavior came from these drafts.
