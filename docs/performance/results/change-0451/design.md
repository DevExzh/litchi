# 0451: Cache-aware OPC decoded read and transfer authorization

Previous goal turn: progress, committed 0450. Initial worktree has only user
GOAL.md untracked; no owned CPU job is live. Accepted ADR tree is unchanged
from the earlier complete read. Preserve all source, semantic, budget and
publication contracts; this batch supplies OPC adoption before PPTX integration.

Add PartView::data_and_authorize_precompressed returning managed PartData plus
an OPC-authorized compressed token. Reserve compressed capture/writer staging
before entering the existing cache. The elected loader uses the combined ZIP
primitive and publishes decoded bytes through the existing flight/rollback path.
Hits and waiters reuse their decoded allocation and perform fresh compressed
verification against it. Normal data/accounting/session behavior remains shared.
A token pins the decoded payload's memory/object reservations as well as compressed
staging so dropping package/data handles cannot release its live-byte charges.

Charge C+U work exactly once for fused cold capture (C at authorization and U at
cache admission), and C+U for a warm capture (U before verification). Preserve
freshness/error precedence, source-checked publication, contextual cancellation,
limits, same-Part coordination, bypass behavior and caller-owned non-seek sinks.

Validate cold/warm exact output and I/O counters, budget boundaries/drop lifetime,
concurrent ordinary waiters, provisional cancellation/source change rollback,
source mutation during capture, typed refusals and partial output. Use deterministic
read-work comparisons as an enabler gate; do not claim PPTX latency, allocator peak
or native/scaling gains before complete format adoption and matched measurement.
