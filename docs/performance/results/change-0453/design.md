# 0453: Share managed decoded payloads in PPTX copy plans

Previous turn made progress in commit 456e19246. Initial worktree contains only
user-owned GOAL.md untracked. Accepted ADR tree remains
c950b6c8be822561b498d7bbe87c460873dcbf49, read previously in full; README and
memory/performance contract rechecked. No owned build/test/profile job is live.

Hypothesis: a plan holding a verified compressed capture can retain ordinary
managed PartData instead of cloning decoded media. Keep source reservations,
exact semantic identity, chunked cancellation and rerun validation. Preserve
aggregate destination staging admission as the bound for a lazy decoded fallback;
this batch does not claim reduced budget admission. A memory-only compressed
publication refusal may allocate that already-budgeted fallback copy.

First extend the standalone harness with operation allocator regions. Build both
normal and instrumented baseline binaries with this same harness before editing
production. Then implement the internal payload representation and compare
matched complete API times, allocation calls/bytes and region live peaks. Keep
normal timing and instrumented diagnostics separate; do not infer allocator
peaks from fixture-dominated process RSS. Retain exact source and binary custody.

Independent agents audit payload ownership and implement the harness field.
Root owns production and serial CPU jobs. No Rust edits while any owned CPU
job is live. Full non-iWork scope remains active, including native breadth,
cold/scaling, bounded existing append, repackaging and broader CRUD.
