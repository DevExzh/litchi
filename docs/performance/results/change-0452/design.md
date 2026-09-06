# 0452: Retain source-authorized captures through reusable PPTX plans

Previous turn made progress in commit 3778dc08b. Worktree starts with only user
GOAL.md untracked; owned jobs are terminal. Accepted ADR tree remains
c950b6c8be822561b498d7bbe87c460873dcbf49, previously read in full.

Split OPC compressed capture and writer reservations without changing their
combined admission amount. Convert a consumed token to an opaque shareable
retained capture, releasing writer staging. Each publication obtains a new
source-checked token and its own writer reservation. Keep decoded pinning and
private ZIP types. No derived Clone on a consumable publication token.

PPTX first preparation captures images/charts along with decoded reads. It
retains captures beside existing staged byte allocations and reuses them only
when the existing metadata/byte equality check reuses that allocation. Planner
rerun, touched digest, candidate rereads, source lineage/version, chart semantic
validation, and partial-output handling stay intact. Memory-only fallback keeps
the old decoded path available; other typed failures remain fail closed.

Before edits, build and retain an exact baseline harness executable. Validate
repeated/concurrent publication reservations, source freshness, byte preservation,
budget failure cleanup and full affected suites, then compare complete matched
PPTX lifecycle samples and provider work. Do not infer latency from I/O counts or
promote native/default coverage without corresponding evidence.
