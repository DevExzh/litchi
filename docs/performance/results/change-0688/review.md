# Independent review and disposition

Retain the frozen candidate with a scoped disposition. The independent source
reviewer finds no correctness blocker in the implementation, focused tests or
actual assembly. The independent profiler recommends retention after checking
the paired native results, longer controls, counters, RSS and code costs; root
agrees. The independent numerical/scope review found no material discrepancy;
all final evidence and documentation gates and source-binding audits pass.

The production helper preserves the exact marker/conversion/index/next-marker
order and errors. Explicit `usize::try_from` remains; on this x86_64 target it
lowers to an identity conversion. `ENDOFCHAIN` and delayed next-index failure
retain their prior behavior. Two bounded differential tests include arbitrary
resume states and no-loop walks, with independent literal error assertions.
The legacy `file.rs` helper, cursor source loop, source fences, hints and
checkpoints remain unchanged. No new state, allocation policy or unsafe code.

Assembly confirms an actual mechanism: baseline calls the checked link helper
for every step and writes/loads its successful `Result` through a hidden return
pointer. Candidate links inline inside a whole-walk function. The cursor still
calls that function once per walk and it still returns a `Result`; it is not
call-free. The inspected fragments total 1,462 → 1,549 bytes and whole `.text`
grows 1,440 bytes (0.23%). Frame reservations change at several levels, so
local frame shrinkage is not a measured peak stack reduction.

The retention scope is repeated selected-cell queries with meaningful
CFB walks. Paired native medians, independent longer loop controls, matched
profiles and instruction/cycle diagnostics agree on the long-chain gains.
Large indexed open-plus-eight windows are approximately flat. Formula-refusal
windows improve; errors remain uncached.

The costs are retained: disabled-index owned 54016 workflows rise 3.77–4.72%;
Simple stored owned loop means rise 3.27–3.68%; Simple missing owned q8 rises
70 → 80 ns (+14.29%) in both pairs. The latter has no chain lookup, and its
cause is not isolated. Code growth, all mean/tail triggers and RSS are reported.
No measured diagnostic median RSS increase exceeds 5%, with the largest +4.52%;
this does not establish an RSS or stack bound.

All 96 query-allocation groups and 12 counted I/O routes are exact. Six quality
gates plus DOC/PPT consumers pass: 4,394 passed, zero failed, 27 ignored.
The real/generated corpus checks preserve semantics. Broader cold-device,
remote, concurrent, cross-platform, malformed-chain performance and native
Office evidence remain open. `performance_claim: none`; no coverage promotion.
