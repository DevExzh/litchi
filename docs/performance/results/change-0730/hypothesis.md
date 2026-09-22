# 0730 bounded DOC validated-render handoff pilot

0729 attributes 13.00–13.30% of the instrumented small DOC lifecycle to finish;
large-DOC observer controls were noisy. Test removing the duplicate finish
render with ordinary default-feature public workflows, without transferring
instrumented fractions into a speedup claim.

The common owner may synchronously return its already validated batched render.
The DOC owner may retain that existing allocation only under the accepted
ADR0005 retention amendment: finite transaction ceiling, capacity accounting,
observable held bytes, explicit release, no extra refusal on overbudget, and
recomputation fallback. Preserve atomic publication, exact no-ops, independent
final strict/public validation and the existing Reuse policy. A token must be
bound by ownership to the exact package state; mutation must invalidate/replace
it. No general CFB cache or policy switch is proposed.

Capture baseline before production edits. Then freeze both binary identities,
source diffs, probe, fixtures, oracle, and schedule before comparison. Retain all
samples, failures, and controls. Use two exact 0728 DOC cases and the same
45-UTF16-unit paragraph-zero replacement, strong direct semantic and unknown
stream/raw-directory checks. This two-case pilot cannot establish broad producer
coverage; existing format tests and additional retention tests must precede any
adoption. Preserve candidate source if rejected.

Prospective comparison: CPU12, three cycles, both cases alternating order,
A/A baseline controls followed by A/B/B/A within each case/cycle; each process
three warmups and fifty samples. Separate allocator lane uses A/B/B/A per case
with one sample and no warmup. Do not use allocator-instrumented timings.
Report process p50/mean/p95/p99/max, all paired changes, allocation calls/bytes,
peak live and retained bytes. Native control and candidate regressions over5%
in p50 or mean are review flags; primary small-fixture median improvement must
exceed5% with consistent direction across cycles to justify complexity.
Allocation/peak increases over5% require explicit review. Do not hide any flag
in an aggregate or rerun selectively. RSS is not inferred from allocator data.
No default adoption without correctness, ownership/retention, representative
corpus and performance review. This packet may close with a rejected candidate.
