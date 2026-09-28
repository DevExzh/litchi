# 0802 — bounded linear prefix without value replay

This fresh candidate targets both costs exposed by 0801: empty consumption and
an ordered map started after only three successful attributes. It retains the
no-replay first/second/third handling, adds an exact-empty-tail shortcut, and
uses a bounded linear name collection before the ordered-map handoff. The
linear stage is capped at 32 names; the candidate must retain worst-case
bounded checking, duplicate-before-value errors, exact positions and fusion.
Production remains unchanged during this experiment.

The plan preserves the exact 39-case fixture matrix, 18 protected consume cases
and advancement thresholds used in 0801 and 0799. Fresh native measurements
use six paired blocks, 30 samples, three warmups and 4,096 iterations, pinned
to CPU 12. The bootstrap seed is 802080. Two separate Callgrind repeats retain
guest instructions and branch counters. These are hot repeated direct-helper
inputs in one two-leg binary, including opaque construction and checksum work;
no historical timings are pooled, and no public-workflow speedup is inferred.

Advancement requires semantic parity and at least 3% improvement in both
one- and two-attribute consume rows with ratio intervals wholly below 1. Any
protected consume regression greater than 5% with its interval wholly above 1
vetoes advancement. Every other regression remains a visible review trigger.
Passing can authorize only fresh public-workflow, resource and cross-format
trials; this packet cannot adopt a production optimization.

Root owns all builds, minimal five-copy helper tests and captures, serially.
The source archive and exact probe inputs are frozen before measurement.
Independent source review and offline replay supplement semantic differential,
clone, fused-exhaustion and comparison-bound tests. Failed attempts, if any,
are retained with their sources and logs. Cleanup must preserve executable
identities before removing only the owned temporary target.

## Final result

Rejected: both required short-tag benefits pass and empty consumption improves,
but four protected syntax cases regress 7.275–34.707%. All 19 consume and 39
construction regression flags remain retained. No workflow advancement or
production adoption follows. The final helper gates pass 170 tests and Clippy;
all 1,248 reports and 28,392 samples replay after owned-target cleanup.
`summary.md`, `decision.json` and the main 0802 report contain the full results.
