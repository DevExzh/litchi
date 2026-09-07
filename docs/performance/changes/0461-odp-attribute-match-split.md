# 0461 — reject ODP attribute-match split on the practical gate

0461 split the ODP attribute-cache match from value decoding. The namespace-first
predicate remains unchanged; matching cached or scanned attributes decode only
after both namespace and local-name checks succeed, while iterator advancement,
first-match behavior, malformed/duplicate reachability, lazy normalization and
error strings remain unchanged. The source review found no semantic blocker.

The frozen A1/B1/B2/A2 matrix retains 24 reports and 720 samples. Normal p50
candidate-minus-baseline deltas are -2.0578% / -2.2940% for tiny,
-2.0604% / -3.8754% for medium, and -2.1547% / -3.1281% for large in R1/R2.
The predeclared 3% medium/large gate therefore fails in both R1 gated shapes,
even though the independent p50 bootstrap intervals remain below zero. No
adverse >5% elapsed or RSS flags are present, and every allocation metric is
exactly unchanged. The result is partial timing improvement with no retained
speedup claim.

Supplementary large-input phase clocks show p50 changes of -3.0962% / -0.4057%
for transaction, -3.9421% / -4.2911% for snapshot opening, -4.1677% /
-3.0561% for commit, -0.1727% / +1.2645% for add, and -0.4772% / -0.4562%
for publication in R1/R2. These are diagnostic phase observations, separate
from the failed lifecycle gate. Whole-process counters report instructions
-3.2567%, cycles -3.2314%, branch misses -3.3587% and cache misses +0.2581%;
setup, warmups and checks are included, so they are not operation-only or
causal counts.

Authenticated release disassembly confirms the mechanism: the candidate has no
out-of-line `ElementAttrs::lookup` body, inlines the namespace/local-name checks
in `get`, and reduces the `get` stack frame from `0x148` to `0x128`. This is
mechanism evidence only and does not override the failed practical gate.

The candidate passes 372 ODP tests, warning-denied all-target Clippy and scoped
formatting; 387 harness tests pass with one ignored. Source restoration to
`05f432d48`, final Clippy/docs/formatting/boundaries, portable verification,
tamper rejection and owned temporary cleanup pass. The rejected experiment adds
no selector or corpus coverage; the registry remains 439 selectors / 36
defaults, 0460's retained optimization remains accepted, the full non-iWork
goal remains open, and iWork is excluded. See the [comparison
summary](../results/change-0461/summary.json), [phase summary](../results/change-0461/phase-summary.json),
[source review](../results/change-0461/source-review.md), and [assembly
receipts](../results/change-0461/candidate-assembly.json).
