# Next OLE2/OOXML priority

The next candidate should keep the shared successful worksheet traversal but
forward materialization errors after the validation observer accepts EOF. The
measured 0542 late-raw path repeats both validator and parser before the existing
x14ac error scan; three of four rows exceed the unchanged 2x valid-baseline
latency envelope. `next-candidate.patch`, relative to `applied-candidate.patch`,
is an unmeasured proposal to remove that repeat while keeping x14ac retry error
precedence. It also drops retained validator state before every authoritative
fallback. Early parser/reader/cap failures still require full validation first.

Apply the full five-file candidate and follow-up only under a fresh source-bound
campaign. Keep the 0541 first-error matrix, three cap tests, unchanged native and
allocation gates, invalid late-error guards, and explicit provisional-state
bounds. Do not relax the 2x threshold or reuse 0542 timing as proof of the revised
implementation. Require instruction attribution and eager read controls after
the complete pilot passes, then final quality before any runtime retention.

This addresses a measured OOXML bottleneck. OLE2 and OOXML remain ahead of ODF;
ODF optimization stays deferred until that goal completes. The broad program
goal remains active.

Late-validator p50 also rose 164.7–168.8% against the same-invalid baseline,
although it stayed inside the frozen valid-baseline envelope. Forwarding
post-EOF raw errors does not remove parsing performed before a late validation
fault. Retain and review that latency/peak tradeoff in the next campaign rather
than describing the follow-up as a fix for all invalid-input costs.

Before freezing that proposal, add differential coverage specifically for
post-EOF `Complete(Err)` outcomes, as requested in `next-candidate-review.md`.
Also inspect whether `complete_source_parse` can avoid its redundant x14ac
marker scan on successful results; eligibility already proved the source plain.
Any revised draft still needs the full fresh admission campaign.
