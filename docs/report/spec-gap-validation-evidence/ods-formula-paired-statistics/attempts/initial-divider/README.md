# Initial shared-divider capture

This is the first complete capture, retained because matched AVEDEV and DEVSQ
controls showed material latency regressions. It is excluded from the final
candidate comparison; its measurements have not been replaced or selectively
filtered. Both phases include 39 matched controls and 105 candidate cases,
with three warmups and 15 fresh child samples (4,320 rows total).

`source/` contains every original selected frozen file, including the isolated
lock; overlay it on the baseline commit in `gates/freeze.json` to reconstruct
that candidate. `gates/` preserves its seven successful integration checks.
`performance/` preserves the captured profile inputs, raw measurements,
reports, and cleanup receipts. The batch verifier independently checks this
capture against its archived source snapshot.

The subsequent source change replaces the per-bit shifted comparison in the
shared dyadic helper with a bounded limb-wise comparison. Exact quotient,
remainder, and direct subnormal rounding remain the intended semantics.
No timing-only retry is used to clear the original flags.
