# 0361ia — Pages footnote policy repair

This slice returns the Pages footnote policy migration for one aggregate-budget
correction. `FootnoteStorageProjection::decode` now derives the residual
aggregate `WireLimits` first, then clamps each storage `RewriteLimits` ceiling
to that residual. Graph usage is therefore charged once against the admitted
package policy; a tighter storage ceiling remains effective without
double-subtracting the same input, field, nesting or rewrite-work counters.

The body-footnote rewrite path also preflights the exact encoded output length
before reserving or publishing bytes. Parsing uses the current retained input
length, and replacement output is checked at its exact boundary; the focused
one-under case remains refused with the typed wire-byte limit error. Unknown
wire fields and identifier/source behavior are unchanged.

Source: `b389a5bd11`, `d2daf400d7` and `3600111b59`. The final Pages/IWA
library suites and Clippy pass at `3600111b59`; the reduced receipt is
`results/0361i-footnote-validation-receipt.json`. The workspace checks remain
an integration responsibility; no native timing is claimed here.
