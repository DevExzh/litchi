# 0362rna — Preserve forward-only redaction during artifact replay

The generic source-backed restore path matched a foreign recipient's complete
artifact before asking the family to publish. A retained redaction inverse
could therefore bypass the family's fixed `UnsupportedOperation` refusal.

The internal `Change` contract now declares whether foreign restoration is
allowed. Main/story redaction disable it for non-noop inverses. The retained
owner fast path remains exact and non-materializing; independently opened
equal redacted output reaches the existing typed refusal. Both families gain
regressions for the retained owner and equal foreign recipient.

Ordinary paragraph and tail patches retain content-matching inverse replay
under record 0362m. Their retained before/after pair supports exact restoration
on an equal reopened artifact; detached record fallbacks without that pair
keep their existing limitations. Comments clarify this distinction and a tail
regression checks exact reopened inverse output.

Source commit: `7753b2c0a2`. Formatting and artifact guards passed. Combined
compiler, semantic, corpus, and performance qualification remain pending.
This repair applies existing redaction rules; the comprehensive item-22 audit
remains the sole deferred work.

The assembled compiler exposed the detachment family's older custom restore
branch. Detachment now uses the shared source-freshness and recipient-policy
path, including its tighter-hyperlink-budget regression. A retained owner
patch authenticates a foreign complete base once and rejects mismatches before
scoped publication. Detached record patches without a retained owner pair
keep their family-scoped validation. A story-text regression changes only an
unrelated archive part to prove this distinction. Foreign redaction inverses
still return their fixed unsupported-operation refusal before hashing; this
preserves the existing wrong-base and equal-base refusal classifications.
