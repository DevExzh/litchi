# Next OOXML work: avoid reconstructing source worksheet layout

The current measured rewrite-to-`scan_with_limit` edge consumes 53.24–54.94%
of exact `MultiSourceEdit::commit` instructions across the four shapes and
two repeats. This edge is larger than XML validation (30.40–35.11%) or
provenance merge (4.64–6.03%). These are nested instruction shares, not elapsed
time shares or removable fractions. Full numeric custody is in
`profile-analysis.json`; `source-review.md` identifies the different proofs
the source parser, edit scanner, and output validator supply.

The next concrete design task is a private source-bound layout proof produced
during the existing eligible source traversal. Determine whether that traversal
can retain the source offsets and edit-eligibility facts needed by the value-only
writer, allowing commit to avoid a second complete source layout scan. A semantic
Store alone cannot authorize this: it lacks lexical offsets, opaque spans, and
several structural facts consumed by the writer.

Before implementing a candidate, enumerate every `Layout` field and scanner
refusal, identify its producing event/state, and prove the proposed handoff
against the unchanged complete scanner. Original input positions must remain
valid across namespace handling and MCE processing. Unsupported or uncertain
input, new rows, shared formulas, identity/version mismatch, and incomplete
proof must retain authoritative fallback with exact error precedence. The
complete emitted-output validator and independent changed-cell readback remain
required. The rejected 0527 row arena is not a retained enabler.

The cost must be assessed across planning plus commit and publication. A
handoff that merely moves the scan into planning is insufficient. Freeze a
fresh matched campaign before admission, with one-cell/one-percent, managed,
noncompact, vendor, exact-no-op and refusal cases; capture planning and commit
allocation overlap and peak bytes as well as native workflow results. Bound
any provisional metadata before allocation, preserve cancellation/execution
checks, and test source sharing and byte-exact output. Inspect emitted code
when selecting the implementation so source-level duplication is not mistaken
for executed redundant work.

If the facts cannot be produced without equivalent extra traversal, excessive
retention, or weaker diagnostics, close the handoff proposal and use the
measured XML-validation branch as the next owner. In particular, visible
allocator/name operations there do not by themselves prove an allocation or
latency opportunity. No production optimization is admitted by this diagnostic.
OLE2/OOXML remain active; ODF stays deferred until their goal is complete.
