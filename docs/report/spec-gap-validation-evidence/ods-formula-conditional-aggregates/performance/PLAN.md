# ODS conditional aggregate evaluator performance plan

The profile follows ADR 0005's measurement contract and the repository's
five-percent matched-control review threshold. It keeps production source,
test/oracle inputs, harness source, lockfiles, toolchain identity, raw child
receipts, source manifests, and cleanup receipts independently reproducible.

The baseline revision is
`5b125e9ea5870a56b2648ad02917894a5549c571`. Its lane contains only functions
that are valid in both revisions: arithmetic, `SIN`, `IMSUM`, `DSUM`, and
`SUM`, with scalar, literal-array, and borrowing-reference controls. The
conditional family has no valid baseline implementation, so its rows are
candidate-only evidence rather than refusal timings.

The candidate matrix has six ordinary function paths, text criteria, ordered
reference-list and 3-D geometry rows, an anchored destination with explicit
SUMIF clipping, three projected `SUMIFS` rows at 64, 256, and 1024 rows, and
five formula-error rows, plus one reference-cell resource refusal.
Reference ranges use four-cell lanes: `A:D` for the first numeric criterion,
`E:H` for the second, `I:L` for selected numeric values, and `M:P` for exact
text criteria. Criteria and destination reads are counted by the resolver. The
projected rows retain the expected selected-value read shape so a per-output
recomputation would be visible in both work and provider-read metrics.

Each sample runs in a fresh release child with three warmups and one measured
iteration batch. The `evaluate` phase reuses a parsed expression, while
`parse-evaluate` parses inside the timed batch. Fixed repeat counts make short
control calls measurable while large reference matrices run once per sample.
Allocator requested/released bytes and live bytes must balance. Every case
performs an untimed finite-result or formula-error preflight before timing.

Builds and control preflights may run while implementation work continues. The
root agent must freeze the selected source closure and authored profile inputs
before either final lane; no profile input or compiled source may change during
capture. A matched time, RSS, allocation, or work movement beyond five percent
receives scoped review before any regression claim.
