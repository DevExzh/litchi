# 0362u — Integrate content replay and primary Commit APIs

This candidate merges replay `1c7eb1e84c` and its fourteen incoming commits
onto the validated metadata assembly `5532099873`. It contains the File
content-proof work in 0362m, primary PPTX/DOCX/RTF Commit migration in 0362p,
and the coordinator repairs in 0362pa. It is not trunk acceptance.

Two harness files conflicted. The resolutions preserve the assembled XLSB
trait import and the already-explicit detector limits, while retaining every
incoming Commit consumer migration. The incoming history was checked against
the repository's pre-rewrite commit map. The ignored workspace lock changes
only by adding the ops SHA-256 dependency; tracked tool lock migrations remain
with their consumers.

The return checker from trunk `0dfb00b590` is included in this candidate.
Both DOCX profiles register the new metadata and source replay results;
PPTX minimal/full profiles now verify all 48/49 owner replay returns. RTF
registers both eager and source-backed tail append. The five profiles pass
against fresh feature-specific rustdoc captures with warnings denied.
The current inventory has 33 DOCX rows and three RTF owner-return rows;
PPTX has none. The remaining RTF primary/durable/redaction replay migrations
and DOCX raw/model work are separate required slices, not exceptions.

Source validation currently passes the regression in 0362pa, all 38 return
checker tests, nine-package all-target/all-feature Clippy and formatting.
Complete merged crate tests found a stale ODF insertion-error assertion after
the range-replacement extension; it now checks the exact fixed range error.
They also exposed an unrelated classification regression introduced by W4
commit `e86b97e61e`: DOCX document Durable, Composition and InvalidDurable
failures changed from invalid_input to invalid_argument. This integration
restores invalid_input and preserves the existing consumer assertion. Patch
conflicts continue to report invalid_input. The repaired ops library suite
passes all 969 tests.

The remaining merged crate tests are running in bounded groups. Their runner
removes only completed executable artifacts and checks target headroom before
each command; protected release/measurement artifacts are retained. Final
test receipts, allocation/native qualification and the final workspace gate
remain required. The complete command logs are retained outside the checkout
under `(deleted-coord-artifact)`.

The first merged run exposed seven facade RTF test consumers that still
required opened-object identity. They now verify forward replay on an
independently opened matching base, inverse replay after reopening the output,
and refusal against changed content. All seven suites pass in resumed groups
78–86. This follows the coordinator's content-identity decision; the changed
base is still rejected. Group 87 then exposed XLSB refusal classifications
that require a source repair and typed consumer assertions. No test output or
capability golden was blessed to hide that failure.

The resumed-session handoff in `(deleted-coord-artifact)` supersedes the
old full-run schedule: preserve completed and failed group receipts, integrate
the committed RTF, DOCX, durable-Ink and raw-policy slices, then run affected
validation on the combined source. The interrupted 225-group run is partial
evidence, not a complete test or acceptance gate.

The trunk merge brings in `a18593385c`, including 0362qa's shared generic
return checker. Its four Python implementation/test files match trunk
exactly; the assembled glossary retains the newer DOCX/PPTX/RTF registrations
and all W4/File profiles. Curated matrix exclusions are the union of both
parents, without overwriting a conflicting disposition. All 40 return-checker
tests and `unified_bench_matrix.py --check` pass after this resolution.

RTF integration through `5909588b99` includes the focused DOCX edit-surface
cleanup from 0362r. The tail-patch conflict retains both the new borrowed span
and the coordinator's direction-aware diagnostics: removal replay inserts
zero bytes. All 63 logical-tail, primary, durable-composition and redaction
tests pass, including forward/inverse/no-op diagnostic assertions. All five
RTF result rules now register in the assembled policy and pass against the
worker's fresh unchanged public-surface capture. The command receipts are in
`results/0362u/rtf-integration.json`; full combined checks remain required.

DOCX paragraph-copy integration through `9f1b3eb45f` applies cleanly after
0362r. It retains the required raw codec limits and all migrated consumers,
and replaces the private recursive `TailChange` enum with one direction bit.
That removes the inverse Box allocation and the unfiltered Clippy failure
recorded by 0362ua. The worker's final source passes 2,270 all-target tests,
87 focused tail/index tests and strict all-feature Clippy; combined validation
is still required and is not inferred from that standalone result.

Durable-Ink integration through `f97f0eef7a` reconciles the two Commit
implementations. It keeps the source-owner sharing from `76148455d5`, retained
Ink policy, early source admission and source lifetime checks, and all existing
no-op patch/diagnostic assertions. Both owner replay methods now register their
`ink::Commit` result in the glossary. Fresh minimal/full DOCX captures pass
42/43 result contracts with rustdoc warnings denied. The combined source passes
144 Ink, tail-append and paragraph-copy tests, plus unfiltered all-target,
all-feature Clippy for DOCX, ops and XLSB. This also closes the combined-source
TailChange Clippy failure retained in 0362ua. Receipts and external capture
hashes are in `results/0362u/ink-integration.json`. The six-route matched
allocation repair remains recorded in 0362lc; native timing and full assembly
acceptance remain required.

ZIP integration through `a4d69035be` preserves the historical typed prefill
refusal without widening a rejected read. The spill threshold retains the
64 KiB recommendation for small declared directories; valid larger records
still use bounded spill storage. On the combined source, the full
all-feature ZIP suite passes 657 tests with 2 existing ignored tests,
including ZIP64 promotion and rustdoc; strict all-target/all-feature Clippy
and the benchmark evidence matrix also pass. The exact receipts are in
`results/0362u/zip-integration.json`. This closes that reproduced old-gate
failure, while the final combined workspace gate remains required.
