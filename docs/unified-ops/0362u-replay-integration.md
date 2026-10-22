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
under `.coord/replay-integration/`.

The first merged run exposed seven facade RTF test consumers that still
required opened-object identity. They now verify forward replay on an
independently opened matching base, inverse replay after reopening the output,
and refusal against changed content. All seven suites pass in resumed groups
78–86. This follows the coordinator's content-identity decision; the changed
base is still rejected. Group 87 then exposed XLSB refusal classifications
that require a source repair and typed consumer assertions. No test output or
capability golden was blessed to hide that failure.

The resumed-session handoff in `.coord/triage-2026-10-22.md` supersedes the
old full-run schedule: preserve completed and failed group receipts, integrate
the committed RTF, DOCX, durable-Ink and raw-policy slices, then run affected
validation on the combined source. The interrupted 225-group run is partial
evidence, not a complete test or acceptance gate.
