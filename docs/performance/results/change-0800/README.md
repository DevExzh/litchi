# 0800 — restore duplicate-error parity at the bounded handoff

The original no-replay performance experiment was deferred when a controlled
probe reproduced a correctness defect in the current baseline. This batch fixes
that defect in all five production helper copies. No candidate performance
measurement or production optimization is claimed.

quick-xml consumes the first non-whitespace key byte before looking for `=` or
XML whitespace. The fallback's error remapping scanned from that byte instead,
so a leading `=` became an empty key. After 32 successful attributes, a repeated
`=n0` key with a malformed value could therefore return `ExpectedQuote`,
`ExpectedValue`, or `UnquotedValue` instead of the contracted `Duplicated` error.
The correction includes the first byte in the name while excluding it from the
delimiter search. Exact duplicate positions, the first-error/fused contract, XML
whitespace, borrowed values, and the existing complexity bound remain intact.
This is lexical compatibility with quick-xml, not XML Name validation.

`correction/` and `correction.patch` retain the exact before/after production
sources. The initial exploratory edge probe retains its source, output and
command disclosure; timestamps were not recorded. A separate controlled locked
run compiles exact baseline and corrected helpers together and asserts the old
mismatch at accepted prefixes 32/33/34 and corrected parity for every row.

Root runs formatting, the controlled edge probe, full tests for the five affected
production crates, and all-target Clippy serially. The new shared test covers
540 duplicate-error cases and 270 lexical/nonduplicate controls per helper copy,
including handoff boundaries, unusual keys, XML whitespace, exact error positions,
clones, and fused exhaustion. A failed first test attempt is retained: its test
matrix incorrectly treated `==n0` as a complete key. The parser correction did
not change when the test case was corrected.

`deferred-candidate/` and `deferred-candidate.patch` are source-only, untested
archives based on the previous commit. They are not applied, measured, or
approved. Before any later experiment, rebase them on this corrected baseline,
reconcile their retained old fallback, and repeat correctness gates. The unused
performance harness draft was removed before any candidate build or capture.

The packet seals production changes, source and architecture witnesses, quality
receipts, failure records, reviews, and target cleanup. Unrelated files and
worktrees remain intact. No performance baseline, CRUD/producer, cold/range, or
concurrency coverage is promoted. iWork remains excluded.
