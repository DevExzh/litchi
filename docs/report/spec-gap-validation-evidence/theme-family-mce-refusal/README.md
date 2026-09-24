# Theme-family MCE mutation refusal

This follow-up closes an ambiguity in selective family mutations: an MCE-wrapped
extension list or admitted extension can be relevant even when it contains no
family yet. The scanner now records that ownership boundary immediately, and an
`AlternateContent` directly inside an admitted extension also blocks mutation.
Choice and Fallback are treated conservatively; this helper does not select or
rewrite effective MCE branches.

Shared add, replacement, and removal refuse those ambiguous sources. Removal
checks ambiguity before its absent-family no-op path. Read-only projection still
preserves the XML and does not infer a direct family through MCE ancestry.
Foreign elements and unrecognized extension URI subtrees remain opaque rather
than becoming typed ownership paths. Exact host transaction no-ops may retain
the original source without attempting a selective XML mutation.

Empty-container removal closure from `670af0b07` remains in force for unambiguous
sources. Earlier closure gate receipts and the original family profile are
historical evidence for their recorded revisions. This directory records fresh
compile-first validation; no new performance or native Office claim is made.

Validation passed: compile-first all-feature/all-target checks, 1,141 tests
including doctests (12 existing ignored tests), strict Clippy and rustdoc,
formatting, and diff checks. The regression matrix covers 13 MCE shapes,
including self-closing containers, plus foreign ancestry and URI preservation.
The independent review approved the production hashes in `review.json`;
`verification.json` binds the receipts to 870 source inputs.

Reproduce from the repository root:

```sh
python3 docs/report/spec-gap-validation-evidence/theme-family-mce-refusal/run_checks.py
python3 docs/report/spec-gap-validation-evidence/theme-family-mce-refusal/verify.py
```
