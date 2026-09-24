# ODS OpenFormula function-name recognition

This batch follows `21ad91183` and addresses the audit's incomplete standard
OpenFormula function catalog. The prior catalog contained 161 names. The
normative source is the bundled ODF 1.4 Part 4 specification, chapter 6.

Recognition is inert: accepting a function name produces a formula token and
does not evaluate arguments, access a source, refresh a cache, or execute a
function. The broader expression grammar, argument/arity validation, dynamic
arrays, host-defined functions, and evaluation remain separate capabilities.
All 161 previously recognized names are part of the expanded standard catalog;
there are no additional compatibility-only catalog entries. The existing
lookup helper's Unicode-uppercase behavior is preserved through bounded
normalization, while the tokenizer retains its ASCII identifier grammar.

The lexer must distinguish a known function invocation from an A1-shaped
identifier. Digit-bearing function names followed by an opening parenthesis
are function calls; ordinary bare cell references retain their existing
interpretation. The speculative ASCII identifier scan stops after exceeding the
longest canonical name (19 bytes), avoiding a second full scan of long unknown
identifiers. UTF-8 string values and doubled quotes are decoded without corrupting their
characters; unterminated literals are refused. Original formula text remains
retained. Public
`is_valid_function` remains a recognition query, not an expression validator.

The performance comparison separately measures recognition queries and parsing
already supported formulas. Newly recognized names have no successful parser
baseline and are reported as additional coverage rather than a speedup.
Lookup allocation counts, elapsed distributions, and process RSS are distinct
metrics. No spreadsheet evaluation or publication occurs in the harness.

This batch contributes to the audit and to the CRUD checklist's structural
query, inert formula authoring/retention, and length-bounded function-name lookup scenarios. It does
not certify the complete OpenFormula language or close the broader goal.

## Validation

The final source passed **783 tests across 46 targets**, warning-denied Clippy,
warning-denied rustdoc and doctests (zero doctests), and formatting. Commands and
exit codes are retained in [gate-results.json](gates/gate-results.json); the
[before](gates/source-before.json) and [after](gates/source-after.json) source
manifests are identical. The full gate ran in the shared workspace; isolated
candidate checks are recorded with the performance evidence.

The [independent catalog verification](gates/catalog-verification.json) confirms
393 standard names, including all 161 previous names and **232 additional names**.
Both the production catalog and the independent test list match the specification
exactly. The retained [extractor](extractor.py) reproduces the
[normative table](normative-functions.tsv) byte for byte. A second extraction
using lxml independently confirmed the heading set and section anchors.
See the [specification review](spec-review.md) and
[lexical/resource review](lexical-review.md) for scope and boundary details.

## Performance evidence

The [performance report](performance/report.md) retains 23 comparable lanes per
variant and 12 candidate-only coverage lanes, source and binary manifests,
independent patch replay, allocator observations, elapsed distributions, process
RSS, and a scoped hardware-counter run. All measured candidate lookup calls
allocate zero heap bytes. Common SUM/VLOOKUP parser batches improve 13–24%; the
cell-heavy control is approximately unchanged. The report explicitly reviews
the one process-RSS trigger (+128 KiB) and limits these claims to the instrumented
microbenchmark. The isolated candidate passes another 30 formula unit tests and
7 integration tests.

Run [verify.py](verify.py) from the repository root to verify the retained raw
measurements and final source provenance. This adds no production dependency.

The [root verification receipt](root-verification.json) records catalog checks,
final hashes, isolated binary verification, patch replay, and owned scratch
cleanup. [raw.sha256](raw.sha256) covers all retained raw measurement and gate
receipts; verify it from this evidence directory with `sha256sum -c raw.sha256`.
