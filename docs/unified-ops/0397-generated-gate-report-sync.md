# 0397 — Synchronize generated gate reports

The complete gate on `1bc8f22499d2691e85ebed37686c8381f0c985da`
finished with 49 passing rows and four failing rows: goldens (42), tooling
tests (48), benchmark matrix (51), and corpus dispositions (52). These
failures have two generated-report causes. Runtime tests, strict Clippy,
rustdoc, feature matrices, fuzz checks and the remaining gate rows passed
on that same unchanged source. The failed receipt remains external under
`(deleted-coord-artifact)`; this is not a green full gate.

The OfficeCLI capability report still recorded the previous Numbers registry
snapshot hash. Regenerating it changes that hash only. No registry, corpus
disposition or capability claim changes.

The benchmark matrix did not ingest the existing eighteen chart-text
instruction comparisons in `results/0375b/integrated-callgrind.json`.
Its new parser records the actual host and leaf counts, computes the ratios,
and retains their historical source `4aa57b3c19` and instruction-only scope.
The rows do not establish wall-time, unified `File`, native-application, or
current-trunk qualification. No evidence exemption or checker relaxation
was added.

Validation: both generators pass `--check`; the benchmark-matrix and
OfficeCLI capability test modules pass (58 tests). Focused parser controls
verify count orientation, the five-percent instruction budget, historical
scope, and rejection of invalid or duplicate routes. Repository guards,
ledger and crate boundaries pass. All generated outputs remain below 1 MiB.

The coordinator must rerun all 53 gate rows on the final integrated
candidate; the historical passing rows are not combined with this repair.
