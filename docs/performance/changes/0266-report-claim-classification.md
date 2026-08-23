# Change 0266: fail-closed historical REPORT claim classification

## Status

Landed in `5d9120a1116ae1a6993d51469d219e1b27df3887`. This change gives the
retained historical tables in `docs/performance/REPORT.md` an explicit,
machine-checked disposition. It does not add a performance result or change
the library or harness.

## Scope

The report's two audited tables are explicitly titled historical and
descriptive, not current claims. The surrounding report prose and individual
change sections remain outside the sidecar's table scope. The new
`docs/performance/report-claim-classification-v1.json` sidecar binds the exact
report path, section heading, preamble, header, row count, ordinal, label, and
row digest for both tables. Its allowed states are `strict_claim`,
`historical`, `descriptive`, and `withheld`.

The sidecar currently covers 167 rows: 88 historical stable-tranche rows and
79 historical accepted-result rows. Their dispositions are 145 `historical`,
14 `descriptive`, 8 `withheld`, and 0 `strict_claim`; no historical row is
linked into the strict claim registry.

## Fail-closed gate

`tools/check_report_claim_classification.py` parses only the two audited
Markdown tables and rejects changed or reordered rows, header/preamble drift,
unknown states, malformed or duplicate-key/non-finite JSON, path escapes, and
symlink rebinding. It also validates the canonical strict claim registry and
requires any future strict-row links to be exact. The CI workflow runs this
checker and its focused test module.

This is report-integrity and claim-boundary evidence only. It makes no latency,
allocation, RSS, I/O, production, or performance-improvement claim.
