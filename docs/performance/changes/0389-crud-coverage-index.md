# Change 0389: non-iWork CRUD coverage index

Date: 2026-09-03

Status: implemented as a contract-only index; full-run timing validation pending

`performance_claim: none`

`claim_authorized: false`

The Phase-1 CRUD matrix now has a machine-readable, non-iWork coverage index at
`docs/performance/crud-coverage-index-v1.json`. Its fifteen categories point to
exact selectors from the current Rust registry, checked schema-2 corpus IDs
where the default catalog applies, and explicit generated-per-run shapes for
opt-in cases. Every row carries one of the bounded status values
`measured`, `correctness-only`, `unsupported`, or `not-applicable`.

The mapping is intentionally representative, not an exhaustive certification
of every checklist row. A category status must match all of its scenario rows:
`measured` is a default-baseline timing contract: its selector/corpus identity
is statically bound, but the identity artifact is not timing evidence. A
scheduled/manual full run must validate at least 15 samples per measured row
from `target/perf/container-baseline.json`. `correctness-only` retains runnable
selector/API evidence but makes no retained baseline timing claim. Unsupported
and not-applicable rows carry an explicit reason. Documentation paths in the
index are navigation only; their existence does not substantiate behavior.

`tools/validate_crud_coverage_index.py` is a standard-library-only gate. It
rejects duplicate JSON keys, stale or invented selectors, catalog identity
drift, malformed corpus references, missing checklist references, unapproved
documentation navigation, invalid status/timing-contract bindings, and any
iWork selector or corpus named by this non-iWork index. The workflow runs it in smoke
and release jobs and includes the index in the scheduled/manual baseline
artifact. Dynamic formula evaluation and refresh remain explicitly unsupported
because the security contract keeps execution capability-bound and inert.
