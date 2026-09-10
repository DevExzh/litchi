# 0498 evidence layout

`plan.md` records the implementation and measurement contract before production
edits. `adr-refresh.json` verifies the unchanged accepted ADR set against 0497.
`environment.json` records the toolchain, host, affinity, and shared-host limits.

`baseline-freeze.json` identifies the retained before executable, its exact
harness source (`baseline-harness.rs`), and the first corrected unmanaged and
managed serial controls. The before production revision is `56e8ddbca`.
The historical CSV label `managed-after-only` means an explicit execution
context; that capability already existed before this change.

The initial `baseline-owned-few-large.csv` includes checksum work in its timer
and is exploratory only. Use the `*-corrected.csv` controls. In those files,
`record=sample` includes both warmups and measurements: indices 0–2 are warmups
and indices 3–32 are measured, separately for each repeat. Their summary rows
exclude warmups. Checksum verification is outside the corrected timer. Serial
controls retain every selected payload handle until after observation.

`gate-commands.json` and `run-gate.py` record reproducible scoped checks. Run one
with `python3 -B docs/performance/results/change-0498/run-gate.py NAME` from the
repository. Each invocation records its command, exit status, elapsed time,
output, and output hash under `checks/`. The script deliberately pins the
repository's 0498 toolchain and owned target; it does not isolate the host.

`protected-work.json` identifies unrelated user work observed during this
batch. These files are outside the change's staging and cleanup scope.

The final candidate is identified by `candidate-source.json` and
`after-freeze.json`. `final/` contains the thirty matched child captures and
receipts; `final-verification.json` checks all 1,800 measurements, 180 warmups,
byte signatures, source/work counters, released budgets, and raw hashes.
`analyze-final.py` produces `final-analysis.json` and `final-analysis.md`;
`verify-analysis.py` independently recomputes all 72 comparison scopes.
`profiles/` and `profile-summary.md` contain the six whole-child CPU captures.
`final-gates.json` identifies the passed checks and retained initial failures.

Reproduce the final matrix with `capture-final.py` after rebuilding and freezing
the executable. The capture script refuses to overwrite existing child outputs;
use a fresh evidence directory or retain the prior capture elsewhere first.
`profile-final.py` supplies the separate whole-child counter supplement.

`cleanup.json` records removal of 1,367,396,352 allocated bytes from dedicated
0498 scratch. Only the two hashed replay executables remain in the cache's
`retained/` directory. Unrelated work and the shared repository target were
outside cleanup scope. `inventory.json` records the final evidence-file hashes.
