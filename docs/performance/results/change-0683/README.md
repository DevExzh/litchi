# Change 0683 evidence packet

This packet evaluates compact retained XLSX selection records, exact
shared-string index staging, and a profiled SpreadsheetML escape-search
optimization. Baseline: `0a292428d598b5235ca0075f2b1a5d5b931af31d`.
The selected scanner finishes XML, archive and dependency validation before
`visit_cells` invokes a callback. No streaming-publication contract is added.

## Reproduction and bindings

[Probe instructions](probe/README.md) specify the two checkouts, deterministic
corpus generator, build commands, measurement boundaries and allocator gauges.
`baseline.json` binds the goal and accepted architecture constraints;
`environment.json` records the host/toolchain and the shared-host caveat.
`corpus-manifest.json` binds the nine generated inputs; measurement manifests
also bind the real POI fallback control. `measurements/{baseline,candidate}`
retain exact commands, source/probe/binary hashes and raw-output hashes.
`audit.py` verifies source, probe, constraint and raw bindings, identical
three-repeat allocation counts, complete timing legs and matching semantics.
`comparison.json` retains every group's p50/mean/p95/p99, all control drift,
allocation gauges and regression flags. Twenty samples do not establish tail
latency guarantees. Cold means reopened owners with ordinary OS file caching.

`final-verified/` contains the final owner/facade quality checks and per-command
source hashes. `run-integration.py` reproduces them. `evidence/` contains seven
passing claim, coverage and dependency gates run by the retained change-0675
runner; these validate their named manifests, not full-goal completion.
`final-evidence/` reruns structural claims, report classification, CRUD coverage
and the non-iWork manifest after final documentation edits. Strict claims and
claim-gate tests are reused because their registry and implementation did not
change; dependency boundaries are reused because crate edges did not change.
`review.md` records independent source review. `run-diagnostics.py` reproduces
the separate native 1,000-sample long-inline ABBA diagnostic; `diagnostics/`
binds its binary/corpus hashes and raw outputs. Hardware counts include process
setup and three warmups, and are not allocator-instrumented measurements.

## Build-order clarification

The frozen, hash-bound `probe/README.md` contains inconsistent ordering prose:
its shell example builds both binaries first, while the following paragraph
says baseline A/A precedes the candidate build. The reproduction requirement is
that builds and measurements do not overlap, and baseline measurements precede
candidate measurements. For the initial capture both binaries were built first.
For the final expanded capture the existing baseline binary was reused,
baseline measurement finished, then `final-build.log` records the final
candidate rebuild, followed by candidate measurement. Either serial build
ordering in the example is valid. Captured probe bytes remain unchanged so the
measurement manifests can continue to verify them exactly.

## Initial candidate and retained failures

`initial-candidate/` freezes the first candidate's probe, raw measurements,
comparison, passing quality checks, production diff and added integration test.
Its long-inline owning-query regression led to the decoder extension; its
profiles and native ABBA reports remain in `diagnostics/`. Paths in archived
command manifests describe the original capture locations, before moving files
under `initial-candidate/`. The audit verifies the archived raw/probe hashes
against their new containing directory. Raw `perf.data` files are temporary;
text symbol reports and command records are retained.

Top-level `integration/` retains initial test-only lint failures;
`final-integration/` retains the later fixture namespace failure, corrected in
`focused-tests.log` and final quality results. `before-build-initial.log`
retains the early probe build failure. `before-build.log`, `after-build.log`
and `final-build.log` distinguish baseline, initial candidate and final builds.
The final combined patch, not record compaction alone, owns final timing.

The [change record](../../0683-xlsx-selected-record-compaction.md) reports
results and limitations. `performance_claim: none`; no registry or CRUD
coverage promotion. OLE2/OOXML work remains active; iWork is excluded.
