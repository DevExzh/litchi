# 0203 — Unified XLS reads

The XLS facade now implements `get`, `find`, compiled queries, node views and
text/annotated/outline/statistics views over the retained native `Package`.
Workbook, sheet, cell, range, row and column nodes share the workbook vocabulary
with XLSX/XLSB. Direct cell lookup uses the native selected-cell route; traversal
materializes a sparse worksheet grid scoped to the read operation. It does not
install a second persistent workbook model in the facade owner.

Eager and positional owners expose the complete tab catalog, preserving chart
tab positions, names and visibility independently of worksheet payload ordinals.
The native selected-cell and visitor projections retain formula classification,
including indexed replay. Views count formulas from that classification and
read the actual workbook date system. Source fences cover node projections,
query traversal, child iteration and view output. Dense range serialization
checks its requested cell budget before expansion.

Capabilities and the generated OfficeCLI delta now include XLS reads. XLS is
compared against the XLSX vocabulary as a parity target, without claiming that
OfficeCLI supports XLS. Formula token text, separate cached-value properties,
row/column layout, active-tab metadata, header predicates and extended workbook
elements remain explicit gaps. Chart/macro tab payloads are not worksheet view
content. Unified mutation operations remain pending.

## Verification

The combined XLS and XLS/encryption facade suites pass 2,257 tests and doctests,
with one ignored. The all-feature facade suite also passes 1,494 tests and
doctests. Final focused checks also exercise cached formula errors in
statistics/outline/annotated output and the BIFF8 A1:IV65536 boundary. New
contracts cover eager/positional parity, sparse values,
formula statistics, normalized names, bounds, dense quotas, signed reads,
encrypted unlocks, stale sources, real chart-tab catalogs, hidden sheets and
1904 dates. Warning-denied all-target Clippy and rustdoc pass. The minimal
facade feature build passes. The boundary check passes for 66 packages and 268
internal dependencies, with 11 existing debt items. The parity generator's 15
Python tests pass. The XLS glossary audit passes 24 banned names and one
transaction family.

Reproduce with `cargo test -p litchi-xls -p litchi-ops --no-default-features
--features litchi-ops/xls,litchi-ops/encryption`; Clippy uses those packages and
features plus `--all-targets -- -D warnings`. Offline builds use an external
lockfile selecting cached cc 1.6.0 and an external target; repository lockfiles
are unchanged.

## Read measurement

The `xls_read_bench` example pairs native selected-cell lookups with facade
`get`, and native cell iteration with compiled facade queries at 1, 100, 1,000
and 10,000 numeric cells. It checks every value, count and checksum before
measurement and checks repeat results. Generation, publication, opening, caller
path construction and query compilation are outside timing. This is an eager,
single-column read measurement; it does not qualify positional reads, opening,
memory use, mixed corpora, writes or Python/CLI.

Five release samples, median microseconds per complete scan:

| Cells | Native get | Unified get | Ratio | Native iteration | Unified query | Ratio |
|---:|---:|---:|---:|---:|---:|---:|
| 1 | 0.030 | 0.420 | 14.00× | 0.050 | 0.730 | 14.60× |
| 100 | 1.420 | 40.781 | 28.72× | 0.670 | 22.030 | 32.88× |
| 1,000 | 17.980 | 412.892 | 22.96× | 6.310 | 226.701 | 35.93× |
| 10,000 | 325.201 | 4281.329 | 13.17× | 74.860 | 2410.111 | 32.19× |

These results **miss** the facade overhead target. Values and checksums agree.
The reduced [measurement JSON](results/0203-xls-read.json) contains every sample.
Reproduce from the commit containing record 0203 with
`LITCHI_XLS_BENCH_SAMPLES=5 cargo run --release -p litchi-ops --features xls
--example xls_read_bench`. The measured binary SHA-256 is
`b53253bee2a6d162b6dfe07c3dbc863422616bc2b49f95fc3946e5317ce2bdb4`; no binary is committed.

The first projection does not satisfy the early-termination/axis-cost contract:
query traversal builds the full sparse grid before yielding its first cell,
even when a caller stops after one match. A retained native cursor with bounded
projection is required before read-performance acceptance. Tab traversal also
copies the complete native catalog per selected tab; eliminating that quadratic
catalog work is a separate read follow-up.

## Remaining scope

Unified XLS mutations, batching/history, Package security publication, full
format coverage, Python/CLI and the remaining performance/acceptance gates are
still required by the active goal.
