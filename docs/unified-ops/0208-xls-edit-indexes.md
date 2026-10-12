# 0208 — XLS edit inventory and staging indexes

The native XLS edit inventory now retains immutable coordinate indexes beside
its source-order entries. Cell lookup, duplicate detection, mutation admission,
semantic replay and candidate readback use the same indexes. Public cell
iteration and physical source slots keep their original order. Fixed numeric
candidate reuse shares the coordinate indexes because those edits cannot move
cells; structural candidates rebuild them from the admitted output.

Fixed and structural staging keep sorted sidecar indexes. They accelerate
coordinate and overlap searches while retaining the canonical operation
vectors, deterministic patch order and the 4,096-operation bound. Staging and
join reserve index capacity before changing logical transaction state. Joins
rebuild sidecars after sorting. Joining an effective change into an existing
no-op now reuses that coordinate's slot, so subsequent staging cannot leave a
stale duplicate operation. Joins also reject fixed/structural overlaps in both
directions.

Canonical `litchi_xls::Edit::source_cell` exposes an original-source lookup from
the edit's already admitted inventory. It deliberately reports the pre-edit
cell after staging and retains source-freshness checks around lookup, including
failures. Unified coalesced commands resolve their paths and use this native
lookup for presence, formula classification and scalar values. This avoids
repeated source payload decoding without adding a second facade cell model.

The indexes use fallible vector reservation. Lookup is logarithmic, but staging
an out-of-order coordinate can shift a vector tail; this is not a general
logarithmic-insertion or linear-scaling claim. The paired benchmark now accepts
`LITCHI_XLS_BENCH_REVERSE=1` to measure descending coordinate order as well as
its default ascending order. The index memory cost is retained with the native
snapshot; peak-memory qualification remains open.

## Verification

The full native suite passes 1,791 tests and doctests, including the public
source-cell example and eager/source-backed lookup tests. Lookups perform no
additional source payload reads, report original values after staging and
preserve stale-source precedence even for missing worksheets. The all-feature
facade suite passes 1,548 tests and doctests, including coalesced XLS barriers,
limits, error attribution, replay and exact inverses.

The native regressions cover duplicate refusal, source-order iteration,
no-op joins followed by staging, fixed/structural conflicts in both join
directions, and two-sheet out-of-order edits followed by join, clone and
further staging. Fixed and structural results and exact inverse bytes are
checked. All six focused index tests pass after the final two-sheet regression
was added.

Warning-denied all-target Clippy and rustdoc pass for XLS and the facade.
Minimal and XLS-only feature checks and the XLS glossary audit pass (24 banned
names and one transaction family). The crate-boundary audit passes for 66
packages and 268 internal dependencies with 11 existing debt items. The
OfficeCLI generated table remains current and its 15 generator tests pass.
Both repository lockfiles are unchanged. Offline checks use an external
lockfile selecting cached cc 1.6.0 and an external build target.

Reproduce with `cargo test -p litchi-xls --all-features`,
`cargo test -p litchi-ops --all-features`, and the focused native
`cell_values::index_tests` module. No Python/CLI or complete acceptance claim
is made.

## Paired measurements

The unchanged ascending workload was captured on the preceding source and
again after indexing. All 20 cases in each five-sample
[baseline](results/0208-xls-batch-before.json),
[indexed ascending run](results/0208-xls-batch-after.json), and
[indexed descending run](results/0208-xls-batch-reverse.json) pass complete
coordinate/value checks, exact paired output bytes, source immutability and
native/unified commit-count equivalence. Generation, opening, payload
construction, serialization and verification are outside timing. The native
4,096-operation bound remains in force: 1,000 operations use one commit and
10,000 use three. Medians in milliseconds:

| Owner | Operation | Calls | Batch before | Batch after | Native after | After batch/native |
|---|---|---:|---:|---:|---:|---:|
| Eager | set | 1,000 | 6.637 | 5.823 | 4.975 | 1.1706× |
| Eager | set | 10,000 | 192.758 | 112.320 | 101.029 | 1.1118× |
| Eager | remove | 1,000 | 5.232 | 4.771 | 4.059 | 1.1753× |
| Eager | remove | 10,000 | 142.074 | 96.007 | 87.016 | 1.1033× |
| Filesystem | set | 1,000 | 9.857 | 8.267 | 4.740 | 1.7441× |
| Filesystem | set | 10,000 | 197.216 | 110.905 | 86.126 | 1.2877× |
| Filesystem | remove | 1,000 | 8.792 | 7.420 | 4.022 | 1.8449× |
| Filesystem | remove | 10,000 | 159.265 | 103.129 | 82.490 | 1.2502× |

At 10,000 operations, ascending atomic batches improve by 32–44% in this
run. Native staging, including edit admission, falls from 35–50 ms to about
3 ms. Descending staging takes 5.7–5.9 ms at 10,000, exposing the remaining
vector-shift cost; its full atomic batches take 99–118 ms. There is no
pre-index descending sample, so no descending speedup claim is made.

Publication still dominates: native 10,000-operation commit time is roughly
79–98 ms. Whole native edits grow about 18–21× from 1,000 to 10,000 operations,
which still does not meet general linear scaling. All measured batch/native
ratios in the ascending run are below 2×, but filesystem 1,000-operation
batches remain about 1.74–1.84× native. These five-sample synthetic numeric
workloads do not establish broad performance or peak-memory acceptance.
Resource/text workloads, repeated targets and arbitrary insertion orders
remain unqualified.

Baseline source: `ef16b597055e65205da3fb9903a34fb60ccb97ab`;
binary SHA-256:
`2160c91fc08c731fc58d63d088b0d5bf8dee62ca9087e0acb130b470935d35ed`.
Indexed source: `0f0eb7bd2b0e45468f3849083236e49b3f70d34b`;
binary SHA-256:
`8c759476a3d9c2c6b45758755c05543228b6be8bb13c5e530675e4f6d696cb3b`.
Build using `cargo build --release -p litchi-ops --features xls --example
xls_batch_bench`; run `target/release/examples/xls_batch_bench` with
`LITCHI_XLS_BENCH_SAMPLES=5`, adding `LITCHI_XLS_BENCH_REVERSE=1` for descending
order. Each JSON retains its exact source, command and binary checksum.
No binary is committed.

## Remaining scope

Source-backed freshness cost, out-of-order insertion cost, resource/formula
staging, repeated-coordinate coalescing, source query streaming and broad
performance qualification remain follow-ups. Worksheet and axis CRUD,
formula/style editing, security publication, other formats, Python/CLI and
the complete goal remain open.

The base remains an ancestor at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; no additional rebase was needed.
