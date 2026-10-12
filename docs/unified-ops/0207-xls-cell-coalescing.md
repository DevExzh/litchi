# 0207 — XLS cell transaction coalescing

Consecutive direct scalar cell sets, clears and removals in `File::edit` and
atomic `File::batch` now share a native XLS package edit. Distinct targets may
span worksheets. The facade retains command plans and attribution, while the
native owner retains all cell state, validation, dependency closure and
publication. No parallel document model or implicit runtime is introduced.

Repeated coordinates close the run and resolve again against the published
candidate. This preserves insert/remove/set sequences. Reads, ranges, queries,
worksheet renames, empty property maps and formula handling remain barriers.
Per-item `changed` flags distinguish scalar no-ops from mutations. Best-effort
batches retain their existing independent-item execution and rollback policy.

Deferred native failures are located by replaying bounded command prefixes
only on the failure path. A failed staging operation first closes any earlier
prefix so a preceding deferred error keeps its original item index. Source,
signature and encryption admission remain in force throughout the run.

## Bounds and native run boundaries

The native XLS owner has a hard limit of 4,096 staged changes. The first
10,000-operation measurement exposed that limit. A typed native operation
limit now closes a nonempty run, publishes it privately, and retries the same
command against the new candidate. A limit on an empty run remains an error.
Other native errors and caller limits are not converted into run boundaries.
This preserves support for larger public batches without silently widening
the native policy. The 1,000-operation scenario uses one native commit;
10,000 distinct numeric operations use three bounded native commits.

The public mutation budget counts calls, including no-ops. XLS previously
compared that bound with native cell/resource counts at publication, so a
text setter could consume multiple public operations. Publication now checks
the public command count and retains the native count separately in diagnostic
history. Staged-payload and retained-snapshot bounds remain independent.

## Verification

The new suite demonstrably fails on the preceding implementation's commit
counts: five distinct mutations produced five commits, and 1,000 produced
1,000. With coalescing they each produce one. A new 4,100-operation boundary
test passes with two commits for both sets and removals, through eager and
real filesystem owners. The original source and exact inverse are checked.

The all-feature facade suite passes 1,547 tests and doctests. Following the
final mutation-budget correction, all 49 focused XLS write/removal/batch tests
pass, including 13 batch tests. Coverage includes cross-sheet runs, repeated
coordinates, query/read/rename barriers, no-op item flags, atomic/best-effort
attribution, a deferred rich SST dependency refusal, source freshness,
signed/encrypted admission, budgets, exact inverse and replay.

Warning-denied all-target Clippy and rustdoc pass. Minimal and XLS-only
feature checks pass, as does the XLS glossary audit (24 banned names and one
transaction family). Scoped formatting passes. The crate-boundary checker passes
for 66 packages and 268 internal dependencies (11 existing debt items).
The generated OfficeCLI table remains current, and all 15 generator tests
pass. No capability claims or property registry entries changed. Repository
lockfiles remain unchanged; offline validation uses an external lockfile
selecting cached cc 1.6.0 and an external build target.

Reproduce with `cargo test -p litchi-ops --all-features` and the focused
`xls_unified_batch`, `xls_unified_write`, and `xls_unified_remove` integration
tests. No Python/CLI or full acceptance completion is claimed.

## Measurement method

`crates/litchi-ops/examples/xls_batch_bench.rs` pairs native package edits
with both unified closure and atomic batch APIs for 1, 10, 100, 1,000 and
10,000 distinct numeric sets and removals. Eager and filesystem owners use
the same input bytes. Above the native staging limit, all legs use equivalent
bounded transactions and assert equal commit counts. Generation, opening,
payload construction, serialization and correctness verification are outside
timing. Native whole, staging (including edit admission), and commit costs
are reported separately. Every output checks complete coordinates/values,
byte equality and source immutability.

## Paired results

All 20 scenarios in the [five-sample run](results/0207-xls-batch.json) and
[nine-sample repeat](results/0207-xls-batch-repeat.json) pass byte equality,
complete coordinate/value checks, source immutability and native/unified
commit-count equivalence. Repeat medians in milliseconds:

| Owner | Operation | Calls | Native | Closure | Atomic batch | Batch/native |
|---|---|---:|---:|---:|---:|---:|
| Eager | set | 1,000 | 5.766 | 6.482 | 6.626 | 1.1491× |
| Eager | set | 10,000 | 170.429 | 187.554 | 188.899 | 1.1084× |
| Eager | remove | 1,000 | 4.507 | 5.129 | 5.242 | 1.1631× |
| Eager | remove | 10,000 | 130.046 | 140.304 | 143.784 | 1.1056× |
| Filesystem | set | 1,000 | 5.582 | 9.956 | 10.037 | 1.7981× |
| Filesystem | set | 10,000 | 157.941 | 200.945 | 204.910 | 1.2974× |
| Filesystem | remove | 1,000 | 4.413 | 8.732 | 8.795 | 1.9931× |
| Filesystem | remove | 10,000 | 126.820 | 158.485 | 163.534 | 1.2895× |

The 1,000-operation atomic batch ratios are 1.15–1.16× for eager owners and
1.80–1.99× for filesystem owners in the repeat. The initial filesystem
removal ratio was 2.001×, just above the 2× review threshold. That small
margin is not stable qualification: source-path/freshness overhead remains
a follow-up. Smaller source removal batches also exceeded 2× in the initial
100-operation case; all samples remain in the evidence.

The 10,000-operation measurements do not establish linear scaling: native
eager sets grew from about 5.8 ms at 1,000 to 170 ms at 10,000, and the facade
followed that growth. Native inventory lookups and staging overlap searches
are known per-command linear scans, while the larger case also requires
three native publications. Indexing those paths while preserving source
order and duplicate refusal is the next performance work; these samples do
not isolate the cost of each cause. Single-cell closure/batch timings are
not a replacement for the one-shot write benchmark, and no text, repeated
coordinate or peak-allocation performance claim is made.

Source commit: `0ebc78408e4a3246095e374a902e80263da6f2e0`.
Build with `cargo build --release -p litchi-ops --features xls --example
xls_batch_bench`; run `target/release/examples/xls_batch_bench` with
`LITCHI_XLS_BENCH_SAMPLES=5` or `9`. Binary SHA-256:
`2160c91fc08c731fc58d63d088b0d5bf8dee62ca9087e0acb130b470935d35ed`.
No binary is committed.

## Remaining scope

The native inventory lookup and staged-change overlap/search paths still
contain linear scans per command. Their indexing, source-backed direct
lookup overhead, repeated-coordinate coalescing and broad batch scaling
qualification remain follow-ups. Worksheet and axis CRUD, formula/style
editing, security publication, source query streaming, the remaining formats,
Python/CLI and the complete goal remain open.

The base remains an ancestor at
`32ad64e3207655b4ba558fc2793d9a7396655cdf`; no additional rebase was needed.
