# ODS adaptive logical-cell locator

Date: 2026-08-11

Production base: `16014fda8e542e39649ddabc650bca07a70743e3`

Scope: repeated public `litchi_ods::Spreadsheet::cell` queries. OLE2, OOXML,
RTF and the other ODF family production crates are unchanged. iWork/IWA was
explicitly excluded while its crates are changing independently.

## Profile gate and benchmark

The existing large ODS full-cell benchmark made 32,768 public facade calls,
but also cloned every cell string and joined the aggregate. A new opt-in
`ods_semantic_cell_sweep` case opens a fresh `Spreadsheet` outside timing,
then performs the same row-major public lookups without cloning text. It
counts stored cells and black-boxes every borrowed cell; complete grid and
text verification still runs outside timing. The locator is not prewarmed.

The frozen production baseline attributed 7.47% of whole-process exclusive
cycles to `Spreadsheet::cell`, despite including package open plus two
out-of-timer verification scans. The inlined path linearly searched physical
row and cell runs on every lookup. That materially cleared the profile-first
gate; XML open and validation were not changed.

The harness now has 110 selectable cases while its 36-case / 198-record
default matrix is unchanged. The ODF push smoke has 22 cases and the
tiny/large release matrix has 41 records because each family's create-small
case remains tiny-only.

## Change and bounds

`Spreadsheet` now owns a private atomic query counter and
`OnceLock<Option<CellLocator>>`. The first 63 successful named-sheet lookups
retain the old linear path. The 64th initializes one sheet-aligned locator;
concurrent callers share the same initialization. Open/list/point-query
workloads therefore perform no eager index construction.

The locator stores no sheet names, rows, cells, strings, or expanded logical
grid. It retains one 12-byte descriptor per physical row, plus cumulative
`u32` endpoints only when a row or cell repeat is not one. Ordinary dense
sheets map physical rows and cells directly. The generated 2 x 128 x 128
large corpus requests 3,216 bytes of index storage per warmed snapshot rather
than 32,768 cell endpoints.

A two-pass builder computes its requested storage with checked arithmetic and
refuses more than 4 MiB before allocation. Every vector uses
`try_reserve_exact`; overflow, conversion, budget or allocation failure stores
`None` and permanently uses the established linear lookup for that snapshot.
Validated ODS logical endpoints fit `u32`; out-of-domain queries still return
`CellView::Missing`. Every indexed hit borrows the original `Cell`, preserving
pointer identity and `CellView` semantics.

The locator cannot alter malformed-input order: `from_package` still completes
definitions, sheets, metadata and calculation-settings parsing before it
constructs the empty lazy state. Every facade mutation already replaces
`self` with a fully reopened `Spreadsheet`; the replacement therefore starts
with an empty counter and locator. No public API, package bytes, transaction,
patch, preservation rule, dependency edge, unsafe code, runtime or ambient
I/O changed. `Spreadsheet` remains `Send + Sync`.

## Matched latency measurement

The common-harness baseline executable SHA-256 is
`54838dac48192af2dc0a91e0bea7cf2e1667484f1b4bae46bd565f1e0154763e`.
The final executable SHA-256 is
`8c92c33ebe285bf5a3138c90e3eea5e04b4ffa8536da61cfe572f473c84ff40e`.

Environment: release Rust 1.95.0, Linux 6.8.0-101-generic, x86_64 AMD EPYC
9575F VM, Rust system allocator, CPU 4 pinned with `taskset`, and
`perf_event_paranoid=1`. The deterministic large corpus has two 128 by 128
sheets and 32,768 cells. Its 98,892-byte archive SHA-256 is
`7f0c43561602aedac7c5e91915f55b3515371d327ae69ac7fc0fe42b655db3f2`.

The primary ABBA matrix used 20 warmups and 500 samples per leg. Pooling 1,000
samples per state gives:

| Public query | Before p50 | After p50 | p50 delta | Mean delta | p95 delta | p99 delta |
|---|---:|---:|---:|---:|---:|---:|
| Large cell sweep | 2.049 ms | 0.374 ms | **-81.74%** | **-80.72%** | -75.43% | -69.94% |
| Medium cell sweep | 44.666 us | 13.723 us | **-69.28%** | **-67.33%** | -58.87% | -56.18% |
| Large full cell text | 3.047 ms | 1.443 ms | **-52.65%** | **-52.30%** | -48.30% | -47.99% |
| Medium full cell text | 74.954 us | 49.017 us | **-34.60%** | **-35.35%** | -32.54% | -20.48% |

The approximate independent-sample 95% interval for the large sweep mean
delta is `[-80.98%, -80.45%]` of the before mean. Both after legs are lower
than both before legs on p50 and mean. After profiling reduces the combined
exclusive `Spreadsheet::cell`, locator-build and indexed-view samples to
0.56% (0.19% + 0.30% + 0.07%), from the former 7.47% cell frame.

## Guardrails, memory and counters

A separate 30-warmup, 1,000-sample-per-leg ABBA isolates the one-cell guard.
Pooling 2,000 samples per state, large p50/mean improve 6.84%/5.23%; p95 moves
0.42%. Medium values also improve, although their 80-151 ns medians are timer
resolution territory. An equally sized open-only ABBA keeps large p50/mean at
+0.20%/+0.18% and medium at +0.84%/+1.00%; all open p50/mean/tails remain
within 3%. List-after-open stays a 20-60 ns operation and cannot execute the
cell-query branch.

Matched Heaptrack processes used five warmups and 100 large sweeps:

| Metric | Before | After | Delta |
|---|---:|---:|---:|
| Allocation calls | 74,691,573 | 74,691,891 | +318 (+0.0004%) |
| Temporary allocations | 17,840,929 | 17,840,824 | -105 |
| Peak heap | 41.70 MiB | 41.70 MiB | unchanged |
| Heaptrack RSS | 47.28 MiB | 47.28 MiB | unchanged |
| Leaked bytes | 1.78 KiB | 1.78 KiB | unchanged |

The three retained allocations per warmed dense snapshot are the sheet array
and two row-descriptor arrays. Uninstrumented GNU Time ABBA reports maximum
RSS of 42,976/43,012 KiB before and 42,980/42,976 KiB after, flat at page
granularity. Minor faults are 21,185/21,187 versus 21,187/21,187; every leg has
zero major faults.

Matched process-wide `perf stat` ABBA at the same 105 iterations per leg gives:

| Counter | Before A+B | After A+B | Delta |
|---|---:|---:|---:|
| task clock | 13,229.35 ms | 11,869.82 ms | -10.28% |
| cycles | 65,110,020,901 | 58,766,764,972 | -9.74% |
| instructions | 265,952,946,872 | 248,069,097,055 | -6.72% |
| branches | 56,504,834,877 | 51,462,893,155 | -8.92% |
| branch misses | 58,096,868 | 34,962,306 | -39.82% |
| cache references | 1,406,215,835 | 1,288,917,276 | -8.34% |
| cache misses | 75,432,642 | 69,474,058 | -7.90% |
| minor / major faults | 42,328 / 0 | 42,326 / 0 | flat |

## Correctness verification

- indexed and linear results are differential-tested across multiple sheets,
  direct/repeated/empty row and cell runs, every run boundary, out-of-range
  coordinates, and original-cell pointer identity;
- zero-budget fallback, maximum validated row/cell repeat endpoints, the exact
  64-query trigger, concurrent first construction, `Send + Sync`, and facade
  replacement invalidation have focused tests;
- complete `litchi-ods --all-features` tests pass: 241 tests across unit,
  transaction, malformed/security, real-corpus and integration suites;
- production-library warning-denied Clippy, 23 harness tests, harness
  warning-denied all-target Clippy, formatting, diff checks, JSON/hash checks
  and the release 22-case ODF smoke pass.

The unchanged broader all-target ODS Clippy gate still reports its six known
unrelated test/module lints. Warning-denied ODS rustdoc still reports the known
broken `super::Cell::text` link in `model/hyperlink.rs`. Neither file changed.

Raw ABBA, focused guard, Heaptrack, GNU Time, `perf stat`, profile and smoke
evidence is under `docs/performance/results/`; all committed evidence digests
are in `ods-cell-locator-sha256.txt`.

## Next non-iWork audits

1. OLE2: profile the XLS-only terminal validated-render handoff without
   retaining a shared editor-wide byte cache.
2. OOXML: profile XLSX action-plan flattening in medium and dense-wide 1%
   commit/save cases before changing the writer.
3. RTF: add byte-1252, LZFu, LibreOffice watermark and relative-font-size
   coverage before another production specialization.
4. ODF: keep source-backed reads, ODT/ODP repeated scans and unchanged-member
   publication separate; the ODS facade cell-scan item is complete.

iWork remains deferred while the `iwa-*` crates are modified independently.
