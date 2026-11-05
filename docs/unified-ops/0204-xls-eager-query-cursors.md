# 0204 — Eager XLS query cursors

Eager XLS cell-only query walks now borrow the native cell iterator directly.
They no longer construct a facade `BTreeMap`, row/column lists or a complete
coordinate vector before yielding the first result. Each step projects only
its selected cell; the iterator borrows the immutable file owner. The native
worksheet iterator itself now walks `BTreeMap::Values` instead of collecting a
`Vec<&Cell>` at construction. Source fences and fused termination remain on the
facade iterator. A concrete native `stored_cells()` iterator avoids boxing
each yielded cell; the facade cursor remains `Send + Sync`, enforced at compile
time.

Native `Workbook::sheet_count()` and selector-based `Workbook::sheet()` expose
the complete tab catalog without cloning all entries. The positional source
owner has the corresponding fenced selected-tab projection. Ordinal facade
lookup and catalog traversal use these methods, preserving chart/macro tab
positions independently of parsed worksheet ordinals. Named facade lookup
still scans the name axis with its existing Unicode comparison and ambiguity
rules; it no longer repeats the catalog copy after finding a match.

## Verification

The combined native XLS and XLS/encryption facade suites pass 2,261 tests and
doctests, with one ignored. Allocation regressions compare both requested byte
counts and allocation counts for eager `.first()`/`.limit(1)` at 1 and 10,000
cells, native first-cell iteration at those sizes, and ordinal `get` at 1 and
100 sheets. Fixture construction and opening are outside those measurements;
path/value assertions verify the selected cell. The tests reject sheet-sized
or catalog-sized intermediate allocations. Catalog tests compare selected
name/position results, missing entries and complete counts on both native
owner routes using real hidden-tab and chart-tab fixtures.

The all-feature facade suite passes 1,498 tests and doctests. Warning-denied
all-target Clippy and rustdoc, the minimal facade feature build, and the XLS
glossary audit (24 banned names, one transaction family) pass. The boundary
check passes for 66 packages and 268 internal dependencies with 11 existing
debt items.

Reproduce with `cargo test -p litchi-xls -p litchi-ops --no-default-features
--features litchi-ops/xls,litchi-ops/encryption`. Offline verification uses an
external lockfile selecting cached cc 1.6.0 and an external target; repository
lockfiles are unchanged.

## Paired eager read measurement

Five release samples, median microseconds per complete scan:

| Cells | Native get | Unified get | Ratio | Native iteration | Unified query | Ratio |
|---:|---:|---:|---:|---:|---:|---:|
| 1 | 0.030 | 0.410 | 13.67× | 0.040 | 0.550 | 13.75× |
| 100 | 1.340 | 37.280 | 27.82× | 0.640 | 15.160 | 23.69× |
| 1,000 | 17.410 | 379.882 | 21.82× | 6.150 | 150.331 | 24.44× |
| 10,000 | 333.711 | 3943.147 | 11.82× | 75.540 | 1578.066 | 20.89× |

Every paired value, count and checksum agrees. The 10,000-cell facade query
median is 1.58 ms versus record 0203's 2.41 ms (about 35% lower in these runs).
That comparison is observational, not a controlled noise qualification. The
native/facade legs above share a corpus, build and run; facade overhead remains
substantial. The numeric single-column eager corpus excludes generation,
opening, path construction and query compilation. It does not qualify source
reads, mixed data, memory peaks, Python/CLI or the full performance gates.

The [measurement JSON](results/0204-xls-read.json) contains all samples.
Source commit: `fd3144e785380a548d179aa5623040dcc199f8ab`. Reproduce with
`LITCHI_XLS_BENCH_SAMPLES=5 cargo run --release -p litchi-ops --features xls
--example xls_read_bench`. The measured binary SHA-256 is
`ed50f9e09d78ceab3a622548212d0a85a79aa73eb6cd29c30f9925341f0ec3de`; no binary is committed.

## Remaining scope

This closes eager **cell-only** query materialization and repeated catalog
copies. Source-backed queries, row/column/range projections and mixed-kind
walks still materialize sparse grids. They remain required work, not accepted
performance exceptions. A source cursor must retain bounded parser state
across calls, handle formula STRING continuations and multi-cell records,
keep source-change and limit checks, and preserve query order and duplicate-cell
semantics. The current visitor reports physical records, including duplicate
coordinates, while the grid uses the last record in row-major order; substituting
a physical record iterator without resolving that difference would change
results. Early stop also needs an explicit native validation contract for the
unvisited suffix. None of those guarantees is established by the eager tests.

Unified XLS mutations, batching/history, native security publication, remaining
format coverage, Python/CLI and the full goal's performance gates remain open.
