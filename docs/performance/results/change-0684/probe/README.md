# Change 0684 XLS query-cache probe

This is a standalone probe for the source-backed XLS selected-cell cache. It
uses only the public `litchi_xls::SourceBackedWorkbook` API, so the same source
builds against the before checkout (`5805d54a1`) and a candidate checkout. It
does not enable a candidate-only feature and it does not implement the cache.

The probe has two entry points:

* `route` measures a fresh owner and one of five bounded routes: `cold1`,
  `cold2`, `cold3`, `prepared`, or `visit`.
* `corpus` opens every `.xls`/`.xlt` below a root, retaining typed open and
  worksheet refusals and comparing selected queries with a complete visitor
  oracle where the worksheet admits the visitor.

The `prepared` route creates one fresh owner and runs `q1(A)`,
`q2-build-trigger(B)`, and `q3-warm-repeat-q1(A)`. The three cold routes each
create a new source-backed owner and perform one `A` query. `visit` times a
`visit_cells` call with an empty callback and reports the actual callback count;
it performs a second untimed visitor pass to retain a digest and returned
values for semantic and allocation checks. Query projection, digest work, and
the visitor oracle are outside the elapsed time and source counters reported
for the measured call.

`owned` counts reads over an immutable in-memory `ReadAt`. `file` counts
positional reads over the original fixture and also reports the union of
logical byte ranges observed by the wrapper. Both adapters use a stable
synthetic `SourceVersion`; they measure logical source behavior and warm-cache
file access, not physical I/O. `owned-native` and `file-native` use the
production `litchi-core` source adapters without counter atomics or range-union
bookkeeping; their source counters are intentionally unavailable and their
elapsed values are clean native timing controls for the counted routes. Every
route sample opens a fresh owner, so an index cannot leak between samples.

Build the probe in each checkout with the checkout's own standalone lock and
target directory. The root coordinator owns the single Cargo lane and should
run these commands serially:

```sh
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-target-0684-before \
  cargo build --manifest-path docs/performance/results/change-0684/probe/Cargo.toml \
  --release --locked --offline --bin xls-index-probe-0684

CARGO_TARGET_DIR=/home/zhuhe/code/litchi-target-0684-after \
  cargo build --manifest-path docs/performance/results/change-0684/probe/Cargo.toml \
  --release --locked --offline --bin xls-index-probe-0684
```

The `debug = 2` release profile is intentional: it keeps symbols for a later
native profile while retaining release optimization. Record the resulting
binary hashes and the source/manifest hashes beside each capture.

A primary route invocation is:

```sh
/home/zhuhe/code/litchi-target-0684-before/release/xls-index-probe-0684 \
  route --input test-data/ole/xls/WithCustomViews.xls --route prepared \
  --mode owned --worksheet 0 --row 0 --column 0 \
  --second-row 0 --second-column 1 --warmups 2 --samples 7
```

Use `--mode file` with the same arguments for the counted positional-file
route, or `--mode owned-native`/`file-native` for uninstrumented timing
controls. The stored target above is on `WithCustomViews.xls`, worksheet 0
(`Plan1`); do not use its empty worksheet 1 as the representative route. The
source-backed route for `ConditionalFormattingSamples.xls` is a refusal
control on worksheets whose shared-formula metadata the reader declines. The
route output retains the exact error text in `outcome.error`.

The corpus differential is bounded but covers the cases that matter for cache
publication:

```sh
/home/zhuhe/code/litchi-target-0684-before/release/xls-index-probe-0684 \
  corpus --root test-data --mode owned --sample-coordinates 16 --max-queries 512 \
  > corpus-before-owned.json

/home/zhuhe/code/litchi-target-0684-before/release/xls-index-probe-0684 \
  corpus --root test-data --mode file --sample-coordinates 16 --max-queries 512 \
  > corpus-before-file.json
```

Run the same two commands with the candidate binary. Corpus JSON uses paths
relative to the supplied root and the fixed `<corpus-root>` marker, so the
before/after reports can be compared without checkout-path normalization. Each eligible worksheet
is first walked completely. The selected set includes every duplicate
coordinate, first and last stored coordinates, `(0,0)` and deterministic
missing edge targets, plus up to 16 evenly spaced stored coordinates. The
`max-queries` bound applies only to ordinary sampled coordinates; duplicate,
edge, and missing coordinates are retained even when the ordinary sample is
truncated. Every selected coordinate is queried once and the first selected
coordinate is queried again as `warm-repeat`. A query agrees only when the
returned `CellValue` equals the visitor's last occurrence at that coordinate;
errors and missing cells remain distinct. `query_mismatches` must be zero for
an accepted differential, while `query_truncated` states whether ordinary
sampling was capped.

The JSON is evidence about route behavior and semantic parity. It is not a
registered timing claim. Native `perf stat` captures, allocator captures, and
the candidate-only cache-budget controls belong to the packet-level drivers
outside this small public-API probe.
