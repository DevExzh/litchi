# Change 0686 native XLS index-budget retry probe

This standalone package measures repeated selected-cell queries when the
worksheet occurrence index has a bounded logical budget. It is intentionally
separate from the workspace and uses only the public `litchi-xls`
source-backed API. The binary uses the production `OwnedSource` and
`FileSource` adapters directly, with the system allocator and no counted
`ReadAt` wrapper or global allocation instrumentation. This keeps the native
timing path free of instrumentation atomics.

The package is a diagnostic harness, not a production dependency. The root
coordinator owns Cargo builds and generates the package lock file in the
measurement checkout.

## Build

From the repository root, the coordinator can build it with:

```sh
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-target-0686-after \
  cargo build --manifest-path docs/performance/results/change-0686/probe/Cargo.toml \
  --release --locked --offline
```

The binary is then at
`/home/zhuhe/code/litchi-target-0686-after/release/xls-index-retry-probe-0686`.

## Run

```sh
taskset -c 12 \
  /home/zhuhe/code/litchi-target-0686-after/release/xls-index-retry-probe-0686 \
  --input test-data/poi/test-data/spreadsheet/54016.xls \
  --budget 1048576 --mode owned --worksheet 0 --row 0 --column 0 \
  --queries 8 --warmups 3 --samples 30 > retry.json
```

`--queries` defaults to 8, `--warmups` to 3, and `--samples` to 30.
`--budget` is the logical `SourceBackedLimits::max_query_index_bytes` ceiling;
zero disables the optional index. Coordinates are zero-based. `owned` reads
the fixture once before the sample loop and creates a fresh
`OwnedSource::from_arc` and workbook owner for every warmup and measured
sample. `file` opens a fresh `FileSource` before each sample's open timer.

The report records the input hash and coordinates, one timed owner-open record,
and one timed record for every query in every measured sample. Each query
contains a full semantic outcome: `value`, `missing`, or `error`. Values retain
formula metadata and recursively retain cached values; floating-point values
are represented by their exact IEEE-754 bits. `all_queries_agree` makes a
budget-retry semantic mismatch visible without replacing the full outcomes.

Source construction is outside the open timer. The open timer covers
`SourceBackedWorkbook::from_read_at_with_limits`; each query timer covers only
`cell_value_by_index`. Outcome projection and JSON serialization happen after
those timers. The returned `CellValue` remains alive through outcome capture
and is retained for the sample before the next query.

Allocation calls, requested bytes, peak live bytes, and retained live bytes are
collected by the separate change-0686 allocation probe. Combining that probe
with this native package preserves an uninstrumented timing path; this package
does not accept an allocator flag.
