# Paired unified query cost

This standalone harness (record 0262) pairs unified `File::get` and
`File::query` (full match, `first()` and `limit(10)` early termination) with
hand-written loops over the source-backed leaves, on generated XLSX
(1,000 and 50,000 rows) and PPTX (10×8 and 500×40) corpora and two DOCX
fixtures from `test-data`, each opened from memory (`bytes`) and from a file
(`file`). Route names end in `_v1`.

Record 0267 adds new routes only: `xlsx_*_cold_*`, `xlsb_*_cold_*` and
`xls_*_cold_*` open a fresh handle in every call and stop at the first or
tenth `cell[value>1000]` match (the cold first-match cost of a worksheet no
earlier read has decoded), against the leaf's own early-stopping read of a
freshly opened workbook; and warm `xlsb_*` get/first/full routes like the
XLSX ones. The XLSB and XLS workbooks have the XLSX corpus's shape and are
written by the leaves' builders.

```sh
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=target/bench cargo build --release \
  --manifest-path tools/perf-baseline/unified-query/Cargo.toml
UNIFIED_QUERY_DIR=/some/scratch taskset -c 0 \
  target/bench/release/unified-query-baseline time all 15 > out.jsonl
```

Each JSON line holds per-sample unified and leaf times (in-process ABBA) and
the median ratio. Only public APIs are used, so binaries built before and
after a change can be alternated. Never commit the compiled binary.

Record 0287 adds the `all-formats` cargo feature, which compiles the unified
layer with every non-iWork format as the `litchi` facade with all features
does, so the unified `Node` has the size a full build gives it (408 bytes
before record 0287, 184 after). Routes and their names are unchanged.

```sh
CARGO_TARGET_DIR=target/bench cargo build --release --features all-formats \
  --manifest-path tools/perf-baseline/unified-query/Cargo.toml
```

Record 0315 adds new routes only: `xlsb_*_all_cell_owned_v1` and
`xlsb_*_all_cell_fenced_v1`, the fair leaf legs of the full XLSB `cell`
query that record 0280 added for XLSX and PPTX. The leaf keeps every
stored cell as an owned handle (the worksheet position, a shared grid
handle and the coordinate) and, in the fenced leg, reads the leaf's
`source_version` once per match.
