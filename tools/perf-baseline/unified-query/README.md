# Paired unified query cost

This standalone harness (record 0262) pairs unified `File::get` and
`File::query` (full match, `first()` and `limit(10)` early termination) with
hand-written loops over the source-backed leaves, on generated XLSX
(1,000 and 50,000 rows) and PPTX (10×8 and 500×40) corpora and two DOCX
fixtures from `test-data`, each opened from memory (`bytes`) and from a file
(`file`). Route names end in `_v1`.

```sh
CARGO_NET_OFFLINE=true CARGO_TARGET_DIR=target/bench cargo build --release \
  --manifest-path tools/perf-baseline/unified-query/Cargo.toml
UNIFIED_QUERY_DIR=/some/scratch taskset -c 0 \
  target/bench/release/unified-query-baseline time all 15 > out.jsonl
```

Each JSON line holds per-sample unified and leaf times (in-process ABBA) and
the median ratio. Only public APIs are used, so binaries built before and
after a change can be alternated. Never commit the compiled binary.
