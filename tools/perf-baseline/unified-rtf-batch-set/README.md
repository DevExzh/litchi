# Paired RTF ordered-batch set benchmark

Run the standalone release harness with:

```sh
cargo run --manifest-path tools/perf-baseline/unified-rtf-batch-set/Cargo.toml --release
```

The harness measures four compatible ordered mutation families: paragraph text,
paragraph alignment plus local layout, run text, and compound run formatting.
Each family uses counts of 1, 10, 100, 1,000 and 10,000. Paragraph families
use separate paragraphs; run families use both separate paragraphs and one
dense paragraph containing the selected runs.

Native updates are staged through their ordered batch APIs. The unified side
executes one direct path `set` per prepared target inside one edit closure.
Both routes use the same source bytes and matching finite operation and
payload limits. Paths, properties, native updates and format patches are
prepared before timing. Two warm-up rounds followed by nine ABBA rounds give
18 samples per case and route.

Every sample checks exact native/unified bytes, native operation count,
`diagnostics.commits() == 1` for the unified batch, reopened paragraph/run
semantics, and source immutability outside the timed region. The JSON schema is
`rtf_batch_set_benchmark.v1`.

This benchmark measures compatible direct-set batching and its scaling with
the number of selected nodes. It does not qualify unsupported source closure,
query fan-out, allocation counts, or unrelated RTF operation families. Keep
the executable and raw benchmark output outside the repository; receipts may
record the source commit, build command, binary size, and SHA-256.
