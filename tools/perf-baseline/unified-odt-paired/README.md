# Paired ODT leaf/unified benchmark (record 0276)

Run the standalone release harness pinned to an idle core with:

```sh
taskset -c 9 cargo run --manifest-path tools/perf-baseline/unified-odt-paired/Cargo.toml --release -- results.json
```

Each scenario times the leaf `litchi_odt` call and the unified `litchi_ops`
call on the same generated document (100, 1,000 and 10,000 paragraphs) in
ABBA order, two warmups and nine rounds, and reports the medians:

- `open_first_get`: open from bytes and read the first paragraph's text;
- `query_all`: every paragraph's text from a warm handle;
- `single_set`: replace one mid-document paragraph's text and commit;
- `batch_bold_N` (N = 1, 10, 100, 1,000, 10,000): bold on N paragraphs, one
  leaf `Edit` against one coalesced unified `File::edit`.

The harness asserts outside the clock that the unified 1k batch publishes the
same bytes as the leaf batch. The optional argument writes the JSON rows.
