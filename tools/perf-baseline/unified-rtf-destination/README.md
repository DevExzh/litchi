# Paired RTF destination-story writes (record 0275)

Run the standalone release harness with:

```sh
cargo run --manifest-path tools/perf-baseline/unified-rtf-destination/Cargo.toml --release
```

The `rtf_destination_benchmark.v1` route compares the native RTF destination
editor with the unified API for table-cell, footnote, header/footer-paragraph
and comment text, at 1, 100 and 1,000 targets, and native exact-source picture
removal at 1 and 64 pictures (the native per-transaction maximum).

Routes per case:

- **native** stages one native operation per target and commits once;
- **unified batch** issues one direct path `set` per target inside an edit
  closure; compatible destination writes coalesce into one native publication
  (picture removal issues one direct `remove` per picture, from the last);
- **native shared** and **unified fan-out** write one shared value to every
  target through the native editor and through one query `set`/`remove`.
  Header/footer paragraphs have no fan-out route: the query walker does not
  descend into header/footer stories;
- **unified direct** issues one `File::set`/`remove` without an edit closure,
  for single-target cases only.

Sources, paths, properties and native operations are prepared before timing;
staging and commit are timed. Every sample's output is checked byte-for-byte
against the native oracle outside the clock, and source immutability is
checked per case. Two warm-up rounds precede nine `ABCDEEDCBA` rounds (18
samples per route).

Run three trials pinned to one available CPU and keep the JSON outside the
repository; receipts record the source commit, compiler, CPU and binary
size/SHA-256. Do not commit the binary.
