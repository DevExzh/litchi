# Paired RTF run text replacement

Run `cargo run --manifest-path tools/perf-baseline/unified-rtf-run-text/Cargo.toml --release`.

This shares the facade example source. Twenty cases cover 1 to 10,000 groups
of three mixed runs, in separate paragraphs or one dense paragraph. Each route
replaces the last or all italic runs with ASCII text and four formatting facets
in one native publication. Prepared owners, selectors, replacement strings and
property values are outside timing. Two warm-ups and nine ABBA rounds yield
18 samples per route. Exact native/unified bytes, reopened text and formatting,
and immutable source bytes are checked outside timing.

Run three CPU-pinned trials and retain raw samples, compiler/CPU, source
commit/tree, binary size/SHA-256 and regeneration commands. Never commit binaries.
Unicode correctness is covered separately by integration tests. This does not
qualify repeated-command coalescing, allocation counts or all source closures.
