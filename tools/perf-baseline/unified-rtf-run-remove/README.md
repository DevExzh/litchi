# Paired RTF run removal

Run `cargo run --manifest-path tools/perf-baseline/unified-rtf-run-remove/Cargo.toml --release`.

Twenty cases cover 1 to 10,000 groups of three mixed-format runs, in separate
paragraphs or one dense paragraph. Native empty replacements and unified remove
select the last middle run or every italic run, in one native publication.
Owners, selectors and native updates are prepared before timing. Two warm-ups
and nine ABBA rounds yield eighteen samples per route. Serialization, exact byte
comparison, reopened text/format/count checks and source immutability checks are
outside timing. Three CPU-pinned trials retain raw samples and source/compiler/
binary size/SHA-256 metadata; never commit the benchmark executable.

The matrix measures selector fan-out, not repeated-command coalescing,
allocation counts, or every admitted source shape.
