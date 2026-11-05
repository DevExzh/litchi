# Paired RTF run insertion

Run `cargo run --manifest-path tools/perf-baseline/unified-rtf-run-add/Cargo.toml --release`.

Thirty-six cases cover start, middle and end insertion into a selected paragraph
with 0, 1, 10, 100, 1,000 or 10,000 groups of three runs, in separate paragraphs
or one dense paragraph. Native and unified inserts use matching text and local
formatting. Owners, paths, node properties and native formatting are prepared
outside timing. Two warm-up rounds and nine ABBA rounds yield 18 samples per
route. Exact native/unified bytes, returned created paths, reopened run text and
formatting, paragraph counts and source immutability are checked outside timing.

This matrix measures individual insertions against increasing document size;
it does not qualify repeated-command coalescing, allocation counts, or all
admitted source shapes. Receipts record source/compiler/CPU and binary SHA-256;
never commit the executable.
