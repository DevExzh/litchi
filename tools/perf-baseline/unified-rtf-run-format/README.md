# Paired RTF run formatting

Run `cargo run --manifest-path tools/perf-baseline/unified-rtf-run-format/Cargo.toml --release`.
The standalone target shares the facade example source.

Sources contain 1, 10, 100, 1,000 or 10,000 groups of three mixed character
runs, either in separate paragraphs or densely in one paragraph. Native and unified routes apply bold, size,
underline and strike in one atomic update to the last group's middle run
or to every italic middle run. Unselected run flags and all text are checked.
Native and unified semantic profiles both allow 10,000 operations and 16 MiB
of replacements.

Owners, selectors and properties are prepared outside timing. Two warm-up
and nine measured ABBA rounds produce 18 samples per route and size. Native
staging/commit and unified set are timed; serialization, exact paired byte
comparison, reopened character/text checks and output destruction are excluded.

Run three trials pinned to one CPU. Receipts include all samples, source
commit/tree, compiler, binary size/SHA-256 and regeneration commands. Do not
commit the binary. This measures selector fanout and single compound updates;
it does not qualify repeated-command coalescing or allocation counts.
