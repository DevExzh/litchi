# Paired RTF paragraph layout

Run `cargo run --manifest-path tools/perf-baseline/unified-rtf-layout/Cargo.toml --release`.
The standalone target shares the facade example source.

The `rtf_layout_benchmark.v1` route compares a native layout batch against
unified `set` on the last paragraph by path and all paragraphs by query. Each
command writes spacing, indentation and three pagination flags in one native
layout delta per selected paragraph. Each source has 1, 10, 100, 1,000 or 10,000 paragraphs. Both owners are warmed.
Native typed updates and facade selectors/properties are prepared outside
timing; facade query expansion remains inside timing. Matching native limits
allow 10,000 operations and 16 MiB of replacements.

Two warm-up and nine measured ABBA rounds produce 18 samples per route/case.
Staging and commit are timed. Serialization, exact native/facade byte checks,
reopened text/layout validation and output destruction are excluded.
Every sample is checked and input owners remain unchanged.

Run three trials pinned to one available CPU. Preserve JSON, source commit/tree,
compiler and binary size/SHA-256, without committing binaries. This qualifies
one command and query fanout, not repeated-command batching or allocations.
