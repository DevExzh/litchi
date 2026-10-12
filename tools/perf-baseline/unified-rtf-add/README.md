# Paired RTF paragraph insertion

Run `cargo run --manifest-path tools/perf-baseline/unified-rtf-add/Cargo.toml --release`
from the repository root. The standalone target shares the facade example source.

The `rtf_add_benchmark.v1` route measures native positioned insertion against
unified `add` at the body end on 0, 1, 10, 100, 1,000 and 10,000 paragraphs.
Both routes reuse warmed immutable owners and matching native limits. The body path and new-node
arguments are prepared outside timing. Two warm-up ABBA rounds precede nine measured ABBA
rounds (18 samples per route). Staging and commit are timed; setup, serialization,
exact native/facade byte checks, reopened paragraph order and output destruction
are excluded. Every output is validated and source immutability is checked.

Run three trials on the same available pinned CPU and retain the JSON output,
source commit/tree, compiler and binary size/SHA-256. Do not commit binaries.
Shared-host timings are evidence for these single insertions; they do not qualify
repeated-command batching or allocation behavior.
