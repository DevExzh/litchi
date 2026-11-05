# Paired RTF paragraph swap

Run `cargo run --manifest-path tools/perf-baseline/unified-rtf-swap/Cargo.toml --release`
from the repository root. The standalone target shares the facade example source.

The `rtf_swap_benchmark.v1` route compares one native swap with unified `swap`
on the first and last paragraphs of 1, 10, 100, 1,000 and 10,000-paragraph
ordinary RTF documents. The one-paragraph case is an exact no-op. Both owners
are warmed; selectors and matching limits are prepared outside timing.

Two warm-up ABBA rounds precede nine measured ABBA rounds (18 samples per route).
Staging and commit are timed; setup, serialization, exact native/facade byte
checks, reopened paragraph order and output destruction are excluded. Every
sample is validated and both input owners remain unchanged.

Run three trials on one available pinned CPU. Keep the JSON, source commit/tree,
compiler and binary size/SHA-256. Never commit the binary. These single-operation
measurements do not qualify repeated-command batching or allocation behavior.
