# Paired RTF run reads

Run `cargo run --manifest-path tools/perf-baseline/unified-rtf-runs/Cargo.toml --release`.
The standalone target shares the facade example source.

Each source contains 1, 10, 100, 1,000 or 10,000 ordinary paragraphs with three
native text runs apiece. Native and unified routes read the last run by path,
traverse every bold run, or stop at the first bold run. Owners, the path and
compiled query plan are prepared outside timing. Lookup, predicate evaluation
and borrowed-text length checks are timed; complete text-oracle comparisons,
source-byte checks and serialization are outside timing.

Two warm-up and nine measured ABBA rounds produce 18 samples per route/case.
Samples repeat each read 1,000 times for first-match and up to 10 paragraphs,
100 times for 100 paragraphs, and 10 times otherwise. Raw sample durations
measure the batch; medians are normalized per read. Every timed sample must
match the repeated native count/text-byte checksum. Untimed
checks compare the full bold-run text sequence and last-run text exactly.

Run three trials pinned to one available CPU. Preserve samples, source
commit/tree, compiler, binary size/SHA-256 and commands without committing the
binary. These measurements cover read traversal and early termination, not
mutations, open/save costs or allocation counts.
