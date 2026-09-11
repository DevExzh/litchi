# XLSX semantic commit guard probe

This is a standalone public-API probe for the 0514 XLSX commit optimization
decision. It does not modify the performance harness or production crates. The
probe creates one deterministic two-sheet numeric workbook per selected shape:
`tiny` and `medium` use 8x8 and 32x32 worksheets, while `dense-wide` uses two
256x256 worksheets. The generated archive is created once per process, hashed
with SHA-256, and reused as the source for every fresh iteration.

The seven scenarios are a cold first-cell read, cold same-value one-cell and
one-percent edits, warmed Store same-value one-cell and one-percent edits, and
warmed Store changed one-cell and one-percent edits. A Store warmup enumerates
the complete public `Worksheet::cells(Rect::ALL)` view for each worksheet
touched by an edit. The edit is prepared after that warmup, and commit clocks
surround only `litchi_xlsx::Edit::commit`. The cold first-cell scenario clocks
only the first public `Worksheet::cell` call that forces an ordinary Store
parse. Output serialization for same-value exact-byte checks, changed-cell
readback, first-cell numeric readback, and returned `Commit`/cell-view drops
happen after the clock. The report retains every measured nanosecond and its
zero-based stable sample index, along with min/p50/p95/p99/max and mean
statistics.

Commit scenarios retain the preceding iteration's final `Commit` until it is
replaced after the next measured region; this lifetime is identical in both
roles. Compare memory within the same scenario and corpus, and report absolute
entry/peak values separately from `region_peak_live_bytes - live_bytes_before`.
The latter is incremental callback-ordered live demand, not document peak or RSS.

Each measured scenario also retains one allocation `Sample` per measured
elapsed value. The sample is unavailable in the normal binary and measured in
the allocator-feature binary using the shared, unchanged
`tools/perf-baseline/src/allocation_metrics.rs` observer and
`tools/perf-baseline/src/bin/support/counting_allocator.rs` wrapper. The
allocation region begins immediately before `Instant::now` and finishes after
the elapsed value is collected, so post-clock oracles and destruction do not
enter the allocation vector. The raw elapsed vector is chronological; the
reported percentile uses a sorted copy and index
`floor((n - 1) * p / 100)`.

The repository-relative manifest is intentionally an independent Cargo
workspace. The same frozen lockfile is used by both roles; `guard-run.py` records the
probe and production source manifests and executable hashes separately.
Normal and allocator builds and all captures are retained in the parent
evidence directory. Report paths use `create_new` so a retained capture
cannot be overwritten accidentally.

Example command after the controlling driver has built the manifest:

```text
litchi-xlsx-commit-guard --samples 100 --warmup 3 \
  --shape tiny,medium,dense-wide \
  --scenario cold-first-cell-read,cold-same-one-cell,cold-same-one-percent,warm-same-one-cell,warm-same-one-percent,warm-changed-one-cell,warm-changed-one-percent \
  --json guard-report.json
```
