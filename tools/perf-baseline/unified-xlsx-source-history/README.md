# XLSX source-history adapter builds

The three retained-history allocator binaries are versioned probes, not a
compatibility surface for the current XLSX API:

- `xlsx_merge_history_alloc`
- `xlsx_tab_state_history_alloc`
- `xlsx_source_reader_paths_alloc`

The current adapter is selected unless the historical feature is selected
without its current counterpart. Consequently `--all-features` always builds
the current `OpenOptions` path. The old constructor names remain only in the
historical branch used for a pinned pre-W4 reproduction.

## Current all-feature build

Run from the candidate worktree, with both the target directory and binaries
outside the checkout:

```sh
CARGO_NET_OFFLINE=true cargo build --release --locked --all-features \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin xlsx_merge_history_alloc \
  --bin xlsx_tab_state_history_alloc \
  --bin xlsx_source_reader_paths_alloc \
  --target-dir /outside/uo2-xlsx-history-current-target
```

Record the candidate commit and each binary's SHA-256 beside the report.

## Historical reproduction

The pre-W4 source is pinned to
`0befb112d96d2409e6f40c9abfeb4813665905ad`. Build that checkout's own
historical adapter; do not copy current consumer files into it. Selecting only
`legacy-xlsx-source-history` keeps the old owner/editor route and its measured
workflow intact:

```sh
git worktree add --detach /outside/uo2-xlsx-history-baseline \
  0befb112d96d2409e6f40c9abfeb4813665905ad
CARGO_NET_OFFLINE=true cargo build --release --locked --no-default-features \
  --features allocator-metrics,legacy-xlsx-source-history \
  --manifest-path /outside/uo2-xlsx-history-baseline/tools/perf-baseline/Cargo.toml \
  --bin xlsx_merge_history_alloc \
  --bin xlsx_tab_state_history_alloc \
  --bin xlsx_source_reader_paths_alloc \
  --target-dir /outside/uo2-xlsx-history-baseline-target
```

Use the same fixture, iteration, warmup, CPU-pinning, and report options for
both builds. The baseline and candidate reports must retain the source
revision, binary SHA-256, output oracle, replay/inverse checks, and allocator
interval. This preserves the historical measurement without compiling a
legacy constructor in the current all-feature workspace.
