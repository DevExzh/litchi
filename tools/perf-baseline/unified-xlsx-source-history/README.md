# XLSX source-history adapter builds

The three retained-history allocator binaries are versioned probes, not a
compatibility surface for the current XLSX API:

- `xlsx_merge_history_alloc`
- `xlsx_tab_state_history_alloc`
- `xlsx_source_reader_paths_alloc`

The current workspace contains only the current `OpenOptions` adapters. There
is no legacy compatibility feature and no old-constructor branch in these
files. Historical source-history capability is preserved by the exact e63
checkout and its external, recorded overlay; it is not hidden behind
`--all-features` precedence in the current manifest.

## Current default and all-feature builds

The existing coordinator target is
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/hist-current`.
The default feature set and `--all-features` both select the current
`SourceBackedWorkbook`/`OpenOptions` path:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/env.sh
CARGO_NET_OFFLINE=true cargo build --release --locked \
  --features allocator-metrics \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin xlsx_merge_history_alloc \
  --bin xlsx_tab_state_history_alloc \
  --bin xlsx_source_reader_paths_alloc \
  --target-dir /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/hist-current

CARGO_NET_OFFLINE=true cargo build --release --locked --all-features \
  --manifest-path tools/perf-baseline/Cargo.toml \
  --bin xlsx_merge_history_alloc \
  --bin xlsx_tab_state_history_alloc \
  --bin xlsx_source_reader_paths_alloc \
  --target-dir /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/targets/hist-current
```

Record the candidate commit and each binary's SHA-256 beside the report.

## Historical reproduction

The actual rewritten pre-W4 revision in the retained records is
`e63a1fe3534bcaaae2ce367339f7f8275877baef`, not
`0befb112d96d2409e6f40c9abfeb4813665905ad`. At e63,
`crates/litchi-xlsx/src/workbook/source_merge.rs` exports
`workbook::SourceBackedEditor` with `from_read_at` and the
`from_read_at_with_limits*` family, and
`crates/litchi-xlsx/src/tab_state/source.rs` exports
`tab_state::SourceBackedEditor` with the same constructor family and
`edit(&self)`. Those signatures are verified in the exact e63 checkout; they
are not inferred from the current source.

The retained historical merge adapter and its manifest overlay are already in
the external worktree
`/home/zhuhe/code/litchi-worktrees/uo2-e63-prew4-alloc-v1`. Its source SHA-256
is `3929b1fbdaf4771d0667bc8a79cf0e274b72eebca6c064cb4a64d9c900c8f1b5`, and
the manifest overlay SHA-256 is
`d89eff26ca16ebb28e4b153ef790eed57379a0ad1516775dd8046de2e1ecc545`.
The exact base archive is outside the checkout at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/e63-prew4-allocator-base-v2.tar`
(SHA-256
`c9281495a28f5abaea8326091071c2f4e77a5015a38596709c287fcd74cee42a`), with
path-list SHA-256
`d69e80ecb137da9d9f040afa0fe26e3f2b46a734279b758b5942e81ce0469d27`.
The snapshot overlays for the tab-state and reader probes are retained at
`/home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/final-snapshots/e63`; they
are not current-workspace compatibility code.

Build the historical merge adapter only from that external e63 overlay. The
existing historical target is
`/home/zhuhe/code/litchi-worktrees/uo2-e63-prew4-target-v1`:

```sh
source /home/zhuhe/code/litchi-worktrees/.coord/w4-xlsx/env.sh
CARGO_NET_OFFLINE=true cargo build --release --locked --no-default-features \
  --features allocator-metrics,legacy-xlsx-source-history \
  --manifest-path /home/zhuhe/code/litchi-worktrees/uo2-e63-prew4-alloc-v1/tools/perf-baseline/Cargo.toml \
  --bin xlsx_merge_history_alloc \
  --target-dir /home/zhuhe/code/litchi-worktrees/uo2-e63-prew4-target-v1
```

The protected e63 merge binary is outside the repository at
`/home/zhuhe/code/litchi-worktrees/uo2-e63-prew4-target-v1/release/xlsx_merge_history_alloc`;
its recorded SHA-256 is
`12ff3a682f35f5c2292ed9232afe59aa78f3393c548584ab863d45e7802b8014`.
Use the same fixture, iteration, warmup, CPU-pinning, and report options for
the current and historical runs. Retain the source revision, binary
SHA-256, output oracle, replay/inverse checks, and allocator interval in the
external report. The historical route remains a diagnostic comparison and
does not reintroduce a legacy constructor into the current workspace.
