# Change 0260: fresh-child XLSX filesystem roots

## Status

Landed in `ae67e808d`. This change strengthens input-mode and cache-state
evidence for two existing XLSX selectors. It does not establish a latency,
throughput, allocation, RSS, or physical-I/O improvement claim.

## Scope

`xlsx_file_open` and `xlsx_file_open_lifecycle` now run every warmup and
retained sample in a fresh child process through the shared filesystem
runner. Both selectors use one pinned deterministic medium XLSX corpus:

- generator: `litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1`;
- shape: four sheets, 48 rows by 48 columns per sheet, and 9,216 cells;
- physical package: 17 ZIP members and 4,226,429 archive bytes;
- source SHA-256:
  `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036`.

The open selector times only `litchi::Workbook::open(Path)`. The lifecycle
selector additionally times worksheet names, worksheet count, and full-text
projection. Elapsed time, operation-local allocations, process metrics, and
cold-verification proof are captured while the exact timed workbook and
lifecycle projection remain live. Their correctness validation and
destruction occur only after those snapshots.

## Correctness and evidence boundaries

The measured workbook is checked after timing against an independently opened
typed `litchi_xlsx::Workbook` oracle and a separate OPC/property oracle. The
child reports both the final source hash and a deterministic semantic hash.
The fixed corpus manifest and archive hash are checked before any sample can
run, so generator drift cannot silently redefine the workload.

Warm and cold-requested cases use the ordinary source archive. A
`cold-verified` run uses a page-aligned ZIP copy only when the filesystem and
`fincore` admission checks succeed, the file starts with zero resident, dirty,
and writeback bytes, and the measured child observes a positive process
`read_bytes` delta. Otherwise the report records an explicit ineligible
status. The aligned source is hashed independently. This establishes a
page-cache/process-I/O proof only; it is not evidence that bytes reached a
particular physical storage device.

The facade path has no operation-local range-source counter, so logical-read
metrics are explicitly classified as `not_applicable_filesystem_xlsx` rather
than reported as zero work. The selectors remain opt-in. The default tranche
stays at 36 cases and 198 records, and the selectable case count remains 381.

## Verification

- `cargo test --manifest-path tools/perf-baseline/Cargo.toml --test xlsx_filesystem`:
  2 passed.
- `cargo fmt --all -- --check`: passed.
- `git diff --check`: passed.
- Strict validation of the four retained performance claims passed.
- Independent static review confirmed the fresh-child routing, timing and
  destruction boundaries, exact-object oracle, corpus pins, aligned-source
  hashes, cache-state proof, and unchanged selector/default cardinality.

A debug one-sample smoke ran both selectors in warm and explicit disk-root
`cold-verified` modes. Both cold samples were eligible, began with zero
resident/dirty/writeback bytes, and observed an 81,920-byte process
`read_bytes` delta. Warm samples retained the pinned source hash; cold samples
retained the aligned-source hash; all four samples produced the same semantic
SHA-256,
`020fdd140d2959ea4f480676a3d4d0bf840927e25251cb6cad37a043ab80627e`.
The smoke ran from the final commit with pre-existing unrelated worktree edits,
so it is functional validation only and is not retained claim evidence.

## Follow-up

Any XLSX filesystem performance claim still requires clean release binaries,
fixed CPU placement, A1/B1/B2/A2 ordering, sufficient warmups and retained
samples, drift checks, and separately retained resource evidence. Future work
should add task- and byte-bounded source-backed XLSX roots before making remote
or range-I/O claims.
