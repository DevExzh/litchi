# Change 0127: ODS source-backed repeated-cell sweep evidence

Status: harnessed for matched query and source-read evidence; no release
performance acceptance claim.

This change adds two opt-in selectors:

```text
ods_file_eager_cell_sweep
ods_file_source_cell_sweep
```

Both selectors use the existing deterministic media-rich ODS corpus:
two 32 by 32 worksheets, eight deterministic incompressible 2 MiB
`Pictures/` members, generator `litchi-ods-media-publication-v1`, archive
size 16,790,689 bytes, and archive SHA-256
`46b7f61cb74639115f6d120dc6498b97d6b310d51c78c4fb85ac60d6fc758b14`.
The corpus builder and package bytes are unchanged.

## Timed scope

Each owner is opened before its timer: eager uses
`litchi_ods::Spreadsheet::from_bytes`, and source-backed uses
`litchi_ods::SourceBackedSpreadsheet::from_path`. The timed interval performs
four identical row-major sweeps over all 2,048 logical coordinates, including
the adaptive locator's threshold transition. Every query is black-boxed.
Digesting, stored-cell counting, source identity checks, archive/member/media
verification, and the source replay are outside the timer.

## Source replay and gates

For every measured source sample, an independent instrumented positional
source is fully opened first. Counters are reset after preparation; the same
four sweeps then observe zero source reads, zero source bytes, and zero
previous-range overlap. Preparation counters are retained and required to be
deterministic across samples. Eager evidence marks the source replay as
`not_applicable`.

The two controls must produce the same 8,192 stored-cell count and semantic
cell digest. Each sample also checks the unchanged source file byte-for-byte
and by SHA-256, member topology, complete semantic grid, manifest media types,
and every retained media payload through the existing ODS media archive gate.

The selectors remain opt-in; the default matrix stays at 36 cases / 198
records. This is query correctness and logical source-read evidence only. It
does not claim latency improvement, physical I/O, decompressed bytes,
allocations, RSS, cold-cache behavior, or release/ABBA acceptance. Any such
claim requires a clean matched release ABBA run with the required resource
and preservation evidence.
