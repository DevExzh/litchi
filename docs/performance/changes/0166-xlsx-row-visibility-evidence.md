# Change 0166: XLSX row-visibility lifecycle evidence

Date: 2026-08-17

Status: opt-in correctness and phase evidence only. No latency, speedup,
allocation/RSS, physical-I/O, cold-cache, decompression, or real-producer
claim is accepted. This change is limited to `tools/perf-baseline`; it does
not change production crates and does not add iWork coverage.

## Scope

The harness adds four selectors:

- `xlsx_eager_row_visibility_edit_save`
- `xlsx_source_backed_row_visibility_edit_save`
- `xlsx_eager_row_visibility_batch_edit_save`
- `xlsx_source_backed_row_visibility_batch_edit_save`

They use a deterministic one-sheet, media-rich corpus in `medium` and `large`
shapes. The shapes contain respectively 512 × 16 and 2,048 × 32 numeric cells,
plus eight untouched 512-KiB media entries. The scalar pair hides one existing
visible row. The batch pair unhides exactly 256 existing rows, initially
hidden. The four selectors are opt-in, so the default 36-case/198-record
matrix is unchanged and the selectable matrix advances from 311 to 315 names.

## Timing and evidence boundary

Each retained record reports separate open, plan/stage, commit, sequential
publication, and lifecycle vectors plus a semantic SHA-256. The lifecycle
vector reopens the untimed expected artifact; exact length and SHA-256 bind it
to the measured zero-retention publication. Source-backed records add logical
owned-source `ReadAt` counters, selected/unselected worksheet read attribution,
source-version checks, and pre-publication cache diagnostics. Those counters
are logical ingress evidence, not physical-I/O or allocation measurements; the
cache fields have no post-publication scope. Publication uses a fixed windowed
hashing sink retaining zero output bytes.

## Gates

Common untimed checks require each eager/source-backed output to match its own
expected length/SHA-256, semantic row-state parity across implementations, and
raw identity of untouched ZIP members. The source-backed transaction alone
checks exact no-op, foreign and stale package/source refusal, signed, protected,
formula, markup-compatibility, macro, and unsupported-relationship refusal,
plus partial-sink and zero-output refusal behavior; those fields are omitted
from eager records. These gates are
fail-closed evidence for the narrow generated corpus; they do not certify
general worksheet, row-structural, formula, extension, or third-party-producer
editing.

## Reproduction

```sh
cargo run --release --locked --manifest-path tools/perf-baseline/Cargo.toml -- \
  --warmup 3 --samples 15 \
  --case xlsx_eager_row_visibility_edit_save,\
xlsx_source_backed_row_visibility_edit_save,\
xlsx_eager_row_visibility_batch_edit_save,\
xlsx_source_backed_row_visibility_batch_edit_save \
  --xlsx-row-visibility-shape medium,large \
  --json target/perf/xlsx-row-visibility-0166.json
```

The focused harness test preserves the default matrix, validates all four
selectors, and checks the complete gate summary. A release ABBA comparison and
any broader resource or producer claim would require a separate change.
