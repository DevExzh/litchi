# Change 0189: XLSX edit-composition evidence

Date: 2026-08-18

## Scope

This change adds four opt-in benchmark-harness selectors for the existing
owned `litchi_xlsx::Workbook` transaction surface:

- `xlsx_join_disjoint_commit_save`;
- `xlsx_join_conflict_plan`;
- `xlsx_three_way_disjoint_commit_save`;
- `xlsx_three_way_conflict_resolve_save`.

They close a previously unrepresented CRUD-matrix boundary: composing
independently prepared edits, refusing overlap without last-writer-wins, and
resolving a conservative three-way conflict explicitly. No production format,
container, patch, or save implementation changed.

## Corpus and timing boundary

The selectors reuse the deterministic four-sheet media-rich XLSX cell-CRUD
corpora:

| Shape | Source SHA-256 | Logical cells |
|---|---|---:|
| `medium` | `dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036` | 9,216 |
| `dense-sparse` | `893ad3f5dd6a98aec44bc541a140048072c84c579b4b9e332431f779b097cb1a` | 17,792 |

Each pair of one-cell branches is prepared concurrently from clones of one
immutable workbook lineage before timing. The measured phase vectors are:

- `join_ns` for disjoint join or typed overlap refusal;
- `plan_ns`, `resolve_ns`, and `finish_ns` where applicable;
- `commit_ns` for successful composition;
- `publication_ns` for sequential `Workbook::write_to`;
- `reopen_ns`, recorded separately and excluded from the composed elapsed
  interval.

Publication uses an output-retaining sink with an explicit output ceiling and
a 64 KiB maximum write-call size. It is not a fixed-memory or zero-retention
writer claim.

## Untimed correctness gates

The focused selector test and every smoke record verify the applicable subset
of these contracts:

- concurrent branch construction retains the exact shared source lineage;
- independently reopened equal bytes are refused as a different lineage by
  both join and three-way planning;
- overlapping join returns a recoverable structured conflict without changing
  the accepted branch;
- disjoint three-way work is automatic, an unresolved conflict cannot finish,
  and `Left`, `Right`, and `Neither` are explicit and semantically checked;
- source bytes remain immutable and an exact empty edit remains byte-exact;
- committed output is deterministic, reopens, retains every expected cell,
  and preserves the untouched media/package inventory;
- deterministic durable JSON is parsed and applied, its inverse restores the
  exact source, and stale/foreign sources are refused;
- save-bearing cases publish the expected byte count and digest without a
  write larger than 64 KiB;
- the refusal-only selector omits output, durable, reopen, and sink claims
  instead of marking them vacuously successful.

The one-sample debug correctness smoke produced eight records: four selectors
across both corpus shapes. Output digests for disjoint join and disjoint
three-way are identical within each shape, while explicit left-conflict
resolution has its own deterministic digest. The smoke is a correctness gate,
not statistical performance evidence.

## Verification

```text
CARGO_INCREMENTAL=0 cargo check --locked \
  --manifest-path tools/perf-baseline/Cargo.toml

CARGO_INCREMENTAL=0 RUSTFLAGS='-D warnings -D deprecated' \
  cargo clippy --locked --manifest-path tools/perf-baseline/Cargo.toml \
  --all-targets

CARGO_INCREMENTAL=0 cargo test --locked \
  --manifest-path tools/perf-baseline/Cargo.toml \
  xlsx_edit_composition_selectors_are_opt_in_and_gate_complete -- --nocapture

CARGO_INCREMENTAL=0 cargo test --locked \
  --manifest-path tools/perf-baseline/Cargo.toml \
  selectable_case_count_matches_current_enumeration -- --nocapture
```

The opt-in smoke used `--warmup 0 --samples 1` with all four selectors and
`--xlsx-cell-crud-shape medium,dense-sparse`. The selectable matrix is now 336
cases; the default remains 36 cases and 198 records.

## Claims and remaining gaps

This tranche is correctness and phase evidence only. It makes no latency,
throughput, allocation, RSS, cache, physical-I/O, cold-filesystem, source-backed,
or atomic-filesystem-save claim. It does not compare join with three-way as a
performance control. Representative release ABBA evidence, larger branch
effect sets, merge-limit boundary records, allocation/resource accounting, and
other formats' patch composition remain future work.
