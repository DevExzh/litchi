# Change 0684 candidate-only query-index budget probe

This small standalone package exercises the candidate's public
`SourceBackedLimits::with_max_query_index_bytes` control. It is intentionally
separate from the before/after probe because the budget setter does not exist
at commit `5805d54a1`. It must be built from the candidate checkout only.

The probe opens one immutable in-memory counted `ReadAt`, applies one logical
index budget, and performs five repeated `cell_value_by_index` calls on the
same worksheet and coordinate. It reports the owner-open and per-query
`read_calls`, `read_bytes`, `version_calls`, and `len_calls`, along with the
actual compact value or exact displayed source error. Every returned value is
compared with the first result using full `CellValue` equality; missing values
and errors remain distinct. Returned values stay alive until the result is
assembled. `elapsed_ns` is a diagnostic only because the wrapper counts every
source call; logical source counters are the primary budget comparison.

Build from the candidate worktree, serially with the root coordinator's Cargo
lane:

```sh
CARGO_TARGET_DIR=/home/zhuhe/code/litchi-target-0684-after \
  cargo build --manifest-path docs/performance/results/change-0684/budget-probe/Cargo.toml \
  --release --locked --offline
```

Run all four requested budgets against a stored and a missing coordinate. The
source-backed owner is candidate-only, so use the candidate checkout's fixture
path or an identical staged corpus:

```sh
for budget in 0 1 1048576 2097152; do
  /home/zhuhe/code/litchi-target-0684-after/release/xls-index-budget-probe-0684 \
    --input test-data/poi/test-data/spreadsheet/54016.xls \
    --worksheet 0 --row 0 --column 0 --budget "$budget" --queries 5 \
    > "budget-54016-stored-${budget}.json"
  /home/zhuhe/code/litchi-target-0684-after/release/xls-index-budget-probe-0684 \
    --input test-data/poi/test-data/spreadsheet/54016.xls \
    --worksheet 0 --row 0 --column 108 --budget "$budget" --queries 5 \
    > "budget-54016-missing-${budget}.json"
done
```

The `54016.xls` missing case above is worksheet 0 `(row 0, column 108)`.
For the other requested missing controls, use `WithCustomViews.xls` worksheet
0 (`Plan1`) `(row 0, column 12)` and `Simple.xls` worksheet 0 `(row 0,
column 1)`. Repeat the matrix for their stored coordinates using coordinates
selected by the frozen corpus report. The
budgets are logical retained-index ceilings: zero disables the optional cache,
one byte is a below-minimum fallback control, and 1 MiB/2 MiB price the
candidate's bounded admission. This probe does not claim that a budget fits a
particular worksheet; its JSON records the actual query read behavior and
whether all five semantic outcomes agree.
