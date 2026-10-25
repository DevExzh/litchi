# 0362ul66l — Numbers tile output-policy repair

Status: source READY for the next scoped F5 retry.

The first focused F5 retry at `6417a19464c064e3fac533e9c837901c4251992e`
reached the Numbers package tests but refused the two mixed-scalar table-cell
transactions with `WireWork observed 20 maximum 12`.  The value was the
previous cell's 12-byte minimal BNC length, which had accidentally been used
as the output ceiling while planning a replacement.  A valid number, Boolean,
date, or duration replacement can be larger than that source cell.

`plan_cell_mutation` now receives the owner-authorized tile output ceiling.
Preflight therefore admits the complete candidate size before any writer
allocation, while execution still passes the exact planned size to the raw
writer and keeps its fallible exact reservation.  Input and tail ceilings
remain derived from the retained source, and no raw item or semantic
materialization limit is widened.

Validation with the approved scoped target reused the prior artifacts:

```text
cargo test -q -p litchi-numbers --all-features --no-fail-fast \
  --test table_cells --test table_data_list_reader_integration \
  --test tile_reader_integration
17 table-cell tests: passed
25 table-data-list tests: passed
2 tile-reader tests: passed
```

The temporary diagnostic instrumentation used to identify the limit was
removed before this record and commit.
