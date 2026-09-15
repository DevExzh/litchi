# The synthetic XLSB fixtures used for scaling evidence

Change 0587's XLSB survey named the corpus ceiling as the blocker for every XLSB
scaling question: the largest of the fifteen `.xlsb` fixtures under `test-data/`
is `testVarious.xlsb` at 22,715 bytes, with **one** worksheet and **48** stored
cells, so a one-cell read and a full-sheet scan are 0.3% apart. This change
lifts that ceiling with a generator rather than a checked-in blob.

`tools/perf-baseline/src/bin/xlsb_synthetic_fixture.rs` (added by this change,
behind the existing `xlsb-crud` feature) builds a workbook of a caller-chosen
shape through the public `litchi_xlsb::writer` API and proves it reopens through
`litchi_xlsb::Workbook::new` before writing it. One cell in every sixteen is
left empty so the sheet stays sparse, which `xlsb_crud`'s
`sparse_iteration_without_rectangular_expansion` gate requires.

The two fixtures measured by this change were **not** checked in: they are
scratch files, reproducible byte for byte from the generator, which is checked
in. Both were deleted from the session scratchpad after the packet was assembled.

```
cargo build --release --locked --features xlsb-crud --bin xlsb_synthetic_fixture

./xlsb_synthetic_fixture --out synthetic-4x500x8.xlsb   --sheets 4 --rows  500 --columns  8
{"bytes":63635,"sha256":"001028d36186e7fa933db77a669dfb50d76d34dc6607c26d08fdb4fdbb8c2a41",
 "sheets":4,"rows":500,"columns":8,"stored_cells_sheet0":3750}

./xlsb_synthetic_fixture --out synthetic-4x2000x12.xlsb --sheets 4 --rows 2000 --columns 12
{"bytes":323710,"sha256":"05ea78ab6117e40d879f07b072d5832c48d8d87486c78fd76f8a6a1ba23caa57",
 "sheets":4,"rows":2000,"columns":12,"stored_cells_sheet0":22500}
```

| fixture | bytes | sheets | stored cells, selected sheet | total stored cells | OPC parts |
| --- | ---: | ---: | ---: | ---: | ---: |
| `testVarious.xlsb` (repository, real producer) | 22,715 | 1 | 48 | 48 | 17 |
| `cond_format.xlsb` (repository, real producer) | 8,253 | 1 | 16 | 16 | 8 |
| `synthetic-4x500x8.xlsb` | 63,635 | 4 | 3,750 | 15,000 | 13 |
| `synthetic-4x2000x12.xlsb` | 323,710 | 4 | 22,500 | 90,000 | 13 |

## What the synthetic fixtures are not

They are **producer-free**: `WorkbookWriter` emits no pivot cache, structured
table, chart sheet, drawing, connection, external link or VBA project, and only
one shared string table's worth of styles. That is exactly why the eager
`Workbook::from_opc_package_with_external_link_limits` is cheap on them and
expensive on `testVarious.xlsb`, and why this change's saving is large on the
real producer's file and small on the synthetic ones. They answer "does the
commit path scale with cell count", not "what does a large real workbook cost".
A large real-producer `.xlsb` remains absent from this corpus.
