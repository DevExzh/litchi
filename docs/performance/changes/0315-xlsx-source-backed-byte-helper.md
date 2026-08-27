# Change 0315: XLSX source-backed byte helper

## Scope

The facade byte helpers `sheet::open_workbook_from_bytes` and
`sheet::open_workbook_from_bytes_with_limits` now use the read-only
`SourceBackedPackage` and `SourceBackedWorkbook` path used by filesystem XLSX
opens. The standalone owned `xlsx::Workbook` and explicit `xlsx::Package`
constructors remain unchanged because they provide CRUD, editing, and output
ownership.

## Contract

The helper checks `ReadLimits::max_input_bytes()` before taking ownership of a
borrowed slice. It performs a fallible `Vec` reservation and maps allocation
failure to `OpcError::Allocation`, then opens the owned vector with
`SourceBackedPackage::from_vec_with_limits`.

`SourceBackedWorkbook::from_source_backed_package` is the strict XLSX boundary.
The helper does not use multi-format detection or an eager fallback, so XLSB
and other OOXML packages remain rejected by the XLSX-only API.

Only the package catalog and required workbook metadata are read at open.
Worksheet payloads remain deferred until the selected worksheet is read. The
owned source keeps the returned trait object independent of the caller's input
slice lifetime.

## Verification focus

Focused tests cover selected-sheet reads with a corrupt unselected worksheet,
input ownership after the backing vector is dropped, non-XLSX/XLSB rejection,
and the exact input-byte and part-byte limit fields.

## Status and evidence

Implemented and validated in one isolated target with serialized Cargo
execution (`CARGO_BUILD_JOBS=1`). Evidence:

- `litchi` no-default-feature XLSX library tests: 60 passed.
- `litchi-xlsx` plus `litchi-xlsb` library tests: 63 passed.
- Strict `litchi` no-default-feature XLSX Clippy run: warnings denied and passed.
- `litchi` no-default-feature XLSX no-deps rustdoc run: warnings denied and passed.
- `rustfmt` and diff checks passed.

This note makes no RSS, OOM, throughput, or other performance claim.
