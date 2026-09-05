# 0426 spreadsheet facade fixture corrections

This source-only correction covers the two spreadsheet facade failures carried
forward from the 0425 baseline at `f22917bbe973709e93dba9d1d4bc85a4dc0778ce`.
It follows the accepted preservation, content-derived detection, lazy-source,
facade/package-ownership, and migration-verification constraints in ADRs
0001, 0005, 0006, 0008, 0010, 0011, and 0024. No production behavior, public
API, archive policy, or performance claim is changed.

No Cargo command, build, test, profiler, or CPU workload was run. The only
verification performed here was Rustfmt 1.9.0 from toolchain 1.98.1 with
`skip_children=true` on the two owned Rust files, followed by `git diff
--check`.

## XLSX selected-read fixture

`crates/litchi/src/sheet/adapters/xlsx.rs` now uses
`corrupt_stored_payload_byte` for
`source_bytes_catalog_and_selection_defer_corrupt_unselected_payload`.
The helper walks the fixture's stored local records, locates the second
worksheet member, and flips one byte in the existing
`SECOND-WORKSHEET-PAYLOAD-MARKER` comment. It does not change the local header,
central directory, declared sizes, or either CRC field. The fixture therefore
passes the archive-wide strict layout proof and leaves the verified checksum
failure to the selected second-member read. The existing central-CRC helper is
retained for the separate filesystem CRC-deferral test.

The first selected read now asserts the concrete `First!A1` value `Int(7)`;
the second selected read retains its error assertion. This keeps the intended
deferred-selection boundary strong: catalog and first-sheet access must remain
usable, while selecting the damaged second payload must fail. The helper is a
test fixture operation over bytes already owned by the test and does not add a
facade dependency on a ZIP implementation.

## XLSB wrong-facade fixture

`crates/litchi/src/sheet/workbook.rs` now makes `minimal_xlsx_marked_opc` a
minimal valid OPC package. Its content-types catalog declares the `rels`
default and the fixture includes `_rels/.rels` with one
`officeDocument` relationship targeting `xl/workbook.xml`; the empty workbook
sheet catalog remains unchanged.

The test continues to call
`open_xlsb_workbook_dyn_with_limits` through a temporary path and requires the
typed `litchi_core::Error::NotOfficeFile`. With a valid package, content-derived
XLSX detection can finish constructing the source-backed XLSX owner, after
which the XLSB facade's existing Xlsx branch must refuse without reopening the
pathname. The correction does not infer format from the temporary filename or
broaden the assertion to an arbitrary parse error.

## Scope and follow-up

These are test-fixture/oracle corrections for baseline failures identified in
the 0425 source-only review. The legacy XLS Formula-record suffix failure is
outside this patch and remains with the format owner for separate analysis.
After the coordinator's serialized verification, the two focused tests should
be run against the corrected fixtures and the broader facade receipt should
record whether any remaining failures are baseline debt. No result is claimed
until those checks are run.
