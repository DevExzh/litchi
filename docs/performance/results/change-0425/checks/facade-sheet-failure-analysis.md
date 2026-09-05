# 0425 spreadsheet facade failure analysis

This note covers the three spreadsheet failures in
[`facade-test.log`](facade-test.log.gz). The candidate and the clean locked
baseline at `340cc91ae2bdec338dfe7682b4b5d8c219a2d288` report the same test
names and the same assertion errors. The baseline receipt is
[`baseline-facade-tests.json`](baseline-facade-tests.json); the candidate
receipt is [`facade-test.json`](facade-test.json). No Cargo command, test,
build, profiler, CPU workload, or source edit was used for this diagnosis.

The affected facade, OPC, XLSX, detection, and formula-metadata paths have no
0425 source transition. The only nearby XLS change is the mechanical
`chunks_exact` to `as_chunks` representation in `crates/litchi-xls/src/formula.rs`;
the Formula-record metadata parser that reports the legacy XLS error is
unchanged. These are baseline failures and do not support a 0425 regression or
performance claim.

## Deferred XLSX selection

`sheet::adapters::xlsx::tests::source_bytes_catalog_and_selection_defer_corrupt_unselected_payload`
fails at `crates/litchi/src/sheet/adapters/xlsx.rs:864`:

```text
assertion failed: first.cell_value(0, 0).is_ok()
```

The helper at `xlsx.rs:745` does not corrupt the second worksheet payload. It
flips the central-directory CRC field for
`xl/worksheets/sheet2.xml`, while leaving the local header and stored payload
in place. The test then selects `First` before selecting `Second`.

The failure is explained by the strict source-backed read path. A selected
worksheet read enters `SourceBackedPackage` through
`litchi-opc/src/source_backed.rs` and the verified indexed reader. Its first
read builds the archive-wide strict layout proof in
`soapberry-zip/src/office.rs:3042-3120`. That proof calls
`validate_strict_entry_layout` for every physical ZIP member, not only the
selected member. The metadata validator in `soapberry-zip/src/archive.rs`
checks the local/central name, flags, method, sizes, and CRC agreement, so the
bad CRC on `sheet2.xml` is rejected while `sheet1.xml` is being selected.

This archive-wide behavior is an intentional accepted contract: change 0351
states that every physical span participates in the strict layout proof, and
change 0359 identifies that proof as archive-wide indexed state. The test is
therefore stale relative to the current verifier policy, although its
selection/deferred-payload intent remains useful.

The bounded test follow-up is to mutate a byte inside the stored second-sheet
payload while preserving its local/central framing and CRC metadata. The
strict layout proof can then succeed, the first sheet can be read, and the
second selected read can observe the payload/CRC failure. If deferred central
metadata corruption is a required contract instead, that needs a separate
reviewed verifier-policy change against the 0351 archive-wide validation
constraint; weakening the proof only to make this assertion pass would hide a
real archive error.

## XLSB API on a marked XLSX package

`sheet::workbook::source_xlsb_path_tests::xlsb_dynamic_open_rejects_known_xlsx_without_reopening_path`
fails at `crates/litchi/src/sheet/workbook.rs:1274` because the returned error
is not `Error::NotOfficeFile`.

The test fixture at `workbook.rs:1241` is not a complete OPC package. It writes
`[Content_Types].xml` with the XLSX main content type and `xl/workbook.xml`,
but omits the package-level `_rels/.rels` office-document relationship. The
source detector does recognize content-type markers without opening ordinary
part payloads (`detection_smart/ooxml.rs:215-249`), but the XLSX source facade
then calls `SourceBackedPackage::main_document_part()`. That method requires a
unique package-level office-document relationship and returns a typed invalid
relationship error when it is missing. Consequently the XLSX construction
fails before `open_xlsb_workbook_dyn_with_limits` can receive the
`WorkbookSourcePathDetection::Xlsx` branch and convert it to
`NotOfficeFile`.

The test was introduced before this 0425 batch and the same failure occurs at
the clean baseline. The next focused correction should make the fixture a
valid minimal OPC package by adding `_rels/.rels` with one office-document
relationship targeting `/xl/workbook.xml` (and retain the empty sheet catalog
if a zero-sheet workbook is accepted). Then keep the typed `NotOfficeFile`
assertion. A future focused run should format the actual error if the valid
fixture still fails, rather than broadening the assertion to any error. This
preserves the no-reopen/source-identity behavior that the test name is meant
to cover.

## Legacy XLS formula extraction

`sheet::workbook::tests::test_workbook_formulas_xls` fails at
`crates/litchi/src/sheet/workbook.rs:1890`:

```text
Failed to extract text: Parse(InvalidLength { expected: 48, found: 58 })
```

The checked-in `FormulaEvalTestData.xls` opens, but its BIFF `Formula`
record at workbook-stream offset `43084` has a 58-byte payload. The record
header fields report 26 bytes of formula tokens, so the unchanged metadata
codec computes `FORMULA_FIXED_SIZE` 22 plus 26 and rejects the record as
`expected: 48, found: 58` in
`crates/litchi-xls/src/formula_metadata/codec.rs:35-62`. Other records in the
same fixture show the same shape (for example, 62 bytes with 30 token bytes),
including a ten-byte suffix after the declared token stream.

This failure predates 0425 and is not caused by the fixed-width iterator
syntax change in `formula.rs`: the error is raised by Formula-record framing
before the changed formula token/string iteration is reached. The format
owner should identify the suffix under the BIFF producer/spec variant and
choose between explicit support and an explicit unsupported-fixture oracle.
Supporting it must retain the declared token length, payload bound, reserved
flag checks, and truncation/overlong error ordering. A generic relaxation of
`formula_end != data.len()` would accept malformed records and is not a
bounded fix.

## Disposition

All three owned spreadsheet failures are reproduced at the clean locked
baseline. They remain open facade correctness work and should not be counted
as evidence against the 0425 chunk migration. The bounded next steps are
test-fixture/oracle corrections for the XLSX and XLSB cases, followed by a
format-owner decision on the BIFF Formula suffix. No production change or
performance measurement is justified by this diagnosis.
