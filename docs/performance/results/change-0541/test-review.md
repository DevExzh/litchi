# 0541 combined first-error test review

This is an independent review of the public source-backed XLSX error-order
tests. The review is limited to test scope and fixture semantics; no build or
Rust test command was run here.

## Public path and error ownership

The intended path is the right one: construct a `SourceBackedEditor` with
`from_read_at`, then call `edit_sheets` with the selected sheet names. Opening
the source-backed package validates the ZIP/OPC catalog, workbook content
types, and relationship graph. Worksheet payload validation is deferred until
`edit_sheets`, where the selected worksheet is passed through the value-only
validator before raw worksheet preprocessing, raw parsing, style checks, and
scalar-cell validation.

That ordering makes the fixture boundary important. The workbook XML, package
relationships, workbook relationship, worksheet content type, and selected
sheet relationship must all be valid. Only the worksheet part's blob should
contain the injected fault. A malformed `.rels`, bad content type, duplicate
package owner, or invalid workbook relationship produces an OPC/package error
before the worksheet validator is reached and therefore tests a different
boundary.

## What the current matrix covers

The 18 independent rows provide useful public-path coverage for:

- raw worksheet decoding and parser semantics: row/cell-reference mismatch,
  style lexical decoding, boolean decoding, formula syntax, and an unknown
  cell type that reaches scalar-cell materialization;
- raw OOXML preprocessing: an MCE namespace marker plus a processing
  instruction, with a typed `Error::MarkupCompatibility` assertion;
- value-only validation: cell metadata, unsupported elements, mixed
  SpreadsheetML dialects, unqualified and qualified attributes, duplicate and
  malformed attributes, incomplete/mismatched markup, DTD refusal, and a
  reference outside a scalar value; and
- public typed errors: exact `Error::Invalid` messages where stable, bounded
  `Error::Invalid` containment checks for parser diagnostics, and `Error::Xml`
  for an unsupported entity.

The retry and source-byte assertions are also appropriate. They demonstrate
that an error is returned through the public editor operation and that the
read-only source remains usable after the failed attempt.

## Gaps to close before treating the matrix as complete

1. **The MCE preprocessing pair is now present; retain its three-way shape.**
   `full_validation_precedes_mce_preprocessing_and_raw_parsing_errors` has an
   independent marker-plus-PI case, a marker/PI before a raw cell error, and a
   marker/PI plus a raw error and late `<mergeCells/>`. The first two should
   return typed `Error::MarkupCompatibility` with
   `mce::Error::NonConformant("DTD and processing instructions are rejected")`;
   the last must return the validator's exact `Error::Invalid` message. This
   proves validation wins before preprocessing, which in turn precedes raw
   parsing. A focused run reported all three MCE cases passing.

   A PI without the MCE URI is a valid control and should remain in the
   namespace/PI acceptance test. Adding the URI to that control would change
   its expected result because the raw preprocessor deliberately rejects it.

2. **Add an independent numeric materialization row.** The combined
   raw-before-validator case uses `not-a-number`, but it does not independently
   assert the raw error. Add a standalone `<v>not-a-number</v>` case with the
   exact `Error::Invalid("invalid worksheet number 'not-a-number'")` result.
   This makes the pair auditable instead of relying on the combined test to
   establish the raw side. The existing boolean and unknown-cell rows cover
   other decode/materialization classes; a date case is optional.

3. **Exercise the second selected sheet.** The matrix still puts every fault
   in Sheet1 and makes Sheet2 valid. `assert_error_case` now retries Sheet1
   and then successfully edits Sheet2 alone, which is a useful failed-state
   isolation check, but it does not prove that the multi-sheet traversal
   reaches a faulty second worksheet. Add the inverse arrangement: valid Sheet1,
   faulty Sheet2, and `edit_sheets(["Sheet1", "Sheet2"])`. Assert the same
   typed error and unchanged source on the first attempt and on retry. This
   confirms that the selected-sheet loop reaches the second worksheet after
   successfully materializing the first, and that no partial result is
   published. A two-invalid-sheet case is useful only if its expected first
   sheet is documented as workbook/selection order; it is not a substitute for
   the valid-first/invalid-second case.

4. **Make combined rows traceable to independent owners.** For each named
   combined pair, retain an independent raw/preprocessing fixture and an
   independent validator fixture with the same fault. The two existing order
   cases correctly check a raw fault before and after a late attribute error,
   and the MCE test now establishes the preprocessing owner independently and
   in combination. The numeric raw fault should still have its own exact
   standalone assertion. This avoids a matrix that reports the expected first
   error without proving that both competing errors are independently
   reachable.

X14ac extension failures are not a good missing public pair here. The
value-only validator rejects the qualified extension attributes that the raw
X14ac capture recognizes, so a failing X14ac capture is generally unreachable
through this public value-only operation. Existing raw-module tests can own
that boundary; public tests should not label a validator rejection as an
X14ac parser failure.

## Recommended first-error pairs

The following pairs keep the worksheet/package boundary valid and identify the
owner of each error. The combined worksheet can place the raw fault before or
after the validator fault to test both source orders; the expected result is
the validator error in either arrangement.

| Raw or preprocessing fault | Independent validator fault | Combined assertion |
| --- | --- | --- |
| `A2` in row `1`, or invalid numeric `<v>` | unknown `future` attribute on `c` | exact validator `Error::Invalid` |
| invalid boolean lexical value | late `<mergeCells>` | exact validator `Error::Invalid` |
| formula with a leading `=` | mixed SpreadsheetML dialect | exact validator `Error::Invalid` |
| MCE-bearing processing instruction | late unsupported/dependency element | typed validator `Error::Invalid`, never MCE preprocessing |
| unknown cell type reaching scalar validation | value-only metadata attribute | exact validator `Error::Invalid` |

The matrix need not duplicate every row in both positions, but it should retain
at least one parser/materialization-before-validator and one validator-before-
parser/materialization ordering, plus the preprocessing pair. The second-sheet
case should use one of these same pairs.

## Assertion and implementation pitfalls

- Match the public `Error` variant first. `Error::Xml` contains an
  `XmlError`; `Error::MarkupCompatibility` contains an MCE error. Exact string
  comparisons are suitable for the stable validator messages. XML/parser
  diagnostics that include offsets or library wording should use a stable
  containment check while still requiring the correct typed variant.
- `edit_sheets` selectors in the existing integration helpers are built with
  `"Sheet1".into()`/`"Sheet2".into()`. Keep that explicit form if inference
  rejects an array of bare `&str` values.
- The `VersionedSource` is held as an `Arc` whose concrete type is coerced to
  the `ReadAt` trait object expected by the editor. If `Arc::clone` cannot infer
  that coercion, clone the concrete `Arc` and let the argument coercion happen,
  or cast explicitly.
- The validator performs its own end-name check, so the mismatched closing-tag
  row should expect its value-only validation message. If a focused run shows
  a different library diagnostic, preserve the `Error::Invalid` owner and use
  a stable substring rather than changing the fixture to an OPC error.
- A publication-oriented `OpcPackage`/`PackageWriter` may reject malformed
  worksheet XML while constructing the fixture. That would move the failure
  into fixture setup and prevent the public worksheet path from running. The
  current helper correctly starts from the valid `two_sheets()` archive and
  replaces only worksheet members with `StreamingArchiveWriter`; retain this
  lower-level construction for duplicate or otherwise malformed XML rows.

## Review disposition

The current test direction is sound and the existing independent rows are
valuable. The MCE preprocessing pair and failed-state isolation check are now
in place. Before calling the combined matrix complete, add an independently
asserted numeric raw case and the valid-first/invalid-second selected-sheet
case. These additions directly test validation-first versus
preprocessing/parser/materialization ownership while respecting the fact that
OPC package validation precedes all worksheet-level checks.
