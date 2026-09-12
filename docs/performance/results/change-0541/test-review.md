# 0541 combined first-error test review

This is an independent review of the public source-backed XLSX error-order
tests. The review is limited to test scope and fixture semantics; no build or
Rust test command was run here.

The reviewed attempt 4 source at
`crates/litchi-xlsx/tests/source_backed_cell_values/planning_error_order.rs`
was inspected read-only. Its SHA-256 is
`05cf40a51c537bde5ff6c6c4d8e68bcf2ef2a791a050952d586844894a34bb50`.
Its focused validation was confirmed as six passing test functions. Full
quality was pending at review time; the root follow-up below records the final
formatting-only revision and completed checks.

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

The proposed `raw_and_validator_cases` now has 19 independent worksheet rows:
the original 18 plus an exact numeric materialization row. It provides useful
public-path coverage for:

- raw worksheet decoding and parser semantics: row/cell-reference mismatch,
  style lexical decoding, boolean decoding, formula syntax, and an unknown
  cell type that reaches scalar-cell materialization;
- raw OOXML preprocessing: an MCE namespace marker plus a processing
  instruction, with a typed `Error::MarkupCompatibility` assertion; the
  separate MCE test contributes three additional cases;
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

The proposed revision also adds two valid-first/faulty-second worksheet cases
and a two-fault workbook-order case. These cover reaching a bad second selected
worksheet and the fact that `resolve_selectors` processes selected positions
in workbook order, even when the caller supplies selectors in reverse order.

## Coverage additions and pending verification

1. **The MCE preprocessing pair is now present; retain its three-way shape.**
   `full_validation_precedes_mce_preprocessing_and_raw_parsing_errors` has an
   independent marker-plus-PI case, a marker before a raw cell error whose PI
   appears after that raw cell in byte order, and a marker/PI plus a raw error
   and late `<mergeCells/>`. The first two should
   return typed `Error::MarkupCompatibility` with
   `mce::Error::NonConformant("DTD and processing instructions are rejected")`;
   the last must return the validator's exact `Error::Invalid` message. This
   proves validation wins before preprocessing, which in turn precedes raw
   parsing. The PI's position after the raw cell bytes does not change the
   result: preprocessing runs before raw parser/materialization, so its MCE
   error still wins. The final focused run reports all six test functions
   passing, including these three MCE cases; full quality validation remains
   pending.

   A PI without the MCE URI is a valid control and should remain in the
   namespace/PI acceptance test. Adding the URI to that control would change
   its expected result because the raw preprocessor deliberately rejects it.

2. **The independent numeric materialization row is now present.** The
   proposed revision asserts the standalone `<v>not-a-number</v>` case as
   exact `Error::Invalid("invalid worksheet number 'not-a-number'")` and adds
   the same fault to the raw-before-late-validator cross-product. The existing
   boolean and unknown-cell rows cover other decode/materialization classes; a
   date case remains optional. The focused validation passed; full quality
   validation remains pending.

3. **The second selected-sheet cases are now present.** The proposed
   `later_selected_worksheet_failure_does_not_publish_the_first_snapshot`
   test puts valid Sheet1 before raw-faulty or validator-faulty Sheet2 and
   verifies recovery of Sheet1 after each failure. The proposed
   `first_error_across_selected_worksheets_follows_workbook_order` test also
   checks two faulty sheets with reversed caller selectors and independently
   reaches Sheet2's validator error. This closes the prior traversal gap;
   focused validation passed, while full quality validation remains pending.

4. **Combined rows are now traceable to independent owners.** For each named
   combined pair, retain an independent raw/preprocessing fixture and an
   independent validator fixture with the same fault. The two existing order
   cases correctly check a raw fault before and after a late attribute error,
   and the MCE test now establishes the preprocessing owner independently and
   in combination; the proposed numeric row does the same for numeric
   materialization. This avoids a matrix that reports the expected first error
   without proving that both competing errors are independently reachable.

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

The matrix need not duplicate every row in both positions, but it retains at
least one parser/materialization-before-validator and one validator-before-
parser/materialization ordering, plus the preprocessing pair. The second-sheet
cases use raw and validator faults from the same families.

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

The proposed revision addresses the prior coverage gaps: 19 independent
worksheet rows, three separate MCE cases, standalone and combined numeric
materialization, valid-first/faulty-second traversal, and reversed-selector
workbook-order checks. The focused result is confirmed as six passing test
functions; the review is favorable subject to root's full quality receipts.
No further semantic coverage findings remain. These additions directly test
validation-first versus preprocessing/parser/materialization ownership while
respecting the fact that OPC package validation precedes all worksheet-level
checks.

## Root formatting follow-up

Attempt 5 differs from the reviewed attempt 4 only by the five match-arm
commas required by the repository rustfmt configuration. The exact retained
source comparison confirms that no fixture or assertion changed. Its focused
run also passes all six tests. Final source SHA-256: `715a4d8f97a0ca03fc6c212cfe4d177805ef70c4e9e0fb6ec16cde44c1b080db`.

Root final validation is complete: attempt 5 passes six focused tests, 1,298
full-suite executions and all six quality checks. Exact command/source bindings
are retained in its receipts; the earlier pending notes describe review time.
