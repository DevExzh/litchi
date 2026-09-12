# 0543 post-EOF raw differential test review

This is an independent, read-only review of the live public regression test
and its module wiring. No build or Rust test command was run for this review,
and the 0543 baseline/candidate executions remain the coordinator's
responsibility.

At review time the live test source is
`crates/litchi-xlsx/tests/source_backed_cell_values/post_eof_raw_differential.rs`
with SHA-256
`58cc03176b0b5903f89713147b16d43f68539d9145a49df3c2d04920b9943721`.

## Wiring and fixture boundary

`source_backed_cell_values.rs` wires the file with an explicit `#[path]` child
module. The child uses `super::*` for the existing integration fixture
helpers, `SML`, `SHEET`, `two_sheets`, `PackageWriter`, `ArchiveReader`, and
the public editor types. The module contains one public test function and
drives both deterministic shapes; there is no duplicate module or separate
fixture owner.

`source_for_shape` starts with the already valid two-worksheet package and
replaces only Sheet1's worksheet blob with valid XML containing the final
typed-boolean fault. The workbook, content types, package relationships,
workbook relationship, and Sheet2 remain valid, so source opening and catalog
validation cannot consume the intended worksheet-level error. The malformed
value is also valid ZIP member data; the package writer has no reason to
reject the fixture while constructing it.

The earlier attempted `editor.snapshot("Sheet2")` assertion was not compatible
with the single-worksheet `snapshot` API when the source contains two sheets.
The live revision removes that call. It now uses the supported multi-sheet
no-op transaction and reads Sheet2 from the published archive, which keeps the
unaffected-owner check at the correct public boundary.

## Why the fault is post-EOF materialization

The generated stream has a complete SpreadsheetML root and valid namespace,
dimension, rows, cell references, and scalar text. The final cell has
`t="b"` and `<v>maybe</v>`. The value-only validator checks that the value text
is in an allowed scalar context but does not parse the boolean lexical value,
so it accepts the entire stream and its EOF/root condition.

In the established raw parser, the final `RawCell` is stored when its closing
`c` event is handled. The parser then consumes the closing worksheet and EOF;
only after that loop does it run the materialization loop. `parse_value` then
matches the final cell's `b` type and returns the exact
`Error::Invalid("invalid worksheet boolean 'maybe'")` error. All earlier
numeric cells materialize successfully, so the error is specifically a
post-EOF semantic/materialization result rather than an XML reader or parser
transition failure.

The 0543 candidate preserves this boundary. Its shared reader returns a
`Complete(parser.finish_parse(...))` result only after EOF. The candidate's
caller finishes the validator and forwards that owned result through
`complete_source_parse`; an `Err` from final materialization therefore remains
the historical typed raw error. The fixture contains no x14ac marker, so the
historical retry has no competing extension failure to replace the boolean
error.

## Bounds and safe prefix

The candidate constants set an 8 MiB shared-source ceiling and a 131,072-event
provisional ceiling. The ordinary worksheet parser also has a 1,000,000-event
and 256-level depth limit, while cell values are bounded at 32,767 characters.
The default MCE input/output limits are much larger. Exact byte/event counts
below are derived from the deterministic generator (including one EOF event),
not from a test run:

| Shape | Dimensions | Prefix cells | XML bytes | Events including EOF | Shared source ratio | Event ratio |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| medium | 96 × 96 | 9,215 | 246,229 | 46,278 | 2.94% | 35.3% |
| dense-sparse | 128 × 128 | 16,383 | 449,347 | 82,182 | 5.36% | 62.7% |

Both streams are valid UTF-8, have maximum depth five, and keep numeric values
at or below 16,383. Their final dimension references (`CR96` and `DX128`) are
inside the SpreadsheetML grid. They contain no MCE namespace, AlternateContent,
x14ac namespace, or `dyDescent` marker, so they satisfy the candidate's plain
shared-traversal predicate. The worksheet payloads are also below the default
8 MiB source-payload cache bound and the default OPC part/input limits.

The large prefix is sufficient to put the boolean at the end of a substantial
valid stream while remaining well below the candidate's provisional event cap;
it does not accidentally exercise the cap fallback. The valid prefix uses
ordinary numeric cells with monotonically increasing row-major addresses, no
styles, formulas, metadata, shared strings, merges, or extension state, so no
earlier parser/materialization error competes with `maybe`.

## Error, source, and no-op assertions

`assert_post_eof_error` first requires the public `Error::Invalid` variant and
then compares the complete message with
`invalid worksheet boolean 'maybe'`. Each shape retries the same editor twice,
checking the original archive bytes and the `VersionedSource` identity/revision
after every refusal. This exercises both repeatability and the absence of
source mutation.

After the failed Sheet1 transactions, the test edits Sheet2 through the
multi-sheet API. It requires an unchanged one-sheet commit, an empty patch,
and the original A1 value (`20`). It then publishes that no-op multi-commit and
requires byte-for-byte equality with the original archive. The live assertion
also reads Sheet2 from the published archive and compares it with the original
worksheet XML. These checks establish that the failed selected transaction did
not expose a partial snapshot, stage a patch, or pollute later publication.

## Review disposition

The module wiring, valid-prefix construction, post-EOF ownership claim, cap
margin, exact typed error, source-version checks, and unaffected-sheet no-op
publication checks are all semantically aligned. The prior multi-sheet
`snapshot` API mismatch is resolved in the live source. No further material
coverage issue was found. Final status remains pending the coordinator's
serial baseline/candidate build and differential test receipts.
