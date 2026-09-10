# Independent XLSB drawing-skipped projection review

Date: 2026-09-10

Scope: the v3 drawing-skipped projection batch, current primary integration,
and its workbook publication, durable patch, no-op, and cross-workbook transfer
boundaries. This review made no source edits.

## Verdict

**Cleared after the scalar-transfer fix.** The current public projection,
durable-patch, no-op, and transfer boundaries are explicit and coherent.

## Resolved finding: scalar transfer now keeps the donor opaque

The original implementation stored only `workbook_bytes(source)` in its
`TransferCell` operation and reconstructed those bytes using the target
drawing policy. A skipped source with an opaque or malformed
`SpreadsheetDrawing` therefore failed before scalar cell/style/formula
transfer into a valid eager target:

```text
source_drawings_parsed=false target_drawings_parsed=true
result=Err(InvalidFormat("drawing part root is not xdr:wsDr"))
```

`transfer_cell` now reconstructs the donor explicitly with
`DrawingLoadPolicy::Skipped` (`crates/litchi-xlsb/src/cell_values/root.rs:1370-1386`),
while the publication reparse continues to use the target's policy. This
matches the documented cell/style/shared-string/formula dependency closure and
keeps donor drawings outside the scalar transfer operation. The new regression
covers a malformed skipped donor transferred into an eager target.

## Reviewed and clear

- `Workbook::drawings_parsed()` distinguishes an eager empty inventory from a
  skipped inventory, so `sheet_drawings().is_empty()` is no longer ambiguous.
- `WorkbookPatch::from_bytes_without_drawing_parse*` makes the absent policy
  field in the durable wire format an explicit caller choice. Default
  `from_bytes*` remains eager and rejects malformed drawing images; the
  skipped decoder accepts them and target application retains the target
  policy.
- Workbook-owned reparse paths for cell values, structure replay/authoring,
  rename, calculation-chain, sparklines, cell watches, generic `edit_opc`,
  and data/XML-map publication carry the existing drawing policy. Package-only
  facades intentionally remain eager because they receive only an
  `OpcPackage`.
- The inverse/no-op shortcut is correctly guarded by both
  `operations.is_empty()` and `before == after`; an inverse has a different
  endpoint and still validates/applies it.

The focused `drawing_skip_projection` integration tests pass (4/4), and an
independent probe of the malformed skipped-donor to eager-target case now
returns success. The parent receipt reports the complete library/harness,
strict lint, rustdoc, and format gates passing.
