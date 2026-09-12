# 0543 post-EOF raw-error differential test

This test is the missing public source-editor oracle called for by the 0542
next-candidate review. It is intentionally a behavior-only integration test:
the same module must pass against the restored baseline and against the
unmeasured follow-up after the full candidate is applied. It does not inspect
private traversal types or make a performance claim.

The test wires
`crates/litchi-xlsx/tests/source_backed_cell_values/post_eof_raw_differential.rs`
into the existing `source_backed_cell_values.rs` integration-test owner. For
each of two deterministic worksheet shapes it replaces Sheet1 in the existing
two-sheet fixture with a plain, eligible worksheet:

| shape | dimensions | valid cells before the final fault |
| --- | ---: | ---: |
| medium | 96 × 96 | 9,215 |
| dense-sparse | 128 × 128 | 16,383 |

Every prefix cell is a valid numeric scalar. The final cell has `t="b"` and
`<v>maybe</v>`. The XML stream reaches its closing worksheet element, so the
validator accepts EOF before raw materialization reports the exact typed
`Error::Invalid("invalid worksheet boolean 'maybe'")` result. The source has
no MCE, x14ac, or other extension marker, keeping the case on the plain
post-EOF path described by `next-candidate.patch`.

The public contract is checked through `SourceBackedEditor::edit_sheets`:

1. The same editor is tried twice for Sheet1. Both attempts must return the
   same `Invalid` variant and exact message.
2. After each refusal, the original archive bytes and `VersionedSource`
   identity/revision are unchanged.
3. Sheet2 remains independently readable with its exact original worksheet
   XML and A1 value. Its later empty transaction has no changes and an empty
   patch.
4. Publishing that unaffected no-op commit produces the original archive
   byte-for-byte. This proves that a refused selected transaction did not
   leave a partial snapshot or staged publication state.

The test should be run unchanged against both source manifests before any
candidate retention. Existing 0541 early-parser and late-validator cases
remain necessary because this module specifically covers `Complete(Err)` after
accepted EOF; it does not replace those immediate-fallback cases. No build or
test was run while preparing this patch; the root coordinator owns serial
baseline/candidate builds and execution.
