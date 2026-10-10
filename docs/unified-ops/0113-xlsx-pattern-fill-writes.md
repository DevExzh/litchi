# 0113 — XLSX solid and pattern-fill writes

Unified XLSX cell and range writes now accept `fill`, `fillPattern`, and
`fillBg`, with `bgcolor` and `bg` foreground aliases. The native
`WorksheetEdit::fill` stages a tri-state `FillChange` using the existing
SpreadsheetML pattern and color types. The facade plans each target against
its native effective format, including the base style for missing cells.
Ordinary readback's base-style omissions remain unchanged.

A foreground assignment preserves an existing non-solid pattern and its
background; otherwise it authors a solid fill. A non-null pattern assignment
replaces the pattern and clears unspecified colors. Null pattern assignment
removes only the authored pattern attribute. A background assignment requires
an existing non-solid pattern or an explicit pattern. Supplied foreground and
background values are both honored or rejected atomically. Pattern spellings
accept case, spaces and hyphens, including OfficeCLI's gray aliases.

The facade accepts RGB and spreadsheet theme colors with tint. Native edits
also support indexed, automatic and default colors. Gradient authoring and
conversion of existing gradients remain explicit property-level capability
gaps with follow-up work; this checkpoint does not complete XLSX formatting
or the full goal.

The native stylesheet transaction reuses fill records, clones only affected
cell formats and preserves unrelated attributes, opaque children and other
formatting. Missing fill collections are created in schema order after
fonts. Virtual default records are materialized only for effective changes,
so clearing an absent pattern remains an exact no-op even with zero fill
budget. Property indexing avoids a nested scan of fill records. Source and
generated record counts and retained bytes are bounded.

Initial checks exposed missing imports, an unused source-owner field and a
fill serialization error; these were corrected.
Review also corrected missing-collection ordering and eager virtual-default
allocation. Native regressions now pass: 75 stylesheet codec and 50 cell-edit
tests. They cover source preservation, native reparse, exact no-ops,
collection creation, limits, joins, replay and inverse.

Reproduction (offline):

```sh
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib raw::styles_edit
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --lib workbook::edit::tests::cells
CARGO_NET_OFFLINE=true cargo test -p litchi-xlsx --doc fill
CARGO_NET_OFFLINE=true cargo test -p litchi-ops --no-default-features --features docx,xlsx,xlsb,encryption
```

The rewritten base remains an ancestor; no rebase replay is needed against
that current base. The full goal and performance acceptance remain open.

Both capability configurations passed seven checks each. The XLSX minimal
API glossary profile passed, and the regenerated OfficeCLI delta passed its
check and all 14 generator tests. This glossary check covers only that
profile; it is not a full-workspace conformance claim.

Facade fixture corrections distinguish an authored `patternType="none"` from
an absent pattern attribute. The first null assignment removes that token;
repeating the clear must preserve exact bytes. Ordered batch review confirmed
that fill setters flush preceding edits before planning against native styles.

All 995 combined DOCX/XLSX/XLSB/encryption tests and doctests passed, including
60 facade cell-write tests. The native fill example passed. Scoped formatting
and whitespace checks passed; the 470-path artifact audit found no oversized
evidence or compiled binaries. Both mandatory staged-file checks printed
nothing.

Warning-denied all-target Clippy passed. Staging failed because shared Git
metadata is read-only: Git could not create
`/home/zhuhe/code/litchi/.git/worktrees/unified-ops-api/index.lock`. No commit
was created; verified changes remain in the working tree.

The borrowed Clippy fixture, generated root Cargo.lock, Python cache,
temporary logs and approximately 20 GiB of build outputs were removed. No
temporary worktree was created. The full goal remains active.
