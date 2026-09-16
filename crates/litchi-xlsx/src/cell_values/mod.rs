//! Source-backed scalar-cell and formula edits for selected worksheet cells.
//!
//! This deliberately narrow capability changes or removes scalar cells already
//! stored in selected worksheets, or inserts one bounded numeric cell at an
//! absent coordinate. Clearing retains the `<c>` owner and its local style;
//! removal deletes that complete owner while retaining its row and the
//! producer's conservative dimension. Cacheless scalar formulas and direct
//! dates are supported; every effective mutation invalidates workbook
//! calculation properties and removes a captured calculation chain atomically.
//! It creates only unstyled numeric cell records and never creates styles,
//! shared strings, or formulas. Workbooks outside the statically provable
//! closure are refused.
//!
//! # What this module admits
//!
//! Change 0657 replaced the element, attribute and relationship allow-lists
//! with the property they were a proxy for: a value-only rewrite composes the
//! worksheet's `<dimension ref>` attribute and its `<sheetData>` span, and
//! copies every other byte of the worksheet and the workbook from the source.
//! An element, attribute or relationship the editor does not model is
//! therefore *unfamiliar*, and an unfamiliar thing is copied through
//! unchanged rather than refused: compatibility markers, `<sheetPr>`,
//! `<mergeCells>`, `<conditionalFormatting>`, `<dataValidations>`,
//! `<hyperlinks>`, `<pageSetup>`, `<extLst>` and their vendor payloads,
//! `docProps`, a worksheet's printer settings or drawing, a workbook's
//! defined names or external links.
//!
//! What stays refused is what the edit's own meaning depends on:
//!
//! | construct | verdict |
//! | --- | --- |
//! | a shared-string part | refused: the value of a `t="s"` cell lives there |
//! | a pivot cache | refused: it holds a copy of the source cells |
//! | a table or query table | refused: a column is named after its header cell |
//! | cell metadata (`cm`, `vm`) | refused: it describes the value being replaced |
//! | an external relationship | refused: no closure over a target outside the package |
//! | a relationship reference inside `<sheetData>` | refused: a removed cell would orphan it |
//! | a protected sheet | the edit is refused |
//! | a data validation covering the address | the edit is refused |
//! | a shared or array formula covering the address | the edit is refused |
//! | a merged range covering the address | the edit is refused unless it anchors the range |
//! | the calculation chain | dropped with its relationship, as before |
//!
//! Every other construct is admitted, and an admitted edit changes the edited
//! `<c>` record, the `<dimension ref>` attribute, the workbook's calculation
//! properties and the calculation chain — and nothing else.

#[cfg(test)]
#[path = "admission_tests.rs"]
mod admission_tests;
#[cfg(test)]
#[path = "facts_oracle_tests.rs"]
mod facts_oracle_tests;
mod patch;
mod snapshot;
mod source;
mod validation;

pub use patch::Patch;
pub use patch::{Commit, Diagnostics, MultiCommit, MultiDiagnostics, MultiPatch};
pub use snapshot::{MAX_MULTI_WORKSHEET_BYTES, MAX_SHEET_OWNERS, MultiSnapshot, Snapshot};
pub use source::{
    CellValueEdit, MAX_BATCH_EDITS, MultiSourceEdit, SheetCellValueEdit, SourceBackedEditor,
    SourceEdit,
};
