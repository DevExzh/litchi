//! Lossless structural edits of worksheet cell records.
//!
//! The owner inventories all ordinary scalar cell families from `[MS-XLSB]`
//! sections 2.4.319-330 plus formula cached-result records from sections
//! 2.4.684-687. Existing values retain their BIFF12 record family, while
//! explicit insert/remove operations author or remove complete records.
//! Length-changing strings and inert formula token/cache changes rebuild only
//! selected records; unknown records remain byte-exact. Row spans and worksheet
//! dimensions expand with structural insertions, and whole-workbook publication
//! validates style, shared-string, formula, and package dependency closure.
//! Patches carry deterministic semantic deltas, exact before/after images,
//! bounded durable transfer, three-way conflict reporting, and undo/redo.

mod resources;
mod root;
pub mod workbook;
mod worksheet;

pub use root::{
    WorkbookCommit, WorkbookEdit, WorkbookHistory, WorkbookMergeConflict, WorkbookMergeOutcome,
    WorkbookPatch,
};
pub use worksheet::{
    CellError, CellFormula, Change, Commit, Edit, History, Limits, MergeConflict, MergeOutcome,
    Number, Patch, Reference, Snapshot, StoredCell, StyleIndex, TransferLimits, Value,
};

pub use crate::package::error::{Error, Result};
