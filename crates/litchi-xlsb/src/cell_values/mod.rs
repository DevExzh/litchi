//! Lossless, length-stable edits of existing worksheet cell fields.
//!
//! The owner inventories all ordinary scalar cell families from `[MS-XLSB]`
//! sections 2.4.319-330 plus formula cached-result records from sections
//! 2.4.684-687. Edits retain the existing record family and byte length: RK
//! values must remain exactly representable, strings retain their UTF-16 code-
//! unit count, and formulas themselves are never changed or evaluated. Cell
//! creation, rich-string rewrites, and shared-string-table mutation remain out
//! of scope. Unknown records and all fields outside the selected value/style
//! bytes remain exact.

pub mod workbook;
mod worksheet;

pub use worksheet::{
    CellError, Commit, Edit, Limits, Number, Patch, Reference, Snapshot, StoredCell, StyleIndex,
    Value,
};

pub use crate::package::error::{Error, Result};
