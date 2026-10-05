//! The module holds the vocabulary the workbook formats share:
//!
//! - The generic sheet traits: [`WorkbookTrait`](crate::sheet::WorkbookTrait),
//!   [`Worksheet`](crate::sheet::Worksheet), [`Cell`](crate::sheet::Cell)
//!   and their iterators, over [`CellValue`](crate::sheet::CellValue).
//! - [`text`](crate::sheet::text): CSV, TSV, PRN, SYLK, DIF and fixed-width
//!   text workbooks.
//! - [`view`](crate::sheet::view): worksheet-view semantics shared by XLSX
//!   and XLSB.
//! - `FormulaEvaluator` (the `eval` feature): formula evaluation over any
//!   `WorkbookTrait`.
//!
//! # Reading workbooks
//!
//! Open an XLS, XLSX, XLSB or ODS workbook with `litchi::open`. Its workbook
//! view (`File::workbook`) lists the worksheets and implements
//! `WorkbookTrait` by static dispatch, so generic sheet code and the formula
//! evaluator borrow it directly. The traits report boxed errors.
//!
//! ```
//! # #[cfg(feature = "xlsx")]
//! # {
//! use litchi::sheet::WorkbookTrait;
//!
//! # let path = concat!(
//! #     env!("CARGO_MANIFEST_DIR"),
//! #     "/../../test-data/ooxml/xlsx/ExcelTables.xlsx"
//! # );
//! let file = litchi::open(path)?;
//! let workbook = file.workbook()?;
//! assert_eq!(workbook.worksheet_names()?, ["ExcelTable"]);
//!
//! // The same view through the generic sheet trait.
//! let sheet = WorkbookTrait::worksheet_by_index(&workbook, 0)?;
//! assert_eq!(sheet.name(), "ExcelTable");
//! # }
//! # Ok::<(), Box<dyn std::error::Error + Send + Sync>>(())
//! ```
//!
//! The `Workbook` type and the boxed `open_*_workbook` functions of this
//! module predate `litchi::File` and are removed in a later step of the
//! migration to it; the text workbooks stay.

/// Canonical worksheet-view semantics shared by XLSX and XLSB.
pub mod view {
    pub use litchi_sheet::view::*;
    pub use litchi_sheet::{Cell, Rect};
}

pub use litchi_core::sheet::{traits, types};

// Unified workbook facade implementations require a concrete workbook format.
#[cfg(feature = "eval")]
pub mod eval {
    pub use litchi_eval::*;
}
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "numbers"
))]
mod adapters;
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "numbers"
))]
pub mod functions;
#[cfg(feature = "xlsx")]
mod selection;
pub mod text;
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "numbers"
))]
pub(crate) mod workbook;
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "numbers"
))]
mod workbook_types;

// Re-exports
#[cfg(feature = "xlsx")]
pub use crate::xlsx::{Address, Area, At, Rect};
#[cfg(feature = "eval")]
pub use eval::FormulaEvaluator;
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "numbers"
))]
pub use functions::*;
#[cfg(feature = "xlsx")]
pub use selection::{SelectedCell, SelectedCellView, SelectedWorksheet, WorksheetSelector};
pub use traits::{Cell, CellIterator, RowIterator, WorkbookTrait, Worksheet, WorksheetIterator};
pub use types::{CellValue, Result};
#[cfg(any(
    feature = "xls",
    feature = "xlsx",
    feature = "xlsb",
    feature = "ods",
    feature = "numbers"
))]
pub use workbook::Workbook;
