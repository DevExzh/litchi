//! The module holds the vocabulary the workbook formats share:
//!
//! - The generic sheet traits: [`WorkbookTrait`](crate::sheet::WorkbookTrait),
//!   [`Worksheet`](crate::sheet::Worksheet) and [`Cell`](crate::sheet::Cell),
//!   over [`CellValue`](crate::sheet::CellValue). They dispatch statically:
//!   each implementation names its worksheet, cell and iterator types as
//!   associated types, and the iterators are standard [`Iterator`]s.
//! - `text` (the `sheet-text` feature): the `litchi-sheet-text` crate's CSV,
//!   TSV, PRN, SYLK, DIF and fixed-width text workbooks.
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
//! evaluator borrow it directly; its worksheets and cells are the
//! `litchi_ops::view::sheet` types. The traits report boxed errors.
//!
//! ```
//! # #[cfg(feature = "xlsx")]
//! # {
//! use litchi::sheet::{WorkbookTrait, Worksheet as _};
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
//! The text workbooks are concrete types: open one with
//! `text::Workbook`, or read rows with `text::{delimited, sylk, dif,
//! fixed_width}::read` (the `sheet-text` feature).

/// Canonical worksheet-view semantics shared by XLSX and XLSB.
pub mod view {
    pub use litchi_sheet::view::*;
    pub use litchi_sheet::{Cell, Rect};
}

pub use litchi_core::sheet::{traits, types};

/// Formula evaluation over any [`WorkbookTrait`].
#[cfg(feature = "eval")]
pub mod eval {
    pub use litchi_eval::*;
}
/// Delimited, SYLK, DIF and fixed-width text workbooks (`litchi-sheet-text`).
#[cfg(feature = "sheet-text")]
pub mod text {
    pub use litchi_sheet_text::*;
}

#[cfg(feature = "eval")]
pub use eval::FormulaEvaluator;
pub use litchi_core::sheet::{Error, LimitKind, Result, error};
pub use traits::{Cell, WorkbookTrait, Worksheet};
pub use types::CellValue;
