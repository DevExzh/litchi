//! Spreadsheet abstraction traits and value types shared across formats.
pub mod traits;
pub mod types;

pub use traits::{Cell, CellIterator, RowIterator, WorkbookTrait, Worksheet, WorksheetIterator};
pub use types::{CellValue, Result};
