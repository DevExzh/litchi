//! Spreadsheet abstraction traits and value types shared across formats.
pub mod traits;
pub mod types;

pub use traits::{Cell, WorkbookTrait, Worksheet};
pub use types::{CellValue, Result};
