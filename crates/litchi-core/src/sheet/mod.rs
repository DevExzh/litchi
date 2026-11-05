//! Spreadsheet abstraction traits and value types shared across formats.
pub mod error;
pub mod traits;
pub mod types;

pub use error::{Error, LimitKind, Result};
pub use traits::{Cell, WorkbookTrait, Worksheet};
pub use types::CellValue;
