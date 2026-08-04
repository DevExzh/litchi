//! Typed SpreadsheetML custom-data properties.
//!
//! The model and bounded XML codec are split by responsibility so callers
//! can reuse the package-neutral types without taking on OPC lifecycle code.

mod codec;
mod model;

pub use codec::{parse_properties, validate_workbook_root, write_properties};
pub use model::{ExtensionList, Properties};
