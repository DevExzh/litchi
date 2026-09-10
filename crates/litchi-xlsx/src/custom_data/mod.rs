//! Typed `SpreadsheetML` custom-data properties.
//!
//! The model and bounded XML codec are split by responsibility so callers
//! can reuse the package-neutral types without taking on OPC lifecycle code.

mod codec;
mod model;
mod package;

pub use codec::{parse_properties, validate_workbook_root, write_properties};
pub use model::{CustomData, CustomDataView, ExtensionList, Properties, RemovalDisposition, Store};
pub use package::{Commit, Limits, Part, Patch, Snapshot, Transaction};

/// Custom Data Properties part content type.
pub const PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customDataProperties+xml";
/// Custom Data payload content type.
pub const DATA_CONTENT_TYPE: &str = "application/binary";
/// Workbook-to-properties relationship type.
pub const PROPERTIES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps";
/// Properties-to-payload relationship type.
pub const DATA_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData";
