//! Shared X14 Custom Data Properties model and bounded XML codec.
//!
//! OPC graph ownership stays in the concrete XLSX/XLSB crates.  This module
//! validates only the delegated `datastoreItem` XML vocabulary and retains
//! exact extension/source bytes for those owners.

pub mod codec;
mod model;

pub use codec::{
    Limits, canonical_extension, parse_properties, rewrite_extension_list, rewrite_id,
    validate_source_properties, validate_workbook_root, write_properties,
};
pub use model::{CustomData, CustomDataView, ExtensionList, Properties, RemovalDisposition};

/// Custom Data Properties part content type.
pub const PROPERTIES_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.customDataProperties+xml";
/// Custom Data payload part content type.
pub const DATA_CONTENT_TYPE: &str = "application/binary";
/// Workbook-to-properties relationship type.
pub const PROPERTIES_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customDataProps";
/// Properties-to-payload relationship type.
pub const DATA_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customData";
