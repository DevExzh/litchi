//! Shared X14 Custom Data Properties model and bounded XML codec.
//!
//! OPC graph ownership stays in the concrete XLSX/XLSB crates.  This module
//! validates only the delegated `datastoreItem` XML vocabulary and retains
//! exact extension/source bytes for those owners.

pub mod codec;
mod model;

pub use codec::{
    Limits, canonical_extension, canonical_extension_with_limits, parse_properties,
    parse_properties_with_limits, rewrite_extension_list, rewrite_extension_list_with_limits,
    rewrite_id, rewrite_id_with_limits, validate_source_properties,
    validate_source_properties_with_limits, validate_workbook_root,
    validate_workbook_root_with_limits, write_properties, write_properties_with_limits,
};
pub use model::{CustomData, CustomDataView, ExtensionList, Properties, RemovalDisposition, Store};

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
