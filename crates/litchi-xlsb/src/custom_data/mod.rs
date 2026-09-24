//! Source-bound XLSB Custom Data and Custom Data Properties lifecycle.
//!
//! The package adapter owns the OPC graph and BIFF12 connection closure.  The
//! X14 values and XML codec are shared with the OOXML spreadsheet hosts, while
//! `bindings` is a narrow XLSB-only source lens for `BrtBeginExtConn14`.

mod model;
mod package;

pub(crate) mod bindings;

pub use litchi_ooxml_common::custom_data::{
    CustomData, CustomDataView, DATA_CONTENT_TYPE, DATA_RELATIONSHIP_TYPE, ExtensionList,
    PROPERTIES_CONTENT_TYPE, PROPERTIES_RELATIONSHIP_TYPE, Properties, RemovalDisposition, Store,
};
pub use litchi_ooxml_common::custom_data::{
    parse_properties, validate_workbook_root, write_properties,
};
pub use model::{Part, Storage, StorageId, StorageSelector, StorageSelectorInput};
pub use package::{Commit, Limits, Patch, Snapshot, Transaction};

/// Result type for XLSB Custom Data operations.
pub type Result<T> = crate::package::error::Result<T>;

/// Validate one delegated `ST_Xstring` UID without exposing the XML codec in
/// ordinary callers.
pub(crate) fn validate_id(value: &str) -> Result<()> {
    litchi_ooxml_common::custom_data::codec::validate_source_properties(&Properties {
        id: value.to_owned(),
        extension_list: None,
    })
    .map_err(Into::into)
}
