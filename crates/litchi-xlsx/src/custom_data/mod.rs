//! Typed `SpreadsheetML` custom-data properties.
//!
//! The model and bounded XML codec are split by responsibility so callers
//! can reuse the package-neutral types without taking on OPC lifecycle code.

mod codec;
mod model;
mod package;

pub use codec::{XmlLimits, parse_properties, validate_workbook_root, write_properties};
pub use model::{CustomData, CustomDataView, ExtensionList, Properties, RemovalDisposition, Store};
pub use package::{Commit, Limits, Part, Patch, Snapshot, Transaction};

pub use litchi_ooxml_common::custom_data::{
    DATA_CONTENT_TYPE, DATA_RELATIONSHIP_TYPE, PROPERTIES_CONTENT_TYPE,
    PROPERTIES_RELATIONSHIP_TYPE,
};
