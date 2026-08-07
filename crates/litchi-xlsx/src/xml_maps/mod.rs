//! Layered SpreadsheetML Custom XML Maps ownership.
//!
//! Typed contextual values live in [`model`], bounded XML conversion lives in
//! [`codec`], workbook relationship traversal lives in [`package`], and focused
//! regression coverage lives in [`tests`]. Schema and binding payloads remain
//! inert, bounded XML rather than being resolved or executed.

mod codec;
mod model;
mod package;
pub mod patch;
pub mod snapshot;
pub mod transaction;

#[cfg(test)]
mod tests;

pub use codec::{parse_xml_map_info, serialize_xml_map_info};
pub use codec::{
    parse_xml_map_info as parse_map_info, serialize_xml_map_info as serialize_map_info,
};
pub use model::{
    DataBinding, XmlMap, XmlMapConformance, XmlMapDataBinding, XmlMapInfo, XmlMapLimits,
    XmlMapSchema, XmlSchema,
};
pub use package::{
    load_from_package, load_from_package_with_conformance, remove_from_package, store_in_package,
};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::Transaction;

fn invalid(message: impl Into<String>) -> Box<dyn std::error::Error + Send + Sync> {
    std::io::Error::new(std::io::ErrorKind::InvalidData, message.into()).into()
}
