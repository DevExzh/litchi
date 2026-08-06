//! Layered SpreadsheetML OLE-object ownership.
//!
//! Semantic anchors and inert payloads live in the model module, worksheet XML in
//! the codec module, and OPC relationship/part ownership in the package module.

mod codec;
mod model;
mod package;
pub mod patch;
pub mod snapshot;
pub mod transaction;
pub mod validation;

use crate::error::Error;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(name: &str) -> Error {
    invalid(format!("worksheet OLE {name} limit exceeded"))
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

pub use codec::{parse_ole_objects, write_ole_objects};
pub use model::{
    OleObject, OleObjectAnchor, OleObjectAspect, OleObjectConformance, OleObjectMarker,
    OleObjectProperties, OleObjectRelationshipKind, OleObjectResource, OleObjectTarget,
    OleObjectUpdate, OleObjects,
};
pub use package::{load_ole_objects, store_ole_objects};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::Transaction;
pub use validation::{graph as validate_graph, objects as validate};

#[cfg(test)]
mod tests;
