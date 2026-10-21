//! Layered `SpreadsheetML` OLE-object ownership.
//!
//! Semantic anchors and inert payloads live in the model module, worksheet XML in
//! the codec module, and OPC relationship/part ownership in the package module.
//!
//! [`crate::Package::edit_ole_objects`] starts an [`Edit`] of one worksheet's
//! OLE metadata; [`Edit::commit`] returns a [`Commit`] holding the new
//! package, the exact [`Patch`] and the [`Diagnostics`], and
//! [`crate::Package::apply_ole_objects`] replays a patch. The source package
//! never changes.

mod codec;
mod edit;
mod model;
mod package;
mod patch;
mod snapshot;
pub mod validation;

pub(super) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(super) fn limit(name: &str) -> Error {
    use model::{
        MAX_DEPTH, MAX_NODES, MAX_OBJECTS, MAX_PAYLOAD_BYTES, MAX_STRING_BYTES,
        MAX_TOTAL_PAYLOAD_BYTES, MAX_XML_BYTES,
    };

    let (kind, maximum) = if name.contains("object") {
        (LimitKind::Objects, MAX_OBJECTS)
    } else if name.contains("total payload") {
        (LimitKind::TotalPayloadBytes, MAX_TOTAL_PAYLOAD_BYTES)
    } else if name.contains("payload") {
        (LimitKind::PayloadBytes, MAX_PAYLOAD_BYTES)
    } else if name.contains("depth") {
        (LimitKind::Depth, MAX_DEPTH)
    } else if name.contains("structure") || name.contains("node") {
        (LimitKind::Nodes, MAX_NODES)
    } else if name.contains("string") {
        (LimitKind::StringBytes, MAX_STRING_BYTES)
    } else {
        (LimitKind::XmlBytes, MAX_XML_BYTES)
    };
    Error::limit(kind, maximum.saturating_add(1), maximum)
}

pub(super) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

pub use codec::{parse_ole_objects, write_ole_objects};
pub use edit::Edit;
#[cfg(test)]
pub(crate) use edit::Transaction;
pub use model::{
    Aspect, OleObject, OleObjectAnchor, OleObjectConformance, OleObjectMarker, OleObjectProperties,
    OleObjectRelationshipKind, OleObjectResource, OleObjectTarget, OleObjectUpdate, OleObjects,
};
pub use package::load_ole_objects;
#[cfg(test)]
pub(crate) use package::store_ole_objects;
pub use patch::{Commit, Diagnostics, Error, LimitKind, Patch};
pub use snapshot::Snapshot;
pub use validation::{graph as validate_graph, objects as validate};

#[cfg(test)]
mod tests;
