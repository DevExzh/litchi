//! Layered `SpreadsheetML` external data connection owner.
//!
//! Typed connection models live in the model module, bounded XML conversion in the codec
//! module, and workbook relationship integration in the package module. External targets
//! remain inert data and are never dereferenced.
//!
//! [`crate::Package::edit_connections`] starts an [`Edit`] of the workbook's
//! connection catalog; [`Edit::commit`] returns a [`Commit`] holding the new
//! package, the exact [`Patch`] and the [`Diagnostics`], and
//! [`crate::Package::apply_connections`] replays a patch. The source package
//! never changes.

mod codec;
pub(crate) mod embedded_data;
mod model;
mod namespace;
mod package;

/// Typed and package-level validation for inert connection metadata.
pub mod validation {
    use litchi_opc::OpcPackage;

    use super::model::Connections;
    use super::package;
    use crate::error::Result;

    pub fn connections(value: &Connections) -> Result<()> {
        value.to_xml(false).map(|_| ())
    }

    pub fn graph(package: &OpcPackage) -> Result<Option<Connections>> {
        package::validate_graph(package)
    }
}

#[cfg(test)]
mod tests;

pub use model::*;
use model::{
    MAX_CONNECTIONS, MAX_DOM_DEPTH, MAX_DOM_NODES, MAX_PARAMETERS, MAX_STRING_BYTES,
    MAX_TEXT_FIELDS, MAX_WEB_TABLES, MAX_XML_BYTES,
};
#[cfg(test)]
pub(crate) use package::Transaction;
pub use package::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Snapshot, load_from_package, validate_graph,
};

fn invalid(v: impl Into<String>) -> Error {
    let value = v.into();
    if value.contains("limit") || value.contains("exceed") {
        let (maximum, name) = if value.contains("depth") {
            (MAX_DOM_DEPTH, "connections XML depth")
        } else if value.contains("node") || value.contains("resource") {
            (MAX_DOM_NODES, "connections XML nodes")
        } else if value.contains("event") {
            (MAX_DOM_NODES.saturating_mul(4), "connections XML events")
        } else if value.contains("attribute") {
            (MAX_DOM_NODES, "connections XML attributes")
        } else if value.contains("namespace") {
            (MAX_XML_BYTES, "connections XML namespace bytes")
        } else if value.contains("parameter") {
            (MAX_PARAMETERS, "connection parameters")
        } else if value.contains("web table") {
            (MAX_WEB_TABLES, "connection web tables")
        } else if value.contains("text field") {
            (MAX_TEXT_FIELDS, "connection text fields")
        } else if value.contains("string") {
            (MAX_STRING_BYTES, "connections XML string bytes")
        } else if value.contains("output") || value.contains("serialized") {
            (MAX_XML_BYTES, "connections output bytes")
        } else if value.contains("connection") {
            (MAX_CONNECTIONS, "connections")
        } else {
            (MAX_XML_BYTES, "connections XML bytes")
        };
        return limit(name, maximum.saturating_add(1), maximum);
    }
    crate::error::invalid(value)
}

pub(crate) fn limit(name: &str, observed: usize, maximum: usize) -> Error {
    let kind = if name.contains("depth") {
        LimitKind::Depth
    } else if name.contains("node") {
        LimitKind::Nodes
    } else if name.contains("event") {
        LimitKind::Events
    } else if name.contains("attribute") {
        LimitKind::Attributes
    } else if name.contains("namespace") {
        LimitKind::NamespaceBytes
    } else if name.contains("parameter") {
        LimitKind::Parameters
    } else if name.contains("web table") {
        LimitKind::WebTables
    } else if name.contains("text field") {
        LimitKind::TextFields
    } else if name.contains("relationship XML") {
        LimitKind::RelationshipBytes
    } else if name.contains("relationship") {
        LimitKind::RelationshipCount
    } else if name.contains("temporary") {
        LimitKind::TemporaryBytes
    } else if name.contains("output") {
        LimitKind::OutputBytes
    } else if name.contains("extension") {
        LimitKind::ExtensionBytes
    } else if name.contains("string") {
        LimitKind::StringBytes
    } else if name.contains("connection") {
        if name.contains("XML") || name.contains("xml") {
            LimitKind::XmlBytes
        } else {
            LimitKind::Connections
        }
    } else {
        LimitKind::XmlBytes
    };
    Error::limit(kind, observed, maximum)
}
