//! MS-XLDM workbook Data Model owner.
//!
//! [`crate::Package::data_model`] reads the source-bound graph;
//! [`crate::Package::edit_data_model`] starts an [`Edit`] whose
//! [`Edit::commit`] returns a [`Commit`] holding the new package, the exact
//! [`Patch`] and the [`Diagnostics`]; [`crate::Package::apply_data_model`]
//! replays a patch. The source package never changes.
//!
//! The semantic descriptor is separated from its bounded XML codec and from
//! the OPC singleton graph. The binary MS-XLDM payload stays inert and is
//! validated only by the existing [`crate::package::xldm`] inspector.

mod codec;
mod model;
mod package;
mod removal;

pub use codec::{parse_data_model, write_data_model};
pub use codec::{parse_model_time_groupings, write_model_time_groupings};
pub use model::{
    CalculatedTimeColumn, Definition, Model, ModelTimeGrouping, ModelTimeGroupingContentType,
    ModelTimeGroupings, ModelView, OpaqueXml, Payload, Relationship, Table,
};
pub use package::{Commit, Diagnostics, Edit, Error, Patch, Snapshot, load_data_model};

crate::record::limit_kinds!("Data Model", DataModel {
    XmlBytes => "Data Model XML bytes",
    RewriteBytes => "Data Model rewrite bytes",
    ExtensionBytes => "Data Model extension bytes",
    NamespaceBytes => "Data Model namespace bytes",
    PayloadBytes => "Data Model payload bytes",
    StringBytes => "Data Model string bytes",
    TotalStringBytes => "Data Model total string bytes",
    Nodes => "Data Model XML nodes",
    Depth => "Data Model XML depth",
    Tables => "Data Model tables",
    Relationships => "Data Model relationships",
    TimeGroupings => "Data Model time groupings",
    CalculatedTimeColumns => "Data Model calculated time columns",
});

/// OPC content type for an MS-XLDM payload.
pub const DATA_MODEL_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.model+data";
/// Workbook extension URI identifying the Data Model descriptor.
pub const DATA_MODEL_EXTENSION_URI: &str = "{FCE2AD5D-F65C-4FA6-A056-5C36A1767C68}";
/// Extension URI assigned by [MS-XLSX] §2.2.4.11 to `ModelTimeGroupings`.
pub const MODEL_TIME_GROUPINGS_EXTENSION_URI: &str = "{9835A34E-60A6-4A7C-AAB8-D5F71C897F49}";
/// Namespace assigned by [MS-XLSX] §2.4.71 to `modelTimeGroupings`.
pub const MODEL_TIME_GROUPINGS_NAMESPACE: &str =
    "http://schemas.microsoft.com/office/spreadsheetml/2014/11/main";
/// The only valid XLSX Data Model payload part name.
pub const DATA_MODEL_PART_NAME: &str = "/xl/model/item.data";

pub(crate) const SML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
pub(crate) const STRICT_SML: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
pub(crate) const X15: &str = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";
pub(crate) const CONNECTIONS_RELATIONSHIP_TYPE: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections";
pub(crate) const STRICT_CONNECTIONS_RELATIONSHIP_TYPE: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/connections";
pub(crate) const CONNECTIONS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml";

pub(crate) const MAX_XML_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_REWRITE_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_EXTENSION_BYTES: usize = 4 * 1024 * 1024;
pub(crate) const MAX_PAYLOAD_BYTES: usize = 512 * 1024 * 1024;
pub(crate) const MAX_STRING_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_TOTAL_STRING_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_NODES: usize = 200_000;
pub(crate) const MAX_DEPTH: usize = 128;
pub(crate) const MAX_TABLES: usize = 65_536;
pub(crate) const MAX_RELATIONSHIPS: usize = 65_536;
pub(crate) const MAX_TIME_GROUPINGS: usize = 65_536;
pub(crate) const MAX_CALCULATED_TIME_COLUMNS: usize = 65_536;

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn limit(name: &str) -> Error {
    let (kind, maximum) = if name.contains("payload") {
        (LimitKind::PayloadBytes, MAX_PAYLOAD_BYTES)
    } else if name.contains("rewrite") {
        (LimitKind::RewriteBytes, MAX_REWRITE_BYTES)
    } else if name.contains("calculatedTimeColumn") {
        (
            LimitKind::CalculatedTimeColumns,
            MAX_CALCULATED_TIME_COLUMNS,
        )
    } else if name.contains("modelTimeGrouping") || name.contains("time grouping") {
        if name.contains("count") {
            (LimitKind::TimeGroupings, MAX_TIME_GROUPINGS)
        } else {
            (LimitKind::ExtensionBytes, MAX_EXTENSION_BYTES)
        }
    } else if name.contains("table") {
        (LimitKind::Tables, MAX_TABLES)
    } else if name.contains("relationship") {
        (LimitKind::Relationships, MAX_RELATIONSHIPS)
    } else if name.contains("depth") {
        (LimitKind::Depth, MAX_DEPTH)
    } else if name.contains("structure") || name.contains("node") {
        (LimitKind::Nodes, MAX_NODES)
    } else if name.contains("string") || name.contains("attribute") {
        if name.contains("total") {
            (LimitKind::TotalStringBytes, MAX_TOTAL_STRING_BYTES)
        } else {
            (LimitKind::StringBytes, MAX_STRING_BYTES)
        }
    } else if name.contains("namespace") {
        (LimitKind::NamespaceBytes, MAX_EXTENSION_BYTES)
    } else if name.contains("extension") || name.contains("context") || name.contains("opaque") {
        (LimitKind::ExtensionBytes, MAX_EXTENSION_BYTES)
    } else {
        (LimitKind::XmlBytes, MAX_XML_BYTES)
    };
    Error::limit(kind, maximum.saturating_add(1), maximum)
}

pub(crate) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
