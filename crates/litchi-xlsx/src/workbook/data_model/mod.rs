//! MS-XLDM workbook Data Model owner.
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
pub use package::{Commit, Patch, Snapshot, Transaction, load_data_model, store_data_model};

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

pub(crate) fn invalid(message: impl Into<String>) -> crate::Error {
    crate::Error::Invalid(message.into())
}

pub(crate) fn limit(name: &str) -> crate::Error {
    invalid(format!("Data Model {name} limit exceeded"))
}

pub(crate) fn xml_error(error: impl std::fmt::Display) -> crate::Error {
    crate::Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}
