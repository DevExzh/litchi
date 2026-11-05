//! XLSB Spreadsheet Data Model support.
//!
//! The workbook binary stream owns typed Data Model metadata records. The
//! relationship-free OPC model part is exposed as a bounded, inert payload
//! handle and is inspected lazily. No DAX, refresh, external source access, or
//! calculation engine is invoked by this module.

mod codec;
mod model;
mod package;
mod patch;
mod proof;
mod snapshot;
mod transaction;
mod workbook;

pub use codec::{
    HARD_MAX_CONNECTION_BYTES, HARD_MAX_CONNECTIONS, HARD_MAX_GRAPH_PARTS,
    HARD_MAX_GRAPH_RELATIONSHIPS, HARD_MAX_METADATA_BYTES, HARD_MAX_PART_BYTES, SemanticLimits,
};
pub use litchi_core::edit::Diagnostics;
pub use litchi_xldm::{Xldm140ColumnBinding, Xldm140TimeGroupingBinding};
pub use model::TimeGroupingContentType as ContentType;
pub use model::{
    Definition, Model, ModelPart, Relationship, Table, TimeGrouping, TimeGroupingColumn,
    TimeGroupingContentType,
};
pub use package::{DATA_MODEL_CONTENT_TYPE, DATA_MODEL_PART_NAME};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::{Edit, EditOptions};

pub(crate) use package::{
    validate_definition_connections, validate_model_for_write, validate_model_payload_and_groupings,
};

/// Alias matching the concise limits naming used by other XLSB feature modules.
pub use crate::package::error::{Error, Result};

/// A bounded dimension of this family; an exceeded bound is reported as
/// [`crate::package::PackageError::Limit`] with this kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes of the Data Model payload part.
    PayloadBytes,
    /// Bytes of the workbook Data Model records.
    MetadataBytes,
    /// Connections of the Data Model closure.
    Connections,
    /// Bytes of the External Data Connections part of the closure.
    ConnectionsBytes,
    /// Parts of the Data Model closure.
    GraphParts,
    /// Relationships of the Data Model closure.
    GraphRelationships,
    /// Bytes of the captured Data Model source closure.
    SourceBytes,
    /// Bytes of the workbook part.
    WorkbookBytes,
    /// Bytes of one Data Model table name.
    TableNameBytes,
    /// A bounded resource of the MS-XLDM model itself.
    Model(litchi_xldm::LimitKind),
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::PayloadBytes => "Data Model payload bytes",
            Self::MetadataBytes => "Data Model workbook record bytes",
            Self::Connections => "Data Model connections",
            Self::ConnectionsBytes => "Data Model connections bytes",
            Self::GraphParts => "Data Model graph parts",
            Self::GraphRelationships => "Data Model graph relationships",
            Self::SourceBytes => "Data Model source bytes",
            Self::WorkbookBytes => "Data Model workbook bytes",
            Self::TableNameBytes => "Data Model table name bytes",
            Self::Model(kind) => kind.as_str(),
        })
    }
}

pub(crate) fn serialize_workbook_records(definition: &Definition) -> Result<Vec<u8>> {
    codec::serialize_new_workbook(definition)
}
