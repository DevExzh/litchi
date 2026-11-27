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
    DependencyBytes => "Data Model dependency XML bytes",
    RewriteBytes => "Data Model rewrite bytes",
    ExtensionBytes => "Data Model extension bytes",
    NamespaceBytes => "Data Model namespace bytes",
    NamespaceBindings => "Data Model namespace bindings",
    PayloadBytes => "Data Model payload bytes",
    StringBytes => "Data Model string bytes",
    TotalStringBytes => "Data Model total string bytes",
    Nodes => "Data Model XML nodes",
    Depth => "Data Model XML depth",
    Events => "Data Model XML events",
    FormulaBytes => "Data Model formula bytes",
    Tables => "Data Model tables",
    Relationships => "Data Model relationships",
    TimeGroupings => "Data Model time groupings",
    CalculatedTimeColumns => "Data Model calculated time columns",
    XldmSourceBytes => "XLDM source bytes",
    XldmOutputBytes => "XLDM output bytes",
    XldmPathTextBytes => "XLDM path text bytes",
    XldmDirectoryBytes => "XLDM directory bytes",
    XldmFiles => "XLDM files",
    XldmPartitions => "XLDM partitions",
    XldmPathBytes => "XLDM path bytes",
    XldmXmlNodes => "XLDM XML nodes",
    XldmXmlDepth => "XLDM XML depth",
    XldmXmlTextBytes => "XLDM XML text bytes",
    XldmMetadataBytes => "XLDM metadata bytes",
    XldmMetadataNodes => "XLDM metadata XML nodes",
    XldmMetadataDepth => "XLDM metadata XML depth",
    XldmMetadataTextBytes => "XLDM metadata text bytes",
    XldmMetadataFiles => "XLDM metadata files",
    XldmProofItems => "XLDM OLAP proof items",
    XldmProofStringBytes => "XLDM OLAP proof string bytes",
    XldmProofSourceBytes => "XLDM OLAP proof source bytes",
    XldmProofWork => "XLDM OLAP proof graph work",
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
pub(crate) const MAX_DEPENDENCY_BYTES: usize = 256 * 1024 * 1024;
pub(crate) const MAX_FORMULA_EVENTS: usize = 10_000_000;
pub(crate) const MAX_REWRITE_BYTES: usize = 32 * 1024 * 1024;
pub(crate) const MAX_EXTENSION_BYTES: usize = 4 * 1024 * 1024;
pub(crate) const MAX_PAYLOAD_BYTES: usize = 512 * 1024 * 1024;
pub(crate) const MAX_STRING_BYTES: usize = 1024 * 1024;
pub(crate) const MAX_TOTAL_STRING_BYTES: usize = 16 * 1024 * 1024;
pub(crate) const MAX_NODES: usize = 200_000;
pub(crate) const MAX_DEPTH: usize = 128;
pub(crate) const MAX_NAMESPACE_BINDINGS: usize = 4096;
pub(crate) const MAX_TABLES: usize = 65_536;
pub(crate) const MAX_RELATIONSHIPS: usize = 65_536;
pub(crate) const MAX_TIME_GROUPINGS: usize = 65_536;
pub(crate) const MAX_CALCULATED_TIME_COLUMNS: usize = 65_536;

/// Resource limits used by one Data Model edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    max_dependency_bytes: usize,
    max_dependency_depth: usize,
    max_formula_events: usize,
    max_formula_bytes: usize,
}

impl Limits {
    /// Construct the default Data Model edit limits.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            max_dependency_bytes: MAX_DEPENDENCY_BYTES,
            max_dependency_depth: MAX_DEPTH,
            max_formula_events: MAX_FORMULA_EVENTS,
            max_formula_bytes: MAX_STRING_BYTES,
        }
    }

    /// Maximum aggregate XML bytes inspected while removing a Data Model.
    #[must_use]
    pub const fn max_dependency_bytes(self) -> usize {
        self.max_dependency_bytes
    }

    /// Maximum XML depth inspected while removing a Data Model.
    #[must_use]
    pub const fn max_dependency_depth(self) -> usize {
        self.max_dependency_depth
    }

    /// Maximum XML events inspected while checking formula dependencies.
    #[must_use]
    pub const fn max_formula_events(self) -> usize {
        self.max_formula_events
    }

    /// Maximum bytes retained for one formula while checking dependencies.
    #[must_use]
    pub const fn max_formula_bytes(self) -> usize {
        self.max_formula_bytes
    }

    /// Set the aggregate dependency XML byte limit.
    pub fn with_max_dependency_bytes(mut self, value: usize) -> crate::error::Result<Self> {
        self.max_dependency_bytes =
            checked_limit(LimitKind::DependencyBytes, value, MAX_DEPENDENCY_BYTES)?;
        Ok(self)
    }

    /// Set the dependency XML depth limit.
    pub fn with_max_dependency_depth(mut self, value: usize) -> crate::error::Result<Self> {
        self.max_dependency_depth = checked_limit(LimitKind::Depth, value, MAX_DEPTH)?;
        Ok(self)
    }

    /// Set the formula XML event limit.
    pub fn with_max_formula_events(mut self, value: usize) -> crate::error::Result<Self> {
        self.max_formula_events = checked_limit(LimitKind::Events, value, MAX_FORMULA_EVENTS)?;
        Ok(self)
    }

    /// Set the formula byte limit.
    pub fn with_max_formula_bytes(mut self, value: usize) -> crate::error::Result<Self> {
        self.max_formula_bytes = checked_limit(LimitKind::FormulaBytes, value, MAX_STRING_BYTES)?;
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::new()
    }
}

fn checked_limit(kind: LimitKind, value: usize, maximum: usize) -> crate::error::Result<usize> {
    if value == 0 || value > maximum {
        return Err(Error::InvalidLimit {
            kind: kind.into(),
            value: u64::try_from(value).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        });
    }
    Ok(value)
}

/// Options for one Data Model edit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct EditOptions {
    limits: Option<Limits>,
}

impl EditOptions {
    /// Bound Data Model dependency checks by `limits`.
    #[must_use]
    pub const fn limits(mut self, limits: Limits) -> Self {
        self.limits = Some(limits);
        self
    }

    /// Return the configured limits, if an override was supplied.
    #[must_use]
    pub const fn limits_value(self) -> Option<Limits> {
        self.limits
    }
}

pub(crate) fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

pub(crate) fn limit(kind: LimitKind, observed: usize, maximum: usize) -> Error {
    Error::limit(kind, observed, maximum)
}

pub(crate) fn limit_lower_bound(kind: LimitKind, maximum: usize) -> Error {
    Error::limit(kind, maximum.checked_add(1).unwrap_or(usize::MAX), maximum)
}

pub(crate) fn xml_error(error: impl std::fmt::Display) -> Error {
    Error::Xml(litchi_ooxml_common::XmlError::Malformed(error.to_string()))
}

#[cfg(test)]
mod tests {
    use super::{LimitKind, MAX_DEPENDENCY_BYTES, MAX_FORMULA_EVENTS, limit};
    use crate::Error;

    #[test]
    fn removal_budgets_have_typed_kinds() {
        assert!(matches!(
            limit(
                LimitKind::DependencyBytes,
                MAX_DEPENDENCY_BYTES + 1,
                MAX_DEPENDENCY_BYTES
            ),
            Error::Limit {
                kind: crate::LimitKind::DataModel(LimitKind::DependencyBytes),
                ..
            }
        ));
        assert!(matches!(
            limit(
                LimitKind::Events,
                MAX_FORMULA_EVENTS + 1,
                MAX_FORMULA_EVENTS,
            ),
            Error::Limit {
                kind: crate::LimitKind::DataModel(LimitKind::Events),
                ..
            }
        ));
        assert!(matches!(
            limit(LimitKind::FormulaBytes, 2, 1),
            Error::Limit {
                kind: crate::LimitKind::DataModel(LimitKind::FormulaBytes),
                ..
            }
        ));
    }
}
