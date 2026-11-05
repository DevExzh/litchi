//! Inert XLSB Calculation Chain ownership ([MS-XLSB] 2.1.7.4).
//!
//! The local XLSB specification defines the Calculation Chain as an optional
//! Workbook-owned performance cache, but does not provide its binary record
//! grammar. This owner therefore validates only the OPC graph and bounded
//! opaque bytes. Generic BIFF12 framing is available only as an explicit
//! diagnostic and is not an admission requirement. The payload remains
//! opaque and is never evaluated, recalculated, refreshed, or rewritten.

mod model;
mod package;
mod patch;
mod snapshot;
mod transaction;
mod workbook;

#[cfg(test)]
mod tests;

pub use litchi_core::edit::Diagnostics;
pub use model::{
    ChainPart, HARD_MAX_GRAPH_PARTS, HARD_MAX_PART_BYTES, HARD_MAX_RECORDS, HARD_MAX_RELATIONSHIPS,
    SemanticLimits,
};
pub use package::{CONTENT_TYPE, RELATIONSHIP_TYPE};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::{Edit, EditOptions};

/// Alias matching the concise limits naming used by other XLSB feature owners.
pub use crate::package::error::{Error, Result};

/// A bounded dimension of this family; an exceeded bound is reported as
/// [`crate::package::PackageError::Limit`] with this kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes of the Calculation Chain part.
    PartBytes,
    /// Records of the Calculation Chain part.
    Records,
    /// Relationships of the Calculation Chain closure.
    Relationships,
    /// Parts of the Calculation Chain closure.
    GraphParts,
    /// Bytes of the workbook part read while binding the chain.
    WorkbookBytes,
    /// Bytes of the captured Calculation Chain source closure.
    SourceBytes,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::PartBytes => "Calculation Chain part bytes",
            Self::Records => "Calculation Chain records",
            Self::Relationships => "Calculation Chain relationships",
            Self::GraphParts => "Calculation Chain graph parts",
            Self::WorkbookBytes => "Calculation Chain workbook bytes",
            Self::SourceBytes => "Calculation Chain source closure bytes",
        })
    }
}
