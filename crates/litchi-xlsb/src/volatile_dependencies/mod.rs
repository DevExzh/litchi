//! Inert XLSB Volatile Dependencies support ([MS-XLSB] 2.1.7.60).
//!
//! The owner reads the Workbook implicit part lazily at the package boundary,
//! keeps cached RTD/cube values inert, and never refreshes, evaluates, or
//! contacts an external service. Unsupported source records remain safe for
//! exact no-ops and owner removal; typed rewrites refuse them.

mod codec;
mod model;
mod package;
mod patch;
mod snapshot;
mod transaction;
mod workbook;

#[cfg(test)]
mod tests;

#[cfg(test)]
mod codec_review_tests;

pub use litchi_core::edit::Diagnostics;
pub use model::{
    CachedValue, CellReference, Dependencies, DependencyKind, ErrorCode, HARD_MAX_ITEMS,
    HARD_MAX_PART_BYTES, HARD_MAX_RECORDS, HARD_MAX_RELATIONSHIPS, HARD_MAX_STRING_UNITS,
    HARD_MAX_TOTAL_STRING_UNITS, MainTopic, SemanticLimits, Topic, VolatileType,
};
pub use package::{CONTENT_TYPE, DEFAULT_PART_NAME, RELATIONSHIP_TYPE};
pub use patch::{Commit, Patch};
pub use snapshot::Snapshot;
pub use transaction::{Edit, EditOptions};

/// Alias matching other XLSB feature owners.
pub use crate::package::error::{Error, Result};

/// A bounded dimension of this family; an exceeded bound is reported as
/// [`crate::package::PackageError::Limit`] with this kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes of the Volatile Dependencies part.
    PartBytes,
    /// Records of the Volatile Dependencies part.
    Records,
    /// Type collections of the owner.
    TypeCollections,
    /// Main collections of the owner.
    MainCollections,
    /// Topic collections of the owner.
    TopicCollections,
    /// Subtopic records of the owner.
    SubtopicRecords,
    /// Cell references of the owner.
    CellReferences,
    /// UTF-16 code units of one string.
    StringUnits,
    /// UTF-16 code units of every string.
    TotalStringUnits,
    /// Relationships of the Volatile Dependencies closure.
    Relationships,
    /// Bytes of the encoded owner.
    OutputBytes,
    /// Bytes of the workbook part read while binding sheet ordinals.
    WorkbookBytes,
    /// Records of the workbook part read while binding sheet ordinals.
    WorkbookRecords,
    /// Bytes of the captured lexical source.
    SourceBytes,
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(match self {
            Self::PartBytes => "Volatile Dependencies part bytes",
            Self::Records => "Volatile Dependencies records",
            Self::TypeCollections => "Volatile Dependencies type collections",
            Self::MainCollections => "Volatile Dependencies main collections",
            Self::TopicCollections => "Volatile Dependencies topic collections",
            Self::SubtopicRecords => "Volatile Dependencies subtopic records",
            Self::CellReferences => "Volatile Dependencies cell references",
            Self::StringUnits => "Volatile Dependencies string units",
            Self::TotalStringUnits => "Volatile Dependencies total string units",
            Self::Relationships => "Volatile Dependencies relationships",
            Self::OutputBytes => "Volatile Dependencies output bytes",
            Self::WorkbookBytes => "Volatile Dependencies workbook bytes",
            Self::WorkbookRecords => "Volatile Dependencies workbook records",
            Self::SourceBytes => "Volatile Dependencies source bytes",
        })
    }
}

/// Validate graph ownership and the modeled stream before catalog remapping.
pub(crate) fn reorder_owner(
    package: &litchi_opc::OpcPackage,
    sheet_count: usize,
) -> Result<Option<litchi_opc::PackURI>> {
    let workbook_name = package.main_document_part()?.partname();
    let graph = package::discover_graph(package, workbook_name, SemanticLimits::DEFAULT)?;
    let Some(graph) = graph else {
        return Ok(None);
    };
    let parsed = codec::read(
        package.get_part(&graph.part_name)?.blob(),
        SemanticLimits::DEFAULT,
        sheet_count,
    )?;
    if !parsed.rewrite_safe {
        return Err(Error::refused(
            litchi_core::edit::RefusalKind::UnsupportedContent,
            "edit_volatile_dependencies",
            "the Volatile Dependencies source contains records this edit cannot preserve",
        ));
    }
    Ok(Some(graph.part_name))
}
