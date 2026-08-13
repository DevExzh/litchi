//! Selector-first exact-source row and column size transactions.

use std::fmt;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use thiserror::Error as ThisError;

use super::{Dimension, Size};
use crate::Package;

/// A content-free semantic location associated with a dimension transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// One row or column of a rooted table.
    Dimension {
        /// Zero-based rooted sheet position.
        sheet: usize,
        /// Zero-based table position within the sheet.
        table: usize,
        /// Selected row or column.
        dimension: Dimension,
    },
}

/// A finite resource governed by the focused transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source package bytes.
    InputBytes,
    /// Complete candidate package bytes.
    OutputBytes,
    /// Physical package members.
    Entries,
    /// Bytes in one physical member.
    EntryBytes,
    /// Aggregate member bytes.
    TotalEntryBytes,
    /// Physical container names and metadata.
    PackageBytes,
    /// Bytes in one decoded payload container.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages or metadata items inspected.
    PayloadItems,
    /// Native references inspected.
    PayloadReferences,
    /// Bytes inspected by a strict projection.
    WireBytes,
    /// Bytes emitted by a strict rewrite.
    WireOutputBytes,
    /// Fields inspected by a strict projection.
    WireFields,
    /// Strict projection nesting depth.
    WireNesting,
    /// Strict projection and transaction work.
    WireWork,
    /// Aggregate focused transaction work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A content-redacted row or column size transaction failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum TransactionError {
    /// No rooted sheet matched the selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No table on the selected sheet matched the selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// A name selector matched more than one rooted table.
    #[error("the selected Numbers sheet has an ambiguous table name")]
    AmbiguousSelector,
    /// The requested row or column is outside the selected table.
    #[error("the requested Numbers table dimension is out of bounds at {path:?}")]
    OutOfBounds {
        /// Selected semantic location.
        path: Path,
        /// Declared length of the selected axis.
        length: u32,
    },
    /// A changed edit targeted an effectively locked table.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked {
        /// Selected semantic location.
        path: Path,
    },
    /// The source lacks exact physical provenance for changed publication.
    #[error("this Numbers source does not support exact table-dimension editing")]
    UnsupportedSource,
    /// Rooted ownership, routing, metadata, or wire framing is invalid.
    #[error("the Numbers table-dimension source is invalid at {path:?}")]
    InvalidSource {
        /// Content-free semantic location.
        path: Path,
    },
    /// A finite resource ceiling was exceeded.
    #[error(
        "Numbers table-dimension {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
        /// Content-free semantic location.
        path: Path,
    },
    /// A bounded allocation failed before publication.
    #[error("could not allocate {amount} units for the Numbers table-dimension transaction")]
    Allocation {
        /// Requested bytes or elements.
        amount: usize,
        /// Content-free semantic location.
        path: Path,
    },
    /// Candidate reopening, reselection, or locality verification failed.
    #[error("the edited Numbers table dimension failed semantic verification")]
    Verification,
    /// The patch was applied to a package other than its exact source.
    #[error("the table-dimension patch does not match the exact source package")]
    PatchConflict,
}

/// Mutable size staged against one immutable package snapshot.
pub struct Edit<'a> {
    pub(crate) source: &'a Package,
    pub(crate) sheet_position: usize,
    pub(crate) table_position: usize,
    pub(crate) dimension: Dimension,
    pub(crate) before: Size,
    pub(crate) size: Size,
    pub(crate) evidence: Evidence,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path())
            .field("size", &self.size)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the selected semantic location.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Dimension {
            sheet: self.sheet_position,
            table: self.table_position,
            dimension: self.dimension,
        }
    }

    /// Return the size that would be published.
    #[must_use]
    pub const fn size(&self) -> Size {
        self.size
    }

    /// Replace the staged size without touching package bytes.
    #[must_use]
    pub fn set(mut self, size: Size) -> Self {
        self.size = size;
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, lock, resource, or verification error.
    pub fn commit(self) -> Result<Commit, TransactionError> {
        crate::package::table_dimension::commit(self)
    }
}

/// Private physical evidence retained by a process-local patch.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct Evidence {
    pub(crate) model_component: usize,
    pub(crate) model_object: usize,
    pub(crate) model_message: usize,
    pub(crate) model_identifier: u64,
    pub(crate) bucket_component: usize,
    pub(crate) bucket_object: usize,
    pub(crate) bucket_message: usize,
    pub(crate) bucket_identifier: u64,
}

/// A reversible, process-local exact-source dimension patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    pub(crate) artifacts: OwnedExactArtifacts,
    pub(crate) sheet_position: usize,
    pub(crate) table_position: usize,
    pub(crate) dimension: Dimension,
    pub(crate) before: Size,
    pub(crate) after: Size,
    pub(crate) evidence: Evidence,
    pub(crate) touched_components: usize,
    pub(crate) source_previews: usize,
    pub(crate) target_previews: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected semantic location.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Dimension {
            sheet: self.sheet_position,
            table: self.table_position,
            dimension: self.dimension,
        }
    }
    /// Return the exact semantic source size.
    #[must_use]
    pub const fn before(&self) -> Size {
        self.before
    }
    /// Return the exact semantic target size.
    #[must_use]
    pub const fn after(&self) -> Size {
        self.after
    }
    /// Return the selected row or column.
    #[must_use]
    pub const fn dimension(&self) -> Dimension {
        self.dimension
    }
    /// Return the diagnostic source fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }
    /// Return the diagnostic target fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
    /// Return whether this patch retains exact byte identity.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }
    /// Return the exact target-to-source inverse in constant time.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            sheet_position: self.sheet_position,
            table_position: self.table_position,
            dimension: self.dimension,
            before: self.after,
            after: self.before,
            evidence: self.evidence,
            touched_components: self.touched_components,
            source_previews: self.target_previews,
            target_previews: self.source_previews,
        }
    }
}

/// Compact publication evidence.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    pub(crate) changed: bool,
    pub(crate) touched_components: usize,
    pub(crate) deleted_previews: usize,
    pub(crate) full_reparse_performed: bool,
}

impl Diagnostics {
    /// Return whether publication changed package bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    /// Return the number of rewritten payload components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
    /// Return the number of canonical previews deleted in this direction.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
    /// Return whether the candidate was fully reopened and reselected.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one transaction.
#[must_use = "a dimension commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    pub(crate) package: Package,
    pub(crate) patch: Patch,
    pub(crate) diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the fully verified package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    /// Consume this result and return the package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}
