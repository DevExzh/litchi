//! Content-free commit facts and limit vocabulary of the primary transaction.

/// A finite dimension of the primary workbook transaction.
///
/// An exceeded bound is reported as [`crate::Error::Limit`] with the
/// dimension, the observed amount and the fixed maximum. The staged style
/// memory budget, which is shared with the native style planner, keeps
/// reporting [`crate::Error::ResourceLimit`] with
/// [`litchi_core::Resource::Memory`].
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Retained native snapshot bytes needed by one workbook edit.
    SnapshotBytes,
    /// Cells in one copy or move of a cell range.
    TransferCells,
    /// Stored cells scanned for formula dependencies of one range transfer.
    TransferDependencyScan,
    /// Bytes of one SVG attachment input.
    SvgInputBytes,
}

impl LimitKind {
    /// A short structural description of the dimension.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::SnapshotBytes => "retained snapshot bytes",
            Self::TransferCells => "cells in one range transfer",
            Self::TransferDependencyScan => "stored cells scanned for one range transfer",
            Self::SvgInputBytes => "SVG attachment input bytes",
        }
    }
}

impl std::fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str(self.as_str())
    }
}

/// Content-free facts about one primary workbook publication.
///
/// The counts describe the reversible patch; they never carry sheet names,
/// cell values or other authored content.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Diagnostics {
    pub(super) changed: bool,
    pub(super) changes: usize,
    pub(super) package_changes: usize,
    pub(super) changed_parts: usize,
}

impl Diagnostics {
    /// Whether the published package differs from its source.
    ///
    /// `false` exactly when [`super::Patch::is_noop`] holds: the commit
    /// returned its source snapshot, so the published package keeps the
    /// source bytes and any signature byte for byte.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of semantic worksheet and catalog changes in the patch.
    #[must_use]
    pub const fn changes(self) -> usize {
        self.changes
    }

    /// Number of package-owned semantic changes in the patch.
    #[must_use]
    pub const fn package_changes(self) -> usize {
        self.package_changes
    }

    /// Number of OPC parts whose payload the publication replaced.
    #[must_use]
    pub const fn changed_parts(self) -> usize {
        self.changed_parts
    }
}
