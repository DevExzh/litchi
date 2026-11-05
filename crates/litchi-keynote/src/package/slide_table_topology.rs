//! Inserting and removing slide-table rows and columns (record 0412e).
//!
//! The structural rewrite is the shared table-topology engine
//! (`litchi_iwa_table::topology`), the engine the Numbers family runs: it
//! moves storage tiles, header buckets, the row/column UID map, border
//! layers, the model's counts and every formula coordinate that addresses
//! the table. This module resolves the Keynote slide and table selectors, clears a
//! removed row's or column's comments through the comment transaction, and
//! publishes an exact-source patch.

use std::fmt;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_table::structure::{Fault, FaultLimit};
use litchi_iwa_table::topology::{self as engine, Geometry, Operation, Target};
use thiserror::Error as ThisError;

use super::Package;
use super::slide_table_cell_rich_text::owner;
use super::slide_table_core as core;
use crate::SlideSelector;
use crate::slide::table::TableSelector;

pub use litchi_iwa_common::table::topology::{
    ColumnDeletion, ColumnInsertion, RowDeletion, RowInsertion,
};
pub use litchi_iwa_table::topology::Unsupported;

/// A finite resource governed by the transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source or candidate package bytes.
    PackageBytes,
    /// Bytes in one decoded component.
    PayloadBytes,
    /// Native object identifiers.
    Identifiers,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::Identifiers => "object identifiers",
        })
    }
}

/// A typed, content-redacted slide-table topology failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`]
    /// classifies it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted slide-table topology source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("{kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} table-topology items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("table-topology candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("table-topology patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The selectors did not resolve one rooted slide table.
    #[error("slide-table selector did not resolve")]
    TableNotFound,
    /// The section-relative index is outside its section.
    #[error("index {index} is outside a section of {length}")]
    PositionOutOfRange {
        /// Requested section-relative index.
        index: usize,
        /// The section's length (an insertion may equal it).
        length: usize,
    },
    /// A header or footer section would exceed five rows or columns.
    #[error("a header or footer section holds at most five rows or columns")]
    SectionFull,
    /// The edit would remove the table's last row or column, or a body row
    /// or column its header and footer sections need.
    #[error("the table must keep at least one row and one column")]
    LastAxis,
    /// The edit touches table content this transaction does not rewrite.
    #[error("table topology edits are not supported for {reason}")]
    Unsupported {
        /// What the table holds.
        reason: Unsupported,
    },
    /// Commit was requested before staging an operation.
    #[error("no table-topology operation is staged")]
    NoStagedOperation,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support table topology changes")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::LastAxis => RefusalKind::UnsupportedTarget,
            Self::TableNotFound => RefusalKind::NotFound,
            Self::PositionOutOfRange { .. } => RefusalKind::InvalidValue,
            Self::SectionFull | Self::Unsupported { .. } | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<Fault> for Error {
    fn from(fault: Fault) -> Self {
        match fault {
            Fault::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    FaultLimit::PackageBytes => LimitKind::PackageBytes,
                    FaultLimit::Identifiers => LimitKind::Identifiers,
                    _ => LimitKind::PayloadBytes,
                },
                observed,
                maximum,
            },
            Fault::Allocation { amount } => Self::Allocation { amount },
            Fault::Unsupported | Fault::Dependency => Self::Refused(Refusal::Unsupported {
                reason: Unsupported::StorageLayout,
            }),
            Fault::Verification => Self::Verification,
            _ => Self::InvalidSource,
        }
    }
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        use engine::{Error as E, Refusal as R};
        match error {
            E::Refused(R::PositionOutOfRange { index, length }) => {
                Self::Refused(Refusal::PositionOutOfRange { index, length })
            },
            E::Refused(R::SectionFull) => Self::Refused(Refusal::SectionFull),
            E::Refused(R::LastAxis) => Self::Refused(Refusal::LastAxis),
            E::Refused(R::Unsupported { reason }) => Self::Refused(Refusal::Unsupported { reason }),
            E::Fault(fault) => Self::from(fault),
            E::Allocation { amount } => Self::Allocation { amount },
            _ => Self::InvalidSource,
        }
    }
}

/// A reversible, exact-source slide-table topology patch.
///
/// The patch retains the exact source and target artifacts; it authorizes
/// application only to its source, and its inverse restores the source byte
/// for byte.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the patch is an exact artifact no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics for one slide-table topology publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
}

/// One validated slide-table topology publication.
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume this publication and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

/// One immutable edit that stages row and column insertions and removals in
/// one slide table, applied in order.
pub struct Edit<'a> {
    source: &'a Package,
    slide: usize,
    table: usize,
    geometry: Geometry,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide", &self.slide)
            .field("table", &self.table)
            .field("operations", &self.operations)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(mut self, operation: Operation) -> Result<Self, Error> {
        let step = engine::resolve(operation, self.geometry)?;
        self.operations
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.operations.push(operation);
        self.geometry = step.after();
        Ok(self)
    }

    /// Stage a blank row inserted at a section-relative position. An index
    /// equal to the section's length appends to it; a header or footer
    /// insertion grows that section.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::SectionFull`].
    pub fn insert_row(self, insertion: RowInsertion) -> Result<Self, Error> {
        self.stage(Operation::InsertRow(insertion))
    }

    /// Stage a blank column inserted at a section-relative position.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::SectionFull`].
    pub fn insert_column(self, insertion: ColumnInsertion) -> Result<Self, Error> {
        self.stage(Operation::InsertColumn(insertion))
    }

    /// Stage the removal of the row at a section-relative position, with its
    /// cells and their comments.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::LastAxis`].
    pub fn remove_row(self, deletion: RowDeletion) -> Result<Self, Error> {
        self.stage(Operation::RemoveRow(deletion))
    }

    /// Stage the removal of the column at a section-relative position, with
    /// its cells and their comments.
    ///
    /// # Errors
    ///
    /// [`Refusal::PositionOutOfRange`], [`Refusal::LastAxis`].
    pub fn remove_column(self, deletion: ColumnDeletion) -> Result<Self, Error> {
        self.stage(Operation::RemoveColumn(deletion))
    }

    /// Validate and atomically publish the staged operations, in order.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit;
    /// [`Refusal::Unsupported`] for table content this transaction does not
    /// rewrite; source, resource and verification errors. Failure never
    /// publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.operations.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for operation in &self.operations {
            let base = current.as_ref().unwrap_or(self.source);
            let (package, components) = run(base, self.slide, self.table, *operation)?;
            touched = touched.saturating_add(components);
            current = Some(package);
        }
        let package = current.ok_or(Error::Refused(Refusal::NoStagedOperation))?;
        Ok(Commit {
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(owner(self.source), owner(&package)),
            },
            diagnostics: Diagnostics {
                changed: true,
                touched_components: touched,
            },
            package,
        })
    }
}

/// Resolve a slide table to its slide and table positions and native
/// drawable and model.
fn locate(
    source: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
) -> Result<(usize, usize, Target), Error> {
    let mut budget = core::Budget::new(source).map_err(locate_error)?;
    let selected = core::select_table(source, slide, table, &mut budget).map_err(locate_error)?;
    Ok((
        selected.slide_position.get(),
        selected.table_position.get(),
        Target {
            info: selected.table_info.identifier,
            model: selected.model.identifier,
        },
    ))
}

fn locate_error(error: core::Error) -> Error {
    match error {
        core::Error::SlideNameNotFound
        | core::Error::SlidePositionNotFound(_)
        | core::Error::TablePositionNotFound(_)
        | core::Error::EmptySlideName
        | core::Error::AmbiguousSelector => Error::Refused(Refusal::TableNotFound),
        core::Error::UnsupportedSource => Error::Refused(Refusal::UnsupportedSource),
        core::Error::Allocation(amount) => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// Clear the comments on the removed cells through the comment transaction
/// (the host removes their storage graphs with the axis); `None` when none
/// of the cells holds one.
fn clear_comments(
    source: &Package,
    slide: usize,
    table: usize,
    cells: &[At],
) -> Result<Option<Package>, Error> {
    let refused = |_error| {
        Error::Refused(Refusal::Unsupported {
            reason: Unsupported::Comments,
        })
    };
    let mut current: Option<Package> = None;
    for position in cells {
        let base = current.as_ref().unwrap_or(source);
        if base
            .slide_table_cell_comment(
                SlideSelector::index(slide),
                TableSelector::index(table),
                *position,
            )
            .map_err(refused)?
            .is_none()
        {
            continue;
        }
        let commit = base
            .edit_slide_table_cell_comment(
                SlideSelector::index(slide),
                TableSelector::index(table),
                *position,
            )
            .map_err(refused)?
            .clear()
            .map_err(refused)?
            .commit()
            .map_err(refused)?;
        current = Some(commit.into_package());
    }
    Ok(current)
}

/// Run one operation: check, clear the removed cells' comments, rewrite,
/// publish and verify.
fn run(
    source: &Package,
    slide: usize,
    table: usize,
    operation: Operation,
) -> Result<(Package, usize), Error> {
    let select = |package: &Package| {
        locate(
            package,
            SlideSelector::index(slide),
            TableSelector::index(table),
        )
    };
    let (_, _, target) = select(source)?;
    let before = engine::read_geometry(source, &target)?;
    let step = engine::resolve(operation, before)?;
    engine::check(source, &target, step, before)?;
    let cells = engine::removed_cells(source, &target, step, before)?;
    let uncommented = clear_comments(source, slide, table, &cells)?;
    let base = uncommented.as_ref().unwrap_or(source);
    let (_, _, target) = select(base)?;
    // The removed cells go with the axis: their values, styles, strings and
    // formulas are released in the structural pass, and a formula that
    // reads one reads `#REF!` (record 0329).
    let mut options = engine::Publish::default();
    options.drop_cells = !cells.is_empty();
    options.fit_drawable = true;
    let published = engine::publish(base, &target, step, before, options)?;
    let (_, _, after_target) = select(&published.package)?;
    if engine::read_geometry(&published.package, &after_target)? != step.after() {
        return Err(Error::Verification);
    }
    Ok((published.package, published.touched_components))
}

impl Package {
    /// Start an immutable row and column insertion and removal edit on one
    /// slide table.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`]; source errors for a semantic-only or
    /// malformed source.
    pub fn edit_slide_table_topology<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let (slide, table, target) = locate(self, slide.into(), table.into())?;
        let geometry = engine::read_geometry(self, &target)?;
        Ok(Edit {
            source: self,
            slide,
            table,
            geometry,
            operations: Vec::new(),
        })
    }

    /// Apply a slide-table topology patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_slide_table_topology(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch.artifacts.authorizes_owner(&owner(self)) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::default(),
            });
        }
        let target = patch.artifacts.target_owner();
        let candidate = Package::from_source_with_options(
            std::sync::Arc::from(target.as_ref()),
            self.state.options,
        )
        .map_err(|_error| Error::Verification)?;
        candidate.validate().map_err(|_error| Error::Verification)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components: 0,
            },
        })
    }
}
