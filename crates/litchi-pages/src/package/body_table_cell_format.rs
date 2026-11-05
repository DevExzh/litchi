//! Display formats of stored body-table cells (record 0379).
//!
//! One family takes the neutral [`DataFormat`] for every display format:
//! Number, Currency, Percentage, Scientific, Fraction, Date and Time,
//! Duration, Text and Custom. The native graph surgery is the shared
//! `litchi_iwa_table::cell_format` engine the Numbers owners use; this module
//! resolves the body-table selector, publishes the engine's members through
//! exact-source reassembly and reopens the candidate. A change between two
//! families runs as two engine rewrites (the first to Automatic) on a private
//! intermediate candidate; the patch keeps only the source and final
//! artifacts, so its inverse restores the source byte for byte.
//!
//! Checkbox, Star Rating, Slider and Stepper cells are written through the
//! engine's scalar-control writer (record 0412); Pop-Up Menus through the
//! shared Pop-Up Menu publication, which also allocates the menu model and
//! rewrites the package metadata registry (record 0412a).

use std::fmt;

use litchi_iwa_archive::package::{OwnedExactArtifacts, SharedBytes};
use litchi_iwa_common::table::cell::data_format::DataFormat;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_table::cell_format::{self as engine, TransactionBudget, data_format, publish};
use thiserror::Error as ThisError;

use super::{Package, table_lock};
use crate::BodyTableSelector;

/// A finite resource governed by a body-table cell-format transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete candidate package output bytes.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// Bytes in one physical entry.
    EntryBytes,
    /// Aggregate physical entry bytes.
    TotalEntryBytes,
    /// Physical package metadata bytes.
    PackageBytes,
    /// Decoded native payload bytes.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native framing items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Strict codec input bytes.
    WireBytes,
    /// Strict codec output bytes.
    WireOutputBytes,
    /// Strict codec reference-envelope bytes.
    WireReferenceBytes,
    /// Strict codec selected-text bytes.
    WireTextBytes,
    /// Strict codec fields inspected.
    WireFields,
    /// Strict codec nesting depth.
    WireNesting,
    /// Strict codec work.
    WireWork,
    /// Private scratch bytes.
    ScratchBytes,
    /// Private candidate bytes retained during execution.
    RetainedBytes,
    /// Bounded allocation units.
    Allocations,
    /// Compressed physical member bytes.
    CompressedBytes,
    /// Aggregate transaction work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A typed, content-redacted body-table cell-format failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted table graph is malformed or ambiguous.
    #[error("the body-table cell-format source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error("body-table cell-format {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} body-table cell-format units")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the staged format.
    #[error("body-table cell-format candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source, or
    /// the cell changed since the edit was staged.
    #[error("body-table cell-format patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No body table matches the selector.
    #[error("no body table matches the selector")]
    TableNotFound,
    /// The cell is outside the table or holds no stored cell record.
    #[error("the body-table cell is outside the table or not stored")]
    CellNotFound,
    /// The table is locked.
    #[error("the body table is locked")]
    TableLocked,
    /// A native dependency of the cell is outside the shared engine.
    #[error("the body-table cell has an unsupported native dependency")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support body-table cell formats")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound | Self::CellNotFound => RefusalKind::NotFound,
            Self::TableLocked => RefusalKind::Protected,
            Self::UnsupportedDependency | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

const fn map_limit_kind(kind: engine::LimitKind) -> LimitKind {
    use engine::LimitKind as E;
    match kind {
        E::InputBytes => LimitKind::InputBytes,
        E::OutputBytes => LimitKind::OutputBytes,
        E::Entries => LimitKind::Entries,
        E::EntryBytes => LimitKind::EntryBytes,
        E::TotalEntryBytes => LimitKind::TotalEntryBytes,
        E::PackageBytes => LimitKind::PackageBytes,
        E::PayloadBytes => LimitKind::PayloadBytes,
        E::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
        E::PayloadObjects => LimitKind::PayloadObjects,
        E::PayloadMessages => LimitKind::PayloadMessages,
        E::PayloadItems => LimitKind::PayloadItems,
        E::PayloadReferences => LimitKind::PayloadReferences,
        E::WireBytes => LimitKind::WireBytes,
        E::WireOutputBytes => LimitKind::WireOutputBytes,
        E::WireReferenceBytes => LimitKind::WireReferenceBytes,
        E::WireTextBytes => LimitKind::WireTextBytes,
        E::WireFields => LimitKind::WireFields,
        E::WireNesting => LimitKind::WireNesting,
        E::WireWork => LimitKind::WireWork,
        E::ScratchBytes => LimitKind::ScratchBytes,
        E::RetainedBytes => LimitKind::RetainedBytes,
        E::Allocations => LimitKind::Allocations,
        E::CompressedBytes => LimitKind::CompressedBytes,
        E::TransactionWork => LimitKind::TransactionWork,
    }
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        match error {
            engine::Error::Refused(refusal) => Self::Refused(match refusal {
                engine::Refusal::CellNotFound => Refusal::CellNotFound,
                engine::Refusal::TableLocked { .. } => Refusal::TableLocked,
                engine::Refusal::UnsupportedDependency { .. } => Refusal::UnsupportedDependency,
                engine::Refusal::UnsupportedSource => Refusal::UnsupportedSource,
            }),
            // The engine reads a stored cell record; a cell without one is
            // reported as an invalid tile slot.
            engine::Error::InvalidSource { .. } => Self::InvalidSource,
            engine::Error::Limit {
                kind,
                observed,
                maximum,
                ..
            } => Self::Limit {
                kind: map_limit_kind(kind),
                observed,
                maximum,
            },
            engine::Error::Allocation { amount, .. } => Self::Allocation { amount },
            engine::Error::Verification => Self::Verification,
            engine::Error::PatchConflict => Self::PatchConflict,
        }
    }
}

fn map_target_error(error: table_lock::Error) -> Error {
    match error {
        table_lock::Error::Refused(
            table_lock::Refusal::TableNotFound | table_lock::Refusal::AmbiguousTableName,
        ) => Error::Refused(Refusal::TableNotFound),
        table_lock::Error::Refused(table_lock::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        table_lock::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// A reversible exact-source body-table cell-format patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    table: usize,
    at: At,
    before: DataFormat,
    after: DataFormat,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("table", &self.table)
            .field("at", &self.at)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The zero-based body-table position the patch edits.
    #[must_use]
    pub const fn table(&self) -> usize {
        self.table
    }

    /// The edited cell.
    #[must_use]
    pub const fn at(&self) -> At {
        self.at
    }

    /// The source format.
    #[must_use]
    pub const fn before(&self) -> &DataFormat {
        &self.before
    }

    /// The target format.
    #[must_use]
    pub const fn after(&self) -> &DataFormat {
        &self.after
    }

    /// Whether the patch is an exact no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            table: self.table,
            at: self.at,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Content-free diagnostics for one body-table cell-format publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    rewrites: usize,
}

impl Diagnostics {
    /// Whether the exact package artifact changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Engine rewrites published: one, or two for a change between families.
    #[must_use]
    pub const fn rewrites(self) -> usize {
        self.rewrites
    }
}

/// One validated body-table cell-format publication.
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

/// One staged format change of one stored body-table cell.
pub struct Edit<'a> {
    source: &'a Package,
    table: usize,
    at: At,
    before: DataFormat,
    after: DataFormat,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("table", &self.table)
            .field("at", &self.at)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// The format observed when the edit was created.
    #[must_use]
    pub const fn before(&self) -> &DataFormat {
        &self.before
    }

    /// The staged format.
    #[must_use]
    pub const fn after(&self) -> &DataFormat {
        &self.after
    }

    /// Stage `format` for the cell ([`DataFormat::Automatic`] clears it).
    #[must_use]
    pub fn set(mut self, format: impl Into<DataFormat>) -> Self {
        self.after = format.into();
        self
    }

    /// Stage iWork's automatic format for the cell.
    #[must_use]
    pub fn reset(mut self) -> Self {
        self.after = DataFormat::Automatic;
        self
    }

    /// Validate and atomically publish the staged format.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableLocked`]; the engine's source,
    /// dependency and resource failures; [`Error::Verification`]. A failure
    /// never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let source_owner = owner(self.source);
        if self.before == self.after {
            return Ok(Commit {
                package: self.source.clone(),
                patch: Patch {
                    artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                    table: self.table,
                    at: self.at,
                    before: self.before,
                    after: self.after,
                },
                diagnostics: Diagnostics::default(),
            });
        }
        let cross_family = matches!(
            (data_format::family(&self.before), data_format::family(&self.after)),
            (Some(left), Some(right)) if left != right
        );
        let (package, rewrites) = if cross_family {
            let cleared = rewrite(
                self.source,
                self.table,
                self.at,
                &self.before,
                &DataFormat::Automatic,
            )?;
            (
                rewrite(
                    &cleared,
                    self.table,
                    self.at,
                    &DataFormat::Automatic,
                    &self.after,
                )?,
                2,
            )
        } else {
            (
                rewrite(self.source, self.table, self.at, &self.before, &self.after)?,
                1,
            )
        };
        let target_owner = owner(&package);
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
                table: self.table,
                at: self.at,
                before: self.before,
                after: self.after,
            },
            diagnostics: Diagnostics {
                changed: true,
                rewrites,
            },
        })
    }
}

fn owner(package: &Package) -> SharedBytes {
    SharedBytes::from_shared_slice(package.state.source.shared_source())
}

/// Resolve the selected body table and cell into the engine's target.
fn resolve(
    package: &Package,
    selector: BodyTableSelector<'_>,
    at: At,
) -> Result<engine::Target, Error> {
    let table = package
        .resolve_body_table(selector)
        .map_err(map_target_error)?;
    if at.row() >= table.table_rows || at.column() >= table.table_columns {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    Ok(engine::Target {
        scope: 0,
        table: table.table_position,
        position: at,
        model_identifier: table.model_identifier.get(),
        component_index: table.model_component_index,
        object_index: table.model_object_index,
        message_index: table.model_message_index,
        message_type: table.model_message_type,
        locked: table.explicit_locked == Some(true),
    })
}

fn read_at(package: &Package, target: engine::Target) -> Result<DataFormat, Error> {
    Ok(data_format::read(package, target, target.path())?)
}

/// One engine rewrite of one family, published and reopened.
fn rewrite(
    source: &Package,
    table: usize,
    at: At,
    before: &DataFormat,
    after: &DataFormat,
) -> Result<Package, Error> {
    let target = resolve(source, BodyTableSelector::index(table), at)?;
    if read_at(source, target)? != *before {
        return Err(Error::PatchConflict);
    }
    let catalog = &source.state.source;
    if !catalog.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let path = target.path();
    let pop_up = |format: &DataFormat| match format {
        DataFormat::PopUpMenu(menu) => Some(menu.clone()),
        _ => None,
    };
    if matches!(before, DataFormat::PopUpMenu(_)) || matches!(after, DataFormat::PopUpMenu(_)) {
        // The shared Pop-Up Menu publication (record 0412a): the menu model,
        // its metadata registry entry and the cell's control spec, with the
        // host's root-preview policy (kept) and list-key policy (append).
        let published = engine::pop_up_publish::publish(
            source,
            target,
            pop_up(before).as_ref(),
            pop_up(after).as_ref(),
            false,
            engine::ListKeys::Append,
            false,
            |bytes| {
                Package::from_bytes_with_archive_limits(&bytes, catalog.limits())
                    .map_err(|_| engine::Error::Verification)
            },
            |candidate| {
                resolve(candidate, BodyTableSelector::index(table), at)
                    .map_err(|_| engine::Error::Verification)
            },
        )?;
        return Ok(published.package);
    }
    // Pages and Keynote append list keys at the watermark and hand a removed
    // last key back; the Numbers owners reuse free keys (record 0379).
    let mut budget =
        TransactionBudget::for_cell_control(source).with_list_keys(engine::ListKeys::Append);
    budget.charge_package_source(catalog, path)?;
    let bound =
        engine::control::preflight_copy_on_write_scalar_control(source, target, path, &mut budget)?;
    let output = data_format::rewrite_family(source, target, before, after, path, &mut budget)?;
    let family = data_format::family(before)
        .or_else(|| data_format::family(after))
        .ok_or(Error::Verification)?;
    let bytes = publish::reassemble(
        source,
        &output.members,
        data_format::maximum_members(family),
        bound,
        path,
        &mut budget,
    )?;
    let candidate = Package::from_bytes_with_archive_limits(&bytes, catalog.limits())
        .map_err(|_| Error::Verification)?;
    budget.charge_candidate_reopen(&candidate.state.source, path)?;
    let candidate_target = resolve(&candidate, BodyTableSelector::index(table), at)?;
    if read_at(&candidate, candidate_target)? != *after {
        return Err(Error::Verification);
    }
    Ok(candidate)
}

impl Package {
    /// Read the display format of one stored body-table cell.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::CellNotFound`] (outside the
    /// table, or a cell with no stored record), [`Refusal::UnsupportedDependency`]
    /// for a format no engine family admits, and source failures.
    pub fn body_table_cell_format<'table>(
        &self,
        table: impl Into<BodyTableSelector<'table>>,
        at: At,
    ) -> Result<DataFormat, Error> {
        let target = resolve(self, table.into(), at)?;
        read_at(self, target)
    }

    /// Start a format edit of one stored body-table cell.
    ///
    /// # Errors
    ///
    /// As [`Self::body_table_cell_format`].
    pub fn edit_body_table_cell_format<'table>(
        &self,
        table: impl Into<BodyTableSelector<'table>>,
        at: At,
    ) -> Result<Edit<'_>, Error> {
        let target = resolve(self, table.into(), at)?;
        let before = read_at(self, target)?;
        Ok(Edit {
            source: self,
            table: target.table,
            at,
            after: before.clone(),
            before,
        })
    }

    /// Replay a body-table cell-format patch on its exact source.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] on another artifact or a changed cell;
    /// [`Error::Verification`] when the target does not reopen to the
    /// patch's format.
    pub fn apply_body_table_cell_format(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch.artifacts.authorizes_owner(&owner(self)) {
            return Err(Error::PatchConflict);
        }
        let target = resolve(self, BodyTableSelector::index(patch.table), patch.at)?;
        if read_at(self, target)? != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::default(),
            });
        }
        let target_owner = patch.artifacts.target_owner();
        let candidate = Package::from_bytes_with_archive_limits(
            target_owner.as_ref(),
            self.state.source.limits(),
        )
        .map_err(|_| Error::Verification)?;
        let candidate_target =
            resolve(&candidate, BodyTableSelector::index(patch.table), patch.at)?;
        if read_at(&candidate, candidate_target)? != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                rewrites: 1,
            },
        })
    }
}
