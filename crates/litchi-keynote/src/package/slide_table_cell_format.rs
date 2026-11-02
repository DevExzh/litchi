//! Display formats of stored slide-table cells (record 0379).
//!
//! One family takes the neutral [`DataFormat`] for every display format:
//! Number, Currency, Percentage, Scientific, Fraction, Date and Time,
//! Duration, Text and Custom. The native graph surgery is the shared
//! `litchi_iwa_table::cell_format` engine the Numbers owners use; this module
//! resolves the slide and table selectors, publishes the engine's members through
//! exact-source reassembly and reopens the candidate. The dedicated
//! `edit_slide_table_cell_number_format` family stays; this family covers
//! every display format. A change between two
//! families runs as two engine rewrites (the first to Automatic) on a private
//! intermediate candidate; the patch keeps only the source and final
//! artifacts, so its inverse restores the source byte for byte.
//!
//! Checkbox, Star Rating, Slider and Stepper cells are written through the
//! engine's scalar-control writer (record 0412). Pop-Up Menus also allocate
//! a menu model and rewrite the package metadata registry; they are read
//! but refused for writing ([`Refusal::UnsupportedFormat`]).

use std::fmt;

use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::table::cell::data_format::DataFormat;
use litchi_iwa_common::table::coordinate::At;
use litchi_iwa_table::cell_format::{self as engine, TransactionBudget, data_format, publish};
use thiserror::Error as ThisError;

use super::{Package, slide_table_core as core};
use crate::SlideSelector;
use crate::slide::table::TableSelector;

/// A finite resource governed by a slide-table cell-format transaction.
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

/// A typed, content-redacted slide-table cell-format failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted table graph is malformed or ambiguous.
    #[error("the slide-table cell-format source is invalid")]
    InvalidSource,
    /// A finite transaction limit was exceeded.
    #[error(
        "slide-table cell-format {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} slide-table cell-format units")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the staged format.
    #[error("slide-table cell-format candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source, or
    /// the cell changed since the edit was staged.
    #[error("slide-table cell-format patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No slide table matches the selector.
    #[error("no slide or slide table matches the selectors")]
    TableNotFound,
    /// The cell is outside the table or holds no stored cell record.
    #[error("the slide-table cell is outside the table or not stored")]
    CellNotFound,
    /// The table is locked.
    #[error("the slide table is locked")]
    TableLocked,
    /// The format is a Pop-Up Menu, whose writes also allocate a menu model
    /// and rewrite the package metadata registry.
    #[error("this slide-table cell format is not written by this route")]
    UnsupportedFormat,
    /// A native dependency of the cell is outside the shared engine.
    #[error("the slide-table cell has an unsupported native dependency")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support slide-table cell formats")]
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
            Self::UnsupportedFormat => RefusalKind::UnsupportedOperation,
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

fn map_target_error(error: core::Error) -> Error {
    match error {
        core::Error::SlideNameNotFound
        | core::Error::SlidePositionNotFound(_)
        | core::Error::TablePositionNotFound(_)
        | core::Error::EmptySlideName
        | core::Error::AmbiguousSelector => Error::Refused(Refusal::TableNotFound),
        core::Error::UnsupportedSource => Error::Refused(Refusal::UnsupportedSource),
        core::Error::UnsupportedDependency | core::Error::UnsupportedTopology => {
            Error::Refused(Refusal::UnsupportedDependency)
        },
        core::Error::Allocation(amount) => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

/// A reversible exact-source slide-table cell-format patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    slide: usize,
    table: usize,
    at: At,
    before: DataFormat,
    after: DataFormat,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("slide", &self.slide)
            .field("table", &self.table)
            .field("at", &self.at)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The zero-based slide position the patch edits.
    #[must_use]
    pub const fn slide(&self) -> usize {
        self.slide
    }

    /// The zero-based table position on that slide.
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
            slide: self.slide,
            table: self.table,
            at: self.at,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Content-free diagnostics for one slide-table cell-format publication.
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

/// One validated slide-table cell-format publication.
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

/// One staged format change of one stored slide-table cell.
pub struct Edit<'a> {
    source: &'a Package,
    slide: usize,
    table: usize,
    at: At,
    before: DataFormat,
    after: DataFormat,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide", &self.slide)
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
    /// [`Refusal::UnsupportedFormat`] for a Pop-Up Menu change; [`Refusal::TableLocked`]; the engine's source,
    /// dependency and resource failures; [`Error::Verification`]. A failure
    /// never publishes a partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        let source_owner = owner(self.source)?;
        if self.before == self.after {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_owner), source_owner),
                    slide: self.slide,
                    table: self.table,
                    at: self.at,
                    before: self.before,
                    after: self.after,
                },
                diagnostics: Diagnostics::default(),
            });
        }
        for format in [&self.before, &self.after] {
            if let Some(family) = data_format::family(format)
                && family == data_format::Family::PopUpMenu
            {
                return Err(Error::Refused(Refusal::UnsupportedFormat));
            }
        }
        let cross_family = matches!(
            (data_format::family(&self.before), data_format::family(&self.after)),
            (Some(left), Some(right)) if left != right
        );
        let (package, rewrites) = if cross_family {
            let cleared = rewrite(
                self.source,
                self.slide,
                self.table,
                self.at,
                &self.before,
                &DataFormat::Automatic,
            )?;
            (
                rewrite(
                    &cleared,
                    self.slide,
                    self.table,
                    self.at,
                    &DataFormat::Automatic,
                    &self.after,
                )?,
                2,
            )
        } else {
            (
                rewrite(
                    self.source,
                    self.slide,
                    self.table,
                    self.at,
                    &self.before,
                    &self.after,
                )?,
                1,
            )
        };
        let target_owner = owner(&package)?;
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(source_owner, target_owner),
                slide: self.slide,
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

fn owner(package: &Package) -> Result<Arc<[u8]>, Error> {
    Ok(core::physical_catalog(package)
        .map_err(map_target_error)?
        .shared_source())
}

/// Resolve the selected slide table and cell into the engine's target.
fn resolve(
    package: &Package,
    slide: SlideSelector<'_>,
    table: TableSelector,
    at: At,
) -> Result<engine::Target, Error> {
    let mut budget = core::Budget::new(package).map_err(map_target_error)?;
    let selected =
        core::select_table(package, slide, table, &mut budget).map_err(map_target_error)?;
    if at.row() >= selected.rows || at.column() >= selected.columns {
        return Err(Error::Refused(Refusal::CellNotFound));
    }
    let component_index = package
        .state
        .source
        .components()
        .iter()
        .position(|component| component.name() == &*selected.model.component)
        .ok_or(Error::InvalidSource)?;
    let message_type = package
        .state
        .source
        .components()
        .get_index(component_index)
        .and_then(|component| component.archive().objects.get(selected.model.object_index))
        .and_then(|object| object.messages.get(selected.model_message_index))
        .map(|message| message.type_)
        .ok_or(Error::InvalidSource)?;
    Ok(engine::Target {
        scope: selected.slide_position.get(),
        table: selected.table_position.get(),
        position: at,
        model_identifier: selected.model.identifier,
        component_index,
        object_index: selected.model.object_index,
        message_index: selected.model_message_index,
        message_type,
        locked: selected.locked,
    })
}

fn reopen(source: &Package, bytes: Arc<[u8]>) -> Result<Package, Error> {
    let candidate = Package::from_source_with_options(bytes, source.state.options)
        .map_err(|_| Error::Verification)?;
    candidate.validate().map_err(|_| Error::Verification)?;
    Ok(candidate)
}

fn read_at(package: &Package, target: engine::Target) -> Result<DataFormat, Error> {
    Ok(data_format::read(package, target, target.path())?)
}

/// One engine rewrite of one family, published and reopened.
fn rewrite(
    source: &Package,
    slide: usize,
    table: usize,
    at: At,
    before: &DataFormat,
    after: &DataFormat,
) -> Result<Package, Error> {
    let target = resolve(
        source,
        SlideSelector::index(slide),
        TableSelector::index(table),
        at,
    )?;
    if read_at(source, target)? != *before {
        return Err(Error::PatchConflict);
    }
    let catalog = core::physical_catalog(source).map_err(map_target_error)?;
    if !catalog.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let path = target.path();
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
    let candidate = reopen(source, Arc::from(bytes))?;
    budget.charge_candidate_reopen(
        core::physical_catalog(&candidate).map_err(|_| Error::Verification)?,
        path,
    )?;
    let candidate_target = resolve(
        &candidate,
        SlideSelector::index(slide),
        TableSelector::index(table),
        at,
    )?;
    if read_at(&candidate, candidate_target)? != *after {
        return Err(Error::Verification);
    }
    Ok(candidate)
}

impl Package {
    /// Read the display format of one stored slide-table cell.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`], [`Refusal::CellNotFound`] (outside the
    /// table, or a cell with no stored record), [`Refusal::UnsupportedDependency`]
    /// for a format no engine family admits, and source failures.
    pub fn slide_table_cell_format<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        at: At,
    ) -> Result<DataFormat, Error> {
        let target = resolve(self, slide.into(), table.into(), at)?;
        read_at(self, target)
    }

    /// Start a format edit of one stored slide-table cell.
    ///
    /// # Errors
    ///
    /// As [`Self::slide_table_cell_format`].
    pub fn edit_slide_table_cell_format<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
        at: At,
    ) -> Result<Edit<'_>, Error> {
        let target = resolve(self, slide.into(), table.into(), at)?;
        let before = read_at(self, target)?;
        Ok(Edit {
            source: self,
            slide: target.scope,
            table: target.table,
            at,
            after: before.clone(),
            before,
        })
    }

    /// Replay a slide-table cell-format patch on its exact source.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] on another artifact or a changed cell;
    /// [`Error::Verification`] when the target does not reopen to the
    /// patch's format.
    pub fn apply_slide_table_cell_format(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch.artifacts.authorizes_source(&owner(self)?) {
            return Err(Error::PatchConflict);
        }
        let target = resolve(
            self,
            SlideSelector::index(patch.slide),
            TableSelector::index(patch.table),
            patch.at,
        )?;
        if read_at(self, target)? != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::default(),
            });
        }
        let candidate = reopen(self, patch.artifacts.target())?;
        let candidate_target = resolve(
            &candidate,
            SlideSelector::index(patch.slide),
            TableSelector::index(patch.table),
            patch.at,
        )?;
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
