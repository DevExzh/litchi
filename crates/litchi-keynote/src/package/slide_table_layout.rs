//! Slide-table layout (record 0412g): resizing a slide table at its
//! trailing edges and setting its drawable frame, with every row and column
//! sized uniformly to fill the frame, as Keynote's host moves a slide
//! table's bounds and dimensions together.
//!
//! The rewrites are the shared table engine's
//! (`litchi_iwa_table::topology::{stage_resize, stage_frame,
//! stage_uniform_dimensions}`, the resize lifted from Numbers); this module
//! resolves the Keynote selectors and publishes an exact-source patch.

use std::fmt;

use litchi_iwa_archive::package::OwnedExactArtifacts;
use litchi_iwa_table::structure::{Fault, FaultLimit};
use litchi_iwa_table::structure::{Published, Working, metadata::Transition};
use litchi_iwa_table::topology::{self as engine, Frame, Target};
use thiserror::Error as ThisError;

use super::Package;
use super::slide_table_cell_rich_text::owner;
use super::slide_table_core as core;
use crate::SlideSelector;
use crate::slide::table::TableSelector;

use crate::slide::drawable::Geometry;

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

/// A typed, content-redacted slide-table layout failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`]
    /// classifies it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted slide-table layout source is invalid")]
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
    #[error("could not allocate {amount} table-layout items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("table-layout candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("table-layout patch conflicts with this package")]
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
    /// The size is zero, leaves no body row or column, or needs more row
    /// and column identities than a table may hold.
    #[error("the table cannot take this size")]
    InvalidSize,
    /// The rows or columns a shrink removes hold stored cells.
    #[error("the removed rows or columns hold stored cells")]
    StoredCells,
    /// The frame is not finite or has no positive size.
    #[error("the table frame is not a finite positive rectangle")]
    InvalidGeometry,
    /// Commit was requested before staging an operation.
    #[error("no table-layout operation is staged")]
    NoStagedOperation,
    /// The table's storage uses a layout this route does not rewrite.
    #[error("source representation does not support table layout changes")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::TableNotFound => RefusalKind::NotFound,
            Self::InvalidSize | Self::StoredCells | Self::InvalidGeometry => {
                RefusalKind::InvalidValue
            },
            Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::UnsupportedSource => RefusalKind::UnsupportedContent,
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
            Fault::Unsupported | Fault::Dependency => Self::Refused(Refusal::UnsupportedSource),
            Fault::Verification => Self::Verification,
            _ => Self::InvalidSource,
        }
    }
}

impl From<engine::Error> for Error {
    fn from(error: engine::Error) -> Self {
        use engine::{Error as E, Refusal as R};
        match error {
            E::Refused(R::InvalidSize) => Self::Refused(Refusal::InvalidSize),
            E::Refused(R::StoredCells) => Self::Refused(Refusal::StoredCells),
            E::Refused(_) => Self::Refused(Refusal::UnsupportedSource),
            E::Fault(fault) => Self::from(fault),
            E::Allocation { amount } => Self::Allocation { amount },
            _ => Self::InvalidSource,
        }
    }
}

/// A reversible, exact-source slide-table layout patch.
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

/// Content-free diagnostics for one slide-table layout publication.
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

/// One validated slide-table layout publication.
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

/// One staged layout operation.
#[derive(Debug, Clone, Copy, PartialEq)]
enum Operation {
    Resize { rows: u32, columns: u32 },
    Frame(Geometry),
}

/// One immutable edit that stages resizes and frame changes of one slide
/// table, applied in order.
pub struct Edit<'a> {
    source: &'a Package,
    slide: usize,
    table: usize,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("slide", &self.slide)
            .field("table", &self.table)
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn stage(mut self, operation: Operation) -> Result<Self, Error> {
        self.operations
            .try_reserve(1)
            .map_err(|_error| Error::Allocation { amount: 1 })?;
        self.operations.push(operation);
        Ok(self)
    }

    /// Stage `rows` x `columns` (header and footer rows and header columns
    /// included) as the table's size: trailing rows and columns are added
    /// blank or removed when empty, and every row and column is sized
    /// uniformly to fill the table's frame.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidSize`] for a zero size.
    pub fn resize(self, rows: u32, columns: u32) -> Result<Self, Error> {
        if rows == 0 || columns == 0 {
            return Err(Error::Refused(Refusal::InvalidSize));
        }
        self.stage(Operation::Resize { rows, columns })
    }

    /// Stage the table's frame; every row and column is sized uniformly to
    /// fill it. A frame without rotation keeps the table's rotation.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidGeometry`] for a frame without a finite position
    /// and a positive finite size.
    pub fn set_frame(self, geometry: Geometry) -> Result<Self, Error> {
        let finite = [geometry.x, geometry.y, geometry.width, geometry.height]
            .iter()
            .all(|value| value.is_finite())
            && geometry.width > 0.0
            && geometry.height > 0.0
            && geometry.rotation.is_none_or(f32::is_finite);
        if !finite {
            return Err(Error::Refused(Refusal::InvalidGeometry));
        }
        self.stage(Operation::Frame(geometry))
    }

    /// Validate and atomically publish the staged operations, in order. An
    /// operation that leaves the table as it is is skipped.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit, the refusals,
    /// source, resource and verification errors. Failure never publishes a
    /// partial package.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.operations.is_empty() {
            return Err(Error::Refused(Refusal::NoStagedOperation));
        }
        let mut current: Option<Package> = None;
        let mut touched = 0_usize;
        for operation in &self.operations {
            let base = current.as_ref().unwrap_or(self.source);
            if let Some(published) = run(base, self.slide, self.table, *operation)? {
                touched = touched.saturating_add(published.touched_components);
                current = Some(published.package);
            }
        }
        let Some(package) = current else {
            let owner = owner(self.source);
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
                },
                diagnostics: Diagnostics::default(),
            });
        };
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

/// Run one operation and verify it; `None` when it changes nothing.
fn run(
    source: &Package,
    slide: usize,
    table: usize,
    operation: Operation,
) -> Result<Option<Published<Package>>, Error> {
    let select = |package: &Package| {
        locate(
            package,
            SlideSelector::index(slide),
            TableSelector::index(table),
        )
    };
    let (_, _, target) = select(source)?;
    let mut working = Working::new(source)?;
    working.keep_previews();
    let before = engine::frame(&working, &target)?;
    let fit = match operation {
        Operation::Resize { rows, columns } => {
            if !engine::stage_resize(&mut working, &target, rows, columns)? {
                return Ok(None);
            }
            before.size
        },
        Operation::Frame(geometry) => {
            let replacement = Frame {
                position: Some((geometry.x, geometry.y)),
                size: Some((geometry.width, geometry.height)),
                flags: before.flags,
                angle: geometry.rotation.or(before.angle),
            };
            if replacement == before {
                return Ok(None);
            }
            engine::stage_frame(&mut working, &target, replacement)?;
            replacement.size
        },
    };
    if let Some((width, height)) = fit {
        engine::stage_uniform_dimensions(&mut working, &target, width, height)?;
    }
    let published = working.publish(&Transition::default())?;
    let (_, _, after) = select(&published.package)?;
    let geometry = engine::read_geometry(&published.package, &after)?;
    let frame = engine::frame(&Working::new(&published.package)?, &after)?;
    let expected = match operation {
        Operation::Resize { rows, columns } => (rows, columns) == (geometry.rows, geometry.columns),
        Operation::Frame(geometry) => {
            frame.position == Some((geometry.x, geometry.y))
                && frame.size == Some((geometry.width, geometry.height))
        },
    };
    if !expected {
        return Err(Error::Verification);
    }
    Ok(Some(published))
}

impl Package {
    /// Start an immutable layout edit (resize, frame) of one slide table.
    ///
    /// # Errors
    ///
    /// [`Refusal::TableNotFound`]; source errors.
    pub fn edit_slide_table_layout<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let (slide, table, _) = locate(self, slide.into(), table.into())?;
        Ok(Edit {
            source: self,
            slide,
            table,
            operations: Vec::new(),
        })
    }

    /// Apply a slide-table layout patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_slide_table_layout(&self, patch: &Patch) -> Result<Commit, Error> {
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
