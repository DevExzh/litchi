//! Exact-source, selector-first chart lifecycle of a Keynote slide (record
//! 0372; a port of the migration host's `set_slide_chart_kind`,
//! `set_slide_chart_direction` and `set_slide_chart_geometry`).
//!
//! A chart is addressed by its slide and a [`ChartSelector`] (its position
//! in the slide's chart order, or its exact visible title). An [`Edit`]
//! stages operations against one immutable package; [`Edit::commit`]
//! applies them in one structural workspace, publishes one candidate,
//! reopens it and reads every touched chart back.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::chart::{Direction, kind::Kind};
use thiserror::Error as ThisError;

use super::slide_chart_title::{self, select_chart};
use super::slide_drawable::Geometry;
use super::structure::StructureError;
use super::structure::wire::Fields;
use super::structure::workspace::Workspace;
use super::{Package, PhysicalSource, ReadError};
use crate::{ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = 5_021;
/// `TSCH.ChartDrawableArchive.super` (`TSD.DrawableArchive`).
const CHART_DRAWABLE: u32 = 1;
const DRAWABLE_GEOMETRY: u32 = 1;
/// The `TSCH.ChartArchive` extension of the chart drawable.
const CHART_UNITY: u32 = 10_000;
const CHART_TYPE: u32 = 1;
const CHART_SERIES_DIRECTION: u32 = 5;

/// A finite resource governed by a chart-lifecycle transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
    /// Native objects in the package.
    Objects,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str("objects")
    }
}

/// A chart-lifecycle failure that never exposes native identifiers.
#[derive(Debug, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses.
    #[error(transparent)]
    Refused(Refusal),
    /// The slide or chart selector did not resolve.
    #[error(transparent)]
    Selection(slide_chart_title::Error),
    /// The package cannot be edited safely.
    #[error("the Keynote source cannot be edited safely")]
    InvalidSource,
    /// A resource limit was exceeded.
    #[error(
        "Keynote chart-lifecycle {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// The exceeded resource.
        kind: LimitKind,
        /// The observed amount.
        observed: u64,
        /// The configured maximum.
        maximum: u64,
    },
    /// The published candidate did not read back as staged.
    #[error("the Keynote chart-lifecycle candidate failed verification")]
    Verification,
    /// The patch does not match the exact source package.
    #[error("the Keynote chart-lifecycle patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request the chart-lifecycle family refuses.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The package was not opened from an exact physical source.
    #[error("this Keynote source does not support structural edits")]
    UnsupportedSource,
    /// The chart kind is undefined or not a concrete iWork kind.
    #[error("the chart kind must be a supported concrete iWork kind")]
    UnsupportedKind,
    /// The series direction is not a recognized orientation.
    #[error("the chart series direction must be rows or columns")]
    UnsupportedDirection,
    /// The frame is not finite with positive dimensions.
    #[error("the chart frame must be finite with positive dimensions")]
    InvalidGeometry,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource => RefusalKind::UnsupportedContent,
            Self::UnsupportedKind | Self::UnsupportedDirection | Self::InvalidGeometry => {
                RefusalKind::InvalidValue
            },
        }
    }
}

#[derive(Clone, Copy)]
enum Change {
    Kind(Kind),
    Direction(Direction),
    Geometry(Geometry),
}

#[derive(Clone, Copy)]
struct Operation {
    chart: u64,
    slide: crate::Position,
    position: crate::Position,
    change: Change,
}

/// One bounded chart-lifecycle edit staged against an immutable package.
///
/// ```no_run
/// use litchi_keynote::chart::lifecycle::Kind;
/// use litchi_keynote::{ChartSelector, Package, SlideSelector};
///
/// # fn main() -> Result<(), Box<dyn std::error::Error>> {
/// let package = Package::open("deck.key")?;
/// let mut edit = package.edit_slide_chart_lifecycle();
/// edit.set_kind(SlideSelector::index(0), ChartSelector::index(0), Kind::Bar2d)?;
/// let _commit = edit.commit()?;
/// # Ok(())
/// # }
/// ```
pub struct Edit<'a> {
    source: &'a Package,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Edit")
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    const fn new(source: &'a Package) -> Self {
        Self {
            source,
            operations: Vec::new(),
        }
    }

    fn stage(
        &mut self,
        slide: SlideSelector<'_>,
        chart: ChartSelector<'_>,
        change: Change,
    ) -> Result<&mut Self, Error> {
        let selection = select_chart(self.source, slide, chart, true).map_err(Error::Selection)?;
        self.operations.push(Operation {
            chart: selection.chart_identifier,
            slide: selection.slide_position,
            position: selection.chart_position,
            change,
        });
        Ok(self)
    }

    /// Stage a new native chart kind, keeping the chart's data (host
    /// `set_slide_chart_kind`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::UnsupportedKind`].
    pub fn set_kind<'s, 'c>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        chart: impl Into<ChartSelector<'c>>,
        kind: Kind,
    ) -> Result<&mut Self, Error> {
        if kind == Kind::Undefined || kind.is_unsupported() {
            return Err(Error::Refused(Refusal::UnsupportedKind));
        }
        self.stage(slide.into(), chart.into(), Change::Kind(kind))
    }

    /// Stage whether rows or columns form the chart's series (host
    /// `set_slide_chart_direction`).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::UnsupportedDirection`].
    pub fn set_direction<'s, 'c>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        chart: impl Into<ChartSelector<'c>>,
        direction: Direction,
    ) -> Result<&mut Self, Error> {
        if direction.is_unsupported() {
            return Err(Error::Refused(Refusal::UnsupportedDirection));
        }
        self.stage(slide.into(), chart.into(), Change::Direction(direction))
    }

    /// Stage a new frame for the chart (host `set_slide_chart_geometry`). A
    /// frame without a rotation keeps the chart's rotation.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Refusal::InvalidGeometry`].
    pub fn set_geometry<'s, 'c>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        chart: impl Into<ChartSelector<'c>>,
        geometry: Geometry,
    ) -> Result<&mut Self, Error> {
        if !geometry.is_valid() {
            return Err(Error::Refused(Refusal::InvalidGeometry));
        }
        self.stage(slide.into(), chart.into(), Change::Geometry(geometry))
    }

    /// Apply every staged operation, publish one candidate, reopen it and
    /// read the touched charts back.
    ///
    /// # Errors
    ///
    /// Returns [`Error::Verification`] when the candidate does not read back
    /// as staged and typed source errors.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_operations(self.source, &self.operations)
    }
}

fn lift(_error: StructureError) -> Error {
    Error::InvalidSource
}

fn physical_source(package: &Package) -> Result<Arc<[u8]>, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) if source.source_is_exact() => Ok(source.shared_source()),
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => {
            Err(Error::Refused(Refusal::UnsupportedSource))
        },
    }
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::Objects,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => Error::InvalidSource,
    }
}

/// Edit the chart drawable payload's nested message at `path`.
fn edit_nested(
    payload: &[u8],
    path: &[u32],
    wire: litchi_iwa_common::WireLimits,
    edit: &mut dyn FnMut(&mut Fields) -> Result<(), StructureError>,
) -> Result<Vec<u8>, StructureError> {
    super::slide_drawable::edit_nested_payload(payload, path, wire, edit)
}

fn chart_values(
    payload: &[u8],
    wire: litchi_iwa_common::WireLimits,
) -> Result<(Option<u64>, Option<u64>, Option<Geometry>), Error> {
    let fields = Fields::parse(payload, wire).map_err(lift)?;
    let unity = fields.unique(CHART_UNITY).map_err(lift)?;
    let (kind, direction) = match unity {
        Some(unity) => {
            let unity = Fields::parse(unity.message().map_err(lift)?, wire).map_err(lift)?;
            let value = |number| -> Result<Option<u64>, Error> {
                unity
                    .unique(number)
                    .map_err(lift)?
                    .map(|field| field.varint().map_err(lift))
                    .transpose()
            };
            (value(CHART_TYPE)?, value(CHART_SERIES_DIRECTION)?)
        },
        None => (None, None),
    };
    let geometry = match fields.unique(CHART_DRAWABLE).map_err(lift)? {
        Some(drawable) => {
            super::slide_drawable::read_geometry(drawable.message().map_err(lift)?, wire)
                .map_err(lift)?
        },
        None => None,
    };
    Ok((kind, direction, geometry))
}

#[allow(clippy::cast_sign_loss, reason = "protobuf enum varints")]
fn native(value: i32) -> u64 {
    i64::from(value) as u64
}

fn commit_operations(source: &Package, operations: &[Operation]) -> Result<Commit, Error> {
    let source_bytes = physical_source(source)?;
    if operations.is_empty() {
        return Ok(Commit {
            package: source.clone(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
            },
            diagnostics: Diagnostics::default(),
        });
    }
    let mut workspace = Workspace::new(source).map_err(lift)?;
    let wire = workspace.wire;
    for operation in operations {
        let payload = workspace
            .payload(operation.chart, CHART_MESSAGE_TYPE)
            .map_err(lift)?
            .to_vec();
        let data = match operation.change {
            Change::Kind(kind) => edit_nested(&payload, &[CHART_UNITY], wire, &mut |fields| {
                fields.set_varint(CHART_TYPE, native(kind.native_value()));
                Ok(())
            }),
            Change::Direction(direction) => {
                edit_nested(&payload, &[CHART_UNITY], wire, &mut |fields| {
                    fields.set_varint(CHART_SERIES_DIRECTION, native(direction.native_value()));
                    Ok(())
                })
            },
            Change::Geometry(geometry) => {
                let current = chart_values(&payload, wire)?.2;
                let geometry = Geometry {
                    rotation: geometry
                        .rotation
                        .or(current.and_then(|current| current.rotation)),
                    ..geometry
                };
                edit_nested(
                    &payload,
                    &[CHART_DRAWABLE, DRAWABLE_GEOMETRY],
                    wire,
                    &mut |fields| {
                        super::slide_drawable::patch_geometry_fields(fields, geometry, wire)
                    },
                )
            },
        }
        .map_err(lift)?;
        super::slide_drawable::replace_message_payload(
            &mut workspace,
            operation.chart,
            CHART_MESSAGE_TYPE,
            data,
        )
        .map_err(lift)?;
    }
    let bytes = workspace.publish().map_err(lift)?;
    let candidate = source.structure_reopen(bytes).map_err(map_read_error)?;
    for operation in operations {
        let selection = select_chart(
            &candidate,
            SlideSelector::position(operation.slide),
            ChartSelector::index(operation.position.get()),
            true,
        )
        .map_err(|_error| Error::Verification)?;
        let object = candidate
            .object(selection.chart_identifier)
            .ok_or(Error::Verification)?;
        let payload = object
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .ok_or(Error::Verification)?;
        let (kind, direction, geometry) = chart_values(&payload.data, wire)?;
        let ok = match operation.change {
            Change::Kind(expected) => kind == Some(native(expected.native_value())),
            Change::Direction(expected) => direction == Some(native(expected.native_value())),
            Change::Geometry(expected) => geometry.is_some_and(|geometry| {
                (geometry.x, geometry.y, geometry.width, geometry.height)
                    == (expected.x, expected.y, expected.width, expected.height)
            }),
        };
        if !ok {
            return Err(Error::Verification);
        }
    }
    let target = physical_source(&candidate)?;
    let patch = Patch {
        artifacts: ExactArtifacts::new(source_bytes, target),
    };
    Ok(Commit {
        package: candidate,
        diagnostics: Diagnostics {
            changed: !patch.is_noop(),
            operations: operations.len(),
        },
        patch,
    })
}

/// Exact-source-checked reversible chart-lifecycle patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Return whether source and target are byte-identical.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Compact evidence for one verified chart-lifecycle commit.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    operations: usize,
}

impl Diagnostics {
    /// Return whether the package changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return how many operations were applied.
    #[must_use]
    pub const fn operations(self) -> usize {
        self.operations
    }
}

/// The verified result of one chart-lifecycle edit.
#[must_use]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the committed package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Take the committed package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the commit diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }

    /// Split the commit into its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
}

impl Package {
    /// Start one chart-lifecycle edit.
    #[must_use]
    pub const fn edit_slide_chart_lifecycle(&self) -> Edit<'_> {
        Edit::new(self)
    }

    /// Read one chart's series direction.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors and [`Error::InvalidSource`].
    pub fn slide_chart_direction<'s, 'c>(
        &self,
        slide: impl Into<SlideSelector<'s>>,
        chart: impl Into<ChartSelector<'c>>,
    ) -> Result<Direction, Error> {
        let selection =
            select_chart(self, slide.into(), chart.into(), false).map_err(Error::Selection)?;
        let wire = self.wire_limits().map_err(|_| Error::InvalidSource)?;
        let object = self
            .object(selection.chart_identifier)
            .ok_or(Error::InvalidSource)?;
        let payload = object
            .messages
            .iter()
            .find(|message| message.type_ == CHART_MESSAGE_TYPE)
            .ok_or(Error::InvalidSource)?;
        let direction = chart_values(&payload.data, wire)?.1.unwrap_or(0);
        #[allow(clippy::cast_possible_truncation, reason = "protobuf enum varints")]
        Ok(Direction::from_native(direction as i32))
    }

    /// Apply a chart-lifecycle patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when this package is not the
    /// patch's exact source.
    pub fn apply_slide_chart_lifecycle(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = physical_source(self)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        let candidate = self
            .structure_reopen(patch.artifacts.target().to_vec())
            .map_err(map_read_error)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                operations: 0,
            },
        })
    }
}
