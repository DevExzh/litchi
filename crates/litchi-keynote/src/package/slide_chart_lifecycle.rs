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
const CHART_GRID: u32 = 7;
const CHART_IS_DIRTY: u32 = 24;

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
    /// The data grid cannot be encoded (labels or values out of range).
    #[error("the chart data grid cannot be encoded")]
    InvalidData,
    /// The chart's graph is referenced from outside the chart.
    #[error("the Keynote chart is referenced from outside its private graph")]
    SharedGraph,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource | Self::SharedGraph => RefusalKind::UnsupportedContent,
            Self::UnsupportedKind
            | Self::UnsupportedDirection
            | Self::InvalidGeometry
            | Self::InvalidData => RefusalKind::InvalidValue,
        }
    }
}

#[derive(Clone)]
enum Change {
    Data(litchi_iwa_common::chart::data::ChartData),
    Kind(Kind),
    Direction(Direction),
    Geometry(Geometry),
    Remove,
    Duplicate,
}

#[derive(Clone)]
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

    /// Stage the removal of one chart with its private title, caption and
    /// styles (host `remove_slide_chart`): the chart's styles leave the
    /// document stylesheet, its private preset leaves the theme, and data
    /// the chart was the final owner of is culled.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors; [`Edit::commit`] refuses a chart
    /// referenced from outside its private graph.
    pub fn remove<'s, 'c>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        chart: impl Into<ChartSelector<'c>>,
    ) -> Result<&mut Self, Error> {
        self.stage(slide.into(), chart.into(), Change::Remove)
    }

    /// Stage a complete replacement of the chart's inline data grid, its
    /// labels and dimensions included (host `set_slide_chart_data`; a
    /// same-shape value update is also available through
    /// [`crate::Package::edit_slide_chart_data`]).
    ///
    /// # Errors
    ///
    /// Returns typed selector errors.
    pub fn set_data<'s, 'c>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        chart: impl Into<ChartSelector<'c>>,
        data: litchi_iwa_common::chart::data::ChartData,
    ) -> Result<&mut Self, Error> {
        self.stage(slide.into(), chart.into(), Change::Data(data))
    }

    /// Stage a copy of one chart on its slide, offset by ten points (host
    /// `duplicate_slide_chart`): the copy gets fresh drawable, title and
    /// caption, style and preset identities, registered like the source's,
    /// and its own data grid and frame.
    ///
    /// # Errors
    ///
    /// Returns typed selector errors.
    pub fn duplicate<'s, 'c>(
        &mut self,
        slide: impl Into<SlideSelector<'s>>,
        chart: impl Into<ChartSelector<'c>>,
    ) -> Result<&mut Self, Error> {
        self.stage(slide.into(), chart.into(), Change::Duplicate)
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
    let mut removed = 0;
    for operation in operations {
        if matches!(operation.change, Change::Remove) {
            remove_chart(&mut workspace, source, operation.chart)?;
            removed += 1;
            continue;
        }
        if matches!(operation.change, Change::Duplicate) {
            duplicate_chart(&mut workspace, source, operation.chart)?;
            continue;
        }
        let payload = workspace
            .payload(operation.chart, CHART_MESSAGE_TYPE)
            .map_err(lift)?
            .to_vec();
        let data = match &operation.change {
            Change::Remove | Change::Duplicate => unreachable!("handled above"),
            Change::Data(data) => {
                let request =
                    litchi_iwa_protos::chart_grid_creation_codec::ChartGridCreationRequest::new(
                        data.row_names(),
                        data.column_names(),
                        data.values(),
                        operation.chart,
                    );
                let options =
                    litchi_iwa_protos::chart_grid_creation_codec::EncodeOptions::for_request(
                        &request,
                    );
                let grid = litchi_iwa_protos::chart_grid_creation_codec::encode_chart_grid(
                    request, options,
                )
                .map_err(|_error| Error::Refused(Refusal::InvalidData))?;
                edit_nested(&payload, &[CHART_UNITY], wire, &mut |fields| {
                    fields.set_message(CHART_GRID, grid.bytes());
                    fields.set_varint(CHART_IS_DIRTY, 0);
                    Ok(())
                })
            },
            &Change::Kind(kind) => edit_nested(&payload, &[CHART_UNITY], wire, &mut |fields| {
                fields.set_varint(CHART_TYPE, native(kind.native_value()));
                Ok(())
            }),
            &Change::Direction(direction) => {
                edit_nested(&payload, &[CHART_UNITY], wire, &mut |fields| {
                    fields.set_varint(CHART_SERIES_DIRECTION, native(direction.native_value()));
                    Ok(())
                })
            },
            &Change::Geometry(geometry) => {
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
    if removed > 0 {
        workspace.release_removed_identifier_suffix();
    }
    let bytes = workspace.publish().map_err(lift)?;
    let candidate = source.structure_reopen(bytes).map_err(map_read_error)?;
    for operation in operations
        .iter()
        .filter(|operation| !matches!(operation.change, Change::Remove))
    {
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
        let ok = match &operation.change {
            Change::Data(expected) => {
                candidate
                    .slide_chart_data(
                        SlideSelector::position(operation.slide),
                        ChartSelector::index(operation.position.get()),
                    )
                    .map_err(|_error| Error::Verification)?
                    == *expected
            },
            &Change::Kind(expected) => kind == Some(native(expected.native_value())),
            &Change::Direction(expected) => direction == Some(native(expected.native_value())),
            &Change::Geometry(expected) => geometry.is_some_and(|geometry| {
                (geometry.x, geometry.y, geometry.width, geometry.height)
                    == (expected.x, expected.y, expected.width, expected.height)
            }),
            Change::Remove | Change::Duplicate => true,
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

const SHOW_MESSAGE_TYPE: u32 = 2;
const SHOW_THEME: u32 = 2;
const THEME_MESSAGE_TYPE: u32 = 10;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const THEME_SUPER: u32 = 1;
const THEME_DOCUMENT_STYLESHEET: u32 = 4;
const THEME_CHART_PRESETS: u32 = 120;
const STYLESHEET_STYLES: u32 = 1;

fn reference_of(
    fields: &Fields,
    number: u32,
    wire: litchi_iwa_common::WireLimits,
) -> Result<Option<u64>, Error> {
    fields
        .unique(number)
        .map_err(lift)?
        .map(|field| {
            super::structure::wire::reference_identifier(field.message().map_err(lift)?, wire)
                .map_err(lift)
        })
        .transpose()
}

/// Remove a repeated reference field's entries naming any of `removed`,
/// keeping the header aggregate in step.
fn drop_references(
    workspace: &mut Workspace<'_>,
    object: u64,
    message_type: u32,
    path: &[u32],
    number: u32,
    removed: &std::collections::BTreeSet<u64>,
) -> Result<usize, Error> {
    let wire = workspace.wire;
    let payload = workspace
        .payload(object, message_type)
        .map_err(lift)?
        .to_vec();
    let mut count = 0;
    let data = edit_nested(&payload, path, wire, &mut |fields| {
        let mut kept = Vec::new();
        for field in fields.all(number) {
            let raw = field.message()?.to_vec();
            if removed.contains(&super::structure::wire::reference_identifier(&raw, wire)?) {
                count += 1;
            } else {
                kept.push(raw);
            }
        }
        if count > 0 {
            fields.replace_all(number, super::structure::wire::LENGTH_DELIMITED, &kept);
        }
        Ok(())
    })
    .map_err(lift)?;
    if count == 0 {
        return Ok(0);
    }
    super::slide_drawable::replace_message_payload(workspace, object, message_type, data)
        .map_err(lift)?;
    let target = workspace.object_mut(object).map_err(lift)?;
    let index = target
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or(Error::InvalidSource)?;
    let info = &mut target.archive_info.message_infos[index];
    info.object_references.retain(|id| !removed.contains(id));
    for field in &mut info.field_infos {
        field.object_references.retain(|id| !removed.contains(id));
    }
    Ok(count)
}

/// The chart's private graph inside its component, in the host's
/// `chart_graph` order: the drawable, its caption graph, its title
/// stand-in, then the chart's styles in `TSCH.ChartArchive` field order;
/// anything else of the closure follows in identifier order.
fn chart_graph(workspace: &Workspace<'_>, chart: u64) -> Result<(u64, Vec<u64>), Error> {
    let wire = workspace.wire;
    let member = workspace.component_of(chart).ok_or(Error::InvalidSource)?;
    let internal = workspace
        .member_objects(&member)
        .into_iter()
        .collect::<std::collections::BTreeSet<_>>();
    let payload = workspace.payload(chart, CHART_MESSAGE_TYPE).map_err(lift)?;
    let fields = Fields::parse(payload, wire).map_err(lift)?;
    let drawable = fields
        .unique(CHART_DRAWABLE)
        .map_err(lift)?
        .ok_or(Error::InvalidSource)?;
    let drawable = Fields::parse(drawable.message().map_err(lift)?, wire).map_err(lift)?;
    let parent = reference_of(&drawable, 2, wire)?.ok_or(Error::InvalidSource)?;
    let reachable = |root: u64| -> Result<Vec<u64>, Error> {
        let mut found = vec![root];
        let mut seen = std::collections::BTreeSet::from([root, parent]);
        let mut index = 0;
        while index < found.len() {
            for target in
                super::slide_drawable::outgoing_references(workspace, found[index]).map_err(lift)?
            {
                if internal.contains(&target) && seen.insert(target) {
                    found.push(target);
                }
            }
            index += 1;
        }
        Ok(found)
    };
    let closure = reachable(chart)?
        .into_iter()
        .collect::<std::collections::BTreeSet<_>>();
    let mut ordered = vec![chart];
    let push = |id: u64, ordered: &mut Vec<u64>| {
        if closure.contains(&id) && !ordered.contains(&id) {
            ordered.push(id);
        }
    };
    if let Some(caption) = reference_of(&drawable, 11, wire)? {
        if internal.contains(&caption) {
            for id in reachable(caption)? {
                push(id, &mut ordered);
            }
        }
    }
    if let Some(title) = reference_of(&drawable, 10, wire)? {
        push(title, &mut ordered);
    }
    if let Some(unity) = fields.unique(CHART_UNITY).map_err(lift)? {
        let unity = Fields::parse(unity.message().map_err(lift)?, wire).map_err(lift)?;
        for number in [4, 9, 10, 11, 12, 13, 14, 15, 16, 17] {
            for field in unity.all(number) {
                let id = super::structure::wire::reference_identifier(
                    field.message().map_err(lift)?,
                    wire,
                )
                .map_err(lift)?;
                push(id, &mut ordered);
            }
        }
        for number in [18, 19] {
            if let Some(sparse) = unity.unique(number).map_err(lift)? {
                let sparse = Fields::parse(sparse.message().map_err(lift)?, wire).map_err(lift)?;
                for entry in sparse.all(2) {
                    let entry =
                        Fields::parse(entry.message().map_err(lift)?, wire).map_err(lift)?;
                    if let Some(id) = reference_of(&entry, 2, wire)? {
                        push(id, &mut ordered);
                    }
                }
            }
        }
    }
    for id in &closure {
        push(*id, &mut ordered);
    }
    Ok((parent, ordered))
}

/// The theme, its chart-preset list and the document stylesheet.
fn theme_and_stylesheet(
    workspace: &Workspace<'_>,
    source: &Package,
) -> Result<(u64, Vec<u64>, u64, Vec<u64>), Error> {
    let wire = workspace.wire;
    let show = source
        .root_show_identifier()
        .map_err(|_| Error::InvalidSource)?;
    let show = Fields::parse(
        workspace.payload(show, SHOW_MESSAGE_TYPE).map_err(lift)?,
        wire,
    )
    .map_err(lift)?;
    let theme = reference_of(&show, SHOW_THEME, wire)?.ok_or(Error::InvalidSource)?;
    let theme_fields = Fields::parse(
        workspace.payload(theme, THEME_MESSAGE_TYPE).map_err(lift)?,
        wire,
    )
    .map_err(lift)?;
    let theme_super = theme_fields
        .unique(THEME_SUPER)
        .map_err(lift)?
        .ok_or(Error::InvalidSource)?;
    let theme_super = Fields::parse(theme_super.message().map_err(lift)?, wire).map_err(lift)?;
    let stylesheet =
        reference_of(&theme_super, THEME_DOCUMENT_STYLESHEET, wire)?.ok_or(Error::InvalidSource)?;
    let references = |fields: &Fields, number| -> Result<Vec<u64>, Error> {
        fields
            .all(number)
            .map(|field| {
                super::structure::wire::reference_identifier(field.message().map_err(lift)?, wire)
                    .map_err(lift)
            })
            .collect()
    };
    let presets = match theme_super.unique(THEME_CHART_PRESETS).map_err(lift)? {
        Some(presets) => references(
            &Fields::parse(presets.message().map_err(lift)?, wire).map_err(lift)?,
            1,
        )?,
        None => Vec::new(),
    };
    let styles = references(
        &Fields::parse(
            workspace
                .payload(stylesheet, STYLESHEET_MESSAGE_TYPE)
                .map_err(lift)?,
            wire,
        )
        .map_err(lift)?,
        STYLESHEET_STYLES,
    )?;
    Ok((theme, presets, stylesheet, styles))
}

/// Append references to a repeated field at `path` and to the header.
fn append_references(
    workspace: &mut Workspace<'_>,
    object: u64,
    message_type: u32,
    path: &[u32],
    number: u32,
    added: &[u64],
) -> Result<(), Error> {
    if added.is_empty() {
        return Ok(());
    }
    let wire = workspace.wire;
    let payload = workspace
        .payload(object, message_type)
        .map_err(lift)?
        .to_vec();
    let data = edit_nested(&payload, path, wire, &mut |fields| {
        let mut values = fields
            .all(number)
            .map(|field| field.message().map(<[u8]>::to_vec))
            .collect::<Result<Vec<_>, _>>()?;
        values.extend(
            added
                .iter()
                .map(|id| super::structure::wire::reference(*id)),
        );
        fields.replace_all(number, super::structure::wire::LENGTH_DELIMITED, &values);
        Ok(())
    })
    .map_err(lift)?;
    super::slide_drawable::replace_message_payload(workspace, object, message_type, data)
        .map_err(lift)?;
    let target = workspace.object_mut(object).map_err(lift)?;
    let index = target
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or(Error::InvalidSource)?;
    let info = &mut target.archive_info.message_infos[index];
    for id in added {
        if !info.object_references.contains(id) {
            info.object_references.push(*id);
        }
    }
    Ok(())
}

const DUPLICATE_OFFSET: f32 = 10.0;

/// Duplicate one chart (host `duplicate_slide_chart`).
fn duplicate_chart(
    workspace: &mut Workspace<'_>,
    source: &Package,
    chart: u64,
) -> Result<(), Error> {
    let wire = workspace.wire;
    let member = workspace.component_of(chart).ok_or(Error::InvalidSource)?;
    let (slide, objects) = chart_graph(workspace, chart)?;
    let mut remap = std::collections::HashMap::with_capacity(objects.len());
    for object in &objects {
        remap.insert(*object, workspace.allocate().map_err(lift)?);
    }
    for object in &objects {
        let original = workspace.require(*object).map_err(lift)?.clone();
        let clone = workspace
            .clone_object(&original, remap[object], &remap)
            .map_err(lift)?;
        workspace.insert(&member, clone).map_err(lift)?;
    }
    let copy = remap[&chart];
    // The copy's frame, offset like Keynote's Duplicate.
    let payload = workspace
        .payload(copy, CHART_MESSAGE_TYPE)
        .map_err(lift)?
        .to_vec();
    let geometry = chart_values(&payload, wire)?
        .2
        .ok_or(Error::InvalidSource)?;
    let offset = Geometry {
        x: geometry.x + DUPLICATE_OFFSET,
        y: geometry.y + DUPLICATE_OFFSET,
        ..geometry
    };
    let data = edit_nested(
        &payload,
        &[CHART_DRAWABLE, DRAWABLE_GEOMETRY],
        wire,
        &mut |fields| super::slide_drawable::patch_geometry_fields(fields, offset, wire),
    )
    .map_err(lift)?;
    super::slide_drawable::replace_message_payload(workspace, copy, CHART_MESSAGE_TYPE, data)
        .map_err(lift)?;
    super::slide_drawable::attach_drawable(workspace, slide, copy).map_err(lift)?;
    // Styles and preset registered like the source's.
    let (theme, presets, stylesheet, styles) = theme_and_stylesheet(workspace, source)?;
    let new_styles = objects
        .iter()
        .filter(|object| styles.contains(object))
        .map(|object| remap[object])
        .collect::<Vec<_>>();
    append_references(
        workspace,
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        &[],
        STYLESHEET_STYLES,
        &new_styles,
    )?;
    let new_presets = objects
        .iter()
        .filter(|object| presets.contains(object))
        .map(|object| remap[object])
        .collect::<Vec<_>>();
    append_references(
        workspace,
        theme,
        THEME_MESSAGE_TYPE,
        &[THEME_SUPER, THEME_CHART_PRESETS],
        1,
        &new_presets,
    )?;
    // Registry: external references, UUIDs and data owners of the copy.
    let component_of = |workspace: &Workspace<'_>, object: u64| -> Result<u64, Error> {
        let member = workspace.component_of(object).ok_or(Error::InvalidSource)?;
        Ok(workspace
            .registry
            .component_for_member(&member)
            .ok_or(Error::InvalidSource)?
            .identifier)
    };
    let owner = component_of(workspace, chart)?;
    let stylesheet_component = component_of(workspace, stylesheet)?;
    let theme_component = component_of(workspace, theme)?;
    for (component, added) in [
        (stylesheet_component, &new_styles),
        (theme_component, &new_presets),
    ] {
        if component == owner {
            continue;
        }
        let entry = workspace
            .registry
            .component_mut(component)
            .ok_or(Error::InvalidSource)?;
        for id in added.iter() {
            entry.add_external_reference(super::structure::metadata::ExternalReference {
                component: owner,
                object: Some(*id),
                is_weak: None,
            });
        }
    }
    let registered = workspace
        .registry
        .component(owner)
        .ok_or(Error::InvalidSource)?
        .clone();
    let mut fresh = Vec::new();
    for object in &objects {
        if let Some(uuid) = registered.uuid_of(*object) {
            fresh.push((
                remap[object],
                workspace
                    .clone_uuid(Some(uuid), remap[object])
                    .map_err(lift)?,
            ));
        }
    }
    let data = registered.data_owned_by(&objects);
    let entry = workspace
        .registry
        .component_mut(owner)
        .ok_or(Error::InvalidSource)?;
    for (object, uuid) in fresh {
        entry.add_object_uuid(object, uuid);
    }
    for datum in data {
        for object in &objects {
            if registered.data_owned_by(&[*object]).contains(&datum) {
                entry.add_data_owner(datum, remap[object]);
            }
        }
    }
    Ok(())
}

/// Remove one chart (host `remove_slide_chart`).
fn remove_chart(workspace: &mut Workspace<'_>, source: &Package, chart: u64) -> Result<(), Error> {
    let wire = workspace.wire;
    let member = workspace.component_of(chart).ok_or(Error::InvalidSource)?;
    // The chart's private graph: everything reachable inside its component.
    let internal = workspace
        .member_objects(&member)
        .into_iter()
        .collect::<std::collections::BTreeSet<_>>();
    let parent = {
        let payload = workspace.payload(chart, CHART_MESSAGE_TYPE).map_err(lift)?;
        let fields = Fields::parse(payload, wire).map_err(lift)?;
        let drawable = fields
            .unique(CHART_DRAWABLE)
            .map_err(lift)?
            .ok_or(Error::InvalidSource)?;
        let drawable = Fields::parse(drawable.message().map_err(lift)?, wire).map_err(lift)?;
        reference_of(&drawable, 2, wire)?.ok_or(Error::InvalidSource)?
    };
    let mut closure = std::collections::BTreeSet::from([chart]);
    let mut pending = vec![chart];
    while let Some(current) = pending.pop() {
        for target in
            super::slide_drawable::outgoing_references(workspace, current).map_err(lift)?
        {
            if target != parent && internal.contains(&target) && closure.insert(target) {
                pending.push(target);
            }
        }
    }
    // Theme presets and the document stylesheet's style list.
    let show = source
        .root_show_identifier()
        .map_err(|_| Error::InvalidSource)?;
    let show = Fields::parse(
        workspace.payload(show, SHOW_MESSAGE_TYPE).map_err(lift)?,
        wire,
    )
    .map_err(lift)?;
    let theme = reference_of(&show, SHOW_THEME, wire)?.ok_or(Error::InvalidSource)?;
    let theme_payload = workspace
        .payload(theme, THEME_MESSAGE_TYPE)
        .map_err(lift)?
        .to_vec();
    let theme_fields = Fields::parse(&theme_payload, wire).map_err(lift)?;
    let theme_super = theme_fields
        .unique(THEME_SUPER)
        .map_err(lift)?
        .ok_or(Error::InvalidSource)?;
    let theme_super = Fields::parse(theme_super.message().map_err(lift)?, wire).map_err(lift)?;
    let stylesheet =
        reference_of(&theme_super, THEME_DOCUMENT_STYLESHEET, wire)?.ok_or(Error::InvalidSource)?;
    drop_references(
        workspace,
        stylesheet,
        STYLESHEET_MESSAGE_TYPE,
        &[],
        STYLESHEET_STYLES,
        &closure,
    )?;
    if theme_super
        .unique(THEME_CHART_PRESETS)
        .map_err(lift)?
        .is_some()
    {
        drop_references(
            workspace,
            theme,
            THEME_MESSAGE_TYPE,
            &[THEME_SUPER, THEME_CHART_PRESETS],
            1,
            &closure,
        )?;
    }
    // Every component's external references to the chart's objects.
    let component = workspace
        .registry
        .component_for_member(&member)
        .ok_or(Error::InvalidSource)?
        .identifier;
    let others = workspace
        .registry
        .components()
        .filter(|other| {
            other.external_references().iter().any(|reference| {
                reference.component == component
                    && reference
                        .object
                        .is_some_and(|object| closure.contains(&object))
            })
        })
        .map(|other| other.identifier)
        .collect::<Vec<_>>();
    for other in others {
        let entry = workspace
            .registry
            .component_mut(other)
            .ok_or(Error::InvalidSource)?;
        let kept = entry
            .external_references()
            .iter()
            .filter(|reference| {
                !(reference.component == component
                    && reference
                        .object
                        .is_some_and(|object| closure.contains(&object)))
            })
            .copied()
            .collect();
        entry.set_external_references(kept);
    }
    super::slide_drawable::remove_drawable_closure(workspace, parent, chart).map_err(|error| {
        match error {
            super::slide_drawable::Error::Refused(_) => Error::Refused(Refusal::SharedGraph),
            _ => Error::InvalidSource,
        }
    })?;
    Ok(())
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
