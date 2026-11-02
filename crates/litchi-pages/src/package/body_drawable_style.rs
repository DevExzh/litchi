//! Drawing styles of body shapes (record 0371; the Pages leaf of the host's
//! `{,set_,reset_}body_shape_{fill,stroke,shadow,effects,text_layout}`) over
//! the shared shape-style variation engine (`litchi_iwa_style::shape`,
//! record 0375a).
//!
//! A shape's own value lives in a copy-on-write variation of its
//! `TSWP.ShapeStyleArchive`. The engine plans and applies each write on a
//! workspace over this package's parsed components: an exclusive variation
//! is rewritten in place, a shared style gets a new child variation (stored
//! beside its stylesheet, registered there, given a UUID, the shape pointed
//! at it), and a variation that no longer overrides anything collapses into
//! its parent. This module selects the shape, publishes the workspace's
//! changes once and reopens.

use std::fmt;
use std::sync::Arc;

use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::{
    Error as WireError, LimitKind as WireLimitKind, WireLimits, decode_varint_from_bytes,
};
use litchi_iwa_core::Archive;
use litchi_iwa_style::graph::{Components, Workspace};
use litchi_iwa_style::shape::{
    self as engine,
    property::{Fills, Shadows, ShapeEffects, Strokes, TextColumns, TextLayout},
};
use thiserror::Error as ThisError;

use litchi_iwa_common::shape::fill::ImageTechnique;
use litchi_iwa_common::shape::geometry::Size as FillSize;

use super::Package;
use super::body_drawable_graph::{self as graph, Budget, GraphError, ObjectIndex, kind_name};
use crate::drawable::style::{Endpoints, Fill, Property, PropertyKind};
use crate::selector::DrawableSelector;

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The selector resolved no drawable.
    #[error("the Pages body has no {kind} at position {}", position.get())]
    NotFound {
        /// The selected kind.
        kind: &'static str,
        /// The requested position.
        position: litchi_core::Position,
    },
    /// Only ordinary body shapes carry these styles.
    #[error("only body shapes carry drawing styles")]
    NotAShape,
    /// An image fill is replaced or reset, or a second image fill is
    /// staged in one edit (its data reference and media are not released
    /// here).
    #[error("an existing image fill cannot be replaced here")]
    ImageFill,
    /// The image-fill bytes are not an image, or the file name is not a
    /// safe basename.
    #[error("the image-fill data or file name is unusable")]
    InvalidImage,
    /// The shape's style graph holds overrides this family cannot rewrite.
    #[error("the shape's style cannot be rewritten safely")]
    UnsupportedStyle,
    /// Nothing was staged.
    #[error("no drawing-style change is staged")]
    NoStagedOperation,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact drawing-style edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound { .. } => RefusalKind::NotFound,
            Self::NotAShape | Self::NoStagedOperation => RefusalKind::UnsupportedOperation,
            Self::ImageFill | Self::UnsupportedStyle | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
            Self::InvalidImage => RefusalKind::InvalidValue,
        }
    }
}

/// A finite resource charged by one drawing-style edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected.
    InputBytes,
    /// Bytes in the rewritten payloads or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and rewrite work.
    WireWork,
    /// Semantic section or body-storage cardinality.
    PayloadItems,
    /// Semantic UTF-8 text retained by the reopened document.
    TextBytes,
    /// Temporary graph scratch bytes.
    ScratchBytes,
    /// Fallible graph allocation events.
    AllocationEvents,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::PayloadItems => "payload items",
            Self::TextBytes => "text bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::AllocationEvents => "allocation events",
        })
    }
}

/// Failure from a drawing-style read or edit.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The drawable or style graph is malformed.
    #[error("the selected Pages shape's style graph is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error("Pages drawing style {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for a Pages drawing-style edit")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the new style.
    #[error("the restyled Pages shape failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages drawing-style patch does not match the exact source package")]
    PatchConflict,
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

impl From<GraphError> for Error {
    fn from(error: GraphError) -> Self {
        match error {
            GraphError::InvalidSource => Self::InvalidSource,
            GraphError::Allocation { amount } => Self::Allocation { amount },
            GraphError::Limit {
                kind,
                observed,
                maximum,
            } => Self::Limit {
                kind: match kind {
                    graph::GraphLimit::InputBytes => LimitKind::InputBytes,
                    graph::GraphLimit::OutputBytes => LimitKind::OutputBytes,
                    graph::GraphLimit::WireFields => LimitKind::WireFields,
                    graph::GraphLimit::WireNesting => LimitKind::WireNesting,
                    graph::GraphLimit::WireWork => LimitKind::WireWork,
                    graph::GraphLimit::SemanticItems => LimitKind::PayloadItems,
                    graph::GraphLimit::SemanticTextBytes => LimitKind::TextBytes,
                },
                observed,
                maximum,
            },
        }
    }
}

fn from_engine(error: litchi_iwa_style::graph::Error) -> Error {
    match error {
        litchi_iwa_style::graph::Error::Invalid(_) => Refusal::UnsupportedStyle.into(),
        litchi_iwa_style::graph::Error::Limit {
            kind: litchi_iwa_style::graph::GraphLimitKind::ScratchBytes,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::ScratchBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_style::graph::Error::Limit {
            kind: litchi_iwa_style::graph::GraphLimitKind::AllocationEvents,
            observed,
            maximum,
        } => Error::Limit {
            kind: LimitKind::AllocationEvents,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => Error::InvalidSource,
    }
}

/// This package's parsed components, as the style engine reads them.
struct PackageComponents<'a>(&'a Package);

impl Components for PackageComponents<'_> {
    fn count(&self) -> usize {
        self.0.state.source.components().len()
    }

    fn component(&self, index: usize) -> Option<(&str, &Archive)> {
        self.0
            .state
            .source
            .components()
            .get_index(index)
            .map(|component| (component.name(), component.archive()))
    }
}

/// The shape drawable `selector` names.
fn drawable(package: &Package, selector: DrawableSelector) -> Result<u64, Error> {
    if !matches!(
        selector,
        DrawableSelector::Shape(_) | DrawableSelector::TextBox(_)
    ) {
        return Err(Refusal::NotAShape.into());
    }
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let selected =
        graph::select_drawable(&index, &body, selector, &mut budget)?.ok_or(Refusal::NotFound {
            kind: kind_name(selector),
            position: selector.position(),
        })?;
    Ok(selected.anchor.drawable.get())
}

fn workspace(package: &Package) -> Result<(PackageComponents<'_>, litchi_iwa_core::Limits), Error> {
    let limits = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_error| Error::InvalidSource)?;
    Ok((PackageComponents(package), limits))
}

fn read(
    source: &Package,
    workspace: &Workspace<'_>,
    shape: u64,
    kind: PropertyKind,
) -> Result<Property, Error> {
    Ok(match kind {
        PropertyKind::Fill => {
            Property::Fill(engine::value(workspace, shape, &Fills).map_err(from_engine)?)
        },
        PropertyKind::Stroke => {
            Property::Stroke(engine::value(workspace, shape, &Strokes).map_err(from_engine)?)
        },
        PropertyKind::Shadow => {
            Property::Shadow(engine::value(workspace, shape, &Shadows).map_err(from_engine)?)
        },
        PropertyKind::Effects => {
            Property::Effects(engine::value(workspace, shape, &ShapeEffects).map_err(from_engine)?)
        },
        PropertyKind::TextLayout => {
            Property::TextLayout(engine::value(workspace, shape, &TextLayout).map_err(from_engine)?)
        },
        PropertyKind::TextColumns => Property::TextColumns(
            engine::value(workspace, shape, &TextColumns).map_err(from_engine)?,
        ),
        PropertyKind::LineEndpoints => {
            Property::LineEndpoints(endpoints(workspace, &line_ends(source, shape)?, shape)?)
        },
    })
}

// ---- Line endpoints (the host's `shapes::line_end`; the slot values are
// the shared `litchi_iwa_style::line_end` payloads) ----

const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const SHAPE_ARCHIVE: u32 = 1;
const SHAPE_PATH_SOURCE: u32 = 3;
const SHAPE_HEAD_LINE_END: u32 = 4;
const SHAPE_TAIL_LINE_END: u32 = 5;
const PATH_SOURCE_HORIZONTAL_FLIP: u32 = 1;

/// The shape's own (legacy per-shape) head and tail decorations and its
/// horizontal flip.
struct LineEnds {
    head: Option<Vec<u8>>,
    tail: Option<Vec<u8>>,
    flipped: bool,
    limits: WireLimits,
}

fn single<'a>(
    budget: &mut Budget,
    data: &'a [u8],
    number: u32,
    depth: usize,
) -> Result<Option<&'a [u8]>, Error> {
    let view = budget.parse(data, depth)?;
    let mut found = None;
    for field in view.fields() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 2 || found.replace(field.payload()).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    Ok(found)
}

fn last_varint(
    budget: &mut Budget,
    data: &[u8],
    number: u32,
    depth: usize,
) -> Result<Option<u64>, Error> {
    let view = budget.parse(data, depth)?;
    let mut found = None;
    for field in view.fields() {
        if field.number() != number {
            continue;
        }
        if field.wire_type() != 0 {
            return Err(Error::InvalidSource);
        }
        let (value, length) =
            decode_varint_from_bytes(field.payload()).map_err(|_error| Error::InvalidSource)?;
        if length != field.payload().len() {
            return Err(Error::InvalidSource);
        }
        found = Some(value);
    }
    Ok(found)
}

/// Own one nested line-end payload only after charging its complete wire
/// traversal and copy against the body-drawable budget.  The endpoint codec
/// receives this owned payload later; it must not be able to turn a borrowed
/// native field into an uncharged allocation.
fn owned_payload(budget: &mut Budget, payload: &[u8], depth: usize) -> Result<Vec<u8>, Error> {
    budget.parse(payload, depth)?;
    budget.work(payload.len())?;
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(payload.len())
        .map_err(|_| Error::Allocation {
            amount: payload.len(),
        })?;
    owned.extend_from_slice(payload);
    Ok(owned)
}

fn line_ends(package: &Package, shape: u64) -> Result<LineEnds, Error> {
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let located = index.get(std::num::NonZeroU64::new(shape).ok_or(Error::InvalidSource)?)?;
    let info = located
        .object
        .messages
        .iter()
        .find(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE)
        .ok_or(Refusal::NotAShape)?;
    let archive = single(&mut budget, &info.data, SHAPE_ARCHIVE, 0)?.ok_or(Error::InvalidSource)?;
    let flipped = match single(&mut budget, archive, SHAPE_PATH_SOURCE, 1)? {
        Some(source) => last_varint(&mut budget, source, PATH_SOURCE_HORIZONTAL_FLIP, 2)?
            .is_some_and(|flip| flip != 0),
        None => false,
    };
    let head = match single(&mut budget, archive, SHAPE_HEAD_LINE_END, 1)? {
        Some(payload) => Some(owned_payload(&mut budget, payload, 2)?),
        None => None,
    };
    let tail = match single(&mut budget, archive, SHAPE_TAIL_LINE_END, 1)? {
        Some(payload) => Some(owned_payload(&mut budget, payload, 2)?),
        None => None,
    };
    Ok(LineEnds {
        head,
        tail,
        flipped,
        limits: budget.wire_limits(),
    })
}

fn decode_endpoint(
    payload: Option<&[u8]>,
    limits: WireLimits,
) -> Result<litchi_iwa_common::shape::line::Endpoint, Error> {
    litchi_iwa_style::line_end::decode(payload, limits).map_err(|error| match error {
        litchi_iwa_style::line_end::DecodeError::Unknown(_error) => {
            Refusal::UnsupportedStyle.into()
        },
        litchi_iwa_style::line_end::DecodeError::Wire(error) => match error {
            WireError::LimitExceeded {
                kind,
                observed,
                limit,
            } => match kind {
                WireLimitKind::InputBytes => Error::Limit {
                    kind: LimitKind::InputBytes,
                    observed: observed as u64,
                    maximum: limit as u64,
                },
                WireLimitKind::Fields => Error::Limit {
                    kind: LimitKind::WireFields,
                    observed: observed as u64,
                    maximum: limit as u64,
                },
                WireLimitKind::OutputBytes => Error::Limit {
                    kind: LimitKind::OutputBytes,
                    observed: observed as u64,
                    maximum: limit as u64,
                },
                WireLimitKind::Nesting => Error::Limit {
                    kind: LimitKind::WireNesting,
                    observed: observed as u64,
                    maximum: limit as u64,
                },
                WireLimitKind::RewriteWork => Error::Limit {
                    kind: LimitKind::WireWork,
                    observed: observed as u64,
                    maximum: limit as u64,
                },
                WireLimitKind::TableRows
                | WireLimitKind::TableColumns
                | WireLimitKind::TableCells
                | WireLimitKind::MaterializedCells => Error::InvalidSource,
            },
            WireError::Allocation { amount, .. } => Error::Allocation { amount },
            WireError::InvalidFormat(_) => Refusal::UnsupportedStyle.into(),
            WireError::InvalidLimit { .. } => Error::InvalidSource,
            _ => Error::InvalidSource,
        },
        _ => Error::InvalidSource,
    })
}

/// The effective head and tail of a style chain, each inherited on its own.
fn inherited_line_ends(
    workspace: &Workspace<'_>,
    style: u64,
) -> Result<(Option<Vec<u8>>, Option<Vec<u8>>), Error> {
    Ok((
        engine::effective(workspace, style, engine::Slot::HeadLineEnd).map_err(from_engine)?,
        engine::effective(workspace, style, engine::Slot::TailLineEnd).map_err(from_engine)?,
    ))
}

fn endpoints(workspace: &Workspace<'_>, own: &LineEnds, shape: u64) -> Result<Endpoints, Error> {
    let (mut head, mut tail) = (own.head.clone(), own.tail.clone());
    if head.is_none() || tail.is_none() {
        let style = engine::target(workspace, shape)
            .map_err(from_engine)?
            .style();
        let (inherited_head, inherited_tail) = inherited_line_ends(workspace, style)?;
        head = head.or(inherited_head);
        tail = tail.or(inherited_tail);
    }
    let mut endpoints = Endpoints {
        start: decode_endpoint(tail.as_deref(), own.limits)?,
        end: decode_endpoint(head.as_deref(), own.limits)?,
    };
    if own.flipped {
        std::mem::swap(&mut endpoints.start, &mut endpoints.end);
    }
    Ok(endpoints)
}

/// Store `wanted` as the line's decorations (the host's
/// `set_shape_line_endpoints`); returns whether anything changed.
fn set_endpoints(
    workspace: &mut Workspace<'_>,
    own: &LineEnds,
    shape: u64,
    wanted: Endpoints,
) -> Result<bool, Error> {
    if endpoints(workspace, own, shape)? == wanted {
        return Ok(false);
    }
    let mut stored = wanted;
    if own.flipped {
        std::mem::swap(&mut stored.start, &mut stored.end);
    }
    let target = engine::target(workspace, shape).map_err(from_engine)?;
    let payload = engine::Styles::shape_style(&*workspace, target.style())
        .ok_or(Error::InvalidSource)?
        .to_vec();
    let disposable = target.exclusive()
        && engine::direct_slots(&payload)
            .map_err(from_engine)?
            .is_some();
    let mut clear = false;
    if disposable && wanted == Endpoints::default() {
        let parent = engine::parent_of(&payload)
            .map_err(from_engine)?
            .ok_or(Refusal::UnsupportedStyle)?;
        let (head, tail) = inherited_line_ends(workspace, parent)?;
        let inherited = Endpoints {
            start: decode_endpoint(tail.as_deref(), own.limits)?,
            end: decode_endpoint(head.as_deref(), own.limits)?,
        };
        clear = inherited == Endpoints::default();
    }
    let writes = if clear {
        vec![
            (engine::Slot::HeadLineEnd, None),
            (engine::Slot::TailLineEnd, None),
        ]
    } else {
        vec![
            (
                engine::Slot::HeadLineEnd,
                Some(engine::Value::message(litchi_iwa_style::line_end::encode(
                    stored.end,
                ))),
            ),
            (
                engine::Slot::TailLineEnd,
                Some(engine::Value::message(litchi_iwa_style::line_end::encode(
                    stored.start,
                ))),
            ),
        ]
    };
    let plan = engine::plan_set(&*workspace, target, &writes).map_err(from_engine)?;
    engine::apply(workspace, shape, plan).map_err(from_engine)?;
    if endpoints(workspace, own, shape)? != wanted {
        return Err(Error::Verification);
    }
    Ok(true)
}

/// One staged change.
#[derive(Debug, Clone, PartialEq)]
enum Change {
    Set(Property),
    Reset(PropertyKind),
    ImageFill {
        bytes: Arc<[u8]>,
        file_name: Box<str>,
        technique: ImageTechnique,
        tint: Option<litchi_iwa_common::color::Rgba>,
    },
}

/// A new embedded image an image fill references (record 0413b).
struct NewImage {
    data: u64,
    bytes: Arc<[u8]>,
    file_name: Box<str>,
    stored_name: String,
}

/// Drawing-style changes staged against one body shape.
pub struct Edit<'a> {
    source: &'a Package,
    selector: DrawableSelector,
    shape: u64,
    changes: Vec<Change>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("selector", &self.selector)
            .field("changes", &self.changes.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage the shape's own value of one property (an image fill is
    /// refused).
    ///
    /// # Errors
    ///
    /// [`Refusal::ImageFill`] for an image fill.
    pub fn set(&mut self, property: Property) -> Result<&mut Self, Error> {
        if matches!(property, Property::Fill(Fill::Image(_))) {
            return Err(Refusal::ImageFill.into());
        }
        self.changes.push(Change::Set(property));
        Ok(self)
    }

    /// Stage embedding `bytes` as a new `Data/` member and making them the
    /// shape's or text box's image fill, laid out by `technique` over its
    /// frame size and optionally tinted (the host's
    /// `set_body_shape_image_fill`). The new data asset is owned by the
    /// shape's style variation.
    ///
    /// # Errors
    ///
    /// At commit: [`Refusal::InvalidImage`] for bytes that are not an image
    /// or an unsafe file name; [`Refusal::ImageFill`] when the shape
    /// already shows an image fill or a second one is staged.
    pub fn set_image_fill(
        &mut self,
        bytes: &[u8],
        file_name: &str,
        technique: ImageTechnique,
        tint: Option<litchi_iwa_common::color::Rgba>,
    ) -> &mut Self {
        self.changes.push(Change::ImageFill {
            bytes: bytes.into(),
            file_name: file_name.into(),
            technique,
            tint,
        });
        self
    }

    /// Stage removing the shape's own value of one property, so it shows its
    /// style's value again.
    pub fn reset(&mut self, kind: PropertyKind) -> &mut Self {
        self.changes.push(Change::Reset(kind));
        self
    }

    /// Apply every staged change in order, publish once and reopen.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] without a change; the engine's typed
    /// refusals; [`Error::Verification`] when the reopened shape does not
    /// show the changes.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.changes.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        if !self.source.state.source.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        let (components, limits) = workspace(self.source)?;
        let mut workspace = Workspace::new(&components, limits).map_err(from_engine)?;
        let mut applied = 0;
        let own = line_ends(self.source, self.shape)?;
        let mut image: Option<NewImage> = None;
        for change in &self.changes {
            let changed = match change {
                Change::ImageFill {
                    bytes,
                    file_name,
                    technique,
                    tint,
                } => {
                    if image.is_some()
                        || matches!(
                            engine::value(&workspace, self.shape, &Fills).map_err(from_engine)?,
                            Fill::Image(_)
                        )
                    {
                        return Err(Refusal::ImageFill.into());
                    }
                    let staged = stage_image(self.source, self.selector, bytes, file_name)?;
                    let size = self
                        .source
                        .body_drawable_properties(self.selector)
                        .map_err(|_| Error::InvalidSource)?
                        .geometry()
                        .and_then(|geometry| geometry.size)
                        .ok_or(Refusal::UnsupportedStyle)?;
                    let fill = embedded_fill(staged.data, *technique, size, *tint)?;
                    let writes = [(
                        engine::Slot::Fill,
                        Some(engine::Value::message(
                            litchi_iwa_style::drawing::encode_fill(&fill),
                        )),
                    )];
                    let target = engine::target(&workspace, self.shape).map_err(from_engine)?;
                    let plan =
                        engine::plan_set(&workspace, target, &writes).map_err(from_engine)?;
                    let changed = plan.changed();
                    engine::apply(&mut workspace, self.shape, plan).map_err(from_engine)?;
                    image = Some(staged);
                    Ok(changed)
                },
                Change::Set(Property::LineEndpoints(wanted)) => {
                    applied +=
                        usize::from(set_endpoints(&mut workspace, &own, self.shape, *wanted)?);
                    continue;
                },
                Change::Reset(PropertyKind::LineEndpoints) => {
                    applied += usize::from(set_endpoints(
                        &mut workspace,
                        &own,
                        self.shape,
                        Endpoints::default(),
                    )?);
                    continue;
                },
                Change::Set(Property::Fill(fill)) => {
                    engine::set(&mut workspace, self.shape, &Fills, fill)
                },
                Change::Set(Property::Stroke(stroke)) => {
                    engine::set(&mut workspace, self.shape, &Strokes, stroke)
                },
                Change::Set(Property::Shadow(shadow)) => {
                    engine::set(&mut workspace, self.shape, &Shadows, shadow)
                },
                Change::Set(Property::Effects(effects)) => {
                    engine::set(&mut workspace, self.shape, &ShapeEffects, effects)
                },
                Change::Set(Property::TextLayout(layout)) => {
                    engine::set(&mut workspace, self.shape, &TextLayout, layout)
                },
                Change::Set(Property::TextColumns(columns)) => {
                    engine::set(&mut workspace, self.shape, &TextColumns, columns)
                },
                Change::Reset(PropertyKind::Fill) => {
                    engine::reset(&mut workspace, self.shape, &Fills)
                },
                Change::Reset(PropertyKind::Stroke) => {
                    engine::reset(&mut workspace, self.shape, &Strokes)
                },
                Change::Reset(PropertyKind::Shadow) => {
                    engine::reset(&mut workspace, self.shape, &Shadows)
                },
                Change::Reset(PropertyKind::Effects) => {
                    engine::reset(&mut workspace, self.shape, &ShapeEffects)
                },
                Change::Reset(PropertyKind::TextLayout) => {
                    engine::reset(&mut workspace, self.shape, &TextLayout)
                },
                Change::Reset(PropertyKind::TextColumns) => {
                    engine::reset(&mut workspace, self.shape, &TextColumns)
                },
            }
            .map_err(from_engine)?;
            applied += usize::from(changed);
        }
        let expected = self
            .changes
            .iter()
            .filter_map(|change| match change {
                Change::Set(property) => Some(property.clone()),
                Change::Reset(_) | Change::ImageFill { .. } => None,
            })
            .collect::<Vec<_>>();
        let changes = workspace.finish();
        let package = if changes.is_empty() {
            self.source.clone()
        } else {
            if !changes.insertions.is_empty() || !changes.deletions.is_empty() {
                return Err(Refusal::UnsupportedStyle.into());
            }
            let mut archives: Vec<(Arc<str>, Archive)> = changes
                .edits
                .into_iter()
                .map(|(name, archive)| (Arc::<str>::from(name), archive))
                .collect();
            match &image {
                Some(image) => {
                    own_image(self.source, &mut archives, self.shape, image)?;
                    let member = format!("Data/{}", image.stored_name);
                    graph::publish_changes(
                        self.source,
                        archives,
                        &[(member.as_str(), image.bytes.as_ref())],
                        &[],
                        &[],
                    )?
                },
                None => graph::publish_changes(self.source, archives, &[], &[], &[])?,
            }
        };
        // The last staged value of each property reads back.
        let (components, limits) = self::workspace(&package)?;
        let reopened = Workspace::new(&components, limits).map_err(from_engine)?;
        let shape = drawable(&package, self.selector)?;
        for property in expected.iter().rev() {
            let later = expected
                .iter()
                .rposition(|other| other.kind() == property.kind())
                .map(|index| &expected[index]);
            if later == Some(property)
                && read(&package, &reopened, shape, property.kind())? != *property
            {
                return Err(Error::Verification);
            }
        }
        let target = package.state.source.shared_source();
        let changed = !Arc::ptr_eq(&target, &self.source.state.source.shared_source());
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(self.source.state.source.shared_source(), target),
            },
            diagnostics: Diagnostics { changed, applied },
        })
    }
}

/// Exact-source reversible drawing-style patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Patch").finish_non_exhaustive()
    }
}

impl Patch {
    /// Whether the exact artifact is unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
        }
    }
}

/// Content-free diagnostics of one drawing-style edit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    applied: usize,
}

impl Diagnostics {
    /// Whether the package changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// How many staged changes changed the shape.
    #[must_use]
    pub const fn applied(self) -> usize {
        self.applied
    }
}

/// The validated immutable result of one drawing-style edit.
#[must_use = "a drawing-style commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the reopened package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow the publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one drawing property of a body shape, inherited values
    /// included.
    ///
    /// # Errors
    ///
    /// [`Refusal::NotFound`] and [`Refusal::NotAShape`] for the selector, and
    /// the engine's typed refusals for an unsupported style graph.
    pub fn body_drawable_style(
        &self,
        selector: impl Into<DrawableSelector>,
        kind: PropertyKind,
    ) -> Result<Property, Error> {
        let shape = drawable(self, selector.into())?;
        let (components, limits) = workspace(self)?;
        let workspace = Workspace::new(&components, limits).map_err(from_engine)?;
        read(self, &workspace, shape, kind)
    }

    /// Begin staging drawing-style changes on one body shape.
    ///
    /// # Errors
    ///
    /// [`Refusal::NotFound`] and [`Refusal::NotAShape`] for the selector.
    pub fn edit_body_drawable_style(
        &self,
        selector: impl Into<DrawableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        let shape = drawable(self, selector)?;
        Ok(Edit {
            source: self,
            selector,
            shape,
            changes: Vec::new(),
        })
    }

    /// Replay an exact-source drawing-style patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact.
    pub fn apply_body_drawable_style(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        let package = if patch.is_noop() {
            self.clone()
        } else {
            graph::reopen(self, patch.artifacts.target())?
        };
        Ok(Commit {
            package,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: !patch.is_noop(),
                applied: usize::from(!patch.is_noop()),
            },
        })
    }
}

/// Validate an image fill's bytes and name and allocate its data asset
/// (the host media draft's `insert_unreferenced`: the next data
/// identifier, `stem-identifier.extension`).
fn stage_image(
    source: &Package,
    selector: DrawableSelector,
    bytes: &Arc<[u8]>,
    file_name: &str,
) -> Result<NewImage, Error> {
    let _ = selector;
    if litchi_iwa_common::media::Type::from_bytes(bytes) != litchi_iwa_common::media::Type::Image {
        return Err(Refusal::InvalidImage.into());
    }
    let metadata = super::body_storage_objects::metadata(source)?.ok_or(Error::InvalidSource)?;
    let data = super::body_storage_objects::max_data_identifier(&metadata.payload)?
        .checked_add(1)
        .ok_or(Error::InvalidSource)?;
    let stored_name = super::body_media_lifecycle::materialized_name(file_name, data)
        .map_err(|_| Refusal::InvalidImage)?;
    if source
        .state
        .source
        .package()
        .iter()
        .any(|entry| entry.name() == format!("Data/{stored_name}"))
    {
        return Err(Error::InvalidSource);
    }
    Ok(NewImage {
        data,
        bytes: Arc::clone(bytes),
        file_name: file_name.into(),
        stored_name,
    })
}

/// The embedded image fill of data `data` over a `size` frame.
fn embedded_fill(
    data: u64,
    technique: ImageTechnique,
    size: crate::drawable::Size,
    tint: Option<litchi_iwa_common::color::Rgba>,
) -> Result<Fill, Error> {
    let image = litchi_iwa_common::shape::fill::Image::embedded(
        litchi_iwa_common::shape::fill::ImageData::new(data).map_err(|_| Error::InvalidSource)?,
        technique,
        FillSize {
            width: size.width,
            height: size.height,
        },
    )
    .map_err(|_| Refusal::UnsupportedStyle)?;
    Ok(Fill::Image(match tint {
        Some(tint) => image.with_tint(tint),
        None => image,
    }))
}

/// Give the shape's (new) style variation the image's data reference and
/// register the asset: its `DataInfo` and the variation's data-reference
/// owner in the package metadata (the media family's data-owner
/// transition, record 0371; Keynote's image fills, record 0372).
fn own_image(
    source: &Package,
    archives: &mut Vec<(Arc<str>, Archive)>,
    shape: u64,
    image: &NewImage,
) -> Result<(), Error> {
    use sha1::Digest as _;
    let objects = super::body_storage_objects::metadata(source)?.ok_or(Error::InvalidSource)?;
    let editable = |archives: &mut Vec<(Arc<str>, Archive)>, name: &str| -> Result<usize, Error> {
        if let Some(position) = archives
            .iter()
            .position(|(member, _)| member.as_ref() == name)
        {
            return Ok(position);
        }
        archives.push((Arc::from(name), graph::editable_archive(source, name)?));
        Ok(archives.len() - 1)
    };
    // The variation the shape now points at.
    let style = archives
        .iter()
        .find_map(|(_, archive)| archive.object(shape))
        .or_else(|| {
            source
                .state
                .source
                .components()
                .iter()
                .find_map(|component| component.archive().object(shape))
        })
        .and_then(|object| object.messages.first())
        .map(|message| shape_style_of(&message.data))
        .transpose()?
        .flatten()
        .ok_or(Error::InvalidSource)?;
    let style_member = archives
        .iter()
        .find(|(_, archive)| archive.object(style).is_some())
        .map(|(name, _)| Arc::clone(name))
        .ok_or(Refusal::UnsupportedStyle)?;
    let position = editable(archives, &style_member)?;
    let variation = archives[position]
        .1
        .object_mut(style)
        .ok_or(Error::InvalidSource)?;
    let info = variation
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource)?;
    if !info.data_references.contains(&image.data) {
        info.data_references.push(image.data);
    }
    let component = objects
        .component_for_entry(&style_member)?
        .ok_or(Error::InvalidSource)?;
    let position = editable(archives, &objects.component)?;
    let metadata = archives[position]
        .1
        .object_mut(objects.object)
        .ok_or(Error::InvalidSource)?;
    let payload = metadata
        .messages
        .get(objects.message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    let payload = super::body_storage_objects::append_data_info(
        &payload,
        image.data,
        &sha1::Sha1::digest(image.bytes.as_ref()),
        &image.file_name,
        &image.stored_name,
        image.bytes.len() as u64,
    )?;
    let payload = super::body_storage_objects::adjust_data_reference(
        &payload, component, image.data, style, true,
    )?;
    let limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource)?;
    metadata
        .replace_message_preserving_header(
            objects.message_index,
            litchi_iwa_core::RawMessage {
                type_: 11_006,
                data: payload,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource)?;
    Ok(())
}

/// `TSWP.ShapeInfoArchive.super.style` of a shape-info payload.
fn shape_style_of(payload: &[u8]) -> Result<Option<u64>, Error> {
    let limits = WireLimits::default();
    let parse = |payload| {
        litchi_iwa_common::wire::WireView::parse(payload, limits).map_err(|_| Error::InvalidSource)
    };
    let view = parse(payload)?;
    let Some(shape) = view.fields().find(|field| field.number() == 1) else {
        return Ok(None);
    };
    let shape = parse(shape.payload())?;
    let Some(style) = shape.fields().find(|field| field.number() == 2) else {
        return Ok(None);
    };
    let reference = parse(style.payload())?;
    Ok(reference
        .fields()
        .find(|field| field.number() == 1)
        .and_then(|field| decode_varint_from_bytes(field.payload()).ok())
        .map(|(value, _)| value))
}
