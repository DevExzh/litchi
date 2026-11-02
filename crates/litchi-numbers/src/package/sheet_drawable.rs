//! Implementation of the public [`crate::sheet::drawable`] structural
//! transaction: the back-to-front list of the drawables a sheet owns, the
//! removal of a drawable and its private graph, and z-order moves
//! (record 0378).
//!
//! An [`Edit`] stages any number of operations against one sheet of an
//! immutable package. Selectors resolve against that base snapshot and the
//! edit remembers native identity privately, so a later operation is never
//! shifted by an earlier one. [`Edit::commit`] applies the operations in
//! order in one structural working set, publishes one candidate, reopens it
//! and reads the sheet back.
//!
//! This is the behaviour of the host editor's `sheet_drawable_order`,
//! `set_sheet_drawable_order`, `move_sheet_drawable` and
//! `remove_sheet_{text_box,shape}` routes.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "the focused transaction keeps its public types beside the staged rewrite"
)]

use std::collections::{BTreeSet, HashMap};
use std::fmt;

use litchi_core::Position;
use thiserror::Error as ThisError;

use super::Package;
use super::sheet_drawable_chart as chart;
use super::sheet_drawable_comment as comment;
use super::sheet_drawable_comment_graph::Graph;
use super::sheet_drawable_duplicate as duplicate;
use super::sheet_drawable_graph::{self as graph, SheetDrawables};
use super::sheet_drawable_media as media;
use super::structure::{
    self, Working,
    collect::{Census, verify_removed},
    metadata::{self, Transition},
    references,
};
use crate::selector::SheetSelector;
use crate::sheet::drawable::{
    AudioOptions, Drawable, DrawableSelector, Geometry, GeometryFlags, ImageOptions, Kind,
    MovieOptions,
};
use litchi_iwa_common::shape::line::Segment;
use litchi_iwa_common::shape::path::Preset;
use litchi_iwa_protos::shape_style_codec::LINE_FLAGS;

use super::sheet_drawable_author::{self as author, Context};

structure::structural_family!("sheet-drawable");

/// A typed, content-redacted sheet-drawable failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted source is malformed or ambiguous.
    #[error("rooted sheet-drawable source is invalid")]
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
    #[error("could not allocate {amount} sheet-drawable items or bytes")]
    Allocation {
        /// Requested amount.
        amount: usize,
    },
    /// The reopened candidate did not reproduce the requested state.
    #[error("sheet-drawable candidate verification failed")]
    Verification,
    /// The patch was applied to an artifact other than its exact source.
    #[error("sheet-drawable patch conflicts with this package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The sheet selector did not resolve one sheet.
    #[error("sheet selector did not resolve")]
    SheetNotFound,
    /// No drawable matches the selector.
    #[error("the sheet has no drawable at {selector:?}")]
    DrawableNotFound {
        /// The selector that did not resolve.
        selector: DrawableSelector,
    },
    /// The z-order destination is outside the sheet's drawable list.
    #[error("z-order position {position:?} is outside {count} drawables")]
    ZPositionOutOfRange {
        /// The requested back-to-front position.
        position: Position,
        /// The number of drawables on the sheet.
        count: usize,
    },
    /// A new drawable's frame needs a finite position and a positive size
    /// (and a rounded-rectangle radius at most half the shorter side).
    #[error("invalid frame for a new drawable")]
    InvalidGeometry,
    /// A new media drawable's file name, media type or data does not fit
    /// (an empty or unsafe name, a type other than the drawable's, or a
    /// signature contradicting the extension).
    #[error("invalid media for a new drawable")]
    InvalidMedia,
    /// A new chart's kind is not a supported concrete kind, or its data
    /// cannot form that kind.
    #[error("invalid kind or data for a new chart")]
    InvalidChart,
    /// The drawable was already removed by an earlier operation of this edit.
    #[error("the drawable was removed earlier in this edit")]
    AlreadyRemoved,
    /// The operation does not apply to the selected kind (removing or
    /// duplicating a table, or duplicating a text box without new text,
    /// through this family).
    #[error("{operation} does not apply to a Numbers sheet {kind}")]
    UnsupportedOperation {
        /// The refused verb.
        operation: &'static str,
        /// The selected kind.
        kind: &'static str,
    },
    /// Commit was requested before staging an operation.
    #[error("no sheet-drawable operation is staged")]
    NoStagedOperation,
    /// A surviving object depends on the removed graph.
    #[error("a surviving object depends on the removed drawable")]
    UnsupportedDependency,
    /// The source cannot publish an exact changed artifact.
    #[error("source representation does not support sheet-drawable edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::SheetNotFound | Self::DrawableNotFound { .. } | Self::AlreadyRemoved => {
                RefusalKind::NotFound
            },
            Self::ZPositionOutOfRange { .. }
            | Self::InvalidGeometry
            | Self::InvalidMedia
            | Self::InvalidChart => RefusalKind::InvalidValue,
            Self::UnsupportedOperation { .. } | Self::NoStagedOperation => {
                RefusalKind::UnsupportedOperation
            },
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

/// One staged operation, with the native identity resolved from the base.
#[derive(Debug, Clone, PartialEq)]
enum Operation {
    Remove {
        drawable: u64,
    },
    MoveTo {
        drawable: u64,
        position: usize,
    },
    Add(Box<NewShape>),
    AddMedia(Box<media::NewMedia>),
    AddChart(Box<chart::NewChart>),
    Duplicate {
        drawable: u64,
        /// The copy's text (a text box duplicate).
        text: Option<String>,
    },
}

/// One staged text box, preset shape or straight line.
#[derive(Debug, Clone, PartialEq)]
struct NewShape {
    path: author::Path,
    geometry: Geometry,
    text: String,
    text_box: bool,
}

/// Drawable operations staged against one sheet of an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    base: SheetDrawables,
    operations: Vec<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("sheet", &self.base.sheet.position)
            .field("operations", &self.operations.len())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn resolve(&self, selector: DrawableSelector) -> Result<(u64, Kind), Error> {
        let entry = self
            .base
            .select(selector)
            .ok_or(Refusal::DrawableNotFound { selector })?;
        if self.operations.iter().any(|operation| {
            matches!(operation, Operation::Remove { drawable } if *drawable == entry.identifier)
        }) {
            return Err(Refusal::AlreadyRemoved.into());
        }
        Ok((entry.identifier, entry.kind))
    }

    /// The number of drawables the sheet holds after the staged operations.
    fn staged_count(&self) -> usize {
        let removed = self
            .operations
            .iter()
            .filter(|operation| matches!(operation, Operation::Remove { .. }))
            .count();
        let added = self
            .operations
            .iter()
            .filter(|operation| {
                matches!(
                    operation,
                    Operation::Add(_)
                        | Operation::AddMedia(_)
                        | Operation::AddChart(_)
                        | Operation::Duplicate { .. }
                )
            })
            .count();
        self.base.entries.len() - removed + added
    }

    fn stage_shape(
        &mut self,
        preset: Preset,
        geometry: Geometry,
        text: &str,
        text_box: bool,
    ) -> Result<&mut Self, Error> {
        let valid = geometry.is_valid()
            && geometry.position.is_some()
            && geometry
                .size
                .is_some_and(|size| size.width > 0.0 && size.height > 0.0);
        if !valid {
            return Err(Refusal::InvalidGeometry.into());
        }
        let mut owned = String::new();
        owned
            .try_reserve_exact(text.len())
            .map_err(|_error| Error::Allocation { amount: text.len() })?;
        owned.push_str(text);
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::Add(Box::new(NewShape {
            path: author::Path::Preset(preset),
            geometry,
            text: owned,
            text_box,
        })));
        Ok(self)
    }

    /// Stage a new text box holding `text` at `geometry` (position and a
    /// positive size required), styled by the theme's text-box preset. It
    /// joins the front of the sheet's list.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidGeometry`], [`Error::Allocation`].
    pub fn add_text_box(&mut self, geometry: Geometry, text: &str) -> Result<&mut Self, Error> {
        self.stage_shape(Preset::Rectangle, geometry, text, true)
    }

    /// Stage a new preset shape holding `text` at `geometry`, styled by the
    /// theme's text-box preset (as Numbers' own shape insertion does). It
    /// joins the front of the sheet's list.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidGeometry`] (also a rounded-rectangle radius above
    /// half the shorter side, at commit), [`Error::Allocation`].
    pub fn add_shape(
        &mut self,
        preset: Preset,
        geometry: Geometry,
        text: &str,
    ) -> Result<&mut Self, Error> {
        self.stage_shape(preset, geometry, text, false)
    }

    /// Stage a new native straight line along `segment` (the host's
    /// `add_sheet_line`), styled by the theme's text-box preset with empty
    /// text. It joins the front of the sheet's list. Its decorations and
    /// stroke are set through [`crate::Package::edit_sheet_drawable_style`].
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidGeometry`] for a segment with a non-finite point.
    pub fn add_line(&mut self, segment: Segment) -> Result<&mut Self, Error> {
        let (position, size, angle) = segment.geometry();
        let geometry = Geometry {
            position: Some(position),
            size: Some(size),
            flags: Some(GeometryFlags::from_bits(LINE_FLAGS)),
            angle: Some(angle),
        };
        if !geometry.is_valid() || size.width <= 0.0 {
            return Err(Refusal::InvalidGeometry.into());
        }
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::Add(Box::new(NewShape {
            path: author::Path::Line(segment.length()),
            geometry,
            text: String::new(),
            text_box: false,
        })));
        Ok(self)
    }

    /// Stage the removal of a text box, ordinary shape, image, movie or
    /// audio clip and the objects only it references (its text storage,
    /// title and caption stand-ins, direct comment graph), with the embedded
    /// data nothing else references. A chart's whole private graph goes
    /// (mediator, preset and styles included); its styles leave the
    /// document stylesheet and its preset the theme.
    ///
    /// # Errors
    ///
    /// [`Refusal::DrawableNotFound`], [`Refusal::AlreadyRemoved`], and
    /// [`Refusal::UnsupportedOperation`] for the other kinds.
    pub fn remove(&mut self, drawable: impl Into<DrawableSelector>) -> Result<&mut Self, Error> {
        let (identifier, kind) = self.resolve(drawable.into())?;
        if !matches!(
            kind,
            Kind::TextBox | Kind::Shape | Kind::Image | Kind::Movie | Kind::Audio | Kind::Chart
        ) {
            return Err(Refusal::UnsupportedOperation {
                operation: "remove",
                kind: kind.name(),
            }
            .into());
        }
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::Remove {
            drawable: identifier,
        });
        Ok(self)
    }

    fn stage_media(&mut self, media: media::NewMedia) -> Result<&mut Self, Error> {
        media.validate().map_err(|invalid| match invalid {
            media::Invalid::Media => Error::Refused(Refusal::InvalidMedia),
            media::Invalid::Placement => Error::Refused(Refusal::InvalidGeometry),
        })?;
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::AddMedia(Box::new(media)));
        Ok(self)
    }

    /// Stage a new image from a file name (a safe basename whose extension
    /// names an image type) and its bytes, placed by `options`, styled by
    /// the theme's first image style; the bytes become a new embedded asset.
    /// It joins the front of the sheet's list.
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidMedia`], [`Refusal::InvalidGeometry`],
    /// [`Error::Allocation`].
    pub fn add_image(
        &mut self,
        file_name: &str,
        data: &[u8],
        options: ImageOptions,
    ) -> Result<&mut Self, Error> {
        self.stage_media(media::NewMedia::Image {
            file_name: owned(file_name)?,
            data: structure::copy_bytes(data)?,
            options,
        })
    }

    /// Stage a new movie with its poster image (each a file name and its
    /// bytes), placed and timed by `options`, styled by the theme's first
    /// movie style. It joins the front of the sheet's list.
    ///
    /// # Errors
    ///
    /// As [`Self::add_image`].
    pub fn add_movie(
        &mut self,
        (file_name, data): (&str, &[u8]),
        (poster_name, poster): (&str, &[u8]),
        options: MovieOptions,
    ) -> Result<&mut Self, Error> {
        self.stage_media(media::NewMedia::Movie {
            movie: (owned(file_name)?, structure::copy_bytes(data)?),
            poster: (owned(poster_name)?, structure::copy_bytes(poster)?),
            options,
        })
    }

    /// Stage a new chart of `kind` over the inline `data`, framed by
    /// `geometry` (the host's `add_sheet_chart`): the chart graph of the
    /// shared chart author in the CalculationEngine component, its styles
    /// in the document stylesheet and its preset in the theme. It joins the
    /// front of the sheet's list; its series are the data's rows.
    ///
    /// ```no_run
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::chart::lifecycle::Kind;
    /// use litchi_numbers::sheet::drawable::{Geometry, Point, Size};
    /// use litchi_numbers::ChartData;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let data = ChartData::new(
    ///     vec!["2025".into()],
    ///     vec!["North".into(), "South".into()],
    ///     vec![vec![Some(1.0), Some(2.0)]],
    /// )?;
    /// let mut edit = package.edit_sheet_drawables(0)?;
    /// edit.add_chart(
    ///     Kind::Bar2d,
    ///     data,
    ///     Geometry::new(
    ///         Point { x: 40.0, y: 40.0 },
    ///         Size { width: 320.0, height: 240.0 },
    ///     ),
    /// )?;
    /// let commit = edit.commit()?;
    /// # let _ = commit;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::InvalidChart`] for an unassignable kind or data that
    /// cannot form it; [`Refusal::InvalidGeometry`]; [`Error::Allocation`].
    pub fn add_chart(
        &mut self,
        kind: litchi_iwa_common::chart::kind::Kind,
        data: litchi_iwa_common::chart::data::ChartData,
        geometry: Geometry,
    ) -> Result<&mut Self, Error> {
        let chart = chart::NewChart {
            kind,
            data,
            geometry,
        };
        chart.validate()?;
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::AddChart(Box::new(chart)));
        Ok(self)
    }

    /// Stage a new audio clip from a file name and its bytes. It joins the
    /// front of the sheet's list.
    ///
    /// # Errors
    ///
    /// As [`Self::add_image`].
    pub fn add_audio(
        &mut self,
        file_name: &str,
        data: &[u8],
        options: AudioOptions,
    ) -> Result<&mut Self, Error> {
        self.stage_media(media::NewMedia::Audio {
            file_name: owned(file_name)?,
            data: structure::copy_bytes(data)?,
            options,
        })
    }

    /// Stage a copy of a shape, line, image, movie or audio clip: its graph
    /// (the drawable, its title and caption stand-ins and a shape's text
    /// storage) under fresh identifiers, offset by 10 points on both axes
    /// as Numbers places a duplicate, sharing its style, comment and
    /// embedded data. A chart's copy also clones its mediator, preset and
    /// styles, registered like the source's. The copy joins the front of the
    /// sheet's list.
    ///
    /// # Errors
    ///
    /// [`Refusal::DrawableNotFound`], [`Refusal::AlreadyRemoved`], and
    /// [`Refusal::UnsupportedOperation`] for the other kinds.
    pub fn duplicate(&mut self, drawable: impl Into<DrawableSelector>) -> Result<&mut Self, Error> {
        let (identifier, kind) = self.resolve(drawable.into())?;
        if !matches!(
            kind,
            Kind::Shape | Kind::Image | Kind::Movie | Kind::Audio | Kind::Chart
        ) {
            return Err(Refusal::UnsupportedOperation {
                operation: "duplicate",
                kind: kind.name(),
            }
            .into());
        }
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::Duplicate {
            drawable: identifier,
            text: None,
        });
        Ok(self)
    }

    /// Stage a copy of a text box holding `text`: its shape, caption and
    /// title stand-ins and storage under fresh identifiers, offset by 10
    /// points on both axes, sharing its styles; the copy's text is then
    /// replaced as Numbers' Duplicate followed by typing does. The copy
    /// joins the front of the sheet's list.
    ///
    /// # Errors
    ///
    /// [`Refusal::DrawableNotFound`], [`Refusal::AlreadyRemoved`],
    /// [`Refusal::UnsupportedOperation`] for the other kinds,
    /// [`Error::Allocation`].
    pub fn duplicate_text_box(
        &mut self,
        drawable: impl Into<DrawableSelector>,
        text: &str,
    ) -> Result<&mut Self, Error> {
        let (identifier, kind) = self.resolve(drawable.into())?;
        if kind != Kind::TextBox {
            return Err(Refusal::UnsupportedOperation {
                operation: "duplicate_text_box",
                kind: kind.name(),
            }
            .into());
        }
        let mut owned = String::new();
        owned
            .try_reserve_exact(text.len())
            .map_err(|_error| Error::Allocation { amount: text.len() })?;
        owned.push_str(text);
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::Duplicate {
            drawable: identifier,
            text: Some(owned),
        });
        Ok(self)
    }

    /// Stage a move of the drawable to the zero-based back-to-front
    /// `position` of the sheet's list (interpreted after the drawable is
    /// taken out, as the native Arrange commands do). Selectors of later
    /// operations still resolve against the base snapshot.
    ///
    /// # Errors
    ///
    /// [`Refusal::DrawableNotFound`], [`Refusal::AlreadyRemoved`],
    /// [`Refusal::ZPositionOutOfRange`].
    pub fn move_to(
        &mut self,
        drawable: impl Into<DrawableSelector>,
        position: Position,
    ) -> Result<&mut Self, Error> {
        let (identifier, _) = self.resolve(drawable.into())?;
        let count = self.staged_count();
        if position.get() >= count {
            return Err(Refusal::ZPositionOutOfRange { position, count }.into());
        }
        structure::reserve(&mut self.operations, 1)?;
        self.operations.push(Operation::MoveTo {
            drawable: identifier,
            position: position.get(),
        });
        Ok(self)
    }

    /// Validate and atomically publish the staged operations.
    ///
    /// # Errors
    ///
    /// [`Refusal::NoStagedOperation`] for an empty edit; source, resource and
    /// verification errors. Failure never publishes a partial package. An
    /// edit whose operations leave the sheet unchanged is an exact no-op.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.operations.is_empty() {
            return Err(Refusal::NoStagedOperation.into());
        }
        let structural = self
            .operations
            .iter()
            .any(|operation| !matches!(operation, Operation::MoveTo { .. }));
        if structural {
            let (published, order) = publish_structural(self.source, &self.base, &self.operations)?;
            let reopened = graph::sheet_drawables(&published.package, self.base.sheet.position)?;
            if reopened.identifiers() != order {
                return Err(Error::Verification);
            }
            let published = self.retext_duplicates(published, &order)?;
            return Commit::new(self.source, published);
        }
        let before = self.base.identifiers();
        let mut order = before.clone();
        for operation in &self.operations {
            if let Operation::MoveTo { drawable, position } = *operation {
                move_in(&mut order, drawable, position)?;
            }
        }
        if order == before {
            return Commit::unchanged(self.source);
        }
        let published = publish_order(self.source, &self.base, &order)?;
        let reopened = graph::sheet_drawables(&published.package, self.base.sheet.position)?;
        if reopened.identifiers() != order {
            return Err(Error::Verification);
        }
        Commit::new(self.source, published)
    }
}

impl Edit<'_> {
    /// Replace the text of every text-box duplicate in `published` (the
    /// host's `set_text` on the copy's storage after the graph clone).
    fn retext_duplicates(
        &self,
        mut published: structure::Published,
        order: &[u64],
    ) -> Result<structure::Published, Error> {
        let base = self.base.identifiers();
        let mut copies = order.iter().filter(|identifier| !base.contains(identifier));
        for operation in &self.operations {
            let (Operation::Add(_)
            | Operation::AddMedia(_)
            | Operation::AddChart(_)
            | Operation::Duplicate { .. }) = operation
            else {
                continue;
            };
            let copy = *copies.next().ok_or(Error::InvalidSource)?;
            let Operation::Duplicate {
                text: Some(text), ..
            } = operation
            else {
                continue;
            };
            let position = order
                .iter()
                .position(|identifier| *identifier == copy)
                .ok_or(Error::InvalidSource)?;
            let mut edit = published
                .package
                .edit_sheet_drawable_plain_text(
                    self.base.sheet.position,
                    DrawableSelector::index(position),
                )
                .map_err(from_plain_text)?;
            edit.set(text).map_err(from_plain_text)?;
            let commit = edit.commit().map_err(from_plain_text)?;
            let touched = published.touched_components;
            published = structure::Published {
                package: commit.into_package(),
                touched_components: touched,
                deleted_previews: published.deleted_previews,
            };
        }
        Ok(published)
    }
}

fn owned(text: &str) -> Result<String, Error> {
    let mut value = String::new();
    value
        .try_reserve_exact(text.len())
        .map_err(|_error| Error::Allocation { amount: text.len() })?;
    value.push_str(text);
    Ok(value)
}

fn from_plain_text(error: super::sheet_drawable_plain_text::Error) -> Error {
    use super::sheet_drawable_plain_text::Error as E;
    match error {
        E::Allocation { amount } => Error::Allocation { amount },
        E::Verification => Error::Verification,
        _ => Error::InvalidSource,
    }
}

fn move_in(order: &mut Vec<u64>, drawable: u64, position: usize) -> Result<(), Error> {
    let current = order
        .iter()
        .position(|identifier| *identifier == drawable)
        .ok_or(Error::InvalidSource)?;
    let moved = order.remove(current);
    if position > order.len() {
        return Err(Refusal::ZPositionOutOfRange {
            position: Position::new(position),
            count: order.len() + 1,
        }
        .into());
    }
    order.insert(position, moved);
    Ok(())
}

/// A pure z-order change: the sheet's list is rewritten in place.
fn publish_order(
    source: &Package,
    base: &SheetDrawables,
    order: &[u64],
) -> Result<structure::Published, Error> {
    let sheet = &base.sheet;
    let mut rewrite = |object: &mut litchi_iwa_core::ArchiveObject,
                       limits: litchi_iwa_core::Limits| {
        let data = references::rewrite(
            &object.messages[sheet.message_index].data,
            sheet.path,
            order,
        )?;
        references::replace(
            object,
            sheet.message_index,
            data,
            &[],
            &[],
            sheet.path,
            limits,
        )
    };
    Ok(graph::publish_in_place(
        source,
        &mut [(sheet.identifier, &mut rewrite)],
    )?)
}

/// Removals, additions and moves in one structural working set; returns
/// the candidate and the sheet's final list.
#[allow(
    clippy::too_many_lines,
    reason = "one staged publication: removals, creations, the sheet list, the registry"
)]
fn publish_structural(
    source: &Package,
    base: &SheetDrawables,
    operations: &[Operation],
) -> Result<(structure::Published, Vec<u64>), Error> {
    let sheet = &base.sheet;
    let mut working = Working::new(source)?;
    let mut registry = metadata::registry(&working)?;
    let sheet_member = working
        .locate(sheet.identifier)?
        .ok_or(Error::InvalidSource)?;
    let mut order = base.identifiers();
    let mut removed = Vec::new();
    let mut removed_charts = Vec::new();
    let mut created: Vec<u64> = Vec::new();
    let mut context = None;
    let mut transition = Transition::default();
    for operation in operations {
        match operation {
            Operation::Remove { drawable } => {
                order.retain(|identifier| identifier != drawable);
                if base
                    .entries
                    .iter()
                    .any(|entry| entry.identifier == *drawable && entry.kind == Kind::Chart)
                {
                    removed_charts.push(*drawable);
                } else {
                    removed.push(*drawable);
                }
            },
            Operation::AddChart(new) => {
                let registry = registry
                    .as_ref()
                    .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
                let created_id = chart::add(
                    &mut working,
                    registry,
                    &mut transition,
                    (sheet.identifier, &sheet_member),
                    new,
                )?;
                order.push(created_id);
                created.push(created_id);
            },
            Operation::MoveTo { drawable, position } => move_in(&mut order, *drawable, *position)?,
            Operation::Add(shape) => {
                if context.is_none() {
                    context = Some(Context::read(source)?);
                }
                let context = context.as_ref().ok_or(Error::InvalidSource)?;
                let identifiers = working.allocate(4)?;
                let identifiers: [u64; 4] = identifiers
                    .try_into()
                    .map_err(|_error| Error::InvalidSource)?;
                let authored = author::author(
                    context,
                    identifiers,
                    sheet.identifier,
                    shape.geometry,
                    shape.path,
                    &shape.text,
                    shape.text_box,
                    working.archive_limits(),
                )?
                .ok_or(Refusal::InvalidGeometry)?;
                for object in authored.objects {
                    working.insert(&sheet_member, object)?;
                }
                if let Some(registry) = &registry {
                    let component = registry.component_for_entry(&sheet_member)?.identifier;
                    for identifier in authored.identifiers {
                        let seed = metadata::identity_seed(&working, identifier)?;
                        transition.register_uuid(registry, component, identifier, seed)?;
                    }
                    let style_member = working
                        .locate(context.shape_style)?
                        .ok_or(Error::InvalidSource)?;
                    let style_component = registry.component_for_entry(&style_member)?.identifier;
                    transition.reference(
                        registry,
                        component,
                        style_component,
                        context.shape_style,
                    )?;
                }
                order.push(authored.identifiers[0]);
                created.push(authored.identifiers[0]);
            },
            Operation::AddMedia(new) => {
                let registry = registry
                    .as_ref()
                    .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
                let created_id = media::author(
                    &mut working,
                    registry,
                    &mut transition,
                    (sheet.identifier, &sheet_member),
                    new,
                )?;
                order.push(created_id);
                created.push(created_id);
            },
            Operation::Duplicate { drawable, .. } => {
                let entry = base
                    .entries
                    .iter()
                    .find(|entry| entry.identifier == *drawable)
                    .ok_or(Error::InvalidSource)?;
                if entry.kind == Kind::Chart {
                    let registry = registry
                        .as_ref()
                        .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
                    let copy = chart::duplicate(
                        &mut working,
                        registry,
                        &mut transition,
                        (sheet.identifier, &sheet_member),
                        *drawable,
                    )?;
                    order.push(copy);
                    created.push(copy);
                    continue;
                }
                let (index, message_type) = entry.message.ok_or(Error::InvalidSource)?;
                let copy = duplicate_graph(
                    &mut working,
                    registry.as_ref(),
                    &mut transition,
                    &sheet_member,
                    (*drawable, index, message_type),
                )?;
                order.push(copy);
                created.push(copy);
            },
        }
    }
    if !removed.is_empty() || !removed_charts.is_empty() {
        // The host clears a removed drawable's direct comment first.
        let taken = registry
            .take()
            .ok_or(Error::Refused(Refusal::UnsupportedSource))?;
        let mut graph = Graph {
            working: &mut working,
            registry: taken,
            transition: std::mem::take(&mut transition),
        };
        for drawable in removed.iter().chain(&removed_charts) {
            comment::clear_in(&mut graph, *drawable)?;
        }
        let Graph {
            registry: restored,
            transition: cleared,
            ..
        } = graph;
        registry = Some(restored);
        transition = cleared;
    }
    let mut closure = BTreeSet::new();
    if !removed.is_empty() {
        let census = Census::new(&working)?;
        closure = census.private_closure(&working, &removed)?;
        // The census reads the source graph: the comment storages the clear
        // above already removed leave the closure.
        closure.retain(|identifier| matches!(working.locate(*identifier), Ok(Some(_))));
        let detached: Vec<(u64, u64)> = removed
            .iter()
            .map(|drawable| (sheet.identifier, *drawable))
            .collect();
        census.require_detached(&closure, &detached)?;
    }
    let chart_objects = match &registry {
        Some(registry) => chart::remove(
            &mut working,
            registry,
            &mut transition,
            sheet.identifier,
            &removed_charts,
        )?,
        None if removed_charts.is_empty() => Vec::new(),
        None => return Err(Refusal::UnsupportedSource.into()),
    };
    removed.extend(&removed_charts);
    let limits = working.archive_limits();
    let object = working.object_mut(sheet.identifier)?;
    let data = references::rewrite(
        &object.messages[sheet.message_index].data,
        sheet.path,
        &order,
    )?;
    references::replace(
        object,
        sheet.message_index,
        data,
        &removed,
        &created,
        sheet.path,
        limits,
    )?;
    if !closure.is_empty() {
        let list: Vec<u64> = closure.iter().copied().collect();
        if let Some(registry) = &registry {
            transition.remove_objects(registry, &list)?;
        }
        transition.release_suffix(&list)?;
        for identifier in &list {
            working.remove(*identifier)?;
        }
        if let Some(registry) = &registry {
            transition.release_unreferenced_data(&working, registry)?;
        }
    }
    closure.extend(chart_objects);
    let published = working.publish(&transition)?;
    verify_removed(&published.package, &closure)?;
    Ok((published, order))
}

/// Copy the graph of `drawable` into the sheet's member under fresh
/// identifiers (the host's `clone_sheet_image` / `clone_sheet_movie`):
/// clones in graph order, the copy offset, UUIDs of the registered objects
/// and the copy's data-reference owners. Returns the copy's identifier.
fn duplicate_graph(
    working: &mut Working<'_>,
    registry: Option<&metadata::Registry>,
    transition: &mut Transition,
    sheet_member: &str,
    (drawable, index, message_type): (u64, usize, u32),
) -> Result<u64, Error> {
    if working.locate(drawable)?.as_deref() != Some(sheet_member) {
        return Err(Refusal::UnsupportedDependency.into());
    }
    let objects = duplicate::graph_objects(working, drawable, index, message_type)?;
    for object in &objects {
        if working.locate(*object)?.as_deref() != Some(sheet_member) {
            return Err(Refusal::UnsupportedDependency.into());
        }
    }
    let identifiers = working.allocate(objects.len())?;
    let remap: HashMap<u64, u64> = objects
        .iter()
        .copied()
        .zip(identifiers.iter().copied())
        .collect();
    let limits = working.archive_limits();
    for object in &objects {
        let mut copy = duplicate::clone_object(working.object(*object)?, &remap, limits)?;
        if *object == drawable {
            let data =
                duplicate::offset(&copy.messages[index].data, message_type, duplicate::OFFSET)?;
            copy.replace_message_preserving_header(
                index,
                litchi_iwa_core::RawMessage {
                    type_: message_type,
                    data,
                },
                limits,
            )
            .map_err(|error| structure::map_core(&error))?;
        }
        working.insert(sheet_member, copy)?;
    }
    let copy = *remap.get(&drawable).ok_or(Error::InvalidSource)?;
    if let Some(registry) = registry {
        let component = registry.component_for_entry(sheet_member)?.identifier;
        for object in &objects {
            if registry
                .uuids
                .iter()
                .any(|binding| binding.component == component && binding.object == *object)
            {
                let new = remap[object];
                let seed = metadata::identity_seed(working, new)?;
                transition.register_uuid(registry, component, new, seed)?;
            }
        }
        let mut data: Vec<u64> = Vec::new();
        for info in &working.object(drawable)?.archive_info.message_infos {
            for identifier in info.data_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.data_references),
            ) {
                if !data.contains(identifier) {
                    structure::reserve(&mut data, 1)?;
                    data.push(*identifier);
                }
            }
        }
        data.sort_unstable();
        for identifier in data {
            structure::reserve(&mut transition.data_owner_additions, 1)?;
            transition.data_owner_additions.push(metadata::DataOwner {
                component,
                data: identifier,
                object: copy,
                count: 1,
            });
        }
    }
    Ok(copy)
}

fn resolve_sheet(package: &Package, selector: SheetSelector<'_>) -> Result<usize, Error> {
    graph::sheet_position(package, selector)?.ok_or_else(|| Refusal::SheetNotFound.into())
}

impl Package {
    /// The drawables one sheet owns, from the back-most to the front-most
    /// layer (`TN.SheetArchive.drawable_infos`), with their kind and frame.
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`] and [`Error::InvalidSource`] for a
    /// malformed drawable list.
    pub fn sheet_drawable_list<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
    ) -> Result<Box<[Drawable]>, Error> {
        let position = resolve_sheet(self, sheet.into())?;
        let drawables = graph::sheet_drawables(self, position)?;
        Ok(graph::list(self, &drawables)?.into_boxed_slice())
    }

    /// Start a structural drawable edit of one sheet: removals and z-order
    /// moves, published together by [`Edit::commit`].
    ///
    /// ```no_run
    /// use litchi_core::Position;
    /// use litchi_numbers::Package;
    /// use litchi_numbers::sheet::drawable::DrawableSelector;
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.numbers")?;
    /// let mut edit = package.edit_sheet_drawables(0)?;
    /// edit.move_to(DrawableSelector::text_box(0), Position::new(0))?;
    /// let commit = edit.commit()?;
    /// let restored = commit.package().apply_sheet_drawables(&commit.patch().inverse())?;
    /// # let _ = restored;
    /// # Ok(())
    /// # }
    /// ```
    ///
    /// # Errors
    ///
    /// [`Refusal::SheetNotFound`] and [`Error::InvalidSource`].
    pub fn edit_sheet_drawables<'sheet>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
    ) -> Result<Edit<'_>, Error> {
        let position = resolve_sheet(self, sheet.into())?;
        Ok(Edit {
            source: self,
            base: graph::sheet_drawables(self, position)?,
            operations: Vec::new(),
        })
    }

    /// Apply a sheet-drawable patch to its exact source package.
    ///
    /// # Errors
    ///
    /// [`Error::PatchConflict`] when this package is not the patch's exact
    /// source.
    pub fn apply_sheet_drawables(&self, patch: &Patch) -> Result<Commit, Error> {
        apply_patch(self, patch)
    }
}
