//! Selector-first lifecycle transactions of body images, movies and audio
//! clips: removal (this module's first slice), with duplication, data
//! replacement and creation following the same private graph proof.
//!
//! The graph of one body media drawable is what the migration host
//! deletes (record 0371): the drawable, the private objects of its title
//! and caption (a stand-in, or the caption graph), and its body attachment,
//! all in the drawable's component. Removal deletes the anchoring U+FFFC
//! and its attachment entry, drops the drawable from the document z-order,
//! releases the graph's data-reference owners, UUID registrations and
//! identifier watermark, and deletes every embedded data asset no object,
//! component or data-metadata entry references any more.

use std::collections::HashSet;
use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_common::media::embedded::{AssetId, Embedded};
use litchi_iwa_common::wire::remove_repeated_length_delimited_field_where;
use litchi_iwa_core::{Archive, RawMessage};
use litchi_iwa_text::TextPosition;
use thiserror::Error as ThisError;

use litchi_iwa_common::shape::path::Preset;

use super::Package;
use super::body_drawable_caption::{self as caption, Kind as LabelKind};
use super::body_drawable_graph::{
    self as graph, Budget, GraphError, GraphLimit, ObjectIndex, kind_name,
};
use super::body_storage_objects as objects;

mod chart;
use crate::selector::DrawableSelector;

const ZORDER_MESSAGE_TYPE: u32 = 10_015;
const ZORDER_DRAWABLES_FIELD: u32 = 1;
const DRAWABLE_COMMENT_FIELD: u32 = 6;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const DATA_METADATA_MAP_MESSAGE_TYPE: u32 = 11_015;
/// The document component, which registers body drawables' UUIDs and data
/// references.
const DOCUMENT_COMPONENT_IDENTIFIER: u64 = 1;

/// A finite resource charged by one body-media lifecycle operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Bytes in native payloads inspected by the graph proof.
    InputBytes,
    /// Bytes in the rewritten payload or package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal, text rewrite and reassembly work.
    WireWork,
    /// Semantic section or body-storage cardinality.
    PayloadItems,
    /// Semantic UTF-8 text retained by the reopened document.
    TextBytes,
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
        })
    }
}

/// Failure from a body-media lifecycle transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted media graph is malformed.
    #[error("the selected Pages body-media source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error("Pages body-media {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for a Pages body-media operation")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested change.
    #[error("the edited Pages body media failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body-media patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family refuses, classified by [`Refusal::kind`].
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No body drawable of the selected kind is at the selected position.
    #[error("the Pages body has no {kind} at position {position:?}")]
    NotFound {
        /// The selected kind.
        kind: &'static str,
        /// The selected position.
        position: Position,
    },
    /// The selected kind is not body media (shapes have their own family).
    #[error("a Pages body {kind} is not body media")]
    NotMedia {
        /// The selected kind.
        kind: &'static str,
    },
    /// The drawable carries a comment, which this family does not remove.
    #[error("the selected Pages body media carries a comment")]
    Commented,
    /// The media graph shares an object or spans components in a way this
    /// family does not prove.
    #[error("the selected Pages body media has a shared or split graph")]
    SharedGraph,
    /// Replacement data is empty.
    #[error("replacement media data is empty")]
    EmptyData,
    /// Only a movie has a poster image.
    #[error("a Pages body {kind} has no poster image")]
    NoPoster {
        /// The selected kind.
        kind: &'static str,
    },
    /// The replacement's signature names another media type than the asset.
    #[error("the replacement media type does not match the asset")]
    MismatchedType,
    /// The asset is not materialized in the package.
    #[error("the media asset is not materialized in this package")]
    NotMaterialized,
    /// The file name is not a safe basename with a stem, or its extension
    /// or the data's signature names another media type than the kind.
    #[error("the media file name or type does not fit the new drawable")]
    InvalidMedia,
    /// A new drawable's position is not finite or its size not positive, or
    /// its preset does not fit the frame.
    #[error("the new Pages drawable's frame is invalid")]
    InvalidGeometry,
    /// A new chart's kind is not a supported concrete kind, or its data
    /// cannot form that kind.
    #[error("invalid kind or data for a new chart")]
    InvalidChart,
    /// The theme lists no style preset for the new drawable.
    #[error("the Pages theme has no style preset for the new drawable")]
    MissingThemeStyle,
    /// The selector resolves no section.
    #[error("the Pages document has no such section")]
    SectionNotFound,
    /// No embedded asset reachable from the document has the identifier.
    #[error("no reachable Pages media asset has the identifier")]
    AssetNotFound,
    /// No operation was staged.
    #[error("no body-media operation is staged")]
    NoStagedOperation,
    /// A second operation was staged on one edit.
    #[error("one body-media edit stages one operation")]
    OperationAlreadyStaged,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact body-media edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound { .. } | Self::AssetNotFound | Self::SectionNotFound => {
                RefusalKind::NotFound
            },
            Self::NotMedia { .. } | Self::NoPoster { .. } => RefusalKind::UnsupportedTarget,
            Self::EmptyData
            | Self::MismatchedType
            | Self::InvalidMedia
            | Self::InvalidGeometry
            | Self::InvalidChart => RefusalKind::InvalidValue,
            Self::NotMaterialized | Self::MissingThemeStyle => RefusalKind::UnsupportedContent,
            Self::Commented | Self::SharedGraph | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
            Self::NoStagedOperation | Self::OperationAlreadyStaged => {
                RefusalKind::UnsupportedOperation
            },
        }
    }
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
                    GraphLimit::InputBytes => LimitKind::InputBytes,
                    GraphLimit::OutputBytes => LimitKind::OutputBytes,
                    GraphLimit::WireFields => LimitKind::WireFields,
                    GraphLimit::WireNesting => LimitKind::WireNesting,
                    GraphLimit::WireWork => LimitKind::WireWork,
                    GraphLimit::SemanticItems => LimitKind::PayloadItems,
                    GraphLimit::SemanticTextBytes => LimitKind::TextBytes,
                },
                observed,
                maximum,
            },
        }
    }
}

impl From<caption::Error> for Error {
    fn from(error: caption::Error) -> Self {
        match error {
            caption::Error::Allocation { amount } => Self::Allocation { amount },
            caption::Error::Limit {
                observed, maximum, ..
            } => Self::Limit {
                kind: LimitKind::WireWork,
                observed,
                maximum,
            },
            _ => Self::InvalidSource,
        }
    }
}

/// The staged operation.
#[derive(Debug, Clone, PartialEq)]
enum Operation {
    Remove(DrawableSelector),
    Duplicate(DrawableSelector, TextPosition, Option<Box<str>>),
    Replace(DrawableSelector, Asset, Arc<[u8]>),
    ReplaceAsset(AssetId, Arc<[u8]>),
    PruneUnreferenced,
    Add(TextPosition, NewMedia),
    AddTextShape(TextPosition, NewTextShape),
    AddChart(TextPosition, Box<chart::NewChart>),
}

/// A new body text box, preset shape or straight line (the host's
/// `add_text_box`, `add_body_shape` and `add_body_line`).
#[derive(Debug, Clone, PartialEq)]
struct NewTextShape {
    text: Box<str>,
    position: crate::drawable::Point,
    size: crate::drawable::Size,
    angle: f32,
    /// The encoded `TSD.PathSourceArchive`.
    path: Vec<u8>,
    text_box: bool,
}

/// A new body image, movie or audio clip.
#[derive(Debug, Clone, PartialEq)]
enum NewMedia {
    Image {
        file_name: Box<str>,
        data: Arc<[u8]>,
        options: crate::image::Options,
    },
    Movie {
        file_name: Box<str>,
        data: Arc<[u8]>,
        poster_name: Box<str>,
        poster: Arc<[u8]>,
        options: crate::movie::Options,
    },
    Audio {
        file_name: Box<str>,
        data: Arc<[u8]>,
        options: crate::audio::Options,
    },
}

/// Which embedded asset of a drawable a replacement targets.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Asset {
    /// The image, movie or audio data.
    Data,
    /// A movie's poster image.
    Poster,
}

/// One body-media lifecycle operation staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    operation: Option<Operation>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("operation", &self.operation)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Stage the removal of one body image, movie, audio clip, shape or
    /// text box with its private graph (a shape's or text box's owned
    /// storage included) and the data assets only it uses. A direct comment
    /// on the drawable is cleared first, as the host does.
    ///
    /// # Errors
    ///
    /// Refuses a second operation.
    pub fn remove(&mut self, selector: impl Into<DrawableSelector>) -> Result<&mut Self, Error> {
        self.stage(Operation::Remove(selector.into()))
    }

    /// Stage a copy of one body image, movie or audio clip anchored at a
    /// UTF-16 body position: its graph is cloned with fresh identifiers,
    /// offset as Pages places a duplicate (12 pt, 30 pt for audio), and its
    /// data assets are shared with the source.
    ///
    /// # Errors
    ///
    /// Refuses a shape and a second operation.
    pub fn duplicate(
        &mut self,
        selector: impl Into<DrawableSelector>,
        anchor: TextPosition,
    ) -> Result<&mut Self, Error> {
        self.stage(Operation::Duplicate(selector.into(), anchor, None))
    }

    /// Stage a copy of one body text box anchored at a UTF-16 body position,
    /// offset 12 pt from its source, holding `text` instead of the source's
    /// text (the host's `duplicate_text_box`).
    ///
    /// # Errors
    ///
    /// Refuses a selector that is not a text box and a second operation.
    pub fn duplicate_text_box(
        &mut self,
        selector: impl Into<DrawableSelector>,
        anchor: TextPosition,
        text: &str,
    ) -> Result<&mut Self, Error> {
        let selector = selector.into();
        if !matches!(selector, DrawableSelector::TextBox(_)) {
            return Err(Refusal::NotFound {
                kind: kind_name(selector),
                position: selector.position(),
            }
            .into());
        }
        self.stage(Operation::Duplicate(
            selector,
            anchor,
            Some(Box::from(text)),
        ))
    }

    /// Stage new bytes for the media data of one body image, movie or audio
    /// clip. The data asset keeps its identifier, so every drawable sharing
    /// it (a duplicate) shows the replacement.
    ///
    /// # Errors
    ///
    /// Refuses a shape, empty data and a second operation.
    pub fn replace_data(
        &mut self,
        selector: impl Into<DrawableSelector>,
        data: &[u8],
    ) -> Result<&mut Self, Error> {
        let selector = selector.into();
        media_kind(selector)?;
        if data.is_empty() {
            return Err(Refusal::EmptyData.into());
        }
        self.stage(Operation::Replace(selector, Asset::Data, Arc::from(data)))
    }

    /// Stage new bytes for the poster image of one body movie.
    ///
    /// # Errors
    ///
    /// Refuses a non-movie, empty data and a second operation.
    pub fn replace_poster(
        &mut self,
        selector: impl Into<DrawableSelector>,
        data: &[u8],
    ) -> Result<&mut Self, Error> {
        let selector = selector.into();
        if !matches!(selector, DrawableSelector::Movie(_)) {
            return Err(Refusal::NoPoster {
                kind: kind_name(selector),
            }
            .into());
        }
        if data.is_empty() {
            return Err(Refusal::EmptyData.into());
        }
        self.stage(Operation::Replace(selector, Asset::Poster, Arc::from(data)))
    }

    /// Stage the removal of every data record nothing references: no
    /// component data-reference owner, object data reference or
    /// `DataMetadataMap` entry names it. Each record's `DataInfo` leaves the
    /// package metadata and its `Data/` member is deleted (the host's
    /// `media::Draft::remove_unreferenced`, over the shared
    /// `litchi_iwa_style::media::unreferenced_data` rule, record 0417). A
    /// package without such records commits as an exact no-op.
    ///
    /// # Errors
    ///
    /// Refuses a second operation.
    pub fn prune_unreferenced(&mut self) -> Result<&mut Self, Error> {
        self.stage(Operation::PruneUnreferenced)
    }

    /// Stage new bytes for one embedded asset reachable from the document,
    /// keeping its identifier.
    ///
    /// # Errors
    ///
    /// Refuses empty data and a second operation.
    pub fn replace_asset(&mut self, asset: AssetId, data: &[u8]) -> Result<&mut Self, Error> {
        if data.is_empty() {
            return Err(Refusal::EmptyData.into());
        }
        self.stage(Operation::ReplaceAsset(asset, Arc::from(data)))
    }

    /// Stage a new body image anchored at a UTF-16 body position, from a
    /// file name (a safe basename whose extension names an image type) and
    /// its bytes, placed by `options`.
    ///
    /// # Errors
    ///
    /// Refuses empty data and a second operation; the file name and type
    /// are checked at commit.
    pub fn add_image(
        &mut self,
        anchor: TextPosition,
        file_name: &str,
        data: &[u8],
        options: crate::image::Options,
    ) -> Result<&mut Self, Error> {
        if data.is_empty() {
            return Err(Refusal::EmptyData.into());
        }
        self.stage(Operation::Add(
            anchor,
            NewMedia::Image {
                file_name: Box::from(file_name),
                data: Arc::from(data),
                options,
            },
        ))
    }

    /// Stage a new body movie with its poster image.
    ///
    /// # Errors
    ///
    /// As [`Self::add_image`].
    pub fn add_movie(
        &mut self,
        anchor: TextPosition,
        (file_name, data): (&str, &[u8]),
        (poster_name, poster): (&str, &[u8]),
        options: crate::movie::Options,
    ) -> Result<&mut Self, Error> {
        if data.is_empty() || poster.is_empty() {
            return Err(Refusal::EmptyData.into());
        }
        self.stage(Operation::Add(
            anchor,
            NewMedia::Movie {
                file_name: Box::from(file_name),
                data: Arc::from(data),
                poster_name: Box::from(poster_name),
                poster: Arc::from(poster),
                options,
            },
        ))
    }

    /// Stage a new body audio clip.
    ///
    /// # Errors
    ///
    /// As [`Self::add_image`].
    pub fn add_audio(
        &mut self,
        anchor: TextPosition,
        file_name: &str,
        data: &[u8],
        options: crate::audio::Options,
    ) -> Result<&mut Self, Error> {
        if data.is_empty() {
            return Err(Refusal::EmptyData.into());
        }
        self.stage(Operation::Add(
            anchor,
            NewMedia::Audio {
                file_name: Box::from(file_name),
                data: Arc::from(data),
                options,
            },
        ))
    }

    /// Stage a new body text box anchored at a UTF-16 body position, with
    /// `text` in its own storage and a frame at `position` × `size` (the
    /// host's `add_text_box`): a rectangle path in the theme's first
    /// text-box style.
    ///
    /// # Errors
    ///
    /// Refuses a non-finite position, a non-positive size and a second
    /// operation.
    pub fn add_text_box(
        &mut self,
        anchor: TextPosition,
        text: &str,
        position: crate::drawable::Point,
        size: crate::drawable::Size,
    ) -> Result<&mut Self, Error> {
        self.add_text_shape(anchor, text, position, size, Preset::Rectangle, true)
    }

    /// Stage a new body shape drawing `preset`, anchored at a UTF-16 body
    /// position, with `text` in its own storage (the host's
    /// `add_body_shape` and `add_body_rectangle`): the theme's first shape
    /// style.
    ///
    /// # Errors
    ///
    /// Refuses a non-finite position, a non-positive size, a preset that
    /// does not fit the frame and a second operation.
    pub fn add_shape(
        &mut self,
        anchor: TextPosition,
        text: &str,
        position: crate::drawable::Point,
        size: crate::drawable::Size,
        preset: Preset,
    ) -> Result<&mut Self, Error> {
        self.add_text_shape(anchor, text, position, size, preset, false)
    }

    /// Stage a new native straight line along `segment`, anchored at a
    /// UTF-16 body position (the host's `add_body_line`), in the theme's
    /// first shape style.
    ///
    /// # Errors
    ///
    /// Refuses a second operation.
    pub fn add_line(
        &mut self,
        anchor: TextPosition,
        segment: litchi_iwa_common::shape::line::Segment,
    ) -> Result<&mut Self, Error> {
        let (position, size, angle) = segment.geometry();
        self.stage(Operation::AddTextShape(
            anchor,
            NewTextShape {
                text: Box::from(""),
                position,
                size,
                angle,
                path: litchi_iwa_style::shape_path::line_path_source(segment.length()),
                text_box: false,
            },
        ))
    }

    /// Stage a new chart of `kind` over the inline `data`, anchored at a
    /// UTF-16 body position and framed by `position` and `size` (the host's
    /// `add_body_chart`): the shared chart author's graph in the body
    /// component, its styles in the document stylesheet and its preset in
    /// the theme. Its series are the data's rows.
    ///
    /// ```no_run
    /// use litchi_pages::chart::lifecycle::Kind;
    /// use litchi_pages::drawable::{Point, Size};
    /// use litchi_pages::{ChartData, Package, TextPosition};
    ///
    /// # fn main() -> Result<(), Box<dyn std::error::Error>> {
    /// let package = Package::open("input.pages")?;
    /// let data = ChartData::new(
    ///     vec!["2025".into()],
    ///     vec!["North".into(), "South".into()],
    ///     vec![vec![Some(1.0), Some(2.0)]],
    /// )?;
    /// let mut edit = package.edit_body_media();
    /// edit.add_chart(
    ///     TextPosition::ZERO,
    ///     Kind::Bar2d,
    ///     data,
    ///     Point { x: 72.0, y: 144.0 },
    ///     Size { width: 320.0, height: 240.0 },
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
    /// cannot form it; [`Refusal::InvalidGeometry`]; a second operation.
    pub fn add_chart(
        &mut self,
        anchor: TextPosition,
        kind: litchi_iwa_common::chart::kind::Kind,
        data: litchi_iwa_common::chart::data::ChartData,
        position: crate::drawable::Point,
        size: crate::drawable::Size,
    ) -> Result<&mut Self, Error> {
        let new = chart::NewChart {
            kind,
            data,
            position,
            size,
        };
        new.validate()?;
        self.stage(Operation::AddChart(anchor, Box::new(new)))
    }

    fn add_text_shape(
        &mut self,
        anchor: TextPosition,
        text: &str,
        position: crate::drawable::Point,
        size: crate::drawable::Size,
        preset: Preset,
        text_box: bool,
    ) -> Result<&mut Self, Error> {
        if !position.x.is_finite()
            || !position.y.is_finite()
            || !size.width.is_finite()
            || !size.height.is_finite()
            || size.width <= 0.0
            || size.height <= 0.0
        {
            return Err(Refusal::InvalidGeometry.into());
        }
        let path = litchi_iwa_style::shape_path::path_source(preset, size.width, size.height)
            .ok_or(Refusal::InvalidGeometry)?;
        self.stage(Operation::AddTextShape(
            anchor,
            NewTextShape {
                text: Box::from(text),
                position,
                size,
                angle: 0.0,
                path,
                text_box,
            },
        ))
    }

    fn stage(&mut self, operation: Operation) -> Result<&mut Self, Error> {
        if self.operation.is_some() {
            return Err(Refusal::OperationAlreadyStaged.into());
        }
        self.operation = Some(operation);
        Ok(self)
    }

    /// Validate and publish the staged operation.
    pub fn commit(self) -> Result<Commit, Error> {
        let operation = self.operation.clone().ok_or(Refusal::NoStagedOperation)?;
        if !self.source.state.source.source_is_exact() {
            return Err(Refusal::UnsupportedSource.into());
        }
        let (package, diagnostics) = match operation {
            Operation::Remove(selector) => match remove(self.source, selector) {
                Err(Error::Refused(Refusal::Commented)) => {
                    // The host clears a removed drawable's direct comment
                    // first (its comment storages and author lifecycle);
                    // the patch spans both steps (record 0413b).
                    let mut clear = self
                        .source
                        .edit_body_drawable_comment(selector)
                        .map_err(comment_error)?;
                    clear.clear().map_err(comment_error)?;
                    let cleared = clear.commit().map_err(comment_error)?.into_package();
                    let (package, mut diagnostics) = remove(&cleared, selector)?;
                    diagnostics.changed = true;
                    (package, diagnostics)
                },
                other => other?,
            },
            Operation::Duplicate(selector, anchor, text) => {
                let (package, diagnostics) = duplicate(self.source, selector, anchor)?;
                match text {
                    Some(text) => (set_copy_text(&package, anchor, &text)?, diagnostics),
                    None => (package, diagnostics),
                }
            },
            Operation::Replace(selector, asset, data) => {
                let identifier = asset_identifier(self.source, selector, asset)?;
                replace(self.source, identifier, &data)?
            },
            Operation::Add(anchor, media) => add(self.source, anchor, &media)?,
            Operation::AddTextShape(anchor, shape) => add_text_shape(self.source, anchor, &shape)?,
            Operation::AddChart(anchor, new) => chart::add(self.source, anchor, &new)?,
            Operation::ReplaceAsset(asset, data) => {
                if !media_assets(self.source, graph::ROOT_OBJECT_IDENTIFIER)?
                    .iter()
                    .any(|embedded| embedded.data() == asset)
                {
                    return Err(Refusal::AssetNotFound.into());
                }
                replace(self.source, asset.get(), &data)?
            },
            Operation::PruneUnreferenced => match prune_unreferenced(self.source)? {
                Some(published) => published,
                None => (self.source.clone(), Diagnostics::default()),
            },
        };
        let target = package.state.source.shared_source();
        Ok(Commit {
            package,
            patch: Patch {
                artifacts: ExactArtifacts::new(self.source.state.source.shared_source(), target),
            },
            diagnostics,
        })
    }
}

/// Exact-source reversible body-media lifecycle patch.
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

/// Content-free diagnostics of one body-media lifecycle publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct Diagnostics {
    changed: bool,
    created_objects: usize,
    removed_objects: usize,
    removed_assets: usize,
}

impl Diagnostics {
    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Native objects the publication created.
    #[must_use]
    pub const fn created_objects(self) -> usize {
        self.created_objects
    }

    /// Native objects the publication deleted.
    #[must_use]
    pub const fn removed_objects(self) -> usize {
        self.removed_objects
    }

    /// Embedded data assets the publication deleted.
    #[must_use]
    pub const fn removed_assets(self) -> usize {
        self.removed_assets
    }
}

/// The validated immutable result of one body-media lifecycle transaction.
#[must_use = "a body-media commit contains the validated package snapshot"]
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

/// One body image, movie or audio clip as the media lifecycle sees it:
/// where it is anchored and which embedded assets it shows. Its frame,
/// properties and playback are read through
/// [`Package::body_drawable_properties`] with [`Media::selector`].
#[derive(Debug, Clone, PartialEq)]
pub struct Media {
    selector: DrawableSelector,
    anchor: TextPosition,
    data: AssetId,
    thumbnail: Option<AssetId>,
    poster: Option<AssetId>,
    original_size: Option<crate::drawable::Size>,
    natural_size: Option<crate::drawable::Size>,
}

impl Media {
    /// The selector of this drawable (kind and text-flow position).
    #[must_use]
    pub const fn selector(&self) -> DrawableSelector {
        self.selector
    }

    /// The UTF-16 body position of the anchoring U+FFFC.
    #[must_use]
    pub const fn anchor(&self) -> TextPosition {
        self.anchor
    }

    /// The image, movie or audio data asset.
    #[must_use]
    pub const fn data(&self) -> AssetId {
        self.data
    }

    /// An image's thumbnail asset.
    #[must_use]
    pub const fn thumbnail(&self) -> Option<AssetId> {
        self.thumbnail
    }

    /// A movie's poster image asset.
    #[must_use]
    pub const fn poster(&self) -> Option<AssetId> {
        self.poster
    }

    /// The size the media was imported at (images and movies).
    #[must_use]
    pub const fn original_size(&self) -> Option<crate::drawable::Size> {
        self.original_size
    }

    /// The media's intrinsic size (images and movies).
    #[must_use]
    pub const fn natural_size(&self) -> Option<crate::drawable::Size> {
        self.natural_size
    }
}

impl Package {
    /// The body's images, movies and audio clips: images first, then movies,
    /// then audio clips, each in text-flow order.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidSource`] for an invalid body graph.
    pub fn body_media(&self) -> Result<Vec<Media>, Error> {
        let mut budget = Budget::new(self)?;
        let index = ObjectIndex::new(self, &mut budget)?;
        let body = graph::resolve_body(&index, &mut budget)?;
        let mut media = Vec::new();
        let kinds: [fn(usize) -> DrawableSelector; 3] = [
            DrawableSelector::image,
            DrawableSelector::movie,
            DrawableSelector::audio,
        ];
        for kind in kinds {
            for position in 0.. {
                let selector = kind(position);
                let Some(selected) = graph::select_drawable(&index, &body, selector, &mut budget)?
                else {
                    break;
                };
                media.push(media_facts(selector, &selected, &mut budget)?);
            }
        }
        Ok(media)
    }

    /// The embedded media assets reachable from the document's object graph
    /// (the host's `media_assets`), ordered by identifier.
    ///
    /// # Errors
    ///
    /// Returns [`Error::InvalidSource`] for invalid package metadata.
    pub fn media_assets(&self) -> Result<Vec<Embedded>, Error> {
        media_assets(self, graph::ROOT_OBJECT_IDENTIFIER)
    }

    /// The embedded media assets reachable from one section's object graph
    /// (the host's `section_media_assets`).
    ///
    /// # Errors
    ///
    /// [`Refusal::SectionNotFound`] when the selector resolves no section.
    pub fn section_media_assets<'selector>(
        &self,
        section: impl Into<crate::SectionSelector<'selector>>,
    ) -> Result<Vec<Embedded>, Error> {
        let position = self
            .select_section(section.into())
            .map_err(|_| Refusal::SectionNotFound)?
            .ok_or(Refusal::SectionNotFound)?
            .index();
        let identifier = super::section_text::native_body(self)
            .map_err(|_| Error::InvalidSource)?
            .section_identifier(position)
            .ok_or(Error::InvalidSource)?;
        media_assets(self, identifier)
    }

    /// The bytes of one embedded asset reachable from the document.
    ///
    /// # Errors
    ///
    /// [`Refusal::AssetNotFound`] for an unreachable asset,
    /// [`Refusal::NotMaterialized`] for one without a `Data/` member.
    pub fn extract_media(&self, asset: AssetId) -> Result<Vec<u8>, Error> {
        let embedded = media_assets(self, graph::ROOT_OBJECT_IDENTIFIER)?
            .into_iter()
            .find(|embedded| embedded.data() == asset)
            .ok_or(Refusal::AssetNotFound)?;
        let path = embedded.package_path().ok_or(Refusal::NotMaterialized)?;
        self.state
            .source
            .package()
            .iter()
            .find(|entry| entry.name() == path)
            .map(|entry| entry.data().to_vec())
            .ok_or_else(|| Refusal::NotMaterialized.into())
    }

    /// Begin a lifecycle operation on the body's images, movies and audio
    /// clips.
    #[must_use]
    pub const fn edit_body_media(&self) -> Edit<'_> {
        Edit {
            source: self,
            operation: None,
        }
    }

    /// Replay an exact-source body-media lifecycle patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact.
    pub fn apply_body_media(&self, patch: &Patch) -> Result<Commit, Error> {
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
                ..Diagnostics::default()
            },
        })
    }
}

fn media_kind(selector: DrawableSelector) -> Result<(), Error> {
    if matches!(
        selector,
        DrawableSelector::Shape(_) | DrawableSelector::TextBox(_)
    ) {
        return Err(Refusal::NotMedia {
            kind: kind_name(selector),
        }
        .into());
    }
    Ok(())
}

/// A drawable-comment failure while clearing the comment of a removed
/// drawable.
fn comment_error(error: super::text_box_drawable_comment::Error) -> Error {
    use super::text_box_drawable_comment::Error as Comment;
    match error {
        Comment::Limit {
            observed, maximum, ..
        } => Error::Limit {
            kind: LimitKind::WireWork,
            observed,
            maximum,
        },
        Comment::Allocation { amount } => Error::Allocation { amount },
        Comment::Verification => Error::Verification,
        Comment::Refused(_) => Refusal::SharedGraph.into(),
        _ => Error::InvalidSource,
    }
}

/// The private graph of one body media drawable.
struct MediaGraph {
    drawable: NonZeroU64,
    attachment: NonZeroU64,
    character_index: u32,
    component: Arc<str>,
    /// Drawable, title objects, caption objects, attachment (host order).
    objects: Vec<u64>,
    /// `(data, object)` for every data reference the graph's messages list.
    data_references: Vec<(u64, u64)>,
    /// Graph objects registered with a UUID in the document component.
    uuid_objects: Vec<u64>,
    selector: DrawableSelector,
    zorder: NonZeroU64,
}

fn resolve(package: &Package, selector: DrawableSelector) -> Result<MediaGraph, Error> {
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let selected =
        graph::select_drawable(&index, &body, selector, &mut budget)?.ok_or(Refusal::NotFound {
            kind: kind_name(selector),
            position: selector.position(),
        })?;
    let view = budget.parse(selected.drawable_payload, 3)?;
    if graph::unique_field(&view, DRAWABLE_COMMENT_FIELD, 2)?.is_some() {
        return Err(Refusal::Commented.into());
    }
    let drawable = selected.anchor.drawable;
    let mut labels = [Vec::new(), Vec::new()];
    for (slot, kind) in labels
        .iter_mut()
        .zip([LabelKind::Title, LabelKind::Caption])
    {
        *slot = caption::label_objects(package, selector, kind)?
            .into_iter()
            .map(NonZeroU64::get)
            .collect();
    }
    let [titles, captions] = labels;
    let attachment = selected.anchor.attachment.get();
    // The host graph orders (its identifiers for a copy follow them):
    // media drawable, titles, captions, attachment; text box drawable,
    // storage, attachment, title, caption; shape drawable, caption, title,
    // storage, attachment.
    let mut graph_objects = vec![drawable.get()];
    match (selector, selected.storage) {
        (DrawableSelector::TextBox(_), Some(storage)) => {
            graph_objects.push(storage.get());
            graph_objects.push(attachment);
            graph_objects.extend(&titles);
            graph_objects.extend(&captions);
        },
        (DrawableSelector::Shape(_), Some(storage)) => {
            graph_objects.extend(&captions);
            graph_objects.extend(&titles);
            graph_objects.push(storage.get());
            graph_objects.push(attachment);
        },
        _ => {
            graph_objects.extend(&titles);
            graph_objects.extend(&captions);
            graph_objects.push(attachment);
        },
    }
    let unique = graph_objects.iter().collect::<HashSet<_>>();
    if unique.len() != graph_objects.len() {
        return Err(Refusal::SharedGraph.into());
    }
    let mut data_references = Vec::new();
    for identifier in &graph_objects {
        let located = index.get(NonZeroU64::new(*identifier).ok_or(Error::InvalidSource)?)?;
        if located.component_index != selected.drawable.component_index {
            return Err(Refusal::SharedGraph.into());
        }
        for info in &located.object.archive_info.message_infos {
            data_references.extend(
                info.data_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| &field.data_references),
                    )
                    .map(|data| (*data, *identifier)),
            );
        }
    }
    let (_, zorder) = root_zorder(package, &mut budget)?;
    let mut uuid_objects = Vec::new();
    if let Some(metadata) = objects::metadata(package)? {
        for identifier in &graph_objects {
            if metadata.uuid_owner(*identifier)? == Some(DOCUMENT_COMPONENT_IDENTIFIER) {
                uuid_objects.push(*identifier);
            }
        }
    }
    Ok(MediaGraph {
        uuid_objects,
        selector,
        drawable,
        attachment: selected.anchor.attachment,
        character_index: selected.anchor.character_index,
        component: index.component_name(selected.drawable.component_index)?,
        objects: graph_objects,
        data_references,
        zorder: zorder.ok_or(Error::InvalidSource)?,
    })
}

fn root_zorder(package: &Package, budget: &mut Budget) -> Result<((), Option<NonZeroU64>), Error> {
    Ok(((), graph::root_zorder(package, budget)?))
}

/// The z-order payload without `drawable`, and the message index.
fn zorder_without(
    archive: &mut Archive,
    zorder: NonZeroU64,
    drawable: u64,
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    let object = archive
        .object_mut(zorder.get())
        .ok_or(Error::InvalidSource)?;
    let (message_index, message) = graph::unique_message(object, ZORDER_MESSAGE_TYPE)?;
    let payload = message.data.clone();
    let mut count = 0usize;
    for reference in litchi_iwa_common::wire::repeated_length_delimited_payloads(
        &payload,
        ZORDER_DRAWABLES_FIELD,
    )
    .map_err(graph::map_wire_error)?
    {
        let budget_view = litchi_iwa_common::wire::WireView::parse(
            reference,
            litchi_iwa_common::WireLimits::default(),
        )
        .map_err(graph::map_wire_error)?;
        let identifier = graph::unique_field(&budget_view, 1, 0)?
            .map(|field| graph::canonical_varint(field.payload()))
            .transpose()?;
        count += usize::from(identifier == Some(drawable));
    }
    if count != 1 {
        return Err(Refusal::SharedGraph.into());
    }
    let data = remove_repeated_length_delimited_field_where(
        &payload,
        ZORDER_DRAWABLES_FIELD,
        |reference| {
            let view = litchi_iwa_common::wire::WireView::parse(
                reference,
                litchi_iwa_common::WireLimits::default(),
            )?;
            Ok::<bool, litchi_iwa_common::Error>(
                view.fields()
                    .find(|field| field.number() == 1)
                    .and_then(|field| {
                        litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()
                    })
                    .is_some_and(|(value, _)| value == drawable),
            )
        },
    )
    .map_err(graph::map_wire_error)?;
    object
        .replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: ZORDER_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(graph::map_core_error)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or(Error::InvalidSource)?;
    info.object_references
        .retain(|identifier| *identifier != drawable);
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|identifier| *identifier != drawable);
    }
    Ok(())
}

fn editable<'a>(
    source: &Package,
    archives: &'a mut Vec<(Arc<str>, Archive)>,
    component: &str,
) -> Result<&'a mut Archive, Error> {
    let position = match archives
        .iter()
        .position(|(name, _)| name.as_ref() == component)
    {
        Some(position) => position,
        None => {
            archives.push((
                Arc::from(component),
                graph::editable_archive(source, component)?,
            ));
            archives.len() - 1
        },
    };
    Ok(&mut archives[position].1)
}

fn remove(source: &Package, selector: DrawableSelector) -> Result<(Package, Diagnostics), Error> {
    let media = resolve(source, selector)?;
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();

    // The anchoring U+FFFC and its attachment entry.
    let (mut storage, _) = objects::body_storage(source)?;
    let at = media.character_index;
    if storage.text.encode_utf16().nth(at as usize) != Some(0xfffc) {
        return Err(Error::InvalidSource);
    }
    let removed = objects::splice_text(source, &mut storage, at, at + 1, "")?;
    if removed != [media.attachment.get()] {
        return Err(Refusal::SharedGraph.into());
    }
    {
        let archive = editable(source, &mut archives, &storage.component)?;
        archive
            .object_mut(storage.identifier.get())
            .ok_or(Error::InvalidSource)?
            .replace_message_pruning_object_references_preserving_header(
                storage.message_index,
                RawMessage {
                    type_: storage.message_type,
                    data: storage.payload.clone(),
                },
                &[media.attachment.get()],
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
    }

    // The document z-order.
    {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let component = index.component_name(index.get(media.zorder)?.component_index)?;
        let archive = editable(source, &mut archives, &component)?;
        zorder_without(archive, media.zorder, media.drawable.get(), archive_limits)?;
    }

    // The graph objects.
    {
        let archive = editable(source, &mut archives, &media.component)?;
        for identifier in &media.objects {
            archive
                .remove_object(*identifier)
                .ok_or(Error::InvalidSource)?;
        }
    }
    let removed_set = media.objects.iter().copied().collect::<HashSet<_>>();
    for (_, archive) in &archives {
        if archive.objects.iter().any(|object| {
            object.archive_info.message_infos.iter().any(|info| {
                info.object_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| &field.object_references),
                    )
                    .any(|identifier| removed_set.contains(identifier))
            })
        }) {
            return Err(Refusal::SharedGraph.into());
        }
    }
    for component in source.state.source.components().iter() {
        if archives
            .iter()
            .any(|(name, _)| name.as_ref() == component.name())
        {
            continue;
        }
        if component.archive().objects.iter().any(|object| {
            object.archive_info.message_infos.iter().any(|info| {
                info.object_references
                    .iter()
                    .chain(
                        info.field_infos
                            .iter()
                            .flat_map(|field| &field.object_references),
                    )
                    .any(|identifier| removed_set.contains(identifier))
            })
        }) {
            return Err(Refusal::SharedGraph.into());
        }
    }

    // Package metadata and data assets.
    let mut deletions = Vec::new();
    let mut removed_assets = 0usize;
    if let Some(metadata) = objects::metadata(source)? {
        let mut payload = metadata.payload.clone();
        for (data, object) in &media.data_references {
            payload = objects::adjust_data_reference(
                &payload,
                DOCUMENT_COMPONENT_IDENTIFIER,
                *data,
                *object,
                false,
            )?;
        }
        for identifier in &media.objects {
            if metadata.uuid_owner(*identifier)? == Some(DOCUMENT_COMPONENT_IDENTIFIER) {
                payload = objects::remove_object_uuid(
                    &payload,
                    DOCUMENT_COMPONENT_IDENTIFIER,
                    *identifier,
                )?;
            }
        }
        let remaining = objects::max_identifier(source, &archives);
        payload = objects::released(&payload, metadata.last, &media.objects, remaining)?;

        let mut referenced = objects::referenced_data(&payload)?;
        for component in source.state.source.components().iter() {
            let staged = archives
                .iter()
                .find(|(name, _)| name.as_ref() == component.name())
                .map_or(component.archive(), |(_, archive)| archive);
            for object in &staged.objects {
                for info in &object.archive_info.message_infos {
                    referenced.extend(info.data_references.iter().copied());
                    referenced.extend(
                        info.field_infos
                            .iter()
                            .flat_map(|field| field.data_references.iter().copied()),
                    );
                }
                if metadata.data_metadata_map.is_some()
                    && object.archive_info.identifier == metadata.data_metadata_map
                {
                    for message in &object.messages {
                        if message.type_ == DATA_METADATA_MAP_MESSAGE_TYPE {
                            referenced.extend(objects::data_metadata_entries(&message.data)?);
                        }
                    }
                }
            }
        }
        let mut candidates = media
            .data_references
            .iter()
            .map(|(data, _)| *data)
            .collect::<Vec<_>>();
        candidates.sort_unstable();
        candidates.dedup();
        for data in candidates {
            if referenced.contains(&data) {
                continue;
            }
            if let Some(member) = objects::data_member(&payload, data)? {
                if source
                    .state
                    .source
                    .package()
                    .iter()
                    .any(|entry| entry.name() == member)
                {
                    deletions.push(member);
                }
            }
            payload = objects::remove_data_info(&payload, data)?;
            removed_assets += 1;
        }
        if payload != metadata.payload {
            let archive = editable(source, &mut archives, &metadata.component)?;
            archive
                .object_mut(metadata.object)
                .ok_or(Error::InvalidSource)?
                .replace_message_preserving_header(
                    metadata.message_index,
                    RawMessage {
                        type_: METADATA_MESSAGE_TYPE,
                        data: payload,
                    },
                    archive_limits,
                )
                .map_err(graph::map_core_error)?;
        }
    }
    let deletions = deletions.iter().map(String::as_str).collect::<Vec<_>>();
    let candidate = graph::publish_changes(source, archives, &[], &[], &deletions)?;
    verify_removed(&candidate, selector, &media)?;
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            removed_objects: media.objects.len(),
            removed_assets,
            ..Diagnostics::default()
        },
    ))
}

fn verify_removed(
    candidate: &Package,
    selector: DrawableSelector,
    media: &MediaGraph,
) -> Result<(), Error> {
    let mut budget = Budget::new(candidate)?;
    let index = ObjectIndex::new(candidate, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    if body
        .anchors
        .iter()
        .any(|anchor| anchor.drawable == media.drawable)
    {
        return Err(Error::Verification);
    }
    // The selector still resolves (the next drawable of the kind) or not;
    // either is valid, so only the graph's absence is checked.
    let _ = selector;
    for identifier in &media.objects {
        if NonZeroU64::new(*identifier).is_some_and(|identifier| index.get(identifier).is_ok()) {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const MOVIE_MESSAGE_TYPE: u32 = 3_007;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const BODY_DUPLICATE_OFFSET: f32 = 12.0;
const AUDIO_DUPLICATE_OFFSET: f32 = 30.0;
const ATTACHMENT_H_OFFSET_FIELD: u32 = 3;
const ATTACHMENT_V_OFFSET_FIELD: u32 = 5;
const STORAGE_ATTACHMENTS_FIELD: u32 = 9;

/// The reference paths of a message type the host's graph clone remaps.
fn reference_paths(message_type: u32) -> Option<&'static [&'static [u32]]> {
    const IMAGE: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[3],
        &[5],
        &[6],
        &[8],
    ];
    const MOVIE: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[1, 10],
        &[1, 11],
        &[2],
        &[10],
        &[11],
        &[19],
    ];
    const STORAGE: &[&[u32]] = &[
        &[2],
        &[5, 1, 2],
        &[7, 1, 2],
        &[8, 1, 2],
        &[9, 1, 2],
        &[11, 1, 2],
        &[12, 1, 2],
        &[15, 1, 2],
        &[16, 1, 2],
        &[17, 1, 2],
        &[18, 1, 2],
        &[21, 1, 2],
        &[22, 1, 2],
        &[23, 1, 2],
        &[27, 1, 2],
        &[28, 1, 2],
        &[25, 1, 2],
        &[26, 1, 2],
    ];
    const ATTACHMENT: &[&[u32]] = &[&[1]];
    // `TSWP.ShapeInfoArchive`: the drawable's parent, comment, title and
    // caption, the shape's style, and both storage slots.
    const SHAPE_INFO: &[&[u32]] = &[
        &[1, 1, 2],
        &[1, 1, 6],
        &[1, 1, 9],
        &[1, 1, 10],
        &[1, 1, 11],
        &[1, 2],
        &[2],
        &[3],
        &[4],
    ];
    match message_type {
        2_011 => Some(SHAPE_INFO),
        IMAGE_MESSAGE_TYPE => Some(IMAGE),
        MOVIE_MESSAGE_TYPE => Some(MOVIE),
        2_001 | 2_022 => Some(STORAGE),
        graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE => Some(ATTACHMENT),
        _ => None,
    }
}

fn remap_paths(
    data: &[u8],
    paths: &[&[u32]],
    remap: &std::collections::HashMap<u64, u64>,
) -> Result<Vec<u8>, Error> {
    let mut output = data.to_vec();
    for path in paths {
        output = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
            &output,
            path,
            |reference| {
                let view = litchi_iwa_common::wire::WireView::parse(
                    reference,
                    litchi_iwa_common::WireLimits::default(),
                )?;
                let current = view
                    .fields()
                    .find(|field| field.number() == 1)
                    .and_then(|field| {
                        litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()
                    })
                    .map(|(value, _)| value);
                let Some(replacement) = current.and_then(|current| remap.get(&current)) else {
                    return Ok(reference.to_vec());
                };
                litchi_iwa_common::wire::patch_varint_field(reference, 1, true, Some(*replacement))
            },
        )
        .map_err(graph::map_wire_error)?;
    }
    Ok(output)
}

/// A clone of one graph object under `remap` (the host's
/// `clone_pages_drawable_graph_object`).
fn clone_object(
    source: &Package,
    object: &litchi_iwa_core::ArchiveObject,
    remap: &std::collections::HashMap<u64, u64>,
) -> Result<litchi_iwa_core::ArchiveObject, Error> {
    let old = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
    let new = *remap.get(&old).ok_or(Error::InvalidSource)?;
    let mut messages = Vec::with_capacity(object.messages.len());
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        let data = match message.type_ {
            CAPTION_INFO_MESSAGE_TYPE => {
                let pairs = remap
                    .iter()
                    .map(|(old, new)| (*old, *new))
                    .collect::<Vec<_>>();
                litchi_iwa_protos::pages_movie_caption_codec::rewrite_caption_info(
                    &message.data,
                    litchi_iwa_protos::pages_movie_caption_codec::CaptionInfoWrite::new(&pairs),
                    litchi_iwa_protos::pages_movie_caption_codec::DecodeOptions::for_source(
                        &message.data,
                    )
                    .with_max_output_bytes(message.data.len().saturating_mul(10)),
                )
                .map_err(|_| Error::InvalidSource)?
            },
            STANDIN_MESSAGE_TYPE => message.data.clone(),
            message_type => match reference_paths(message_type) {
                Some(paths) => remap_paths(&message.data, paths, remap)?,
                None => {
                    if info
                        .object_references
                        .iter()
                        .any(|identifier| remap.contains_key(identifier))
                    {
                        return Err(Refusal::SharedGraph.into());
                    }
                    message.data.clone()
                },
            },
        };
        messages.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    let limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut cloned = litchi_iwa_core::ArchiveObject::new(new, messages, limits)
        .map_err(graph::map_core_error)?;
    cloned.archive_info.should_merge = object.archive_info.should_merge;
    for ((target, original), message) in cloned
        .archive_info
        .message_infos
        .iter_mut()
        .zip(&object.archive_info.message_infos)
        .zip(&cloned.messages)
    {
        let length = u32::try_from(message.data.len()).map_err(|_| Error::InvalidSource)?;
        *target = original.clone();
        target.length = length;
        for identifier in &mut target.object_references {
            if let Some(replacement) = remap.get(identifier) {
                *identifier = *replacement;
            }
        }
        for field in &mut target.field_infos {
            for identifier in &mut field.object_references {
                if let Some(replacement) = remap.get(identifier) {
                    *identifier = *replacement;
                }
            }
        }
    }
    Ok(cloned)
}

/// The host's `update_reference_list` with only an addition.
fn add_reference(references: &mut Vec<u64>, new: u64) {
    if !references.contains(&new) {
        references.push(new);
    }
}

fn reference_payload(identifier: u64) -> Vec<u8> {
    let mut payload = vec![0x08];
    litchi_iwa_common::varint::encode_varint_into(&mut payload, identifier);
    payload
}

/// The body storage payload with a new attachment entry at `anchor`
/// (the host's `add_body_drawable_attachment`).
fn with_attachment_entry(payload: &[u8], anchor: u32, attachment: u64) -> Result<Vec<u8>, Error> {
    let mut entry = Vec::new();
    litchi_iwa_common::wire::append_varint_field(&mut entry, 1, u64::from(anchor))
        .map_err(graph::map_wire_error)?;
    litchi_iwa_common::wire::append_length_delimited_field(
        &mut entry,
        2,
        &reference_payload(attachment),
        litchi_iwa_common::WireLimits::default(),
    )
    .map_err(graph::map_wire_error)?;
    let view =
        litchi_iwa_common::wire::WireView::parse(payload, litchi_iwa_common::WireLimits::default())
            .map_err(graph::map_wire_error)?;
    if graph::unique_field(&view, STORAGE_ATTACHMENTS_FIELD, 2)?.is_none() {
        let mut table = Vec::new();
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut table,
            1,
            &entry,
            litchi_iwa_common::WireLimits::default(),
        )
        .map_err(graph::map_wire_error)?;
        return litchi_iwa_common::wire::patch_length_delimited_field(
            payload,
            STORAGE_ATTACHMENTS_FIELD,
            false,
            Some(&table),
        )
        .map_err(|error| graph::map_wire_error(error).into());
    }
    litchi_iwa_common::wire::transform_length_delimited_field(
        payload,
        STORAGE_ATTACHMENTS_FIELD,
        |table| {
            let mut entries = Vec::new();
            for (order, data) in
                litchi_iwa_common::wire::repeated_length_delimited_payloads(table, 1)?
                    .into_iter()
                    .enumerate()
            {
                let view = litchi_iwa_common::wire::WireView::parse(
                    data,
                    litchi_iwa_common::WireLimits::default(),
                )?;
                let index = view
                    .fields()
                    .find(|field| field.number() == 1)
                    .and_then(|field| {
                        litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()
                    })
                    .map_or(0, |(value, _)| value);
                if index == u64::from(anchor) {
                    return Err(litchi_iwa_common::Error::InvalidFormat(
                        "attachment already at the anchor".to_owned(),
                    ));
                }
                entries.push((index, order, data.to_vec()));
            }
            let order = entries.len();
            entries.push((u64::from(anchor), order, entry.clone()));
            entries.sort_by_key(|(index, order, _)| (*index, *order));
            let entries = entries
                .into_iter()
                .map(|(_, _, data)| data)
                .collect::<Vec<_>>();
            litchi_iwa_common::wire::rewrite_repeated_length_delimited_fields(table, 1, &entries)
        },
    )
    .map_err(|error: litchi_iwa_common::Error| graph::map_wire_error(error).into())
}

fn duplicate(
    source: &Package,
    selector: DrawableSelector,
    anchor: TextPosition,
) -> Result<(Package, Diagnostics), Error> {
    let media = resolve(source, selector)?;
    let anchor = anchor.utf16_index();
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let metadata = objects::metadata(source)?;
    let first = objects::next_identifier(source, metadata.as_ref())?;
    let mut remap = std::collections::HashMap::with_capacity(media.objects.len());
    for (offset, identifier) in media.objects.iter().enumerate() {
        let replacement = first
            .checked_add(offset as u64)
            .ok_or(Error::InvalidSource)?;
        remap.insert(*identifier, replacement);
    }
    let new_drawable = remap[&media.drawable.get()];
    let new_attachment = remap[&media.attachment.get()];
    let offset = if matches!(selector, DrawableSelector::Audio(_)) {
        AUDIO_DUPLICATE_OFFSET
    } else {
        BODY_DUPLICATE_OFFSET
    };
    let position = source
        .body_drawable_properties(selector)
        .map_err(|_| Error::InvalidSource)?
        .geometry()
        .and_then(|geometry| geometry.position)
        .ok_or(Error::InvalidSource)?;
    let moved = litchi_iwa_protos::shape_properties_codec::Point::new(
        position.x + offset,
        position.y + offset,
    );
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    {
        let archive = editable(source, &mut archives, &media.component)?;
        for identifier in &media.objects {
            let object = source
                .state
                .source
                .components()
                .get(media.component.as_ref())
                .and_then(|component| component.archive().object(*identifier))
                .ok_or(Error::InvalidSource)?;
            let mut cloned = clone_object(source, object, &remap)?;
            if *identifier == media.drawable.get() {
                let message_type = graph::message_type(selector);
                let index = cloned
                    .messages
                    .iter()
                    .position(|message| message.type_ == message_type)
                    .ok_or(Error::InvalidSource)?;
                let shape = matches!(
                    selector,
                    DrawableSelector::Shape(_) | DrawableSelector::TextBox(_)
                );
                let rewrite = |drawable: &[u8]| {
                    litchi_iwa_protos::shape_properties_codec::rewrite_drawable_archive(
                            drawable,
                            litchi_iwa_protos::shape_properties_codec::GeometryWrite::preserve()
                                .with_position(Some(moved)),
                            litchi_iwa_protos::shape_properties_codec::DrawablePropertiesWrite::preserve(),
                            litchi_iwa_protos::shape_properties_codec::DecodeOptions::for_source(
                                drawable,
                            ),
                        )
                        .map(litchi_iwa_protos::shape_properties_codec::RewriteOutput::into_bytes)
                        .map_err(|_| {
                            litchi_iwa_common::Error::InvalidFormat("drawable geometry".to_owned())
                        })
                };
                let data = if shape {
                    litchi_iwa_common::wire::transform_length_delimited_field(
                        &cloned.messages[index].data,
                        1,
                        |shape| {
                            litchi_iwa_common::wire::transform_length_delimited_field(
                                shape, 1, rewrite,
                            )
                        },
                    )
                } else {
                    litchi_iwa_common::wire::transform_length_delimited_field(
                        &cloned.messages[index].data,
                        1,
                        rewrite,
                    )
                }
                .map_err(graph::map_wire_error)?;
                cloned
                    .replace_message_preserving_header(
                        index,
                        RawMessage {
                            type_: message_type,
                            data,
                        },
                        archive_limits,
                    )
                    .map_err(graph::map_core_error)?;
            }
            if *identifier == media.attachment.get() {
                let (index, message) =
                    graph::unique_message(&cloned, graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE)?;
                let payload = message.data.clone();
                let view = litchi_iwa_common::wire::WireView::parse(
                    &payload,
                    litchi_iwa_common::WireLimits::default(),
                )
                .map_err(graph::map_wire_error)?;
                let offset_of = |number| -> Result<f32, Error> {
                    let field =
                        graph::unique_field(&view, number, 5)?.ok_or(Error::InvalidSource)?;
                    let bytes: [u8; 4] = field
                        .payload()
                        .try_into()
                        .map_err(|_| Error::InvalidSource)?;
                    Ok(f32::from_le_bytes(bytes) + offset)
                };
                let h = offset_of(ATTACHMENT_H_OFFSET_FIELD)?;
                let v = offset_of(ATTACHMENT_V_OFFSET_FIELD)?;
                let data = litchi_iwa_common::wire::patch_fixed32_field(
                    &payload,
                    ATTACHMENT_H_OFFSET_FIELD,
                    true,
                    Some(h.to_bits()),
                )
                .and_then(|data| {
                    litchi_iwa_common::wire::patch_fixed32_field(
                        &data,
                        ATTACHMENT_V_OFFSET_FIELD,
                        true,
                        Some(v.to_bits()),
                    )
                })
                .map_err(graph::map_wire_error)?;
                cloned
                    .replace_message_preserving_header(
                        index,
                        RawMessage {
                            type_: graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
                            data,
                        },
                        archive_limits,
                    )
                    .map_err(graph::map_core_error)?;
            }
            archive
                .insert_object(cloned, archive_limits)
                .map_err(graph::map_core_error)?;
        }
    }

    // The anchoring U+FFFC and the attachment entry.
    let (mut storage, _) = objects::body_storage(source)?;
    let removed = objects::splice_text(source, &mut storage, anchor, anchor, "\u{fffc}")?;
    if !removed.is_empty() {
        return Err(Error::InvalidSource);
    }
    let payload = with_attachment_entry(&storage.payload, anchor, new_attachment)?;
    {
        let existing = {
            let mut budget = Budget::new(source)?;
            let index = ObjectIndex::new(source, &mut budget)?;
            graph::resolve_body(&index, &mut budget)?
                .anchors
                .iter()
                .map(|anchor| anchor.attachment.get())
                .collect::<HashSet<_>>()
        };
        let archive = editable(source, &mut archives, &storage.component)?;
        let object = archive
            .object_mut(storage.identifier.get())
            .ok_or(Error::InvalidSource)?;
        object
            .replace_message_preserving_header(
                storage.message_index,
                RawMessage {
                    type_: storage.message_type,
                    data: payload,
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(storage.message_index)
            .ok_or(Error::InvalidSource)?;
        add_reference(&mut info.object_references, new_attachment);
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|identifier| existing.contains(identifier))
            {
                add_reference(&mut field.object_references, new_attachment);
            }
        }
    }

    // The document z-order.
    {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let component = index.component_name(index.get(media.zorder)?.component_index)?;
        let archive = editable(source, &mut archives, &component)?;
        let object = archive
            .object_mut(media.zorder.get())
            .ok_or(Error::InvalidSource)?;
        let (message_index, message) = graph::unique_message(object, ZORDER_MESSAGE_TYPE)?;
        let previous = litchi_iwa_common::wire::repeated_length_delimited_payloads(
            &message.data,
            ZORDER_DRAWABLES_FIELD,
        )
        .map_err(graph::map_wire_error)?
        .into_iter()
        .filter_map(|reference| {
            litchi_iwa_common::wire::WireView::parse(
                reference,
                litchi_iwa_common::WireLimits::default(),
            )
            .ok()
            .and_then(|view| {
                view.fields()
                    .find(|field| field.number() == 1)
                    .and_then(|field| {
                        litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok()
                    })
                    .map(|(value, _)| value)
            })
        })
        .collect::<HashSet<_>>();
        let data = litchi_iwa_common::wire::append_repeated_length_delimited_field(
            &message.data,
            ZORDER_DRAWABLES_FIELD,
            &reference_payload(new_drawable),
        )
        .map_err(graph::map_wire_error)?;
        object
            .replace_message_preserving_header(
                message_index,
                RawMessage {
                    type_: ZORDER_MESSAGE_TYPE,
                    data,
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(message_index)
            .ok_or(Error::InvalidSource)?;
        add_reference(&mut info.object_references, new_drawable);
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|identifier| previous.contains(identifier))
            {
                add_reference(&mut field.object_references, new_drawable);
            }
        }
    }

    // Package metadata: watermark, UUIDs, shared data-reference owners.
    if let Some(metadata) = &metadata {
        let last = remap.values().copied().max().ok_or(Error::InvalidSource)?;
        let uuids = media
            .uuid_objects
            .iter()
            .map(|identifier| remap[identifier])
            .collect::<Vec<_>>();
        let mut payload =
            metadata.registered(source, DOCUMENT_COMPONENT_IDENTIFIER, &uuids, last)?;
        for (data, object) in &media.data_references {
            payload = objects::adjust_data_reference(
                &payload,
                DOCUMENT_COMPONENT_IDENTIFIER,
                *data,
                remap[object],
                true,
            )?;
        }
        let archive = editable(source, &mut archives, &metadata.component)?;
        archive
            .object_mut(metadata.object)
            .ok_or(Error::InvalidSource)?
            .replace_message_preserving_header(
                metadata.message_index,
                RawMessage {
                    type_: METADATA_MESSAGE_TYPE,
                    data: payload,
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
    }
    let candidate = graph::publish_archives(source, archives)?;
    let created = resolve_drawable(&candidate, media.selector, NonZeroU64::new(new_drawable))?;
    if !created {
        return Err(Error::Verification);
    }
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            created_objects: media.objects.len(),
            ..Diagnostics::default()
        },
    ))
}

/// Whether the candidate body anchors `drawable` as media of the selected
/// kind.
fn resolve_drawable(
    candidate: &Package,
    selector: DrawableSelector,
    drawable: Option<NonZeroU64>,
) -> Result<bool, Error> {
    let drawable = drawable.ok_or(Error::InvalidSource)?;
    let mut budget = Budget::new(candidate)?;
    let index = ObjectIndex::new(candidate, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let _ = selector;
    Ok(body
        .anchors
        .iter()
        .any(|anchor| anchor.drawable == drawable))
}

const IMAGE_DATA_FIELD: u32 = 11;
const MOVIE_DATA_FIELD: u32 = 14;
const MOVIE_POSTER_FIELD: u32 = 15;
const DATA_INFO_DIGEST_FIELD: u32 = 2;
const DATA_INFO_LENGTH_FIELD: u32 = 18;
const METADATA_DATAS_FIELD: u32 = 4;

/// The data identifier `asset` of the selected drawable references.
fn asset_identifier(
    package: &Package,
    selector: DrawableSelector,
    asset: Asset,
) -> Result<u64, Error> {
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let selected =
        graph::select_drawable(&index, &body, selector, &mut budget)?.ok_or(Refusal::NotFound {
            kind: kind_name(selector),
            position: selector.position(),
        })?;
    let field = match (selector, asset) {
        (DrawableSelector::Image(_), _) => IMAGE_DATA_FIELD,
        (_, Asset::Data) => MOVIE_DATA_FIELD,
        (_, Asset::Poster) => MOVIE_POSTER_FIELD,
    };
    let view = budget.parse(selected.payload, 1)?;
    let reference = graph::unique_field(&view, field, 2)?.ok_or(Error::InvalidSource)?;
    let reference = budget.parse(reference.payload(), 2)?;
    let identifier = graph::unique_field(&reference, 1, 0)?.ok_or(Error::InvalidSource)?;
    Ok(graph::canonical_varint(identifier.payload())?)
}

/// The `DataInfo` with a new digest and materialized length (the host's
/// `patch_data_info`: existing fields rewritten in place, missing ones
/// appended).
fn patched_data_info(info: &[u8], digest: &[u8], length: u64) -> Result<Vec<u8>, Error> {
    let view =
        litchi_iwa_common::wire::WireView::parse(info, litchi_iwa_common::WireLimits::default())
            .map_err(graph::map_wire_error)?;
    let mut output = Vec::with_capacity(info.len() + digest.len());
    let (mut digests, mut lengths) = (0usize, 0usize);
    for field in view.fields() {
        match field.number() {
            DATA_INFO_DIGEST_FIELD => {
                digests += 1;
                litchi_iwa_common::wire::append_length_delimited_field(
                    &mut output,
                    DATA_INFO_DIGEST_FIELD,
                    digest,
                    litchi_iwa_common::WireLimits::default(),
                )
                .map_err(graph::map_wire_error)?;
            },
            DATA_INFO_LENGTH_FIELD => {
                lengths += 1;
                litchi_iwa_common::wire::append_varint_field(
                    &mut output,
                    DATA_INFO_LENGTH_FIELD,
                    length,
                )
                .map_err(graph::map_wire_error)?;
            },
            _ => output.extend_from_slice(field.raw()),
        }
    }
    if digests > 1 || lengths > 1 {
        return Err(Error::InvalidSource);
    }
    if digests == 0 {
        litchi_iwa_common::wire::append_length_delimited_field(
            &mut output,
            DATA_INFO_DIGEST_FIELD,
            digest,
            litchi_iwa_common::WireLimits::default(),
        )
        .map_err(graph::map_wire_error)?;
    }
    if lengths == 0 {
        litchi_iwa_common::wire::append_varint_field(&mut output, DATA_INFO_LENGTH_FIELD, length)
            .map_err(graph::map_wire_error)?;
    }
    Ok(output)
}

fn replace(
    source: &Package,
    identifier: u64,
    data: &[u8],
) -> Result<(Package, Diagnostics), Error> {
    use sha1::Digest as _;
    let metadata = objects::metadata(source)?.ok_or(Error::InvalidSource)?;
    let member =
        objects::data_member(&metadata.payload, identifier)?.ok_or(Refusal::NotMaterialized)?;
    let current = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or(Refusal::NotMaterialized)?;
    let declared = litchi_iwa_common::media::Type::from_extension(
        std::path::Path::new(&member)
            .extension()
            .and_then(|extension| extension.to_str())
            .unwrap_or(""),
    );
    let detected = litchi_iwa_common::media::Type::from_bytes(data);
    if declared != litchi_iwa_common::media::Type::Unknown
        && detected != litchi_iwa_common::media::Type::Unknown
        && declared != detected
    {
        return Err(Refusal::MismatchedType.into());
    }
    let _ = current;
    let digest = sha1::Sha1::digest(data);
    let length = data.len() as u64;
    let mut patched = 0usize;
    let payload = litchi_iwa_common::wire::transform_length_delimited_fields_at_path(
        &metadata.payload,
        &[METADATA_DATAS_FIELD],
        |info| {
            let view = litchi_iwa_common::wire::WireView::parse(
                info,
                litchi_iwa_common::WireLimits::default(),
            )?;
            let id = view
                .fields()
                .find(|field| field.number() == 1)
                .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
                .map(|(value, _)| value);
            if id != Some(identifier) {
                return Ok(info.to_vec());
            }
            patched += 1;
            patched_data_info(info, &digest, length)
                .map_err(|_| litchi_iwa_common::Error::InvalidFormat("DataInfo".to_owned()))
        },
    )
    .map_err(graph::map_wire_error)?;
    if patched != 1 {
        return Err(Error::InvalidSource);
    }
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    let archive = editable(source, &mut archives, &metadata.component)?;
    archive
        .object_mut(metadata.object)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            metadata.message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
            archive_limits,
        )
        .map_err(graph::map_core_error)?;
    let candidate = graph::publish_changes(source, archives, &[], &[(member.as_str(), data)], &[])?;
    let written = candidate
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or(Error::Verification)?;
    if written.data() != data {
        return Err(Error::Verification);
    }
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            ..Diagnostics::default()
        },
    ))
}

/// Remove every unreferenced data record (see [`Edit::prune_unreferenced`]);
/// `None` when there is none.
fn prune_unreferenced(source: &Package) -> Result<Option<(Package, Diagnostics)>, Error> {
    let Some(metadata) = objects::metadata(source)? else {
        return Ok(None);
    };
    let all = source
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .collect::<Vec<_>>();
    let package = source.state.source.package();
    let Some(pruning) = litchi_iwa_style::media::unreferenced_data(
        &metadata.payload,
        all.iter().copied(),
        |path| {
            package
                .iter()
                .find(|entry| entry.name() == path)
                .map(|entry| entry.data().len() as u64)
        },
    )
    .map_err(|_error| Error::InvalidSource)?
    else {
        return Ok(None);
    };
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    let archive = editable(source, &mut archives, &metadata.component)?;
    archive
        .object_mut(metadata.object)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            metadata.message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: pruning.metadata,
            },
            archive_limits,
        )
        .map_err(graph::map_core_error)?;
    let deletions = pruning
        .records
        .iter()
        .filter_map(|(_, member)| member.as_deref())
        .collect::<Vec<_>>();
    let candidate = graph::publish_changes(source, archives, &[], &[], &deletions)?;
    if candidate
        .state
        .source
        .package()
        .iter()
        .any(|entry| deletions.contains(&entry.name()))
    {
        return Err(Error::Verification);
    }
    Ok(Some((
        candidate,
        Diagnostics {
            changed: true,
            ..Diagnostics::default()
        },
    )))
}

/// The assets reachable from the document root (object 1) through object
/// references, through a message data reference or a component
/// data-reference owner of a reachable object.
fn media_assets(package: &Package, root: u64) -> Result<Vec<Embedded>, Error> {
    let Some(metadata) = objects::metadata(package)? else {
        return Ok(Vec::new());
    };
    let all = package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .collect::<Vec<_>>();
    let objects_iter = || all.iter().copied();
    let assets =
        litchi_iwa_style::media::embedded_assets(&metadata.payload, objects_iter(), |path| {
            package
                .state
                .source
                .package()
                .iter()
                .find(|entry| entry.name() == path)
                .map(|entry| entry.data().len() as u64)
        })
        .map_err(|_| Error::InvalidSource)?;
    let mut outgoing = std::collections::HashMap::<u64, Vec<u64>>::new();
    let mut message_data = std::collections::HashMap::<u64, Vec<u64>>::new();
    for object in objects_iter() {
        let identifier = object.archive_info.identifier.ok_or(Error::InvalidSource)?;
        let references = outgoing.entry(identifier).or_default();
        for info in &object.archive_info.message_infos {
            references.extend(info.object_references.iter().copied());
            message_data
                .entry(identifier)
                .or_default()
                .extend(info.data_references.iter().copied());
        }
    }
    let mut reachable = HashSet::new();
    let mut queue = std::collections::VecDeque::from([root]);
    while let Some(identifier) = queue.pop_front() {
        if !reachable.insert(identifier) {
            continue;
        }
        if let Some(references) = outgoing.get(&identifier) {
            queue.extend(references.iter().copied());
        }
    }
    let mut live = HashSet::new();
    for identifier in &reachable {
        if let Some(data) = message_data.get(identifier) {
            live.extend(data.iter().copied());
        }
    }
    for (data, object) in objects::data_owners(&metadata.payload)? {
        if reachable.contains(&object) {
            live.insert(data);
        }
    }
    Ok(assets
        .into_iter()
        .filter(|asset| live.contains(&asset.data().get()))
        .collect())
}

const IMAGE_THUMBNAIL_FIELD: u32 = 12;
const IMAGE_ORIGINAL_SIZE_FIELD: u32 = 4;
const IMAGE_NATURAL_SIZE_FIELD: u32 = 9;
const MOVIE_ORIGINAL_SIZE_FIELD: u32 = 20;
const MOVIE_NATURAL_SIZE_FIELD: u32 = 21;

fn data_reference(
    view: &litchi_iwa_common::wire::WireView<'_>,
    field: u32,
    budget: &mut Budget,
) -> Result<Option<AssetId>, Error> {
    let Some(reference) = graph::unique_field(view, field, 2)? else {
        return Ok(None);
    };
    let reference = budget.parse(reference.payload(), 2)?;
    let identifier = graph::unique_field(&reference, 1, 0)?.ok_or(Error::InvalidSource)?;
    AssetId::new(graph::canonical_varint(identifier.payload())?)
        .map(Some)
        .map_err(|_| Error::InvalidSource)
}

fn size_field(
    view: &litchi_iwa_common::wire::WireView<'_>,
    field: u32,
    budget: &mut Budget,
) -> Result<Option<crate::drawable::Size>, Error> {
    let Some(size) = graph::unique_field(view, field, 2)? else {
        return Ok(None);
    };
    let size = budget.parse(size.payload(), 2)?;
    let float = |number| -> Result<f32, Error> {
        let field = graph::unique_field(&size, number, 5)?.ok_or(Error::InvalidSource)?;
        let bytes: [u8; 4] = field
            .payload()
            .try_into()
            .map_err(|_| Error::InvalidSource)?;
        Ok(f32::from_le_bytes(bytes))
    };
    Ok(Some(crate::drawable::Size {
        width: float(1)?,
        height: float(2)?,
    }))
}

fn media_facts(
    selector: DrawableSelector,
    selected: &graph::Selected<'_>,
    budget: &mut Budget,
) -> Result<Media, Error> {
    let view = budget.parse(selected.payload, 1)?;
    let image = matches!(selector, DrawableSelector::Image(_));
    let data_field = if image {
        IMAGE_DATA_FIELD
    } else {
        MOVIE_DATA_FIELD
    };
    let (original, natural) = if image {
        (IMAGE_ORIGINAL_SIZE_FIELD, IMAGE_NATURAL_SIZE_FIELD)
    } else {
        (MOVIE_ORIGINAL_SIZE_FIELD, MOVIE_NATURAL_SIZE_FIELD)
    };
    let audio = matches!(selector, DrawableSelector::Audio(_));
    Ok(Media {
        selector,
        anchor: TextPosition::from_utf16_index(selected.anchor.character_index as usize)
            .map_err(|_| Error::InvalidSource)?,
        data: data_reference(&view, data_field, budget)?.ok_or(Error::InvalidSource)?,
        thumbnail: if image {
            data_reference(&view, IMAGE_THUMBNAIL_FIELD, budget)?
        } else {
            None
        },
        poster: if image || audio {
            None
        } else {
            data_reference(&view, MOVIE_POSTER_FIELD, budget)?
        },
        original_size: if audio {
            None
        } else {
            size_field(&view, original, budget)?
        },
        natural_size: if audio {
            None
        } else {
            size_field(&view, natural, budget)?
        },
    })
}

const ROOT_MESSAGE_TYPE: u32 = 10_000;
const ROOT_ZORDER_FIELD: u32 = 20;
const THEME_MESSAGE_TYPE: u32 = 10_001;
const STANDARD_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_VERSION: [u32; 3] = [10, 1, 0];

/// A minimal protobuf writer for the fixed native payloads authored here.
#[derive(Default)]
struct Pb(Vec<u8>);

impl Pb {
    fn key(&mut self, field: u32, wire: u8) {
        litchi_iwa_common::varint::encode_varint_into(
            &mut self.0,
            (u64::from(field) << 3) | u64::from(wire),
        );
    }

    fn varint(mut self, field: u32, value: u64) -> Self {
        self.key(field, 0);
        litchi_iwa_common::varint::encode_varint_into(&mut self.0, value);
        self
    }

    fn boolean(self, field: u32, value: bool) -> Self {
        self.varint(field, u64::from(value))
    }

    fn float(mut self, field: u32, value: f32) -> Self {
        self.key(field, 5);
        self.0.extend_from_slice(&value.to_le_bytes());
        self
    }

    fn message(mut self, field: u32, payload: &[u8]) -> Self {
        self.key(field, 2);
        litchi_iwa_common::varint::encode_varint_into(&mut self.0, payload.len() as u64);
        self.0.extend_from_slice(payload);
        self
    }

    fn reference(self, field: u32, identifier: u64) -> Self {
        let reference = Pb::default().varint(1, identifier);
        self.message(field, &reference.0)
    }

    fn pair(self, field: u32, first: f32, second: f32) -> Self {
        let pair = Pb::default().float(1, first).float(2, second);
        self.message(field, &pair.0)
    }
}

struct CreationIds {
    zorder: Option<u64>,
    drawable: u64,
    title: u64,
    caption: u64,
    attachment: u64,
}

#[allow(clippy::too_many_arguments)]
fn drawable_payload(
    ids: &CreationIds,
    body: u64,
    position: crate::drawable::Point,
    size: crate::drawable::Size,
    aspect_ratio_locked: bool,
) -> Vec<u8> {
    let geometry = Pb::default()
        .pair(1, position.x, position.y)
        .pair(2, size.width, size.height)
        .varint(3, 3)
        .float(4, 0.0);
    let wrap = Pb::default()
        .varint(1, 4)
        .varint(2, 2)
        .varint(3, 1)
        .float(4, 12.0)
        .float(5, 0.5)
        .boolean(6, false);
    Pb::default()
        .message(1, &geometry.0)
        .reference(2, body)
        .message(3, &wrap.0)
        .boolean(5, false)
        .boolean(7, aspect_ratio_locked)
        .reference(10, ids.title)
        .reference(11, ids.caption)
        .boolean(12, false)
        .boolean(13, false)
        .0
}

fn new_object(
    source: &Package,
    identifier: u64,
    message_type: u32,
    data: Vec<u8>,
    version: [u32; 3],
    references: Vec<u64>,
    data_references: Vec<u64>,
) -> Result<litchi_iwa_core::ArchiveObject, Error> {
    let mut object = objects::new_object(source, identifier, message_type, data)?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource)?;
    info.versions = version.to_vec().into();
    info.object_references = references;
    info.data_references = data_references;
    Ok(object)
}

/// The first image (7) or movie (8) style preset of the document theme.
fn style_preset(package: &Package, field: u32) -> Result<NonZeroU64, Error> {
    let mut budget = Budget::new(package)?;
    let (theme, _) = graph::root_theme(package, &mut budget)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let theme = index.get(theme.ok_or(Error::InvalidSource)?)?;
    let (_, message) = graph::unique_message(theme.object, THEME_MESSAGE_TYPE)?;
    let wrapper = budget.parse(&message.data, 1)?;
    let base = graph::unique_field(&wrapper, 1, 2)?.ok_or(Error::InvalidSource)?;
    let base = budget.parse(base.payload(), 2)?;
    let drawing = graph::unique_field(&base, 100, 2)?.ok_or(Error::InvalidSource)?;
    let drawing = budget.parse(drawing.payload(), 3)?;
    let preset = drawing
        .fields()
        .find(|candidate| candidate.number() == field)
        .ok_or(Error::InvalidSource)?;
    if preset.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    Ok(graph::parse_reference(preset.payload(), &mut budget)?)
}

/// A new embedded asset's `Data/` file name (`stem-identifier.extension`).
pub(super) fn materialized_name(preferred: &str, identifier: u64) -> Result<String, Error> {
    let path = std::path::Path::new(preferred);
    if path.file_name().and_then(|name| name.to_str()) != Some(preferred)
        || preferred.contains(['\0', '\\'])
    {
        return Err(Refusal::InvalidMedia.into());
    }
    let stem = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.is_empty())
        .ok_or(Refusal::InvalidMedia)?;
    Ok(
        match path.extension().and_then(|extension| extension.to_str()) {
            Some(extension) if !extension.is_empty() => format!("{stem}-{identifier}.{extension}"),
            _ => format!("{stem}-{identifier}"),
        },
    )
}

fn check_type(
    preferred: &str,
    data: &[u8],
    expected: litchi_iwa_common::media::Type,
) -> Result<(), Error> {
    let declared = litchi_iwa_common::media::Type::from_extension(
        std::path::Path::new(preferred)
            .extension()
            .and_then(|extension| extension.to_str())
            .unwrap_or(""),
    );
    let detected = litchi_iwa_common::media::Type::from_bytes(data);
    if declared != expected
        || (detected != litchi_iwa_common::media::Type::Unknown && detected != declared)
    {
        return Err(Refusal::InvalidMedia.into());
    }
    Ok(())
}

#[allow(clippy::too_many_lines)]
fn add(
    source: &Package,
    anchor: TextPosition,
    media: &NewMedia,
) -> Result<(Package, Diagnostics), Error> {
    use litchi_iwa_common::media::Type;
    use sha1::Digest as _;
    let (kind, assets): (DrawableSelector, Vec<(&str, &[u8], Type)>) = match media {
        NewMedia::Image {
            file_name, data, ..
        } => (
            DrawableSelector::image(0),
            vec![(file_name.as_ref(), data.as_ref(), Type::Image)],
        ),
        NewMedia::Movie {
            file_name,
            data,
            poster_name,
            poster,
            ..
        } => (
            DrawableSelector::movie(0),
            vec![
                (file_name.as_ref(), data.as_ref(), Type::Video),
                (poster_name.as_ref(), poster.as_ref(), Type::Image),
            ],
        ),
        NewMedia::Audio {
            file_name, data, ..
        } => (
            DrawableSelector::audio(0),
            vec![(file_name.as_ref(), data.as_ref(), Type::Audio)],
        ),
    };
    for (name, data, expected) in &assets {
        check_type(name, data, *expected)?;
    }
    let anchor = anchor.utf16_index();
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let style = style_preset(
        source,
        if matches!(kind, DrawableSelector::Image(_)) {
            7
        } else {
            8
        },
    )?;
    let metadata = objects::metadata(source)?.ok_or(Error::InvalidSource)?;
    let (body_identifier, left_margin, zorder, body_component) = {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let body = graph::resolve_body(&index, &mut budget)?;
        let zorder = graph::root_zorder(source, &mut budget)?;
        let component = index.component_name(index.get(body.root.body)?.component_index)?;
        (
            body.root.body.get(),
            body.root.left_margin,
            zorder,
            component,
        )
    };
    let first = objects::next_identifier(source, Some(&metadata))?;
    let graph_first = first + u64::from(zorder.is_none());
    let ids = CreationIds {
        zorder: zorder.is_none().then_some(first),
        drawable: graph_first,
        title: graph_first + 1,
        caption: graph_first + 2,
        attachment: graph_first + 3,
    };

    // Data assets.
    let mut metadata_payload = metadata.payload.clone();
    let mut next_data = objects::max_data_identifier(&metadata_payload)?;
    let mut members = Vec::new();
    let mut data_ids = Vec::new();
    for (name, data, _) in &assets {
        next_data = next_data.checked_add(1).ok_or(Error::InvalidSource)?;
        let file_name = materialized_name(name, next_data)?;
        let member = format!("Data/{file_name}");
        if source
            .state
            .source
            .package()
            .iter()
            .any(|entry| entry.name() == member)
        {
            return Err(Error::InvalidSource);
        }
        metadata_payload = objects::append_data_info(
            &metadata_payload,
            next_data,
            &sha1::Sha1::digest(data),
            name,
            &file_name,
            data.len() as u64,
        )?;
        members.push((member, *data));
        data_ids.push(next_data);
    }

    // The drawable graph.
    let (position, size, aspect) = match media {
        NewMedia::Image { options, .. } => (options.position(), options.size(), false),
        NewMedia::Movie { options, .. } => (options.position(), options.size(), true),
        NewMedia::Audio { options, .. } => (
            options.position(),
            crate::drawable::Size {
                width: 0.0,
                height: 0.0,
            },
            true,
        ),
    };
    let drawable = drawable_payload(&ids, body_identifier, position, size, aspect);
    let data_reference = |identifier: u64| Pb::default().varint(1, identifier).0;
    let payload = match media {
        NewMedia::Image { options, .. } => {
            let natural = options.natural_size();
            Pb::default()
                .message(1, &drawable)
                .reference(3, style.get())
                .pair(4, natural.width, natural.height)
                .varint(7, 0)
                .pair(9, natural.width, natural.height)
                .message(11, &data_reference(data_ids[0]))
                .0
        },
        _ => {
            let audio = matches!(media, NewMedia::Audio { .. });
            let (duration, natural) = match media {
                NewMedia::Movie { options, .. } => {
                    (options.duration_seconds(), options.natural_size())
                },
                NewMedia::Audio { options, .. } => (
                    options.duration_seconds(),
                    crate::drawable::Size {
                        width: 0.0,
                        height: 0.0,
                    },
                ),
                NewMedia::Image { .. } => return Err(Error::InvalidSource),
            };
            let mut payload = Pb::default()
                .message(1, &drawable)
                .float(3, 0.0)
                .float(4, duration)
                .float(5, 0.0)
                .float(7, 1.0)
                .boolean(9, audio)
                .varint(13, 0)
                .message(14, &data_reference(data_ids[0]));
            if !audio {
                payload = payload.message(15, &data_reference(data_ids[1]));
            }
            payload
                .boolean(18, false)
                .reference(19, style.get())
                .pair(20, natural.width, natural.height)
                .pair(21, natural.width, natural.height)
                .boolean(23, !audio)
                .varint(24, 0)
                .0
        },
    };
    let message_type = graph::message_type(kind);
    let attachment = Pb::default()
        .reference(1, ids.drawable)
        .varint(2, 0)
        .float(3, position.x - left_margin.unwrap_or_default())
        .varint(4, 1)
        .float(5, position.y)
        .0;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    {
        let archive = editable(source, &mut archives, &body_component)?;
        if let Some(zorder) = ids.zorder {
            create_zorder(source, archive, zorder, archive_limits)?;
        }
        for object in [
            new_object(
                source,
                ids.drawable,
                message_type,
                payload,
                STANDARD_VERSION,
                vec![ids.title, ids.caption, style.get()],
                data_ids.clone(),
            )?,
            new_object(
                source,
                ids.title,
                STANDIN_MESSAGE_TYPE,
                Vec::new(),
                STANDIN_VERSION,
                Vec::new(),
                Vec::new(),
            )?,
            new_object(
                source,
                ids.caption,
                STANDIN_MESSAGE_TYPE,
                Vec::new(),
                STANDIN_VERSION,
                Vec::new(),
                Vec::new(),
            )?,
            new_object(
                source,
                ids.attachment,
                graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
                attachment,
                STANDARD_VERSION,
                vec![ids.drawable],
                Vec::new(),
            )?,
        ] {
            archive
                .insert_object(object, archive_limits)
                .map_err(graph::map_core_error)?;
        }
    }

    // The U+FFFC, the attachment entry and the z-order entry.
    insert_anchor(
        source,
        &mut archives,
        anchor,
        ids.attachment,
        archive_limits,
    )?;
    let zorder_identifier = zorder.map_or(ids.zorder.unwrap_or(0), NonZeroU64::get);
    append_zorder(
        source,
        &mut archives,
        zorder_identifier,
        &body_component,
        ids.drawable,
        archive_limits,
    )?;

    // Package metadata.
    let staged_metadata = objects::Metadata::with_payload(&metadata, metadata_payload);
    let mut payload = staged_metadata.registered(
        source,
        DOCUMENT_COMPONENT_IDENTIFIER,
        &[ids.drawable, ids.title, ids.caption],
        ids.attachment,
    )?;
    for data in &data_ids {
        payload = objects::adjust_data_reference(
            &payload,
            DOCUMENT_COMPONENT_IDENTIFIER,
            *data,
            ids.drawable,
            true,
        )?;
    }
    {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let style_component = index.component_name(index.get(style)?.component_index)?;
        if let Some(component) = metadata.component_for_entry(&style_component)?
            && component != DOCUMENT_COMPONENT_IDENTIFIER
        {
            payload = objects::add_external_reference(
                &payload,
                DOCUMENT_COMPONENT_IDENTIFIER,
                component,
                style.get(),
            )?;
        }
    }
    {
        let archive = editable(source, &mut archives, &metadata.component)?;
        archive
            .object_mut(metadata.object)
            .ok_or(Error::InvalidSource)?
            .replace_message_preserving_header(
                metadata.message_index,
                RawMessage {
                    type_: METADATA_MESSAGE_TYPE,
                    data: payload,
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
    }
    let insertions = members
        .iter()
        .map(|(name, data)| (name.as_str(), *data))
        .collect::<Vec<_>>();
    let candidate = graph::publish_changes(source, archives, &insertions, &[], &[])?;
    if !resolve_drawable(&candidate, kind, NonZeroU64::new(ids.drawable))? {
        return Err(Error::Verification);
    }
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            created_objects: 4 + usize::from(ids.zorder.is_some()),
            ..Diagnostics::default()
        },
    ))
}

/// The body storage's text context a new drawable storage inherits (the
/// host's `body_text_storage`): its stylesheet, first paragraph and list
/// styles and first language.
struct BodyTextContext {
    stylesheet: u64,
    paragraph: Option<u64>,
    list: Option<u64>,
    language: Option<Vec<u8>>,
}

fn first_table_entry(
    storage: &[u8],
    field: u32,
    budget: &mut Budget,
) -> Result<Option<Vec<u8>>, Error> {
    let view = budget.parse(storage, 1)?;
    let Some(table) = graph::unique_field(&view, field, 2)? else {
        return Ok(None);
    };
    let table = budget.parse(table.payload(), 2)?;
    for entry in table.fields().filter(|entry| entry.number() == 1) {
        let entry = budget.parse(entry.payload(), 3)?;
        if let Some(object) = graph::unique_field(&entry, 2, 2)? {
            return Ok(Some(object.payload().to_vec()));
        }
    }
    Ok(None)
}

fn body_text_context(source: &Package) -> Result<BodyTextContext, Error> {
    let (storage, mut budget) = objects::body_storage(source)?;
    let view = budget.parse(&storage.payload, 1)?;
    let stylesheet = graph::unique_field(&view, 2, 2)?.ok_or(Error::InvalidSource)?;
    let stylesheet = graph::parse_reference(stylesheet.payload(), &mut budget)?.get();
    let reference = |payload: Option<Vec<u8>>, budget: &mut Budget| -> Result<Option<u64>, Error> {
        payload
            .map(|payload| graph::parse_reference(&payload, budget).map(NonZeroU64::get))
            .transpose()
            .map_err(Error::from)
    };
    let paragraph = first_table_entry(&storage.payload, 5, &mut budget)?;
    let paragraph = reference(paragraph, &mut budget)?;
    let list = first_table_entry(&storage.payload, 7, &mut budget)?;
    let list = reference(list, &mut budget)?;
    let language = first_table_entry(&storage.payload, 19, &mut budget)?;
    Ok(BodyTextContext {
        stylesheet,
        paragraph,
        list,
        language,
    })
}

/// The host's `body_text_storage` payload and its object references.
fn text_storage_payload(text: &str, context: &BodyTextContext) -> (Vec<u8>, Vec<u64>) {
    let attribute = |object: Option<u64>| {
        let mut entry = Pb::default().varint(1, 0);
        if let Some(object) = object {
            entry = entry.reference(2, object);
        }
        Pb::default().message(1, &entry.0)
    };
    let zero = || Pb::default().message(1, &Pb::default().varint(1, 0).varint(2, 0).varint(3, 0).0);
    let mut storage = Pb::default()
        .reference(2, context.stylesheet)
        .message(3, text.as_bytes())
        .boolean(10, true)
        .message(5, &attribute(context.paragraph).0)
        .message(6, &zero().0)
        .message(7, &attribute(context.list).0)
        .message(14, &zero().0);
    if let Some(language) = &context.language {
        storage = storage.message(
            19,
            &Pb::default()
                .message(1, &Pb::default().varint(1, 0).message(2, language).0)
                .0,
        );
    }
    storage = storage
        .message(24, &zero().0)
        .message(28, &attribute(None).0);
    let mut references: Vec<u64> = [Some(context.stylesheet), context.paragraph, context.list]
        .into_iter()
        .flatten()
        .collect();
    references.sort_unstable();
    references.dedup();
    (storage.0, references)
}

/// Create a body text box, preset shape or line (the host's
/// `add_text_box` and `add_body_shape_path`): the shape, two caption
/// stand-ins, its own storage and the body attachment, under fresh
/// identifiers after an optional new z-order; the U+FFFC and attachment
/// entry; the z-order entry; UUIDs, the style's external reference and the
/// identifier watermark.
#[allow(
    clippy::too_many_lines,
    reason = "one creation: graph, anchor, metadata"
)]
fn add_text_shape(
    source: &Package,
    anchor: TextPosition,
    shape: &NewTextShape,
) -> Result<(Package, Diagnostics), Error> {
    let anchor = anchor.utf16_index();
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let style = style_preset(source, if shape.text_box { 6 } else { 5 })
        .map_err(|_| Refusal::MissingThemeStyle)?;
    let context = body_text_context(source)?;
    let metadata = objects::metadata(source)?.ok_or(Error::InvalidSource)?;
    let (body_identifier, left_margin, zorder, body_component) = {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let body = graph::resolve_body(&index, &mut budget)?;
        let zorder = graph::root_zorder(source, &mut budget)?;
        let component = index.component_name(index.get(body.root.body)?.component_index)?;
        (
            body.root.body.get(),
            body.root.left_margin,
            zorder,
            component,
        )
    };
    let first = objects::next_identifier(source, Some(&metadata))?;
    let graph_first = first + u64::from(zorder.is_none());
    // The host's order: drawable, caption, title, storage, attachment.
    let (drawable, caption, title, storage, attachment) = (
        graph_first,
        graph_first + 1,
        graph_first + 2,
        graph_first + 3,
        graph_first + 4,
    );
    let ids = CreationIds {
        zorder: zorder.is_none().then_some(first),
        drawable,
        title,
        caption,
        attachment,
    };
    let frame = Pb::default()
        .pair(1, shape.position.x, shape.position.y)
        .pair(2, shape.size.width, shape.size.height)
        .varint(3, 3)
        .float(4, shape.angle);
    let wrap = Pb::default()
        .varint(1, 4)
        .varint(2, 2)
        .varint(3, 1)
        .float(4, 12.0)
        .float(5, 0.5)
        .boolean(6, false);
    let drawable_archive = Pb::default()
        .message(1, &frame.0)
        .reference(2, body_identifier)
        .message(3, &wrap.0)
        .boolean(5, false)
        .boolean(7, false)
        .reference(10, title)
        .reference(11, caption)
        .boolean(12, false)
        .boolean(13, false);
    let shape_archive = Pb::default()
        .message(1, &drawable_archive.0)
        .reference(2, style.get())
        .message(3, &shape.path)
        .float(6, 0.0);
    let shape_info = Pb::default()
        .message(1, &shape_archive.0)
        .reference(2, storage)
        .reference(4, storage)
        .boolean(6, shape.text_box)
        .0;
    let (storage_payload, storage_references) = text_storage_payload(&shape.text, &context);
    let attachment_payload = Pb::default()
        .reference(1, drawable)
        .varint(2, 0)
        .float(3, shape.position.x - left_margin.unwrap_or_default())
        .varint(4, 1)
        .float(5, shape.position.y)
        .0;
    let mut archives: Vec<(Arc<str>, Archive)> = Vec::new();
    {
        let archive = editable(source, &mut archives, &body_component)?;
        if let Some(zorder) = ids.zorder {
            create_zorder(source, archive, zorder, archive_limits)?;
        }
        for object in [
            new_object(
                source,
                drawable,
                graph::message_type(DrawableSelector::shape(0)),
                shape_info,
                STANDARD_VERSION,
                vec![caption, title, style.get(), storage],
                Vec::new(),
            )?,
            new_object(
                source,
                caption,
                STANDIN_MESSAGE_TYPE,
                Vec::new(),
                STANDIN_VERSION,
                Vec::new(),
                Vec::new(),
            )?,
            new_object(
                source,
                title,
                STANDIN_MESSAGE_TYPE,
                Vec::new(),
                STANDIN_VERSION,
                Vec::new(),
                Vec::new(),
            )?,
            new_object(
                source,
                storage,
                2_001,
                storage_payload,
                STANDARD_VERSION,
                storage_references,
                Vec::new(),
            )?,
            new_object(
                source,
                attachment,
                graph::DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
                attachment_payload,
                STANDARD_VERSION,
                vec![drawable],
                Vec::new(),
            )?,
        ] {
            archive
                .insert_object(object, archive_limits)
                .map_err(graph::map_core_error)?;
        }
    }
    insert_anchor(source, &mut archives, anchor, attachment, archive_limits)?;
    let zorder_identifier = zorder.map_or(ids.zorder.unwrap_or(0), NonZeroU64::get);
    append_zorder(
        source,
        &mut archives,
        zorder_identifier,
        &body_component,
        drawable,
        archive_limits,
    )?;
    let mut payload = metadata.registered(
        source,
        DOCUMENT_COMPONENT_IDENTIFIER,
        &[drawable, caption, title, storage],
        attachment,
    )?;
    {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        let style_component = index.component_name(index.get(style)?.component_index)?;
        if let Some(component) = metadata.component_for_entry(&style_component)?
            && component != DOCUMENT_COMPONENT_IDENTIFIER
        {
            payload = objects::add_external_reference(
                &payload,
                DOCUMENT_COMPONENT_IDENTIFIER,
                component,
                style.get(),
            )?;
        }
    }
    {
        let archive = editable(source, &mut archives, &metadata.component)?;
        archive
            .object_mut(metadata.object)
            .ok_or(Error::InvalidSource)?
            .replace_message_preserving_header(
                metadata.message_index,
                RawMessage {
                    type_: METADATA_MESSAGE_TYPE,
                    data: payload,
                },
                archive_limits,
            )
            .map_err(graph::map_core_error)?;
    }
    let candidate = graph::publish_changes(source, archives, &[], &[], &[])?;
    let kind = if shape.text_box {
        DrawableSelector::text_box(0)
    } else {
        DrawableSelector::shape(0)
    };
    if !resolve_drawable(&candidate, kind, NonZeroU64::new(drawable))? {
        return Err(Error::Verification);
    }
    Ok((
        candidate,
        Diagnostics {
            changed: true,
            created_objects: 5 + usize::from(ids.zorder.is_some()),
            ..Diagnostics::default()
        },
    ))
}

/// Replace the text of the text box anchored at UTF-16 `anchor` (a copy
/// just made) through the shared plain-text family.
fn set_copy_text(package: &Package, anchor: TextPosition, text: &str) -> Result<Package, Error> {
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let mut position = 0;
    loop {
        let selector = DrawableSelector::text_box(position);
        let selected = graph::select_drawable(&index, &body, selector, &mut budget)?
            .ok_or(Error::Verification)?;
        if selected.anchor.character_index == anchor.utf16_index() {
            let mut edit = package
                .edit_body_drawable_text(selector)
                .map_err(|_| Error::Verification)?;
            edit.set(text).map_err(|_| Error::Verification)?;
            return Ok(edit
                .commit()
                .map_err(|_| Error::Verification)?
                .into_package());
        }
        position += 1;
    }
}

/// A new empty drawable z-order object, referenced from the document root.
fn create_zorder(
    source: &Package,
    archive: &mut Archive,
    zorder: u64,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    archive
        .insert_object(
            new_object(
                source,
                zorder,
                ZORDER_MESSAGE_TYPE,
                Vec::new(),
                STANDARD_VERSION,
                Vec::new(),
                Vec::new(),
            )?,
            archive_limits,
        )
        .map_err(graph::map_core_error)?;
    let root = archive
        .object_mut(graph::ROOT_OBJECT_IDENTIFIER)
        .ok_or(Error::InvalidSource)?;
    let (index, message) = graph::unique_message(root, ROOT_MESSAGE_TYPE)?;
    let data = litchi_iwa_common::wire::patch_length_delimited_field(
        &message.data,
        ROOT_ZORDER_FIELD,
        false,
        Some(&Pb::default().varint(1, zorder).0),
    )
    .map_err(graph::map_wire_error)?;
    root.replace_message_preserving_header(
        index,
        RawMessage {
            type_: ROOT_MESSAGE_TYPE,
            data,
        },
        archive_limits,
    )
    .map_err(graph::map_core_error)?;
    add_reference(
        &mut root
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or(Error::InvalidSource)?
            .object_references,
        zorder,
    );
    Ok(())
}

/// Insert a U+FFFC at `anchor` and its attachment entry (shared by
/// duplication and creation).
fn insert_anchor(
    source: &Package,
    archives: &mut Vec<(Arc<str>, Archive)>,
    anchor: u32,
    attachment: u64,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    let (mut storage, _) = objects::body_storage(source)?;
    let removed = objects::splice_text(source, &mut storage, anchor, anchor, "\u{fffc}")?;
    if !removed.is_empty() {
        return Err(Error::InvalidSource);
    }
    let payload = with_attachment_entry(&storage.payload, anchor, attachment)?;
    let existing = {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        graph::resolve_body(&index, &mut budget)?
            .anchors
            .iter()
            .map(|anchor| anchor.attachment.get())
            .collect::<HashSet<_>>()
    };
    let archive = editable(source, archives, &storage.component)?;
    let object = archive
        .object_mut(storage.identifier.get())
        .ok_or(Error::InvalidSource)?;
    object
        .replace_message_preserving_header(
            storage.message_index,
            RawMessage {
                type_: storage.message_type,
                data: payload,
            },
            archive_limits,
        )
        .map_err(graph::map_core_error)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(storage.message_index)
        .ok_or(Error::InvalidSource)?;
    add_reference(&mut info.object_references, attachment);
    for field in &mut info.field_infos {
        if field
            .object_references
            .iter()
            .any(|identifier| existing.contains(identifier))
        {
            add_reference(&mut field.object_references, attachment);
        }
    }
    Ok(())
}

/// Append `drawable` to the document z-order (shared by duplication and
/// creation); `fallback` is the component of a z-order created in this
/// candidate.
fn append_zorder(
    source: &Package,
    archives: &mut Vec<(Arc<str>, Archive)>,
    zorder: u64,
    fallback: &str,
    drawable: u64,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    let component = {
        let mut budget = Budget::new(source)?;
        let index = ObjectIndex::new(source, &mut budget)?;
        match NonZeroU64::new(zorder).map(|identifier| index.get(identifier)) {
            Some(Ok(located)) => index.component_name(located.component_index)?,
            _ => Arc::from(fallback),
        }
    };
    let archive = editable(source, archives, &component)?;
    let object = archive.object_mut(zorder).ok_or(Error::InvalidSource)?;
    let (message_index, message) = graph::unique_message(object, ZORDER_MESSAGE_TYPE)?;
    let previous = litchi_iwa_common::wire::repeated_length_delimited_payloads(
        &message.data,
        ZORDER_DRAWABLES_FIELD,
    )
    .map_err(graph::map_wire_error)?
    .into_iter()
    .filter_map(|reference| {
        litchi_iwa_common::wire::WireView::parse(
            reference,
            litchi_iwa_common::WireLimits::default(),
        )
        .ok()
        .and_then(|view| {
            view.fields()
                .find(|field| field.number() == 1)
                .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
                .map(|(value, _)| value)
        })
    })
    .collect::<HashSet<_>>();
    let data = litchi_iwa_common::wire::append_repeated_length_delimited_field(
        &message.data,
        ZORDER_DRAWABLES_FIELD,
        &reference_payload(drawable),
    )
    .map_err(graph::map_wire_error)?;
    object
        .replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: ZORDER_MESSAGE_TYPE,
                data,
            },
            archive_limits,
        )
        .map_err(graph::map_core_error)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(message_index)
        .ok_or(Error::InvalidSource)?;
    add_reference(&mut info.object_references, drawable);
    for field in &mut info.field_infos {
        if field
            .object_references
            .iter()
            .any(|identifier| previous.contains(identifier))
        {
            add_reference(&mut field.object_references, drawable);
        }
    }
    Ok(())
}
