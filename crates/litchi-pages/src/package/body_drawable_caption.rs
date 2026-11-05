//! Selector-first title and caption transactions of body images, movies and
//! shapes.
//!
//! A drawable's title and caption are references (`TSD.DrawableArchive`
//! fields 10 and 11). A present label is a `TSA.CaptionInfoArchive` with a
//! private text storage, shape style and placement in the drawable's
//! component; an absent label is an empty `TSD.StandinCaptionArchive`.
//! Setting a present label rewrites its storage text; setting an absent one
//! authors the native inline caption graph (the shared
//! `keynote_chart_caption_graph_codec`, the same four payloads the host
//! writes) and points the drawable at it; clearing a present label points
//! the drawable at a fresh stand-in and keeps the old graph, as the
//! application does for undo. Created objects take identifiers above every
//! object and metadata registration, are registered with UUIDs in the
//! document component and raise the metadata watermark (record 0371).

use std::fmt;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::ExactArtifacts;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage};
use litchi_iwa_protos::keynote_chart_caption_graph_codec as graph_codec;
use litchi_iwa_style::caption;
use thiserror::Error as ThisError;

use super::Package;
use super::body_chart_arrangement::{self as arrangement, ArrangementBudget};
use super::body_drawable_graph::{
    self as graph, Budget, GraphError, GraphLimit, ObjectIndex, kind_name,
};
use super::body_storage_objects as objects;
use crate::selector::{BodyChartSelector, DrawableSelector};

use caption::{
    CAPTION_INFO_MESSAGE_TYPE, CAPTION_PLACEMENT_MESSAGE_TYPE, SHAPE_STYLE_MESSAGE_TYPE,
    STANDIN_MESSAGE_TYPE, STORAGE_MESSAGE_TYPE,
};
const THEME_MESSAGE_TYPE: u32 = 10_001;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const THEME_SUPER_FIELD: u32 = 1;
const THEME_DOCUMENT_STYLESHEET_FIELD: u32 = 4;
/// The document component, which registers created objects' UUIDs.
const DOCUMENT_COMPONENT_IDENTIFIER: u64 = 1;

/// Which of a drawable's two labels an operation targets.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// The label above the drawable.
    Title,
    /// The label below the drawable.
    Caption,
}

impl Kind {
    /// The shared engine's label.
    const fn label(self) -> caption::Label {
        match self {
            Self::Title => caption::Label::Title,
            Self::Caption => caption::Label::Caption,
        }
    }
}

/// A finite resource charged by one body-drawable label operation.
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

/// Failure from a body-drawable label read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies
    /// it (record 0351 §2.4).
    #[error(transparent)]
    Refused(Refusal),
    /// The rooted drawable graph or its label graph is malformed.
    #[error("the selected Pages body-drawable label source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Pages body-drawable label {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for a Pages body-drawable label")]
    Allocation {
        /// Requested units.
        amount: usize,
    },
    /// Reopening the candidate did not reproduce the requested labels.
    #[error("the edited Pages body-drawable label failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body-drawable label patch does not match the exact source package")]
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
    /// The selected kind has no title or caption in this family (audio).
    #[error("a Pages body {kind} has no title or caption")]
    NoLabel {
        /// The selected kind.
        kind: &'static str,
    },
    /// The text contains an object-replacement, footnote or section-break
    /// marker, which only object families may write.
    #[error("the label text contains a reserved native text marker")]
    ReservedMarker,
    /// The present label's text references an object that replacing the
    /// text would drop.
    #[error("the label text references a native object")]
    DependentContent,
    /// Authoring a new label needs a displayed width, and the drawable has
    /// none (or a non-positive one).
    #[error("the drawable has no displayed width for a new label")]
    NoWidth,
    /// A second operation was staged on one edit.
    #[error("one label edit stages one operation")]
    OperationAlreadyStaged,
    /// The source does not keep an exact artifact for a changed publication.
    #[error("this Pages source does not support exact body-drawable label edits")]
    UnsupportedSource,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound { .. } => RefusalKind::NotFound,
            Self::NoLabel { .. } => RefusalKind::UnsupportedTarget,
            Self::ReservedMarker => RefusalKind::InvalidValue,
            Self::DependentContent | Self::NoWidth | Self::UnsupportedSource => {
                RefusalKind::UnsupportedContent
            },
            Self::OperationAlreadyStaged => RefusalKind::UnsupportedOperation,
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

/// One label change staged against an immutable package.
pub struct Edit<'a> {
    source: &'a Package,
    selector: DrawableSelector,
    slot: Slot,
    operation: Option<Option<Box<str>>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("selector", &self.selector)
            .field("kind", &self.slot.kind)
            .field("staged", &self.operation.is_some())
            .finish_non_exhaustive()
    }
}

/// Validate one staged label text (shared with the chart caption family).
pub(super) fn stage(
    operation: &mut Option<Option<Box<str>>>,
    text: Option<&str>,
) -> Result<(), Error> {
    if operation.is_some() {
        return Err(Refusal::OperationAlreadyStaged.into());
    }
    if text.is_some_and(|text| {
        text.chars()
            .any(|character| matches!(character, '\u{fffc}' | '\u{000e}' | '\u{0004}'))
    }) {
        return Err(Refusal::ReservedMarker.into());
    }
    *operation = Some(text.map(Box::from));
    Ok(())
}

impl Edit<'_> {
    /// The label text when the edit began; `None` when the label is absent.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.slot.text.as_deref()
    }

    /// Stage the label's text, creating the label when it is absent.
    ///
    /// # Errors
    ///
    /// Refuses a reserved text marker and a second operation.
    pub fn set(&mut self, text: &str) -> Result<&mut Self, Error> {
        stage(&mut self.operation, Some(text))?;
        Ok(self)
    }

    /// Stage the label's removal; removing an absent label changes nothing.
    ///
    /// # Errors
    ///
    /// Refuses a second operation.
    pub fn clear(&mut self) -> Result<&mut Self, Error> {
        stage(&mut self.operation, None)?;
        Ok(self)
    }

    /// Validate and publish the staged change.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// Exact-source reversible body-drawable label patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    selector: DrawableSelector,
    kind: Kind,
    before: Option<Arc<str>>,
    after: Option<Arc<str>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("selector", &self.selector)
            .field("kind", &self.kind)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// The selected drawable.
    #[must_use]
    pub const fn selector(&self) -> DrawableSelector {
        self.selector
    }

    /// The selected label.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// The label text required before this patch applies.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// The label text this patch produces.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Whether both the label and the exact artifact are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// The exact target-to-source inverse.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selector: self.selector,
            kind: self.kind,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Content-free diagnostics of one body-drawable label publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    pub(super) changed: bool,
    pub(super) created_objects: usize,
}

impl Diagnostics {
    /// Whether package bytes and the label changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Native objects the publication created (a new label graph or a
    /// stand-in).
    #[must_use]
    pub const fn created_objects(self) -> usize {
        self.created_objects
    }
}

/// The validated immutable result of one body-drawable label transaction.
#[must_use = "a body-drawable label commit contains the validated package snapshot"]
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

/// What a label belongs to: a body image, movie or shape, or a body chart
/// (caption only).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Selection {
    Drawable(DrawableSelector),
    Chart(BodyChartSelector),
}

impl Selection {
    const fn nesting(self) -> caption::Nesting {
        match self {
            Self::Drawable(DrawableSelector::Shape(_) | DrawableSelector::TextBox(_)) => {
                caption::Nesting::Shape
            },
            Self::Drawable(_) => caption::Nesting::Drawable,
            Self::Chart(_) => caption::Nesting::Chart,
        }
    }
}

/// The private native target of one label.
#[derive(Debug, Clone, PartialEq)]
pub(super) struct Slot {
    selection: Selection,
    pub(super) kind: Kind,
    drawable: NonZeroU64,
    component: Arc<str>,
    message_index: usize,
    message_type: u32,
    reference: NonZeroU64,
    storage: Option<NonZeroU64>,
    pub(super) text: Option<Arc<str>>,
    width: Option<f32>,
    /// The label's private objects: the stand-in, or the CaptionInfo, its
    /// style, storage and placement.
    objects: Vec<NonZeroU64>,
}

impl Package {
    /// Read one label (title or caption) of a body image, movie or shape;
    /// `None` when the label is absent.
    ///
    /// # Errors
    ///
    /// Returns [`Error`] when no drawable matches, the kind has no labels,
    /// or the graph is invalid.
    pub fn body_drawable_caption(
        &self,
        selector: impl Into<DrawableSelector>,
        kind: Kind,
    ) -> Result<Option<String>, Error> {
        Ok(resolve(self, selector.into(), kind)?
            .text
            .map(|text| text.as_ref().to_owned()))
    }

    /// Begin a change of one label of a body image, movie or shape.
    ///
    /// # Errors
    ///
    /// As [`Self::body_drawable_caption`].
    pub fn edit_body_drawable_caption(
        &self,
        selector: impl Into<DrawableSelector>,
        kind: Kind,
    ) -> Result<Edit<'_>, Error> {
        let selector = selector.into();
        Ok(Edit {
            source: self,
            selector,
            slot: resolve(self, selector, kind)?,
            operation: None,
        })
    }

    /// Replay an exact-source body-drawable label patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] when the patch was made against
    /// another artifact or the label no longer matches its `before` state.
    pub fn apply_body_drawable_caption(&self, patch: &Patch) -> Result<Commit, Error> {
        if !patch
            .artifacts
            .authorizes_source(&self.state.source.shared_source())
        {
            return Err(Error::PatchConflict);
        }
        let current = resolve(self, patch.selector, patch.kind)?;
        if current.text != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics {
                    changed: false,
                    created_objects: 0,
                },
            });
        }
        let candidate = graph::reopen(self, patch.artifacts.target())?;
        if resolve(&candidate, patch.selector, patch.kind)?.text != patch.after {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                created_objects: 0,
            },
        })
    }
}

/// The private objects of one label of a body image, movie, audio clip or
/// shape (an audio clip's labels are stand-ins), for the media lifecycle.
pub(crate) fn label_objects(
    package: &Package,
    selector: DrawableSelector,
    kind: Kind,
) -> Result<Vec<NonZeroU64>, Error> {
    Ok(resolve_any(package, selector, kind)?.objects)
}

fn resolve(package: &Package, selector: DrawableSelector, kind: Kind) -> Result<Slot, Error> {
    if matches!(selector, DrawableSelector::Audio(_)) {
        return Err(Refusal::NoLabel {
            kind: kind_name(selector),
        }
        .into());
    }
    resolve_any(package, selector, kind)
}

/// The drawable a label belongs to, after its selection proof.
struct Owner<'a> {
    drawable: NonZeroU64,
    component_index: usize,
    /// The `TSD.DrawableArchive` inside the drawable's message.
    archive: &'a [u8],
    message_index: usize,
    message_type: u32,
}

fn resolve_any(package: &Package, selector: DrawableSelector, kind: Kind) -> Result<Slot, Error> {
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let body = graph::resolve_body(&index, &mut budget)?;
    let selected =
        graph::select_drawable(&index, &body, selector, &mut budget)?.ok_or(Refusal::NotFound {
            kind: kind_name(selector),
            position: selector.position(),
        })?;
    let owner = Owner {
        drawable: selected.anchor.drawable,
        component_index: selected.drawable.component_index,
        archive: selected.drawable_payload,
        message_index: selected.message_index,
        message_type: graph::message_type(selector),
    };
    label_slot(
        package,
        &index,
        &mut budget,
        Selection::Drawable(selector),
        kind,
        &owner,
    )
}

/// The caption of the body chart `selector` (the rooted chart selection of
/// the chart families; the chart message stores its `TSD.DrawableArchive`
/// at field 1).
pub(super) fn resolve_chart(package: &Package, selector: BodyChartSelector) -> Result<Slot, Error> {
    let mut arrangement_budget = ArrangementBudget::new(package).map_err(map_arrangement_error)?;
    let chart = arrangement::resolve_target(package, selector, &mut arrangement_budget)
        .map_err(map_arrangement_error)?;
    let payload =
        arrangement::selected_chart_payload(package, &chart).map_err(map_arrangement_error)?;
    let mut budget = Budget::new(package)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let located = index.get(chart.drawable_identifier)?;
    if located.component_index != chart.component_index {
        return Err(Error::InvalidSource);
    }
    let view = budget.parse(payload, 2)?;
    let archive = graph::unique_field(&view, 1, 2)?
        .ok_or(Error::InvalidSource)?
        .payload();
    let owner = Owner {
        drawable: chart.drawable_identifier,
        component_index: chart.component_index,
        archive,
        message_index: chart.drawable_message_index,
        message_type: chart.drawable_message_type,
    };
    label_slot(
        package,
        &index,
        &mut budget,
        Selection::Chart(selector),
        Kind::Caption,
        &owner,
    )
}

fn map_arrangement_error(error: arrangement::Error) -> Error {
    match error {
        arrangement::Error::Refused(
            arrangement::Refusal::ChartNotFound { position }
            | arrangement::Refusal::ChartPositionNotFound { position },
        ) => Refusal::NotFound {
            kind: "chart",
            position,
        }
        .into(),
        arrangement::Error::Refused(arrangement::Refusal::UnsupportedSource) => {
            Refusal::UnsupportedSource.into()
        },
        arrangement::Error::Limit {
            kind,
            observed,
            maximum,
        } => {
            use arrangement::LimitKind as Arrangement;
            Error::Limit {
                kind: match kind {
                    Arrangement::OutputBytes | Arrangement::WireOutputBytes => {
                        LimitKind::OutputBytes
                    },
                    Arrangement::WireFields => LimitKind::WireFields,
                    Arrangement::WireNesting => LimitKind::WireNesting,
                    Arrangement::PayloadObjects
                    | Arrangement::PayloadMessages
                    | Arrangement::PayloadItems
                    | Arrangement::PayloadReferences => LimitKind::PayloadItems,
                    Arrangement::WireWork
                    | Arrangement::WireAllocations
                    | Arrangement::WireRetainedBytes
                    | Arrangement::WireScratchBytes => LimitKind::WireWork,
                    _ => LimitKind::InputBytes,
                },
                observed,
                maximum,
            }
        },
        arrangement::Error::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn label_slot(
    package: &Package,
    index: &ObjectIndex<'_>,
    budget: &mut Budget,
    selection: Selection,
    kind: Kind,
    owner: &Owner<'_>,
) -> Result<Slot, Error> {
    let drawable = owner.drawable;
    let component = index.component_name(owner.component_index)?;
    let view = budget.parse(owner.archive, 3)?;
    let reference =
        graph::unique_field(&view, kind.label().field(), 2)?.ok_or(Error::InvalidSource)?;
    let reference = graph::parse_reference(reference.payload(), budget)?;
    let width = caption::archive_width(owner.archive).map_err(|_| Error::InvalidSource)?;
    let located = index.get(reference)?;
    if located.component_index != owner.component_index {
        return Err(Error::InvalidSource);
    }
    let Some((info_index, info)) =
        graph::unique_optional_message(located.object, CAPTION_INFO_MESSAGE_TYPE)?
    else {
        graph::unique_message(located.object, STANDIN_MESSAGE_TYPE)?;
        if located.object.messages.len() != 1 {
            return Err(Error::InvalidSource);
        }
        return Ok(Slot {
            selection,
            kind,
            drawable,
            component,
            message_index: owner.message_index,
            message_type: owner.message_type,
            reference,
            storage: None,
            text: None,
            width,
            objects: vec![reference],
        });
    };
    graph::validate_message_metadata(located.object, info_index)?;
    let payload = info.data.as_slice();
    budget.work(payload.len())?;
    let info = caption::info_graph(payload, reference.get(), drawable.get(), kind.label())
        .ok_or(Error::InvalidSource)?;
    let nonzero = |identifier| NonZeroU64::new(identifier).ok_or(Error::InvalidSource);
    let (style, storage, placement) = (
        nonzero(info.style)?,
        nonzero(info.storage)?,
        nonzero(info.placement)?,
    );
    for (identifier, message_type, private) in [
        (style, SHAPE_STYLE_MESSAGE_TYPE, false),
        (storage, STORAGE_MESSAGE_TYPE, true),
        (placement, CAPTION_PLACEMENT_MESSAGE_TYPE, true),
    ] {
        let object = index.get(identifier)?;
        graph::unique_message(object.object, message_type)?;
        if private && object.component_index != owner.component_index {
            return Err(Error::InvalidSource);
        }
    }
    let text = objects::storage_by_identifier(package, storage.get())?.text;
    Ok(Slot {
        selection,
        kind,
        drawable,
        component,
        message_index: owner.message_index,
        message_type: owner.message_type,
        reference,
        storage: Some(storage),
        text: Some(Arc::from(text)),
        width,
        objects: vec![reference, style, storage, placement],
    })
}

/// Re-resolve `slot` against `package` (the edit's source).
fn reresolve(package: &Package, slot: &Slot) -> Result<Slot, Error> {
    match slot.selection {
        Selection::Drawable(selector) => resolve(package, selector, slot.kind),
        Selection::Chart(selector) => resolve_chart(package, selector),
    }
}

/// The theme's document stylesheet and first caption paragraph-style
/// preset, and the document language (the host's `image_caption_theme`).
pub(crate) fn caption_theme(package: &Package) -> Result<(u64, u64, Option<String>), Error> {
    let mut budget = Budget::new(package)?;
    let (theme, language) = graph::root_theme(package, &mut budget)?;
    let index = ObjectIndex::new(package, &mut budget)?;
    let theme = index.get(theme.ok_or(Error::InvalidSource)?)?;
    let (_, message) = graph::unique_message(theme.object, THEME_MESSAGE_TYPE)?;
    let wrapper = budget.parse(&message.data, 1)?;
    let base_field =
        graph::unique_field(&wrapper, THEME_SUPER_FIELD, 2)?.ok_or(Error::InvalidSource)?;
    let base = budget.parse(base_field.payload(), 2)?;
    let stylesheet = graph::unique_field(&base, THEME_DOCUMENT_STYLESHEET_FIELD, 2)?
        .ok_or(Error::InvalidSource)?;
    let stylesheet = graph::parse_reference(stylesheet.payload(), &mut budget)?;
    let paragraph_style = caption::theme_paragraph_style(base_field.payload())
        .and_then(NonZeroU64::new)
        .ok_or(Error::InvalidSource)?;
    index.get(stylesheet)?;
    index.get(paragraph_style)?;
    Ok((stylesheet.get(), paragraph_style.get(), language))
}

/// The four native objects of a new inline label graph, starting at `first`
/// (the shared `caption::graph_objects`).
fn label_graph(
    source: &Package,
    slot: &Slot,
    first: u64,
    text: &str,
) -> Result<(Vec<ArchiveObject>, [u64; 4]), Error> {
    let width = slot
        .width
        .filter(|width| width.is_finite() && *width > 0.0)
        .ok_or(Refusal::NoWidth)?;
    let (stylesheet, paragraph_style, language) = caption_theme(source)?;
    let ids = caption::GraphIds::allocate(first).ok_or(Error::InvalidSource)?;
    let (objects, _) = caption::graph_objects(
        ids,
        slot.drawable.get(),
        width,
        text,
        caption::Theme {
            stylesheet,
            paragraph_style,
            language: language.as_deref(),
        },
        slot.kind.label(),
        caption::Layout::Inline,
        graph_codec::EncodeOptions::for_text(text),
        archive_limits(source)?,
    )
    .map_err(|error| match error {
        caption::BuildError::Encode(graph_codec::EncodeError::Resource(_)) => Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: text.len() as u64,
            maximum: text.len() as u64,
        },
        caption::BuildError::Encode(graph_codec::EncodeError::Allocation { amount }) => {
            Error::Allocation { amount }
        },
        caption::BuildError::Archive(error) => Error::from(graph::map_core_error(error)),
        caption::BuildError::Encode(_) => Error::InvalidSource,
    })?;
    Ok((objects, ids.all()))
}

fn archive_limits(source: &Package) -> Result<litchi_iwa_core::Limits, Error> {
    source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(|error| Error::from(graph::map_archive_error(error)))
}

fn unchanged(edit: Edit<'_>) -> Commit {
    let source_bytes = edit.source.state.source.shared_source();
    Commit {
        package: edit.source.clone(),
        patch: Patch {
            artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
            selector: edit.selector,
            kind: edit.slot.kind,
            before: edit.slot.text.clone(),
            after: edit.slot.text,
        },
        diagnostics: Diagnostics {
            changed: false,
            created_objects: 0,
        },
    }
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let source = edit.source;
    let Some(operation) = edit.operation.clone() else {
        return Ok(unchanged(edit));
    };
    let after: Option<Arc<str>> = operation.map(Arc::from);
    if after == edit.slot.text {
        return Ok(unchanged(edit));
    }
    let Some((candidate, created)) = publish(source, &edit.slot, after.as_ref())? else {
        return Ok(unchanged(edit));
    };
    let target_bytes = candidate.state.source.shared_source();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: ExactArtifacts::new(source.state.source.shared_source(), target_bytes),
            selector: edit.selector,
            kind: edit.slot.kind,
            before: edit.slot.text,
            after,
        },
        diagnostics: Diagnostics {
            changed: true,
            created_objects: created,
        },
    })
}

/// Publish `after` as the label of `staged` (resolved from `source`) and
/// strictly reopen: the candidate and the number of created objects, or
/// `None` when nothing changes. Shared with the chart caption family.
pub(super) fn publish(
    source: &Package,
    staged: &Slot,
    after: Option<&Arc<str>>,
) -> Result<Option<(Package, usize)>, Error> {
    if !source.state.source.source_is_exact() {
        return Err(Refusal::UnsupportedSource.into());
    }
    let slot = reresolve(source, staged)?;
    if slot != *staged {
        return Err(Error::PatchConflict);
    }
    let other = match slot.selection {
        Selection::Drawable(selector) => {
            let other = match slot.kind {
                Kind::Title => Kind::Caption,
                Kind::Caption => Kind::Title,
            };
            Some((other, resolve(source, selector, other)?.text))
        },
        // A chart's title is the chart-title family's, not a label graph.
        Selection::Chart(_) => None,
    };
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(graph::map_archive_error)?;
    let mut archive = graph::editable_archive(source, &slot.component)?;
    let mut created = Vec::new();
    match (after, slot.storage) {
        (Some(text), Some(storage)) => {
            // A present label: rewrite its storage text.
            let mut located = objects::storage_by_identifier(source, storage.get())?;
            let end = located.utf16_len();
            let removed = objects::splice_text(source, &mut located, 0, end, text)?;
            if !removed.is_empty() {
                return Err(Refusal::DependentContent.into());
            }
            archive
                .object_mut(storage.get())
                .ok_or(Error::InvalidSource)?
                .replace_message_preserving_header(
                    located.message_index,
                    RawMessage {
                        type_: located.message_type,
                        data: located.payload,
                    },
                    archive_limits,
                )
                .map_err(graph::map_core_error)?;
        },
        (Some(text), None) => {
            // An absent label: author the inline graph.
            let metadata = objects::metadata(source)?;
            let first = objects::next_identifier(source, metadata.as_ref())?;
            let (graph_objects, ids) = label_graph(source, &slot, first, text)?;
            for object in graph_objects {
                archive
                    .insert_object(object, archive_limits)
                    .map_err(graph::map_core_error)?;
            }
            repoint_drawable(&mut archive, &slot, ids[1], archive_limits)?;
            created.extend(ids);
        },
        (None, Some(_)) => {
            // A present label: point the drawable at a fresh stand-in.
            let metadata = objects::metadata(source)?;
            let standin = objects::next_identifier(source, metadata.as_ref())?;
            let object = caption::standin(standin, archive_limits)
                .map_err(|error| Error::from(graph::map_core_error(error)))?;
            archive
                .insert_object(object, archive_limits)
                .map_err(graph::map_core_error)?;
            repoint_drawable(&mut archive, &slot, standin, archive_limits)?;
            created.push(standin);
        },
        (None, None) => return Ok(None),
    }
    let mut archives: Vec<(Arc<str>, Archive)> = vec![(Arc::clone(&slot.component), archive)];
    if let (Some(last), Some(metadata)) = (created.last(), objects::metadata(source)?) {
        // Body drawables register in the document component; a chart in
        // the component that stores it (the host's `source.component_id`).
        let component = match slot.selection {
            Selection::Drawable(_) => DOCUMENT_COMPONENT_IDENTIFIER,
            Selection::Chart(_) => metadata
                .component_for_entry(&slot.component)?
                .ok_or(Error::InvalidSource)?,
        };
        let payload = metadata.registered(source, component, &created, *last)?;
        let mut metadata_archive = graph::editable_archive(source, &metadata.component)?;
        metadata_archive
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
        archives.push((Arc::clone(&metadata.component), metadata_archive));
    }
    let candidate = graph::publish_archives(source, archives)?;
    if reresolve(&candidate, &slot)?.text.as_ref() != after {
        return Err(Error::Verification);
    }
    if let (Selection::Drawable(selector), Some((other, before))) = (slot.selection, other)
        && resolve(&candidate, selector, other)?.text != before
    {
        return Err(Error::Verification);
    }
    Ok(Some((candidate, created.len())))
}

fn repoint_drawable(
    archive: &mut Archive,
    slot: &Slot,
    replacement: u64,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    let object = archive
        .object_mut(slot.drawable.get())
        .ok_or(Error::InvalidSource)?;
    let payload = object
        .messages
        .get(slot.message_index)
        .ok_or(Error::InvalidSource)?
        .data
        .clone();
    let nesting = slot.selection.nesting();
    let data = caption::repoint(&payload, nesting, slot.kind.label(), replacement)
        .map_err(|error| Error::from(graph::map_wire_error(error)))?;
    object
        .replace_message_preserving_header(
            slot.message_index,
            RawMessage {
                type_: slot.message_type,
                data,
            },
            archive_limits,
        )
        .map_err(graph::map_core_error)?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(slot.message_index)
        .ok_or(Error::InvalidSource)?;
    caption::retarget_references(info, nesting, slot.reference.get(), replacement)
        .map_err(|_| Error::InvalidSource)
}
