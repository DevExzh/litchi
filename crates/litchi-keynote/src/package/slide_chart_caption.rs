//! Exact-source, selector-first Keynote chart-caption text transactions.
//!
//! This owner deliberately admits only replacement of an existing, exclusive
//! native caption text storage. Creating a caption graph or replacing one with
//! a stand-in changes object identity, UUID registration, and package metadata;
//! those compatibility operations remain outside this focused transaction.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction redacts native graph failures at the semantic boundary."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{SourceCatalog, package::ExactArtifacts};
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len, wire::WireView};
use litchi_iwa_core::ArchiveObject;
use litchi_iwa_protos::{keynote_chart_caption_codec, pages_movie_caption_codec};
use thiserror::Error;

use super::{
    Package, PhysicalSource, ReadError, STORAGE_MESSAGE_TYPE, SemanticLimitKind,
    slide_chart_title::{ChartSelection, select_chart},
    unique_payload,
};
use crate::{ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CAPTION_INFO_MESSAGE_TYPE: u32 = 633;
const CAPTION_PLACEMENT_MESSAGE_TYPE: u32 = 634;
const SHAPE_STYLE_MESSAGE_TYPE: u32 = 2_025;
const PREVIEW_ENTRY_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const MAX_CAPTION_BYTES: usize = 64 * 1024 * 1024;

/// A finite resource governed while a chart-caption transaction is prepared
/// or published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartCaptionLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic text-storage objects.
    TextStorages,
    /// Semantic text fragments.
    TextFragments,
    /// Aggregate semantic text bytes.
    TextBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf scan and rewrite work.
    WireWork,
    /// UTF-8 bytes in one caption value.
    CaptionBytes,
}

impl fmt::Display for ChartCaptionLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::CaptionBytes => "caption bytes",
        })
    }
}

/// A content-redacted failure raised by a chart-caption transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartCaptionError {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical chart-caption edits")]
    UnsupportedSource,
    /// The requested operation needs caption graph creation or stand-in
    /// replacement, which this focused text owner intentionally does not do.
    #[error("the requested Keynote chart-caption graph operation is not yet supported")]
    UnsupportedDependency,
    /// An exact-name slide or chart selector was ambiguous.
    #[error("the Keynote chart-caption selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// Missing checked semantic slide position.
        position: Position,
    },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked semantic chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound {
        /// Missing checked semantic chart position.
        position: Position,
    },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// The source chart, caption graph, or text storage is malformed or unsafe.
    #[error("the Keynote chart-caption source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote chart-caption {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: ChartCaptionLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote chart-caption transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full candidate reopening did not reproduce the requested caption.
    #[error("the edited Keynote chart caption failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote chart-caption patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable chart-caption value staged against an immutable package.
pub struct ChartCaptionEdit<'a> {
    source: &'a Package,
    selection: CaptionSelection,
    after: Option<String>,
}

impl fmt::Debug for ChartCaptionEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartCaptionEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("chart_position", &self.selection.chart_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> ChartCaptionEdit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Self, ChartCaptionError> {
        let selection = select_caption(source, slide_selector.into(), chart_selector.into(), true)?;
        let after = selection.text.as_deref().map(copy_caption).transpose()?;
        Ok(Self {
            source,
            selection,
            after,
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.selection.chart_position
    }

    /// Borrow the caption text observed when this edit began.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    /// Borrow the caption text staged for publication.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Stage whole-text replacement of an existing native caption.
    ///
    /// If the selected chart currently has no native caption storage, commit
    /// returns [`ChartCaptionError::UnsupportedDependency`] rather than
    /// synthesizing an incomplete graph.
    pub fn set(mut self, caption: impl AsRef<str>) -> Result<Self, ChartCaptionError> {
        self.after = Some(copy_caption(caption.as_ref())?);
        Ok(self)
    }

    /// Stage caption removal.
    ///
    /// A changed removal needs native stand-in allocation and is therefore
    /// refused by this replacement-only owner. Clearing an already absent
    /// caption remains an exact no-op.
    pub fn clear(mut self) -> Result<Self, ChartCaptionError> {
        self.after = None;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<ChartCaptionCommit, ChartCaptionError> {
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let current = select_caption(
            self.source,
            SlideSelector::position(self.selection.slide_position),
            ChartSelector::index(self.selection.chart_position.get()),
            true,
        )?;
        if !current.same_identity(&self.selection) || current.text != self.selection.text {
            return Err(ChartCaptionError::InvalidSource);
        }
        if self.selection.text == self.after {
            self.source.validate().map_err(map_read_error)?;
            return Ok(ChartCaptionCommit {
                package: self.source.snapshot(),
                patch: ChartCaptionPatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    selection: self.selection,
                    after: self.after,
                    touched_components: 0,
                    deleted_previews: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: ChartCaptionDiagnostics::unchanged(),
            });
        }
        let (Some(before), Some(after), Some(storage_identifier)) = (
            self.selection.text.as_deref(),
            self.after.as_deref(),
            self.selection.storage_identifier,
        ) else {
            return Err(ChartCaptionError::UnsupportedDependency);
        };
        if contains_dependent_marker(before) || contains_dependent_marker(after) {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        if !catalog.source_is_exact() {
            return Err(ChartCaptionError::UnsupportedSource);
        }
        self.source.validate().map_err(map_read_error)?;
        let end = before.encode_utf16().count();
        let deleted_previews = preview_count(catalog);
        let (package, touched_components) = super::slide_text::rewrite_owned_storage_text(
            self.source,
            storage_identifier,
            self.selection.slide_node_identifier,
            0..end,
            after,
        )
        .map_err(map_slide_text_error)?;
        let candidate = select_caption(
            &package,
            SlideSelector::position(self.selection.slide_position),
            ChartSelector::index(self.selection.chart_position.get()),
            true,
        )?;
        if !candidate.same_identity(&self.selection) || candidate.text != self.after {
            return Err(ChartCaptionError::Verification);
        }
        verify_caption_candidate(
            self.source,
            &package,
            &self.selection,
            self.after.as_deref(),
            true,
        )?;
        let target = physical_catalog(&package)?.shared_source();
        Ok(ChartCaptionCommit {
            package,
            patch: ChartCaptionPatch {
                artifacts: ExactArtifacts::new(source_bytes, target),
                selection: self.selection,
                after: self.after,
                touched_components,
                deleted_previews,
                target_requires_invalidated_previews: true,
            },
            diagnostics: ChartCaptionDiagnostics::published(touched_components, deleted_previews),
        })
    }
}

/// An exact-source-checked reversible semantic chart-caption patch.
#[derive(Clone, PartialEq, Eq)]
pub struct ChartCaptionPatch {
    artifacts: ExactArtifacts,
    selection: CaptionSelection,
    after: Option<String>,
    touched_components: usize,
    deleted_previews: usize,
    target_requires_invalidated_previews: bool,
}

impl fmt::Debug for ChartCaptionPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartCaptionPatch")
            .field("slide_position", &self.selection.slide_position)
            .field("chart_position", &self.selection.chart_position)
            .field("has_before", &self.selection.text.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl ChartCaptionPatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.selection.chart_position
    }

    /// Borrow the caption required from the source package.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.selection.text.as_deref()
    }

    /// Borrow the caption produced by the target package.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch preserves semantic caption state and bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.selection.text == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from target back to source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        let mut selection = self.selection.clone();
        selection.text = self.after.clone();
        Self {
            artifacts: self.artifacts.inverse(),
            selection,
            after: self.selection.text.clone(),
            touched_components: self.touched_components,
            deleted_previews: 0,
            target_requires_invalidated_previews: self.touched_components != 0
                && !self.target_requires_invalidated_previews,
        }
    }
}

/// Compact evidence describing one chart-caption commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartCaptionDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl ChartCaptionDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return how many stale root rendering previews were deleted.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable chart-caption transaction.
#[must_use = "a Keynote chart-caption commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct ChartCaptionCommit {
    package: Package,
    patch: ChartCaptionPatch,
    diagnostics: ChartCaptionDiagnostics,
}

impl ChartCaptionCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &ChartCaptionPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &ChartCaptionDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the caption text of one selected chart.
    ///
    /// `None` means that the chart has no native caption storage; `Some("")`
    /// is an existing empty storage.
    pub fn slide_chart_caption<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Option<String>, ChartCaptionError> {
        Ok(select_caption(self, slide_selector.into(), chart_selector.into(), true)?.text)
    }

    /// Start an exact immutable edit of one selected chart caption.
    pub fn edit_slide_chart_caption<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartCaptionEdit<'_>, ChartCaptionError> {
        ChartCaptionEdit::new(self, slide_selector, chart_selector)
    }

    /// Apply an exact-source-checked chart-caption patch.
    pub fn apply_slide_chart_caption(
        &self,
        patch: &ChartCaptionPatch,
    ) -> Result<ChartCaptionCommit, ChartCaptionError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(ChartCaptionError::PatchConflict);
        }
        let current = select_caption(
            self,
            SlideSelector::position(patch.selection.slide_position),
            ChartSelector::index(patch.selection.chart_position.get()),
            true,
        )?;
        if !current.same_identity(&patch.selection) || current.text != patch.selection.text {
            return Err(ChartCaptionError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(ChartCaptionCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: ChartCaptionDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartCaptionError::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let selected = select_caption(
            &candidate,
            SlideSelector::position(patch.selection.slide_position),
            ChartSelector::index(patch.selection.chart_position.get()),
            true,
        )?;
        if !selected.same_identity(&patch.selection) || selected.text != patch.after {
            return Err(ChartCaptionError::Verification);
        }
        verify_caption_candidate(
            self,
            &candidate,
            &patch.selection,
            patch.after.as_deref(),
            patch.target_requires_invalidated_previews,
        )?;
        Ok(ChartCaptionCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ChartCaptionDiagnostics::published(
                patch.touched_components,
                patch.deleted_previews,
            ),
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct CaptionSelection {
    slide_position: Position,
    chart_position: Position,
    slide_identifier: u64,
    slide_node_identifier: u64,
    chart_identifier: u64,
    non_style_identifier: u64,
    slide_component_name: String,
    reference_identifier: Option<u64>,
    caption_info_identifier: Option<u64>,
    storage_identifier: Option<u64>,
    placement_identifier: Option<u64>,
    style_identifier: Option<u64>,
    text: Option<String>,
}

impl CaptionSelection {
    fn same_identity(&self, other: &Self) -> bool {
        self.slide_position == other.slide_position
            && self.chart_position == other.chart_position
            && self.slide_identifier == other.slide_identifier
            && self.slide_node_identifier == other.slide_node_identifier
            && self.chart_identifier == other.chart_identifier
            && self.non_style_identifier == other.non_style_identifier
            && self.slide_component_name == other.slide_component_name
            && self.reference_identifier == other.reference_identifier
            && self.caption_info_identifier == other.caption_info_identifier
            && self.storage_identifier == other.storage_identifier
            && self.placement_identifier == other.placement_identifier
            && self.style_identifier == other.style_identifier
    }
}

fn select_caption(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    chart_selector: ChartSelector<'_>,
    mutation_guards: bool,
) -> Result<CaptionSelection, ChartCaptionError> {
    let chart = select_chart(package, slide_selector, chart_selector, mutation_guards)
        .map_err(map_chart_title_error)?;
    caption_selection(package, chart, mutation_guards)
}

fn caption_selection(
    package: &Package,
    chart: ChartSelection,
    mutation_guards: bool,
) -> Result<CaptionSelection, ChartCaptionError> {
    let record = package
        .slide_record_at(chart.slide_position.get())
        .map_err(map_read_error)?
        .ok_or(ChartCaptionError::InvalidSource)?;
    if record.slide_identifier != chart.slide_identifier {
        return Err(ChartCaptionError::InvalidSource);
    }
    let (component, object) = package
        .object_with_component(chart.chart_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if component != chart.slide_component_name || object.messages.len() != 1 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let payload = unique_payload(&object.messages, &[CHART_MESSAGE_TYPE], "Keynote chart")
        .map_err(map_read_error)?;
    let reference_identifier = chart_caption_reference(package, payload)?;
    let empty = |reference_identifier| CaptionSelection {
        slide_position: chart.slide_position,
        chart_position: chart.chart_position,
        slide_identifier: chart.slide_identifier,
        slide_node_identifier: record.node_identifier,
        chart_identifier: chart.chart_identifier,
        non_style_identifier: chart.non_style_identifier,
        slide_component_name: chart.slide_component_name.clone(),
        reference_identifier,
        caption_info_identifier: None,
        storage_identifier: None,
        placement_identifier: None,
        style_identifier: None,
        text: None,
    };
    let Some(reference_identifier) = reference_identifier else {
        return Ok(empty(None));
    };
    let (info_component, info_object) = package
        .object_with_component(reference_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if info_component != chart.slide_component_name || info_object.messages.len() != 1 {
        return Err(ChartCaptionError::InvalidSource);
    }
    if info_object.messages[0].type_ == STANDIN_MESSAGE_TYPE {
        validate_selected_message_metadata(info_object, 0)?;
        return Ok(empty(Some(reference_identifier)));
    }
    if info_object.messages[0].type_ != CAPTION_INFO_MESSAGE_TYPE {
        return Err(ChartCaptionError::InvalidSource);
    }
    validate_selected_message_metadata(info_object, 0)?;
    let snapshot = pages_movie_caption_codec::decode_caption_info(
        &info_object.messages[0].data,
        caption_info_decode_options(package, &info_object.messages[0].data)?,
    )
    .map_err(map_caption_info_codec_error)?;
    let storage_identifier = snapshot
        .owned_storage_identifier()
        .ok_or(ChartCaptionError::InvalidSource)?;
    if storage_identifier == 0
        || snapshot.deprecated_storage_identifier() != Some(storage_identifier)
        || snapshot.parent_identifier() != chart.chart_identifier
        || snapshot.is_text_box() != Some(true)
        || snapshot.child_info_kind() != Some(1)
    {
        return Err(ChartCaptionError::InvalidSource);
    }
    let placement_identifier = snapshot
        .placement_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(ChartCaptionError::InvalidSource)?;
    let style_identifier = snapshot
        .style_identifier()
        .filter(|identifier| *identifier != 0)
        .ok_or(ChartCaptionError::InvalidSource)?;
    let identifiers = [
        chart.chart_identifier,
        reference_identifier,
        storage_identifier,
        placement_identifier,
        style_identifier,
    ];
    for (index, identifier) in identifiers.iter().enumerate() {
        if identifiers[..index].contains(identifier) {
            return Err(ChartCaptionError::InvalidSource);
        }
    }
    require_private_object(
        package,
        storage_identifier,
        STORAGE_MESSAGE_TYPE,
        Some(chart.slide_component_name.as_str()),
    )?;
    require_private_object(
        package,
        placement_identifier,
        CAPTION_PLACEMENT_MESSAGE_TYPE,
        Some(chart.slide_component_name.as_str()),
    )?;
    require_private_object(package, style_identifier, SHAPE_STYLE_MESSAGE_TYPE, None)?;
    if mutation_guards {
        prove_exclusive_caption_storage(package, reference_identifier, storage_identifier)?;
    }
    let text = super::slide_text::read_owned_storage_text(package, storage_identifier)
        .map_err(map_slide_text_error)?;
    Ok(CaptionSelection {
        slide_position: chart.slide_position,
        chart_position: chart.chart_position,
        slide_identifier: chart.slide_identifier,
        slide_node_identifier: record.node_identifier,
        chart_identifier: chart.chart_identifier,
        non_style_identifier: chart.non_style_identifier,
        slide_component_name: chart.slide_component_name,
        reference_identifier: Some(reference_identifier),
        caption_info_identifier: Some(reference_identifier),
        storage_identifier: Some(storage_identifier),
        placement_identifier: Some(placement_identifier),
        style_identifier: Some(style_identifier),
        text: Some(text),
    })
}

fn require_private_object(
    package: &Package,
    identifier: u64,
    message_type: u32,
    expected_component: Option<&str>,
) -> Result<(), ChartCaptionError> {
    let (component, object) = package
        .object_with_component(identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if expected_component.is_some_and(|expected| expected != component)
        || object.messages.len() != 1
        || object.messages[0].type_ != message_type
    {
        return Err(ChartCaptionError::InvalidSource);
    }
    validate_selected_message_metadata(object, 0)
}

fn validate_selected_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), ChartCaptionError> {
    let message = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || message.base_message_index.is_some()
        || !message.diff_merge_version.is_empty()
        || message.diff_field_path.is_some()
        || !message.fields_to_remove.is_empty()
        || !message.diff_read_version.is_empty()
    {
        return Err(ChartCaptionError::InvalidSource);
    }
    Ok(())
}

fn prove_exclusive_caption_storage(
    package: &Package,
    caption_info_identifier: u64,
    storage_identifier: u64,
) -> Result<(), ChartCaptionError> {
    let mut payload_owner_seen = false;
    let mut metadata_owner_seen = false;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(ChartCaptionError::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == CAPTION_INFO_MESSAGE_TYPE {
                    let snapshot = pages_movie_caption_codec::decode_caption_info(
                        &message.data,
                        caption_info_decode_options(package, &message.data)?,
                    )
                    .map_err(map_caption_info_codec_error)?;
                    let owns = snapshot.owned_storage_identifier() == Some(storage_identifier)
                        || snapshot.deprecated_storage_identifier() == Some(storage_identifier);
                    if owns
                        && (std::mem::replace(&mut payload_owner_seen, true)
                            || owner_identifier != caption_info_identifier)
                    {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                }
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(ChartCaptionError::InvalidSource)?;
                let aggregate_count = info
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == storage_identifier)
                    .count();
                let aggregate_data = info
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == storage_identifier)
                    .count();
                let mut field_count = 0usize;
                let mut field_data = 0usize;
                for field in &info.field_infos {
                    let occurrences = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == storage_identifier)
                        .count();
                    if occurrences != 0 && !matches!(field.path.as_slice(), [1, 2] | [1, 4]) {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                    field_count = field_count
                        .checked_add(occurrences)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                    field_data = field_data
                        .checked_add(
                            field
                                .data_references
                                .iter()
                                .filter(|identifier| **identifier == storage_identifier)
                                .count(),
                        )
                        .ok_or(ChartCaptionError::InvalidSource)?;
                }
                if aggregate_count != 0
                    || field_count != 0
                    || aggregate_data != 0
                    || field_data != 0
                {
                    if owner_identifier != caption_info_identifier
                        || message.type_ != CAPTION_INFO_MESSAGE_TYPE
                        || std::mem::replace(&mut metadata_owner_seen, true)
                        || aggregate_count != 1
                        || field_count > 2
                        || aggregate_data != 0
                        || field_data != 0
                    {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                }
            }
        }
    }
    if payload_owner_seen && metadata_owner_seen {
        Ok(())
    } else {
        Err(ChartCaptionError::InvalidSource)
    }
}

fn chart_caption_reference(
    package: &Package,
    payload: &[u8],
) -> Result<Option<u64>, ChartCaptionError> {
    let options = chart_caption_decode_options(package, payload)?;
    let identifier = keynote_chart_caption_codec::decode_chart_caption_identifier(payload, options)
        .map_err(map_chart_caption_codec_error)?;
    let Some(identifier) = identifier else {
        return Ok(None);
    };
    if identifier == 0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let outer = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let super_payload = unique_payload_field(&outer, 1)?.ok_or(ChartCaptionError::InvalidSource)?;
    let drawable = WireView::parse_with_limits(super_payload, limits).map_err(map_wire_error)?;
    let reference = unique_payload_field(&drawable, 11)?.ok_or(ChartCaptionError::InvalidSource)?;
    let projected = super::validate_reference_payload(reference, limits, "Keynote chart caption")
        .map_err(map_wire_error)?;
    if projected != identifier {
        return Err(ChartCaptionError::InvalidSource);
    }
    let view = WireView::parse_with_limits(reference, limits).map_err(map_wire_error)?;
    for field in view.fields().filter(|field| field.number() == 3) {
        let (value, bytes) = decode_varint_from_bytes(field.payload())
            .map_err(|_error| ChartCaptionError::InvalidSource)?;
        if bytes != field.payload().len() || bytes != encoded_len(value) || value != 0 {
            return Err(ChartCaptionError::InvalidSource);
        }
    }
    Ok(Some(identifier))
}

fn unique_payload_field<'a>(
    view: &WireView<'a>,
    number: u32,
) -> Result<Option<&'a [u8]>, ChartCaptionError> {
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(ChartCaptionError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        selected = Some(field.payload());
    }
    Ok(selected)
}

fn verify_caption_candidate(
    source: &Package,
    candidate: &Package,
    selection: &CaptionSelection,
    expected: Option<&str>,
    require_invalidated_previews: bool,
) -> Result<(), ChartCaptionError> {
    if source.state.total_objects != candidate.state.total_objects {
        return Err(ChartCaptionError::Verification);
    }
    let selected = select_caption(
        candidate,
        SlideSelector::position(selection.slide_position),
        ChartSelector::index(selection.chart_position.get()),
        true,
    )?;
    if !selected.same_identity(selection) || selected.text.as_deref() != expected {
        return Err(ChartCaptionError::Verification);
    }
    let storage_identifier = selection
        .storage_identifier
        .ok_or(ChartCaptionError::Verification)?;
    super::slide_text::verify_owned_storage_candidate(
        source,
        candidate,
        storage_identifier,
        selection.slide_node_identifier,
        require_invalidated_previews,
    )
    .map_err(map_slide_text_error)
}

fn chart_caption_decode_options(
    package: &Package,
    payload: &[u8],
) -> Result<keynote_chart_caption_codec::DecodeOptions, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_error| ChartCaptionError::InvalidSource)?;
    Ok(keynote_chart_caption_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn caption_info_decode_options(
    package: &Package,
    payload: &[u8],
) -> Result<pages_movie_caption_codec::DecodeOptions, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_error| ChartCaptionError::InvalidSource)?;
    Ok(pages_movie_caption_codec::DecodeOptions::new(
        payload.len().min(limits.max_input_bytes()),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion,
    ))
}

fn copy_caption(value: &str) -> Result<String, ChartCaptionError> {
    if value.len() > MAX_CAPTION_BYTES {
        return Err(ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::CaptionBytes,
            observed: usize_to_u64(value.len()),
            maximum: usize_to_u64(MAX_CAPTION_BYTES),
        });
    }
    let mut copy = String::new();
    copy.try_reserve_exact(value.len())
        .map_err(|_error| ChartCaptionError::Allocation {
            amount: value.len(),
        })?;
    copy.push_str(value);
    Ok(copy)
}

fn contains_dependent_marker(text: &str) -> bool {
    text.contains('\u{000e}') || text.contains('\u{fffc}')
}

fn preview_count(catalog: &SourceCatalog) -> usize {
    PREVIEW_ENTRY_NAMES
        .iter()
        .filter(|name| catalog.package().iter().any(|entry| entry.name() == **name))
        .count()
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, ChartCaptionError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(ChartCaptionError::UnsupportedSource),
    }
}

fn map_chart_title_error(error: super::slide_chart_title::ChartTitleError) -> ChartCaptionError {
    match error {
        super::slide_chart_title::ChartTitleError::UnsupportedSource => {
            ChartCaptionError::UnsupportedSource
        },
        super::slide_chart_title::ChartTitleError::AmbiguousSelector => {
            ChartCaptionError::AmbiguousSelector
        },
        super::slide_chart_title::ChartTitleError::EmptySlideName => {
            ChartCaptionError::EmptySlideName
        },
        super::slide_chart_title::ChartTitleError::SlideNameNotFound => {
            ChartCaptionError::SlideNameNotFound
        },
        super::slide_chart_title::ChartTitleError::SlidePositionNotFound { position } => {
            ChartCaptionError::SlidePositionNotFound { position }
        },
        super::slide_chart_title::ChartTitleError::ChartNameNotFound => {
            ChartCaptionError::ChartNameNotFound
        },
        super::slide_chart_title::ChartTitleError::ChartPositionNotFound { position } => {
            ChartCaptionError::ChartPositionNotFound { position }
        },
        super::slide_chart_title::ChartTitleError::EmptyChartName => {
            ChartCaptionError::EmptyChartName
        },
        super::slide_chart_title::ChartTitleError::LimitExceeded {
            observed, maximum, ..
        } => ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            observed,
            maximum,
        },
        super::slide_chart_title::ChartTitleError::Allocation { amount } => {
            ChartCaptionError::Allocation { amount }
        },
        super::slide_chart_title::ChartTitleError::InvalidSource
        | super::slide_chart_title::ChartTitleError::Verification
        | super::slide_chart_title::ChartTitleError::PatchConflict => {
            ChartCaptionError::InvalidSource
        },
    }
}

fn map_read_error(error: ReadError) -> ChartCaptionError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartCaptionError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => ChartCaptionLimitKind::Entries,
                SemanticLimitKind::Slides => ChartCaptionLimitKind::Slides,
                SemanticLimitKind::References => ChartCaptionLimitKind::References,
                SemanticLimitKind::TextStorages => ChartCaptionLimitKind::TextStorages,
                SemanticLimitKind::TextFragments => ChartCaptionLimitKind::TextFragments,
                SemanticLimitKind::TextBytes => ChartCaptionLimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartCaptionError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => ChartCaptionLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => ChartCaptionLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => ChartCaptionLimitKind::WireNesting,
                super::PayloadLimitKind::Work => ChartCaptionLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => ChartCaptionError::Allocation { amount },
        _ => ChartCaptionError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> ChartCaptionError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ChartCaptionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => ChartCaptionLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => ChartCaptionLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields => ChartCaptionLimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => ChartCaptionLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => ChartCaptionLimitKind::WireWork,
                _ => ChartCaptionLimitKind::WireBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        _ => ChartCaptionError::InvalidSource,
    }
}

fn map_chart_caption_codec_error(
    error: keynote_chart_caption_codec::DecodeError,
) -> ChartCaptionError {
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            keynote_chart_caption_codec::WireResourceLimit::Bytes { observed, maximum } => {
                ChartCaptionError::LimitExceeded {
                    kind: ChartCaptionLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            keynote_chart_caption_codec::WireResourceLimit::Nesting { observed, maximum } => {
                ChartCaptionError::LimitExceeded {
                    kind: ChartCaptionLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => ChartCaptionError::InvalidSource,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    ChartCaptionError::InvalidSource
}

fn map_caption_info_codec_error(
    error: pages_movie_caption_codec::DecodeError,
) -> ChartCaptionError {
    if let Some((observed, maximum)) = error.message_byte_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    ChartCaptionError::InvalidSource
}

fn map_slide_text_error(error: super::slide_text::SlideTextError) -> ChartCaptionError {
    match error {
        super::slide_text::SlideTextError::UnsupportedSource => {
            ChartCaptionError::UnsupportedSource
        },
        super::slide_text::SlideTextError::DependentContent
        | super::slide_text::SlideTextError::ObjectMarkerReplacement => {
            ChartCaptionError::UnsupportedDependency
        },
        super::slide_text::SlideTextError::LimitExceeded {
            observed, maximum, ..
        } => ChartCaptionError::LimitExceeded {
            kind: ChartCaptionLimitKind::WireWork,
            observed,
            maximum,
        },
        super::slide_text::SlideTextError::Allocation { amount } => {
            ChartCaptionError::Allocation { amount }
        },
        super::slide_text::SlideTextError::Verification => ChartCaptionError::Verification,
        _ => ChartCaptionError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
