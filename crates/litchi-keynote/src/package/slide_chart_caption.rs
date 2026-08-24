//! Exact-source, selector-first Keynote chart-caption text transactions.
//!
//! This owner admits the canonical inline chart-caption graph transition:
//! creation from an exclusive stand-in and removal by retaining the old graph
//! while retargeting the chart to a fresh stand-in.  Existing storage text is
//! still edited through the focused storage owner.  Cross-component graphs,
//! shared ownership, and metadata shapes that cannot be proven exactly remain
//! outside this transaction.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction redacts native graph failures at the semantic boundary."
)]

use std::collections::HashSet;
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{decode_varint_from_bytes, varint::encoded_len, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::package_metadata_codec::{
    self as metadata_codec, AdditionSaveTokenBatch, Batch as MetadataBatch, ComponentDescriptor,
    ComponentSelector, ObjectUuidAddition, PackageMetadataVisitor,
    RewriteOptions as MetadataRewriteOptions, SaveTokenBatch, UuidBits,
    inspect_package_metadata_with_visitor, rewrite_package_metadata_additions_and_save_tokens,
};
use litchi_iwa_protos::{
    keynote_chart_caption_codec, keynote_chart_caption_graph_codec as graph_codec,
    pages_movie_caption_codec,
};
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
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const PACKAGE_METADATA_MEMBER_NAME: &str = "Index/Metadata.iwa";
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
    /// The requested operation needs a graph shape outside the supported
    /// canonical inline stand-in transition.
    #[error("the requested Keynote chart-caption graph operation is unsupported")]
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

    /// Stage whole-text replacement or creation of a native caption.
    pub fn set(mut self, caption: impl AsRef<str>) -> Result<Self, ChartCaptionError> {
        self.after = Some(copy_caption(caption.as_ref())?);
        Ok(self)
    }

    /// Stage caption removal.
    ///
    /// Stage removal of an active caption.  Commit replaces the edge with a
    /// fresh canonical stand-in and retains the old graph for exact history
    /// and inverse semantics. Clearing an already absent caption is an exact
    /// no-op.
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
                    target_selection: self.selection.clone(),
                    selection: self.selection,
                    after: self.after,
                    touched_components: 0,
                    deleted_previews: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: ChartCaptionDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartCaptionError::UnsupportedSource);
        }
        self.source.validate().map_err(map_read_error)?;
        let (package, touched_components, deleted_previews) =
            rewrite_chart_caption_operation(self.source, &self.selection, self.after.as_deref())?;
        let candidate = select_caption(
            &package,
            SlideSelector::position(self.selection.slide_position),
            ChartSelector::index(self.selection.chart_position.get()),
            true,
        )?;
        if !candidate.same_chart_identity(&self.selection) || candidate.text != self.after {
            return Err(ChartCaptionError::Verification);
        }
        verify_caption_candidate(
            self.source,
            &package,
            &self.selection,
            &candidate,
            self.after.as_deref(),
            self.selection.text != self.after,
        )?;
        let target = physical_catalog(&package)?.shared_source();
        Ok(ChartCaptionCommit {
            package,
            patch: ChartCaptionPatch {
                artifacts: ExactArtifacts::new(source_bytes, target),
                selection: self.selection,
                target_selection: candidate,
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
    target_selection: CaptionSelection,
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
        self.selection == self.target_selection && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from target back to source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.target_selection.clone(),
            target_selection: self.selection.clone(),
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
        if !selected.same_identity(&patch.target_selection) || selected.text != patch.after {
            return Err(ChartCaptionError::Verification);
        }
        verify_caption_candidate(
            self,
            &candidate,
            &patch.selection,
            &patch.target_selection,
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
    fn same_chart_identity(&self, other: &Self) -> bool {
        self.slide_position == other.slide_position
            && self.chart_position == other.chart_position
            && self.slide_identifier == other.slide_identifier
            && self.slide_node_identifier == other.slide_node_identifier
            && self.chart_identifier == other.chart_identifier
            && self.non_style_identifier == other.non_style_identifier
            && self.slide_component_name == other.slide_component_name
    }

    fn same_identity(&self, other: &Self) -> bool {
        self.same_chart_identity(other)
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
        if !info_object.messages[0].data.is_empty() {
            return Err(ChartCaptionError::InvalidSource);
        }
        if mutation_guards {
            prove_exclusive_caption_standin(package, chart.chart_identifier, reference_identifier)?;
        }
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
    require_private_object(
        package,
        style_identifier,
        SHAPE_STYLE_MESSAGE_TYPE,
        Some(chart.slide_component_name.as_str()),
    )?;
    if mutation_guards {
        prove_exclusive_caption_storage(
            package,
            chart.chart_identifier,
            reference_identifier,
            storage_identifier,
        )?;
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
    chart_identifier: u64,
    caption_info_identifier: u64,
    storage_identifier: u64,
) -> Result<(), ChartCaptionError> {
    let mut payload_owner_seen = false;
    let mut metadata_owner_seen = false;
    let mut chart_edge_seen = false;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(ChartCaptionError::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == CHART_MESSAGE_TYPE
                    && chart_caption_reference(package, &message.data)?
                        == Some(caption_info_identifier)
                {
                    if std::mem::replace(&mut chart_edge_seen, true)
                        || owner_identifier != chart_identifier
                    {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                }
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
    if payload_owner_seen && metadata_owner_seen && chart_edge_seen {
        Ok(())
    } else {
        Err(ChartCaptionError::InvalidSource)
    }
}

fn prove_exclusive_caption_standin(
    package: &Package,
    chart_identifier: u64,
    standin_identifier: u64,
) -> Result<(), ChartCaptionError> {
    let mut payload_edges = 0usize;
    let mut aggregate_edges = 0usize;
    let mut field_edges = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner_identifier = object
                .archive_info
                .identifier
                .ok_or(ChartCaptionError::InvalidSource)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ == CHART_MESSAGE_TYPE
                    && chart_caption_reference(package, &message.data)? == Some(standin_identifier)
                {
                    payload_edges = payload_edges
                        .checked_add(1)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                    if payload_edges > 1 || owner_identifier != chart_identifier {
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
                    .filter(|identifier| **identifier == standin_identifier)
                    .count();
                let aggregate_data = info
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == standin_identifier)
                    .count();
                if aggregate_data != 0 {
                    return Err(ChartCaptionError::UnsupportedDependency);
                }
                if aggregate_count != 0 {
                    if owner_identifier != chart_identifier
                        || message.type_ != CHART_MESSAGE_TYPE
                        || aggregate_count != 1
                    {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                    aggregate_edges = aggregate_edges
                        .checked_add(aggregate_count)
                        .ok_or(ChartCaptionError::InvalidSource)?;
                }
                for field in &info.field_infos {
                    let field_count = field
                        .object_references
                        .iter()
                        .filter(|identifier| **identifier == standin_identifier)
                        .count();
                    let field_data = field
                        .data_references
                        .iter()
                        .filter(|identifier| **identifier == standin_identifier)
                        .count();
                    if field_data != 0 {
                        return Err(ChartCaptionError::UnsupportedDependency);
                    }
                    if field_count != 0 {
                        if owner_identifier != chart_identifier
                            || message.type_ != CHART_MESSAGE_TYPE
                            || field_count != 1
                            || !matches!(
                                field.path.path.as_slice(),
                                [11, 1] | [1, 11, 1] | [1, 1, 11, 1]
                            )
                        {
                            return Err(ChartCaptionError::UnsupportedDependency);
                        }
                        field_edges = field_edges
                            .checked_add(field_count)
                            .ok_or(ChartCaptionError::InvalidSource)?;
                    }
                }
            }
        }
    }
    if payload_edges == 1 && aggregate_edges == 1 && field_edges <= 1 {
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

/// Execute the graph transition used when a chart moves between its canonical
/// empty stand-in and an inline CaptionInfo graph.  Existing storage edits are
/// intentionally delegated to `slide_text`; this path only adds a graph or a
/// fresh stand-in and never removes the old graph.
fn rewrite_chart_caption_operation(
    source: &Package,
    selection: &CaptionSelection,
    after: Option<&str>,
) -> Result<(Package, usize, usize), ChartCaptionError> {
    let creating = selection.storage_identifier.is_none() && after.is_some();
    let removing = selection.storage_identifier.is_some() && after.is_none();
    if !creating && !removing {
        let storage = selection
            .storage_identifier
            .ok_or(ChartCaptionError::UnsupportedDependency)?;
        let desired = after.ok_or(ChartCaptionError::UnsupportedDependency)?;
        let end = selection
            .text
            .as_deref()
            .ok_or(ChartCaptionError::InvalidSource)?
            .encode_utf16()
            .count();
        if selection
            .text
            .as_deref()
            .is_some_and(contains_dependent_marker)
            || contains_dependent_marker(desired)
        {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        let deleted = physical_catalog(source).map(preview_count)?;
        let (package, touched) = super::slide_text::rewrite_owned_storage_text(
            source,
            storage,
            selection.slide_node_identifier,
            0..end,
            desired,
        )
        .map_err(map_slide_text_error)?;
        return Ok((package, touched, deleted));
    }

    let catalog = physical_catalog(source)?;
    let metadata_name = metadata_member_name(catalog, source)?;
    let slide_name = selection.slide_component_name.clone();
    if metadata_name == slide_name {
        return Err(ChartCaptionError::InvalidSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    let mut archives = Vec::new();
    for name in [&slide_name, &metadata_name] {
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == name.as_str())
            .ok_or(ChartCaptionError::InvalidSource)?;
        if entry.is_opaque() {
            return Err(ChartCaptionError::InvalidSource);
        }
        let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
            .map_err(map_core_error)?;
        let archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        archives.push((name.clone(), archive));
    }

    let metadata_source = metadata_payload(archive_ref(&archives, &metadata_name)?)?;
    let metadata_facts =
        caption_metadata_facts(source, &metadata_source.2, metadata_locator(&slide_name))?;
    let slide_selector = metadata_facts.selector(metadata_locator(&slide_name))?;
    let first_identifier = next_caption_identifier(source, metadata_facts.last_identifier)?;
    let (new_identifiers, replacement_identifier, graph_objects) = if creating {
        let first = first_identifier;
        let ids = CaptionGraphIds::allocate(first)?;
        let theme = caption_theme(source, selection)?;
        let width = caption_drawable_width(source, selection)?;
        let objects = caption_graph_objects(
            source,
            ids,
            selection.chart_identifier,
            width,
            after.ok_or(ChartCaptionError::InvalidSource)?,
            theme.stylesheet,
            theme.paragraph_style,
            theme.language.as_deref(),
        )?;
        (ids, ids.info, objects)
    } else {
        let standin = first_identifier;
        let object = canonical_standin(standin)?;
        (CaptionGraphIds::standin(standin), standin, vec![object])
    };

    let slide_archive = archive_mut(&mut archives, &slide_name)?;
    patch_chart_caption_edge(
        slide_archive,
        selection.chart_identifier,
        selection.reference_identifier,
        replacement_identifier,
        source,
        archive_limits,
    )?;
    for object in graph_objects {
        slide_archive
            .insert_object_with_limits(object, archive_limits)
            .map_err(map_core_error)?;
    }

    let new_last = new_identifiers.last();
    let mut additions = Vec::new();
    for identifier in new_identifiers.all() {
        if identifier == 0 {
            continue;
        }
        additions.push(ObjectUuidAddition::new(
            slide_selector,
            identifier,
            fresh_uuid(&metadata_facts.uuids, identifier),
        ));
    }
    let selectors = [slide_selector];
    let token_batch = SaveTokenBatch::new(&selectors);
    let addition_batch =
        MetadataBatch::new(metadata_facts.last_identifier, new_last, &additions, &[]);
    let metadata_output = rewrite_package_metadata_additions_and_save_tokens(
        &metadata_source.2,
        AdditionSaveTokenBatch::new(addition_batch, token_batch),
        metadata_options(source, additions.len())?,
    )
    .map_err(map_metadata_error)?
    .into_bytes();
    replace_metadata_payload(
        archive_mut(&mut archives, &metadata_name)?,
        metadata_source.0,
        metadata_source.1,
        metadata_output,
        archive_limits,
    )?;

    let mut compressed = Vec::new();
    for (name, archive) in &archives {
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let data = SnappyStream::compress(&bytes).map_err(map_core_error)?;
        compressed.push((name.clone(), data));
    }
    let edits = compressed
        .iter()
        .map(|(name, data)| EntryEdit::new(name.as_str(), data.as_slice()))
        .collect::<Vec<_>>();
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(map_rendering_error)?;
    let output = catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, previews.names(), physical_limits)
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((candidate, 2, previews.len()))
}

fn archive_ref<'a>(
    archives: &'a [(String, Archive)],
    name: &str,
) -> Result<&'a Archive, ChartCaptionError> {
    archives
        .iter()
        .find(|(candidate, _)| candidate == name)
        .map(|(_, archive)| archive)
        .ok_or(ChartCaptionError::InvalidSource)
}

fn archive_mut<'a>(
    archives: &'a mut [(String, Archive)],
    name: &str,
) -> Result<&'a mut Archive, ChartCaptionError> {
    archives
        .iter_mut()
        .find(|(candidate, _)| candidate == name)
        .map(|(_, archive)| archive)
        .ok_or(ChartCaptionError::InvalidSource)
}

fn metadata_member_name(
    catalog: &SourceCatalog,
    package: &Package,
) -> Result<String, ChartCaptionError> {
    let snappy_limits = package
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let mut found = false;
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    for entry in catalog.package().iter() {
        if entry.is_opaque() || !entry.name().ends_with(".iwa") {
            continue;
        }
        let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
            .map_err(map_core_error)?;
        let archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        for object in &archive.objects {
            for message in &object.messages {
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                    if entry.name() != PACKAGE_METADATA_MEMBER_NAME || found {
                        return Err(ChartCaptionError::InvalidSource);
                    }
                    found = true;
                }
            }
        }
    }
    found
        .then_some(PACKAGE_METADATA_MEMBER_NAME.to_owned())
        .ok_or(ChartCaptionError::InvalidSource)
}

fn metadata_payload(archive: &Archive) -> Result<(u64, usize, Vec<u8>), ChartCaptionError> {
    let mut selected = None;
    for object in &archive.objects {
        let identifier = object
            .archive_info
            .identifier
            .ok_or(ChartCaptionError::InvalidSource)?;
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                continue;
            }
            if selected
                .replace((identifier, index, message.data.clone()))
                .is_some()
            {
                return Err(ChartCaptionError::InvalidSource);
            }
        }
    }
    selected.ok_or(ChartCaptionError::InvalidSource)
}

fn replace_metadata_payload(
    archive: &mut Archive,
    object_identifier: u64,
    index: usize,
    data: Vec<u8>,
    limits: litchi_iwa_core::Limits,
) -> Result<(), ChartCaptionError> {
    let object = archive
        .object_mut(object_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if object.messages.get(index).map(|message| message.type_)
        != Some(PACKAGE_METADATA_MESSAGE_TYPE)
    {
        return Err(ChartCaptionError::InvalidSource);
    }
    object
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

fn metadata_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

fn metadata_options(
    package: &Package,
    max_additions: usize,
) -> Result<MetadataRewriteOptions, ChartCaptionError> {
    let wire = package.wire_limits().map_err(map_wire_error)?;
    let recursion = u32::try_from(wire.max_nesting()).unwrap_or(u32::MAX);
    let semantic = package.semantic_limits();
    Ok(MetadataRewriteOptions::new(
        wire.max_input_bytes(),
        wire.max_output_bytes(),
        wire.max_fields(),
        wire.max_rewrite_work(),
        recursion,
        semantic.max_objects(),
        semantic.max_references(),
        max_additions.max(1),
    ))
}

struct CaptionMetadataFacts {
    target_locator: String,
    selected_identifier: Option<u64>,
    selected_count: usize,
    uuids: HashSet<(u64, u64)>,
    last_identifier: u64,
}

impl CaptionMetadataFacts {
    fn selector<'a>(&self, locator: &'a str) -> Result<ComponentSelector<'a>, ChartCaptionError> {
        if self.selected_count != 1 {
            return Err(ChartCaptionError::InvalidSource);
        }
        Ok(ComponentSelector::new(
            self.selected_identifier
                .ok_or(ChartCaptionError::InvalidSource)?,
            locator,
        ))
    }
}

impl PackageMetadataVisitor for CaptionMetadataFacts {
    fn visit_component(
        &mut self,
        component: ComponentDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        // The exact locator is checked by `selector`; retain every current
        // component identifier only through the selected-component callback
        // installed by `caption_metadata_facts`.
        if component.is_current() && component.effective_locator() == self.target_locator {
            self.selected_identifier = Some(component.identifier());
            self.selected_count = self.selected_count.saturating_add(1);
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        self.uuids
            .insert((binding.uuid().lower(), binding.uuid().upper()));
        Ok(())
    }
}

fn caption_metadata_facts(
    package: &Package,
    source: &[u8],
    target_locator: &str,
) -> Result<CaptionMetadataFacts, ChartCaptionError> {
    let options = metadata_options(package, 4)?;
    let mut facts = CaptionMetadataFacts {
        target_locator: target_locator.to_owned(),
        selected_identifier: None,
        selected_count: 0,
        uuids: HashSet::new(),
        last_identifier: 0,
    };
    let inspection = inspect_package_metadata_with_visitor(source, options, &mut facts)
        .map_err(map_metadata_error)?;
    facts.last_identifier = inspection.last_object_identifier();
    Ok(facts)
}

#[derive(Debug, Clone, Copy)]
struct CaptionGraphIds {
    style: u64,
    info: u64,
    storage: u64,
    placement: u64,
}

impl CaptionGraphIds {
    fn allocate(first: u64) -> Result<Self, ChartCaptionError> {
        Ok(Self {
            style: first,
            info: first
                .checked_add(1)
                .ok_or(ChartCaptionError::InvalidSource)?,
            storage: first
                .checked_add(2)
                .ok_or(ChartCaptionError::InvalidSource)?,
            placement: first
                .checked_add(3)
                .ok_or(ChartCaptionError::InvalidSource)?,
        })
    }

    const fn standin(identifier: u64) -> Self {
        Self {
            style: 0,
            info: 0,
            storage: 0,
            placement: identifier,
        }
    }

    const fn all(self) -> [u64; 4] {
        [self.style, self.info, self.storage, self.placement]
    }

    const fn last(self) -> u64 {
        self.placement
    }
}

fn next_caption_identifier(
    package: &Package,
    metadata_last_identifier: u64,
) -> Result<u64, ChartCaptionError> {
    let mut maximum = metadata_last_identifier;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            maximum = maximum.max(
                object
                    .archive_info
                    .identifier
                    .ok_or(ChartCaptionError::InvalidSource)?,
            );
        }
    }
    maximum
        .checked_add(1)
        .ok_or(ChartCaptionError::InvalidSource)
}

fn fresh_uuid(existing: &HashSet<(u64, u64)>, identifier: u64) -> UuidBits {
    let mut lower = identifier.max(1);
    let mut upper = identifier.rotate_left(17) ^ 0x9e37_79b9_7f4a_7c15;
    while existing.contains(&(lower, upper)) || (lower == 0 && upper == 0) {
        lower = lower.wrapping_add(1).max(1);
        upper = upper.rotate_left(7) ^ 0xd1b5_4a32_d192_ed03;
    }
    UuidBits::new(lower, upper)
}

fn canonical_standin(identifier: u64) -> Result<ArchiveObject, ChartCaptionError> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: STANDIN_MESSAGE_TYPE,
            data: graph_codec::canonical_standin_payload().to_vec(),
        }],
    )
    .map_err(map_core_error)?;
    object.archive_info.message_infos[0].versions = vec![10, 1, 0];
    Ok(object)
}

fn caption_graph_objects(
    package: &Package,
    ids: CaptionGraphIds,
    chart_identifier: u64,
    drawable_width: f32,
    text: &str,
    stylesheet_identifier: u64,
    paragraph_style_identifier: u64,
    language: Option<&str>,
) -> Result<Vec<ArchiveObject>, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let wire = graph_codec::EncodeOptions::for_text(text)
        .with_max_output_bytes(limits.max_output_bytes())
        .with_max_text_bytes(limits.max_input_bytes())
        .with_max_fields(limits.max_fields())
        .with_max_work_bytes(limits.max_rewrite_work())
        .with_max_depth(u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX));
    let output = graph_codec::encode_caption_graph_with_report(
        graph_codec::CaptionGraphWrite {
            drawable_identifier: chart_identifier,
            style_identifier: ids.style,
            info_identifier: ids.info,
            storage_identifier: ids.storage,
            placement_identifier: ids.placement,
            stylesheet_identifier,
            paragraph_style_identifier,
            drawable_width,
            text,
            language,
        },
        wire,
    )
    .map_err(map_graph_error)?;
    let payloads = output.into_payloads().into_parts();
    let types = [
        SHAPE_STYLE_MESSAGE_TYPE,
        CAPTION_INFO_MESSAGE_TYPE,
        STORAGE_MESSAGE_TYPE,
        CAPTION_PLACEMENT_MESSAGE_TYPE,
    ];
    let identifiers = [ids.style, ids.info, ids.storage, ids.placement];
    let references = [
        vec![paragraph_style_identifier],
        vec![ids.style, ids.storage, ids.placement],
        vec![paragraph_style_identifier],
        Vec::new(),
    ];
    let mut objects = Vec::new();
    objects
        .try_reserve_exact(4)
        .map_err(|_error| ChartCaptionError::Allocation { amount: 4 })?;
    for (index, (payload, references)) in payloads.into_iter().zip(references).enumerate() {
        let mut object = ArchiveObject::new(
            identifiers[index],
            vec![RawMessage {
                type_: types[index],
                data: payload,
            }],
        )
        .map_err(map_core_error)?;
        object.archive_info.message_infos[0].versions = vec![1, 0, 5];
        object.archive_info.message_infos[0].object_references = references;
        objects.push(object);
    }
    Ok(objects)
}

#[derive(Debug, Clone)]
struct CaptionTheme {
    stylesheet: u64,
    paragraph_style: u64,
    language: Option<String>,
}

fn caption_theme(
    package: &Package,
    _selection: &CaptionSelection,
) -> Result<CaptionTheme, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let document = package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .find(|object| object.primary_message_type() == Some(1))
        .ok_or(ChartCaptionError::InvalidSource)?;
    let document_payload = document
        .messages
        .first()
        .ok_or(ChartCaptionError::InvalidSource)?
        .data
        .as_slice();
    let document_view =
        WireView::parse_with_limits(document_payload, limits).map_err(map_wire_error)?;
    let show_payload =
        unique_payload_field(&document_view, 2)?.ok_or(ChartCaptionError::InvalidSource)?;
    let show_identifier = reference_identifier(show_payload, limits)?;
    let show = package
        .object(show_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    let show_payload = show
        .messages
        .first()
        .ok_or(ChartCaptionError::InvalidSource)?
        .data
        .as_slice();
    let show_view = WireView::parse_with_limits(show_payload, limits).map_err(map_wire_error)?;
    let theme_identifier = reference_identifier(
        unique_payload_field(&show_view, 2)?.ok_or(ChartCaptionError::InvalidSource)?,
        limits,
    )?;
    let stylesheet_identifier = reference_identifier(
        unique_payload_field(&show_view, 5)?.ok_or(ChartCaptionError::InvalidSource)?,
        limits,
    )?;
    let theme = package
        .object(theme_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    let theme_payload = theme
        .messages
        .first()
        .ok_or(ChartCaptionError::InvalidSource)?
        .data
        .as_slice();
    let theme_view = WireView::parse_with_limits(theme_payload, limits).map_err(map_wire_error)?;
    let theme_super =
        unique_payload_field(&theme_view, 1)?.ok_or(ChartCaptionError::InvalidSource)?;
    let theme_super_view =
        WireView::parse_with_limits(theme_super, limits).map_err(map_wire_error)?;
    let presets =
        unique_payload_field(&theme_super_view, 210)?.ok_or(ChartCaptionError::InvalidSource)?;
    let presets_view = WireView::parse_with_limits(presets, limits).map_err(map_wire_error)?;
    let paragraph_style = first_reference_identifier(&presets_view, 1, limits)?;
    let document_super =
        unique_payload_field(&document_view, 3)?.ok_or(ChartCaptionError::InvalidSource)?;
    let document_super_view =
        WireView::parse_with_limits(document_super, limits).map_err(map_wire_error)?;
    let language = match unique_payload_field(&document_super_view, 3)? {
        Some(bytes) => Some(
            std::str::from_utf8(bytes)
                .map_err(|_| ChartCaptionError::InvalidSource)?
                .to_owned(),
        ),
        None => None,
    };
    Ok(CaptionTheme {
        stylesheet: stylesheet_identifier,
        paragraph_style,
        language,
    })
}

fn reference_identifier(
    payload: &[u8],
    limits: litchi_iwa_common::WireLimits,
) -> Result<u64, ChartCaptionError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let field = view
        .fields()
        .find(|field| field.number() == 1)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if field.wire_type() != 0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let (value, bytes) =
        decode_varint_from_bytes(field.payload()).map_err(|_| ChartCaptionError::InvalidSource)?;
    if bytes != field.payload().len() || bytes != encoded_len(value) || value == 0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    Ok(value)
}

fn first_reference_identifier(
    view: &WireView<'_>,
    number: u32,
    limits: litchi_iwa_common::WireLimits,
) -> Result<u64, ChartCaptionError> {
    let mut first = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 {
            return Err(ChartCaptionError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let identifier = reference_identifier(field.payload(), limits)?;
        if first.is_none() {
            first = Some(identifier);
        }
    }
    first.ok_or(ChartCaptionError::InvalidSource)
}

fn caption_drawable_width(
    package: &Package,
    selection: &CaptionSelection,
) -> Result<f32, ChartCaptionError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let object = package
        .object(selection.chart_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    let payload = &object
        .messages
        .first()
        .ok_or(ChartCaptionError::InvalidSource)?
        .data;
    let outer = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let drawable = unique_payload_field(&outer, 1)?.ok_or(ChartCaptionError::InvalidSource)?;
    let drawable_view = WireView::parse_with_limits(drawable, limits).map_err(map_wire_error)?;
    let geometry =
        unique_payload_field(&drawable_view, 1)?.ok_or(ChartCaptionError::InvalidSource)?;
    let geometry_view = WireView::parse_with_limits(geometry, limits).map_err(map_wire_error)?;
    let size = unique_payload_field(&geometry_view, 2)?.ok_or(ChartCaptionError::InvalidSource)?;
    let size_view = WireView::parse_with_limits(size, limits).map_err(map_wire_error)?;
    let width = size_view
        .fields()
        .find(|field| field.number() == 1 && field.wire_type() == 5)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if width.payload().len() != 4 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let value = f32::from_le_bytes(width.payload().try_into().unwrap_or([0; 4]));
    if !value.is_finite() || value <= 0.0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    Ok(value)
}

fn patch_chart_caption_edge(
    archive: &mut Archive,
    chart_identifier: u64,
    expected_identifier: Option<u64>,
    replacement_identifier: u64,
    package: &Package,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<(), ChartCaptionError> {
    let expected_identifier = expected_identifier.ok_or(ChartCaptionError::InvalidSource)?;
    if expected_identifier == replacement_identifier || replacement_identifier == 0 {
        return Err(ChartCaptionError::InvalidSource);
    }
    let object = archive
        .object_mut(chart_identifier)
        .ok_or(ChartCaptionError::InvalidSource)?;
    if object.messages.len() != 1 || object.messages[0].type_ != CHART_MESSAGE_TYPE {
        return Err(ChartCaptionError::InvalidSource);
    }
    validate_selected_message_metadata(object, 0)?;
    {
        let info = object
            .archive_info
            .message_infos
            .get_mut(0)
            .ok_or(ChartCaptionError::InvalidSource)?;
        if info.data_references.contains(&expected_identifier)
            || info.data_references.contains(&replacement_identifier)
            || info
                .object_references
                .iter()
                .filter(|identifier| **identifier == expected_identifier)
                .count()
                != 1
            || info.object_references.contains(&replacement_identifier)
        {
            return Err(ChartCaptionError::UnsupportedDependency);
        }
        for field in &mut info.field_infos {
            let count = field
                .object_references
                .iter()
                .filter(|identifier| **identifier == expected_identifier)
                .count();
            if count == 0 {
                if field.data_references.contains(&expected_identifier) {
                    return Err(ChartCaptionError::UnsupportedDependency);
                }
                continue;
            }
            if !matches!(
                field.path.path.as_slice(),
                [11, 1] | [1, 11, 1] | [1, 1, 11, 1]
            ) || count != 1
                || field.data_references.contains(&expected_identifier)
            {
                return Err(ChartCaptionError::UnsupportedDependency);
            }
            for identifier in &mut field.object_references {
                if *identifier == expected_identifier {
                    *identifier = replacement_identifier;
                }
            }
        }
        for identifier in &mut info.object_references {
            if *identifier == expected_identifier {
                *identifier = replacement_identifier;
            }
        }
    }
    let payload = object.messages[0].data.clone();
    let options = chart_caption_decode_options(package, &payload)?;
    let (rewritten, _) = keynote_chart_caption_codec::rewrite_chart_caption_with_report(
        &payload,
        keynote_chart_caption_codec::ChartCaptionWrite::new(replacement_identifier),
        options,
    )
    .map_err(map_chart_caption_codec_error)?;
    object
        .replace_message_preserving_header_with_limits(
            0,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

fn verify_graph_transition(
    source: &Package,
    candidate: &Package,
    before: &CaptionSelection,
    target: &CaptionSelection,
) -> Result<(), ChartCaptionError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let metadata_name = metadata_member_name(source_catalog, source)?;
    if metadata_member_name(candidate_catalog, candidate)? != metadata_name {
        return Err(ChartCaptionError::Verification);
    }
    let slide_name = before.slide_component_name.as_str();
    for entry in source_catalog.package().iter() {
        if PREVIEW_ENTRY_NAMES.contains(&entry.name())
            || entry.name() == slide_name
            || entry.name() == metadata_name
        {
            continue;
        }
        let candidate_entry = candidate_catalog
            .package()
            .iter()
            .find(|candidate_entry| candidate_entry.name() == entry.name())
            .ok_or(ChartCaptionError::Verification)?;
        if entry.data() != candidate_entry.data() {
            return Err(ChartCaptionError::Verification);
        }
    }
    let source_metadata = metadata_payload(&archive_for_member(source, &metadata_name)?)?;
    let candidate_metadata = metadata_payload(&archive_for_member(candidate, &metadata_name)?)?;
    if source_metadata.2 == candidate_metadata.2 {
        return Err(ChartCaptionError::Verification);
    }
    if before.storage_identifier.is_none() {
        // A stand-in -> active transition normally retains the source
        // stand-in.  The exact inverse of an active -> stand-in transition
        // has the same semantic orientation, but its source stand-in is the
        // newly allocated object and therefore is deliberately absent from
        // the restored candidate.  In both cases the target graph must be
        // present in full.
        if target.storage_identifier.is_none()
            || target.style_identifier.is_none()
            || target.placement_identifier.is_none()
            || target.caption_info_identifier.is_none()
            || ![
                target.caption_info_identifier,
                target.storage_identifier,
                target.placement_identifier,
                target.style_identifier,
            ]
            .into_iter()
            .flatten()
            .all(|identifier| candidate.object(identifier).is_some())
        {
            return Err(ChartCaptionError::Verification);
        }
        let source_standin_present = candidate
            .object(
                before
                    .reference_identifier
                    .ok_or(ChartCaptionError::Verification)?,
            )
            .is_some();
        if !source_standin_present
            && ![
                target.caption_info_identifier,
                target.storage_identifier,
                target.placement_identifier,
                target.style_identifier,
            ]
            .into_iter()
            .flatten()
            .all(|identifier| candidate.object(identifier).is_some())
        {
            return Err(ChartCaptionError::Verification);
        }
    } else if target.storage_identifier.is_some() {
        return Err(ChartCaptionError::Verification);
    } else {
        // Removal creates a fresh stand-in while retaining the old graph.
        // The inverse of creation restores the original stand-in and drops
        // that old graph, so either the target stand-in or the source graph
        // is the retained side of this exact transition.
        let target_standin = candidate
            .object(
                target
                    .reference_identifier
                    .ok_or(ChartCaptionError::Verification)?,
            )
            .is_some();
        if !target_standin {
            return Err(ChartCaptionError::Verification);
        }
        let source_graph_present = [
            before.caption_info_identifier,
            before.storage_identifier,
            before.placement_identifier,
            before.style_identifier,
        ]
        .into_iter()
        .flatten()
        .all(|identifier| candidate.object(identifier).is_some());
        if !source_graph_present && target.reference_identifier.is_none() {
            return Err(ChartCaptionError::Verification);
        }
    }
    Ok(())
}

fn archive_for_member(package: &Package, name: &str) -> Result<Archive, ChartCaptionError> {
    let catalog = physical_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(ChartCaptionError::Verification)?;
    let physical_limits = package.state.options.archive();
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    Archive::parse_with_limits(
        stream.as_bytes(),
        physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)
}

fn verify_caption_candidate(
    source: &Package,
    candidate: &Package,
    before: &CaptionSelection,
    target: &CaptionSelection,
    expected: Option<&str>,
    require_invalidated_previews: bool,
) -> Result<(), ChartCaptionError> {
    let source_count = source.state.total_objects;
    let candidate_count = candidate.state.total_objects;
    let count_matches = match (before.storage_identifier, target.storage_identifier) {
        // Forward creation adds four graph objects; an exact inverse of a
        // removal restores those four objects from a source that only has its
        // fresh stand-in.
        (None, Some(_)) => {
            candidate_count == source_count.saturating_add(4)
                || source_count == candidate_count.saturating_add(1)
        },
        // Forward removal adds one stand-in; an exact inverse of creation
        // removes the four graph objects while restoring the old stand-in.
        (Some(_), None) => {
            candidate_count == source_count.saturating_add(1)
                || source_count == candidate_count.saturating_add(4)
        },
        _ => candidate_count == source_count,
    };
    if !count_matches {
        return Err(ChartCaptionError::Verification);
    }
    let selected = select_caption(
        candidate,
        SlideSelector::position(before.slide_position),
        ChartSelector::index(before.chart_position.get()),
        true,
    )?;
    if !selected.same_identity(target) || selected.text.as_deref() != expected {
        return Err(ChartCaptionError::Verification);
    }
    if let (Some(storage_identifier), Some(target_storage)) =
        (before.storage_identifier, target.storage_identifier)
    {
        if storage_identifier != target_storage {
            return Err(ChartCaptionError::Verification);
        }
        super::slide_text::verify_owned_storage_candidate(
            source,
            candidate,
            storage_identifier,
            before.slide_node_identifier,
            require_invalidated_previews,
        )
        .map_err(map_slide_text_error)?;
    } else {
        verify_graph_transition(source, candidate, before, target)?;
        if require_invalidated_previews
            && !super::rendering_invalidation::root_previews_absent(
                candidate.state.source.package(),
            )
            .map_err(map_rendering_error)?
        {
            return Err(ChartCaptionError::Verification);
        }
    }
    Ok(())
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

fn map_archive_error(_error: litchi_iwa_archive::Error) -> ChartCaptionError {
    ChartCaptionError::InvalidSource
}

fn map_core_error(_error: litchi_iwa_core::Error) -> ChartCaptionError {
    ChartCaptionError::InvalidSource
}

fn map_metadata_error(_error: metadata_codec::RewriteError) -> ChartCaptionError {
    ChartCaptionError::InvalidSource
}

fn map_graph_error(_error: graph_codec::EncodeError) -> ChartCaptionError {
    ChartCaptionError::InvalidSource
}

fn map_rendering_error(
    _error: super::rendering_invalidation::RenderingInvalidationError,
) -> ChartCaptionError {
    ChartCaptionError::InvalidSource
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
