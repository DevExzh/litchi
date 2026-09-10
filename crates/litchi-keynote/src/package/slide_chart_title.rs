//! Exact-source, selector-first Keynote chart-title transactions.
//!
//! The public surface in this module is deliberately semantic: callers select
//! a slide and then a chart by checked position or exact visible title. The
//! native chart drawable, title stand-in, non-style object, component name,
//! and generated extension remain private to this adapter.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction redacts lower-layer failures and keeps native graph adapters private."
)]

use std::collections::HashMap;
use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, encode_varint_into,
    varint::encoded_len,
    wire::{WireField, WireView, parse_wire_fields_with_limits},
};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferenceScope, RawMessage, SnappyStream,
};
use litchi_iwa_protos::keynote_chart_title_codec::{
    ChartTitleWrite, DecodeError as ChartTitleDecodeError, DecodeLimit as ChartTitleDecodeLimit,
    DecodeOptions, WireResourceLimit as ChartTitleWireResourceLimit, decode_visible_chart_title,
    rewrite_chart_title as rewrite_chart_title_payload,
};
use thiserror::Error;

use super::chart_axis_support::{self, AxisSupportBudget, AxisSupportError};
use super::{
    Package, PhysicalSource, ReadError, SLIDE_MESSAGE_TYPE, SemanticLimitKind, SlideRecord,
    unique_payload,
};
use crate::{ChartCatalog, ChartSelector, ChartSelectorError, SlideSelector, SlideSelectorError};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const DRAWABLE_TITLE_FIELD: u32 = 10;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const CHART_NON_STYLE_FIELD: u32 = 10;
const CHART_MEDIATOR_FIELD: u32 = 8;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const GENERATED_CHART_NON_STYLE_EXTENSION_FIELD: u32 = 10_000;
const MAX_CHART_TITLE_BYTES: usize = 64 * 1024 * 1024;

/// A finite resource governed while a chart-title transaction is prepared or
/// published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartTitleLimitKind {
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
    /// Aggregate protobuf rewrite work.
    WireWork,
    /// UTF-8 bytes in one visible chart title.
    TitleBytes,
}

impl fmt::Display for ChartTitleLimitKind {
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
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::TitleBytes => "chart title bytes",
        })
    }
}

/// A content-redacted failure raised by a chart-title transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartTitleError {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical chart-title edits")]
    UnsupportedSource,
    /// An exact-name slide or chart selector was ambiguous.
    #[error("the Keynote chart-title selector is ambiguous")]
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
    /// The source chart graph, title stand-in, or selected native payload is
    /// malformed or unsupported.
    #[error("the Keynote chart-title source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote chart-title {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: ChartTitleLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote chart-title transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full candidate reopening did not reproduce the requested title.
    #[error("the edited Keynote chart title failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote chart-title patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable chart-title value staged against an immutable package snapshot.
pub struct ChartTitleEdit<'a> {
    source: &'a Package,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    non_style_identifier: u64,
    slide_identifier: u64,
    before: Option<String>,
    after: Option<String>,
}

impl fmt::Debug for ChartTitleEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartTitleEdit")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("has_before", &self.before.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> ChartTitleEdit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Self, ChartTitleError> {
        let selection = select_chart(source, slide_selector.into(), chart_selector.into(), true)?;
        let before = selection.title;
        let after = before.as_deref().map(copy_title).transpose()?;
        Ok(Self {
            source,
            slide_position: selection.slide_position,
            chart_position: selection.chart_position,
            chart_identifier: selection.chart_identifier,
            non_style_identifier: selection.non_style_identifier,
            slide_identifier: selection.slide_identifier,
            before,
            after,
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Borrow the visible title observed when this edit began.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Borrow the title staged for publication.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Stage a visible chart title.
    ///
    /// The input is copied into bounded edit-owned storage immediately, so a
    /// caller may pass either a borrowed string slice or an owned `String`
    /// without changing the transaction's source ownership or wire behavior.
    pub fn set(mut self, title: impl AsRef<str>) -> Result<Self, ChartTitleError> {
        self.after = Some(copy_title(title.as_ref())?);
        Ok(self)
    }

    /// Stage removal of the visible chart title.
    pub fn clear(mut self) -> Result<Self, ChartTitleError> {
        self.after = None;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<ChartTitleCommit, ChartTitleError> {
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let mut budget = ChartGraphScanBudget::new(self.source)?;
        budget
            .charge_selection_scans(self.source, true)
            .map_err(map_axis_support_error)?;
        let source_selection = select_chart_with_budget(
            self.source,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        if source_selection.chart_identifier != self.chart_identifier
            || source_selection.non_style_identifier != self.non_style_identifier
            || source_selection.slide_identifier != self.slide_identifier
            || source_selection.title != self.before
        {
            return Err(ChartTitleError::InvalidSource);
        }

        if self.before == self.after {
            self.source.validate().map_err(map_read_error)?;
            return Ok(ChartTitleCommit {
                package: self.source.snapshot(),
                patch: ChartTitlePatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    slide_position: self.slide_position,
                    chart_position: self.chart_position,
                    chart_identifier: self.chart_identifier,
                    non_style_identifier: self.non_style_identifier,
                    slide_identifier: self.slide_identifier,
                    before: self.before,
                    after: self.after,
                    deleted_previews: 0,
                    preview_count: 0,
                    target_requires_invalidated_previews: false,
                    before_message: Arc::from(Vec::<u8>::new()),
                    after_message: Arc::from(Vec::<u8>::new()),
                },
                diagnostics: ChartTitleDiagnostics::unchanged(),
            });
        }

        if !catalog.source_is_exact() {
            return Err(ChartTitleError::UnsupportedSource);
        }
        self.source.validate().map_err(map_read_error)?;
        validate_chart_title_reference_metadata(self.source, &mut budget)?;
        let rewrite = rewrite_chart_title(
            self.source,
            &source_selection,
            self.after.as_deref(),
            &mut budget,
        )?;
        rewrite.package.validate().map_err(map_read_error)?;
        let target = physical_catalog(&rewrite.package)?.shared_source();
        budget.charge(target.len())?;
        budget
            .charge_selection_scans(&rewrite.package, true)
            .map_err(map_axis_support_error)?;
        let candidate_selection = select_chart_with_budget(
            &rewrite.package,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        if candidate_selection.chart_identifier != self.chart_identifier
            || candidate_selection.non_style_identifier != self.non_style_identifier
            || candidate_selection.slide_identifier != self.slide_identifier
            || candidate_selection.title != self.after
        {
            return Err(ChartTitleError::Verification);
        }
        verify_chart_candidate(
            self.source,
            &rewrite.package,
            self.slide_position,
            self.chart_position,
            self.after.as_deref(),
            &rewrite.selected_message,
            true,
            &mut budget,
        )
        .map_err(map_candidate_verification_error)?;
        let ChartTitleRewrite {
            package,
            before_message,
            selected_message,
            deleted_previews,
        } = rewrite;
        let after_message: Arc<[u8]> = Arc::from(selected_message);
        Ok(ChartTitleCommit {
            package,
            patch: ChartTitlePatch {
                artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
                slide_position: self.slide_position,
                chart_position: self.chart_position,
                chart_identifier: self.chart_identifier,
                non_style_identifier: self.non_style_identifier,
                slide_identifier: self.slide_identifier,
                before: self.before,
                after: self.after,
                deleted_previews,
                preview_count: deleted_previews,
                target_requires_invalidated_previews: true,
                before_message,
                after_message,
            },
            diagnostics: ChartTitleDiagnostics::published(deleted_previews),
        })
    }
}

/// An exact-source-checked reversible semantic chart-title patch.
#[derive(Clone, PartialEq, Eq)]
pub struct ChartTitlePatch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    non_style_identifier: u64,
    slide_identifier: u64,
    before: Option<String>,
    after: Option<String>,
    deleted_previews: usize,
    preview_count: usize,
    target_requires_invalidated_previews: bool,
    before_message: Arc<[u8]>,
    after_message: Arc<[u8]>,
}

impl fmt::Debug for ChartTitlePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartTitlePatch")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("has_before", &self.before.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl ChartTitlePatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Borrow the visible title required from the source package.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Borrow the visible title produced by the target package.
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

    /// Return whether this patch preserves the exact source bytes and title.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            chart_position: self.chart_position,
            chart_identifier: self.chart_identifier,
            non_style_identifier: self.non_style_identifier,
            slide_identifier: self.slide_identifier,
            before: self.after.clone(),
            after: self.before.clone(),
            deleted_previews: if self.target_requires_invalidated_previews {
                0
            } else {
                self.preview_count
            },
            preview_count: self.preview_count,
            target_requires_invalidated_previews: !self.artifacts.is_byte_noop()
                && !self.target_requires_invalidated_previews,
            before_message: Arc::clone(&self.after_message),
            after_message: Arc::clone(&self.before_message),
        }
    }
}

/// Compact evidence describing one chart-title commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartTitleDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl ChartTitleDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
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

/// The fully verified result of one immutable chart-title transaction.
#[must_use = "a Keynote chart-title commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct ChartTitleCommit {
    package: Package,
    patch: ChartTitlePatch,
    diagnostics: ChartTitleDiagnostics,
}

impl ChartTitleCommit {
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
    pub const fn patch(&self) -> &ChartTitlePatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &ChartTitleDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the visible chart titles owned by one selected slide in z-order.
    pub fn slide_chart_catalog<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<ChartCatalog, ChartTitleError> {
        let position = resolve_slide_position(self, selector.into())?;
        let graphs = chart_graphs(self, position, true)?;
        let count = graphs.len();
        ChartCatalog::try_from_owned_titles(graphs.into_iter().map(|graph| graph.title))
            .map_err(|_error| ChartTitleError::Allocation { amount: count })
    }

    /// Read the visible title of one selected chart.
    pub fn slide_chart_title<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Option<String>, ChartTitleError> {
        Ok(select_chart(self, slide_selector.into(), chart_selector.into(), true)?.title)
    }

    /// Start an exact immutable edit of one selected chart title.
    pub fn edit_slide_chart_title<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartTitleEdit<'_>, ChartTitleError> {
        ChartTitleEdit::new(self, slide_selector, chart_selector)
    }

    /// Apply an exact-source-checked chart-title patch.
    pub fn apply_slide_chart_title(
        &self,
        patch: &ChartTitlePatch,
    ) -> Result<ChartTitleCommit, ChartTitleError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(ChartTitleError::PatchConflict);
        }
        let mut budget = ChartGraphScanBudget::new(self)?;
        budget
            .charge_selection_scans(self, true)
            .map_err(map_axis_support_error)?;
        let source_selection = select_chart_with_budget(
            self,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        if source_selection.chart_identifier != patch.chart_identifier
            || source_selection.non_style_identifier != patch.non_style_identifier
            || source_selection.slide_identifier != patch.slide_identifier
            || source_selection.title != patch.before
        {
            return Err(ChartTitleError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(ChartTitleCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: ChartTitleDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartTitleError::PatchConflict);
        }
        validate_chart_title_reference_metadata(self, &mut budget)?;
        verify_selected_chart_title_message(
            self,
            &source_selection,
            patch.before_message.as_ref(),
            &mut budget,
        )?;
        budget.charge_candidate_reopen(self, patch.artifacts.target().len(), None)?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        budget.charge(patch.artifacts.target().len())?;
        budget
            .charge_selection_scans(&candidate, true)
            .map_err(map_axis_support_error)?;
        let candidate_selection = select_chart_with_budget(
            &candidate,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        if candidate_selection.chart_identifier != patch.chart_identifier
            || candidate_selection.non_style_identifier != patch.non_style_identifier
            || candidate_selection.slide_identifier != patch.slide_identifier
            || candidate_selection.title != patch.after
        {
            return Err(ChartTitleError::Verification);
        }
        verify_chart_candidate(
            self,
            &candidate,
            patch.slide_position,
            patch.chart_position,
            patch.after.as_deref(),
            patch.after_message.as_ref(),
            patch.target_requires_invalidated_previews,
            &mut budget,
        )
        .map_err(map_candidate_verification_error)?;
        Ok(ChartTitleCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ChartTitleDiagnostics::published(patch.deleted_previews),
        })
    }
}

#[derive(Debug, Clone)]
pub(super) struct ChartSelection {
    pub(super) slide_position: Position,
    pub(super) chart_position: Position,
    pub(super) slide_identifier: u64,
    pub(super) chart_identifier: u64,
    pub(super) non_style_identifier: u64,
    pub(super) non_style_message_index: usize,
    pub(super) title: Option<String>,
    pub(super) slide_component_name: String,
    pub(super) non_style_component_name: String,
}

#[derive(Debug, Clone)]
struct ChartGraph {
    slide_identifier: u64,
    chart_identifier: u64,
    non_style_identifier: u64,
    non_style_message_index: usize,
    title_identifier: u64,
    title: Option<String>,
    slide_component_name: String,
    non_style_component_name: String,
}

pub(super) fn select_chart(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    chart_selector: ChartSelector<'_>,
    mutation_guards: bool,
) -> Result<ChartSelection, ChartTitleError> {
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let graphs = chart_graphs(package, slide_position, mutation_guards)?;
    let catalog = ChartCatalog::try_from_titles(graphs.iter().map(|graph| graph.title.as_deref()))
        .map_err(|_error| ChartTitleError::Allocation {
            amount: graphs.len(),
        })?;
    let chart_position = match chart_selector {
        ChartSelector::Index(index) => graphs.get(index).map(|_graph| Position::new(index)).ok_or(
            ChartTitleError::ChartPositionNotFound {
                position: Position::new(index),
            },
        )?,
        ChartSelector::Name(name) => {
            let selected = catalog
                .select_position(name)
                .map_err(map_chart_selector_error)?
                .ok_or(ChartTitleError::ChartNameNotFound)?;
            Position::new(selected)
        },
    };
    let graph = graphs
        .get(chart_position.get())
        .ok_or(ChartTitleError::InvalidSource)?;
    let title = graph.title.as_deref().map(copy_title).transpose()?;
    Ok(ChartSelection {
        slide_position,
        chart_position,
        slide_identifier: graph.slide_identifier,
        chart_identifier: graph.chart_identifier,
        non_style_identifier: graph.non_style_identifier,
        non_style_message_index: graph.non_style_message_index,
        title,
        slide_component_name: graph.slide_component_name.clone(),
        non_style_component_name: graph.non_style_component_name.clone(),
    })
}

/// Resolve every chart owned by one slide through the checked mutation graph.
///
/// The returned selections retain the same chart-drawable z-order used by
/// [`select_chart`], but build the graph and ownership index only once. This
/// is used by focused chart properties that need a complete per-slide view.
pub(super) fn select_charts(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    mutation_guards: bool,
) -> Result<Box<[ChartSelection]>, ChartTitleError> {
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let graphs = chart_graphs(package, slide_position, mutation_guards)?;
    let mut selections = Vec::new();
    selections
        .try_reserve_exact(graphs.len())
        .map_err(|_error| ChartTitleError::Allocation {
            amount: graphs.len(),
        })?;
    for (chart_position, graph) in graphs.into_iter().enumerate() {
        selections.push(ChartSelection {
            slide_position,
            chart_position: Position::new(chart_position),
            slide_identifier: graph.slide_identifier,
            chart_identifier: graph.chart_identifier,
            non_style_identifier: graph.non_style_identifier,
            non_style_message_index: graph.non_style_message_index,
            title: graph.title,
            slide_component_name: graph.slide_component_name,
            non_style_component_name: graph.non_style_component_name,
        });
    }
    Ok(selections.into_boxed_slice())
}

/// Resolve a chart selector while sharing the caller's aggregate axis
/// budget. The ordinary chart-title API intentionally keeps its historical
/// local accounting; axis owners enter through this private path so every
/// temporary graph/catalog allocation and nested wire scan is charged to the
/// same transaction ledger.
pub(super) fn select_chart_with_budget(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    chart_selector: ChartSelector<'_>,
    mutation_guards: bool,
    budget: &mut dyn AxisSupportBudget,
) -> Result<ChartSelection, AxisSupportError> {
    let slide_position = resolve_slide_position(package, slide_selector)
        .map_err(chart_axis_support::map_chart_title_error)?;
    let graphs = chart_graphs_with_budget(package, slide_position, mutation_guards, budget)?;
    charge_chart_catalog(&graphs, budget)?;
    let catalog = ChartCatalog::try_from_titles(graphs.iter().map(|graph| graph.title.as_deref()))
        .map_err(|_error| {
            chart_axis_support::map_chart_title_error(ChartTitleError::Allocation {
                amount: graphs.len(),
            })
        })?;
    charge_chart_catalog_selection(&graphs, chart_selector, budget)?;
    let chart_position = match chart_selector {
        ChartSelector::Index(index) => graphs.get(index).map(|_graph| Position::new(index)).ok_or(
            chart_axis_support::map_chart_title_error(ChartTitleError::ChartPositionNotFound {
                position: Position::new(index),
            }),
        )?,
        ChartSelector::Name(name) => {
            let selected = catalog
                .select_position(name)
                .map_err(map_chart_selector_error)
                .map_err(chart_axis_support::map_chart_title_error)?
                .ok_or_else(|| {
                    chart_axis_support::map_chart_title_error(ChartTitleError::ChartNameNotFound)
                })?;
            Position::new(selected)
        },
    };
    let graph = graphs
        .into_iter()
        .nth(chart_position.get())
        .ok_or(AxisSupportError::InvalidSource)?;
    let title = graph
        .title
        .as_deref()
        .map(|title| copy_title_with_budget(title, budget))
        .transpose()?;
    Ok(ChartSelection {
        slide_position,
        chart_position,
        slide_identifier: graph.slide_identifier,
        chart_identifier: graph.chart_identifier,
        non_style_identifier: graph.non_style_identifier,
        non_style_message_index: graph.non_style_message_index,
        title,
        slide_component_name: graph.slide_component_name,
        non_style_component_name: graph.non_style_component_name,
    })
}

fn selector_allocation_units(bytes: usize) -> Result<usize, AxisSupportError> {
    bytes
        .checked_add(size_of::<u64>() - 1)
        .and_then(|value| value.checked_div(size_of::<u64>()))
        .ok_or(AxisSupportError::InvalidSource)
}

fn charge_selector_allocation(
    budget: &mut dyn AxisSupportBudget,
    bytes: usize,
) -> Result<(), AxisSupportError> {
    let units = selector_allocation_units(bytes)?;
    if bytes == 0 {
        budget.charge_reference_vector(0)
    } else {
        budget.charge_reference_vector(units.max(1))
    }
}

fn charge_selector_allocations(
    budget: &mut dyn AxisSupportBudget,
    count: usize,
    bytes: usize,
) -> Result<(), AxisSupportError> {
    if count == 0 {
        return Ok(());
    }
    charge_selector_allocation(budget, bytes)?;
    for _ in 1..count {
        charge_selector_allocation(budget, 0)?;
    }
    Ok(())
}

fn charge_selector_string(
    budget: &mut dyn AxisSupportBudget,
    value: &str,
) -> Result<(), AxisSupportError> {
    budget.charge_work(value.len())?;
    if !value.is_empty() {
        charge_selector_allocation(budget, value.len())?;
    }
    Ok(())
}

fn copy_title_with_budget(
    title: &str,
    budget: &mut dyn AxisSupportBudget,
) -> Result<String, AxisSupportError> {
    if title.len() > MAX_CHART_TITLE_BYTES {
        return Err(AxisSupportError::LimitExceeded {
            kind: chart_axis_support::AxisSupportLimitKind::TitleBytes,
            observed: usize_to_u64(title.len()),
            maximum: usize_to_u64(MAX_CHART_TITLE_BYTES),
        });
    }
    charge_selector_string(budget, title)?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(title.len())
        .map_err(|_error| AxisSupportError::Allocation {
            amount: title.len(),
        })?;
    owned.push_str(title);
    Ok(owned)
}

fn charge_chart_catalog(
    graphs: &[ChartGraph],
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    // `try_from_titles` calls `try_reserve(1)` for each descriptor and then
    // converts the vector to a boxed slice. Charge the conservative upper
    // bound before entering that collection, including every borrowed-title
    // copy performed by `try_boxed_str`.
    let descriptor_bytes = graphs
        .len()
        .checked_mul(size_of::<crate::ChartDescriptor>())
        .and_then(|value| value.checked_mul(8))
        .ok_or(AxisSupportError::InvalidSource)?;
    budget.charge_work(graphs.len())?;
    let catalog_allocations = if graphs.is_empty() {
        0
    } else {
        graphs
            .len()
            .checked_add(1)
            .ok_or(AxisSupportError::InvalidSource)?
    };
    charge_selector_allocations(budget, catalog_allocations, descriptor_bytes)?;
    for graph in graphs {
        if let Some(title) = graph.title.as_deref() {
            charge_selector_string(budget, title)?;
        }
    }
    Ok(())
}

fn charge_chart_catalog_selection(
    graphs: &[ChartGraph],
    selector: ChartSelector<'_>,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    let ChartSelector::Name(name) = selector else {
        return Ok(());
    };
    let title_bytes = graphs.iter().try_fold(0usize, |total, graph| {
        total
            .checked_add(graph.title.as_deref().map_or(0, str::len))
            .ok_or(AxisSupportError::InvalidSource)
    })?;
    budget.charge_work(
        graphs
            .len()
            .checked_add(title_bytes)
            .and_then(|value| value.checked_add(name.len()))
            .ok_or(AxisSupportError::InvalidSource)?,
    )
}

fn chart_graph_error(error: ChartTitleError) -> AxisSupportError {
    chart_axis_support::map_chart_title_error(error)
}

pub(super) fn map_axis_support_error(error: AxisSupportError) -> ChartTitleError {
    match error {
        AxisSupportError::Selector(selector) => match selector {
            chart_axis_support::AxisSupportSelectorError::UnsupportedSource => {
                ChartTitleError::UnsupportedSource
            },
            chart_axis_support::AxisSupportSelectorError::AmbiguousSelector => {
                ChartTitleError::AmbiguousSelector
            },
            chart_axis_support::AxisSupportSelectorError::EmptySlideName => {
                ChartTitleError::EmptySlideName
            },
            chart_axis_support::AxisSupportSelectorError::SlideNameNotFound => {
                ChartTitleError::SlideNameNotFound
            },
            chart_axis_support::AxisSupportSelectorError::SlidePositionNotFound { position } => {
                ChartTitleError::SlidePositionNotFound { position }
            },
            chart_axis_support::AxisSupportSelectorError::ChartNameNotFound => {
                ChartTitleError::ChartNameNotFound
            },
            chart_axis_support::AxisSupportSelectorError::ChartPositionNotFound { position } => {
                ChartTitleError::ChartPositionNotFound { position }
            },
            chart_axis_support::AxisSupportSelectorError::EmptyChartName => {
                ChartTitleError::EmptyChartName
            },
        },
        AxisSupportError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ChartTitleError::LimitExceeded {
            kind: match kind {
                chart_axis_support::AxisSupportLimitKind::InputBytes => {
                    ChartTitleLimitKind::InputBytes
                },
                chart_axis_support::AxisSupportLimitKind::OutputBytes => {
                    ChartTitleLimitKind::OutputBytes
                },
                chart_axis_support::AxisSupportLimitKind::WireBytes => {
                    ChartTitleLimitKind::WireBytes
                },
                chart_axis_support::AxisSupportLimitKind::Entries => ChartTitleLimitKind::Entries,
                chart_axis_support::AxisSupportLimitKind::EntryBytes => {
                    ChartTitleLimitKind::EntryBytes
                },
                chart_axis_support::AxisSupportLimitKind::TotalBytes => {
                    ChartTitleLimitKind::TotalBytes
                },
                chart_axis_support::AxisSupportLimitKind::Slides => ChartTitleLimitKind::Slides,
                chart_axis_support::AxisSupportLimitKind::References => {
                    ChartTitleLimitKind::References
                },
                chart_axis_support::AxisSupportLimitKind::TextStorages => {
                    ChartTitleLimitKind::TextStorages
                },
                chart_axis_support::AxisSupportLimitKind::TextFragments => {
                    ChartTitleLimitKind::TextFragments
                },
                chart_axis_support::AxisSupportLimitKind::TextBytes => {
                    ChartTitleLimitKind::TextBytes
                },
                chart_axis_support::AxisSupportLimitKind::WireFields => {
                    ChartTitleLimitKind::WireFields
                },
                chart_axis_support::AxisSupportLimitKind::WireNesting => {
                    ChartTitleLimitKind::WireNesting
                },
                chart_axis_support::AxisSupportLimitKind::WireWork => ChartTitleLimitKind::WireWork,
                chart_axis_support::AxisSupportLimitKind::TitleBytes => {
                    ChartTitleLimitKind::TitleBytes
                },
            },
            observed,
            maximum,
        },
        AxisSupportError::Allocation { amount } => ChartTitleError::Allocation { amount },
        AxisSupportError::InvalidSource => ChartTitleError::InvalidSource,
    }
}

fn map_candidate_verification_error(error: ChartTitleError) -> ChartTitleError {
    match error {
        ChartTitleError::InvalidSource | ChartTitleError::Verification => {
            ChartTitleError::Verification
        },
        other => other,
    }
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, ChartTitleError> {
    match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(map_read_error)?
            .map(|_record| position)
            .ok_or(ChartTitleError::SlidePositionNotFound { position }),
        SlideSelector::Name(name) => {
            let selector = SlideSelector::try_name(name).map_err(map_slide_selector_error)?;
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(map_slide_selector_error)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(ChartTitleError::SlideNameNotFound)
        },
    }
}

fn chart_graphs(
    package: &Package,
    slide_position: Position,
    mutation_guards: bool,
) -> Result<Vec<ChartGraph>, ChartTitleError> {
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(ChartTitleError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (slide_component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(ChartTitleError::InvalidSource)?;
    let slide_payload = unique_payload(&slide.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
        .map_err(|_error| ChartTitleError::InvalidSource)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let owned = repeated_references(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits)?;
    let z_order = repeated_references(slide_payload, SLIDE_DRAWABLES_Z_ORDER_FIELD, limits)?;
    let mut graphs = Vec::new();
    graphs
        .try_reserve(z_order.len())
        .map_err(|_error| ChartTitleError::Allocation {
            amount: z_order.len(),
        })?;
    let reference_counts = ChartGraphReferenceCounts::new(&owned, &z_order)?;
    // The mutation guard below needs to prove that each selected non-style
    // object and title stand-in have exactly one chart owner. Keep the
    // package-wide index lazy so a slide with no charts never pays for the
    // ownership scan, then share it across every chart in this graph build.
    let mut graph_owners = None;
    for identifier in z_order.iter().copied() {
        let Some((component_name, drawable)) = package.object_with_component(identifier) else {
            return Err(ChartTitleError::InvalidSource);
        };
        let mut chart_message = None;
        for message in &drawable.messages {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            if chart_message.replace(message).is_some() {
                return Err(ChartTitleError::InvalidSource);
            }
        }
        let Some(chart_message) = chart_message else {
            continue;
        };
        if component_name != slide_component_name {
            return Err(ChartTitleError::InvalidSource);
        }
        let graph = chart_graph(
            package,
            record,
            slide_component_name,
            &reference_counts,
            &chart_message.data,
            identifier,
        )?;
        if mutation_guards {
            if graph_owners.is_none() {
                graph_owners = Some(scan_chart_graph_owners(package)?);
            }
            let owners = graph_owners
                .as_ref()
                .ok_or(ChartTitleError::InvalidSource)?;
            if owners.non_style.get(&graph.non_style_identifier).copied() != Some(1)
                || owners.title_standin.get(&graph.title_identifier).copied() != Some(1)
                || graph.chart_identifier == 0
            {
                return Err(ChartTitleError::InvalidSource);
            }
            if owners
                .non_style_metadata
                .get(&graph.non_style_identifier)
                .is_some_and(|references| !references.is_fully_allowed())
            {
                return Err(ChartTitleError::InvalidSource);
            }
        }
        graphs.push(graph);
    }
    Ok(graphs)
}

fn chart_graphs_with_budget(
    package: &Package,
    slide_position: Position,
    mutation_guards: bool,
    budget: &mut dyn AxisSupportBudget,
) -> Result<Vec<ChartGraph>, AxisSupportError> {
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)
        .map_err(chart_graph_error)?
        .ok_or(ChartTitleError::SlidePositionNotFound {
            position: slide_position,
        })
        .map_err(chart_graph_error)?;
    let (slide_component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(AxisSupportError::InvalidSource)?;
    let slide_payload = unique_payload(&slide.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
        .map_err(|_error| AxisSupportError::InvalidSource)?;
    let limits = package
        .wire_limits()
        .map_err(map_wire_error)
        .map_err(chart_graph_error)?;
    let owned = chart_axis_support::repeated_references(
        slide_payload,
        SLIDE_OWNED_DRAWABLES_FIELD,
        limits,
        budget,
    )?;
    let z_order = chart_axis_support::repeated_references(
        slide_payload,
        SLIDE_DRAWABLES_Z_ORDER_FIELD,
        limits,
        budget,
    )?;
    let mut graphs = Vec::new();
    if !z_order.is_empty() {
        let graph_bytes = z_order
            .len()
            .checked_mul(size_of::<ChartGraph>())
            .and_then(|value| value.checked_mul(4))
            .ok_or(AxisSupportError::InvalidSource)?;
        charge_selector_allocation(budget, graph_bytes)?;
    }
    graphs
        .try_reserve(z_order.len())
        .map_err(|_error| AxisSupportError::Allocation {
            amount: z_order.len(),
        })?;
    let reference_counts = ChartGraphReferenceCounts::new_with_budget(&owned, &z_order, budget)?;
    // Keep the package-wide owner index lazy so a slide with no charts does
    // not retain its maps. The shared path charges each wire parse and map
    // insertion directly on the caller's ledger.
    let mut graph_owners = None;
    for identifier in z_order.iter().copied() {
        budget.charge_work(1)?;
        let Some((component_name, drawable)) = package.object_with_component(identifier) else {
            return Err(AxisSupportError::InvalidSource);
        };
        budget.charge_work(drawable.messages.len())?;
        let mut chart_message = None;
        for message in &drawable.messages {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            if chart_message.replace(message).is_some() {
                return Err(AxisSupportError::InvalidSource);
            }
        }
        let Some(chart_message) = chart_message else {
            continue;
        };
        if component_name != slide_component_name {
            return Err(AxisSupportError::InvalidSource);
        }
        let graph = chart_graph_with_budget(
            package,
            record,
            slide_component_name,
            &reference_counts,
            &chart_message.data,
            identifier,
            budget,
        )?;
        if mutation_guards {
            if graph_owners.is_none() {
                graph_owners = Some(scan_chart_graph_owners_with_budget(package, budget)?);
            }
            let owners = graph_owners
                .as_ref()
                .ok_or(AxisSupportError::InvalidSource)?;
            budget.charge_work(3)?;
            if owners.non_style.get(&graph.non_style_identifier).copied() != Some(1)
                || owners.title_standin.get(&graph.title_identifier).copied() != Some(1)
                || graph.chart_identifier == 0
            {
                return Err(AxisSupportError::InvalidSource);
            }
            if owners
                .non_style_metadata
                .get(&graph.non_style_identifier)
                .is_some_and(|references| !references.is_fully_allowed())
            {
                return Err(AxisSupportError::InvalidSource);
            }
        }
        graphs.push(graph);
    }
    Ok(graphs)
}

fn chart_graph(
    package: &Package,
    record: SlideRecord,
    slide_component_name: &str,
    reference_counts: &ChartGraphReferenceCounts,
    chart_data: &[u8],
    chart_identifier: u64,
) -> Result<ChartGraph, ChartTitleError> {
    if reference_counts.owned.get(&chart_identifier).copied() != Some(1)
        || reference_counts.z_order.get(&chart_identifier).copied() != Some(1)
    {
        return Err(ChartTitleError::InvalidSource);
    }
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let outer = parse_wire_fields_with_limits(chart_data, limits).map_err(map_wire_error)?;
    let super_payload = unique_length_delimited_field(&outer, chart_data, DRAWABLE_SUPER_FIELD)?
        .ok_or(ChartTitleError::InvalidSource)?;
    let chart_payload = unique_length_delimited_field(&outer, chart_data, CHART_EXTENSION_FIELD)?
        .ok_or(ChartTitleError::InvalidSource)?;
    let drawable = parse_wire_fields_with_limits(super_payload, limits).map_err(map_wire_error)?;
    let parent = required_reference_field(&drawable, super_payload, DRAWABLE_PARENT_FIELD, limits)?;
    let title_identifier =
        required_reference_field(&drawable, super_payload, DRAWABLE_TITLE_FIELD, limits)?;
    if parent != record.slide_identifier || title_identifier == 0 {
        return Err(ChartTitleError::InvalidSource);
    }
    let chart_fields =
        parse_wire_fields_with_limits(chart_payload, limits).map_err(map_wire_error)?;
    if let Some(mediator_payload) =
        unique_length_delimited_field(&chart_fields, chart_payload, CHART_MEDIATOR_FIELD)?
    {
        if super::validate_reference_payload(mediator_payload, limits, "Keynote chart mediator")
            .map_err(map_wire_error)?
            != 0
        {
            return Err(ChartTitleError::InvalidSource);
        }
    }
    let non_style_identifier =
        required_reference_field(&chart_fields, chart_payload, CHART_NON_STYLE_FIELD, limits)?;
    if non_style_identifier == 0
        || title_identifier == chart_identifier
        || title_identifier == non_style_identifier
        || non_style_identifier == chart_identifier
    {
        return Err(ChartTitleError::InvalidSource);
    }
    let (title_component, title_object) = package
        .object_with_component(title_identifier)
        .ok_or(ChartTitleError::InvalidSource)?;
    let (non_style_component, non_style_object) = package
        .object_with_component(non_style_identifier)
        .ok_or(ChartTitleError::InvalidSource)?;
    if title_component != slide_component_name {
        return Err(ChartTitleError::InvalidSource);
    }
    let (title_message_index, _) =
        exactly_one_message_with_index(title_object, STANDIN_MESSAGE_TYPE)?;
    chart_axis_support::validate_selected_message_metadata(title_object, title_message_index)
        .map_err(map_axis_support_error)?;
    let (non_style_message_index, non_style_message) =
        exactly_one_message_with_index(non_style_object, CHART_NON_STYLE_MESSAGE_TYPE)?;
    chart_axis_support::validate_selected_message_metadata(
        non_style_object,
        non_style_message_index,
    )
    .map_err(map_axis_support_error)?;
    let (chart_component, chart_object) = package
        .object_with_component(chart_identifier)
        .ok_or(ChartTitleError::InvalidSource)?;
    if chart_component != slide_component_name {
        return Err(ChartTitleError::InvalidSource);
    }
    let (chart_message_index, _) =
        exactly_one_message_with_index(chart_object, CHART_MESSAGE_TYPE)?;
    chart_axis_support::validate_selected_message_metadata(chart_object, chart_message_index)
        .map_err(map_axis_support_error)?;
    let title = read_chart_title(non_style_message.data.as_slice(), limits)?;
    Ok(ChartGraph {
        slide_identifier: record.slide_identifier,
        chart_identifier,
        non_style_identifier,
        non_style_message_index,
        title_identifier,
        title,
        slide_component_name: slide_component_name.to_owned(),
        non_style_component_name: non_style_component.to_owned(),
    })
}

fn chart_graph_with_budget(
    package: &Package,
    record: SlideRecord,
    slide_component_name: &str,
    reference_counts: &ChartGraphReferenceCounts,
    chart_data: &[u8],
    chart_identifier: u64,
    budget: &mut dyn AxisSupportBudget,
) -> Result<ChartGraph, AxisSupportError> {
    budget.charge_work(2)?;
    if reference_counts.owned.get(&chart_identifier).copied() != Some(1)
        || reference_counts.z_order.get(&chart_identifier).copied() != Some(1)
    {
        return Err(AxisSupportError::InvalidSource);
    }
    let limits = package
        .wire_limits()
        .map_err(map_wire_error)
        .map_err(chart_graph_error)?;
    let outer = chart_axis_support::accounted_wire_fields(chart_data, limits, budget)?;
    let super_payload = chart_axis_support::unique_length_delimited_field(
        &outer,
        chart_data,
        DRAWABLE_SUPER_FIELD,
        budget,
    )?
    .ok_or(AxisSupportError::InvalidSource)?;
    let chart_payload = chart_axis_support::unique_length_delimited_field(
        &outer,
        chart_data,
        CHART_EXTENSION_FIELD,
        budget,
    )?
    .ok_or(AxisSupportError::InvalidSource)?;
    let chart_fields = chart_axis_support::accounted_wire_fields(chart_payload, limits, budget)?;
    if let Some(mediator_payload) = chart_axis_support::unique_length_delimited_field(
        &chart_fields,
        chart_payload,
        CHART_MEDIATOR_FIELD,
        budget,
    )? {
        budget.charge_wire_vector(mediator_payload.len())?;
        budget.finish_wire_scan(mediator_payload.len())?;
        if super::validate_reference_payload(mediator_payload, limits, "Keynote chart mediator")
            .map_err(map_wire_error)
            .map_err(chart_graph_error)?
            != 0
        {
            return Err(AxisSupportError::InvalidSource);
        }
    }
    let parent = chart_axis_support::required_reference(
        super_payload,
        DRAWABLE_PARENT_FIELD,
        limits,
        budget,
    )?;
    let title_identifier = chart_axis_support::required_reference(
        super_payload,
        DRAWABLE_TITLE_FIELD,
        limits,
        budget,
    )?;
    if parent != record.slide_identifier || title_identifier == 0 {
        return Err(AxisSupportError::InvalidSource);
    }
    let non_style_identifier = chart_axis_support::required_reference(
        chart_payload,
        CHART_NON_STYLE_FIELD,
        limits,
        budget,
    )?;
    if non_style_identifier == 0
        || title_identifier == chart_identifier
        || title_identifier == non_style_identifier
        || non_style_identifier == chart_identifier
    {
        return Err(AxisSupportError::InvalidSource);
    }
    let (title_component, title_object) = package
        .object_with_component(title_identifier)
        .ok_or(AxisSupportError::InvalidSource)?;
    let (non_style_component, non_style_object) = package
        .object_with_component(non_style_identifier)
        .ok_or(AxisSupportError::InvalidSource)?;
    if title_component != slide_component_name {
        return Err(AxisSupportError::InvalidSource);
    }
    let (title_message_index, _title_message) =
        chart_axis_support::unique_message(title_object, STANDIN_MESSAGE_TYPE, budget)?;
    chart_axis_support::validate_selected_message_metadata(title_object, title_message_index)?;
    let (non_style_message_index, non_style_message) =
        chart_axis_support::unique_message(non_style_object, CHART_NON_STYLE_MESSAGE_TYPE, budget)?;
    chart_axis_support::validate_selected_message_metadata(
        non_style_object,
        non_style_message_index,
    )?;
    let (chart_component, chart_object) = package
        .object_with_component(chart_identifier)
        .ok_or(AxisSupportError::InvalidSource)?;
    if chart_component != slide_component_name {
        return Err(AxisSupportError::InvalidSource);
    }
    let (chart_message_index, _) =
        chart_axis_support::unique_message(chart_object, CHART_MESSAGE_TYPE, budget)?;
    chart_axis_support::validate_selected_message_metadata(chart_object, chart_message_index)?;
    let title = read_chart_title_with_budget(non_style_message.data.as_slice(), limits, budget)?;
    charge_selector_string(budget, slide_component_name)?;
    charge_selector_string(budget, non_style_component)?;
    Ok(ChartGraph {
        slide_identifier: record.slide_identifier,
        chart_identifier,
        non_style_identifier,
        non_style_message_index,
        title_identifier,
        title,
        slide_component_name: slide_component_name.to_owned(),
        non_style_component_name: non_style_component.to_owned(),
    })
}

#[derive(Debug, Default)]
struct ChartGraphReferenceCounts {
    owned: HashMap<u64, usize>,
    z_order: HashMap<u64, usize>,
}

impl ChartGraphReferenceCounts {
    fn new(owned: &[u64], z_order: &[u64]) -> Result<Self, ChartTitleError> {
        let mut counts = Self::default();
        counts
            .owned
            .try_reserve(owned.len())
            .map_err(|_error| ChartTitleError::Allocation {
                amount: owned.len(),
            })?;
        counts
            .z_order
            .try_reserve(z_order.len())
            .map_err(|_error| ChartTitleError::Allocation {
                amount: z_order.len(),
            })?;
        for &identifier in owned {
            increment_chart_reference_count(&mut counts.owned, identifier)?;
        }
        for &identifier in z_order {
            increment_chart_reference_count(&mut counts.z_order, identifier)?;
        }
        Ok(counts)
    }

    fn new_with_budget(
        owned: &[u64],
        z_order: &[u64],
        budget: &mut dyn AxisSupportBudget,
    ) -> Result<Self, AxisSupportError> {
        let mut counts = Self::default();
        charge_selector_hash_map(owned.len(), budget)?;
        charge_selector_hash_map(z_order.len(), budget)?;
        budget.charge_work(
            owned
                .len()
                .checked_add(z_order.len())
                .ok_or(AxisSupportError::InvalidSource)?,
        )?;
        counts
            .owned
            .try_reserve(owned.len())
            .map_err(|_error| AxisSupportError::Allocation {
                amount: owned.len(),
            })?;
        counts
            .z_order
            .try_reserve(z_order.len())
            .map_err(|_error| AxisSupportError::Allocation {
                amount: z_order.len(),
            })?;
        for &identifier in owned {
            increment_chart_reference_count_with_budget(&mut counts.owned, identifier, budget)?;
        }
        for &identifier in z_order {
            increment_chart_reference_count_with_budget(&mut counts.z_order, identifier, budget)?;
        }
        Ok(counts)
    }
}

fn increment_chart_reference_count(
    counts: &mut HashMap<u64, usize>,
    identifier: u64,
) -> Result<(), ChartTitleError> {
    let count = counts.entry(identifier).or_insert(0);
    *count = count.checked_add(1).ok_or(ChartTitleError::InvalidSource)?;
    Ok(())
}

fn increment_chart_reference_count_with_budget(
    counts: &mut HashMap<u64, usize>,
    identifier: u64,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    budget.charge_work(1)?;
    let count = counts.entry(identifier).or_insert(0);
    *count = count
        .checked_add(1)
        .ok_or(AxisSupportError::InvalidSource)?;
    Ok(())
}

fn charge_selector_hash_map(
    entries: usize,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    if entries == 0 {
        return Ok(());
    }
    // HashMap::try_reserve(entries) allocates a bucket table whose capacity
    // is implementation-defined. Four key/value slots per requested entry
    // cover the current SwissTable growth/control-byte overhead while
    // remaining a checked conservative bound for future allocators.
    let bytes = entries
        .checked_mul(size_of::<(u64, usize)>())
        .and_then(|value| value.checked_mul(4))
        .ok_or(AxisSupportError::InvalidSource)?;
    charge_selector_allocation(budget, bytes)
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
) -> Result<Vec<u64>, ChartTitleError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut references = Vec::new();
    references
        .try_reserve(fields.len())
        .map_err(|_error| ChartTitleError::Allocation {
            amount: fields.len(),
        })?;
    for field in fields.fields() {
        if field.number() != field_number {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(ChartTitleError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        references.push(
            super::validate_reference_payload(field.payload(), limits, "Keynote chart drawable")
                .map_err(map_wire_error)?,
        );
    }
    Ok(references)
}

fn unique_length_delimited_field<'a>(
    fields: &[WireField],
    source: &'a [u8],
    number: u32,
) -> Result<Option<&'a [u8]>, ChartTitleError> {
    let mut selected = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number() == number)
    {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(ChartTitleError::InvalidSource);
        }
        field
            .validate_canonical_framing(source)
            .map_err(map_wire_error)?;
        selected = Some(field.payload(source).map_err(map_wire_error)?);
    }
    Ok(selected)
}

fn required_reference_field(
    fields: &[WireField],
    source: &[u8],
    number: u32,
    limits: WireLimits,
) -> Result<u64, ChartTitleError> {
    let payload = unique_length_delimited_field(fields, source, number)?
        .ok_or(ChartTitleError::InvalidSource)?;
    super::validate_reference_payload(payload, limits, "Keynote chart reference")
        .map_err(map_wire_error)
}

fn exactly_one_message_with_index(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &RawMessage), ChartTitleError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace((index, message)).is_some() {
            return Err(ChartTitleError::InvalidSource);
        }
    }
    selected.ok_or(ChartTitleError::InvalidSource)
}

pub(super) struct ChartGraphScanBudget {
    limits: WireLimits,
    maximum_work: usize,
    work: usize,
}

impl ChartGraphScanBudget {
    pub(super) fn new(package: &Package) -> Result<Self, ChartTitleError> {
        let limits = package.wire_limits().map_err(map_wire_error)?;
        Ok(Self {
            maximum_work: limits
                .max_rewrite_work()
                .checked_mul(16)
                .ok_or(ChartTitleError::InvalidSource)?,
            limits,
            work: 0,
        })
    }

    fn charge(&mut self, amount: usize) -> Result<(), ChartTitleError> {
        let observed = self
            .work
            .checked_add(amount)
            .ok_or(ChartTitleError::InvalidSource)?;
        if observed > self.maximum_work {
            return Err(ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::WireWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(self.maximum_work),
            });
        }
        self.work = observed;
        Ok(())
    }

    pub(super) fn charge_chart_metadata_report(
        &mut self,
        report: litchi_iwa_protos::chart_metadata_codec::DecodeReport,
    ) -> Result<(), ChartTitleError> {
        let amount = report
            .source_bytes()
            .checked_add(report.fields())
            .and_then(|value| value.checked_add(report.work_bytes()))
            .and_then(|value| value.checked_add(report.failure_work_bytes()))
            .and_then(|value| value.checked_add(report.max_depth() as usize))
            .and_then(|value| value.checked_add(report.label_count()))
            .and_then(|value| value.checked_add(report.text_bytes()))
            .and_then(|value| value.checked_add(report.allocations()))
            .and_then(|value| value.checked_add(report.retained_bytes()))
            .ok_or(ChartTitleError::InvalidSource)?;
        self.charge(amount)
    }

    pub(super) fn chart_metadata_residual_limits(&self) -> (WireLimits, usize) {
        (self.limits, self.maximum_work.saturating_sub(self.work))
    }

    fn parse(&mut self, payload: &[u8]) -> Result<Vec<WireField>, ChartTitleError> {
        self.charge(payload.len())?;
        let fields = parse_wire_fields_with_limits(payload, self.limits).map_err(map_wire_error)?;
        self.charge(fields.len())?;
        Ok(fields)
    }

    fn charge_native_decompress_parse(
        &mut self,
        compressed: usize,
        decompressed: usize,
        objects: usize,
    ) -> Result<(), ChartTitleError> {
        self.charge(
            compressed
                .checked_add(decompressed)
                .ok_or(ChartTitleError::InvalidSource)?,
        )?;
        self.charge(
            decompressed
                .checked_mul(2)
                .and_then(|value| value.checked_add(objects))
                .and_then(|value| value.checked_add(2))
                .ok_or(ChartTitleError::InvalidSource)?,
        )
    }

    fn charge_archive_snappy_plan(
        &mut self,
        archive: &Archive,
        archive_limits: litchi_iwa_core::Limits,
        snappy_limits: litchi_iwa_core::SnappyLimits,
    ) -> Result<(), ChartTitleError> {
        let encoded = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let compressed = SnappyStream::maximum_compressed_len(encoded).map_err(map_core_error)?;
        if compressed > snappy_limits.max_compressed_stream() {
            return Err(ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::EntryBytes,
                observed: usize_to_u64(compressed),
                maximum: usize_to_u64(snappy_limits.max_compressed_stream()),
            });
        }
        let total = encoded
            .checked_add(compressed)
            .ok_or(ChartTitleError::InvalidSource)?;
        self.charge(total)?;
        self.charge(total)
    }

    fn charge_candidate_reopen(
        &mut self,
        source: &Package,
        candidate_bytes: usize,
        replacement: Option<(&str, &Archive)>,
    ) -> Result<(), ChartTitleError> {
        let physical_limits = source.state.options.archive();
        let candidate_u64 =
            u64::try_from(candidate_bytes).map_err(|_error| ChartTitleError::InvalidSource)?;
        if candidate_u64 > physical_limits.max_input_bytes() {
            return Err(ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::InputBytes,
                observed: candidate_u64,
                maximum: physical_limits.max_input_bytes(),
            });
        }
        let catalog = physical_catalog(source)?.package();
        if catalog.len() > physical_limits.max_entries() {
            return Err(ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::Entries,
                observed: usize_to_u64(catalog.len()),
                maximum: usize_to_u64(physical_limits.max_entries()),
            });
        }
        let archive_limits = physical_limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let mut decompressed = 0usize;
        let mut largest_component = 0usize;
        let mut objects = 0usize;
        let mut messages = 0usize;
        let mut metadata_fields = 0usize;
        let mut references = 0usize;
        for component in source.state.source.components().iter() {
            let archive = replacement
                .filter(|(name, _archive)| *name == component.name())
                .map_or(component.archive(), |(_name, archive)| archive);
            let bytes = archive
                .encoded_len_with_limits(archive_limits)
                .map_err(map_core_error)?;
            if bytes > physical_limits.max_iwa_stream_bytes() {
                return Err(ChartTitleError::LimitExceeded {
                    kind: ChartTitleLimitKind::EntryBytes,
                    observed: usize_to_u64(bytes),
                    maximum: usize_to_u64(physical_limits.max_iwa_stream_bytes()),
                });
            }
            decompressed = decompressed
                .checked_add(bytes)
                .ok_or(ChartTitleError::InvalidSource)?;
            largest_component = largest_component.max(bytes);
            objects = objects
                .checked_add(archive.objects.len())
                .ok_or(ChartTitleError::InvalidSource)?;
            for object in &archive.objects {
                messages = messages
                    .checked_add(object.messages.len())
                    .ok_or(ChartTitleError::InvalidSource)?;
                metadata_fields = metadata_fields
                    .checked_add(object.archive_info.message_infos.len())
                    .ok_or(ChartTitleError::InvalidSource)?;
                for info in &object.archive_info.message_infos {
                    references = references
                        .checked_add(info.object_references.len())
                        .and_then(|value| value.checked_add(info.data_references.len()))
                        .ok_or(ChartTitleError::InvalidSource)?;
                    metadata_fields = metadata_fields
                        .checked_add(info.field_infos.len())
                        .ok_or(ChartTitleError::InvalidSource)?;
                    for field in &info.field_infos {
                        references = references
                            .checked_add(field.object_references.len())
                            .and_then(|value| value.checked_add(field.data_references.len()))
                            .ok_or(ChartTitleError::InvalidSource)?;
                    }
                }
            }
        }
        let logical = objects
            .checked_mul(3 * size_of::<usize>())
            .and_then(|value| value.checked_add(messages.checked_mul(size_of::<RawMessage>())?))
            .and_then(|value| value.checked_add(references.checked_mul(size_of::<u64>())?))
            .ok_or(ChartTitleError::InvalidSource)?;
        let catalog_index = catalog
            .len()
            .checked_mul(size_of::<[usize; 8]>())
            .ok_or(ChartTitleError::InvalidSource)?;
        let retained = candidate_bytes
            .checked_add(
                decompressed
                    .checked_mul(2)
                    .ok_or(ChartTitleError::InvalidSource)?,
            )
            .and_then(|value| value.checked_add(logical))
            .and_then(|value| value.checked_add(catalog_index))
            .ok_or(ChartTitleError::InvalidSource)?;
        let scratch = largest_component
            .checked_add(candidate_bytes)
            .and_then(|value| value.checked_add(catalog_index))
            .ok_or(ChartTitleError::InvalidSource)?;
        let allocations = catalog
            .len()
            .checked_mul(2)
            .and_then(|value| value.checked_add(source.state.source.components().len()))
            .and_then(|value| value.checked_add(objects.checked_mul(3)?))
            .and_then(|value| value.checked_add(messages.checked_mul(3)?))
            .and_then(|value| value.checked_add(metadata_fields))
            .and_then(|value| value.checked_add(4))
            .ok_or(ChartTitleError::InvalidSource)?;
        let work = candidate_bytes
            .checked_add(decompressed)
            .and_then(|value| value.checked_add(objects))
            .and_then(|value| value.checked_add(messages))
            .and_then(|value| value.checked_add(metadata_fields))
            .and_then(|value| value.checked_add(references))
            .ok_or(ChartTitleError::InvalidSource)?;
        self.charge(work)?;
        self.charge(retained)?;
        self.charge(scratch)?;
        self.charge(allocations)
    }
}

impl AxisSupportBudget for ChartGraphScanBudget {
    fn charge_selection_scans(
        &mut self,
        package: &Package,
        mutation_guards: bool,
    ) -> Result<(), AxisSupportError> {
        let passes = usize::from(mutation_guards)
            .checked_add(1)
            .ok_or(AxisSupportError::InvalidSource)?;
        let amount = package
            .state
            .source
            .components()
            .len()
            .checked_add(package.state.total_objects)
            .and_then(|value| value.checked_mul(passes))
            .ok_or(AxisSupportError::InvalidSource)?;
        self.charge(amount)
            .map_err(chart_axis_support::map_chart_title_error)
    }

    fn charge_input(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        self.charge(amount)
            .map_err(chart_axis_support::map_chart_title_error)
    }

    fn charge_wire_vector(&mut self, payload: usize) -> Result<(), AxisSupportError> {
        let amount = payload
            .checked_add(
                payload
                    .checked_mul(size_of::<WireField>())
                    .ok_or(AxisSupportError::InvalidSource)?,
            )
            .ok_or(AxisSupportError::InvalidSource)?;
        self.charge(amount)
            .map_err(chart_axis_support::map_chart_title_error)
    }

    fn finish_wire_scan(&mut self, fields: usize) -> Result<(), AxisSupportError> {
        self.charge(fields)
            .map_err(chart_axis_support::map_chart_title_error)
    }

    fn charge_reference_vector(&mut self, capacity: usize) -> Result<(), AxisSupportError> {
        let bytes = capacity
            .checked_mul(size_of::<u64>())
            .ok_or(AxisSupportError::InvalidSource)?;
        self.charge(bytes.max(1))
            .map_err(chart_axis_support::map_chart_title_error)
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        self.charge(amount)
            .map_err(chart_axis_support::map_chart_title_error)
    }

    fn charge_scan_pass(
        &mut self,
        package: &Package,
        retained_vectors: usize,
    ) -> Result<(), AxisSupportError> {
        let amount = package
            .state
            .source
            .components()
            .len()
            .checked_add(package.state.total_objects)
            .and_then(|value| value.checked_add(retained_vectors))
            .ok_or(AxisSupportError::InvalidSource)?;
        self.charge(amount)
            .map_err(chart_axis_support::map_chart_title_error)
    }

    fn charge_locality_scan(&mut self, package: &Package) -> Result<(), AxisSupportError> {
        let amount = package
            .state
            .source
            .components()
            .len()
            .checked_add(package.state.total_objects)
            .and_then(|value| value.checked_mul(2))
            .ok_or(AxisSupportError::InvalidSource)?;
        self.charge(amount)
            .map_err(chart_axis_support::map_chart_title_error)
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        self.charge(amount)
            .map_err(chart_axis_support::map_chart_title_error)
    }

    fn metadata_options(
        &self,
        _package: &Package,
    ) -> Result<litchi_iwa_protos::package_metadata_codec::RewriteOptions, AxisSupportError> {
        Err(AxisSupportError::InvalidSource)
    }

    fn charge_metadata_report(
        &mut self,
        report: litchi_iwa_protos::package_metadata_codec::RewriteReport,
    ) -> Result<(), AxisSupportError> {
        let amount = report
            .input_bytes()
            .checked_add(report.output_bytes())
            .and_then(|value| value.checked_add(report.fields()))
            .and_then(|value| value.checked_add(report.work_bytes()))
            .and_then(|value| value.checked_add(report.components_scanned()))
            .and_then(|value| value.checked_add(report.references_scanned()))
            .ok_or(AxisSupportError::InvalidSource)?;
        self.charge(amount)
            .map_err(chart_axis_support::map_chart_title_error)
    }
}

#[derive(Debug, Default)]
struct ChartGraphOwners {
    non_style: HashMap<u64, usize>,
    title_standin: HashMap<u64, usize>,
    non_style_metadata: HashMap<u64, MetadataReferenceCount>,
}

#[derive(Debug, Default, Clone, Copy)]
struct MetadataReferenceCount {
    total: usize,
    allowed_chart_message: usize,
    allowed_chart_registry: usize,
}

impl MetadataReferenceCount {
    fn is_fully_allowed(self) -> bool {
        self.allowed_chart_message <= self.total
            && self.allowed_chart_registry <= self.total - self.allowed_chart_message
            && self.total - self.allowed_chart_message - self.allowed_chart_registry == 0
    }
}

fn scan_chart_graph_owners(package: &Package) -> Result<ChartGraphOwners, ChartTitleError> {
    let mut budget = ChartGraphScanBudget::new(package)?;
    let mut owners = ChartGraphOwners::default();
    for component in package.state.source.components().iter() {
        budget.charge(1)?;
        for object in &component.archive().objects {
            budget.charge(1)?;
            for message in &object.messages {
                budget.charge(1)?;
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let fields = budget.parse(message.data.as_slice())?;
                let drawable_payload = unique_length_delimited_field(
                    &fields,
                    message.data.as_slice(),
                    DRAWABLE_SUPER_FIELD,
                )?
                .ok_or(ChartTitleError::InvalidSource)?;
                let drawable_fields = budget.parse(drawable_payload)?;
                let title_identifier = required_reference_field(
                    &drawable_fields,
                    drawable_payload,
                    DRAWABLE_TITLE_FIELD,
                    budget.limits,
                )?;
                let Some(chart_payload) = unique_length_delimited_field(
                    &fields,
                    message.data.as_slice(),
                    CHART_EXTENSION_FIELD,
                )?
                else {
                    return Err(ChartTitleError::InvalidSource);
                };
                let chart_fields = budget.parse(chart_payload)?;
                let non_style_identifier = required_reference_field(
                    &chart_fields,
                    chart_payload,
                    CHART_NON_STYLE_FIELD,
                    budget.limits,
                )?;
                increment_graph_owner(&mut owners.non_style, non_style_identifier)?;
                increment_graph_owner(&mut owners.title_standin, title_identifier)?;
            }
        }
    }
    scan_non_style_metadata_owners(package, &mut owners, &mut budget)?;
    Ok(owners)
}

// A changed title needs a complete inbound ownership proof. Keep this
// stricter metadata policy local to publication: other chart owners and
// read/no-op paths preserve opaque headers under their own contracts.
fn validate_chart_title_reference_metadata(
    package: &Package,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), ChartTitleError> {
    for component in package.state.source.components().iter() {
        budget.charge(1)?;
        for object in &component.archive().objects {
            chart_axis_support::inspect_archive_references_strict(package, object, budget)
                .map_err(map_axis_support_error)?;
        }
    }
    Ok(())
}

fn scan_non_style_metadata_owners(
    package: &Package,
    owners: &mut ChartGraphOwners,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), ChartTitleError> {
    for component in package.state.source.components().iter() {
        budget.charge(1)?;
        for object in &component.archive().objects {
            budget.charge(1)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                budget.charge(1)?;
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(ChartTitleError::InvalidSource)?;
                let mut metadata_slots = info
                    .object_references
                    .len()
                    .checked_add(info.data_references.len())
                    .and_then(|value| value.checked_add(info.field_infos.len()))
                    .ok_or(ChartTitleError::InvalidSource)?;
                for field in &info.field_infos {
                    metadata_slots = metadata_slots
                        .checked_add(field.object_references.len())
                        .and_then(|value| value.checked_add(field.data_references.len()))
                        .ok_or(ChartTitleError::InvalidSource)?;
                }
                budget.charge(
                    metadata_slots
                        .checked_mul(2)
                        .ok_or(ChartTitleError::InvalidSource)?,
                )?;
                let expected_chart_non_style = if message.type_ == CHART_MESSAGE_TYPE {
                    let fields = budget.parse(message.data.as_slice())?;
                    let chart_payload = unique_length_delimited_field(
                        &fields,
                        message.data.as_slice(),
                        CHART_EXTENSION_FIELD,
                    )?
                    .ok_or(ChartTitleError::InvalidSource)?;
                    Some(required_reference_field(
                        &budget.parse(chart_payload)?,
                        chart_payload,
                        CHART_NON_STYLE_FIELD,
                        budget.limits,
                    )?)
                } else {
                    None
                };
                let mut expected_aggregate_matches = 0usize;
                for identifier in &info.object_references {
                    if owners.non_style.contains_key(identifier)
                        && expected_chart_non_style == Some(*identifier)
                    {
                        expected_aggregate_matches = expected_aggregate_matches
                            .checked_add(1)
                            .ok_or(ChartTitleError::InvalidSource)?;
                    }
                }
                let mut data_matches = false;
                for identifier in &info.data_references {
                    data_matches |= owners.non_style.contains_key(identifier);
                }
                let mut expected_field_matches = 0usize;
                let mut foreign_field_matches = false;
                for field in &info.field_infos {
                    for identifier in &field.object_references {
                        if !owners.non_style.contains_key(identifier) {
                            continue;
                        }
                        if expected_chart_non_style == Some(*identifier) {
                            expected_field_matches = expected_field_matches
                                .checked_add(1)
                                .ok_or(ChartTitleError::InvalidSource)?;
                        } else {
                            foreign_field_matches = true;
                        }
                    }
                    for identifier in &field.data_references {
                        if owners.non_style.contains_key(identifier) {
                            foreign_field_matches = true;
                        }
                    }
                }
                let allowed_chart_message = expected_chart_non_style.is_some()
                    && !data_matches
                    && !foreign_field_matches
                    && expected_aggregate_matches == 1
                    && expected_field_matches <= 1;
                for identifier in &info.object_references {
                    if owners.non_style.contains_key(identifier) {
                        budget.charge(1)?;
                        let allowed_registry = is_chart_stylesheet_registration(
                            package,
                            object,
                            info,
                            message_index,
                            *identifier,
                            budget.limits,
                            budget,
                        )?;
                        record_non_style_metadata_reference(
                            &mut owners.non_style_metadata,
                            *identifier,
                            allowed_chart_message && expected_chart_non_style == Some(*identifier),
                            allowed_registry,
                        )?;
                    }
                }
                for identifier in &info.data_references {
                    if owners.non_style.contains_key(identifier) {
                        budget.charge(1)?;
                        record_non_style_metadata_reference(
                            &mut owners.non_style_metadata,
                            *identifier,
                            false,
                            false,
                        )?;
                    }
                }
                for field in &info.field_infos {
                    for identifier in &field.object_references {
                        if owners.non_style.contains_key(identifier) {
                            budget.charge(1)?;
                            let allowed_chart_field = allowed_chart_message
                                && expected_chart_non_style == Some(*identifier)
                                && expected_field_matches == 1;
                            record_non_style_metadata_reference(
                                &mut owners.non_style_metadata,
                                *identifier,
                                allowed_chart_field,
                                false,
                            )?;
                        }
                    }
                    for identifier in &field.data_references {
                        if owners.non_style.contains_key(identifier) {
                            budget.charge(1)?;
                            record_non_style_metadata_reference(
                                &mut owners.non_style_metadata,
                                *identifier,
                                false,
                                false,
                            )?;
                        }
                    }
                }
            }
        }
    }
    Ok(())
}

/// Recognize the one inbound edge created by the chart-style registration
/// lifecycle. A source-built chart registers its non-style object in the
/// document stylesheet, so the stylesheet aggregate metadata is a legitimate
/// owner in addition to the chart's own aggregate edge. The wire membership,
/// metadata shape, and cardinality are all checked here; a message merely
/// having type 401 is not sufficient authority.
fn is_chart_stylesheet_registration(
    package: &Package,
    object: &ArchiveObject,
    info: &litchi_iwa_core::MessageInfo,
    message_index: usize,
    identifier: u64,
    limits: WireLimits,
    budget: &mut ChartGraphScanBudget,
) -> Result<bool, ChartTitleError> {
    let Some(object_identifier) = object.archive_info.identifier else {
        return Ok(false);
    };
    let occurrence = ArchiveReferenceOccurrence {
        object_identifier,
        message_index,
        scope: ArchiveReferenceScope::Message,
        kind: ArchiveReferenceKind::Object,
        referenced_identifier: identifier,
    };
    if chart_axis_support::stylesheet_registration_component(package, occurrence).is_none() {
        return Ok(false);
    }
    let mut metadata_scan = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .and_then(|value| value.checked_add(info.field_infos.len()))
        .ok_or(ChartTitleError::InvalidSource)?;
    for field in &info.field_infos {
        metadata_scan = metadata_scan
            .checked_add(field.object_references.len())
            .and_then(|value| value.checked_add(field.data_references.len()))
            .ok_or(ChartTitleError::InvalidSource)?;
    }
    budget.charge(
        metadata_scan
            .checked_mul(2)
            .ok_or(ChartTitleError::InvalidSource)?,
    )?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
        || info.data_references.contains(&identifier)
        || info.field_infos.iter().any(|field| {
            field.object_references.contains(&identifier)
                || field.data_references.contains(&identifier)
        })
    {
        return Ok(false);
    }
    let message = object
        .messages
        .get(message_index)
        .ok_or(ChartTitleError::InvalidSource)?;
    let fields = budget.parse(message.data.as_slice())?;
    let mut matches = 0usize;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number() == STYLESHEET_STYLES_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(ChartTitleError::InvalidSource);
        }
        field
            .validate_canonical_framing(message.data.as_slice())
            .map_err(map_wire_error)?;
        let payload = field
            .payload(message.data.as_slice())
            .map_err(map_wire_error)?;
        if chart_axis_support::validate_reference_payload(payload, limits, budget)
            .map_err(map_axis_support_error)?
            == identifier
        {
            matches = matches
                .checked_add(1)
                .ok_or(ChartTitleError::InvalidSource)?;
        }
    }
    Ok(matches == 1)
}

fn is_chart_stylesheet_registration_with_budget(
    package: &Package,
    object: &ArchiveObject,
    info: &litchi_iwa_core::MessageInfo,
    message_index: usize,
    identifier: u64,
    limits: WireLimits,
    budget: &mut dyn AxisSupportBudget,
) -> Result<bool, AxisSupportError> {
    let Some(object_identifier) = object.archive_info.identifier else {
        return Ok(false);
    };
    let occurrence = ArchiveReferenceOccurrence {
        object_identifier,
        message_index,
        scope: ArchiveReferenceScope::Message,
        kind: ArchiveReferenceKind::Object,
        referenced_identifier: identifier,
    };
    if chart_axis_support::stylesheet_registration_component(package, occurrence).is_none() {
        return Ok(false);
    }
    let mut metadata_scan = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .and_then(|value| value.checked_add(info.field_infos.len()))
        .ok_or(AxisSupportError::InvalidSource)?;
    for field in &info.field_infos {
        metadata_scan = metadata_scan
            .checked_add(field.object_references.len())
            .and_then(|value| value.checked_add(field.data_references.len()))
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    budget.charge_work(
        metadata_scan
            .checked_mul(2)
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
        || info.data_references.contains(&identifier)
        || info.field_infos.iter().any(|field| {
            field.object_references.contains(&identifier)
                || field.data_references.contains(&identifier)
        })
    {
        return Ok(false);
    }
    let message = object
        .messages
        .get(message_index)
        .ok_or(AxisSupportError::InvalidSource)?;
    let fields = chart_axis_support::accounted_wire_fields(&message.data, limits, budget)?;
    budget.charge_work(fields.len())?;
    let mut matches = 0usize;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number() == STYLESHEET_STYLES_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(AxisSupportError::InvalidSource);
        }
        field
            .validate_canonical_framing(&message.data)
            .map_err(|_error| AxisSupportError::InvalidSource)?;
        let payload = field
            .payload(&message.data)
            .map_err(|_error| AxisSupportError::InvalidSource)?;
        if chart_axis_support::validate_reference_payload(payload, limits, budget)? == identifier {
            matches = matches
                .checked_add(1)
                .ok_or(AxisSupportError::InvalidSource)?;
        }
    }
    Ok(matches == 1)
}

fn record_non_style_metadata_reference(
    owners: &mut HashMap<u64, MetadataReferenceCount>,
    identifier: u64,
    allowed_chart_message: bool,
    allowed_chart_registry: bool,
) -> Result<(), ChartTitleError> {
    if !owners.contains_key(&identifier) {
        owners
            .try_reserve(1)
            .map_err(|_error| ChartTitleError::Allocation { amount: 1 })?;
        owners.insert(identifier, MetadataReferenceCount::default());
    }
    let entry = owners
        .get_mut(&identifier)
        .ok_or(ChartTitleError::InvalidSource)?;
    entry.total = entry
        .total
        .checked_add(1)
        .ok_or(ChartTitleError::InvalidSource)?;
    if allowed_chart_message {
        entry.allowed_chart_message = entry
            .allowed_chart_message
            .checked_add(1)
            .ok_or(ChartTitleError::InvalidSource)?;
    }
    if allowed_chart_registry {
        entry.allowed_chart_registry = entry
            .allowed_chart_registry
            .checked_add(1)
            .ok_or(ChartTitleError::InvalidSource)?;
    }
    Ok(())
}

fn scan_chart_graph_owners_with_budget(
    package: &Package,
    budget: &mut dyn AxisSupportBudget,
) -> Result<ChartGraphOwners, AxisSupportError> {
    let limits = package
        .wire_limits()
        .map_err(map_wire_error)
        .map_err(chart_graph_error)?;
    let mut owners = ChartGraphOwners::default();
    for component in package.state.source.components().iter() {
        budget.charge_work(1)?;
        for object in &component.archive().objects {
            budget.charge_work(1)?;
            for message in &object.messages {
                budget.charge_work(1)?;
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let fields = chart_axis_support::accounted_wire_fields(
                    message.data.as_slice(),
                    limits,
                    budget,
                )?;
                let drawable_payload = chart_axis_support::unique_length_delimited_field(
                    &fields,
                    message.data.as_slice(),
                    DRAWABLE_SUPER_FIELD,
                    budget,
                )?
                .ok_or(AxisSupportError::InvalidSource)?;
                let title_identifier = chart_axis_support::required_reference(
                    drawable_payload,
                    DRAWABLE_TITLE_FIELD,
                    limits,
                    budget,
                )?;
                let chart_payload = chart_axis_support::unique_length_delimited_field(
                    &fields,
                    message.data.as_slice(),
                    CHART_EXTENSION_FIELD,
                    budget,
                )?
                .ok_or(AxisSupportError::InvalidSource)?;
                let non_style_identifier = chart_axis_support::required_reference(
                    chart_payload,
                    CHART_NON_STYLE_FIELD,
                    limits,
                    budget,
                )?;
                increment_graph_owner_with_budget(
                    &mut owners.non_style,
                    non_style_identifier,
                    budget,
                )?;
                increment_graph_owner_with_budget(
                    &mut owners.title_standin,
                    title_identifier,
                    budget,
                )?;
            }
        }
    }
    scan_non_style_metadata_owners_with_budget(package, &mut owners, budget)?;
    Ok(owners)
}

fn scan_non_style_metadata_owners_with_budget(
    package: &Package,
    owners: &mut ChartGraphOwners,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    let limits = package
        .wire_limits()
        .map_err(map_wire_error)
        .map_err(chart_graph_error)?;
    for component in package.state.source.components().iter() {
        budget.charge_work(1)?;
        for object in &component.archive().objects {
            budget.charge_work(1)?;
            for (message_index, message) in object.messages.iter().enumerate() {
                budget.charge_work(1)?;
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(AxisSupportError::InvalidSource)?;
                let mut metadata_slots = info
                    .object_references
                    .len()
                    .checked_add(info.data_references.len())
                    .and_then(|value| value.checked_add(info.field_infos.len()))
                    .ok_or(AxisSupportError::InvalidSource)?;
                for field in &info.field_infos {
                    metadata_slots = metadata_slots
                        .checked_add(field.object_references.len())
                        .and_then(|value| value.checked_add(field.data_references.len()))
                        .ok_or(AxisSupportError::InvalidSource)?;
                }
                budget.charge_work(
                    metadata_slots
                        .checked_mul(2)
                        .ok_or(AxisSupportError::InvalidSource)?,
                )?;
                let expected_chart_non_style = if message.type_ == CHART_MESSAGE_TYPE {
                    let fields = chart_axis_support::accounted_wire_fields(
                        message.data.as_slice(),
                        limits,
                        budget,
                    )?;
                    let chart_payload = chart_axis_support::unique_length_delimited_field(
                        &fields,
                        message.data.as_slice(),
                        CHART_EXTENSION_FIELD,
                        budget,
                    )?
                    .ok_or(AxisSupportError::InvalidSource)?;
                    Some(chart_axis_support::required_reference(
                        chart_payload,
                        CHART_NON_STYLE_FIELD,
                        limits,
                        budget,
                    )?)
                } else {
                    None
                };
                let mut expected_aggregate_matches = 0usize;
                for identifier in &info.object_references {
                    if owners.non_style.contains_key(identifier)
                        && expected_chart_non_style == Some(*identifier)
                    {
                        expected_aggregate_matches = expected_aggregate_matches
                            .checked_add(1)
                            .ok_or(AxisSupportError::InvalidSource)?;
                    }
                }
                let mut data_matches = false;
                for identifier in &info.data_references {
                    data_matches |= owners.non_style.contains_key(identifier);
                }
                let mut expected_field_matches = 0usize;
                let mut foreign_field_matches = false;
                for field in &info.field_infos {
                    for identifier in &field.object_references {
                        if !owners.non_style.contains_key(identifier) {
                            continue;
                        }
                        if expected_chart_non_style == Some(*identifier) {
                            expected_field_matches = expected_field_matches
                                .checked_add(1)
                                .ok_or(AxisSupportError::InvalidSource)?;
                        } else {
                            foreign_field_matches = true;
                        }
                    }
                    for identifier in &field.data_references {
                        if owners.non_style.contains_key(identifier) {
                            foreign_field_matches = true;
                        }
                    }
                }
                let allowed_chart_message = expected_chart_non_style.is_some()
                    && !data_matches
                    && !foreign_field_matches
                    && expected_aggregate_matches == 1
                    && expected_field_matches <= 1;
                for identifier in &info.object_references {
                    if owners.non_style.contains_key(identifier) {
                        let allowed_registry = is_chart_stylesheet_registration_with_budget(
                            package,
                            object,
                            info,
                            message_index,
                            *identifier,
                            limits,
                            budget,
                        )?;
                        record_non_style_metadata_reference_with_budget(
                            &mut owners.non_style_metadata,
                            *identifier,
                            allowed_chart_message && expected_chart_non_style == Some(*identifier),
                            allowed_registry,
                            budget,
                        )?;
                    }
                }
                for identifier in &info.data_references {
                    if owners.non_style.contains_key(identifier) {
                        record_non_style_metadata_reference_with_budget(
                            &mut owners.non_style_metadata,
                            *identifier,
                            false,
                            false,
                            budget,
                        )?;
                    }
                }
                for field in &info.field_infos {
                    for identifier in &field.object_references {
                        if owners.non_style.contains_key(identifier) {
                            let allowed_chart_field = allowed_chart_message
                                && expected_chart_non_style == Some(*identifier)
                                && expected_field_matches == 1;
                            record_non_style_metadata_reference_with_budget(
                                &mut owners.non_style_metadata,
                                *identifier,
                                allowed_chart_field,
                                false,
                                budget,
                            )?;
                        }
                    }
                    for identifier in &field.data_references {
                        if owners.non_style.contains_key(identifier) {
                            record_non_style_metadata_reference_with_budget(
                                &mut owners.non_style_metadata,
                                *identifier,
                                false,
                                false,
                                budget,
                            )?;
                        }
                    }
                }
            }
        }
    }
    Ok(())
}

fn record_non_style_metadata_reference_with_budget(
    owners: &mut HashMap<u64, MetadataReferenceCount>,
    identifier: u64,
    allowed_chart_message: bool,
    allowed_chart_registry: bool,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    budget.charge_work(1)?;
    if !owners.contains_key(&identifier) {
        charge_selector_hash_map(1, budget)?;
        owners
            .try_reserve(1)
            .map_err(|_error| AxisSupportError::Allocation { amount: 1 })?;
        owners.insert(identifier, MetadataReferenceCount::default());
    }
    let entry = owners
        .get_mut(&identifier)
        .ok_or(AxisSupportError::InvalidSource)?;
    entry.total = entry
        .total
        .checked_add(1)
        .ok_or(AxisSupportError::InvalidSource)?;
    if allowed_chart_message {
        entry.allowed_chart_message = entry
            .allowed_chart_message
            .checked_add(1)
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    if allowed_chart_registry {
        entry.allowed_chart_registry = entry
            .allowed_chart_registry
            .checked_add(1)
            .ok_or(AxisSupportError::InvalidSource)?;
    }
    Ok(())
}

fn increment_graph_owner(
    owners: &mut HashMap<u64, usize>,
    identifier: u64,
) -> Result<(), ChartTitleError> {
    if let Some(count) = owners.get_mut(&identifier) {
        *count = count.checked_add(1).ok_or(ChartTitleError::InvalidSource)?;
    } else {
        owners
            .try_reserve(1)
            .map_err(|_error| ChartTitleError::Allocation { amount: 1 })?;
        owners.insert(identifier, 1);
    }
    Ok(())
}

fn increment_graph_owner_with_budget(
    owners: &mut HashMap<u64, usize>,
    identifier: u64,
    budget: &mut dyn AxisSupportBudget,
) -> Result<(), AxisSupportError> {
    if let Some(count) = owners.get_mut(&identifier) {
        *count = count
            .checked_add(1)
            .ok_or(AxisSupportError::InvalidSource)?;
    } else {
        charge_selector_hash_map(1, budget)?;
        owners
            .try_reserve(1)
            .map_err(|_error| AxisSupportError::Allocation { amount: 1 })?;
        owners.insert(identifier, 1);
    }
    budget.charge_work(1)
}

struct ChartTitleRewrite {
    package: Package,
    before_message: Arc<[u8]>,
    selected_message: Vec<u8>,
    deleted_previews: usize,
}

fn verify_selected_chart_title_message(
    source: &Package,
    selection: &ChartSelection,
    expected: &[u8],
    budget: &mut ChartGraphScanBudget,
) -> Result<(), ChartTitleError> {
    let (_component_name, object) = source
        .object_with_component(selection.non_style_identifier)
        .ok_or(ChartTitleError::PatchConflict)?;
    let message = object
        .messages
        .get(selection.non_style_message_index)
        .filter(|message| message.type_ == CHART_NON_STYLE_MESSAGE_TYPE)
        .ok_or(ChartTitleError::PatchConflict)?;
    budget.charge(
        message
            .data
            .len()
            .checked_add(expected.len())
            .ok_or(ChartTitleError::InvalidSource)?,
    )?;
    if message.data.as_slice() != expected {
        return Err(ChartTitleError::PatchConflict);
    }
    Ok(())
}

fn clone_payload_with_budget(
    data: &[u8],
    budget: &mut ChartGraphScanBudget,
) -> Result<Vec<u8>, ChartTitleError> {
    budget.charge(data.len())?;
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(data.len())
        .map_err(|_error| ChartTitleError::Allocation { amount: data.len() })?;
    owned.extend_from_slice(data);
    Ok(owned)
}

fn rewrite_chart_title(
    source: &Package,
    selection: &ChartSelection,
    after: Option<&str>,
    budget: &mut ChartGraphScanBudget,
) -> Result<ChartTitleRewrite, ChartTitleError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.non_style_component_name)
        .ok_or(ChartTitleError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(ChartTitleError::InvalidSource);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = source
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let parsed_component = source
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selection.non_style_component_name)
        .ok_or(ChartTitleError::InvalidSource)?;
    let decompressed_bound = parsed_component
        .archive()
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.charge_native_decompress_parse(
        entry.data().len(),
        decompressed_bound,
        parsed_component.archive().objects.len(),
    )?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    let (message_index, message_data) = {
        let object = archive
            .object(selection.non_style_identifier)
            .ok_or(ChartTitleError::InvalidSource)?;
        let (message_index, message) =
            exactly_one_message_with_index(object, CHART_NON_STYLE_MESSAGE_TYPE)?;
        if message_index != selection.non_style_message_index {
            return Err(ChartTitleError::InvalidSource);
        }
        budget.charge(message.data.len())?;
        let mut message_data = Vec::new();
        message_data
            .try_reserve_exact(message.data.len())
            .map_err(|_error| ChartTitleError::Allocation {
                amount: message.data.len(),
            })?;
        message_data.extend_from_slice(&message.data);
        (message_index, message_data)
    };
    let limits = source.wire_limits().map_err(map_wire_error)?;
    let patched = patch_chart_title(message_data.as_slice(), after, limits, budget)?;
    let before_message: Arc<[u8]> = Arc::from(message_data);
    let selected_message = clone_payload_with_budget(&patched, budget)?;
    archive
        .object_mut(selection.non_style_identifier)
        .ok_or(ChartTitleError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: CHART_NON_STYLE_MESSAGE_TYPE,
                data: patched,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    budget.charge_archive_snappy_plan(&archive, archive_limits, snappy_limits)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_error| ChartTitleError::InvalidSource)?;
    let edit = EntryEdit::new(
        selection.non_style_component_name.as_str(),
        compressed.as_slice(),
    );
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(
            std::slice::from_ref(&edit),
            previews.names(),
            source.state.options.archive(),
        )
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.charge_candidate_reopen(
        source,
        requirements.output_bytes(),
        Some((selection.non_style_component_name.as_str(), &archive)),
    )?;
    budget.charge(
        requirements
            .output_bytes()
            .checked_add(requirements.scratch_bytes())
            .and_then(|value| value.checked_add(requirements.offset_count()))
            .ok_or(ChartTitleError::InvalidSource)?,
    )?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let package = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok(ChartTitleRewrite {
        package,
        before_message,
        selected_message,
        deleted_previews: previews.len(),
    })
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), ChartTitleError> {
    for object in &archive.objects {
        let offset = usize::try_from(object.header_offset)
            .map_err(|_error| ChartTitleError::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(ChartTitleError::InvalidSource)?;
        let (header_bytes, prefix_bytes) =
            decode_varint_from_bytes(remaining).map_err(|_error| ChartTitleError::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(ChartTitleError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(
                u64::try_from(prefix_bytes).map_err(|_error| ChartTitleError::InvalidSource)?,
            )
            .ok_or(ChartTitleError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(ChartTitleError::InvalidSource)?
                != object.data_offset
        {
            return Err(ChartTitleError::InvalidSource);
        }
    }
    Ok(())
}

fn read_chart_title(data: &[u8], limits: WireLimits) -> Result<Option<String>, ChartTitleError> {
    let fields = parse_wire_fields_with_limits(data, limits).map_err(map_wire_error)?;
    let mut extension_field = None;
    let mut extension_count = 0usize;
    for field in &fields {
        if field.number() == GENERATED_CHART_NON_STYLE_EXTENSION_FIELD {
            extension_count = extension_count
                .checked_add(1)
                .ok_or(ChartTitleError::InvalidSource)?;
            extension_field = Some(*field);
        }
    }
    let Some(field) = extension_field else {
        return Ok(None);
    };
    if extension_count != 1 || field.wire_type() != 2 {
        return Err(ChartTitleError::InvalidSource);
    }
    field
        .validate_canonical_framing(data)
        .map_err(map_wire_error)?;
    let extension = field.payload(data).map_err(map_wire_error)?;
    decode_visible_chart_title(extension, chart_title_decode_options(limits)?)
        .map_err(map_chart_title_codec_error)?
        .map(copy_title)
        .transpose()
}

fn read_chart_title_with_budget(
    data: &[u8],
    limits: WireLimits,
    budget: &mut dyn AxisSupportBudget,
) -> Result<Option<String>, AxisSupportError> {
    let fields = chart_axis_support::accounted_wire_fields(data, limits, budget)?;
    let mut extension_field = None;
    let mut extension_count = 0usize;
    budget.charge_work(fields.len())?;
    for field in &fields {
        if field.number() == GENERATED_CHART_NON_STYLE_EXTENSION_FIELD {
            extension_count = extension_count
                .checked_add(1)
                .ok_or(AxisSupportError::InvalidSource)?;
            extension_field = Some(*field);
        }
    }
    let Some(field) = extension_field else {
        return Ok(None);
    };
    if extension_count != 1 || field.wire_type() != 2 {
        return Err(AxisSupportError::InvalidSource);
    }
    field
        .validate_canonical_framing(data)
        .map_err(map_wire_error)
        .map_err(chart_graph_error)?;
    let extension = field
        .payload(data)
        .map_err(map_wire_error)
        .map_err(chart_graph_error)?;
    // `decode_visible_chart_title` returns a borrowed view, but its strict
    // decoder still traverses every encoded field. Charge the codec's
    // source-sized worst case before invoking it; the title's owned copy is
    // charged separately before `String::try_reserve_exact` below.
    let worst_fields = extension
        .len()
        .checked_mul(4)
        .ok_or(AxisSupportError::InvalidSource)?
        .max(1);
    let worst_work = extension
        .len()
        .checked_mul(8)
        .ok_or(AxisSupportError::InvalidSource)?
        .max(1);
    budget.charge_input(extension.len())?;
    budget.finish_wire_scan(worst_fields)?;
    budget.charge_work(
        worst_work
            .checked_sub(worst_fields)
            .ok_or(AxisSupportError::InvalidSource)?,
    )?;
    let title = decode_visible_chart_title(
        extension,
        chart_title_decode_options(limits).map_err(chart_graph_error)?,
    )
    .map_err(map_chart_title_codec_error)
    .map_err(chart_graph_error)?;
    title
        .map(|title| copy_title_with_budget(title, budget))
        .transpose()
}

fn verify_chart_candidate(
    source: &Package,
    candidate: &Package,
    slide_position: Position,
    chart_position: Position,
    expected_title: Option<&str>,
    expected_selected_message: &[u8],
    target_requires_invalidated_previews: bool,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), ChartTitleError> {
    if source.state.total_objects != candidate.state.total_objects {
        return Err(ChartTitleError::Verification);
    }
    budget.charge(2)?;
    let source_graphs = chart_graphs_with_budget(source, slide_position, true, budget)
        .map_err(map_axis_support_error)?;
    let candidate_graphs = chart_graphs_with_budget(candidate, slide_position, true, budget)
        .map_err(map_axis_support_error)?;
    if source_graphs.len() != candidate_graphs.len() {
        return Err(ChartTitleError::Verification);
    }
    for (index, (source_graph, candidate_graph)) in source_graphs
        .iter()
        .zip(candidate_graphs.iter())
        .enumerate()
    {
        if source_graph.slide_identifier != candidate_graph.slide_identifier
            || source_graph.chart_identifier != candidate_graph.chart_identifier
            || source_graph.non_style_identifier != candidate_graph.non_style_identifier
            || source_graph.title_identifier != candidate_graph.title_identifier
            || source_graph.slide_component_name != candidate_graph.slide_component_name
            || source_graph.non_style_component_name != candidate_graph.non_style_component_name
        {
            return Err(ChartTitleError::Verification);
        }
        let expected = if index == chart_position.get() {
            expected_title
        } else {
            source_graph.title.as_deref()
        };
        if candidate_graph.title.as_deref() != expected {
            return Err(ChartTitleError::Verification);
        }
    }
    let source_show = source.show().map_err(map_read_error)?;
    let candidate_show = candidate.show().map_err(map_read_error)?;
    if source_show != candidate_show {
        return Err(ChartTitleError::Verification);
    }
    let selected = source_graphs
        .get(chart_position.get())
        .ok_or(ChartTitleError::Verification)?;
    chart_axis_support::verify_package_locality_for_component(
        source,
        candidate,
        selected.non_style_component_name.as_str(),
        selected.non_style_identifier,
        selected.non_style_message_index,
        target_requires_invalidated_previews,
        Some(expected_selected_message),
        budget,
    )
    .map_err(map_axis_support_error)?;
    Ok(())
}

fn patch_chart_title(
    data: &[u8],
    title: Option<&str>,
    limits: WireLimits,
    budget: &mut ChartGraphScanBudget,
) -> Result<Vec<u8>, ChartTitleError> {
    let fields = budget.parse(data)?;
    let extension_field = fields
        .iter()
        .copied()
        .find(|field| field.number() == GENERATED_CHART_NON_STYLE_EXTENSION_FIELD);
    if fields
        .iter()
        .filter(|field| field.number() == GENERATED_CHART_NON_STYLE_EXTENSION_FIELD)
        .count()
        > 1
    {
        return Err(ChartTitleError::InvalidSource);
    }
    if let Some(field) = extension_field {
        if field.wire_type() != 2 {
            return Err(ChartTitleError::InvalidSource);
        }
        field
            .validate_canonical_framing(data)
            .map_err(map_wire_error)?;
    }
    let Some(extension) = extension_field
        .map(|field| field.payload(data).map_err(map_wire_error))
        .transpose()?
    else {
        let Some(title) = title else {
            return clone_title_source(data, limits, budget);
        };
        precharge_chart_title_rewrite(
            &[],
            ChartTitleWrite::new(Some(true), Some(title)),
            limits,
            budget,
        )?;
        let extension = rewrite_chart_title_payload(
            &[],
            ChartTitleWrite::new(Some(true), Some(title)),
            chart_title_decode_options(limits)?,
        )
        .map_err(map_chart_title_codec_error)?;
        return rewrite_chart_non_style_extension(data, None, &extension, limits, budget);
    };

    if title.is_none() {
        budget.charge(
            extension
                .len()
                .checked_mul(4)
                .ok_or(ChartTitleError::InvalidSource)?,
        )?;
        if decode_visible_chart_title(extension, chart_title_decode_options(limits)?)
            .map_err(map_chart_title_codec_error)?
            .is_none()
        {
            return clone_title_source(data, limits, budget);
        }
    }
    let write = ChartTitleWrite::new(Some(title.is_some()), title);
    precharge_chart_title_rewrite(extension, write, limits, budget)?;
    let extension =
        rewrite_chart_title_payload(extension, write, chart_title_decode_options(limits)?)
            .map_err(map_chart_title_codec_error)?;
    rewrite_chart_non_style_extension(data, extension_field, &extension, limits, budget)
}

fn precharge_chart_title_rewrite(
    source: &[u8],
    write: ChartTitleWrite<'_>,
    limits: WireLimits,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), ChartTitleError> {
    let fields = budget.parse(source)?;
    let mut output_length = 0usize;
    let mut saw_visible = false;
    let mut saw_title = false;
    for field in fields.iter().copied() {
        field
            .validate_canonical_framing(source)
            .map_err(map_wire_error)?;
        let replacement_length = match field.number() {
            21 => {
                if field.wire_type() != 0 || saw_visible {
                    return Err(ChartTitleError::InvalidSource);
                }
                saw_visible = true;
                write.title_visible().map_or(Ok(0), |value| {
                    encoded_len(21_u64 << 3)
                        .checked_add(encoded_len(u64::from(value)))
                        .ok_or(ChartTitleError::InvalidSource)
                })?
            },
            23 => {
                if field.wire_type() != 2 || saw_title {
                    return Err(ChartTitleError::InvalidSource);
                }
                saw_title = true;
                write.title().map_or(Ok(0), |value| {
                    let length = u64::try_from(value.len())
                        .map_err(|_error| ChartTitleError::InvalidSource)?;
                    encoded_len((23_u64 << 3) | 2)
                        .checked_add(encoded_len(length))
                        .and_then(|length| length.checked_add(value.len()))
                        .ok_or(ChartTitleError::InvalidSource)
                })?
            },
            _ => field.end() - field.start(),
        };
        output_length = output_length
            .checked_add(replacement_length)
            .ok_or(ChartTitleError::InvalidSource)?;
    }
    if !saw_visible && let Some(value) = write.title_visible() {
        output_length = output_length
            .checked_add(encoded_len(21_u64 << 3))
            .and_then(|length| length.checked_add(encoded_len(u64::from(value))))
            .ok_or(ChartTitleError::InvalidSource)?;
    }
    if !saw_title && let Some(value) = write.title() {
        let title_length =
            u64::try_from(value.len()).map_err(|_error| ChartTitleError::InvalidSource)?;
        output_length = output_length
            .checked_add(encoded_len((23_u64 << 3) | 2))
            .and_then(|length| length.checked_add(encoded_len(title_length)))
            .and_then(|length| length.checked_add(value.len()))
            .ok_or(ChartTitleError::InvalidSource)?;
    }
    if output_length > limits.max_output_bytes() {
        return Err(ChartTitleError::LimitExceeded {
            kind: ChartTitleLimitKind::OutputBytes,
            observed: usize_to_u64(output_length),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    let codec_work = source
        .len()
        .checked_mul(10)
        .and_then(|value| {
            output_length
                .checked_mul(4)
                .and_then(|output| value.checked_add(output))
        })
        .ok_or(ChartTitleError::InvalidSource)?;
    budget.charge(codec_work)
}

fn rewrite_chart_non_style_extension(
    data: &[u8],
    extension_field: Option<WireField>,
    replacement: &[u8],
    limits: WireLimits,
    budget: &mut ChartGraphScanBudget,
) -> Result<Vec<u8>, ChartTitleError> {
    let replacement_length =
        u64::try_from(replacement.len()).map_err(|_error| ChartTitleError::InvalidSource)?;
    let key_length = extension_field.map_or_else(
        || encoded_len((u64::from(GENERATED_CHART_NON_STYLE_EXTENSION_FIELD) << 3) | 2),
        |field| field.key_end() - field.start(),
    );
    let replacement_field_length = key_length
        .checked_add(encoded_len(replacement_length))
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or(ChartTitleError::InvalidSource)?;
    let output_length = extension_field
        .map_or_else(
            || data.len().checked_add(replacement_field_length),
            |field| {
                data.len()
                    .checked_sub(field.end() - field.start())
                    .and_then(|length| length.checked_add(replacement_field_length))
            },
        )
        .ok_or(ChartTitleError::InvalidSource)?;
    if output_length > limits.max_output_bytes() {
        return Err(ChartTitleError::LimitExceeded {
            kind: ChartTitleLimitKind::OutputBytes,
            observed: usize_to_u64(output_length),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    budget.charge(output_length)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_error| ChartTitleError::Allocation {
            amount: output_length,
        })?;
    match extension_field {
        Some(field) => {
            output.extend_from_slice(&data[..field.start()]);
            output.extend_from_slice(&data[field.start()..field.key_end()]);
            encode_varint_into(&mut output, replacement_length);
            output.extend_from_slice(replacement);
            output.extend_from_slice(&data[field.end()..]);
        },
        None => {
            output.extend_from_slice(data);
            encode_varint_into(
                &mut output,
                (u64::from(GENERATED_CHART_NON_STYLE_EXTENSION_FIELD) << 3) | 2,
            );
            encode_varint_into(&mut output, replacement_length);
            output.extend_from_slice(replacement);
        },
    }
    debug_assert_eq!(output.len(), output_length);
    Ok(output)
}

fn clone_title_source(
    data: &[u8],
    limits: WireLimits,
    budget: &mut ChartGraphScanBudget,
) -> Result<Vec<u8>, ChartTitleError> {
    if data.len() > limits.max_output_bytes() {
        return Err(ChartTitleError::LimitExceeded {
            kind: ChartTitleLimitKind::OutputBytes,
            observed: usize_to_u64(data.len()),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    budget.charge(data.len())?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(data.len())
        .map_err(|_error| ChartTitleError::Allocation { amount: data.len() })?;
    output.extend_from_slice(data);
    Ok(output)
}

fn chart_title_decode_options(limits: WireLimits) -> Result<DecodeOptions, ChartTitleError> {
    let recursion_limit =
        u32::try_from(limits.max_nesting()).map_err(|_error| ChartTitleError::LimitExceeded {
            kind: ChartTitleLimitKind::WireNesting,
            observed: usize_to_u64(limits.max_nesting()),
            maximum: u64::from(u32::MAX),
        })?;
    Ok(DecodeOptions::new(
        limits.max_input_bytes().max(1),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion_limit,
    )
    .with_max_output_bytes(limits.max_output_bytes())
    .with_max_title_bytes(MAX_CHART_TITLE_BYTES.min(limits.max_output_bytes())))
}

fn copy_title(title: &str) -> Result<String, ChartTitleError> {
    if title.len() > MAX_CHART_TITLE_BYTES {
        return Err(ChartTitleError::LimitExceeded {
            kind: ChartTitleLimitKind::TitleBytes,
            observed: usize_to_u64(title.len()),
            maximum: usize_to_u64(MAX_CHART_TITLE_BYTES),
        });
    }
    let mut owned = String::new();
    owned
        .try_reserve_exact(title.len())
        .map_err(|_error| ChartTitleError::Allocation {
            amount: title.len(),
        })?;
    owned.push_str(title);
    Ok(owned)
}

fn map_chart_title_codec_error(error: ChartTitleDecodeError) -> ChartTitleError {
    if let Some(limit) = error.resource_limit() {
        return match limit {
            ChartTitleDecodeLimit::Bytes { observed, maximum } => ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::WireBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartTitleDecodeLimit::Fields { observed, maximum } => ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::WireFields,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartTitleDecodeLimit::Work { observed, maximum } => ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::WireWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartTitleDecodeLimit::Output { observed, maximum } => ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::OutputBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartTitleDecodeLimit::Title { observed, maximum } => ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::TitleBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartTitleDecodeLimit::Nesting { observed, maximum } => {
                ChartTitleError::LimitExceeded {
                    kind: ChartTitleLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => ChartTitleError::InvalidSource,
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            ChartTitleWireResourceLimit::Bytes { observed, maximum } => {
                ChartTitleError::LimitExceeded {
                    kind: ChartTitleLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartTitleWireResourceLimit::Nesting { observed, maximum } => {
                ChartTitleError::LimitExceeded {
                    kind: ChartTitleLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => ChartTitleError::InvalidSource,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return ChartTitleError::Allocation { amount };
    }
    ChartTitleError::InvalidSource
}

fn map_chart_selector_error(error: ChartSelectorError) -> ChartTitleError {
    match error {
        ChartSelectorError::EmptyName => ChartTitleError::EmptyChartName,
        ChartSelectorError::DuplicateChartTitle { .. } => ChartTitleError::AmbiguousSelector,
    }
}

fn map_slide_selector_error(error: SlideSelectorError) -> ChartTitleError {
    match error {
        SlideSelectorError::EmptySlideName => ChartTitleError::EmptySlideName,
        SlideSelectorError::DuplicateSlideName { .. } => ChartTitleError::AmbiguousSelector,
    }
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, ChartTitleError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(ChartTitleError::UnsupportedSource),
    }
}

fn map_read_error(error: ReadError) -> ChartTitleError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartTitleError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => ChartTitleLimitKind::Entries,
                SemanticLimitKind::Slides => ChartTitleLimitKind::Slides,
                SemanticLimitKind::References => ChartTitleLimitKind::References,
                SemanticLimitKind::TextStorages => ChartTitleLimitKind::TextStorages,
                SemanticLimitKind::TextFragments => ChartTitleLimitKind::TextFragments,
                SemanticLimitKind::TextBytes => ChartTitleLimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartTitleError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => ChartTitleLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => ChartTitleLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => ChartTitleLimitKind::WireNesting,
                super::PayloadLimitKind::Work => ChartTitleLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => ChartTitleError::Allocation { amount },
        ReadError::Archive(_)
        | ReadError::Detection(_)
        | ReadError::NotKeynote
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_)
        | ReadError::Io(_) => ChartTitleError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> ChartTitleError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => ChartTitleLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => ChartTitleLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => ChartTitleLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes => ChartTitleLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => ChartTitleLimitKind::TotalBytes,
                _ => ChartTitleLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            ChartTitleError::Allocation { amount }
        },
        _ => ChartTitleError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> ChartTitleError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => ChartTitleLimitKind::TotalBytes,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => ChartTitleLimitKind::Entries,
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    ChartTitleLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::HeaderFields => ChartTitleLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => ChartTitleLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyFrames => ChartTitleLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            ChartTitleError::Allocation { amount: requested }
        },
        _ => ChartTitleError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> ChartTitleError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ChartTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => ChartTitleLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => ChartTitleLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    ChartTitleLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => ChartTitleLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => ChartTitleLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            ChartTitleError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => ChartTitleError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
#[path = "slide_chart_title/verification_tests.rs"]
mod verification_tests;
