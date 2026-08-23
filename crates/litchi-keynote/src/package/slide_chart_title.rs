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
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_chart_title_codec::{
    ChartTitleWrite, DecodeError as ChartTitleDecodeError, DecodeLimit as ChartTitleDecodeLimit,
    DecodeOptions, WireResourceLimit as ChartTitleWireResourceLimit, decode_visible_chart_title,
    rewrite_chart_title as rewrite_chart_title_payload,
};
use thiserror::Error;

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
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_DRAWABLES_Z_ORDER_FIELD: u32 = 42;
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
    pub fn set(mut self, title: &str) -> Result<Self, ChartTitleError> {
        self.after = Some(copy_title(title)?);
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
        let source_selection = select_chart(
            self.source,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            true,
        )?;
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
                },
                diagnostics: ChartTitleDiagnostics::unchanged(),
            });
        }

        if !catalog.source_is_exact() {
            return Err(ChartTitleError::UnsupportedSource);
        }
        self.source.validate().map_err(map_read_error)?;
        let package = rewrite_chart_title(self.source, &source_selection, self.after.as_deref())?;
        package.validate().map_err(map_read_error)?;
        let target = physical_catalog(&package)?.shared_source();
        let candidate_selection = select_chart(
            &package,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            true,
        )?;
        if candidate_selection.chart_identifier != self.chart_identifier
            || candidate_selection.non_style_identifier != self.non_style_identifier
            || candidate_selection.slide_identifier != self.slide_identifier
            || candidate_selection.title != self.after
        {
            return Err(ChartTitleError::Verification);
        }
        verify_chart_candidate(
            self.source,
            &package,
            self.slide_position,
            self.chart_position,
            self.after.as_deref(),
        )?;
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
            },
            diagnostics: ChartTitleDiagnostics::published(),
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
        }
    }
}

/// Compact evidence describing one chart-title commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartTitleDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl ChartTitleDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
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
        let source_selection = select_chart(
            self,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            true,
        )?;
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
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let candidate_selection = select_chart(
            &candidate,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            true,
        )?;
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
        )?;
        Ok(ChartTitleCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ChartTitleDiagnostics::published(),
        })
    }
}

#[derive(Debug, Clone)]
struct ChartSelection {
    slide_position: Position,
    chart_position: Position,
    slide_identifier: u64,
    chart_identifier: u64,
    non_style_identifier: u64,
    title: Option<String>,
    slide_component_name: String,
}

#[derive(Debug, Clone)]
struct ChartGraph {
    slide_identifier: u64,
    chart_identifier: u64,
    non_style_identifier: u64,
    title: Option<String>,
    slide_component_name: String,
}

fn select_chart(
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
        title,
        slide_component_name: graph.slide_component_name.clone(),
    })
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
    // The mutation guard below needs to prove that each selected non-style
    // object has exactly one chart owner. Keep the package-wide index lazy so
    // a slide with no charts never pays for the ownership scan, then share it
    // across every chart in this graph build.
    let mut non_style_owners = None;
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
            &owned,
            &z_order,
            &chart_message.data,
            identifier,
        )?;
        if mutation_guards {
            if non_style_owners.is_none() {
                non_style_owners = Some(scan_non_style_owners(package)?);
            }
            let owners = non_style_owners
                .as_ref()
                .ok_or(ChartTitleError::InvalidSource)?;
            if owners.get(&graph.non_style_identifier).copied() != Some(1)
                || graph.chart_identifier == 0
            {
                return Err(ChartTitleError::InvalidSource);
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
    owned: &[u64],
    z_order: &[u64],
    chart_data: &[u8],
    chart_identifier: u64,
) -> Result<ChartGraph, ChartTitleError> {
    if owned
        .iter()
        .filter(|candidate| **candidate == chart_identifier)
        .count()
        != 1
        || z_order
            .iter()
            .filter(|candidate| **candidate == chart_identifier)
            .count()
            != 1
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
    if title_component != slide_component_name || non_style_component != slide_component_name {
        return Err(ChartTitleError::InvalidSource);
    }
    exactly_one_message(title_object, STANDIN_MESSAGE_TYPE)?;
    let non_style_message = exactly_one_message(non_style_object, CHART_NON_STYLE_MESSAGE_TYPE)?;
    let title = read_chart_title(non_style_message.data.as_slice(), limits)?;
    Ok(ChartGraph {
        slide_identifier: record.slide_identifier,
        chart_identifier,
        non_style_identifier,
        title,
        slide_component_name: slide_component_name.to_owned(),
    })
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

fn exactly_one_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<&RawMessage, ChartTitleError> {
    let mut selected = None;
    for message in &object.messages {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(message).is_some() {
            return Err(ChartTitleError::InvalidSource);
        }
    }
    selected.ok_or(ChartTitleError::InvalidSource)
}

struct ChartGraphScanBudget {
    limits: WireLimits,
    work: usize,
}

impl ChartGraphScanBudget {
    fn new(package: &Package) -> Result<Self, ChartTitleError> {
        Ok(Self {
            limits: package.wire_limits().map_err(map_wire_error)?,
            work: 0,
        })
    }

    fn charge(&mut self, amount: usize) -> Result<(), ChartTitleError> {
        let observed = self
            .work
            .checked_add(amount)
            .ok_or(ChartTitleError::InvalidSource)?;
        if observed > self.limits.max_rewrite_work() {
            return Err(ChartTitleError::LimitExceeded {
                kind: ChartTitleLimitKind::WireWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(self.limits.max_rewrite_work()),
            });
        }
        self.work = observed;
        Ok(())
    }

    fn parse(&mut self, payload: &[u8]) -> Result<Vec<WireField>, ChartTitleError> {
        self.charge(payload.len())?;
        let fields = parse_wire_fields_with_limits(payload, self.limits).map_err(map_wire_error)?;
        self.charge(fields.len())?;
        Ok(fields)
    }
}

fn scan_non_style_owners(package: &Package) -> Result<HashMap<u64, usize>, ChartTitleError> {
    let mut budget = ChartGraphScanBudget::new(package)?;
    let mut owners = HashMap::new();
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
                if let Some(count) = owners.get_mut(&non_style_identifier) {
                    *count = count.checked_add(1).ok_or(ChartTitleError::InvalidSource)?;
                } else {
                    owners
                        .try_reserve(1)
                        .map_err(|_error| ChartTitleError::Allocation { amount: 1 })?;
                    owners.insert(non_style_identifier, 1);
                }
            }
        }
    }
    Ok(owners)
}

fn rewrite_chart_title(
    source: &Package,
    selection: &ChartSelection,
    after: Option<&str>,
) -> Result<Package, ChartTitleError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.slide_component_name)
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
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        source
            .state
            .options
            .archive()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    let (message_index, message_data) = {
        let object = archive
            .object(selection.non_style_identifier)
            .ok_or(ChartTitleError::InvalidSource)?;
        let mut message_index = None;
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ != CHART_NON_STYLE_MESSAGE_TYPE {
                continue;
            }
            if message_index.replace(index).is_some() {
                return Err(ChartTitleError::InvalidSource);
            }
        }
        let message_index = message_index.ok_or(ChartTitleError::InvalidSource)?;
        (message_index, object.messages[message_index].data.clone())
    };
    let limits = source.wire_limits().map_err(map_wire_error)?;
    let patched = patch_chart_title(message_data.as_slice(), after, limits)?;
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
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    let edit = EntryEdit::new(
        selection.slide_component_name.as_str(),
        compressed.as_slice(),
    );
    let output = catalog
        .package()
        .reassemble_to_bytes(&[edit], source.state.options.archive())
        .map_err(map_archive_error)?;
    Package::from_source_with_options(output.into(), source.state.options).map_err(map_read_error)
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
            extension_count += 1;
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

fn verify_chart_candidate(
    source: &Package,
    candidate: &Package,
    slide_position: Position,
    chart_position: Position,
    expected_title: Option<&str>,
) -> Result<(), ChartTitleError> {
    if source.state.total_objects != candidate.state.total_objects {
        return Err(ChartTitleError::Verification);
    }
    let source_graphs = chart_graphs(source, slide_position, true)?;
    let candidate_graphs = chart_graphs(candidate, slide_position, true)?;
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
            || source_graph.slide_component_name != candidate_graph.slide_component_name
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
    Ok(())
}

fn patch_chart_title(
    data: &[u8],
    title: Option<&str>,
    limits: WireLimits,
) -> Result<Vec<u8>, ChartTitleError> {
    let fields = parse_wire_fields_with_limits(data, limits).map_err(map_wire_error)?;
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
            return clone_title_source(data, limits);
        };
        let extension = rewrite_chart_title_payload(
            &[],
            ChartTitleWrite::new(Some(true), Some(title)),
            chart_title_decode_options(limits)?,
        )
        .map_err(map_chart_title_codec_error)?;
        return rewrite_chart_non_style_extension(data, None, &extension, limits);
    };

    if title.is_none()
        && decode_visible_chart_title(extension, chart_title_decode_options(limits)?)
            .map_err(map_chart_title_codec_error)?
            .is_none()
    {
        return clone_title_source(data, limits);
    }
    let extension = rewrite_chart_title_payload(
        extension,
        ChartTitleWrite::new(Some(title.is_some()), title),
        chart_title_decode_options(limits)?,
    )
    .map_err(map_chart_title_codec_error)?;
    rewrite_chart_non_style_extension(data, extension_field, &extension, limits)
}

fn rewrite_chart_non_style_extension(
    data: &[u8],
    extension_field: Option<WireField>,
    replacement: &[u8],
    limits: WireLimits,
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

fn clone_title_source(data: &[u8], limits: WireLimits) -> Result<Vec<u8>, ChartTitleError> {
    if data.len() > limits.max_output_bytes() {
        return Err(ChartTitleError::LimitExceeded {
            kind: ChartTitleLimitKind::OutputBytes,
            observed: usize_to_u64(data.len()),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
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
        SlideSelectorError::EmptySlideName | SlideSelectorError::DuplicateSlideName { .. } => {
            ChartTitleError::AmbiguousSelector
        },
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
