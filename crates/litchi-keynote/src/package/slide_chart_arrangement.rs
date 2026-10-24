//! Exact-source, selector-first Keynote chart Arrange-panel transactions.
//!
//! The focused owner changes only the two interaction flags stored by the
//! selected chart drawable: `locked` and `aspect_ratio_locked`. Native chart
//! identifiers, archive names, generated protobuf values, and package
//! records remain private to this adapter. The source artifact stays the
//! preservation authority throughout the transaction.
//!
//! Changed transactions use one aggregate ledger for focused graph selection,
//! lazy arrangement codec work, candidate readback, rewrite staging, and
//! physical locality. Package-wide semantic validation remains a separately
//! bounded admission precondition, and the shared locality framing proof's
//! temporary Snappy streams are reserved locally from their exact archive
//! bounds before decompression.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The package boundary redacts native graph failures."
)]

use std::fmt;
use std::mem::size_of;
use std::sync::{Arc, OnceLock};

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::chart_arrangement_codec::{
    ChartArrangementWrite, DecodeError as ChartArrangementDecodeError,
    DecodeLimit as ChartArrangementDecodeLimit, DecodeOptions,
    WireResourceLimit as ChartArrangementWireResourceLimit, decode_chart_arrangement,
    decode_chart_arrangement_with_report, prepare_chart_arrangement_rewrite,
};
use thiserror::Error as ThisError;

use super::chart_axis_support::{self, AxisSupportBudget, AxisSupportError};
use super::edit::{OwnedChartSelector, OwnedSlideSelector, SelectorError};
use super::slide_chart_title::{
    ChartSelection, Error as ChartTitleError, select_chart, select_chart_with_budget, select_charts,
};
use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{ChartArrangement, ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const MAX_ARRANGEMENT_CODEC_BYTES: usize = 64 * 1024 * 1024;
const MAX_TRANSACTION_ALLOCATIONS: usize = 64;

/// A finite resource governed by one chart Arrange-panel operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Logical codec or transaction allocations.
    Allocations,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
    /// Bytes retained by the selected chart drawable payload.
    ArrangementBytes,
    /// Source plus candidate bytes retained by the lazy codec.
    RetainedBytes,
    /// Temporary candidate bytes required by the lazy codec.
    ScratchBytes,
    /// Caller bytes retained by the deferred selectors.
    SelectorBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::Allocations => "allocations",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::ArrangementBytes => "chart arrangement bytes",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::SelectorBytes => "selector bytes",
        })
    }
}

/// A content-redacted failure raised by a chart Arrange-panel operation.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The selected chart graph or drawable payload was malformed.
    #[error("the Keynote chart-arrangement source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote chart arrangement {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote chart-arrangement transaction")]
    Allocation {
        /// Elements or bytes requested by the failed allocation.
        amount: usize,
    },
    /// Full candidate reopening did not reproduce the requested state.
    #[error("the edited Keynote chart arrangement failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote chart-arrangement patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The source was not retained as an exact physical package.
    #[error("this Keynote source does not support physical chart arrangement edits")]
    UnsupportedSource,
    /// An exact-name slide or chart selector was ambiguous.
    #[error("the Keynote chart-arrangement selector is ambiguous")]
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
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. } => RefusalKind::UnsupportedContent,
            Self::AmbiguousSelector { .. }
            | Self::EmptySlideName { .. }
            | Self::EmptyChartName { .. } => RefusalKind::InvalidValue,
            Self::SlideNameNotFound { .. }
            | Self::SlidePositionNotFound { .. }
            | Self::ChartNameNotFound { .. }
            | Self::ChartPositionNotFound { .. } => RefusalKind::NotFound,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One mutable chart Arrange-panel value staged against an immutable package
/// snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    slide: OwnedSlideSelector,
    chart: OwnedChartSelector,
    selection: OnceLock<(ChartSelection, ChartArrangement)>,
    after: Option<ChartArrangement>,
    locked: Option<bool>,
    constrain_proportions: Option<bool>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("has_selection", &self.selection.get().is_some())
            .field(
                "has_staged",
                &(self.after.is_some()
                    || self.locked.is_some()
                    || self.constrain_proportions.is_some()),
            )
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Self, Error> {
        let maximum_selector_bytes = source.semantic_limits().max_text_bytes();
        let mut selector_bytes = 0;
        let slide = OwnedSlideSelector::from_selector(
            slide_selector.into(),
            maximum_selector_bytes,
            &mut selector_bytes,
        )
        .map_err(map_selector_error)?;
        let chart = OwnedChartSelector::from_selector(
            chart_selector.into(),
            maximum_selector_bytes,
            &mut selector_bytes,
        )
        .map_err(map_selector_error)?;
        Ok(Self {
            source,
            slide,
            chart,
            selection: OnceLock::new(),
            after: None,
            locked: None,
            constrain_proportions: None,
        })
    }

    fn resolve(&self) -> Result<&(ChartSelection, ChartArrangement), Error> {
        if let Some(selection) = self.selection.get() {
            return Ok(selection);
        }
        let selection = select_chart(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
            true,
        )
        .map_err(map_chart_title_error)?;
        let before = read_selected_arrangement(self.source, &selection)?;
        let _ = self.selection.set((selection, before));
        self.selection.get().ok_or(Error::InvalidSource)
    }

    /// Return the selected semantic slide position.
    #[must_use = "handle the deferred chart-arrangement result"]
    pub fn slide_position(&self) -> Result<Position, Error> {
        Ok(self.resolve()?.0.slide_position)
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use = "handle the deferred chart-arrangement result"]
    pub fn chart_position(&self) -> Result<Position, Error> {
        Ok(self.resolve()?.0.chart_position)
    }

    /// Return the arrangement observed when this edit first resolves.
    #[must_use = "handle the deferred chart-arrangement result"]
    pub fn before(&self) -> Result<ChartArrangement, Error> {
        Ok(self.resolve()?.1)
    }

    /// Return the arrangement staged for publication.
    #[must_use = "handle the deferred chart-arrangement result"]
    pub fn after(&self) -> Result<ChartArrangement, Error> {
        if let Some(after) = self.after {
            return Ok(after);
        }
        let mut after = self.resolve()?.1;
        if let Some(locked) = self.locked {
            after = after.with_locked(locked);
        }
        if let Some(constrain_proportions) = self.constrain_proportions {
            after = after.with_constrain_proportions(constrain_proportions);
        }
        Ok(after)
    }

    /// Stage a complete replacement of the chart Arrange-panel state.
    #[must_use]
    pub fn set(mut self, arrangement: ChartArrangement) -> Self {
        self.after = Some(arrangement);
        self.locked = None;
        self.constrain_proportions = None;
        self
    }

    /// Stage the requested lock state while retaining the other flag.
    ///
    /// Requesting `false` preserves an already-absent proto2 field, so an
    /// effective false-to-false edit can remain an exact byte no-op.
    #[must_use]
    pub fn set_locked(mut self, locked: bool) -> Self {
        if let Some(after) = self.after {
            self.after = Some(after.with_locked(locked));
        } else {
            self.locked = Some(locked);
        }
        self
    }

    /// Stage the requested aspect-ratio constraint while retaining the other
    /// flag.
    ///
    /// Requesting `false` preserves an already-absent proto2 field, so an
    /// effective false-to-false edit can remain an exact byte no-op.
    #[must_use]
    pub fn set_constrain_proportions(mut self, constrain_proportions: bool) -> Self {
        if let Some(after) = self.after {
            self.after = Some(after.with_constrain_proportions(constrain_proportions));
        } else {
            self.constrain_proportions = Some(constrain_proportions);
        }
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<Commit, Error> {
        let (source_selection, before) = self.resolve()?.clone();
        let after = self.after()?;
        let slide_position = source_selection.slide_position;
        let chart_position = source_selection.chart_position;
        let chart_identifier = source_selection.chart_identifier;
        let slide_identifier = source_selection.slide_identifier;
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        if before == after {
            let current = read_selected_arrangement(self.source, &source_selection)?;
            if current != before {
                return Err(Error::InvalidSource);
            }
            self.source.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    slide_position,
                    chart_position,
                    chart_identifier,
                    slide_identifier,
                    before,
                    after,
                    source_payload: None,
                    target_payload: None,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }

        if !catalog.source_is_exact() {
            return Err(Error::Refused(Refusal::UnsupportedSource));
        }
        self.source.validate().map_err(map_read_error)?;
        let mut budget = ArrangementTransactionBudget::new(self.source)?;
        budget.source_package(self.source.source_bytes().len())?;
        budget
            .charge_selection_scans(self.source, true)
            .map_err(map_axis_support_error)?;
        let source_selection = select_chart_with_budget(
            self.source,
            SlideSelector::position(slide_position),
            ChartSelector::index(chart_position.get()),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        let (current, source_payload_bytes) = read_selected_arrangement_with_transaction_budget(
            self.source,
            &source_selection,
            &mut budget,
        )?;
        if source_selection.chart_identifier != chart_identifier
            || source_selection.slide_identifier != slide_identifier
            || current != before
        {
            return Err(Error::InvalidSource);
        }
        let source_payload = copy_payload_with_budget(source_payload_bytes, &mut budget)?;
        let rewritten =
            rewrite_chart_arrangement(self.source, &source_selection, after, &mut budget)?;
        verify_chart_candidate(
            self.source,
            &rewritten.package,
            &source_selection,
            after,
            rewritten.selected_payload.as_ref(),
            &mut budget,
        )?;
        let target = physical_catalog(&rewritten.package)?.shared_source();
        Ok(Commit {
            package: rewritten.package,
            patch: Patch {
                artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
                slide_position,
                chart_position,
                chart_identifier,
                slide_identifier,
                before,
                after,
                source_payload: Some(source_payload),
                target_payload: Some(rewritten.selected_payload),
            },
            diagnostics: Diagnostics::published(),
        })
    }
}

fn map_selector_error(error: SelectorError) -> Error {
    match error {
        SelectorError::Limit { observed, maximum } => Error::Limit {
            kind: LimitKind::SelectorBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        SelectorError::Allocation { amount } => Error::Allocation { amount },
    }
}

/// An exact-source-checked reversible semantic chart-arrangement patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    slide_identifier: u64,
    before: ChartArrangement,
    after: ChartArrangement,
    source_payload: Option<Arc<[u8]>>,
    target_payload: Option<Arc<[u8]>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
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

    /// Return the arrangement required from the source package.
    #[must_use]
    pub const fn before(&self) -> ChartArrangement {
        self.before
    }

    /// Return the arrangement produced by the target package.
    #[must_use]
    pub const fn after(&self) -> ChartArrangement {
        self.after
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

    /// Return whether this patch preserves exact source bytes and state.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_payload.is_none()
            && self.target_payload.is_none()
            && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            chart_position: self.chart_position,
            chart_identifier: self.chart_identifier,
            slide_identifier: self.slide_identifier,
            before: self.after,
            after: self.before,
            source_payload: self.target_payload.as_ref().map(Arc::clone),
            target_payload: self.source_payload.as_ref().map(Arc::clone),
        }
    }
}

/// Compact evidence describing one chart-arrangement commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews: 0,
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

    /// Return how many rendering previews were removed by this transaction.
    ///
    /// Arrange state does not change chart content, so preview invalidation is
    /// intentionally never requested and this value is always zero.
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

/// The fully verified result of one immutable chart-arrangement transaction.
#[must_use = "a Keynote chart-arrangement commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
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

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the selected Keynote chart's Arrange-panel state.
    ///
    /// Absent native proto2 booleans have the effective value `false`.
    pub fn slide_chart_arrangement<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartArrangement, Error> {
        let selection = select_chart(self, slide_selector.into(), chart_selector.into(), true)
            .map_err(map_chart_title_error)?;
        read_selected_arrangement(self, &selection)
    }

    /// Read every chart's Arrange-panel state in one checked slide traversal.
    ///
    /// Chart graph ownership and selector validation are performed once, then
    /// each chart drawable is decoded exactly once in slide chart-drawable
    /// z-order. Element `N` is the value selected by
    /// `ChartSelector::index(N)`. The returned collection contains only
    /// archive-free semantic values. Absent native proto2 booleans read as
    /// `false`.
    pub fn slide_chart_arrangements<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
    ) -> Result<Box<[ChartArrangement]>, Error> {
        let selections =
            select_charts(self, slide_selector.into(), true).map_err(map_chart_title_error)?;
        let limits = self.wire_limits().map_err(map_wire_error)?;
        let mut budget = ArrangementReadBudget::new(limits);
        let mut arrangements = Vec::new();
        arrangements
            .try_reserve_exact(selections.len())
            .map_err(|_error| Error::Allocation {
                amount: selections.len(),
            })?;
        for selection in selections.iter() {
            let (arrangement, _) =
                read_selected_arrangement_with_budget(self, selection, limits, &mut budget)?;
            arrangements.push(arrangement);
        }
        Ok(arrangements.into_boxed_slice())
    }

    /// Start an exact immutable edit of one selected chart's Arrange-panel
    /// state.
    ///
    /// Entry retains only the source and the archive-free selectors. Chart
    /// graph selection and arrangement decoding are deferred until a value is
    /// read, an aspect is staged, or the edit is committed. Selector names are
    /// admitted against the package's finite selector-byte limit at entry.
    pub fn edit_slide_chart_arrangement<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Edit<'_>, Error> {
        Edit::new(self, slide_selector, chart_selector)
    }

    /// Apply an exact-source-checked chart-arrangement patch.
    pub fn apply_slide_chart_arrangement(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            let selection = select_chart(
                self,
                SlideSelector::position(patch.slide_position),
                ChartSelector::index(patch.chart_position.get()),
                true,
            )
            .map_err(map_chart_title_error)?;
            let current = read_selected_arrangement(self, &selection)?;
            if selection.chart_identifier != patch.chart_identifier
                || selection.slide_identifier != patch.slide_identifier
                || current != patch.before
            {
                return Err(Error::PatchConflict);
            }
            self.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        let mut budget = ArrangementTransactionBudget::new(self)?;
        budget.source_package(self.source_bytes().len())?;
        budget
            .charge_selection_scans(self, true)
            .map_err(map_axis_support_error)?;
        let selection = select_chart_with_budget(
            self,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        let (current, _) =
            read_selected_arrangement_with_transaction_budget(self, &selection, &mut budget)?;
        if selection.chart_identifier != patch.chart_identifier
            || selection.slide_identifier != patch.slide_identifier
            || current != patch.before
        {
            return Err(Error::PatchConflict);
        }
        budget.candidate_reopen(patch.artifacts.target().len())?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        verify_chart_candidate(
            self,
            &candidate,
            &selection,
            patch.after,
            patch
                .target_payload
                .as_deref()
                .ok_or(Error::PatchConflict)?,
            &mut budget,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }
}

fn read_selected_arrangement(
    package: &Package,
    selection: &ChartSelection,
) -> Result<ChartArrangement, Error> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    read_selected_arrangement_with_limits(package, selection, limits)
}

fn read_selected_arrangement_with_limits(
    package: &Package,
    selection: &ChartSelection,
    limits: WireLimits,
) -> Result<ChartArrangement, Error> {
    let (_, object) = package
        .object_with_component(selection.chart_identifier)
        .ok_or(Error::InvalidSource)?;
    let message = exactly_one_message(object, CHART_MESSAGE_TYPE)?;
    let snapshot = decode_chart_arrangement(
        message.data.as_slice(),
        arrangement_decode_options(message.data.as_slice(), limits)?,
    )
    .map_err(map_arrangement_codec_error)?;
    Ok(ChartArrangement::new(
        snapshot.is_locked(),
        snapshot.is_constrained(),
    ))
}

#[derive(Debug, Clone, Copy)]
struct ArrangementReadBudget {
    input_limit: usize,
    field_limit: usize,
    work_limit: usize,
    input_used: usize,
    fields_used: usize,
    work_used: usize,
}

#[derive(Debug, Clone, Copy)]
struct ArrangementTransactionBudget {
    input_limit: usize,
    output_limit: usize,
    field_limit: usize,
    graph_field_limit: usize,
    work_limit: usize,
    retained_limit: usize,
    scratch_limit: usize,
    graph_allocation_limit: usize,
    input_used: usize,
    output_used: usize,
    field_used: usize,
    graph_field_used: usize,
    work_used: usize,
    allocations_used: usize,
    graph_allocations_used: usize,
    retained_used: usize,
    scratch_used: usize,
}

impl ArrangementTransactionBudget {
    fn new(package: &Package) -> Result<Self, Error> {
        let physical = package.state.options.archive();
        let input_limit =
            usize::try_from(physical.max_input_bytes()).map_err(|_| Error::InvalidSource)?;
        let aggregate_limit = input_limit.checked_mul(4).ok_or(Error::InvalidSource)?;
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let graph_allocation_limit = graph_allocation_limit(package, aggregate_limit)?;
        Ok(Self {
            input_limit: aggregate_limit,
            output_limit: aggregate_limit,
            field_limit: wire.max_fields(),
            graph_field_limit: wire.max_rewrite_work(),
            work_limit: wire.max_rewrite_work(),
            retained_limit: aggregate_limit,
            scratch_limit: aggregate_limit,
            graph_allocation_limit,
            input_used: 0,
            output_used: 0,
            field_used: 0,
            graph_field_used: 0,
            work_used: 0,
            allocations_used: 0,
            graph_allocations_used: 0,
            retained_used: 0,
            scratch_used: 0,
        })
    }

    fn source_package(&mut self, bytes: usize) -> Result<(), Error> {
        self.input(bytes)?;
        self.work(bytes)
    }

    fn input(&mut self, amount: usize) -> Result<(), Error> {
        self.input_used = charge_aggregate_limit(
            LimitKind::InputBytes,
            self.input_used,
            amount,
            self.input_limit,
        )?;
        Ok(())
    }

    fn output(&mut self, amount: usize) -> Result<(), Error> {
        self.output_used = charge_aggregate_limit(
            LimitKind::OutputBytes,
            self.output_used,
            amount,
            self.output_limit,
        )?;
        Ok(())
    }

    fn fields(&mut self, amount: usize) -> Result<(), Error> {
        self.field_used = charge_aggregate_limit(
            LimitKind::WireFields,
            self.field_used,
            amount,
            self.field_limit,
        )?;
        Ok(())
    }

    fn graph_fields(&mut self, amount: usize) -> Result<(), Error> {
        self.graph_field_used = charge_aggregate_limit(
            LimitKind::WireFields,
            self.graph_field_used,
            amount,
            self.graph_field_limit,
        )?;
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), Error> {
        self.work_used =
            charge_aggregate_limit(LimitKind::WireWork, self.work_used, amount, self.work_limit)?;
        Ok(())
    }

    fn allocations(&mut self, amount: usize) -> Result<(), Error> {
        self.allocations_used = charge_aggregate_limit(
            LimitKind::Allocations,
            self.allocations_used,
            amount,
            MAX_TRANSACTION_ALLOCATIONS,
        )?;
        Ok(())
    }

    fn graph_allocations(&mut self, amount: usize) -> Result<(), Error> {
        self.graph_allocations_used = charge_aggregate_limit(
            LimitKind::Allocations,
            self.graph_allocations_used,
            amount,
            self.graph_allocation_limit,
        )?;
        Ok(())
    }

    fn retained(&mut self, amount: usize) -> Result<(), Error> {
        self.retained_used = charge_aggregate_limit(
            LimitKind::RetainedBytes,
            self.retained_used,
            amount,
            self.retained_limit,
        )?;
        Ok(())
    }

    fn scratch(&mut self, amount: usize) -> Result<(), Error> {
        self.scratch_used = charge_aggregate_limit(
            LimitKind::ScratchBytes,
            self.scratch_used,
            amount,
            self.scratch_limit,
        )?;
        Ok(())
    }

    fn codec(
        &mut self,
        requirements: litchi_iwa_protos::chart_arrangement_codec::RewriteExecutionRequirements,
    ) -> Result<(), Error> {
        self.output(requirements.output_bytes)?;
        self.fields(requirements.fields)?;
        self.work(requirements.work_bytes)?;
        self.allocations(requirements.allocations)?;
        self.retained(requirements.retained_bytes)?;
        self.scratch(requirements.scratch_bytes)
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), Error> {
        self.output(requirements.output_bytes())?;
        self.work(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())
    }

    fn candidate_reopen(&mut self, bytes: usize) -> Result<(), Error> {
        self.input(bytes)?;
        self.work(bytes)?;
        self.allocations(1)?;
        self.retained(bytes)
    }

    fn verify_packages(
        &mut self,
        source_bytes: usize,
        candidate_bytes: usize,
    ) -> Result<(), Error> {
        let amount = source_bytes
            .checked_add(candidate_bytes)
            .ok_or(Error::InvalidSource)?;
        self.work(amount)
    }
}

fn graph_allocation_limit(package: &Package, retained_limit: usize) -> Result<usize, Error> {
    // Graph scans allocate bounded wire/reference vectors. Their logical
    // allocation count is therefore finite when derived from the configured
    // semantic graph ceiling and the retained-byte ledger, without walking
    // every object a second time merely to establish a budget.
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    // A reference vector stores `u64` elements while a wire-span vector stores
    // `WireField`s. Use the smaller element size so the retained-byte term
    // bounds either allocation family, including a one-element vector.
    let graph_element_bytes = size_of::<litchi_iwa_common::wire::WireField>().min(size_of::<u64>());
    package
        .state
        .source
        .components()
        .len()
        .checked_add(package.state.total_objects)
        .and_then(|value| value.checked_add(archive_limits.max_messages()))
        .and_then(|value| value.checked_add(package.semantic_limits().max_references()))
        .and_then(|value| {
            value.checked_add(retained_limit.checked_div(graph_element_bytes).unwrap_or(0))
        })
        .and_then(|value| value.checked_add(1))
        .ok_or(Error::InvalidSource)
}

impl AxisSupportBudget for ArrangementTransactionBudget {
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
        self.work(amount).map_err(axis_support_budget_error)
    }

    fn charge_input(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        self.input(amount).map_err(axis_support_budget_error)
    }

    fn charge_wire_vector(&mut self, payload: usize) -> Result<(), AxisSupportError> {
        let vector_bytes = payload
            .checked_mul(size_of::<litchi_iwa_common::wire::WireField>())
            .ok_or(AxisSupportError::InvalidSource)?;
        self.input(payload).map_err(axis_support_budget_error)?;
        self.work(payload).map_err(axis_support_budget_error)?;
        self.graph_allocations(1)
            .map_err(axis_support_budget_error)?;
        self.retained(vector_bytes)
            .map_err(axis_support_budget_error)?;
        self.scratch(vector_bytes)
            .map_err(axis_support_budget_error)
    }

    fn finish_wire_scan(&mut self, fields: usize) -> Result<(), AxisSupportError> {
        self.graph_fields(fields)
            .map_err(axis_support_budget_error)?;
        self.work(fields).map_err(axis_support_budget_error)
    }

    fn charge_reference_vector(&mut self, capacity: usize) -> Result<(), AxisSupportError> {
        let bytes = capacity
            .checked_mul(size_of::<u64>())
            .ok_or(AxisSupportError::InvalidSource)?;
        self.graph_allocations(1)
            .map_err(axis_support_budget_error)?;
        self.retained(bytes).map_err(axis_support_budget_error)?;
        self.scratch(bytes).map_err(axis_support_budget_error)
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        self.work(amount).map_err(axis_support_budget_error)
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
        self.work(amount).map_err(axis_support_budget_error)?;
        self.graph_allocations(retained_vectors)
            .map_err(axis_support_budget_error)
    }

    fn charge_locality_scan(&mut self, package: &Package) -> Result<(), AxisSupportError> {
        let amount = package
            .state
            .source
            .components()
            .len()
            .checked_add(
                package
                    .state
                    .total_objects
                    .checked_mul(2)
                    .ok_or(AxisSupportError::InvalidSource)?,
            )
            .ok_or(AxisSupportError::InvalidSource)?;
        self.work(amount).map_err(axis_support_budget_error)
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        self.work(amount).map_err(axis_support_budget_error)
    }

    fn metadata_options(
        &self,
        package: &Package,
    ) -> Result<litchi_iwa_protos::package_metadata_codec::RewriteOptions, AxisSupportError> {
        let limits = package
            .wire_limits()
            .map_err(map_wire_error)
            .map_err(axis_support_budget_error)?;
        let remaining_input = self
            .input_limit
            .checked_sub(self.input_used)
            .ok_or(AxisSupportError::InvalidSource)?
            .min(limits.max_input_bytes());
        let remaining_output = self
            .output_limit
            .checked_sub(self.output_used)
            .ok_or(AxisSupportError::InvalidSource)?
            .min(limits.max_output_bytes());
        let remaining_fields = self
            .graph_field_limit
            .checked_sub(self.graph_field_used)
            .ok_or(AxisSupportError::InvalidSource)?
            .min(limits.max_fields());
        let remaining_work = self
            .work_limit
            .checked_sub(self.work_used)
            .ok_or(AxisSupportError::InvalidSource)?
            .min(limits.max_rewrite_work());
        let remaining_allocations = self
            .graph_allocation_limit
            .checked_sub(self.graph_allocations_used)
            .ok_or(AxisSupportError::InvalidSource)?;
        let recursion =
            u32::try_from(limits.max_nesting()).map_err(|_| AxisSupportError::InvalidSource)?;
        Ok(
            litchi_iwa_protos::package_metadata_codec::RewriteOptions::new(
                remaining_input,
                remaining_output,
                remaining_fields,
                remaining_work,
                recursion,
                package.semantic_limits().max_objects(),
                package.semantic_limits().max_references(),
                remaining_allocations,
            ),
        )
    }

    fn charge_metadata_report(
        &mut self,
        report: litchi_iwa_protos::package_metadata_codec::RewriteReport,
    ) -> Result<(), AxisSupportError> {
        self.input(report.input_bytes())
            .map_err(axis_support_budget_error)?;
        self.output(report.output_bytes())
            .map_err(axis_support_budget_error)?;
        self.graph_fields(report.fields())
            .map_err(axis_support_budget_error)?;
        let work = report
            .work_bytes()
            .checked_add(report.components_scanned())
            .and_then(|value| value.checked_add(report.references_scanned()))
            .ok_or(AxisSupportError::InvalidSource)?;
        self.work(work).map_err(axis_support_budget_error)?;
        self.graph_allocations(report.allocations())
            .map_err(axis_support_budget_error)?;
        self.retained(report.retained_bytes())
            .map_err(axis_support_budget_error)?;
        self.scratch(report.scratch_bytes())
            .map_err(axis_support_budget_error)
    }
}

impl ArrangementReadBudget {
    const fn new(limits: WireLimits) -> Self {
        Self {
            input_limit: limits.max_input_bytes(),
            field_limit: limits.max_fields(),
            work_limit: limits.max_rewrite_work(),
            input_used: 0,
            fields_used: 0,
            work_used: 0,
        }
    }

    const fn remaining_fields(self) -> usize {
        self.field_limit.saturating_sub(self.fields_used)
    }

    const fn remaining_work(self) -> usize {
        self.work_limit.saturating_sub(self.work_used)
    }

    fn check_input(self, amount: usize) -> Result<(), Error> {
        check_aggregate_limit(
            LimitKind::WireBytes,
            self.input_used,
            amount,
            self.input_limit,
        )
    }

    fn charge(
        &mut self,
        report: litchi_iwa_protos::chart_arrangement_codec::DecodeReport,
    ) -> Result<(), Error> {
        self.input_used = charge_aggregate_limit(
            LimitKind::WireBytes,
            self.input_used,
            report.source_bytes(),
            self.input_limit,
        )?;
        self.fields_used = charge_aggregate_limit(
            LimitKind::WireFields,
            self.fields_used,
            report.fields(),
            self.field_limit,
        )?;
        self.work_used = charge_aggregate_limit(
            LimitKind::WireWork,
            self.work_used,
            report.work_bytes(),
            self.work_limit,
        )?;
        Ok(())
    }
}

fn read_selected_arrangement_with_budget<'a>(
    package: &'a Package,
    selection: &ChartSelection,
    limits: WireLimits,
    budget: &mut ArrangementReadBudget,
) -> Result<(ChartArrangement, &'a [u8]), Error> {
    let (_, object) = package
        .object_with_component(selection.chart_identifier)
        .ok_or(Error::InvalidSource)?;
    let message = exactly_one_message(object, CHART_MESSAGE_TYPE)?;
    budget.check_input(message.data.len())?;
    let options = arrangement_decode_options_with_budget(
        message.data.as_slice(),
        limits,
        budget.remaining_fields(),
        budget.remaining_work(),
    )?;
    let (snapshot, report) = decode_chart_arrangement_with_report(message.data.as_slice(), options)
        .map_err(map_arrangement_codec_error)?;
    budget.charge(report)?;
    Ok((
        ChartArrangement::new(snapshot.is_locked(), snapshot.is_constrained()),
        message.data.as_slice(),
    ))
}

fn read_selected_arrangement_with_transaction_budget<'a>(
    package: &'a Package,
    selection: &ChartSelection,
    budget: &mut ArrangementTransactionBudget,
) -> Result<(ChartArrangement, &'a [u8]), Error> {
    let (_, object) = package
        .object_with_component(selection.chart_identifier)
        .ok_or(Error::InvalidSource)?;
    let message = exactly_one_message(object, CHART_MESSAGE_TYPE)?;
    budget.input(message.data.len())?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let options = arrangement_decode_options_with_budget(
        message.data.as_slice(),
        limits,
        budget.field_limit.saturating_sub(budget.field_used),
        budget.work_limit.saturating_sub(budget.work_used),
    )?;
    let (snapshot, report) = decode_chart_arrangement_with_report(message.data.as_slice(), options)
        .map_err(map_arrangement_codec_error)?;
    budget.fields(report.fields())?;
    budget.work(report.work_bytes())?;
    budget.graph_allocations(report.allocations())?;
    budget.retained(report.retained_bytes())?;
    budget.scratch(report.scratch_bytes())?;
    Ok((
        ChartArrangement::new(snapshot.is_locked(), snapshot.is_constrained()),
        message.data.as_slice(),
    ))
}

fn copy_payload_with_budget(
    payload: &[u8],
    budget: &mut ArrangementTransactionBudget,
) -> Result<Arc<[u8]>, Error> {
    budget.allocations(1)?;
    budget.retained(payload.len())?;
    Ok(Arc::<[u8]>::from(payload))
}

fn check_aggregate_limit(
    kind: LimitKind,
    used: usize,
    amount: usize,
    maximum: usize,
) -> Result<(), Error> {
    charge_aggregate_limit(kind, used, amount, maximum).map(|_observed| ())
}

fn charge_aggregate_limit(
    kind: LimitKind,
    used: usize,
    amount: usize,
    maximum: usize,
) -> Result<usize, Error> {
    let observed = used.checked_add(amount).ok_or_else(|| Error::Limit {
        kind,
        observed: u64::MAX,
        maximum: usize_to_u64(maximum),
    })?;
    if observed > maximum {
        return Err(Error::Limit {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        });
    }
    Ok(observed)
}

struct RewrittenChartArrangement {
    package: Package,
    selected_payload: Arc<[u8]>,
}

fn rewrite_chart_arrangement(
    source: &Package,
    selection: &ChartSelection,
    replacement: ChartArrangement,
    budget: &mut ArrangementTransactionBudget,
) -> Result<RewrittenChartArrangement, Error> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.slide_component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.input(entry.data().len())?;
    budget.work(entry.data().len())?;
    let stream = SnappyStream::decompress(entry.data(), snappy_limits).map_err(map_core_error)?;
    budget.output(stream.as_bytes().len())?;
    budget.work(stream.as_bytes().len())?;
    budget.allocations(1)?;
    budget.retained(stream.as_bytes().len())?;
    budget.scratch(stream.as_bytes().len())?;
    budget.allocations(1)?;
    budget.retained(stream.as_bytes().len())?;
    let mut archive = Archive::parse(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    let (message_index, message_data) = {
        let object = archive
            .object(selection.chart_identifier)
            .ok_or(Error::InvalidSource)?;
        let mut index = None;
        for (candidate, message) in object.messages.iter().enumerate() {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            if index.replace(candidate).is_some() {
                return Err(Error::InvalidSource);
            }
        }
        let index = index.ok_or(Error::InvalidSource)?;
        (index, object.messages[index].data.as_slice())
    };
    let limits = source.wire_limits().map_err(map_wire_error)?;
    let patched = patch_chart_arrangement(message_data, replacement, limits, budget)?;
    budget.allocations(1)?;
    budget.retained(patched.len())?;
    let selected_payload = Arc::<[u8]>::from(patched.as_slice());
    archive
        .object_mut(selection.chart_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: patched,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_bound = archive
        .encoded_len(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound = SnappyStream::maximum_compressed_len(encoded_bound, snappy_limits)
        .map_err(map_core_error)?;
    let maximum_entry =
        usize::try_from(physical_limits.max_entry_bytes()).map_err(|_| Error::InvalidSource)?;
    if encoded_bound > maximum_entry || compressed_bound > maximum_entry {
        return Err(Error::Limit {
            kind: LimitKind::EntryBytes,
            observed: usize_to_u64(encoded_bound.max(compressed_bound)),
            maximum: usize_to_u64(maximum_entry),
        });
    }
    let package_bound = source
        .source_bytes()
        .len()
        .checked_sub(entry.data().len())
        .and_then(|bytes| bytes.checked_add(compressed_bound))
        .ok_or(Error::InvalidSource)?;
    let maximum_package =
        usize::try_from(physical_limits.max_input_bytes()).map_err(|_| Error::InvalidSource)?;
    if package_bound > maximum_package {
        return Err(Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: usize_to_u64(package_bound),
            maximum: usize_to_u64(maximum_package),
        });
    }
    budget.output(encoded_bound)?;
    budget.work(encoded_bound)?;
    budget.allocations(1)?;
    budget.retained(encoded_bound)?;
    budget.scratch(encoded_bound)?;
    let bytes = archive.to_bytes(archive_limits).map_err(map_core_error)?;
    budget.output(compressed_bound)?;
    budget.work(compressed_bound)?;
    budget.allocations(1)?;
    budget.retained(compressed_bound)?;
    budget.scratch(compressed_bound)?;
    let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
    let edits = [EntryEdit::new(
        selection.slide_component_name.as_str(),
        compressed.as_slice(),
    )];
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.reassembly(requirements)?;
    budget.candidate_reopen(requirements.output_bytes())?;
    let execution_limits = requirements.exact_limits();
    let output = prepared
        .execute(execution_limits)
        .map_err(map_archive_error)?;
    let package = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok(RewrittenChartArrangement {
        package,
        selected_payload,
    })
}

fn patch_chart_arrangement(
    data: &[u8],
    replacement: ChartArrangement,
    limits: WireLimits,
    budget: &mut ArrangementTransactionBudget,
) -> Result<Vec<u8>, Error> {
    budget.input(data.len())?;
    budget.work(data.len())?;
    let options = arrangement_decode_options_with_budget(
        data,
        limits,
        budget.field_limit.saturating_sub(budget.field_used),
        budget.work_limit.saturating_sub(budget.work_used),
    )?;
    let prepared = prepare_chart_arrangement_rewrite(
        data,
        ChartArrangementWrite::new(replacement.locked(), replacement.constrain_proportions()),
        options,
    )
    .map_err(map_arrangement_codec_error)?;
    budget.codec(prepared.execution_requirements())?;
    prepared
        .execute(prepared.execution_requirements().exact())
        .map(|output| output.into_output())
        .map_err(map_arrangement_codec_error)
}

fn verify_chart_candidate(
    source: &Package,
    candidate: &Package,
    source_selection: &ChartSelection,
    expected: ChartArrangement,
    expected_payload: &[u8],
    budget: &mut ArrangementTransactionBudget,
) -> Result<(), Error> {
    budget.verify_packages(source.source_bytes().len(), candidate.source_bytes().len())?;
    // Semantic validation remains the package's separately bounded ingress
    // precondition. The transaction ledger below covers the focused graph,
    // codec, rewrite, and physical-locality proof without duplicating the
    // package-wide semantic cache in a second accounting model.
    candidate.validate().map_err(map_read_error)?;
    if source.state.total_objects != candidate.state.total_objects {
        return Err(Error::Verification);
    }
    budget
        .charge_selection_scans(candidate, true)
        .map_err(map_axis_support_error)?;
    let candidate_selection = select_chart_with_budget(
        candidate,
        SlideSelector::position(source_selection.slide_position),
        ChartSelector::index(source_selection.chart_position.get()),
        true,
        budget,
    )
    .map_err(map_axis_support_error)?;
    if candidate_selection.chart_identifier != source_selection.chart_identifier
        || candidate_selection.slide_identifier != source_selection.slide_identifier
        || source_selection.non_style_identifier != candidate_selection.non_style_identifier
        || source_selection.title != candidate_selection.title
        || source_selection.slide_component_name != candidate_selection.slide_component_name
    {
        return Err(Error::Verification);
    }
    let (actual, candidate_payload) =
        read_selected_arrangement_with_transaction_budget(candidate, &candidate_selection, budget)?;
    if actual != expected || candidate_payload != expected_payload {
        return Err(Error::Verification);
    }
    let (_, source_object) = source
        .object_with_component(source_selection.chart_identifier)
        .ok_or(Error::Verification)?;
    budget.work(source_object.messages.len())?;
    let selected_message_index = unique_message_index(source_object, CHART_MESSAGE_TYPE)?;
    let source_catalog = physical_catalog(source)?.package();
    let candidate_catalog = physical_catalog(candidate)?.package();
    let preview_work = source_catalog
        .len()
        .checked_add(candidate_catalog.len())
        .and_then(|entries| entries.checked_mul(3))
        .and_then(|work| {
            preview_entry_work(source_catalog)
                .and_then(|source_work| {
                    preview_entry_work(candidate_catalog).map(|candidate_work| {
                        work.checked_add(source_work)
                            .and_then(|work| work.checked_add(candidate_work))
                    })
                })
                .flatten()
        })
        .ok_or(Error::InvalidSource)?;
    budget.work(preview_work)?;
    if !super::rendering_invalidation::root_previews_preserved(source_catalog, candidate_catalog)
        .map_err(|_| Error::Verification)?
    {
        return Err(Error::Verification);
    }
    charge_locality_stream_buffers(
        source,
        candidate,
        &source_selection.slide_component_name,
        budget,
    )?;
    chart_axis_support::verify_package_locality_for_component(
        source,
        candidate,
        &source_selection.slide_component_name,
        source_selection.chart_identifier,
        selected_message_index,
        false,
        Some(expected_payload),
        budget,
    )
    .map_err(map_axis_support_error)
}

fn charge_locality_stream_buffers(
    source: &Package,
    candidate: &Package,
    selected_component_name: &str,
    budget: &mut ArrangementTransactionBudget,
) -> Result<(), Error> {
    let source_component = source
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selected_component_name)
        .ok_or(Error::Verification)?;
    let candidate_component = candidate
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selected_component_name)
        .ok_or(Error::Verification)?;
    let source_bound = parsed_archive_stream_bound(source_component.archive())?;
    let candidate_bound = parsed_archive_stream_bound(candidate_component.archive())?;
    let streams = source_bound
        .checked_add(candidate_bound)
        .ok_or(Error::InvalidSource)?;
    // The shared framing proof materializes one bounded Snappy stream for
    // each package. Reserve both output buffers before it starts; its wire
    // work ledger remains responsible for the byte scans and framing parses.
    budget.graph_allocations(2)?;
    budget.retained(streams)?;
    budget.scratch(streams)
}

fn parsed_archive_stream_bound(archive: &Archive) -> Result<usize, Error> {
    archive.objects.iter().try_fold(0usize, |bound, object| {
        let end = usize::try_from(object.data_offset)
            .map_err(|_| Error::InvalidSource)?
            .checked_add(usize::try_from(object.data_length).map_err(|_| Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
        Ok(bound.max(end))
    })
}

fn preview_entry_work(catalog: &litchi_iwa_archive::package::Catalog) -> Option<usize> {
    catalog
        .iter()
        .filter(|entry| super::rendering_invalidation::is_root_preview_name(entry.name()))
        .try_fold(0usize, |work, entry| {
            let metadata = entry.metadata();
            let bytes = entry
                .name()
                .len()
                .checked_add(entry.raw_name().len())
                .and_then(|value| value.checked_add(entry.data().len()))
                .and_then(|value| value.checked_add(entry.raw_record().local_record().len()))
                .and_then(|value| value.checked_add(entry.raw_record().compressed_data().len()))
                .and_then(|value| {
                    value.checked_add(entry.raw_record().central_directory_record().len())
                })
                .and_then(|value| value.checked_add(metadata.local().name().len()))
                .and_then(|value| value.checked_add(metadata.local().extra().len()))
                .and_then(|value| value.checked_add(metadata.local().comment().len()))
                .and_then(|value| value.checked_add(metadata.central().name().len()))
                .and_then(|value| value.checked_add(metadata.central().extra().len()))
                .and_then(|value| value.checked_add(metadata.central().comment().len()))
                .and_then(|value| value.checked_add(32))?;
            work.checked_add(bytes)
        })
}

fn unique_message_index(object: &ArchiveObject, message_type: u32) -> Result<usize, Error> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(index).is_some() {
            return Err(Error::Verification);
        }
    }
    selected.ok_or(Error::Verification)
}

fn exactly_one_message(object: &ArchiveObject, message_type: u32) -> Result<&RawMessage, Error> {
    let mut selected = None;
    for message in &object.messages {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(message).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), Error> {
    for object in &archive.objects {
        let offset = usize::try_from(object.header_offset).map_err(|_| Error::InvalidSource)?;
        let remaining = source.get(offset..).ok_or(Error::InvalidSource)?;
        let (header_bytes, prefix_bytes) = litchi_iwa_common::decode_varint_from_bytes(remaining)
            .map_err(|_| Error::InvalidSource)?;
        if prefix_bytes != litchi_iwa_common::varint::encoded_len(header_bytes) {
            return Err(Error::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(u64::try_from(prefix_bytes).map_err(|_| Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(Error::InvalidSource)?
                != object.data_offset
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn arrangement_decode_options(source: &[u8], limits: WireLimits) -> Result<DecodeOptions, Error> {
    arrangement_decode_options_with_budget(
        source,
        limits,
        limits.max_fields(),
        limits.max_rewrite_work(),
    )
}

fn arrangement_decode_options_with_budget(
    source: &[u8],
    limits: WireLimits,
    max_fields: usize,
    max_work: usize,
) -> Result<DecodeOptions, Error> {
    let recursion_limit = u32::try_from(limits.max_nesting()).map_err(|_| Error::Limit {
        kind: LimitKind::WireNesting,
        observed: usize_to_u64(limits.max_nesting()),
        maximum: u64::from(u32::MAX),
    })?;
    let source_limit = source.len().max(1);
    let codec_input = source_limit.min(MAX_ARRANGEMENT_CODEC_BYTES);
    let codec_output = limits
        .max_output_bytes()
        .min(MAX_ARRANGEMENT_CODEC_BYTES)
        .max(codec_input);
    let codec_fields = max_fields.min(MAX_ARRANGEMENT_CODEC_BYTES);
    let codec_work = max_work.min(MAX_ARRANGEMENT_CODEC_BYTES);
    let codec_retained = codec_input
        .checked_add(codec_output)
        .unwrap_or(MAX_ARRANGEMENT_CODEC_BYTES)
        .min(MAX_ARRANGEMENT_CODEC_BYTES);
    Ok(
        DecodeOptions::new(codec_input, codec_fields, codec_work, recursion_limit)
            .with_max_output_bytes(codec_output)
            .with_max_allocations(1)
            .with_max_retained_bytes(codec_retained.max(codec_input))
            .with_max_scratch_bytes(codec_output.max(1)),
    )
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::Refused(Refusal::UnsupportedSource)),
    }
}

fn map_chart_title_error(error: ChartTitleError) -> Error {
    match error {
        ChartTitleError::Refused(crate::chart::title::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        ChartTitleError::Refused(crate::chart::title::Refusal::AmbiguousSelector) => {
            Error::Refused(Refusal::AmbiguousSelector)
        },
        ChartTitleError::Refused(crate::chart::title::Refusal::EmptySlideName) => {
            Error::Refused(Refusal::EmptySlideName)
        },
        ChartTitleError::Refused(crate::chart::title::Refusal::SlideNameNotFound) => {
            Error::Refused(Refusal::SlideNameNotFound)
        },
        ChartTitleError::Refused(crate::chart::title::Refusal::SlidePositionNotFound {
            position,
        }) => Error::Refused(Refusal::SlidePositionNotFound { position }),
        ChartTitleError::Refused(crate::chart::title::Refusal::ChartNameNotFound) => {
            Error::Refused(Refusal::ChartNameNotFound)
        },
        ChartTitleError::Refused(crate::chart::title::Refusal::ChartPositionNotFound {
            position,
        }) => Error::Refused(Refusal::ChartPositionNotFound { position }),
        ChartTitleError::Refused(crate::chart::title::Refusal::EmptyChartName) => {
            Error::Refused(Refusal::EmptyChartName)
        },
        ChartTitleError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_chart_title_limit(kind),
            observed,
            maximum,
        },
        ChartTitleError::Allocation { amount } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn axis_support_budget_error(error: Error) -> AxisSupportError {
    match error {
        Error::Refused(Refusal::UnsupportedSource) => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::UnsupportedSource,
        ),
        Error::Refused(Refusal::AmbiguousSelector) => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::AmbiguousSelector,
        ),
        Error::Refused(Refusal::EmptySlideName) => {
            AxisSupportError::Selector(chart_axis_support::AxisSupportSelectorError::EmptySlideName)
        },
        Error::Refused(Refusal::SlideNameNotFound) => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::SlideNameNotFound,
        ),
        Error::Refused(Refusal::SlidePositionNotFound { position }) => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::SlidePositionNotFound { position },
        ),
        Error::Refused(Refusal::ChartNameNotFound) => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::ChartNameNotFound,
        ),
        Error::Refused(Refusal::ChartPositionNotFound { position }) => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::ChartPositionNotFound { position },
        ),
        Error::Refused(Refusal::EmptyChartName) => {
            AxisSupportError::Selector(chart_axis_support::AxisSupportSelectorError::EmptyChartName)
        },
        Error::Limit {
            kind,
            observed,
            maximum,
        } => match kind {
            LimitKind::Allocations => AxisSupportError::Allocation {
                amount: usize::try_from(observed).unwrap_or(usize::MAX),
            },
            kind => AxisSupportError::LimitExceeded {
                kind: match kind {
                    LimitKind::InputBytes => chart_axis_support::AxisSupportLimitKind::InputBytes,
                    LimitKind::OutputBytes => chart_axis_support::AxisSupportLimitKind::OutputBytes,
                    LimitKind::WireBytes => chart_axis_support::AxisSupportLimitKind::WireBytes,
                    LimitKind::Entries => chart_axis_support::AxisSupportLimitKind::Entries,
                    LimitKind::EntryBytes => chart_axis_support::AxisSupportLimitKind::EntryBytes,
                    LimitKind::TotalBytes => chart_axis_support::AxisSupportLimitKind::TotalBytes,
                    LimitKind::Slides => chart_axis_support::AxisSupportLimitKind::Slides,
                    LimitKind::References => chart_axis_support::AxisSupportLimitKind::References,
                    LimitKind::WireFields => chart_axis_support::AxisSupportLimitKind::WireFields,
                    LimitKind::WireNesting => chart_axis_support::AxisSupportLimitKind::WireNesting,
                    LimitKind::WireWork => chart_axis_support::AxisSupportLimitKind::WireWork,
                    LimitKind::ArrangementBytes
                    | LimitKind::RetainedBytes
                    | LimitKind::ScratchBytes
                    | LimitKind::SelectorBytes => {
                        chart_axis_support::AxisSupportLimitKind::TotalBytes
                    },
                    LimitKind::Allocations => unreachable!(),
                },
                observed,
                maximum,
            },
        },
        Error::Allocation { amount } => AxisSupportError::Allocation { amount },
        Error::InvalidSource | Error::Verification | Error::PatchConflict => {
            AxisSupportError::InvalidSource
        },
    }
}

fn map_axis_support_error(error: AxisSupportError) -> Error {
    match error {
        AxisSupportError::Selector(selector) => match selector {
            chart_axis_support::AxisSupportSelectorError::UnsupportedSource => {
                Error::Refused(Refusal::UnsupportedSource)
            },
            chart_axis_support::AxisSupportSelectorError::AmbiguousSelector => {
                Error::Refused(Refusal::AmbiguousSelector)
            },
            chart_axis_support::AxisSupportSelectorError::EmptySlideName => {
                Error::Refused(Refusal::EmptySlideName)
            },
            chart_axis_support::AxisSupportSelectorError::SlideNameNotFound => {
                Error::Refused(Refusal::SlideNameNotFound)
            },
            chart_axis_support::AxisSupportSelectorError::SlidePositionNotFound { position } => {
                Error::Refused(Refusal::SlidePositionNotFound { position })
            },
            chart_axis_support::AxisSupportSelectorError::ChartNameNotFound => {
                Error::Refused(Refusal::ChartNameNotFound)
            },
            chart_axis_support::AxisSupportSelectorError::ChartPositionNotFound { position } => {
                Error::Refused(Refusal::ChartPositionNotFound { position })
            },
            chart_axis_support::AxisSupportSelectorError::EmptyChartName => {
                Error::Refused(Refusal::EmptyChartName)
            },
        },
        AxisSupportError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                chart_axis_support::AxisSupportLimitKind::InputBytes => LimitKind::InputBytes,
                chart_axis_support::AxisSupportLimitKind::OutputBytes => LimitKind::OutputBytes,
                chart_axis_support::AxisSupportLimitKind::WireBytes => LimitKind::WireBytes,
                chart_axis_support::AxisSupportLimitKind::Entries => LimitKind::Entries,
                chart_axis_support::AxisSupportLimitKind::EntryBytes => LimitKind::EntryBytes,
                chart_axis_support::AxisSupportLimitKind::TotalBytes => LimitKind::TotalBytes,
                chart_axis_support::AxisSupportLimitKind::Slides => LimitKind::Slides,
                chart_axis_support::AxisSupportLimitKind::References => LimitKind::References,
                chart_axis_support::AxisSupportLimitKind::TextStorages
                | chart_axis_support::AxisSupportLimitKind::TextFragments
                | chart_axis_support::AxisSupportLimitKind::TextBytes
                | chart_axis_support::AxisSupportLimitKind::TitleBytes => {
                    LimitKind::ArrangementBytes
                },
                chart_axis_support::AxisSupportLimitKind::WireFields => LimitKind::WireFields,
                chart_axis_support::AxisSupportLimitKind::WireNesting => LimitKind::WireNesting,
                chart_axis_support::AxisSupportLimitKind::WireWork => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
        AxisSupportError::Allocation { amount } => Error::Allocation { amount },
        AxisSupportError::InvalidSource => Error::InvalidSource,
    }
}

fn map_chart_title_limit(kind: super::slide_chart_title::LimitKind) -> LimitKind {
    use super::slide_chart_title::LimitKind as ChartTitleLimitKind;
    match kind {
        ChartTitleLimitKind::InputBytes => LimitKind::InputBytes,
        ChartTitleLimitKind::OutputBytes => LimitKind::OutputBytes,
        ChartTitleLimitKind::WireBytes => LimitKind::WireBytes,
        ChartTitleLimitKind::Entries => LimitKind::Entries,
        ChartTitleLimitKind::EntryBytes => LimitKind::EntryBytes,
        ChartTitleLimitKind::TotalBytes => LimitKind::TotalBytes,
        ChartTitleLimitKind::Slides => LimitKind::Slides,
        ChartTitleLimitKind::References => LimitKind::References,
        ChartTitleLimitKind::SelectorBytes => LimitKind::SelectorBytes,
        ChartTitleLimitKind::WireFields => LimitKind::WireFields,
        ChartTitleLimitKind::WireNesting => LimitKind::WireNesting,
        ChartTitleLimitKind::WireWork => LimitKind::WireWork,
        _ => LimitKind::ArrangementBytes,
    }
}

fn map_arrangement_codec_error(error: ChartArrangementDecodeError) -> Error {
    if let Some(limit) = error.resource_limit() {
        return match limit {
            ChartArrangementDecodeLimit::Bytes { observed, maximum } => Error::Limit {
                kind: LimitKind::WireBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartArrangementDecodeLimit::Fields { observed, maximum } => Error::Limit {
                kind: LimitKind::WireFields,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartArrangementDecodeLimit::Work { observed, maximum } => Error::Limit {
                kind: LimitKind::WireWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartArrangementDecodeLimit::Output { observed, maximum } => Error::Limit {
                kind: LimitKind::OutputBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartArrangementDecodeLimit::Nesting { observed, maximum } => Error::Limit {
                kind: LimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            },
            ChartArrangementDecodeLimit::Allocations { observed, maximum } => Error::Limit {
                kind: LimitKind::Allocations,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartArrangementDecodeLimit::Retained { observed, maximum } => Error::Limit {
                kind: LimitKind::RetainedBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartArrangementDecodeLimit::Scratch { observed, maximum } => Error::Limit {
                kind: LimitKind::ScratchBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            _ => Error::InvalidSource,
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            ChartArrangementWireResourceLimit::Bytes { observed, maximum } => Error::Limit {
                kind: LimitKind::WireBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            ChartArrangementWireResourceLimit::Nesting { observed, maximum } => Error::Limit {
                kind: LimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            },
            _ => Error::InvalidSource,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    Error::InvalidSource
}

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                SemanticLimitKind::Objects => LimitKind::Entries,
                SemanticLimitKind::Slides => LimitKind::Slides,
                SemanticLimitKind::References => LimitKind::References,
                SemanticLimitKind::TextStorages
                | SemanticLimitKind::TextFragments
                | SemanticLimitKind::TextBytes => LimitKind::WireFields,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalBytes,
                _ => LimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => LimitKind::TotalBytes,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn changed_verification_accepts_exact_work_and_rejects_one_under()
    -> Result<(), Box<dyn std::error::Error>> {
        let source_bytes = include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/iwork/keynote/chart-arrangement-retirement-resaved.key"
        ));
        let source = Package::from_bytes(source_bytes.as_slice())?;
        let commit = source
            .edit_slide_chart_arrangement(0usize, 0usize)?
            .set(ChartArrangement::default())
            .commit()?;
        let expected_payload = commit
            .patch()
            .target_payload
            .as_deref()
            .ok_or("changed patch has no target payload")?;

        let mut preamble = ArrangementTransactionBudget::new(&source)?;
        preamble.source_package(source.source_bytes().len())?;
        preamble
            .charge_selection_scans(&source, true)
            .map_err(map_axis_support_error)?;
        let selection = select_chart_with_budget(
            &source,
            SlideSelector::position(Position::new(0)),
            ChartSelector::index(0),
            true,
            &mut preamble,
        )
        .map_err(map_axis_support_error)?;
        let (current, _) =
            read_selected_arrangement_with_transaction_budget(&source, &selection, &mut preamble)?;
        assert_eq!(current, commit.patch().before);

        let mut exact = preamble;
        verify_chart_candidate(
            &source,
            commit.package(),
            &selection,
            commit.patch().after,
            expected_payload,
            &mut exact,
        )?;
        assert!(exact.work_used > preamble.work_used);

        let mut exact_ceiling = preamble;
        exact_ceiling.work_limit = exact.work_used;
        verify_chart_candidate(
            &source,
            commit.package(),
            &selection,
            commit.patch().after,
            expected_payload,
            &mut exact_ceiling,
        )?;

        let mut one_under = preamble;
        one_under.work_limit = exact.work_used - 1;
        let error = verify_chart_candidate(
            &source,
            commit.package(),
            &selection,
            commit.patch().after,
            expected_payload,
            &mut one_under,
        )
        .expect_err("verification must consume the complete work envelope");
        assert!(matches!(
            error,
            Error::Limit {
                kind: LimitKind::WireWork,
                ..
            }
        ));
        Ok(())
    }
}
