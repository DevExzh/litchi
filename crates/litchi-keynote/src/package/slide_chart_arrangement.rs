//! Exact-source, selector-first Keynote chart Arrange-panel transactions.
//!
//! The focused owner changes only the two interaction flags stored by the
//! selected chart drawable: `locked` and `aspect_ratio_locked`. Native chart
//! identifiers, archive names, generated protobuf values, and package
//! records remain private to this adapter. The source artifact stays the
//! preservation authority throughout the transaction.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "The package boundary redacts native graph failures."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_chart_arrangement_codec::{
    ChartArrangementWrite, DecodeError as ChartArrangementDecodeError,
    DecodeLimit as ChartArrangementDecodeLimit, DecodeOptions,
    WireResourceLimit as ChartArrangementWireResourceLimit, decode_chart_arrangement,
    decode_chart_arrangement_with_report, prepare_chart_arrangement_rewrite,
};
use thiserror::Error;

use super::slide_chart_title::{ChartSelection, ChartTitleError, select_chart, select_charts};
use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{ChartArrangement, ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const MAX_ARRANGEMENT_CODEC_BYTES: usize = 64 * 1024 * 1024;
const MAX_TRANSACTION_ALLOCATIONS: usize = 64;

/// A finite resource governed by one chart Arrange-panel operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartArrangementLimitKind {
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
}

impl fmt::Display for ChartArrangementLimitKind {
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
        })
    }
}

/// A content-redacted failure raised by a chart Arrange-panel operation.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartArrangementError {
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
    /// The selected chart graph or drawable payload was malformed.
    #[error("the Keynote chart-arrangement source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote chart arrangement {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: ChartArrangementLimitKind,
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

/// One mutable chart Arrange-panel value staged against an immutable package
/// snapshot.
pub struct ChartArrangementEdit<'a> {
    source: &'a Package,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    slide_identifier: u64,
    before: ChartArrangement,
    after: ChartArrangement,
}

impl fmt::Debug for ChartArrangementEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartArrangementEdit")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl<'a> ChartArrangementEdit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Self, ChartArrangementError> {
        let selection = select_chart(source, slide_selector.into(), chart_selector.into(), true)
            .map_err(map_chart_title_error)?;
        let before = read_selected_arrangement(source, &selection)?;
        Ok(Self {
            source,
            slide_position: selection.slide_position,
            chart_position: selection.chart_position,
            chart_identifier: selection.chart_identifier,
            slide_identifier: selection.slide_identifier,
            before,
            after: before,
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

    /// Return the arrangement observed when this edit began.
    #[must_use]
    pub const fn before(&self) -> ChartArrangement {
        self.before
    }

    /// Return the arrangement staged for publication.
    #[must_use]
    pub const fn after(&self) -> ChartArrangement {
        self.after
    }

    /// Stage a complete replacement of the chart Arrange-panel state.
    #[must_use]
    pub const fn set(mut self, arrangement: ChartArrangement) -> Self {
        self.after = arrangement;
        self
    }

    /// Stage the requested lock state while retaining the other flag.
    ///
    /// Requesting `false` preserves an already-absent proto2 field, so an
    /// effective false-to-false edit can remain an exact byte no-op.
    #[must_use]
    pub const fn set_locked(mut self, locked: bool) -> Self {
        self.after = self.after.with_locked(locked);
        self
    }

    /// Stage the requested aspect-ratio constraint while retaining the other
    /// flag.
    ///
    /// Requesting `false` preserves an already-absent proto2 field, so an
    /// effective false-to-false edit can remain an exact byte no-op.
    #[must_use]
    pub const fn set_constrain_proportions(mut self, constrain_proportions: bool) -> Self {
        self.after = self.after.with_constrain_proportions(constrain_proportions);
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<ChartArrangementCommit, ChartArrangementError> {
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let source_selection = select_chart(
            self.source,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            true,
        )
        .map_err(map_chart_title_error)?;
        let current = read_selected_arrangement(self.source, &source_selection)?;
        if source_selection.chart_identifier != self.chart_identifier
            || source_selection.slide_identifier != self.slide_identifier
            || current != self.before
        {
            return Err(ChartArrangementError::InvalidSource);
        }

        if self.before == self.after {
            self.source.validate().map_err(map_read_error)?;
            return Ok(ChartArrangementCommit {
                package: self.source.snapshot(),
                patch: ChartArrangementPatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    slide_position: self.slide_position,
                    chart_position: self.chart_position,
                    chart_identifier: self.chart_identifier,
                    slide_identifier: self.slide_identifier,
                    before: self.before,
                    after: self.after,
                },
                diagnostics: ChartArrangementDiagnostics::unchanged(),
            });
        }

        if !catalog.source_is_exact() {
            return Err(ChartArrangementError::UnsupportedSource);
        }
        self.source.validate().map_err(map_read_error)?;
        let mut budget = ArrangementTransactionBudget::new(self.source)?;
        budget.source_package(self.source.source_bytes().len())?;
        let package =
            rewrite_chart_arrangement(self.source, &source_selection, self.after, &mut budget)?;
        package.validate().map_err(map_read_error)?;
        verify_chart_candidate(
            self.source,
            &package,
            self.slide_position,
            self.chart_position,
            self.chart_identifier,
            self.slide_identifier,
            self.after,
            &mut budget,
        )?;
        let target = physical_catalog(&package)?.shared_source();
        Ok(ChartArrangementCommit {
            package,
            patch: ChartArrangementPatch {
                artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
                slide_position: self.slide_position,
                chart_position: self.chart_position,
                chart_identifier: self.chart_identifier,
                slide_identifier: self.slide_identifier,
                before: self.before,
                after: self.after,
            },
            diagnostics: ChartArrangementDiagnostics::published(),
        })
    }
}

/// An exact-source-checked reversible semantic chart-arrangement patch.
#[derive(Clone, PartialEq, Eq)]
pub struct ChartArrangementPatch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    slide_identifier: u64,
    before: ChartArrangement,
    after: ChartArrangement,
}

impl fmt::Debug for ChartArrangementPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartArrangementPatch")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl ChartArrangementPatch {
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
            slide_identifier: self.slide_identifier,
            before: self.after,
            after: self.before,
        }
    }
}

/// Compact evidence describing one chart-arrangement commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartArrangementDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl ChartArrangementDiagnostics {
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
pub struct ChartArrangementCommit {
    package: Package,
    patch: ChartArrangementPatch,
    diagnostics: ChartArrangementDiagnostics,
}

impl ChartArrangementCommit {
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
    pub const fn patch(&self) -> &ChartArrangementPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &ChartArrangementDiagnostics {
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
    ) -> Result<ChartArrangement, ChartArrangementError> {
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
    ) -> Result<Box<[ChartArrangement]>, ChartArrangementError> {
        let selections =
            select_charts(self, slide_selector.into(), true).map_err(map_chart_title_error)?;
        let limits = self.wire_limits().map_err(map_wire_error)?;
        let mut budget = ArrangementReadBudget::new(limits);
        let mut arrangements = Vec::new();
        arrangements
            .try_reserve_exact(selections.len())
            .map_err(|_error| ChartArrangementError::Allocation {
                amount: selections.len(),
            })?;
        for selection in selections.iter() {
            arrangements.push(read_selected_arrangement_with_budget(
                self,
                selection,
                limits,
                &mut budget,
            )?);
        }
        Ok(arrangements.into_boxed_slice())
    }

    /// Start an exact immutable edit of one selected chart's Arrange-panel
    /// state.
    pub fn edit_slide_chart_arrangement<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartArrangementEdit<'_>, ChartArrangementError> {
        ChartArrangementEdit::new(self, slide_selector, chart_selector)
    }

    /// Apply an exact-source-checked chart-arrangement patch.
    pub fn apply_slide_chart_arrangement(
        &self,
        patch: &ChartArrangementPatch,
    ) -> Result<ChartArrangementCommit, ChartArrangementError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(ChartArrangementError::PatchConflict);
        }
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
            return Err(ChartArrangementError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(ChartArrangementCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: ChartArrangementDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartArrangementError::PatchConflict);
        }
        let mut budget = ArrangementTransactionBudget::new(self)?;
        budget.source_package(self.source_bytes().len())?;
        budget.candidate_reopen(patch.artifacts.target().len())?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        verify_chart_candidate(
            self,
            &candidate,
            patch.slide_position,
            patch.chart_position,
            patch.chart_identifier,
            patch.slide_identifier,
            patch.after,
            &mut budget,
        )?;
        Ok(ChartArrangementCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ChartArrangementDiagnostics::published(),
        })
    }
}

fn read_selected_arrangement(
    package: &Package,
    selection: &ChartSelection,
) -> Result<ChartArrangement, ChartArrangementError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    read_selected_arrangement_with_limits(package, selection, limits)
}

fn read_selected_arrangement_with_limits(
    package: &Package,
    selection: &ChartSelection,
    limits: WireLimits,
) -> Result<ChartArrangement, ChartArrangementError> {
    let (_, object) = package
        .object_with_component(selection.chart_identifier)
        .ok_or(ChartArrangementError::InvalidSource)?;
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
    work_limit: usize,
    retained_limit: usize,
    scratch_limit: usize,
    input_used: usize,
    output_used: usize,
    work_used: usize,
    allocations_used: usize,
    retained_used: usize,
    scratch_used: usize,
}

impl ArrangementTransactionBudget {
    fn new(package: &Package) -> Result<Self, ChartArrangementError> {
        let physical = package.state.options.archive();
        let input_limit = usize::try_from(physical.max_input_bytes())
            .map_err(|_| ChartArrangementError::InvalidSource)?;
        let aggregate_limit = input_limit
            .checked_mul(4)
            .ok_or(ChartArrangementError::InvalidSource)?;
        let wire = package.wire_limits().map_err(map_wire_error)?;
        Ok(Self {
            input_limit: aggregate_limit,
            output_limit: aggregate_limit,
            work_limit: wire.max_rewrite_work(),
            retained_limit: aggregate_limit,
            scratch_limit: aggregate_limit,
            input_used: 0,
            output_used: 0,
            work_used: 0,
            allocations_used: 0,
            retained_used: 0,
            scratch_used: 0,
        })
    }

    fn source_package(&mut self, bytes: usize) -> Result<(), ChartArrangementError> {
        self.input(bytes)?;
        self.work(bytes)
    }

    fn input(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        self.input_used = charge_aggregate_limit(
            ChartArrangementLimitKind::InputBytes,
            self.input_used,
            amount,
            self.input_limit,
        )?;
        Ok(())
    }

    fn output(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        self.output_used = charge_aggregate_limit(
            ChartArrangementLimitKind::OutputBytes,
            self.output_used,
            amount,
            self.output_limit,
        )?;
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        self.work_used = charge_aggregate_limit(
            ChartArrangementLimitKind::WireWork,
            self.work_used,
            amount,
            self.work_limit,
        )?;
        Ok(())
    }

    fn allocations(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        self.allocations_used = charge_aggregate_limit(
            ChartArrangementLimitKind::Allocations,
            self.allocations_used,
            amount,
            MAX_TRANSACTION_ALLOCATIONS,
        )?;
        Ok(())
    }

    fn retained(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        self.retained_used = charge_aggregate_limit(
            ChartArrangementLimitKind::RetainedBytes,
            self.retained_used,
            amount,
            self.retained_limit,
        )?;
        Ok(())
    }

    fn scratch(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        self.scratch_used = charge_aggregate_limit(
            ChartArrangementLimitKind::ScratchBytes,
            self.scratch_used,
            amount,
            self.scratch_limit,
        )?;
        Ok(())
    }

    fn codec(
        &mut self,
        requirements: litchi_iwa_protos::keynote_chart_arrangement_codec::RewriteExecutionRequirements,
    ) -> Result<(), ChartArrangementError> {
        self.output(requirements.output_bytes)?;
        self.work(requirements.work_bytes)?;
        self.allocations(requirements.allocations)?;
        self.retained(requirements.retained_bytes)?;
        self.scratch(requirements.scratch_bytes)
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), ChartArrangementError> {
        self.output(requirements.output_bytes())?;
        self.work(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())
    }

    fn candidate_reopen(&mut self, bytes: usize) -> Result<(), ChartArrangementError> {
        self.input(bytes)?;
        self.work(bytes)?;
        self.allocations(1)?;
        self.retained(bytes)
    }

    fn verify_packages(
        &mut self,
        source_bytes: usize,
        candidate_bytes: usize,
    ) -> Result<(), ChartArrangementError> {
        let amount = source_bytes
            .checked_add(candidate_bytes)
            .ok_or(ChartArrangementError::InvalidSource)?;
        self.work(amount)
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

    fn check_input(self, amount: usize) -> Result<(), ChartArrangementError> {
        check_aggregate_limit(
            ChartArrangementLimitKind::WireBytes,
            self.input_used,
            amount,
            self.input_limit,
        )
    }

    fn charge(
        &mut self,
        report: litchi_iwa_protos::keynote_chart_arrangement_codec::DecodeReport,
    ) -> Result<(), ChartArrangementError> {
        self.input_used = charge_aggregate_limit(
            ChartArrangementLimitKind::WireBytes,
            self.input_used,
            report.source_bytes(),
            self.input_limit,
        )?;
        self.fields_used = charge_aggregate_limit(
            ChartArrangementLimitKind::WireFields,
            self.fields_used,
            report.fields(),
            self.field_limit,
        )?;
        self.work_used = charge_aggregate_limit(
            ChartArrangementLimitKind::WireWork,
            self.work_used,
            report.work_bytes(),
            self.work_limit,
        )?;
        Ok(())
    }
}

fn read_selected_arrangement_with_budget(
    package: &Package,
    selection: &ChartSelection,
    limits: WireLimits,
    budget: &mut ArrangementReadBudget,
) -> Result<ChartArrangement, ChartArrangementError> {
    let (_, object) = package
        .object_with_component(selection.chart_identifier)
        .ok_or(ChartArrangementError::InvalidSource)?;
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
    Ok(ChartArrangement::new(
        snapshot.is_locked(),
        snapshot.is_constrained(),
    ))
}

fn check_aggregate_limit(
    kind: ChartArrangementLimitKind,
    used: usize,
    amount: usize,
    maximum: usize,
) -> Result<(), ChartArrangementError> {
    charge_aggregate_limit(kind, used, amount, maximum).map(|_observed| ())
}

fn charge_aggregate_limit(
    kind: ChartArrangementLimitKind,
    used: usize,
    amount: usize,
    maximum: usize,
) -> Result<usize, ChartArrangementError> {
    let observed =
        used.checked_add(amount)
            .ok_or_else(|| ChartArrangementError::LimitExceeded {
                kind,
                observed: u64::MAX,
                maximum: usize_to_u64(maximum),
            })?;
    if observed > maximum {
        return Err(ChartArrangementError::LimitExceeded {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        });
    }
    Ok(observed)
}

fn rewrite_chart_arrangement(
    source: &Package,
    selection: &ChartSelection,
    replacement: ChartArrangement,
    budget: &mut ArrangementTransactionBudget,
) -> Result<Package, ChartArrangementError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.slide_component_name)
        .ok_or(ChartArrangementError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(ChartArrangementError::InvalidSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.input(entry.data().len())?;
    budget.work(entry.data().len())?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    budget.output(stream.as_bytes().len())?;
    budget.work(stream.as_bytes().len())?;
    budget.allocations(1)?;
    budget.retained(stream.as_bytes().len())?;
    budget.scratch(stream.as_bytes().len())?;
    budget.allocations(1)?;
    budget.retained(stream.as_bytes().len())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    let (message_index, message_data) = {
        let object = archive
            .object(selection.chart_identifier)
            .ok_or(ChartArrangementError::InvalidSource)?;
        let mut index = None;
        for (candidate, message) in object.messages.iter().enumerate() {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            if index.replace(candidate).is_some() {
                return Err(ChartArrangementError::InvalidSource);
            }
        }
        let index = index.ok_or(ChartArrangementError::InvalidSource)?;
        (index, object.messages[index].data.as_slice())
    };
    let limits = source.wire_limits().map_err(map_wire_error)?;
    let patched = patch_chart_arrangement(message_data, replacement, limits, budget)?;
    archive
        .object_mut(selection.chart_identifier)
        .ok_or(ChartArrangementError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: patched,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_bound = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_bound).map_err(map_core_error)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(ChartArrangementError::LimitExceeded {
            kind: ChartArrangementLimitKind::EntryBytes,
            observed: usize_to_u64(compressed_bound),
            maximum: usize_to_u64(snappy_limits.max_compressed_stream()),
        });
    }
    let maximum_entry = usize::try_from(physical_limits.max_entry_bytes())
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    if encoded_bound > maximum_entry || compressed_bound > maximum_entry {
        return Err(ChartArrangementError::LimitExceeded {
            kind: ChartArrangementLimitKind::EntryBytes,
            observed: usize_to_u64(encoded_bound.max(compressed_bound)),
            maximum: usize_to_u64(maximum_entry),
        });
    }
    let package_bound = source
        .source_bytes()
        .len()
        .checked_sub(entry.data().len())
        .and_then(|bytes| bytes.checked_add(compressed_bound))
        .ok_or(ChartArrangementError::InvalidSource)?;
    let maximum_package = usize::try_from(physical_limits.max_input_bytes())
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    if package_bound > maximum_package {
        return Err(ChartArrangementError::LimitExceeded {
            kind: ChartArrangementLimitKind::OutputBytes,
            observed: usize_to_u64(package_bound),
            maximum: usize_to_u64(maximum_package),
        });
    }
    budget.output(encoded_bound)?;
    budget.work(encoded_bound)?;
    budget.allocations(1)?;
    budget.retained(encoded_bound)?;
    budget.scratch(encoded_bound)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.output(compressed_bound)?;
    budget.work(compressed_bound)?;
    budget.allocations(1)?;
    budget.retained(compressed_bound)?;
    budget.scratch(compressed_bound)?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
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
    Package::from_source_with_options(output.into(), source.state.options).map_err(map_read_error)
}

fn patch_chart_arrangement(
    data: &[u8],
    replacement: ChartArrangement,
    limits: WireLimits,
    budget: &mut ArrangementTransactionBudget,
) -> Result<Vec<u8>, ChartArrangementError> {
    let options = arrangement_decode_options(data, limits)?;
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
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    slide_identifier: u64,
    expected: ChartArrangement,
    budget: &mut ArrangementTransactionBudget,
) -> Result<(), ChartArrangementError> {
    budget.verify_packages(source.source_bytes().len(), candidate.source_bytes().len())?;
    if source.state.total_objects != candidate.state.total_objects {
        return Err(ChartArrangementError::Verification);
    }
    let source_selection = select_chart(
        source,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        true,
    )
    .map_err(map_chart_title_error)?;
    let candidate_selection = select_chart(
        candidate,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        true,
    )
    .map_err(map_chart_title_error)?;
    if source_selection.chart_identifier != chart_identifier
        || source_selection.slide_identifier != slide_identifier
        || candidate_selection.chart_identifier != chart_identifier
        || candidate_selection.slide_identifier != slide_identifier
        || source_selection.non_style_identifier != candidate_selection.non_style_identifier
        || source_selection.title != candidate_selection.title
        || source_selection.slide_component_name != candidate_selection.slide_component_name
    {
        return Err(ChartArrangementError::Verification);
    }
    if read_selected_arrangement(candidate, &candidate_selection)? != expected {
        return Err(ChartArrangementError::Verification);
    }
    if source.show().map_err(map_read_error)? != candidate.show().map_err(map_read_error)? {
        return Err(ChartArrangementError::Verification);
    }
    verify_chart_package_locality(
        source,
        candidate,
        &source_selection.slide_component_name,
        chart_identifier,
    )
}

fn verify_chart_package_locality(
    source: &Package,
    candidate: &Package,
    selected_component_name: &str,
    selected_identifier: u64,
) -> Result<(), ChartArrangementError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let mut source_entries = source_catalog.package().iter();
    let mut candidate_entries = candidate_catalog.package().iter();
    loop {
        match (source_entries.next(), candidate_entries.next()) {
            (Some(source_entry), Some(candidate_entry)) => {
                let selected = source_entry.name() == selected_component_name;
                if source_entry.name() != candidate_entry.name()
                    || (!selected && !same_unselected_package_entry(source_entry, candidate_entry))
                    || (selected && !same_selected_package_entry(source_entry, candidate_entry))
                {
                    return Err(ChartArrangementError::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(ChartArrangementError::Verification),
        }
    }

    let source_components = source_catalog.components();
    let candidate_components = candidate_catalog.components();
    if source_components.len() != candidate_components.len() {
        return Err(ChartArrangementError::Verification);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut selected_component_seen = false;
    for (source_component, candidate_component) in
        source_components.iter().zip(candidate_components.iter())
    {
        if source_component.name() != candidate_component.name()
            || source_component.archive().objects.len()
                != candidate_component.archive().objects.len()
        {
            return Err(ChartArrangementError::Verification);
        }
        if source_component.name() != selected_component_name {
            for (source_object, candidate_object) in source_component
                .archive()
                .objects
                .iter()
                .zip(&candidate_component.archive().objects)
            {
                if !source_object.same_content_ignoring_offsets(candidate_object) {
                    return Err(ChartArrangementError::Verification);
                }
            }
            continue;
        }
        selected_component_seen = true;
        verify_selected_chart_component(
            source_component.archive(),
            candidate_component.archive(),
            selected_identifier,
            archive_limits,
        )?;
    }
    if !selected_component_seen {
        return Err(ChartArrangementError::Verification);
    }
    Ok(())
}

fn verify_selected_chart_component(
    source: &Archive,
    candidate: &Archive,
    selected_identifier: u64,
    archive_limits: litchi_iwa_core::ArchiveLimits,
) -> Result<(), ChartArrangementError> {
    let mut selected_seen = false;
    for (source_object, candidate_object) in source.objects.iter().zip(&candidate.objects) {
        if source_object.archive_info.identifier != candidate_object.archive_info.identifier {
            return Err(ChartArrangementError::Verification);
        }
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(ChartArrangementError::Verification)?;
        if identifier != selected_identifier {
            if !source_object.same_content_ignoring_offsets(candidate_object) {
                return Err(ChartArrangementError::Verification);
            }
            continue;
        }
        if std::mem::replace(&mut selected_seen, true) {
            return Err(ChartArrangementError::Verification);
        }
        let selected_message_index = unique_message_index(source_object, CHART_MESSAGE_TYPE)?;
        verify_selected_chart_object(
            source_object,
            candidate_object,
            selected_message_index,
            archive_limits,
        )?;
    }
    if !selected_seen {
        return Err(ChartArrangementError::Verification);
    }
    Ok(())
}

fn verify_selected_chart_object(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    selected_message_index: usize,
    archive_limits: litchi_iwa_core::ArchiveLimits,
) -> Result<(), ChartArrangementError> {
    if source.archive_info.identifier != candidate.archive_info.identifier
        || source.archive_info.should_merge != candidate.archive_info.should_merge
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.message_infos.len() != candidate.archive_info.message_infos.len()
    {
        return Err(ChartArrangementError::Verification);
    }
    for (index, ((source_message, candidate_message), (source_info, candidate_info))) in source
        .messages
        .iter()
        .zip(&candidate.messages)
        .zip(
            source
                .archive_info
                .message_infos
                .iter()
                .zip(&candidate.archive_info.message_infos),
        )
        .enumerate()
    {
        if index == selected_message_index {
            if source_message.type_ != candidate_message.type_
                || !message_info_equal_except_length(source_info, candidate_info)
            {
                return Err(ChartArrangementError::Verification);
            }
        } else if source_message != candidate_message || source_info != candidate_info {
            return Err(ChartArrangementError::Verification);
        }
    }

    let candidate_message = candidate
        .messages
        .get(selected_message_index)
        .ok_or(ChartArrangementError::Verification)?
        .clone();
    let mut expected = source.clone();
    expected
        .replace_message_preserving_header_with_limits(
            selected_message_index,
            candidate_message,
            archive_limits,
        )
        .map_err(map_core_error)?;
    expected.header_length = candidate.header_length;
    expected.data_length = candidate.data_length;
    if !expected.same_content_ignoring_offsets(candidate) {
        return Err(ChartArrangementError::Verification);
    }
    Ok(())
}

fn unique_message_index(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<usize, ChartArrangementError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(index).is_some() {
            return Err(ChartArrangementError::Verification);
        }
    }
    selected.ok_or(ChartArrangementError::Verification)
}

fn exactly_one_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<&RawMessage, ChartArrangementError> {
    let mut selected = None;
    for message in &object.messages {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(message).is_some() {
            return Err(ChartArrangementError::InvalidSource);
        }
    }
    selected.ok_or(ChartArrangementError::InvalidSource)
}

fn same_unselected_package_entry(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.data() == candidate.data()
        && source.metadata() == candidate.metadata()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && source.raw_record().compressed_data() == candidate.raw_record().compressed_data()
        && same_central_record_except_offset(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn same_selected_package_entry(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    if source.name() != candidate.name()
        || source.raw_name() != candidate.raw_name()
        || source.is_opaque() != candidate.is_opaque()
        || !same_header_metadata(source.metadata().local(), candidate.metadata().local())
        || !same_header_metadata(source.metadata().central(), candidate.metadata().central())
    {
        return false;
    }
    compatible_local_zip_record(source, candidate)
        && compatible_central_zip_record(source, candidate)
}

fn same_header_metadata(
    source: &litchi_iwa_archive::package::HeaderMetadata,
    candidate: &litchi_iwa_archive::package::HeaderMetadata,
) -> bool {
    source.version_needed() == candidate.version_needed()
        && source.flags() == candidate.flags()
        && source.compression_method() == candidate.compression_method()
        && source.last_modified() == candidate.last_modified()
        && source.name() == candidate.name()
        && source.extra() == candidate.extra()
        && source.comment() == candidate.comment()
}

fn compatible_local_zip_record(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    let source_record = source.raw_record().local_record();
    let candidate_record = candidate.raw_record().local_record();
    if source_record.len() < 30 || candidate_record.len() < 30 {
        return false;
    }
    let Some(source_name_len) = read_u16(source_record, 26) else {
        return false;
    };
    let Some(source_extra_len) = read_u16(source_record, 28) else {
        return false;
    };
    let Some(candidate_name_len) = read_u16(candidate_record, 26) else {
        return false;
    };
    let Some(candidate_extra_len) = read_u16(candidate_record, 28) else {
        return false;
    };
    let source_header_len = 30usize
        .checked_add(usize::from(source_name_len))
        .and_then(|length| length.checked_add(usize::from(source_extra_len)));
    let candidate_header_len = 30usize
        .checked_add(usize::from(candidate_name_len))
        .and_then(|length| length.checked_add(usize::from(candidate_extra_len)));
    let (Some(source_header_len), Some(candidate_header_len)) =
        (source_header_len, candidate_header_len)
    else {
        return false;
    };
    if source_header_len != candidate_header_len
        || source_header_len > source_record.len()
        || candidate_header_len > candidate_record.len()
    {
        return false;
    }
    let Some(source_suffix_start) =
        source_header_len.checked_add(source.raw_record().compressed_data().len())
    else {
        return false;
    };
    let Some(candidate_suffix_start) =
        candidate_header_len.checked_add(candidate.raw_record().compressed_data().len())
    else {
        return false;
    };
    if source_suffix_start > source_record.len() || candidate_suffix_start > candidate_record.len()
    {
        return false;
    }
    if !equal_except_ranges(
        &source_record[..source_header_len],
        &candidate_record[..candidate_header_len],
        std::slice::from_ref(&(14..26)),
    ) {
        return false;
    }
    descriptor_shape_compatible(
        source.metadata().local().flags(),
        &source_record[source_suffix_start..],
        &candidate_record[candidate_suffix_start..],
    )
}

fn compatible_central_zip_record(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    let source_record = source.raw_record().central_directory_record();
    let candidate_record = candidate.raw_record().central_directory_record();
    if source_record.len() < 46
        || candidate_record.len() != source_record.len()
        || read_u16(source_record, 28) != read_u16(candidate_record, 28)
        || read_u16(source_record, 30) != read_u16(candidate_record, 30)
        || read_u16(source_record, 32) != read_u16(candidate_record, 32)
    {
        return false;
    }
    equal_except_ranges(source_record, candidate_record, &[(16..28), (42..46)])
}

fn descriptor_shape_compatible(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source.is_empty() && candidate.is_empty();
    }
    if source.len() != candidate.len() || !matches!(source.len(), 12 | 16) {
        return false;
    }
    if source.len() == 16 {
        source[..4] == candidate[..4]
            && read_u32(source, 0) == Some(0x0807_4b50)
            && read_u32(candidate, 0) == Some(0x0807_4b50)
    } else {
        true
    }
}

fn equal_except_ranges(
    source: &[u8],
    candidate: &[u8],
    ignored: &[std::ops::Range<usize>],
) -> bool {
    source.len() == candidate.len()
        && source.iter().enumerate().all(|(index, byte)| {
            ignored.iter().any(|range| range.contains(&index)) || Some(byte) == candidate.get(index)
        })
}

fn same_central_record_except_offset(source: &[u8], candidate: &[u8]) -> bool {
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= LOCAL_HEADER_OFFSET.end
        && source[..LOCAL_HEADER_OFFSET.start] == candidate[..LOCAL_HEADER_OFFSET.start]
        && source[LOCAL_HEADER_OFFSET.end..] == candidate[LOCAL_HEADER_OFFSET.end..]
}

fn read_u16(bytes: &[u8], offset: usize) -> Option<u16> {
    let value = bytes.get(offset..offset.checked_add(2)?)?;
    Some(u16::from_le_bytes([value[0], value[1]]))
}

fn read_u32(bytes: &[u8], offset: usize) -> Option<u32> {
    let value = bytes.get(offset..offset.checked_add(4)?)?;
    Some(u32::from_le_bytes([value[0], value[1], value[2], value[3]]))
}

fn message_info_equal_except_length(
    source: &litchi_iwa_core::MessageInfo,
    candidate: &litchi_iwa_core::MessageInfo,
) -> bool {
    source.type_ == candidate.type_
        && source.versions == candidate.versions
        && source.field_infos == candidate.field_infos
        && source.object_references == candidate.object_references
        && source.data_references == candidate.data_references
        && source.base_message_index == candidate.base_message_index
        && source.diff_merge_version == candidate.diff_merge_version
        && source.diff_field_path == candidate.diff_field_path
        && source.fields_to_remove == candidate.fields_to_remove
        && source.diff_read_version == candidate.diff_read_version
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), ChartArrangementError> {
    for object in &archive.objects {
        let offset = usize::try_from(object.header_offset)
            .map_err(|_| ChartArrangementError::InvalidSource)?;
        let remaining = source
            .get(offset..)
            .ok_or(ChartArrangementError::InvalidSource)?;
        let (header_bytes, prefix_bytes) = litchi_iwa_common::decode_varint_from_bytes(remaining)
            .map_err(|_| ChartArrangementError::InvalidSource)?;
        if prefix_bytes != litchi_iwa_common::varint::encoded_len(header_bytes) {
            return Err(ChartArrangementError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(
                u64::try_from(prefix_bytes).map_err(|_| ChartArrangementError::InvalidSource)?,
            )
            .ok_or(ChartArrangementError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(ChartArrangementError::InvalidSource)?
                != object.data_offset
        {
            return Err(ChartArrangementError::InvalidSource);
        }
    }
    Ok(())
}

fn arrangement_decode_options(
    source: &[u8],
    limits: WireLimits,
) -> Result<DecodeOptions, ChartArrangementError> {
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
) -> Result<DecodeOptions, ChartArrangementError> {
    let recursion_limit =
        u32::try_from(limits.max_nesting()).map_err(|_| ChartArrangementError::LimitExceeded {
            kind: ChartArrangementLimitKind::WireNesting,
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

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, ChartArrangementError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(ChartArrangementError::UnsupportedSource),
    }
}

fn map_chart_title_error(error: ChartTitleError) -> ChartArrangementError {
    match error {
        ChartTitleError::UnsupportedSource => ChartArrangementError::UnsupportedSource,
        ChartTitleError::AmbiguousSelector => ChartArrangementError::AmbiguousSelector,
        ChartTitleError::EmptySlideName => ChartArrangementError::EmptySlideName,
        ChartTitleError::SlideNameNotFound => ChartArrangementError::SlideNameNotFound,
        ChartTitleError::SlidePositionNotFound { position } => {
            ChartArrangementError::SlidePositionNotFound { position }
        },
        ChartTitleError::ChartNameNotFound => ChartArrangementError::ChartNameNotFound,
        ChartTitleError::ChartPositionNotFound { position } => {
            ChartArrangementError::ChartPositionNotFound { position }
        },
        ChartTitleError::EmptyChartName => ChartArrangementError::EmptyChartName,
        ChartTitleError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ChartArrangementError::LimitExceeded {
            kind: map_chart_title_limit(kind),
            observed,
            maximum,
        },
        ChartTitleError::Allocation { amount } => ChartArrangementError::Allocation { amount },
        _ => ChartArrangementError::InvalidSource,
    }
}

fn map_chart_title_limit(
    kind: super::slide_chart_title::ChartTitleLimitKind,
) -> ChartArrangementLimitKind {
    use super::slide_chart_title::ChartTitleLimitKind;
    match kind {
        ChartTitleLimitKind::InputBytes => ChartArrangementLimitKind::InputBytes,
        ChartTitleLimitKind::OutputBytes => ChartArrangementLimitKind::OutputBytes,
        ChartTitleLimitKind::WireBytes => ChartArrangementLimitKind::WireBytes,
        ChartTitleLimitKind::Entries => ChartArrangementLimitKind::Entries,
        ChartTitleLimitKind::EntryBytes => ChartArrangementLimitKind::EntryBytes,
        ChartTitleLimitKind::TotalBytes => ChartArrangementLimitKind::TotalBytes,
        ChartTitleLimitKind::Slides => ChartArrangementLimitKind::Slides,
        ChartTitleLimitKind::References => ChartArrangementLimitKind::References,
        ChartTitleLimitKind::WireFields => ChartArrangementLimitKind::WireFields,
        ChartTitleLimitKind::WireNesting => ChartArrangementLimitKind::WireNesting,
        ChartTitleLimitKind::WireWork => ChartArrangementLimitKind::WireWork,
        _ => ChartArrangementLimitKind::ArrangementBytes,
    }
}

fn map_arrangement_codec_error(error: ChartArrangementDecodeError) -> ChartArrangementError {
    if let Some(limit) = error.resource_limit() {
        return match limit {
            ChartArrangementDecodeLimit::Bytes { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartArrangementDecodeLimit::Fields { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::WireFields,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartArrangementDecodeLimit::Work { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::WireWork,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartArrangementDecodeLimit::Output { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::OutputBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartArrangementDecodeLimit::Nesting { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            ChartArrangementDecodeLimit::Allocations { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::Allocations,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartArrangementDecodeLimit::Retained { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::RetainedBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartArrangementDecodeLimit::Scratch { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::ScratchBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            _ => ChartArrangementError::InvalidSource,
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            ChartArrangementWireResourceLimit::Bytes { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartArrangementWireResourceLimit::Nesting { observed, maximum } => {
                ChartArrangementError::LimitExceeded {
                    kind: ChartArrangementLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => ChartArrangementError::InvalidSource,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return ChartArrangementError::Allocation { amount };
    }
    ChartArrangementError::InvalidSource
}

fn map_read_error(error: ReadError) -> ChartArrangementError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartArrangementError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => ChartArrangementLimitKind::Entries,
                SemanticLimitKind::Slides => ChartArrangementLimitKind::Slides,
                SemanticLimitKind::References => ChartArrangementLimitKind::References,
                SemanticLimitKind::TextStorages
                | SemanticLimitKind::TextFragments
                | SemanticLimitKind::TextBytes => ChartArrangementLimitKind::WireFields,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartArrangementError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => ChartArrangementLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => ChartArrangementLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => ChartArrangementLimitKind::WireNesting,
                super::PayloadLimitKind::Work => ChartArrangementLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => ChartArrangementError::Allocation { amount },
        _ => ChartArrangementError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> ChartArrangementError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartArrangementError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => ChartArrangementLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    ChartArrangementLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => ChartArrangementLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes => ChartArrangementLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => ChartArrangementLimitKind::TotalBytes,
                _ => ChartArrangementLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            ChartArrangementError::Allocation { amount }
        },
        _ => ChartArrangementError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> ChartArrangementError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartArrangementError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => ChartArrangementLimitKind::TotalBytes,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => ChartArrangementLimitKind::Entries,
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    ChartArrangementLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::HeaderFields => ChartArrangementLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => ChartArrangementLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyFrames => ChartArrangementLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            ChartArrangementError::Allocation { amount: requested }
        },
        _ => ChartArrangementError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> ChartArrangementError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ChartArrangementError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => ChartArrangementLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => ChartArrangementLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    ChartArrangementLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => ChartArrangementLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => ChartArrangementLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            ChartArrangementError::Allocation { amount }
        },
        _ => ChartArrangementError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
