//! Selector-first, archive-free Keynote chart data reads.
//!
//! The chart graph is resolved by the existing chart-title authority. This
//! adapter only owns the final semantic copy of the selected rectangular grid;
//! native identifiers, archive objects, and generated protobuf values remain
//! private to the package and protocol crates.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "The package boundary maps native failures to content-redacted semantic errors."
)]

use std::{
    fmt,
    mem::size_of,
    sync::{Arc, OnceLock},
};

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::chart::data::{ChartData, DataError};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::chart_data_codec::{self, DecodeOptions, LabelList};
use thiserror::Error as ThisError;

use super::Package;
use super::chart_axis_support::{self, AxisSupportBudget, AxisSupportError};
use super::edit::{OwnedChartSelector, OwnedSlideSelector, SelectorError};
use super::slide_chart_title::{
    self, ChartGraphScanBudget, ChartSelection, Error as ChartTitleError,
    LimitKind as ChartTitleLimitKind, select_chart_with_budget,
};
use crate::{ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = chart_data_codec::MODERN_CHART_DRAWABLE_MESSAGE_TYPE;
const MAX_DATA_CELL_COUNT: usize = 1_000_000;
const MAX_DATA_LABEL_COUNT: usize = 1_000_000;
const MAX_DATA_TEXT_BYTES: usize = 64 * 1024 * 1024;
const MAX_DATA_CODEC_BYTES: usize = 64 * 1024 * 1024;

/// A finite resource governed while one chart-data view is prepared.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// Native payload bytes inspected by the selector.
    PayloadBytes,
    /// Native references inspected by the selector.
    PayloadReferences,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate graph and codec work.
    WireWork,
    /// Decoder allocation units.
    WireAllocations,
    /// Decoder-retained bytes.
    WireRetainedBytes,
    /// Numeric grid cells visited.
    Cells,
    /// Semantic slide count.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic text-storage objects.
    TextStorages,
    /// Semantic text fragments.
    TextFragments,
    /// Aggregate semantic text bytes.
    TextBytes,
    /// ZIP/IWA entries and objects.
    Entries,
    /// Bytes in one ZIP/IWA entry or message.
    EntryBytes,
    /// Aggregate retained package bytes.
    TotalBytes,
    /// Allocations needed by the data view.
    Allocations,
    /// Bytes retained by the data view.
    RetainedBytes,
    /// Caller bytes retained by the deferred selectors.
    SelectorBytes,
    /// Number of borrowed row/column labels.
    LabelCount,
    /// Candidate output bytes.
    OutputBytes,
    /// Candidate scratch bytes.
    ScratchBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::WireBytes => "wire bytes",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadReferences => "payload references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::WireAllocations => "wire allocations",
            Self::WireRetainedBytes => "wire retained bytes",
            Self::Cells => "cells",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::SelectorBytes => "selector bytes",
            Self::LabelCount => "label count",
            Self::OutputBytes => "output bytes",
            Self::ScratchBytes => "scratch bytes",
        })
    }
}

/// A content-redacted failure raised by a chart-data read.
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The selected chart graph or data payload was malformed.
    #[error("the Keynote chart data source is malformed or unsupported")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote chart data {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category that exceeded its limit.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded semantic allocation failed before the result was published.
    #[error("could not allocate {amount} units for Keynote chart data")]
    Allocation { amount: usize },
    /// Full candidate reopening did not reproduce the requested state.
    #[error("the edited Keynote chart data failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote chart-data patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// The source was not retained as an exact physical package.
    #[error("this Keynote source does not support physical chart data reads")]
    UnsupportedSource,
    /// An exact-name selector was ambiguous.
    #[error("the Keynote chart data selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound { position: Position },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// The requested edit changes one or more row or column labels.
    #[error("Keynote chart data edits must preserve row and column labels")]
    LabelsChanged,
    /// The requested edit changes the rectangular grid dimensions.
    #[error("Keynote chart data edits must preserve chart dimensions")]
    ShapeChanged,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. }
            | Self::LabelsChanged { .. }
            | Self::ShapeChanged { .. } => RefusalKind::UnsupportedContent,
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

/// One mutable semantic chart-data value staged against an immutable Keynote
/// package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    slide: OwnedSlideSelector,
    chart: OwnedChartSelector,
    selection: OnceLock<(ChartSelection, Arc<ChartData>)>,
    after: Option<ChartData>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("has_selection", &self.selection.get().is_some())
            .field("has_staged", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    fn resolve(&self) -> Result<&(ChartSelection, Arc<ChartData>), Error> {
        if let Some(selection) = self.selection.get() {
            return Ok(selection);
        }
        let mut budget = ChartGraphScanBudget::new(self.source).map_err(map_chart_title_error)?;
        budget
            .charge_selection_scans(self.source, true)
            .map_err(map_axis_support_error)?;
        let selection = select_chart_with_budget(
            self.source,
            self.slide.as_selector(),
            self.chart.as_selector(),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        let before = read_selected_data(self.source, selection.clone(), &mut budget)?;
        preflight_data_budget(&budget, LimitKind::Allocations, 1)?;
        budget
            .charge_work(size_of::<ChartData>())
            .map_err(map_axis_support_error)?;
        let _ = self.selection.set((selection, Arc::new(before)));
        self.selection.get().ok_or(Error::InvalidSource)
    }

    /// Return the selected semantic slide position.
    #[must_use = "handle the deferred chart-data result"]
    pub fn slide_position(&self) -> Result<Position, Error> {
        Ok(self.resolve()?.0.slide_position)
    }

    /// Return the selected semantic chart position.
    #[must_use = "handle the deferred chart-data result"]
    pub fn chart_position(&self) -> Result<Position, Error> {
        Ok(self.resolve()?.0.chart_position)
    }

    /// Return the chart data observed when this edit first resolves.
    #[must_use = "handle the deferred chart-data result"]
    pub fn before(&self) -> Result<&ChartData, Error> {
        Ok(self.resolve()?.1.as_ref())
    }

    /// Return the chart data currently staged for publication.
    #[must_use = "handle the deferred chart-data result"]
    pub fn after(&self) -> Result<&ChartData, Error> {
        match self.after.as_ref() {
            Some(after) => Ok(after),
            None => Ok(self.resolve()?.1.as_ref()),
        }
    }

    /// Replace the staged chart data.
    #[must_use]
    pub fn set(mut self, data: ChartData) -> Self {
        // Keep the caller-owned value by value until commit can charge its
        // complete logical footprint.  In particular, do not hide an
        // infallible Arc allocation behind this fluent setter.
        self.after = Some(data);
        self
    }

    /// Validate and publish the staged exact-source edit.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_data_edit(self)
    }
}

/// Exact-source-checked reversible chart-data patch.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    slide_identifier: u64,
    before: Arc<ChartData>,
    after: Arc<ChartData>,
    source_payload: Option<Arc<[u8]>>,
    target_payload: Option<Arc<[u8]>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", self.before.as_ref())
            .field("after", self.after.as_ref())
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Return the semantic data required from the source package.
    #[must_use]
    pub fn before(&self) -> &ChartData {
        self.before.as_ref()
    }

    /// Return the semantic data produced by this patch.
    #[must_use]
    pub fn after(&self) -> &ChartData {
        self.after.as_ref()
    }

    /// Return the exact source artifact fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the exact target artifact fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether the patch is an exact byte and semantic no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        chart_data_bits_equal(self.before.as_ref(), self.after.as_ref())
            && self.artifacts.is_byte_noop()
            && self.source_payload.is_none()
            && self.target_payload.is_none()
    }

    /// Return the exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            chart_position: self.chart_position,
            chart_identifier: self.chart_identifier,
            slide_identifier: self.slide_identifier,
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            source_payload: self.target_payload.as_ref().map(Arc::clone),
            target_payload: self.source_payload.as_ref().map(Arc::clone),
        }
    }
}

/// Compact evidence describing one committed chart-data transaction.
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

    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Return whether the package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of deleted root rendering previews.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully reopened immutable result of one chart-data transaction.
#[must_use = "a Keynote chart-data commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the fully reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package snapshot.
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

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one selected chart's modern inline rectangular grid.
    ///
    /// Numeric cells are returned as finite `f64` values and missing,
    /// date-only, or duration-only native cells become `None`. Legacy chart
    /// payloads and graphs that do not prove the modern inline grid are
    /// rejected as [`Error::InvalidSource`].
    pub fn slide_chart_data<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartData, Error> {
        let mut budget = ChartGraphScanBudget::new(self).map_err(map_chart_title_error)?;
        budget
            .charge_selection_scans(self, true)
            .map_err(map_axis_support_error)?;
        let selection = select_chart_with_budget(
            self,
            slide_selector.into(),
            chart_selector.into(),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        read_selected_data(self, selection, &mut budget)
    }

    /// Begin a selector-first immutable chart-data edit.
    ///
    /// Entry retains only the source and the archive-free selectors. Chart
    /// graph selection and data decoding are deferred until a value is read
    /// or the edit is committed. Selector names are admitted against the
    /// package's finite selector-byte limit at entry.
    pub fn edit_slide_chart_data<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Edit<'_>, Error> {
        let maximum_selector_bytes = self.semantic_limits().max_text_bytes();
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
        Ok(Edit {
            source: self,
            slide,
            chart,
            selection: OnceLock::new(),
            after: None,
        })
    }

    /// Apply an exact-source-checked reversible chart-data patch.
    pub fn apply_slide_chart_data(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        let mut budget = ChartGraphScanBudget::new(self).map_err(map_chart_title_error)?;
        // Exact source authorization may walk every source byte when the
        // patch came from a distinct Arc. Reserve that fingerprint comparison
        // before entering the authorization call.
        budget
            .charge_work(source.len())
            .map_err(map_axis_support_error)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
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
        if selection.chart_identifier != patch.chart_identifier
            || selection.slide_identifier != patch.slide_identifier
        {
            return Err(Error::PatchConflict);
        }
        let current = read_selected_data(self, selection.clone(), &mut budget)?;
        if !compare_chart_data_with_budget(&current, patch.before.as_ref(), &mut budget)? {
            return Err(Error::PatchConflict);
        }
        if let Some(expected) = patch.source_payload.as_deref() {
            let actual = selected_chart_payload(self, &selection, &mut budget)?;
            if !compare_payload_with_budget(actual, expected, &mut budget)? {
                return Err(Error::PatchConflict);
            }
        }
        if patch_is_noop_with_budget(patch, &mut budget)? {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        validate_data_change(patch.before.as_ref(), patch.after.as_ref(), &mut budget)?;
        let target = patch
            .target_payload
            .as_deref()
            .ok_or(Error::PatchConflict)?;
        let target_bytes = patch.artifacts.target();
        budget
            .charge_work(target_bytes.len())
            .map_err(map_axis_support_error)?;
        let candidate = Package::from_source_with_options(target_bytes, self.state.options)
            .map_err(|_error| Error::Verification)?;
        let candidate_selection = select_chart_with_budget(
            &candidate,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        if candidate_selection.chart_identifier != patch.chart_identifier
            || candidate_selection.slide_identifier != patch.slide_identifier
        {
            return Err(Error::Verification);
        }
        let candidate_data =
            read_selected_data(&candidate, candidate_selection.clone(), &mut budget)?;
        if !compare_chart_data_with_budget(&candidate_data, patch.after.as_ref(), &mut budget)? {
            return Err(Error::Verification);
        }
        let actual_target = selected_chart_payload(&candidate, &candidate_selection, &mut budget)?;
        if !compare_payload_with_budget(actual_target, target, &mut budget)? {
            return Err(Error::Verification);
        }
        verify_chart_candidate(
            self,
            &candidate,
            &selection,
            patch.target_payload.as_deref(),
            false,
            &mut budget,
        )?;
        let deleted_previews = preview_count_removed(self, &candidate, &mut budget)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(deleted_previews),
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

fn commit_data_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let (selection_base, before) = edit.resolve()?.clone();
    let source = edit.source;
    let after_value = edit.after;
    let slide_position = selection_base.slide_position;
    let chart_position = selection_base.chart_position;
    let chart_identifier = selection_base.chart_identifier;
    let slide_identifier = selection_base.slide_identifier;
    let catalog = physical_catalog(source)?;
    let source_bytes = catalog.shared_source();
    let mut budget = ChartGraphScanBudget::new(source).map_err(map_chart_title_error)?;
    let after_is_noop = match after_value.as_ref() {
        Some(after) => compare_chart_data_with_budget(before.as_ref(), after, &mut budget)?,
        None => true,
    };

    if after_is_noop {
        budget
            .charge_selection_scans(source, true)
            .map_err(map_axis_support_error)?;
        let selection = select_chart_with_budget(
            source,
            SlideSelector::position(slide_position),
            ChartSelector::index(chart_position.get()),
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        let current = read_selected_data(source, selection.clone(), &mut budget)?;
        if selection.chart_identifier != chart_identifier
            || selection.slide_identifier != slide_identifier
            || !compare_chart_data_with_budget(&current, before.as_ref(), &mut budget)?
        {
            return Err(Error::PatchConflict);
        }
        return Ok(Commit {
            package: source.snapshot(),
            patch: Patch {
                artifacts: exact_artifacts_with_budget(
                    Arc::clone(&source_bytes),
                    source_bytes,
                    &mut budget,
                )?,
                slide_position,
                chart_position,
                chart_identifier,
                slide_identifier,
                before: Arc::clone(&before),
                after: Arc::clone(&before),
                source_payload: None,
                target_payload: None,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }

    let after = after_value.ok_or(Error::InvalidSource)?;
    validate_data_change(before.as_ref(), &after, &mut budget)?;
    if !catalog.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }

    charge_chart_data_ownership(&after, &mut budget)?;
    budget
        .charge_selection_scans(source, true)
        .map_err(map_axis_support_error)?;
    let selection = select_chart_with_budget(
        source,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        true,
        &mut budget,
    )
    .map_err(map_axis_support_error)?;
    if selection.chart_identifier != chart_identifier
        || selection.slide_identifier != slide_identifier
    {
        return Err(Error::PatchConflict);
    }
    let current = read_selected_data(source, selection.clone(), &mut budget)?;
    if !compare_chart_data_with_budget(&current, before.as_ref(), &mut budget)? {
        return Err(Error::PatchConflict);
    }
    let source_payload = selected_chart_payload(source, &selection, &mut budget)?;
    let options = data_options(&budget, source_payload)?;
    let prepared =
        chart_data_codec::prepare_chart_data_rewrite(source_payload, after.values(), options)
            .map_err(map_data_rewrite_error)?;
    charge_data_report(&mut budget, prepared.prepare_report())?;
    let requirements = prepared.execution_requirements();
    charge_rewrite_requirements(&mut budget, requirements)?;
    let rewritten = prepared
        .execute(requirements.exact_limits())
        .map_err(map_data_rewrite_error)?
        .into_output();

    let (candidate, deleted_previews) =
        rewrite_chart_payload(source, &selection, rewritten, &mut budget)?;
    let candidate_selection = select_chart_with_budget(
        &candidate,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        true,
        &mut budget,
    )
    .map_err(map_axis_support_error)?;
    if candidate_selection.chart_identifier != chart_identifier
        || candidate_selection.slide_identifier != slide_identifier
    {
        return Err(Error::Verification);
    }
    let candidate_data = read_selected_data(&candidate, candidate_selection.clone(), &mut budget)?;
    if !compare_chart_data_with_budget(&candidate_data, &after, &mut budget)? {
        return Err(Error::Verification);
    }
    let target_payload = selected_chart_payload(&candidate, &candidate_selection, &mut budget)?;
    verify_chart_candidate(
        source,
        &candidate,
        &selection,
        Some(target_payload),
        true,
        &mut budget,
    )?;
    let (source_payload, target_payload) =
        retain_payload_pair(source_payload, target_payload, &mut budget)?;
    let target_bytes = physical_catalog(&candidate)?.shared_source();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: exact_artifacts_with_budget(source_bytes, target_bytes, &mut budget)?,
            slide_position,
            chart_position,
            chart_identifier,
            slide_identifier,
            before,
            after: Arc::new(after),
            source_payload: Some(source_payload),
            target_payload: Some(target_payload),
        },
        diagnostics: Diagnostics::published(deleted_previews),
    })
}

fn validate_data_change(
    before: &ChartData,
    after: &ChartData,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), Error> {
    charge_chart_data_compare(before, after, budget)?;
    if before.row_names() != after.row_names() || before.column_names() != after.column_names() {
        return Err(Error::Refused(Refusal::LabelsChanged));
    }
    if before.values().len() != after.values().len()
        || before
            .values()
            .iter()
            .zip(after.values())
            .any(|(left, right)| left.len() != right.len())
    {
        return Err(Error::Refused(Refusal::ShapeChanged));
    }
    Ok(())
}

fn chart_data_bits_equal(left: &ChartData, right: &ChartData) -> bool {
    left.bitwise_eq(right)
}

fn compare_chart_data_with_budget(
    left: &ChartData,
    right: &ChartData,
    budget: &mut ChartGraphScanBudget,
) -> Result<bool, Error> {
    charge_chart_data_compare(left, right, budget)?;
    Ok(chart_data_bits_equal(left, right))
}

fn charge_chart_data_compare(
    left: &ChartData,
    right: &ChartData,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), Error> {
    let amount = chart_data_compare_work(left)?
        .checked_add(chart_data_compare_work(right)?)
        .ok_or(Error::InvalidSource)?;
    preflight_data_budget(budget, LimitKind::WireWork, amount)?;
    budget.charge_work(amount).map_err(map_axis_support_error)
}

fn chart_data_compare_work(data: &ChartData) -> Result<usize, Error> {
    let mut amount = data
        .row_names()
        .len()
        .checked_add(data.column_names().len())
        .and_then(|value| value.checked_add(data.values().len()))
        .ok_or(Error::InvalidSource)?;
    for label in data.row_names().iter().chain(data.column_names()) {
        amount = amount
            .checked_add(label.len())
            .ok_or(Error::InvalidSource)?;
    }
    for row in data.values() {
        amount = amount.checked_add(row.len()).ok_or(Error::InvalidSource)?;
    }
    Ok(amount)
}

fn compare_payload_with_budget(
    left: &[u8],
    right: &[u8],
    budget: &mut ChartGraphScanBudget,
) -> Result<bool, Error> {
    let amount = left
        .len()
        .checked_add(right.len())
        .ok_or(Error::InvalidSource)?;
    preflight_data_budget(budget, LimitKind::WireWork, amount)?;
    budget.charge_work(amount).map_err(map_axis_support_error)?;
    Ok(left == right)
}

fn patch_is_noop_with_budget(
    patch: &Patch,
    budget: &mut ChartGraphScanBudget,
) -> Result<bool, Error> {
    if !compare_chart_data_with_budget(patch.before.as_ref(), patch.after.as_ref(), budget)? {
        return Ok(false);
    }
    let source = patch.artifacts.source();
    let target = patch.artifacts.target();
    if !compare_payload_with_budget(source.as_ref(), target.as_ref(), budget)? {
        return Ok(false);
    }
    Ok(patch.source_payload.is_none() && patch.target_payload.is_none())
}

fn exact_artifacts_with_budget(
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    budget: &mut ChartGraphScanBudget,
) -> Result<ExactArtifacts, Error> {
    let amount = if Arc::ptr_eq(&source, &target) {
        source.len()
    } else {
        source
            .len()
            .checked_add(target.len())
            .ok_or(Error::InvalidSource)?
    };
    preflight_data_budget(budget, LimitKind::WireWork, amount)?;
    budget.charge_work(amount).map_err(map_axis_support_error)?;
    Ok(ExactArtifacts::new(source, target))
}

/// Charge the complete owned semantic value before it is retained by a
/// transaction patch. `ChartData` intentionally exposes borrowed slices, so
/// the retained-byte policy charges logical string/row lengths plus their
/// fixed `String`/`Vec` slots. The two `usize` words account conservatively for
/// the Arc control block surrounding the model; no allocator capacity is
/// treated as caller-visible retained content.
fn charge_chart_data_ownership(
    data: &ChartData,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), Error> {
    let mut retained = size_of::<ChartData>()
        .checked_add(
            2usize
                .checked_mul(size_of::<usize>())
                .ok_or(Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)?;
    let mut allocations = 1usize;
    let mut work = 0usize;

    for labels in [data.row_names(), data.column_names()] {
        retained = retained
            .checked_add(
                labels
                    .len()
                    .checked_mul(size_of::<String>())
                    .ok_or(Error::InvalidSource)?,
            )
            .ok_or(Error::InvalidSource)?;
        allocations = allocations
            .checked_add(usize::from(!labels.is_empty()))
            .ok_or(Error::InvalidSource)?;
        work = work.checked_add(labels.len()).ok_or(Error::InvalidSource)?;
        for label in labels {
            retained = retained
                .checked_add(label.len())
                .ok_or(Error::InvalidSource)?;
            allocations = allocations
                .checked_add(usize::from(!label.is_empty()))
                .ok_or(Error::InvalidSource)?;
            work = work.checked_add(label.len()).ok_or(Error::InvalidSource)?;
        }
    }

    let rows = data.values();
    retained = retained
        .checked_add(
            rows.len()
                .checked_mul(size_of::<Vec<Option<f64>>>())
                .ok_or(Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)?;
    allocations = allocations
        .checked_add(usize::from(!rows.is_empty()))
        .ok_or(Error::InvalidSource)?;
    work = work.checked_add(rows.len()).ok_or(Error::InvalidSource)?;
    for row in rows {
        retained = retained
            .checked_add(
                row.len()
                    .checked_mul(size_of::<Option<f64>>())
                    .ok_or(Error::InvalidSource)?,
            )
            .ok_or(Error::InvalidSource)?;
        allocations = allocations
            .checked_add(usize::from(!row.is_empty()))
            .ok_or(Error::InvalidSource)?;
        work = work.checked_add(row.len()).ok_or(Error::InvalidSource)?;
    }

    let charge = retained
        .checked_add(allocations)
        .and_then(|value| value.checked_add(work))
        .ok_or(Error::InvalidSource)?;
    preflight_data_budget(budget, LimitKind::Allocations, allocations)?;
    preflight_data_budget(budget, LimitKind::RetainedBytes, retained)?;
    preflight_data_budget(budget, LimitKind::WireWork, charge)?;
    budget.charge_work(charge).map_err(map_axis_support_error)
}

fn data_options(budget: &ChartGraphScanBudget, source: &[u8]) -> Result<DecodeOptions, Error> {
    let (limits, remaining_work) = budget.chart_metadata_residual_limits();
    if remaining_work == 0 {
        return Err(Error::Limit {
            kind: LimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let maximum_depth = u32::try_from(limits.max_nesting()).map_err(|_error| Error::Limit {
        kind: LimitKind::WireNesting,
        observed: usize_to_u64(limits.max_nesting()),
        maximum: u64::from(u32::MAX),
    })?;
    let input = source.len().max(1).min(limits.max_input_bytes());
    let bounded_work = remaining_work.clamp(1, MAX_DATA_CODEC_BYTES);
    let fields = limits.max_fields().min(bounded_work).max(1);
    let work = bounded_work;
    let cells = limits
        .max_fields()
        .min(bounded_work)
        .clamp(1, MAX_DATA_CELL_COUNT);
    let labels = limits
        .max_fields()
        .min(bounded_work)
        .clamp(1, MAX_DATA_LABEL_COUNT);
    let text = limits
        .max_input_bytes()
        .min(bounded_work)
        .clamp(1, MAX_DATA_TEXT_BYTES);
    Ok(
        DecodeOptions::new(input, fields, work, maximum_depth, cells, labels, text)
            .with_max_output_bytes(bounded_work)
            .with_max_allocations(bounded_work)
            .with_max_retained_bytes(bounded_work)
            .with_max_scratch_bytes(bounded_work),
    )
}

fn charge_rewrite_requirements(
    budget: &mut ChartGraphScanBudget,
    requirements: chart_data_codec::RewriteExecutionRequirements,
) -> Result<(), Error> {
    let amount = requirements
        .output_bytes
        .checked_add(requirements.fields)
        .and_then(|value| value.checked_add(requirements.work_bytes))
        .and_then(|value| value.checked_add(requirements.max_depth as usize))
        .and_then(|value| value.checked_add(requirements.allocations))
        .and_then(|value| value.checked_add(requirements.retained_bytes))
        .and_then(|value| value.checked_add(requirements.scratch_bytes))
        .ok_or(Error::InvalidSource)?;
    preflight_data_budget(budget, LimitKind::OutputBytes, amount)?;
    budget.charge_work(amount).map_err(map_axis_support_error)
}

fn retain_payload_pair(
    source: &[u8],
    target: &[u8],
    budget: &mut ChartGraphScanBudget,
) -> Result<(Arc<[u8]>, Arc<[u8]>), Error> {
    let amount = source
        .len()
        .checked_add(target.len())
        .ok_or(Error::InvalidSource)?;
    let allocations = usize::from(!source.is_empty()) + usize::from(!target.is_empty());
    preflight_data_budget(budget, LimitKind::Allocations, allocations)?;
    preflight_data_budget(budget, LimitKind::RetainedBytes, amount)?;
    budget.charge_work(amount).map_err(map_axis_support_error)?;
    Ok((Arc::<[u8]>::from(source), Arc::<[u8]>::from(target)))
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, Error> {
    chart_axis_support::physical_catalog(package).map_err(map_axis_support_error)
}

fn selected_chart_payload<'source>(
    package: &'source Package,
    selection: &ChartSelection,
    budget: &mut ChartGraphScanBudget,
) -> Result<&'source [u8], Error> {
    let (component_name, object) = package
        .object_with_component(selection.chart_identifier)
        .ok_or(Error::InvalidSource)?;
    if component_name != selection.slide_component_name {
        return Err(Error::InvalidSource);
    }
    let (_message_index, message) =
        chart_axis_support::unique_message(object, CHART_MESSAGE_TYPE, budget)
            .map_err(map_axis_support_error)?;
    Ok(message.data.as_slice())
}

fn root_preview_deletions_with_budget(
    catalog: &litchi_iwa_archive::package::Catalog,
    budget: &mut ChartGraphScanBudget,
) -> Result<super::rendering_invalidation::RootPreviewPlan, Error> {
    const ROOT_PREVIEW_NAMES: usize = 3;
    let scan_work = catalog
        .len()
        .checked_mul(ROOT_PREVIEW_NAMES)
        .and_then(|value| value.checked_add(ROOT_PREVIEW_NAMES))
        .ok_or(Error::InvalidSource)?;
    let retained = ROOT_PREVIEW_NAMES
        .checked_mul(size_of::<&str>())
        .ok_or(Error::InvalidSource)?;
    // `root_preview_deletions` builds a bounded boxed slice after scanning the
    // catalog. Charge its complete fixed-name envelope before entering that
    // allocator so an over-budget plan cannot partially materialize.
    preflight_data_budget(budget, LimitKind::Allocations, 1)?;
    preflight_data_budget(budget, LimitKind::RetainedBytes, retained)?;
    preflight_data_budget(budget, LimitKind::WireWork, scan_work)?;
    budget
        .charge_work(scan_work)
        .map_err(map_axis_support_error)?;
    super::rendering_invalidation::root_preview_deletions(catalog).map_err(|error| match error {
        super::rendering_invalidation::RenderingInvalidationError::Allocation { amount } => {
            Error::Allocation { amount }
        },
        super::rendering_invalidation::RenderingInvalidationError::InvalidSource => {
            Error::InvalidSource
        },
    })
}

fn preview_count_removed(
    source: &Package,
    candidate: &Package,
    budget: &mut ChartGraphScanBudget,
) -> Result<usize, Error> {
    let source_names =
        root_preview_deletions_with_budget(physical_catalog(source)?.package(), budget)?;
    let candidate_names =
        root_preview_deletions_with_budget(physical_catalog(candidate)?.package(), budget)?;
    budget
        .charge_work(
            source_names
                .len()
                .checked_add(candidate_names.len())
                .and_then(|value| value.checked_mul(3))
                .ok_or(Error::InvalidSource)?,
        )
        .map_err(map_axis_support_error)?;
    Ok(source_names
        .names()
        .iter()
        .filter(|name| !candidate_names.names().contains(name))
        .count())
}

/// Publish a replacement chart drawable message through the native Keynote
/// component pipeline. The caller owns semantic selection and has already
/// prepared the source-preserving chart-data payload; this helper only
/// changes the selected message, removes stale root previews, and reopens the
/// exact candidate.
fn rewrite_chart_payload(
    source: &Package,
    selection: &ChartSelection,
    rewritten: Vec<u8>,
    budget: &mut ChartGraphScanBudget,
) -> Result<(Package, usize), Error> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.slide_component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let component = source
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selection.slide_component_name)
        .ok_or(Error::InvalidSource)?;
    let (message_index, original) = {
        let object = component
            .archive()
            .objects
            .iter()
            .find(|object| object.archive_info.identifier == Some(selection.chart_identifier))
            .ok_or(Error::InvalidSource)?;
        let mut selected = None;
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ != CHART_MESSAGE_TYPE {
                continue;
            }
            if selected.replace((index, message.data.as_slice())).is_some() {
                return Err(Error::InvalidSource);
            }
        }
        selected.ok_or(Error::InvalidSource)?
    };
    let original_len = original.len();
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    let stream_bound = parsed_archive_stream_bound(component.archive())?;
    let replacement_bound = stream_bound
        .checked_sub(original_len)
        .and_then(|value| value.checked_add(rewritten.len()))
        .and_then(|value| value.checked_add(64))
        .ok_or(Error::InvalidSource)?;
    let compressed_bound = SnappyStream::maximum_compressed_len(replacement_bound, snappy_limits)
        .map_err(map_core_error)?;
    if replacement_bound > usize::try_from(physical_limits.max_entry_bytes()).unwrap_or(usize::MAX)
        || compressed_bound > snappy_limits.max_compressed_stream()
    {
        return Err(Error::Limit {
            kind: LimitKind::EntryBytes,
            observed: usize_to_u64(replacement_bound.max(compressed_bound)),
            maximum: usize_to_u64(
                usize::try_from(physical_limits.max_entry_bytes())
                    .unwrap_or(usize::MAX)
                    .min(snappy_limits.max_compressed_stream()),
            ),
        });
    }
    let package_bound = source
        .source_bytes()
        .len()
        .checked_sub(entry.data().len())
        .and_then(|value| value.checked_add(compressed_bound))
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
    let previews = root_preview_deletions_with_budget(catalog.package(), budget)?;
    let precharged = entry
        .data()
        .len()
        .checked_add(stream_bound)
        .and_then(|value| value.checked_add(replacement_bound))
        .and_then(|value| value.checked_add(compressed_bound))
        .and_then(|value| value.checked_add(package_bound))
        .and_then(|value| value.checked_add(source.source_bytes().len()))
        .and_then(|value| value.checked_add(64 * previews.len()))
        .ok_or(Error::InvalidSource)?;
    preflight_data_budget(budget, LimitKind::ScratchBytes, precharged)?;
    budget
        .charge_work(precharged)
        .map_err(map_axis_support_error)?;

    let stream = SnappyStream::decompress(entry.data(), snappy_limits).map_err(map_core_error)?;
    let mut archive = Archive::parse(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    let object = archive
        .objects
        .iter()
        .find(|object| object.archive_info.identifier == Some(selection.chart_identifier))
        .ok_or(Error::InvalidSource)?;
    if object
        .messages
        .get(message_index)
        .map(|message| message.type_)
        != Some(CHART_MESSAGE_TYPE)
    {
        return Err(Error::InvalidSource);
    }
    archive
        .object_mut(selection.chart_identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_bound = archive
        .encoded_len(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound = SnappyStream::maximum_compressed_len(encoded_bound, snappy_limits)
        .map_err(map_core_error)?;
    if encoded_bound > usize::try_from(physical_limits.max_entry_bytes()).unwrap_or(usize::MAX)
        || compressed_bound > snappy_limits.max_compressed_stream()
    {
        return Err(Error::Limit {
            kind: LimitKind::EntryBytes,
            observed: usize_to_u64(encoded_bound.max(compressed_bound)),
            maximum: usize_to_u64(
                usize::try_from(physical_limits.max_entry_bytes())
                    .unwrap_or(usize::MAX)
                    .min(snappy_limits.max_compressed_stream()),
            ),
        });
    }
    let bytes = archive.to_bytes(archive_limits).map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
    let edit = EntryEdit::new(selection.slide_component_name.as_str(), &compressed);
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(
            std::slice::from_ref(&edit),
            previews.names(),
            physical_limits,
        )
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    let reassembly_work = requirements
        .output_bytes()
        .checked_add(requirements.allocations())
        .and_then(|value| value.checked_add(requirements.retained_bytes()))
        .and_then(|value| value.checked_add(requirements.scratch_bytes()))
        .ok_or(Error::InvalidSource)?;
    preflight_data_budget(budget, LimitKind::OutputBytes, reassembly_work)?;
    budget
        .charge_work(reassembly_work)
        .map_err(map_axis_support_error)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let package = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(|_error| Error::Verification)?;
    Ok((package, previews.len()))
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

fn verify_chart_candidate(
    source: &Package,
    candidate: &Package,
    source_selection: &ChartSelection,
    expected_payload: Option<&[u8]>,
    previews_must_be_absent: bool,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), Error> {
    // Reopening already enforces the physical ZIP/IWA profile. The selected
    // graph readback plus the complete locality proof below provide the
    // focused semantic invariant; a package-wide `validate` here would walk
    // unrelated lazy document state without adding publication evidence.
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
        || candidate_selection.non_style_identifier != source_selection.non_style_identifier
        || candidate_selection.title != source_selection.title
        || candidate_selection.slide_component_name != source_selection.slide_component_name
    {
        return Err(Error::Verification);
    }
    let (_, source_object) = source
        .object_with_component(source_selection.chart_identifier)
        .ok_or(Error::Verification)?;
    let (message_index, _) =
        chart_axis_support::unique_message(source_object, CHART_MESSAGE_TYPE, budget)
            .map_err(map_axis_support_error)?;
    chart_axis_support::verify_package_locality_for_component(
        source,
        candidate,
        &source_selection.slide_component_name,
        source_selection.chart_identifier,
        message_index,
        previews_must_be_absent,
        expected_payload,
        budget,
    )
    .map_err(map_axis_support_error)
}

fn read_selected_data(
    package: &Package,
    selection: ChartSelection,
    budget: &mut ChartGraphScanBudget,
) -> Result<ChartData, Error> {
    let (component_name, chart_object) = package
        .object_with_component(selection.chart_identifier)
        .ok_or(Error::InvalidSource)?;
    if component_name != selection.slide_component_name {
        return Err(Error::InvalidSource);
    }
    let (message_index, message) =
        chart_axis_support::unique_message(chart_object, CHART_MESSAGE_TYPE, budget)
            .map_err(map_axis_support_error)?;
    chart_axis_support::validate_selected_message_metadata(chart_object, message_index)
        .map_err(map_axis_support_error)?;

    let (limits, residual_work) = budget.chart_metadata_residual_limits();
    if residual_work == 0 {
        return Err(Error::Limit {
            kind: LimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let maximum_depth = u32::try_from(limits.max_nesting()).map_err(|_error| Error::Limit {
        kind: LimitKind::WireNesting,
        observed: usize_to_u64(limits.max_nesting()),
        maximum: u64::from(u32::MAX),
    })?;
    let options = DecodeOptions::new(
        message.data.len().max(1).min(limits.max_input_bytes()),
        limits.max_fields().min(residual_work).max(1),
        residual_work,
        maximum_depth,
        limits
            .max_fields()
            .min(residual_work)
            .clamp(1, MAX_DATA_CELL_COUNT),
        limits
            .max_fields()
            .min(residual_work)
            .clamp(1, MAX_DATA_LABEL_COUNT),
        limits
            .max_input_bytes()
            .min(residual_work)
            .clamp(1, MAX_DATA_TEXT_BYTES),
    );
    let (snapshot, report) =
        chart_data_codec::decode_modern_with_report(message.data.as_slice(), &options).map_err(
            |error| {
                let charge = charge_data_report(budget, error.report());
                match charge {
                    Ok(()) => map_chart_data_decode_error(error),
                    Err(budget_error) => budget_error,
                }
            },
        )?;
    charge_data_report(budget, report)?;
    charge_snapshot_projection(snapshot, budget)?;

    let row_names = copy_labels(snapshot.row_labels(), budget)?;
    let column_names = copy_labels(snapshot.column_labels(), budget)?;
    let values = copy_values(snapshot, budget)?;
    ChartData::new(row_names, column_names, values).map_err(map_data_error)
}

fn charge_data_report(
    budget: &mut ChartGraphScanBudget,
    report: chart_data_codec::DecodeReport,
) -> Result<(), Error> {
    let amount = report
        .source_bytes()
        .checked_add(report.fields())
        .and_then(|value| value.checked_add(report.work_bytes()))
        .and_then(|value| value.checked_add(report.max_depth() as usize))
        .and_then(|value| value.checked_add(report.cell_count()))
        .and_then(|value| value.checked_add(report.label_count()))
        .and_then(|value| value.checked_add(report.text_bytes()))
        .and_then(|value| value.checked_add(report.allocations()))
        .and_then(|value| value.checked_add(report.retained_bytes()))
        .ok_or(Error::InvalidSource)?;
    budget.charge_work(amount).map_err(map_axis_support_error)
}

/// Reserve the repeated borrowed-view walks performed while materializing the
/// common model.  The codec report accounts for its own strict preflight and
/// scalar projection; this envelope covers the two label passes, row framing,
/// and value-field walks owned by this package adapter.
fn charge_snapshot_projection(
    snapshot: chart_data_codec::ChartDataSnapshot<'_>,
    budget: &mut ChartGraphScanBudget,
) -> Result<(), Error> {
    let cells = snapshot
        .row_count()
        .checked_mul(snapshot.column_count())
        .ok_or(Error::InvalidSource)?;
    let amount = snapshot
        .grid_source()
        .len()
        .max(1)
        .checked_mul(10)
        .and_then(|value| value.checked_add(cells))
        .ok_or(Error::InvalidSource)?;
    preflight_data_budget(budget, LimitKind::WireWork, amount)?;
    budget.charge_work(amount).map_err(map_axis_support_error)
}

fn copy_labels<'source>(
    labels: LabelList<'source>,
    budget: &mut ChartGraphScanBudget,
) -> Result<Vec<String>, Error> {
    let count = labels.len();
    let mut text_bytes = 0usize;
    let mut string_allocations = 0usize;
    let mut seen = 0usize;
    for label in labels.iter() {
        seen = seen.checked_add(1).ok_or(Error::InvalidSource)?;
        text_bytes = text_bytes
            .checked_add(label.len())
            .ok_or(Error::InvalidSource)?;
        string_allocations = string_allocations
            .checked_add(usize::from(!label.is_empty()))
            .ok_or(Error::InvalidSource)?;
    }
    if seen != count {
        return Err(Error::InvalidSource);
    }
    let allocations = usize::from(count != 0)
        .checked_add(string_allocations)
        .ok_or(Error::InvalidSource)?;
    let retained = count
        .checked_mul(size_of::<String>())
        .and_then(|value| value.checked_add(text_bytes))
        .ok_or(Error::InvalidSource)?;
    let charge = count
        .checked_add(text_bytes)
        .and_then(|value| value.checked_add(allocations))
        .and_then(|value| value.checked_add(retained))
        .ok_or(Error::InvalidSource)?;
    preflight_data_budget(budget, LimitKind::Allocations, allocations)?;
    preflight_data_budget(budget, LimitKind::RetainedBytes, retained)?;
    preflight_data_budget(budget, LimitKind::WireWork, charge)?;

    let mut copied = Vec::new();
    copied
        .try_reserve_exact(count)
        .map_err(|_error| Error::Allocation { amount: count })?;
    for label in labels.iter() {
        let mut owned = String::new();
        owned
            .try_reserve_exact(label.len())
            .map_err(|_error| Error::Allocation {
                amount: label.len(),
            })?;
        owned.push_str(label);
        copied.push(owned);
    }
    if copied.len() != count {
        return Err(Error::InvalidSource);
    }
    budget.charge_work(charge).map_err(map_axis_support_error)?;
    Ok(copied)
}

fn copy_values(
    snapshot: chart_data_codec::ChartDataSnapshot<'_>,
    budget: &mut ChartGraphScanBudget,
) -> Result<Vec<Vec<Option<f64>>>, Error> {
    let rows = snapshot.row_count();
    let columns = snapshot.column_count();
    let cells = rows.checked_mul(columns).ok_or(Error::InvalidSource)?;
    let value_work = rows
        .checked_add(cells)
        .and_then(|value| value.checked_add(rows.checked_mul(size_of::<Vec<Option<f64>>>())?))
        .and_then(|value| value.checked_add(cells.checked_mul(size_of::<Option<f64>>())?))
        .ok_or(Error::InvalidSource)?;
    let allocations = usize::from(rows != 0)
        .checked_add(rows)
        .ok_or(Error::InvalidSource)?;
    let retained = rows
        .checked_mul(size_of::<Vec<Option<f64>>>())
        .and_then(|value| value.checked_add(cells.checked_mul(size_of::<Option<f64>>())?))
        .ok_or(Error::InvalidSource)?;
    let charge = value_work
        .checked_add(allocations)
        .ok_or(Error::InvalidSource)?;
    preflight_data_budget(budget, LimitKind::Allocations, allocations)?;
    preflight_data_budget(budget, LimitKind::RetainedBytes, retained)?;
    preflight_data_budget(budget, LimitKind::WireWork, charge)?;

    let mut copied = Vec::new();
    copied
        .try_reserve_exact(rows)
        .map_err(|_error| Error::Allocation { amount: rows })?;
    for row in snapshot.rows().iter() {
        if row.len() != columns {
            return Err(Error::InvalidSource);
        }
        let mut values = Vec::new();
        values
            .try_reserve_exact(columns)
            .map_err(|_error| Error::Allocation { amount: columns })?;
        for value in row.values() {
            values.push(value);
        }
        if values.len() != columns {
            return Err(Error::InvalidSource);
        }
        copied.push(values);
    }
    if copied.len() != rows {
        return Err(Error::InvalidSource);
    }
    budget.charge_work(charge).map_err(map_axis_support_error)?;
    Ok(copied)
}

fn preflight_data_budget(
    budget: &ChartGraphScanBudget,
    kind: LimitKind,
    amount: usize,
) -> Result<(), Error> {
    // ChartGraphScanBudget deliberately exposes one aggregate operation
    // ledger shared by selector and metadata readers.  Allocation and
    // retained-byte checks use that same finite remainder, but run before
    // try_reserve so a rejected semantic copy cannot partially materialize.
    let (_, remaining) = budget.chart_metadata_residual_limits();
    if amount > remaining {
        return Err(Error::Limit {
            kind,
            observed: usize_to_u64(amount),
            maximum: usize_to_u64(remaining),
        });
    }
    Ok(())
}

fn map_data_error(_error: DataError) -> Error {
    Error::InvalidSource
}

fn map_axis_support_error(error: AxisSupportError) -> Error {
    map_chart_title_error(slide_chart_title::map_axis_support_error(error))
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
        ChartTitleError::InvalidSource
        | ChartTitleError::Verification
        | ChartTitleError::PatchConflict => Error::InvalidSource,
        ChartTitleError::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_title_limit_kind(kind),
            observed,
            maximum,
        },
        ChartTitleError::Allocation { amount } => Error::Allocation { amount },
    }
}

fn map_title_limit_kind(kind: ChartTitleLimitKind) -> LimitKind {
    match kind {
        ChartTitleLimitKind::InputBytes => LimitKind::InputBytes,
        ChartTitleLimitKind::OutputBytes => LimitKind::WireBytes,
        ChartTitleLimitKind::WireBytes => LimitKind::WireBytes,
        ChartTitleLimitKind::Entries => LimitKind::Entries,
        ChartTitleLimitKind::EntryBytes => LimitKind::EntryBytes,
        ChartTitleLimitKind::TotalBytes => LimitKind::TotalBytes,
        ChartTitleLimitKind::Slides => LimitKind::Slides,
        ChartTitleLimitKind::References => LimitKind::References,
        ChartTitleLimitKind::TextStorages => LimitKind::TextStorages,
        ChartTitleLimitKind::TextFragments => LimitKind::TextFragments,
        ChartTitleLimitKind::TextBytes | ChartTitleLimitKind::TitleBytes => LimitKind::TextBytes,
        ChartTitleLimitKind::SelectorBytes => LimitKind::SelectorBytes,
        ChartTitleLimitKind::WireFields => LimitKind::WireFields,
        ChartTitleLimitKind::WireNesting => LimitKind::WireNesting,
        ChartTitleLimitKind::WireWork => LimitKind::WireWork,
    }
}

fn map_chart_data_decode_error(error: chart_data_codec::DecodeError) -> Error {
    if let Some((observed, maximum)) = error.input_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.cell_limit_values() {
        return Error::Limit {
            kind: LimitKind::Cells,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.label_limit_values() {
        return Error::Limit {
            kind: LimitKind::LabelCount,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.text_limit_values() {
        return Error::Limit {
            kind: LimitKind::TextBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.depth_limit_values() {
        return Error::Limit {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        };
    }
    Error::InvalidSource
}

fn map_data_rewrite_error(error: chart_data_codec::RewriteError) -> Error {
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    if error.is_shape_mismatch() {
        return Error::Refused(Refusal::ShapeChanged);
    }
    if error.is_non_finite_numeric() {
        return Error::InvalidSource;
    }
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    match limit {
        chart_data_codec::RewriteLimit::Bytes { observed, maximum } => Error::Limit {
            kind: LimitKind::WireBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        chart_data_codec::RewriteLimit::Fields { observed, maximum } => Error::Limit {
            kind: LimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        chart_data_codec::RewriteLimit::Work { observed, maximum } => Error::Limit {
            kind: LimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        chart_data_codec::RewriteLimit::Output { observed, maximum } => Error::Limit {
            kind: LimitKind::OutputBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        chart_data_codec::RewriteLimit::Nesting { observed, maximum } => Error::Limit {
            kind: LimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        chart_data_codec::RewriteLimit::Allocations { observed, maximum } => Error::Limit {
            kind: LimitKind::WireAllocations,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        chart_data_codec::RewriteLimit::Retained { observed, maximum } => Error::Limit {
            kind: LimitKind::WireRetainedBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        chart_data_codec::RewriteLimit::Scratch { observed, maximum } => Error::Limit {
            kind: LimitKind::ScratchBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
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
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes
                | litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
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
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => LimitKind::EntryBytes,
                litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems
                | litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::Entries,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::HeaderMemoryBytes => LimitKind::RetainedBytes,
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

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn semantic_copy_budget_refuses_retained_bytes_before_allocation() {
        let source = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/keynote/chart-caption-native.key"),
        )
        .expect("native chart fixture");
        let package = Package::from_bytes(&source).expect("native package");
        let mut budget = ChartGraphScanBudget::new(&package).expect("chart budget");
        let (_, remaining) = budget.chart_metadata_residual_limits();
        assert!(remaining > 1);
        budget
            .charge_work(remaining - 1)
            .expect("reserve all but one work unit");

        assert_eq!(
            preflight_data_budget(&budget, LimitKind::RetainedBytes, 2),
            Err(Error::Limit {
                kind: LimitKind::RetainedBytes,
                observed: 2,
                maximum: 1,
            })
        );
    }

    #[test]
    fn semantic_projection_budget_refuses_a_low_work_ceiling() {
        let source = std::fs::read(
            std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../test-data/iwork/keynote/chart-caption-native.key"),
        )
        .expect("native chart fixture");
        let package = Package::from_bytes(&source).expect("native package");
        let grid = one_cell_grid();
        let (snapshot, _report) =
            chart_data_codec::decode_grid_with_report(&grid, &DecodeOptions::for_source(&grid))
                .expect("one-cell chart grid");
        let mut budget = ChartGraphScanBudget::new(&package).expect("chart budget");
        let (_, remaining) = budget.chart_metadata_residual_limits();
        let projection = snapshot
            .grid_source()
            .len()
            .checked_mul(10)
            .and_then(|value| value.checked_add(1))
            .expect("projection work");
        assert!(remaining > projection);
        budget
            .charge_work(remaining - projection + 1)
            .expect("leave less than the projection envelope");

        assert!(matches!(
            charge_snapshot_projection(snapshot, &mut budget),
            Err(Error::Limit {
                kind: LimitKind::WireWork,
                observed,
                maximum,
            }) if observed == projection as u64 && maximum == (projection - 1) as u64
        ));
    }

    fn one_cell_grid() -> Vec<u8> {
        fn varint(mut value: u32, output: &mut Vec<u8>) {
            while value >= 0x80 {
                output.push((value as u8) | 0x80);
                value >>= 7;
            }
            output.push(value as u8);
        }

        fn text_field(number: u32, value: &[u8], output: &mut Vec<u8>) {
            varint((number << 3) | 2, output);
            varint(value.len() as u32, output);
            output.extend_from_slice(value);
        }

        let mut value = Vec::new();
        varint(9, &mut value);
        value.extend_from_slice(&1.0_f64.to_le_bytes());

        let mut row = Vec::new();
        text_field(1, &value, &mut row);

        let mut grid = Vec::new();
        text_field(1, b"Row", &mut grid);
        text_field(2, b"Value", &mut grid);
        text_field(3, &row, &mut grid);
        grid
    }
}
