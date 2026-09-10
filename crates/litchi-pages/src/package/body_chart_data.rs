//! Selector-first Pages body-chart data reads.
//!
//! The Arrange reader proves the rooted body-chart ownership path before this
//! module asks the shared chart-data codec to inspect the selected drawable.
//! Native identifiers and protobuf objects therefore stay private to the
//! package boundary.  The returned value is the common archive-free
//! [`litchi_iwa_common::chart::data::ChartData`] grid.

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{OwnedExactArtifacts, SharedBytes};
use litchi_iwa_common::chart::data::{ChartData, DataError};
use litchi_iwa_protos::chart_data_codec::{self, DecodeLimit, DecodeOptions};
use thiserror::Error;

use super::Package;
use super::body_chart_arrangement::{
    self, ArrangementBudget, BodyChartArrangementError, BodyChartArrangementLimitKind, ChartTarget,
};
use crate::selector::BodyChartSelector;

const MAX_DATA_CELLS: usize = 1_000_000;
const MAX_DATA_LABEL_COUNT: usize = 1_000_000;
const MAX_DATA_TEXT_BYTES: usize = 64 * 1024 * 1024;

/// Finite resources governed by one Pages body-chart data read.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyChartDataLimitKind {
    /// Complete package input bytes inspected by rooted selection.
    InputBytes,
    /// Native payload bytes inspected.
    PayloadBytes,
    /// Native references inspected.
    PayloadReferences,
    /// Strict chart-data wire bytes.
    WireBytes,
    /// Strict chart-data wire fields.
    WireFields,
    /// Strict chart-data wire nesting.
    WireNesting,
    /// Strict chart-data wire work.
    WireWork,
    /// Borrowed decoder allocations.
    WireAllocations,
    /// Borrowed decoder retained bytes.
    WireRetainedBytes,
    /// Numeric cell count.
    CellCount,
    /// Number of row and column labels.
    LabelCount,
    /// UTF-8 label bytes.
    TextBytes,
}

impl fmt::Display for BodyChartDataLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::PayloadBytes => "payload bytes",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::WireAllocations => "wire allocations",
            Self::WireRetainedBytes => "wire retained bytes",
            Self::CellCount => "cell count",
            Self::LabelCount => "label count",
            Self::TextBytes => "text bytes",
        })
    }
}

/// Failure while reading one rooted Pages body chart's semantic data grid.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum BodyChartDataError {
    /// No rooted body chart matched the source-order selector.
    #[error("the Pages body has no chart at position {position:?}")]
    ChartNotFound { position: Position },
    /// The source does not retain a supported exact native artifact.
    #[error("this Pages source does not support body-chart data reads")]
    UnsupportedSource,
    /// The rooted graph or selected chart-data payload was malformed.
    #[error("the selected Pages body-chart data source is invalid")]
    InvalidSource,
    /// The requested edit changes row or column labels, which this first
    /// source-preserving numeric transaction does not own.
    #[error("Pages body-chart data edits must preserve row and column labels")]
    LabelsChanged,
    /// The requested edit changes the grid dimensions, which this first
    /// source-preserving numeric transaction does not own.
    #[error("Pages body-chart data edits must preserve chart dimensions")]
    ShapeChanged,
    /// Reopening the candidate did not reproduce the requested numeric grid.
    #[error("the edited Pages body-chart data failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact or stale
    /// chart source.
    #[error("the Pages body-chart data patch does not match the exact source package")]
    PatchConflict,
    /// A finite read resource ceiling was exceeded.
    #[error("Pages body-chart data {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: BodyChartDataLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded semantic allocation failed.
    #[error("could not allocate {amount} units for Pages body-chart data")]
    Allocation { amount: usize },
}

/// One mutable semantic body-chart data value staged against an immutable
/// Pages package snapshot.
pub struct BodyChartDataEdit<'a> {
    source: &'a Package,
    target: ChartTarget,
    before: Arc<ChartData>,
    after: Option<ChartData>,
}

impl fmt::Debug for BodyChartDataEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyChartDataEdit")
            .field("position", &self.target.position)
            .field("before", self.before.as_ref())
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyChartDataEdit<'_> {
    /// Return the selected body-chart source position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.target.position
    }

    /// Return the chart data read before this edit.
    #[must_use]
    pub fn before(&self) -> &ChartData {
        self.before.as_ref()
    }

    /// Return the chart data currently staged for publication.
    #[must_use]
    pub fn after(&self) -> &ChartData {
        self.after.as_ref().unwrap_or(self.before.as_ref())
    }

    /// Replace the staged chart data.
    #[must_use]
    pub fn set(mut self, data: ChartData) -> Self {
        self.after = Some(data);
        self
    }

    /// Validate and publish the staged exact-source edit.
    pub fn commit(self) -> Result<BodyChartDataCommit, BodyChartDataError> {
        commit_data_edit(self)
    }
}

/// Exact-source checked reversible body-chart data patch.
#[derive(Clone, PartialEq)]
pub struct BodyChartDataPatch {
    artifacts: OwnedExactArtifacts,
    target: ChartTarget,
    before: Arc<ChartData>,
    after: Arc<ChartData>,
    touched_components: usize,
    source_payload: Option<Arc<[u8]>>,
    target_payload: Option<Arc<[u8]>>,
    source_previews: usize,
    target_previews: usize,
}

impl fmt::Debug for BodyChartDataPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyChartDataPatch")
            .field("position", &self.target.position)
            .field("before", self.before.as_ref())
            .field("after", self.after.as_ref())
            .finish_non_exhaustive()
    }
}

impl BodyChartDataPatch {
    /// Return the selected body-chart source position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.target.position
    }

    /// Return the semantic data required before this patch applies.
    #[must_use]
    pub fn before(&self) -> &ChartData {
        self.before.as_ref()
    }

    /// Return the semantic data produced by this patch.
    #[must_use]
    pub fn after(&self) -> &ChartData {
        self.after.as_ref()
    }

    /// Return a diagnostic fingerprint of the exact source artifact.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return a diagnostic fingerprint of the exact target artifact.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return the number of rewritten native components.
    #[must_use]
    pub const fn touched_components(&self) -> usize {
        self.touched_components
    }

    /// Return whether both semantic data and exact bytes are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.bitwise_eq(&self.after) && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            target: self.target.clone(),
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
            touched_components: self.touched_components,
            source_payload: self.target_payload.clone(),
            target_payload: self.source_payload.clone(),
            source_previews: self.target_previews,
            target_previews: self.source_previews,
        }
    }
}

/// Compact evidence describing one committed body-chart data transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyChartDataDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl BodyChartDataDiagnostics {
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

    /// Return whether the committed package differs semantically from source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of canonical root previews deleted in this
    /// publication direction.
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

/// Fully reopened immutable result of one body-chart data transaction.
#[must_use = "a Pages body-chart data commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyChartDataCommit {
    package: Package,
    patch: BodyChartDataPatch,
    diagnostics: BodyChartDataDiagnostics,
}

impl BodyChartDataCommit {
    /// Borrow the fully reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its reopened package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyChartDataPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyChartDataDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted body chart's modern inline numeric data grid.
    ///
    /// Root selection, strict chart-data decoding, and semantic materializing
    /// allocations consume one aggregate bounded budget.  The decoder keeps
    /// native labels and values borrowed until this method returns the
    /// archive-free common model. Date- and duration-only native cells are
    /// represented as missing values; legacy pre-Unity grids are refused.
    pub fn body_chart_data(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<ChartData, BodyChartDataError> {
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = body_chart_arrangement::resolve_target(self, selector.into(), &mut budget)
            .map_err(map_arrangement_error)?;
        read_target_data(self, &target, &mut budget)
    }

    /// Begin a selector-first immutable body-chart data edit.
    pub fn edit_body_chart_data(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<BodyChartDataEdit<'_>, BodyChartDataError> {
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let target = body_chart_arrangement::resolve_target(self, selector.into(), &mut budget)
            .map_err(map_arrangement_error)?;
        let before = read_target_data(self, &target, &mut budget)?;
        budget
            .preflight_allocations(1)
            .map_err(map_arrangement_error)?;
        budget
            .preflight_retained(size_of::<ChartData>())
            .map_err(map_arrangement_error)?;
        let before = Arc::new(before);
        budget.allocations(1).map_err(map_arrangement_error)?;
        budget
            .retained(size_of::<ChartData>())
            .map_err(map_arrangement_error)?;
        Ok(BodyChartDataEdit {
            source: self,
            target,
            before,
            after: None,
        })
    }

    /// Apply an exact-source checked reversible body-chart data patch.
    pub fn apply_body_chart_data(
        &self,
        patch: &BodyChartDataPatch,
    ) -> Result<BodyChartDataCommit, BodyChartDataError> {
        let mut budget = ArrangementBudget::new(self).map_err(map_arrangement_error)?;
        let source = self.state.source.shared_source();
        let source_owner = SharedBytes::from_shared_slice(Arc::clone(&source));
        budget
            .work(
                source
                    .len()
                    .checked_add(patch.artifacts.source_owner().as_ref().len())
                    .ok_or(BodyChartDataError::InvalidSource)?,
            )
            .map_err(map_arrangement_error)?;
        if !patch.artifacts.authorizes_owner(&source_owner) {
            return Err(BodyChartDataError::PatchConflict);
        }
        let current = body_chart_arrangement::resolve_target(
            self,
            BodyChartSelector::position(patch.target.position),
            &mut budget,
        )
        .map_err(map_arrangement_error)?;
        if !current.same_identity(&patch.target) {
            return Err(BodyChartDataError::PatchConflict);
        }
        let current_data = read_target_data(self, &current, &mut budget)?;
        budget
            .work(chart_data_compare_work(
                &current_data,
                patch.before.as_ref(),
            )?)
            .map_err(map_arrangement_error)?;
        if !current_data.bitwise_eq(&patch.before) {
            return Err(BodyChartDataError::PatchConflict);
        }
        if let Some(expected) = patch.source_payload.as_deref() {
            let actual = body_chart_arrangement::selected_chart_payload(self, &patch.target)
                .map_err(|_| BodyChartDataError::PatchConflict)?;
            budget
                .work(
                    actual
                        .len()
                        .checked_add(expected.len())
                        .ok_or(BodyChartDataError::InvalidSource)?,
                )
                .map_err(map_arrangement_error)?;
            if actual != expected {
                return Err(BodyChartDataError::PatchConflict);
            }
        }
        let patch_semantic_work =
            chart_data_compare_work(patch.before.as_ref(), patch.after.as_ref())?;
        budget
            .work(patch_semantic_work)
            .map_err(map_arrangement_error)?;
        let patch_bytes_work = patch
            .artifacts
            .source_owner()
            .as_ref()
            .len()
            .checked_add(patch.artifacts.target_owner().as_ref().len())
            .ok_or(BodyChartDataError::InvalidSource)?;
        budget
            .work(patch_bytes_work)
            .map_err(map_arrangement_error)?;
        if patch.before.bitwise_eq(&patch.after) && patch.artifacts.is_byte_noop() {
            return Ok(BodyChartDataCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyChartDataDiagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact() {
            return Err(BodyChartDataError::PatchConflict);
        }
        budget
            .work(chart_data_axis_validation_work(
                patch.before.as_ref(),
                patch.after.as_ref(),
            )?)
            .map_err(map_arrangement_error)?;
        validate_data_change(&patch.before, &patch.after)?;
        let target_owner = patch.artifacts.target_owner();
        let target_len = target_owner.as_ref().len();
        budget
            .candidate_reopen(target_len)
            .map_err(map_arrangement_error)?;
        let target_bytes = Arc::<[u8]>::from(target_owner.as_ref());
        let catalog = litchi_iwa_archive::SourceCatalog::from_shared_bytes_with_limits(
            target_bytes,
            self.state.source.limits(),
        )
        .map_err(|_| BodyChartDataError::Verification)?;
        let candidate =
            Package::from_source_catalog(catalog).map_err(|_| BodyChartDataError::Verification)?;
        let verified = body_chart_arrangement::resolve_target(
            &candidate,
            BodyChartSelector::position(patch.target.position),
            &mut budget,
        )
        .map_err(map_arrangement_error)?;
        if !verified.same_identity(&patch.target) {
            return Err(BodyChartDataError::Verification);
        }
        let verified_data = read_target_data(&candidate, &verified, &mut budget)?;
        budget
            .work(chart_data_compare_work(
                &verified_data,
                patch.after.as_ref(),
            )?)
            .map_err(map_arrangement_error)?;
        if !verified_data.bitwise_eq(&patch.after) {
            return Err(BodyChartDataError::Verification);
        }
        body_chart_arrangement::verify_locality(
            self,
            &candidate,
            &patch.target,
            patch.target_payload.as_deref(),
            Some((patch.source_previews, patch.target_previews)),
            body_chart_arrangement::ChartPayloadMutation::NumericData,
            &mut budget,
        )
        .map_err(map_arrangement_error)?;
        Ok(BodyChartDataCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyChartDataDiagnostics::published(
                patch.source_previews.saturating_sub(patch.target_previews),
            ),
        })
    }
}

fn commit_data_edit(
    mut edit: BodyChartDataEdit<'_>,
) -> Result<BodyChartDataCommit, BodyChartDataError> {
    let source = edit.source;
    let source_bytes = source.state.source.shared_source();
    let target = edit.target;
    let before = edit.before;
    let mut budget = ArrangementBudget::new(source).map_err(map_arrangement_error)?;

    // The edit owns the immutable `before` grid through an Arc. Its child
    // buffers were charged while the edit was created, but commit starts a
    // fresh aggregate ledger, so admit the complete retained model before
    // any comparison or patch construction below.
    admit_chart_data(&mut budget, before.as_ref(), true)?;

    let Some(after_data) = edit.after.take() else {
        budget
            .work(source_bytes.len())
            .map_err(map_arrangement_error)?;
        let source_owner = SharedBytes::from_shared_slice(Arc::clone(&source_bytes));
        return Ok(BodyChartDataCommit {
            package: source.snapshot(),
            patch: BodyChartDataPatch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                target,
                before: Arc::clone(&before),
                after: before,
                touched_components: 0,
                source_payload: None,
                target_payload: None,
                source_previews: 0,
                target_previews: 0,
            },
            diagnostics: BodyChartDataDiagnostics::unchanged(),
        });
    };

    let after_retained = chart_data_retained_bytes(&after_data)?;
    budget
        .preflight_retained(after_retained)
        .map_err(map_arrangement_error)?;
    budget
        .retained(after_retained)
        .map_err(map_arrangement_error)?;
    let compare_work = chart_data_compare_work(before.as_ref(), &after_data)?;
    budget.work(compare_work).map_err(map_arrangement_error)?;
    if before.bitwise_eq(&after_data) {
        // The caller-owned `after_data` is dropped after this exact no-op
        // check, so it does not need an Arc allocation admission. Its
        // retained children still had to be admitted before the walk.
        budget
            .work(source_bytes.len())
            .map_err(map_arrangement_error)?;
        let source_owner = SharedBytes::from_shared_slice(Arc::clone(&source_bytes));
        return Ok(BodyChartDataCommit {
            package: source.snapshot(),
            patch: BodyChartDataPatch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                target,
                before: Arc::clone(&before),
                after: before,
                touched_components: 0,
                source_payload: None,
                target_payload: None,
                source_previews: 0,
                target_previews: 0,
            },
            diagnostics: BodyChartDataDiagnostics::unchanged(),
        });
    }

    // Validate labels and dimensions with a second budgeted walk; the first
    // comparison above deliberately uses exact float bits to preserve signed
    // zero changes.
    budget
        .work(chart_data_axis_validation_work(
            before.as_ref(),
            &after_data,
        )?)
        .map_err(map_arrangement_error)?;
    validate_data_change(before.as_ref(), &after_data)?;
    if !source.state.source.source_is_exact() {
        return Err(BodyChartDataError::UnsupportedSource);
    }
    // `after_data` was supplied by the caller. Its children are already
    // admitted above; only the Arc control block is allocated by this API.
    budget
        .preflight_allocations(1)
        .map_err(map_arrangement_error)?;
    let after = Arc::new(after_data);
    budget.allocations(1).map_err(map_arrangement_error)?;
    let current = body_chart_arrangement::resolve_target(
        source,
        BodyChartSelector::position(target.position),
        &mut budget,
    )
    .map_err(map_arrangement_error)?;
    if !current.same_identity(&target) {
        return Err(BodyChartDataError::PatchConflict);
    }
    let current_data = read_target_data(source, &current, &mut budget)?;
    budget
        .work(chart_data_compare_work(&current_data, before.as_ref())?)
        .map_err(map_arrangement_error)?;
    if !current_data.bitwise_eq(&before) {
        return Err(BodyChartDataError::PatchConflict);
    }
    let source_payload = body_chart_arrangement::selected_chart_payload(source, &target)
        .map_err(map_arrangement_error)?;
    let options = data_options(&budget, source_payload)?;
    let prepared =
        chart_data_codec::prepare_chart_data_rewrite(source_payload, after.values(), options)
            .map_err(map_data_rewrite_error)?;
    charge_data_report(&mut budget, prepared.prepare_report())?;
    let requirements = prepared.execution_requirements();
    budget
        .charge_rewrite_values(
            requirements.output_bytes,
            requirements.fields,
            requirements.work_bytes,
            requirements.max_depth,
            requirements.allocations,
            requirements.retained_bytes,
            requirements.scratch_bytes,
        )
        .map_err(map_arrangement_error)?;
    let rewritten = prepared
        .execute(requirements.exact())
        .map_err(map_data_rewrite_error)?
        .into_output();
    let publication = body_chart_arrangement::rewrite_chart_payload(
        source,
        &target,
        rewritten,
        body_chart_arrangement::PreviewDeletionMode::DeleteCanonicalRoot,
        &mut budget,
    )
    .map_err(map_arrangement_error)?;
    let source_previews = publication.source_previews;
    let target_previews = publication.target_previews;
    let candidate = publication.package;
    let verified = body_chart_arrangement::resolve_target(
        &candidate,
        BodyChartSelector::position(target.position),
        &mut budget,
    )
    .map_err(map_arrangement_error)?;
    if !verified.same_identity(&target) {
        return Err(BodyChartDataError::Verification);
    }
    let verified_data = read_target_data(&candidate, &verified, &mut budget)?;
    budget
        .work(chart_data_compare_work(&verified_data, after.as_ref())?)
        .map_err(map_arrangement_error)?;
    if !verified_data.bitwise_eq(&after) {
        return Err(BodyChartDataError::Verification);
    }
    let target_payload = body_chart_arrangement::selected_chart_payload(&candidate, &verified)
        .map_err(map_arrangement_error)?;
    body_chart_arrangement::verify_locality(
        source,
        &candidate,
        &target,
        Some(target_payload),
        Some((source_previews, target_previews)),
        body_chart_arrangement::ChartPayloadMutation::NumericData,
        &mut budget,
    )
    .map_err(map_arrangement_error)?;
    let (source_payload, target_payload) =
        body_chart_arrangement::retain_payload_pair(source_payload, target_payload, &mut budget)
            .map_err(map_arrangement_error)?;
    let target_bytes = candidate.state.source.shared_source();
    budget
        .work(
            source_bytes
                .len()
                .checked_add(target_bytes.len())
                .ok_or(BodyChartDataError::InvalidSource)?,
        )
        .map_err(map_arrangement_error)?;
    let source_owner = SharedBytes::from_shared_slice(source_bytes);
    let target_owner = SharedBytes::from_shared_slice(Arc::clone(&target_bytes));
    Ok(BodyChartDataCommit {
        package: candidate,
        patch: BodyChartDataPatch {
            artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            target,
            before,
            after,
            touched_components: 1,
            source_payload: Some(source_payload),
            target_payload: Some(target_payload),
            source_previews,
            target_previews,
        },
        diagnostics: BodyChartDataDiagnostics::published(
            source_previews.saturating_sub(target_previews),
        ),
    })
}

fn validate_data_change(before: &ChartData, after: &ChartData) -> Result<(), BodyChartDataError> {
    if before.row_names().len() != after.row_names().len()
        || before.column_names().len() != after.column_names().len()
        || before.values().len() != after.values().len()
        || before
            .values()
            .iter()
            .zip(after.values())
            .any(|(before, after)| before.len() != after.len())
    {
        return Err(BodyChartDataError::ShapeChanged);
    }
    if before.row_names() != after.row_names() || before.column_names() != after.column_names() {
        return Err(BodyChartDataError::LabelsChanged);
    }
    Ok(())
}

/// Return the logical retained footprint of one archive-free chart grid.
///
/// `ChartData` intentionally exposes borrowed slices rather than allocator
/// capacities. The model therefore charges the bytes represented by its
/// validated lengths: the model header, `String` and row-`Vec` headers, label
/// text, and row-major cells. This is the same shape admitted by the focused
/// reader and is sufficient to reject oversized caller-owned replacements
/// before they enter a patch or an Arc.
fn chart_data_retained_bytes(data: &ChartData) -> Result<usize, BodyChartDataError> {
    let label_count = data
        .row_names()
        .len()
        .checked_add(data.column_names().len())
        .ok_or(BodyChartDataError::InvalidSource)?;
    let label_headers = label_count
        .checked_mul(size_of::<String>())
        .ok_or(BodyChartDataError::InvalidSource)?;
    let label_text = data
        .row_names()
        .iter()
        .chain(data.column_names())
        .try_fold(0usize, |bytes, label| {
            bytes
                .checked_add(label.len())
                .ok_or(BodyChartDataError::InvalidSource)
        })?;
    let row_headers = data
        .values()
        .len()
        .checked_mul(size_of::<Vec<Option<f64>>>())
        .ok_or(BodyChartDataError::InvalidSource)?;
    let cells = data.values().iter().try_fold(0usize, |cells, row| {
        cells
            .checked_add(row.len())
            .ok_or(BodyChartDataError::InvalidSource)
    })?;
    let cell_bytes = cells
        .checked_mul(size_of::<Option<f64>>())
        .ok_or(BodyChartDataError::InvalidSource)?;
    size_of::<ChartData>()
        .checked_add(label_headers)
        .and_then(|bytes| bytes.checked_add(label_text))
        .and_then(|bytes| bytes.checked_add(row_headers))
        .and_then(|bytes| bytes.checked_add(cell_bytes))
        .ok_or(BodyChartDataError::InvalidSource)
}

/// Charge the retained model and, when requested, one owning Arc allocation.
fn admit_chart_data(
    budget: &mut ArrangementBudget,
    data: &ChartData,
    arc_allocation: bool,
) -> Result<usize, BodyChartDataError> {
    let retained = chart_data_retained_bytes(data)?;
    if arc_allocation {
        budget
            .preflight_allocations(1)
            .map_err(map_arrangement_error)?;
    }
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;
    if arc_allocation {
        budget.allocations(1).map_err(map_arrangement_error)?;
    }
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(retained)
}

/// Bound one exact semantic comparison before its labels and cells are read.
fn chart_data_compare_work(
    left: &ChartData,
    right: &ChartData,
) -> Result<usize, BodyChartDataError> {
    chart_data_retained_bytes(left)?
        .checked_add(chart_data_retained_bytes(right)?)
        .ok_or(BodyChartDataError::InvalidSource)
}

/// Bound the labels, row headers, and shape probes used by edit validation.
/// Numeric cells are deliberately excluded because `bitwise_eq` already
/// charged the full value walk immediately before this narrower validation.
fn chart_data_axis_validation_work(
    left: &ChartData,
    right: &ChartData,
) -> Result<usize, BodyChartDataError> {
    fn one(data: &ChartData) -> Result<usize, BodyChartDataError> {
        let label_count = data
            .row_names()
            .len()
            .checked_add(data.column_names().len())
            .ok_or(BodyChartDataError::InvalidSource)?;
        let label_text = data
            .row_names()
            .iter()
            .chain(data.column_names())
            .try_fold(0usize, |bytes, label| {
                bytes
                    .checked_add(label.len())
                    .ok_or(BodyChartDataError::InvalidSource)
            })?;
        let label_headers = label_count
            .checked_mul(size_of::<String>())
            .ok_or(BodyChartDataError::InvalidSource)?;
        let row_headers = data
            .values()
            .len()
            .checked_mul(size_of::<Vec<Option<f64>>>())
            .ok_or(BodyChartDataError::InvalidSource)?;
        size_of::<ChartData>()
            .checked_add(label_headers)
            .and_then(|bytes| bytes.checked_add(label_text))
            .and_then(|bytes| bytes.checked_add(row_headers))
            .ok_or(BodyChartDataError::InvalidSource)
    }

    one(left)?
        .checked_add(one(right)?)
        .ok_or(BodyChartDataError::InvalidSource)
}

fn read_target_data(
    package: &Package,
    target: &ChartTarget,
    budget: &mut ArrangementBudget,
) -> Result<ChartData, BodyChartDataError> {
    let source = body_chart_arrangement::selected_chart_payload(package, target)
        .map_err(map_arrangement_error)?;
    let options = data_options(budget, source)?;
    let (snapshot, report) = match chart_data_codec::decode_modern_with_report(source, &options) {
        Ok(decoded) => decoded,
        Err(error) => {
            let report = error.report();
            charge_data_report(budget, report)?;
            return Err(map_data_codec_error(error));
        },
    };
    charge_data_report(budget, report)?;
    charge_materialization_work(budget, snapshot, report)?;

    let row_names = own_labels(snapshot.row_labels(), budget)?;
    let column_names = own_labels(snapshot.column_labels(), budget)?;
    let values = own_values(snapshot.rows(), snapshot.column_count(), budget)?;

    ChartData::new(row_names, column_names, values).map_err(map_data_error)
}

fn data_options(
    budget: &ArrangementBudget,
    source: &[u8],
) -> Result<DecodeOptions, BodyChartDataError> {
    let limits = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    let source_bytes = source.len().max(1).min(limits.max_input_bytes());
    let fields = limits.max_fields().max(1);
    let work = limits.max_rewrite_work().max(1);
    let max_depth =
        u32::try_from(limits.max_nesting()).map_err(|_| BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireNesting,
            observed: limits.max_nesting() as u64,
            maximum: u32::MAX as u64,
        })?;
    let max_cells = limits.max_fields().clamp(1, MAX_DATA_CELLS);
    let max_labels = limits.max_fields().clamp(1, MAX_DATA_LABEL_COUNT);
    let max_text = limits.max_input_bytes().clamp(1, MAX_DATA_TEXT_BYTES);

    let output = budget
        .residual_output_bytes()
        .map_err(map_arrangement_error)?
        .max(1);
    let allocations = budget
        .residual_allocations()
        .map_err(map_arrangement_error)?
        .max(1);
    let retained = budget
        .residual_retained_bytes()
        .map_err(map_arrangement_error)?
        .max(1);
    Ok(DecodeOptions::new(
        source_bytes,
        fields,
        work,
        max_depth,
        max_cells,
        max_labels,
        max_text,
    )
    .with_max_output_bytes(output)
    .with_max_allocations(allocations)
    .with_max_retained_bytes(retained)
    .with_max_scratch_bytes(retained))
}

fn charge_data_report(
    budget: &mut ArrangementBudget,
    report: chart_data_codec::DecodeReport,
) -> Result<(), BodyChartDataError> {
    budget
        .input(report.source_bytes())
        .map_err(map_arrangement_error)?;
    budget
        .fields(report.fields())
        .map_err(map_arrangement_error)?;
    budget
        .work(report.work_bytes())
        .map_err(map_arrangement_error)?;
    budget
        .allocations(report.allocations())
        .map_err(map_arrangement_error)?;
    budget
        .retained(report.retained_bytes())
        .map_err(map_arrangement_error)?;

    let limits = budget
        .residual_wire_limits()
        .map_err(map_arrangement_error)?;
    if usize::try_from(report.max_depth()).unwrap_or(usize::MAX) > limits.max_nesting() {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireNesting,
            observed: u64::from(report.max_depth()),
            maximum: limits.max_nesting() as u64,
        });
    }
    if report.cell_count() > MAX_DATA_CELLS {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::CellCount,
            observed: report.cell_count() as u64,
            maximum: MAX_DATA_CELLS as u64,
        });
    }
    if report.label_count() > MAX_DATA_LABEL_COUNT {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::LabelCount,
            observed: report.label_count() as u64,
            maximum: MAX_DATA_LABEL_COUNT as u64,
        });
    }
    if report.text_bytes() > MAX_DATA_TEXT_BYTES {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::TextBytes,
            observed: report.text_bytes() as u64,
            maximum: MAX_DATA_TEXT_BYTES as u64,
        });
    }
    Ok(())
}

fn charge_materialization_work(
    budget: &mut ArrangementBudget,
    snapshot: chart_data_codec::ChartDataSnapshot<'_>,
    report: chart_data_codec::DecodeReport,
) -> Result<(), BodyChartDataError> {
    // The borrowed iterators intentionally re-scan source spans after strict
    // decode: labels are walked once to preflight text/allocation shape and
    // once to copy, while rows and values are walked once to materialize the
    // common model. Reserve a source-sized envelope before any such walk so
    // the aggregate budget also bounds this lazy replay work.
    let work = snapshot
        .grid_source()
        .len()
        .checked_mul(10)
        .and_then(|amount| amount.checked_add(report.text_bytes()))
        .and_then(|amount| amount.checked_add(report.cell_count()))
        .ok_or(BodyChartDataError::InvalidSource)?;
    budget.work(work).map_err(map_arrangement_error)
}

fn own_labels(
    labels: chart_data_codec::LabelList<'_>,
    budget: &mut ArrangementBudget,
) -> Result<Vec<String>, BodyChartDataError> {
    let count = labels.len();
    if count > MAX_DATA_LABEL_COUNT {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::LabelCount,
            observed: count as u64,
            maximum: MAX_DATA_LABEL_COUNT as u64,
        });
    }
    let (text_bytes, string_allocations, seen) = labels.iter().try_fold(
        (0usize, 0usize, 0usize),
        |(text_bytes, string_allocations, seen), label| {
            let text_bytes = text_bytes
                .checked_add(label.len())
                .ok_or(BodyChartDataError::InvalidSource)?;
            let string_allocations = string_allocations
                .checked_add(usize::from(!label.is_empty()))
                .ok_or(BodyChartDataError::InvalidSource)?;
            let seen = seen
                .checked_add(1)
                .ok_or(BodyChartDataError::InvalidSource)?;
            Ok((text_bytes, string_allocations, seen))
        },
    )?;
    if seen != count {
        return Err(BodyChartDataError::InvalidSource);
    }
    if text_bytes > MAX_DATA_TEXT_BYTES {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::TextBytes,
            observed: text_bytes as u64,
            maximum: MAX_DATA_TEXT_BYTES as u64,
        });
    }
    let allocations = usize::from(count != 0)
        .checked_add(string_allocations)
        .ok_or(BodyChartDataError::InvalidSource)?;
    let retained = count
        .checked_mul(size_of::<String>())
        .and_then(|amount| amount.checked_add(text_bytes))
        .ok_or(BodyChartDataError::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;

    let mut owned = Vec::new();
    owned
        .try_reserve_exact(count)
        .map_err(|_| BodyChartDataError::Allocation { amount: count })?;
    for label in labels.iter() {
        let mut value = String::new();
        value
            .try_reserve_exact(label.len())
            .map_err(|_| BodyChartDataError::Allocation {
                amount: label.len(),
            })?;
        value.push_str(label);
        owned.push(value);
    }
    if owned.len() != count {
        return Err(BodyChartDataError::InvalidSource);
    }
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(owned)
}

fn own_values(
    rows: chart_data_codec::GridRows<'_>,
    columns: usize,
    budget: &mut ArrangementBudget,
) -> Result<Vec<Vec<Option<f64>>>, BodyChartDataError> {
    let row_count = rows.len();
    let cells = row_count
        .checked_mul(columns)
        .ok_or(BodyChartDataError::InvalidSource)?;
    if cells > MAX_DATA_CELLS {
        return Err(BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::CellCount,
            observed: cells as u64,
            maximum: MAX_DATA_CELLS as u64,
        });
    }
    let allocations = usize::from(row_count != 0)
        .checked_add(row_count)
        .ok_or(BodyChartDataError::InvalidSource)?;
    let retained = row_count
        .checked_mul(size_of::<Vec<Option<f64>>>())
        .and_then(|amount| amount.checked_add(cells.checked_mul(size_of::<Option<f64>>())?))
        .ok_or(BodyChartDataError::InvalidSource)?;
    budget
        .preflight_allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget
        .preflight_retained(retained)
        .map_err(map_arrangement_error)?;

    let mut owned = Vec::new();
    owned
        .try_reserve_exact(row_count)
        .map_err(|_| BodyChartDataError::Allocation { amount: row_count })?;
    for row in rows.iter() {
        if row.len() != columns {
            return Err(BodyChartDataError::InvalidSource);
        }
        let mut values = Vec::new();
        values
            .try_reserve_exact(columns)
            .map_err(|_| BodyChartDataError::Allocation { amount: columns })?;
        for value in row.values() {
            values.push(value);
        }
        if values.len() != columns {
            return Err(BodyChartDataError::InvalidSource);
        }
        owned.push(values);
    }
    if owned.len() != row_count {
        return Err(BodyChartDataError::InvalidSource);
    }
    budget
        .allocations(allocations)
        .map_err(map_arrangement_error)?;
    budget.retained(retained).map_err(map_arrangement_error)?;
    Ok(owned)
}

fn map_arrangement_error(error: BodyChartArrangementError) -> BodyChartDataError {
    match error {
        BodyChartArrangementError::ChartNotFound { position }
        | BodyChartArrangementError::ChartPositionNotFound { position } => {
            BodyChartDataError::ChartNotFound { position }
        },
        BodyChartArrangementError::UnsupportedSource => BodyChartDataError::UnsupportedSource,
        BodyChartArrangementError::InvalidSource => BodyChartDataError::InvalidSource,
        BodyChartArrangementError::Verification => BodyChartDataError::Verification,
        BodyChartArrangementError::PatchConflict => BodyChartDataError::PatchConflict,
        BodyChartArrangementError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyChartDataError::LimitExceeded {
            kind: map_arrangement_limit(kind),
            observed,
            maximum,
        },
        BodyChartArrangementError::Allocation { amount } => {
            BodyChartDataError::Allocation { amount }
        },
    }
}

fn map_arrangement_limit(kind: BodyChartArrangementLimitKind) -> BodyChartDataLimitKind {
    match kind {
        BodyChartArrangementLimitKind::InputBytes
        | BodyChartArrangementLimitKind::PayloadBytes
        | BodyChartArrangementLimitKind::TotalPayloadBytes => BodyChartDataLimitKind::InputBytes,
        BodyChartArrangementLimitKind::PayloadReferences => {
            BodyChartDataLimitKind::PayloadReferences
        },
        BodyChartArrangementLimitKind::WireBytes
        | BodyChartArrangementLimitKind::WireOutputBytes => BodyChartDataLimitKind::WireBytes,
        BodyChartArrangementLimitKind::WireFields
        | BodyChartArrangementLimitKind::PayloadItems
        | BodyChartArrangementLimitKind::PayloadMessages
        | BodyChartArrangementLimitKind::PayloadObjects => BodyChartDataLimitKind::WireFields,
        BodyChartArrangementLimitKind::WireNesting => BodyChartDataLimitKind::WireNesting,
        BodyChartArrangementLimitKind::WireWork => BodyChartDataLimitKind::WireWork,
        BodyChartArrangementLimitKind::WireAllocations => BodyChartDataLimitKind::WireAllocations,
        BodyChartArrangementLimitKind::WireRetainedBytes
        | BodyChartArrangementLimitKind::WireScratchBytes => {
            BodyChartDataLimitKind::WireRetainedBytes
        },
        BodyChartArrangementLimitKind::OutputBytes
        | BodyChartArrangementLimitKind::Entries
        | BodyChartArrangementLimitKind::EntryBytes
        | BodyChartArrangementLimitKind::TotalEntryBytes
        | BodyChartArrangementLimitKind::PackageBytes => BodyChartDataLimitKind::InputBytes,
    }
}

fn map_data_codec_error(error: chart_data_codec::DecodeError) -> BodyChartDataError {
    let Some(limit) = error.resource_limit() else {
        return BodyChartDataError::InvalidSource;
    };
    match limit {
        DecodeLimit::Bytes { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Fields { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Work { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Nesting { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::WireNesting,
            observed: u64::from(observed),
            maximum: u64::from(maximum),
        },
        DecodeLimit::Cells { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::CellCount,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Labels { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::LabelCount,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        DecodeLimit::Text { observed, maximum } => BodyChartDataError::LimitExceeded {
            kind: BodyChartDataLimitKind::TextBytes,
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => BodyChartDataError::InvalidSource,
    }
}

fn map_data_rewrite_error(error: chart_data_codec::RewriteError) -> BodyChartDataError {
    if let Some(amount) = error.allocation_amount() {
        return BodyChartDataError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return BodyChartDataError::InvalidSource;
    };
    match limit {
        chart_data_codec::RewriteLimit::Bytes { observed, maximum } => {
            BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        chart_data_codec::RewriteLimit::Fields { observed, maximum } => {
            BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireFields,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        chart_data_codec::RewriteLimit::Work { observed, maximum } => {
            BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        chart_data_codec::RewriteLimit::Output { observed, maximum } => {
            BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        chart_data_codec::RewriteLimit::Nesting { observed, maximum } => {
            BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        chart_data_codec::RewriteLimit::Allocations { observed, maximum } => {
            BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireAllocations,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        chart_data_codec::RewriteLimit::Retained { observed, maximum }
        | chart_data_codec::RewriteLimit::Scratch { observed, maximum } => {
            BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireRetainedBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        _ => BodyChartDataError::InvalidSource,
    }
}

fn map_data_error(_error: DataError) -> BodyChartDataError {
    BodyChartDataError::InvalidSource
}

#[cfg(test)]
mod tests {
    use super::*;

    const NATIVE: &[u8] =
        include_bytes!("../../../../test-data/iwork/pages/chart-data-native.pages");
    const EDITED_NATIVE: &[u8] =
        include_bytes!("../../../../test-data/iwork/pages/chart-data-edited-native.pages");

    #[test]
    fn selected_payload_cell_limit_is_reported_before_materializing_values() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let mut budget = ArrangementBudget::new(&package).expect("arrangement budget");
        let target = body_chart_arrangement::resolve_target(
            &package,
            BodyChartSelector::index(0),
            &mut budget,
        )
        .expect("native chart target");
        let source = body_chart_arrangement::selected_chart_payload(&package, &target)
            .expect("selected chart payload");
        let options = data_options(&budget, source)
            .expect("data options")
            .with_max_cells(1);
        let error = chart_data_codec::decode_modern_with_report(source, &options)
            .expect_err("the two-by-four grid exceeds one cell");
        assert!(matches!(
            error.resource_limit(),
            Some(DecodeLimit::Cells {
                observed: 2,
                maximum: 1
            })
        ));
    }

    #[test]
    fn postdecode_replay_work_is_refused_by_the_aggregate_budget() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let mut selection_budget = ArrangementBudget::new(&package).expect("arrangement budget");
        let target = body_chart_arrangement::resolve_target(
            &package,
            BodyChartSelector::index(0),
            &mut selection_budget,
        )
        .expect("native chart target");
        let source = body_chart_arrangement::selected_chart_payload(&package, &target)
            .expect("selected chart payload");
        let options = data_options(&selection_budget, source).expect("data options");
        let (snapshot, report) =
            chart_data_codec::decode_modern_with_report(source, &options).expect("chart data");

        let mut exhausted = ArrangementBudget::new(&package).expect("fresh arrangement budget");
        let BodyChartArrangementError::LimitExceeded { maximum, .. } =
            exhausted.work(usize::MAX).expect_err("finite work ceiling")
        else {
            panic!("expected a work limit");
        };
        exhausted
            .work(usize::try_from(maximum).expect("addressable work limit"))
            .expect("consume the work budget");
        let result = charge_materialization_work(&mut exhausted, snapshot, report);
        assert!(matches!(
            result,
            Err(BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireWork,
                ..
            })
        ));
    }

    #[test]
    fn value_copy_refuses_exhausted_retention_budget() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let mut selection_budget = ArrangementBudget::new(&package).expect("arrangement budget");
        let target = body_chart_arrangement::resolve_target(
            &package,
            BodyChartSelector::index(0),
            &mut selection_budget,
        )
        .expect("native chart target");
        let source = body_chart_arrangement::selected_chart_payload(&package, &target)
            .expect("selected chart payload");
        let options = data_options(&selection_budget, source).expect("data options");
        let snapshot = chart_data_codec::decode_modern(source, &options).expect("chart data");

        let mut budget = ArrangementBudget::new(&package).expect("fresh arrangement budget");
        budget
            .retained(
                usize::try_from(package.state.source.limits().max_input_bytes())
                    .expect("addressable retention limit"),
            )
            .expect("consume the retention budget");
        let result = own_values(snapshot.rows(), snapshot.column_count(), &mut budget);
        assert!(matches!(
            result,
            Err(BodyChartDataError::LimitExceeded {
                kind: BodyChartDataLimitKind::WireRetainedBytes,
                ..
            })
        ));
    }

    #[test]
    fn native_resaved_numeric_fixture_preserves_grid_axes_and_unedited_cells() {
        let source = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let edited = Package::from_bytes(EDITED_NATIVE).expect("native resaved Pages fixture");
        assert_eq!(edited.state.source.shared_source().as_ref(), EDITED_NATIVE);
        let before = source
            .body_chart_data(BodyChartSelector::index(0))
            .expect("native source chart data");
        let after = edited
            .body_chart_data(BodyChartSelector::index(0))
            .expect("native resaved chart data");
        assert_eq!(after.row_names(), before.row_names());
        assert_eq!(after.column_names(), before.column_names());
        assert_eq!(after.values()[0][0], Some(27.5));
        assert_eq!(after.values()[0][1], Some(12.75));
        assert_eq!(after.values()[0][2..], before.values()[0][2..]);
        assert_eq!(after.values()[1], before.values()[1]);
    }

    #[test]
    fn same_shape_numeric_edit_reopens_and_inverse_restores_source() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let source_bytes = package.state.source.shared_source();
        let before = package
            .body_chart_data(BodyChartSelector::index(0))
            .expect("native chart data");
        let replacement = ChartData::new(
            before.row_names().to_vec(),
            before.column_names().to_vec(),
            vec![
                vec![Some(27.5), Some(12.75), Some(53.0), Some(96.0)],
                vec![Some(55.0), Some(43.0), Some(70.0), Some(58.0)],
            ],
        )
        .expect("same-shape finite replacement");
        let commit = package
            .edit_body_chart_data(BodyChartSelector::index(0))
            .expect("body chart edit")
            .set(replacement.clone())
            .commit()
            .expect("body chart data commit");
        assert_eq!(commit.diagnostics().touched_components(), 1);
        assert_eq!(commit.diagnostics().deleted_previews(), 3);
        assert!(commit.diagnostics().full_reparse_performed());
        for name in ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"] {
            assert!(
                commit
                    .package()
                    .state
                    .source
                    .package()
                    .iter()
                    .all(|entry| entry.name() != name)
            );
        }
        assert!(
            commit
                .package()
                .body_chart_data(BodyChartSelector::index(0))
                .expect("reopened chart data")
                .bitwise_eq(&replacement)
        );
        let reopened = commit
            .package()
            .body_chart_data(BodyChartSelector::index(0))
            .expect("golden chart data readback");
        assert_eq!(
            reopened.values()[0],
            [Some(27.5), Some(12.75), Some(53.0), Some(96.0)]
        );
        assert_eq!(reopened.values()[1], before.values()[1]);

        let restored = commit
            .package()
            .apply_body_chart_data(&commit.patch().inverse())
            .expect("inverse chart data patch");
        assert_eq!(restored.diagnostics().deleted_previews(), 0);
        assert!(
            restored
                .package()
                .body_chart_data(BodyChartSelector::index(0))
                .expect("restored chart data")
                .bitwise_eq(&before)
        );
        assert_eq!(
            restored.package().state.source.shared_source().as_ref(),
            source_bytes.as_ref()
        );
    }

    #[test]
    fn exact_numeric_noop_preserves_previews_and_source_bytes() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let source_bytes = package.state.source.shared_source();
        let before = package
            .body_chart_data(BodyChartSelector::index(0))
            .expect("native chart data");
        let commit = package
            .edit_body_chart_data(BodyChartSelector::index(0))
            .expect("body chart edit")
            .set(before.clone())
            .commit()
            .expect("body chart no-op");
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().deleted_previews(), 0);
        assert!(commit.patch().is_noop());
        assert_eq!(
            commit.package().state.source.shared_source().as_ref(),
            source_bytes.as_ref()
        );
        for name in ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"] {
            assert!(
                commit
                    .package()
                    .state
                    .source
                    .package()
                    .iter()
                    .any(|entry| entry.name() == name)
            );
        }
    }

    #[test]
    fn data_edit_refuses_label_and_dimension_changes_before_publication() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let before = package
            .body_chart_data(BodyChartSelector::index(0))
            .expect("native chart data");
        let labels_changed = ChartData::new(
            vec![String::from("Changed"), String::from("Region 2")],
            before.column_names().to_vec(),
            before.values().to_vec(),
        )
        .expect("same-shape label replacement");
        assert!(matches!(
            package
                .edit_body_chart_data(BodyChartSelector::index(0))
                .expect("body chart edit")
                .set(labels_changed)
                .commit(),
            Err(BodyChartDataError::LabelsChanged)
        ));

        let dimensions_changed = ChartData::new(
            before.row_names().to_vec(),
            before.column_names()[..3].to_vec(),
            before
                .values()
                .iter()
                .map(|row| row[..3].to_vec())
                .collect(),
        )
        .expect("valid but reshaped replacement");
        assert!(matches!(
            package
                .edit_body_chart_data(BodyChartSelector::index(0))
                .expect("body chart edit")
                .set(dimensions_changed)
                .commit(),
            Err(BodyChartDataError::ShapeChanged)
        ));
    }

    #[test]
    fn applying_published_patch_to_its_target_is_a_conflict() {
        let package = Package::from_bytes(NATIVE).expect("native Pages fixture");
        let before = package
            .body_chart_data(BodyChartSelector::index(0))
            .expect("native chart data");
        let replacement = ChartData::new(
            before.row_names().to_vec(),
            before.column_names().to_vec(),
            before
                .values()
                .iter()
                .enumerate()
                .map(|(row, values)| {
                    values
                        .iter()
                        .enumerate()
                        .map(|(column, value)| {
                            if row == 0 && column == 0 {
                                Some(27.5)
                            } else {
                                *value
                            }
                        })
                        .collect()
                })
                .collect(),
        )
        .expect("same-shape finite replacement");
        let commit = package
            .edit_body_chart_data(BodyChartSelector::index(0))
            .expect("body chart edit")
            .set(replacement)
            .commit()
            .expect("body chart data commit");
        assert!(matches!(
            commit.package().apply_body_chart_data(commit.patch()),
            Err(BodyChartDataError::PatchConflict)
        ));
    }
}
