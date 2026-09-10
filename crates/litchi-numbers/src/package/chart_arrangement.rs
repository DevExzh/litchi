//! Exact-source Numbers chart Arrange-panel transactions.
//!
//! The focused owner admits only the two interaction flags stored by an
//! existing chart drawable: `locked` and `aspect_ratio_locked`.  Semantic
//! sheet/chart positions are resolved before the native graph is entered;
//! native identifiers, archive names, generated values, and payload bytes
//! remain private to this module.  The retained package source is the
//! preservation authority for no-ops, unknown fields, and inverse patches.

#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "The package boundary converts bounded native arithmetic and redacts graph failures."
)]

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::{Entry, EntryEdit, OwnedExactArtifacts};
use litchi_iwa_common::chart::arrangement::ChartArrangement;
use litchi_iwa_common::wire::{
    WireDescent, WirePreflight, WireView, preflight_wire_tree_with_limits,
};
use litchi_iwa_common::{
    Error as CommonError, LimitKind as CommonLimitKind, WireLimits, decode_varint_from_bytes,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::chart_arrangement_codec::{
    self as codec, ChartArrangementWrite, DecodeLimit, DecodeOptions, DecodeReport,
    RewriteExecutionRequirements,
};
use litchi_iwa_protos::numbers_sheet_order_codec;
use litchi_iwa_protos::{chart_data_codec, chart_metadata_codec};
use thiserror::Error;

use super::{Error as PackageError, Package};
use crate::{ChartSelector, SheetSelector};

const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;
const CHART_MESSAGE_TYPE: u32 = 5_021;
const FORM_SHEET_SUPER_FIELD: u32 = 1;
const SHEET_DRAWABLE_FIELD: u32 = 2;
const CHART_DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const MAX_TRANSACTION_ALLOCATIONS: usize = 64;
// The neutral chart codec rejects explicit finite limits above these values.
// Keep the focused owner's residual package budgets below that ceiling before
// constructing codec options, so a large package limit cannot become a codec
// validation error after admission has already succeeded.
const MAX_CODEC_LIMIT: usize = 64 * 1024 * 1024;
const MAX_CODEC_NESTING: u32 = 64;
const MAX_METADATA_LABEL_COUNT: usize = 1_000_000;
// snap 1.1.2 grows Encoder::big to 16,384 u16 entries for a nontrivial
// block. Keep the full 32 KiB workspace in the focused ledger even when the
// input happens to stay below that implementation's private threshold.
const SNAPPY_ENCODER_WORKSPACE_BYTES: usize = 32 * 1024;
// `WireView` keeps a compact span for every parsed field. The span type is
// private to litchi-iwa-common, so reserve a conservative bound per parsed
// field before entering the parser. A large opaque length-delimited value does
// not create one span per payload byte.
const WIRE_VIEW_SPAN_BYTES_PER_FIELD: usize = size_of::<[usize; 8]>();

/// Resource category reported by a focused Numbers chart-arrangement
/// operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartArrangementLimitKind {
    /// Source or candidate wire bytes.
    WireBytes,
    /// Parsed wire fields.
    WireFields,
    /// Nested wire depth.
    WireNesting,
    /// Aggregate graph or codec work.
    WireWork,
    /// Candidate output bytes.
    OutputBytes,
    /// Logical temporary allocations.
    Allocations,
    /// Bytes retained across one transaction.
    RetainedBytes,
    /// Candidate scratch bytes.
    ScratchBytes,
    /// Native graph references inspected by the selector.
    References,
}

impl fmt::Display for ChartArrangementLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::OutputBytes => "output bytes",
            Self::Allocations => "allocations",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::References => "references",
        })
    }
}

/// Failure raised by a selector-first Numbers chart-arrangement operation.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartArrangementError {
    /// The package does not retain an exact physical source suitable for an
    /// edit.
    #[error("this Numbers source does not support physical chart-arrangement edits")]
    UnsupportedSource,
    /// The selected graph crosses an unsupported native dependency boundary.
    #[error("the requested Numbers chart-arrangement graph is unsupported")]
    UnsupportedDependency,
    /// A semantic selector was ambiguous.
    #[error("the Numbers chart-arrangement selector is ambiguous")]
    AmbiguousSelector,
    /// An empty sheet name was supplied.
    #[error("the Numbers sheet selector name cannot be empty")]
    EmptySheetName,
    /// No sheet matched the requested name.
    #[error("the Numbers workbook has no sheet matching the requested name")]
    SheetNameNotFound,
    /// No sheet matched the requested position.
    #[error("the Numbers workbook has no sheet at position {position:?}")]
    SheetPositionNotFound { position: Position },
    /// No chart matched the requested position.
    #[error("the selected Numbers sheet has no chart at position {position:?}")]
    ChartPositionNotFound { position: Position },
    /// The source graph or selected payload is malformed.
    #[error("the Numbers chart-arrangement source is invalid")]
    InvalidSource,
    /// A finite operation budget was exceeded.
    #[error(
        "Numbers chart arrangement {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource that exceeded its ceiling.
        kind: ChartArrangementLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A required temporary allocation could not be reserved.
    #[error("could not allocate {amount} units for Numbers chart arrangement")]
    Allocation { amount: usize },
    /// Candidate reopening or semantic readback did not reproduce the target.
    #[error("the edited Numbers chart arrangement failed semantic verification")]
    Verification,
    /// The patch was created from another exact package artifact.
    #[error("the Numbers chart-arrangement patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable chart Arrange-panel value staged against an immutable Numbers
/// package snapshot.
pub struct ChartArrangementEdit<'a> {
    source: &'a Package,
    selection: ChartSelection,
    after: ChartArrangement,
}

impl fmt::Debug for ChartArrangementEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartArrangementEdit")
            .field("sheet_position", &self.selection.sheet_position)
            .field("chart_position", &self.selection.chart_position)
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl<'a> ChartArrangementEdit<'a> {
    fn new<'sheet>(
        source: &'a Package,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        chart_selector: impl Into<ChartSelector>,
    ) -> Result<Self, ChartArrangementError> {
        let mut budget = ChartBudget::for_package(source)?;
        let selection = select_chart_with_budget(
            source,
            sheet_selector.into(),
            chart_selector.into(),
            false,
            &mut budget,
        )?;
        Ok(Self {
            source,
            after: selection.before,
            selection,
        })
    }

    /// Return the selected sheet source position.
    #[must_use]
    pub const fn sheet_position(&self) -> Position {
        self.selection.sheet_position
    }

    /// Return the selected chart source position within the sheet.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.selection.chart_position
    }

    /// Return the arrangement observed when this edit began.
    #[must_use]
    pub const fn before(&self) -> ChartArrangement {
        self.selection.before
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
    #[must_use]
    pub const fn set_locked(mut self, locked: bool) -> Self {
        self.after = self.after.with_locked(locked);
        self
    }

    /// Stage the requested aspect-ratio constraint while retaining the other
    /// flag.
    #[must_use]
    pub const fn set_constrain_proportions(mut self, value: bool) -> Self {
        self.after = self.after.with_constrain_proportions(value);
        self
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<ChartArrangementCommit, ChartArrangementError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible semantic chart-arrangement patch.
#[derive(Clone, PartialEq, Eq)]
pub struct ChartArrangementPatch {
    artifacts: OwnedExactArtifacts,
    selection: ChartSelection,
    before: ChartArrangement,
    after: ChartArrangement,
    source_payload: Option<Arc<[u8]>>,
    target_payload: Option<Arc<[u8]>>,
}

impl fmt::Debug for ChartArrangementPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartArrangementPatch")
            .field("sheet_position", &self.selection.sheet_position)
            .field("chart_position", &self.selection.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl ChartArrangementPatch {
    /// Return the selected sheet source position.
    #[must_use]
    pub const fn sheet_position(&self) -> Position {
        self.selection.sheet_position
    }

    /// Return the selected chart source position.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.selection.chart_position
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

    /// Return the source artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target artifact's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether the patch preserves the exact source artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact reversible patch from target back to source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            before: self.after,
            after: self.before,
            source_payload: self.target_payload.clone(),
            target_payload: self.source_payload.clone(),
        }
    }
}

/// Compact evidence describing one chart-arrangement publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartArrangementDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl ChartArrangementDiagnostics {
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

    /// Return whether the package bytes and semantic state changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one immutable chart-arrangement transaction.
#[must_use = "a Numbers chart-arrangement commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct ChartArrangementCommit {
    package: Package,
    patch: ChartArrangementPatch,
    diagnostics: ChartArrangementDiagnostics,
}

impl ChartArrangementCommit {
    /// Borrow the fully reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this publication and return its package snapshot.
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

/// Compatibility aliases for callers that use the format-qualified names.
pub type SheetChartArrangementEdit<'a> = ChartArrangementEdit<'a>;
/// Compatibility alias for the format-qualified patch name.
pub type SheetChartArrangementPatch = ChartArrangementPatch;
/// Compatibility alias for the format-qualified commit name.
pub type SheetChartArrangementCommit = ChartArrangementCommit;
/// Compatibility alias for the format-qualified diagnostics name.
pub type SheetChartArrangementDiagnostics = ChartArrangementDiagnostics;
/// Compatibility alias for the format-qualified error name.
pub type SheetChartArrangementError = ChartArrangementError;
/// Compatibility alias for the format-qualified limit vocabulary.
pub type SheetChartArrangementLimitKind = ChartArrangementLimitKind;

#[derive(Clone, PartialEq, Eq)]
pub(super) struct ChartSelection {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
    sheet_position: Position,
    chart_position: Position,
    sheet_identifier: u64,
    chart_identifier: u64,
    component_name: Arc<str>,
    before: ChartArrangement,
}

impl fmt::Debug for ChartSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartSelection")
            .field("sheet_position", &self.sheet_position)
            .field("chart_position", &self.chart_position)
            .finish_non_exhaustive()
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct ChartBudget {
    max_input: usize,
    max_output: usize,
    max_fields: usize,
    max_work: usize,
    max_nesting: usize,
    max_references: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    nesting: usize,
    references: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
}

impl ChartBudget {
    pub(super) fn for_package(package: &Package) -> Result<Self, ChartArrangementError> {
        let archive = package.state.options.archive();
        let core = archive
            .effective_archive_limits()
            .map_err(|_| ChartArrangementError::InvalidSource)?;
        let source = core
            .max_message_bytes()
            .min(core.max_archive_bytes())
            .min(archive.max_iwa_stream_bytes())
            .clamp(1, WireLimits::MAX_INPUT_BYTES);
        let limits = WireLimits::default()
            .with_input_bytes(source)
            .and_then(|limits| {
                limits.with_fields(source.saturating_mul(4).clamp(1, WireLimits::MAX_FIELDS))
            })
            .and_then(|limits| {
                limits.with_output_bytes(
                    source
                        .saturating_add(64)
                        .clamp(1, WireLimits::MAX_OUTPUT_BYTES),
                )
            })
            .and_then(|limits| {
                limits.with_rewrite_work(
                    source
                        .saturating_mul(8)
                        .clamp(1, WireLimits::MAX_REWRITE_WORK),
                )
            })
            .map_err(|_| ChartArrangementError::InvalidSource)?;
        let semantic = package.state.options.semantic();
        let max_allocations = semantic
            .max_objects()
            .checked_add(semantic.max_references())
            .and_then(|amount| amount.checked_add(MAX_TRANSACTION_ALLOCATIONS))
            .ok_or(ChartArrangementError::InvalidSource)?;
        let aggregate = source.saturating_mul(8).max(source);
        Ok(Self {
            max_input: aggregate,
            max_output: aggregate,
            max_fields: limits.max_fields(),
            max_work: limits.max_rewrite_work(),
            max_nesting: limits.max_nesting(),
            max_references: semantic.max_references(),
            max_allocations,
            max_retained: aggregate,
            max_scratch: aggregate,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: ChartArrangementLimitKind,
    ) -> Result<(), ChartArrangementError> {
        let observed = current
            .checked_add(amount)
            .ok_or(ChartArrangementError::InvalidSource)?;
        if observed > maximum {
            return Err(ChartArrangementError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn preflight(
        current: usize,
        amount: usize,
        maximum: usize,
        kind: ChartArrangementLimitKind,
    ) -> Result<(), ChartArrangementError> {
        let observed = current
            .checked_add(amount)
            .ok_or(ChartArrangementError::InvalidSource)?;
        if observed > maximum {
            return Err(ChartArrangementError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        Ok(())
    }

    fn source(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            ChartArrangementLimitKind::WireBytes,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            ChartArrangementLimitKind::OutputBytes,
        )
    }

    fn fields(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::add(
            &mut self.fields,
            amount,
            self.max_fields,
            ChartArrangementLimitKind::WireFields,
        )
    }

    fn work(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            ChartArrangementLimitKind::WireWork,
        )
    }

    fn references(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            ChartArrangementLimitKind::References,
        )
    }

    pub(super) fn allocations(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            ChartArrangementLimitKind::Allocations,
        )
    }

    pub(super) fn retained(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            ChartArrangementLimitKind::RetainedBytes,
        )
    }

    fn scratch(&mut self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            ChartArrangementLimitKind::ScratchBytes,
        )
    }

    fn scan(
        &mut self,
        payload: &[u8],
        fields: usize,
        depth: usize,
    ) -> Result<(), ChartArrangementError> {
        self.source(payload.len())?;
        self.fields(fields)?;
        self.work(payload.len())?;
        if depth > self.max_nesting {
            return Err(ChartArrangementError::LimitExceeded {
                kind: ChartArrangementLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
        Ok(())
    }

    fn preflight_scan(&self, report: WirePreflight) -> Result<(), ChartArrangementError> {
        Self::preflight(
            self.input,
            report.scanned_bytes(),
            self.max_input,
            ChartArrangementLimitKind::WireBytes,
        )?;
        Self::preflight(
            self.fields,
            report.fields(),
            self.max_fields,
            ChartArrangementLimitKind::WireFields,
        )?;
        Self::preflight(
            self.work,
            report.scanned_bytes(),
            self.max_work,
            ChartArrangementLimitKind::WireWork,
        )?;
        if report.max_depth() > self.max_nesting {
            return Err(ChartArrangementError::LimitExceeded {
                kind: ChartArrangementLimitKind::WireNesting,
                observed: report.max_depth() as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn charge_preflight_scan(
        &mut self,
        report: WirePreflight,
    ) -> Result<(), ChartArrangementError> {
        self.source(report.scanned_bytes())?;
        self.fields(report.fields())?;
        self.work(report.scanned_bytes())?;
        self.nesting = self.nesting.max(report.max_depth());
        Ok(())
    }

    pub(super) fn preflight_allocations(&self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::preflight(
            self.allocations,
            amount,
            self.max_allocations,
            ChartArrangementLimitKind::Allocations,
        )
    }

    pub(super) fn preflight_retained(&self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::preflight(
            self.retained,
            amount,
            self.max_retained,
            ChartArrangementLimitKind::RetainedBytes,
        )
    }

    fn preflight_scratch(&self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::preflight(
            self.scratch,
            amount,
            self.max_scratch,
            ChartArrangementLimitKind::ScratchBytes,
        )
    }

    fn preflight_output(&self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::preflight(
            self.output,
            amount,
            self.max_output,
            ChartArrangementLimitKind::OutputBytes,
        )
    }

    fn preflight_work(&self, amount: usize) -> Result<(), ChartArrangementError> {
        Self::preflight(
            self.work,
            amount,
            self.max_work,
            ChartArrangementLimitKind::WireWork,
        )
    }

    pub(super) fn residual_wire_limits(&self) -> Result<WireLimits, ChartArrangementError> {
        let input = self
            .max_input
            .checked_sub(self.input)
            .ok_or(ChartArrangementError::InvalidSource)?
            .clamp(1, WireLimits::MAX_INPUT_BYTES);
        let fields = self
            .max_fields
            .checked_sub(self.fields)
            .ok_or(ChartArrangementError::InvalidSource)?
            .clamp(1, WireLimits::MAX_FIELDS);
        let work = self
            .max_work
            .checked_sub(self.work)
            .ok_or(ChartArrangementError::InvalidSource)?
            .clamp(1, WireLimits::MAX_REWRITE_WORK);
        let nesting = self
            .max_nesting
            .checked_sub(self.nesting)
            .unwrap_or(1)
            .clamp(1, WireLimits::MAX_NESTING);
        WireLimits::default()
            .with_input_bytes(input)
            .and_then(|limits| limits.with_fields(fields))
            .and_then(|limits| limits.with_rewrite_work(work))
            .and_then(|limits| limits.with_nesting(nesting))
            .map_err(|_| ChartArrangementError::InvalidSource)
    }

    fn codec_options(&self, source: &[u8]) -> Result<DecodeOptions, ChartArrangementError> {
        let limits = self.residual_wire_limits()?;
        let depth = u32::try_from(limits.max_nesting().min(MAX_CODEC_NESTING as usize))
            .unwrap_or(MAX_CODEC_NESTING);
        let max_input = limits
            .max_input_bytes()
            .min(source.len().max(1))
            .min(MAX_CODEC_LIMIT);
        let max_fields = limits.max_fields().clamp(1, MAX_CODEC_LIMIT);
        let max_work = limits.max_rewrite_work().clamp(1, MAX_CODEC_LIMIT);
        let max_output = limits.max_output_bytes().clamp(1, MAX_CODEC_LIMIT);
        let max_allocations = self
            .max_allocations
            .saturating_sub(self.allocations)
            .clamp(1, MAX_CODEC_LIMIT);
        let max_retained = self
            .max_retained
            .saturating_sub(self.retained)
            .clamp(1, MAX_CODEC_LIMIT);
        let max_scratch = self
            .max_scratch
            .saturating_sub(self.scratch)
            .clamp(1, MAX_CODEC_LIMIT);
        Ok(DecodeOptions::new(max_input, max_fields, max_work, depth)
            .with_max_output_bytes(max_output)
            .with_max_allocations(max_allocations)
            .with_max_retained_bytes(max_retained)
            .with_max_scratch_bytes(max_scratch))
    }

    /// Build the residual limits for the borrowed chart-metadata projection.
    ///
    /// Chart metadata is a read-only sibling of the Arrange projection, but
    /// both operations must consume the same selector traversal ledger.  Keep
    /// this conversion here so a future metadata caller cannot accidentally
    /// reset the package-wide wire ceilings when it enters its codec.
    pub(super) fn metadata_codec_options(
        &self,
    ) -> Result<chart_metadata_codec::DecodeOptions, ChartArrangementError> {
        let limits = self.residual_wire_limits()?;
        let depth = u32::try_from(limits.max_nesting().min(MAX_CODEC_NESTING as usize))
            .unwrap_or(MAX_CODEC_NESTING);
        let max_input = limits.max_input_bytes().clamp(1, MAX_CODEC_LIMIT);
        let max_fields = limits.max_fields().clamp(1, MAX_CODEC_LIMIT);
        let max_work = limits.max_rewrite_work().clamp(1, MAX_CODEC_LIMIT);
        // Labels are projection output, not native graph edges. Keep their
        // codec ceiling independent from the selector's reference ledger;
        // the decoded strings and vector slots are charged separately by the
        // retained/allocation counters in the metadata adapter.
        let max_labels = limits
            .max_fields()
            .min(MAX_METADATA_LABEL_COUNT)
            .clamp(1, MAX_CODEC_LIMIT);
        let max_text = self
            .max_retained
            .saturating_sub(self.retained)
            .clamp(1, MAX_CODEC_LIMIT);
        Ok(chart_metadata_codec::DecodeOptions::new(
            max_input, max_fields, max_work, depth, max_labels, max_text,
        ))
    }

    /// Charge one borrowed metadata decode, including work spent before a
    /// rejected source is returned by the codec.
    pub(super) fn metadata_codec_report(
        &mut self,
        report: chart_metadata_codec::DecodeReport,
    ) -> Result<(), ChartArrangementError> {
        self.source(report.source_bytes())?;
        self.fields(report.fields())?;
        self.work(
            report
                .work_bytes()
                .checked_add(report.failure_work_bytes())
                .ok_or(ChartArrangementError::InvalidSource)?,
        )?;
        let depth = usize::try_from(report.max_depth()).unwrap_or(usize::MAX);
        if depth > self.max_nesting {
            return Err(ChartArrangementError::LimitExceeded {
                kind: ChartArrangementLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
        self.retained(report.text_bytes())?;
        self.retained(report.retained_bytes())?;
        self.allocations(report.allocations())
    }

    /// Build residual limits for the borrowed chart-data projection.
    ///
    /// The grid decoder has an independent cell ceiling in addition to its
    /// wire and text axes.  Keep that ceiling derived from the same finite
    /// field budget so a large package limit cannot turn into an unbounded
    /// rectangular allocation at the package boundary.
    pub(super) fn data_codec_options(
        &self,
    ) -> Result<chart_data_codec::DecodeOptions, ChartArrangementError> {
        let limits = self.residual_wire_limits()?;
        let depth = u32::try_from(limits.max_nesting().min(MAX_CODEC_NESTING as usize))
            .unwrap_or(MAX_CODEC_NESTING);
        let max_input = limits.max_input_bytes().clamp(1, MAX_CODEC_LIMIT);
        let max_fields = limits.max_fields().clamp(1, MAX_CODEC_LIMIT);
        let max_work = limits.max_rewrite_work().clamp(1, MAX_CODEC_LIMIT);
        let max_cells = limits.max_fields().clamp(1, MAX_CODEC_LIMIT);
        let max_labels = limits
            .max_fields()
            .min(MAX_METADATA_LABEL_COUNT)
            .clamp(1, MAX_CODEC_LIMIT);
        let max_text = self
            .max_retained
            .saturating_sub(self.retained)
            .clamp(1, MAX_CODEC_LIMIT);
        Ok(chart_data_codec::DecodeOptions::new(
            max_input, max_fields, max_work, depth, max_cells, max_labels, max_text,
        ))
    }

    /// Charge one chart-data decode against the selector's aggregate ledger.
    pub(super) fn data_codec_report(
        &mut self,
        report: chart_data_codec::DecodeReport,
    ) -> Result<(), ChartArrangementError> {
        self.source(report.source_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        // The codec's cell ceiling is independent from protobuf field count.
        // Charge one bounded unit for each validated cell before the package
        // adapter starts walking the borrowed rows to own them.
        self.work(report.cell_count())?;
        let depth = usize::try_from(report.max_depth()).unwrap_or(usize::MAX);
        if depth > self.max_nesting {
            return Err(ChartArrangementError::LimitExceeded {
                kind: ChartArrangementLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
        self.retained(report.text_bytes())?;
        self.retained(report.retained_bytes())?;
        self.allocations(report.allocations())
    }

    /// Charge the bounded wire work needed to walk the lazy chart-data views
    /// before any fallible semantic allocation starts.
    pub(super) fn data_materialization_work(
        &mut self,
        amount: usize,
    ) -> Result<(), ChartArrangementError> {
        self.work(amount)
    }

    /// Build the title codec's residual options from the same aggregate
    /// ledger used by chart metadata and chart Arrange reads.
    pub(super) fn title_codec_options(
        &self,
        source: &[u8],
    ) -> Result<litchi_iwa_protos::keynote_chart_title_codec::DecodeOptions, ChartArrangementError>
    {
        let limits = self.residual_wire_limits()?;
        let depth = u32::try_from(limits.max_nesting().min(MAX_CODEC_NESTING as usize))
            .unwrap_or(MAX_CODEC_NESTING);
        let max_input = limits
            .max_input_bytes()
            .min(source.len().max(1))
            .min(MAX_CODEC_LIMIT);
        let max_fields = limits.max_fields().clamp(1, MAX_CODEC_LIMIT);
        let max_work = limits.max_rewrite_work().clamp(1, MAX_CODEC_LIMIT);
        let max_text = self
            .max_retained
            .saturating_sub(self.retained)
            .clamp(1, MAX_CODEC_LIMIT);
        Ok(
            litchi_iwa_protos::keynote_chart_title_codec::DecodeOptions::new(
                max_input, max_fields, max_work, depth,
            )
            .with_max_output_bytes(max_text)
            .with_max_title_bytes(max_text),
        )
    }

    /// Charge a successful title projection against the shared wire ledger.
    pub(super) fn title_codec_report(
        &mut self,
        report: litchi_iwa_protos::keynote_chart_title_codec::DecodeReport,
    ) -> Result<(), ChartArrangementError> {
        self.source(report.source_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        let depth = usize::try_from(report.max_depth()).unwrap_or(usize::MAX);
        if depth > self.max_nesting {
            return Err(ChartArrangementError::LimitExceeded {
                kind: ChartArrangementLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
        Ok(())
    }

    /// Charge a conservative failed title decode when the title codec cannot
    /// return its final report. The selected extension has already passed the
    /// outer wire-view ledger; this reservation covers the title codec's
    /// strict scan and Buffa attempt without allowing an error path to bypass
    /// the aggregate budget.
    pub(super) fn title_codec_failure(
        &mut self,
        source: &[u8],
    ) -> Result<(), ChartArrangementError> {
        let fields = source.len().saturating_mul(2).max(1);
        let work = source.len().saturating_mul(4).max(1);
        self.source(source.len())?;
        self.fields(fields)?;
        self.work(work)
    }

    fn codec_report(&mut self, report: DecodeReport) -> Result<(), ChartArrangementError> {
        self.source(report.source_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())
    }

    fn codec_requirements(
        &mut self,
        requirements: RewriteExecutionRequirements,
    ) -> Result<(), ChartArrangementError> {
        self.output(requirements.output_bytes)?;
        self.fields(requirements.fields)?;
        self.work(requirements.work_bytes)?;
        let depth = usize::try_from(requirements.max_depth).unwrap_or(usize::MAX);
        if depth > self.max_nesting {
            return Err(ChartArrangementError::LimitExceeded {
                kind: ChartArrangementLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
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
        self.source(bytes)?;
        self.work(bytes)?;
        self.allocations(1)?;
        self.retained(bytes)
    }
}

impl Package {
    /// Read one sheet chart's Arrange-panel state through semantic selectors.
    pub fn sheet_chart_arrangement<'sheet>(
        &self,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        chart_selector: impl Into<ChartSelector>,
    ) -> Result<ChartArrangement, ChartArrangementError> {
        let mut budget = ChartBudget::for_package(self)?;
        Ok(select_chart_with_budget(
            self,
            sheet_selector.into(),
            chart_selector.into(),
            false,
            &mut budget,
        )?
        .before)
    }

    /// Read all ordinary chart Arrange-panel states in one checked sheet
    /// traversal.
    pub fn sheet_chart_arrangements<'sheet>(
        &self,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
    ) -> Result<Box<[ChartArrangement]>, ChartArrangementError> {
        let mut budget = ChartBudget::for_package(self)?;
        let charts = discover_sheet_charts(
            self,
            resolve_sheet_position(self, sheet_selector.into())?,
            false,
            &mut budget,
        )?;
        let mut arrangements = Vec::new();
        arrangements.try_reserve_exact(charts.len()).map_err(|_| {
            ChartArrangementError::Allocation {
                amount: charts.len(),
            }
        })?;
        budget.allocations(1)?;
        budget.retained(
            charts
                .len()
                .checked_mul(size_of::<ChartArrangement>())
                .ok_or(ChartArrangementError::InvalidSource)?,
        )?;
        for chart in charts {
            arrangements.push(chart.before);
        }
        Ok(arrangements.into_boxed_slice())
    }

    /// Begin an exact immutable edit of one sheet chart's Arrange-panel
    /// state.
    pub fn edit_sheet_chart_arrangement<'sheet>(
        &self,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        chart_selector: impl Into<ChartSelector>,
    ) -> Result<ChartArrangementEdit<'_>, ChartArrangementError> {
        ChartArrangementEdit::new(self, sheet_selector, chart_selector)
    }

    /// Apply an exact-source checked chart-arrangement patch.
    pub fn apply_sheet_chart_arrangement(
        &self,
        patch: &ChartArrangementPatch,
    ) -> Result<ChartArrangementCommit, ChartArrangementError> {
        let catalog = physical_catalog(self)?;
        let source_owner = catalog.__source_owner();
        if !patch.artifacts.authorizes_owner(&source_owner) {
            return Err(ChartArrangementError::PatchConflict);
        }
        let mut budget = ChartBudget::for_package(self)?;
        budget.source(self.source_bytes().len())?;
        let current = select_chart_with_budget(
            self,
            SheetSelector::position(patch.selection.sheet_position),
            ChartSelector::position(patch.selection.chart_position),
            true,
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(ChartArrangementError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(ChartArrangementCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: ChartArrangementDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartArrangementError::PatchConflict);
        }
        let target_owner = patch.artifacts.target_owner();
        budget.candidate_reopen(target_owner.len())?;
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| ChartArrangementError::Verification)?;
        let selected = select_chart_with_budget(
            &candidate,
            SheetSelector::position(patch.selection.sheet_position),
            ChartSelector::position(patch.selection.chart_position),
            true,
            &mut budget,
        )?;
        if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
            return Err(ChartArrangementError::Verification);
        }
        verify_locality(
            self,
            &candidate,
            &patch.selection,
            patch.target_payload.as_deref(),
            &mut budget,
        )?;
        Ok(ChartArrangementCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ChartArrangementDiagnostics::published(),
        })
    }
}

fn resolve_sheet_position(
    package: &Package,
    selector: SheetSelector<'_>,
) -> Result<Position, ChartArrangementError> {
    match selector {
        SheetSelector::Name(name) => {
            if name.is_empty() {
                return Err(ChartArrangementError::EmptySheetName);
            }
            let mut matches = package
                .document()
                .sheets()
                .iter()
                .filter(|sheet| sheet.name() == name);
            let Some(sheet) = matches.next() else {
                return Err(ChartArrangementError::SheetNameNotFound);
            };
            if matches.next().is_some() {
                return Err(ChartArrangementError::AmbiguousSelector);
            }
            Ok(Position::new(sheet.index()))
        },
        SheetSelector::Index(index) => {
            if package.document().sheets().get(index).is_none() {
                return Err(ChartArrangementError::SheetPositionNotFound {
                    position: Position::new(index),
                });
            }
            Ok(Position::new(index))
        },
    }
}

pub(super) fn select_chart_with_budget(
    package: &Package,
    sheet_selector: SheetSelector<'_>,
    chart_selector: ChartSelector,
    require_editable: bool,
    budget: &mut ChartBudget,
) -> Result<ChartSelection, ChartArrangementError> {
    let sheet_position = resolve_sheet_position(package, sheet_selector)?;
    let charts = discover_sheet_charts(package, sheet_position, require_editable, budget)?;
    charts.into_iter().nth(chart_selector.as_index()).ok_or(
        ChartArrangementError::ChartPositionNotFound {
            position: chart_selector.as_position(),
        },
    )
}

fn root_document_with_budget(
    package: &Package,
    budget: &mut ChartBudget,
) -> Result<numbers_sheet_order_codec::DocumentSheetOrderSnapshot, ChartArrangementError> {
    let payload =
        Package::root_document_payload(&package.state.components).map_err(map_package_error)?;
    let fields = payload
        .len()
        .saturating_mul(2)
        .clamp(1, WireLimits::MAX_FIELDS);
    let work = payload
        .len()
        .saturating_mul(4)
        .clamp(1, WireLimits::MAX_REWRITE_WORK);
    let references = package
        .state
        .options
        .semantic()
        .max_sheets()
        .saturating_add(1);
    let retained = payload
        .len()
        .checked_add(
            references
                .checked_mul(size_of::<numbers_sheet_order_codec::ReferenceSnapshot>())
                .ok_or(ChartArrangementError::InvalidSource)?,
        )
        .ok_or(ChartArrangementError::InvalidSource)?;
    ChartBudget::preflight(
        budget.input,
        payload.len(),
        budget.max_input,
        ChartArrangementLimitKind::WireBytes,
    )?;
    ChartBudget::preflight(
        budget.fields,
        fields,
        budget.max_fields,
        ChartArrangementLimitKind::WireFields,
    )?;
    ChartBudget::preflight(
        budget.work,
        work,
        budget.max_work,
        ChartArrangementLimitKind::WireWork,
    )?;
    budget.preflight_allocations(1)?;
    budget.preflight_retained(retained)?;
    let document = Package::root_document(&package.state.components).map_err(map_package_error)?;
    budget.source(payload.len())?;
    budget.fields(fields)?;
    budget.work(work)?;
    budget.allocations(1)?;
    budget.retained(retained)?;
    Ok(document)
}

fn discover_sheet_charts(
    package: &Package,
    sheet_position: Position,
    _require_editable: bool,
    budget: &mut ChartBudget,
) -> Result<Vec<ChartSelection>, ChartArrangementError> {
    let document = root_document_with_budget(package, budget)?;
    let sheet_references = document.sheet_references();
    let sheet_reference_count = sheet_references.len();
    ChartBudget::preflight(
        budget.references,
        sheet_reference_count,
        budget.max_references,
        ChartArrangementLimitKind::References,
    )?;
    ChartBudget::preflight(
        budget.work,
        sheet_reference_count,
        budget.max_work,
        ChartArrangementLimitKind::WireWork,
    )?;
    budget.references(sheet_reference_count)?;
    budget.work(sheet_reference_count)?;
    let sheet_reference = sheet_references.get(sheet_position.get()).ok_or(
        ChartArrangementError::SheetPositionNotFound {
            position: sheet_position,
        },
    )?;
    let sheet_identifier = sheet_reference.identifier();
    if sheet_identifier == 0 || sheet_reference.deprecated_is_external() == Some(true) {
        return Err(ChartArrangementError::InvalidSource);
    }
    if sheet_references
        .iter()
        .filter(|reference| reference.identifier() == sheet_identifier)
        .count()
        != 1
    {
        return Err(ChartArrangementError::InvalidSource);
    }
    let resolved_sheet = package
        .state
        .index
        .resolve_ref_id(&package.state.components, sheet_identifier)
        .map_err(map_package_error)?
        .ok_or(ChartArrangementError::InvalidSource)?;
    let sheet_object = object_from_resolved(package, resolved_sheet)?;
    if sheet_object.archive_info.identifier != Some(sheet_identifier) {
        return Err(ChartArrangementError::InvalidSource);
    }
    validate_message_metadata(sheet_object)?;
    let has_sheet = unique_typed_message(sheet_object, SHEET_MESSAGE_TYPE)?.is_some();
    let has_form_sheet =
        unique_typed_message(sheet_object, FORM_BASED_SHEET_MESSAGE_TYPE)?.is_some();
    if has_sheet == has_form_sheet {
        return Err(ChartArrangementError::InvalidSource);
    }
    let sheet_message_type = if has_sheet {
        SHEET_MESSAGE_TYPE
    } else {
        FORM_BASED_SHEET_MESSAGE_TYPE
    };
    let (_, sheet_message) = unique_typed_message(sheet_object, sheet_message_type)?
        .ok_or(ChartArrangementError::InvalidSource)?;
    let drawable_identifiers =
        sheet_drawable_identifiers(sheet_message_type, sheet_message.data.as_slice(), budget)?;
    let chart_capacity = drawable_identifiers.len();
    budget.preflight_allocations(1)?;
    budget.preflight_retained(
        chart_capacity
            .checked_mul(size_of::<ChartSelection>())
            .ok_or(ChartArrangementError::InvalidSource)?,
    )?;
    let mut charts = Vec::new();
    charts
        .try_reserve_exact(chart_capacity)
        .map_err(|_| ChartArrangementError::Allocation {
            amount: chart_capacity,
        })?;
    budget.allocations(1)?;
    budget.retained(
        chart_capacity
            .checked_mul(size_of::<ChartSelection>())
            .ok_or(ChartArrangementError::InvalidSource)?,
    )?;

    // A repeated sheet drawable reference would make semantic chart
    // positions ambiguous and could otherwise expose the same native owner
    // twice through the focused selector. Reserve the bounded duplicate
    // detector before traversing or publishing any chart selection.
    budget.preflight_allocations(1)?;
    budget.preflight_retained(
        chart_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(ChartArrangementError::InvalidSource)?,
    )?;
    let mut seen_drawables = Vec::new();
    seen_drawables
        .try_reserve_exact(chart_capacity)
        .map_err(|_| ChartArrangementError::Allocation {
            amount: chart_capacity,
        })?;
    budget.allocations(1)?;
    budget.retained(
        chart_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(ChartArrangementError::InvalidSource)?,
    )?;

    for drawable_identifier in drawable_identifiers {
        budget.work(seen_drawables.len())?;
        if seen_drawables.contains(&drawable_identifier) {
            return Err(ChartArrangementError::InvalidSource);
        }
        seen_drawables.push(drawable_identifier);
        budget.references(1)?;
        let resolved = package
            .state
            .index
            .resolve_ref_id(&package.state.components, drawable_identifier)
            .map_err(map_package_error)?
            .ok_or(ChartArrangementError::InvalidSource)?;
        let object = object_from_resolved(package, resolved)?;
        if object.archive_info.identifier != Some(drawable_identifier) {
            return Err(ChartArrangementError::InvalidSource);
        }
        let Some((message_index, message)) = unique_typed_message(object, CHART_MESSAGE_TYPE)?
        else {
            continue;
        };
        let info = object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(ChartArrangementError::InvalidSource)?;
        validate_chart_message_metadata(object, info, sheet_identifier)?;
        let parent = chart_parent_identifier(message.data.as_slice(), budget)?;
        if parent != sheet_identifier {
            return Err(ChartArrangementError::InvalidSource);
        }
        let before = decode_arrangement(message.data.as_slice(), budget)?;
        let component = package
            .state
            .components
            .catalog()
            .get_index(resolved.component_index)
            .ok_or(ChartArrangementError::InvalidSource)?;
        let component_name: Arc<str> = Arc::from(component.name());
        budget.retained(component.name().len())?;
        budget.allocations(1)?;
        let chart_position = Position::new(charts.len());
        charts.push(ChartSelection {
            sheet_position,
            chart_position,
            sheet_identifier,
            chart_identifier: drawable_identifier,
            component_index: resolved.component_index,
            object_index: resolved.object_index,
            message_index,
            component_name,
            before,
        });
    }
    Ok(charts)
}

fn sheet_drawable_identifiers(
    message_type: u32,
    source: &[u8],
    budget: &mut ChartBudget,
) -> Result<Vec<u64>, ChartArrangementError> {
    let maximum = budget.max_references.saturating_sub(budget.references);
    let mut count = 0usize;
    let preflight =
        preflight_wire_tree_with_limits(source, budget.residual_wire_limits()?, |visit| {
            let is_drawable = match message_type {
                SHEET_MESSAGE_TYPE => {
                    visit.path().is_empty() && visit.field().number() == SHEET_DRAWABLE_FIELD
                },
                FORM_BASED_SHEET_MESSAGE_TYPE => {
                    visit.path() == [FORM_SHEET_SUPER_FIELD]
                        && visit.field().number() == SHEET_DRAWABLE_FIELD
                },
                _ => {
                    return Err(CommonError::InvalidFormat(
                        "invalid Numbers sheet type".into(),
                    ));
                },
            };
            if is_drawable && visit.field().wire_type() != 2 {
                return Err(CommonError::InvalidFormat(
                    "invalid Numbers drawable reference".into(),
                ));
            }
            if is_drawable {
                count = count
                    .checked_add(1)
                    .ok_or_else(|| CommonError::InvalidFormat("drawable count overflow".into()))?;
                if count > maximum {
                    return Err(CommonError::LimitExceeded {
                        kind: CommonLimitKind::Fields,
                        observed: count,
                        limit: maximum,
                    });
                }
            }
            let descend = (message_type == FORM_BASED_SHEET_MESSAGE_TYPE
                && visit.path().is_empty()
                && visit.field().number() == FORM_SHEET_SUPER_FIELD)
                || is_drawable;
            Ok(if descend {
                WireDescent::Descend
            } else {
                WireDescent::Skip
            })
        })
        .map_err(map_common_error)?;
    budget.preflight_scan(preflight)?;
    let identifier_bound = count
        .checked_mul(size_of::<u64>())
        .ok_or(ChartArrangementError::InvalidSource)?;
    budget.preflight_allocations(3)?;
    budget.preflight_retained(
        source
            .len()
            .checked_add(identifier_bound)
            .and_then(|amount| amount.checked_add(source.len()))
            .ok_or(ChartArrangementError::InvalidSource)?,
    )?;
    budget.charge_preflight_scan(preflight)?;
    // `names::preflight_sheet_payload` performs its own bounded tree walk
    // while extracting canonical identifiers. Charge an equal or larger
    // focused report before and after that call; this walk descends into every
    // drawable envelope, so its report is a conservative bound for the
    // names projection's shallower pass.
    budget.preflight_scan(preflight)?;
    let (_name, identifiers) = super::names::preflight_sheet_payload(message_type, source, maximum)
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    budget.charge_preflight_scan(preflight)?;
    if identifiers.len() != count {
        return Err(ChartArrangementError::InvalidSource);
    }
    budget.allocations(3)?;
    budget.retained(
        source
            .len()
            .checked_add(
                identifiers
                    .len()
                    .checked_mul(size_of::<u64>())
                    .ok_or(ChartArrangementError::InvalidSource)?,
            )
            .ok_or(ChartArrangementError::InvalidSource)?,
    )?;
    budget.references(identifiers.len())?;
    budget.work(identifiers.len())?;
    Ok(identifiers)
}

pub(super) fn object_from_resolved<'a>(
    package: &'a Package,
    resolved: crate::package::index::Resolved<'a>,
) -> Result<&'a ArchiveObject, ChartArrangementError> {
    package
        .state
        .components
        .catalog()
        .get_index(resolved.component_index)
        .and_then(|component| component.archive().objects.get(resolved.object_index))
        .ok_or(ChartArrangementError::InvalidSource)
}

pub(super) fn unique_typed_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<Option<(usize, &RawMessage)>, ChartArrangementError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace((index, message)).is_some() {
            return Err(ChartArrangementError::InvalidSource);
        }
    }
    Ok(selected)
}

pub(super) fn validate_message_metadata(
    object: &ArchiveObject,
) -> Result<(), ChartArrangementError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(ChartArrangementError::InvalidSource);
    }
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(ChartArrangementError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_chart_message_metadata(
    object: &ArchiveObject,
    info: &litchi_iwa_core::MessageInfo,
    sheet_identifier: u64,
) -> Result<(), ChartArrangementError> {
    if info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
        || object.archive_info.should_merge == Some(true)
    {
        return Err(ChartArrangementError::InvalidSource);
    }
    if !chart_parent_metadata_is_owned(info, sheet_identifier) {
        return Err(ChartArrangementError::InvalidSource);
    }
    Ok(())
}

fn chart_parent_metadata_is_owned(info: &litchi_iwa_core::MessageInfo, identifier: u64) -> bool {
    if identifier == 0
        || info.data_references.contains(&identifier)
        || info
            .field_infos
            .iter()
            .any(|field| field.data_references.contains(&identifier))
    {
        return false;
    }
    let aggregate = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    let mut field_occurrence = 0usize;
    for field in &info.field_infos {
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if field.path.as_slice() != [CHART_DRAWABLE_SUPER_FIELD, DRAWABLE_PARENT_FIELD]
            || occurrences != 1
        {
            return false;
        }
        field_occurrence = field_occurrence.saturating_add(occurrences);
    }
    matches!((aggregate, field_occurrence), (0, 0) | (1, 0 | 1))
}

fn chart_parent_identifier(
    payload: &[u8],
    budget: &mut ChartBudget,
) -> Result<u64, ChartArrangementError> {
    let root = parse_wire_view_with_budget(payload, 1, budget)?;
    let mut super_payload = None;
    for field in root.fields() {
        if field.number() != CHART_DRAWABLE_SUPER_FIELD {
            continue;
        }
        if super_payload.is_some() || field.wire_type() != 2 {
            return Err(ChartArrangementError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(map_common_error)?;
        super_payload = Some(field.payload());
    }
    let super_payload = super_payload.ok_or(ChartArrangementError::InvalidSource)?;
    let drawable = parse_wire_view_with_budget(super_payload, 2, budget)?;
    let mut parent_payload = None;
    for field in drawable.fields() {
        if field.number() != DRAWABLE_PARENT_FIELD {
            continue;
        }
        if parent_payload.is_some() || field.wire_type() != 2 {
            return Err(ChartArrangementError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(map_common_error)?;
        parent_payload = Some(field.payload());
    }
    let parent_payload = parent_payload.ok_or(ChartArrangementError::InvalidSource)?;
    let reference = parse_wire_view_with_budget(parent_payload, 3, budget)?;
    let mut identifier = None;
    for field in reference.fields() {
        if field.number() != REFERENCE_IDENTIFIER_FIELD {
            continue;
        }
        if identifier.is_some() || field.wire_type() != 0 {
            return Err(ChartArrangementError::InvalidSource);
        }
        field.validate_canonical_key().map_err(map_common_error)?;
        let (value, consumed) = decode_varint_from_bytes(field.payload())
            .map_err(|_| ChartArrangementError::InvalidSource)?;
        if consumed != field.payload().len()
            || litchi_iwa_common::varint::encoded_len(value) != consumed
            || value == 0
        {
            return Err(ChartArrangementError::InvalidSource);
        }
        identifier = Some(value);
    }
    identifier.ok_or(ChartArrangementError::InvalidSource)
}

pub(super) fn parse_wire_view_with_budget<'source>(
    payload: &'source [u8],
    depth: usize,
    budget: &mut ChartBudget,
) -> Result<WireView<'source>, ChartArrangementError> {
    // Count only top-level fields before constructing `WireView`; its private
    // span vector grows with fields, rather than with the size of an opaque
    // length-delimited payload. The preflight itself is allocation-free but
    // still belongs to the transaction's source/work ledger.
    let preflight =
        preflight_wire_tree_with_limits(payload, budget.residual_wire_limits()?, |_visit| {
            Ok(WireDescent::Skip)
        })
        .map_err(map_common_error)?;
    budget.preflight_scan(preflight)?;
    let span_bytes = preflight
        .fields()
        .checked_mul(4)
        .and_then(|capacity| capacity.checked_mul(WIRE_VIEW_SPAN_BYTES_PER_FIELD))
        .ok_or(ChartArrangementError::InvalidSource)?;
    // The span parser reserves incrementally; bound growth events and the
    // cumulative capacities, including its initial four-slot allocation.
    let span_allocations = preflight.fields();
    budget.preflight_allocations(span_allocations)?;
    budget.preflight_retained(span_bytes)?;
    budget.charge_preflight_scan(preflight)?;
    let view = WireView::parse_with_limits(payload, budget.residual_wire_limits()?)
        .map_err(map_common_error)?;
    budget.allocations(span_allocations)?;
    budget.retained(span_bytes)?;
    budget.scan(payload, view.len(), depth)?;
    Ok(view)
}

fn decode_arrangement(
    payload: &[u8],
    budget: &mut ChartBudget,
) -> Result<ChartArrangement, ChartArrangementError> {
    let options = budget.codec_options(payload)?;
    let (snapshot, report) =
        codec::decode_chart_arrangement_with_report(payload, options).map_err(map_codec_error)?;
    budget.codec_report(report)?;
    if !snapshot.has_drawable() {
        return Err(ChartArrangementError::InvalidSource);
    }
    Ok(ChartArrangement::new(
        snapshot.is_locked(),
        snapshot.is_constrained(),
    ))
}

fn same_selection(left: &ChartSelection, right: &ChartSelection) -> bool {
    left.sheet_position == right.sheet_position
        && left.chart_position == right.chart_position
        && left.sheet_identifier == right.sheet_identifier
        && left.chart_identifier == right.chart_identifier
        && left.component_index == right.component_index
        && left.object_index == right.object_index
        && left.message_index == right.message_index
        && left.component_name == right.component_name
}

fn commit_edit(
    source: &Package,
    selection: &ChartSelection,
    after: ChartArrangement,
) -> Result<ChartArrangementCommit, ChartArrangementError> {
    let catalog = physical_catalog(source)?;
    let source_owner = catalog.__source_owner();
    let mut budget = ChartBudget::for_package(source)?;
    budget.source(source.source_bytes().len())?;
    let current = select_chart_with_budget(
        source,
        SheetSelector::position(selection.sheet_position),
        ChartSelector::position(selection.chart_position),
        true,
        &mut budget,
    )?;
    if !same_selection(&current, selection) || current.before != selection.before {
        return Err(ChartArrangementError::InvalidSource);
    }
    if selection.before == after {
        return Ok(ChartArrangementCommit {
            package: source.snapshot(),
            patch: ChartArrangementPatch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                selection: selection.clone(),
                before: selection.before,
                after,
                source_payload: None,
                target_payload: None,
            },
            diagnostics: ChartArrangementDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(ChartArrangementError::UnsupportedSource);
    }
    let candidate = rewrite_chart(source, selection, after, &mut budget)?;
    let target_catalog = physical_catalog(&candidate)?;
    let target_owner = target_catalog.__source_owner();
    let selected = select_chart_with_budget(
        &candidate,
        SheetSelector::position(selection.sheet_position),
        ChartSelector::position(selection.chart_position),
        true,
        &mut budget,
    )?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(ChartArrangementError::Verification);
    }
    let source_payload = selected_chart_payload(source, selection)?;
    let target_payload = selected_chart_payload(&candidate, selection)?;
    verify_locality(
        source,
        &candidate,
        selection,
        Some(target_payload),
        &mut budget,
    )?;
    let (source_payload, target_payload) =
        retain_payload_pair(source_payload, target_payload, &mut budget)?;
    Ok(ChartArrangementCommit {
        package: candidate,
        patch: ChartArrangementPatch {
            artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            selection: selection.clone(),
            before: selection.before,
            after,
            source_payload: Some(source_payload),
            target_payload: Some(target_payload),
        },
        diagnostics: ChartArrangementDiagnostics::published(),
    })
}

fn rewrite_chart(
    source: &Package,
    selection: &ChartSelection,
    after: ChartArrangement,
    budget: &mut ChartBudget,
) -> Result<Package, ChartArrangementError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.component_name.as_ref())
        .ok_or(ChartArrangementError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(ChartArrangementError::UnsupportedSource);
    }
    let component = package_component(source, selection.component_index)?;
    if component.name() != selection.component_name.as_ref() {
        return Err(ChartArrangementError::InvalidSource);
    }
    let source_archive = component.archive();
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    let snappy_limits = source
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    let object = source_archive
        .objects
        .get(selection.object_index)
        .ok_or(ChartArrangementError::InvalidSource)?;
    if object.archive_info.identifier != Some(selection.chart_identifier) {
        return Err(ChartArrangementError::InvalidSource);
    }
    validate_message_metadata(object)?;
    let (message_index, message) = unique_typed_message(object, CHART_MESSAGE_TYPE)?
        .ok_or(ChartArrangementError::InvalidSource)?;
    if message_index != selection.message_index {
        return Err(ChartArrangementError::InvalidSource);
    }
    validate_chart_message_metadata(
        object,
        object
            .archive_info
            .message_infos
            .get(message_index)
            .ok_or(ChartArrangementError::InvalidSource)?,
        selection.sheet_identifier,
    )?;
    let original = message.data.as_slice();
    let before = decode_arrangement(original, budget)?;
    if before != selection.before {
        return Err(ChartArrangementError::InvalidSource);
    }
    let source_encoded_len = source_archive
        .encoded_len_with_limits(archive_limits)
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    let options = budget.codec_options(original)?;
    let prepared = codec::prepare_chart_arrangement_rewrite(
        original,
        ChartArrangementWrite::new(after.locked(), after.constrain_proportions()),
        options,
    )
    .map_err(map_codec_error)?;
    budget.codec_report(prepared.prepare_report())?;
    let requirements = prepared.execution_requirements();
    budget.codec_requirements(requirements)?;
    let encoded_bound = rewritten_archive_bound(
        source_encoded_len,
        original.len(),
        requirements.output_bytes,
    )?;
    let compressed_bound = SnappyStream::maximum_compressed_len(encoded_bound)
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    let (snappy_allocations, snappy_scratch) = snappy_compression_requirements(encoded_bound)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(ChartArrangementError::LimitExceeded {
            kind: ChartArrangementLimitKind::OutputBytes,
            observed: compressed_bound as u64,
            maximum: snappy_limits.max_compressed_stream() as u64,
        });
    }
    budget.preflight_output(encoded_bound)?;
    budget.preflight_output(compressed_bound)?;
    budget.preflight_work(encoded_bound.saturating_add(compressed_bound))?;
    let (clone_allocations, clone_retained, clone_work) =
        archive_clone_requirements(source_archive, source_encoded_len)?;
    let (header_allocations, header_retained, header_work) =
        archive_header_rewrite_requirements(object)?;
    let (serialization_allocations, serialization_retained, serialization_work) =
        archive_serialization_requirements(source_archive)?;
    let total_allocations = clone_allocations
        .checked_add(header_allocations)
        .and_then(|amount| amount.checked_add(serialization_allocations))
        // `Archive::to_bytes_with_limits` creates the final decompressed
        // archive buffer, followed by Snappy's output buffer and one
        // temporary compressed buffer per write chunk.
        .and_then(|amount| amount.checked_add(1))
        .and_then(|amount| amount.checked_add(snappy_allocations))
        .ok_or(ChartArrangementError::InvalidSource)?;
    let total_work = clone_work
        .checked_add(header_work)
        .and_then(|amount| amount.checked_add(serialization_work))
        .ok_or(ChartArrangementError::InvalidSource)?;
    budget.preflight_allocations(total_allocations)?;
    budget.preflight_work(total_work)?;
    let temporary_retained = clone_retained
        .checked_add(header_retained)
        .and_then(|amount| amount.checked_add(serialization_retained))
        .and_then(|amount| amount.checked_add(encoded_bound))
        .and_then(|amount| amount.checked_add(compressed_bound))
        .and_then(|amount| amount.checked_add(snappy_scratch))
        .ok_or(ChartArrangementError::InvalidSource)?;
    budget.preflight_retained(temporary_retained)?;
    budget.preflight_scratch(snappy_scratch)?;
    budget.allocations(total_allocations)?;
    budget.work(total_work)?;
    budget.retained(clone_retained)?;
    budget.retained(header_retained)?;
    budget.retained(serialization_retained)?;
    budget.scratch(snappy_scratch)?;
    let mut archive = source_archive.clone();
    let rewritten = prepared
        .execute(requirements.exact_limits())
        .map_err(map_codec_error)?
        .into_output();
    let verified = decode_arrangement(&rewritten, budget)?;
    if verified != after {
        return Err(ChartArrangementError::Verification);
    }
    archive
        .objects
        .get_mut(selection.object_index)
        .ok_or(ChartArrangementError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    budget.output(bytes.len())?;
    budget.retained(bytes.len())?;
    let compressed =
        SnappyStream::compress(&bytes).map_err(|_| ChartArrangementError::InvalidSource)?;
    budget.retained(compressed.len())?;
    let edits = [EntryEdit::new(
        selection.component_name.as_ref(),
        &compressed,
    )];
    let prepared_reassembly = catalog
        .prepare_reassembly_with_deletions(&edits, &[], source.state.options.archive())
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    let reassembly_requirements = prepared_reassembly.execution_requirements();
    budget.reassembly(reassembly_requirements)?;
    budget.candidate_reopen(reassembly_requirements.output_bytes())?;
    let output = prepared_reassembly
        .execute(reassembly_requirements.exact_limits())
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    Package::from_owned_bytes_with_options(output, source.state.options)
        .map_err(|_| ChartArrangementError::Verification)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &ChartSelection,
    expected_payload: Option<&[u8]>,
    budget: &mut ChartBudget,
) -> Result<(), ChartArrangementError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    if source_catalog.package().len() != candidate_catalog.package().len() {
        return Err(ChartArrangementError::Verification);
    }
    let selected_component = selection.component_name.as_ref();
    let mut entry_work = 0usize;
    for (source_entry, candidate_entry) in source_catalog
        .package()
        .iter()
        .zip(candidate_catalog.package().iter())
    {
        let source_entry_work =
            zip_entry_work(source_entry, source_entry.name() != selected_component)?;
        let candidate_entry_work = zip_entry_work(
            candidate_entry,
            candidate_entry.name() != selected_component,
        )?;
        entry_work = entry_work
            .checked_add(source_entry_work)
            .and_then(|amount| amount.checked_add(candidate_entry_work))
            .ok_or(ChartArrangementError::InvalidSource)?;
    }
    budget.preflight_work(entry_work)?;
    budget.work(entry_work)?;
    for (source_entry, candidate_entry) in source_catalog
        .package()
        .iter()
        .zip(candidate_catalog.package().iter())
    {
        if source_entry.name() != candidate_entry.name() {
            return Err(ChartArrangementError::Verification);
        }
        if !same_entry_fixed_metadata(source_entry, candidate_entry) {
            return Err(ChartArrangementError::Verification);
        }
        if source_entry.name() == selected_component {
            if !same_selected_entry_records(source_entry, candidate_entry) {
                return Err(ChartArrangementError::Verification);
            }
        } else if !same_unselected_entry_records(source_entry, candidate_entry) {
            return Err(ChartArrangementError::Verification);
        }
    }
    let source_archive = package_component(source, selection.component_index)?.archive();
    let candidate_archive = package_component(candidate, selection.component_index)?.archive();
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(ChartArrangementError::Verification);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    let mut object_work = 0usize;
    let mut selected_clone = None;
    for (object_index, source_object) in source_archive.objects.iter().enumerate() {
        let candidate_object = candidate_archive
            .objects
            .get(object_index)
            .ok_or(ChartArrangementError::Verification)?;
        let source_object_work = archive_object_work(source_object)?;
        let candidate_object_work = archive_object_work(candidate_object)?;
        object_work = object_work
            .checked_add(source_object_work)
            .and_then(|amount| amount.checked_add(candidate_object_work))
            .ok_or(ChartArrangementError::InvalidSource)?;
        if object_index == selection.object_index {
            let candidate_message = candidate_object
                .messages
                .get(selection.message_index)
                .ok_or(ChartArrangementError::Verification)?;
            if candidate_message.type_ != CHART_MESSAGE_TYPE {
                return Err(ChartArrangementError::Verification);
            }
            let requirements = archive_object_clone_requirements(
                source_object,
                candidate_object,
                candidate_message.data.len(),
            )?;
            object_work = object_work
                .checked_add(requirements.2)
                .ok_or(ChartArrangementError::InvalidSource)?;
            selected_clone = Some(requirements);
        }
    }
    if selected_clone.is_none() {
        return Err(ChartArrangementError::Verification);
    }
    budget.preflight_work(object_work)?;
    if let Some((allocations, retained, _work)) = selected_clone {
        budget.preflight_allocations(allocations)?;
        budget.preflight_retained(retained)?;
        budget.allocations(allocations)?;
        budget.retained(retained)?;
    }
    budget.work(object_work)?;
    for (object_index, source_object) in source_archive.objects.iter().enumerate() {
        let candidate_object = candidate_archive
            .objects
            .get(object_index)
            .ok_or(ChartArrangementError::Verification)?;
        if object_index != selection.object_index {
            if !source_object.same_content_ignoring_offsets(candidate_object) {
                return Err(ChartArrangementError::Verification);
            }
            continue;
        }
        let candidate_message = candidate_object
            .messages
            .get(selection.message_index)
            .ok_or(ChartArrangementError::Verification)?;
        if candidate_message.type_ != CHART_MESSAGE_TYPE {
            return Err(ChartArrangementError::Verification);
        }
        if let Some(expected_payload) = expected_payload
            && candidate_message.data.as_slice() != expected_payload
        {
            return Err(ChartArrangementError::Verification);
        }
        let mut expected = source_object.clone();
        expected
            .replace_message_preserving_header_with_limits(
                selection.message_index,
                RawMessage {
                    type_: CHART_MESSAGE_TYPE,
                    data: candidate_message.data.clone(),
                },
                archive_limits,
            )
            .map_err(|_| ChartArrangementError::Verification)?;
        expected.header_length = candidate_object.header_length;
        expected.data_length = candidate_object.data_length;
        if !expected.same_content_ignoring_offsets(candidate_object) {
            return Err(ChartArrangementError::Verification);
        }
    }
    Ok(())
}

fn zip_entry_work(entry: &Entry, include_data: bool) -> Result<usize, ChartArrangementError> {
    let metadata = entry.metadata();
    let mut work = entry
        .name()
        .len()
        .checked_add(entry.raw_name().len())
        .and_then(|amount| amount.checked_add(entry.raw_record().local_record().len()))
        .and_then(|amount| amount.checked_add(entry.raw_record().central_directory_record().len()))
        .ok_or(ChartArrangementError::InvalidSource)?;
    if include_data {
        work = work
            .checked_add(entry.data().len())
            .ok_or(ChartArrangementError::InvalidSource)?;
    }
    for header in [metadata.local(), metadata.central()] {
        work = work
            .checked_add(header.name().len())
            .and_then(|amount| amount.checked_add(header.extra().len()))
            .and_then(|amount| amount.checked_add(header.comment().len()))
            .ok_or(ChartArrangementError::InvalidSource)?;
    }
    Ok(work)
}

fn same_header_metadata(
    left: &litchi_iwa_archive::package::HeaderMetadata,
    right: &litchi_iwa_archive::package::HeaderMetadata,
) -> bool {
    left.version_needed() == right.version_needed()
        && left.flags() == right.flags()
        && left.compression_method() == right.compression_method()
        && left.last_modified() == right.last_modified()
        && left.name() == right.name()
        && left.extra() == right.extra()
        && left.comment() == right.comment()
}

fn same_entry_fixed_metadata(left: &Entry, right: &Entry) -> bool {
    left.name() == right.name()
        && left.raw_name() == right.raw_name()
        && left.is_opaque() == right.is_opaque()
        && same_header_metadata(left.metadata().local(), right.metadata().local())
        && same_header_metadata(left.metadata().central(), right.metadata().central())
}

fn same_selected_entry_records(left: &Entry, right: &Entry) -> bool {
    same_selected_central_record(
        left.raw_record().central_directory_record(),
        right.raw_record().central_directory_record(),
    )
}

fn same_selected_central_record(left: &[u8], right: &[u8]) -> bool {
    // Recompression may change CRC and the two ZIP size fields (16..28), but
    // it must preserve every other central-directory byte, including the
    // selected member's DOS timestamp (14..16) and local-header offset
    // (42..46).
    left.len() >= 46
        && right.len() == left.len()
        && left[..16] == right[..16]
        && left[28..] == right[28..]
}

#[cfg(test)]
mod tests {
    use super::same_selected_central_record;

    #[test]
    fn selected_central_record_mask_preserves_timestamp_and_allows_sizes() {
        let source: Vec<u8> = (0..46).map(|value| value as u8).collect();

        let mut timestamp = source.clone();
        timestamp[14] ^= 1;
        assert!(!same_selected_central_record(&source, &timestamp));

        let mut crc = source.clone();
        crc[16] ^= 1;
        assert!(same_selected_central_record(&source, &crc));

        let mut high_size = source.clone();
        high_size[26] ^= 1;
        high_size[27] ^= 1;
        assert!(same_selected_central_record(&source, &high_size));

        let mut filename_length = source.clone();
        filename_length[28] ^= 1;
        assert!(!same_selected_central_record(&source, &filename_length));
    }
}

fn same_unselected_entry_records(left: &Entry, right: &Entry) -> bool {
    if left.data() != right.data()
        || left.metadata() != right.metadata()
        || left.raw_record().local_record() != right.raw_record().local_record()
    {
        return false;
    }
    let left_central = left.raw_record().central_directory_record();
    let right_central = right.raw_record().central_directory_record();
    // Reassembly can move the local record of an unselected member after the
    // edited member. ZIP central-directory bytes 42..46 are the sole allowed
    // difference; all raw records and physical metadata remain authoritative.
    left_central.len() >= 46
        && right_central.len() == left_central.len()
        && left_central[..42] == right_central[..42]
        && left_central[46..] == right_central[46..]
}

fn package_component(
    package: &Package,
    component_index: usize,
) -> Result<&litchi_iwa_archive::Component, ChartArrangementError> {
    package
        .state
        .components
        .catalog()
        .get_index(component_index)
        .ok_or(ChartArrangementError::InvalidSource)
}

pub(super) fn selected_chart_payload<'source>(
    package: &'source Package,
    selection: &ChartSelection,
) -> Result<&'source [u8], ChartArrangementError> {
    let object = package_component(package, selection.component_index)?
        .archive()
        .objects
        .get(selection.object_index)
        .ok_or(ChartArrangementError::InvalidSource)?;
    if object.archive_info.identifier != Some(selection.chart_identifier) {
        return Err(ChartArrangementError::InvalidSource);
    }
    let (message_index, message) = unique_typed_message(object, CHART_MESSAGE_TYPE)?
        .ok_or(ChartArrangementError::InvalidSource)?;
    if message_index != selection.message_index {
        return Err(ChartArrangementError::InvalidSource);
    }
    Ok(message.data.as_slice())
}

fn retain_payload_pair(
    source: &[u8],
    target: &[u8],
    budget: &mut ChartBudget,
) -> Result<(Arc<[u8]>, Arc<[u8]>), ChartArrangementError> {
    let allocations = usize::from(!source.is_empty()) + usize::from(!target.is_empty());
    let retained = source
        .len()
        .checked_add(target.len())
        .ok_or(ChartArrangementError::InvalidSource)?;
    budget.preflight_allocations(allocations)?;
    budget.preflight_retained(retained)?;
    let source = Arc::<[u8]>::from(source);
    let target = Arc::<[u8]>::from(target);
    budget.allocations(allocations)?;
    budget.retained(retained)?;
    Ok((source, target))
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, ChartArrangementError> {
    package
        .state
        .components
        .physical()
        .ok_or(ChartArrangementError::UnsupportedSource)
}

fn rewritten_archive_bound(
    source_encoded_len: usize,
    original_len: usize,
    replacement_len: usize,
) -> Result<usize, ChartArrangementError> {
    source_encoded_len
        .checked_sub(original_len)
        .and_then(|value| value.checked_add(replacement_len))
        .and_then(|value| value.checked_add(64))
        .ok_or(ChartArrangementError::InvalidSource)
}

/// Bound the allocations and transient workspace used by the final IWA
/// Snappy write. `SnappyStream::compress` reserves one output vector and
/// obtains one temporary compressed vector for each independently encoded
/// chunk. The neutral stream API exposes the chunk size and output bound, so
/// account for both before invoking the allocation-bearing compressor.
fn snappy_compression_requirements(
    input_len: usize,
) -> Result<(usize, usize), ChartArrangementError> {
    let frame_count = if input_len == 0 {
        0
    } else {
        input_len
            .checked_add(SnappyStream::WRITE_CHUNK_SIZE - 1)
            .and_then(|length| length.checked_div(SnappyStream::WRITE_CHUNK_SIZE))
            .ok_or(ChartArrangementError::InvalidSource)?
    };
    let chunk_output = if input_len == 0 {
        0
    } else {
        SnappyStream::maximum_compressed_len(input_len.min(SnappyStream::WRITE_CHUNK_SIZE))
            .map_err(|_| ChartArrangementError::InvalidSource)?
            .checked_sub(4)
            .ok_or(ChartArrangementError::InvalidSource)?
    };
    let encoder_workspace = usize::from(input_len != 0)
        .checked_mul(SNAPPY_ENCODER_WORKSPACE_BYTES)
        .ok_or(ChartArrangementError::InvalidSource)?;
    let allocations = frame_count
        .checked_add(1)
        .and_then(|amount| amount.checked_add(usize::from(input_len != 0)))
        .ok_or(ChartArrangementError::InvalidSource)?;
    let scratch = chunk_output
        .checked_add(encoder_workspace)
        .ok_or(ChartArrangementError::InvalidSource)?;
    Ok((allocations, scratch))
}

fn archive_header_rewrite_requirements(
    object: &ArchiveObject,
) -> Result<(usize, usize, usize), ChartArrangementError> {
    let header =
        usize::try_from(object.header_length).map_err(|_| ChartArrangementError::InvalidSource)?;
    let retained = header
        .checked_add(64)
        .and_then(|value| value.checked_mul(3))
        .ok_or(ChartArrangementError::InvalidSource)?;
    Ok((3, retained, archive_object_work(object)?))
}

fn archive_serialization_requirements(
    archive: &Archive,
) -> Result<(usize, usize, usize), ChartArrangementError> {
    let mut retained = 0usize;
    let mut work = 0usize;
    for object in &archive.objects {
        let header = usize::try_from(object.header_length)
            .map_err(|_| ChartArrangementError::InvalidSource)?;
        retained = retained
            .checked_add(header)
            .and_then(|amount| amount.checked_add(64))
            .ok_or(ChartArrangementError::InvalidSource)?;
        work = work
            .checked_add(archive_object_work(object)?)
            .ok_or(ChartArrangementError::InvalidSource)?;
    }
    let allocations = archive
        .objects
        .len()
        .checked_add(2)
        .ok_or(ChartArrangementError::InvalidSource)?;
    Ok((allocations, retained, work))
}

fn archive_clone_requirements(
    archive: &Archive,
    encoded_len: usize,
) -> Result<(usize, usize, usize), ChartArrangementError> {
    let mut allocations = 0usize;
    let mut retained = encoded_len;
    let mut work = 0usize;
    add_clone_vec::<ArchiveObject>(archive.objects.len(), &mut allocations, &mut retained)?;
    for object in &archive.objects {
        work = work
            .checked_add(archive_object_work(object)?)
            .ok_or(ChartArrangementError::InvalidSource)?;
        add_clone_vec::<RawMessage>(object.messages.len(), &mut allocations, &mut retained)?;
        add_clone_vec::<litchi_iwa_core::MessageInfo>(
            object.archive_info.message_infos.len(),
            &mut allocations,
            &mut retained,
        )?;
        if object.header_length != 0 {
            allocations = allocations
                .checked_add(2)
                .ok_or(ChartArrangementError::InvalidSource)?;
            retained = retained
                .checked_add(
                    usize::try_from(object.header_length)
                        .map_err(|_| ChartArrangementError::InvalidSource)?
                        .checked_mul(2)
                        .ok_or(ChartArrangementError::InvalidSource)?,
                )
                .ok_or(ChartArrangementError::InvalidSource)?;
        }
        for message in &object.messages {
            add_clone_bytes(message.data.len(), &mut allocations, &mut retained)?;
        }
        for info in &object.archive_info.message_infos {
            add_clone_vec::<u32>(info.versions.len(), &mut allocations, &mut retained)?;
            add_clone_vec::<litchi_iwa_core::FieldInfo>(
                info.field_infos.len(),
                &mut allocations,
                &mut retained,
            )?;
            add_clone_vec::<u64>(
                info.object_references.len(),
                &mut allocations,
                &mut retained,
            )?;
            add_clone_vec::<u64>(info.data_references.len(), &mut allocations, &mut retained)?;
            add_clone_vec::<u32>(
                info.diff_merge_version.len(),
                &mut allocations,
                &mut retained,
            )?;
            if let Some(path) = info.diff_field_path.as_ref() {
                add_clone_vec::<u32>(path.path.len(), &mut allocations, &mut retained)?;
            }
            add_clone_vec::<litchi_iwa_core::FieldPath>(
                info.fields_to_remove.len(),
                &mut allocations,
                &mut retained,
            )?;
            for path in &info.fields_to_remove {
                add_clone_vec::<u32>(path.path.len(), &mut allocations, &mut retained)?;
            }
            add_clone_vec::<u32>(
                info.diff_read_version.len(),
                &mut allocations,
                &mut retained,
            )?;
            for field in &info.field_infos {
                add_clone_vec::<u32>(field.path.path.len(), &mut allocations, &mut retained)?;
                add_clone_vec::<u64>(
                    field.object_references.len(),
                    &mut allocations,
                    &mut retained,
                )?;
                add_clone_vec::<u64>(field.data_references.len(), &mut allocations, &mut retained)?;
                add_clone_vec::<u32>(
                    field.known_field_version.len(),
                    &mut allocations,
                    &mut retained,
                )?;
                if let Some(identifier) = field.known_field_feature_identifier.as_ref() {
                    add_clone_bytes(identifier.len(), &mut allocations, &mut retained)?;
                }
            }
        }
    }
    Ok((allocations, retained, work))
}

/// Bound the full decoded object walk performed by archive cloning,
/// serialization, header replacement, and physical-content comparison.
///
/// The archive core keeps nested metadata in several vectors. Counting only
/// top-level objects or messages would admit a source whose metadata walk can
/// exhaust the focused transaction ledger before the first checked
/// allocation. This helper charges every nested vector and scalar byte
/// inspected by those operations.
fn archive_object_work(object: &ArchiveObject) -> Result<usize, ChartArrangementError> {
    let mut work = size_of::<ArchiveObject>();
    work = work
        .checked_add(object.messages.len())
        .and_then(|amount| amount.checked_add(object.archive_info.message_infos.len()))
        .ok_or(ChartArrangementError::InvalidSource)?;
    for message in &object.messages {
        work = work
            .checked_add(message.data.len())
            .and_then(|amount| amount.checked_add(size_of::<RawMessage>()))
            .ok_or(ChartArrangementError::InvalidSource)?;
    }
    for info in &object.archive_info.message_infos {
        work = work
            .checked_add(size_of::<litchi_iwa_core::MessageInfo>())
            .and_then(|amount| amount.checked_add(info.versions.len()))
            .and_then(|amount| amount.checked_add(info.field_infos.len()))
            .and_then(|amount| amount.checked_add(info.object_references.len()))
            .and_then(|amount| amount.checked_add(info.data_references.len()))
            .and_then(|amount| amount.checked_add(info.diff_merge_version.len()))
            .and_then(|amount| amount.checked_add(info.fields_to_remove.len()))
            .and_then(|amount| amount.checked_add(info.diff_read_version.len()))
            .ok_or(ChartArrangementError::InvalidSource)?;
        if let Some(path) = info.diff_field_path.as_ref() {
            work = work
                .checked_add(path.path.len())
                .ok_or(ChartArrangementError::InvalidSource)?;
        }
        for path in &info.fields_to_remove {
            work = work
                .checked_add(path.path.len())
                .ok_or(ChartArrangementError::InvalidSource)?;
        }
        for field in &info.field_infos {
            work = work
                .checked_add(size_of::<litchi_iwa_core::FieldInfo>())
                .and_then(|amount| amount.checked_add(field.path.path.len()))
                .and_then(|amount| amount.checked_add(field.object_references.len()))
                .and_then(|amount| amount.checked_add(field.data_references.len()))
                .and_then(|amount| amount.checked_add(field.known_field_version.len()))
                .ok_or(ChartArrangementError::InvalidSource)?;
            if let Some(identifier) = field.known_field_feature_identifier.as_ref() {
                work = work
                    .checked_add(identifier.len())
                    .ok_or(ChartArrangementError::InvalidSource)?;
            }
        }
    }
    Ok(work)
}

/// Reserve the complete metadata and header footprint of the selected object
/// clone used by the locality proof. The replacement message bytes are
/// included because the proof stages an owned `RawMessage` before comparing
/// the candidate object.
fn archive_object_clone_requirements(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    replacement_len: usize,
) -> Result<(usize, usize, usize), ChartArrangementError> {
    let mut allocations = 0usize;
    let mut retained = size_of::<ArchiveObject>();
    add_clone_vec::<RawMessage>(source.messages.len(), &mut allocations, &mut retained)?;
    add_clone_vec::<litchi_iwa_core::MessageInfo>(
        source.archive_info.message_infos.len(),
        &mut allocations,
        &mut retained,
    )?;
    for message in &source.messages {
        add_clone_bytes(message.data.len(), &mut allocations, &mut retained)?;
    }
    for info in &source.archive_info.message_infos {
        add_clone_vec::<u32>(info.versions.len(), &mut allocations, &mut retained)?;
        add_clone_vec::<litchi_iwa_core::FieldInfo>(
            info.field_infos.len(),
            &mut allocations,
            &mut retained,
        )?;
        add_clone_vec::<u64>(
            info.object_references.len(),
            &mut allocations,
            &mut retained,
        )?;
        add_clone_vec::<u64>(info.data_references.len(), &mut allocations, &mut retained)?;
        add_clone_vec::<u32>(
            info.diff_merge_version.len(),
            &mut allocations,
            &mut retained,
        )?;
        if let Some(path) = info.diff_field_path.as_ref() {
            add_clone_vec::<u32>(path.path.len(), &mut allocations, &mut retained)?;
        }
        add_clone_vec::<litchi_iwa_core::FieldPath>(
            info.fields_to_remove.len(),
            &mut allocations,
            &mut retained,
        )?;
        for path in &info.fields_to_remove {
            add_clone_vec::<u32>(path.path.len(), &mut allocations, &mut retained)?;
        }
        add_clone_vec::<u32>(
            info.diff_read_version.len(),
            &mut allocations,
            &mut retained,
        )?;
        for field in &info.field_infos {
            add_clone_vec::<u32>(field.path.path.len(), &mut allocations, &mut retained)?;
            add_clone_vec::<u64>(
                field.object_references.len(),
                &mut allocations,
                &mut retained,
            )?;
            add_clone_vec::<u64>(field.data_references.len(), &mut allocations, &mut retained)?;
            add_clone_vec::<u32>(
                field.known_field_version.len(),
                &mut allocations,
                &mut retained,
            )?;
            if let Some(identifier) = field.known_field_feature_identifier.as_ref() {
                add_clone_bytes(identifier.len(), &mut allocations, &mut retained)?;
            }
        }
    }
    if source.header_length != 0 {
        allocations = allocations
            .checked_add(2)
            .ok_or(ChartArrangementError::InvalidSource)?;
        retained = retained
            .checked_add(
                usize::try_from(source.header_length)
                    .map_err(|_| ChartArrangementError::InvalidSource)?
                    .checked_mul(2)
                    .ok_or(ChartArrangementError::InvalidSource)?,
            )
            .ok_or(ChartArrangementError::InvalidSource)?;
    }
    add_clone_bytes(replacement_len, &mut allocations, &mut retained)?;
    let source_header =
        usize::try_from(source.header_length).map_err(|_| ChartArrangementError::InvalidSource)?;
    let candidate_header = usize::try_from(candidate.header_length)
        .map_err(|_| ChartArrangementError::InvalidSource)?;
    let header_scratch = source_header
        .checked_add(candidate_header)
        .and_then(|bytes| bytes.checked_mul(3))
        .ok_or(ChartArrangementError::InvalidSource)?;
    allocations = allocations
        .checked_add(3)
        .ok_or(ChartArrangementError::InvalidSource)?;
    retained = retained
        .checked_add(header_scratch)
        .ok_or(ChartArrangementError::InvalidSource)?;
    let work = archive_object_work(source)?
        .checked_add(replacement_len)
        .and_then(|amount| amount.checked_add(header_scratch))
        .ok_or(ChartArrangementError::InvalidSource)?;
    Ok((allocations, retained, work))
}

fn add_clone_vec<T>(
    length: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), ChartArrangementError> {
    if length == 0 {
        return Ok(());
    }
    *allocations = allocations
        .checked_add(1)
        .ok_or(ChartArrangementError::InvalidSource)?;
    *retained = retained
        .checked_add(
            length
                .checked_mul(size_of::<T>())
                .ok_or(ChartArrangementError::InvalidSource)?,
        )
        .ok_or(ChartArrangementError::InvalidSource)?;
    Ok(())
}

fn add_clone_bytes(
    length: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), ChartArrangementError> {
    add_clone_vec::<u8>(length, allocations, retained)
}

fn map_common_error(error: CommonError) -> ChartArrangementError {
    match error {
        CommonError::LimitExceeded {
            kind,
            observed,
            limit,
        } => ChartArrangementError::LimitExceeded {
            kind: match kind {
                CommonLimitKind::InputBytes => ChartArrangementLimitKind::WireBytes,
                CommonLimitKind::Fields => ChartArrangementLimitKind::WireFields,
                CommonLimitKind::OutputBytes => ChartArrangementLimitKind::OutputBytes,
                CommonLimitKind::Nesting => ChartArrangementLimitKind::WireNesting,
                CommonLimitKind::RewriteWork => ChartArrangementLimitKind::WireWork,
                _ => ChartArrangementLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        CommonError::Allocation { amount, .. } => ChartArrangementError::Allocation { amount },
        _ => ChartArrangementError::InvalidSource,
    }
}

fn map_package_error(error: PackageError) -> ChartArrangementError {
    match error {
        PackageError::Common(error) => map_common_error(error),
        _ => ChartArrangementError::InvalidSource,
    }
}

fn map_codec_error(error: codec::DecodeError) -> ChartArrangementError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            DecodeLimit::Bytes { observed, maximum } => {
                (ChartArrangementLimitKind::WireBytes, observed, maximum)
            },
            DecodeLimit::Fields { observed, maximum } => {
                (ChartArrangementLimitKind::WireFields, observed, maximum)
            },
            DecodeLimit::Work { observed, maximum } => {
                (ChartArrangementLimitKind::WireWork, observed, maximum)
            },
            DecodeLimit::Output { observed, maximum } => {
                (ChartArrangementLimitKind::OutputBytes, observed, maximum)
            },
            DecodeLimit::Nesting { observed, maximum } => (
                ChartArrangementLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            DecodeLimit::Allocations { observed, maximum } => {
                (ChartArrangementLimitKind::Allocations, observed, maximum)
            },
            DecodeLimit::Retained { observed, maximum } => {
                (ChartArrangementLimitKind::RetainedBytes, observed, maximum)
            },
            DecodeLimit::Scratch { observed, maximum } => {
                (ChartArrangementLimitKind::ScratchBytes, observed, maximum)
            },
            _ => return ChartArrangementError::InvalidSource,
        };
        return ChartArrangementError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return ChartArrangementError::Allocation { amount };
    }
    ChartArrangementError::InvalidSource
}

fn physical_component_archive(
    package: &Package,
    component_index: usize,
) -> Result<&Archive, ChartArrangementError> {
    Ok(package_component(package, component_index)?.archive())
}

// Keep the archive helper visible to rustdoc's private-item lint without
// exposing native identifiers through the public package API.
#[allow(
    dead_code,
    reason = "Retained for the host migration's graph diagnostics."
)]
fn _component_archive(
    package: &Package,
    component_index: usize,
) -> Result<&Archive, ChartArrangementError> {
    physical_component_archive(package, component_index)
}
