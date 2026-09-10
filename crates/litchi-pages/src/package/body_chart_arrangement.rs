//! Selector-first, source-preserving Pages body-chart Arrange transactions.
//!
//! The public value is the small archive-free [`ChartArrangement`] state used
//! by the Arrange panel.  Selection remains semantic: this adapter discovers
//! ordinary chart attachments from the rooted body storage, proves their
//! attachment/drawable/body/z-order ownership, and only then hands the
//! selected drawable payload to the bounded lazy Buffa codec.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "The Pages package boundary redacts native graph failures."
)]

use std::fmt;
use std::mem::size_of;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts, SharedBytes};
use litchi_iwa_archive::{Error as ArchiveError, LimitKind as ArchiveLimitKind, SourceCatalog};
use litchi_iwa_common::chart::arrangement::ChartArrangement;
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    varint::encoded_len,
    wire::{WireDescent, WireFieldView, WireView, preflight_wire_tree_with_limits},
};
use litchi_iwa_core::{ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    chart_arrangement_codec as codec, pages_body_codec, pages_drawable_order_codec,
};
use thiserror::Error;

use super::{Package, page_layout};
use crate::selector::BodyChartSelector;

const DOCUMENT_COMPONENT: &str = "Index/Document.iwa";
const ROOT_OBJECT_IDENTIFIER: u64 = 1;
const ROOT_MESSAGE_TYPE: u32 = 10_000;
const BODY_TEXT_FIELD: u32 = 3;
const BODY_ATTACHMENTS_FIELD: u32 = 9;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const ATTACHMENT_MESSAGE_TYPE: u32 = 2_003;
const ATTACHMENT_DRAWABLE_FIELD: u32 = 1;
const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_DRAWABLE_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const DRAWABLE_ORDER_MESSAGE_TYPE: u32 = 10_015;
const ROOT_PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const MAX_CODEC_LIMIT: usize = 64 * 1024 * 1024;
// `WireView` retains one compact private span per field.  Its span type is
// intentionally private to the common crate, so reserve a conservative
// pointer-sized bound before parsing rather than admitting an uncharged Vec
// growth at this format boundary.
const WIRE_VIEW_SPAN_BYTES_PER_FIELD: usize = size_of::<[usize; 8]>();

/// Finite resources governed by one Pages body-chart Arrange operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyChartArrangementLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// ZIP entries inspected or retained.
    Entries,
    /// One ZIP entry's bytes.
    EntryBytes,
    /// Aggregate ZIP entry bytes.
    TotalEntryBytes,
    /// ZIP metadata bytes.
    PackageBytes,
    /// One native payload's bytes.
    PayloadBytes,
    /// Aggregate native payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native metadata items inspected.
    PayloadItems,
    /// Native references inspected.
    PayloadReferences,
    /// Strict wire input bytes.
    WireBytes,
    /// Codec output bytes.
    WireOutputBytes,
    /// Strict wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire work.
    WireWork,
    /// Codec allocations.
    WireAllocations,
    /// Codec retained bytes.
    WireRetainedBytes,
    /// Codec scratch bytes.
    WireScratchBytes,
}

impl fmt::Display for BodyChartArrangementLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "ZIP entries",
            Self::EntryBytes => "ZIP entry bytes",
            Self::TotalEntryBytes => "total ZIP entry bytes",
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::WireAllocations => "wire allocations",
            Self::WireRetainedBytes => "wire retained bytes",
            Self::WireScratchBytes => "wire scratch bytes",
        })
    }
}

/// Failure from a Pages body-chart Arrange read or exact-source transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum BodyChartArrangementError {
    /// No ordinary body chart matched the checked source-order selector.
    #[error("the Pages body has no chart at position {position:?}")]
    ChartNotFound { position: Position },
    /// Alias used by callers that name checked positions explicitly.
    #[error("the Pages body has no chart at position {position:?}")]
    ChartPositionNotFound { position: Position },
    /// The source does not retain an exact physical artifact suitable for a
    /// changed publication.
    #[error("this Pages source does not support exact body-chart arrangement edits")]
    UnsupportedSource,
    /// The rooted chart graph or selected payload was malformed or ambiguous.
    #[error("the selected Pages body-chart arrangement source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Pages body-chart arrangement {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: BodyChartArrangementLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for Pages body-chart arrangement")]
    Allocation { amount: usize },
    /// Reopening the candidate did not reproduce the requested semantic state.
    #[error("the edited Pages body-chart arrangement failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body-chart arrangement patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable semantic chart Arrange value staged against an immutable
/// package snapshot.
pub struct BodyChartArrangementEdit<'a> {
    source: &'a Package,
    target: ChartTarget,
    before: ChartArrangement,
    after: ChartArrangement,
}

impl fmt::Debug for BodyChartArrangementEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyChartArrangementEdit")
            .field("position", &self.target.position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyChartArrangementEdit<'_> {
    /// Return the selected body-chart source position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.target.position
    }

    /// Return the Arrange state read before this edit.
    #[must_use]
    pub const fn before(&self) -> ChartArrangement {
        self.before
    }

    /// Return the Arrange state currently staged for publication.
    #[must_use]
    pub const fn after(&self) -> ChartArrangement {
        self.after
    }

    /// Replace the staged Arrange state.
    #[must_use]
    pub const fn set(mut self, arrangement: ChartArrangement) -> Self {
        self.after = arrangement;
        self
    }

    /// Validate and publish the staged exact-source edit.
    pub fn commit(self) -> Result<BodyChartArrangementCommit, BodyChartArrangementError> {
        commit_edit(self)
    }
}

/// Exact-source checked reversible body-chart Arrange patch.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyChartArrangementPatch {
    artifacts: OwnedExactArtifacts,
    target: ChartTarget,
    before: ChartArrangement,
    after: ChartArrangement,
    touched_components: usize,
    source_payload: Option<Arc<[u8]>>,
    target_payload: Option<Arc<[u8]>>,
}

impl fmt::Debug for BodyChartArrangementPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyChartArrangementPatch")
            .field("position", &self.target.position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyChartArrangementPatch {
    /// Return the selected body-chart source position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.target.position
    }

    /// Return the semantic state required before this patch applies.
    #[must_use]
    pub const fn before(&self) -> ChartArrangement {
        self.before
    }

    /// Return the semantic state produced by this patch.
    #[must_use]
    pub const fn after(&self) -> ChartArrangement {
        self.after
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

    /// Return whether both semantic state and exact bytes are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            target: self.target.clone(),
            before: self.after,
            after: self.before,
            touched_components: self.touched_components,
            source_payload: self.target_payload.clone(),
            target_payload: self.source_payload.clone(),
        }
    }
}

/// Compact evidence describing one committed body-chart transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyChartArrangementDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl BodyChartArrangementDiagnostics {
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

    /// Return the number of preview members removed by this transaction.
    ///
    /// Pages chart arrangement rewrites preserve preview members, so this is
    /// always zero for a successfully published edit.
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

/// Fully reopened immutable result of one body-chart Arrange transaction.
#[must_use = "a Pages body-chart arrangement commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyChartArrangementCommit {
    package: Package,
    patch: BodyChartArrangementPatch,
    diagnostics: BodyChartArrangementDiagnostics,
}

impl BodyChartArrangementCommit {
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
    pub const fn patch(&self) -> &BodyChartArrangementPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyChartArrangementDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq, Eq)]
pub(super) struct ChartTarget {
    pub(super) position: Position,
    pub(super) character_index: u32,
    pub(super) body_identifier: NonZeroU64,
    pub(super) body_component_index: usize,
    pub(super) body_object_index: usize,
    pub(super) body_message_index: usize,
    pub(super) body_message_type: u32,
    pub(super) attachment_identifier: NonZeroU64,
    pub(super) attachment_component_index: usize,
    pub(super) attachment_object_index: usize,
    pub(super) attachment_message_index: usize,
    pub(super) drawable_identifier: NonZeroU64,
    pub(super) component_index: usize,
    pub(super) component_name: Arc<str>,
    pub(super) drawable_object_index: usize,
    pub(super) drawable_message_index: usize,
    pub(super) drawable_message_type: u32,
    pub(super) drawable_order_identifier: NonZeroU64,
    pub(super) before: ChartArrangement,
}

/// Controls whether a chart payload publication invalidates the canonical
/// root previews generated by iWork.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PreviewDeletionMode {
    /// Preserve all package previews, as Arrange-only edits do.
    Preserve,
    /// Delete the fixed canonical root preview members after a numeric edit.
    DeleteCanonicalRoot,
}

/// Selects the schema-local mutation allowed by the payload locality check.
///
/// Arrange edits update the drawable's small layout envelope. Numeric data
/// edits update only the modern chart extension's grid values. Keeping the
/// distinction explicit prevents the shared publication helper from
/// weakening either locality proof.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum ChartPayloadMutation {
    /// The two Arrange scalar fields in the drawable envelope may change.
    Arrangement,
    /// Numeric field 1 in each modern chart-grid value may change or appear.
    NumericData,
}

/// Physical chart publication plus the preview counts needed by locality and
/// reversible patch verification.
pub(super) struct ChartPayloadPublication {
    pub(super) package: Package,
    pub(super) source_previews: usize,
    pub(super) target_previews: usize,
}

impl ChartTarget {
    /// Compare the physical graph witness while deliberately ignoring the
    /// selected chart's mutable Arrange projection.  A reopened candidate is
    /// expected to differ in `before`; comparing the whole derived target
    /// would reject every successful changed commit as a false verification
    /// failure.
    pub(super) fn same_identity(&self, other: &Self) -> bool {
        self.position == other.position
            && self.character_index == other.character_index
            && self.body_identifier == other.body_identifier
            && self.body_component_index == other.body_component_index
            && self.body_object_index == other.body_object_index
            && self.body_message_index == other.body_message_index
            && self.body_message_type == other.body_message_type
            && self.attachment_identifier == other.attachment_identifier
            && self.attachment_component_index == other.attachment_component_index
            && self.attachment_object_index == other.attachment_object_index
            && self.attachment_message_index == other.attachment_message_index
            && self.drawable_identifier == other.drawable_identifier
            && self.component_index == other.component_index
            && self.component_name == other.component_name
            && self.drawable_object_index == other.drawable_object_index
            && self.drawable_message_index == other.drawable_message_index
            && self.drawable_message_type == other.drawable_message_type
            && self.drawable_order_identifier == other.drawable_order_identifier
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct LocatedObject<'a> {
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) object: &'a ArchiveObject,
}

#[derive(Debug, Clone, Copy)]
struct BodyEntry {
    character_index: u32,
    attachment_identifier: NonZeroU64,
}

/// Bounds focused graph projections, codec work, and physical rewrite staging.
/// Package reconstruction retains its separate physical and semantic ingress limits.
#[derive(Debug, Clone, Copy)]
pub(super) struct ArrangementBudget {
    limits: WireLimits,
    max_input: usize,
    max_output: usize,
    max_fields: usize,
    max_work: usize,
    max_nesting: usize,
    max_references: usize,
    max_allocations: usize,
    max_retained: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    references: usize,
    allocations: usize,
    retained: usize,
}

impl ArrangementBudget {
    pub(super) fn new(package: &Package) -> Result<Self, BodyChartArrangementError> {
        let physical = package.state.source.limits();
        let archive = physical
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let message = archive.max_message_bytes().max(1);
        let limits = WireLimits::default()
            .with_input_bytes(message.min(WireLimits::MAX_INPUT_BYTES))
            .and_then(|value| value.with_output_bytes(message.min(WireLimits::MAX_OUTPUT_BYTES)))
            .and_then(|value| {
                value.with_rewrite_work(
                    message
                        .saturating_mul(16)
                        .clamp(1, WireLimits::MAX_REWRITE_WORK),
                )
            })
            .map_err(map_wire_error)?;
        let max_output = usize::try_from(physical.max_input_bytes())
            .map_err(|_| BodyChartArrangementError::InvalidSource)?
            .max(1);
        let max_work = usize::try_from(physical.max_total_bytes())
            .map_err(|_| BodyChartArrangementError::InvalidSource)?
            .saturating_mul(32)
            .max(limits.max_rewrite_work());
        Ok(Self {
            limits,
            max_input: physical.max_iwa_stream_bytes().saturating_mul(8).max(1),
            max_output,
            max_fields: limits.max_fields(),
            max_work,
            max_nesting: limits.max_nesting(),
            max_references: archive.max_metadata_items().saturating_mul(16).max(1),
            max_allocations: archive.max_metadata_items().saturating_mul(16).max(1),
            max_retained: max_output,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            references: 0,
            allocations: 0,
            retained: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: BodyChartArrangementLimitKind,
    ) -> Result<(), BodyChartArrangementError> {
        let observed = current
            .checked_add(amount)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        if observed > maximum {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    pub(super) fn input(&mut self, amount: usize) -> Result<(), BodyChartArrangementError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            BodyChartArrangementLimitKind::InputBytes,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), BodyChartArrangementError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            BodyChartArrangementLimitKind::OutputBytes,
        )
    }

    pub(super) fn fields(&mut self, amount: usize) -> Result<(), BodyChartArrangementError> {
        Self::add(
            &mut self.fields,
            amount,
            self.max_fields,
            BodyChartArrangementLimitKind::WireFields,
        )
    }

    pub(super) fn work(&mut self, amount: usize) -> Result<(), BodyChartArrangementError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            BodyChartArrangementLimitKind::WireWork,
        )
    }

    pub(super) fn references(&mut self, amount: usize) -> Result<(), BodyChartArrangementError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            BodyChartArrangementLimitKind::PayloadReferences,
        )
    }

    pub(super) fn allocations(&mut self, amount: usize) -> Result<(), BodyChartArrangementError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            BodyChartArrangementLimitKind::WireAllocations,
        )
    }

    pub(super) fn retained(&mut self, amount: usize) -> Result<(), BodyChartArrangementError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            BodyChartArrangementLimitKind::WireRetainedBytes,
        )
    }

    pub(super) fn parse<'a>(
        &mut self,
        source: &'a [u8],
        depth: usize,
    ) -> Result<WireView<'a>, BodyChartArrangementError> {
        if depth > self.max_nesting {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        let preflight =
            preflight_wire_tree_with_limits(source, self.residual_wire_limits()?, |_| {
                Ok(WireDescent::Skip)
            })
            .map_err(map_wire_error)?;
        self.input(preflight.scanned_bytes())?;
        self.work(preflight.scanned_bytes())?;
        self.fields(preflight.fields())?;
        let span_bytes = preflight
            .fields()
            .checked_mul(4)
            .and_then(|capacity| capacity.checked_mul(WIRE_VIEW_SPAN_BYTES_PER_FIELD))
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        // Include incremental reserve events and cumulative span capacities.
        let span_allocations = preflight.fields();
        self.preflight_allocations(span_allocations)?;
        self.preflight_retained(span_bytes)?;
        let view = WireView::parse_with_limits(source, self.residual_wire_limits()?)
            .map_err(map_wire_error)?;
        self.allocations(span_allocations)?;
        self.retained(span_bytes)?;
        self.input(source.len())?;
        self.fields(view.len())?;
        self.work(source.len().saturating_add(view.len()))?;
        Ok(view)
    }

    pub(super) fn residual_wire_limits(&self) -> Result<WireLimits, BodyChartArrangementError> {
        let input = self
            .max_input
            .checked_sub(self.input)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        let fields = self
            .max_fields
            .checked_sub(self.fields)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        let work = self
            .max_work
            .checked_sub(self.work)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        if input == 0 || fields == 0 || work == 0 {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::WireWork,
                observed: 1,
                maximum: 0,
            });
        }
        self.limits
            .with_input_bytes(self.limits.max_input_bytes().min(input))
            .and_then(|value| value.with_fields(self.limits.max_fields().min(fields)))
            .and_then(|value| value.with_rewrite_work(self.limits.max_rewrite_work().min(work)))
            .and_then(|value| value.with_nesting(self.limits.max_nesting().min(self.max_nesting)))
            .map_err(map_wire_error)
    }

    pub(super) fn residual_output_bytes(&self) -> Result<usize, BodyChartArrangementError> {
        self.max_output
            .checked_sub(self.output)
            .ok_or(BodyChartArrangementError::InvalidSource)
    }

    pub(super) fn residual_allocations(&self) -> Result<usize, BodyChartArrangementError> {
        self.max_allocations
            .checked_sub(self.allocations)
            .ok_or(BodyChartArrangementError::InvalidSource)
    }

    pub(super) fn residual_retained_bytes(&self) -> Result<usize, BodyChartArrangementError> {
        self.max_retained
            .checked_sub(self.retained)
            .ok_or(BodyChartArrangementError::InvalidSource)
    }

    fn codec_options(
        &self,
        source: &[u8],
    ) -> Result<codec::DecodeOptions, BodyChartArrangementError> {
        let limits = self.residual_wire_limits()?;
        let output = self
            .max_output
            .checked_sub(self.output)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        let allocations = self
            .max_allocations
            .checked_sub(self.allocations)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        let retained = self
            .max_retained
            .checked_sub(self.retained)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        let recursion = u32::try_from(limits.max_nesting())
            .map_err(|_| BodyChartArrangementError::InvalidSource)?;
        Ok(codec::DecodeOptions::new(
            limits
                .max_input_bytes()
                .min(source.len().max(1))
                .min(MAX_CODEC_LIMIT),
            limits.max_fields().min(MAX_CODEC_LIMIT),
            limits.max_rewrite_work().min(MAX_CODEC_LIMIT),
            recursion,
        )
        .with_max_output_bytes(output.clamp(1, MAX_CODEC_LIMIT))
        .with_max_allocations(allocations.clamp(1, MAX_CODEC_LIMIT))
        .with_max_retained_bytes(retained.clamp(1, MAX_CODEC_LIMIT))
        .with_max_scratch_bytes(retained.clamp(1, MAX_CODEC_LIMIT)))
    }

    fn charge_arrangement_report(
        &mut self,
        report: codec::DecodeReport,
    ) -> Result<(), BodyChartArrangementError> {
        self.input(report.source_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.retained(report.scratch_bytes())?;
        if usize::try_from(report.max_depth()).unwrap_or(usize::MAX) > self.max_nesting {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::WireNesting,
                observed: u64::from(report.max_depth()),
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn charge_rewrite(
        &mut self,
        requirements: codec::RewriteExecutionRequirements,
    ) -> Result<(), BodyChartArrangementError> {
        self.charge_rewrite_values(
            requirements.output_bytes,
            requirements.fields,
            requirements.work_bytes,
            requirements.max_depth,
            requirements.allocations,
            requirements.retained_bytes,
            requirements.scratch_bytes,
        )
    }

    pub(super) fn charge_rewrite_values(
        &mut self,
        output_bytes: usize,
        fields: usize,
        work_bytes: usize,
        max_depth: u32,
        allocations: usize,
        retained_bytes: usize,
        scratch_bytes: usize,
    ) -> Result<(), BodyChartArrangementError> {
        self.output(output_bytes)?;
        self.fields(fields)?;
        self.work(work_bytes)?;
        self.allocations(allocations)?;
        self.retained(retained_bytes)?;
        self.retained(scratch_bytes)?;
        if usize::try_from(max_depth).unwrap_or(usize::MAX) > self.max_nesting {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::WireNesting,
                observed: u64::from(max_depth),
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn order_options(
        &self,
        source: &[u8],
    ) -> Result<pages_drawable_order_codec::DecodeOptions, BodyChartArrangementError> {
        let limits = self.residual_wire_limits()?;
        let source_len = source.len().max(1);
        Ok(pages_drawable_order_codec::DecodeOptions::new(
            source_len.min(limits.max_input_bytes()),
            source_len
                .saturating_mul(2)
                .min(limits.max_output_bytes())
                .max(1),
            source_len.saturating_mul(8).min(limits.max_fields()).max(1),
            limits.max_rewrite_work().clamp(1, MAX_CODEC_LIMIT),
            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
            self.max_references.max(1),
        ))
    }

    fn charge_order_report(
        &mut self,
        report: pages_drawable_order_codec::DecodeReport,
    ) -> Result<(), BodyChartArrangementError> {
        self.input(report.input_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.references(report.references())?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        if usize::try_from(report.max_depth()).unwrap_or(usize::MAX) > self.max_nesting {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::WireNesting,
                observed: u64::from(report.max_depth()),
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn preflight_output(&self, amount: usize) -> Result<(), BodyChartArrangementError> {
        let observed = self
            .output
            .checked_add(amount)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        if observed > self.max_output {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::OutputBytes,
                observed: observed as u64,
                maximum: self.max_output as u64,
            });
        }
        Ok(())
    }

    fn preflight_work(&self, amount: usize) -> Result<(), BodyChartArrangementError> {
        let observed = self
            .work
            .checked_add(amount)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        if observed > self.max_work {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::WireWork,
                observed: observed as u64,
                maximum: self.max_work as u64,
            });
        }
        Ok(())
    }

    pub(super) fn preflight_allocations(
        &self,
        amount: usize,
    ) -> Result<(), BodyChartArrangementError> {
        let observed = self
            .allocations
            .checked_add(amount)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        if observed > self.max_allocations {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::WireAllocations,
                observed: observed as u64,
                maximum: self.max_allocations as u64,
            });
        }
        Ok(())
    }

    pub(super) fn preflight_retained(
        &self,
        amount: usize,
    ) -> Result<(), BodyChartArrangementError> {
        let observed = self
            .retained
            .checked_add(amount)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        if observed > self.max_retained {
            return Err(BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::WireRetainedBytes,
                observed: observed as u64,
                maximum: self.max_retained as u64,
            });
        }
        Ok(())
    }

    // This ledger admits the transaction's artifact copy. SourceCatalog and
    // semantic Package reconstruction separately enforce their ingress limits.
    pub(super) fn candidate_reopen(
        &mut self,
        bytes: usize,
    ) -> Result<(), BodyChartArrangementError> {
        self.input(bytes)?;
        self.work(bytes)?;
        self.preflight_allocations(1)?;
        self.preflight_retained(bytes)?;
        self.allocations(1)?;
        self.retained(bytes)
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), BodyChartArrangementError> {
        self.output(requirements.output_bytes())?;
        self.work(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.retained(requirements.scratch_bytes())
    }
}

impl Package {
    /// Read one rooted body chart's Arrange-panel state by source-order index.
    pub fn body_chart_arrangement(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<ChartArrangement, BodyChartArrangementError> {
        let mut budget = ArrangementBudget::new(self)?;
        let target = resolve_target(self, selector.into(), &mut budget)?;
        Ok(target.before)
    }

    /// Read every rooted body chart's Arrange-panel state in one checked body
    /// traversal. The order is the ordinary chart-attachment order in the
    /// body storage, and contains no native identifiers.
    pub fn body_chart_arrangements(
        &self,
    ) -> Result<Box<[ChartArrangement]>, BodyChartArrangementError> {
        let mut budget = ArrangementBudget::new(self)?;
        let targets = resolve_targets(self, &mut budget)?;
        budget.allocations(usize::from(!targets.is_empty()))?;
        budget.retained(
            targets
                .len()
                .checked_mul(size_of::<ChartArrangement>())
                .ok_or(BodyChartArrangementError::InvalidSource)?,
        )?;
        let mut arrangements = Vec::new();
        arrangements.try_reserve_exact(targets.len()).map_err(|_| {
            BodyChartArrangementError::Allocation {
                amount: targets.len(),
            }
        })?;
        for target in targets {
            arrangements.push(target.before);
        }
        Ok(arrangements.into_boxed_slice())
    }

    /// Begin a selector-first immutable body-chart Arrange edit.
    pub fn edit_body_chart_arrangement(
        &self,
        selector: impl Into<BodyChartSelector>,
    ) -> Result<BodyChartArrangementEdit<'_>, BodyChartArrangementError> {
        let mut budget = ArrangementBudget::new(self)?;
        let target = resolve_target(self, selector.into(), &mut budget)?;
        let before = target.before;
        Ok(BodyChartArrangementEdit {
            source: self,
            target,
            before,
            after: before,
        })
    }

    /// Apply an exact-source checked reversible body-chart Arrange patch.
    pub fn apply_body_chart_arrangement(
        &self,
        patch: &BodyChartArrangementPatch,
    ) -> Result<BodyChartArrangementCommit, BodyChartArrangementError> {
        let source = self.state.source.shared_source();
        let source_owner = SharedBytes::from_shared_slice(Arc::clone(&source));
        if !patch.artifacts.authorizes_owner(&source_owner) {
            return Err(BodyChartArrangementError::PatchConflict);
        }
        let mut budget = ArrangementBudget::new(self)?;
        let current = resolve_target(
            self,
            BodyChartSelector::position(patch.target.position),
            &mut budget,
        )?;
        if !current.same_identity(&patch.target) || current.before != patch.before {
            return Err(BodyChartArrangementError::PatchConflict);
        }
        if let Some(expected) = patch.source_payload.as_deref() {
            let actual = selected_chart_payload(self, &patch.target)
                .map_err(|_| BodyChartArrangementError::PatchConflict)?;
            if actual != expected {
                return Err(BodyChartArrangementError::PatchConflict);
            }
        }
        if patch.is_noop() {
            return Ok(BodyChartArrangementCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyChartArrangementDiagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact() {
            return Err(BodyChartArrangementError::PatchConflict);
        }
        let target_owner = patch.artifacts.target_owner();
        let target_len = target_owner.as_ref().len();
        // `Arc::<[u8]>::from(&[u8])` copies the complete exact target.  Admit
        // that copy before constructing it so an apply cannot allocate after
        // the transaction has crossed its last checked boundary.
        budget.candidate_reopen(target_len)?;
        let target_bytes = Arc::<[u8]>::from(target_owner.as_ref());
        let catalog =
            SourceCatalog::from_shared_bytes_with_limits(target_bytes, self.state.source.limits())
                .map_err(map_archive_error)?;
        let candidate = Package::from_source_catalog(catalog)
            .map_err(|_| BodyChartArrangementError::Verification)?;
        let verified = resolve_target(
            &candidate,
            BodyChartSelector::position(patch.target.position),
            &mut budget,
        )?;
        if !verified.same_identity(&patch.target) || verified.before != patch.after {
            return Err(BodyChartArrangementError::Verification);
        }
        verify_locality(
            self,
            &candidate,
            &patch.target,
            patch.target_payload.as_deref(),
            None,
            ChartPayloadMutation::Arrangement,
            &mut budget,
        )?;
        Ok(BodyChartArrangementCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyChartArrangementDiagnostics::published(),
        })
    }
}

fn commit_edit(
    edit: BodyChartArrangementEdit<'_>,
) -> Result<BodyChartArrangementCommit, BodyChartArrangementError> {
    let source = edit.source;
    let source_bytes = source.state.source.shared_source();
    if edit.before == edit.after {
        let source_owner = SharedBytes::from_shared_slice(Arc::clone(&source_bytes));
        return Ok(BodyChartArrangementCommit {
            package: source.snapshot(),
            patch: BodyChartArrangementPatch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                target: edit.target,
                before: edit.before,
                after: edit.after,
                touched_components: 0,
                source_payload: None,
                target_payload: None,
            },
            diagnostics: BodyChartArrangementDiagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(BodyChartArrangementError::UnsupportedSource);
    }
    let mut budget = ArrangementBudget::new(source)?;
    let current = resolve_target(
        source,
        BodyChartSelector::position(edit.target.position),
        &mut budget,
    )?;
    if current != edit.target || current.before != edit.before {
        return Err(BodyChartArrangementError::PatchConflict);
    }
    let candidate = rewrite_chart(source, &edit.target, edit.after, &mut budget)?;
    let verified = resolve_target(
        &candidate,
        BodyChartSelector::position(edit.target.position),
        &mut budget,
    )?;
    if !verified.same_identity(&edit.target) || verified.before != edit.after {
        return Err(BodyChartArrangementError::Verification);
    }
    let source_payload = selected_chart_payload(source, &edit.target)?;
    let target_payload = selected_chart_payload(&candidate, &edit.target)?;
    verify_locality(
        source,
        &candidate,
        &edit.target,
        Some(target_payload),
        None,
        ChartPayloadMutation::Arrangement,
        &mut budget,
    )?;
    let (source_payload, target_payload) =
        retain_payload_pair(source_payload, target_payload, &mut budget)?;
    let target_bytes = candidate.state.source.shared_source();
    let source_owner = SharedBytes::from_shared_slice(source_bytes);
    let target_owner = SharedBytes::from_shared_slice(Arc::clone(&target_bytes));
    Ok(BodyChartArrangementCommit {
        package: candidate,
        patch: BodyChartArrangementPatch {
            artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            target: edit.target,
            before: edit.before,
            after: edit.after,
            touched_components: 1,
            source_payload: Some(source_payload),
            target_payload: Some(target_payload),
        },
        diagnostics: BodyChartArrangementDiagnostics::published(),
    })
}

pub(super) fn resolve_target(
    package: &Package,
    selector: BodyChartSelector,
    budget: &mut ArrangementBudget,
) -> Result<ChartTarget, BodyChartArrangementError> {
    let requested = selector.as_position();
    resolve_targets(package, budget)?
        .into_iter()
        .find(|target| target.position == requested)
        .ok_or(BodyChartArrangementError::ChartNotFound {
            position: requested,
        })
}

pub(super) fn resolve_targets(
    package: &Package,
    budget: &mut ArrangementBudget,
) -> Result<Vec<ChartTarget>, BodyChartArrangementError> {
    let (
        body,
        body_identifier,
        body_payload,
        body_message_index,
        body_message_type,
        order_identifier,
    ) = body_storage_payload(package, budget)?;
    let body_view = budget.parse(body_payload, 1)?;
    let Some(table) = unique_field(&body_view, BODY_ATTACHMENTS_FIELD, 2)? else {
        return Ok(Vec::new());
    };
    let table_view = budget.parse(table.payload(), 2)?;
    let mut entries = Vec::new();
    budget.allocations(usize::from(!table_view.is_empty()))?;
    budget.retained(
        table_view
            .len()
            .checked_mul(size_of::<BodyEntry>())
            .ok_or(BodyChartArrangementError::InvalidSource)?,
    )?;
    entries.try_reserve_exact(table_view.len()).map_err(|_| {
        BodyChartArrangementError::Allocation {
            amount: table_view.len(),
        }
    })?;
    for field in table_view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRIES_FIELD)
    {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        entries.push(parse_body_entry(field.payload(), budget)?);
    }
    let sorting_work = entries
        .len()
        .checked_mul(entries.len().checked_ilog2().unwrap_or(0) as usize + 1)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    budget.work(sorting_work)?;
    entries.sort_unstable_by_key(|entry| entry.character_index);
    if entries
        .windows(2)
        .any(|window| window[0].character_index == window[1].character_index)
    {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    validate_body_anchors(body_payload, &entries, budget)?;

    let mut seen_attachments = Vec::new();
    let mut seen_drawables = Vec::new();
    budget.allocations(usize::from(!entries.is_empty()) * 2)?;
    budget.retained(
        entries
            .len()
            .checked_mul(size_of::<NonZeroU64>() * 2)
            .ok_or(BodyChartArrangementError::InvalidSource)?,
    )?;
    seen_attachments
        .try_reserve_exact(entries.len())
        .map_err(|_| BodyChartArrangementError::Allocation {
            amount: entries.len(),
        })?;
    seen_drawables
        .try_reserve_exact(entries.len())
        .map_err(|_| BodyChartArrangementError::Allocation {
            amount: entries.len(),
        })?;

    let mut targets = Vec::new();
    budget.allocations(usize::from(!entries.is_empty()))?;
    budget.retained(
        entries
            .len()
            .checked_mul(size_of::<ChartTarget>())
            .ok_or(BodyChartArrangementError::InvalidSource)?,
    )?;
    targets.try_reserve_exact(entries.len()).map_err(|_| {
        BodyChartArrangementError::Allocation {
            amount: entries.len(),
        }
    })?;
    for entry in entries {
        budget.work(
            seen_attachments
                .len()
                .checked_add(seen_drawables.len())
                .ok_or(BodyChartArrangementError::InvalidSource)?,
        )?;
        if entry.attachment_identifier == order_identifier
            || seen_drawables.contains(&entry.attachment_identifier)
        {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        if !object_metadata_is_owned(
            body.object,
            body_message_index,
            entry.attachment_identifier,
            &[BODY_ATTACHMENTS_FIELD],
            true,
        ) {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        let attachment = locate_unique_object(package, entry.attachment_identifier, budget)?;
        if attachment.object.archive_info.identifier == Some(ROOT_OBJECT_IDENTIFIER)
            || attachment.object.archive_info.identifier == Some(body_identifier.get())
            || !push_unique(&mut seen_attachments, entry.attachment_identifier)
        {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        let Some((attachment_message_index, attachment_message)) =
            unique_optional_message(attachment.object, ATTACHMENT_MESSAGE_TYPE)?
        else {
            continue;
        };
        validate_message_metadata(attachment.object, attachment_message_index)?;
        let drawable_identifier = parse_object_reference(
            unique_field_payload(attachment_message.data.as_slice(), 1, 2, budget)?,
            budget,
        )?;
        if !object_metadata_is_owned(
            attachment.object,
            attachment_message_index,
            drawable_identifier,
            &[ATTACHMENT_DRAWABLE_FIELD],
            true,
        ) {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        budget.work(
            seen_attachments
                .len()
                .checked_add(seen_drawables.len())
                .ok_or(BodyChartArrangementError::InvalidSource)?,
        )?;
        if drawable_identifier == order_identifier
            || seen_attachments.contains(&drawable_identifier)
        {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        let drawable = locate_unique_object(package, drawable_identifier, budget)?;
        if drawable.component_index != attachment.component_index
            || drawable.object.archive_info.identifier == Some(ROOT_OBJECT_IDENTIFIER)
            || drawable.object.archive_info.identifier == Some(body_identifier.get())
            || !push_unique(&mut seen_drawables, drawable_identifier)
        {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        let Some((drawable_message_index, drawable_message)) =
            unique_optional_message(drawable.object, CHART_MESSAGE_TYPE)?
        else {
            continue;
        };
        validate_message_metadata(drawable.object, drawable_message_index)?;
        let parent = parse_chart_parent(drawable_message.data.as_slice(), budget)?;
        if parent != body_identifier
            || !object_metadata_is_owned(
                drawable.object,
                drawable_message_index,
                parent,
                &[CHART_DRAWABLE_FIELD, DRAWABLE_PARENT_FIELD],
                false,
            )
        {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        validate_drawable_order(package, order_identifier, drawable_identifier, budget)?;
        let arrangement = decode_arrangement(drawable_message.data.as_slice(), budget)?;
        let component_name = package
            .state
            .source
            .components()
            .get_index(drawable.component_index)
            .ok_or(BodyChartArrangementError::InvalidSource)?
            .name();
        budget.allocations(1)?;
        budget.retained(component_name.len())?;
        let position = Position::new(targets.len());
        targets.push(ChartTarget {
            position,
            character_index: entry.character_index,
            body_identifier,
            body_component_index: body.component_index,
            body_object_index: body.object_index,
            body_message_index,
            body_message_type,
            attachment_identifier: entry.attachment_identifier,
            attachment_component_index: attachment.component_index,
            attachment_object_index: attachment.object_index,
            attachment_message_index,
            drawable_identifier,
            component_index: drawable.component_index,
            component_name: Arc::from(component_name),
            drawable_object_index: drawable.object_index,
            drawable_message_index,
            drawable_message_type: drawable_message.type_,
            drawable_order_identifier: order_identifier,
            before: arrangement,
        });
    }
    Ok(targets)
}

fn body_storage_payload<'a>(
    package: &'a Package,
    budget: &mut ArrangementBudget,
) -> Result<
    (
        LocatedObject<'a>,
        NonZeroU64,
        &'a [u8],
        usize,
        u32,
        NonZeroU64,
    ),
    BodyChartArrangementError,
> {
    let document = package
        .state
        .source
        .components()
        .get(DOCUMENT_COMPONENT)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    let root = document
        .archive()
        .object(ROOT_OBJECT_IDENTIFIER)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    budget.work(archive_object_work(root)?)?;
    let (root_message_index, root_message) = unique_message(root, ROOT_MESSAGE_TYPE)?;
    validate_message_metadata(root, root_message_index)?;
    let root_payload = root_message.data.as_slice();
    let options = root_projection_options(root_payload, budget, true)?;
    let body_facts = pages_body_codec::decode_document_body(root_payload, options)
        .map_err(|_| BodyChartArrangementError::InvalidSource)?;
    let options = root_projection_options(root_payload, budget, false)?;
    let root_facts = pages_body_codec::decode_document_root(root_payload, options)
        .map_err(|_| BodyChartArrangementError::InvalidSource)?;
    let body_identifier = body_facts
        .body_storage()
        .map(|reference| reference.identifier())
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    let order_identifier = root_facts
        .drawables_zorder()
        .map(|reference| reference.identifier())
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    if body_identifier == order_identifier
        || body_identifier.get() == ROOT_OBJECT_IDENTIFIER
        || order_identifier.get() == ROOT_OBJECT_IDENTIFIER
    {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    if !object_metadata_is_owned(root, root_message_index, body_identifier, &[4], true)
        || !object_metadata_is_owned(root, root_message_index, order_identifier, &[20], true)
    {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    let body = locate_unique_object(package, body_identifier, budget)?;
    let (body_message_index, body_message) = unique_text_message(body.object)?;
    validate_message_metadata(body.object, body_message_index)?;
    let body_message_type = body_message.type_;
    Ok((
        body,
        body_identifier,
        body_message.data.as_slice(),
        body_message_index,
        body_message_type,
        order_identifier,
    ))
}

// Mirror the root codec's selected descent: references are decoded once,
// TSA is inspected one level deep, and its TSK payload stays opaque. The
// preflight report bounds both this scan and the codec's strict/lazy passes.
fn root_projection_options(
    source: &[u8],
    budget: &mut ArrangementBudget,
    body_only: bool,
) -> Result<pages_body_codec::DecodeOptions, BodyChartArrangementError> {
    // Selected descent is one level, requiring one small path vector.
    budget.allocations(1)?;
    budget.retained(size_of::<[u32; 4]>())?;
    let report = preflight_wire_tree_with_limits(source, budget.residual_wire_limits()?, |visit| {
        let selected = visit.path().is_empty()
            && if body_only {
                matches!(visit.field().number(), 4 | 5)
            } else {
                matches!(visit.field().number(), 3 | 6 | 15 | 20 | 48)
            };
        Ok(if selected {
            WireDescent::Descend
        } else {
            WireDescent::Skip
        })
    })
    .map_err(map_wire_error)?;
    budget.input(report.scanned_bytes())?;
    budget.fields(report.fields())?;
    budget.work(report.scanned_bytes())?;
    let options = pages_body_options(budget, source)?;
    budget.input(report.scanned_bytes())?;
    budget.fields(report.fields())?;
    budget.work(
        report
            .scanned_bytes()
            .checked_mul(2)
            .ok_or(BodyChartArrangementError::InvalidSource)?,
    )?;
    Ok(options)
}

fn pages_body_options(
    budget: &ArrangementBudget,
    source: &[u8],
) -> Result<pages_body_codec::DecodeOptions, BodyChartArrangementError> {
    let limits = budget.residual_wire_limits()?;
    Ok(pages_body_codec::DecodeOptions::new(
        limits.max_input_bytes().min(source.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting())
            .map_err(|_| BodyChartArrangementError::InvalidSource)?,
    ))
}

fn parse_body_entry(
    source: &[u8],
    budget: &mut ArrangementBudget,
) -> Result<BodyEntry, BodyChartArrangementError> {
    let view = budget.parse(source, 3)?;
    let character = unique_field(&view, ENTRY_CHARACTER_INDEX_FIELD, 0)?
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    let (character, width) = decode_varint_from_bytes(character.payload())
        .map_err(|_| BodyChartArrangementError::InvalidSource)?;
    if width != encoded_len(character) || character > u64::from(u32::MAX) {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    let object = unique_field(&view, ENTRY_OBJECT_FIELD, 2)?
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    Ok(BodyEntry {
        character_index: character as u32,
        attachment_identifier: parse_object_reference(object.payload(), budget)?,
    })
}

fn validate_body_anchors(
    payload: &[u8],
    entries: &[BodyEntry],
    budget: &mut ArrangementBudget,
) -> Result<(), BodyChartArrangementError> {
    let view = budget.parse(payload, 1)?;
    let mut next = 0usize;
    let mut utf16_index = 0usize;
    for field in view
        .fields()
        .filter(|field| field.number() == BODY_TEXT_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let text = std::str::from_utf8(field.payload())
            .map_err(|_| BodyChartArrangementError::InvalidSource)?;
        for character in text.chars() {
            if entries
                .get(next)
                .is_some_and(|entry| entry.character_index as usize == utf16_index)
            {
                if character != '\u{fffc}' {
                    return Err(BodyChartArrangementError::InvalidSource);
                }
                next = next
                    .checked_add(1)
                    .ok_or(BodyChartArrangementError::InvalidSource)?;
            }
            utf16_index = utf16_index
                .checked_add(character.len_utf16())
                .ok_or(BodyChartArrangementError::InvalidSource)?;
        }
    }
    if next != entries.len() {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    budget.work(utf16_index)
}

fn parse_object_reference(
    source: &[u8],
    budget: &mut ArrangementBudget,
) -> Result<NonZeroU64, BodyChartArrangementError> {
    let view = budget.parse(source, 4)?;
    let field = unique_field(&view, 1, 0)?.ok_or(BodyChartArrangementError::InvalidSource)?;
    let (identifier, width) = decode_varint_from_bytes(field.payload())
        .map_err(|_| BodyChartArrangementError::InvalidSource)?;
    if width != encoded_len(identifier) {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    for number in [2, 3] {
        if let Some(field) = unique_field(&view, number, 0)? {
            let (value, consumed) = decode_varint_from_bytes(field.payload())
                .map_err(|_| BodyChartArrangementError::InvalidSource)?;
            if consumed != encoded_len(value) || (number == 3 && value != 0) {
                return Err(BodyChartArrangementError::InvalidSource);
            }
        }
    }
    NonZeroU64::new(identifier).ok_or(BodyChartArrangementError::InvalidSource)
}

fn parse_chart_parent(
    source: &[u8],
    budget: &mut ArrangementBudget,
) -> Result<NonZeroU64, BodyChartArrangementError> {
    let chart = budget.parse(source, 1)?;
    let super_payload = unique_field(&chart, CHART_DRAWABLE_FIELD, 2)?
        .ok_or(BodyChartArrangementError::InvalidSource)?
        .payload();
    let drawable = budget.parse(super_payload, 2)?;
    let parent = unique_field(&drawable, DRAWABLE_PARENT_FIELD, 2)?
        .ok_or(BodyChartArrangementError::InvalidSource)?
        .payload();
    parse_object_reference(parent, budget)
}

fn decode_arrangement(
    source: &[u8],
    budget: &mut ArrangementBudget,
) -> Result<ChartArrangement, BodyChartArrangementError> {
    let options = budget.codec_options(source)?;
    let (snapshot, report) =
        codec::decode_chart_arrangement_with_report(source, options).map_err(map_codec_error)?;
    budget.charge_arrangement_report(report)?;
    if !snapshot.has_drawable() {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    Ok(ChartArrangement::new(
        snapshot.is_locked(),
        snapshot.is_constrained(),
    ))
}

pub(super) fn locate_unique_object<'a>(
    package: &'a Package,
    identifier: NonZeroU64,
    budget: &mut ArrangementBudget,
) -> Result<LocatedObject<'a>, BodyChartArrangementError> {
    budget.work(package.state.object_count)?;
    let mut found = None;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            if object.archive_info.identifier == Some(identifier.get()) {
                if found.is_some() {
                    return Err(BodyChartArrangementError::InvalidSource);
                }
                found = Some(LocatedObject {
                    component_index,
                    object_index,
                    object,
                });
            }
        }
    }
    let found = found.ok_or(BodyChartArrangementError::InvalidSource)?;
    budget.work(archive_object_work(found.object)?)?;
    Ok(found)
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &RawMessage), BodyChartArrangementError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        if info.type_ != message.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        if message.type_ == message_type && selected.replace((index, message)).is_some() {
            return Err(BodyChartArrangementError::InvalidSource);
        }
    }
    selected.ok_or(BodyChartArrangementError::InvalidSource)
}

pub(super) fn unique_optional_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<Option<(usize, &RawMessage)>, BodyChartArrangementError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        if info.type_ != message.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        if message.type_ == message_type && selected.replace((index, message)).is_some() {
            return Err(BodyChartArrangementError::InvalidSource);
        }
    }
    Ok(selected)
}

fn unique_text_message(
    object: &ArchiveObject,
) -> Result<(usize, &RawMessage), BodyChartArrangementError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if matches!(message.type_, 2_001 | 2_022) && selected.replace((index, message)).is_some() {
            return Err(BodyChartArrangementError::InvalidSource);
        }
    }
    selected.ok_or(BodyChartArrangementError::InvalidSource)
}

fn unique_field<'a>(
    view: &WireView<'a>,
    number: u32,
    wire_type: u8,
) -> Result<Option<WireFieldView<'a>>, BodyChartArrangementError> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != wire_type {
            return Err(BodyChartArrangementError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if found.replace(field).is_some() {
            return Err(BodyChartArrangementError::InvalidSource);
        }
    }
    Ok(found)
}

fn unique_field_payload<'a>(
    source: &'a [u8],
    number: u32,
    wire_type: u8,
    budget: &mut ArrangementBudget,
) -> Result<&'a [u8], BodyChartArrangementError> {
    let view = budget.parse(source, 1)?;
    unique_field(&view, number, wire_type)?
        .map(|field| field.payload())
        .ok_or(BodyChartArrangementError::InvalidSource)
}

pub(super) fn validate_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), BodyChartArrangementError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    Ok(())
}

pub(super) fn object_metadata_is_owned(
    object: &ArchiveObject,
    message_index: usize,
    identifier: NonZeroU64,
    accepted_path: &[u32],
    require_declared: bool,
) -> bool {
    object
        .archive_info
        .message_infos
        .get(message_index)
        .is_some_and(|info| {
            let aggregate = info
                .object_references
                .iter()
                .filter(|candidate| **candidate == identifier.get())
                .count();
            if info.data_references.contains(&identifier.get())
                || info
                    .field_infos
                    .iter()
                    .any(|field| field.data_references.contains(&identifier.get()))
                || aggregate > 1
            {
                return false;
            }
            let mut field = 0usize;
            for field_info in &info.field_infos {
                let occurrences = field_info
                    .object_references
                    .iter()
                    .filter(|candidate| **candidate == identifier.get())
                    .count();
                if occurrences == 0 {
                    continue;
                }
                if field_info.path.as_slice() != accepted_path || occurrences != 1 {
                    return false;
                }
                field = field.saturating_add(occurrences);
            }
            !require_declared || aggregate.saturating_add(field) != 0
        })
}

fn validate_drawable_order(
    package: &Package,
    order_identifier: NonZeroU64,
    drawable_identifier: NonZeroU64,
    budget: &mut ArrangementBudget,
) -> Result<(), BodyChartArrangementError> {
    let order = locate_unique_object(package, order_identifier, budget)?;
    let (message_index, message) = unique_message(order.object, DRAWABLE_ORDER_MESSAGE_TYPE)?;
    validate_message_metadata(order.object, message_index)?;
    let options = budget.order_options(&message.data)?;
    let (snapshot, report) = pages_drawable_order_codec::decode_drawable_order_with_report(
        message.data.as_slice(),
        options,
    )
    .map_err(map_drawable_order_error)?;
    budget.charge_order_report(report)?;
    if snapshot
        .identifiers()
        .filter(|identifier| *identifier == drawable_identifier.get())
        .count()
        != 1
    {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    Ok(())
}

fn rewrite_chart(
    source: &Package,
    target: &ChartTarget,
    after: ChartArrangement,
    budget: &mut ArrangementBudget,
) -> Result<Package, BodyChartArrangementError> {
    let original = selected_chart_payload(source, target)?;
    let rewritten = rewrite_arrangement(original, after, budget)?;
    Ok(rewrite_chart_payload(
        source,
        target,
        rewritten,
        PreviewDeletionMode::Preserve,
        budget,
    )?
    .package)
}

/// Publish one already validated chart-message replacement through the
/// bounded Pages component pipeline.
///
/// Arrange and focused chart-data transactions share the physical publication
/// boundary.  Keeping this helper here means both paths preserve the native
/// object header, rewrite one component, retain unrelated ZIP members, and
/// reopen the exact candidate under the same aggregate budget.
pub(super) fn rewrite_chart_payload(
    source: &Package,
    target: &ChartTarget,
    rewritten: Vec<u8>,
    preview_mode: PreviewDeletionMode,
    budget: &mut ArrangementBudget,
) -> Result<ChartPayloadPublication, BodyChartArrangementError> {
    let component = source
        .state
        .source
        .components()
        .get_index(target.component_index)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    if component.name() != target.component_name.as_ref() {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    let entry = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == target.component_name.as_ref())
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyChartArrangementError::UnsupportedSource);
    }
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let object = component
        .archive()
        .objects
        .get(target.drawable_object_index)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier.get()) {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    let original = object
        .messages
        .get(target.drawable_message_index)
        .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
        .ok_or(BodyChartArrangementError::InvalidSource)?
        .data
        .as_slice();
    let encoded_len = component
        .archive()
        .objects
        .iter()
        .try_fold(0usize, |bound, object| {
            let end = usize::try_from(object.data_offset)
                .ok()
                .and_then(|offset| {
                    usize::try_from(object.data_length)
                        .ok()
                        .and_then(|length| offset.checked_add(length))
                })
                .ok_or(BodyChartArrangementError::InvalidSource)?;
            Ok::<_, BodyChartArrangementError>(bound.max(end))
        })?;
    let replacement_bound = encoded_len
        .checked_sub(original.len())
        .and_then(|value| value.checked_add(rewritten.len()))
        .and_then(|value| value.checked_add(64))
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(replacement_bound).map_err(map_core_error)?;
    let package_bound = source
        .state
        .source
        .source_bytes()
        .len()
        .checked_sub(entry.data().len())
        .and_then(|value| value.checked_add(compressed_bound))
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    let (clone_allocations, clone_retained, clone_work) =
        archive_clone_requirements(&component.archive().objects, encoded_len)?;
    let (header_allocations, header_retained, header_work) =
        archive_header_rewrite_requirements(object)?;
    let (serialization_allocations, serialization_retained, serialization_work) =
        archive_serialization_requirements(component.archive())?;
    let physical = source.state.source.limits();
    let maximum_entry = usize::try_from(physical.max_entry_bytes())
        .map_err(|_| BodyChartArrangementError::InvalidSource)?;
    let maximum_compressed = physical
        .snappy_limits()
        .map_err(map_archive_error)?
        .max_compressed_stream();
    if replacement_bound > maximum_entry || compressed_bound > maximum_entry.min(maximum_compressed)
    {
        return Err(BodyChartArrangementError::LimitExceeded {
            kind: BodyChartArrangementLimitKind::EntryBytes,
            observed: replacement_bound.max(compressed_bound) as u64,
            maximum: maximum_entry.min(maximum_compressed) as u64,
        });
    }
    let frame_count = replacement_bound.div_ceil(SnappyStream::WRITE_CHUNK_SIZE);
    let compression_scratch =
        SnappyStream::maximum_compressed_len(replacement_bound.min(SnappyStream::WRITE_CHUNK_SIZE))
            .map_err(map_core_error)?
            .checked_add(32 * 1024)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
    // Pinned snap 1.1.2 retains a 16,384-entry u16 encoder table. Include it
    // alongside the decompressed stream, output, and transient frame buffers.
    let framing_allocations = frame_count
        .checked_add(3)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    let total_allocations = clone_allocations
        .checked_add(header_allocations)
        .and_then(|value| value.checked_add(serialization_allocations))
        .and_then(|value| value.checked_add(framing_allocations))
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    let total_work = clone_work
        .checked_add(header_work)
        .and_then(|value| value.checked_add(serialization_work))
        .and_then(|value| value.checked_add(replacement_bound))
        .and_then(|value| value.checked_add(compressed_bound))
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    let temporary_retained = clone_retained
        .checked_add(header_retained)
        .and_then(|value| value.checked_add(serialization_retained))
        .and_then(|value| value.checked_add(replacement_bound))
        .and_then(|value| value.checked_add(compressed_bound))
        .and_then(|value| value.checked_add(compression_scratch))
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    budget.preflight_output(replacement_bound)?;
    budget.preflight_output(compressed_bound)?;
    budget.preflight_output(package_bound)?;
    budget.preflight_work(total_work)?;
    budget.preflight_allocations(total_allocations)?;
    budget.preflight_retained(temporary_retained)?;
    budget.output(replacement_bound)?;
    budget.output(compressed_bound)?;
    budget.allocations(total_allocations)?;
    budget.work(total_work)?;
    budget.retained(clone_retained)?;
    budget.retained(header_retained)?;
    budget.retained(serialization_retained)?;
    budget.retained(replacement_bound)?;
    budget.retained(compressed_bound)?;
    budget.retained(compression_scratch)?;

    let (mut archive, _) = page_layout::editable_archive(source, target.component_name.as_ref())
        .map_err(|_| BodyChartArrangementError::InvalidSource)?;
    let object = archive
        .objects
        .get_mut(target.drawable_object_index)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    object
        .replace_message_preserving_header_with_limits(
            target.drawable_message_index,
            RawMessage {
                type_: CHART_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let compressed = page_layout::compress_archive(archive, archive_limits)
        .map_err(|_| BodyChartArrangementError::InvalidSource)?;
    let edits = [EntryEdit::new(target.component_name.as_ref(), &compressed)];
    let deleted_previews = match preview_mode {
        PreviewDeletionMode::Preserve => Vec::new(),
        PreviewDeletionMode::DeleteCanonicalRoot => root_preview_names(source, budget)?,
    };
    let prepared = source
        .state
        .source
        .package()
        .prepare_reassembly_with_deletions(&edits, &deleted_previews, source.state.source.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.reassembly(requirements)?;
    budget.candidate_reopen(requirements.output_bytes())?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let catalog =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), source.state.source.limits())
            .map_err(map_archive_error)?;
    let package = Package::from_source_catalog(catalog)
        .map_err(|_| BodyChartArrangementError::Verification)?;
    let target_previews = match preview_mode {
        PreviewDeletionMode::Preserve => 0,
        PreviewDeletionMode::DeleteCanonicalRoot => root_preview_names(&package, budget)?.len(),
    };
    Ok(ChartPayloadPublication {
        package,
        source_previews: deleted_previews.len(),
        target_previews,
    })
}

fn rewrite_arrangement(
    source: &[u8],
    after: ChartArrangement,
    budget: &mut ArrangementBudget,
) -> Result<Vec<u8>, BodyChartArrangementError> {
    let options = budget.codec_options(source)?;
    let prepared = codec::prepare_chart_arrangement_rewrite(
        source,
        codec::ChartArrangementWrite::new(after.locked(), after.constrain_proportions()),
        options,
    )
    .map_err(map_codec_error)?;
    budget.charge_arrangement_report(prepared.prepare_report())?;
    let requirements = prepared.execution_requirements();
    budget.charge_rewrite(requirements)?;
    prepared
        .execute(requirements.exact())
        .map(|output| output.into_output())
        .map_err(map_codec_error)
}

pub(super) fn selected_chart_payload<'source>(
    package: &'source Package,
    target: &ChartTarget,
) -> Result<&'source [u8], BodyChartArrangementError> {
    let component = package
        .state
        .source
        .components()
        .get_index(target.component_index)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    if component.name() != target.component_name.as_ref() {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    let object = component
        .archive()
        .objects
        .get(target.drawable_object_index)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    if object.archive_info.identifier != Some(target.drawable_identifier.get()) {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    let message = object
        .messages
        .get(target.drawable_message_index)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    if message.type_ != CHART_MESSAGE_TYPE {
        return Err(BodyChartArrangementError::InvalidSource);
    }
    Ok(message.data.as_slice())
}

/// Collect the canonical root preview members under the same bounded
/// allocation and work ledger as the surrounding chart transaction.
fn root_preview_names(
    package: &Package,
    budget: &mut ArrangementBudget,
) -> Result<Vec<&'static str>, BodyChartArrangementError> {
    budget.preflight_allocations(ROOT_PREVIEW_NAMES.len())?;
    budget.preflight_retained(
        ROOT_PREVIEW_NAMES
            .len()
            .checked_mul(size_of::<&'static str>())
            .ok_or(BodyChartArrangementError::InvalidSource)?,
    )?;
    budget.work(
        package
            .state
            .source
            .package()
            .len()
            .checked_mul(ROOT_PREVIEW_NAMES.len())
            .ok_or(BodyChartArrangementError::InvalidSource)?,
    )?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(ROOT_PREVIEW_NAMES.len())
        .map_err(|_| BodyChartArrangementError::Allocation {
            amount: ROOT_PREVIEW_NAMES.len(),
        })?;
    for name in ROOT_PREVIEW_NAMES {
        if package
            .state
            .source
            .package()
            .iter()
            .any(|entry| entry.name() == name)
        {
            names.push(name);
        }
    }
    budget.allocations(ROOT_PREVIEW_NAMES.len())?;
    budget.retained(
        ROOT_PREVIEW_NAMES
            .len()
            .checked_mul(size_of::<&'static str>())
            .ok_or(BodyChartArrangementError::InvalidSource)?,
    )?;
    Ok(names)
}

pub(super) fn retain_payload_pair(
    source: &[u8],
    target: &[u8],
    budget: &mut ArrangementBudget,
) -> Result<(Arc<[u8]>, Arc<[u8]>), BodyChartArrangementError> {
    let allocations = usize::from(!source.is_empty()) + usize::from(!target.is_empty());
    let retained = source
        .len()
        .checked_add(target.len())
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    budget.preflight_allocations(allocations)?;
    budget.preflight_retained(retained)?;
    let source = Arc::<[u8]>::from(source);
    let target = Arc::<[u8]>::from(target);
    budget.allocations(allocations)?;
    budget.retained(retained)?;
    Ok((source, target))
}

fn verify_payload_locality(
    source: &[u8],
    candidate: &[u8],
    mutation: ChartPayloadMutation,
    budget: &mut ArrangementBudget,
) -> Result<(), BodyChartArrangementError> {
    match mutation {
        ChartPayloadMutation::Arrangement => {
            verify_arrangement_payload_locality(source, candidate, budget)
        },
        ChartPayloadMutation::NumericData => {
            verify_chart_data_payload_locality(source, candidate, budget)
        },
    }
}

fn verify_arrangement_payload_locality(
    source: &[u8],
    candidate: &[u8],
    budget: &mut ArrangementBudget,
) -> Result<(), BodyChartArrangementError> {
    let source_view = budget.parse(source, 1)?;
    let candidate_view = budget.parse(candidate, 1)?;
    let source_super = unique_field(&source_view, CHART_DRAWABLE_FIELD, 2)?
        .ok_or(BodyChartArrangementError::Verification)?
        .payload();
    let candidate_super = unique_field(&candidate_view, CHART_DRAWABLE_FIELD, 2)?
        .ok_or(BodyChartArrangementError::Verification)?
        .payload();
    if !same_raw_except_numbers(
        source_view.fields(),
        candidate_view.fields(),
        &[CHART_DRAWABLE_FIELD],
        budget,
    )? {
        return Err(BodyChartArrangementError::Verification);
    }
    let source_super_view = budget.parse(source_super, 2)?;
    let candidate_super_view = budget.parse(candidate_super, 2)?;
    let source_locked = unique_field(&source_super_view, 5, 0)?;
    let candidate_locked = unique_field(&candidate_super_view, 5, 0)?;
    let source_constrain = unique_field(&source_super_view, 7, 0)?;
    let candidate_constrain = unique_field(&candidate_super_view, 7, 0)?;
    for field in [
        source_locked,
        candidate_locked,
        source_constrain,
        candidate_constrain,
    ]
    .into_iter()
    .flatten()
    {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let (value, width) = decode_varint_from_bytes(field.payload())
            .map_err(|_| BodyChartArrangementError::Verification)?;
        if width != encoded_len(value) || value > 1 {
            return Err(BodyChartArrangementError::Verification);
        }
    }
    if !same_raw_except_numbers(
        source_super_view.fields(),
        candidate_super_view.fields(),
        &[5, 7],
        budget,
    )? {
        return Err(BodyChartArrangementError::Verification);
    }
    Ok(())
}

const MODERN_CHART_EXTENSION_FIELD: u32 = 10_000;
const MODERN_CHART_GRID_FIELD: u32 = 7;
const MODERN_CHART_GRID_ROW_FIELD: u32 = 3;
const MODERN_CHART_GRID_VALUE_FIELD: u32 = 1;

fn verify_chart_data_payload_locality(
    source: &[u8],
    candidate: &[u8],
    budget: &mut ArrangementBudget,
) -> Result<(), BodyChartArrangementError> {
    let source_view = budget.parse(source, 1)?;
    let candidate_view = budget.parse(candidate, 1)?;
    let source_chart = unique_field(&source_view, MODERN_CHART_EXTENSION_FIELD, 2)?
        .ok_or(BodyChartArrangementError::Verification)?
        .payload();
    let candidate_chart = unique_field(&candidate_view, MODERN_CHART_EXTENSION_FIELD, 2)?
        .ok_or(BodyChartArrangementError::Verification)?
        .payload();
    if !same_raw_except_numbers(
        source_view.fields(),
        candidate_view.fields(),
        &[MODERN_CHART_EXTENSION_FIELD],
        budget,
    )? {
        return Err(BodyChartArrangementError::Verification);
    }

    let source_chart_view = budget.parse(source_chart, 2)?;
    let candidate_chart_view = budget.parse(candidate_chart, 2)?;
    let source_grid = unique_field(&source_chart_view, MODERN_CHART_GRID_FIELD, 2)?
        .ok_or(BodyChartArrangementError::Verification)?
        .payload();
    let candidate_grid = unique_field(&candidate_chart_view, MODERN_CHART_GRID_FIELD, 2)?
        .ok_or(BodyChartArrangementError::Verification)?
        .payload();
    if !same_raw_except_numbers(
        source_chart_view.fields(),
        candidate_chart_view.fields(),
        &[MODERN_CHART_GRID_FIELD],
        budget,
    )? {
        return Err(BodyChartArrangementError::Verification);
    }

    let source_grid_view = budget.parse(source_grid, 3)?;
    let candidate_grid_view = budget.parse(candidate_grid, 3)?;
    let mut source_rows = source_grid_view
        .fields()
        .filter(|field| field.number() == MODERN_CHART_GRID_ROW_FIELD);
    let mut candidate_rows = candidate_grid_view
        .fields()
        .filter(|field| field.number() == MODERN_CHART_GRID_ROW_FIELD);
    if !same_raw_except_numbers(
        source_grid_view.fields(),
        candidate_grid_view.fields(),
        &[MODERN_CHART_GRID_ROW_FIELD],
        budget,
    )? {
        return Err(BodyChartArrangementError::Verification);
    }
    loop {
        match (source_rows.next(), candidate_rows.next()) {
            (Some(source_row), Some(candidate_row)) => {
                if source_row.wire_type() != 2 || candidate_row.wire_type() != 2 {
                    return Err(BodyChartArrangementError::Verification);
                }
                source_row
                    .validate_canonical_framing()
                    .map_err(map_wire_error)?;
                candidate_row
                    .validate_canonical_framing()
                    .map_err(map_wire_error)?;
                verify_chart_data_row(source_row.payload(), candidate_row.payload(), budget)?;
            },
            (None, None) => break,
            _ => return Err(BodyChartArrangementError::Verification),
        }
    }
    Ok(())
}

fn verify_chart_data_row(
    source: &[u8],
    candidate: &[u8],
    budget: &mut ArrangementBudget,
) -> Result<(), BodyChartArrangementError> {
    let source_view = budget.parse(source, 4)?;
    let candidate_view = budget.parse(candidate, 4)?;
    if !same_raw_except_numbers(
        source_view.fields(),
        candidate_view.fields(),
        &[MODERN_CHART_GRID_VALUE_FIELD],
        budget,
    )? {
        return Err(BodyChartArrangementError::Verification);
    }
    let mut source_values = source_view
        .fields()
        .filter(|field| field.number() == MODERN_CHART_GRID_VALUE_FIELD);
    let mut candidate_values = candidate_view
        .fields()
        .filter(|field| field.number() == MODERN_CHART_GRID_VALUE_FIELD);
    loop {
        match (source_values.next(), candidate_values.next()) {
            (Some(source_value), Some(candidate_value)) => {
                if source_value.wire_type() != 2 || candidate_value.wire_type() != 2 {
                    return Err(BodyChartArrangementError::Verification);
                }
                source_value
                    .validate_canonical_framing()
                    .map_err(map_wire_error)?;
                candidate_value
                    .validate_canonical_framing()
                    .map_err(map_wire_error)?;
                verify_chart_data_value(source_value.payload(), candidate_value.payload(), budget)?;
            },
            (None, None) => break,
            _ => return Err(BodyChartArrangementError::Verification),
        }
    }
    Ok(())
}

fn verify_chart_data_value(
    source: &[u8],
    candidate: &[u8],
    budget: &mut ArrangementBudget,
) -> Result<(), BodyChartArrangementError> {
    let source_view = budget.parse(source, 5)?;
    let candidate_view = budget.parse(candidate, 5)?;
    let _source_numeric = unique_field(&source_view, MODERN_CHART_GRID_VALUE_FIELD, 1)?;
    let _candidate_numeric = unique_field(&candidate_view, MODERN_CHART_GRID_VALUE_FIELD, 1)?;
    if !same_raw_except_numbers(
        source_view.fields(),
        candidate_view.fields(),
        &[MODERN_CHART_GRID_VALUE_FIELD],
        budget,
    )? {
        return Err(BodyChartArrangementError::Verification);
    }
    Ok(())
}

fn same_raw_except_numbers<'left, 'right>(
    left: impl Iterator<Item = WireFieldView<'left>>,
    right: impl Iterator<Item = WireFieldView<'right>>,
    excluded: &[u32],
    budget: &mut ArrangementBudget,
) -> Result<bool, BodyChartArrangementError> {
    let mut left = left.filter(|field| !excluded.contains(&field.number()));
    let mut right = right.filter(|field| !excluded.contains(&field.number()));
    loop {
        match (left.next(), right.next()) {
            (Some(left), Some(right)) => {
                budget.work(
                    left.raw()
                        .len()
                        .checked_add(right.raw().len())
                        .ok_or(BodyChartArrangementError::InvalidSource)?,
                )?;
                if left.raw() != right.raw() {
                    return Ok(false);
                }
            },
            (None, None) => return Ok(true),
            _ => return Ok(false),
        }
    }
}

fn add_clone_vec(
    length: usize,
    element_size: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), BodyChartArrangementError> {
    if length == 0 {
        return Ok(());
    }
    *allocations = allocations
        .checked_add(1)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    *retained = retained
        .checked_add(
            length
                .checked_mul(element_size)
                .ok_or(BodyChartArrangementError::InvalidSource)?,
        )
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    Ok(())
}

fn add_clone_bytes(
    length: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), BodyChartArrangementError> {
    add_clone_vec(length, 1, allocations, retained)
}

/// Account for a complete decoded object walk used by archive clone and
/// locality proof operations.  Archive metadata contains several nested
/// vectors; counting only top-level objects would leave those walks outside
/// the transaction ledger.
fn archive_object_work(object: &ArchiveObject) -> Result<usize, BodyChartArrangementError> {
    let mut work = size_of::<ArchiveObject>();
    work = work
        .checked_add(object.messages.len())
        .and_then(|value| value.checked_add(object.archive_info.message_infos.len()))
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    for message in &object.messages {
        work = work
            .checked_add(size_of::<RawMessage>())
            .and_then(|value| value.checked_add(message.data.len()))
            .ok_or(BodyChartArrangementError::InvalidSource)?;
    }
    for info in &object.archive_info.message_infos {
        work = work
            .checked_add(size_of::<litchi_iwa_core::MessageInfo>())
            .and_then(|value| value.checked_add(info.versions.len()))
            .and_then(|value| value.checked_add(info.field_infos.len()))
            .and_then(|value| value.checked_add(info.object_references.len()))
            .and_then(|value| value.checked_add(info.data_references.len()))
            .and_then(|value| value.checked_add(info.diff_merge_version.len()))
            .and_then(|value| value.checked_add(info.fields_to_remove.len()))
            .and_then(|value| value.checked_add(info.diff_read_version.len()))
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        if let Some(path) = info.diff_field_path.as_ref() {
            work = work
                .checked_add(path.path.len())
                .ok_or(BodyChartArrangementError::InvalidSource)?;
        }
        for path in &info.fields_to_remove {
            work = work
                .checked_add(path.path.len())
                .ok_or(BodyChartArrangementError::InvalidSource)?;
        }
        for field in &info.field_infos {
            work = work
                .checked_add(size_of::<litchi_iwa_core::FieldInfo>())
                .and_then(|value| value.checked_add(field.path.path.len()))
                .and_then(|value| value.checked_add(field.object_references.len()))
                .and_then(|value| value.checked_add(field.data_references.len()))
                .and_then(|value| value.checked_add(field.known_field_version.len()))
                .ok_or(BodyChartArrangementError::InvalidSource)?;
            if let Some(identifier) = field.known_field_feature_identifier.as_ref() {
                work = work
                    .checked_add(identifier.len())
                    .ok_or(BodyChartArrangementError::InvalidSource)?;
            }
        }
    }
    Ok(work)
}

fn archive_clone_requirements(
    objects: &[ArchiveObject],
    encoded_len: usize,
) -> Result<(usize, usize, usize), BodyChartArrangementError> {
    let mut allocations = 0usize;
    let mut retained = encoded_len;
    let mut work = 0usize;
    add_clone_vec(
        objects.len(),
        size_of::<ArchiveObject>(),
        &mut allocations,
        &mut retained,
    )?;
    for object in objects {
        work = work
            .checked_add(archive_object_work(object)?)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        add_clone_vec(
            object.messages.len(),
            size_of::<RawMessage>(),
            &mut allocations,
            &mut retained,
        )?;
        add_clone_vec(
            object.archive_info.message_infos.len(),
            size_of::<litchi_iwa_core::MessageInfo>(),
            &mut allocations,
            &mut retained,
        )?;
        if object.header_length != 0 {
            allocations = allocations
                .checked_add(2)
                .ok_or(BodyChartArrangementError::InvalidSource)?;
            retained = retained
                .checked_add(
                    usize::try_from(object.header_length)
                        .map_err(|_| BodyChartArrangementError::InvalidSource)?
                        .checked_mul(2)
                        .ok_or(BodyChartArrangementError::InvalidSource)?,
                )
                .ok_or(BodyChartArrangementError::InvalidSource)?;
        }
        for message in &object.messages {
            add_clone_bytes(message.data.len(), &mut allocations, &mut retained)?;
        }
        for info in &object.archive_info.message_infos {
            add_clone_vec(
                info.versions.len(),
                size_of::<u32>(),
                &mut allocations,
                &mut retained,
            )?;
            add_clone_vec(
                info.field_infos.len(),
                size_of::<litchi_iwa_core::FieldInfo>(),
                &mut allocations,
                &mut retained,
            )?;
            add_clone_vec(
                info.object_references.len(),
                size_of::<u64>(),
                &mut allocations,
                &mut retained,
            )?;
            add_clone_vec(
                info.data_references.len(),
                size_of::<u64>(),
                &mut allocations,
                &mut retained,
            )?;
            add_clone_vec(
                info.diff_merge_version.len(),
                size_of::<u32>(),
                &mut allocations,
                &mut retained,
            )?;
            if let Some(path) = info.diff_field_path.as_ref() {
                add_clone_vec(
                    path.path.len(),
                    size_of::<u32>(),
                    &mut allocations,
                    &mut retained,
                )?;
            }
            add_clone_vec(
                info.fields_to_remove.len(),
                size_of::<litchi_iwa_core::FieldPath>(),
                &mut allocations,
                &mut retained,
            )?;
            for path in &info.fields_to_remove {
                add_clone_vec(
                    path.path.len(),
                    size_of::<u32>(),
                    &mut allocations,
                    &mut retained,
                )?;
            }
            add_clone_vec(
                info.diff_read_version.len(),
                size_of::<u32>(),
                &mut allocations,
                &mut retained,
            )?;
            for field in &info.field_infos {
                add_clone_vec(
                    field.path.path.len(),
                    size_of::<u32>(),
                    &mut allocations,
                    &mut retained,
                )?;
                add_clone_vec(
                    field.object_references.len(),
                    size_of::<u64>(),
                    &mut allocations,
                    &mut retained,
                )?;
                add_clone_vec(
                    field.data_references.len(),
                    size_of::<u64>(),
                    &mut allocations,
                    &mut retained,
                )?;
                add_clone_vec(
                    field.known_field_version.len(),
                    size_of::<u32>(),
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

fn archive_header_rewrite_requirements(
    object: &ArchiveObject,
) -> Result<(usize, usize, usize), BodyChartArrangementError> {
    let header = usize::try_from(object.header_length)
        .map_err(|_| BodyChartArrangementError::InvalidSource)?;
    let retained = header
        .checked_add(64)
        .and_then(|value| value.checked_mul(3))
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    Ok((3, retained, archive_object_work(object)?))
}

fn archive_serialization_requirements(
    archive: &litchi_iwa_core::Archive,
) -> Result<(usize, usize, usize), BodyChartArrangementError> {
    let mut retained = 0usize;
    let mut work = 0usize;
    for object in &archive.objects {
        let header = usize::try_from(object.header_length)
            .map_err(|_| BodyChartArrangementError::InvalidSource)?;
        retained = retained
            .checked_add(header)
            .and_then(|amount| amount.checked_add(64))
            .ok_or(BodyChartArrangementError::InvalidSource)?;
        work = work
            .checked_add(archive_object_work(object)?)
            .ok_or(BodyChartArrangementError::InvalidSource)?;
    }
    let allocations = archive
        .objects
        .len()
        .checked_add(2)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    Ok((allocations, retained, work))
}

fn archive_object_clone_requirements(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    replacement_len: usize,
) -> Result<(usize, usize, usize), BodyChartArrangementError> {
    let (mut allocations, mut retained, work) =
        archive_clone_requirements(std::slice::from_ref(source), 0)?;
    let (header_allocations, header_retained, header_work) =
        archive_header_rewrite_requirements(source)?;
    allocations = allocations
        .checked_add(header_allocations)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    retained = retained
        .checked_add(header_retained)
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    add_clone_bytes(replacement_len, &mut allocations, &mut retained)?;
    let work = work
        .checked_add(header_work)
        .and_then(|value| value.checked_add(replacement_len))
        .and_then(|value| value.checked_add(archive_object_work(candidate).ok()?))
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    Ok((allocations, retained, work))
}

pub(super) fn verify_locality(
    source: &Package,
    candidate: &Package,
    target: &ChartTarget,
    expected_payload: Option<&[u8]>,
    preview_counts: Option<(usize, usize)>,
    mutation: ChartPayloadMutation,
    budget: &mut ArrangementBudget,
) -> Result<(), BodyChartArrangementError> {
    let source_entries = source.state.source.package();
    let candidate_entries = candidate.state.source.package();
    let source_preview_names = root_preview_names(source, budget)?;
    let candidate_preview_names = root_preview_names(candidate, budget)?;
    let (source_preview_count, candidate_preview_count) =
        preview_counts.unwrap_or((source_preview_names.len(), source_preview_names.len()));
    if source_preview_names.len() != source_preview_count
        || candidate_preview_names.len() != candidate_preview_count
        || (preview_counts.is_none() && source_preview_names != candidate_preview_names)
    {
        return Err(BodyChartArrangementError::Verification);
    }
    for name in ROOT_PREVIEW_NAMES {
        let source_preview = source_entries.iter().find(|entry| entry.name() == name);
        let candidate_preview = candidate_entries.iter().find(|entry| entry.name() == name);
        if let (Some(before), Some(after)) = (source_preview, candidate_preview) {
            let work = [
                before.name().len(),
                before.raw_name().len(),
                before.data().len(),
                before.raw_record().local_record().len(),
                before.raw_record().central_directory_record().len(),
                before.raw_record().compressed_data().len(),
                after.name().len(),
                after.raw_name().len(),
                after.data().len(),
                after.raw_record().local_record().len(),
                after.raw_record().central_directory_record().len(),
                after.raw_record().compressed_data().len(),
            ]
            .into_iter()
            .try_fold(0usize, |total, amount| total.checked_add(amount))
            .ok_or(BodyChartArrangementError::InvalidSource)?;
            budget.work(work)?;
            if before.raw_name() != after.raw_name()
                || before.data() != after.data()
                || before.metadata() != after.metadata()
                || before.raw_record().local_record() != after.raw_record().local_record()
                || !zip_record_equal_outside(
                    before.raw_record().central_directory_record(),
                    after.raw_record().central_directory_record(),
                    &[(42, 46)],
                )
                || before.raw_record().compressed_data() != after.raw_record().compressed_data()
            {
                return Err(BodyChartArrangementError::Verification);
            }
        }
    }
    let expected_entry_count = source_entries
        .len()
        .checked_sub(source_preview_names.len())
        .and_then(|count| count.checked_add(candidate_preview_names.len()))
        .ok_or(BodyChartArrangementError::InvalidSource)?;
    if candidate_entries.len() != expected_entry_count {
        return Err(BodyChartArrangementError::Verification);
    }
    for (before, after) in source_entries
        .iter()
        .filter(|entry| !source_preview_names.contains(&entry.name()))
        .zip(
            candidate_entries
                .iter()
                .filter(|entry| !candidate_preview_names.contains(&entry.name())),
        )
    {
        let mut work = 0usize;
        for entry in [before, after] {
            for amount in [
                entry.name().len(),
                entry.raw_name().len(),
                entry.data().len(),
                entry.raw_record().local_record().len(),
                entry.raw_record().central_directory_record().len(),
            ] {
                work = work
                    .checked_add(amount)
                    .ok_or(BodyChartArrangementError::InvalidSource)?;
            }
        }
        budget.work(
            work.checked_mul(2)
                .ok_or(BodyChartArrangementError::InvalidSource)?,
        )?;
        if before.name() != after.name()
            || before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
            || before.metadata().local() != after.metadata().local()
            || before.metadata().central() != after.metadata().central()
        {
            return Err(BodyChartArrangementError::Verification);
        }
        if before.name() == target.component_name.as_ref() {
            if !zip_local_header_preserved(
                before.raw_record().local_record(),
                after.raw_record().local_record(),
            ) || !zip_record_equal_outside(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
                &[(16, 28), (42, 46)],
            ) {
                return Err(BodyChartArrangementError::Verification);
            }
        } else if before.data() != after.data()
            || before.raw_record().local_record() != after.raw_record().local_record()
            || !zip_record_equal_outside(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
                &[(42, 46)],
            )
        {
            return Err(BodyChartArrangementError::Verification);
        }
    }
    let source_components = source.state.source.components();
    let candidate_components = candidate.state.source.components();
    if source_components.len() != candidate_components.len() {
        return Err(BodyChartArrangementError::Verification);
    }
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut selected_seen = false;
    for (before_component, after_component) in
        source_components.iter().zip(candidate_components.iter())
    {
        if before_component.name() != after_component.name()
            || before_component.archive().objects.len() != after_component.archive().objects.len()
        {
            return Err(BodyChartArrangementError::Verification);
        }
        for (object_index, (before, after)) in before_component
            .archive()
            .objects
            .iter()
            .zip(&after_component.archive().objects)
            .enumerate()
        {
            let work = archive_object_work(before)?
                .checked_add(archive_object_work(after)?)
                .ok_or(BodyChartArrangementError::InvalidSource)?;
            budget.work(work)?;
            let selected = before_component.name() == target.component_name.as_ref()
                && object_index == target.drawable_object_index;
            if !selected {
                if !before.same_content_ignoring_offsets(after) {
                    return Err(BodyChartArrangementError::Verification);
                }
                continue;
            }
            selected_seen = true;
            let message = after
                .messages
                .get(target.drawable_message_index)
                .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
                .ok_or(BodyChartArrangementError::Verification)?;
            if let Some(expected) = expected_payload {
                budget.work(
                    message
                        .data
                        .len()
                        .checked_add(expected.len())
                        .ok_or(BodyChartArrangementError::InvalidSource)?,
                )?;
                if message.data != expected {
                    return Err(BodyChartArrangementError::Verification);
                }
            }
            verify_payload_locality(
                before
                    .messages
                    .get(target.drawable_message_index)
                    .filter(|message| message.type_ == CHART_MESSAGE_TYPE)
                    .ok_or(BodyChartArrangementError::Verification)?
                    .data
                    .as_slice(),
                message.data.as_slice(),
                mutation,
                budget,
            )?;
            let (allocations, retained, clone_work) =
                archive_object_clone_requirements(before, after, message.data.len())?;
            budget.allocations(allocations)?;
            budget.retained(retained)?;
            budget.work(clone_work)?;
            let mut expected = before.clone();
            expected
                .replace_message_preserving_header_with_limits(
                    target.drawable_message_index,
                    RawMessage {
                        type_: message.type_,
                        data: message.data.clone(),
                    },
                    archive_limits,
                )
                .map_err(map_core_error)?;
            expected.header_length = after.header_length;
            expected.data_length = after.data_length;
            if !expected.same_content_ignoring_offsets(after) {
                return Err(BodyChartArrangementError::Verification);
            }
        }
    }
    if !selected_seen {
        return Err(BodyChartArrangementError::Verification);
    }
    Ok(())
}

fn zip_record_equal_outside(before: &[u8], after: &[u8], excluded: &[(usize, usize)]) -> bool {
    if before.len() != after.len() {
        return false;
    }
    let mut cursor = 0;
    for &(start, end) in excluded {
        if start < cursor
            || end < start
            || end > before.len()
            || before[cursor..start] != after[cursor..start]
        {
            return false;
        }
        cursor = end;
    }
    before[cursor..] == after[cursor..]
}

fn zip_local_header_preserved(before: &[u8], after: &[u8]) -> bool {
    let Some(before_header) = zip_local_header_len(before) else {
        return false;
    };
    let Some(after_header) = zip_local_header_len(after) else {
        return false;
    };
    before_header == after_header
        && zip_record_equal_outside(
            &before[..before_header],
            &after[..after_header],
            &[(14, 26)],
        )
}

fn zip_local_header_len(record: &[u8]) -> Option<usize> {
    if record.len() < 30 || record.get(..4) != Some(b"PK\x03\x04") {
        return None;
    }
    let name = usize::from(u16::from_le_bytes([record[26], record[27]]));
    let extra = usize::from(u16::from_le_bytes([record[28], record[29]]));
    30usize
        .checked_add(name)?
        .checked_add(extra)
        .filter(|length| *length <= record.len())
}

fn push_unique(values: &mut Vec<NonZeroU64>, value: NonZeroU64) -> bool {
    if values.contains(&value) {
        false
    } else {
        values.push(value);
        true
    }
}

fn map_codec_error(error: codec::DecodeError) -> BodyChartArrangementError {
    if let Some(amount) = error.allocation_amount() {
        return BodyChartArrangementError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return BodyChartArrangementError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::Bytes { observed, maximum } => {
            (BodyChartArrangementLimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::Output { observed, maximum } => (
            BodyChartArrangementLimitKind::WireOutputBytes,
            observed,
            maximum,
        ),
        codec::DecodeLimit::Fields { observed, maximum } => {
            (BodyChartArrangementLimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::Work { observed, maximum } => {
            (BodyChartArrangementLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => {
            return BodyChartArrangementError::LimitExceeded {
                kind: BodyChartArrangementLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            };
        },
        codec::DecodeLimit::Allocations { observed, maximum } => (
            BodyChartArrangementLimitKind::WireAllocations,
            observed,
            maximum,
        ),
        codec::DecodeLimit::Retained { observed, maximum } => (
            BodyChartArrangementLimitKind::WireRetainedBytes,
            observed,
            maximum,
        ),
        codec::DecodeLimit::Scratch { observed, maximum } => (
            BodyChartArrangementLimitKind::WireScratchBytes,
            observed,
            maximum,
        ),
        _ => return BodyChartArrangementError::InvalidSource,
    };
    BodyChartArrangementError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn map_drawable_order_error(
    error: pages_drawable_order_codec::DecodeError,
) -> BodyChartArrangementError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            pages_drawable_order_codec::WireResourceLimit::InputBytes { observed, maximum } => {
                (BodyChartArrangementLimitKind::WireBytes, observed, maximum)
            },
            pages_drawable_order_codec::WireResourceLimit::OutputBytes { observed, maximum } => (
                BodyChartArrangementLimitKind::WireOutputBytes,
                observed,
                maximum,
            ),
            pages_drawable_order_codec::WireResourceLimit::Fields { observed, maximum } => {
                (BodyChartArrangementLimitKind::WireFields, observed, maximum)
            },
            pages_drawable_order_codec::WireResourceLimit::WorkBytes { observed, maximum } => {
                (BodyChartArrangementLimitKind::WireWork, observed, maximum)
            },
            pages_drawable_order_codec::WireResourceLimit::Nesting { observed, maximum } => {
                return BodyChartArrangementError::LimitExceeded {
                    kind: BodyChartArrangementLimitKind::WireNesting,
                    observed: observed as u64,
                    maximum: maximum as u64,
                };
            },
            pages_drawable_order_codec::WireResourceLimit::References { observed, maximum } => (
                BodyChartArrangementLimitKind::PayloadReferences,
                observed,
                maximum,
            ),
            _ => return BodyChartArrangementError::InvalidSource,
        };
        return BodyChartArrangementError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return BodyChartArrangementError::Allocation { amount };
    }
    BodyChartArrangementError::InvalidSource
}

fn map_wire_error(error: litchi_iwa_common::Error) -> BodyChartArrangementError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => BodyChartArrangementError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    BodyChartArrangementLimitKind::WireBytes
                },
                litchi_iwa_common::LimitKind::OutputBytes => {
                    BodyChartArrangementLimitKind::WireOutputBytes
                },
                litchi_iwa_common::LimitKind::Fields => BodyChartArrangementLimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => BodyChartArrangementLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => {
                    BodyChartArrangementLimitKind::WireWork
                },
                _ => BodyChartArrangementLimitKind::PayloadItems,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            BodyChartArrangementError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => BodyChartArrangementError::InvalidSource,
    }
}

fn map_archive_error(error: ArchiveError) -> BodyChartArrangementError {
    match error {
        ArchiveError::Limit {
            kind,
            observed,
            maximum,
        } => BodyChartArrangementError::LimitExceeded {
            kind: match kind {
                ArchiveLimitKind::InputBytes => BodyChartArrangementLimitKind::InputBytes,
                ArchiveLimitKind::OutputBytes => BodyChartArrangementLimitKind::OutputBytes,
                ArchiveLimitKind::Entries => BodyChartArrangementLimitKind::Entries,
                ArchiveLimitKind::MemberNameBytes | ArchiveLimitKind::MetadataBytes => {
                    BodyChartArrangementLimitKind::PackageBytes
                },
                ArchiveLimitKind::CompressedEntryBytes | ArchiveLimitKind::EntryBytes => {
                    BodyChartArrangementLimitKind::EntryBytes
                },
                ArchiveLimitKind::TotalBytes => BodyChartArrangementLimitKind::TotalEntryBytes,
                ArchiveLimitKind::IwaStreamBytes => BodyChartArrangementLimitKind::PayloadBytes,
                ArchiveLimitKind::IwaTotalBytes => BodyChartArrangementLimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
        },
        ArchiveError::Allocation { amount, .. } => BodyChartArrangementError::Allocation { amount },
        ArchiveError::Iwa(error) => map_core_error(error),
        ArchiveError::Reassembly(_) => BodyChartArrangementError::UnsupportedSource,
        ArchiveError::Io(_)
        | ArchiveError::Zip { .. }
        | ArchiveError::InvalidLimits(_)
        | ArchiveError::Encrypted
        | ArchiveError::SourceChanged { .. }
        | ArchiveError::DirectoryChanged { .. }
        | ArchiveError::InvalidBundle(_) => BodyChartArrangementError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyChartArrangementError {
    match error {
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyChartArrangementError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyChartArrangementError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => {
                    BodyChartArrangementLimitKind::PayloadObjects
                },
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    BodyChartArrangementLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderFields => {
                    BodyChartArrangementLimitKind::WireFields
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    BodyChartArrangementLimitKind::WireNesting
                },
                _ => BodyChartArrangementLimitKind::PayloadBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        _ => BodyChartArrangementError::InvalidSource,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn zip_masks_preserve_timestamps_and_allow_complete_size_fields() {
        let central = [0u8; 46];
        for index in 0..central.len() {
            let mut changed = central;
            changed[index] = 1;
            assert_eq!(
                zip_record_equal_outside(&central, &changed, &[(16, 28)]),
                (16..28).contains(&index)
            );
            assert_eq!(
                zip_record_equal_outside(&central, &changed, &[(42, 46)]),
                (42..46).contains(&index)
            );
        }
        let mut local = [0u8; 30];
        local[..4].copy_from_slice(b"PK\x03\x04");
        for index in 10..26 {
            let mut changed = local;
            changed[index] = 1;
            assert_eq!(
                zip_local_header_preserved(&local, &changed),
                (14..26).contains(&index)
            );
        }
    }
}
