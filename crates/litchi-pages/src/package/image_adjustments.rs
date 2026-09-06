//! Selector-first, source-preserving Pages body-image adjustment transactions.
//!
//! A body image is discovered from the rooted body storage rather than from a
//! caller supplied native identifier. The graph proof borrows the parsed
//! [`SourceCatalog`], and only the selected ImageArchive payload is handed to
//! the neutral bounded Buffa codec. Unknown ImageArchive fields therefore
//! remain source-owned and are preserved by every rewrite.

use std::fmt;
use std::mem::size_of;
use std::num::NonZeroU64;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{image_adjustments_codec, pages_body_codec, pages_drawable_order_codec};
use thiserror::Error;

use super::Package;
use crate::selector::ImageSelector;

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
const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const IMAGE_DRAWABLE_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const DRAWABLE_TITLE_FIELD: u32 = 10;
const DRAWABLE_CAPTION_FIELD: u32 = 11;
const IMAGE_DATA_FIELD: u32 = 11;
const DRAWABLE_ORDER_MESSAGE_TYPE: u32 = 10_015;

/// A finite resource charged by one body-image adjustment operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyImageAdjustmentsLimitKind {
    /// Bytes in native payloads inspected by the graph proof.
    InputBytes,
    /// Bytes in the rewritten package artifact.
    OutputBytes,
    /// Parsed wire fields.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Aggregate wire traversal and reassembly work.
    WireWork,
    /// Native object references inspected.
    References,
    /// Temporary allocation units.
    Allocations,
    /// Retained temporary bytes.
    Retained,
}

impl fmt::Display for BodyImageAdjustmentsLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::References => "references",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
        })
    }
}

/// Failure from a Pages body-image adjustment read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum BodyImageAdjustmentsError {
    /// No ordinary body image matched the checked source-order selector.
    #[error("the Pages body has no image at position {position:?}")]
    ImageNotFound { position: Position },
    /// The source does not retain an exact physical artifact suitable for a
    /// changed publication.
    #[error("this Pages source does not support exact body-image adjustment edits")]
    UnsupportedSource,
    /// The rooted image graph or selected payload is malformed or ambiguous.
    #[error("the selected Pages body-image adjustment source is invalid")]
    InvalidSource,
    /// A finite operation resource ceiling was exceeded.
    #[error(
        "Pages body-image adjustments {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: BodyImageAdjustmentsLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded temporary allocation failed.
    #[error("could not allocate {amount} units for Pages body-image adjustments")]
    Allocation { amount: usize },
    /// Reopening the candidate did not reproduce the requested semantic state.
    #[error("the edited Pages body-image adjustments failed semantic verification")]
    Verification,
    /// The patch was produced from another exact package artifact.
    #[error("the Pages body-image adjustment patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable image-adjustment value staged against an immutable package.
pub struct BodyImageAdjustmentsEdit<'a> {
    source: &'a Package,
    target: ImageTarget,
    before: ImageAdjustments,
    after: ImageAdjustments,
}

impl fmt::Debug for BodyImageAdjustmentsEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyImageAdjustmentsEdit")
            .field("position", &self.target.position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyImageAdjustmentsEdit<'_> {
    /// Return the selected body-image source position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.target.position
    }

    /// Return the adjustment controls read before this edit.
    #[must_use]
    pub const fn before(&self) -> ImageAdjustments {
        self.before
    }

    /// Return the adjustment controls currently staged for publication.
    #[must_use]
    pub const fn after(&self) -> ImageAdjustments {
        self.after
    }

    /// Replace all three optional inspector controls.
    pub fn set(mut self, adjustments: ImageAdjustments) -> Result<Self, BodyImageAdjustmentsError> {
        self.after = adjustments;
        Ok(self)
    }

    /// Validate and publish the staged exact-source edit.
    pub fn commit(self) -> Result<BodyImageAdjustmentsCommit, BodyImageAdjustmentsError> {
        commit_edit(self)
    }
}

/// Exact-source checked reversible body-image adjustment patch.
#[derive(Clone, PartialEq)]
pub struct BodyImageAdjustmentsPatch {
    artifacts: ExactArtifacts,
    target: ImageTarget,
    before: ImageAdjustments,
    after: ImageAdjustments,
    touched_components: usize,
}

impl fmt::Debug for BodyImageAdjustmentsPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyImageAdjustmentsPatch")
            .field("position", &self.target.position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl BodyImageAdjustmentsPatch {
    /// Return the selected body-image source position.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.target.position
    }

    /// Return the semantic controls required before this patch applies.
    #[must_use]
    pub const fn before(&self) -> ImageAdjustments {
        self.before
    }

    /// Return the semantic controls produced by this patch.
    #[must_use]
    pub const fn after(&self) -> ImageAdjustments {
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

    /// Return whether both the semantic value and exact source are unchanged.
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
        }
    }
}

/// Content-free diagnostics from one body-image adjustment publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct BodyImageAdjustmentsDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl BodyImageAdjustmentsDiagnostics {
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

    /// Return whether package bytes and semantic state changed.
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

/// Fully validated immutable result of one body-image adjustment transaction.
#[must_use = "a body-image adjustment commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct BodyImageAdjustmentsCommit {
    package: Package,
    patch: BodyImageAdjustmentsPatch,
    diagnostics: BodyImageAdjustmentsDiagnostics,
}

impl BodyImageAdjustmentsCommit {
    /// Borrow the fully reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact-source reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyImageAdjustmentsPatch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyImageAdjustmentsDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq, Eq)]
struct ImageTarget {
    position: Position,
    body_identifier: NonZeroU64,
    attachment_identifier: NonZeroU64,
    drawable_identifier: NonZeroU64,
    data_identifier: NonZeroU64,
    character_index: u32,
    message_index: usize,
    component_name: Arc<str>,
}

impl fmt::Debug for ImageTarget {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ImageTarget")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

#[derive(Clone, Copy)]
struct LocatedObject<'a> {
    component_index: usize,
    object: &'a ArchiveObject,
}

#[derive(Debug, Clone, Copy)]
struct ImageBudget {
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

impl ImageBudget {
    #[cfg(test)]
    fn for_test() -> Self {
        let limits = WireLimits::default();
        Self {
            limits,
            max_input: usize::MAX,
            max_output: usize::MAX,
            max_fields: usize::MAX,
            max_work: usize::MAX,
            max_nesting: usize::MAX,
            max_references: usize::MAX,
            max_allocations: usize::MAX,
            max_retained: usize::MAX,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            references: 0,
            allocations: 0,
            retained: 0,
        }
    }

    fn new(package: &Package) -> Result<Self, BodyImageAdjustmentsError> {
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
        let max_input = physical.max_iwa_stream_bytes().saturating_mul(4).max(1);
        let max_output = usize::try_from(physical.max_input_bytes())
            .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?
            .max(1);
        Ok(Self {
            limits,
            max_input,
            max_output,
            max_fields: limits.max_fields(),
            max_work: usize::try_from(physical.max_total_bytes())
                .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?
                .saturating_mul(32)
                .max(limits.max_rewrite_work()),
            max_nesting: limits.max_nesting(),
            max_references: archive.max_metadata_items().saturating_mul(8).max(1),
            max_allocations: archive.max_metadata_items().saturating_mul(8).max(1),
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
        kind: BodyImageAdjustmentsLimitKind,
    ) -> Result<(), BodyImageAdjustmentsError> {
        let observed = current
            .checked_add(amount)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        if observed > maximum {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn scan(
        &mut self,
        payload: &[u8],
        fields: usize,
        depth: usize,
    ) -> Result<(), BodyImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            payload.len(),
            self.max_input,
            BodyImageAdjustmentsLimitKind::InputBytes,
        )?;
        Self::add(
            &mut self.fields,
            fields,
            self.max_fields,
            BodyImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            payload.len().saturating_add(fields),
            self.max_work,
            BodyImageAdjustmentsLimitKind::WireWork,
        )?;
        if depth > self.limits.max_nesting() {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.limits.max_nesting() as u64,
            });
        }
        Ok(())
    }

    fn references(&mut self, amount: usize) -> Result<(), BodyImageAdjustmentsError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            BodyImageAdjustmentsLimitKind::References,
        )
    }

    fn allocations(&mut self, amount: usize) -> Result<(), BodyImageAdjustmentsError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            BodyImageAdjustmentsLimitKind::Allocations,
        )
    }

    fn retained(&mut self, amount: usize) -> Result<(), BodyImageAdjustmentsError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            BodyImageAdjustmentsLimitKind::Retained,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), BodyImageAdjustmentsError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            BodyImageAdjustmentsLimitKind::OutputBytes,
        )
    }

    fn work(&mut self, amount: usize) -> Result<(), BodyImageAdjustmentsError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            BodyImageAdjustmentsLimitKind::WireWork,
        )
    }

    fn parse<'a>(
        &mut self,
        payload: &'a [u8],
        depth: usize,
    ) -> Result<WireView<'a>, BodyImageAdjustmentsError> {
        let view = WireView::parse_with_limits(payload, self.residual_wire_limits()?)
            .map_err(map_wire_error)?;
        self.scan(payload, view.len(), depth)?;
        Ok(view)
    }

    fn residual_wire_limits(&self) -> Result<WireLimits, BodyImageAdjustmentsError> {
        let input = self.max_input.checked_sub(self.input).ok_or(
            BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::InputBytes,
                observed: self.max_input.saturating_add(1) as u64,
                maximum: self.max_input as u64,
            },
        )?;
        if input == 0 {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::InputBytes,
                observed: self.max_input.saturating_add(1) as u64,
                maximum: self.max_input as u64,
            });
        }
        let fields = self.max_fields.checked_sub(self.fields).ok_or(
            BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::WireFields,
                observed: self.max_fields.saturating_add(1) as u64,
                maximum: self.max_fields as u64,
            },
        )?;
        if fields == 0 {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::WireFields,
                observed: self.max_fields.saturating_add(1) as u64,
                maximum: self.max_fields as u64,
            });
        }
        let work = self.max_work.checked_sub(self.work).ok_or(
            BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::WireWork,
                observed: self.max_work.saturating_add(1) as u64,
                maximum: self.max_work as u64,
            },
        )?;
        if work == 0 {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::WireWork,
                observed: self.max_work.saturating_add(1) as u64,
                maximum: self.max_work as u64,
            });
        }
        let output = self.max_output.checked_sub(self.output).ok_or(
            BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::OutputBytes,
                observed: self.max_output.saturating_add(1) as u64,
                maximum: self.max_output as u64,
            },
        )?;
        if output == 0 {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::OutputBytes,
                observed: self.max_output.saturating_add(1) as u64,
                maximum: self.max_output as u64,
            });
        }
        self.limits
            .with_input_bytes(self.limits.max_input_bytes().min(input))
            .and_then(|value| value.with_fields(self.limits.max_fields().min(fields)))
            .and_then(|value| value.with_output_bytes(self.limits.max_output_bytes().min(output)))
            .and_then(|value| value.with_rewrite_work(self.limits.max_rewrite_work().min(work)))
            .and_then(|value| value.with_nesting(self.limits.max_nesting().min(self.max_nesting)))
            .map_err(map_wire_error)
    }

    fn codec_options(
        &self,
        source: &[u8],
    ) -> Result<image_adjustments_codec::DecodeOptions, BodyImageAdjustmentsError> {
        let limits = self.residual_wire_limits()?;
        let output = self.max_output.checked_sub(self.output).ok_or(
            BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::OutputBytes,
                observed: self.max_output.saturating_add(1) as u64,
                maximum: self.max_output as u64,
            },
        )?;
        if output == 0 {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::OutputBytes,
                observed: self.max_output.saturating_add(1) as u64,
                maximum: self.max_output as u64,
            });
        }
        let recursion = u32::try_from(limits.max_nesting())
            .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
        Ok(image_adjustments_codec::DecodeOptions::new(
            limits.max_input_bytes().min(source.len().max(1)),
            limits.max_fields(),
            limits.max_rewrite_work(),
            recursion,
        )
        .with_max_output_bytes(output))
    }

    fn charge_rewrite_requirements(
        &mut self,
        requirements: image_adjustments_codec::RewriteExecutionRequirements,
    ) -> Result<(), BodyImageAdjustmentsError> {
        self.output(requirements.output_bytes)?;
        Self::add(
            &mut self.fields,
            requirements.fields,
            self.max_fields,
            BodyImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            requirements.work_bytes,
            self.max_work,
            BodyImageAdjustmentsLimitKind::WireWork,
        )?;
        if usize::try_from(requirements.max_depth).unwrap_or(usize::MAX) > self.max_nesting {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::WireNesting,
                observed: u64::from(requirements.max_depth),
                maximum: self.max_nesting as u64,
            });
        }
        self.allocations(requirements.allocations)?;
        self.retained(requirements.retained_bytes)?;
        self.retained(requirements.scratch_bytes)
    }

    fn preflight(
        current: usize,
        amount: usize,
        maximum: usize,
        kind: BodyImageAdjustmentsLimitKind,
    ) -> Result<(), BodyImageAdjustmentsError> {
        let observed = current
            .checked_add(amount)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        if observed > maximum {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        Ok(())
    }

    fn preflight_output(&self, amount: usize) -> Result<(), BodyImageAdjustmentsError> {
        Self::preflight(
            self.output,
            amount,
            self.max_output,
            BodyImageAdjustmentsLimitKind::OutputBytes,
        )
    }

    fn preflight_work(&self, amount: usize) -> Result<(), BodyImageAdjustmentsError> {
        Self::preflight(
            self.work,
            amount,
            self.max_work,
            BodyImageAdjustmentsLimitKind::WireWork,
        )
    }

    fn preflight_allocations(&self, amount: usize) -> Result<(), BodyImageAdjustmentsError> {
        Self::preflight(
            self.allocations,
            amount,
            self.max_allocations,
            BodyImageAdjustmentsLimitKind::Allocations,
        )
    }

    fn preflight_retained(&self, amount: usize) -> Result<(), BodyImageAdjustmentsError> {
        Self::preflight(
            self.retained,
            amount,
            self.max_retained,
            BodyImageAdjustmentsLimitKind::Retained,
        )
    }

    fn candidate_reopen(
        &mut self,
        package: &Package,
        bytes: usize,
    ) -> Result<(), BodyImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            bytes,
            self.max_input,
            BodyImageAdjustmentsLimitKind::InputBytes,
        )?;
        Self::add(
            &mut self.work,
            bytes,
            self.max_work,
            BodyImageAdjustmentsLimitKind::WireWork,
        )?;
        let (allocations, retained) = package_reopen_requirements(package, bytes)?;
        self.preflight_allocations(allocations)?;
        self.preflight_retained(retained)?;
        self.allocations(allocations)?;
        self.retained(retained)
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), BodyImageAdjustmentsError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.retained(requirements.scratch_bytes())
    }

    fn charge_drawable_order_report(
        &mut self,
        report: pages_drawable_order_codec::DecodeReport,
    ) -> Result<(), BodyImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            report.input_bytes(),
            self.max_input,
            BodyImageAdjustmentsLimitKind::InputBytes,
        )?;
        Self::add(
            &mut self.fields,
            report.fields(),
            self.max_fields,
            BodyImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            report.work_bytes(),
            self.max_work,
            BodyImageAdjustmentsLimitKind::WireWork,
        )?;
        self.references(report.references())?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        if usize::try_from(report.max_depth()).unwrap_or(usize::MAX) > self.limits.max_nesting() {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::WireNesting,
                observed: u64::from(report.max_depth()),
                maximum: self.limits.max_nesting() as u64,
            });
        }
        Ok(())
    }

    fn charge_image_report(
        &mut self,
        report: image_adjustments_codec::DecodeReport,
    ) -> Result<(), BodyImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            report.input_bytes(),
            self.max_input,
            BodyImageAdjustmentsLimitKind::InputBytes,
        )?;
        Self::add(
            &mut self.fields,
            report.fields(),
            self.max_fields,
            BodyImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            report.work_bytes(),
            self.max_work,
            BodyImageAdjustmentsLimitKind::WireWork,
        )?;
        Self::add(
            &mut self.allocations,
            report.allocations(),
            self.max_allocations,
            BodyImageAdjustmentsLimitKind::Allocations,
        )?;
        self.retained(report.retained_bytes())?;
        self.retained(report.scratch_bytes())?;
        if usize::try_from(report.max_depth()).unwrap_or(usize::MAX) > self.limits.max_nesting() {
            return Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::WireNesting,
                observed: u64::from(report.max_depth()),
                maximum: self.limits.max_nesting() as u64,
            });
        }
        Ok(())
    }
}

impl Package {
    /// Read one rooted body image's adjustments by source-order selector.
    pub fn body_image_adjustments(
        &self,
        selector: impl Into<ImageSelector>,
    ) -> Result<ImageAdjustments, BodyImageAdjustmentsError> {
        let mut budget = ImageBudget::new(self)?;
        Ok(resolve_image(self, selector.into(), &mut budget)?.1)
    }

    /// Begin a selector-first immutable body-image adjustment edit.
    pub fn edit_body_image_adjustments(
        &self,
        selector: impl Into<ImageSelector>,
    ) -> Result<BodyImageAdjustmentsEdit<'_>, BodyImageAdjustmentsError> {
        let mut budget = ImageBudget::new(self)?;
        let (target, before) = resolve_image(self, selector.into(), &mut budget)?;
        Ok(BodyImageAdjustmentsEdit {
            source: self,
            target,
            before,
            after: before,
        })
    }

    /// Apply an exact-source checked reversible body-image adjustment patch.
    pub fn apply_body_image_adjustments(
        &self,
        patch: &BodyImageAdjustmentsPatch,
    ) -> Result<BodyImageAdjustmentsCommit, BodyImageAdjustmentsError> {
        let source = self.state.source.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(BodyImageAdjustmentsError::PatchConflict);
        }
        let mut budget = ImageBudget::new(self)?;
        let (current, before) = resolve_image(
            self,
            ImageSelector::position(patch.target.position),
            &mut budget,
        )?;
        if current != patch.target || before != patch.before {
            return Err(BodyImageAdjustmentsError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(BodyImageAdjustmentsCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyImageAdjustmentsDiagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact() {
            return Err(BodyImageAdjustmentsError::PatchConflict);
        }
        let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
            patch.artifacts.target(),
            self.state.source.limits(),
        )
        .map_err(map_archive_error)?;
        let candidate = Package::from_source_catalog(candidate_source)
            .map_err(|_| BodyImageAdjustmentsError::Verification)?;
        candidate
            .validate()
            .map_err(|_| BodyImageAdjustmentsError::Verification)?;
        let (candidate_target, candidate_after) = resolve_image(
            &candidate,
            ImageSelector::position(patch.target.position),
            &mut budget,
        )?;
        if candidate_target != patch.target || candidate_after != patch.after {
            return Err(BodyImageAdjustmentsError::Verification);
        }
        verify_locality(self, &candidate, &patch.target, &mut budget)?;
        Ok(BodyImageAdjustmentsCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyImageAdjustmentsDiagnostics::published(),
        })
    }
}

fn commit_edit(
    edit: BodyImageAdjustmentsEdit<'_>,
) -> Result<BodyImageAdjustmentsCommit, BodyImageAdjustmentsError> {
    let source = edit.source;
    let source_bytes = source.state.source.shared_source();
    if edit.before == edit.after {
        return Ok(BodyImageAdjustmentsCommit {
            package: source.snapshot(),
            patch: BodyImageAdjustmentsPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                target: edit.target,
                before: edit.before,
                after: edit.after,
                touched_components: 0,
            },
            diagnostics: BodyImageAdjustmentsDiagnostics::unchanged(),
        });
    }
    if !source.state.source.source_is_exact() {
        return Err(BodyImageAdjustmentsError::UnsupportedSource);
    }
    let mut budget = ImageBudget::new(source)?;
    let (current, before) = resolve_image(
        source,
        ImageSelector::position(edit.target.position),
        &mut budget,
    )?;
    if current != edit.target || before != edit.before {
        return Err(BodyImageAdjustmentsError::PatchConflict);
    }
    let candidate = rewrite_image(source, &edit.target, edit.after, &mut budget)?;
    candidate
        .validate()
        .map_err(|_| BodyImageAdjustmentsError::Verification)?;
    let (candidate_target, candidate_after) = resolve_image(
        &candidate,
        ImageSelector::position(edit.target.position),
        &mut budget,
    )?;
    if candidate_target != edit.target || candidate_after != edit.after {
        return Err(BodyImageAdjustmentsError::Verification);
    }
    verify_locality(source, &candidate, &edit.target, &mut budget)?;
    let target_bytes = candidate.state.source.shared_source();
    Ok(BodyImageAdjustmentsCommit {
        package: candidate,
        patch: BodyImageAdjustmentsPatch {
            artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target_bytes)),
            target: edit.target,
            before: edit.before,
            after: edit.after,
            touched_components: 1,
        },
        diagnostics: BodyImageAdjustmentsDiagnostics::published(),
    })
}

fn resolve_image(
    package: &Package,
    selector: ImageSelector,
    budget: &mut ImageBudget,
) -> Result<(ImageTarget, ImageAdjustments), BodyImageAdjustmentsError> {
    let requested = selector.as_position();
    let (body_identifier, body_payload, drawable_order_identifier) =
        body_storage_payload(package, budget)?;
    let body_view = budget.parse(body_payload, 1)?;
    let table = unique_field(&body_view, BODY_ATTACHMENTS_FIELD, 2)?;
    let Some(table) = table else {
        return Err(BodyImageAdjustmentsError::ImageNotFound {
            position: requested,
        });
    };
    let table_view = budget.parse(table.payload(), 2)?;
    let mut entries = Vec::new();
    budget.allocations(table_view.len())?;
    entries
        .try_reserve(table_view.len())
        .map_err(|_| BodyImageAdjustmentsError::Allocation {
            amount: table_view.len(),
        })?;
    for field in table_view
        .fields()
        .filter(|field| field.number() == TABLE_ENTRIES_FIELD)
    {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let entry = parse_body_entry(field.payload(), budget)?;
        entries.push(entry);
    }
    entries.sort_unstable_by_key(|entry| entry.character_index);
    if entries
        .windows(2)
        .any(|window| window[0].character_index == window[1].character_index)
    {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    validate_body_anchors(body_payload, &entries, budget)?;

    let mut seen_attachments = Vec::new();
    let mut seen_drawables = Vec::new();
    let mut images_seen = 0usize;
    let mut found = None;
    budget.allocations(entries.len().saturating_mul(2))?;
    seen_attachments
        .try_reserve_exact(entries.len())
        .map_err(|_| BodyImageAdjustmentsError::Allocation {
            amount: entries.len(),
        })?;
    seen_drawables
        .try_reserve_exact(entries.len())
        .map_err(|_| BodyImageAdjustmentsError::Allocation {
            amount: entries.len(),
        })?;
    for entry in entries {
        let attachment = locate_unique_object(package, entry.attachment_identifier, budget)?;
        if attachment.object.archive_info.identifier == Some(ROOT_OBJECT_IDENTIFIER)
            || attachment.object.archive_info.identifier == Some(body_identifier.get())
            || !push_unique(&mut seen_attachments, entry.attachment_identifier)
        {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
        let Some((attachment_index, attachment_message)) =
            unique_optional_message(attachment.object, ATTACHMENT_MESSAGE_TYPE)?
        else {
            // Body storage can contain other drawable attachment kinds (for
            // example movies and inline text). They are part of the rooted
            // storage graph but are not ordinary image candidates.
            continue;
        };
        validate_message_metadata(attachment.object, attachment_index)?;
        let drawable_identifier =
            parse_attachment_drawable(attachment_message.data.as_slice(), budget)?;
        if !object_metadata_is_owned(
            attachment.object,
            attachment_index,
            drawable_identifier,
            &[ATTACHMENT_DRAWABLE_FIELD],
            true,
        ) {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
        let drawable = locate_unique_object(package, drawable_identifier, budget)?;
        if drawable.component_index != attachment.component_index
            || drawable.object.archive_info.identifier == Some(ROOT_OBJECT_IDENTIFIER)
            || drawable.object.archive_info.identifier == Some(body_identifier.get())
            || !push_unique(&mut seen_drawables, drawable_identifier)
        {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
        let Some((image_index, image_message)) =
            unique_optional_message(drawable.object, IMAGE_MESSAGE_TYPE)?
        else {
            continue;
        };
        validate_message_metadata(drawable.object, image_index)?;
        let (parent, data_identifier, private_references) =
            parse_image_metadata(image_message.data.as_slice(), budget)?;
        if parent != body_identifier
            || !image_data_metadata_is_owned(drawable.object, image_index, data_identifier)
            || !object_metadata_is_owned(
                drawable.object,
                image_index,
                parent,
                &[IMAGE_DRAWABLE_FIELD, DRAWABLE_PARENT_FIELD],
                false,
            )
        {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
        validate_private_graph(
            package,
            &attachment,
            &drawable,
            &private_references,
            body_identifier,
            budget,
        )?;
        if images_seen == requested.get() {
            validate_drawable_order(
                package,
                drawable_order_identifier,
                drawable_identifier,
                budget,
            )?;
            budget.references(2)?;
            let before = decode_adjustments(image_message.data.as_slice(), budget)?;
            let component_name = package
                .state
                .source
                .components()
                .get_index(drawable.component_index)
                .ok_or(BodyImageAdjustmentsError::InvalidSource)?
                .name();
            found = Some((
                ImageTarget {
                    position: requested,
                    body_identifier,
                    attachment_identifier: entry.attachment_identifier,
                    drawable_identifier,
                    data_identifier,
                    character_index: entry.character_index,
                    message_index: image_index,
                    component_name: Arc::from(component_name),
                },
                before,
            ));
        }
        images_seen = images_seen
            .checked_add(1)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    }
    found.ok_or(BodyImageAdjustmentsError::ImageNotFound {
        position: requested,
    })
}

fn body_storage_payload<'a>(
    package: &'a Package,
    budget: &mut ImageBudget,
) -> Result<(NonZeroU64, &'a [u8], NonZeroU64), BodyImageAdjustmentsError> {
    let document = package
        .state
        .source
        .components()
        .get(DOCUMENT_COMPONENT)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let root = document
        .archive()
        .object(ROOT_OBJECT_IDENTIFIER)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let (_, root_message) = unique_message(root, ROOT_MESSAGE_TYPE)?;
    let root_payload = root_message.data.as_slice();
    let body_options = pages_body_options(budget, root_payload)?;
    budget.scan(root_payload, root_payload.len(), 1)?;
    let body_facts = pages_body_codec::decode_document_body(root_payload, body_options)
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    let root_options = pages_body_options(budget, root_payload)?;
    budget.scan(root_payload, root_payload.len(), 1)?;
    let root_facts = pages_body_codec::decode_document_root(root_payload, root_options)
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    let body_identifier = body_facts
        .body_storage()
        .map(|reference| reference.identifier())
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let drawable_order_identifier = root_facts
        .drawables_zorder()
        .map(|reference| reference.identifier())
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let body = locate_unique_object(package, body_identifier, budget)?;
    let (_, payload) = unique_text_message(body.object, body_identifier)?;
    Ok((body_identifier, payload, drawable_order_identifier))
}

fn pages_body_options(
    budget: &ImageBudget,
    source: &[u8],
) -> Result<pages_body_codec::DecodeOptions, BodyImageAdjustmentsError> {
    let limits = budget.residual_wire_limits()?;
    let recursion_limit = u32::try_from(limits.max_nesting())
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    Ok(pages_body_codec::DecodeOptions::new(
        limits.max_input_bytes().min(source.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion_limit,
    ))
}

#[derive(Debug, Clone, Copy)]
struct BodyEntry {
    character_index: u32,
    attachment_identifier: NonZeroU64,
}

fn parse_body_entry(
    source: &[u8],
    budget: &mut ImageBudget,
) -> Result<BodyEntry, BodyImageAdjustmentsError> {
    let view = budget.parse(source, 3)?;
    let character = unique_field(&view, ENTRY_CHARACTER_INDEX_FIELD, 0)?
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let (character, width) = decode_varint_from_bytes(character.payload())
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    if width != encoded_len(character) || character > u64::from(u32::MAX) {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    let object = unique_field(&view, ENTRY_OBJECT_FIELD, 2)?
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    Ok(BodyEntry {
        character_index: character as u32,
        attachment_identifier: parse_object_reference(object.payload(), budget)?,
    })
}

fn validate_body_anchors(
    payload: &[u8],
    entries: &[BodyEntry],
    budget: &mut ImageBudget,
) -> Result<(), BodyImageAdjustmentsError> {
    // This is a second traversal of the body payload.  Charge it against the
    // aggregate ledger as the UTF-16 pass performs real wire parsing and
    // retains a bounded span view for the duration of the check.
    let view = budget.parse(payload, 1)?;
    let mut next = 0usize;
    let mut utf16_index = 0usize;
    for field in view
        .fields()
        .filter(|field| field.number() == BODY_TEXT_FIELD)
    {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let text = std::str::from_utf8(field.payload())
            .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
        for character in text.chars() {
            if entries
                .get(next)
                .is_some_and(|entry| entry.character_index as usize == utf16_index)
            {
                if character != '\u{fffc}' {
                    return Err(BodyImageAdjustmentsError::InvalidSource);
                }
                next = next
                    .checked_add(1)
                    .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
            }
            utf16_index = utf16_index
                .checked_add(character.len_utf16())
                .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        }
    }
    if next != entries.len() {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    budget.work(utf16_index)
}

fn parse_attachment_drawable(
    source: &[u8],
    budget: &mut ImageBudget,
) -> Result<NonZeroU64, BodyImageAdjustmentsError> {
    let view = budget.parse(source, 3)?;
    let field = unique_field(&view, ATTACHMENT_DRAWABLE_FIELD, 2)?
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    parse_object_reference(field.payload(), budget)
}

fn parse_image_metadata(
    source: &[u8],
    budget: &mut ImageBudget,
) -> Result<(NonZeroU64, NonZeroU64, Vec<NonZeroU64>), BodyImageAdjustmentsError> {
    let image = budget.parse(source, 1)?;
    let drawable = unique_field(&image, IMAGE_DRAWABLE_FIELD, 2)?
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let drawable = budget.parse(drawable.payload(), 2)?;
    let parent = unique_field(&drawable, DRAWABLE_PARENT_FIELD, 2)?
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let data = unique_field(&image, IMAGE_DATA_FIELD, 2)?
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let parent = parse_object_reference(parent.payload(), budget)?;
    let data = parse_data_reference(data.payload(), budget)?;
    let mut private_references = Vec::new();
    budget.allocations(1)?;
    private_references
        .try_reserve(2)
        .map_err(|_| BodyImageAdjustmentsError::Allocation { amount: 2 })?;
    for number in [DRAWABLE_TITLE_FIELD, DRAWABLE_CAPTION_FIELD] {
        let Some(field) = unique_field(&drawable, number, 2)? else {
            continue;
        };
        private_references.push(parse_object_reference(field.payload(), budget)?);
    }
    Ok((parent, data, private_references))
}

fn parse_object_reference(
    source: &[u8],
    budget: &mut ImageBudget,
) -> Result<NonZeroU64, BodyImageAdjustmentsError> {
    let view = budget.parse(source, 4)?;
    let field = unique_field(&view, 1, 0)?.ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let (identifier, width) = decode_varint_from_bytes(field.payload())
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    if width != encoded_len(identifier) {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    if let Some(field) = unique_field(&view, 2, 0)? {
        let (value, consumed) = decode_varint_from_bytes(field.payload())
            .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
        if consumed != encoded_len(value) {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
    }
    if let Some(field) = unique_field(&view, 3, 0)? {
        let (value, consumed) = decode_varint_from_bytes(field.payload())
            .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
        if consumed != encoded_len(value) || value != 0 {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
    }
    NonZeroU64::new(identifier).ok_or(BodyImageAdjustmentsError::InvalidSource)
}

fn parse_data_reference(
    source: &[u8],
    budget: &mut ImageBudget,
) -> Result<NonZeroU64, BodyImageAdjustmentsError> {
    let view = budget.parse(source, 4)?;
    let field = unique_field(&view, 1, 0)?.ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let (identifier, width) = decode_varint_from_bytes(field.payload())
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    if width != encoded_len(identifier) {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    // The complete DataReference payload remains source-authoritative. Only
    // its required identifier participates in ownership proof; future fields
    // are opaque and preserved by the selected ImageArchive rewrite.
    NonZeroU64::new(identifier).ok_or(BodyImageAdjustmentsError::InvalidSource)
}

fn unique_field<'a>(
    view: &WireView<'a>,
    number: u32,
    wire_type: u8,
) -> Result<Option<litchi_iwa_common::wire::WireFieldView<'a>>, BodyImageAdjustmentsError> {
    let mut found = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != wire_type {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if found.replace(field).is_some() {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
    }
    Ok(found)
}

fn locate_unique_object<'a>(
    package: &'a Package,
    identifier: NonZeroU64,
    budget: &mut ImageBudget,
) -> Result<LocatedObject<'a>, BodyImageAdjustmentsError> {
    // The source catalog does not expose an indexed object lookup. Charge the
    // complete archive-wide scan before borrowing any match so hostile
    // identifiers cannot turn repeated graph checks into unbounded work.
    budget.work(package.state.object_count)?;
    let mut found = None;
    for (component_index, component) in package.state.source.components().iter().enumerate() {
        for object in &component.archive().objects {
            if object.archive_info.identifier == Some(identifier.get()) {
                if found.is_some() {
                    return Err(BodyImageAdjustmentsError::InvalidSource);
                }
                found = Some(LocatedObject {
                    component_index,
                    object,
                });
            }
        }
    }
    found.ok_or(BodyImageAdjustmentsError::InvalidSource)
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &RawMessage), BodyImageAdjustmentsError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        if info.type_ != message.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
        if message.type_ == message_type && selected.replace((index, message)).is_some() {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
    }
    selected.ok_or(BodyImageAdjustmentsError::InvalidSource)
}

fn unique_optional_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<Option<(usize, &RawMessage)>, BodyImageAdjustmentsError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        if info.type_ != message.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
        if message.type_ == message_type && selected.replace((index, message)).is_some() {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
    }
    Ok(selected)
}

fn unique_text_message(
    object: &ArchiveObject,
    _identifier: NonZeroU64,
) -> Result<(usize, &[u8]), BodyImageAdjustmentsError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if matches!(message.type_, 2_001 | 2_022)
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
    }
    selected.ok_or(BodyImageAdjustmentsError::InvalidSource)
}

fn validate_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), BodyImageAdjustmentsError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    Ok(())
}

fn image_data_metadata_is_owned(
    object: &ArchiveObject,
    message_index: usize,
    identifier: NonZeroU64,
) -> bool {
    object
        .archive_info
        .message_infos
        .get(message_index)
        .is_some_and(|info| {
            let aggregate = info
                .data_references
                .iter()
                .filter(|candidate| **candidate == identifier.get())
                .count();
            if aggregate != 1 {
                return false;
            }
            let mut field = 0usize;
            for field_info in &info.field_infos {
                let occurrences = field_info
                    .data_references
                    .iter()
                    .filter(|candidate| **candidate == identifier.get())
                    .count();
                if occurrences == 0 {
                    continue;
                }
                if field_info.path.as_slice() != [IMAGE_DATA_FIELD] || occurrences != 1 {
                    return false;
                }
                field = field.saturating_add(occurrences);
            }
            field <= 1
        })
}

fn object_metadata_is_owned(
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
            {
                return false;
            }
            if aggregate > 1 {
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

fn validate_private_graph(
    package: &Package,
    attachment: &LocatedObject<'_>,
    drawable: &LocatedObject<'_>,
    private_references: &[NonZeroU64],
    body_identifier: NonZeroU64,
    budget: &mut ImageBudget,
) -> Result<(), BodyImageAdjustmentsError> {
    if attachment.component_index != drawable.component_index {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    let mut identifiers = Vec::new();
    budget.allocations(private_references.len().saturating_add(3))?;
    identifiers
        .try_reserve(private_references.len().saturating_add(3))
        .map_err(|_| BodyImageAdjustmentsError::Allocation {
            amount: private_references.len().saturating_add(3),
        })?;
    for identifier in [
        NonZeroU64::new(ROOT_OBJECT_IDENTIFIER).ok_or(BodyImageAdjustmentsError::InvalidSource)?,
        body_identifier,
        attachment
            .object
            .archive_info
            .identifier
            .and_then(NonZeroU64::new)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?,
        drawable
            .object
            .archive_info
            .identifier
            .and_then(NonZeroU64::new)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?,
    ] {
        if !push_unique(&mut identifiers, identifier) {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
    }
    for identifier in private_references {
        let object = locate_unique_object(package, *identifier, budget)?;
        if object.component_index != drawable.component_index
            || !push_unique(&mut identifiers, *identifier)
        {
            return Err(BodyImageAdjustmentsError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_drawable_order(
    package: &Package,
    identifier: NonZeroU64,
    drawable_identifier: NonZeroU64,
    budget: &mut ImageBudget,
) -> Result<(), BodyImageAdjustmentsError> {
    let order = locate_unique_object(package, identifier, budget)?;
    let (message_index, message) = unique_message(order.object, DRAWABLE_ORDER_MESSAGE_TYPE)?;
    validate_message_metadata(order.object, message_index)?;
    let source = message.data.as_slice();
    let options = drawable_order_options(
        source,
        budget.residual_wire_limits()?,
        budget.max_references,
    );
    let (snapshot, report) =
        pages_drawable_order_codec::decode_drawable_order_with_report(source, options)
            .map_err(map_drawable_order_error)?;
    budget.charge_drawable_order_report(report)?;
    let count = snapshot
        .identifiers()
        .filter(|candidate| *candidate == drawable_identifier.get())
        .count();
    if count != 1 {
        return Err(BodyImageAdjustmentsError::InvalidSource);
    }
    Ok(())
}

fn drawable_order_options(
    source: &[u8],
    limits: WireLimits,
    maximum_references: usize,
) -> pages_drawable_order_codec::DecodeOptions {
    let source_len = source.len().max(1);
    let max_fields = source_len.saturating_mul(8).clamp(1, limits.max_fields());
    let max_work = source_len
        .saturating_mul(16)
        .clamp(1, limits.max_rewrite_work());
    let max_references = source_len.min(maximum_references.max(1)).max(1);
    let recursion_limit = u32::try_from(limits.max_nesting().min(64)).unwrap_or(64);
    pages_drawable_order_codec::DecodeOptions::new(
        source_len.min(limits.max_input_bytes()),
        source_len.min(limits.max_output_bytes()),
        max_fields,
        max_work,
        recursion_limit,
        max_references,
    )
}

fn map_drawable_order_error(
    error: pages_drawable_order_codec::DecodeError,
) -> BodyImageAdjustmentsError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            pages_drawable_order_codec::WireResourceLimit::InputBytes { observed, maximum } => {
                (BodyImageAdjustmentsLimitKind::InputBytes, observed, maximum)
            },
            pages_drawable_order_codec::WireResourceLimit::OutputBytes { observed, maximum } => (
                BodyImageAdjustmentsLimitKind::OutputBytes,
                observed,
                maximum,
            ),
            pages_drawable_order_codec::WireResourceLimit::Fields { observed, maximum } => {
                (BodyImageAdjustmentsLimitKind::WireFields, observed, maximum)
            },
            pages_drawable_order_codec::WireResourceLimit::WorkBytes { observed, maximum } => {
                (BodyImageAdjustmentsLimitKind::WireWork, observed, maximum)
            },
            pages_drawable_order_codec::WireResourceLimit::Nesting { observed, maximum } => (
                BodyImageAdjustmentsLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            pages_drawable_order_codec::WireResourceLimit::References { observed, maximum } => {
                (BodyImageAdjustmentsLimitKind::References, observed, maximum)
            },
            _ => (BodyImageAdjustmentsLimitKind::WireWork, 0, 0),
        };
        return BodyImageAdjustmentsError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return BodyImageAdjustmentsError::Allocation { amount };
    }
    BodyImageAdjustmentsError::InvalidSource
}

fn decode_adjustments(
    source: &[u8],
    budget: &mut ImageBudget,
) -> Result<ImageAdjustments, BodyImageAdjustmentsError> {
    let options = budget.codec_options(source)?;
    let (snapshot, report) =
        image_adjustments_codec::decode_image_adjustments_with_report(source, options)
            .map_err(map_codec_error)?;
    budget.charge_image_report(report)?;
    adjustments_from_snapshot(snapshot).map_err(|_| BodyImageAdjustmentsError::InvalidSource)
}

fn rewrite_adjustments(
    source: &[u8],
    adjustments: ImageAdjustments,
    budget: &mut ImageBudget,
) -> Result<Vec<u8>, BodyImageAdjustmentsError> {
    let options = budget.codec_options(source)?;
    let prepared = image_adjustments_codec::prepare_image_adjustments_rewrite(
        source,
        image_adjustments_write(adjustments),
        options,
    )
    .map_err(map_codec_error)?;
    budget.charge_image_report(prepared.prepare_report())?;
    let requirements = prepared.execution_requirements();
    budget.charge_rewrite_requirements(requirements)?;
    prepared
        .execute(requirements.exact())
        .map(|output| output.into_output())
        .map_err(map_codec_error)
}

#[cfg(feature = "internal-iwork-source")]
fn codec_options(source: &[u8], limits: WireLimits) -> image_adjustments_codec::DecodeOptions {
    image_adjustments_codec::DecodeOptions::new(
        limits.max_input_bytes().min(source.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_output_bytes(limits.max_output_bytes())
}

fn image_adjustments_write(
    adjustments: ImageAdjustments,
) -> image_adjustments_codec::ImageAdjustmentsWrite {
    image_adjustments_codec::ImageAdjustmentsWrite::from_values(
        adjustments.exposure().map(ImageAdjustment::value),
        adjustments.saturation().map(ImageAdjustment::value),
        adjustments
            .enhancement()
            .map(|value| matches!(value, ImageEnhancement::Enabled)),
    )
}

fn adjustments_from_snapshot(
    snapshot: image_adjustments_codec::ImageAdjustmentsSnapshot<'_>,
) -> Result<ImageAdjustments, litchi_iwa_common::shape::image::Error> {
    Ok(ImageAdjustments::new()
        .with_exposure(snapshot.exposure().map(ImageAdjustment::new).transpose()?)
        .with_saturation(
            snapshot
                .saturation()
                .map(ImageAdjustment::new)
                .transpose()?,
        )
        .with_enhancement(snapshot.enhance().map(|value| {
            if value {
                ImageEnhancement::Enabled
            } else {
                ImageEnhancement::Disabled
            }
        })))
}

fn rewrite_image(
    source: &Package,
    target: &ImageTarget,
    after: ImageAdjustments,
    budget: &mut ImageBudget,
) -> Result<Package, BodyImageAdjustmentsError> {
    let component = source
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == target.component_name.as_ref())
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let object = component
        .archive()
        .object(target.drawable_identifier.get())
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let original = object
        .messages
        .get(target.message_index)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?
        .data
        .as_slice();
    let source_archive = component.archive();
    let rewritten = rewrite_adjustments(original, after, budget)?;
    let encoded_len = source_archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let replacement_bound = rewritten_archive_bound(encoded_len, original.len(), rewritten.len())?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(replacement_bound).map_err(map_core_error)?;
    let snappy_limits = source
        .state
        .source
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(BodyImageAdjustmentsError::LimitExceeded {
            kind: BodyImageAdjustmentsLimitKind::OutputBytes,
            observed: compressed_bound as u64,
            maximum: snappy_limits.max_compressed_stream() as u64,
        });
    }
    let entry = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    let package_bound = source
        .state
        .source
        .source_bytes()
        .len()
        .checked_sub(entry.data().len())
        .and_then(|value| value.checked_add(compressed_bound))
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    budget.preflight_output(replacement_bound)?;
    budget.preflight_output(package_bound)?;
    budget.preflight_work(
        replacement_bound
            .checked_add(compressed_bound)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?,
    )?;
    let (clone_allocations, clone_retained) =
        archive_clone_requirements(source_archive, encoded_len)?;
    let (header_allocations, header_retained) = archive_header_rewrite_requirements(object)?;
    let (serialization_allocations, serialization_retained) =
        archive_serialization_requirements(source_archive)?;
    let total_allocations = clone_allocations
        .checked_add(header_allocations)
        .and_then(|amount| amount.checked_add(serialization_allocations))
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    budget.preflight_allocations(total_allocations)?;
    let temporary_retained = clone_retained
        .checked_add(header_retained)
        .and_then(|amount| amount.checked_add(serialization_retained))
        .and_then(|amount| amount.checked_add(replacement_bound))
        .and_then(|amount| amount.checked_add(compressed_bound))
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    budget.preflight_retained(temporary_retained)?;
    budget.allocations(total_allocations)?;
    budget.retained(clone_retained)?;
    budget.retained(header_retained)?;
    budget.retained(serialization_retained)?;
    let mut archive = source_archive.clone();
    archive
        .object_mut(target.drawable_identifier.get())
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            target.message_index,
            RawMessage {
                type_: IMAGE_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.output(bytes.len())?;
    budget.retained(bytes.len())?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    budget.retained(compressed.len())?;
    if compressed.len() as u64 > source.state.source.limits().max_entry_bytes() {
        return Err(BodyImageAdjustmentsError::LimitExceeded {
            kind: BodyImageAdjustmentsLimitKind::OutputBytes,
            observed: compressed.len() as u64,
            maximum: source.state.source.limits().max_entry_bytes(),
        });
    }
    let edit = EntryEdit::new(component.name(), &compressed);
    let edits = [edit];
    let prepared = source
        .state
        .source
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], source.state.source.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    let max_input = source.state.source.limits().max_input_bytes();
    if requirements.output_bytes() as u64 > max_input {
        return Err(BodyImageAdjustmentsError::LimitExceeded {
            kind: BodyImageAdjustmentsLimitKind::OutputBytes,
            observed: requirements.output_bytes() as u64,
            maximum: max_input,
        });
    }
    budget.reassembly(requirements)?;
    budget.candidate_reopen(source, requirements.output_bytes())?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let catalog =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), source.state.source.limits())
            .map_err(map_archive_error)?;
    Package::from_source_catalog(catalog).map_err(|_| BodyImageAdjustmentsError::Verification)
}

fn rewritten_archive_bound(
    source_encoded_len: usize,
    original_len: usize,
    replacement_len: usize,
) -> Result<usize, BodyImageAdjustmentsError> {
    const MAX_REWRITE_FRAMING_GROWTH: usize = 64;
    source_encoded_len
        .checked_sub(original_len)
        .and_then(|value| value.checked_add(replacement_len))
        .and_then(|value| value.checked_add(MAX_REWRITE_FRAMING_GROWTH))
        .ok_or(BodyImageAdjustmentsError::InvalidSource)
}

fn archive_header_rewrite_requirements(
    object: &ArchiveObject,
) -> Result<(usize, usize), BodyImageAdjustmentsError> {
    let header = usize::try_from(object.header_length)
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    let retained = header
        .checked_add(64)
        .and_then(|value| value.checked_mul(3))
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    // The core replacement primitive builds canonical, rewritten, and
    // post-mutation header buffers before publishing the new message.  The
    // focused owner reserves that bounded staging before entering the
    // infallible clone/mutation path.
    Ok((3, retained))
}

fn archive_serialization_requirements(
    archive: &Archive,
) -> Result<(usize, usize), BodyImageAdjustmentsError> {
    let mut retained = 0usize;
    for object in &archive.objects {
        let header = usize::try_from(object.header_length)
            .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(header)
            .and_then(|value| value.checked_add(64))
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    }
    let allocations = archive
        .objects
        .len()
        .checked_add(2)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    Ok((allocations, retained))
}

fn archive_clone_requirements(
    archive: &Archive,
    encoded_len: usize,
) -> Result<(usize, usize), BodyImageAdjustmentsError> {
    let mut allocations = 0usize;
    let mut retained = encoded_len;
    add_clone_vec::<ArchiveObject>(archive.objects.len(), &mut allocations, &mut retained)?;
    for object in &archive.objects {
        add_clone_vec::<RawMessage>(object.messages.len(), &mut allocations, &mut retained)?;
        add_clone_vec::<litchi_iwa_core::MessageInfo>(
            object.archive_info.message_infos.len(),
            &mut allocations,
            &mut retained,
        )?;
        if object.header_length != 0 {
            allocations = allocations
                .checked_add(2)
                .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
            retained = retained
                .checked_add(
                    usize::try_from(object.header_length)
                        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?
                        .checked_mul(2)
                        .ok_or(BodyImageAdjustmentsError::InvalidSource)?,
                )
                .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
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
    Ok((allocations, retained))
}

fn package_reopen_requirements(
    package: &Package,
    candidate_bytes: usize,
) -> Result<(usize, usize), BodyImageAdjustmentsError> {
    let archive_limits = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let components = package.state.source.components();
    let entries = package.state.source.package().iter();
    let mut allocations = 3usize;
    let mut retained = candidate_bytes;
    let mut object_count = 0usize;
    for entry in entries {
        allocations = allocations
            .checked_add(1)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        let entry_slots = 2usize
            .checked_mul(size_of::<usize>())
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(entry.name().len())
            .and_then(|value| value.checked_add(entry_slots))
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    }
    if !components.is_empty() {
        allocations = allocations
            .checked_add(1)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        let component_slots = components
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_mul(size_of::<usize>()))
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(component_slots)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    }
    for component in components.iter() {
        allocations = allocations
            .checked_add(1)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(component.name().len())
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        object_count = object_count
            .checked_add(component.archive().objects.len())
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        let encoded_len = component
            .archive()
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let (archive_allocations, archive_retained) =
            archive_clone_requirements(component.archive(), encoded_len)?;
        allocations = allocations
            .checked_add(archive_allocations)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(archive_retained)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    }
    if object_count != 0 {
        allocations = allocations
            .checked_add(1)
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(
                object_count
                    .checked_mul(size_of::<(u64, usize, usize)>())
                    .ok_or(BodyImageAdjustmentsError::InvalidSource)?,
            )
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    }
    allocations = allocations
        .checked_add(1)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    retained = retained
        .checked_add(candidate_bytes)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    Ok((allocations, retained))
}

fn add_clone_vec<T>(
    length: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), BodyImageAdjustmentsError> {
    if length == 0 {
        return Ok(());
    }
    *allocations = allocations
        .checked_add(1)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    *retained = retained
        .checked_add(
            length
                .checked_mul(size_of::<T>())
                .ok_or(BodyImageAdjustmentsError::InvalidSource)?,
        )
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    Ok(())
}

fn add_clone_bytes(
    length: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), BodyImageAdjustmentsError> {
    add_clone_vec::<u8>(length, allocations, retained)
}

fn archive_object_clone_requirements(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    replacement_len: usize,
) -> Result<(usize, usize), BodyImageAdjustmentsError> {
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
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(
                usize::try_from(source.header_length)
                    .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?
                    .checked_mul(2)
                    .ok_or(BodyImageAdjustmentsError::InvalidSource)?,
            )
            .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    }
    add_clone_bytes(replacement_len, &mut allocations, &mut retained)?;
    let source_header = usize::try_from(source.header_length)
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    let candidate_header = usize::try_from(candidate.header_length)
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    let header_scratch = source_header
        .checked_add(candidate_header)
        .and_then(|value| value.checked_mul(3))
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    allocations = allocations
        .checked_add(3)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    retained = retained
        .checked_add(header_scratch)
        .ok_or(BodyImageAdjustmentsError::InvalidSource)?;
    Ok((allocations, retained))
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    target: &ImageTarget,
    budget: &mut ImageBudget,
) -> Result<(), BodyImageAdjustmentsError> {
    let source_entries = source.state.source.package();
    let candidate_entries = candidate.state.source.package();
    if source_entries.len() != candidate_entries.len() {
        return Err(BodyImageAdjustmentsError::Verification);
    }
    for (before, after) in source_entries.iter().zip(candidate_entries.iter()) {
        budget.work(
            before
                .name()
                .len()
                .saturating_add(before.data().len())
                .saturating_add(1),
        )?;
        if before.name() != after.name() {
            return Err(BodyImageAdjustmentsError::Verification);
        }
        if before.name() != target.component_name.as_ref() && before.data() != after.data() {
            return Err(BodyImageAdjustmentsError::Verification);
        }
    }
    let source_components = source.state.source.components();
    let candidate_components = candidate.state.source.components();
    if source_components.len() != candidate_components.len() {
        return Err(BodyImageAdjustmentsError::Verification);
    }
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut target_archives = None;
    for (before, after) in source_components.iter().zip(candidate_components.iter()) {
        if before.name() != after.name() {
            return Err(BodyImageAdjustmentsError::Verification);
        }
        if before.name() == target.component_name.as_ref() {
            target_archives = Some((before.archive(), after.archive()));
        } else {
            if before.archive().objects.len() != after.archive().objects.len() {
                return Err(BodyImageAdjustmentsError::Verification);
            }
            for (source_object, candidate_object) in before
                .archive()
                .objects
                .iter()
                .zip(&after.archive().objects)
            {
                budget.work(archive_object_comparison_work(
                    source_object,
                    candidate_object,
                )?)?;
                if !source_object.same_content_ignoring_offsets(candidate_object) {
                    return Err(BodyImageAdjustmentsError::Verification);
                }
            }
        }
    }
    let (before_archive, after_archive) =
        target_archives.ok_or(BodyImageAdjustmentsError::Verification)?;
    if before_archive.objects.len() != after_archive.objects.len() {
        return Err(BodyImageAdjustmentsError::Verification);
    }
    for (before_object, after_object) in before_archive.objects.iter().zip(&after_archive.objects) {
        budget.work(archive_object_comparison_work(before_object, after_object)?)?;
        let identifier = before_object
            .archive_info
            .identifier
            .ok_or(BodyImageAdjustmentsError::Verification)?;
        if after_object.archive_info.identifier != Some(identifier) {
            return Err(BodyImageAdjustmentsError::Verification);
        }
        if identifier != target.drawable_identifier.get() {
            if !before_object.same_content_ignoring_offsets(after_object) {
                return Err(BodyImageAdjustmentsError::Verification);
            }
            continue;
        }
        let candidate_message = after_object
            .messages
            .get(target.message_index)
            .ok_or(BodyImageAdjustmentsError::Verification)?;
        let (allocations, retained) = archive_object_clone_requirements(
            before_object,
            after_object,
            candidate_message.data.len(),
        )?;
        budget.preflight_allocations(allocations)?;
        budget.preflight_retained(retained)?;
        budget.allocations(allocations)?;
        budget.retained(retained)?;
        let mut expected = before_object.clone();
        expected
            .replace_message_preserving_header_with_limits(
                target.message_index,
                RawMessage {
                    type_: candidate_message.type_,
                    data: candidate_message.data.clone(),
                },
                archive_limits,
            )
            .map_err(map_core_error)?;
        expected.header_length = after_object.header_length;
        expected.data_length = after_object.data_length;
        if !expected.same_content_ignoring_offsets(after_object) {
            return Err(BodyImageAdjustmentsError::Verification);
        }
    }
    Ok(())
}

fn archive_object_comparison_work(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
) -> Result<usize, BodyImageAdjustmentsError> {
    let source_header = usize::try_from(source.header_length)
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    let candidate_header = usize::try_from(candidate.header_length)
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    let source_data = usize::try_from(source.data_length)
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    let candidate_data = usize::try_from(candidate.data_length)
        .map_err(|_| BodyImageAdjustmentsError::InvalidSource)?;
    source_header
        .checked_add(candidate_header)
        .and_then(|value| value.checked_add(source_data))
        .and_then(|value| value.checked_add(candidate_data))
        .and_then(|value| value.checked_add(source.messages.len()))
        .and_then(|value| value.checked_add(candidate.messages.len()))
        .and_then(|value| value.checked_add(1))
        .ok_or(BodyImageAdjustmentsError::InvalidSource)
}

fn push_unique(values: &mut Vec<NonZeroU64>, value: NonZeroU64) -> bool {
    if values.contains(&value) {
        false
    } else {
        values.push(value);
        true
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> BodyImageAdjustmentsError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => BodyImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    BodyImageAdjustmentsLimitKind::InputBytes
                },
                litchi_iwa_common::LimitKind::Fields => BodyImageAdjustmentsLimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => {
                    BodyImageAdjustmentsLimitKind::OutputBytes
                },
                litchi_iwa_common::LimitKind::Nesting => BodyImageAdjustmentsLimitKind::WireNesting,
                _ => BodyImageAdjustmentsLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            BodyImageAdjustmentsError::Allocation { amount }
        },
        _ => BodyImageAdjustmentsError::InvalidSource,
    }
}

fn map_codec_error(error: image_adjustments_codec::DecodeError) -> BodyImageAdjustmentsError {
    if let Some(limit) = error.limit_kind() {
        let (observed, maximum) = error.limit_values().unwrap_or((0, 0));
        return BodyImageAdjustmentsError::LimitExceeded {
            kind: match limit {
                image_adjustments_codec::DecodeLimit::InputBytes => {
                    BodyImageAdjustmentsLimitKind::InputBytes
                },
                image_adjustments_codec::DecodeLimit::OutputBytes
                | image_adjustments_codec::DecodeLimit::Retained => {
                    BodyImageAdjustmentsLimitKind::OutputBytes
                },
                image_adjustments_codec::DecodeLimit::Fields => {
                    BodyImageAdjustmentsLimitKind::WireFields
                },
                image_adjustments_codec::DecodeLimit::Nesting => {
                    BodyImageAdjustmentsLimitKind::WireNesting
                },
                _ => BodyImageAdjustmentsLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    BodyImageAdjustmentsError::InvalidSource
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyImageAdjustmentsError {
    match error {
        litchi_iwa_archive::Error::Limit {
            observed,
            maximum,
            kind,
        } => BodyImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    BodyImageAdjustmentsLimitKind::OutputBytes
                },
                _ => BodyImageAdjustmentsLimitKind::InputBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyImageAdjustmentsError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => BodyImageAdjustmentsError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyImageAdjustmentsError {
    match error {
        litchi_iwa_core::Error::Limit {
            observed,
            maximum,
            kind,
        } => BodyImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => {
                    BodyImageAdjustmentsLimitKind::InputBytes
                },
                _ => BodyImageAdjustmentsLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyImageAdjustmentsError::Allocation { amount: requested }
        },
        _ => BodyImageAdjustmentsError::InvalidSource,
    }
}

#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub use image_adjustments_bridge::{
    __decode_image_adjustments_payload, __rewrite_image_adjustments_payload, ImageAdjustmentsError,
};

#[cfg(feature = "internal-iwork-source")]
mod image_adjustments_bridge {
    use super::*;

    /// Typed failures at the hidden Pages ImageArchive adjustment seam.
    #[doc(hidden)]
    #[derive(Debug, Clone, PartialEq, Eq, Error)]
    pub enum ImageAdjustmentsError {
        /// The bounded neutral codec rejected the source or candidate.
        #[error("image adjustments codec rejected the payload: {0}")]
        Codec(#[from] image_adjustments_codec::DecodeError),
        /// The projected value failed the common semantic boundary.
        #[error("invalid image adjustment value: {0}")]
        Semantic(#[from] litchi_iwa_common::shape::image::Error),
    }

    #[doc(hidden)]
    pub fn __decode_image_adjustments_payload(
        source: &[u8],
        limits: WireLimits,
    ) -> Result<ImageAdjustments, ImageAdjustmentsError> {
        let snapshot = image_adjustments_codec::decode_image_adjustments(
            source,
            codec_options(source, limits),
        )?;
        Ok(adjustments_from_snapshot(snapshot)?)
    }

    #[doc(hidden)]
    pub fn __rewrite_image_adjustments_payload(
        source: &[u8],
        adjustments: ImageAdjustments,
        limits: WireLimits,
    ) -> Result<Vec<u8>, ImageAdjustmentsError> {
        image_adjustments_codec::rewrite_image_adjustments(
            source,
            image_adjustments_write(adjustments),
            codec_options(source, limits),
        )
        .map_err(ImageAdjustmentsError::Codec)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn adjustment_source_with_unknowns() -> Vec<u8> {
        vec![
            0x72, 0x0c, // ImageArchive.imageAdjustments
            0x0d, 0x00, 0x00, 0x00, 0x00, // exposure = 0
            0x15, 0x00, 0x00, 0x00, 0x00, // saturation = 0
            0x68, 0x00, // enhance = false
            0x98, 0x06, 0x01, // opaque outer field 99
        ]
    }

    fn rewrite_budget_template(source: &[u8], after: ImageAdjustments) -> ImageBudget {
        let options =
            image_adjustments_codec::DecodeOptions::for_source(source).with_max_output_bytes(1024);
        let prepared = image_adjustments_codec::prepare_image_adjustments_rewrite(
            source,
            image_adjustments_write(after),
            options,
        )
        .unwrap();
        let report = prepared.prepare_report();
        let requirements = prepared.execution_requirements();
        let input = report.input_bytes().max(1);
        let output = requirements.output_bytes.max(1);
        let fields = report.fields().checked_add(requirements.fields).unwrap();
        let work = report
            .work_bytes()
            .checked_add(requirements.work_bytes)
            .unwrap();
        let allocations = report
            .allocations()
            .checked_add(requirements.allocations)
            .unwrap();
        let retained = report
            .retained_bytes()
            .checked_add(report.scratch_bytes())
            .and_then(|value| value.checked_add(requirements.retained_bytes))
            .and_then(|value| value.checked_add(requirements.scratch_bytes))
            .unwrap();
        let nesting = usize::try_from(requirements.max_depth).unwrap().max(1);
        let limits = WireLimits::default()
            .with_input_bytes(input)
            .and_then(|value| value.with_output_bytes(output))
            .and_then(|value| value.with_fields(fields.max(1)))
            .and_then(|value| value.with_rewrite_work(work.max(1)))
            .and_then(|value| value.with_nesting(nesting))
            .unwrap();
        ImageBudget {
            limits,
            max_input: input,
            max_output: output,
            max_fields: fields.max(1),
            max_work: work.max(1),
            max_nesting: nesting,
            max_references: usize::MAX,
            max_allocations: allocations.max(1),
            max_retained: retained.max(1),
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            references: 0,
            allocations: 0,
            retained: 0,
        }
    }

    fn assert_budget_limit(
        result: Result<Vec<u8>, BodyImageAdjustmentsError>,
        expected: BodyImageAdjustmentsLimitKind,
    ) {
        let error = result.expect_err("the one-under budget must reject the rewrite");
        assert!(matches!(
            error,
            BodyImageAdjustmentsError::LimitExceeded { kind, .. } if kind == expected
        ));
    }

    #[test]
    fn object_reference_preserves_unknown_and_canonical_deprecated_fields() {
        let mut budget = ImageBudget::for_test();
        let source = [
            0x08, 0x07, // identifier
            0x10, 0x01, // deprecated type
            0x18, 0x00, // deprecated external=false
            0x20, 0x01, // opaque future field
        ];
        assert_eq!(
            parse_object_reference(&source, &mut budget).unwrap().get(),
            7
        );
    }

    #[test]
    fn data_reference_preserves_opaque_future_fields() {
        let mut budget = ImageBudget::for_test();
        let source = [0x08, 0x18, 0x20, 0x01, 0x2a, 0x01, 0x00];
        assert_eq!(
            parse_data_reference(&source, &mut budget).unwrap().get(),
            24
        );
    }

    #[test]
    fn wire_parse_accepts_inclusive_cost_and_rejects_each_one_under_limit() {
        let source = [0x08, 0x01, 0x10, 0x02];
        let fields = WireView::parse(&source).unwrap().len();
        let work = source.len().checked_add(fields).unwrap();
        let limits = WireLimits::default()
            .with_input_bytes(source.len())
            .and_then(|value| value.with_fields(fields))
            .and_then(|value| value.with_rewrite_work(work))
            .unwrap();
        let template = ImageBudget {
            limits,
            max_input: source.len(),
            max_output: usize::MAX,
            max_fields: fields,
            max_work: work,
            max_nesting: limits.max_nesting(),
            max_references: usize::MAX,
            max_allocations: usize::MAX,
            max_retained: usize::MAX,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            references: 0,
            allocations: 0,
            retained: 0,
        };
        let mut exact = template;
        assert!(exact.parse(&source, 1).is_ok());

        let mut input = template;
        input.max_input -= 1;
        input.limits = input.limits.with_input_bytes(input.max_input).unwrap();
        assert_budget_limit(
            input.parse(&source, 1).map(|_| Vec::new()),
            BodyImageAdjustmentsLimitKind::InputBytes,
        );

        let mut fields_budget = template;
        fields_budget.max_fields -= 1;
        fields_budget.limits = fields_budget
            .limits
            .with_fields(fields_budget.max_fields)
            .unwrap();
        assert_budget_limit(
            fields_budget.parse(&source, 1).map(|_| Vec::new()),
            BodyImageAdjustmentsLimitKind::WireFields,
        );

        let mut work_budget = template;
        work_budget.max_work -= 1;
        work_budget.limits = work_budget
            .limits
            .with_rewrite_work(work_budget.max_work)
            .unwrap();
        assert_budget_limit(
            work_budget.parse(&source, 1).map(|_| Vec::new()),
            BodyImageAdjustmentsLimitKind::WireWork,
        );
    }

    #[test]
    fn prepared_rewrite_charges_unknown_work_and_accepts_only_inclusive_budget() {
        let source = adjustment_source_with_unknowns();
        let after = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        let template = rewrite_budget_template(&source, after);

        let mut exact = template;
        let output = rewrite_adjustments(&source, after, &mut exact).unwrap();
        assert!(output.ends_with(&[0x98, 0x06, 0x01]));

        let mut input = template;
        input.max_input -= 1;
        input.limits = input.limits.with_input_bytes(input.max_input).unwrap();
        assert_budget_limit(
            rewrite_adjustments(&source, after, &mut input),
            BodyImageAdjustmentsLimitKind::InputBytes,
        );

        let mut fields = template;
        fields.max_fields -= 1;
        fields.limits = fields.limits.with_fields(fields.max_fields).unwrap();
        assert_budget_limit(
            rewrite_adjustments(&source, after, &mut fields),
            BodyImageAdjustmentsLimitKind::WireFields,
        );

        let mut work = template;
        work.max_work -= 1;
        work.limits = work.limits.with_rewrite_work(work.max_work).unwrap();
        assert_budget_limit(
            rewrite_adjustments(&source, after, &mut work),
            BodyImageAdjustmentsLimitKind::WireWork,
        );

        let mut output_budget = template;
        output_budget.max_output -= 1;
        output_budget.limits = output_budget
            .limits
            .with_output_bytes(output_budget.max_output)
            .unwrap();
        assert_budget_limit(
            rewrite_adjustments(&source, after, &mut output_budget),
            BodyImageAdjustmentsLimitKind::OutputBytes,
        );

        let mut allocations = template;
        allocations.max_allocations -= 1;
        assert_budget_limit(
            rewrite_adjustments(&source, after, &mut allocations),
            BodyImageAdjustmentsLimitKind::Allocations,
        );

        let mut retained = template;
        retained.max_retained -= 1;
        assert_budget_limit(
            rewrite_adjustments(&source, after, &mut retained),
            BodyImageAdjustmentsLimitKind::Retained,
        );
    }

    #[test]
    fn archive_rewrite_budget_includes_clone_headers_and_serialization_staging() {
        let payload = vec![1, 2, 3, 4];
        let mut object = ArchiveObject::new(
            7,
            vec![RawMessage {
                type_: IMAGE_MESSAGE_TYPE,
                data: payload.clone(),
            }],
        )
        .unwrap();
        object.header_length = 11;
        let archive = Archive {
            objects: vec![object.clone()],
        };
        let (clone_allocations, clone_retained) = archive_clone_requirements(&archive, 37).unwrap();
        let (header_allocations, header_retained) =
            archive_header_rewrite_requirements(&object).unwrap();
        let (serialization_allocations, serialization_retained) =
            archive_serialization_requirements(&archive).unwrap();
        let allocations = clone_allocations
            .checked_add(header_allocations)
            .and_then(|value| value.checked_add(serialization_allocations))
            .unwrap();
        let retained = clone_retained
            .checked_add(header_retained)
            .and_then(|value| value.checked_add(serialization_retained))
            .unwrap();
        assert_eq!(header_allocations, 3);
        assert_eq!(header_retained, (11 + 64) * 3);
        assert!(clone_allocations >= 6);
        assert!(clone_retained >= 37 + size_of::<ArchiveObject>() + payload.len());

        let mut budget = ImageBudget::for_test();
        budget.max_allocations = allocations;
        budget.max_retained = retained;
        assert!(budget.preflight_allocations(allocations).is_ok());
        assert!(budget.preflight_retained(retained).is_ok());
        budget.max_allocations -= 1;
        assert!(matches!(
            budget.preflight_allocations(allocations),
            Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::Allocations,
                ..
            })
        ));
        let mut retained_budget = ImageBudget::for_test();
        retained_budget.max_retained = retained - 1;
        assert!(matches!(
            retained_budget.preflight_retained(retained),
            Err(BodyImageAdjustmentsError::LimitExceeded {
                kind: BodyImageAdjustmentsLimitKind::Retained,
                ..
            })
        ));
    }

    #[test]
    fn parent_metadata_can_be_payload_only_but_attachment_requires_declaration() {
        let mut object = ArchiveObject::new(
            9,
            vec![RawMessage {
                type_: ATTACHMENT_MESSAGE_TYPE,
                data: Vec::new(),
            }],
        )
        .unwrap();
        let identifier = NonZeroU64::new(7).unwrap();
        assert!(object_metadata_is_owned(
            &object,
            0,
            identifier,
            &[IMAGE_DRAWABLE_FIELD, DRAWABLE_PARENT_FIELD],
            false,
        ));
        assert!(!object_metadata_is_owned(
            &object,
            0,
            identifier,
            &[ATTACHMENT_DRAWABLE_FIELD],
            true,
        ));
        object.archive_info.message_infos[0]
            .object_references
            .push(identifier.get());
        assert!(object_metadata_is_owned(
            &object,
            0,
            identifier,
            &[ATTACHMENT_DRAWABLE_FIELD],
            true,
        ));
    }
}
