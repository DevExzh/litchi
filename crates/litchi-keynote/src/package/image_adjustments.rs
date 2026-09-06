//! Selector-first Keynote image-adjustment reads and exact-source edits.
//!
//! The complete `TSD.ImageArchive` payload remains the preservation authority.
//! This owner projects only the three controls exposed by the common image
//! vocabulary and delegates wire validation/rewrite to the neutral bounded
//! Buffa codec. Native object identifiers and generated protobuf values do not
//! cross the supported API.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    reason = "The focused boundary deliberately redacts lower-layer failures."
)]

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::image_adjustments_codec::{
    self as image_adjustments_codec, ImageAdjustmentsSnapshot, ImageAdjustmentsWrite,
};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::slide::image::ImageSelector;
use crate::{SlideSelector, SlideSelectorError};

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const IMAGE_SUPER_FIELD: u32 = 1;
const IMAGE_DATA_FIELD: u32 = 11;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const IMAGE_FLAGS_FIELD: u32 = 7;
const LAYOUT_IMAGE_FLAG: u32 = 1;

/// Resource category reported by a focused image-adjustment operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideImageAdjustmentsLimitKind {
    /// Source or candidate wire bytes.
    WireBytes,
    /// Parsed fields.
    WireFields,
    /// Nested message depth.
    WireNesting,
    /// Aggregate scan or rewrite work.
    WireWork,
    /// Referenced object count.
    References,
    /// Temporary allocations.
    Allocations,
    /// Retained temporary bytes.
    Retained,
    /// Package or component bytes.
    OutputBytes,
}

impl fmt::Display for SlideImageAdjustmentsLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::References => "references",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::OutputBytes => "output bytes",
        })
    }
}

/// Failure raised by a selector-first image-adjustment read or edit.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideImageAdjustmentsError {
    /// The package does not retain an exact physical source suitable for edits.
    #[error("this Keynote source does not support physical image-adjustment edits")]
    UnsupportedSource,
    /// The selected graph is valid but outside this focused scalar owner.
    #[error("the requested Keynote image-adjustment graph operation is unsupported")]
    UnsupportedDependency,
    /// A semantic selector was ambiguous.
    #[error("the Keynote image-adjustment selector is ambiguous")]
    AmbiguousSelector,
    /// A name selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide matched the requested name.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// No slide matched the requested position.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// No image matched the requested source-order position.
    #[error("the selected Keynote slide has no image at position {position:?}")]
    ImagePositionNotFound { position: Position },
    /// The source graph or payload is malformed.
    #[error("the Keynote image-adjustment source is invalid")]
    InvalidSource,
    /// A finite operation budget was exceeded.
    #[error(
        "Keynote image-adjustment {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideImageAdjustmentsLimitKind,
        observed: u64,
        maximum: u64,
    },
    /// A required temporary allocation could not be reserved.
    #[error("could not allocate {amount} units for Keynote image adjustments")]
    Allocation { amount: usize },
    /// The reopened candidate did not retain the requested semantic value.
    #[error("the edited Keynote image adjustments failed semantic verification")]
    Verification,
    /// The patch was created from another exact package artifact.
    #[error("the Keynote image-adjustment patch does not match the exact source package")]
    PatchConflict,
}

/// Typed failures at the hidden host migration seam.
#[doc(hidden)]
#[cfg(feature = "internal-iwork-source")]
#[derive(Debug, Clone, PartialEq, Eq, Error)]
pub enum ImageAdjustmentsError {
    /// The strict neutral codec rejected the source or candidate.
    #[error("image-adjustments codec rejected the payload: {0}")]
    Codec(#[from] image_adjustments_codec::DecodeError),
    /// The common semantic image boundary rejected a projected value.
    #[error("invalid image adjustment value: {0}")]
    Semantic(#[from] litchi_iwa_common::shape::image::Error),
}

/// One mutable image-adjustment value staged against an immutable package.
pub struct SlideImageAdjustmentsEdit<'a> {
    source: &'a Package,
    selection: ImageSelection,
    after: ImageAdjustments,
}

impl fmt::Debug for SlideImageAdjustmentsEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideImageAdjustmentsEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("image_position", &self.selection.image_position)
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl<'a> SlideImageAdjustmentsEdit<'a> {
    fn new<'slide>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        image_selector: impl Into<ImageSelector>,
    ) -> Result<Self, SlideImageAdjustmentsError> {
        let mut budget = ImageBudget::for_package(source)?;
        let selection = select_image_with_budget(
            source,
            slide_selector.into(),
            image_selector.into(),
            true,
            &mut budget,
        )?;
        Ok(Self {
            source,
            after: selection.before,
            selection,
        })
    }

    /// Return the selected slide source position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected image source position.
    #[must_use]
    pub const fn image_position(&self) -> Position {
        self.selection.image_position
    }

    /// Return the controls read before this edit.
    #[must_use]
    pub const fn before(&self) -> ImageAdjustments {
        self.selection.before
    }

    /// Return the controls currently staged for publication.
    #[must_use]
    pub const fn after(&self) -> ImageAdjustments {
        self.after
    }

    /// Replace all three optional inspector controls.
    pub fn set(
        mut self,
        adjustments: ImageAdjustments,
    ) -> Result<Self, SlideImageAdjustmentsError> {
        self.after = adjustments;
        Ok(self)
    }

    /// Commit this exact-source edit after reopening and verifying its candidate.
    pub fn commit(self) -> Result<SlideImageAdjustmentsCommit, SlideImageAdjustmentsError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible image-adjustment patch.
#[derive(Clone, PartialEq)]
pub struct SlideImageAdjustmentsPatch {
    artifacts: ExactArtifacts,
    selection: ImageSelection,
    before: ImageAdjustments,
    after: ImageAdjustments,
    touched_components: usize,
}

impl fmt::Debug for SlideImageAdjustmentsPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideImageAdjustmentsPatch")
            .field("slide_position", &self.selection.slide_position)
            .field("image_position", &self.selection.image_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideImageAdjustmentsPatch {
    /// Return the selected slide source position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected image source position.
    #[must_use]
    pub const fn image_position(&self) -> Position {
        self.selection.image_position
    }

    /// Return the source controls required before applying this patch.
    #[must_use]
    pub const fn before(&self) -> ImageAdjustments {
        self.before
    }

    /// Return the controls produced by this patch.
    #[must_use]
    pub const fn after(&self) -> ImageAdjustments {
        self.after
    }

    /// Return a diagnostic fingerprint of the exact source artifact.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return a diagnostic fingerprint of the committed target artifact.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether the patch retains the exact source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact-source inverse operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            before: self.after,
            after: self.before,
            touched_components: self.touched_components,
        }
    }
}

/// Compact evidence describing one image-adjustment publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideImageAdjustmentsDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl SlideImageAdjustmentsDiagnostics {
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

    /// Return the number of IWA components rewritten.
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

/// Fully verified result of one image-adjustment transaction.
#[must_use = "a Keynote image-adjustment commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideImageAdjustmentsCommit {
    package: Package,
    patch: SlideImageAdjustmentsPatch,
    diagnostics: SlideImageAdjustmentsDiagnostics,
}

impl SlideImageAdjustmentsCommit {
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

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideImageAdjustmentsPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideImageAdjustmentsDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq)]
struct ImageSelection {
    slide_position: Position,
    image_position: Position,
    slide_identifier: u64,
    node_identifier: u64,
    image_identifier: u64,
    message_index: usize,
    slide_component_name: Arc<str>,
    layout: bool,
    before: ImageAdjustments,
}

impl fmt::Debug for ImageSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ImageSelection")
            .field("slide_position", &self.slide_position)
            .field("image_position", &self.image_position)
            .field("layout", &self.layout)
            .finish_non_exhaustive()
    }
}

#[derive(Debug, Clone, Copy)]
struct ImageEntry {
    identifier: u64,
    message_index: usize,
    layout: bool,
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
    nesting: usize,
    references: usize,
    allocations: usize,
    retained: usize,
}

impl ImageBudget {
    fn for_package(package: &Package) -> Result<Self, SlideImageAdjustmentsError> {
        let limits = package.wire_limits().map_err(map_wire_error)?;
        let physical = package.state.options.archive();
        let source = usize::try_from(physical.max_input_bytes())
            .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
        let aggregate = source
            .checked_mul(4)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        let semantic = package.semantic_limits();
        Ok(Self {
            limits,
            max_input: aggregate,
            max_output: aggregate,
            max_fields: limits.max_fields(),
            max_work: limits.max_rewrite_work(),
            max_nesting: limits.max_nesting(),
            max_references: semantic.max_references(),
            max_allocations: semantic
                .max_objects()
                .checked_add(semantic.max_references())
                .ok_or(SlideImageAdjustmentsError::InvalidSource)?,
            max_retained: aggregate,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: SlideImageAdjustmentsLimitKind,
    ) -> Result<(), SlideImageAdjustmentsError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        if observed > maximum {
            return Err(SlideImageAdjustmentsError::LimitExceeded {
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
    ) -> Result<(), SlideImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            payload.len(),
            self.max_input,
            SlideImageAdjustmentsLimitKind::WireBytes,
        )?;
        Self::add(
            &mut self.fields,
            fields,
            self.max_fields,
            SlideImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            payload.len(),
            self.max_work,
            SlideImageAdjustmentsLimitKind::WireWork,
        )?;
        self.nesting(depth)
    }

    fn source(&mut self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            SlideImageAdjustmentsLimitKind::WireBytes,
        )
    }

    fn work(&mut self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            SlideImageAdjustmentsLimitKind::WireWork,
        )
    }

    fn nesting(&mut self, depth: usize) -> Result<(), SlideImageAdjustmentsError> {
        if depth > self.max_nesting {
            return Err(SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
        Ok(())
    }

    fn references(&mut self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SlideImageAdjustmentsLimitKind::References,
        )
    }

    fn allocations(&mut self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SlideImageAdjustmentsLimitKind::Allocations,
        )
    }

    fn retained(&mut self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SlideImageAdjustmentsLimitKind::Retained,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            SlideImageAdjustmentsLimitKind::OutputBytes,
        )
    }

    fn preflight(
        current: usize,
        amount: usize,
        maximum: usize,
        kind: SlideImageAdjustmentsLimitKind,
    ) -> Result<(), SlideImageAdjustmentsError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        if observed > maximum {
            return Err(SlideImageAdjustmentsError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        Ok(())
    }

    fn preflight_allocations(&self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        Self::preflight(
            self.allocations,
            amount,
            self.max_allocations,
            SlideImageAdjustmentsLimitKind::Allocations,
        )
    }

    fn preflight_retained(&self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        Self::preflight(
            self.retained,
            amount,
            self.max_retained,
            SlideImageAdjustmentsLimitKind::Retained,
        )
    }

    fn residual_wire_limits(&self) -> Result<WireLimits, SlideImageAdjustmentsError> {
        let input = self.max_input.checked_sub(self.input).ok_or(
            SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::WireBytes,
                observed: self.max_input.saturating_add(1) as u64,
                maximum: self.max_input as u64,
            },
        )?;
        if input == 0 {
            return Err(SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::WireBytes,
                observed: self.max_input.saturating_add(1) as u64,
                maximum: self.max_input as u64,
            });
        }
        let fields = self.max_fields.checked_sub(self.fields).ok_or(
            SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::WireFields,
                observed: self.max_fields.saturating_add(1) as u64,
                maximum: self.max_fields as u64,
            },
        )?;
        if fields == 0 {
            return Err(SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::WireFields,
                observed: self.max_fields.saturating_add(1) as u64,
                maximum: self.max_fields as u64,
            });
        }
        let work = self.max_work.checked_sub(self.work).ok_or(
            SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::WireWork,
                observed: self.max_work.saturating_add(1) as u64,
                maximum: self.max_work as u64,
            },
        )?;
        if work == 0 {
            return Err(SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::WireWork,
                observed: self.max_work.saturating_add(1) as u64,
                maximum: self.max_work as u64,
            });
        }
        self.limits
            .with_input_bytes(self.limits.max_input_bytes().min(input))
            .and_then(|value| value.with_fields(self.limits.max_fields().min(fields)))
            .and_then(|value| value.with_rewrite_work(self.limits.max_rewrite_work().min(work)))
            .and_then(|value| value.with_nesting(self.limits.max_nesting().min(self.max_nesting)))
            .map_err(map_wire_error)
    }

    fn codec_options(
        &self,
        source: &[u8],
    ) -> Result<image_adjustments_codec::DecodeOptions, SlideImageAdjustmentsError> {
        let limits = self.residual_wire_limits()?;
        let recursion = u32::try_from(limits.max_nesting())
            .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
        let output = self.max_output.checked_sub(self.output).ok_or(
            SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::OutputBytes,
                observed: self.max_output.saturating_add(1) as u64,
                maximum: self.max_output as u64,
            },
        )?;
        if output == 0 {
            return Err(SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::OutputBytes,
                observed: self.max_output.saturating_add(1) as u64,
                maximum: self.max_output as u64,
            });
        }
        Ok(image_adjustments_codec::DecodeOptions::new(
            limits.max_input_bytes().min(source.len().max(1)),
            limits.max_fields(),
            limits.max_rewrite_work(),
            recursion,
        )
        .with_max_output_bytes(output))
    }

    fn codec_report(
        &mut self,
        report: image_adjustments_codec::DecodeReport,
    ) -> Result<(), SlideImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            report.input_bytes(),
            self.max_input,
            SlideImageAdjustmentsLimitKind::WireBytes,
        )?;
        Self::add(
            &mut self.fields,
            report.fields(),
            self.max_fields,
            SlideImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            report.work_bytes(),
            self.max_work,
            SlideImageAdjustmentsLimitKind::WireWork,
        )?;
        Self::add(
            &mut self.allocations,
            report.allocations(),
            self.max_allocations,
            SlideImageAdjustmentsLimitKind::Allocations,
        )?;
        Self::add(
            &mut self.retained,
            report.retained_bytes(),
            self.max_retained,
            SlideImageAdjustmentsLimitKind::Retained,
        )?;
        Self::add(
            &mut self.retained,
            report.scratch_bytes(),
            self.max_retained,
            SlideImageAdjustmentsLimitKind::Retained,
        )?;
        let depth = usize::try_from(report.max_depth()).unwrap_or(usize::MAX);
        self.nesting(depth)
    }

    fn codec_requirements(
        &mut self,
        requirements: image_adjustments_codec::RewriteExecutionRequirements,
    ) -> Result<(), SlideImageAdjustmentsError> {
        self.output(requirements.output_bytes)?;
        Self::add(
            &mut self.fields,
            requirements.fields,
            self.max_fields,
            SlideImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            requirements.work_bytes,
            self.max_work,
            SlideImageAdjustmentsLimitKind::WireWork,
        )?;
        self.nesting(usize::try_from(requirements.max_depth).unwrap_or(usize::MAX))?;
        self.allocations(requirements.allocations)?;
        self.retained(requirements.retained_bytes)?;
        self.retained(requirements.scratch_bytes)?;
        Ok(())
    }

    fn preflight_output(&self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        let observed = self
            .output
            .checked_add(amount)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        if observed > self.max_output {
            return Err(SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::OutputBytes,
                observed: observed as u64,
                maximum: self.max_output as u64,
            });
        }
        Ok(())
    }

    fn preflight_work(&self, amount: usize) -> Result<(), SlideImageAdjustmentsError> {
        let observed = self
            .work
            .checked_add(amount)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        if observed > self.max_work {
            return Err(SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::WireWork,
                observed: observed as u64,
                maximum: self.max_work as u64,
            });
        }
        Ok(())
    }

    fn candidate_reopen(
        &mut self,
        package: &Package,
        bytes: usize,
    ) -> Result<(), SlideImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            bytes,
            self.max_input,
            SlideImageAdjustmentsLimitKind::WireBytes,
        )?;
        Self::add(
            &mut self.work,
            bytes,
            self.max_work,
            SlideImageAdjustmentsLimitKind::WireWork,
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
    ) -> Result<(), SlideImageAdjustmentsError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.retained(requirements.scratch_bytes())?;
        Ok(())
    }
}

impl Package {
    /// Read image adjustments through typed slide and image selectors.
    pub fn slide_image_adjustments<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        image_selector: impl Into<ImageSelector>,
    ) -> Result<ImageAdjustments, SlideImageAdjustmentsError> {
        let mut budget = ImageBudget::for_package(self)?;
        Ok(select_image_with_budget(
            self,
            slide_selector.into(),
            image_selector.into(),
            false,
            &mut budget,
        )?
        .before)
    }

    /// Begin an exact immutable edit of one slide-owned image's adjustments.
    pub fn edit_slide_image_adjustments<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        image_selector: impl Into<ImageSelector>,
    ) -> Result<SlideImageAdjustmentsEdit<'_>, SlideImageAdjustmentsError> {
        SlideImageAdjustmentsEdit::new(self, slide_selector, image_selector)
    }

    /// Apply an exact-source checked image-adjustment patch.
    pub fn apply_slide_image_adjustments(
        &self,
        patch: &SlideImageAdjustmentsPatch,
    ) -> Result<SlideImageAdjustmentsCommit, SlideImageAdjustmentsError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(SlideImageAdjustmentsError::PatchConflict);
        }
        let mut budget = ImageBudget::for_package(self)?;
        budget.source(self.source_bytes().len())?;
        let current = select_image_with_budget(
            self,
            SlideSelector::position(patch.selection.slide_position),
            ImageSelector::position(patch.selection.image_position),
            true,
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(SlideImageAdjustmentsError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideImageAdjustmentsCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideImageAdjustmentsDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(SlideImageAdjustmentsError::PatchConflict);
        }
        reopen_target_patch(self, patch, &mut budget)
    }
}

/// Decode one borrowed complete `TSD.ImageArchive` payload for the legacy host.
#[cfg(feature = "internal-iwork-source")]
#[doc(hidden)]
pub fn __decode_image_adjustments_payload(
    source: &[u8],
    wire_limits: WireLimits,
) -> Result<ImageAdjustments, ImageAdjustmentsError> {
    decode_image_adjustments_payload(source, wire_limits)
}

#[cfg(feature = "internal-iwork-source")]
fn decode_image_adjustments_payload(
    source: &[u8],
    wire_limits: WireLimits,
) -> Result<ImageAdjustments, ImageAdjustmentsError> {
    Ok(decode_image_adjustments_snapshot(source, wire_limits)?.0)
}

#[cfg(feature = "internal-iwork-source")]
fn decode_image_adjustments_snapshot(
    source: &[u8],
    wire_limits: WireLimits,
) -> Result<(ImageAdjustments, image_adjustments_codec::DecodeReport), ImageAdjustmentsError> {
    let recursion_limit = u32::try_from(wire_limits.max_nesting()).unwrap_or(u32::MAX);
    let options = image_adjustments_codec::DecodeOptions::new(
        wire_limits.max_input_bytes().min(source.len().max(1)),
        wire_limits.max_fields(),
        wire_limits.max_rewrite_work(),
        recursion_limit,
    )
    .with_max_output_bytes(wire_limits.max_output_bytes());
    let (snapshot, report) =
        image_adjustments_codec::decode_image_adjustments_with_report(source, options)?;
    Ok((adjustments_from_snapshot(snapshot)?, report))
}

fn image_adjustments_write(adjustments: ImageAdjustments) -> ImageAdjustmentsWrite {
    ImageAdjustmentsWrite::from_values(
        adjustments.exposure().map(ImageAdjustment::value),
        adjustments.saturation().map(ImageAdjustment::value),
        adjustments
            .enhancement()
            .map(|value| matches!(value, ImageEnhancement::Enabled)),
    )
}

fn adjustments_from_snapshot(
    snapshot: ImageAdjustmentsSnapshot<'_>,
) -> Result<ImageAdjustments, litchi_iwa_common::shape::image::Error> {
    Ok(ImageAdjustments::new()
        .with_exposure(snapshot.exposure().map(ImageAdjustment::new).transpose()?)
        .with_saturation(
            snapshot
                .saturation()
                .map(ImageAdjustment::new)
                .transpose()?,
        )
        .with_enhancement(snapshot.enhance().map(image_enhancement_from_native)))
}

fn image_enhancement_from_native(value: bool) -> ImageEnhancement {
    if value {
        ImageEnhancement::Enabled
    } else {
        ImageEnhancement::Disabled
    }
}

fn select_image_with_budget(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    image_selector: ImageSelector,
    require_editable: bool,
    budget: &mut ImageBudget,
) -> Result<ImageSelection, SlideImageAdjustmentsError> {
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideImageAdjustmentsError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (slide_component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    let slide_message = unique_message(slide, SLIDE_MESSAGE_TYPE)?;
    let image_position = image_selector.as_position();
    let mut image_count = 0usize;
    let mut selected = None;
    visit_references_with_callback(slide_message.1, budget, |identifier, budget| {
        let (component, image) = package
            .object_with_component(identifier)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        if image
            .messages
            .iter()
            .all(|message| message.type_ != IMAGE_MESSAGE_TYPE)
        {
            return Ok(());
        }
        if component != slide_component_name {
            return Err(SlideImageAdjustmentsError::InvalidSource);
        }
        let (message_index, payload) = unique_message(image, IMAGE_MESSAGE_TYPE)?;
        let (parent, layout) = image_metadata(payload, budget)?;
        validate_image_parent(parent, record.slide_identifier)?;
        if image_count == image_position.get() {
            selected = Some(ImageEntry {
                identifier,
                message_index,
                layout,
            });
        }
        image_count = image_count
            .checked_add(1)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        Ok(())
    })?;
    let entry = selected.ok_or(SlideImageAdjustmentsError::ImagePositionNotFound {
        position: image_position,
    })?;
    if require_editable && entry.layout {
        return Err(SlideImageAdjustmentsError::UnsupportedDependency);
    }
    let image = package
        .object(entry.identifier)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    let payload = image
        .messages
        .get(entry.message_index)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?
        .data
        .as_slice();
    let before = decode_with_budget(payload, budget)?;
    ensure_unique_image_owner(package, record.slide_identifier, entry.identifier, budget)?;
    Ok(ImageSelection {
        slide_position,
        image_position,
        slide_identifier: record.slide_identifier,
        node_identifier: record.node_identifier,
        image_identifier: entry.identifier,
        message_index: entry.message_index,
        slide_component_name: Arc::from(slide_component_name),
        layout: entry.layout,
        before,
    })
}

fn decode_with_budget(
    payload: &[u8],
    budget: &mut ImageBudget,
) -> Result<ImageAdjustments, SlideImageAdjustmentsError> {
    let options = budget.codec_options(payload)?;
    let (snapshot, report) =
        image_adjustments_codec::decode_image_adjustments_with_report(payload, options)
            .map_err(map_codec_error)?;
    let adjustments = adjustments_from_snapshot(snapshot)
        .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
    budget.codec_report(report)?;
    Ok(adjustments)
}

fn commit_edit(
    source: &Package,
    selection: &ImageSelection,
    after: ImageAdjustments,
) -> Result<SlideImageAdjustmentsCommit, SlideImageAdjustmentsError> {
    let catalog = physical_catalog(source)?;
    let source_bytes = catalog.shared_source();
    let mut budget = ImageBudget::for_package(source)?;
    budget.source(source.source_bytes().len())?;
    let current = select_image_with_budget(
        source,
        SlideSelector::position(selection.slide_position),
        ImageSelector::position(selection.image_position),
        true,
        &mut budget,
    )?;
    if !same_selection(&current, selection) || current.before != selection.before {
        return Err(SlideImageAdjustmentsError::InvalidSource);
    }
    if selection.before == after {
        return Ok(SlideImageAdjustmentsCommit {
            package: source.snapshot(),
            patch: SlideImageAdjustmentsPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                selection: selection.clone(),
                before: selection.before,
                after,
                touched_components: 0,
            },
            diagnostics: SlideImageAdjustmentsDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(SlideImageAdjustmentsError::UnsupportedSource);
    }
    let candidate = rewrite_image(source, selection, after, &mut budget)?;
    candidate.validate().map_err(map_read_error)?;
    let target = physical_catalog(&candidate)?.shared_source();
    let selected = select_image_with_budget(
        &candidate,
        SlideSelector::position(selection.slide_position),
        ImageSelector::position(selection.image_position),
        true,
        &mut budget,
    )?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(SlideImageAdjustmentsError::Verification);
    }
    verify_locality(source, &candidate, selection, &mut budget)?;
    Ok(SlideImageAdjustmentsCommit {
        package: candidate,
        patch: SlideImageAdjustmentsPatch {
            artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
            selection: selection.clone(),
            before: selection.before,
            after,
            touched_components: 1,
        },
        diagnostics: SlideImageAdjustmentsDiagnostics::published(),
    })
}

fn reopen_target_patch(
    source: &Package,
    patch: &SlideImageAdjustmentsPatch,
    budget: &mut ImageBudget,
) -> Result<SlideImageAdjustmentsCommit, SlideImageAdjustmentsError> {
    budget.candidate_reopen(source, patch.artifacts.target().len())?;
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    let selected = select_image_with_budget(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        ImageSelector::position(patch.selection.image_position),
        true,
        budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(SlideImageAdjustmentsError::Verification);
    }
    verify_locality(source, &candidate, &patch.selection, budget)?;
    Ok(SlideImageAdjustmentsCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideImageAdjustmentsDiagnostics::published(),
    })
}

fn rewrite_image(
    source: &Package,
    selection: &ImageSelection,
    after: ImageAdjustments,
    budget: &mut ImageBudget,
) -> Result<Package, SlideImageAdjustmentsError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.slide_component_name.as_ref())
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideImageAdjustmentsError::InvalidSource);
    }
    let source_archive = component_archive(source, selection.slide_component_name.as_ref())?;
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
    let object = source_archive
        .object(selection.image_identifier)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    let original = object
        .messages
        .get(selection.message_index)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?
        .data
        .as_slice();
    let source_encoded_len = source_archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let before = decode_with_budget(original, budget)?;
    if before != selection.before {
        return Err(SlideImageAdjustmentsError::InvalidSource);
    }
    let options = budget.codec_options(original)?;
    let prepared = image_adjustments_codec::prepare_image_adjustments_rewrite(
        original,
        image_adjustments_write(after),
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
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_bound).map_err(map_core_error)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(SlideImageAdjustmentsError::LimitExceeded {
            kind: SlideImageAdjustmentsLimitKind::WireBytes,
            observed: compressed_bound as u64,
            maximum: snappy_limits.max_compressed_stream() as u64,
        });
    }
    let package_bound = source
        .source_bytes()
        .len()
        .checked_sub(entry.data().len())
        .and_then(|value| value.checked_add(compressed_bound))
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    budget.preflight_output(encoded_bound)?;
    budget.preflight_output(package_bound)?;
    budget.preflight_work(
        encoded_bound
            .checked_add(compressed_bound)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?,
    )?;
    let (clone_allocations, clone_retained) =
        archive_clone_requirements(source_archive, source_encoded_len)?;
    let (header_allocations, header_retained) = archive_header_rewrite_requirements(object)?;
    let (serialization_allocations, serialization_retained) =
        archive_serialization_requirements(source_archive)?;
    let total_allocations = clone_allocations
        .checked_add(header_allocations)
        .and_then(|amount| amount.checked_add(serialization_allocations))
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    budget.preflight_allocations(total_allocations)?;
    let temporary_retained = clone_retained
        .checked_add(header_retained)
        .and_then(|amount| amount.checked_add(serialization_retained))
        .and_then(|amount| amount.checked_add(encoded_bound))
        .and_then(|amount| amount.checked_add(compressed_bound))
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    budget.preflight_retained(temporary_retained)?;
    budget.allocations(total_allocations)?;
    budget.retained(clone_retained)?;
    budget.retained(header_retained)?;
    budget.retained(serialization_retained)?;
    let mut archive = source_archive.clone();
    let rewritten = prepared
        .execute(requirements.exact_limits())
        .map_err(map_codec_error)?
        .into_output();
    let verified = decode_with_budget(&rewritten, budget)?;
    if verified != after {
        return Err(SlideImageAdjustmentsError::Verification);
    }
    archive
        .object_mut(selection.image_identifier)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.message_index,
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
    let edits = [EntryEdit::new(
        selection.slide_component_name.as_ref(),
        &compressed,
    )];
    let prepared_reassembly = catalog
        .prepare_reassembly_with_deletions(&edits, &[], source.state.options.archive())
        .map_err(map_archive_error)?;
    let reassembly_requirements = prepared_reassembly.execution_requirements();
    budget.reassembly(reassembly_requirements)?;
    budget.candidate_reopen(source, reassembly_requirements.output_bytes())?;
    let output = prepared_reassembly
        .execute(reassembly_requirements.exact_limits())
        .map_err(map_archive_error)?;
    Package::from_source_with_options(output.into(), source.state.options).map_err(map_read_error)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &ImageSelection,
    budget: &mut ImageBudget,
) -> Result<(), SlideImageAdjustmentsError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    if source_catalog.package().len() != candidate_catalog.package().len() {
        return Err(SlideImageAdjustmentsError::Verification);
    }
    for source_entry in source_catalog.package().iter() {
        budget.work(source_entry.name().len().saturating_add(1))?;
        let candidate_entry = candidate_catalog
            .package()
            .iter()
            .find(|entry| entry.name() == source_entry.name())
            .ok_or(SlideImageAdjustmentsError::Verification)?;
        if source_entry.name() != selection.slide_component_name.as_ref()
            && source_entry.data() != candidate_entry.data()
        {
            return Err(SlideImageAdjustmentsError::Verification);
        }
    }
    let source_archive = component_archive(source, selection.slide_component_name.as_ref())?;
    let candidate_archive = component_archive(candidate, selection.slide_component_name.as_ref())?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(SlideImageAdjustmentsError::Verification);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    for source_object in &source_archive.objects {
        budget.work(source_object.messages.len())?;
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(SlideImageAdjustmentsError::Verification)?;
        let candidate_object = candidate_archive
            .object(identifier)
            .ok_or(SlideImageAdjustmentsError::Verification)?;
        if identifier == selection.image_identifier {
            let candidate_message = candidate_object
                .messages
                .get(selection.message_index)
                .ok_or(SlideImageAdjustmentsError::Verification)?;
            let (clone_allocations, clone_retained) = archive_object_clone_requirements(
                source_object,
                candidate_object,
                candidate_message.data.len(),
            )?;
            budget.preflight_allocations(clone_allocations)?;
            budget.preflight_retained(clone_retained)?;
            budget.allocations(clone_allocations)?;
            budget.retained(clone_retained)?;
            let mut expected = source_object.clone();
            let replacement = RawMessage {
                type_: candidate_message.type_,
                data: candidate_message.data.clone(),
            };
            expected
                .replace_message_preserving_header_with_limits(
                    selection.message_index,
                    replacement,
                    archive_limits,
                )
                .map_err(map_core_error)?;
            expected.header_length = candidate_object.header_length;
            expected.data_length = candidate_object.data_length;
            if !expected.same_content_ignoring_offsets(candidate_object) {
                return Err(SlideImageAdjustmentsError::Verification);
            }
        } else if !source_object.same_content_ignoring_offsets(candidate_object) {
            return Err(SlideImageAdjustmentsError::Verification);
        }
    }
    Ok(())
}

fn component_archive<'package>(
    package: &'package Package,
    name: &str,
) -> Result<&'package Archive, SlideImageAdjustmentsError> {
    physical_catalog(package)?;
    package
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == name)
        .map(|component| component.archive())
        .ok_or(SlideImageAdjustmentsError::InvalidSource)
}

fn rewritten_archive_bound(
    source_encoded_len: usize,
    original_len: usize,
    replacement_len: usize,
) -> Result<usize, SlideImageAdjustmentsError> {
    const MAX_REWRITE_FRAMING_GROWTH: usize = 64;
    source_encoded_len
        .checked_sub(original_len)
        .and_then(|value| value.checked_add(replacement_len))
        .and_then(|value| value.checked_add(MAX_REWRITE_FRAMING_GROWTH))
        .ok_or(SlideImageAdjustmentsError::InvalidSource)
}

fn archive_header_rewrite_requirements(
    object: &litchi_iwa_core::ArchiveObject,
) -> Result<(usize, usize), SlideImageAdjustmentsError> {
    let header = usize::try_from(object.header_length)
        .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
    let retained = header
        .checked_add(64)
        .and_then(|header| header.checked_mul(3))
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    // replace_message_preserving_header_with_limits constructs canonical,
    // rewritten, and post-mutation header buffers before publishing the
    // changed message.  The core API currently exposes no fallible clone of
    // those internal buffers, so the focused owner admits their bounded
    // staging before entering that mutation.
    Ok((3, retained))
}

fn archive_serialization_requirements(
    archive: &Archive,
) -> Result<(usize, usize), SlideImageAdjustmentsError> {
    let mut retained = 0usize;
    for object in &archive.objects {
        let header = usize::try_from(object.header_length)
            .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(header)
            .and_then(|amount| amount.checked_add(64))
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    }
    // One output Vec, one Snappy output Vec, and one temporary encoded header
    // per archive object.  Snappy's internal workspace remains governed by
    // its physical profile; the compressed result is charged separately by
    // the caller after the operation returns.
    let allocations = archive
        .objects
        .len()
        .checked_add(2)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    Ok((allocations, retained))
}

fn archive_clone_requirements(
    archive: &Archive,
    encoded_len: usize,
) -> Result<(usize, usize), SlideImageAdjustmentsError> {
    // `encoded_len` covers the source payload and its retained physical
    // framing.  The remaining accounting covers the decoded container
    // vectors, nested metadata vectors, and cloned raw-header boxes.  This
    // keeps the clone reservation an upper bound without reparsing the IWA
    // stream or exposing raw headers from the neutral archive layer.
    let mut allocations = 0usize;
    let mut retained = encoded_len;
    add_clone_vec::<litchi_iwa_core::ArchiveObject>(
        archive.objects.len(),
        &mut allocations,
        &mut retained,
    )?;
    for object in &archive.objects {
        add_clone_vec::<RawMessage>(object.messages.len(), &mut allocations, &mut retained)?;
        add_clone_vec::<litchi_iwa_core::MessageInfo>(
            object.archive_info.message_infos.len(),
            &mut allocations,
            &mut retained,
        )?;
        // The neutral archive keeps raw and canonical header boxes private.
        // A parsed object with a framed header can retain at most both boxes;
        // reserve that upper bound from the public framing length.
        if object.header_length != 0 {
            allocations = allocations
                .checked_add(2)
                .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
            retained = retained
                .checked_add(
                    usize::try_from(object.header_length)
                        .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?
                        .checked_mul(2)
                        .ok_or(SlideImageAdjustmentsError::InvalidSource)?,
                )
                .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
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
    source_bytes: usize,
) -> Result<(usize, usize), SlideImageAdjustmentsError> {
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let components = package.state.source.components();
    let entries = package.state.source.package().iter();
    let mut allocations = 3usize; // SourceCatalog, component storage, and Package state.
    let mut retained = source_bytes;
    let mut object_count = 0usize;

    // Shared source bytes are retained by the reopened package, while its
    // catalog and component/index metadata are rebuilt around the same
    // immutable allocation.  Count the known names and fixed-width index
    // slots before entering the parser.  The archive/catalog crates enforce
    // their own detailed physical ceilings; this focused reservation covers
    // the metadata sizes visible at this boundary without pretending to
    // account for every private parser temporary.
    for entry in entries {
        allocations = allocations
            .checked_add(1)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        let entry_slots = 2usize
            .checked_mul(size_of::<usize>())
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(entry.name().len())
            .and_then(|value| value.checked_add(entry_slots))
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    }
    if !components.is_empty() {
        allocations = allocations
            .checked_add(1)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        let component_slots = components
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_mul(size_of::<usize>()))
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(component_slots)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    }
    for component in components.iter() {
        allocations = allocations
            .checked_add(1)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(component.name().len())
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        object_count = object_count
            .checked_add(component.archive().objects.len())
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        let encoded_len = component
            .archive()
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let (archive_allocations, archive_retained) =
            archive_clone_requirements(component.archive(), encoded_len)?;
        allocations = allocations
            .checked_add(archive_allocations)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(archive_retained)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    }
    if object_count != 0 {
        allocations = allocations
            .checked_add(1)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(
                object_count
                    .checked_mul(size_of::<(u64, usize, usize)>())
                    .ok_or(SlideImageAdjustmentsError::InvalidSource)?,
            )
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    }
    // Conversion of the reassembly Vec into the package's shared source may
    // require a second immutable allocation before the temporary Vec drops.
    allocations = allocations
        .checked_add(1)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    retained = retained
        .checked_add(source_bytes)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    Ok((allocations, retained))
}

fn add_clone_vec<T>(
    length: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), SlideImageAdjustmentsError> {
    if length == 0 {
        return Ok(());
    }
    *allocations = allocations
        .checked_add(1)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    *retained = retained
        .checked_add(
            length
                .checked_mul(size_of::<T>())
                .ok_or(SlideImageAdjustmentsError::InvalidSource)?,
        )
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    Ok(())
}

fn add_clone_bytes(
    length: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), SlideImageAdjustmentsError> {
    add_clone_vec::<u8>(length, allocations, retained)
}

fn archive_object_clone_requirements(
    source: &litchi_iwa_core::ArchiveObject,
    candidate: &litchi_iwa_core::ArchiveObject,
    replacement_len: usize,
) -> Result<(usize, usize), SlideImageAdjustmentsError> {
    let mut allocations = 0usize;
    let mut retained = size_of::<litchi_iwa_core::ArchiveObject>();
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
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(
                usize::try_from(source.header_length)
                    .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?
                    .checked_mul(2)
                    .ok_or(SlideImageAdjustmentsError::InvalidSource)?,
            )
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    }
    // The selected candidate message is cloned into the expected object.  A
    // header-preserving replacement also builds bounded before/rewritten/after
    // header buffers before it publishes the mutation.
    add_clone_bytes(replacement_len, &mut allocations, &mut retained)?;
    let source_header = usize::try_from(source.header_length)
        .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
    let candidate_header = usize::try_from(candidate.header_length)
        .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
    let header_scratch = source_header
        .checked_add(candidate_header)
        .and_then(|bytes| bytes.checked_mul(3))
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    allocations = allocations
        .checked_add(3)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    retained = retained
        .checked_add(header_scratch)
        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    Ok((allocations, retained))
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideImageAdjustmentsError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideImageAdjustmentsError::EmptySlideName);
            }
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(map_slide_selector_error)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideImageAdjustmentsError::SlideNameNotFound)
        },
    }
}

fn visit_references_with_callback<F>(
    payload: &[u8],
    budget: &mut ImageBudget,
    mut callback: F,
) -> Result<(), SlideImageAdjustmentsError>
where
    F: FnMut(u64, &mut ImageBudget) -> Result<(), SlideImageAdjustmentsError>,
{
    let fields = WireView::parse_with_limits(payload, budget.residual_wire_limits()?)
        .map_err(map_wire_error)?;
    budget.scan(payload, fields.len(), 1)?;
    let count = fields
        .fields()
        .filter(|field| field.number() == SLIDE_OWNED_DRAWABLES_FIELD)
        .count();
    // Validate every reference before invoking the caller.  This keeps the
    // graph's strict error ordering while avoiding a vector proportional to
    // the number of slide-owned objects.
    for field in fields
        .fields()
        .filter(|field| field.number() == SLIDE_OWNED_DRAWABLES_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(SlideImageAdjustmentsError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let nested = WireView::parse_with_limits(field.payload(), budget.residual_wire_limits()?)
            .map_err(map_wire_error)?;
        budget.scan(field.payload(), nested.len(), 2)?;
        strict_reference_view(&nested)?;
    }
    budget.references(count)?;
    for field in fields
        .fields()
        .filter(|field| field.number() == SLIDE_OWNED_DRAWABLES_FIELD)
    {
        let nested = WireView::parse_with_limits(field.payload(), budget.residual_wire_limits()?)
            .map_err(map_wire_error)?;
        callback(strict_reference_view(&nested)?, budget)?;
    }
    Ok(())
}

fn strict_reference_view(fields: &WireView<'_>) -> Result<u64, SlideImageAdjustmentsError> {
    let mut identifier = None;
    for reference in fields.fields() {
        reference.validate_canonical_key().map_err(map_wire_error)?;
        reference
            .validate_canonical_framing()
            .map_err(map_wire_error)?;
        match reference.number() {
            1 => {
                if reference.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideImageAdjustmentsError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(reference.payload())
                    .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
                if value == 0 || width != encoded_len(value) {
                    return Err(SlideImageAdjustmentsError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(SlideImageAdjustmentsError::InvalidSource),
            _ => {},
        }
    }
    identifier.ok_or(SlideImageAdjustmentsError::InvalidSource)
}

fn image_metadata(
    payload: &[u8],
    budget: &mut ImageBudget,
) -> Result<(u64, bool), SlideImageAdjustmentsError> {
    let fields = WireView::parse_with_limits(payload, budget.residual_wire_limits()?)
        .map_err(map_wire_error)?;
    budget.scan(payload, fields.len(), 1)?;
    let mut parent = None;
    let mut seen_super = false;
    let mut data = None;
    let mut flags = None;
    for field in fields.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            IMAGE_SUPER_FIELD => {
                if seen_super || field.wire_type() != 2 {
                    return Err(SlideImageAdjustmentsError::InvalidSource);
                }
                seen_super = true;
                let nested =
                    WireView::parse_with_limits(field.payload(), budget.residual_wire_limits()?)
                        .map_err(map_wire_error)?;
                budget.scan(field.payload(), nested.len(), 2)?;
                let mut found = None;
                for inner in nested.fields() {
                    inner.validate_canonical_key().map_err(map_wire_error)?;
                    inner.validate_canonical_framing().map_err(map_wire_error)?;
                    if inner.number() == DRAWABLE_PARENT_FIELD {
                        if found.is_some() || inner.wire_type() != 2 {
                            return Err(SlideImageAdjustmentsError::InvalidSource);
                        }
                        found = Some(strict_reference_payload(inner.payload(), budget)?);
                    }
                }
                parent = found;
            },
            IMAGE_FLAGS_FIELD => {
                if flags.is_some() || field.wire_type() != 0 {
                    return Err(SlideImageAdjustmentsError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
                if width != encoded_len(value) {
                    return Err(SlideImageAdjustmentsError::InvalidSource);
                }
                let value =
                    u32::try_from(value).map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
                flags = Some(value);
            },
            IMAGE_DATA_FIELD => {
                if data.is_some() || field.wire_type() != 2 {
                    return Err(SlideImageAdjustmentsError::InvalidSource);
                }
                data = Some(strict_reference_payload(field.payload(), budget)?);
            },
            _ => {},
        }
    }
    let _data_identifier = data.ok_or(SlideImageAdjustmentsError::InvalidSource)?;
    Ok((
        parent.ok_or(SlideImageAdjustmentsError::InvalidSource)?,
        flags.unwrap_or(0) & LAYOUT_IMAGE_FLAG != 0,
    ))
}

fn validate_image_parent(parent: u64, expected: u64) -> Result<(), SlideImageAdjustmentsError> {
    if parent == expected {
        Ok(())
    } else {
        Err(SlideImageAdjustmentsError::InvalidSource)
    }
}

fn strict_reference_payload(
    payload: &[u8],
    budget: &mut ImageBudget,
) -> Result<u64, SlideImageAdjustmentsError> {
    let fields = WireView::parse_with_limits(payload, budget.residual_wire_limits()?)
        .map_err(map_wire_error)?;
    budget.scan(payload, fields.len(), 3)?;
    let mut identifier = None;
    for field in fields.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideImageAdjustmentsError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideImageAdjustmentsError::InvalidSource)?;
                if value == 0 || width != encoded_len(value) {
                    return Err(SlideImageAdjustmentsError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(SlideImageAdjustmentsError::InvalidSource),
            _ => {},
        }
    }
    identifier.ok_or(SlideImageAdjustmentsError::InvalidSource)
}

fn unique_message(
    object: &litchi_iwa_core::ArchiveObject,
    message_type: u32,
) -> Result<(usize, &[u8]), SlideImageAdjustmentsError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideImageAdjustmentsError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        let info = object
            .archive_info
            .message_infos
            .get(index)
            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
        if info.type_ != message.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(SlideImageAdjustmentsError::InvalidSource);
        }
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(SlideImageAdjustmentsError::InvalidSource);
        }
    }
    selected.ok_or(SlideImageAdjustmentsError::InvalidSource)
}

fn ensure_unique_image_owner(
    package: &Package,
    slide_identifier: u64,
    image_identifier: u64,
    budget: &mut ImageBudget,
) -> Result<(), SlideImageAdjustmentsError> {
    let mut occurrences = 0usize;
    let mut selected_occurrences = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object
                .messages
                .iter()
                .all(|message| message.type_ != SLIDE_MESSAGE_TYPE)
            {
                continue;
            }
            let (_, slide_payload) = unique_message(object, SLIDE_MESSAGE_TYPE)?;
            visit_references_with_callback(slide_payload, budget, |reference, _budget| {
                if reference == image_identifier {
                    occurrences = occurrences
                        .checked_add(1)
                        .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
                    if object.archive_info.identifier == Some(slide_identifier) {
                        selected_occurrences = selected_occurrences
                            .checked_add(1)
                            .ok_or(SlideImageAdjustmentsError::InvalidSource)?;
                    }
                }
                Ok(())
            })?;
        }
    }
    if occurrences != 1 || selected_occurrences != 1 {
        return Err(SlideImageAdjustmentsError::InvalidSource);
    }
    Ok(())
}

fn same_selection(left: &ImageSelection, right: &ImageSelection) -> bool {
    left.slide_position == right.slide_position
        && left.image_position == right.image_position
        && left.slide_identifier == right.slide_identifier
        && left.node_identifier == right.node_identifier
        && left.image_identifier == right.image_identifier
        && left.message_index == right.message_index
        && left.slide_component_name == right.slide_component_name
        && left.layout == right.layout
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideImageAdjustmentsError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideImageAdjustmentsError::UnsupportedSource),
    }
}

fn map_slide_selector_error(error: SlideSelectorError) -> SlideImageAdjustmentsError {
    match error {
        SlideSelectorError::DuplicateSlideName { .. } => {
            SlideImageAdjustmentsError::AmbiguousSelector
        },
        SlideSelectorError::EmptySlideName => SlideImageAdjustmentsError::EmptySlideName,
    }
}

fn map_read_error(error: ReadError) -> SlideImageAdjustmentsError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideImageAdjustmentsLimitKind::References,
                _ => SlideImageAdjustmentsLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideImageAdjustmentsLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideImageAdjustmentsLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideImageAdjustmentsLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideImageAdjustmentsLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideImageAdjustmentsError::Allocation { amount },
        _ => SlideImageAdjustmentsError::InvalidSource,
    }
}

fn map_codec_error(error: image_adjustments_codec::DecodeError) -> SlideImageAdjustmentsError {
    if let Some(limit) = error.limit_kind() {
        let (observed, maximum) = error.limit_values().unwrap_or((0, 0));
        let kind = match limit {
            image_adjustments_codec::DecodeLimit::InputBytes => {
                SlideImageAdjustmentsLimitKind::WireBytes
            },
            image_adjustments_codec::DecodeLimit::OutputBytes
            | image_adjustments_codec::DecodeLimit::Retained => {
                SlideImageAdjustmentsLimitKind::OutputBytes
            },
            image_adjustments_codec::DecodeLimit::Fields => {
                SlideImageAdjustmentsLimitKind::WireFields
            },
            image_adjustments_codec::DecodeLimit::Nesting => {
                SlideImageAdjustmentsLimitKind::WireNesting
            },
            image_adjustments_codec::DecodeLimit::Work
            | image_adjustments_codec::DecodeLimit::Allocations
            | image_adjustments_codec::DecodeLimit::Scratch => {
                SlideImageAdjustmentsLimitKind::WireWork
            },
            _ => SlideImageAdjustmentsLimitKind::WireWork,
        };
        return SlideImageAdjustmentsError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    SlideImageAdjustmentsError::InvalidSource
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SlideImageAdjustmentsError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SlideImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    SlideImageAdjustmentsLimitKind::WireBytes
                },
                litchi_iwa_common::LimitKind::Fields => SlideImageAdjustmentsLimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => {
                    SlideImageAdjustmentsLimitKind::OutputBytes
                },
                litchi_iwa_common::LimitKind::Nesting => {
                    SlideImageAdjustmentsLimitKind::WireNesting
                },
                litchi_iwa_common::LimitKind::RewriteWork
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    SlideImageAdjustmentsLimitKind::WireWork
                },
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideImageAdjustmentsError::Allocation { amount }
        },
        _ => SlideImageAdjustmentsError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideImageAdjustmentsError {
    match error {
        litchi_iwa_archive::Error::Limit {
            observed,
            maximum,
            kind,
        } => SlideImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    SlideImageAdjustmentsLimitKind::WireBytes
                },
                litchi_iwa_archive::LimitKind::Entries
                | litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    SlideImageAdjustmentsLimitKind::WireFields
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideImageAdjustmentsLimitKind::OutputBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideImageAdjustmentsError::Allocation { amount }
        },
        _ => SlideImageAdjustmentsError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideImageAdjustmentsError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideImageAdjustmentsError::LimitExceeded {
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
                    SlideImageAdjustmentsLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems => {
                    SlideImageAdjustmentsLimitKind::WireFields
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideImageAdjustmentsLimitKind::WireNesting
                },
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideImageAdjustmentsError::Allocation { amount: requested }
        },
        _ => SlideImageAdjustmentsError::InvalidSource,
    }
}

#[cfg(test)]
mod tests {
    use super::{
        ImageBudget, ImageSelector, SlideImageAdjustmentsError, SlideImageAdjustmentsLimitKind,
        image_metadata, strict_reference_payload, validate_image_parent,
    };
    use litchi_core::Position;
    use litchi_iwa_common::{WireLimits, encode_varint_into};

    fn bytes_field(output: &mut Vec<u8>, field: u32, payload: &[u8]) {
        encode_varint_into(output, u64::from(field) << 3 | 2);
        encode_varint_into(output, payload.len() as u64);
        output.extend_from_slice(payload);
    }

    fn varint_field(output: &mut Vec<u8>, field: u32, value: u64) {
        encode_varint_into(output, u64::from(field) << 3);
        encode_varint_into(output, value);
    }

    fn reference(identifier: u64) -> Vec<u8> {
        let mut output = Vec::new();
        varint_field(&mut output, 1, identifier);
        output
    }

    fn image_with_parent(parent: u64) -> Vec<u8> {
        let mut super_payload = Vec::new();
        bytes_field(&mut super_payload, 2, &reference(parent));
        let mut image = Vec::new();
        bytes_field(&mut image, 1, &super_payload);
        bytes_field(&mut image, 11, &reference(7));
        image
    }

    fn budget(limits: WireLimits) -> ImageBudget {
        ImageBudget {
            limits,
            max_input: limits.max_input_bytes(),
            max_output: limits.max_output_bytes(),
            max_fields: limits.max_fields(),
            max_work: limits.max_rewrite_work(),
            max_nesting: limits.max_nesting(),
            max_references: limits.max_fields(),
            max_allocations: limits.max_fields(),
            max_retained: limits.max_output_bytes(),
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
        }
    }

    #[test]
    fn image_selector_keeps_source_order_typed() {
        let from_index = ImageSelector::index(3);
        let from_position = ImageSelector::from(Position::new(3));
        assert_eq!(from_index, from_position);
        assert_eq!(from_index.as_index(), 3);
        assert_eq!(from_index.as_position(), Position::new(3));
    }

    #[test]
    fn image_metadata_rejects_duplicate_super_after_an_empty_super() {
        let mut source = Vec::new();
        bytes_field(&mut source, 1, &[]);
        let valid = image_with_parent(42);
        // image_with_parent already starts with the outer super field.
        source.extend_from_slice(&valid);
        let mut limits = budget(WireLimits::default());
        assert!(matches!(
            image_metadata(&source, &mut limits),
            Err(SlideImageAdjustmentsError::InvalidSource)
        ));
    }

    #[test]
    fn image_parent_mismatch_is_rejected_before_decode() {
        let mut limits = budget(WireLimits::default());
        let (parent, layout) =
            image_metadata(&image_with_parent(42), &mut limits).expect("strict image metadata");
        assert!(!layout);
        assert!(matches!(
            validate_image_parent(parent, 7),
            Err(SlideImageAdjustmentsError::InvalidSource)
        ));
    }

    #[test]
    fn image_metadata_rejects_duplicate_or_wrong_primary_data() {
        let mut duplicate = image_with_parent(42);
        bytes_field(&mut duplicate, 11, &reference(8));
        let mut limits = budget(WireLimits::default());
        assert!(matches!(
            image_metadata(&duplicate, &mut limits),
            Err(SlideImageAdjustmentsError::InvalidSource)
        ));

        let mut wrong_wire = image_with_parent(42);
        varint_field(&mut wrong_wire, 11, 7);
        let mut limits = budget(WireLimits::default());
        assert!(matches!(
            image_metadata(&wrong_wire, &mut limits),
            Err(SlideImageAdjustmentsError::InvalidSource)
        ));
    }

    #[test]
    fn image_metadata_rejects_flags_outside_native_uint32() {
        let mut source = image_with_parent(42);
        varint_field(&mut source, 7, u64::from(u32::MAX) + 1);
        let mut limits = budget(WireLimits::default());
        assert!(matches!(
            image_metadata(&source, &mut limits),
            Err(SlideImageAdjustmentsError::InvalidSource)
        ));
    }

    #[test]
    fn strict_reference_rejects_duplicate_identifier_fields() {
        let mut source = reference(7);
        varint_field(&mut source, 1, 8);
        assert!(matches!(
            strict_reference_payload(&source, &mut budget(WireLimits::default())),
            Err(SlideImageAdjustmentsError::InvalidSource)
        ));
    }

    #[test]
    fn aggregate_budget_fails_before_crossing_each_counter() {
        let limits = WireLimits::default().with_fields(2).unwrap();
        let mut budget = budget(limits);
        budget.scan(&[0x08], 1, 1).expect("first field");
        let error = budget.scan(&[0x08], 2, 1).unwrap_err();
        assert!(matches!(
            error,
            SlideImageAdjustmentsError::LimitExceeded {
                kind: SlideImageAdjustmentsLimitKind::WireFields,
                observed: 3,
                maximum: 2,
            }
        ));
    }
}
