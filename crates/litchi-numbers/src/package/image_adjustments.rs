//! Bounded Numbers image-adjustment projection and source-preserving rewrite.
//!
//! This module is the private native bridge for the legacy `litchi-iwa` host.
//! The complete `TSD.ImageArchive` payload remains byte-authoritative in the
//! host; this focused owner projects only the three semantic Image inspector
//! controls and delegates wire validation and rewriting to the strict Buffa
//! codec.

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement},
    varint::encoded_len,
    wire::{WireDescent, WirePreflight, WireView, preflight_wire_tree_with_limits},
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::image_adjustments_codec::{
    self as codec, ImageAdjustmentsSnapshot, ImageAdjustmentsWrite,
};
use thiserror::Error;

use super::{Error as PackageError, Package, SemanticLimitKind};
use crate::{SheetSelector, shape::image::ImageSelector};

/// Typed failures at the hidden Numbers image-adjustment seam.
#[doc(hidden)]
#[cfg(feature = "internal-iwork-source")]
#[derive(Debug, Clone, PartialEq, Eq, Error)]
pub enum ImageAdjustmentsError {
    /// The bounded Buffa/wire projection rejected the source or candidate.
    #[error("image-adjustments codec rejected the payload: {0}")]
    Codec(#[from] codec::DecodeError),
    /// The projected values failed the common semantic boundary.
    #[error("invalid image adjustment value: {0}")]
    Semantic(#[from] litchi_iwa_common::shape::image::Error),
}

const IMAGE_MESSAGE_TYPE: u32 = 3_005;
const SHEET_MESSAGE_TYPE: u32 = 2;
const FORM_BASED_SHEET_MESSAGE_TYPE: u32 = 3;
const FORM_SHEET_SUPER_FIELD: u32 = 1;
const IMAGE_SUPER_FIELD: u32 = 1;
const IMAGE_DATA_FIELD: u32 = 11;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const MIN_SIGN_EXTENDED_I32: u64 = 0xffff_ffff_8000_0000;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ReferenceKind {
    /// A drawable's `TSP.Reference` parent edge.
    Parent,
    /// A drawable's `TSP.DataReference` media-data edge.
    Data,
}

/// Resource category reported by a focused Numbers image-adjustment
/// operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SheetImageAdjustmentsLimitKind {
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
    /// Package or component output bytes.
    OutputBytes,
}

impl fmt::Display for SheetImageAdjustmentsLimitKind {
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

/// Failure raised by a selector-first Numbers image-adjustment read or edit.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SheetImageAdjustmentsError {
    /// The package does not retain an exact physical source suitable for edits.
    #[error("this Numbers source does not support physical image-adjustment edits")]
    UnsupportedSource,
    /// The selected graph is valid but outside this focused scalar owner.
    #[error("the requested Numbers image-adjustment graph operation is unsupported")]
    UnsupportedDependency,
    /// A semantic selector was ambiguous.
    #[error("the Numbers image-adjustment selector is ambiguous")]
    AmbiguousSelector,
    /// A name selector was empty.
    #[error("the Numbers sheet selector name cannot be empty")]
    EmptySheetName,
    /// No sheet matched the requested name.
    #[error("the Numbers workbook has no sheet matching the requested name")]
    SheetNameNotFound,
    /// No sheet matched the requested position.
    #[error("the Numbers workbook has no sheet at position {position:?}")]
    SheetPositionNotFound { position: Position },
    /// No image matched the requested source-order position.
    #[error("the selected Numbers sheet has no image at position {position:?}")]
    ImagePositionNotFound { position: Position },
    /// The source graph or payload is malformed.
    #[error("the Numbers image-adjustment source is invalid")]
    InvalidSource,
    /// A finite operation budget was exceeded.
    #[error(
        "Numbers image-adjustment {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SheetImageAdjustmentsLimitKind,
        observed: u64,
        maximum: u64,
    },
    /// A required temporary allocation could not be reserved.
    #[error("could not allocate {amount} units for Numbers image adjustments")]
    Allocation { amount: usize },
    /// The reopened candidate did not retain the requested semantic value.
    #[error("the edited Numbers image adjustments failed semantic verification")]
    Verification,
    /// The patch was created from another exact package artifact.
    #[error("the Numbers image-adjustment patch does not match the exact source package")]
    PatchConflict,
}

#[cfg(feature = "internal-iwork-source")]
fn codec_options(source: &[u8], limits: WireLimits) -> codec::DecodeOptions {
    let recursion_limit = u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX);
    codec::DecodeOptions::new(
        limits.max_input_bytes().min(source.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        recursion_limit,
    )
    .with_max_output_bytes(limits.max_output_bytes())
}

/// Decode one borrowed Numbers ImageArchive payload into semantic controls.
#[doc(hidden)]
#[cfg(feature = "internal-iwork-source")]
pub fn __decode_image_adjustments_payload(
    source: &[u8],
    limits: WireLimits,
) -> Result<ImageAdjustments, ImageAdjustmentsError> {
    let snapshot = codec::decode_image_adjustments(source, codec_options(source, limits))?;
    adjustments_from_snapshot(snapshot).map_err(ImageAdjustmentsError::Semantic)
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

#[derive(Debug, Clone, Copy)]
struct SheetDrawablePreflight {
    report: WirePreflight,
    drawable_count: usize,
}

fn preflight_sheet_drawables(
    message_type: u32,
    source: &[u8],
    maximum_drawables: usize,
    budget: &ImageBudget,
) -> Result<SheetDrawablePreflight, SheetImageAdjustmentsError> {
    let mut drawable_count = 0usize;
    let report = preflight_wire_tree_with_limits(source, budget.residual_wire_limits()?, |visit| {
        let is_drawable = match message_type {
            SHEET_MESSAGE_TYPE => visit.path().is_empty() && visit.field().number() == 2,
            FORM_BASED_SHEET_MESSAGE_TYPE => {
                visit.path() == [FORM_SHEET_SUPER_FIELD] && visit.field().number() == 2
            },
            _ => {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "invalid Numbers sheet type".into(),
                ));
            },
        };
        if is_drawable {
            if visit.field().wire_type() != 2 {
                return Err(litchi_iwa_common::Error::InvalidFormat(
                    "invalid Numbers drawable reference".into(),
                ));
            }
            let observed = drawable_count.saturating_add(1);
            if observed > maximum_drawables {
                return Err(litchi_iwa_common::Error::LimitExceeded {
                    kind: litchi_iwa_common::LimitKind::MaterializedCells,
                    observed,
                    limit: maximum_drawables,
                });
            }
            drawable_count = observed;
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
    .map_err(|error| match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind: litchi_iwa_common::LimitKind::MaterializedCells,
            observed,
            limit,
        } => SheetImageAdjustmentsError::LimitExceeded {
            kind: SheetImageAdjustmentsLimitKind::References,
            observed: observed as u64,
            maximum: limit as u64,
        },
        other => map_wire_error(other),
    })?;
    Ok(SheetDrawablePreflight {
        report,
        drawable_count,
    })
}

fn sheet_drawable_identifiers(
    message_type: u32,
    source: &[u8],
    budget: &mut ImageBudget,
) -> Result<Vec<u64>, SheetImageAdjustmentsError> {
    let maximum_drawables = budget.max_references.saturating_sub(budget.references);
    let preflight = preflight_sheet_drawables(message_type, source, maximum_drawables, budget)?;
    let identifier_bytes = preflight
        .drawable_count
        .checked_mul(size_of::<u64>())
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    budget.preflight_scan(preflight.report)?;
    // One path scratch allocation belongs to the bounded preflight, another
    // to the source-ordered projection, and one Vec backs the identifiers.
    budget.preflight_allocations(3)?;
    let retained = source
        .len()
        .checked_add(identifier_bytes)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    budget.preflight_retained(retained)?;
    budget.charge_preflight_scan(preflight.report)?;
    let (_name, identifiers) =
        super::names::preflight_sheet_payload(message_type, source, maximum_drawables)
            .map_err(map_names_error)?;
    if identifiers.len() != preflight.drawable_count {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    // `preflight_sheet_payload` performs its own bounded walk and validates
    // each local TSP.Reference before materializing the identifier Vec.
    budget.charge_preflight_scan(preflight.report)?;
    budget.allocations(3)?;
    budget.retained(retained)?;
    budget.references(identifiers.len())?;
    budget.work(identifiers.len())?;
    Ok(identifiers)
}

fn write_from_adjustments(adjustments: ImageAdjustments) -> ImageAdjustmentsWrite {
    ImageAdjustmentsWrite::from_values(
        adjustments.exposure().map(ImageAdjustment::value),
        adjustments.saturation().map(ImageAdjustment::value),
        adjustments.enhancement().map(image_enhancement_to_native),
    )
}

const fn image_enhancement_from_native(value: bool) -> ImageEnhancement {
    if value {
        ImageEnhancement::Enabled
    } else {
        ImageEnhancement::Disabled
    }
}

const fn image_enhancement_to_native(value: ImageEnhancement) -> bool {
    matches!(value, ImageEnhancement::Enabled)
}

/// One mutable image-adjustment value staged against an immutable Numbers
/// package.
pub struct SheetImageAdjustmentsEdit<'a> {
    source: &'a Package,
    selection: ImageSelection,
    after: ImageAdjustments,
}

impl fmt::Debug for SheetImageAdjustmentsEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SheetImageAdjustmentsEdit")
            .field("sheet_position", &self.selection.sheet_position)
            .field("image_position", &self.selection.image_position)
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl<'a> SheetImageAdjustmentsEdit<'a> {
    fn new<'sheet>(
        source: &'a Package,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        image_selector: impl Into<ImageSelector>,
    ) -> Result<Self, SheetImageAdjustmentsError> {
        let mut budget = ImageBudget::for_package(source)?;
        let selection = select_image_with_budget(
            source,
            sheet_selector.into(),
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

    /// Return the selected sheet source position.
    #[must_use]
    pub const fn sheet_position(&self) -> Position {
        self.selection.sheet_position
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
    ) -> Result<Self, SheetImageAdjustmentsError> {
        self.after = adjustments;
        Ok(self)
    }

    /// Commit this exact-source edit after reopening and verifying its
    /// candidate.
    pub fn commit(self) -> Result<SheetImageAdjustmentsCommit, SheetImageAdjustmentsError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible image-adjustment patch.
#[derive(Clone, PartialEq)]
pub struct SheetImageAdjustmentsPatch {
    artifacts: OwnedExactArtifacts,
    selection: ImageSelection,
    before: ImageAdjustments,
    after: ImageAdjustments,
}

impl fmt::Debug for SheetImageAdjustmentsPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SheetImageAdjustmentsPatch")
            .field("sheet_position", &self.selection.sheet_position)
            .field("image_position", &self.selection.image_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SheetImageAdjustmentsPatch {
    /// Return the selected sheet source position.
    #[must_use]
    pub const fn sheet_position(&self) -> Position {
        self.selection.sheet_position
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

    /// Return whether the patch is both semantically and byte-wise unchanged.
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
        }
    }
}

/// Compact evidence describing one image-adjustment publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SheetImageAdjustmentsDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl SheetImageAdjustmentsDiagnostics {
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
#[must_use = "a Numbers image-adjustment commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SheetImageAdjustmentsCommit {
    package: Package,
    patch: SheetImageAdjustmentsPatch,
    diagnostics: SheetImageAdjustmentsDiagnostics,
}

impl SheetImageAdjustmentsCommit {
    /// Borrow the fully reopened package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &SheetImageAdjustmentsPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SheetImageAdjustmentsDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq)]
struct ImageSelection {
    sheet_position: Position,
    image_position: Position,
    sheet_identifier: u64,
    image_identifier: u64,
    component_index: usize,
    object_index: usize,
    message_index: usize,
    component_name: Arc<str>,
    before: ImageAdjustments,
}

impl fmt::Debug for ImageSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ImageSelection")
            .field("sheet_position", &self.sheet_position)
            .field("image_position", &self.image_position)
            .finish_non_exhaustive()
    }
}

#[derive(Debug, Clone, Copy)]
struct ImageEntry {
    identifier: u64,
    component_index: usize,
    object_index: usize,
    message_index: usize,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ImageMetadata {
    parent_identifier: u64,
    data_identifier: u64,
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
    fn for_package(package: &Package) -> Result<Self, SheetImageAdjustmentsError> {
        let archive = package.state.options.archive();
        let core = archive
            .effective_archive_limits()
            .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
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
            .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
        let semantic = package.state.options.semantic();
        let max_allocations = semantic
            .max_objects()
            .checked_add(semantic.max_references())
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let aggregate = source.saturating_mul(8).max(source);
        Ok(Self {
            limits,
            max_input: aggregate,
            max_output: aggregate,
            max_fields: limits.max_fields(),
            max_work: limits.max_rewrite_work(),
            max_nesting: limits.max_nesting(),
            max_references: semantic.max_references(),
            max_allocations,
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
        kind: SheetImageAdjustmentsLimitKind,
    ) -> Result<(), SheetImageAdjustmentsError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        if observed > maximum {
            return Err(SheetImageAdjustmentsError::LimitExceeded {
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
        kind: SheetImageAdjustmentsLimitKind,
    ) -> Result<(), SheetImageAdjustmentsError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        if observed > maximum {
            return Err(SheetImageAdjustmentsError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        Ok(())
    }

    fn source(&mut self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            SheetImageAdjustmentsLimitKind::WireBytes,
        )
    }

    fn scan(
        &mut self,
        payload: &[u8],
        fields: usize,
        depth: usize,
    ) -> Result<(), SheetImageAdjustmentsError> {
        self.source(payload.len())?;
        Self::add(
            &mut self.fields,
            fields,
            self.max_fields,
            SheetImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            payload.len(),
            self.max_work,
            SheetImageAdjustmentsLimitKind::WireWork,
        )?;
        if depth > self.max_nesting {
            return Err(SheetImageAdjustmentsError::LimitExceeded {
                kind: SheetImageAdjustmentsLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
        Ok(())
    }

    fn preflight_scan(&self, report: WirePreflight) -> Result<(), SheetImageAdjustmentsError> {
        Self::preflight(
            self.input,
            report.scanned_bytes(),
            self.max_input,
            SheetImageAdjustmentsLimitKind::WireBytes,
        )?;
        Self::preflight(
            self.fields,
            report.fields(),
            self.max_fields,
            SheetImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::preflight(
            self.work,
            report.scanned_bytes(),
            self.max_work,
            SheetImageAdjustmentsLimitKind::WireWork,
        )?;
        let depth = report.max_depth();
        if depth > self.max_nesting {
            return Err(SheetImageAdjustmentsError::LimitExceeded {
                kind: SheetImageAdjustmentsLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn charge_preflight_scan(
        &mut self,
        report: WirePreflight,
    ) -> Result<(), SheetImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            report.scanned_bytes(),
            self.max_input,
            SheetImageAdjustmentsLimitKind::WireBytes,
        )?;
        Self::add(
            &mut self.fields,
            report.fields(),
            self.max_fields,
            SheetImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            report.scanned_bytes(),
            self.max_work,
            SheetImageAdjustmentsLimitKind::WireWork,
        )?;
        let depth = report.max_depth();
        if depth > self.max_nesting {
            return Err(SheetImageAdjustmentsError::LimitExceeded {
                kind: SheetImageAdjustmentsLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
        Ok(())
    }

    fn work(&mut self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            SheetImageAdjustmentsLimitKind::WireWork,
        )
    }

    fn references(&mut self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SheetImageAdjustmentsLimitKind::References,
        )
    }

    fn allocations(&mut self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SheetImageAdjustmentsLimitKind::Allocations,
        )
    }

    fn retained(&mut self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SheetImageAdjustmentsLimitKind::Retained,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            SheetImageAdjustmentsLimitKind::OutputBytes,
        )
    }

    fn preflight_output(&self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::preflight(
            self.output,
            amount,
            self.max_output,
            SheetImageAdjustmentsLimitKind::OutputBytes,
        )
    }

    fn preflight_work(&self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::preflight(
            self.work,
            amount,
            self.max_work,
            SheetImageAdjustmentsLimitKind::WireWork,
        )
    }

    fn preflight_allocations(&self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::preflight(
            self.allocations,
            amount,
            self.max_allocations,
            SheetImageAdjustmentsLimitKind::Allocations,
        )
    }

    fn preflight_retained(&self, amount: usize) -> Result<(), SheetImageAdjustmentsError> {
        Self::preflight(
            self.retained,
            amount,
            self.max_retained,
            SheetImageAdjustmentsLimitKind::Retained,
        )
    }

    fn residual_wire_limits(&self) -> Result<WireLimits, SheetImageAdjustmentsError> {
        let input = self.max_input.saturating_sub(self.input).max(1);
        let fields = self.max_fields.saturating_sub(self.fields).max(1);
        let work = self.max_work.saturating_sub(self.work).max(1);
        self.limits
            .with_input_bytes(self.limits.max_input_bytes().min(input))
            .and_then(|limits| limits.with_fields(limits.max_fields().min(fields)))
            .and_then(|limits| limits.with_rewrite_work(limits.max_rewrite_work().min(work)))
            .and_then(|limits| limits.with_nesting(limits.max_nesting().min(self.max_nesting)))
            .map_err(|_| SheetImageAdjustmentsError::InvalidSource)
    }

    fn codec_options(
        &self,
        source: &[u8],
    ) -> Result<codec::DecodeOptions, SheetImageAdjustmentsError> {
        let limits = self.residual_wire_limits()?;
        let output = self.max_output.saturating_sub(self.output).max(1);
        let recursion = u32::try_from(limits.max_nesting())
            .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
        Ok(codec::DecodeOptions::new(
            limits.max_input_bytes().min(source.len().max(1)),
            limits.max_fields(),
            limits.max_rewrite_work(),
            recursion,
        )
        .with_max_output_bytes(output))
    }

    fn codec_report(
        &mut self,
        report: codec::DecodeReport,
    ) -> Result<(), SheetImageAdjustmentsError> {
        Self::add(
            &mut self.input,
            report.input_bytes(),
            self.max_input,
            SheetImageAdjustmentsLimitKind::WireBytes,
        )?;
        Self::add(
            &mut self.fields,
            report.fields(),
            self.max_fields,
            SheetImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            report.work_bytes(),
            self.max_work,
            SheetImageAdjustmentsLimitKind::WireWork,
        )?;
        Self::add(
            &mut self.allocations,
            report.allocations(),
            self.max_allocations,
            SheetImageAdjustmentsLimitKind::Allocations,
        )?;
        Self::add(
            &mut self.retained,
            report
                .retained_bytes()
                .saturating_add(report.scratch_bytes()),
            self.max_retained,
            SheetImageAdjustmentsLimitKind::Retained,
        )?;
        let depth = usize::try_from(report.max_depth()).unwrap_or(usize::MAX);
        if depth > self.max_nesting {
            return Err(SheetImageAdjustmentsError::LimitExceeded {
                kind: SheetImageAdjustmentsLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
        Ok(())
    }

    fn codec_requirements(
        &mut self,
        requirements: codec::RewriteExecutionRequirements,
    ) -> Result<(), SheetImageAdjustmentsError> {
        self.output(requirements.output_bytes)?;
        Self::add(
            &mut self.fields,
            requirements.fields,
            self.max_fields,
            SheetImageAdjustmentsLimitKind::WireFields,
        )?;
        Self::add(
            &mut self.work,
            requirements.work_bytes,
            self.max_work,
            SheetImageAdjustmentsLimitKind::WireWork,
        )?;
        let depth = usize::try_from(requirements.max_depth).unwrap_or(usize::MAX);
        if depth > self.max_nesting {
            return Err(SheetImageAdjustmentsError::LimitExceeded {
                kind: SheetImageAdjustmentsLimitKind::WireNesting,
                observed: depth as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.nesting = self.nesting.max(depth);
        self.allocations(requirements.allocations)?;
        self.retained(
            requirements
                .retained_bytes
                .saturating_add(requirements.scratch_bytes),
        )?;
        Ok(())
    }

    fn candidate_reopen(
        &mut self,
        package: &Package,
        bytes: usize,
    ) -> Result<(), SheetImageAdjustmentsError> {
        self.source(bytes)?;
        self.work(bytes)?;
        let (allocations, retained) = package_reopen_requirements(package, bytes)?;
        self.preflight_allocations(allocations)?;
        self.preflight_retained(retained)?;
        self.allocations(allocations)?;
        self.retained(retained)
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SheetImageAdjustmentsError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(
            requirements
                .retained_bytes()
                .saturating_add(requirements.scratch_bytes()),
        )
    }
}

impl Package {
    /// Read image adjustments through typed sheet and image selectors.
    pub fn sheet_image_adjustments<'sheet>(
        &self,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        image_selector: impl Into<ImageSelector>,
    ) -> Result<ImageAdjustments, SheetImageAdjustmentsError> {
        let mut budget = ImageBudget::for_package(self)?;
        Ok(select_image_with_budget(
            self,
            sheet_selector.into(),
            image_selector.into(),
            false,
            &mut budget,
        )?
        .before)
    }

    /// Begin an exact immutable edit of one sheet-owned image's adjustments.
    pub fn edit_sheet_image_adjustments<'sheet>(
        &self,
        sheet_selector: impl Into<SheetSelector<'sheet>>,
        image_selector: impl Into<ImageSelector>,
    ) -> Result<SheetImageAdjustmentsEdit<'_>, SheetImageAdjustmentsError> {
        SheetImageAdjustmentsEdit::new(self, sheet_selector, image_selector)
    }

    /// Apply an exact-source checked image-adjustment patch.
    pub fn apply_sheet_image_adjustments(
        &self,
        patch: &SheetImageAdjustmentsPatch,
    ) -> Result<SheetImageAdjustmentsCommit, SheetImageAdjustmentsError> {
        let catalog = physical_catalog(self)?;
        let owner = catalog.__source_owner();
        if !patch.artifacts.authorizes_owner(&owner) {
            return Err(SheetImageAdjustmentsError::PatchConflict);
        }
        let mut budget = ImageBudget::for_package(self)?;
        budget.source(self.source_bytes().len())?;
        let current = select_image_with_budget(
            self,
            SheetSelector::position(patch.selection.sheet_position),
            ImageSelector::position(patch.selection.image_position),
            true,
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(SheetImageAdjustmentsError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(SheetImageAdjustmentsCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SheetImageAdjustmentsDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(SheetImageAdjustmentsError::UnsupportedSource);
        }
        reopen_target_patch(self, patch, &mut budget)
    }
}

fn select_image_with_budget(
    package: &Package,
    sheet_selector: SheetSelector<'_>,
    image_selector: ImageSelector,
    require_editable: bool,
    budget: &mut ImageBudget,
) -> Result<ImageSelection, SheetImageAdjustmentsError> {
    let sheet_position = resolve_sheet_position(package, sheet_selector)?;
    let document = Package::root_document(&package.state.components).map_err(map_package_error)?;
    let sheet_reference = document
        .sheet_references()
        .get(sheet_position.get())
        .ok_or(SheetImageAdjustmentsError::SheetPositionNotFound {
            position: sheet_position,
        })?;
    let sheet_identifier = sheet_reference.identifier();
    if sheet_identifier == 0 || sheet_reference.deprecated_is_external() == Some(true) {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    let resolved_sheet = package
        .state
        .index
        .resolve_ref_id(&package.state.components, sheet_identifier)
        .map_err(map_package_error)?
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    let sheet_component = package
        .state
        .components
        .catalog()
        .get_index(resolved_sheet.component_index)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    let sheet_object = sheet_component
        .archive()
        .objects
        .get(resolved_sheet.object_index)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    if sheet_object.archive_info.identifier != Some(sheet_identifier) {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    validate_message_metadata(sheet_object, resolved_sheet.messages.len())?;
    let has_sheet = unique_typed_message(sheet_object, SHEET_MESSAGE_TYPE)?.is_some();
    let has_form_sheet =
        unique_typed_message(sheet_object, FORM_BASED_SHEET_MESSAGE_TYPE)?.is_some();
    let sheet_message_type = sheet_message_type_for(has_sheet, has_form_sheet)?;
    let sheet_payload = unique_typed_message(sheet_object, sheet_message_type)?
        .ok_or(SheetImageAdjustmentsError::InvalidSource)
        .map(|(_, message)| message.data.as_slice())?;
    let drawable_identifiers =
        sheet_drawable_identifiers(sheet_message_type, sheet_payload, budget)?;

    let requested = image_selector.as_position();
    let mut image_count = 0usize;
    let mut selected = None;
    for drawable_identifier in drawable_identifiers {
        let resolved = package
            .state
            .index
            .resolve_ref_id(&package.state.components, drawable_identifier)
            .map_err(map_package_error)?
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let component = package
            .state
            .components
            .catalog()
            .get_index(resolved.component_index)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let object = component
            .archive()
            .objects
            .get(resolved.object_index)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        if object.archive_info.identifier != Some(drawable_identifier) {
            return Err(SheetImageAdjustmentsError::InvalidSource);
        }
        let Some((message_index, message)) = unique_typed_message(object, IMAGE_MESSAGE_TYPE)?
        else {
            continue;
        };
        if require_editable && resolved.component_index != resolved_sheet.component_index {
            return Err(SheetImageAdjustmentsError::UnsupportedDependency);
        }
        let info = validate_selected_image_metadata(object, message_index, budget)?;
        let metadata = image_metadata(message.data.as_slice(), budget)?;
        if !image_parent_metadata_is_owned(info, metadata.parent_identifier)
            || !image_data_metadata_is_owned(info, metadata.data_identifier)
            || metadata.parent_identifier != sheet_identifier
        {
            return Err(SheetImageAdjustmentsError::InvalidSource);
        }
        if image_count == requested.get() {
            selected = Some(ImageEntry {
                identifier: drawable_identifier,
                component_index: resolved.component_index,
                object_index: resolved.object_index,
                message_index,
            });
        }
        image_count = image_count
            .checked_add(1)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    }
    let entry = selected.ok_or(SheetImageAdjustmentsError::ImagePositionNotFound {
        position: requested,
    })?;
    let image_component = package
        .state
        .components
        .catalog()
        .get_index(entry.component_index)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    let image_object = image_component
        .archive()
        .objects
        .get(entry.object_index)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    let message = image_object
        .messages
        .get(entry.message_index)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    let before = decode_with_budget(message.data.as_slice(), budget)?;
    ensure_unique_image_owner(package, sheet_identifier, entry.identifier, budget)?;
    Ok(ImageSelection {
        sheet_position,
        image_position: requested,
        sheet_identifier,
        image_identifier: entry.identifier,
        component_index: entry.component_index,
        object_index: entry.object_index,
        message_index: entry.message_index,
        component_name: Arc::from(image_component.name()),
        before,
    })
}

fn decode_with_budget(
    payload: &[u8],
    budget: &mut ImageBudget,
) -> Result<ImageAdjustments, SheetImageAdjustmentsError> {
    let options = budget.codec_options(payload)?;
    let (snapshot, report) =
        codec::decode_image_adjustments_with_report(payload, options).map_err(map_codec_error)?;
    let adjustments = adjustments_from_snapshot(snapshot)
        .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
    budget.codec_report(report)?;
    Ok(adjustments)
}

fn resolve_sheet_position(
    package: &Package,
    selector: SheetSelector<'_>,
) -> Result<Position, SheetImageAdjustmentsError> {
    match selector {
        SheetSelector::Name(name) => {
            if name.is_empty() {
                return Err(SheetImageAdjustmentsError::EmptySheetName);
            }
            let mut matches = package
                .document()
                .sheets()
                .iter()
                .filter(|sheet| sheet.name() == name);
            let Some(sheet) = matches.next() else {
                return Err(SheetImageAdjustmentsError::SheetNameNotFound);
            };
            if matches.next().is_some() {
                return Err(SheetImageAdjustmentsError::AmbiguousSelector);
            }
            Ok(Position::new(sheet.index()))
        },
        SheetSelector::Index(index) => {
            if package.document().sheets().get(index).is_none() {
                return Err(SheetImageAdjustmentsError::SheetPositionNotFound {
                    position: Position::new(index),
                });
            }
            Ok(Position::new(index))
        },
    }
}

fn unique_typed_message(
    object: &litchi_iwa_core::ArchiveObject,
    message_type: u32,
) -> Result<Option<(usize, &RawMessage)>, SheetImageAdjustmentsError> {
    validate_message_metadata(object, object.messages.len())?;
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace((index, message)).is_some() {
            return Err(SheetImageAdjustmentsError::InvalidSource);
        }
    }
    Ok(selected)
}

fn sheet_message_type_for(
    has_sheet: bool,
    has_form_sheet: bool,
) -> Result<u32, SheetImageAdjustmentsError> {
    if has_sheet == has_form_sheet {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    Ok(if has_sheet {
        SHEET_MESSAGE_TYPE
    } else {
        FORM_BASED_SHEET_MESSAGE_TYPE
    })
}

fn validate_message_metadata(
    object: &litchi_iwa_core::ArchiveObject,
    expected_messages: usize,
) -> Result<(), SheetImageAdjustmentsError> {
    if object.messages.len() != object.archive_info.message_infos.len()
        || object.messages.len() != expected_messages
    {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    for (message, info) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
    {
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(SheetImageAdjustmentsError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_selected_image_metadata<'a>(
    object: &'a litchi_iwa_core::ArchiveObject,
    message_index: usize,
    budget: &mut ImageBudget,
) -> Result<&'a litchi_iwa_core::MessageInfo, SheetImageAdjustmentsError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    let message = object
        .messages
        .get(message_index)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    budget.work(image_metadata_work(info)?)?;
    if info.type_ != message.type_
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    Ok(info)
}

fn image_metadata_work(
    info: &litchi_iwa_core::MessageInfo,
) -> Result<usize, SheetImageAdjustmentsError> {
    let aggregate_data_work = info
        .data_references
        .len()
        .checked_mul(3)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    let field_list_work = info
        .field_infos
        .len()
        .checked_mul(4)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    let aggregate_object_work = info
        .object_references
        .len()
        .checked_mul(2)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    let mut work = 1usize
        // The selected image proof scans aggregate data references once for
        // data ownership and once for the parent/data cross-kind exclusion;
        // this sizing pass is the third traversal.
        .checked_add(aggregate_data_work)
        // FieldInfo is traversed while sizing, for data ownership, for the
        // parent/data cross-kind exclusion, and for parent ownership.
        .and_then(|amount| amount.checked_add(field_list_work))
        // Parent ownership scans aggregate object references once in addition
        // to this sizing pass.
        .and_then(|amount| amount.checked_add(aggregate_object_work))
        .and_then(|amount| amount.checked_add(info.diff_merge_version.len()))
        .and_then(|amount| amount.checked_add(info.diff_read_version.len()))
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    if let Some(path) = info.diff_field_path.as_ref() {
        work = work
            .checked_add(path.path.len())
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    }
    for path in &info.fields_to_remove {
        work = work
            .checked_add(path.path.len())
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    }
    for field_info in &info.field_infos {
        let path_work = field_info
            .path
            .path
            .len()
            .checked_mul(3)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let object_work = field_info
            .object_references
            .len()
            .checked_mul(2)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let data_work = field_info
            .data_references
            .len()
            .checked_mul(3)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        work = work
            // The path is compared by both ownership proofs in addition to
            // this sizing walk.
            .checked_add(path_work)
            .and_then(|amount| amount.checked_add(object_work))
            .and_then(|amount| amount.checked_add(data_work))
            .and_then(|amount| amount.checked_add(field_info.known_field_version.len()))
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        if let Some(identifier) = field_info.known_field_feature_identifier.as_ref() {
            work = work
                .checked_add(identifier.len())
                .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        }
    }
    Ok(work)
}

fn image_data_metadata_is_owned(info: &litchi_iwa_core::MessageInfo, identifier: u64) -> bool {
    if identifier == 0 {
        return false;
    }
    let aggregate = info
        .data_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    if aggregate != 1 {
        return false;
    }
    let mut field_occurrence = 0usize;
    for field_info in &info.field_infos {
        let occurrences = field_info
            .data_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if field_info.path.as_slice() != [IMAGE_DATA_FIELD] || occurrences != 1 {
            return false;
        }
        field_occurrence = field_occurrence.saturating_add(occurrences);
    }
    field_occurrence <= 1
}

fn image_parent_metadata_is_owned(info: &litchi_iwa_core::MessageInfo, identifier: u64) -> bool {
    if identifier == 0 {
        return false;
    }
    if info.data_references.contains(&identifier)
        || info
            .field_infos
            .iter()
            .any(|field_info| field_info.data_references.contains(&identifier))
    {
        return false;
    }
    let aggregate = info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count();
    let mut field_occurrence = 0usize;
    for field_info in &info.field_infos {
        let occurrences = field_info
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if field_info.path.as_slice() != [IMAGE_SUPER_FIELD, DRAWABLE_PARENT_FIELD]
            || occurrences != 1
        {
            return false;
        }
        field_occurrence = field_occurrence.saturating_add(occurrences);
    }
    matches!((aggregate, field_occurrence), (0, 0) | (1, 0 | 1))
}

fn image_metadata(
    payload: &[u8],
    budget: &mut ImageBudget,
) -> Result<ImageMetadata, SheetImageAdjustmentsError> {
    let fields = WireView::parse_with_limits(payload, budget.residual_wire_limits()?)
        .map_err(map_wire_error)?;
    budget.scan(payload, fields.len(), 1)?;
    let mut parent = None;
    let mut data_identifier = None;
    let mut seen_super = false;
    let mut seen_data = false;
    for field in fields.fields() {
        match field.number() {
            IMAGE_SUPER_FIELD => {
                if seen_super || field.wire_type() != 2 {
                    return Err(SheetImageAdjustmentsError::InvalidSource);
                }
                field.validate_canonical_framing().map_err(map_wire_error)?;
                seen_super = true;
                let nested =
                    WireView::parse_with_limits(field.payload(), budget.residual_wire_limits()?)
                        .map_err(map_wire_error)?;
                budget.scan(field.payload(), nested.len(), 2)?;
                for inner in nested.fields() {
                    if inner.number() == DRAWABLE_PARENT_FIELD {
                        if parent.is_some() || inner.wire_type() != 2 {
                            return Err(SheetImageAdjustmentsError::InvalidSource);
                        }
                        inner.validate_canonical_framing().map_err(map_wire_error)?;
                        parent = Some(strict_reference_payload(
                            inner.payload(),
                            budget,
                            ReferenceKind::Parent,
                        )?);
                    }
                }
            },
            IMAGE_DATA_FIELD => {
                if seen_data || field.wire_type() != 2 {
                    return Err(SheetImageAdjustmentsError::InvalidSource);
                }
                field.validate_canonical_framing().map_err(map_wire_error)?;
                seen_data = true;
                data_identifier = Some(strict_reference_payload(
                    field.payload(),
                    budget,
                    ReferenceKind::Data,
                )?);
            },
            _ => {},
        }
    }
    budget.references(2)?;
    if !seen_super || !seen_data {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    Ok(ImageMetadata {
        parent_identifier: parent.ok_or(SheetImageAdjustmentsError::InvalidSource)?,
        data_identifier: data_identifier.ok_or(SheetImageAdjustmentsError::InvalidSource)?,
    })
}

fn strict_reference_payload(
    payload: &[u8],
    budget: &mut ImageBudget,
    kind: ReferenceKind,
) -> Result<u64, SheetImageAdjustmentsError> {
    let fields = WireView::parse_with_limits(payload, budget.residual_wire_limits()?)
        .map_err(map_wire_error)?;
    budget.scan(payload, fields.len(), 3)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in fields.fields() {
        match kind {
            ReferenceKind::Data => {
                // TSP.DataReference defines only identifier. Every other
                // field is opaque here, including the historical field 2/3
                // numbers used by TSP.Reference. The source codec owns their
                // bytes and the image rewrite never re-encodes this payload.
                if field.number() != 1 {
                    continue;
                }
                field.validate_canonical_key().map_err(map_wire_error)?;
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SheetImageAdjustmentsError::InvalidSource);
                }
                let value = canonical_reference_varint(field.payload())?;
                if value == 0 {
                    return Err(SheetImageAdjustmentsError::InvalidSource);
                }
                identifier = Some(value);
            },
            ReferenceKind::Parent => match field.number() {
                1 => {
                    field.validate_canonical_key().map_err(map_wire_error)?;
                    if field.wire_type() != 0 || identifier.is_some() {
                        return Err(SheetImageAdjustmentsError::InvalidSource);
                    }
                    let value = canonical_reference_varint(field.payload())?;
                    if value == 0 {
                        return Err(SheetImageAdjustmentsError::InvalidSource);
                    }
                    identifier = Some(value);
                },
                2 => {
                    field.validate_canonical_key().map_err(map_wire_error)?;
                    if field.wire_type() != 0 || deprecated_type.is_some() {
                        return Err(SheetImageAdjustmentsError::InvalidSource);
                    }
                    let value = canonical_reference_varint(field.payload())?;
                    if value > u64::from(i32::MAX.unsigned_abs()) && value < MIN_SIGN_EXTENDED_I32 {
                        return Err(SheetImageAdjustmentsError::InvalidSource);
                    }
                    deprecated_type = Some(value);
                },
                3 => {
                    field.validate_canonical_key().map_err(map_wire_error)?;
                    if field.wire_type() != 0 {
                        return Err(SheetImageAdjustmentsError::InvalidSource);
                    }
                    let value = canonical_reference_varint(field.payload())?;
                    if value > 1 || external.is_some() {
                        return Err(SheetImageAdjustmentsError::InvalidSource);
                    }
                    external = Some(value != 0);
                },
                _ => {},
            },
        }
    }
    if kind == ReferenceKind::Parent && external == Some(true) {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    identifier.ok_or(SheetImageAdjustmentsError::InvalidSource)
}

fn canonical_reference_varint(payload: &[u8]) -> Result<u64, SheetImageAdjustmentsError> {
    let (value, width) =
        decode_varint_from_bytes(payload).map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
    if width != payload.len() || width != encoded_len(value) {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    Ok(value)
}

fn ensure_unique_image_owner(
    package: &Package,
    sheet_identifier: u64,
    image_identifier: u64,
    budget: &mut ImageBudget,
) -> Result<(), SheetImageAdjustmentsError> {
    let document = Package::root_document(&package.state.components).map_err(map_package_error)?;
    let mut occurrences = 0usize;
    let mut selected_occurrences = 0usize;
    for reference in document.sheet_references() {
        let identifier = reference.identifier();
        if identifier == 0 || reference.deprecated_is_external() == Some(true) {
            return Err(SheetImageAdjustmentsError::InvalidSource);
        }
        let resolved = package
            .state
            .index
            .resolve_ref_id(&package.state.components, identifier)
            .map_err(map_package_error)?
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let component = package
            .state
            .components
            .catalog()
            .get_index(resolved.component_index)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let object = component
            .archive()
            .objects
            .get(resolved.object_index)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        if object.archive_info.identifier != Some(identifier) {
            return Err(SheetImageAdjustmentsError::InvalidSource);
        }
        let has_sheet = unique_typed_message(object, SHEET_MESSAGE_TYPE)?.is_some();
        let has_form_sheet = unique_typed_message(object, FORM_BASED_SHEET_MESSAGE_TYPE)?.is_some();
        let message_type = sheet_message_type_for(has_sheet, has_form_sheet)?;
        let payload = unique_typed_message(object, message_type)?
            .ok_or(SheetImageAdjustmentsError::InvalidSource)
            .map(|(_, message)| message.data.as_slice())?;
        let identifiers = sheet_drawable_identifiers(message_type, payload, budget)?;
        for drawable_identifier in identifiers {
            if drawable_identifier != image_identifier {
                continue;
            }
            let resolved_image = package
                .state
                .index
                .resolve_ref_id(&package.state.components, drawable_identifier)
                .map_err(map_package_error)?
                .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
            let image_component = package
                .state
                .components
                .catalog()
                .get_index(resolved_image.component_index)
                .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
            let image_object = image_component
                .archive()
                .objects
                .get(resolved_image.object_index)
                .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
            let Some((message_index, image_message)) =
                unique_typed_message(image_object, IMAGE_MESSAGE_TYPE)?
            else {
                continue;
            };
            let info = validate_selected_image_metadata(image_object, message_index, budget)?;
            let metadata = image_metadata(image_message.data.as_slice(), budget)?;
            if !image_parent_metadata_is_owned(info, metadata.parent_identifier)
                || !image_data_metadata_is_owned(info, metadata.data_identifier)
                || metadata.parent_identifier != identifier
            {
                return Err(SheetImageAdjustmentsError::InvalidSource);
            }
            occurrences = occurrences
                .checked_add(1)
                .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
            if identifier == sheet_identifier {
                selected_occurrences = selected_occurrences
                    .checked_add(1)
                    .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
            }
        }
    }
    if occurrences != 1 || selected_occurrences != 1 {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    Ok(())
}

fn same_selection(left: &ImageSelection, right: &ImageSelection) -> bool {
    left.sheet_position == right.sheet_position
        && left.image_position == right.image_position
        && left.sheet_identifier == right.sheet_identifier
        && left.image_identifier == right.image_identifier
        && left.component_index == right.component_index
        && left.object_index == right.object_index
        && left.message_index == right.message_index
        && left.component_name == right.component_name
}

fn commit_edit(
    source: &Package,
    selection: &ImageSelection,
    after: ImageAdjustments,
) -> Result<SheetImageAdjustmentsCommit, SheetImageAdjustmentsError> {
    let catalog = physical_catalog(source)?;
    let source_owner = catalog.__source_owner();
    let mut budget = ImageBudget::for_package(source)?;
    budget.source(source.source_bytes().len())?;
    let current = select_image_with_budget(
        source,
        SheetSelector::position(selection.sheet_position),
        ImageSelector::position(selection.image_position),
        true,
        &mut budget,
    )?;
    if !same_selection(&current, selection) || current.before != selection.before {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    if selection.before == after {
        return Ok(SheetImageAdjustmentsCommit {
            package: source.snapshot(),
            patch: SheetImageAdjustmentsPatch {
                artifacts: OwnedExactArtifacts::new(source_owner.clone(), source_owner),
                selection: selection.clone(),
                before: selection.before,
                after,
            },
            diagnostics: SheetImageAdjustmentsDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(SheetImageAdjustmentsError::UnsupportedSource);
    }
    let candidate = rewrite_image(source, selection, after, &mut budget)?;
    let target_catalog = physical_catalog(&candidate)?;
    let target_owner = target_catalog.__source_owner();
    let selected = select_image_with_budget(
        &candidate,
        SheetSelector::position(selection.sheet_position),
        ImageSelector::position(selection.image_position),
        true,
        &mut budget,
    )?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(SheetImageAdjustmentsError::Verification);
    }
    verify_locality(source, &candidate, selection, &mut budget)?;
    Ok(SheetImageAdjustmentsCommit {
        package: candidate,
        patch: SheetImageAdjustmentsPatch {
            artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            selection: selection.clone(),
            before: selection.before,
            after,
        },
        diagnostics: SheetImageAdjustmentsDiagnostics::published(),
    })
}

fn reopen_target_patch(
    source: &Package,
    patch: &SheetImageAdjustmentsPatch,
    budget: &mut ImageBudget,
) -> Result<SheetImageAdjustmentsCommit, SheetImageAdjustmentsError> {
    let target_owner = patch.artifacts.target_owner();
    budget.candidate_reopen(source, target_owner.as_ref().len())?;
    let candidate = Package::from_source_owner_with_options(target_owner, source.state.options)
        .map_err(|_| SheetImageAdjustmentsError::Verification)?;
    let selected = select_image_with_budget(
        &candidate,
        SheetSelector::position(patch.selection.sheet_position),
        ImageSelector::position(patch.selection.image_position),
        true,
        budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(SheetImageAdjustmentsError::Verification);
    }
    verify_locality(source, &candidate, &patch.selection, budget)?;
    Ok(SheetImageAdjustmentsCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SheetImageAdjustmentsDiagnostics::published(),
    })
}

fn rewrite_image(
    source: &Package,
    selection: &ImageSelection,
    after: ImageAdjustments,
    budget: &mut ImageBudget,
) -> Result<Package, SheetImageAdjustmentsError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.component_name.as_ref())
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SheetImageAdjustmentsError::UnsupportedSource);
    }
    let component = source
        .state
        .components
        .catalog()
        .get_index(selection.component_index)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    if component.name() != selection.component_name.as_ref() {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    let source_archive = component.archive();
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
    let snappy_limits = source
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
    let object = source_archive
        .objects
        .get(selection.object_index)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    if object.archive_info.identifier != Some(selection.image_identifier) {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    let (message_index, message) = unique_typed_message(object, IMAGE_MESSAGE_TYPE)?
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    if message_index != selection.message_index {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    let _ = validate_selected_image_metadata(object, message_index, budget)?;
    let original = message.data.as_slice();
    let before = decode_with_budget(original, budget)?;
    if before != selection.before {
        return Err(SheetImageAdjustmentsError::InvalidSource);
    }
    let source_encoded_len = source_archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let options = budget.codec_options(original)?;
    let prepared =
        codec::prepare_image_adjustments_rewrite(original, write_from_adjustments(after), options)
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
        return Err(SheetImageAdjustmentsError::LimitExceeded {
            kind: SheetImageAdjustmentsLimitKind::WireBytes,
            observed: compressed_bound as u64,
            maximum: snappy_limits.max_compressed_stream() as u64,
        });
    }
    budget.preflight_output(encoded_bound)?;
    budget.preflight_output(compressed_bound)?;
    budget.preflight_work(encoded_bound.saturating_add(compressed_bound))?;
    let (clone_allocations, clone_retained) =
        archive_clone_requirements(source_archive, source_encoded_len)?;
    let (header_allocations, header_retained) = archive_header_rewrite_requirements(object)?;
    let (serialization_allocations, serialization_retained) =
        archive_serialization_requirements(source_archive)?;
    let total_allocations = clone_allocations
        .checked_add(header_allocations)
        .and_then(|amount| amount.checked_add(serialization_allocations))
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    budget.preflight_allocations(total_allocations)?;
    let temporary_retained = clone_retained
        .checked_add(header_retained)
        .and_then(|amount| amount.checked_add(serialization_retained))
        .and_then(|amount| amount.checked_add(encoded_bound))
        .and_then(|amount| amount.checked_add(compressed_bound))
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
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
        return Err(SheetImageAdjustmentsError::Verification);
    }
    archive
        .object_mut(selection.image_identifier)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?
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
        selection.component_name.as_ref(),
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
    Package::from_owned_bytes_with_options(output, source.state.options)
        .map_err(|_| SheetImageAdjustmentsError::Verification)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &ImageSelection,
    budget: &mut ImageBudget,
) -> Result<(), SheetImageAdjustmentsError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    if source_catalog.package().len() != candidate_catalog.package().len() {
        return Err(SheetImageAdjustmentsError::Verification);
    }
    for (source_entry, candidate_entry) in source_catalog
        .package()
        .iter()
        .zip(candidate_catalog.package().iter())
    {
        budget.work(source_entry.name().len().saturating_add(1))?;
        if source_entry.name() != candidate_entry.name() {
            return Err(SheetImageAdjustmentsError::Verification);
        }
        if source_entry.name() != selection.component_name.as_ref()
            && source_entry.data() != candidate_entry.data()
        {
            return Err(SheetImageAdjustmentsError::Verification);
        }
    }
    let source_archive = component_archive(source, selection.component_index)?;
    let candidate_archive = component_archive(candidate, selection.component_index)?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(SheetImageAdjustmentsError::Verification);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
    for (object_index, source_object) in source_archive.objects.iter().enumerate() {
        budget.work(source_object.messages.len())?;
        let candidate_object = candidate_archive
            .objects
            .get(object_index)
            .ok_or(SheetImageAdjustmentsError::Verification)?;
        if object_index != selection.object_index {
            if !source_object.same_content_ignoring_offsets(candidate_object) {
                return Err(SheetImageAdjustmentsError::Verification);
            }
            continue;
        }
        let candidate_message = candidate_object
            .messages
            .get(selection.message_index)
            .ok_or(SheetImageAdjustmentsError::Verification)?;
        if candidate_message.type_ != IMAGE_MESSAGE_TYPE {
            return Err(SheetImageAdjustmentsError::Verification);
        }
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
        expected
            .replace_message_preserving_header_with_limits(
                selection.message_index,
                RawMessage {
                    type_: IMAGE_MESSAGE_TYPE,
                    data: candidate_message.data.clone(),
                },
                archive_limits,
            )
            .map_err(map_core_error)?;
        expected.header_length = candidate_object.header_length;
        expected.data_length = candidate_object.data_length;
        if !expected.same_content_ignoring_offsets(candidate_object) {
            return Err(SheetImageAdjustmentsError::Verification);
        }
    }
    Ok(())
}

fn component_archive(
    package: &Package,
    component_index: usize,
) -> Result<&Archive, SheetImageAdjustmentsError> {
    package
        .state
        .components
        .catalog()
        .get_index(component_index)
        .map(|component| component.archive())
        .ok_or(SheetImageAdjustmentsError::InvalidSource)
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, SheetImageAdjustmentsError> {
    package
        .state
        .components
        .physical()
        .ok_or(SheetImageAdjustmentsError::UnsupportedSource)
}

fn rewritten_archive_bound(
    source_encoded_len: usize,
    original_len: usize,
    replacement_len: usize,
) -> Result<usize, SheetImageAdjustmentsError> {
    const MAX_REWRITE_FRAMING_GROWTH: usize = 64;
    source_encoded_len
        .checked_sub(original_len)
        .and_then(|value| value.checked_add(replacement_len))
        .and_then(|value| value.checked_add(MAX_REWRITE_FRAMING_GROWTH))
        .ok_or(SheetImageAdjustmentsError::InvalidSource)
}

fn archive_header_rewrite_requirements(
    object: &litchi_iwa_core::ArchiveObject,
) -> Result<(usize, usize), SheetImageAdjustmentsError> {
    let header = usize::try_from(object.header_length)
        .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
    let retained = header
        .checked_add(64)
        .and_then(|header| header.checked_mul(3))
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    // The core header-preserving replacement stages canonical, rewritten, and
    // post-mutation header buffers before publishing the object. Reserve that
    // bounded work before cloning the archive so a later allocation failure
    // cannot leave a partially admitted transaction.
    Ok((3, retained))
}

fn archive_serialization_requirements(
    archive: &Archive,
) -> Result<(usize, usize), SheetImageAdjustmentsError> {
    let mut retained = 0usize;
    for object in &archive.objects {
        let header = usize::try_from(object.header_length)
            .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(header)
            .and_then(|amount| amount.checked_add(64))
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    }
    // Archive serialization allocates output and Snappy staging vectors plus
    // one temporary encoded header per object. Snappy's own physical limits
    // remain enforced by the core operation and its compressed result is
    // charged separately after it returns.
    let allocations = archive
        .objects
        .len()
        .checked_add(2)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    Ok((allocations, retained))
}

fn archive_clone_requirements(
    archive: &Archive,
    encoded_len: usize,
) -> Result<(usize, usize), SheetImageAdjustmentsError> {
    // `encoded_len` covers the source payload and its retained physical
    // framing. The remaining accounting covers every decoded container and
    // metadata vector cloned by the core archive materialization.
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
        if object.header_length != 0 {
            allocations = allocations
                .checked_add(2)
                .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
            retained = retained
                .checked_add(
                    usize::try_from(object.header_length)
                        .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?
                        .checked_mul(2)
                        .ok_or(SheetImageAdjustmentsError::InvalidSource)?,
                )
                .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
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

fn add_clone_vec<T>(
    length: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), SheetImageAdjustmentsError> {
    if length == 0 {
        return Ok(());
    }
    *allocations = allocations
        .checked_add(1)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    *retained = retained
        .checked_add(
            length
                .checked_mul(size_of::<T>())
                .ok_or(SheetImageAdjustmentsError::InvalidSource)?,
        )
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    Ok(())
}

fn add_clone_bytes(
    length: usize,
    allocations: &mut usize,
    retained: &mut usize,
) -> Result<(), SheetImageAdjustmentsError> {
    add_clone_vec::<u8>(length, allocations, retained)
}

fn archive_object_clone_requirements(
    source: &litchi_iwa_core::ArchiveObject,
    candidate: &litchi_iwa_core::ArchiveObject,
    replacement_len: usize,
) -> Result<(usize, usize), SheetImageAdjustmentsError> {
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
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(
                usize::try_from(source.header_length)
                    .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?
                    .checked_mul(2)
                    .ok_or(SheetImageAdjustmentsError::InvalidSource)?,
            )
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    }
    add_clone_bytes(replacement_len, &mut allocations, &mut retained)?;
    let source_header = usize::try_from(source.header_length)
        .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
    let candidate_header = usize::try_from(candidate.header_length)
        .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
    let header_scratch = source_header
        .checked_add(candidate_header)
        .and_then(|bytes| bytes.checked_mul(3))
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    allocations = allocations
        .checked_add(3)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    retained = retained
        .checked_add(header_scratch)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    Ok((allocations, retained))
}

fn package_reopen_requirements(
    package: &Package,
    source_bytes: usize,
) -> Result<(usize, usize), SheetImageAdjustmentsError> {
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(|_| SheetImageAdjustmentsError::InvalidSource)?;
    let physical = package
        .state
        .components
        .physical()
        .ok_or(SheetImageAdjustmentsError::UnsupportedSource)?;
    let components = package.state.components.catalog();
    let entries = physical.package().iter();
    let mut allocations = 3usize; // SourceCatalog, component storage, and Package state.
    let mut retained = source_bytes;
    let mut object_count = 0usize;

    // The reopened package shares its source bytes but rebuilds the catalog,
    // component index, and semantic package state around that allocation.
    // Reserve the known entry/component/index storage before parsing it.
    for entry in entries {
        allocations = allocations
            .checked_add(1)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let entry_slots = 2usize
            .checked_mul(size_of::<usize>())
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(entry.name().len())
            .and_then(|value| value.checked_add(entry_slots))
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    }
    if !components.is_empty() {
        allocations = allocations
            .checked_add(1)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let component_slots = components
            .len()
            .checked_mul(2)
            .and_then(|count| count.checked_mul(size_of::<usize>()))
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(component_slots)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    }
    for component in components.iter() {
        allocations = allocations
            .checked_add(1)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(component.name().len())
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        object_count = object_count
            .checked_add(component.archive().objects.len())
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        let encoded_len = component
            .archive()
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let (archive_allocations, archive_retained) =
            archive_clone_requirements(component.archive(), encoded_len)?;
        allocations = allocations
            .checked_add(archive_allocations)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(archive_retained)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    }
    if object_count != 0 {
        allocations = allocations
            .checked_add(1)
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
        retained = retained
            .checked_add(
                object_count
                    .checked_mul(size_of::<(u64, usize, usize)>())
                    .ok_or(SheetImageAdjustmentsError::InvalidSource)?,
            )
            .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    }
    // Reassembly output is converted into the package's shared source owner;
    // keep a second immutable allocation covered until the temporary Vec is
    // released by the caller.
    allocations = allocations
        .checked_add(1)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    retained = retained
        .checked_add(source_bytes)
        .ok_or(SheetImageAdjustmentsError::InvalidSource)?;
    Ok((allocations, retained))
}

fn map_package_error(error: PackageError) -> SheetImageAdjustmentsError {
    match error {
        PackageError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SheetImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SheetImageAdjustmentsLimitKind::References,
                _ => SheetImageAdjustmentsLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        PackageError::Common(error) => map_common_error(error),
        _ => SheetImageAdjustmentsError::InvalidSource,
    }
}

fn map_names_error(error: super::names::Error) -> SheetImageAdjustmentsError {
    match error {
        super::names::Error::LimitExceeded {
            kind,
            observed,
            maximum,
        } => SheetImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                super::names::LimitKind::PayloadReferences => {
                    SheetImageAdjustmentsLimitKind::References
                },
                super::names::LimitKind::WireBytes
                | super::names::LimitKind::PayloadBytes
                | super::names::LimitKind::TotalPayloadBytes => {
                    SheetImageAdjustmentsLimitKind::WireBytes
                },
                super::names::LimitKind::WireFields
                | super::names::LimitKind::PayloadItems
                | super::names::LimitKind::PayloadMessages
                | super::names::LimitKind::PayloadObjects => {
                    SheetImageAdjustmentsLimitKind::WireFields
                },
                super::names::LimitKind::WireNesting => SheetImageAdjustmentsLimitKind::WireNesting,
                _ => SheetImageAdjustmentsLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        super::names::Error::Allocation { amount } => {
            SheetImageAdjustmentsError::Allocation { amount }
        },
        _ => SheetImageAdjustmentsError::InvalidSource,
    }
}

fn map_common_error(error: litchi_iwa_common::Error) -> SheetImageAdjustmentsError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SheetImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    SheetImageAdjustmentsLimitKind::WireBytes
                },
                litchi_iwa_common::LimitKind::Fields => SheetImageAdjustmentsLimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => {
                    SheetImageAdjustmentsLimitKind::OutputBytes
                },
                litchi_iwa_common::LimitKind::Nesting => {
                    SheetImageAdjustmentsLimitKind::WireNesting
                },
                litchi_iwa_common::LimitKind::RewriteWork
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    SheetImageAdjustmentsLimitKind::WireWork
                },
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SheetImageAdjustmentsError::Allocation { amount }
        },
        _ => SheetImageAdjustmentsError::InvalidSource,
    }
}

fn map_codec_error(error: codec::DecodeError) -> SheetImageAdjustmentsError {
    if let Some(limit) = error.limit_kind() {
        let (observed, maximum) = error.limit_values().unwrap_or((0, 0));
        let kind = match limit {
            codec::DecodeLimit::InputBytes => SheetImageAdjustmentsLimitKind::WireBytes,
            codec::DecodeLimit::OutputBytes | codec::DecodeLimit::Retained => {
                SheetImageAdjustmentsLimitKind::OutputBytes
            },
            codec::DecodeLimit::Fields => SheetImageAdjustmentsLimitKind::WireFields,
            codec::DecodeLimit::Nesting => SheetImageAdjustmentsLimitKind::WireNesting,
            codec::DecodeLimit::Work
            | codec::DecodeLimit::Allocations
            | codec::DecodeLimit::Scratch => SheetImageAdjustmentsLimitKind::WireWork,
            _ => SheetImageAdjustmentsLimitKind::WireWork,
        };
        return SheetImageAdjustmentsError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    SheetImageAdjustmentsError::InvalidSource
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SheetImageAdjustmentsError {
    map_common_error(error)
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SheetImageAdjustmentsError {
    match error {
        litchi_iwa_archive::Error::Limit {
            observed,
            maximum,
            kind,
        } => SheetImageAdjustmentsError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    SheetImageAdjustmentsLimitKind::WireBytes
                },
                litchi_iwa_archive::LimitKind::Entries
                | litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    SheetImageAdjustmentsLimitKind::WireFields
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SheetImageAdjustmentsLimitKind::OutputBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SheetImageAdjustmentsError::Allocation { amount }
        },
        _ => SheetImageAdjustmentsError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SheetImageAdjustmentsError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SheetImageAdjustmentsError::LimitExceeded {
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
                    SheetImageAdjustmentsLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::HeaderFields
                | litchi_iwa_core::LimitKind::MetadataItems => {
                    SheetImageAdjustmentsLimitKind::WireFields
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SheetImageAdjustmentsLimitKind::WireNesting
                },
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SheetImageAdjustmentsError::Allocation { amount: requested }
        },
        _ => SheetImageAdjustmentsError::InvalidSource,
    }
}

#[cfg(all(test, feature = "internal-iwork-source"))]
mod tests {
    use litchi_iwa_common::WireLimits;
    use litchi_iwa_common::shape::image::{ImageAdjustment, ImageAdjustments, ImageEnhancement};

    use super::{__decode_image_adjustments_payload, codec, codec_options, write_from_adjustments};

    fn rewrite_image_adjustments_payload(
        source: &[u8],
        adjustments: ImageAdjustments,
    ) -> Result<Vec<u8>, codec::DecodeError> {
        Ok(codec::rewrite_image_adjustments(
            source,
            write_from_adjustments(adjustments),
            codec_options(source, WireLimits::default()),
        )?)
    }

    #[test]
    fn focused_bridge_maps_native_controls_and_preserves_unknown_bytes() {
        let source = [
            0x0a, 0x00, // ImageArchive.super
            0x72, 0x10, // ImageArchive.imageAdjustments, length 16
            0x0d, 0x00, 0x00, 0x00, 0x00, // exposure = 0.0
            0x15, 0x00, 0x00, 0x00, 0x00, // saturation = 0.0
            0x68, 0x00, // enhance = false
            0x98, 0x06, 0xde, 0x07, // unknown nested field
            0xa0, 0x06, 0xe8, 0x07, // unknown outer field
        ];
        let baseline = __decode_image_adjustments_payload(&source, WireLimits::default()).unwrap();
        assert_eq!(baseline.exposure(), Some(ImageAdjustment::NEUTRAL));
        assert_eq!(baseline.saturation(), Some(ImageAdjustment::NEUTRAL));
        assert_eq!(baseline.enhancement(), Some(ImageEnhancement::Disabled));

        let replacement = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::new(0.25).unwrap()))
            .with_saturation(Some(ImageAdjustment::new(-0.5).unwrap()))
            .with_enhancement(Some(ImageEnhancement::Enabled));
        let changed = rewrite_image_adjustments_payload(&source, replacement).unwrap();
        assert_eq!(
            __decode_image_adjustments_payload(&changed, WireLimits::default()).unwrap(),
            replacement
        );
        assert!(
            changed
                .windows(4)
                .any(|window| window == [0x98, 0x06, 0xde, 0x07])
        );
        assert!(
            changed
                .windows(4)
                .any(|window| window == [0xa0, 0x06, 0xe8, 0x07])
        );

        let restored = rewrite_image_adjustments_payload(&changed, baseline).unwrap();
        assert_eq!(restored, source);
    }

    #[test]
    fn focused_bridge_preserves_omitted_outer_and_inner_controls() {
        let source = [0x0a, 0x00];
        let baseline = __decode_image_adjustments_payload(&source, WireLimits::default()).unwrap();
        assert_eq!(baseline, ImageAdjustments::default());
        let explicit = ImageAdjustments::new()
            .with_exposure(Some(ImageAdjustment::NEUTRAL))
            .with_saturation(Some(ImageAdjustment::NEUTRAL))
            .with_enhancement(Some(ImageEnhancement::Disabled));
        let changed = rewrite_image_adjustments_payload(&source, explicit).unwrap();
        assert_eq!(
            __decode_image_adjustments_payload(&changed, WireLimits::default()).unwrap(),
            explicit
        );
        let reset =
            rewrite_image_adjustments_payload(&changed, ImageAdjustments::default()).unwrap();
        assert_eq!(reset, source);
    }

    #[test]
    fn focused_bridge_applies_caller_limits() {
        let source = [0x0a, 0x00, 0x72, 0x00];
        let limits = WireLimits::default().with_input_bytes(1).unwrap();
        assert!(__decode_image_adjustments_payload(&source, limits).is_err());
    }
}

#[cfg(test)]
mod graph_tests {
    use litchi_iwa_common::WireLimits;
    use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, MessageInfo, RawMessage};

    use super::{
        ImageAdjustment, ImageAdjustments, ImageBudget, SheetImageAdjustmentsError, codec,
        image_data_metadata_is_owned, image_metadata, image_parent_metadata_is_owned,
        sheet_drawable_identifiers, sheet_message_type_for, validate_selected_image_metadata,
        write_from_adjustments,
    };

    fn test_budget() -> ImageBudget {
        let limits = WireLimits::default();
        ImageBudget {
            limits,
            max_input: 4096,
            max_output: 4096,
            max_fields: 512,
            max_work: 8192,
            max_nesting: limits.max_nesting(),
            max_references: 16,
            max_allocations: 16,
            max_retained: 4096,
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

    fn valid_image_payload() -> Vec<u8> {
        vec![
            0x0a, 0x04, // ImageArchive.super
            0x12, 0x02, // Drawable.parent
            0x08, 0x07, // parent identifier = 7
            0x5a, 0x02, // ImageArchive.data
            0x08, 0x63, // data identifier = 99
        ]
    }

    fn image_object_with_metadata(
        data_references: Vec<u64>,
        field_infos: Vec<FieldInfo>,
    ) -> ArchiveObject {
        let payload = valid_image_payload();
        let mut object = ArchiveObject::new(
            7,
            vec![RawMessage {
                type_: super::IMAGE_MESSAGE_TYPE,
                data: payload.clone(),
            }],
        )
        .unwrap();
        let mut info = MessageInfo::new(super::IMAGE_MESSAGE_TYPE, payload.len() as u32);
        info.data_references = data_references;
        info.field_infos = field_infos;
        object.archive_info.message_infos[0] = info;
        object
    }

    #[test]
    fn graph_keeps_unknown_data_and_image_fields_opaque() {
        // The parent is a TSP.Reference with its deprecated fields present.
        // Its unknown field 9 and the outer unknown field 20 use deliberately
        // overlong framing. The data edge is TSP.DataReference: fields 2 and
        // 3 are unknown there and must not be interpreted as parent metadata.
        let source = [
            0x0a, 0x0b, // ImageArchive.super
            0x12, 0x09, // Drawable.parent (TSP.Reference)
            0x08, 0x07, // identifier = 7
            0x10, 0x00, // deprecated_type = 0
            0x18, 0x00, // deprecated_is_external = false
            0xc8, 0x00, 0x01, // unknown parent field 9, overlong key
            0x5a, 0x08, // ImageArchive.data (TSP.DataReference)
            0x08, 0x63, // identifier = 99
            0x90, 0x00, 0x01, // unknown data field 2, overlong key
            0x18, 0x81, 0x00, // unknown data field 3, overlong value
            0xa2, 0x81, 0x00, 0x01, 0xff, // unknown outer field 20, overlong key
        ];
        let mut budget = test_budget();
        let metadata = image_metadata(&source, &mut budget).unwrap();
        assert_eq!(metadata.parent_identifier, 7);
        assert_eq!(metadata.data_identifier, 99);
    }

    #[test]
    fn image_data_metadata_accepts_aggregate_only_and_path_eleven() {
        let aggregate_only = image_object_with_metadata(vec![99], Vec::new());
        let mut budget = test_budget();
        let info = validate_selected_image_metadata(&aggregate_only, 0, &mut budget).unwrap();
        let metadata = image_metadata(&valid_image_payload(), &mut budget).unwrap();
        assert!(image_data_metadata_is_owned(info, metadata.data_identifier));

        let mut field = FieldInfo::new(vec![super::IMAGE_DATA_FIELD]);
        field.data_references = vec![99];
        let with_field = image_object_with_metadata(vec![99], vec![field]);
        let info = validate_selected_image_metadata(&with_field, 0, &mut budget).unwrap();
        assert!(image_data_metadata_is_owned(info, metadata.data_identifier));
    }

    #[test]
    fn image_data_metadata_rejects_wrong_or_duplicate_field_attribution() {
        let mut wrong_path = FieldInfo::new(vec![super::IMAGE_DATA_FIELD + 1]);
        wrong_path.data_references = vec![99];
        let wrong_path = image_object_with_metadata(vec![99], vec![wrong_path]);
        let mut budget = test_budget();
        let info = validate_selected_image_metadata(&wrong_path, 0, &mut budget).unwrap();
        assert!(!image_data_metadata_is_owned(info, 99));

        let mut duplicate_field = FieldInfo::new(vec![super::IMAGE_DATA_FIELD]);
        duplicate_field.data_references = vec![99, 99];
        let duplicate_field = image_object_with_metadata(vec![99], vec![duplicate_field]);
        let info = validate_selected_image_metadata(&duplicate_field, 0, &mut budget).unwrap();
        assert!(!image_data_metadata_is_owned(info, 99));

        let duplicate_aggregate = image_object_with_metadata(vec![99, 99], Vec::new());
        let info = validate_selected_image_metadata(&duplicate_aggregate, 0, &mut budget).unwrap();
        assert!(!image_data_metadata_is_owned(info, 99));

        let wrong_identifier = image_object_with_metadata(vec![100], Vec::new());
        let info = validate_selected_image_metadata(&wrong_identifier, 0, &mut budget).unwrap();
        assert!(!image_data_metadata_is_owned(info, 99));

        let mut first = FieldInfo::new(vec![super::IMAGE_DATA_FIELD]);
        first.data_references = vec![99];
        let mut second = FieldInfo::new(vec![super::IMAGE_DATA_FIELD]);
        second.data_references = vec![99];
        let duplicate_fields = image_object_with_metadata(vec![99], vec![first, second]);
        let info = validate_selected_image_metadata(&duplicate_fields, 0, &mut budget).unwrap();
        assert!(!image_data_metadata_is_owned(info, 99));
    }

    #[test]
    fn image_parent_metadata_accepts_native_absent_aggregate_only_and_path_one_two() {
        let native_absent = image_object_with_metadata(vec![99], Vec::new());
        let mut budget = test_budget();
        let info = validate_selected_image_metadata(&native_absent, 0, &mut budget).unwrap();
        assert!(image_parent_metadata_is_owned(info, 7));

        let mut aggregate_only = image_object_with_metadata(vec![99], Vec::new());
        aggregate_only.archive_info.message_infos[0].object_references = vec![7];
        let info = validate_selected_image_metadata(&aggregate_only, 0, &mut budget).unwrap();
        let metadata = image_metadata(&valid_image_payload(), &mut budget).unwrap();
        assert!(image_parent_metadata_is_owned(
            info,
            metadata.parent_identifier
        ));

        let mut field =
            FieldInfo::new(vec![super::IMAGE_SUPER_FIELD, super::DRAWABLE_PARENT_FIELD]);
        field.object_references = vec![7];
        let mut with_field = image_object_with_metadata(vec![99], vec![field]);
        with_field.archive_info.message_infos[0].object_references = vec![7];
        let info = validate_selected_image_metadata(&with_field, 0, &mut budget).unwrap();
        assert!(image_parent_metadata_is_owned(
            info,
            metadata.parent_identifier
        ));
    }

    #[test]
    fn image_parent_metadata_rejects_duplicate_and_wrong_attribution() {
        let mut absent = image_object_with_metadata(vec![99], Vec::new());
        let mut budget = test_budget();
        let info = validate_selected_image_metadata(&absent, 0, &mut budget).unwrap();
        assert!(image_parent_metadata_is_owned(info, 7));

        absent.archive_info.message_infos[0].object_references = vec![7, 7];
        let info = validate_selected_image_metadata(&absent, 0, &mut budget).unwrap();
        assert!(!image_parent_metadata_is_owned(info, 7));

        // Unrelated style/caption references do not declare the parent.
        absent.archive_info.message_infos[0].object_references = vec![8];
        let info = validate_selected_image_metadata(&absent, 0, &mut budget).unwrap();
        assert!(image_parent_metadata_is_owned(info, 7));

        let mut field_only =
            FieldInfo::new(vec![super::IMAGE_SUPER_FIELD, super::DRAWABLE_PARENT_FIELD]);
        field_only.object_references = vec![7];
        absent.archive_info.message_infos[0].field_infos = vec![field_only];
        let info = validate_selected_image_metadata(&absent, 0, &mut budget).unwrap();
        assert!(!image_parent_metadata_is_owned(info, 7));

        let mut wrong_path = FieldInfo::new(vec![super::IMAGE_SUPER_FIELD, 3]);
        wrong_path.object_references = vec![7];
        absent.archive_info.message_infos[0].object_references = vec![7];
        absent.archive_info.message_infos[0].field_infos = vec![wrong_path];
        let info = validate_selected_image_metadata(&absent, 0, &mut budget).unwrap();
        assert!(!image_parent_metadata_is_owned(info, 7));

        let mut duplicate_field =
            FieldInfo::new(vec![super::IMAGE_SUPER_FIELD, super::DRAWABLE_PARENT_FIELD]);
        duplicate_field.object_references = vec![7, 7];
        absent.archive_info.message_infos[0].field_infos = vec![duplicate_field];
        let info = validate_selected_image_metadata(&absent, 0, &mut budget).unwrap();
        assert!(!image_parent_metadata_is_owned(info, 7));

        absent.archive_info.message_infos[0].field_infos = Vec::new();
        absent.archive_info.message_infos[0].object_references = vec![7];
        absent.archive_info.message_infos[0].data_references = vec![7];
        let info = validate_selected_image_metadata(&absent, 0, &mut budget).unwrap();
        assert!(!image_parent_metadata_is_owned(info, 7));

        let mut data_collision =
            FieldInfo::new(vec![super::IMAGE_SUPER_FIELD, super::DRAWABLE_PARENT_FIELD]);
        data_collision.object_references = vec![7];
        data_collision.data_references = vec![7];
        absent.archive_info.message_infos[0].data_references = Vec::new();
        absent.archive_info.message_infos[0].field_infos = vec![data_collision];
        let info = validate_selected_image_metadata(&absent, 0, &mut budget).unwrap();
        assert!(!image_parent_metadata_is_owned(info, 7));
    }

    #[test]
    fn selected_image_metadata_rejects_merge_headers_atomically() {
        let cases = [
            "should_merge",
            "base_message_index",
            "diff_merge_version",
            "diff_field_path",
            "fields_to_remove",
            "diff_read_version",
        ];
        for case in cases {
            let mut object = image_object_with_metadata(vec![99], Vec::new());
            match case {
                "should_merge" => object.archive_info.should_merge = Some(true),
                "base_message_index" => {
                    object.archive_info.message_infos[0].base_message_index = Some(0)
                },
                "diff_merge_version" => {
                    object.archive_info.message_infos[0].diff_merge_version = vec![1]
                },
                "diff_field_path" => {
                    object.archive_info.message_infos[0].diff_field_path =
                        Some(vec![super::IMAGE_DATA_FIELD].into())
                },
                "fields_to_remove" => {
                    object.archive_info.message_infos[0].fields_to_remove =
                        vec![vec![super::IMAGE_DATA_FIELD].into()]
                },
                "diff_read_version" => {
                    object.archive_info.message_infos[0].diff_read_version = vec![1]
                },
                _ => unreachable!(),
            }
            let before = object.clone();
            let mut budget = test_budget();
            assert!(matches!(
                validate_selected_image_metadata(&object, 0, &mut budget),
                Err(SheetImageAdjustmentsError::InvalidSource)
            ));
            assert_eq!(object, before);
        }
    }

    #[test]
    fn sheet_graph_rejects_dual_message_models_before_projection() {
        assert_eq!(sheet_message_type_for(true, false).unwrap(), 2);
        assert_eq!(sheet_message_type_for(false, true).unwrap(), 3);
        assert!(matches!(
            sheet_message_type_for(true, true),
            Err(SheetImageAdjustmentsError::InvalidSource)
        ));
        assert!(matches!(
            sheet_message_type_for(false, false),
            Err(SheetImageAdjustmentsError::InvalidSource)
        ));
    }

    #[test]
    fn sheet_preflight_accounts_nested_reference_work_before_materialization() {
        let source = [
            0x0a, 0x01, b'S', // visible sheet name
            0x12, 0x02, 0x08, 0x07, // one local drawable reference
        ];
        let mut budget = test_budget();
        let identifiers = sheet_drawable_identifiers(2, &source, &mut budget).unwrap();
        assert_eq!(identifiers, [7]);
        assert!(budget.input > source.len());
        assert!(budget.fields >= 4);
        assert!(budget.work >= source.len());
        assert!(budget.allocations >= 3);
        assert!(budget.retained >= source.len() + size_of::<u64>());

        let mut tight = test_budget();
        tight.max_input = source.len();
        assert!(matches!(
            sheet_drawable_identifiers(2, &source, &mut tight),
            Err(SheetImageAdjustmentsError::LimitExceeded {
                kind: super::SheetImageAdjustmentsLimitKind::WireBytes,
                ..
            })
        ));
    }

    #[test]
    fn rewrite_preflight_covers_archive_clone_headers_and_codec_candidate() {
        let source_object = ArchiveObject::new(
            7,
            vec![RawMessage {
                type_: super::IMAGE_MESSAGE_TYPE,
                data: vec![0x0a, 0x00],
            }],
        )
        .unwrap();
        let candidate_object = ArchiveObject::new(
            7,
            vec![RawMessage {
                type_: super::IMAGE_MESSAGE_TYPE,
                data: vec![0x0a, 0x02, 0x72, 0x00],
            }],
        )
        .unwrap();
        let archive = Archive {
            objects: vec![source_object.clone()],
        };
        let (archive_allocations, archive_retained) =
            super::archive_clone_requirements(&archive, 128).unwrap();
        assert!(archive_allocations >= 4);
        assert!(archive_retained >= 128);

        let (header_allocations, header_retained) =
            super::archive_header_rewrite_requirements(&source_object).unwrap();
        let (serialization_allocations, serialization_retained) =
            super::archive_serialization_requirements(&archive).unwrap();
        assert_eq!(header_allocations, 3);
        assert!(header_retained >= 192);
        assert!(serialization_allocations >= 3);
        assert!(serialization_retained >= 64);

        let (object_allocations, object_retained) = super::archive_object_clone_requirements(
            &source_object,
            &candidate_object,
            candidate_object.messages[0].data.len(),
        )
        .unwrap();
        assert!(object_allocations >= 5);
        assert!(object_retained >= candidate_object.messages[0].data.len());

        let after = ImageAdjustments::new().with_exposure(Some(ImageAdjustment::new(0.5).unwrap()));
        let options = codec::DecodeOptions::new(64, 64, 512, 8).with_max_output_bytes(64);
        let prepared = codec::prepare_image_adjustments_rewrite(
            &[0x0a, 0x00],
            write_from_adjustments(after),
            options,
        )
        .unwrap();
        let requirements = prepared.execution_requirements();
        let mut budget = test_budget();
        budget.codec_requirements(requirements).unwrap();
        assert!(budget.output >= requirements.output_bytes);
        assert!(budget.fields >= requirements.fields);
        assert!(budget.work >= requirements.work_bytes);
        assert!(budget.allocations >= requirements.allocations);
    }
}
