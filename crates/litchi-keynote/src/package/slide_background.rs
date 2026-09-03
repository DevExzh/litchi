//! Exact-source, selector-first Keynote slide-background transactions.
//!
//! A slide's effective fill is the first fill-bearing style in its native
//! parent chain.  A direct override is deliberately kept separate from that
//! effective value: a variation with an empty `FillArchive` is an explicit
//! no-fill override, while a variation without a fill inherits its parent.

#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The package adapter redacts lower-layer failures at its semantic boundary."
)]

use std::collections::HashSet;
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{
    WireLimits,
    color::{RgbColorSpace, Rgba},
    decode_varint_from_bytes, encode_varint_into,
    shape::fill::{Opacity, StopMidpoint, StopPosition},
    wire::{WireFieldView, WireView},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_slide_background_codec as codec;
use litchi_iwa_protos::package_metadata_codec::{
    self as metadata_codec, Batch as MetadataBatch, ComponentSelector, ExternalReferenceAddition,
    ExternalReferenceRemoval, ObjectUuidAddition, ObjectUuidRemoval, PackageMetadataVisitor,
    RemovalBatch, RewriteOptions as MetadataRewriteOptions, UuidBits,
    inspect_package_metadata_with_visitor, remove_package_metadata, rewrite_package_metadata,
};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SLIDE_MESSAGE_TYPE, SemanticLimitKind};
use crate::{Background, Gradient, Kind, Opaque, SlideSelector, SlideSelectorError, Stop};

const SLIDE_STYLE_MESSAGE_TYPE: u32 = 9;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const GRADIENT_TYPE_FIELD: u32 = 1;
const GRADIENT_STOP_FIELD: u32 = 2;
const GRADIENT_OPACITY_FIELD: u32 = 3;
const GRADIENT_ADVANCED_FIELD: u32 = 4;
const GRADIENT_ANGLE_FIELD: u32 = 5;
const STOP_COLOR_FIELD: u32 = 1;
const STOP_POSITION_FIELD: u32 = 2;
const STOP_MIDPOINT_FIELD: u32 = 3;
const COLOR_MODEL_FIELD: u32 = 1;
const COLOR_RED_FIELD: u32 = 3;
const COLOR_GREEN_FIELD: u32 = 4;
const COLOR_BLUE_FIELD: u32 = 5;
const COLOR_ALPHA_FIELD: u32 = 6;
const COLOR_SPACE_FIELD: u32 = 12;
const ANGLE_RADIANS_FIELD: u32 = 2;
const SLIDE_STYLE_FIELD: u32 = 1;
const STYLE_NAME_FIELD: u32 = 1;
const STYLE_IDENTIFIER_FIELD: u32 = 2;
const STYLE_PARENT_FIELD: u32 = 3;
const STYLE_VARIATION_FIELD: u32 = 4;
const STYLE_STYLESHEET_FIELD: u32 = 5;
const STYLE_SUPER_FIELD: u32 = 1;
const STYLE_OVERRIDE_COUNT_FIELD: u32 = 10;
const STYLE_PROPERTIES_FIELD: u32 = 11;
const PROPERTIES_FILL_FIELD: u32 = 1;
const STYLESHEET_STYLES_FIELD: u32 = 1;
const STYLESHEET_IDENTIFIER_MAP_FIELD: u32 = 2;
const STYLESHEET_PARENT_FIELD: u32 = 3;
const STYLESHEET_CHILDREN_FIELD: u32 = 5;
const STYLESHEET_CAN_CULL_FIELD: u32 = 6;
const STYLESHEET_FIRST_VERSIONED_FIELD: u32 = 7;
const STYLESHEET_LAST_VERSIONED_FIELD: u32 = 22;
const IDENTIFIED_STYLE_REFERENCE_FIELD: u32 = 2;
const VERSIONED_STYLES_FIELD: u32 = 1;
const VERSIONED_IDENTIFIER_MAP_FIELD: u32 = 2;
const VERSIONED_CHILDREN_FIELD: u32 = 3;
const CHILD_PARENT_FIELD: u32 = 1;
const CHILDREN_FIELD: u32 = 2;

/// A finite resource governed while a slide-background transaction runs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideBackgroundLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete rewritten package output bytes.
    OutputBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Bytes in one package member, object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf scan and rewrite work.
    WireWork,
}

impl fmt::Display for SlideBackgroundLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// A content-redacted slide-background package failure.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideBackgroundError {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical slide-background edits")]
    UnsupportedSource,
    /// A checked name selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide matched the exact navigator name.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked position was outside the show.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// A name selector matched multiple slides.
    #[error("the Keynote slide-background selector is ambiguous")]
    AmbiguousSelector,
    /// The selected graph or payload is not safe to edit.
    #[error("the Keynote slide-background source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote slide-background {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideBackgroundLimitKind,
        observed: u64,
        maximum: u64,
    },
    /// A bounded allocation failed.
    #[error("could not allocate {amount} units for the Keynote slide-background transaction")]
    Allocation { amount: usize },
    /// Candidate reopening did not reproduce the requested semantic state.
    #[error("the edited Keynote slide background failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote slide-background patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable semantic slide-background edit staged against an immutable
/// package snapshot.
pub struct SlideBackgroundEdit<'a> {
    source: &'a Package,
    slide_position: Position,
    slide_identifier: u64,
    style_identifier: u64,
    before: Background,
    before_override: Option<Background>,
    desired: Desired,
}

#[derive(Debug, Clone, PartialEq)]
enum Desired {
    Set(Background),
    Reset,
}

impl fmt::Debug for SlideBackgroundEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideBackgroundEdit")
            .field("slide_position", &self.slide_position)
            .field("has_before_override", &self.before_override.is_some())
            .field("desired", &self.desired)
            .finish_non_exhaustive()
    }
}

impl<'a> SlideBackgroundEdit<'a> {
    fn new(source: &'a Package, position: Position) -> Result<Self, SlideBackgroundError> {
        let selection = select_slide(source, position)?;
        let before = selection.effective;
        Ok(Self {
            source,
            slide_position: position,
            slide_identifier: selection.slide_identifier,
            style_identifier: selection.style_identifier,
            before: before.clone(),
            before_override: selection.direct,
            desired: Desired::Set(before),
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Borrow the effective background observed when this edit began.
    #[must_use]
    pub fn before(&self) -> &Background {
        &self.before
    }

    /// Borrow the direct override observed when this edit began.
    #[must_use]
    pub fn before_override(&self) -> Option<&Background> {
        self.before_override.as_ref()
    }

    /// Borrow the staged replacement. `None` means reset inheritance.
    #[must_use]
    pub fn after(&self) -> Option<&Background> {
        match &self.desired {
            Desired::Set(value) => Some(value),
            Desired::Reset => None,
        }
    }

    /// Stage a semantic background value, including explicit [`Background::None`].
    pub fn set(mut self, background: Background) -> Result<Self, SlideBackgroundError> {
        validate_background(self.source, &background)?;
        self.desired = Desired::Set(background);
        Ok(self)
    }

    /// Stage an explicit solid fill.
    pub fn set_solid(self, color: Rgba) -> Result<Self, SlideBackgroundError> {
        self.set(Background::Solid(color))
    }

    /// Stage a validated native gradient.
    pub fn set_gradient(self, gradient: Gradient) -> Result<Self, SlideBackgroundError> {
        self.set(Background::Gradient(gradient))
    }

    /// Stage a bounded opaque native fill payload.
    pub fn set_opaque(self, opaque: Opaque) -> Result<Self, SlideBackgroundError> {
        self.set(Background::Opaque(opaque))
    }

    /// Stage removal of the direct override and restoration of inheritance.
    pub fn clear(mut self) -> Result<Self, SlideBackgroundError> {
        self.desired = Desired::Reset;
        Ok(self)
    }

    /// Alias for [`Self::clear`].
    pub fn reset(self) -> Result<Self, SlideBackgroundError> {
        self.clear()
    }

    /// Validate and publish the immutable candidate.
    pub fn commit(self) -> Result<SlideBackgroundCommit, SlideBackgroundError> {
        commit_edit(self)
    }
}

/// An exact-source-checked reversible semantic slide-background patch.
#[derive(Clone)]
pub struct SlideBackgroundPatch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    slide_identifier: u64,
    style_identifier: u64,
    target_style_identifier: u64,
    before: Background,
    before_override: Option<Background>,
    after: Option<Background>,
    target_effective: Background,
    target_override: Option<Background>,
    reset: bool,
    touched_components: usize,
    deleted_previews: usize,
}

impl fmt::Debug for SlideBackgroundPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideBackgroundPatch")
            .field("slide_position", &self.slide_position)
            .field("reset", &self.reset)
            .field("source_fingerprint", &self.artifacts.source_fingerprint())
            .field("target_fingerprint", &self.artifacts.target_fingerprint())
            .finish_non_exhaustive()
    }
}

impl SlideBackgroundPatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Borrow the effective background required from the source package.
    #[must_use]
    pub fn before(&self) -> &Background {
        &self.before
    }

    /// Borrow the effective background produced by the target package.
    #[must_use]
    pub fn after(&self) -> Option<&Background> {
        self.after.as_ref()
    }

    /// Return the direct override required from the source package.
    #[must_use]
    pub fn before_override(&self) -> Option<&Background> {
        self.before_override.as_ref()
    }

    /// Return whether the patch preserves the exact source bytes and selected
    /// slide-background semantics.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
            && self.before == self.target_effective
            && self.before_override == self.target_override
            && self.style_identifier == self.target_style_identifier
    }

    /// Return compact source provenance.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return compact target provenance.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return an exact reversible patch from the target back to the source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        let noop = self.is_noop();
        let (after, reset) = if noop {
            (self.after.clone(), self.reset)
        } else if self.before_override.is_none() {
            (None, true)
        } else {
            (Some(self.before.clone()), false)
        };
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            slide_identifier: self.slide_identifier,
            style_identifier: self.target_style_identifier,
            target_style_identifier: self.style_identifier,
            before: self.target_effective.clone(),
            before_override: self.target_override.clone(),
            after,
            target_effective: self.before.clone(),
            target_override: self.before_override.clone(),
            reset,
            touched_components: self.touched_components,
            // The exact inverse restores the prior package artifact; it does
            // not perform another preview-deletion pass.
            deleted_previews: 0,
        }
    }
}

impl PartialEq for SlideBackgroundPatch {
    fn eq(&self, other: &Self) -> bool {
        self.artifacts == other.artifacts
            && self.slide_position == other.slide_position
            && self.slide_identifier == other.slide_identifier
            && self.style_identifier == other.style_identifier
            && self.target_style_identifier == other.target_style_identifier
            && self.before == other.before
            && self.before_override == other.before_override
            && self.after == other.after
            && self.target_effective == other.target_effective
            && self.target_override == other.target_override
            && self.reset == other.reset
            && self.touched_components == other.touched_components
            && self.deleted_previews == other.deleted_previews
    }
}

/// Compact publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SlideBackgroundDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideBackgroundDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }
    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        }
    }
    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    /// Number of rewritten IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
    /// Number of root rendering previews removed from the target package.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
    /// Whether the candidate was fully reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one immutable slide-background transaction.
#[must_use = "a Keynote slide-background commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideBackgroundCommit {
    package: Package,
    patch: SlideBackgroundPatch,
    diagnostics: SlideBackgroundDiagnostics,
}

impl SlideBackgroundCommit {
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
    /// Borrow the exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideBackgroundPatch {
        &self.patch
    }
    /// Borrow compact diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideBackgroundDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the effective background, following slide-style inheritance.
    pub fn slide_background<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Background, SlideBackgroundError> {
        let position = resolve_slide_position(self, selector.into())?;
        Ok(select_slide(self, position)?.effective)
    }

    /// Read the background stored directly on the slide's variation style.
    pub fn slide_background_override<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Option<Background>, SlideBackgroundError> {
        let position = resolve_slide_position(self, selector.into())?;
        Ok(select_slide(self, position)?.direct)
    }

    /// Start an immutable exact-source slide-background edit.
    pub fn edit_slide_background<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<SlideBackgroundEdit<'_>, SlideBackgroundError> {
        let position = resolve_slide_position(self, selector.into())?;
        SlideBackgroundEdit::new(self, position)
    }

    /// Apply an exact-source-checked slide-background patch.
    pub fn apply_slide_background(
        &self,
        patch: &SlideBackgroundPatch,
    ) -> Result<SlideBackgroundCommit, SlideBackgroundError> {
        apply_patch(self, patch)
    }
}

#[derive(Debug)]
struct Selection {
    slide_identifier: u64,
    style_identifier: u64,
    slide_component: String,
    style_component: String,
    stylesheet_identifier: u64,
    stylesheet_component: String,
    style: StyleState,
    effective: Background,
    direct: Option<Background>,
}

#[derive(Debug, Clone)]
struct StyleState {
    raw: Vec<u8>,
    parent: Option<u64>,
    variation: Option<bool>,
    stylesheet: Option<u64>,
    override_count: Option<u64>,
    properties: Option<Vec<u8>>,
    fill: Option<Vec<u8>>,
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideBackgroundError> {
    match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(map_read_error)?
            .map(|_| position)
            .ok_or(SlideBackgroundError::SlidePositionNotFound { position }),
        SlideSelector::Name(name) => {
            let selector = SlideSelector::try_name(name).map_err(map_slide_selector_error)?;
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(map_slide_selector_error)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideBackgroundError::SlideNameNotFound)
        },
    }
}

fn select_slide(package: &Package, position: Position) -> Result<Selection, SlideBackgroundError> {
    let record = package
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(SlideBackgroundError::SlidePositionNotFound { position })?;
    let (slide_component, slide_object) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let slide_payload = exactly_one_payload(slide_object, SLIDE_MESSAGE_TYPE)?;
    let style_identifier = required_reference_field(package, &slide_payload, SLIDE_STYLE_FIELD)?;
    let (style_component, _) = package
        .object_with_component(style_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let style = read_style(package, style_identifier)?;
    let stylesheet_identifier = style.stylesheet.unwrap_or(0);
    let stylesheet_component = if stylesheet_identifier == 0 {
        String::new()
    } else {
        package
            .object_with_component(stylesheet_identifier)
            .map(|(component, _)| component.to_owned())
            .ok_or(SlideBackgroundError::InvalidSource)?
    };
    let direct = if style.variation == Some(true) {
        style
            .fill
            .as_deref()
            .map(|fill| decode_background(package, fill))
            .transpose()?
    } else {
        None
    };
    let effective = resolve_effective(package, style_identifier)?;
    Ok(Selection {
        slide_identifier: record.slide_identifier,
        style_identifier,
        slide_component: slide_component.to_owned(),
        style_component: style_component.to_owned(),
        stylesheet_identifier,
        stylesheet_component,
        style,
        effective,
        direct,
    })
}

fn resolve_effective(
    package: &Package,
    initial_style: u64,
) -> Result<Background, SlideBackgroundError> {
    let mut style_identifier = initial_style;
    let mut visited = HashSet::new();
    loop {
        if !visited.insert(style_identifier) {
            return Err(SlideBackgroundError::InvalidSource);
        }
        let style = read_style(package, style_identifier)?;
        if let Some(fill) = style.fill.as_deref() {
            return decode_background(package, fill);
        }
        let Some(parent) = style.parent else {
            return Ok(Background::None);
        };
        style_identifier = parent;
    }
}

fn read_style(package: &Package, identifier: u64) -> Result<StyleState, SlideBackgroundError> {
    let (_component, object) = package
        .object_with_component(identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let raw = exactly_one_payload(object, SLIDE_STYLE_MESSAGE_TYPE)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let outer = WireView::parse_with_limits(&raw, limits).map_err(map_wire_error)?;
    let super_payload = required_bytes(&outer, &raw, STYLE_SUPER_FIELD)?;
    let super_view = WireView::parse_with_limits(super_payload, limits).map_err(map_wire_error)?;
    let parent = optional_reference(&super_view, super_payload, STYLE_PARENT_FIELD, limits)?;
    let variation = optional_bool(&super_view, STYLE_VARIATION_FIELD)?;
    let stylesheet =
        optional_reference(&super_view, super_payload, STYLE_STYLESHEET_FIELD, limits)?;
    let override_count = optional_varint(&outer, STYLE_OVERRIDE_COUNT_FIELD)?;
    let properties = optional_bytes(&outer, &raw, STYLE_PROPERTIES_FIELD)?.map(ToOwned::to_owned);
    let fill = properties
        .as_deref()
        .map(|payload| optional_bytes_in_payload(payload, PROPERTIES_FILL_FIELD, limits))
        .transpose()?
        .flatten()
        .map(ToOwned::to_owned);
    Ok(StyleState {
        raw,
        parent,
        variation,
        stylesheet,
        override_count,
        properties,
        fill,
    })
}

fn decode_background(package: &Package, source: &[u8]) -> Result<Background, SlideBackgroundError> {
    let options = codec_options(package, source)?;
    let wire_limits = package.wire_limits().map_err(map_wire_error)?;
    let (snapshot, _report) =
        codec::decode_slide_background_with_report(source, options).map_err(map_codec_error)?;
    snapshot_to_background(snapshot, options, wire_limits)
}

fn snapshot_to_background(
    snapshot: codec::BackgroundSnapshot<'_>,
    options: codec::DecodeOptions,
    wire_limits: WireLimits,
) -> Result<Background, SlideBackgroundError> {
    match snapshot {
        codec::BackgroundSnapshot::None { .. } => Ok(Background::None),
        codec::BackgroundSnapshot::Opaque { raw }
        | codec::BackgroundSnapshot::Image { raw, .. } => Opaque::from_slice(raw)
            .map(Background::Opaque)
            .map_err(|_| SlideBackgroundError::InvalidSource),
        codec::BackgroundSnapshot::Solid { raw, color } => {
            rgba_from_color(color).map(Background::Solid).or_else(|_| {
                Opaque::from_slice(raw)
                    .map(Background::Opaque)
                    .map_err(|_| SlideBackgroundError::InvalidSource)
            })
        },
        codec::BackgroundSnapshot::Gradient { raw, gradient } => {
            if gradient_wire_has_only_semantic_fields(gradient, wire_limits)? {
                gradient_from_snapshot(gradient, options)
                    .map(Background::Gradient)
                    .or_else(|_| {
                        Opaque::from_slice(raw)
                            .map(Background::Opaque)
                            .map_err(|_| SlideBackgroundError::InvalidSource)
                    })
            } else {
                Opaque::from_slice(raw)
                    .map(Background::Opaque)
                    .map_err(|_| SlideBackgroundError::InvalidSource)
            }
        },
    }
}

fn gradient_wire_has_only_semantic_fields(
    gradient: codec::GradientSnapshot<'_>,
    limits: WireLimits,
) -> Result<bool, SlideBackgroundError> {
    let gradient_view =
        WireView::parse_with_limits(gradient.raw, limits).map_err(map_wire_error)?;
    if !has_only_field_numbers(
        &gradient_view,
        &[
            GRADIENT_TYPE_FIELD,
            GRADIENT_STOP_FIELD,
            GRADIENT_OPACITY_FIELD,
            GRADIENT_ADVANCED_FIELD,
            GRADIENT_ANGLE_FIELD,
        ],
    ) {
        return Ok(false);
    }

    let angle = required_bytes(&gradient_view, gradient.raw, GRADIENT_ANGLE_FIELD)?;
    let angle_view = WireView::parse_with_limits(angle, limits).map_err(map_wire_error)?;
    if !has_exact_field_numbers(&angle_view, &[ANGLE_RADIANS_FIELD]) {
        return Ok(false);
    }

    for field in gradient_view
        .fields()
        .filter(|field| field.number() == GRADIENT_STOP_FIELD)
    {
        if field.wire_type() != 2 {
            return Err(SlideBackgroundError::InvalidSource);
        }
        let stop = field.payload();
        let stop_view = WireView::parse_with_limits(stop, limits).map_err(map_wire_error)?;
        if !has_exact_field_numbers(
            &stop_view,
            &[STOP_COLOR_FIELD, STOP_POSITION_FIELD, STOP_MIDPOINT_FIELD],
        ) {
            return Ok(false);
        }
        let color = required_bytes(&stop_view, stop, STOP_COLOR_FIELD)?;
        let color_view = WireView::parse_with_limits(color, limits).map_err(map_wire_error)?;
        if !has_only_field_numbers(
            &color_view,
            &[
                COLOR_MODEL_FIELD,
                COLOR_RED_FIELD,
                COLOR_GREEN_FIELD,
                COLOR_BLUE_FIELD,
                COLOR_ALPHA_FIELD,
                COLOR_SPACE_FIELD,
            ],
        ) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn rgba_from_color(color: codec::ColorSnapshot) -> Result<Rgba, SlideBackgroundError> {
    if color.model != 1
        || color.cyan.is_some()
        || color.magenta.is_some()
        || color.yellow.is_some()
        || color.black.is_some()
        || color.white.is_some()
    {
        return Err(SlideBackgroundError::InvalidSource);
    }
    let (Some(red), Some(green), Some(blue), Some(space)) =
        (color.red, color.green, color.blue, color.rgb_space)
    else {
        return Err(SlideBackgroundError::InvalidSource);
    };
    let color_space = match space {
        1 => RgbColorSpace::Srgb,
        2 => RgbColorSpace::DisplayP3,
        _ => return Err(SlideBackgroundError::InvalidSource),
    };
    Rgba::new(red, green, blue, color.alpha.unwrap_or(1.0), color_space)
        .map_err(|_| SlideBackgroundError::InvalidSource)
}

fn gradient_from_snapshot(
    gradient: codec::GradientSnapshot<'_>,
    options: codec::DecodeOptions,
) -> Result<Gradient, SlideBackgroundError> {
    let kind = match gradient.gradient_type {
        Some(0) => Kind::Linear,
        Some(1) => Kind::Radial,
        _ => return Err(SlideBackgroundError::InvalidSource),
    };
    let opacity = Opacity::new(
        gradient
            .opacity
            .ok_or(SlideBackgroundError::InvalidSource)?,
    )
    .map_err(|_| SlideBackgroundError::InvalidSource)?;
    let advanced = gradient
        .advanced
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let angle = crate::Angle::from_radians(
        gradient
            .angle_radians
            .ok_or(SlideBackgroundError::InvalidSource)?,
    )
    .map_err(|_| SlideBackgroundError::InvalidSource)?;
    if gradient.transform.is_some() {
        return Err(SlideBackgroundError::InvalidSource);
    }
    let mut stops = Vec::new();
    for stop in gradient.stops(options) {
        let stop = stop.map_err(map_codec_error)?;
        let color = rgba_from_color(stop.color.ok_or(SlideBackgroundError::InvalidSource)?)?;
        let position = StopPosition::new(stop.fraction.ok_or(SlideBackgroundError::InvalidSource)?)
            .map_err(|_| SlideBackgroundError::InvalidSource)?;
        let midpoint =
            StopMidpoint::new(stop.inflection.ok_or(SlideBackgroundError::InvalidSource)?)
                .map_err(|_| SlideBackgroundError::InvalidSource)?;
        stops.push(Stop::new(color, position, midpoint));
    }
    Gradient::from_parts(kind, stops, opacity, advanced, angle)
        .map_err(|_| SlideBackgroundError::InvalidSource)
}

fn validate_background(
    package: &Package,
    background: &Background,
) -> Result<(), SlideBackgroundError> {
    if let Background::Opaque(opaque) = background {
        let snapshot = codec::decode_slide_background(
            opaque.as_bytes(),
            codec_options(package, opaque.as_bytes())?,
        )
        .map_err(map_codec_error)?;
        if matches!(snapshot, codec::BackgroundSnapshot::Image { .. }) {
            // ImageFillArchive carries data/object references that cannot be
            // published by a fresh style object without cloning the native
            // ownership metadata.  Refuse this branch rather than emitting
            // an unindexed resource reference.
            return Err(SlideBackgroundError::InvalidSource);
        }
        let root = WireView::parse_with_limits(opaque.as_bytes(), WireLimits::default())
            .map_err(map_wire_error)?;
        if root.fields().any(|field| field.number() == 3) {
            // A mixed FillArchive (for example color plus image) can be
            // projected as opaque by the codec, but field 3 still carries
            // ImageFillArchive resource references.
            return Err(SlideBackgroundError::InvalidSource);
        }
    }
    Ok(())
}

fn commit_edit(
    edit: SlideBackgroundEdit<'_>,
) -> Result<SlideBackgroundCommit, SlideBackgroundError> {
    let catalog = physical_catalog(edit.source)?;
    let source_bytes = catalog.shared_source();
    let source_selection = select_slide(edit.source, edit.slide_position)?;
    if source_selection.slide_identifier != edit.slide_identifier
        || source_selection.style_identifier != edit.style_identifier
        || source_selection.effective != edit.before
        || source_selection.direct != edit.before_override
    {
        return Err(SlideBackgroundError::InvalidSource);
    }
    let (after, reset) = match &edit.desired {
        Desired::Set(value) => (Some(value.clone()), false),
        Desired::Reset => (None, true),
    };
    let unchanged = if reset {
        edit.before_override.is_none()
    } else if matches!(after.as_ref(), Some(Background::None)) && edit.before_override.is_none() {
        // An empty FillArchive is an explicit no-fill override.  It is not
        // interchangeable with an inherited effective no-fill.
        false
    } else {
        after.as_ref() == Some(&edit.before)
    };
    let before = edit.before.clone();
    let before_override = edit.before_override.clone();
    if unchanged {
        edit.source.validate().map_err(map_read_error)?;
        let patch = SlideBackgroundPatch {
            artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
            slide_position: edit.slide_position,
            slide_identifier: edit.slide_identifier,
            style_identifier: edit.style_identifier,
            target_style_identifier: edit.style_identifier,
            before: before.clone(),
            before_override: before_override.clone(),
            after,
            target_effective: before.clone(),
            target_override: before_override,
            reset,
            touched_components: 0,
            deleted_previews: 0,
        };
        return Ok(SlideBackgroundCommit {
            package: edit.source.snapshot(),
            patch,
            diagnostics: SlideBackgroundDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(SlideBackgroundError::UnsupportedSource);
    }
    edit.source.validate().map_err(map_read_error)?;
    let reset_effective = if reset {
        let parent = source_selection
            .style
            .parent
            .ok_or(SlideBackgroundError::InvalidSource)?;
        Some(resolve_effective(edit.source, parent)?)
    } else {
        None
    };
    let (candidate, touched, deleted_previews) =
        rewrite_package(edit.source, &source_selection, after.as_ref(), reset)?;
    candidate.validate().map_err(map_read_error)?;
    let candidate_selection = select_slide(&candidate, edit.slide_position)?;
    let candidate_ok = if reset {
        candidate_selection.direct.is_none()
            && candidate_selection.effective
                == reset_effective
                    .clone()
                    .ok_or(SlideBackgroundError::Verification)?
    } else {
        candidate_selection.direct.as_ref() == after.as_ref()
            && candidate_selection.effective
                == after
                    .as_ref()
                    .ok_or(SlideBackgroundError::Verification)?
                    .clone()
    };
    if !candidate_ok || candidate_selection.slide_identifier != edit.slide_identifier {
        return Err(SlideBackgroundError::Verification);
    }
    let target_effective = candidate_selection.effective.clone();
    let target = physical_catalog(&candidate)?.shared_source();
    let patch = SlideBackgroundPatch {
        artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
        slide_position: edit.slide_position,
        slide_identifier: edit.slide_identifier,
        style_identifier: edit.style_identifier,
        target_style_identifier: candidate_selection.style_identifier,
        before,
        before_override,
        after,
        target_effective,
        target_override: candidate_selection.direct,
        reset,
        touched_components: touched,
        deleted_previews,
    };
    Ok(SlideBackgroundCommit {
        package: candidate,
        patch,
        diagnostics: SlideBackgroundDiagnostics::published(touched, deleted_previews),
    })
}

fn apply_patch(
    source: &Package,
    patch: &SlideBackgroundPatch,
) -> Result<SlideBackgroundCommit, SlideBackgroundError> {
    let catalog = physical_catalog(source)?;
    let source_bytes = catalog.shared_source();
    if !patch.artifacts.authorizes_source(&source_bytes) {
        return Err(SlideBackgroundError::PatchConflict);
    }
    let selection = select_slide(source, patch.slide_position)?;
    if selection.slide_identifier != patch.slide_identifier
        || selection.style_identifier != patch.style_identifier
        || selection.effective != patch.before
        || selection.direct != patch.before_override
    {
        return Err(SlideBackgroundError::PatchConflict);
    }
    if patch.is_noop() {
        source.validate().map_err(map_read_error)?;
        return Ok(SlideBackgroundCommit {
            package: source.snapshot(),
            patch: patch.clone(),
            diagnostics: SlideBackgroundDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(SlideBackgroundError::PatchConflict);
    }
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    let candidate_selection = select_slide(&candidate, patch.slide_position)?;
    let valid = candidate_selection.direct == patch.target_override
        && candidate_selection.effective == patch.target_effective;
    if !valid
        || candidate_selection.slide_identifier != patch.slide_identifier
        || candidate_selection.style_identifier != patch.target_style_identifier
    {
        return Err(SlideBackgroundError::Verification);
    }
    Ok(SlideBackgroundCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideBackgroundDiagnostics::published(
            patch.touched_components,
            patch.deleted_previews,
        ),
    })
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, SlideBackgroundError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideBackgroundError::UnsupportedSource),
    }
}

fn rewrite_package(
    source: &Package,
    selection: &Selection,
    after: Option<&Background>,
    reset: bool,
) -> Result<(Package, usize, usize), SlideBackgroundError> {
    let catalog = physical_catalog(source)?;
    if selection.stylesheet_identifier == 0 || selection.stylesheet_component.is_empty() {
        return Err(SlideBackgroundError::InvalidSource);
    }
    if selection.style_component != selection.stylesheet_component {
        return Err(SlideBackgroundError::InvalidSource);
    }
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let metadata_component = metadata_component_name(source)?;
    let mut names = vec![
        selection.slide_component.as_str(),
        selection.stylesheet_component.as_str(),
    ];
    if let Some(name) = metadata_component.as_deref() {
        names.push(name);
    }
    let mut archives = Vec::new();
    for name in names {
        if archives
            .iter()
            .any(|(existing, _): &(String, Archive)| existing == name)
        {
            continue;
        }
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == name)
            .ok_or(SlideBackgroundError::InvalidSource)?;
        if entry.is_opaque() {
            return Err(SlideBackgroundError::InvalidSource);
        }
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
        let archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        archives.push((name.to_owned(), archive));
    }
    let new_id = next_object_identifier(source)?;
    let metadata_change = metadata_mutation_for(source, selection, new_id, reset)?;
    if reset {
        mutate_reset(source, selection, &mut archives, new_id, archive_limits)?;
    } else {
        let desired = after.ok_or(SlideBackgroundError::InvalidSource)?;
        mutate_set(
            source,
            selection,
            desired,
            &mut archives,
            new_id,
            archive_limits,
        )?;
    }
    if let Some(name) = metadata_component.as_deref() {
        rewrite_metadata_for_change(
            source,
            archive_mut(&mut archives, name)?,
            selection,
            metadata_change,
            archive_limits,
        )?;
    }
    let mut compressed = Vec::new();
    for (name, archive) in &archives {
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let data = SnappyStream::compress(&bytes).map_err(map_core_error)?;
        compressed.push((name.clone(), data));
    }
    let edits = compressed
        .iter()
        .map(|(name, bytes)| EntryEdit::new(name.as_str(), bytes.as_slice()))
        .collect::<Vec<_>>();
    let plan = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(map_rendering_error)?;
    let output = catalog
        .package()
        .reassemble_with_deletions_to_bytes(&edits, plan.names(), source.state.options.archive())
        .map_err(map_archive_error)?;
    let touched = archives.len();
    let deleted_previews = plan.len();
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((candidate, touched, deleted_previews))
}

fn metadata_mutation_for(
    source: &Package,
    selection: &Selection,
    new_id: u64,
    reset: bool,
) -> Result<MetadataMutation, SlideBackgroundError> {
    if reset {
        if selection.style.variation != Some(true)
            || selection.style.parent.is_none()
            || selection.style.fill.is_none()
        {
            return Err(SlideBackgroundError::InvalidSource);
        }
        let collapsible = style_is_collapsible(source, selection)?;
        let exclusive = style_is_exclusive(source, selection)?;
        if collapsible {
            return Ok(MetadataMutation::Replace {
                old_style: selection.style_identifier,
                new_style: selection
                    .style
                    .parent
                    .ok_or(SlideBackgroundError::InvalidSource)?,
                new_last_identifier: new_id,
                remove_uuid: exclusive,
                add_uuid: false,
            });
        }
        if !exclusive {
            return Ok(MetadataMutation::Replace {
                old_style: selection.style_identifier,
                new_style: new_id,
                new_last_identifier: new_id,
                remove_uuid: false,
                add_uuid: true,
            });
        }
        Ok(MetadataMutation::None)
    } else {
        let disposable =
            style_is_collapsible(source, selection)? && style_is_exclusive(source, selection)?;
        Ok(MetadataMutation::Replace {
            old_style: selection.style_identifier,
            new_style: new_id,
            new_last_identifier: new_id,
            remove_uuid: disposable,
            add_uuid: true,
        })
    }
}

fn next_object_identifier(package: &Package) -> Result<u64, SlideBackgroundError> {
    let mut maximum = 0_u64;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            maximum = maximum.max(
                object
                    .archive_info
                    .identifier
                    .ok_or(SlideBackgroundError::InvalidSource)?,
            );
            for message in &object.messages {
                if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                    continue;
                }
                let view =
                    WireView::parse_with_limits(&message.data, limits).map_err(map_wire_error)?;
                if let Some(last) = optional_varint(&view, 1)? {
                    maximum = maximum.max(last);
                }
                maximum = maximum.max(metadata_identifier_max(&message.data, limits)?);
            }
        }
    }
    maximum
        .checked_add(1)
        .ok_or(SlideBackgroundError::InvalidSource)
}

/// Return the greatest object/component/data identifier carried by the
/// PackageMetadata registries.  The watermark is not authoritative for
/// partially-saved native packages, so allocation must also account for
/// registry records and cross-component references.
fn metadata_identifier_max(source: &[u8], limits: WireLimits) -> Result<u64, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut maximum = 0_u64;
    for field in view.fields() {
        match field.number() {
            1 => maximum = maximum.max(metadata_field_varint(&field)?),
            3 | 11 => {
                maximum = maximum.max(metadata_component_identifier_max(
                    metadata_field_bytes(&field)?,
                    limits,
                )?);
            },
            4 => {
                maximum = maximum.max(metadata_data_info_identifier(
                    metadata_field_bytes(&field)?,
                    limits,
                )?);
            },
            10 => {
                maximum = maximum.max(metadata_reference_identifier(
                    metadata_field_bytes(&field)?,
                    limits,
                )?);
            },
            _ => {},
        }
    }
    Ok(maximum)
}

fn metadata_component_identifier_max(
    source: &[u8],
    limits: WireLimits,
) -> Result<u64, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut maximum = 0_u64;
    for field in view.fields() {
        match field.number() {
            1 | 21 => maximum = maximum.max(metadata_field_varint(&field)?),
            20 => maximum = maximum.max(metadata_field_varint_or_packed(&field)?),
            6 | 18 => {
                maximum = maximum.max(metadata_external_identifier(
                    metadata_field_bytes(&field)?,
                    limits,
                )?);
            },
            7 => {
                maximum = maximum.max(metadata_data_reference_identifier(
                    metadata_field_bytes(&field)?,
                    limits,
                )?);
            },
            11 => {
                maximum = maximum.max(metadata_uuid_identifier(
                    metadata_field_bytes(&field)?,
                    limits,
                )?);
            },
            _ => {},
        }
    }
    Ok(maximum)
}

fn metadata_external_identifier(
    source: &[u8],
    limits: WireLimits,
) -> Result<u64, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut maximum = 0_u64;
    for field in view.fields() {
        match field.number() {
            1 | 2 => maximum = maximum.max(metadata_field_varint(&field)?),
            _ => {},
        }
    }
    Ok(maximum)
}

fn metadata_data_reference_identifier(
    source: &[u8],
    limits: WireLimits,
) -> Result<u64, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut maximum = 0_u64;
    for field in view.fields() {
        match field.number() {
            1 => maximum = maximum.max(metadata_field_varint(&field)?),
            2 => {
                let nested = WireView::parse_with_limits(metadata_field_bytes(&field)?, limits)
                    .map_err(map_wire_error)?;
                for nested_field in nested.fields().filter(|field| field.number() == 1) {
                    maximum = maximum.max(metadata_field_varint(&nested_field)?);
                }
            },
            _ => {},
        }
    }
    Ok(maximum)
}

fn metadata_uuid_identifier(
    source: &[u8],
    limits: WireLimits,
) -> Result<u64, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    view.fields()
        .filter(|field| field.number() == 1)
        .map(|field| metadata_field_varint(&field))
        .next()
        .transpose()
        .map(|value| value.unwrap_or(0))
}

fn metadata_data_info_identifier(
    source: &[u8],
    limits: WireLimits,
) -> Result<u64, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    view.fields()
        .filter(|field| field.number() == 1)
        .map(|field| metadata_field_varint(&field))
        .next()
        .transpose()
        .map(|value| value.unwrap_or(0))
}

fn metadata_reference_identifier(
    source: &[u8],
    limits: WireLimits,
) -> Result<u64, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    view.fields()
        .filter(|field| field.number() == 1)
        .map(|field| metadata_field_varint(&field))
        .next()
        .transpose()
        .map(|value| value.unwrap_or(0))
}

fn metadata_field_bytes<'a>(field: &WireFieldView<'a>) -> Result<&'a [u8], SlideBackgroundError> {
    if field.wire_type() != 2 {
        return Err(SlideBackgroundError::InvalidSource);
    }
    field.validate_canonical_framing().map_err(map_wire_error)?;
    Ok(field.payload())
}

fn metadata_field_varint(field: &WireFieldView<'_>) -> Result<u64, SlideBackgroundError> {
    if field.wire_type() != 0 {
        return Err(SlideBackgroundError::InvalidSource);
    }
    field.validate_canonical_key().map_err(map_wire_error)?;
    canonical_varint(field.payload())
}

fn metadata_field_varint_or_packed(field: &WireFieldView<'_>) -> Result<u64, SlideBackgroundError> {
    if field.wire_type() == 0 {
        return metadata_field_varint(field);
    }
    let payload = metadata_field_bytes(field)?;
    let mut maximum = 0_u64;
    let mut remaining = payload;
    while !remaining.is_empty() {
        let (value, used) =
            decode_varint_from_bytes(remaining).map_err(|_| SlideBackgroundError::InvalidSource)?;
        maximum = maximum.max(value);
        remaining = &remaining[used..];
    }
    Ok(maximum)
}

fn metadata_component_name(package: &Package) -> Result<Option<String>, SlideBackgroundError> {
    let mut selected = None;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object
                .messages
                .iter()
                .any(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
            {
                if selected.replace(component.name().to_owned()).is_some() {
                    return Err(SlideBackgroundError::InvalidSource);
                }
            }
        }
    }
    Ok(selected)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MetadataMutation {
    None,
    Replace {
        old_style: u64,
        new_style: u64,
        new_last_identifier: u64,
        remove_uuid: bool,
        add_uuid: bool,
    },
}

#[derive(Debug, Clone)]
struct MetadataComponent {
    identifier: u64,
    locator: String,
}

#[derive(Debug)]
struct MetadataFacts {
    slide_locator: String,
    stylesheet_locator: String,
    stylesheet_identifier: u64,
    old_style: u64,
    slide_components: Vec<MetadataComponent>,
    stylesheet_components: Vec<MetadataComponent>,
    old_uuids: Vec<UuidBits>,
    old_external_weaknesses: Vec<Option<bool>>,
    uuids: HashSet<(u64, u64)>,
}

impl MetadataFacts {
    fn new(selection: &Selection) -> Self {
        Self {
            slide_locator: metadata_locator(&selection.slide_component).to_owned(),
            stylesheet_locator: metadata_locator(&selection.stylesheet_component).to_owned(),
            stylesheet_identifier: selection.stylesheet_identifier,
            old_style: selection.style_identifier,
            slide_components: Vec::new(),
            stylesheet_components: Vec::new(),
            old_uuids: Vec::new(),
            old_external_weaknesses: Vec::new(),
            uuids: HashSet::new(),
        }
    }

    fn selectors(
        &self,
    ) -> Result<(ComponentSelector<'_>, ComponentSelector<'_>), SlideBackgroundError> {
        let slide = self
            .slide_components
            .first()
            .ok_or(SlideBackgroundError::InvalidSource)?;
        let stylesheet = self
            .stylesheet_components
            .first()
            .ok_or(SlideBackgroundError::InvalidSource)?;
        if self.slide_components.len() != 1
            || self.stylesheet_components.len() != 1
            || stylesheet.identifier != self.stylesheet_identifier
        {
            return Err(SlideBackgroundError::InvalidSource);
        }
        Ok((
            ComponentSelector::new(slide.identifier, &slide.locator),
            ComponentSelector::new(stylesheet.identifier, &stylesheet.locator),
        ))
    }

    fn old_uuid(&self) -> Result<UuidBits, SlideBackgroundError> {
        if self.old_uuids.len() != 1 {
            return Err(SlideBackgroundError::InvalidSource);
        }
        Ok(self.old_uuids[0])
    }

    fn old_external_weakness(&self) -> Result<Option<bool>, SlideBackgroundError> {
        if self.old_external_weaknesses.len() != 1 {
            return Err(SlideBackgroundError::InvalidSource);
        }
        Ok(self.old_external_weaknesses[0])
    }
}

impl PackageMetadataVisitor for MetadataFacts {
    fn visit_component(
        &mut self,
        component: metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        let locator = component.effective_locator();
        let value = MetadataComponent {
            identifier: component.identifier(),
            locator: locator.to_owned(),
        };
        if component.is_current() && locator == self.slide_locator {
            self.slide_components.push(value.clone());
        }
        if component.is_current() && locator == self.stylesheet_locator {
            self.stylesheet_components.push(value);
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        let uuid = binding.uuid();
        self.uuids.insert((uuid.lower(), uuid.upper()));
        let component = binding.component();
        if component.is_current()
            && component.identifier() == self.stylesheet_identifier
            && component.effective_locator() == self.stylesheet_locator
            && binding.object_identifier() == self.old_style
        {
            self.old_uuids.push(uuid);
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        let source = reference.source();
        if source.is_current()
            && source.effective_locator() == self.slide_locator
            && reference.target_component_identifier() == self.stylesheet_identifier
            && reference.object_identifier() == Some(self.old_style)
            && !reference.is_versioned()
        {
            self.old_external_weaknesses.push(reference.is_weak());
        }
        Ok(())
    }
}

#[derive(Debug)]
struct MetadataOwnershipFacts {
    stylesheet_identifier: u64,
    style_identifier: u64,
    selected_slide_locator: String,
    selected: usize,
    other: bool,
}

impl PackageMetadataVisitor for MetadataOwnershipFacts {
    fn visit_external_reference(
        &mut self,
        reference: metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), metadata_codec::RewriteError> {
        if reference.target_component_identifier() != self.stylesheet_identifier
            || reference.object_identifier() != Some(self.style_identifier)
        {
            return Ok(());
        }
        if reference.source().is_current()
            && reference.source().effective_locator() == self.selected_slide_locator
            && !reference.is_versioned()
        {
            self.selected = self.selected.saturating_add(1);
        } else {
            self.other = true;
        }
        Ok(())
    }
}

fn metadata_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

fn metadata_options(
    package: &Package,
    max_additions: usize,
) -> Result<MetadataRewriteOptions, SlideBackgroundError> {
    let wire = package.wire_limits().map_err(map_wire_error)?;
    let recursion_limit = u32::try_from(wire.max_nesting()).unwrap_or(u32::MAX);
    let semantic = package.semantic_limits();
    Ok(MetadataRewriteOptions::new(
        wire.max_input_bytes(),
        wire.max_output_bytes(),
        wire.max_fields(),
        wire.max_rewrite_work(),
        recursion_limit,
        semantic.max_objects(),
        semantic.max_references(),
        max_additions,
    ))
}

fn metadata_payload(archive: &Archive) -> Result<(u64, usize, Vec<u8>), SlideBackgroundError> {
    let mut selected = None;
    for object in &archive.objects {
        let identifier = object
            .archive_info
            .identifier
            .ok_or(SlideBackgroundError::InvalidSource)?;
        for (index, message) in object.messages.iter().enumerate() {
            if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                continue;
            }
            if selected
                .replace((identifier, index, message.data.clone()))
                .is_some()
            {
                return Err(SlideBackgroundError::InvalidSource);
            }
        }
    }
    selected.ok_or(SlideBackgroundError::InvalidSource)
}

fn replace_metadata_payload(
    archive: &mut Archive,
    object_identifier: u64,
    index: usize,
    data: Vec<u8>,
    limits: litchi_iwa_core::Limits,
) -> Result<(), SlideBackgroundError> {
    let object = archive
        .object_mut(object_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    validate_rewrite_metadata(object, index)?;
    object
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

fn metadata_facts(
    package: &Package,
    source: &[u8],
    selection: &Selection,
) -> Result<(MetadataFacts, u64), SlideBackgroundError> {
    let options = metadata_options(package, 8)?;
    let mut facts = MetadataFacts::new(selection);
    let inspection = inspect_package_metadata_with_visitor(source, options, &mut facts)
        .map_err(map_metadata_error)?;
    facts.selectors()?;
    Ok((facts, inspection.last_object_identifier()))
}

fn fresh_uuid(facts: &MetadataFacts, identifier: u64) -> UuidBits {
    let mut lower = identifier.max(1);
    let mut upper = identifier.rotate_left(17) ^ 0x9e37_79b9_7f4a_7c15;
    while facts.uuids.contains(&(lower, upper)) || (lower == 0 && upper == 0) {
        lower = lower.wrapping_add(1).max(1);
        upper = upper.rotate_left(7) ^ 0xd1b5_4a32_d192_ed03;
    }
    UuidBits::new(lower, upper)
}

fn rewrite_metadata_for_change(
    package: &Package,
    archive: &mut Archive,
    selection: &Selection,
    change: MetadataMutation,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<(), SlideBackgroundError> {
    if !matches!(change, MetadataMutation::None) {
        let (object_identifier, index, source) = metadata_payload(archive)?;
        let (facts, last) = metadata_facts(package, &source, selection)?;
        let (slide_selector, stylesheet_selector) = facts.selectors()?;
        let mut metadata = source;
        let mut removals_uuid = Vec::new();
        let mut removals_external = Vec::new();
        let mut additions_uuid = Vec::new();
        let mut additions_external = Vec::new();
        let (old_style, new_style, new_last_identifier, remove_uuid, add_uuid) = match change {
            MetadataMutation::Replace {
                old_style,
                new_style,
                new_last_identifier,
                remove_uuid,
                add_uuid,
            } => (
                old_style,
                new_style,
                new_last_identifier,
                remove_uuid,
                add_uuid,
            ),
            MetadataMutation::None => unreachable!(),
        };
        let old_weak = facts.old_external_weakness()?;
        removals_external.push(ExternalReferenceRemoval::new(
            slide_selector,
            stylesheet_selector,
            old_style,
            old_weak,
        ));
        if remove_uuid {
            removals_uuid.push(ObjectUuidRemoval::new(
                stylesheet_selector,
                old_style,
                facts.old_uuid()?,
            ));
        }
        if !removals_uuid.is_empty() || !removals_external.is_empty() {
            let options = metadata_options(package, 8)?;
            metadata = remove_package_metadata(
                &metadata,
                RemovalBatch::new(last, &removals_uuid, &removals_external, &[]),
                options,
            )
            .map_err(map_metadata_error)?
            .into_bytes();
        }
        if add_uuid {
            let uuid = fresh_uuid(&facts, new_style);
            additions_uuid.push(ObjectUuidAddition::new(
                stylesheet_selector,
                new_style,
                uuid,
            ));
        }
        additions_external.push(ExternalReferenceAddition::new(
            slide_selector,
            stylesheet_selector,
            new_style,
            old_weak,
        ));
        let options = metadata_options(package, 8)?;
        metadata = rewrite_package_metadata(
            &metadata,
            MetadataBatch::new(
                last,
                new_last_identifier,
                &additions_uuid,
                &additions_external,
            ),
            options,
        )
        .map_err(map_metadata_error)?
        .into_bytes();
        replace_metadata_payload(archive, object_identifier, index, metadata, archive_limits)?;
    }
    Ok(())
}

fn new_style_object(
    package: &Package,
    source_identifier: u64,
    identifier: u64,
    data: Vec<u8>,
    parent: u64,
    stylesheet: u64,
) -> Result<ArchiveObject, SlideBackgroundError> {
    let source = package
        .object(source_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    if source.messages.len() != 1
        || source.messages[0].type_ != SLIDE_STYLE_MESSAGE_TYPE
        || source.archive_info.message_infos.len() != 1
    {
        // A replacement style has a new object identity. Unknown companion
        // messages cannot be copied safely without knowing whether their
        // internal references are identity-relative, so fail closed instead
        // of silently dropping or misbinding them.
        return Err(SlideBackgroundError::InvalidSource);
    }
    let source_info = &source.archive_info.message_infos[0];
    validate_rewrite_metadata(source, 0)?;
    if !source_info.data_references.is_empty() {
        return Err(SlideBackgroundError::InvalidSource);
    }
    let source_style = parse_style_metadata(package, &source.messages[0].data)?;
    if source_style.stylesheet != Some(stylesheet)
        || source_info
            .object_references
            .iter()
            .any(|reference| *reference != stylesheet && Some(*reference) != source_style.parent)
    {
        return Err(SlideBackgroundError::InvalidSource);
    }
    for field in &source_info.field_infos {
        if !field.data_references.is_empty() {
            return Err(SlideBackgroundError::InvalidSource);
        }
        let path = field.path.path.as_slice();
        let valid = if path == [STYLE_SUPER_FIELD, STYLE_PARENT_FIELD] {
            field
                .object_references
                .iter()
                .all(|reference| Some(*reference) == source_style.parent)
        } else if path == [STYLE_SUPER_FIELD, STYLE_STYLESHEET_FIELD] {
            field
                .object_references
                .iter()
                .all(|reference| *reference == stylesheet)
        } else {
            field.object_references.is_empty()
        };
        if !valid {
            return Err(SlideBackgroundError::InvalidSource);
        }
    }

    let mut replacement = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: SLIDE_STYLE_MESSAGE_TYPE,
            data,
        }],
    )
    .map_err(map_core_error)?;
    let info = &mut replacement.archive_info.message_infos[0];
    info.versions.clone_from(&source_info.versions);
    info.field_infos.clone_from(&source_info.field_infos);
    info.object_references.extend([parent, stylesheet]);
    for field in &mut info.field_infos {
        let path = field.path.path.as_slice();
        if path == [STYLE_SUPER_FIELD, STYLE_PARENT_FIELD]
            || path == [STYLE_SUPER_FIELD, STYLE_STYLESHEET_FIELD]
        {
            field.object_references.clear();
            field.data_references.clear();
            if path[1] == STYLE_PARENT_FIELD {
                field.object_references.push(parent);
            } else {
                field.object_references.push(stylesheet);
            }
        }
    }
    Ok(replacement)
}

fn mutate_set(
    source: &Package,
    selection: &Selection,
    desired: &Background,
    archives: &mut [(String, Archive)],
    new_id: u64,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<(), SlideBackgroundError> {
    let wire_limits = source.wire_limits().map_err(map_wire_error)?;
    let old_style = &selection.style;
    let stylesheet_id = selection.stylesheet_identifier;
    let disposable =
        style_is_collapsible(source, selection)? && style_is_exclusive(source, selection)?;
    let parent = if disposable {
        old_style
            .parent
            .ok_or(SlideBackgroundError::InvalidSource)?
    } else {
        selection.style_identifier
    };
    let fill = encode_background(source, old_style.fill.as_deref().unwrap_or(&[]), desired)?;
    let style_data =
        new_style_data_from_style(old_style, parent, stylesheet_id, &fill, wire_limits)?;
    let style_object = new_style_object(
        source,
        selection.style_identifier,
        new_id,
        style_data,
        parent,
        stylesheet_id,
    )?;

    let slide_archive = archive_mut(archives, &selection.slide_component)?;
    patch_slide_style(
        slide_archive,
        selection.slide_identifier,
        selection.style_identifier,
        new_id,
        wire_limits,
        archive_limits,
    )?;
    let stylesheet_archive = archive_mut(archives, &selection.stylesheet_component)?;
    patch_stylesheet(
        stylesheet_archive,
        stylesheet_id,
        disposable.then_some(selection.style_identifier),
        Some((parent, new_id)),
        wire_limits,
        archive_limits,
    )?;
    if disposable {
        stylesheet_archive
            .remove_object(selection.style_identifier)
            .ok_or(SlideBackgroundError::InvalidSource)?;
    }
    stylesheet_archive
        .insert_object_with_limits(style_object, archive_limits)
        .map_err(map_core_error)?;
    Ok(())
}

fn mutate_reset(
    source: &Package,
    selection: &Selection,
    archives: &mut [(String, Archive)],
    new_id: u64,
    archive_limits: litchi_iwa_core::Limits,
) -> Result<(), SlideBackgroundError> {
    let wire_limits = source.wire_limits().map_err(map_wire_error)?;
    let style = &selection.style;
    if style.variation != Some(true) || style.parent.is_none() || style.fill.is_none() {
        return Err(SlideBackgroundError::InvalidSource);
    }
    let parent = style.parent.ok_or(SlideBackgroundError::InvalidSource)?;
    validate_style_background_reset_metadata(source, selection)?;
    let collapsible = style_is_collapsible(source, selection)?;
    let exclusive = style_is_exclusive(source, selection)?;
    let stylesheet_id = selection.stylesheet_identifier;
    if collapsible {
        let slide_archive = archive_mut(archives, &selection.slide_component)?;
        patch_slide_style(
            slide_archive,
            selection.slide_identifier,
            selection.style_identifier,
            parent,
            wire_limits,
            archive_limits,
        )?;
        if exclusive {
            let stylesheet_archive = archive_mut(archives, &selection.stylesheet_component)?;
            patch_stylesheet(
                stylesheet_archive,
                stylesheet_id,
                Some(selection.style_identifier),
                None,
                wire_limits,
                archive_limits,
            )?;
            stylesheet_archive
                .remove_object(selection.style_identifier)
                .ok_or(SlideBackgroundError::InvalidSource)?;
        }
        return Ok(());
    }
    let replacement_data = style_without_background(style, wire_limits)?;
    if exclusive {
        let stylesheet_archive = archive_mut(archives, &selection.stylesheet_component)?;
        replace_style_data(
            stylesheet_archive,
            selection.style_identifier,
            replacement_data,
            archive_limits,
        )?;
    } else {
        let replacement = new_style_object(
            source,
            selection.style_identifier,
            new_id,
            replacement_data,
            parent,
            stylesheet_id,
        )?;
        let slide_archive = archive_mut(archives, &selection.slide_component)?;
        patch_slide_style(
            slide_archive,
            selection.slide_identifier,
            selection.style_identifier,
            new_id,
            wire_limits,
            archive_limits,
        )?;
        let stylesheet_archive = archive_mut(archives, &selection.stylesheet_component)?;
        patch_stylesheet(
            stylesheet_archive,
            stylesheet_id,
            None,
            Some((parent, new_id)),
            wire_limits,
            archive_limits,
        )?;
        stylesheet_archive
            .insert_object_with_limits(replacement, archive_limits)
            .map_err(map_core_error)?;
    }
    Ok(())
}

fn archive_mut<'a>(
    archives: &'a mut [(String, Archive)],
    name: &str,
) -> Result<&'a mut Archive, SlideBackgroundError> {
    archives
        .iter_mut()
        .find(|(candidate, _)| candidate == name)
        .map(|(_, archive)| archive)
        .ok_or(SlideBackgroundError::InvalidSource)
}

fn validate_slide_style_metadata_for_rewrite(
    object: &ArchiveObject,
    message_index: usize,
    old_style: u64,
) -> Result<(), SlideBackgroundError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    if info
        .object_references
        .iter()
        .filter(|reference| **reference == old_style)
        .count()
        != 1
    {
        return Err(SlideBackgroundError::InvalidSource);
    }
    if info.data_references.contains(&old_style) {
        return Err(SlideBackgroundError::InvalidSource);
    }

    let mut canonical_references = 0usize;
    for field in &info.field_infos {
        if field.path.path.as_slice() == [SLIDE_STYLE_FIELD]
            && field
                .r#type
                .is_some_and(|field_type| field_type != FieldType::ObjectReference)
        {
            return Err(SlideBackgroundError::InvalidSource);
        }
        let references = field
            .object_references
            .iter()
            .filter(|reference| **reference == old_style)
            .count();
        if references == 0 {
            continue;
        }
        if field.path.path.as_slice() != [SLIDE_STYLE_FIELD] {
            return Err(SlideBackgroundError::InvalidSource);
        }
        if field.data_references.contains(&old_style) {
            return Err(SlideBackgroundError::InvalidSource);
        }
        canonical_references = canonical_references
            .checked_add(references)
            .ok_or(SlideBackgroundError::InvalidSource)?;
    }
    if canonical_references > 1 {
        return Err(SlideBackgroundError::InvalidSource);
    }
    Ok(())
}

fn patch_slide_style(
    archive: &mut Archive,
    slide_identifier: u64,
    old_style: u64,
    new_style: u64,
    wire_limits: WireLimits,
    limits: litchi_iwa_core::Limits,
) -> Result<(), SlideBackgroundError> {
    let object = archive
        .object_mut(slide_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let index = one_message_index(&object.messages, SLIDE_MESSAGE_TYPE)?;
    validate_rewrite_metadata(object, index)?;
    validate_slide_style_metadata_for_rewrite(object, index, old_style)?;
    let data = object.messages[index].data.clone();
    let slide_view = WireView::parse_with_limits(&data, wire_limits).map_err(map_wire_error)?;
    if slide_view
        .fields()
        .any(|field| !matches!(field.number(), 1..=7 | 10..=31 | 34..=45))
    {
        // Unknown slide fields can carry an additional style edge that this
        // focused transaction cannot retarget safely. Preserve the source by
        // refusing the edit instead of rewriting only the canonical field 1.
        return Err(SlideBackgroundError::InvalidSource);
    }
    let payload = required_bytes_from_slice(&data, SLIDE_STYLE_FIELD, wire_limits)?;
    if required_reference_payload(&payload, wire_limits)? != old_style {
        return Err(SlideBackgroundError::InvalidSource);
    }
    let replacement = rewrite_reference_identifier(&payload, new_style, wire_limits)?;
    let rewritten =
        rewrite_length_field(&data, SLIDE_STYLE_FIELD, Some(&replacement), wire_limits)?;
    object
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: SLIDE_MESSAGE_TYPE,
                data: rewritten,
            },
            limits,
        )
        .map_err(map_core_error)?;
    replace_metadata_reference(
        &mut object.archive_info.message_infos[index].object_references,
        old_style,
        new_style,
    );
    for field in &mut object.archive_info.message_infos[index].field_infos {
        if field.path.path.as_slice() == [SLIDE_STYLE_FIELD] {
            replace_metadata_reference(&mut field.object_references, old_style, new_style);
        }
    }
    Ok(())
}

fn replace_style_data(
    archive: &mut Archive,
    style_identifier: u64,
    data: Vec<u8>,
    limits: litchi_iwa_core::Limits,
) -> Result<(), SlideBackgroundError> {
    let object = archive
        .object_mut(style_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let index = one_message_index(&object.messages, SLIDE_STYLE_MESSAGE_TYPE)?;
    validate_rewrite_metadata(object, index)?;
    object
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: SLIDE_STYLE_MESSAGE_TYPE,
                data,
            },
            limits,
        )
        .map_err(map_core_error)?;
    // The field-level metadata for the removed fill is not authoritative for
    // semantic projection, but retaining stale references makes later graph
    // ownership decisions unsafe. Keep parent/stylesheet references only.
    prune_background_metadata(&mut object.archive_info.message_infos[index]);
    Ok(())
}

fn style_without_background(
    style: &StyleState,
    wire_limits: WireLimits,
) -> Result<Vec<u8>, SlideBackgroundError> {
    let properties = style
        .properties
        .as_deref()
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let properties = rewrite_length_field(properties, PROPERTIES_FILL_FIELD, None, wire_limits)?;
    let count = style
        .override_count
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let count = count
        .checked_sub(1)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let mut output = rewrite_length_field(
        &style.raw,
        STYLE_PROPERTIES_FIELD,
        Some(&properties),
        wire_limits,
    )?;
    output = rewrite_varint_field(
        &output,
        STYLE_OVERRIDE_COUNT_FIELD,
        Some(count),
        wire_limits,
    )?;
    check_wire_output(&output, wire_limits)?;
    Ok(output)
}

fn new_style_data_from_style(
    source: &StyleState,
    parent: u64,
    stylesheet: u64,
    fill: &[u8],
    wire_limits: WireLimits,
) -> Result<Vec<u8>, SlideBackgroundError> {
    let super_payload = required_bytes_from_slice(&source.raw, STYLE_SUPER_FIELD, wire_limits)?;
    // A new variation must not inherit the base style's producer identity.
    // Retaining either value makes the child look like an independently
    // named native style and prevents safe later culling. Unknown fields stay
    // source-authoritative and are preserved by these focused rewrites.
    let super_payload = rewrite_length_field(&super_payload, STYLE_NAME_FIELD, None, wire_limits)?;
    let super_payload =
        rewrite_length_field(&super_payload, STYLE_IDENTIFIER_FIELD, None, wire_limits)?;
    let super_payload =
        rewrite_reference_field(&super_payload, STYLE_PARENT_FIELD, parent, wire_limits)?;
    let super_payload =
        rewrite_varint_field(&super_payload, STYLE_VARIATION_FIELD, Some(1), wire_limits)?;
    let super_payload = rewrite_reference_field(
        &super_payload,
        STYLE_STYLESHEET_FIELD,
        stylesheet,
        wire_limits,
    )?;
    let properties = source.properties.as_deref().unwrap_or(&[]);
    let properties =
        rewrite_length_field(properties, PROPERTIES_FILL_FIELD, Some(fill), wire_limits)?;
    let output = rewrite_length_field(
        &source.raw,
        STYLE_SUPER_FIELD,
        Some(&super_payload),
        wire_limits,
    )?;
    let output = rewrite_varint_field(&output, STYLE_OVERRIDE_COUNT_FIELD, Some(1), wire_limits)?;
    rewrite_length_field(
        &output,
        STYLE_PROPERTIES_FIELD,
        Some(&properties),
        wire_limits,
    )
}

#[allow(
    dead_code,
    reason = "Kept as a compact fallback for future synthetic style producers."
)]
fn new_style_data(
    parent: u64,
    stylesheet: u64,
    fill: &[u8],
) -> Result<Vec<u8>, SlideBackgroundError> {
    let mut super_payload = Vec::new();
    append_length(
        &mut super_payload,
        STYLE_PARENT_FIELD,
        &encode_reference(parent),
    )?;
    append_varint(&mut super_payload, STYLE_VARIATION_FIELD, 1)?;
    append_length(
        &mut super_payload,
        STYLE_STYLESHEET_FIELD,
        &encode_reference(stylesheet),
    )?;
    let mut properties = Vec::new();
    append_length(&mut properties, PROPERTIES_FILL_FIELD, fill)?;
    let mut output = Vec::new();
    append_length(&mut output, STYLE_SUPER_FIELD, &super_payload)?;
    append_varint(&mut output, STYLE_OVERRIDE_COUNT_FIELD, 1)?;
    append_length(&mut output, STYLE_PROPERTIES_FIELD, &properties)?;
    Ok(output)
}

fn encode_background(
    package: &Package,
    inherited: &[u8],
    background: &Background,
) -> Result<Vec<u8>, SlideBackgroundError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let options = codec_options(package, inherited)?;
    let rewritten = match background {
        Background::None => codec::rewrite_slide_background_with_report(
            inherited,
            codec::BackgroundWrite::Clear,
            options,
        ),
        Background::Solid(color) => codec::rewrite_slide_background_with_report(
            inherited,
            codec::BackgroundWrite::Solid(codec::ColorWrite::new(
                color.red(),
                color.green(),
                color.blue(),
                color.alpha(),
                match color.color_space() {
                    RgbColorSpace::Srgb => codec::RgbSpace::Srgb,
                    RgbColorSpace::DisplayP3 => codec::RgbSpace::DisplayP3,
                },
            )),
            options,
        ),
        Background::Gradient(gradient) => {
            let stops = gradient
                .stops()
                .iter()
                .map(|stop| codec::GradientStopWrite {
                    color: codec::ColorWrite::new(
                        stop.color().red(),
                        stop.color().green(),
                        stop.color().blue(),
                        stop.color().alpha(),
                        match stop.color().color_space() {
                            RgbColorSpace::Srgb => codec::RgbSpace::Srgb,
                            RgbColorSpace::DisplayP3 => codec::RgbSpace::DisplayP3,
                        },
                    ),
                    fraction: stop.position().get(),
                    inflection: stop.midpoint().get(),
                })
                .collect::<Vec<_>>();
            codec::rewrite_slide_background_with_report(
                inherited,
                codec::BackgroundWrite::Gradient(codec::GradientWrite {
                    kind: match gradient.kind() {
                        Kind::Linear => codec::GradientKind::Linear,
                        Kind::Radial => codec::GradientKind::Radial,
                    },
                    stops: &stops,
                    opacity: gradient.opacity().get(),
                    advanced: gradient.is_advanced(),
                    angle_radians: gradient.angle().radians(),
                }),
                options,
            )
        },
        Background::Opaque(opaque) => codec::rewrite_slide_background_with_report(
            inherited,
            codec::BackgroundWrite::Raw(opaque.as_bytes()),
            options,
        ),
    };
    rewritten
        .map(|(bytes, _report)| bytes)
        .map_err(map_codec_error)
        .and_then(|bytes| {
            if bytes.len() > limits.max_output_bytes() {
                Err(SlideBackgroundError::LimitExceeded {
                    kind: SlideBackgroundLimitKind::OutputBytes,
                    observed: usize_to_u64(bytes.len()),
                    maximum: usize_to_u64(limits.max_output_bytes()),
                })
            } else {
                Ok(bytes)
            }
        })
}

fn style_is_collapsible(
    package: &Package,
    selection: &Selection,
) -> Result<bool, SlideBackgroundError> {
    let style = &selection.style;
    if style.variation != Some(true)
        || style.parent.is_none()
        || style.override_count != Some(1)
        || style.properties.is_none()
        || style.fill.is_none()
    {
        return Ok(false);
    }
    if !stylesheet_allows_culling(package, selection)? {
        return Ok(false);
    }
    if !style_archive_metadata_allows_culling(package, selection)? {
        return Ok(false);
    }
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let outer = WireView::parse_with_limits(&style.raw, limits).map_err(map_wire_error)?;
    let super_payload = required_bytes(&outer, &style.raw, STYLE_SUPER_FIELD)?;
    let properties = style
        .properties
        .as_deref()
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let super_view = WireView::parse_with_limits(super_payload, limits).map_err(map_wire_error)?;
    let properties_view =
        WireView::parse_with_limits(properties, limits).map_err(map_wire_error)?;
    // Unknown fields and any additional native property mean that culling the
    // variation would silently discard producer-authored state.
    Ok(has_exact_field_numbers(
        &outer,
        &[
            STYLE_SUPER_FIELD,
            STYLE_OVERRIDE_COUNT_FIELD,
            STYLE_PROPERTIES_FIELD,
        ],
    ) && has_exact_field_numbers(
        &super_view,
        &[
            STYLE_PARENT_FIELD,
            STYLE_VARIATION_FIELD,
            STYLE_STYLESHEET_FIELD,
        ],
    ) && has_exact_field_numbers(&properties_view, &[PROPERTIES_FILL_FIELD]))
}

fn stylesheet_allows_culling(
    package: &Package,
    selection: &Selection,
) -> Result<bool, SlideBackgroundError> {
    let stylesheet = package
        .object(selection.stylesheet_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let index = one_message_index(&stylesheet.messages, STYLESHEET_MESSAGE_TYPE)?;
    validate_rewrite_metadata(stylesheet, index)?;
    let payload = stylesheet.messages[index].data.clone();
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let view = WireView::parse_with_limits(&payload, limits).map_err(map_wire_error)?;
    if optional_bool(&view, STYLESHEET_CAN_CULL_FIELD)? != Some(true) {
        return Ok(false);
    }

    let target = selection.style_identifier;
    let expected_parent = selection
        .style
        .parent
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let mut direct = 0usize;
    let mut child = 0usize;
    for field in view.fields() {
        match field.number() {
            STYLESHEET_STYLES_FIELD => {
                let reference = field.canonical_payload().map_err(map_wire_error)?;
                if required_reference_payload(reference, limits)? == target {
                    if !reference_payload_allows_removal(reference, limits)? {
                        return Ok(false);
                    }
                    direct = direct.saturating_add(1);
                }
            },
            STYLESHEET_IDENTIFIER_MAP_FIELD => {
                let entry = field.canonical_payload().map_err(map_wire_error)?;
                let entry_view =
                    WireView::parse_with_limits(entry, limits).map_err(map_wire_error)?;
                if !has_exact_field_numbers(&entry_view, &[1, IDENTIFIED_STYLE_REFERENCE_FIELD]) {
                    return Ok(false);
                }
                if required_reference_field_from_slice(
                    entry,
                    IDENTIFIED_STYLE_REFERENCE_FIELD,
                    limits,
                )? == target
                {
                    return Ok(false);
                }
            },
            STYLESHEET_PARENT_FIELD => {
                if required_reference_payload(
                    field.canonical_payload().map_err(map_wire_error)?,
                    limits,
                )? == target
                {
                    return Ok(false);
                }
            },
            STYLESHEET_CHILDREN_FIELD => {
                let entry = field.canonical_payload().map_err(map_wire_error)?;
                let entry_view =
                    WireView::parse_with_limits(entry, limits).map_err(map_wire_error)?;
                if !has_exact_field_numbers(&entry_view, &[CHILD_PARENT_FIELD, CHILDREN_FIELD]) {
                    return Ok(false);
                }
                let parent_payload = required_bytes_from_slice(entry, CHILD_PARENT_FIELD, limits)?;
                let parent = required_reference_payload(&parent_payload, limits)?;
                if parent == target {
                    return Ok(false);
                }
                let children = repeated_payloads(entry, CHILDREN_FIELD, limits)?;
                let mut contains_target = false;
                for reference in &children {
                    if required_reference_payload(reference, limits)? == target {
                        if !reference_payload_allows_removal(reference, limits)? {
                            return Ok(false);
                        }
                        contains_target = true;
                    }
                }
                if contains_target {
                    if parent != expected_parent {
                        return Ok(false);
                    }
                    child = child.saturating_add(1);
                }
            },
            STYLESHEET_FIRST_VERSIONED_FIELD..=STYLESHEET_LAST_VERSIONED_FIELD => {
                if versioned_styles_reference(
                    field.canonical_payload().map_err(map_wire_error)?,
                    target,
                    limits,
                )? {
                    return Ok(false);
                }
            },
            4 | STYLESHEET_CAN_CULL_FIELD => {},
            _ => return Ok(false),
        }
    }
    if direct != 1 || child != 1 {
        return Ok(false);
    }

    let info = stylesheet
        .archive_info
        .message_infos
        .get(index)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    if info
        .object_references
        .iter()
        .filter(|reference| **reference == target)
        .count()
        != 1
    {
        return Ok(false);
    }
    for field in &info.field_infos {
        if !field.object_references.contains(&target) {
            continue;
        }
        let path = field.path.path.as_slice();
        let supported = path == [STYLESHEET_STYLES_FIELD]
            || path == [STYLESHEET_CHILDREN_FIELD, CHILDREN_FIELD];
        if !supported {
            return Ok(false);
        }
    }
    Ok(true)
}

fn versioned_styles_reference(
    source: &[u8],
    target: u64,
    limits: WireLimits,
) -> Result<bool, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    for field in view.fields() {
        match field.number() {
            VERSIONED_STYLES_FIELD => {
                if required_reference_payload(
                    field.canonical_payload().map_err(map_wire_error)?,
                    limits,
                )? == target
                {
                    return Ok(true);
                }
            },
            VERSIONED_IDENTIFIER_MAP_FIELD => {
                let entry = field.canonical_payload().map_err(map_wire_error)?;
                let entry_view =
                    WireView::parse_with_limits(entry, limits).map_err(map_wire_error)?;
                if !has_exact_field_numbers(&entry_view, &[1, IDENTIFIED_STYLE_REFERENCE_FIELD])
                    || required_reference_field_from_slice(
                        entry,
                        IDENTIFIED_STYLE_REFERENCE_FIELD,
                        limits,
                    )? == target
                {
                    return Ok(true);
                }
            },
            VERSIONED_CHILDREN_FIELD => {
                let entry = field.canonical_payload().map_err(map_wire_error)?;
                let entry_view =
                    WireView::parse_with_limits(entry, limits).map_err(map_wire_error)?;
                if !has_exact_field_numbers(&entry_view, &[CHILD_PARENT_FIELD, CHILDREN_FIELD]) {
                    return Ok(true);
                }
                let parent =
                    required_reference_field_from_slice(entry, CHILD_PARENT_FIELD, limits)?;
                if parent == target {
                    return Ok(true);
                }
                for child in repeated_payloads(entry, CHILDREN_FIELD, limits)? {
                    if required_reference_payload(&child, limits)? == target {
                        return Ok(true);
                    }
                }
            },
            _ => return Ok(true),
        }
    }
    Ok(false)
}

fn style_archive_metadata_allows_culling(
    package: &Package,
    selection: &Selection,
) -> Result<bool, SlideBackgroundError> {
    let style = package
        .object(selection.style_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    if style.messages.len() != 1 || style.archive_info.message_infos.len() != 1 {
        return Ok(false);
    }
    if validate_rewrite_metadata(style, 0).is_err() {
        return Ok(false);
    }
    let info = &style.archive_info.message_infos[0];
    if !info.data_references.is_empty() {
        return Ok(false);
    }
    let parent = selection
        .style
        .parent
        .ok_or(SlideBackgroundError::InvalidSource)?;
    if info
        .object_references
        .iter()
        .any(|identifier| *identifier != parent && *identifier != selection.stylesheet_identifier)
    {
        return Ok(false);
    }
    for field in &info.field_infos {
        if !field.data_references.is_empty() {
            return Ok(false);
        }
        if field.object_references.is_empty() {
            continue;
        }
        let path = field.path.path.as_slice();
        let valid = if path == [STYLE_SUPER_FIELD, STYLE_PARENT_FIELD] {
            field.object_references.iter().all(|value| *value == parent)
        } else if path == [STYLE_SUPER_FIELD, STYLE_STYLESHEET_FIELD] {
            field
                .object_references
                .iter()
                .all(|value| *value == selection.stylesheet_identifier)
        } else {
            false
        };
        if !valid {
            return Ok(false);
        }
    }
    Ok(true)
}

fn validate_style_background_reset_metadata(
    package: &Package,
    selection: &Selection,
) -> Result<(), SlideBackgroundError> {
    let style = package
        .object(selection.style_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    if style.messages.len() != 1 || style.archive_info.message_infos.len() != 1 {
        return Err(SlideBackgroundError::InvalidSource);
    }
    validate_rewrite_metadata(style, 0)?;
    let info = &style.archive_info.message_infos[0];
    let parent = selection
        .style
        .parent
        .ok_or(SlideBackgroundError::InvalidSource)?;

    // Parent and stylesheet references are the conventional aggregate style
    // ownership edges. Every other aggregate reference must be attributable
    // to field metadata. Reset removes the fill field metadata, while retained
    // property fields keep their shared references alive.
    for reference in &info.object_references {
        if *reference == parent || *reference == selection.stylesheet_identifier {
            continue;
        }
        if !info
            .field_infos
            .iter()
            .any(|field| field.object_references.contains(reference))
        {
            return Err(SlideBackgroundError::InvalidSource);
        }
    }
    for reference in &info.data_references {
        if !info
            .field_infos
            .iter()
            .any(|field| field.data_references.contains(reference))
        {
            return Err(SlideBackgroundError::InvalidSource);
        }
    }

    // Field-level references must also be present in the aggregate lists.
    // Unknown retained property metadata is source-authoritative and remains
    // safe because prune_background_metadata removes only the fill subtree.
    for field in &info.field_infos {
        let path = field.path.path.as_slice();
        let identity = path == [STYLE_SUPER_FIELD, STYLE_PARENT_FIELD]
            || path == [STYLE_SUPER_FIELD, STYLE_STYLESHEET_FIELD];
        let fill = path.starts_with(&[STYLE_PROPERTIES_FIELD, PROPERTIES_FILL_FIELD]);
        if fill
            && field.object_references.iter().any(|reference| {
                *reference == parent || *reference == selection.stylesheet_identifier
            })
        {
            // Removing fill metadata must never make either surviving style
            // identity edge look unowned at the aggregate level.
            return Err(SlideBackgroundError::InvalidSource);
        }
        if identity {
            let expected = if path[1] == STYLE_PARENT_FIELD {
                parent
            } else {
                selection.stylesheet_identifier
            };
            if field
                .object_references
                .iter()
                .any(|reference| *reference != expected)
                || !field.data_references.is_empty()
            {
                return Err(SlideBackgroundError::InvalidSource);
            }
        }
        if field
            .object_references
            .iter()
            .any(|reference| !info.object_references.contains(reference))
            || field
                .data_references
                .iter()
                .any(|reference| !info.data_references.contains(reference))
        {
            return Err(SlideBackgroundError::InvalidSource);
        }
    }
    Ok(())
}

fn stylesheet_payload_references_target(
    source: &[u8],
    target: u64,
    limits: WireLimits,
) -> Result<bool, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    for field in view.fields() {
        match field.number() {
            STYLESHEET_STYLES_FIELD => {
                if required_reference_payload(
                    field.canonical_payload().map_err(map_wire_error)?,
                    limits,
                )? == target
                {
                    return Ok(true);
                }
            },
            STYLESHEET_IDENTIFIER_MAP_FIELD => {
                let entry = field.canonical_payload().map_err(map_wire_error)?;
                let entry_view =
                    WireView::parse_with_limits(entry, limits).map_err(map_wire_error)?;
                if !has_exact_field_numbers(&entry_view, &[1, IDENTIFIED_STYLE_REFERENCE_FIELD]) {
                    return Ok(true);
                }
                if required_reference_field_from_slice(
                    entry,
                    IDENTIFIED_STYLE_REFERENCE_FIELD,
                    limits,
                )? == target
                {
                    return Ok(true);
                }
            },
            STYLESHEET_PARENT_FIELD => {
                if required_reference_payload(
                    field.canonical_payload().map_err(map_wire_error)?,
                    limits,
                )? == target
                {
                    return Ok(true);
                }
            },
            STYLESHEET_CHILDREN_FIELD => {
                let entry = field.canonical_payload().map_err(map_wire_error)?;
                let entry_view =
                    WireView::parse_with_limits(entry, limits).map_err(map_wire_error)?;
                if !has_exact_field_numbers(&entry_view, &[CHILD_PARENT_FIELD, CHILDREN_FIELD]) {
                    return Ok(true);
                }
                let parent_payload = required_bytes_from_slice(entry, CHILD_PARENT_FIELD, limits)?;
                if required_reference_payload(&parent_payload, limits)? == target {
                    return Ok(true);
                }
                for child in repeated_payloads(entry, CHILDREN_FIELD, limits)? {
                    if required_reference_payload(&child, limits)? == target {
                        return Ok(true);
                    }
                }
            },
            STYLESHEET_FIRST_VERSIONED_FIELD..=STYLESHEET_LAST_VERSIONED_FIELD => {
                if versioned_styles_reference(
                    field.canonical_payload().map_err(map_wire_error)?,
                    target,
                    limits,
                )? {
                    return Ok(true);
                }
            },
            4 | STYLESHEET_CAN_CULL_FIELD => {},
            // A future stylesheet field may carry a style edge that this
            // package cannot attribute.  Refuse to cull in that case.
            _ => return Ok(true),
        }
    }
    Ok(false)
}

fn style_is_exclusive(
    package: &Package,
    selection: &Selection,
) -> Result<bool, SlideBackgroundError> {
    let style_identifier = selection.style_identifier;
    let wire_limits = package.wire_limits().map_err(map_wire_error)?;
    let mut slides = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for (message_index, message) in object.messages.iter().enumerate() {
                let info = object
                    .archive_info
                    .message_infos
                    .get(message_index)
                    .ok_or(SlideBackgroundError::InvalidSource)?;
                let aggregate_owners = info
                    .object_references
                    .iter()
                    .filter(|reference| **reference == style_identifier)
                    .count();
                let unsupported_field_owner = info.field_infos.iter().any(|field| {
                    field.object_references.contains(&style_identifier)
                        && (message.type_ != SLIDE_MESSAGE_TYPE
                            || field.path.path.as_slice() != [SLIDE_STYLE_FIELD])
                        && message.type_ != STYLESHEET_MESSAGE_TYPE
                });
                if unsupported_field_owner
                    || (aggregate_owners != 0
                        && message.type_ != SLIDE_MESSAGE_TYPE
                        && message.type_ != STYLESHEET_MESSAGE_TYPE)
                {
                    // An unrecognized owner means that culling this style
                    // could orphan a native graph edge.
                    return Ok(false);
                }
                if message.type_ == SLIDE_MESSAGE_TYPE {
                    let slide_view = WireView::parse_with_limits(&message.data, wire_limits)
                        .map_err(map_wire_error)?;
                    if slide_view
                        .fields()
                        .any(|field| !matches!(field.number(), 1..=7 | 10..=31 | 34..=45))
                    {
                        return Ok(false);
                    }
                    let selected =
                        optional_reference_from_payload(package, &message.data, SLIDE_STYLE_FIELD)?
                            == Some(style_identifier);
                    if aggregate_owners != usize::from(selected)
                        || info.field_infos.iter().any(|field| {
                            field.object_references.contains(&style_identifier)
                                && field.path.path.as_slice() != [SLIDE_STYLE_FIELD]
                        })
                    {
                        return Ok(false);
                    }
                    if selected {
                        slides += 1;
                    }
                } else if message.type_ == SLIDE_STYLE_MESSAGE_TYPE {
                    let style = parse_style_metadata(package, &message.data)?;
                    if style.parent == Some(style_identifier)
                        || style.stylesheet == Some(style_identifier)
                        || (object.archive_info.identifier != Some(style_identifier)
                            && !other_style_payload_allows_culling(&message.data, wire_limits)?)
                    {
                        return Ok(false);
                    }
                } else if message.type_ == STYLESHEET_MESSAGE_TYPE
                    && (component.name() != selection.stylesheet_component
                        || object.archive_info.identifier != Some(selection.stylesheet_identifier))
                {
                    if aggregate_owners != 0
                        || info
                            .field_infos
                            .iter()
                            .any(|field| field.object_references.contains(&style_identifier))
                        || stylesheet_payload_references_target(
                            &message.data,
                            style_identifier,
                            wire_limits,
                        )?
                    {
                        // A different stylesheet is an independent owner of
                        // the style graph.  Removing the selected style would
                        // leave that owner dangling, so force copy-on-write.
                        return Ok(false);
                    }
                }
            }
        }
    }
    if slides != 1 {
        return Ok(false);
    }
    if let Some(metadata_name) = metadata_component_name(package)? {
        let metadata = package
            .state
            .source
            .components()
            .iter()
            .find(|component| component.name() == metadata_name)
            .and_then(|component| {
                component.archive().objects.iter().find_map(|object| {
                    object
                        .messages
                        .iter()
                        .find(|message| message.type_ == PACKAGE_METADATA_MESSAGE_TYPE)
                        .map(|message| message.data.clone())
                })
            })
            .ok_or(SlideBackgroundError::InvalidSource)?;
        let options = metadata_options(package, 8)?;
        let mut ownership = MetadataOwnershipFacts {
            stylesheet_identifier: selection.stylesheet_identifier,
            style_identifier,
            selected_slide_locator: metadata_locator(&selection.slide_component).to_owned(),
            selected: 0,
            other: false,
        };
        inspect_package_metadata_with_visitor(&metadata, options, &mut ownership)
            .map_err(map_metadata_error)?;
        if ownership.other || ownership.selected != 1 {
            return Ok(false);
        }
    }
    Ok(true)
}

fn other_style_payload_allows_culling(
    source: &[u8],
    limits: WireLimits,
) -> Result<bool, SlideBackgroundError> {
    let outer = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    if !has_only_field_numbers(
        &outer,
        &[
            STYLE_SUPER_FIELD,
            STYLE_OVERRIDE_COUNT_FIELD,
            STYLE_PROPERTIES_FIELD,
        ],
    ) {
        return Ok(false);
    }
    let super_payload = required_bytes(&outer, source, STYLE_SUPER_FIELD)?;
    let super_view = WireView::parse_with_limits(super_payload, limits).map_err(map_wire_error)?;
    if !has_only_field_numbers(&super_view, &[1, 2, 3, 4, 5]) {
        return Ok(false);
    }
    for number in [STYLE_PARENT_FIELD, STYLE_STYLESHEET_FIELD] {
        if let Some(reference) = optional_bytes(&super_view, super_payload, number)?
            && !reference_payload_allows_removal(reference, limits)?
        {
            return Ok(false);
        }
    }
    if let Some(properties) = optional_bytes(&outer, source, STYLE_PROPERTIES_FIELD)? {
        let properties = WireView::parse_with_limits(properties, limits).map_err(map_wire_error)?;
        if !has_only_field_numbers(&properties, &[1, 2, 3, 4, 5, 6, 7]) {
            return Ok(false);
        }
    }
    Ok(true)
}

fn patch_stylesheet(
    archive: &mut Archive,
    stylesheet_identifier: u64,
    remove: Option<u64>,
    insertion: Option<(u64, u64)>,
    wire_limits: WireLimits,
    limits: litchi_iwa_core::Limits,
) -> Result<(), SlideBackgroundError> {
    let object = archive
        .object_mut(stylesheet_identifier)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    let index = one_message_index(&object.messages, STYLESHEET_MESSAGE_TYPE)?;
    validate_rewrite_metadata(object, index)?;
    let data = object.messages[index].data.clone();
    let mut styles = repeated_payloads(&data, STYLESHEET_STYLES_FIELD, wire_limits)?;
    if let Some(old) = remove {
        let style_ids = styles
            .iter()
            .map(|payload| required_reference_payload(payload, wire_limits))
            .collect::<Result<Vec<_>, _>>()?;
        if style_ids.iter().filter(|id| **id == old).count() != 1 {
            return Err(SlideBackgroundError::InvalidSource);
        }
        styles = styles
            .into_iter()
            .zip(style_ids)
            .filter_map(|(payload, id)| (id != old).then_some(payload))
            .collect();
    }
    if let Some((_, new_id)) = insertion {
        if styles
            .iter()
            .map(|payload| required_reference_payload(payload, wire_limits))
            .collect::<Result<Vec<_>, _>>()?
            .contains(&new_id)
        {
            return Err(SlideBackgroundError::InvalidSource);
        }
        styles.push(encode_reference(new_id));
    }
    let mut rewritten =
        rewrite_repeated_length_field(&data, STYLESHEET_STYLES_FIELD, &styles, wire_limits)?;
    let mut children = repeated_payloads(&rewritten, STYLESHEET_CHILDREN_FIELD, wire_limits)?;
    let mut matched_parent = false;
    let mut retained_children = Vec::new();
    for child in children.drain(..) {
        let parent_payload = required_bytes_from_slice(&child, CHILD_PARENT_FIELD, wire_limits)?;
        let parent = required_reference_payload(&parent_payload, wire_limits)?;
        let mut child_refs = repeated_payloads(&child, CHILDREN_FIELD, wire_limits)?;
        if let Some(old) = remove {
            let child_ids = child_refs
                .iter()
                .map(|payload| required_reference_payload(payload, wire_limits))
                .collect::<Result<Vec<_>, _>>()?;
            child_refs = child_refs
                .into_iter()
                .zip(child_ids)
                .filter_map(|(payload, id)| (id != old).then_some(payload))
                .collect();
        }
        if let Some((insert_parent, new_id)) = insertion {
            if parent == insert_parent {
                if matched_parent {
                    return Err(SlideBackgroundError::InvalidSource);
                }
                matched_parent = true;
                child_refs.push(encode_reference(new_id));
            }
        }
        if !child_refs.is_empty() {
            retained_children.push(rewrite_repeated_length_field(
                &child,
                CHILDREN_FIELD,
                &child_refs,
                wire_limits,
            )?);
        }
    }
    if let Some((parent, new_id)) = insertion {
        if !matched_parent {
            let mut child = Vec::new();
            append_length(&mut child, CHILD_PARENT_FIELD, &encode_reference(parent))?;
            append_length(&mut child, CHILDREN_FIELD, &encode_reference(new_id))?;
            retained_children.push(child);
        }
    }
    rewritten = rewrite_repeated_length_field(
        &rewritten,
        STYLESHEET_CHILDREN_FIELD,
        &retained_children,
        wire_limits,
    )?;
    object
        .replace_message_preserving_header_with_limits(
            index,
            RawMessage {
                type_: STYLESHEET_MESSAGE_TYPE,
                data: rewritten,
            },
            limits,
        )
        .map_err(map_core_error)?;
    if let Some(old) = remove {
        prune_object_reference(
            &mut object.archive_info.message_infos[index].object_references,
            old,
        );
    }
    if let Some((_, new_id)) = insertion {
        if !object.archive_info.message_infos[index]
            .object_references
            .contains(&new_id)
        {
            object.archive_info.message_infos[index]
                .object_references
                .push(new_id);
        }
    }
    let info = &mut object.archive_info.message_infos[index];
    update_stylesheet_field_metadata(info, remove, insertion);
    Ok(())
}

fn update_stylesheet_field_metadata(
    info: &mut litchi_iwa_core::MessageInfo,
    remove: Option<u64>,
    insertion: Option<(u64, u64)>,
) {
    let mut style_added = false;
    let mut child_added = false;
    for field in &mut info.field_infos {
        let path = field.path.path.as_slice();
        let is_style_field = path == [STYLESHEET_STYLES_FIELD];
        let is_child_field = path == [STYLESHEET_CHILDREN_FIELD, CHILDREN_FIELD];
        if !is_style_field && !is_child_field {
            continue;
        }
        if let Some(old) = remove {
            prune_object_reference(&mut field.object_references, old);
        }
        let Some((_, new_id)) = insertion else {
            continue;
        };
        if is_style_field {
            if !style_added && !field.object_references.contains(&new_id) {
                field.object_references.push(new_id);
                style_added = true;
            }
        } else if !child_added && !field.object_references.contains(&new_id) {
            field.object_references.push(new_id);
            child_added = true;
        }
    }
}

fn exactly_one_payload(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<Vec<u8>, SlideBackgroundError> {
    let mut selected = None;
    for message in &object.messages {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(message.data.clone()).is_some() {
            return Err(SlideBackgroundError::InvalidSource);
        }
    }
    selected.ok_or(SlideBackgroundError::InvalidSource)
}

fn one_message_index(
    messages: &[RawMessage],
    message_type: u32,
) -> Result<usize, SlideBackgroundError> {
    let mut selected = None;
    for (index, message) in messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(index).is_some() {
            return Err(SlideBackgroundError::InvalidSource);
        }
    }
    selected.ok_or(SlideBackgroundError::InvalidSource)
}

fn validate_rewrite_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), SlideBackgroundError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideBackgroundError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(SlideBackgroundError::InvalidSource);
    }
    Ok(())
}

fn required_reference_field(
    package: &Package,
    source: &[u8],
    number: u32,
) -> Result<u64, SlideBackgroundError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let payload = required_bytes(&view, source, number)?;
    required_reference_payload(payload, limits)
}

fn required_reference_field_from_slice(
    source: &[u8],
    number: u32,
    limits: WireLimits,
) -> Result<u64, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let payload = required_bytes(&view, source, number)?;
    required_reference_payload(payload, limits)
}

fn required_reference_payload(
    source: &[u8],
    limits: WireLimits,
) -> Result<u64, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let field = unique_field(&view, 1)?;
    let field = field.ok_or(SlideBackgroundError::InvalidSource)?;
    if field.wire_type() != 0 {
        return Err(SlideBackgroundError::InvalidSource);
    }
    field.validate_canonical_key().map_err(map_wire_error)?;
    let value = canonical_varint(field.payload())?;
    if value == 0 {
        return Err(SlideBackgroundError::InvalidSource);
    }
    Ok(value)
}

fn reference_payload_allows_removal(
    source: &[u8],
    limits: WireLimits,
) -> Result<bool, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    required_reference_payload(source, limits)?;
    if !has_only_field_numbers(&view, &[1, 2, 3]) {
        return Ok(false);
    }
    optional_varint(&view, 2)?;
    optional_bool(&view, 3)?;
    Ok(true)
}

fn optional_reference(
    view: &WireView<'_>,
    source: &[u8],
    number: u32,
    limits: WireLimits,
) -> Result<Option<u64>, SlideBackgroundError> {
    let Some(payload) = optional_bytes(view, source, number)? else {
        return Ok(None);
    };
    required_reference_payload(payload, limits).map(Some)
}

fn optional_reference_from_payload(
    package: &Package,
    source: &[u8],
    number: u32,
) -> Result<Option<u64>, SlideBackgroundError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let Some(payload) = optional_bytes(&view, source, number)? else {
        return Ok(None);
    };
    required_reference_payload(payload, limits).map(Some)
}

fn parse_style_metadata(
    package: &Package,
    source: &[u8],
) -> Result<StyleState, SlideBackgroundError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let outer = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let super_payload = required_bytes(&outer, source, STYLE_SUPER_FIELD)?;
    let super_view = WireView::parse_with_limits(super_payload, limits).map_err(map_wire_error)?;
    Ok(StyleState {
        raw: source.to_vec(),
        parent: optional_reference(&super_view, super_payload, STYLE_PARENT_FIELD, limits)?,
        variation: optional_bool(&super_view, STYLE_VARIATION_FIELD)?,
        stylesheet: optional_reference(&super_view, super_payload, STYLE_STYLESHEET_FIELD, limits)?,
        override_count: optional_varint(&outer, STYLE_OVERRIDE_COUNT_FIELD)?,
        properties: optional_bytes(&outer, source, STYLE_PROPERTIES_FIELD)?.map(ToOwned::to_owned),
        fill: None,
    })
}

fn required_bytes<'a>(
    view: &WireView<'a>,
    source: &'a [u8],
    number: u32,
) -> Result<&'a [u8], SlideBackgroundError> {
    optional_bytes(view, source, number)?.ok_or(SlideBackgroundError::InvalidSource)
}

fn required_bytes_from_slice(
    source: &[u8],
    number: u32,
    limits: WireLimits,
) -> Result<Vec<u8>, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    optional_bytes(&view, source, number)
        .map(|payload| payload.map(ToOwned::to_owned))?
        .ok_or(SlideBackgroundError::InvalidSource)
}

fn optional_bytes<'a>(
    view: &WireView<'a>,
    _source: &'a [u8],
    number: u32,
) -> Result<Option<&'a [u8]>, SlideBackgroundError> {
    let Some(field) = unique_field(view, number)? else {
        return Ok(None);
    };
    if field.wire_type() != 2 {
        return Err(SlideBackgroundError::InvalidSource);
    }
    field.validate_canonical_framing().map_err(map_wire_error)?;
    Ok(Some(field.payload()))
}

fn optional_bytes_in_payload(
    source: &[u8],
    number: u32,
    limits: WireLimits,
) -> Result<Option<&[u8]>, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    optional_bytes(&view, source, number)
}

fn optional_varint(view: &WireView<'_>, number: u32) -> Result<Option<u64>, SlideBackgroundError> {
    let Some(field) = unique_field(view, number)? else {
        return Ok(None);
    };
    if field.wire_type() != 0 {
        return Err(SlideBackgroundError::InvalidSource);
    }
    field.validate_canonical_key().map_err(map_wire_error)?;
    canonical_varint(field.payload()).map(Some)
}

fn optional_bool(view: &WireView<'_>, number: u32) -> Result<Option<bool>, SlideBackgroundError> {
    let Some(value) = optional_varint(view, number)? else {
        return Ok(None);
    };
    match value {
        0 => Ok(Some(false)),
        1 => Ok(Some(true)),
        _ => Err(SlideBackgroundError::InvalidSource),
    }
}

fn unique_field<'a>(
    view: &WireView<'a>,
    number: u32,
) -> Result<Option<WireFieldView<'a>>, SlideBackgroundError> {
    let mut selected = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if selected.replace(field).is_some() {
            return Err(SlideBackgroundError::InvalidSource);
        }
    }
    Ok(selected)
}

fn field_numbers(view: &WireView<'_>) -> Vec<u32> {
    view.fields().map(WireFieldView::number).collect()
}

fn has_exact_field_numbers(view: &WireView<'_>, expected: &[u32]) -> bool {
    let mut actual = field_numbers(view);
    let mut expected = expected.to_vec();
    actual.sort_unstable();
    expected.sort_unstable();
    actual == expected
}

fn has_only_field_numbers(view: &WireView<'_>, allowed: &[u32]) -> bool {
    view.fields().all(|field| allowed.contains(&field.number()))
}

fn canonical_varint(payload: &[u8]) -> Result<u64, SlideBackgroundError> {
    let (value, used) =
        decode_varint_from_bytes(payload).map_err(|_| SlideBackgroundError::InvalidSource)?;
    if used != payload.len() {
        return Err(SlideBackgroundError::InvalidSource);
    }
    Ok(value)
}

fn encode_reference(identifier: u64) -> Vec<u8> {
    let mut output = Vec::new();
    encode_varint_into(&mut output, u64::from(1_u32) << 3);
    encode_varint_into(&mut output, identifier);
    output
}

fn rewrite_reference_identifier(
    source: &[u8],
    identifier: u64,
    limits: WireLimits,
) -> Result<Vec<u8>, SlideBackgroundError> {
    if identifier == 0 {
        return Err(SlideBackgroundError::InvalidSource);
    }
    required_reference_payload(source, limits)?;
    rewrite_varint_field(source, 1, Some(identifier), limits)
}

fn rewrite_reference_field(
    source: &[u8],
    number: u32,
    identifier: u64,
    limits: WireLimits,
) -> Result<Vec<u8>, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let replacement = match optional_bytes(&view, source, number)? {
        Some(reference) => rewrite_reference_identifier(reference, identifier, limits)?,
        None => encode_reference(identifier),
    };
    rewrite_length_field(source, number, Some(&replacement), limits)
}

fn append_length(
    output: &mut Vec<u8>,
    number: u32,
    payload: &[u8],
) -> Result<(), SlideBackgroundError> {
    litchi_iwa_common::wire::append_length_delimited_field(output, number, payload)
        .map_err(map_wire_error)
}

fn append_varint(
    output: &mut Vec<u8>,
    number: u32,
    value: u64,
) -> Result<(), SlideBackgroundError> {
    litchi_iwa_common::wire::append_varint_field(output, number, value).map_err(map_wire_error)
}

fn rewrite_length_field(
    source: &[u8],
    number: u32,
    replacement: Option<&[u8]>,
    limits: WireLimits,
) -> Result<Vec<u8>, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut output = Vec::new();
    let mut seen = false;
    for field in view.fields() {
        if field.number() != number {
            output.extend_from_slice(field.raw());
            continue;
        }
        if seen || field.wire_type() != 2 {
            return Err(SlideBackgroundError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        seen = true;
        if let Some(payload) = replacement {
            append_length(&mut output, number, payload)?;
        }
    }
    if !seen {
        if let Some(payload) = replacement {
            append_length(&mut output, number, payload)?;
        }
    }
    check_wire_output(&output, limits)?;
    Ok(output)
}

fn rewrite_varint_field(
    source: &[u8],
    number: u32,
    replacement: Option<u64>,
    limits: WireLimits,
) -> Result<Vec<u8>, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut output = Vec::new();
    let mut seen = false;
    for field in view.fields() {
        if field.number() != number {
            output.extend_from_slice(field.raw());
            continue;
        }
        if seen || field.wire_type() != 0 {
            return Err(SlideBackgroundError::InvalidSource);
        }
        field.validate_canonical_key().map_err(map_wire_error)?;
        canonical_varint(field.payload())?;
        seen = true;
        if let Some(value) = replacement {
            append_varint(&mut output, number, value)?;
        }
    }
    if !seen {
        if let Some(value) = replacement {
            append_varint(&mut output, number, value)?;
        }
    }
    check_wire_output(&output, limits)?;
    Ok(output)
}

fn repeated_payloads(
    source: &[u8],
    number: u32,
    limits: WireLimits,
) -> Result<Vec<Vec<u8>>, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut payloads = Vec::new();
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 {
            return Err(SlideBackgroundError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        payloads.push(field.payload().to_vec());
    }
    Ok(payloads)
}

fn rewrite_repeated_length_field(
    source: &[u8],
    number: u32,
    replacements: &[Vec<u8>],
    limits: WireLimits,
) -> Result<Vec<u8>, SlideBackgroundError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut output = Vec::new();
    let mut selected = 0usize;
    for field in view.fields() {
        if field.number() != number {
            output.extend_from_slice(field.raw());
            continue;
        }
        if field.wire_type() != 2 {
            return Err(SlideBackgroundError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if let Some(payload) = replacements.get(selected) {
            append_length(&mut output, number, payload)?;
        }
        selected += 1;
    }
    for payload in replacements.iter().skip(selected) {
        append_length(&mut output, number, payload)?;
    }
    check_wire_output(&output, limits)?;
    Ok(output)
}

fn check_wire_output(output: &[u8], limits: WireLimits) -> Result<(), SlideBackgroundError> {
    if output.len() > limits.max_output_bytes() {
        return Err(SlideBackgroundError::LimitExceeded {
            kind: SlideBackgroundLimitKind::OutputBytes,
            observed: usize_to_u64(output.len()),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    Ok(())
}

fn replace_metadata_reference(references: &mut Vec<u64>, old: u64, new: u64) {
    let mut found = false;
    for reference in references.iter_mut() {
        if *reference == old {
            *reference = new;
            found = true;
        }
    }
    if !found && !references.contains(&new) {
        references.push(new);
    }
}

fn prune_object_reference(references: &mut Vec<u64>, identifier: u64) {
    references.retain(|candidate| *candidate != identifier);
}

fn prune_background_metadata(info: &mut litchi_iwa_core::MessageInfo) {
    let mut removed_objects = HashSet::new();
    let mut removed_data = HashSet::new();
    info.field_infos.retain(|field| {
        if field
            .path
            .path
            .starts_with(&[STYLE_PROPERTIES_FIELD, PROPERTIES_FILL_FIELD])
        {
            removed_objects.extend(field.object_references.iter().copied());
            removed_data.extend(field.data_references.iter().copied());
            false
        } else {
            true
        }
    });
    let retained_objects: HashSet<u64> = info
        .field_infos
        .iter()
        .flat_map(|field| field.object_references.iter().copied())
        .collect();
    let retained_data: HashSet<u64> = info
        .field_infos
        .iter()
        .flat_map(|field| field.data_references.iter().copied())
        .collect();
    info.object_references.retain(|identifier| {
        !removed_objects.contains(identifier) || retained_objects.contains(identifier)
    });
    info.data_references.retain(|identifier| {
        !removed_data.contains(identifier) || retained_data.contains(identifier)
    });
}

fn codec_options(
    package: &Package,
    _source: &[u8],
) -> Result<codec::DecodeOptions, SlideBackgroundError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let maximum = limits.max_input_bytes();
    let nesting = u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX);
    Ok(codec::DecodeOptions::new(maximum, nesting)
        .with_resource_limits(limits.max_fields(), limits.max_rewrite_work()))
}

fn map_codec_error(error: codec::DecodeError) -> SlideBackgroundError {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideBackgroundError::LimitExceeded {
            kind: SlideBackgroundLimitKind::WireFields,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideBackgroundError::LimitExceeded {
            kind: SlideBackgroundLimitKind::WireWork,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            codec::WireResourceLimit::Bytes { observed, maximum } => {
                SlideBackgroundError::LimitExceeded {
                    kind: SlideBackgroundLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            codec::WireResourceLimit::Nesting { observed, maximum } => {
                SlideBackgroundError::LimitExceeded {
                    kind: SlideBackgroundLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => SlideBackgroundError::InvalidSource,
        };
    }
    SlideBackgroundError::InvalidSource
}

fn map_metadata_error(error: metadata_codec::RewriteError) -> SlideBackgroundError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            metadata_codec::RewriteLimit::InputBytes { observed, maximum } => (
                SlideBackgroundLimitKind::WireBytes,
                usize_to_u64(observed),
                usize_to_u64(maximum),
            ),
            metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => (
                SlideBackgroundLimitKind::OutputBytes,
                usize_to_u64(observed),
                usize_to_u64(maximum),
            ),
            metadata_codec::RewriteLimit::Fields { observed, maximum } => (
                SlideBackgroundLimitKind::WireFields,
                usize_to_u64(observed),
                usize_to_u64(maximum),
            ),
            metadata_codec::RewriteLimit::Work { observed, maximum } => (
                SlideBackgroundLimitKind::WireWork,
                usize_to_u64(observed),
                usize_to_u64(maximum),
            ),
            metadata_codec::RewriteLimit::Nesting { observed, maximum } => (
                SlideBackgroundLimitKind::WireNesting,
                u64::from(observed),
                u64::from(maximum),
            ),
            metadata_codec::RewriteLimit::Components { observed, maximum }
            | metadata_codec::RewriteLimit::References { observed, maximum }
            | metadata_codec::RewriteLimit::Additions { observed, maximum } => (
                SlideBackgroundLimitKind::References,
                usize_to_u64(observed),
                usize_to_u64(maximum),
            ),
            _ => return SlideBackgroundError::InvalidSource,
        };
        return SlideBackgroundError::LimitExceeded {
            kind,
            observed,
            maximum,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return SlideBackgroundError::Allocation { amount };
    }
    SlideBackgroundError::InvalidSource
}

fn map_slide_selector_error(error: SlideSelectorError) -> SlideBackgroundError {
    match error {
        SlideSelectorError::EmptySlideName => SlideBackgroundError::EmptySlideName,
        SlideSelectorError::DuplicateSlideName { .. } => SlideBackgroundError::AmbiguousSelector,
    }
}

fn map_read_error(error: ReadError) -> SlideBackgroundError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideBackgroundError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => SlideBackgroundLimitKind::Entries,
                SemanticLimitKind::Slides => SlideBackgroundLimitKind::Slides,
                SemanticLimitKind::References => SlideBackgroundLimitKind::References,
                SemanticLimitKind::TextStorages
                | SemanticLimitKind::TextFragments
                | SemanticLimitKind::TextBytes => SlideBackgroundLimitKind::WireBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideBackgroundError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideBackgroundLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideBackgroundLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideBackgroundLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideBackgroundLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => SlideBackgroundError::Allocation { amount },
        _ => SlideBackgroundError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SlideBackgroundError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SlideBackgroundError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => SlideBackgroundLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => SlideBackgroundLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    SlideBackgroundLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => SlideBackgroundLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => SlideBackgroundLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideBackgroundError::Allocation { amount }
        },
        _ => SlideBackgroundError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideBackgroundError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideBackgroundError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideBackgroundLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => SlideBackgroundLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => SlideBackgroundLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes => SlideBackgroundLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => SlideBackgroundLimitKind::TotalBytes,
                _ => SlideBackgroundLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideBackgroundError::Allocation { amount }
        },
        _ => SlideBackgroundError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideBackgroundError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideBackgroundError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => SlideBackgroundLimitKind::TotalBytes,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => SlideBackgroundLimitKind::Entries,
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    SlideBackgroundLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::HeaderFields => SlideBackgroundLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => SlideBackgroundLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyFrames => SlideBackgroundLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideBackgroundError::Allocation { amount: requested }
        },
        _ => SlideBackgroundError::InvalidSource,
    }
}

fn map_rendering_error(
    error: super::rendering_invalidation::RenderingInvalidationError,
) -> SlideBackgroundError {
    match error {
        super::rendering_invalidation::RenderingInvalidationError::Allocation { amount } => {
            SlideBackgroundError::Allocation { amount }
        },
        super::rendering_invalidation::RenderingInvalidationError::InvalidSource => {
            SlideBackgroundError::InvalidSource
        },
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn byte_equal_but_semantically_different_patch_is_not_a_noop() {
        let source = Arc::<[u8]>::from(&b"same source bytes"[..]);
        let target = Background::Solid(
            Rgba::new(0.25, 0.5, 0.75, 1.0, RgbColorSpace::Srgb).expect("test color is valid"),
        );
        let patch = SlideBackgroundPatch {
            artifacts: ExactArtifacts::new(Arc::clone(&source), source),
            slide_position: Position::new(0),
            slide_identifier: 1,
            style_identifier: 2,
            target_style_identifier: 2,
            before: Background::None,
            before_override: None,
            after: Some(target.clone()),
            target_effective: target.clone(),
            target_override: Some(target),
            reset: false,
            touched_components: 1,
            deleted_previews: 0,
        };

        assert!(patch.artifacts.is_byte_noop());
        assert!(!patch.is_noop());
    }

    #[test]
    fn target_style_mismatch_is_not_a_noop() {
        let source = Arc::<[u8]>::from(&b"same source bytes"[..]);
        let background = Background::Solid(
            Rgba::new(0.25, 0.5, 0.75, 1.0, RgbColorSpace::Srgb).expect("test color is valid"),
        );
        let patch = SlideBackgroundPatch {
            artifacts: ExactArtifacts::new(Arc::clone(&source), source),
            slide_position: Position::new(0),
            slide_identifier: 1,
            style_identifier: 2,
            target_style_identifier: 3,
            before: background.clone(),
            before_override: Some(background.clone()),
            after: Some(background.clone()),
            target_effective: background.clone(),
            target_override: Some(background),
            reset: false,
            touched_components: 0,
            deleted_previews: 0,
        };

        assert!(patch.artifacts.is_byte_noop());
        assert_eq!(patch.before, patch.target_effective);
        assert_eq!(patch.before_override, patch.target_override);
        assert_ne!(patch.style_identifier, patch.target_style_identifier);
        assert!(!patch.is_noop());
    }
}
