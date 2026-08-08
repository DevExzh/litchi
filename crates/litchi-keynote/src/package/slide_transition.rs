//! Exact-source, selector-first Keynote slide-transition transactions.
//!
//! This module deliberately keeps the protobuf representation private.  The
//! public transaction operates on `transition::Settings`; raw wire records are
//! only used to retain extension fields and the native envelope around a
//! transition.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::missing_errors_doc,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction deliberately redacts lower-layer errors and exhaustively maps bounded cross-crate error families."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::{
    WireLimits, encode_varint_into,
    wire::{
        WireView, patch_nested_fixed32_field, patch_nested_fixed64_field,
        patch_nested_length_delimited_field, patch_nested_varint_field,
    },
};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    keynote_slide_transition_codec::{
        DecodeOptions as TransitionDecodeOptions, TransitionSettingsSnapshot,
        decode_slide_node_has_transition, decode_slide_transition,
    },
    tsd, tsp,
};
use prost::Message;
use thiserror::Error;

use super::{
    Package, PhysicalSource, ReadError, SLIDE_MESSAGE_TYPE, SLIDE_NODE_MESSAGE_TYPE, unique_payload,
};
use crate::{
    SlideSelector,
    transition::{
        Acceleration, AnimationParameters, CustomParameters, Direction, Effect, Settings,
        TextDelivery, TimingCurveSlot,
    },
};

const TRANSITION_FIELD: u32 = 4;
const TRANSITION_ATTRIBUTES_FIELD: u32 = 2;
const NODE_HAS_TRANSITION_FIELD: u32 = 7;

// `patch_nested_*` rebuilds each enclosing message.  The focused transition
// rewrite makes sixteen animation-level and nine attributes-level patches.
// Keep an explicit bound ahead of those compatibility helpers, which still
// use their broad default profile internally.
const ANIMATION_PATCHES: usize = 16;
const CUSTOM_PATCHES: usize = 9;
const VARIABLE_PAYLOAD_PATCHES: usize = 9;
const MAX_NESTED_PATCH_GROWTH: usize = 48;

/// A finite resource governed while a transition transaction is prepared.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SlideTransitionLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// One protobuf payload.
    WireBytes,
    /// ZIP members or IWA objects.
    Entries,
    /// One package member or IWA value.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic text storage objects.
    TextStorages,
    /// Semantic text fragments.
    TextFragments,
    /// Aggregate semantic text.
    TextBytes,
    /// Parsed protobuf records.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
}

impl fmt::Display for SlideTransitionLimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// A content-redacted failure raised by a slide-transition transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTransitionError {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical slide-transition edits")]
    UnsupportedSource,
    /// An exact-name selector was ambiguous.
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic source position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// The source package or selected envelope is malformed or unsupported.
    #[error("the Keynote slide transition cannot be edited safely")]
    InvalidSource,
    /// Requested settings fail the archive-free transition invariants.
    #[error("the requested Keynote slide transition is invalid")]
    InvalidSettings,
    /// The selected slide has no modern transition attributes to patch.
    #[error("the Keynote slide transition uses an unsupported legacy representation")]
    UnsupportedTransition,
    /// One of the retained resource ceilings was exceeded.
    #[error(
        "Keynote slide-transition {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideTransitionLimitKind,
        observed: u64,
        maximum: u64,
    },
    /// Bounded allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote slide-transition transaction")]
    Allocation { amount: usize },
    /// Full reopening did not reproduce the requested semantic state.
    #[error("the edited Keynote slide transition failed semantic verification")]
    Verification,
    /// A patch was supplied to a package other than its exact source artifact.
    #[error("the Keynote slide-transition patch does not match the exact source package")]
    PatchConflict,
}

/// Mutable transition settings staged against one immutable package snapshot.
#[derive(Debug)]
pub struct SlideTransitionEdit<'a> {
    source: &'a Package,
    position: Position,
    before: Option<Settings>,
    settings: Option<Settings>,
}

impl<'a> SlideTransitionEdit<'a> {
    fn new(source: &'a Package, position: Position, before: Option<Settings>) -> Self {
        Self {
            source,
            position,
            settings: before.clone(),
            before,
        }
    }

    /// Borrow the transition that would be published by this edit.
    #[must_use]
    pub fn settings(&self) -> Option<&Settings> {
        self.settings.as_ref()
    }

    /// Replace the staged modern transition settings.
    ///
    /// A package with no modern transition is deliberately not synthesized by
    /// this API; callers can only edit the validated envelope they selected.
    ///
    /// # Errors
    ///
    /// Returns an error when `settings` violates the archive-free transition
    /// invariants or the selected source has no editable modern envelope.
    pub fn set_transition(
        &mut self,
        settings: Settings,
    ) -> Result<&mut Self, SlideTransitionError> {
        settings
            .validate()
            .map_err(|_| SlideTransitionError::InvalidSettings)?;
        if self.settings.is_none() {
            return Err(SlideTransitionError::UnsupportedTransition);
        }
        self.settings = Some(settings);
        Ok(self)
    }

    /// Stage Keynote's native no-effect transition representation.
    ///
    /// Unlike removing an envelope, this keeps a valid transition archive and
    /// is what Keynote itself writes when the user chooses "No Transition".
    pub fn clear(&mut self) -> Result<&mut Self, SlideTransitionError> {
        let current = self
            .settings
            .as_ref()
            .ok_or(SlideTransitionError::UnsupportedTransition)?;
        let mut animation = AnimationParameters::new();
        animation.set_random_number_seed(current.animation_parameters().random_number_seed());
        animation.set_writing_direction_is_rtl(
            current.animation_parameters().writing_direction_is_rtl(),
        );
        let mut settings = Settings::new();
        settings
            .set_animation_type(Some("Transition"))
            .map_err(|_error| SlideTransitionError::InvalidSettings)?;
        settings
            .set_effect(Some(Effect::None))
            .map_err(|_error| SlideTransitionError::InvalidSettings)?;
        settings
            .set_duration(Some(1.0))
            .map_err(|_error| SlideTransitionError::InvalidSettings)?;
        settings
            .set_delay(current.delay())
            .map_err(|_error| SlideTransitionError::InvalidSettings)?;
        settings.set_is_automatic(current.is_automatic());
        settings
            .set_animation_parameters(animation)
            .map_err(|_error| SlideTransitionError::InvalidSettings)?;
        settings
            .set_custom_parameters(CustomParameters::new())
            .map_err(|_error| SlideTransitionError::InvalidSettings)?;
        self.settings = Some(settings);
        Ok(self)
    }

    /// Validate and atomically publish the candidate.
    pub fn commit(self) -> Result<SlideTransitionCommit, SlideTransitionError> {
        if let Some(settings) = &self.settings {
            settings
                .validate()
                .map_err(|_| SlideTransitionError::InvalidSettings)?;
        }
        let catalog = physical_catalog(self.source)?;
        let source = catalog.shared_source();
        let source_fingerprint = fingerprint(&source);
        if self.source.slide_transition(self.position)? != self.before {
            return Err(SlideTransitionError::InvalidSource);
        }
        self.source.validate().map_err(map_read_error)?;
        if self.before == self.settings {
            return Ok(SlideTransitionCommit {
                package: self.source.snapshot(),
                patch: SlideTransitionPatch {
                    source: Arc::clone(&source),
                    target: source,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    position: self.position,
                    before: self.before,
                    after: self.settings,
                    touched_components: 0,
                },
                diagnostics: SlideTransitionDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(SlideTransitionError::UnsupportedSource);
        }
        let (package, touched) = rewrite_transition(
            self.source,
            self.position,
            self.before.as_ref(),
            self.settings.as_ref(),
        )?;
        let target = physical_catalog(&package)?.shared_source();
        Ok(SlideTransitionCommit {
            patch: SlideTransitionPatch {
                source,
                target: Arc::clone(&target),
                source_fingerprint,
                target_fingerprint: fingerprint(&target),
                position: self.position,
                before: self.before,
                after: self.settings,
                touched_components: touched,
            },
            package,
            diagnostics: SlideTransitionDiagnostics::published(touched),
        })
    }
}

/// An exact-source-checked reversible transition patch.
#[derive(Clone, PartialEq)]
pub struct SlideTransitionPatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    position: Position,
    before: Option<Settings>,
    after: Option<Settings>,
    touched_components: usize,
}

impl fmt::Debug for SlideTransitionPatch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideTransitionPatch")
            .field("position", &self.position)
            .field("has_before", &self.before.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}
impl SlideTransitionPatch {
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }
    #[must_use]
    pub fn before(&self) -> Option<&Settings> {
        self.before.as_ref()
    }
    #[must_use]
    pub fn after(&self) -> Option<&Settings> {
        self.after.as_ref()
    }
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
    }
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            position: self.position,
            before: self.after.clone(),
            after: self.before.clone(),
            touched_components: self.touched_components,
        }
    }
}

/// Compact publication evidence for one transition commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTransitionDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}
impl SlideTransitionDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }
    const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one transition transaction.
#[must_use = "a Keynote slide-transition commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTransitionCommit {
    package: Package,
    patch: SlideTransitionPatch,
    diagnostics: SlideTransitionDiagnostics,
}
impl SlideTransitionCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    #[must_use]
    pub const fn patch(&self) -> &SlideTransitionPatch {
        &self.patch
    }
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTransitionDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one slide's focused transition settings without exposing IWA IDs.
    pub fn slide_transition<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Option<Settings>, SlideTransitionError> {
        let position = self.resolve_transition_selector(selector.into())?;
        self.transition_at(position)
    }
    /// Start an immutable, selector-first transition edit.
    pub fn edit_slide_transition<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<SlideTransitionEdit<'_>, SlideTransitionError> {
        let position = self.resolve_transition_selector(selector.into())?;
        let before = self.transition_at(position)?;
        Ok(SlideTransitionEdit::new(self, position, before))
    }
    /// Apply an exact-source-checked reversible transition patch.
    pub fn apply_slide_transition(
        &self,
        patch: &SlideTransitionPatch,
    ) -> Result<SlideTransitionCommit, SlideTransitionError> {
        let catalog = physical_catalog(self)?;
        if fingerprint(catalog.source_bytes()) != patch.source_fingerprint
            || catalog.source_bytes() != patch.source.as_ref()
        {
            return Err(SlideTransitionError::PatchConflict);
        }
        if self.transition_at(patch.position)? != patch.before {
            return Err(SlideTransitionError::PatchConflict);
        }
        self.validate().map_err(map_read_error)?;
        if patch.is_noop() {
            if patch.source.as_ref() != patch.target.as_ref()
                || patch.source_fingerprint != patch.target_fingerprint
            {
                return Err(SlideTransitionError::PatchConflict);
            }
            return Ok(SlideTransitionCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTransitionDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() || fingerprint(&patch.target) != patch.target_fingerprint {
            return Err(SlideTransitionError::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(Arc::clone(&patch.target), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        if candidate.transition_at(patch.position)? != patch.after {
            return Err(SlideTransitionError::Verification);
        }
        Ok(SlideTransitionCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SlideTransitionDiagnostics::published(patch.touched_components),
        })
    }

    fn resolve_transition_selector(
        &self,
        selector: SlideSelector<'_>,
    ) -> Result<Position, SlideTransitionError> {
        match selector {
            SlideSelector::Position(position) => {
                if self
                    .slide_record_at(position.get())
                    .map_err(map_read_error)?
                    .is_some()
                {
                    Ok(position)
                } else {
                    Err(SlideTransitionError::SlidePositionNotFound { position })
                }
            },
            SlideSelector::Name(_) => self
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(|_| SlideTransitionError::AmbiguousSelector)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideTransitionError::SlideNameNotFound),
        }
    }

    fn transition_at(&self, position: Position) -> Result<Option<Settings>, SlideTransitionError> {
        let record = self
            .slide_record_at(position.get())
            .map_err(map_read_error)?
            .ok_or(SlideTransitionError::SlidePositionNotFound { position })?;
        let object = self
            .required_object(record.slide_identifier, "Keynote slide")
            .map_err(map_read_error)?;
        let payload = unique_payload(&object.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
            .map_err(map_read_error)?;
        let settings = settings_from_slide(payload, self.wire_limits().map_err(map_wire_error)?)?;
        let node = self
            .required_object(record.node_identifier, "Keynote slide node")
            .map_err(map_read_error)?;
        let node_payload = unique_payload(
            &node.messages,
            &[SLIDE_NODE_MESSAGE_TYPE],
            "Keynote slide node",
        )
        .map_err(map_read_error)?;
        let marker =
            strict_node_transition_flag(node_payload, self.wire_limits().map_err(map_wire_error)?)?;
        if settings.as_ref().is_some_and(Settings::has_effect) != marker && settings.is_some() {
            return Err(SlideTransitionError::InvalidSource);
        }
        Ok(settings)
    }
}

fn rewrite_transition(
    source: &Package,
    position: Position,
    before: Option<&Settings>,
    after: Option<&Settings>,
) -> Result<(Package, usize), SlideTransitionError> {
    let record = source
        .slide_record_at(position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTransitionError::InvalidSource)?;
    let catalog = physical_catalog(source)?;
    let mut targets = [
        (record.slide_identifier, false),
        (record.node_identifier, true),
    ];
    targets.sort_unstable_by_key(|target| target.0);
    let mut compressed_components: Vec<(&str, Vec<u8>)> = Vec::new();
    let mut component_names: Vec<&str> = Vec::new();
    for (identifier, _node) in targets {
        let mut found = catalog
            .components()
            .iter()
            .filter(|component| component.archive().object(identifier).is_some());
        let component = found.next().ok_or(SlideTransitionError::InvalidSource)?;
        if found.next().is_some() {
            return Err(SlideTransitionError::InvalidSource);
        }
        if component_names.iter().any(|name| *name == component.name()) {
            continue;
        }
        component_names.push(component.name());
    }
    // Each component is decompressed exactly once, and both objects are patched
    // before it is reassembled. This handles compact archives where node and
    // slide are co-located without reporting duplicate physical work.
    for name in &component_names {
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == *name)
            .ok_or(SlideTransitionError::InvalidSource)?;
        if entry.is_opaque() {
            return Err(SlideTransitionError::InvalidSource);
        }
        let archive_limits = source
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(map_archive_error)?;
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
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        if archive.object(record.slide_identifier).is_some() {
            let object = archive
                .object(record.slide_identifier)
                .ok_or(SlideTransitionError::InvalidSource)?;
            let (index, message) = one_message(object.messages.as_slice(), SLIDE_MESSAGE_TYPE)?;
            if settings_from_slide(&message.data, source.wire_limits().map_err(map_wire_error)?)?
                != before.cloned()
            {
                return Err(SlideTransitionError::InvalidSource);
            }
            let changed = rewrite_slide_payload(
                &message.data,
                after.ok_or(SlideTransitionError::UnsupportedTransition)?,
                source.wire_limits().map_err(map_wire_error)?,
            )?;
            archive
                .object_mut(record.slide_identifier)
                .ok_or(SlideTransitionError::InvalidSource)?
                .replace_message_preserving_header_with_limits(
                    index,
                    RawMessage {
                        type_: SLIDE_MESSAGE_TYPE,
                        data: changed,
                    },
                    archive_limits,
                )
                .map_err(map_core_error)?;
        }
        let mut node_changed = false;
        if archive.object(record.node_identifier).is_some() {
            let object = archive
                .object(record.node_identifier)
                .ok_or(SlideTransitionError::InvalidSource)?;
            let (index, message) =
                one_message(object.messages.as_slice(), SLIDE_NODE_MESSAGE_TYPE)?;
            let desired = after.is_some_and(Settings::has_effect);
            if strict_node_transition_flag(
                &message.data,
                source.wire_limits().map_err(map_wire_error)?,
            )? != desired
            {
                let changed = rewrite_node_transition_flag(
                    &message.data,
                    desired,
                    source.wire_limits().map_err(map_wire_error)?,
                )?;
                archive
                    .object_mut(record.node_identifier)
                    .ok_or(SlideTransitionError::InvalidSource)?
                    .replace_message_preserving_header_with_limits(
                        index,
                        RawMessage {
                            type_: SLIDE_NODE_MESSAGE_TYPE,
                            data: changed,
                        },
                        archive_limits,
                    )
                    .map_err(map_core_error)?;
                node_changed = true;
            }
        }
        if archive.object(record.slide_identifier).is_none() && !node_changed {
            continue;
        }
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
        compressed_components.push((name, compressed));
    }
    let edits: Vec<_> = compressed_components
        .iter()
        .map(|(name, compressed)| EntryEdit::new(name, compressed.as_slice()))
        .collect();
    let output = catalog
        .package()
        .reassemble_to_bytes(&edits, source.state.options.archive())
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    if candidate.transition_at(position)? != after.cloned() {
        return Err(SlideTransitionError::Verification);
    }
    Ok((candidate, compressed_components.len()))
}

fn one_message(
    messages: &[RawMessage],
    kind: u32,
) -> Result<(usize, &RawMessage), SlideTransitionError> {
    let mut matches = messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == kind);
    let item = matches.next().ok_or(SlideTransitionError::InvalidSource)?;
    if matches.next().is_some() {
        return Err(SlideTransitionError::InvalidSource);
    }
    Ok(item)
}

fn settings_from_slide(
    slide: &[u8],
    limits: WireLimits,
) -> Result<Option<Settings>, SlideTransitionError> {
    let projection = decode_slide_transition(
        slide,
        TransitionDecodeOptions::new(limits.max_input_bytes(), 4),
    )
    .map_err(|_| SlideTransitionError::InvalidSource)?;
    if projection.settings.animation.is_none() {
        Ok(None)
    } else {
        settings_from_projection(&projection.settings).map(Some)
    }
}

fn settings_from_snapshot(
    animation: litchi_iwa_protos::keynote_slide_transition_codec::AnimationSnapshot<'_>,
) -> Result<Settings, SlideTransitionError> {
    let mut settings = Settings::new();
    settings
        .set_animation_type(animation.animation_type)
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    let effect = animation
        .effect
        .map(Effect::from_identifier)
        .transpose()
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    settings
        .set_effect(effect)
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    settings
        .set_duration(animation.duration)
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    settings.set_direction(animation.direction.map(Direction::from_native));
    settings
        .set_delay(animation.delay)
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    settings.set_is_automatic(animation.is_automatic);
    let mut animation_parameters = AnimationParameters::new();
    // The focused settings value treats opaque submessages as exact payloads
    // when supplied through the public API. Generated projections are only
    // used for scalar semantic validation here.
    animation_parameters.set_random_number_seed(animation.random_number_seed);
    animation_parameters
        .set_detail(animation.custom_detail)
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    for (slot, value) in TimingCurveSlot::ALL.into_iter().zip([
        animation.custom_effect_timing_curve_theme_name_1,
        animation.custom_effect_timing_curve_theme_name_2,
        animation.custom_effect_timing_curve_theme_name_3,
    ]) {
        animation_parameters
            .set_timing_curve_theme_name(slot, value)
            .map_err(|_| SlideTransitionError::InvalidSource)?;
    }
    animation_parameters.set_writing_direction_is_rtl(animation.writing_direction_is_rtl);
    settings
        .set_animation_parameters(animation_parameters)
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    Ok(settings)
}

fn settings_from_projection(
    snapshot: &TransitionSettingsSnapshot<'_>,
) -> Result<Settings, SlideTransitionError> {
    let animation = snapshot
        .animation
        .ok_or(SlideTransitionError::UnsupportedTransition)?;
    let mut settings = settings_from_snapshot(animation)?;
    let mut parameters = settings.animation_parameters().clone();
    if let Some(payload) = animation.color {
        tsp::Color::decode(payload).map_err(|_error| SlideTransitionError::InvalidSource)?;
        parameters
            .set_color_payload(Some(payload))
            .map_err(|_error| SlideTransitionError::InvalidSource)?;
    }
    for (slot, payload) in TimingCurveSlot::ALL.into_iter().zip([
        animation.custom_effect_timing_curve_1,
        animation.custom_effect_timing_curve_2,
        animation.custom_effect_timing_curve_3,
    ]) {
        if let Some(payload) = payload {
            tsd::PathSourceArchive::decode(payload)
                .map_err(|_error| SlideTransitionError::InvalidSource)?;
            parameters
                .set_timing_curve_payload(slot, Some(payload))
                .map_err(|_error| SlideTransitionError::InvalidSource)?;
        }
    }
    settings
        .set_animation_parameters(parameters)
        .map_err(|_error| SlideTransitionError::InvalidSource)?;
    let mut custom = CustomParameters::new();
    custom
        .set_twist(snapshot.custom_twist)
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    custom.set_mosaic_size(snapshot.custom_mosaic_size);
    custom.set_mosaic_type(
        snapshot
            .custom_mosaic_type
            .map(crate::transition::MosaicType::from_native),
    );
    custom.set_bounce(snapshot.custom_bounce);
    custom.set_magic_move_fade_unmatched_objects(snapshot.custom_magic_move_fade_unmatched_objects);
    custom.set_acceleration(snapshot.custom_timing_curve.map(Acceleration::from_native));
    custom.set_text_delivery(
        snapshot
            .custom_text_delivery_type
            .map(TextDelivery::from_native),
    );
    custom.set_motion_blur(snapshot.custom_motion_blur);
    custom
        .set_travel_distance(snapshot.custom_travel_distance)
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    settings
        .set_custom_parameters(custom)
        .map_err(|_| SlideTransitionError::InvalidSource)?;
    Ok(settings)
}

fn strict_node_transition_flag(
    source: &[u8],
    limits: WireLimits,
) -> Result<bool, SlideTransitionError> {
    decode_slide_node_has_transition(
        source,
        TransitionDecodeOptions::new(limits.max_input_bytes(), 1),
    )
    .map_err(|_| SlideTransitionError::InvalidSource)
}

fn rewrite_slide_payload(
    source: &[u8],
    after: &Settings,
    limits: WireLimits,
) -> Result<Vec<u8>, SlideTransitionError> {
    let projection = decode_slide_transition(
        source,
        TransitionDecodeOptions::new(limits.max_input_bytes(), 4),
    )
    .map_err(|_error| SlideTransitionError::InvalidSource)?;
    if projection.settings.animation.is_none() {
        return Err(SlideTransitionError::UnsupportedTransition);
    }
    rewrite_transition_envelope(source, &projection.settings, after, limits)
}

fn rewrite_transition_envelope(
    source: &[u8],
    snapshot: &TransitionSettingsSnapshot<'_>,
    settings: &Settings,
    limits: WireLimits,
) -> Result<Vec<u8>, SlideTransitionError> {
    let animation = snapshot
        .animation
        .ok_or(SlideTransitionError::UnsupportedTransition)?;
    preflight_nested_patch_budget(source, snapshot, settings, limits)?;
    let parameters = settings.animation_parameters();
    let custom = settings.custom_parameters();
    let animation_path = |field| [TRANSITION_FIELD, TRANSITION_ATTRIBUTES_FIELD, 8, field];
    let custom_path = |field| [TRANSITION_FIELD, TRANSITION_ATTRIBUTES_FIELD, field];
    let mut data = patch_nested_length_delimited_field(
        source,
        &animation_path(1),
        animation.animation_type.is_some(),
        settings.animation_type().map(str::as_bytes),
    )
    .map_err(map_wire_error)?;
    data = patch_nested_length_delimited_field(
        &data,
        &animation_path(2),
        animation.effect.is_some(),
        settings.effect().map(Effect::identifier).map(str::as_bytes),
    )
    .map_err(map_wire_error)?;
    for (field, current, replacement) in [
        (3, animation.duration, settings.duration()),
        (5, animation.delay, settings.delay()),
    ] {
        data = patch_nested_fixed64_field(
            &data,
            &animation_path(field),
            current.is_some(),
            replacement.map(f64::to_bits),
        )
        .map_err(map_wire_error)?;
    }
    data = patch_nested_varint_field(
        &data,
        &animation_path(4),
        animation.direction.is_some(),
        settings
            .direction()
            .map(|value| u64::from(value.native_value())),
    )
    .map_err(map_wire_error)?;
    data = patch_nested_varint_field(
        &data,
        &animation_path(6),
        animation.is_automatic.is_some(),
        settings.is_automatic().map(u64::from),
    )
    .map_err(map_wire_error)?;
    data = patch_nested_length_delimited_field(
        &data,
        &animation_path(7),
        animation.color.is_some(),
        parameters.color_payload(),
    )
    .map_err(map_wire_error)?;
    for (slot, field, current) in [
        (
            TimingCurveSlot::First,
            8,
            animation.custom_effect_timing_curve_1.is_some(),
        ),
        (
            TimingCurveSlot::Second,
            9,
            animation.custom_effect_timing_curve_2.is_some(),
        ),
        (
            TimingCurveSlot::Third,
            10,
            animation.custom_effect_timing_curve_3.is_some(),
        ),
    ] {
        data = patch_nested_length_delimited_field(
            &data,
            &animation_path(field),
            current,
            parameters.timing_curve_payload(slot),
        )
        .map_err(map_wire_error)?;
    }
    data = patch_nested_varint_field(
        &data,
        &animation_path(11),
        animation.random_number_seed.is_some(),
        parameters.random_number_seed().map(u64::from),
    )
    .map_err(map_wire_error)?;
    data = patch_nested_fixed64_field(
        &data,
        &animation_path(12),
        animation.custom_detail.is_some(),
        parameters.detail().map(f64::to_bits),
    )
    .map_err(map_wire_error)?;
    for (slot, field, current) in [
        (
            TimingCurveSlot::First,
            13,
            animation.custom_effect_timing_curve_theme_name_1.is_some(),
        ),
        (
            TimingCurveSlot::Second,
            14,
            animation.custom_effect_timing_curve_theme_name_2.is_some(),
        ),
        (
            TimingCurveSlot::Third,
            15,
            animation.custom_effect_timing_curve_theme_name_3.is_some(),
        ),
    ] {
        data = patch_nested_length_delimited_field(
            &data,
            &animation_path(field),
            current,
            parameters.timing_curve_theme_name(slot).map(str::as_bytes),
        )
        .map_err(map_wire_error)?;
    }
    data = patch_nested_varint_field(
        &data,
        &animation_path(16),
        animation.writing_direction_is_rtl.is_some(),
        parameters.writing_direction_is_rtl().map(u64::from),
    )
    .map_err(map_wire_error)?;
    for (field, current, replacement) in [
        (9, snapshot.custom_twist, custom.twist()),
        (
            18,
            snapshot.custom_travel_distance,
            custom.travel_distance(),
        ),
    ] {
        data = patch_nested_fixed32_field(
            &data,
            &custom_path(field),
            current.is_some(),
            replacement.map(f32::to_bits),
        )
        .map_err(map_wire_error)?;
    }
    for (field, current, replacement) in [
        (
            10,
            snapshot.custom_mosaic_size.map(u64::from),
            custom.mosaic_size().map(u64::from),
        ),
        (
            11,
            snapshot.custom_mosaic_type.map(u64::from),
            custom
                .mosaic_type()
                .map(|value| u64::from(value.native_value())),
        ),
        (
            12,
            snapshot.custom_bounce.map(u64::from),
            custom.bounce().map(u64::from),
        ),
        (
            13,
            snapshot
                .custom_magic_move_fade_unmatched_objects
                .map(u64::from),
            custom.magic_move_fade_unmatched_objects().map(u64::from),
        ),
        (
            15,
            snapshot
                .custom_timing_curve
                .map(|value| i64::from(value) as u64),
            custom
                .acceleration()
                .map(|value| i64::from(value.native_value()) as u64),
        ),
        (
            16,
            snapshot
                .custom_text_delivery_type
                .map(|value| i64::from(value) as u64),
            custom
                .text_delivery()
                .map(|value| i64::from(value.native_value()) as u64),
        ),
        (
            17,
            snapshot.custom_motion_blur.map(u64::from),
            custom.motion_blur().map(u64::from),
        ),
    ] {
        data =
            patch_nested_varint_field(&data, &custom_path(field), current.is_some(), replacement)
                .map_err(map_wire_error)?;
    }
    check_output(&data, limits)?;
    Ok(data)
}

/// Bound the legacy nested patch helpers before any of them allocates.
///
/// They reconstruct the root and every selected ancestor for each leaf and
/// currently accept only their default `WireLimits`.  The source was already
/// strict-projected, but a final output check alone would let an intermediate
/// replacement exceed this transaction's tighter package profile.  This
/// conservative preflight caps all intermediate outputs and the repeated
/// field scans under the caller's retained limits.
fn preflight_nested_patch_budget(
    source: &[u8],
    snapshot: &TransitionSettingsSnapshot<'_>,
    settings: &Settings,
    limits: WireLimits,
) -> Result<(), SlideTransitionError> {
    let root = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let transition = required_nested_payload(&root, TRANSITION_FIELD)?;
    let transition = WireView::parse_with_limits(transition, limits).map_err(map_wire_error)?;
    let attributes = required_nested_payload(&transition, TRANSITION_ATTRIBUTES_FIELD)?;
    let attributes = WireView::parse_with_limits(attributes, limits).map_err(map_wire_error)?;
    let animation = required_nested_payload(&attributes, 8)?;
    let animation = WireView::parse_with_limits(animation, limits).map_err(map_wire_error)?;

    let root_fields = root.len();
    let transition_fields = transition.len();
    let attributes_fields = attributes.len();
    let animation_fields = animation.len();
    let animation_scan = root_fields
        .checked_add(transition_fields)
        .and_then(|value| value.checked_add(attributes_fields))
        .and_then(|value| value.checked_add(animation_fields))
        .and_then(|value| value.checked_add(ANIMATION_PATCHES))
        .and_then(|value| value.checked_mul(ANIMATION_PATCHES));
    let custom_scan = root_fields
        .checked_add(transition_fields)
        .and_then(|value| value.checked_add(attributes_fields))
        .and_then(|value| value.checked_add(CUSTOM_PATCHES))
        .and_then(|value| value.checked_mul(CUSTOM_PATCHES));
    // Every parsed field is also copied once while rebuilding that message.
    let rewrite_work = animation_scan
        .and_then(|value| custom_scan.and_then(|custom| value.checked_add(custom)))
        .and_then(|value| value.checked_mul(2))
        .ok_or_else(|| rewrite_limit_error(usize::MAX, limits))?;
    if rewrite_work > limits.max_rewrite_work() {
        return Err(rewrite_limit_error(rewrite_work, limits));
    }

    let projected = snapshot
        .animation
        .ok_or(SlideTransitionError::UnsupportedTransition)?;
    let parameters = settings.animation_parameters();
    let output_bound = [
        (
            projected.animation_type.map(str::as_bytes),
            settings.animation_type().map(str::as_bytes),
        ),
        (
            projected.effect.map(str::as_bytes),
            settings.effect().map(Effect::identifier).map(str::as_bytes),
        ),
        (projected.color, parameters.color_payload()),
        (
            projected.custom_effect_timing_curve_1,
            parameters.timing_curve_payload(TimingCurveSlot::First),
        ),
        (
            projected.custom_effect_timing_curve_2,
            parameters.timing_curve_payload(TimingCurveSlot::Second),
        ),
        (
            projected.custom_effect_timing_curve_3,
            parameters.timing_curve_payload(TimingCurveSlot::Third),
        ),
        (
            projected
                .custom_effect_timing_curve_theme_name_1
                .map(str::as_bytes),
            parameters
                .timing_curve_theme_name(TimingCurveSlot::First)
                .map(str::as_bytes),
        ),
        (
            projected
                .custom_effect_timing_curve_theme_name_2
                .map(str::as_bytes),
            parameters
                .timing_curve_theme_name(TimingCurveSlot::Second)
                .map(str::as_bytes),
        ),
        (
            projected
                .custom_effect_timing_curve_theme_name_3
                .map(str::as_bytes),
            parameters
                .timing_curve_theme_name(TimingCurveSlot::Third)
                .map(str::as_bytes),
        ),
    ]
    .into_iter()
    .try_fold(0usize, |total, (before, after)| {
        payload_patch_growth(before, after).and_then(|growth| total.checked_add(growth))
    })
    .and_then(|growth| source.len().checked_add(growth))
    .and_then(|total| {
        total.checked_add(
            (ANIMATION_PATCHES + CUSTOM_PATCHES - VARIABLE_PAYLOAD_PATCHES)
                .checked_mul(MAX_NESTED_PATCH_GROWTH)?,
        )
    })
    .ok_or_else(|| output_limit_error(usize::MAX, limits))?;
    if output_bound > limits.max_output_bytes() {
        return Err(output_limit_error(output_bound, limits));
    }
    Ok(())
}

fn payload_patch_growth(before: Option<&[u8]>, after: Option<&[u8]>) -> Option<usize> {
    let Some(after) = after else {
        return Some(0);
    };
    let growth = before.map_or(after.len(), |before| {
        after.len().saturating_sub(before.len())
    });
    if growth == 0 {
        Some(0)
    } else {
        growth.checked_add(MAX_NESTED_PATCH_GROWTH)
    }
}

fn required_nested_payload<'a>(
    view: &WireView<'a>,
    field_number: u32,
) -> Result<&'a [u8], SlideTransitionError> {
    let mut matches = view.fields().filter(|field| field.number() == field_number);
    let field = matches.next().ok_or(SlideTransitionError::InvalidSource)?;
    if matches.next().is_some() || field.wire_type() != 2 {
        return Err(SlideTransitionError::InvalidSource);
    }
    Ok(field.payload())
}

fn output_limit_error(observed: usize, limits: WireLimits) -> SlideTransitionError {
    SlideTransitionError::LimitExceeded {
        kind: SlideTransitionLimitKind::OutputBytes,
        observed: observed as u64,
        maximum: limits.max_output_bytes() as u64,
    }
}

fn rewrite_limit_error(observed: usize, limits: WireLimits) -> SlideTransitionError {
    SlideTransitionError::LimitExceeded {
        kind: SlideTransitionLimitKind::WireWork,
        observed: observed as u64,
        maximum: limits.max_rewrite_work() as u64,
    }
}

fn rewrite_node_transition_flag(
    source: &[u8],
    value: bool,
    limits: WireLimits,
) -> Result<Vec<u8>, SlideTransitionError> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut seen = false;
    let mut output = Vec::new();
    output
        .try_reserve_exact(source.len())
        .map_err(|_| SlideTransitionError::Allocation {
            amount: source.len(),
        })?;
    for field in view.fields() {
        if field.number() == NODE_HAS_TRANSITION_FIELD {
            if seen || field.wire_type() != 0 {
                return Err(SlideTransitionError::InvalidSource);
            }
            seen = true;
            append_varint_field(&mut output, NODE_HAS_TRANSITION_FIELD, u64::from(value));
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !seen {
        return Err(SlideTransitionError::InvalidSource);
    }
    check_output(&output, limits)?;
    Ok(output)
}

fn append_varint_field(output: &mut Vec<u8>, number: u32, value: u64) {
    encode_varint_into(output, u64::from(number) << 3);
    encode_varint_into(output, value);
}
fn check_output(output: &[u8], limits: WireLimits) -> Result<(), SlideTransitionError> {
    if output.len() > limits.max_output_bytes() {
        Err(SlideTransitionError::LimitExceeded {
            kind: SlideTransitionLimitKind::OutputBytes,
            observed: output.len() as u64,
            maximum: limits.max_output_bytes() as u64,
        })
    } else {
        Ok(())
    }
}

#[allow(
    clippy::unnecessary_wraps,
    reason = "the semantic-only source feature adds a typed unsupported-source branch"
)]
fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideTransitionError> {
    #[cfg(feature = "internal-iwork-source")]
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideTransitionError::UnsupportedSource),
    }
    #[cfg(not(feature = "internal-iwork-source"))]
    {
        let PhysicalSource::Package(source) = &package.state.source;
        Ok(source)
    }
}
fn map_read_error(error: ReadError) -> SlideTransitionError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTransitionError::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => SlideTransitionLimitKind::Entries,
                super::SemanticLimitKind::Slides => SlideTransitionLimitKind::Slides,
                super::SemanticLimitKind::References => SlideTransitionLimitKind::References,
                super::SemanticLimitKind::TextStorages => SlideTransitionLimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => SlideTransitionLimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => SlideTransitionLimitKind::TextBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTransitionError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideTransitionLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideTransitionLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideTransitionLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideTransitionLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTransitionError::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        _ => SlideTransitionError::InvalidSource,
    }
}
fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTransitionError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTransitionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideTransitionLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => SlideTransitionLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => SlideTransitionLimitKind::Entries,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    SlideTransitionLimitKind::TotalBytes
                },
                _ => SlideTransitionLimitKind::EntryBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTransitionError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideTransitionError::InvalidSource,
    }
}
fn map_core_error(error: litchi_iwa_core::Error) -> SlideTransitionError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTransitionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::MessageBytes => SlideTransitionLimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => SlideTransitionLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => SlideTransitionLimitKind::WireNesting,
                _ => SlideTransitionLimitKind::EntryBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTransitionError::Allocation { amount: requested }
        },
        _ => SlideTransitionError::InvalidSource,
    }
}
fn map_wire_error(error: litchi_iwa_common::Error) -> SlideTransitionError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SlideTransitionError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => SlideTransitionLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => SlideTransitionLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    SlideTransitionLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => SlideTransitionLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => SlideTransitionLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideTransitionError::Allocation { amount }
        },
        _ => SlideTransitionError::InvalidSource,
    }
}
fn fingerprint(bytes: &[u8]) -> u64 {
    bytes.iter().fold(0xcbf2_9ce4_8422_2325_u64, |value, byte| {
        (value ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3)
    })
}
