//! Exact-source, selector-first Keynote slide-transition transactions.
//!
//! This module deliberately keeps the protobuf representation private.  The
//! public transaction operates on `transition::Settings`; raw wire records are
//! only used to retain extension fields and the native envelope around a
//! transition.

#![allow(
    clippy::cast_sign_loss,
    clippy::arbitrary_source_item_ordering,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction deliberately redacts lower-layer errors and exhaustively maps bounded cross-crate error families."
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes,
    wire::{
        NestedFieldEdit, NestedFieldReplacement, WireDescent, WireFieldView, WireView,
        patch_nested_fields_batched_with_limits, preflight_wire_tree_with_limits,
    },
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    keynote_document_codec, keynote_show_codec,
    keynote_slide_transition_codec::{
        DecodeOptions as TransitionDecodeOptions, TransitionSettingsSnapshot,
        decode_slide_node_has_transition, decode_slide_transition,
    },
    tsd, tsp,
};
use prost::Message;
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SLIDE_MESSAGE_TYPE, SLIDE_NODE_MESSAGE_TYPE};
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
const DOCUMENT_SHOW_FIELD: u32 = 2;
const SHOW_SLIDE_TREE_FIELD: u32 = 3;
const SLIDE_TREE_SLIDES_FIELD: u32 = 2;
const NODE_SLIDE_FIELD: u32 = 2;

const ANIMATION_PATHS: [[u32; 4]; 16] = [
    [4, 2, 8, 1],
    [4, 2, 8, 2],
    [4, 2, 8, 3],
    [4, 2, 8, 4],
    [4, 2, 8, 5],
    [4, 2, 8, 6],
    [4, 2, 8, 7],
    [4, 2, 8, 8],
    [4, 2, 8, 9],
    [4, 2, 8, 10],
    [4, 2, 8, 11],
    [4, 2, 8, 12],
    [4, 2, 8, 13],
    [4, 2, 8, 14],
    [4, 2, 8, 15],
    [4, 2, 8, 16],
];
const CUSTOM_PATHS: [[u32; 3]; 9] = [
    [4, 2, 9],
    [4, 2, 10],
    [4, 2, 11],
    [4, 2, 12],
    [4, 2, 13],
    [4, 2, 15],
    [4, 2, 16],
    [4, 2, 17],
    [4, 2, 18],
];

/// A finite resource governed while a transition transaction is prepared.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum LimitKind {
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

impl fmt::Display for LimitKind {
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
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum Error {
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
    InvalidSettings {
        /// Content-redacted semantic validation cause.
        #[source]
        source: crate::Error,
    },
    /// A requested opaque color or timing-curve payload is malformed.
    #[error("the requested Keynote slide transition contains an invalid opaque payload")]
    InvalidOpaquePayload,
    /// The selected slide has no modern transition attributes to patch.
    #[error("the Keynote slide transition uses an unsupported legacy representation")]
    UnsupportedTransition,
    /// One of the retained resource ceilings was exceeded.
    #[error(
        "Keynote slide-transition {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: LimitKind,
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
pub struct Edit<'a> {
    source: &'a Package,
    position: Position,
    node_identifier: u64,
    slide_identifier: u64,
    before: Option<Arc<Settings>>,
    settings: Option<Arc<Settings>>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("position", &self.position)
            .field("has_before", &self.before.is_some())
            .field("has_staged", &self.settings.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn new(source: &'a Package, selection: Selection<'_>) -> Self {
        let before = selection.settings.map(Arc::new);
        Self {
            source,
            position: selection.position,
            node_identifier: selection.node_identifier,
            slide_identifier: selection.slide_identifier,
            settings: before.as_ref().map(Arc::clone),
            before,
        }
    }

    /// Borrow the transition that would be published by this edit.
    #[must_use]
    pub fn settings(&self) -> Option<&Settings> {
        self.settings.as_deref()
    }

    /// Replace the staged modern transition settings.
    ///
    /// A package with no modern transition is deliberately not synthesized by
    /// this API; callers can only edit the validated envelope they selected.
    ///
    /// # Errors
    ///
    /// Unknown native discriminants and opaque payloads may be synthesized;
    /// opaque payloads are strictly bounded and structurally validated before
    /// they are retained byte-for-byte. Returns an error when `settings`
    /// violates those invariants or the selected source has no editable modern
    /// envelope.
    ///
    /// # Costs
    ///
    /// Moves `settings` into one shared allocation. Opaque validation is linear
    /// in encoded bytes and does not clone them again.
    pub fn set(mut self, settings: Settings) -> Result<Self, Error> {
        settings
            .validate()
            .map_err(|source| Error::InvalidSettings { source })?;
        if self.settings.is_none() {
            return Err(Error::UnsupportedTransition);
        }
        validate_requested_opaque_settings(
            &settings,
            self.source.wire_limits().map_err(map_wire_error)?,
        )?;
        self.settings = Some(Arc::new(settings));
        Ok(self)
    }

    /// Stage Keynote's native no-effect transition representation.
    ///
    /// Unlike removing an envelope, this keeps a valid transition archive and
    /// is what Keynote itself writes when the user chooses "No Transition".
    ///
    /// # Errors
    ///
    /// Returns a typed settings error if the native no-effect representation
    /// cannot be staged. An absent modern envelope is a successful exact no-op.
    ///
    /// # Costs
    ///
    /// Allocates only compact no-effect settings. The absent path performs no
    /// allocation or package work.
    pub fn clear(mut self) -> Result<Self, Error> {
        let Some(current) = self.settings.as_deref() else {
            return Ok(self);
        };
        let mut animation = AnimationParameters::new();
        animation.set_random_number_seed(current.animation_parameters().random_number_seed());
        animation.set_writing_direction_is_rtl(
            current.animation_parameters().writing_direction_is_rtl(),
        );
        let mut settings = Settings::new();
        settings
            .set_animation_type(Some("Transition"))
            .map_err(|source| Error::InvalidSettings { source })?;
        settings
            .set_effect(Some(Effect::None))
            .map_err(|source| Error::InvalidSettings { source })?;
        settings
            .set_duration(Some(1.0))
            .map_err(|source| Error::InvalidSettings { source })?;
        settings
            .set_delay(current.delay())
            .map_err(|source| Error::InvalidSettings { source })?;
        settings.set_is_automatic(current.is_automatic());
        settings
            .set_animation_parameters(animation)
            .map_err(|source| Error::InvalidSettings { source })?;
        settings
            .set_custom_parameters(CustomParameters::new())
            .map_err(|source| Error::InvalidSettings { source })?;
        self.settings = Some(Arc::new(settings));
        Ok(self)
    }

    /// Validate and atomically publish the candidate.
    ///
    /// # Errors
    ///
    /// Returns a typed source, resource, allocation, settings, or verification
    /// error without publishing a partial package.
    ///
    /// # Costs
    ///
    /// A semantic no-op performs no archive rewrite or candidate reopen. A
    /// change rewrites at most the selected slide/node components and opens one
    /// candidate package.
    pub fn commit(self) -> Result<Commit, Error> {
        if let Some(settings) = &self.settings {
            settings
                .validate()
                .map_err(|source| Error::InvalidSettings { source })?;
        }
        let catalog = physical_catalog(self.source)?;
        let source = catalog.shared_source();
        if self.before.as_deref() == self.settings.as_deref() {
            return Ok(Commit {
                package: self.source.snapshot(),
                patch: Patch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source), source),
                    position: self.position,
                    node_identifier: self.node_identifier,
                    slide_identifier: self.slide_identifier,
                    before: self.before,
                    after: self.settings,
                    touched_components: 0,
                },
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::UnsupportedSource);
        }
        let selection = select_transition(self.source, self.position, true)?;
        if selection.node_identifier != self.node_identifier
            || selection.slide_identifier != self.slide_identifier
            || selection.settings.as_ref() != self.before.as_deref()
        {
            return Err(Error::InvalidSource);
        }
        let (package, touched) =
            rewrite_transition(self.source, &selection, self.settings.as_deref())?;
        let target = physical_catalog(&package)?.shared_source();
        Ok(Commit {
            patch: Patch {
                artifacts: ExactArtifacts::new(source, Arc::clone(&target)),
                position: self.position,
                node_identifier: self.node_identifier,
                slide_identifier: self.slide_identifier,
                before: self.before,
                after: self.settings,
                touched_components: touched,
            },
            package,
            diagnostics: Diagnostics::published(touched),
        })
    }
}

/// An exact-source-checked reversible transition patch.
///
/// Patches are process-local capabilities retaining shared exact source and
/// target artifacts, not a stable serialized interchange format. Cloning and
/// inversion are O(1) in package and opaque-payload bytes.
#[derive(Clone, PartialEq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    position: Position,
    node_identifier: u64,
    slide_identifier: u64,
    before: Option<Arc<Settings>>,
    after: Option<Arc<Settings>>,
    touched_components: usize,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch")
            .field("position", &self.position)
            .field("has_before", &self.before.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}
impl Patch {
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }
    #[must_use]
    pub fn before(&self) -> Option<&Settings> {
        self.before.as_deref()
    }
    #[must_use]
    pub fn after(&self) -> Option<&Settings> {
        self.after.as_deref()
    }
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_deref() == self.after.as_deref() && self.artifacts.is_byte_noop()
    }
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            position: self.position,
            node_identifier: self.node_identifier,
            slide_identifier: self.slide_identifier,
            before: self.after.clone(),
            after: self.before.clone(),
            touched_components: self.touched_components,
        }
    }
}

/// Compact publication evidence for one transition commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}
impl Diagnostics {
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
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}
impl Commit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

struct Selection<'source> {
    position: Position,
    show_identifier: u64,
    node_identifier: u64,
    slide_identifier: u64,
    node_component_name: &'source str,
    slide_component_name: &'source str,
    slide_payload: &'source [u8],
    node_payload: &'source [u8],
    transition_snapshot: TransitionSettingsSnapshot<'source>,
    settings: Option<Settings>,
}

fn select_transition(
    package: &Package,
    position: Position,
    mutation_guards: bool,
) -> Result<Selection<'_>, Error> {
    let catalog = package.state.source.components();
    let wire_limits = package.wire_limits().map_err(map_wire_error)?;
    let semantic_wire_limits = package.semantic_wire_limits().map_err(map_read_error)?;

    let mut roots = catalog
        .iter()
        .filter(|component| component.name().rsplit('/').next() == Some("Document.iwa"));
    let root_component = roots.next().ok_or(Error::InvalidSource)?;
    if roots.next().is_some() {
        return Err(Error::InvalidSource);
    }
    let root = root_component
        .archive()
        .object(1)
        .ok_or(Error::InvalidSource)?;
    let (root_message_index, root_payload) = selected_message(root, super::DOCUMENT_MESSAGE_TYPE)?;
    let recursion_limit =
        u32::try_from(semantic_wire_limits.max_nesting()).map_err(|_error| Error::InvalidSource)?;
    let root_reference = keynote_document_codec::decode_show_reference(
        root_payload,
        keynote_document_codec::DecodeOptions::new(root_payload.len(), recursion_limit)
            .with_max_fields(semantic_wire_limits.max_fields())
            .with_max_work_bytes(semantic_wire_limits.max_rewrite_work()),
    )
    .map_err(map_document_codec_error)?;
    if root_reference.identifier() == 0 || root_reference.deprecated_is_external() == Some(true) {
        return Err(Error::InvalidSource);
    }
    let show_identifier = root_reference.identifier();
    let (show_component_name, show) = unique_component_object(package, show_identifier)?;
    let (show_message_index, show_payload) = selected_message(show, super::SHOW_MESSAGE_TYPE)?;
    let show_snapshot = keynote_show_codec::decode_show(
        show_payload,
        keynote_show_codec::DecodeOptions::new(
            show_payload.len(),
            package.semantic_limits().max_slides(),
            recursion_limit,
        )
        .with_max_fields(semantic_wire_limits.max_fields())
        .with_max_work_bytes(semantic_wire_limits.max_rewrite_work()),
    )
    .map_err(map_show_codec_error)?;
    let node_identifiers = show_snapshot.slide_node_identifiers();
    let node_identifier = *node_identifiers
        .get(position.get())
        .ok_or(Error::SlidePositionNotFound { position })?;
    if node_identifiers
        .iter()
        .filter(|candidate| **candidate == node_identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource);
    }
    let raw_node_identifiers = strict_show_slide_references(show_payload, wire_limits)?;
    if raw_node_identifiers.as_slice() != node_identifiers {
        return Err(Error::InvalidSource);
    }

    let (node_component_name, node) = unique_component_object(package, node_identifier)?;
    let (node_message_index, node_payload) = selected_message(node, SLIDE_NODE_MESSAGE_TYPE)?;
    let mut owner_work = 0usize;
    let slide_identifier = strict_node_slide_reference(node_payload, wire_limits, &mut owner_work)?;
    let mut selected_slide_occurrences = 0usize;
    for candidate_node_identifier in node_identifiers {
        let candidate_slide_identifier = if *candidate_node_identifier == node_identifier {
            slide_identifier
        } else {
            let (_name, candidate_node) =
                unique_component_object(package, *candidate_node_identifier)?;
            let (_index, candidate_payload) =
                selected_message(candidate_node, SLIDE_NODE_MESSAGE_TYPE)?;
            strict_node_slide_reference(candidate_payload, wire_limits, &mut owner_work)?
        };
        if candidate_slide_identifier == slide_identifier {
            selected_slide_occurrences = selected_slide_occurrences
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
        }
    }
    if selected_slide_occurrences != 1 {
        return Err(Error::InvalidSource);
    }

    let (slide_component_name, slide) = unique_component_object(package, slide_identifier)?;
    let (slide_message_index, slide_payload) = selected_message(slide, SLIDE_MESSAGE_TYPE)?;
    let projection = transition_projection(slide_payload, wire_limits)?;
    let transition_snapshot = projection.settings;
    let settings = if transition_snapshot.animation.is_some() {
        Some(settings_from_projection(&transition_snapshot, wire_limits)?)
    } else {
        None
    };
    let marker = strict_node_transition_flag(node_payload, wire_limits)?;
    if settings.as_ref().is_some_and(Settings::has_effect) != marker && settings.is_some() {
        return Err(Error::InvalidSource);
    }

    if mutation_guards {
        if projection.settings.has_legacy_database_fields {
            return Err(Error::InvalidSource);
        }
        validate_reference_metadata(
            root,
            root_message_index,
            show_identifier,
            &[DOCUMENT_SHOW_FIELD],
        )?;
        validate_reference_metadata(
            show,
            show_message_index,
            node_identifier,
            &[SHOW_SLIDE_TREE_FIELD, SLIDE_TREE_SLIDES_FIELD],
        )?;
        validate_reference_metadata(
            node,
            node_message_index,
            slide_identifier,
            &[NODE_SLIDE_FIELD],
        )?;
        for (object, message_index) in [
            (root, root_message_index),
            (show, show_message_index),
            (node, node_message_index),
            (slide, slide_message_index),
        ] {
            validate_selected_metadata(object, message_index)?;
        }
        let component_names = [
            root_component.name(),
            show_component_name,
            node_component_name,
            slide_component_name,
        ];
        for (index, component_name) in component_names.into_iter().enumerate() {
            if component_names[..index].contains(&component_name) {
                continue;
            }
            validate_component_framing(package, component_name)?;
        }
        for payload in [root_payload, show_payload, node_payload, slide_payload] {
            reject_top_level_groups(payload, wire_limits)?;
        }
    }

    Ok(Selection {
        position,
        show_identifier,
        node_identifier,
        slide_identifier,
        node_component_name,
        slide_component_name,
        slide_payload,
        node_payload,
        transition_snapshot,
        settings,
    })
}

fn unique_component_object(
    package: &Package,
    identifier: u64,
) -> Result<(&str, &ArchiveObject), Error> {
    // Package construction rejects duplicate native identities and retains a
    // sorted locator index, so this lookup is globally unique and logarithmic.
    package
        .object_with_component(identifier)
        .ok_or(Error::InvalidSource)
}

fn selected_message(object: &ArchiveObject, message_type: u32) -> Result<(usize, &[u8]), Error> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(Error::InvalidSource);
    }
    let mut selected = None;
    for (index, (message, info)) in object
        .messages
        .iter()
        .zip(&object.archive_info.message_infos)
        .enumerate()
    {
        if message.type_ != info.type_
            || usize::try_from(info.length).ok() != Some(message.data.len())
        {
            return Err(Error::InvalidSource);
        }
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(Error::InvalidSource);
        }
    }
    selected.ok_or(Error::InvalidSource)
}

fn validate_selected_metadata(object: &ArchiveObject, message_index: usize) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn validate_reference_metadata(
    object: &ArchiveObject,
    message_index: usize,
    identifier: u64,
    path: &[u32],
) -> Result<(), Error> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource)?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource);
    }
    let mut path_seen = false;
    for field in &info.field_infos {
        if field.path.as_slice() == path {
            if path_seen || field.object_references.as_slice() != [identifier] {
                return Err(Error::InvalidSource);
            }
            path_seen = true;
        } else if field.object_references.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
    }
    // Native packages may omit optional field-local ownership metadata. The
    // aggregate index above remains mandatory and exact in that case.
    Ok(())
}

fn strict_show_slide_references(source: &[u8], limits: WireLimits) -> Result<Vec<u64>, Error> {
    let tree = strict_nested_payload(source, SHOW_SLIDE_TREE_FIELD, limits)?;
    let view = WireView::parse_with_limits(tree, limits).map_err(map_wire_error)?;
    let count = view
        .fields()
        .filter(|field| field.number() == SLIDE_TREE_SLIDES_FIELD)
        .count();
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(count)
        .map_err(|_allocation| Error::Allocation { amount: count })?;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.number() == SLIDE_TREE_SLIDES_FIELD {
            if field.wire_type() != 2 {
                return Err(Error::InvalidSource);
            }
            identifiers.push(strict_reference(field.payload(), limits)?);
        }
    }
    Ok(identifiers)
}

fn strict_node_slide_reference(
    source: &[u8],
    limits: WireLimits,
    aggregate_work: &mut usize,
) -> Result<u64, Error> {
    let payload = strict_nested_payload(source, NODE_SLIDE_FIELD, limits)?;
    let added = source
        .len()
        .checked_add(payload.len())
        .ok_or(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: limits.max_rewrite_work() as u64,
        })?;
    let observed = aggregate_work
        .checked_add(added)
        .ok_or(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::MAX,
            maximum: limits.max_rewrite_work() as u64,
        })?;
    if observed > limits.max_rewrite_work() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: limits.max_rewrite_work() as u64,
        });
    }
    *aggregate_work = observed;
    strict_reference(payload, limits)
}

fn strict_nested_payload(
    source: &[u8],
    field_number: u32,
    limits: WireLimits,
) -> Result<&[u8], Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut matches = view.fields().filter(|field| field.number() == field_number);
    let field = matches.next().ok_or(Error::InvalidSource)?;
    if matches.next().is_some() || field.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    field.canonical_payload().map_err(map_wire_error)
}

fn strict_reference(source: &[u8], limits: WireLimits) -> Result<u64, Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    let mut deprecated_type_seen = false;
    let mut external = None;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if identifier.is_some() {
                    return Err(Error::InvalidSource);
                }
                identifier = Some(canonical_varint(field)?);
            },
            2 => {
                if std::mem::replace(&mut deprecated_type_seen, true) {
                    return Err(Error::InvalidSource);
                }
                let deprecated_type = canonical_varint(field)?;
                if deprecated_type > i32::MAX as u64 && deprecated_type < 0xffff_ffff_8000_0000 {
                    return Err(Error::InvalidSource);
                }
            },
            3 => {
                if external.is_some() {
                    return Err(Error::InvalidSource);
                }
                external = Some(match canonical_varint(field)? {
                    0 => false,
                    1 => true,
                    _ => return Err(Error::InvalidSource),
                });
            },
            _ if matches!(field.wire_type(), 3 | 4) => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    let identifier = identifier
        .filter(|identifier| *identifier != 0)
        .ok_or(Error::InvalidSource)?;
    if external == Some(true) {
        return Err(Error::InvalidSource);
    }
    Ok(identifier)
}

fn canonical_varint(field: WireFieldView<'_>) -> Result<u64, Error> {
    if field.wire_type() != 0 {
        return Err(Error::InvalidSource);
    }
    let (value, consumed) =
        decode_varint_from_bytes(field.payload()).map_err(|_error| Error::InvalidSource)?;
    if consumed != field.payload().len() || consumed != canonical_varint_len(value) {
        return Err(Error::InvalidSource);
    }
    Ok(value)
}

const fn canonical_varint_len(mut value: u64) -> usize {
    let mut length = 1usize;
    while value >= 0x80 {
        value >>= 7;
        length += 1;
    }
    length
}

fn reject_top_level_groups(source: &[u8], limits: WireLimits) -> Result<(), Error> {
    let view = WireView::parse_with_limits(source, limits).map_err(map_wire_error)?;
    if view
        .fields()
        .any(|field| matches!(field.wire_type(), 3 | 4))
    {
        Err(Error::InvalidSource)
    } else {
        Ok(())
    }
}

fn validate_component_framing(package: &Package, component_name: &str) -> Result<(), Error> {
    let catalog = physical_catalog(package)?;
    let component = catalog
        .components()
        .get(component_name)
        .ok_or(Error::InvalidSource)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::InvalidSource);
    }
    let physical_limits = package.state.options.archive();
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    component
        .archive()
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)
}

fn transition_projection(
    slide: &[u8],
    limits: WireLimits,
) -> Result<litchi_iwa_protos::keynote_slide_transition_codec::SlideTransitionSnapshot<'_>, Error> {
    preflight_selected_transition_payload(slide, limits)?;
    decode_slide_transition(slide, transition_decode_options(slide, limits)?)
        .map_err(map_transition_codec_error)
}

fn transition_decode_options(
    source: &[u8],
    limits: WireLimits,
) -> Result<TransitionDecodeOptions, Error> {
    let recursion_limit =
        u32::try_from(limits.max_nesting()).map_err(|_error| Error::InvalidSource)?;
    Ok(TransitionDecodeOptions::new(source.len(), recursion_limit)
        .with_resource_limits(limits.max_fields(), limits.max_rewrite_work()))
}

fn preflight_selected_transition_payload(source: &[u8], limits: WireLimits) -> Result<(), Error> {
    preflight_wire_tree_with_limits(source, limits, |visit| {
        if matches!(visit.field().wire_type(), 3 | 4) {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "group-bearing transition payload".to_owned(),
            ));
        }
        let descend = match (visit.path(), visit.field().number()) {
            ([], TRANSITION_FIELD)
            | ([TRANSITION_FIELD], TRANSITION_ATTRIBUTES_FIELD)
            | ([TRANSITION_FIELD, TRANSITION_ATTRIBUTES_FIELD], 8) => WireDescent::Descend,
            _ => WireDescent::Skip,
        };
        Ok(descend)
    })
    .map(|_preflight| ())
    .map_err(map_wire_error)
}

fn map_document_codec_error(error: keynote_document_codec::DecodeError) -> Error {
    if let Some(limit) = error.wire_resource_limit() {
        let (kind, observed, maximum) = match limit {
            keynote_document_codec::WireResourceLimit::Bytes { observed, maximum } => {
                (LimitKind::WireBytes, observed as u64, maximum as u64)
            },
            keynote_document_codec::WireResourceLimit::Nesting { observed, maximum } => (
                LimitKind::WireNesting,
                u64::from(observed),
                u64::from(maximum),
            ),
            _ => return Error::InvalidSource,
        };
        return Error::LimitExceeded {
            kind,
            observed,
            maximum,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

fn map_show_codec_error(error: keynote_show_codec::DecodeError) -> Error {
    if let Some(limit) = error.wire_resource_limit() {
        let (kind, observed, maximum) = match limit {
            keynote_show_codec::WireResourceLimit::Bytes { observed, maximum } => {
                (LimitKind::WireBytes, observed as u64, maximum as u64)
            },
            keynote_show_codec::WireResourceLimit::Nesting { observed, maximum } => (
                LimitKind::WireNesting,
                u64::from(observed),
                u64::from(maximum),
            ),
            _ => return Error::InvalidSource,
        };
        return Error::LimitExceeded {
            kind,
            observed,
            maximum,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.slide_reference_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::Slides,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

fn map_transition_codec_error(
    error: litchi_iwa_protos::keynote_slide_transition_codec::DecodeError,
) -> Error {
    if let Some(limit) = error.wire_resource_limit() {
        let (kind, observed, maximum) = match limit {
            litchi_iwa_protos::keynote_slide_transition_codec::WireResourceLimit::Bytes {
                observed,
                maximum,
            } => (LimitKind::WireBytes, observed as u64, maximum as u64),
            litchi_iwa_protos::keynote_slide_transition_codec::WireResourceLimit::Nesting {
                observed,
                maximum,
            } => (
                LimitKind::WireNesting,
                u64::from(observed),
                u64::from(maximum),
            ),
            _ => return Error::InvalidSource,
        };
        return Error::LimitExceeded {
            kind,
            observed,
            maximum,
        };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    Error::InvalidSource
}

impl Package {
    /// Read one slide's focused transition settings without exposing IWA IDs.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, resource, or allocation error when the
    /// rooted selected transition cannot be projected safely.
    ///
    /// # Costs
    ///
    /// Resolves only the rooted show chain and selected node/slide envelopes.
    /// The ownership audit performs indexed O(slides log objects) lookups and
    /// charges aggregate node/reference bytes against [`LimitKind::WireWork`].
    pub fn slide_transition<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Option<Settings>, Error> {
        let position = self.resolve_transition_selector(selector.into())?;
        Ok(select_transition(self, position, false)?.settings)
    }
    /// Start an immutable, selector-first transition edit.
    ///
    /// # Errors
    ///
    /// Returns a typed selector, source, resource, or allocation error when the
    /// rooted selected transition cannot be retained safely.
    ///
    /// # Costs
    ///
    /// Retains one compact settings allocation and borrows the exact package;
    /// opaque values are not cloned a second time. Rooted ownership uses the
    /// same indexed, aggregate-work-bounded audit as [`Self::slide_transition`].
    pub fn edit_slide_transition<'selector>(
        &self,
        selector: impl Into<SlideSelector<'selector>>,
    ) -> Result<Edit<'_>, Error> {
        let position = self.resolve_transition_selector(selector.into())?;
        let selection = select_transition(self, position, false)?;
        Ok(Edit::new(self, selection))
    }
    /// Apply an exact-source-checked reversible transition patch.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for a nonmatching source, or a typed
    /// source, resource, allocation, or verification error for the target.
    ///
    /// # Costs
    ///
    /// Exact-source pointer identity is O(1). A changed patch physically opens
    /// its retained target once and verifies only the rooted selected closure.
    pub fn apply_slide_transition(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            let selection = select_transition(self, patch.position, false)?;
            if selection.node_identifier != patch.node_identifier
                || selection.slide_identifier != patch.slide_identifier
                || selection.settings.as_ref() != patch.before.as_deref()
            {
                return Err(Error::PatchConflict);
            }
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(Error::PatchConflict);
        }
        let source_selection = select_transition(self, patch.position, true)?;
        if source_selection.node_identifier != patch.node_identifier
            || source_selection.slide_identifier != patch.slide_identifier
            || source_selection.settings.as_ref() != patch.before.as_deref()
        {
            return Err(Error::PatchConflict);
        }
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        let touched = verify_candidate_artifacts(
            self,
            &candidate,
            &source_selection,
            patch.after.as_deref(),
        )?;
        if touched != patch.touched_components {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(touched),
        })
    }

    fn resolve_transition_selector(&self, selector: SlideSelector<'_>) -> Result<Position, Error> {
        match selector {
            SlideSelector::Position(position) => Ok(position),
            SlideSelector::Name(_) => self
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(|_| Error::AmbiguousSelector)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(Error::SlideNameNotFound),
        }
    }
}

fn rewrite_transition(
    source: &Package,
    selection: &Selection<'_>,
    after: Option<&Settings>,
) -> Result<(Package, usize), Error> {
    let catalog = physical_catalog(source)?;
    let mut compressed_components: Vec<(&str, Vec<u8>)> = Vec::new();
    let mut component_names: Vec<&str> = Vec::new();
    for component_name in [
        selection.slide_component_name,
        selection.node_component_name,
    ] {
        if component_names.contains(&component_name) {
            continue;
        }
        component_names.push(component_name);
    }
    // Each component is decompressed exactly once, and both objects are patched
    // before it is reassembled. This handles compact archives where node and
    // slide are co-located without reporting duplicate physical work.
    for name in &component_names {
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == *name)
            .ok_or(Error::InvalidSource)?;
        if entry.is_opaque() {
            return Err(Error::InvalidSource);
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
        if archive.object(selection.slide_identifier).is_some() {
            let object = archive
                .object(selection.slide_identifier)
                .ok_or(Error::InvalidSource)?;
            let (index, message) = one_message(object.messages.as_slice(), SLIDE_MESSAGE_TYPE)?;
            if message.data.as_slice() != selection.slide_payload {
                return Err(Error::InvalidSource);
            }
            let changed = rewrite_transition_envelope(
                &message.data,
                &selection.transition_snapshot,
                after.ok_or(Error::UnsupportedTransition)?,
                source.wire_limits().map_err(map_wire_error)?,
            )?;
            archive
                .object_mut(selection.slide_identifier)
                .ok_or(Error::InvalidSource)?
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
        if archive.object(selection.node_identifier).is_some() {
            let object = archive
                .object(selection.node_identifier)
                .ok_or(Error::InvalidSource)?;
            let (index, message) =
                one_message(object.messages.as_slice(), SLIDE_NODE_MESSAGE_TYPE)?;
            let desired = after.is_some_and(Settings::has_effect);
            if message.data.as_slice() != selection.node_payload {
                return Err(Error::InvalidSource);
            }
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
                    .object_mut(selection.node_identifier)
                    .ok_or(Error::InvalidSource)?
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
        if archive.object(selection.slide_identifier).is_none() && !node_changed {
            continue;
        }
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        drop(archive);
        drop(stream);
        let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
        drop(bytes);
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
    drop(edits);
    drop(compressed_components);
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    let touched = verify_candidate_artifacts(source, &candidate, selection, after)?;
    Ok((candidate, touched))
}

fn verify_candidate_artifacts(
    source: &Package,
    candidate: &Package,
    selection: &Selection<'_>,
    expected: Option<&Settings>,
) -> Result<usize, Error> {
    let candidate_selection = select_transition(candidate, selection.position, true)?;
    if candidate_selection.show_identifier != selection.show_identifier
        || candidate_selection.node_identifier != selection.node_identifier
        || candidate_selection.slide_identifier != selection.slide_identifier
        || candidate_selection.settings.as_ref() != expected
    {
        return Err(Error::Verification);
    }
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    let selected_names = [
        selection.slide_component_name,
        selection.node_component_name,
    ];
    let mut source_entries = source_catalog.package().iter();
    let mut candidate_entries = candidate_catalog.package().iter();
    let mut touched = 0usize;
    loop {
        match (source_entries.next(), candidate_entries.next()) {
            (Some(before), Some(after)) if before.name() == after.name() => {
                if selected_names.contains(&before.name()) {
                    if !selected_package_member_preserved(before, after) {
                        return Err(Error::Verification);
                    }
                    if before.raw_record().compressed_data() != after.raw_record().compressed_data()
                    {
                        touched = touched.checked_add(1).ok_or(Error::Verification)?;
                    }
                } else if !package_member_preserved(before, after) {
                    return Err(Error::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(Error::Verification),
        }
    }
    for (index, component_name) in selected_names.into_iter().enumerate() {
        if index == 1 && component_name == selected_names[0] {
            continue;
        }
        verify_selected_component(source, candidate, selection, component_name, expected)?;
    }
    if touched == 0 {
        return Err(Error::Verification);
    }
    Ok(touched)
}

fn verify_selected_component(
    source: &Package,
    candidate: &Package,
    selection: &Selection<'_>,
    component_name: &str,
    expected: Option<&Settings>,
) -> Result<(), Error> {
    let before = source
        .state
        .source
        .components()
        .get(component_name)
        .ok_or(Error::Verification)?;
    let after = candidate
        .state
        .source
        .components()
        .get(component_name)
        .ok_or(Error::Verification)?;
    if before.archive().objects.len() != after.archive().objects.len() {
        return Err(Error::Verification);
    }
    for (old, new) in before
        .archive()
        .objects
        .iter()
        .zip(&after.archive().objects)
    {
        if old.archive_info.identifier != new.archive_info.identifier {
            return Err(Error::Verification);
        }
        match old.archive_info.identifier {
            Some(identifier) if identifier == selection.slide_identifier => {
                let expected = expected.ok_or(Error::Verification)?;
                let limits = source.wire_limits().map_err(map_wire_error)?;
                verify_selected_object(old, new, SLIDE_MESSAGE_TYPE, |old, new| {
                    let rewritten = rewrite_transition_envelope(
                        old,
                        &selection.transition_snapshot,
                        expected,
                        limits,
                    )?;
                    Ok(rewritten == new)
                })?;
            },
            Some(identifier) if identifier == selection.node_identifier => {
                let desired = expected.is_some_and(Settings::has_effect);
                verify_selected_object(old, new, SLIDE_NODE_MESSAGE_TYPE, |old, new| {
                    Ok(node_payload_preserved_except_marker(old, new, desired))
                })?;
            },
            _ if old.archive_info != new.archive_info || old.messages != new.messages => {
                return Err(Error::Verification);
            },
            _ => {},
        }
    }
    Ok(())
}

fn verify_selected_object(
    source: &ArchiveObject,
    candidate: &ArchiveObject,
    message_type: u32,
    selected_payload_valid: impl FnOnce(&[u8], &[u8]) -> Result<bool, Error>,
) -> Result<(), Error> {
    let (source_index, source_payload) = selected_message(source, message_type)?;
    let (candidate_index, candidate_payload) = selected_message(candidate, message_type)?;
    if source_index != candidate_index
        || source.messages.len() != candidate.messages.len()
        || source.archive_info.identifier != candidate.archive_info.identifier
        || source.archive_info.should_merge != candidate.archive_info.should_merge
        || !selected_payload_valid(source_payload, candidate_payload)?
    {
        return Err(Error::Verification);
    }
    for (index, (old, new)) in source.messages.iter().zip(&candidate.messages).enumerate() {
        if old.type_ != new.type_ || (index != source_index && old != new) {
            return Err(Error::Verification);
        }
    }
    let mut expected_archive_info = source.archive_info.clone();
    let expected_length =
        u32::try_from(candidate_payload.len()).map_err(|_error| Error::Verification)?;
    expected_archive_info
        .message_infos
        .get_mut(source_index)
        .ok_or(Error::Verification)?
        .length = expected_length;
    if expected_archive_info != candidate.archive_info {
        return Err(Error::Verification);
    }
    Ok(())
}

fn node_payload_preserved_except_marker(source: &[u8], candidate: &[u8], desired: bool) -> bool {
    let Ok(source_view) = WireView::parse(source) else {
        return false;
    };
    let Ok(candidate_view) = WireView::parse(candidate) else {
        return false;
    };
    let source_other = source_view
        .fields()
        .filter(|field| field.number() != NODE_HAS_TRANSITION_FIELD)
        .map(WireFieldView::raw);
    let candidate_other = candidate_view
        .fields()
        .filter(|field| field.number() != NODE_HAS_TRANSITION_FIELD)
        .map(WireFieldView::raw);
    source_other.eq(candidate_other) && strict_node_marker_wire(candidate).ok() == Some(desired)
}

fn strict_node_marker_wire(source: &[u8]) -> Result<bool, Error> {
    let view = WireView::parse(source).map_err(map_wire_error)?;
    let mut matches = view
        .fields()
        .filter(|field| field.number() == NODE_HAS_TRANSITION_FIELD);
    let field = matches.next().ok_or(Error::InvalidSource)?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource);
    }
    match canonical_varint(field)? {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(Error::InvalidSource),
    }
}

fn package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.raw_record().local_record() == candidate.raw_record().local_record()
        && super::rendering_invalidation::central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_package_member_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    source.name() == candidate.name()
        && source.raw_name() == candidate.raw_name()
        && source.is_opaque() == candidate.is_opaque()
        && source.metadata().local() == candidate.metadata().local()
        && source.metadata().central() == candidate.metadata().central()
        && selected_local_record_preserved(source, candidate)
        && selected_central_record_preserved(
            source.raw_record().central_directory_record(),
            candidate.raw_record().central_directory_record(),
        )
}

fn selected_local_record_preserved(
    source: &litchi_iwa_archive::package::Entry,
    candidate: &litchi_iwa_archive::package::Entry,
) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 14..26;
    let source_record = source.raw_record().local_record();
    let candidate_record = candidate.raw_record().local_record();
    let Some(source_header_length) = zip_local_header_length(source_record) else {
        return false;
    };
    let Some(candidate_header_length) = zip_local_header_length(candidate_record) else {
        return false;
    };
    if source_header_length != candidate_header_length
        || source_record[..CRC_AND_SIZES.start] != candidate_record[..CRC_AND_SIZES.start]
        || source_record[CRC_AND_SIZES.end..source_header_length]
            != candidate_record[CRC_AND_SIZES.end..candidate_header_length]
    {
        return false;
    }
    let Some(source_payload_end) = source_header_length
        .checked_add(source.raw_record().compressed_data().len())
        .filter(|end| *end <= source_record.len())
    else {
        return false;
    };
    let Some(candidate_payload_end) = candidate_header_length
        .checked_add(candidate.raw_record().compressed_data().len())
        .filter(|end| *end <= candidate_record.len())
    else {
        return false;
    };
    selected_local_suffix_preserved(
        source.metadata().local().flags(),
        &source_record[source_payload_end..],
        &candidate_record[candidate_payload_end..],
    )
}

fn zip_local_header_length(record: &[u8]) -> Option<usize> {
    if record.get(..4)? != b"PK\x03\x04" {
        return None;
    }
    let name_length = usize::from(u16::from_le_bytes(record.get(26..28)?.try_into().ok()?));
    let extra_length = usize::from(u16::from_le_bytes(record.get(28..30)?.try_into().ok()?));
    30usize.checked_add(name_length)?.checked_add(extra_length)
}

fn selected_local_suffix_preserved(flags: u16, source: &[u8], candidate: &[u8]) -> bool {
    if flags & 0x0008 == 0 {
        return source == candidate;
    }
    let source_descriptor = usize::from(source.starts_with(b"PK\x07\x08")) * 4;
    let candidate_descriptor = usize::from(candidate.starts_with(b"PK\x07\x08")) * 4;
    source_descriptor == candidate_descriptor
        && source.len() == candidate.len()
        && source.len() >= source_descriptor + 12
        && source[..source_descriptor] == candidate[..candidate_descriptor]
        && source[source_descriptor + 12..] == candidate[candidate_descriptor + 12..]
}

fn selected_central_record_preserved(source: &[u8], candidate: &[u8]) -> bool {
    const CRC_AND_SIZES: std::ops::Range<usize> = 16..28;
    const LOCAL_HEADER_OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= LOCAL_HEADER_OFFSET.end
        && source[..CRC_AND_SIZES.start] == candidate[..CRC_AND_SIZES.start]
        && source[CRC_AND_SIZES.end..LOCAL_HEADER_OFFSET.start]
            == candidate[CRC_AND_SIZES.end..LOCAL_HEADER_OFFSET.start]
        && source[LOCAL_HEADER_OFFSET.end..] == candidate[LOCAL_HEADER_OFFSET.end..]
}

fn one_message(messages: &[RawMessage], kind: u32) -> Result<(usize, &RawMessage), Error> {
    let mut matches = messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == kind);
    let item = matches.next().ok_or(Error::InvalidSource)?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource);
    }
    Ok(item)
}

fn settings_from_snapshot(
    animation: litchi_iwa_protos::keynote_slide_transition_codec::AnimationSnapshot<'_>,
) -> Result<Settings, Error> {
    let mut settings = Settings::new();
    settings
        .set_animation_type(animation.animation_type)
        .map_err(|_| Error::InvalidSource)?;
    let effect = animation
        .effect
        .map(Effect::from_identifier)
        .transpose()
        .map_err(|_| Error::InvalidSource)?;
    settings
        .set_effect(effect)
        .map_err(|_| Error::InvalidSource)?;
    settings
        .set_duration(animation.duration)
        .map_err(|_| Error::InvalidSource)?;
    settings.set_direction(animation.direction.map(Direction::from_native));
    settings
        .set_delay(animation.delay)
        .map_err(|_| Error::InvalidSource)?;
    settings.set_is_automatic(animation.is_automatic);
    let mut animation_parameters = AnimationParameters::new();
    // The focused settings value treats opaque submessages as exact payloads
    // when supplied through the public API. Generated projections are only
    // used for scalar semantic validation here.
    animation_parameters.set_random_number_seed(animation.random_number_seed);
    animation_parameters
        .set_detail(animation.custom_detail)
        .map_err(|_| Error::InvalidSource)?;
    for (slot, value) in TimingCurveSlot::ALL.into_iter().zip([
        animation.custom_effect_timing_curve_theme_name_1,
        animation.custom_effect_timing_curve_theme_name_2,
        animation.custom_effect_timing_curve_theme_name_3,
    ]) {
        animation_parameters
            .set_timing_curve_theme_name(slot, value)
            .map_err(|_| Error::InvalidSource)?;
    }
    animation_parameters.set_writing_direction_is_rtl(animation.writing_direction_is_rtl);
    settings
        .set_animation_parameters(animation_parameters)
        .map_err(|_| Error::InvalidSource)?;
    Ok(settings)
}

fn settings_from_projection(
    snapshot: &TransitionSettingsSnapshot<'_>,
    limits: WireLimits,
) -> Result<Settings, Error> {
    let animation = snapshot.animation.ok_or(Error::UnsupportedTransition)?;
    let mut settings = settings_from_snapshot(animation)?;
    let mut parameters = settings.animation_parameters().clone();
    if let Some(payload) = animation.color {
        preflight_opaque_color(payload, limits)?;
        let color = tsp::Color::decode(payload).map_err(|_error| Error::InvalidSource)?;
        if [
            color.r, color.g, color.b, color.a, color.c, color.m, color.y, color.k, color.w,
        ]
        .into_iter()
        .flatten()
        .any(|component| !component.is_finite())
        {
            return Err(Error::InvalidSource);
        }
        parameters
            .set_color_payload(Some(payload))
            .map_err(|_error| Error::InvalidSource)?;
    }
    for (slot, payload) in TimingCurveSlot::ALL.into_iter().zip([
        animation.custom_effect_timing_curve_1,
        animation.custom_effect_timing_curve_2,
        animation.custom_effect_timing_curve_3,
    ]) {
        if let Some(payload) = payload {
            preflight_opaque_path(payload, limits)?;
            tsd::PathSourceArchive::decode(payload).map_err(|_error| Error::InvalidSource)?;
            parameters
                .set_timing_curve_payload(slot, Some(payload))
                .map_err(|_error| Error::InvalidSource)?;
        }
    }
    settings
        .set_animation_parameters(parameters)
        .map_err(|_error| Error::InvalidSource)?;
    let mut custom = CustomParameters::new();
    custom
        .set_twist(snapshot.custom_twist)
        .map_err(|_| Error::InvalidSource)?;
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
        .map_err(|_| Error::InvalidSource)?;
    settings
        .set_custom_parameters(custom)
        .map_err(|_| Error::InvalidSource)?;
    Ok(settings)
}

fn validate_requested_opaque_settings(
    settings: &Settings,
    limits: WireLimits,
) -> Result<(), Error> {
    let parameters = settings.animation_parameters();
    if let Some(payload) = parameters.color_payload() {
        preflight_opaque_color(payload, limits).map_err(map_requested_opaque_error)?;
        let color = tsp::Color::decode(payload).map_err(|_error| Error::InvalidOpaquePayload)?;
        if [
            color.r, color.g, color.b, color.a, color.c, color.m, color.y, color.k, color.w,
        ]
        .into_iter()
        .flatten()
        .any(|component| !component.is_finite())
        {
            return Err(Error::InvalidOpaquePayload);
        }
    }
    for payload in parameters.timing_curve_payloads().into_iter().flatten() {
        preflight_opaque_path(payload, limits).map_err(map_requested_opaque_error)?;
        tsd::PathSourceArchive::decode(payload).map_err(|_error| Error::InvalidOpaquePayload)?;
    }
    Ok(())
}

fn map_requested_opaque_error(error: Error) -> Error {
    match error {
        Error::LimitExceeded { .. } | Error::Allocation { .. } => error,
        _ => Error::InvalidOpaquePayload,
    }
}

fn preflight_opaque_color(source: &[u8], limits: WireLimits) -> Result<(), Error> {
    preflight_wire_tree_with_limits(source, limits, |visit| {
        if matches!(visit.field().wire_type(), 3 | 4) {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "group-bearing transition color".to_owned(),
            ));
        }
        Ok(WireDescent::Skip)
    })
    .map(|_preflight| ())
    .map_err(map_wire_error)
}

#[allow(
    clippy::unnested_or_patterns,
    reason = "separate path-schema comments are clearer than a mechanically nested pattern"
)]
fn preflight_opaque_path(source: &[u8], limits: WireLimits) -> Result<(), Error> {
    preflight_wire_tree_with_limits(source, limits, |visit| {
        if matches!(visit.field().wire_type(), 3 | 4) {
            return Err(litchi_iwa_common::Error::InvalidFormat(
                "group-bearing transition path".to_owned(),
            ));
        }
        let path = visit.path();
        let field = visit.field().number();
        let descend = match (path, field) {
            // PathSourceArchive variants.
            ([], 3..=8)
            // PointPathSourceArchive / ScalarPathSourceArchive.
            | ([3], 2 | 3)
            | ([4], 3)
            // BezierPathSourceArchive -> Size / TSP.Path.
            | ([5], 2 | 3)
            | ([5, 3], 1)
            | ([5, 3, 1], 2)
            // CalloutPathSourceArchive.
            | ([6], 1 | 2)
            // ConnectionLinePathSourceArchive -> Bezier -> Size / Path.
            | ([7], 1)
            | ([7, 1], 2 | 3)
            | ([7, 1, 3], 1)
            | ([7, 1, 3, 1], 2)
            // EditableBezierPathSourceArchive -> Subpath -> Node -> Points.
            | ([8], 1 | 2)
            | ([8, 1], 1)
            | ([8, 1, 1], 1..=3) => WireDescent::Descend,
            _ => WireDescent::Skip,
        };
        Ok(descend)
    })
    .map(|_preflight| ())
    .map_err(map_wire_error)
}

fn strict_node_transition_flag(source: &[u8], limits: WireLimits) -> Result<bool, Error> {
    decode_slide_node_has_transition(source, transition_decode_options(source, limits)?)
        .map_err(map_transition_codec_error)
}

fn rewrite_transition_envelope(
    source: &[u8],
    snapshot: &TransitionSettingsSnapshot<'_>,
    settings: &Settings,
    limits: WireLimits,
) -> Result<Vec<u8>, Error> {
    let animation = snapshot.animation.ok_or(Error::UnsupportedTransition)?;
    let parameters = settings.animation_parameters();
    let custom = settings.custom_parameters();
    let edits = [
        NestedFieldEdit::new(
            &ANIMATION_PATHS[0],
            animation.animation_type.is_some(),
            NestedFieldReplacement::LengthDelimited(settings.animation_type().map(str::as_bytes)),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[1],
            animation.effect.is_some(),
            NestedFieldReplacement::LengthDelimited(
                settings.effect().map(Effect::identifier).map(str::as_bytes),
            ),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[2],
            animation.duration.is_some(),
            NestedFieldReplacement::Fixed64(settings.duration().map(f64::to_bits)),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[3],
            animation.direction.is_some(),
            NestedFieldReplacement::Varint(
                settings
                    .direction()
                    .map(|value| u64::from(value.native_value())),
            ),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[4],
            animation.delay.is_some(),
            NestedFieldReplacement::Fixed64(settings.delay().map(f64::to_bits)),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[5],
            animation.is_automatic.is_some(),
            NestedFieldReplacement::Varint(settings.is_automatic().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[6],
            animation.color.is_some(),
            NestedFieldReplacement::LengthDelimited(parameters.color_payload()),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[7],
            animation.custom_effect_timing_curve_1.is_some(),
            NestedFieldReplacement::LengthDelimited(
                parameters.timing_curve_payload(TimingCurveSlot::First),
            ),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[8],
            animation.custom_effect_timing_curve_2.is_some(),
            NestedFieldReplacement::LengthDelimited(
                parameters.timing_curve_payload(TimingCurveSlot::Second),
            ),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[9],
            animation.custom_effect_timing_curve_3.is_some(),
            NestedFieldReplacement::LengthDelimited(
                parameters.timing_curve_payload(TimingCurveSlot::Third),
            ),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[10],
            animation.random_number_seed.is_some(),
            NestedFieldReplacement::Varint(parameters.random_number_seed().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[11],
            animation.custom_detail.is_some(),
            NestedFieldReplacement::Fixed64(parameters.detail().map(f64::to_bits)),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[12],
            animation.custom_effect_timing_curve_theme_name_1.is_some(),
            NestedFieldReplacement::LengthDelimited(
                parameters
                    .timing_curve_theme_name(TimingCurveSlot::First)
                    .map(str::as_bytes),
            ),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[13],
            animation.custom_effect_timing_curve_theme_name_2.is_some(),
            NestedFieldReplacement::LengthDelimited(
                parameters
                    .timing_curve_theme_name(TimingCurveSlot::Second)
                    .map(str::as_bytes),
            ),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[14],
            animation.custom_effect_timing_curve_theme_name_3.is_some(),
            NestedFieldReplacement::LengthDelimited(
                parameters
                    .timing_curve_theme_name(TimingCurveSlot::Third)
                    .map(str::as_bytes),
            ),
        ),
        NestedFieldEdit::new(
            &ANIMATION_PATHS[15],
            animation.writing_direction_is_rtl.is_some(),
            NestedFieldReplacement::Varint(parameters.writing_direction_is_rtl().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &CUSTOM_PATHS[0],
            snapshot.custom_twist.is_some(),
            NestedFieldReplacement::Fixed32(custom.twist().map(f32::to_bits)),
        ),
        NestedFieldEdit::new(
            &CUSTOM_PATHS[1],
            snapshot.custom_mosaic_size.is_some(),
            NestedFieldReplacement::Varint(custom.mosaic_size().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &CUSTOM_PATHS[2],
            snapshot.custom_mosaic_type.is_some(),
            NestedFieldReplacement::Varint(
                custom
                    .mosaic_type()
                    .map(|value| u64::from(value.native_value())),
            ),
        ),
        NestedFieldEdit::new(
            &CUSTOM_PATHS[3],
            snapshot.custom_bounce.is_some(),
            NestedFieldReplacement::Varint(custom.bounce().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &CUSTOM_PATHS[4],
            snapshot.custom_magic_move_fade_unmatched_objects.is_some(),
            NestedFieldReplacement::Varint(
                custom.magic_move_fade_unmatched_objects().map(u64::from),
            ),
        ),
        NestedFieldEdit::new(
            &CUSTOM_PATHS[5],
            snapshot.custom_timing_curve.is_some(),
            NestedFieldReplacement::Varint(
                custom
                    .acceleration()
                    .map(|value| i64::from(value.native_value()) as u64),
            ),
        ),
        NestedFieldEdit::new(
            &CUSTOM_PATHS[6],
            snapshot.custom_text_delivery_type.is_some(),
            NestedFieldReplacement::Varint(
                custom
                    .text_delivery()
                    .map(|value| i64::from(value.native_value()) as u64),
            ),
        ),
        NestedFieldEdit::new(
            &CUSTOM_PATHS[7],
            snapshot.custom_motion_blur.is_some(),
            NestedFieldReplacement::Varint(custom.motion_blur().map(u64::from)),
        ),
        NestedFieldEdit::new(
            &CUSTOM_PATHS[8],
            snapshot.custom_travel_distance.is_some(),
            NestedFieldReplacement::Fixed32(custom.travel_distance().map(f32::to_bits)),
        ),
    ];
    let data =
        patch_nested_fields_batched_with_limits(source, &edits, limits).map_err(map_wire_error)?;
    check_output(&data, limits)?;
    Ok(data)
}

fn rewrite_node_transition_flag(
    source: &[u8],
    value: bool,
    limits: WireLimits,
) -> Result<Vec<u8>, Error> {
    let path = [NODE_HAS_TRANSITION_FIELD];
    let edit = NestedFieldEdit::new(
        &path,
        true,
        NestedFieldReplacement::Varint(Some(u64::from(value))),
    );
    let output =
        patch_nested_fields_batched_with_limits(source, &[edit], limits).map_err(map_wire_error)?;
    check_output(&output, limits)?;
    Ok(output)
}
fn check_output(output: &[u8], limits: WireLimits) -> Result<(), Error> {
    if output.len() > limits.max_output_bytes() {
        Err(Error::LimitExceeded {
            kind: LimitKind::OutputBytes,
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
fn physical_catalog(package: &Package) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(Error::UnsupportedSource),
    }
}
fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => LimitKind::Entries,
                super::SemanticLimitKind::Slides => LimitKind::Slides,
                super::SemanticLimitKind::References => LimitKind::References,
                super::SemanticLimitKind::TextStorages => LimitKind::TextStorages,
                super::SemanticLimitKind::TextFragments => LimitKind::TextFragments,
                super::SemanticLimitKind::TextBytes => LimitKind::TextBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => LimitKind::WireBytes,
                super::PayloadLimitKind::Fields => LimitKind::WireFields,
                super::PayloadLimitKind::Nesting => LimitKind::WireNesting,
                super::PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        _ => Error::InvalidSource,
    }
}
fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
                _ => LimitKind::EntryBytes,
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
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::MessageBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                _ => LimitKind::EntryBytes,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}
fn map_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}
