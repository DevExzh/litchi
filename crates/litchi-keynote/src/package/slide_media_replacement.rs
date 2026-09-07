//! Exact-source, selector-first replacement of existing Keynote slide media.
//!
//! The public operation is deliberately small: a caller selects a slide and a
//! source-order movie/audio drawable, chooses its content or poster, and
//! replaces the materialized bytes.  The movie graph and all data references
//! stay untouched.  The package adapter owns the physical media closure,
//! `PackageMetadata` digest/length rewrite, and ZIP reassembly.

#![allow(
    clippy::cast_possible_truncation,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::too_many_lines,
    clippy::wildcard_enum_match_arm,
    reason = "The adapter keeps strict physical admission, rewrite, and readback adjacent."
)]

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{WireLimits, media::Type as MediaType, wire::WireView};
use litchi_iwa_core::{
    Archive, ArchiveLimits, ArchiveObject, RawMessage, SnappyLimits, SnappyStream,
};
use litchi_iwa_protos::{keynote_media_codec, package_metadata_media_codec as metadata_codec};
use sha1::{Digest, Sha1};
use thiserror::Error;

use super::{MOVIE_MESSAGE_TYPE, Package, PhysicalSource, SLIDE_MESSAGE_TYPE, SemanticPath};
use crate::{MovieKind, MovieSelector, SlideSelector};

mod budget;
mod closure;
pub(in crate::package) use budget::{MediaBudget, MediaBudgetUsage};

const MOVIE_DATA_FIELD: u32 = 14;
const POSTER_IMAGE_DATA_FIELD: u32 = 15;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const MOVIE_SUPER_FIELD: u32 = 1;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const DATA_PREFIX: &str = "Data/";
const SHA1_BYTES: usize = 20;
const MAX_REPLACEMENT_BYTES: usize = 1 << 30;

fn map_keynote_decode_error(error: keynote_media_codec::DecodeError) -> SlideMediaDataError {
    error
        .resource_limit()
        .map_or(SlideMediaDataError::InvalidSource, map_keynote_limit)
}

fn map_keynote_limit(limit: keynote_media_codec::DecodeLimit) -> SlideMediaDataError {
    match limit {
        keynote_media_codec::DecodeLimit::Bytes { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::InputBytes, observed, maximum)
        },
        keynote_media_codec::DecodeLimit::Fields { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::WireFields, observed, maximum)
        },
        keynote_media_codec::DecodeLimit::Work { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::WireWork, observed, maximum)
        },
        keynote_media_codec::DecodeLimit::Nesting { observed, maximum } => {
            limit_exceeded_u32(SlideMediaDataLimitKind::WireNesting, observed, maximum)
        },
        _ => SlideMediaDataError::InvalidSource,
    }
}

fn map_metadata_decode_error(error: metadata_codec::DecodeError) -> SlideMediaDataError {
    error
        .resource_limit()
        .map_or(SlideMediaDataError::InvalidSource, map_metadata_limit)
}

fn map_metadata_rewrite_error(error: metadata_codec::RewriteError) -> SlideMediaDataError {
    if let Some(limit) = error.resource_limit() {
        return map_metadata_limit(limit);
    }
    error
        .allocation_request()
        .map_or(SlideMediaDataError::InvalidSource, |amount| {
            SlideMediaDataError::Allocation { amount }
        })
}

fn map_metadata_limit(limit: metadata_codec::DecodeLimit) -> SlideMediaDataError {
    match limit {
        metadata_codec::DecodeLimit::Bytes { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::InputBytes, observed, maximum)
        },
        metadata_codec::DecodeLimit::Fields { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::WireFields, observed, maximum)
        },
        metadata_codec::DecodeLimit::Work { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::WireWork, observed, maximum)
        },
        metadata_codec::DecodeLimit::OutputBytes { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::OutputBytes, observed, maximum)
        },
        metadata_codec::DecodeLimit::Components { observed, maximum }
        | metadata_codec::DecodeLimit::DataRecords { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::Entries, observed, maximum)
        },
        metadata_codec::DecodeLimit::Owners { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::References, observed, maximum)
        },
        metadata_codec::DecodeLimit::DigestBytes { observed, maximum }
        | metadata_codec::DecodeLimit::NameBytes { observed, maximum } => {
            limit_exceeded(SlideMediaDataLimitKind::EntryBytes, observed, maximum)
        },
        metadata_codec::DecodeLimit::Nesting { observed, maximum } => {
            limit_exceeded_u32(SlideMediaDataLimitKind::WireNesting, observed, maximum)
        },
        _ => SlideMediaDataError::InvalidSource,
    }
}

fn limit_exceeded(
    kind: SlideMediaDataLimitKind,
    observed: usize,
    maximum: usize,
) -> SlideMediaDataError {
    SlideMediaDataError::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    }
}

fn limit_exceeded_u32(
    kind: SlideMediaDataLimitKind,
    observed: u32,
    maximum: u32,
) -> SlideMediaDataError {
    SlideMediaDataError::LimitExceeded {
        kind,
        observed: u64::from(observed),
        maximum: u64::from(maximum),
    }
}

/// Which materialized part of a selected slide media drawable to replace.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum MediaPart {
    /// The movie's video bytes, or the audio bytes for an audio drawable.
    Content,
    /// The movie's poster image bytes.
    Poster,
}

/// A finite resource category governed by one slide-media transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideMediaDataLimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete rewritten package output bytes.
    OutputBytes,
    /// ZIP members or IWA objects/messages retained during the operation.
    Entries,
    /// One selected package or wire member.
    EntryBytes,
    /// Aggregate physical package bytes.
    TotalBytes,
    /// Semantic slide references.
    Slides,
    /// Native graph references.
    References,
    /// Media payload bytes.
    MediaBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate wire and rewrite work.
    WireWork,
    /// Fallible operation-local allocation events.
    Allocations,
}

impl fmt::Display for SlideMediaDataLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::MediaBytes => "media bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
        })
    }
}

/// A content-redacted failure from a focused Keynote media operation.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideMediaDataError {
    /// The package was created from semantic components without exact ZIP bytes.
    #[error("this Keynote source does not support physical slide-media edits")]
    UnsupportedSource,
    /// The slide selector name was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide matched the exact navigator name.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked slide position was outside the presentation.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// A name selector was ambiguous.
    #[error("the Keynote slide-media selector is ambiguous")]
    AmbiguousSelector,
    /// No movie/audio drawable matched the source-order selector.
    #[error("the selected Keynote slide has no media at position {position:?}")]
    MoviePositionNotFound { position: Position },
    /// Poster selection on an audio-only drawable is unsupported.
    #[error("Keynote audio media has no poster part")]
    AudioPoster,
    /// The selected graph or package-media closure is unsafe to edit.
    #[error("the Keynote slide-media source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote slide-media {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// The bounded resource category.
        kind: SlideMediaDataLimitKind,
        /// The checked observed value.
        observed: u64,
        /// The configured maximum.
        maximum: u64,
    },
    /// A bounded operation-local allocation failed.
    #[error("could not allocate {amount} units for the Keynote slide-media transaction")]
    Allocation { amount: usize },
    /// Empty replacement data is never a materialized asset.
    #[error("Keynote media replacement data cannot be empty")]
    EmptyReplacement,
    /// The replacement exceeds the bounded focused-owner ceiling.
    #[error("Keynote media replacement exceeds the configured byte limit")]
    ReplacementTooLarge,
    /// Replacement bytes do not retain the selected media family.
    #[error("Keynote media replacement has an incompatible media type")]
    ReplacementType,
    /// Candidate reopening did not reproduce the requested replacement.
    #[error("the edited Keynote slide-media candidate failed semantic verification")]
    Verification,
    /// The patch was created from a different exact source artifact.
    #[error("the Keynote slide-media patch does not match the exact source package")]
    PatchConflict,
    /// A lower-level read failed without exposing raw archive details.
    #[error("could not read the Keynote package for slide-media access")]
    Read,
}

/// A borrowed existing media payload.
pub type SlideMediaData<'a> = &'a [u8];

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct MediaSelection {
    slide_position: Position,
    movie_position: Position,
    slide_identifier: u64,
    slide_node_identifier: u64,
    movie_identifier: u64,
    component_name: String,
    kind: MovieKind,
    content_identifier: Option<u64>,
    poster_identifier: Option<u64>,
    record: Option<MediaRecord>,
}

impl MediaSelection {
    pub(super) fn same_identity(&self, other: &Self) -> bool {
        self.slide_position == other.slide_position
            && self.movie_position == other.movie_position
            && self.slide_identifier == other.slide_identifier
            && self.slide_node_identifier == other.slide_node_identifier
            && self.movie_identifier == other.movie_identifier
            && self.component_name == other.component_name
            && self.kind == other.kind
            && self.content_identifier == other.content_identifier
            && self.poster_identifier == other.poster_identifier
    }

    /// Report whether the selected MovieArchive carries a poster data edge.
    /// The identifier itself remains private to the physical replacement
    /// owner; sibling semantic adapters only need to distinguish an absent
    /// optional poster from an invalid materialized asset.
    pub(super) const fn has_poster(&self) -> bool {
        self.poster_identifier.is_some()
    }

    const fn identifier(&self, part: MediaPart) -> Option<u64> {
        match part {
            MediaPart::Content => self.content_identifier,
            MediaPart::Poster => self.poster_identifier,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct MediaRecord {
    identifier: u64,
    digest: [u8; SHA1_BYTES],
    preferred_name: Box<str>,
    current_name: Box<str>,
    materialized_length: Option<usize>,
    unknown_fields: bool,
}

#[derive(Debug, Clone, Copy)]
struct OwnerRecord {
    component_identifier: u64,
    data_identifier: u64,
    object_identifier: u64,
    count: u32,
    versioned: bool,
}

// The metadata visitor only retains owned data.  Component snapshots contain
// borrowed slices, so this adapter records just the semantic component facts
// in a separate vector instead of storing the generated view.
#[derive(Debug, Clone)]
struct ComponentRecord {
    identifier: u64,
    locator: Box<str>,
    versioned: bool,
    unknown_fields: bool,
}

#[derive(Debug, Default)]
struct OwnedMetadataFacts {
    records: Vec<MediaRecord>,
    components: Vec<ComponentRecord>,
    references: Vec<(u64, u64, u32, bool)>,
    owners: Vec<OwnerRecord>,
}

struct MetadataVisitor<'budget> {
    facts: OwnedMetadataFacts,
    budget: &'budget mut MediaBudget,
}

impl metadata_codec::PackageMetadataMediaVisitor for MetadataVisitor<'_> {
    fn visit_data_info(
        &mut self,
        data_info: metadata_codec::DataInfoSnapshot<'_>,
    ) -> Result<(), metadata_codec::DecodeError> {
        let digest = <[u8; SHA1_BYTES]>::try_from(data_info.digest())
            .map_err(|_| metadata_codec::DecodeError::invalid_for_adapter())?;
        let materialized_length = data_info
            .materialized_length()
            .and_then(|length| usize::try_from(length).ok());
        let current_name = data_info
            .file_name()
            .filter(|name| !name.is_empty())
            .unwrap_or_else(|| data_info.preferred_file_name());
        if self.facts.records.len() == self.facts.records.capacity() {
            return Err(metadata_codec::DecodeError::invalid_for_adapter());
        }
        let preferred_name = copy_boxed_str(data_info.preferred_file_name(), self.budget)
            .map_err(|_| metadata_codec::DecodeError::invalid_for_adapter())?;
        let current_name = copy_boxed_str(current_name, self.budget)
            .map_err(|_| metadata_codec::DecodeError::invalid_for_adapter())?;
        self.facts.records.push(MediaRecord {
            identifier: data_info.identifier(),
            digest,
            preferred_name,
            current_name,
            materialized_length,
            unknown_fields: data_info.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: metadata_codec::ComponentSnapshot<'_>,
    ) -> Result<(), metadata_codec::DecodeError> {
        if self.facts.components.len() == self.facts.components.capacity() {
            return Err(metadata_codec::DecodeError::invalid_for_adapter());
        }
        let locator = copy_boxed_str(component.effective_locator(), self.budget)
            .map_err(|_| metadata_codec::DecodeError::invalid_for_adapter())?;
        self.facts.components.push(ComponentRecord {
            identifier: component.identifier(),
            locator,
            versioned: component.is_versioned(),
            unknown_fields: component.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        component: metadata_codec::ComponentSnapshot<'_>,
        data_reference: metadata_codec::ComponentDataReferenceSnapshot<'_>,
    ) -> Result<(), metadata_codec::DecodeError> {
        if self.facts.references.len() == self.facts.references.capacity() {
            return Err(metadata_codec::DecodeError::invalid_for_adapter());
        }
        let owner_count = u32::try_from(data_reference.owner_count())
            .map_err(|_| metadata_codec::DecodeError::invalid_for_adapter())?;
        self.facts.references.push((
            component.identifier(),
            data_reference.data_identifier(),
            owner_count,
            component.is_versioned(),
        ));
        Ok(())
    }

    fn visit_owner(
        &mut self,
        component: metadata_codec::ComponentSnapshot<'_>,
        data_reference: metadata_codec::ComponentDataReferenceSnapshot<'_>,
        owner: metadata_codec::OwnerSnapshot<'_>,
    ) -> Result<(), metadata_codec::DecodeError> {
        if self.facts.owners.len() == self.facts.owners.capacity() {
            return Err(metadata_codec::DecodeError::invalid_for_adapter());
        }
        self.facts.owners.push(OwnerRecord {
            component_identifier: component.identifier(),
            data_identifier: data_reference.data_identifier(),
            object_identifier: owner.object_identifier(),
            count: owner.count(),
            versioned: component.is_versioned(),
        });
        Ok(())
    }
}

/// Mutable semantic media replacement staged against one immutable package.
pub struct SlideMediaDataEdit<'a> {
    source: &'a Package,
    selection: MediaSelection,
    part: MediaPart,
    before: &'a [u8],
    after: Option<Vec<u8>>,
}

impl fmt::Debug for SlideMediaDataEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideMediaDataEdit")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field("part", &self.part)
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.as_ref().map(|bytes| bytes.len()))
            .finish_non_exhaustive()
    }
}

impl<'a> SlideMediaDataEdit<'a> {
    fn new<'slide>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
        part: MediaPart,
    ) -> Result<Self, SlideMediaDataError> {
        let mut budget = MediaBudget::for_package(source)?;
        let catalog = physical_catalog(source)?;
        budget.charge_catalog(catalog)?;
        let selection = select_media(
            source,
            slide_selector.into(),
            movie_selector.into(),
            part,
            &mut budget,
        )?;
        let before = read_selected_media(source, &selection, part, &mut budget)?;
        Ok(Self {
            source,
            selection,
            part,
            before,
            after: None,
        })
    }

    /// Return the selected slide's typed source-order position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected media drawable's typed source-order position.
    #[must_use]
    pub const fn movie_position(&self) -> Position {
        self.selection.movie_position
    }

    /// Return which media part this transaction edits.
    #[must_use]
    pub const fn part(&self) -> MediaPart {
        self.part
    }

    /// Borrow the exact source bytes captured during selection.
    #[must_use]
    pub fn before(&self) -> &[u8] {
        self.before
    }

    /// Borrow the staged replacement bytes, if one has been provided.
    #[must_use]
    pub fn after(&self) -> Option<&[u8]> {
        self.after.as_deref()
    }

    /// Stage a non-empty replacement while retaining the selected media type.
    pub fn set(mut self, replacement: &[u8]) -> Result<Self, SlideMediaDataError> {
        validate_replacement(
            replacement,
            self.part,
            self.before,
            self.source.state.options.archive(),
        )?;
        let mut budget = MediaBudget::for_package(self.source)?;
        budget.entry_bytes(replacement.len())?;
        budget.allocation(replacement.len())?;
        let mut owned = Vec::new();
        owned.try_reserve_exact(replacement.len()).map_err(|_| {
            SlideMediaDataError::Allocation {
                amount: replacement.len(),
            }
        })?;
        owned.extend_from_slice(replacement);
        self.after = Some(owned);
        Ok(self)
    }

    /// Publish the replacement after physical and semantic candidate checks.
    pub fn commit(self) -> Result<SlideMediaDataCommit, SlideMediaDataError> {
        commit_edit(self)
    }
}

/// An exact-source reversible slide-media replacement patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideMediaDataPatch {
    artifacts: ExactArtifacts,
    selection: MediaSelection,
    part: MediaPart,
    before_digest: [u8; SHA1_BYTES],
    after_digest: [u8; SHA1_BYTES],
    before_length: usize,
    after_length: usize,
    source_preview_count: usize,
    target_preview_count: usize,
    deleted_previews: usize,
}

impl fmt::Debug for SlideMediaDataPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideMediaDataPatch")
            .field("slide_position", &self.selection.slide_position)
            .field("movie_position", &self.selection.movie_position)
            .field("part", &self.part)
            .field("before_length", &self.before_length)
            .field("after_length", &self.after_length)
            .field("deleted_previews", &self.deleted_previews)
            .finish_non_exhaustive()
    }
}

impl SlideMediaDataPatch {
    /// Return the compact diagnostic fingerprint of the exact source artifact.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the compact diagnostic fingerprint of the committed artifact.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return the selected media part.
    #[must_use]
    pub const fn part(&self) -> MediaPart {
        self.part
    }

    /// Return the source payload length.
    #[must_use]
    pub const fn before_length(&self) -> usize {
        self.before_length
    }

    /// Return the target payload length.
    #[must_use]
    pub const fn after_length(&self) -> usize {
        self.after_length
    }

    /// Return the number of root previews removed by the rewrite.
    #[must_use]
    pub const fn deleted_previews(&self) -> usize {
        self.deleted_previews
    }

    /// Return whether the patch retains the exact source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            part: self.part,
            before_digest: self.after_digest,
            after_digest: self.before_digest,
            before_length: self.after_length,
            after_length: self.before_length,
            source_preview_count: self.target_preview_count,
            target_preview_count: self.source_preview_count,
            deleted_previews: self
                .target_preview_count
                .saturating_sub(self.source_preview_count),
        }
    }
}

/// Result of a committed slide-media replacement.
pub struct SlideMediaDataCommit {
    package: Package,
    patch: SlideMediaDataPatch,
    diagnostics: SlideMediaDataDiagnostics,
}

impl SlideMediaDataCommit {
    /// Borrow the fully reopened candidate package.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its package.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the exact reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideMediaDataPatch {
        &self.patch
    }

    /// Borrow operation diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideMediaDataDiagnostics {
        &self.diagnostics
    }
}

/// Deterministic physical diagnostics for one slide-media commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideMediaDataDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
}

impl SlideMediaDataDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
        }
    }

    const fn published(touched_components: usize, deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            deleted_previews,
        }
    }

    /// Return whether the candidate changed any bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of edited ZIP members.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return the number of removed root rendering previews.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
}

impl Package {
    /// Borrow the materialized content or poster for one source-order slide
    /// movie/audio drawable.
    pub fn slide_media_data<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
        part: MediaPart,
    ) -> Result<SlideMediaData<'_>, SlideMediaDataError> {
        let mut budget = MediaBudget::for_package(self)?;
        let catalog = physical_catalog(self)?;
        budget.charge_catalog(catalog)?;
        let selection = select_media(
            self,
            slide_selector.into(),
            movie_selector.into(),
            part,
            &mut budget,
        )?;
        read_selected_media(self, &selection, part, &mut budget)
    }

    /// Start an exact-source replacement of one materialized slide media part.
    pub fn edit_slide_media_data<'slide>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        movie_selector: impl Into<MovieSelector>,
        part: MediaPart,
    ) -> Result<SlideMediaDataEdit<'_>, SlideMediaDataError> {
        SlideMediaDataEdit::new(self, slide_selector, movie_selector, part)
    }

    /// Apply an exact-source-checked slide media replacement patch.
    pub fn apply_slide_media_data(
        &self,
        patch: &SlideMediaDataPatch,
    ) -> Result<SlideMediaDataCommit, SlideMediaDataError> {
        let mut budget = MediaBudget::for_package(self)?;
        let catalog = physical_catalog(self)?;
        budget.charge_catalog(catalog)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(SlideMediaDataError::PatchConflict);
        }
        let current = select_media(
            self,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            patch.part,
            &mut budget,
        )?;
        if !current.same_identity(&patch.selection) {
            return Err(SlideMediaDataError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(|_| SlideMediaDataError::Read)?;
            return Ok(SlideMediaDataCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideMediaDataDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(SlideMediaDataError::PatchConflict);
        }
        let target = patch.artifacts.target();
        budget.output_bytes(target.len())?;
        budget.media_bytes(target.len())?;
        let candidate = Package::from_source_with_options(target, self.state.options)
            .map_err(|_| SlideMediaDataError::PatchConflict)?;
        candidate
            .validate()
            .map_err(|_| SlideMediaDataError::PatchConflict)?;
        let candidate_catalog = physical_catalog(&candidate)?;
        budget.charge_catalog(candidate_catalog)?;
        let candidate_selection = select_media(
            &candidate,
            SlideSelector::position(patch.selection.slide_position),
            MovieSelector::position(patch.selection.movie_position),
            patch.part,
            &mut budget,
        )?;
        if !candidate_selection.same_identity(&patch.selection) {
            return Err(SlideMediaDataError::Verification);
        }
        let bytes = read_selected_media(&candidate, &candidate_selection, patch.part, &mut budget)?;
        let digest: [u8; SHA1_BYTES] = Sha1::digest(bytes).into();
        if bytes.len() != patch.after_length || digest != patch.after_digest {
            return Err(SlideMediaDataError::Verification);
        }
        let candidate_preview_count =
            super::rendering_invalidation::root_preview_deletions(candidate_catalog.package())
                .map_err(|_| SlideMediaDataError::Verification)?
                .len();
        if candidate_preview_count != patch.target_preview_count {
            return Err(SlideMediaDataError::Verification);
        }
        Ok(SlideMediaDataCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SlideMediaDataDiagnostics::published(2, patch.deleted_previews),
        })
    }
}

fn commit_edit(edit: SlideMediaDataEdit<'_>) -> Result<SlideMediaDataCommit, SlideMediaDataError> {
    let SlideMediaDataEdit {
        source,
        selection,
        part,
        before: expected_before,
        after,
    } = edit;
    let replacement = after
        .as_deref()
        .ok_or(SlideMediaDataError::EmptyReplacement)?;
    let catalog = physical_catalog(source)?;
    let mut budget = MediaBudget::for_package(source)?;
    budget.charge_catalog(catalog)?;
    budget.media_bytes(replacement.len())?;
    let source_bytes = catalog.shared_source();
    // `Package` is immutable for the lifetime of this edit. The constructor
    // already admitted the rooted closure and cached its DataInfo record, so
    // commit only rechecks the exact physical member instead of rescanning
    // PackageMetadata a second time.
    let before = read_selected_media(source, &selection, part, &mut budget)?;
    if before != expected_before {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let before_digest: [u8; SHA1_BYTES] = Sha1::digest(before).into();
    let after_digest: [u8; SHA1_BYTES] = Sha1::digest(replacement).into();
    if before == replacement {
        let source_preview_count =
            super::rendering_invalidation::root_preview_deletions(catalog.package())
                .map_err(|_| SlideMediaDataError::InvalidSource)?
                .len();
        let patch = SlideMediaDataPatch {
            artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
            selection,
            part,
            before_digest,
            after_digest,
            before_length: before.len(),
            after_length: replacement.len(),
            source_preview_count,
            target_preview_count: source_preview_count,
            deleted_previews: 0,
        };
        return Ok(SlideMediaDataCommit {
            package: source.snapshot(),
            patch,
            diagnostics: SlideMediaDataDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(SlideMediaDataError::UnsupportedSource);
    }
    let record = selection
        .record
        .as_ref()
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let data_name = copy_data_path(record.current_name.as_ref(), &mut budget)?;
    let preview_plan = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let (package, deleted_previews) =
        rewrite_media(source, &selection, part, replacement, &mut budget)?;
    let candidate_catalog = physical_catalog(&package)?;
    budget.charge_catalog(candidate_catalog)?;
    let target_bytes = physical_catalog(&package)?.shared_source();
    let source_preview_count =
        super::rendering_invalidation::root_preview_deletions(catalog.package())
            .map_err(|_| SlideMediaDataError::InvalidSource)?
            .len();
    let target_preview_count = super::rendering_invalidation::root_preview_deletions(
        physical_catalog(&package)?.package(),
    )
    .map_err(|_| SlideMediaDataError::Verification)?
    .len();
    let candidate = select_media(
        &package,
        SlideSelector::position(selection.slide_position),
        MovieSelector::position(selection.movie_position),
        part,
        &mut budget,
    )?;
    if !candidate.same_identity(&selection) {
        return Err(SlideMediaDataError::Verification);
    }
    let candidate_bytes = read_selected_media(&package, &candidate, part, &mut budget)?;
    if candidate_bytes != replacement || Sha1::digest(candidate_bytes).as_slice() != after_digest {
        return Err(SlideMediaDataError::Verification);
    }
    if !super::rendering_invalidation::root_previews_absent(physical_catalog(&package)?.package())
        .map_err(|_| SlideMediaDataError::Verification)?
    {
        return Err(SlideMediaDataError::Verification);
    }
    verify_zip_locality(
        catalog,
        physical_catalog(&package)?,
        &data_name,
        preview_plan.names(),
    )?;
    let patch = SlideMediaDataPatch {
        artifacts: ExactArtifacts::new(source_bytes, target_bytes),
        selection,
        part,
        before_digest,
        after_digest,
        before_length: before.len(),
        after_length: replacement.len(),
        source_preview_count,
        target_preview_count,
        deleted_previews,
    };
    Ok(SlideMediaDataCommit {
        package,
        patch,
        diagnostics: SlideMediaDataDiagnostics::published(2, deleted_previews),
    })
}

fn validate_replacement(
    replacement: &[u8],
    part: MediaPart,
    before: &[u8],
    limits: litchi_iwa_archive::Limits,
) -> Result<(), SlideMediaDataError> {
    if replacement.is_empty() {
        return Err(SlideMediaDataError::EmptyReplacement);
    }
    let replacement_length = u64::try_from(replacement.len()).unwrap_or(u64::MAX);
    if replacement.len() > MAX_REPLACEMENT_BYTES || replacement_length > limits.max_entry_bytes() {
        return Err(SlideMediaDataError::ReplacementTooLarge);
    }
    let before_type = MediaType::from_bytes(before);
    let after_type = MediaType::from_bytes(replacement);
    if before_type != MediaType::Unknown
        && after_type != MediaType::Unknown
        && before_type != after_type
    {
        return Err(SlideMediaDataError::ReplacementType);
    }
    if part == MediaPart::Poster && before_type == MediaType::Audio {
        return Err(SlideMediaDataError::AudioPoster);
    }
    Ok(())
}

/// Confirm that candidate reassembly changed only the selected data member,
/// PackageMetadata, and the root rendering previews explicitly invalidated by
/// the transaction.  Raw names, local records, compressed payloads, and
/// central-directory metadata for every untouched member remain authoritative.
fn verify_zip_locality(
    source: &SourceCatalog,
    candidate: &SourceCatalog,
    selected_data_name: &str,
    deleted_preview_names: &[&str],
) -> Result<(), SlideMediaDataError> {
    let mut candidates = candidate.package().iter();
    for entry in source.package().iter() {
        if deleted_preview_names.contains(&entry.name()) {
            continue;
        }
        let candidate_entry = candidates.next().ok_or(SlideMediaDataError::Verification)?;
        if entry.name() != candidate_entry.name() {
            return Err(SlideMediaDataError::Verification);
        }
        if entry.name() == selected_data_name || entry.name() == METADATA_COMPONENT {
            if !super::rendering_invalidation::selected_package_member_preserved(
                entry,
                candidate_entry,
            ) {
                return Err(SlideMediaDataError::Verification);
            }
        } else {
            if entry.raw_name() != candidate_entry.raw_name()
                || entry.is_opaque() != candidate_entry.is_opaque()
                || entry.data() != candidate_entry.data()
                || entry.metadata() != candidate_entry.metadata()
                || entry.raw_record().local_record() != candidate_entry.raw_record().local_record()
                || entry.raw_record().compressed_data()
                    != candidate_entry.raw_record().compressed_data()
                || !super::rendering_invalidation::central_record_preserved(
                    entry.raw_record().central_directory_record(),
                    candidate_entry.raw_record().central_directory_record(),
                )
            {
                return Err(SlideMediaDataError::Verification);
            }
        }
    }
    if candidates.next().is_some() {
        return Err(SlideMediaDataError::Verification);
    }
    Ok(())
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideMediaDataError> {
    match selector {
        SlideSelector::Position(position) => package
            .slide_record_at(position.get())
            .map_err(|_| SlideMediaDataError::Read)?
            .map(|_| position)
            .ok_or(SlideMediaDataError::SlidePositionNotFound { position }),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideMediaDataError::EmptySlideName);
            }
            let selected = package
                .show()
                .map_err(|_| SlideMediaDataError::Read)?
                .select_slide(SlideSelector::name(name))
                .map_err(|error| match error {
                    crate::SlideSelectorError::DuplicateSlideName { .. } => {
                        SlideMediaDataError::AmbiguousSelector
                    },
                    crate::SlideSelectorError::EmptySlideName => {
                        SlideMediaDataError::EmptySlideName
                    },
                })?;
            selected
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideMediaDataError::SlideNameNotFound)
        },
    }
}

fn copy_boxed_str(value: &str, budget: &mut MediaBudget) -> Result<Box<str>, SlideMediaDataError> {
    budget.allocation(value.len())?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_| SlideMediaDataError::Allocation {
            amount: value.len(),
        })?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

fn copy_data_path(name: &str, budget: &mut MediaBudget) -> Result<String, SlideMediaDataError> {
    let length = DATA_PREFIX
        .len()
        .checked_add(name.len())
        .ok_or(SlideMediaDataError::InvalidSource)?;
    budget.allocation(length)?;
    let mut path = String::new();
    path.try_reserve_exact(length)
        .map_err(|_| SlideMediaDataError::Allocation { amount: length })?;
    path.push_str(DATA_PREFIX);
    path.push_str(name);
    Ok(path)
}

fn copy_media_record(
    record: &MediaRecord,
    budget: &mut MediaBudget,
) -> Result<MediaRecord, SlideMediaDataError> {
    Ok(MediaRecord {
        identifier: record.identifier,
        digest: record.digest,
        preferred_name: copy_boxed_str(record.preferred_name.as_ref(), budget)?,
        current_name: copy_boxed_str(record.current_name.as_ref(), budget)?,
        materialized_length: record.materialized_length,
        unknown_fields: record.unknown_fields,
    })
}

pub(super) fn select_media(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    movie_selector: MovieSelector,
    part: MediaPart,
    budget: &mut MediaBudget,
) -> Result<MediaSelection, SlideMediaDataError> {
    budget.slides(1)?;
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(|_| SlideMediaDataError::Read)?
        .ok_or(SlideMediaDataError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let (_node_component, _node) = package
        .object_with_component(record.node_identifier)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let slide_payload =
        super::unique_payload(&slide.messages, &[SLIDE_MESSAGE_TYPE], "Keynote slide")
            .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let limits = package
        .semantic_wire_limits()
        .map_err(|_| SlideMediaDataError::Read)?;
    let references = repeated_drawable_references(slide_payload, limits, budget)?;
    let mut movies = Vec::new();
    budget.allocation(
        references
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or(SlideMediaDataError::InvalidSource)?,
    )?;
    movies
        .try_reserve_exact(references.len())
        .map_err(|_| SlideMediaDataError::Allocation {
            amount: references.len(),
        })?;
    for identifier in references {
        let Some((_, movie)) = package.object_with_component(identifier) else {
            return Err(SlideMediaDataError::InvalidSource);
        };
        budget.wire_work(movie.messages.len())?;
        let movie_count = movie
            .messages
            .iter()
            .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
            .count();
        if movie_count == 0 {
            continue;
        }
        if movie_count != 1 {
            return Err(SlideMediaDataError::InvalidSource);
        }
        movies.push(identifier);
    }
    let movie_position = movie_selector.as_position();
    let movie_identifier =
        *movies
            .get(movie_position.get())
            .ok_or(SlideMediaDataError::MoviePositionNotFound {
                position: movie_position,
            })?;
    let (movie_component, movie) = package
        .object_with_component(movie_identifier)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    if movie_component != component_name || movie.messages.len() != 1 {
        return Err(SlideMediaDataError::InvalidSource);
    }
    validate_selected_message_metadata(movie, 0)?;
    let movie_payload = movie
        .messages
        .first()
        .filter(|message| message.type_ == MOVIE_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(SlideMediaDataError::InvalidSource)?;
    budget.input_bytes(movie_payload.len())?;
    budget.wire_work(movie_payload.len())?;
    let (movie_info, movie_references) = super::decode_movie_info(
        movie_payload,
        limits,
        SemanticPath::SlideDrawable {
            slide: slide_position.get(),
            index: movie_position.get(),
        },
    )
    .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.references(movie_references)?;
    let kind = movie_info.kind();
    if !matches!(kind, MovieKind::File | MovieKind::Audio) {
        return Err(SlideMediaDataError::InvalidSource);
    }
    if part == MediaPart::Poster && kind == MovieKind::Audio {
        return Err(SlideMediaDataError::AudioPoster);
    }
    let parent = movie_parent_identifier(movie_payload, limits, budget)?;
    if parent != record.slide_identifier {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let content_identifier =
        unique_movie_data_identifier(movie_payload, MOVIE_DATA_FIELD, limits, budget)?;
    let poster_identifier =
        unique_movie_data_identifier(movie_payload, POSTER_IMAGE_DATA_FIELD, limits, budget)?;
    if content_identifier.is_none() {
        return Err(SlideMediaDataError::InvalidSource);
    }
    if part == MediaPart::Poster && poster_identifier.is_none() {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let component_name = copy_boxed_str(component_name, budget)?;
    let selection = MediaSelection {
        slide_position,
        movie_position,
        slide_identifier: record.slide_identifier,
        slide_node_identifier: record.node_identifier,
        movie_identifier,
        component_name: component_name.into(),
        kind,
        content_identifier,
        poster_identifier,
        record: None,
    };
    let record = validate_media_closure(package, &selection, part, budget)?;
    let mut selection = selection;
    selection.record = Some(record);
    Ok(selection)
}

fn repeated_drawable_references(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut MediaBudget,
) -> Result<Vec<u64>, SlideMediaDataError> {
    budget.input_bytes(payload.len())?;
    budget.wire_work(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.wire_fields(fields.len())?;
    budget.wire_nesting(1)?;
    let reference_capacity = fields
        .fields()
        .filter(|field| field.number() == SLIDE_OWNED_DRAWABLES_FIELD)
        .count();
    budget.references(reference_capacity)?;
    budget.allocation(
        reference_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(SlideMediaDataError::InvalidSource)?,
    )?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(reference_capacity)
        .map_err(|_| SlideMediaDataError::Allocation {
            amount: reference_capacity,
        })?;
    for field in fields.fields() {
        if field.number() != SLIDE_OWNED_DRAWABLES_FIELD {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(SlideMediaDataError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaDataError::InvalidSource)?;
        references.push(
            super::validate_reference_payload(
                field.payload(),
                limits,
                "Keynote slide media drawable",
            )
            .map_err(|_| SlideMediaDataError::InvalidSource)?,
        );
    }
    Ok(references)
}

fn unique_movie_data_identifier(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut MediaBudget,
) -> Result<Option<u64>, SlideMediaDataError> {
    budget.input_bytes(payload.len())?;
    budget.wire_work(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.wire_fields(fields.len())?;
    budget.wire_nesting(1)?;
    let mut identifier = None;
    for field in fields
        .fields()
        .filter(|field| field.number() == field_number)
    {
        if identifier.is_some() || field.wire_type() != 2 {
            return Err(SlideMediaDataError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaDataError::InvalidSource)?;
        let options = budget.keynote_options(field.payload())?;
        let (snapshot, report) =
            keynote_media_codec::decode_data_reference_with_report(field.payload(), options)
                .map_err(map_keynote_decode_error)?;
        budget.keynote_report(report)?;
        identifier = Some(snapshot.identifier());
    }
    Ok(identifier)
}

fn movie_parent_identifier(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut MediaBudget,
) -> Result<u64, SlideMediaDataError> {
    budget.input_bytes(payload.len())?;
    budget.wire_work(payload.len())?;
    let fields = WireView::parse_with_limits(payload, limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.wire_fields(fields.len())?;
    budget.wire_nesting(1)?;
    let mut parent = None;
    for field in fields
        .fields()
        .filter(|field| field.number() == MOVIE_SUPER_FIELD)
    {
        if parent.is_some() || field.wire_type() != 2 {
            return Err(SlideMediaDataError::InvalidSource);
        }
        field
            .validate_canonical_framing()
            .map_err(|_| SlideMediaDataError::InvalidSource)?;
        let super_fields = WireView::parse_with_limits(field.payload(), limits)
            .map_err(|_| SlideMediaDataError::InvalidSource)?;
        budget.input_bytes(field.payload().len())?;
        budget.wire_work(field.payload().len())?;
        budget.wire_fields(super_fields.len())?;
        budget.wire_nesting(2)?;
        for parent_field in super_fields
            .fields()
            .filter(|candidate| candidate.number() == DRAWABLE_PARENT_FIELD)
        {
            if parent.is_some() || parent_field.wire_type() != 2 {
                return Err(SlideMediaDataError::InvalidSource);
            }
            parent_field
                .validate_canonical_framing()
                .map_err(|_| SlideMediaDataError::InvalidSource)?;
            parent = Some(
                super::validate_reference_payload(
                    parent_field.payload(),
                    limits,
                    "Keynote movie parent reference",
                )
                .map_err(|_| SlideMediaDataError::InvalidSource)?,
            );
        }
    }
    parent.ok_or(SlideMediaDataError::InvalidSource)
}

/// Admit only the rooted physical closure that the selected part will edit.
///
/// `validated_media_record` performs the strict metadata, owner, archive-info,
/// ZIP-member, digest, length, basename, and known-type checks.  Keeping this
/// seam explicit makes the selector path auditable: no caller can obtain or
/// mutate bytes before the selected graph edge is rooted in one current
/// component and one unambiguous `DataInfo` record.
fn validate_media_closure(
    package: &Package,
    selection: &MediaSelection,
    part: MediaPart,
    budget: &mut MediaBudget,
) -> Result<MediaRecord, SlideMediaDataError> {
    let catalog = physical_catalog(package)?;
    validated_media_record(package, catalog, selection, part, budget)
}

pub(super) fn validate_selected_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), SlideMediaDataError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(SlideMediaDataError::InvalidSource);
    }
    Ok(())
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, SlideMediaDataError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideMediaDataError::UnsupportedSource),
    }
}

pub(super) fn read_selected_media<'a>(
    package: &'a Package,
    selection: &MediaSelection,
    part: MediaPart,
    budget: &mut MediaBudget,
) -> Result<&'a [u8], SlideMediaDataError> {
    let catalog = physical_catalog(package)?;
    let record = selection
        .record
        .as_ref()
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let _ = part;
    let mut matches = catalog.package().iter().filter(|entry| {
        entry
            .name()
            .strip_prefix(DATA_PREFIX)
            .is_some_and(|name| name == record.current_name.as_ref())
    });
    let entry = matches.next().ok_or(SlideMediaDataError::InvalidSource)?;
    if matches.next().is_some() || entry.is_opaque() {
        return Err(SlideMediaDataError::InvalidSource);
    }
    // Bound the member and account for the digest pass before hashing its
    // bytes.  The digest is part of source admission, so hostile entry data
    // must not receive uncharged work or bypass the caller's residual cap.
    budget.entry_bytes(entry.data().len())?;
    budget.wire_work(entry.data().len())?;
    if record.materialized_length != Some(entry.data().len())
        || Sha1::digest(entry.data()).as_slice() != record.digest
    {
        return Err(SlideMediaDataError::InvalidSource);
    }
    Ok(entry.data())
}

fn validated_media_record(
    package: &Package,
    catalog: &SourceCatalog,
    selection: &MediaSelection,
    part: MediaPart,
    budget: &mut MediaBudget,
) -> Result<MediaRecord, SlideMediaDataError> {
    let identifier = selection.identifier(part).ok_or_else(|| {
        if selection.kind == MovieKind::Audio && part == MediaPart::Poster {
            SlideMediaDataError::AudioPoster
        } else {
            SlideMediaDataError::InvalidSource
        }
    })?;
    let facts = metadata_facts(package, catalog, selection, identifier, budget)?;
    let locator = selection
        .component_name
        .strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let mut components = facts
        .components
        .iter()
        .filter(|component| !component.versioned && component.locator.as_ref() == locator);
    let component = components
        .next()
        .ok_or(SlideMediaDataError::InvalidSource)?;
    if components.next().is_some() {
        return Err(SlideMediaDataError::InvalidSource);
    }
    // The selected record is rewritten by raw field spans; unknown fields are
    // therefore retained byte-for-byte instead of being decoded or dropped.
    let _selected_component_has_unknown_fields = component.unknown_fields;
    let mut records = facts
        .records
        .iter()
        .filter(|record| record.identifier == identifier);
    let record = records.next().ok_or(SlideMediaDataError::InvalidSource)?;
    if records.next().is_some() {
        return Err(SlideMediaDataError::InvalidSource);
    }
    if record.current_name.is_empty()
        || record.current_name.contains(['/', '\\', '\0'])
        || record.current_name.as_ref() == "."
        || record.current_name.as_ref() == ".."
        || record.preferred_name.is_empty()
        || record.preferred_name.contains(['/', '\\', '\0'])
    {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let materialized_length = record
        .materialized_length
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let _selected_record_has_unknown_fields = record.unknown_fields;
    // A selected data name must not alias an unrelated DataInfo declaration.
    if facts
        .records
        .iter()
        .filter(|candidate| candidate.current_name == record.current_name)
        .count()
        != 1
    {
        return Err(SlideMediaDataError::InvalidSource);
    }
    // A shared DataInfo is a graph asset, so validate every current ownership
    // group that names it. The selected movie must be rooted in its own group,
    // while other valid owners (including shared posters) stay attached to the
    // same physical asset and observe the replacement.
    let mut current_reference_count = 0usize;
    let mut selected_group_count = 0usize;
    for &(component_identifier, data_identifier, declared_owner_count, versioned) in
        &facts.references
    {
        if versioned || data_identifier != identifier {
            continue;
        }
        current_reference_count = current_reference_count
            .checked_add(1)
            .ok_or(SlideMediaDataError::InvalidSource)?;
        let mut current_components = facts.components.iter().filter(|candidate| {
            !candidate.versioned && candidate.identifier == component_identifier
        });
        let owner_component = current_components
            .next()
            .ok_or(SlideMediaDataError::InvalidSource)?;
        if current_components.next().is_some() || declared_owner_count == 0 {
            return Err(SlideMediaDataError::InvalidSource);
        }
        let mut owner_count = 0usize;
        let mut selected_owner_count = 0usize;
        for owner in facts.owners.iter().filter(|owner| {
            !owner.versioned
                && owner.component_identifier == component_identifier
                && owner.data_identifier == identifier
        }) {
            if owner.count == 0 || owner.object_identifier == 0 {
                return Err(SlideMediaDataError::InvalidSource);
            }
            let (owner_component_name, owner_object) = package
                .object_with_component(owner.object_identifier)
                .ok_or(SlideMediaDataError::InvalidSource)?;
            let owner_locator = owner_component_name
                .strip_prefix("Index/")
                .and_then(|name| name.strip_suffix(".iwa"))
                .ok_or(SlideMediaDataError::InvalidSource)?;
            if owner_locator != owner_component.locator.as_ref() {
                return Err(SlideMediaDataError::InvalidSource);
            }
            let archive_references = owner_object
                .archive_info
                .message_infos
                .iter()
                .flat_map(|info| info.data_references.iter())
                .filter(|value| **value == identifier)
                .count();
            if archive_references == 0 {
                return Err(SlideMediaDataError::InvalidSource);
            }
            owner_count = owner_count
                .checked_add(1)
                .ok_or(SlideMediaDataError::InvalidSource)?;
            if component_identifier == component.identifier
                && owner.object_identifier == selection.movie_identifier
            {
                selected_owner_count = selected_owner_count
                    .checked_add(1)
                    .ok_or(SlideMediaDataError::InvalidSource)?;
            }
        }
        if owner_count != declared_owner_count as usize {
            return Err(SlideMediaDataError::InvalidSource);
        }
        if component_identifier == component.identifier {
            selected_group_count = selected_group_count
                .checked_add(1)
                .ok_or(SlideMediaDataError::InvalidSource)?;
            if selected_owner_count != 1 {
                return Err(SlideMediaDataError::InvalidSource);
            }
        }
    }
    if current_reference_count == 0 || selected_group_count != 1 {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let movie = package
        .object(selection.movie_identifier)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let (message_index, _message) = movie
        .messages
        .iter()
        .enumerate()
        .find(|(_, message)| message.type_ == MOVIE_MESSAGE_TYPE)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let info = movie
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    if info
        .data_references
        .iter()
        .filter(|value| **value == identifier)
        .count()
        != 1
    {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let name_type = record
        .current_name
        .rsplit_once('.')
        .map_or(MediaType::Unknown, |(_, extension)| {
            MediaType::from_extension(extension)
        });
    let expected_type = match part {
        MediaPart::Content if selection.kind == MovieKind::Audio => MediaType::Audio,
        MediaPart::Content => MediaType::Video,
        MediaPart::Poster => MediaType::Image,
    };
    if name_type != MediaType::Unknown && name_type != expected_type {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let mut entries = catalog.package().iter().filter(|entry| {
        entry
            .name()
            .strip_prefix(DATA_PREFIX)
            .is_some_and(|name| name == record.current_name.as_ref())
    });
    let entry = entries.next().ok_or(SlideMediaDataError::InvalidSource)?;
    if entries.next().is_some()
        || entry.is_opaque()
        || entry.data().len() != materialized_length
        || Sha1::digest(entry.data()).as_slice() != record.digest
    {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let content_type = MediaType::from_bytes(entry.data());
    if content_type != MediaType::Unknown && content_type != expected_type {
        return Err(SlideMediaDataError::InvalidSource);
    }
    copy_media_record(record, budget)
}

/// Derive the metadata Snappy bound from the already-parsed authoritative
/// component. ZIP entry lengths describe compressed bytes and therefore cannot
/// safely cap the decompressed IWA stream. The catalog archive's exact encoded
/// length is a source-preserving, allocation-free preflight for that stream.
fn metadata_archive_inventory_work(archive: &Archive) -> Result<usize, SlideMediaDataError> {
    let mut work = archive.objects.len();
    for (object_index, object) in archive.objects.iter().enumerate() {
        // `encoded_len_with_limits` checks duplicate object identifiers by
        // scanning the preceding prefix. Account for that bounded lookup as
        // well as the object/header/message inventory it traverses.
        work = work
            .checked_add(object_index)
            .and_then(|value| value.checked_add(1))
            .and_then(|value| value.checked_add(object.messages.len()))
            .and_then(|value| value.checked_add(object.archive_info.message_infos.len()))
            .and_then(|value| {
                usize::try_from(object.header_length)
                    .ok()
                    .and_then(|header| value.checked_add(header))
            })
            .ok_or(SlideMediaDataError::InvalidSource)?;
        for message in &object.messages {
            work = work
                .checked_add(message.data.len())
                .and_then(|value| value.checked_add(1))
                .ok_or(SlideMediaDataError::InvalidSource)?;
        }
        for info in &object.archive_info.message_infos {
            work = work
                .checked_add(info.versions.len())
                .and_then(|value| value.checked_add(info.field_infos.len()))
                .and_then(|value| value.checked_add(info.object_references.len()))
                .and_then(|value| value.checked_add(info.data_references.len()))
                .and_then(|value| value.checked_add(info.diff_merge_version.len()))
                .and_then(|value| value.checked_add(usize::from(info.diff_field_path.is_some())))
                .and_then(|value| value.checked_add(info.fields_to_remove.len()))
                .and_then(|value| value.checked_add(info.diff_read_version.len()))
                .ok_or(SlideMediaDataError::InvalidSource)?;
            if let Some(path) = &info.diff_field_path {
                work = work
                    .checked_add(path.path.len())
                    .ok_or(SlideMediaDataError::InvalidSource)?;
            }
            for path in &info.fields_to_remove {
                work = work
                    .checked_add(path.path.len())
                    .ok_or(SlideMediaDataError::InvalidSource)?;
            }
            for field in &info.field_infos {
                work = work
                    .checked_add(1)
                    .and_then(|value| value.checked_add(field.path.path.len()))
                    .and_then(|value| value.checked_add(field.object_references.len()))
                    .and_then(|value| value.checked_add(field.data_references.len()))
                    .and_then(|value| value.checked_add(field.known_field_version.len()))
                    .and_then(|value| {
                        value.checked_add(usize::from(
                            field.known_field_feature_identifier.is_some(),
                        ))
                    })
                    .ok_or(SlideMediaDataError::InvalidSource)?;
            }
        }
    }
    Ok(work)
}

fn metadata_stream_profile(
    package: &Package,
    catalog: &SourceCatalog,
    budget: &mut MediaBudget,
) -> Result<(usize, SnappyLimits, ArchiveLimits), SlideMediaDataError> {
    let metadata_component = catalog
        .components()
        .get(METADATA_COMPONENT)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let archive_limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let metadata_archive = metadata_component.archive();
    // Precharge the complete object/header/reference/payload inventory before
    // the archive preflight itself. This keeps encoded-length validation from
    // becoming an uncharged traversal when a sibling residual work cap is
    // already tight.
    let inventory_work = metadata_archive_inventory_work(metadata_archive)?;
    budget.wire_work(inventory_work)?;
    let stream_capacity = metadata_archive
        .encoded_len_with_limits(archive_limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    if stream_capacity == 0 {
        return Err(SlideMediaDataError::InvalidSource);
    }
    // A newly constructed archive may not carry source offsets/header lengths
    // in its provenance. Cover any stream bytes not represented by the
    // inventory bound before the decompression passes begin.
    let uncovered_work = stream_capacity.checked_sub(inventory_work).unwrap_or(0);
    budget.wire_work(uncovered_work)?;
    let source_limits = package
        .limits()
        .snappy_limits()
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let max_stream = source_limits.max_decompressed_stream().min(stream_capacity);
    let max_chunk = source_limits.max_uncompressed_chunk().min(max_stream);
    let snappy_limits = SnappyLimits::new(max_chunk, max_stream)
        .and_then(|limits| {
            limits.with_input_limits(
                source_limits.max_compressed_chunk(),
                source_limits.max_compressed_stream(),
                source_limits.max_frames(),
            )
        })
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    Ok((stream_capacity, snappy_limits, archive_limits))
}

fn metadata_facts(
    package: &Package,
    catalog: &SourceCatalog,
    selection: &MediaSelection,
    identifier: u64,
    budget: &mut MediaBudget,
) -> Result<OwnedMetadataFacts, SlideMediaDataError> {
    let metadata_archive = catalog
        .components()
        .get(METADATA_COMPONENT)
        .ok_or(SlideMediaDataError::InvalidSource)?
        .archive();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == METADATA_COMPONENT)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideMediaDataError::InvalidSource);
    }
    budget.entry_bytes(entry.data().len())?;
    budget.input_bytes(entry.data().len())?;
    let (stream_capacity, snappy_limits, _archive_limits) =
        metadata_stream_profile(package, catalog, budget)?;
    // Account for the decoded-stream pass, canonical framing pass, and the
    // archive/message census below before allocating or traversing them.
    let stream_work = entry
        .data()
        .len()
        .checked_add(
            stream_capacity
                .checked_mul(3)
                .ok_or(SlideMediaDataError::InvalidSource)?,
        )
        .ok_or(SlideMediaDataError::InvalidSource)?;
    budget.wire_work(stream_work)?;
    budget.entry_bytes(stream_capacity)?;
    budget.allocation(stream_capacity)?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    if stream.as_bytes().len() != stream_capacity {
        return Err(SlideMediaDataError::InvalidSource);
    }
    metadata_archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let mut payload = None;
    for object in &metadata_archive.objects {
        if object.messages.len() != object.archive_info.message_infos.len() {
            return Err(SlideMediaDataError::InvalidSource);
        }
        for (index, message) in object.messages.iter().enumerate() {
            let info = object
                .archive_info
                .message_infos
                .get(index)
                .ok_or(SlideMediaDataError::InvalidSource)?;
            if message.type_ != info.type_
                || usize::try_from(info.length).ok() != Some(message.data.len())
            {
                return Err(SlideMediaDataError::InvalidSource);
            }
            if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                && payload.replace(message.data.as_slice()).is_some()
            {
                return Err(SlideMediaDataError::InvalidSource);
            }
        }
    }
    let payload = payload.ok_or(SlideMediaDataError::InvalidSource)?;
    let options = budget.metadata_options(payload)?;
    let report = metadata_codec::inspect_package_metadata_media(payload, options)
        .map_err(map_metadata_decode_error)?;
    budget.metadata_report(report)?;
    // The visitor is a second strict pass over the same borrowed payload.  Its
    // finite work is charged before it is allowed to allocate the owned facts.
    budget.metadata_report(report)?;
    let mut facts = OwnedMetadataFacts::default();
    let record_capacity = report.data_records();
    let component_capacity = report.components();
    let reference_capacity = report.data_references();
    let owner_capacity = report.owners();
    budget.allocation(
        record_capacity
            .checked_mul(size_of::<MediaRecord>())
            .ok_or(SlideMediaDataError::InvalidSource)?,
    )?;
    budget.allocation(
        component_capacity
            .checked_mul(size_of::<ComponentRecord>())
            .ok_or(SlideMediaDataError::InvalidSource)?,
    )?;
    budget.allocation(
        reference_capacity
            .checked_mul(size_of::<(u64, u64, u32, bool)>())
            .ok_or(SlideMediaDataError::InvalidSource)?,
    )?;
    budget.allocation(
        owner_capacity
            .checked_mul(size_of::<OwnerRecord>())
            .ok_or(SlideMediaDataError::InvalidSource)?,
    )?;
    facts
        .records
        .try_reserve_exact(record_capacity)
        .map_err(|_| SlideMediaDataError::Allocation {
            amount: record_capacity,
        })?;
    facts
        .components
        .try_reserve_exact(component_capacity)
        .map_err(|_| SlideMediaDataError::Allocation {
            amount: component_capacity,
        })?;
    facts
        .references
        .try_reserve_exact(reference_capacity)
        .map_err(|_| SlideMediaDataError::Allocation {
            amount: reference_capacity,
        })?;
    facts
        .owners
        .try_reserve_exact(owner_capacity)
        .map_err(|_| SlideMediaDataError::Allocation {
            amount: owner_capacity,
        })?;
    let mut visitor = MetadataVisitor { facts, budget };
    let visited_report =
        metadata_codec::visit_package_metadata_media(payload, options, &mut visitor)
            .map_err(map_metadata_decode_error)?;
    if visited_report != report {
        return Err(SlideMediaDataError::InvalidSource);
    }
    let MetadataVisitor { facts, budget } = visitor;
    closure::validate_selected_media_closure(
        package, payload, &facts, selection, identifier, budget,
    )?;
    Ok(facts)
}

fn rewrite_media(
    source: &Package,
    selection: &MediaSelection,
    part: MediaPart,
    replacement: &[u8],
    budget: &mut MediaBudget,
) -> Result<(Package, usize), SlideMediaDataError> {
    let catalog = physical_catalog(source)?;
    let record = selection
        .record
        .as_ref()
        .ok_or(SlideMediaDataError::InvalidSource)?;
    let old = catalog
        .package()
        .iter()
        .find(|entry| {
            entry
                .name()
                .strip_prefix(DATA_PREFIX)
                .is_some_and(|name| name == record.current_name.as_ref())
        })
        .ok_or(SlideMediaDataError::InvalidSource)?
        .data();
    validate_replacement(replacement, part, old, source.state.options.archive())?;
    let metadata = rewrite_metadata_entry(
        source,
        catalog,
        selection
            .identifier(part)
            .ok_or(SlideMediaDataError::InvalidSource)?,
        &record.digest,
        record
            .materialized_length
            .ok_or(SlideMediaDataError::InvalidSource)?,
        replacement,
        budget,
    )?;
    let data_name = copy_data_path(record.current_name.as_ref(), budget)?;
    let edits = [
        EntryEdit::new(&data_name, replacement),
        EntryEdit::new(METADATA_COMPONENT, &metadata),
    ];
    let preview_plan = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let prepared = catalog
        .package()
        .prepare_reassembly_with_changes(&[], &edits, preview_plan.names(), catalog.limits())
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let requirements = prepared.execution_requirements();
    budget.reassembly_requirements(requirements)?;
    budget.input_bytes(requirements.output_bytes())?;
    budget.wire_work(requirements.output_bytes())?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let package = Package::from_source_with_options(Arc::from(output), source.state.options)
        .map_err(|_| SlideMediaDataError::Verification)?;
    package
        .validate()
        .map_err(|_| SlideMediaDataError::Verification)?;
    if physical_catalog(&package)?
        .package()
        .iter()
        .find(|entry| entry.name() == METADATA_COMPONENT)
        .is_none_or(|entry| entry.data() != metadata)
    {
        return Err(SlideMediaDataError::Verification);
    }
    Ok((package, preview_plan.len()))
}

fn rewrite_metadata_entry(
    package: &Package,
    catalog: &SourceCatalog,
    target_identifier: u64,
    expected_digest: &[u8; SHA1_BYTES],
    expected_length: usize,
    replacement: &[u8],
    budget: &mut MediaBudget,
) -> Result<Vec<u8>, SlideMediaDataError> {
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == METADATA_COMPONENT)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideMediaDataError::InvalidSource);
    }
    budget.entry_bytes(entry.data().len())?;
    budget.input_bytes(entry.data().len())?;
    let (stream_capacity, snappy_limits, archive_limits) =
        metadata_stream_profile(package, catalog, budget)?;
    // The rewrite keeps the decoded stream live while parsing an owned
    // Archive, checking canonical framing, and scanning its message census.
    // Reserve all four stream-sized passes before decompression begins.
    let stream_work = entry
        .data()
        .len()
        .checked_add(
            stream_capacity
                .checked_mul(4)
                .ok_or(SlideMediaDataError::InvalidSource)?,
        )
        .ok_or(SlideMediaDataError::InvalidSource)?;
    budget.wire_work(stream_work)?;
    budget.entry_bytes(stream_capacity)?;
    budget.allocation(stream_capacity)?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    if stream.as_bytes().len() != stream_capacity {
        return Err(SlideMediaDataError::InvalidSource);
    }
    // Parsing an Archive owns each message payload independently of the
    // decoded Snappy stream. Reserve the full source bound before that copy.
    budget.allocation(stream_capacity)?;
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let mut location = None;
    for (object_index, object) in archive.objects.iter().enumerate() {
        if object.messages.len() != object.archive_info.message_infos.len() {
            return Err(SlideMediaDataError::InvalidSource);
        }
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE
                && location
                    .replace((object_index, message_index, message.data.as_slice()))
                    .is_some()
            {
                return Err(SlideMediaDataError::InvalidSource);
            }
        }
    }
    let (object_index, message_index, payload) =
        location.ok_or(SlideMediaDataError::InvalidSource)?;
    budget.wire_work(replacement.len())?;
    let replacement_digest: [u8; SHA1_BYTES] = Sha1::digest(replacement).into();
    let expected_length =
        u64::try_from(expected_length).map_err(|_| SlideMediaDataError::InvalidSource)?;
    let replacement_length =
        u64::try_from(replacement.len()).map_err(|_| SlideMediaDataError::InvalidSource)?;
    let replacement = metadata_codec::DataInfoContentReplacement::new(
        target_identifier,
        expected_digest,
        &replacement_digest,
        expected_length,
        replacement_length,
    );
    let replacements = [replacement];
    let options = budget.metadata_options(payload)?;
    let prepared = metadata_codec::prepare_package_metadata_media_content_replacements(
        payload,
        &replacements,
        options,
    )
    .map_err(map_metadata_rewrite_error)?;
    budget.metadata_report(prepared.source_report())?;
    let requirements = prepared.execution_requirements();
    budget.metadata_requirements(requirements)?;
    let rewritten = prepared
        .execute(requirements.exact_limits())
        .map_err(map_metadata_rewrite_error)?
        .into_bytes();
    let object = archive
        .objects
        .get_mut(object_index)
        .ok_or(SlideMediaDataError::InvalidSource)?;
    object
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    let serialized_capacity = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.wire_work(serialized_capacity)?;
    budget.allocation(serialized_capacity)?;
    let serialized = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    if serialized.len() != serialized_capacity {
        return Err(SlideMediaDataError::Verification);
    }
    let compressed_capacity = SnappyStream::maximum_compressed_len(serialized.len())
        .map_err(|_| SlideMediaDataError::InvalidSource)?;
    budget.wire_work(compressed_capacity)?;
    budget.allocation(compressed_capacity)?;
    let compressed =
        SnappyStream::compress(&serialized).map_err(|_| SlideMediaDataError::InvalidSource)?;
    if compressed.len() > compressed_capacity {
        return Err(SlideMediaDataError::Verification);
    }
    Ok(compressed)
}
