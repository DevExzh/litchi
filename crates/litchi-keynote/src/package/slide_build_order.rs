//! Immutable, selector-first transactions for Keynote build playback order.
//!
//! This module owns only the physical order of existing build events. Build
//! effects, timing values, and build lifecycle remain outside this focused
//! transaction. The two native order lists (`SlideArchive.builds` and
//! `SlideArchive.buildChunks`) are rewritten together, with each chunk group
//! following its build while the original nested reference bytes remain
//! untouched.

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::str;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use thiserror::Error;

use super::{
    BUILD_MESSAGE_TYPE, Package, PhysicalSource, ReadError, SHOW_MESSAGE_TYPE, SLIDE_MESSAGE_TYPE,
    SLIDE_NODE_MESSAGE_TYPE, decode_show_snapshot, decode_slide_node_projection, unique_payload,
};
use crate::{SlideSelector, SlideSelectorError};

const BUILDS_FIELD: u32 = 2;
const BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;
const DEPRECATED_BUILD_CHUNKS_FIELD: u32 = 3;
const BUILD_CHUNKS_FIELD: u32 = 43;
const OWNED_DRAWABLES_FIELD: u32 = 7;
const DRAWABLE_FIELD: u32 = 1;
const DELIVERY_FIELD: u32 = 2;
const ATTRIBUTES_FIELD: u32 = 4;
const CHUNK_BUILD_FIELD: u32 = 1;
const CHUNK_AUTOMATIC_FIELD: u32 = 5;
const CHUNK_REFERENT_FIELD: u32 = 6;
const CHUNK_IDENTIFIER_FIELD: u32 = 7;
const CHUNK_BUILD_ID_FIELD: u32 = 8;
const ATTRIBUTES_ANIMATION_FIELD: u32 = 18;
const ANIMATION_AUTOMATIC_FIELD: u32 = 6;
const UUID_LOWER_FIELD: u32 = 1;
const UUID_UPPER_FIELD: u32 = 2;

/// A finite resource governed while a Keynote build-order transaction is
/// prepared.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SlideBuildOrderLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members or IWA objects.
    Entries,
    /// Bytes in one package member or IWA value.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic build objects.
    Builds,
    /// Semantic build timing chunks.
    BuildChunks,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
}

impl fmt::Display for SlideBuildOrderLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::Builds => "builds",
            Self::BuildChunks => "build chunks",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// An error raised while staging or committing a Keynote build-order
/// transaction.
///
/// Error values contain only semantic positions and resource measurements;
/// native object identities, member names, and lower wire types stay private.
#[derive(Debug, Error)]
#[non_exhaustive]
pub enum SlideBuildOrderError {
    /// This package was prepared for semantic reading only and has no editable
    /// physical source.
    #[error("this Keynote source does not support physical edits")]
    UnsupportedSource,
    /// The selected source graph uses an unsupported build-order topology.
    #[error("this Keynote build topology is not supported for reordering")]
    UnsupportedTopology,
    /// An exact-name selector matched more than one slide.
    #[error("the Keynote slide selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name selector matched no slide.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked source position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// Missing checked source position.
        position: Position,
    },
    /// A checked build position does not exist.
    #[error("the selected Keynote slide has no build at position {position:?}")]
    BuildPositionNotFound {
        /// Missing checked source build position.
        position: Position,
    },
    /// The requested final destination is outside the unchanged build count.
    #[error(
        "build-order destination {position:?} is outside the selected slide's {build_count} builds"
    )]
    DestinationOutOfRange {
        /// Invalid checked final position.
        position: Position,
        /// Number of builds in the immutable base snapshot.
        build_count: usize,
    },
    /// A second operation was staged in the same bounded transaction.
    #[error("the Keynote build-order transaction already has a staged operation")]
    OperationAlreadyStaged,
    /// Commit was requested without a staged operation.
    #[error("the Keynote build-order transaction has no staged operation")]
    NoStagedOperation,
    /// The source package or selected wire payload is structurally invalid.
    #[error("the Keynote source cannot be reordered safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error("Keynote build-order {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: SlideBuildOrderLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote build-order transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full candidate reopening did not reproduce the requested order.
    #[error("the reordered Keynote candidate failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote build-order patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Intent {
    slide: Position,
    source: Position,
    destination: Position,
}

/// A bounded build-order edit staged against an immutable Keynote snapshot.
#[derive(Debug)]
pub struct SlideBuildOrderEdit<'a> {
    source: &'a Package,
    intent: Option<Intent>,
}

impl<'a> SlideBuildOrderEdit<'a> {
    pub(super) const fn new(source: &'a Package) -> Self {
        Self {
            source,
            intent: None,
        }
    }

    /// Stage one build move on the slide selected by an exact navigator name
    /// or checked semantic slide position.
    ///
    /// `destination` is interpreted in the final build order, matching
    /// `Vec::remove(source); Vec::insert(destination, value)`.
    ///
    /// # Errors
    ///
    /// Returns a typed error when the slide or build is missing or ambiguous,
    /// the destination is outside the base build count, or an operation is
    /// already staged.
    pub fn move_build<'selector>(
        &mut self,
        slide_input: impl Into<SlideSelector<'selector>>,
        source: Position,
        destination: Position,
    ) -> Result<&mut Self, SlideBuildOrderError> {
        if self.intent.is_some() {
            return Err(SlideBuildOrderError::OperationAlreadyStaged);
        }

        let slide = resolve_slide_position(self.source, slide_input)?;
        let snapshot = read_build_order(self.source, slide)?;
        let build_count = snapshot.builds.len();
        if source.get() >= build_count {
            return Err(SlideBuildOrderError::BuildPositionNotFound { position: source });
        }
        if destination.get() >= build_count {
            return Err(SlideBuildOrderError::DestinationOutOfRange {
                position: destination,
                build_count,
            });
        }
        self.intent = Some(Intent {
            slide,
            source,
            destination,
        });
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    ///
    /// An exact same-position move reuses the source package allocation and
    /// bytes. A changed candidate is fully reopened and semantically read back
    /// under the original physical and semantic limits before publication.
    pub fn commit(self) -> Result<SlideBuildOrderCommit, SlideBuildOrderError> {
        let intent = self.intent.ok_or(SlideBuildOrderError::NoStagedOperation)?;
        let source_bytes = physical_shared_source(self.source)?;
        let source_fingerprint = fingerprint(&source_bytes);
        let before = read_build_order(self.source, intent.slide)?;
        ensure_intent_bounds(&before, intent)?;
        if intent.source == intent.destination {
            self.source.validate().map_err(map_read_error)?;
            return Ok(SlideBuildOrderCommit {
                package: self.source.snapshot(),
                patch: SlideBuildOrderPatch {
                    source_bytes: Arc::clone(&source_bytes),
                    target_bytes: source_bytes,
                    source_fingerprint,
                    target_fingerprint: source_fingerprint,
                    slide: intent.slide,
                    source: intent.source,
                    destination: intent.destination,
                },
                diagnostics: SlideBuildOrderDiagnostics::unchanged(),
            });
        }

        editable_source_catalog(self.source)?;
        self.source.validate().map_err(map_read_error)?;
        let after = moved_build_order(&before, intent)?;
        let package = rewrite_build_order(self.source, &before, intent, &after)?;
        let target_fingerprint = fingerprint(package.source_bytes());
        Ok(SlideBuildOrderCommit {
            patch: SlideBuildOrderPatch {
                source_bytes,
                target_bytes: editable_shared_source(&package)?,
                source_fingerprint,
                target_fingerprint,
                slide: intent.slide,
                source: intent.source,
                destination: intent.destination,
            },
            package,
            diagnostics: SlideBuildOrderDiagnostics::published(),
        })
    }
}

/// An exact-source-checked, reversible semantic Keynote build-order patch.
///
/// Native identities and package member names remain private. Public metadata
/// contains only the checked semantic slide and build positions.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideBuildOrderPatch {
    source_bytes: Arc<[u8]>,
    target_bytes: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    slide: Position,
    source: Position,
    destination: Position,
}

impl fmt::Debug for SlideBuildOrderPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideBuildOrderPatch")
            .field("slide", &self.slide)
            .field("source", &self.source)
            .field("destination", &self.destination)
            .finish_non_exhaustive()
    }
}

impl SlideBuildOrderPatch {
    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return the selected slide's semantic source position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide
    }

    /// Return the selected build's source position.
    #[must_use]
    pub const fn source_position(&self) -> Position {
        self.source
    }

    /// Return the selected build's final destination position.
    #[must_use]
    pub const fn destination_position(&self) -> Position {
        self.destination
    }

    /// Return whether this patch preserves the exact source order and bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.source.get() == self.destination.get()
            && self.source_fingerprint == self.target_fingerprint
            && (Arc::ptr_eq(&self.source_bytes, &self.target_bytes)
                || self.source_bytes.as_ref() == self.target_bytes.as_ref())
    }

    /// Return an exact reversible patch from the committed package back to its
    /// immutable source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_bytes: Arc::clone(&self.target_bytes),
            target_bytes: Arc::clone(&self.source_bytes),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            slide: self.slide,
            source: self.destination,
            destination: self.source,
        }
    }
}

/// Compact evidence describing one build-order commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideBuildOrderDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl SlideBuildOrderDiagnostics {
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

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable Keynote build-order transaction.
#[must_use = "a Keynote build-order commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideBuildOrderCommit {
    package: Package,
    patch: SlideBuildOrderPatch,
    diagnostics: SlideBuildOrderDiagnostics,
}

impl SlideBuildOrderCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideBuildOrderPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideBuildOrderDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Start one selector-first Keynote build-order edit.
    #[must_use]
    pub const fn edit_slide_build_order(&self) -> SlideBuildOrderEdit<'_> {
        SlideBuildOrderEdit::new(self)
    }

    /// Apply an exact-source-checked build-order patch.
    ///
    /// The retained target is fully reopened and semantically verified under
    /// this package's original physical and semantic limits.
    pub fn apply_slide_build_order(
        &self,
        patch: &SlideBuildOrderPatch,
    ) -> Result<SlideBuildOrderCommit, SlideBuildOrderError> {
        let source = physical_shared_source(self)?;
        if fingerprint(self.source_bytes()) != patch.source_fingerprint
            || self.source_bytes() != patch.source_bytes.as_ref()
            || source.as_ref() != patch.source_bytes.as_ref()
        {
            return Err(SlideBuildOrderError::PatchConflict);
        }
        let current = read_build_order(self, patch.slide)?;
        let intent = Intent {
            slide: patch.slide,
            source: patch.source,
            destination: patch.destination,
        };
        ensure_intent_bounds(&current, intent)?;

        if patch.is_noop() {
            if patch.source_bytes.as_ref() != patch.target_bytes.as_ref() {
                return Err(SlideBuildOrderError::PatchConflict);
            }
            self.validate().map_err(map_read_error)?;
            return Ok(SlideBuildOrderCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideBuildOrderDiagnostics::unchanged(),
            });
        }

        editable_source_catalog(self)?;
        if fingerprint(&patch.target_bytes) != patch.target_fingerprint {
            return Err(SlideBuildOrderError::PatchConflict);
        }
        self.validate().map_err(map_read_error)?;
        let expected = moved_build_order(&current, intent)?;
        let candidate =
            Package::from_source_with_options(Arc::clone(&patch.target_bytes), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let readback = read_build_order(&candidate, patch.slide)?;
        if !same_order(&readback, &expected) {
            return Err(SlideBuildOrderError::Verification);
        }
        Ok(SlideBuildOrderCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: SlideBuildOrderDiagnostics::published(),
        })
    }
}

#[derive(Debug)]
struct RawReference {
    identifier: u64,
    raw: Box<[u8]>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Uuid {
    lower: u64,
    upper: u64,
}

#[derive(Debug)]
struct BuildFact {
    identifier: u64,
    reference: RawReference,
    default_automatic: Option<bool>,
}

#[derive(Debug)]
struct ChunkFact {
    identifier: u64,
    reference: RawReference,
    build_identifier: u64,
    automatic: Option<bool>,
    referent: Option<bool>,
}

#[derive(Debug)]
struct BuildOrderSnapshot {
    slide_identifier: u64,
    component_name: String,
    builds: Vec<BuildFact>,
    chunks: Vec<ChunkFact>,
    chunk_indexes_by_build: Vec<Vec<usize>>,
}

fn resolve_slide_position<'selector>(
    source: &Package,
    selector_input: impl Into<SlideSelector<'selector>>,
) -> Result<Position, SlideBuildOrderError> {
    let selector = selector_input.into();
    let show = source.show().map_err(map_read_error)?;
    match selector {
        SlideSelector::Position(position) => {
            if position.get() >= show.slides().len() {
                return Err(SlideBuildOrderError::SlidePositionNotFound { position });
            }
            Ok(position)
        },
        SlideSelector::Name(_) => {
            let selected = show
                .select_slide(selector)
                .map_err(map_selector_error)?
                .ok_or(SlideBuildOrderError::SlideNameNotFound)?;
            Ok(Position::new(selected.index()))
        },
    }
}

fn read_build_order(
    source: &Package,
    slide_position: Position,
) -> Result<BuildOrderSnapshot, SlideBuildOrderError> {
    let slide_identifier = slide_identifier_at_position(source, slide_position)?;
    let (component_name, slide_object) = source
        .object_with_component(slide_identifier)
        .ok_or(SlideBuildOrderError::InvalidSource)?;
    let slide_payload = unique_payload(
        &slide_object.messages,
        &[SLIDE_MESSAGE_TYPE],
        "Keynote slide",
    )
    .map_err(map_read_error)?;
    let limits = source.semantic_wire_limits().map_err(map_read_error)?;
    let (build_references, chunk_references, owned_drawables) = parse_slide_order_payload(
        slide_payload,
        limits,
        source.semantic_limits().max_references(),
    )?;

    let mut build_ids = HashSet::new();
    build_ids.try_reserve(build_references.len()).map_err(|_| {
        SlideBuildOrderError::Allocation {
            amount: build_references.len(),
        }
    })?;
    let mut builds = Vec::new();
    builds
        .try_reserve_exact(build_references.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: build_references.len(),
        })?;
    for reference in build_references {
        if !build_ids.insert(reference.identifier) {
            return Err(SlideBuildOrderError::UnsupportedTopology);
        }
        let payload = object_payload_in_component(
            source,
            component_name,
            reference.identifier,
            BUILD_MESSAGE_TYPE,
        )?;
        let default_automatic = parse_build_payload(payload, limits, &owned_drawables)?;
        builds.push(BuildFact {
            identifier: reference.identifier,
            reference,
            default_automatic,
        });
    }

    let mut chunk_ids = HashSet::new();
    chunk_ids.try_reserve(chunk_references.len()).map_err(|_| {
        SlideBuildOrderError::Allocation {
            amount: chunk_references.len(),
        }
    })?;
    let mut chunks = Vec::new();
    chunks
        .try_reserve_exact(chunk_references.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: chunk_references.len(),
        })?;
    for reference in chunk_references {
        if !chunk_ids.insert(reference.identifier) {
            return Err(SlideBuildOrderError::UnsupportedTopology);
        }
        let payload = object_payload_in_component(
            source,
            component_name,
            reference.identifier,
            BUILD_CHUNK_MESSAGE_TYPE,
        )?;
        let (build_identifier, automatic, referent) = parse_chunk_payload(payload, limits)?;
        if !build_ids.contains(&build_identifier) {
            return Err(SlideBuildOrderError::UnsupportedTopology);
        }
        chunks.push(ChunkFact {
            identifier: reference.identifier,
            reference,
            build_identifier,
            automatic,
            referent,
        });
    }

    let mut build_indexes = HashMap::new();
    build_indexes
        .try_reserve(builds.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: builds.len(),
        })?;
    for (index, build) in builds.iter().enumerate() {
        build_indexes.insert(build.identifier, index);
    }
    let mut chunk_indexes_by_build = Vec::new();
    chunk_indexes_by_build
        .try_reserve_exact(builds.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: builds.len(),
        })?;
    for _ in &builds {
        chunk_indexes_by_build.push(Vec::new());
    }
    validate_chunk_grouping(
        &builds,
        &chunks,
        &build_indexes,
        &mut chunk_indexes_by_build,
    )?;
    validate_start_semantics(&builds, &chunks, &chunk_indexes_by_build)?;
    ensure_reference_limit(
        source,
        builds
            .len()
            .checked_add(chunks.len())
            .and_then(|value| value.checked_add(owned_drawables.len()))
            .ok_or(SlideBuildOrderError::InvalidSource)?,
    )?;

    Ok(BuildOrderSnapshot {
        slide_identifier,
        component_name: try_owned_string(component_name)?,
        builds,
        chunks,
        chunk_indexes_by_build,
    })
}

fn slide_identifier_at_position(
    source: &Package,
    position: Position,
) -> Result<u64, SlideBuildOrderError> {
    let show_identifier = source.root_show_identifier().map_err(map_read_error)?;
    if show_identifier == 0 {
        return Err(SlideBuildOrderError::InvalidSource);
    }
    let show_object = source
        .required_object(show_identifier, "Keynote show")
        .map_err(map_read_error)?;
    let show_payload = unique_payload(&show_object.messages, &[SHOW_MESSAGE_TYPE], "Keynote show")
        .map_err(map_read_error)?;
    let limits = source.semantic_wire_limits().map_err(map_read_error)?;
    let snapshot =
        decode_show_snapshot(show_payload, source.semantic_limits().max_slides(), limits)
            .map_err(map_read_error)?;
    let node_identifier = snapshot
        .slide_node_identifiers()
        .get(position.get())
        .copied()
        .ok_or(SlideBuildOrderError::SlidePositionNotFound { position })?;
    let node_object = source
        .required_object(node_identifier, "Keynote slide node")
        .map_err(map_read_error)?;
    let node_payload = unique_payload(
        &node_object.messages,
        &[SLIDE_NODE_MESSAGE_TYPE],
        "Keynote slide node",
    )
    .map_err(map_read_error)?;
    let (slide_identifier, _is_skipped) = decode_slide_node_projection(
        node_payload,
        limits,
        super::SemanticPath::Slide {
            index: position.get(),
        },
    )
    .map_err(map_read_error)?;
    if slide_identifier == 0 {
        return Err(SlideBuildOrderError::UnsupportedTopology);
    }
    Ok(slide_identifier)
}

fn parse_slide_order_payload(
    payload: &[u8],
    limits: WireLimits,
    max_references: usize,
) -> Result<(Vec<RawReference>, Vec<RawReference>, HashSet<u64>), SlideBuildOrderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut build_count = 0usize;
    let mut chunk_count = 0usize;
    let mut owned_count = 0usize;
    for field in view.fields() {
        let count = match field.number() {
            BUILDS_FIELD => &mut build_count,
            DEPRECATED_BUILD_CHUNKS_FIELD => {
                return Err(SlideBuildOrderError::UnsupportedTopology);
            },
            BUILD_CHUNKS_FIELD => &mut chunk_count,
            OWNED_DRAWABLES_FIELD => &mut owned_count,
            _ => continue,
        };
        *count = count
            .checked_add(1)
            .ok_or(SlideBuildOrderError::InvalidSource)?;
    }
    for (kind, observed) in [
        (SlideBuildOrderLimitKind::Builds, build_count),
        (SlideBuildOrderLimitKind::BuildChunks, chunk_count),
    ] {
        if observed > max_references {
            return Err(SlideBuildOrderError::LimitExceeded {
                kind,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(max_references),
            });
        }
    }
    let reference_count = build_count
        .checked_add(chunk_count)
        .and_then(|count| count.checked_add(owned_count))
        .ok_or(SlideBuildOrderError::InvalidSource)?;
    if reference_count > max_references {
        return Err(SlideBuildOrderError::LimitExceeded {
            kind: SlideBuildOrderLimitKind::References,
            observed: usize_to_u64(reference_count),
            maximum: usize_to_u64(max_references),
        });
    }
    let mut builds = Vec::new();
    builds
        .try_reserve_exact(build_count)
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: build_count,
        })?;
    let mut chunks = Vec::new();
    chunks
        .try_reserve_exact(chunk_count)
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: chunk_count,
        })?;
    let mut owned = HashSet::new();
    owned
        .try_reserve(owned_count)
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: owned_count,
        })?;
    for field in view.fields() {
        match field.number() {
            BUILDS_FIELD => builds.push(parse_reference_field(field, limits)?),
            DEPRECATED_BUILD_CHUNKS_FIELD => {
                return Err(SlideBuildOrderError::UnsupportedTopology);
            },
            BUILD_CHUNKS_FIELD => chunks.push(parse_reference_field(field, limits)?),
            OWNED_DRAWABLES_FIELD => {
                let identifier = parse_reference_field(field, limits)?.identifier;
                if !owned.insert(identifier) {
                    return Err(SlideBuildOrderError::UnsupportedTopology);
                }
            },
            _ => {},
        }
    }
    Ok((builds, chunks, owned))
}

fn parse_reference_field(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    limits: WireLimits,
) -> Result<RawReference, SlideBuildOrderError> {
    if field.wire_type() != 2 {
        return Err(SlideBuildOrderError::InvalidSource);
    }
    field.validate_canonical_framing().map_err(map_wire_error)?;
    let identifier = parse_reference_payload(field.payload(), limits)?;
    Ok(RawReference {
        identifier,
        raw: try_owned_bytes(field.raw())?,
    })
}

fn parse_reference_payload(
    payload: &[u8],
    limits: WireLimits,
) -> Result<u64, SlideBuildOrderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    let mut deprecated_type = None;
    let mut external = None;
    for field in view.fields() {
        match field.number() {
            UUID_LOWER_FIELD => {
                field.validate_canonical_key().map_err(map_wire_error)?;
                let value = canonical_varint(field.payload())?;
                if identifier.replace(value).is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
            },
            2 => {
                field.validate_canonical_key().map_err(map_wire_error)?;
                let value = canonical_varint(field.payload())?;
                if value > i32::MAX as u64 && value < 0xffff_ffff_8000_0000 {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                if deprecated_type.replace(value).is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
            },
            3 => {
                field.validate_canonical_key().map_err(map_wire_error)?;
                let value = canonical_varint(field.payload())?;
                if value > 1 || external.replace(value == 1).is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
            },
            _ => {},
        }
    }
    let identifier = identifier
        .filter(|identifier| *identifier != 0)
        .ok_or(SlideBuildOrderError::InvalidSource)?;
    if external == Some(true) {
        return Err(SlideBuildOrderError::UnsupportedTopology);
    }
    Ok(identifier)
}

fn parse_build_payload(
    payload: &[u8],
    limits: WireLimits,
    owned_drawables: &HashSet<u64>,
) -> Result<Option<bool>, SlideBuildOrderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut drawable = None;
    let mut delivery = false;
    let mut attributes = None;
    for field in view.fields() {
        match field.number() {
            DRAWABLE_FIELD => {
                if drawable.is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                let reference = parse_reference_payload_field(field, limits)?;
                drawable = Some(reference);
            },
            DELIVERY_FIELD => {
                if delivery || field.wire_type() != 2 {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                field.validate_canonical_framing().map_err(map_wire_error)?;
                str::from_utf8(field.payload()).map_err(|_| SlideBuildOrderError::InvalidSource)?;
                delivery = true;
            },
            ATTRIBUTES_FIELD => {
                if attributes.is_some() || field.wire_type() != 2 {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                field.validate_canonical_framing().map_err(map_wire_error)?;
                attributes = Some(field.payload());
            },
            _ => {},
        }
    }
    let drawable = drawable.ok_or(SlideBuildOrderError::UnsupportedTopology)?;
    if !owned_drawables.contains(&drawable) || !delivery || attributes.is_none() {
        return Err(SlideBuildOrderError::UnsupportedTopology);
    }
    parse_animation_automatic(attributes.unwrap_or_default(), limits)
}

fn parse_reference_payload_field(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
    limits: WireLimits,
) -> Result<u64, SlideBuildOrderError> {
    if field.wire_type() != 2 {
        return Err(SlideBuildOrderError::InvalidSource);
    }
    field.validate_canonical_framing().map_err(map_wire_error)?;
    parse_reference_payload(field.payload(), limits)
}

fn parse_animation_automatic(
    payload: &[u8],
    limits: WireLimits,
) -> Result<Option<bool>, SlideBuildOrderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut animation = None;
    for field in view.fields() {
        if field.number() != ATTRIBUTES_ANIMATION_FIELD {
            continue;
        }
        if animation.is_some() || field.wire_type() != 2 {
            return Err(SlideBuildOrderError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        animation = Some(field.payload());
    }
    let Some(animation) = animation else {
        return Ok(None);
    };
    let animation_view = WireView::parse_with_limits(animation, limits).map_err(map_wire_error)?;
    let mut automatic = None;
    for field in animation_view.fields() {
        if field.number() != ANIMATION_AUTOMATIC_FIELD {
            continue;
        }
        if automatic.is_some() {
            return Err(SlideBuildOrderError::InvalidSource);
        }
        automatic = Some(parse_bool_field(field)?);
    }
    Ok(automatic)
}

fn parse_chunk_payload(
    payload: &[u8],
    limits: WireLimits,
) -> Result<(u64, Option<bool>, Option<bool>), SlideBuildOrderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut build = None;
    let mut automatic = None;
    let mut referent = None;
    let mut chunk_identifier = None;
    let mut build_id = None;
    for field in view.fields() {
        match field.number() {
            CHUNK_BUILD_FIELD => {
                if build.is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                build = Some(parse_reference_payload_field(field, limits)?);
            },
            CHUNK_AUTOMATIC_FIELD => {
                if automatic.is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                automatic = Some(parse_bool_field(field)?);
            },
            CHUNK_REFERENT_FIELD => {
                if referent.is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                referent = Some(parse_bool_field(field)?);
            },
            CHUNK_IDENTIFIER_FIELD => {
                if chunk_identifier.is_some() || field.wire_type() != 2 {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                field.validate_canonical_framing().map_err(map_wire_error)?;
                chunk_identifier = Some(parse_chunk_identifier(field.payload(), limits)?);
            },
            CHUNK_BUILD_ID_FIELD => {
                if build_id.is_some() || field.wire_type() != 2 {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                field.validate_canonical_framing().map_err(map_wire_error)?;
                build_id = Some(parse_uuid(field.payload(), limits)?);
            },
            _ => {},
        }
    }
    if let (Some(left), Some(right)) = (chunk_identifier, build_id) {
        if left != right {
            return Err(SlideBuildOrderError::UnsupportedTopology);
        }
    }
    let build = build.ok_or(SlideBuildOrderError::UnsupportedTopology)?;
    Ok((build, automatic, referent))
}

fn parse_chunk_identifier(
    payload: &[u8],
    limits: WireLimits,
) -> Result<Uuid, SlideBuildOrderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut uuid = None;
    let mut chunk_id = None;
    for field in view.fields() {
        match field.number() {
            1 => {
                if uuid.is_some() || field.wire_type() != 2 {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                field.validate_canonical_framing().map_err(map_wire_error)?;
                uuid = Some(parse_uuid(field.payload(), limits)?);
            },
            2 => {
                if chunk_id.is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                chunk_id = Some(canonical_int32(field)?);
            },
            _ => {},
        }
    }
    uuid.ok_or(SlideBuildOrderError::UnsupportedTopology)
}

fn parse_uuid(payload: &[u8], limits: WireLimits) -> Result<Uuid, SlideBuildOrderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut lower = None;
    let mut upper = None;
    for field in view.fields() {
        match field.number() {
            UUID_LOWER_FIELD => {
                if lower.is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                lower = Some(canonical_uint64(field)?);
            },
            UUID_UPPER_FIELD => {
                if upper.is_some() {
                    return Err(SlideBuildOrderError::InvalidSource);
                }
                upper = Some(canonical_uint64(field)?);
            },
            _ => {},
        }
    }
    Ok(Uuid {
        lower: lower.ok_or(SlideBuildOrderError::UnsupportedTopology)?,
        upper: upper.ok_or(SlideBuildOrderError::UnsupportedTopology)?,
    })
}

fn parse_bool_field(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> Result<bool, SlideBuildOrderError> {
    if field.wire_type() != 0 {
        return Err(SlideBuildOrderError::InvalidSource);
    }
    field.validate_canonical_key().map_err(map_wire_error)?;
    match canonical_varint(field.payload())? {
        0 => Ok(false),
        1 => Ok(true),
        _ => Err(SlideBuildOrderError::InvalidSource),
    }
}

fn canonical_uint64(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> Result<u64, SlideBuildOrderError> {
    if field.wire_type() != 0 {
        return Err(SlideBuildOrderError::InvalidSource);
    }
    field.validate_canonical_key().map_err(map_wire_error)?;
    canonical_varint(field.payload())
}

fn canonical_int32(
    field: litchi_iwa_common::wire::WireFieldView<'_>,
) -> Result<i32, SlideBuildOrderError> {
    let value = canonical_uint64(field)?;
    if value > i32::MAX as u64 && value < 0xffff_ffff_8000_0000 {
        return Err(SlideBuildOrderError::InvalidSource);
    }
    Ok(value as i32)
}

fn canonical_varint(payload: &[u8]) -> Result<u64, SlideBuildOrderError> {
    let (value, consumed) = litchi_iwa_common::decode_varint_from_bytes(payload)
        .map_err(|_| SlideBuildOrderError::InvalidSource)?;
    if consumed != payload.len() || consumed != litchi_iwa_common::varint::encoded_len(value) {
        return Err(SlideBuildOrderError::InvalidSource);
    }
    Ok(value)
}

fn validate_chunk_grouping(
    builds: &[BuildFact],
    chunks: &[ChunkFact],
    build_indexes: &HashMap<u64, usize>,
    chunk_indexes_by_build: &mut [Vec<usize>],
) -> Result<(), SlideBuildOrderError> {
    let mut chunk_counts = Vec::new();
    chunk_counts
        .try_reserve_exact(builds.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: builds.len(),
        })?;
    chunk_counts.resize(builds.len(), 0usize);
    for chunk in chunks {
        let build_index = *build_indexes
            .get(&chunk.build_identifier)
            .ok_or(SlideBuildOrderError::UnsupportedTopology)?;
        chunk_counts[build_index] = chunk_counts[build_index]
            .checked_add(1)
            .ok_or(SlideBuildOrderError::InvalidSource)?;
    }
    for (indexes, count) in chunk_indexes_by_build.iter_mut().zip(chunk_counts) {
        indexes
            .try_reserve_exact(count)
            .map_err(|_| SlideBuildOrderError::Allocation { amount: count })?;
    }
    let mut closed = HashSet::new();
    closed
        .try_reserve(builds.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: builds.len(),
        })?;
    let mut active = None;
    let mut group_order = Vec::new();
    group_order
        .try_reserve_exact(builds.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: builds.len(),
        })?;
    for (chunk_index, chunk) in chunks.iter().enumerate() {
        let build_index = *build_indexes
            .get(&chunk.build_identifier)
            .ok_or(SlideBuildOrderError::UnsupportedTopology)?;
        chunk_indexes_by_build[build_index].push(chunk_index);
        if active != Some(build_index) {
            if closed.contains(&build_index) {
                return Err(SlideBuildOrderError::UnsupportedTopology);
            }
            if let Some(previous) = active.replace(build_index) {
                closed.insert(previous);
            }
            group_order.push(build_index);
        }
    }
    let mut expected_group_order = Vec::new();
    expected_group_order
        .try_reserve_exact(builds.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: builds.len(),
        })?;
    for (index, chunk_indexes) in chunk_indexes_by_build.iter().enumerate() {
        if !chunk_indexes.is_empty() {
            expected_group_order.push(index);
        }
    }
    if group_order != expected_group_order {
        return Err(SlideBuildOrderError::UnsupportedTopology);
    }
    Ok(())
}

fn validate_start_semantics(
    builds: &[BuildFact],
    chunks: &[ChunkFact],
    chunk_indexes_by_build: &[Vec<usize>],
) -> Result<(), SlideBuildOrderError> {
    let mut event_index = 0usize;
    for (build_index, build) in builds.iter().enumerate() {
        let Some(first_chunk_index) = chunk_indexes_by_build[build_index].first().copied() else {
            continue;
        };
        let chunk = chunks
            .get(first_chunk_index)
            .ok_or(SlideBuildOrderError::InvalidSource)?;
        let automatic = chunk.automatic.or(build.default_automatic).unwrap_or(false);
        let referent = chunk.referent.unwrap_or(true);
        // Automatic/referent on the first event is native After Transition;
        // the same shape later in the order is native After Previous.
        if automatic && !referent && event_index == 0 {
            return Err(SlideBuildOrderError::UnsupportedTopology);
        }
        event_index = event_index
            .checked_add(chunk_indexes_by_build[build_index].len())
            .ok_or(SlideBuildOrderError::InvalidSource)?;
    }
    Ok(())
}

fn try_owned_bytes(source: &[u8]) -> Result<Box<[u8]>, SlideBuildOrderError> {
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(source.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: source.len(),
        })?;
    bytes.extend_from_slice(source);
    Ok(bytes.into_boxed_slice())
}

fn try_owned_string(source: &str) -> Result<String, SlideBuildOrderError> {
    let mut value = String::new();
    value
        .try_reserve_exact(source.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: source.len(),
        })?;
    value.push_str(source);
    Ok(value)
}

fn try_clone_reference(source: &RawReference) -> Result<RawReference, SlideBuildOrderError> {
    Ok(RawReference {
        identifier: source.identifier,
        raw: try_owned_bytes(&source.raw)?,
    })
}

fn try_clone_build_fact(source: &BuildFact) -> Result<BuildFact, SlideBuildOrderError> {
    Ok(BuildFact {
        identifier: source.identifier,
        reference: try_clone_reference(&source.reference)?,
        default_automatic: source.default_automatic,
    })
}

fn try_clone_chunk_fact(source: &ChunkFact) -> Result<ChunkFact, SlideBuildOrderError> {
    Ok(ChunkFact {
        identifier: source.identifier,
        reference: try_clone_reference(&source.reference)?,
        build_identifier: source.build_identifier,
        automatic: source.automatic,
        referent: source.referent,
    })
}

fn object_payload_in_component<'a>(
    source: &'a Package,
    expected_component: &str,
    identifier: u64,
    type_: u32,
) -> Result<&'a [u8], SlideBuildOrderError> {
    let (component, object) = source
        .object_with_component(identifier)
        .ok_or(SlideBuildOrderError::InvalidSource)?;
    if component != expected_component {
        return Err(SlideBuildOrderError::UnsupportedTopology);
    }
    unique_payload(&object.messages, &[type_], "Keynote build object").map_err(map_read_error)
}

fn validate_build_order_snapshot(
    snapshot: &BuildOrderSnapshot,
) -> Result<(), SlideBuildOrderError> {
    if snapshot.builds.len() != snapshot.chunk_indexes_by_build.len()
        || snapshot
            .builds
            .iter()
            .any(|build| build.reference.identifier != build.identifier)
        || snapshot
            .chunks
            .iter()
            .any(|chunk| chunk.reference.identifier != chunk.identifier)
    {
        return Err(SlideBuildOrderError::InvalidSource);
    }
    Ok(())
}

fn ensure_intent_bounds(
    snapshot: &BuildOrderSnapshot,
    intent: Intent,
) -> Result<(), SlideBuildOrderError> {
    validate_build_order_snapshot(snapshot)?;
    if intent.source.get() >= snapshot.builds.len() {
        return Err(SlideBuildOrderError::BuildPositionNotFound {
            position: intent.source,
        });
    }
    if intent.destination.get() >= snapshot.builds.len() {
        return Err(SlideBuildOrderError::DestinationOutOfRange {
            position: intent.destination,
            build_count: snapshot.builds.len(),
        });
    }
    Ok(())
}

fn moved_build_order(
    before: &BuildOrderSnapshot,
    intent: Intent,
) -> Result<BuildOrderSnapshot, SlideBuildOrderError> {
    let mut order = Vec::new();
    order
        .try_reserve_exact(before.builds.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: before.builds.len(),
        })?;
    order.extend(0..before.builds.len());
    let selected = order.remove(intent.source.get());
    order.insert(intent.destination.get(), selected);

    let mut builds = Vec::new();
    builds.try_reserve_exact(before.builds.len()).map_err(|_| {
        SlideBuildOrderError::Allocation {
            amount: before.builds.len(),
        }
    })?;
    for index in &order {
        builds.push(try_clone_build_fact(&before.builds[*index])?);
    }
    let mut chunks = Vec::new();
    chunks.try_reserve_exact(before.chunks.len()).map_err(|_| {
        SlideBuildOrderError::Allocation {
            amount: before.chunks.len(),
        }
    })?;
    for index in &order {
        for chunk_index in &before.chunk_indexes_by_build[*index] {
            chunks.push(try_clone_chunk_fact(&before.chunks[*chunk_index])?);
        }
    }
    let mut chunk_indexes_by_build = Vec::new();
    chunk_indexes_by_build
        .try_reserve_exact(builds.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: builds.len(),
        })?;
    let mut chunk_offset = 0usize;
    for index in &order {
        let mut indexes = Vec::new();
        let chunk_count = before.chunk_indexes_by_build[*index].len();
        indexes
            .try_reserve_exact(chunk_count)
            .map_err(|_| SlideBuildOrderError::Allocation {
                amount: chunk_count,
            })?;
        let chunk_end = chunk_offset
            .checked_add(chunk_count)
            .ok_or(SlideBuildOrderError::InvalidSource)?;
        indexes.extend(chunk_offset..chunk_end);
        chunk_offset = chunk_end;
        chunk_indexes_by_build.push(indexes);
    }
    let moved = BuildOrderSnapshot {
        slide_identifier: before.slide_identifier,
        component_name: try_owned_string(&before.component_name)?,
        builds,
        chunks,
        chunk_indexes_by_build,
    };
    validate_start_semantics(&moved.builds, &moved.chunks, &moved.chunk_indexes_by_build)?;
    Ok(moved)
}

fn same_order(left: &BuildOrderSnapshot, right: &BuildOrderSnapshot) -> bool {
    left.slide_identifier == right.slide_identifier
        && left
            .builds
            .iter()
            .map(|build| build.identifier)
            .eq(right.builds.iter().map(|build| build.identifier))
        && left
            .chunks
            .iter()
            .map(|chunk| chunk.identifier)
            .eq(right.chunks.iter().map(|chunk| chunk.identifier))
}

fn rewrite_build_order(
    source: &Package,
    before: &BuildOrderSnapshot,
    intent: Intent,
    after: &BuildOrderSnapshot,
) -> Result<Package, SlideBuildOrderError> {
    let expected = moved_build_order(before, intent)?;
    if !same_order(&expected, after) {
        return Err(SlideBuildOrderError::Verification);
    }
    let source_catalog = editable_source_catalog(source)?;
    let component = source_catalog
        .components()
        .iter()
        .find(|component| component.name() == before.component_name.as_str())
        .ok_or(SlideBuildOrderError::InvalidSource)?;
    let entry = source_catalog
        .package()
        .iter()
        .find(|entry| entry.name() == before.component_name.as_str())
        .ok_or(SlideBuildOrderError::InvalidSource)?;
    if entry.is_opaque()
        || component
            .archive()
            .object(before.slide_identifier)
            .is_none()
    {
        return Err(SlideBuildOrderError::InvalidSource);
    }
    let archive_limits = source_catalog
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        source_catalog
            .limits()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    let stream_bytes = stream.as_bytes();
    let mut archive =
        Archive::parse_with_limits(stream_bytes, archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream_bytes)
        .map_err(map_core_error)?;
    let object = archive
        .object_mut(before.slide_identifier)
        .ok_or(SlideBuildOrderError::InvalidSource)?;
    let mut message_index = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ == SLIDE_MESSAGE_TYPE && message_index.replace(index).is_some() {
            return Err(SlideBuildOrderError::InvalidSource);
        }
    }
    let message_index = message_index.ok_or(SlideBuildOrderError::InvalidSource)?;
    let original = object.messages[message_index].data.as_slice();
    let rewritten = permute_slide_build_references(
        original,
        before,
        after,
        source.wire_limits().map_err(map_wire_error)?,
    )?;
    let replacement = RawMessage {
        type_: SLIDE_MESSAGE_TYPE,
        data: rewritten,
    };
    object
        .replace_message_preserving_header_with_limits(message_index, replacement, archive_limits)
        .map_err(map_core_error)?;
    let serialized = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&serialized).map_err(map_core_error)?;
    let output = source_catalog
        .package()
        .reassemble_to_bytes(
            &[EntryEdit::new(before.component_name.as_str(), &compressed)],
            source_catalog.limits(),
        )
        .map_err(map_archive_error)?;
    Package::from_source_with_options(output.into(), source.state.options).map_err(map_read_error)
}

fn permute_slide_build_references(
    payload: &[u8],
    before: &BuildOrderSnapshot,
    after: &BuildOrderSnapshot,
    limits: WireLimits,
) -> Result<Vec<u8>, SlideBuildOrderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut build_raw_by_id = HashMap::new();
    build_raw_by_id
        .try_reserve(before.builds.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: before.builds.len(),
        })?;
    for build in &before.builds {
        if build_raw_by_id
            .insert(build.identifier, build.reference.raw.as_ref())
            .is_some()
        {
            return Err(SlideBuildOrderError::UnsupportedTopology);
        }
    }
    let mut chunk_raw_by_id = HashMap::new();
    chunk_raw_by_id
        .try_reserve(before.chunks.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: before.chunks.len(),
        })?;
    for chunk in &before.chunks {
        if chunk_raw_by_id
            .insert(chunk.identifier, chunk.reference.raw.as_ref())
            .is_some()
        {
            return Err(SlideBuildOrderError::UnsupportedTopology);
        }
    }
    let mut requested_builds = after.builds.iter();
    let mut requested_chunks = after.chunks.iter();
    let mut output = Vec::new();
    output
        .try_reserve_exact(payload.len())
        .map_err(|_| SlideBuildOrderError::Allocation {
            amount: payload.len(),
        })?;
    for field in view.fields() {
        match field.number() {
            BUILDS_FIELD => {
                let identifier = requested_builds
                    .next()
                    .map(|build| &build.identifier)
                    .ok_or(SlideBuildOrderError::InvalidSource)?;
                output.extend_from_slice(
                    build_raw_by_id
                        .get(identifier)
                        .copied()
                        .ok_or(SlideBuildOrderError::InvalidSource)?,
                );
            },
            BUILD_CHUNKS_FIELD => {
                let identifier = requested_chunks
                    .next()
                    .map(|chunk| &chunk.identifier)
                    .ok_or(SlideBuildOrderError::InvalidSource)?;
                output.extend_from_slice(
                    chunk_raw_by_id
                        .get(identifier)
                        .copied()
                        .ok_or(SlideBuildOrderError::InvalidSource)?,
                );
            },
            _ => output.extend_from_slice(field.raw()),
        }
    }
    if requested_builds.next().is_some() || requested_chunks.next().is_some() {
        return Err(SlideBuildOrderError::Verification);
    }
    if output.len() > limits.max_output_bytes() {
        return Err(SlideBuildOrderError::LimitExceeded {
            kind: SlideBuildOrderLimitKind::OutputBytes,
            observed: usize_to_u64(output.len()),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    Ok(output)
}

fn ensure_reference_limit(source: &Package, count: usize) -> Result<(), SlideBuildOrderError> {
    let maximum = source.semantic_limits().max_references();
    if count > maximum {
        return Err(SlideBuildOrderError::LimitExceeded {
            kind: SlideBuildOrderLimitKind::References,
            observed: usize_to_u64(count),
            maximum: usize_to_u64(maximum),
        });
    }
    Ok(())
}

fn physical_shared_source(source: &Package) -> Result<Arc<[u8]>, SlideBuildOrderError> {
    match &source.state.source {
        PhysicalSource::Package(package) => Ok(package.shared_source()),
        PhysicalSource::Semantic(_) => Err(SlideBuildOrderError::UnsupportedSource),
    }
}

fn editable_shared_source(source: &Package) -> Result<Arc<[u8]>, SlideBuildOrderError> {
    Ok(editable_source_catalog(source)?.shared_source())
}

fn editable_source_catalog(
    source: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideBuildOrderError> {
    let package = match &source.state.source {
        PhysicalSource::Package(package) => package,
        PhysicalSource::Semantic(_) => return Err(SlideBuildOrderError::UnsupportedSource),
    };
    if !package.source_is_exact() {
        return Err(SlideBuildOrderError::UnsupportedSource);
    }
    Ok(package)
}

fn map_selector_error(error: SlideSelectorError) -> SlideBuildOrderError {
    match error {
        SlideSelectorError::EmptySlideName => SlideBuildOrderError::EmptySlideName,
        SlideSelectorError::DuplicateSlideName { .. } => SlideBuildOrderError::AmbiguousSelector,
    }
}

fn map_read_error(error: ReadError) -> SlideBuildOrderError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideBuildOrderError::LimitExceeded {
            kind: match kind {
                super::SemanticLimitKind::Objects => SlideBuildOrderLimitKind::Entries,
                super::SemanticLimitKind::Slides => SlideBuildOrderLimitKind::Slides,
                super::SemanticLimitKind::References => SlideBuildOrderLimitKind::References,
                super::SemanticLimitKind::TextStorages
                | super::SemanticLimitKind::TextFragments
                | super::SemanticLimitKind::TextBytes => SlideBuildOrderLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideBuildOrderError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => SlideBuildOrderLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => SlideBuildOrderLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => SlideBuildOrderLimitKind::WireNesting,
                super::PayloadLimitKind::Work => SlideBuildOrderLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => SlideBuildOrderError::Allocation { amount },
        ReadError::Archive(archive_error) => map_archive_error(archive_error),
        ReadError::Io(_)
        | ReadError::Detection(_)
        | ReadError::NotKeynote
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_) => SlideBuildOrderError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideBuildOrderError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideBuildOrderError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideBuildOrderLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => SlideBuildOrderLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => SlideBuildOrderLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    SlideBuildOrderLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    SlideBuildOrderLimitKind::TotalBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideBuildOrderError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(core_error) => map_core_error(core_error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => SlideBuildOrderError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideBuildOrderError {
    match error {
        litchi_iwa_core::Error::Limit {
            observed, maximum, ..
        } => SlideBuildOrderError::LimitExceeded {
            kind: SlideBuildOrderLimitKind::EntryBytes,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideBuildOrderError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => SlideBuildOrderError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SlideBuildOrderError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SlideBuildOrderError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => SlideBuildOrderLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => SlideBuildOrderLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    SlideBuildOrderLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => SlideBuildOrderLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => SlideBuildOrderLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideBuildOrderError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => SlideBuildOrderError::InvalidSource,
    }
}

fn fingerprint(bytes: &[u8]) -> u64 {
    // A compact diagnostic value only; exact source bytes retained by the
    // patch remain the authorization boundary.
    let mut value = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        value ^= u64::from(*byte);
        value = value.wrapping_mul(0x1000_0000_01b3);
    }
    value
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::io;

    use litchi_iwa_archive::package::{self, Catalog};
    use litchi_iwa_common::wire::{WireView, append_varint_field};
    use litchi_iwa_core::{ArchiveObject, SnappyStream};
    use litchi_iwa_protos::{kn, tsa, tsk, tsp};
    use prost::Message as _;

    const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
    const SLIDE_MEMBER: &str = "Index/Slide-4.iwa";
    const BUILDS: [u64; 2] = [20, 21];
    const CHUNKS: [u64; 2] = [30, 31];
    const DRAWABLES: [u64; 2] = [50, 51];

    type TestResult<T> = Result<T, Box<dyn std::error::Error>>;

    #[derive(Debug, Clone, Copy)]
    struct FixtureOptions {
        unknown_reference_fields: bool,
        interleaved_chunks: bool,
        first_chunk_with_previous: bool,
        deprecated_chunk_list: bool,
    }

    impl FixtureOptions {
        const VALID: Self = Self {
            unknown_reference_fields: true,
            interleaved_chunks: false,
            first_chunk_with_previous: false,
            deprecated_chunk_list: false,
        };
    }

    fn reference(identifier: u64) -> tsp::Reference {
        tsp::Reference {
            identifier,
            deprecated_type: Some(-1),
            deprecated_is_external: Some(false),
        }
    }

    fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
        Ok(ArchiveObject::new(
            identifier,
            vec![RawMessage { type_, data }],
        )?)
    }

    fn component(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
        Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
    }

    #[allow(
        clippy::cast_possible_truncation,
        reason = "each emitted byte intentionally retains only the low seven varint bits"
    )]
    fn push_varint(mut value: u64, output: &mut Vec<u8>) {
        while value >= 0x80 {
            output.push(((value as u8) & 0x7f) | 0x80);
            value >>= 7;
        }
        output.push(value as u8);
    }

    fn length_delimited_field(number: u32, payload: &[u8]) -> Vec<u8> {
        let mut field = Vec::with_capacity(payload.len().saturating_add(8));
        push_varint((u64::from(number) << 3) | 2, &mut field);
        push_varint(payload.len() as u64, &mut field);
        field.extend_from_slice(payload);
        field
    }

    fn reference_payload(identifier: u64, sentinel: Option<u64>) -> Vec<u8> {
        let mut payload = reference(identifier).encode_to_vec();
        if let Some(sentinel) = sentinel {
            append_varint_field(&mut payload, 99, sentinel).expect("fixture varint fits");
        }
        payload
    }

    fn uuid(index: u64) -> tsp::Uuid {
        tsp::Uuid {
            lower: 0x1000_u64.saturating_add(index),
            upper: 0x2000_u64.saturating_add(index),
        }
    }

    fn build_payload(drawable: u64, effect: &str) -> Vec<u8> {
        kn::BuildArchive {
            drawable: Some(reference(drawable)),
            delivery: effect.to_owned(),
            attributes: kn::BuildAttributesArchive {
                animation_attributes: Some(kn::AnimationAttributesArchive {
                    animation_type: Some("In".to_owned()),
                    effect: Some(effect.to_owned()),
                    duration: Some(0.5),
                    ..Default::default()
                }),
                ..Default::default()
            },
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn chunk_payload(
        build: u64,
        chunk_index: u64,
        automatic: Option<bool>,
        referent: Option<bool>,
    ) -> Vec<u8> {
        let identifier = uuid(chunk_index);
        kn::BuildChunkArchive {
            build: Some(reference(build)),
            delay: Some(0.1),
            duration: Some(0.5),
            automatic,
            referent,
            build_chunk_identifier: Some(kn::BuildChunkIdentifierArchive {
                build_id: Some(identifier),
                build_chunk_id: Some(i32::try_from(chunk_index).expect("fixture index fits")),
            }),
            build_id: Some(identifier),
            ..Default::default()
        }
        .encode_to_vec()
    }

    fn slide_payload(options: FixtureOptions) -> Vec<u8> {
        let mut chunk_ids = CHUNKS.to_vec();
        if options.interleaved_chunks {
            chunk_ids.push(32);
        }
        let canonical = kn::SlideArchive {
            style: reference(90),
            builds: BUILDS.iter().copied().map(reference).collect(),
            build_chunks: chunk_ids.iter().copied().map(reference).collect(),
            transition: kn::TransitionArchive {
                attributes: kn::TransitionAttributesArchive::default(),
            },
            owned_drawables: DRAWABLES.iter().copied().map(reference).collect(),
            name: Some("Build fixture".to_owned()),
            in_document: true,
            ..Default::default()
        }
        .encode_to_vec();
        let view = WireView::parse(&canonical).expect("canonical slide parses");
        let mut build_index = 0usize;
        let mut chunk_index = 0usize;
        let mut output = Vec::with_capacity(canonical.len().saturating_add(64));
        for field in view.fields() {
            match field.number() {
                BUILDS_FIELD => {
                    let identifier = BUILDS[build_index];
                    let sentinel = options
                        .unknown_reference_fields
                        .then_some(0xabc_u64.saturating_add(build_index as u64));
                    output.extend_from_slice(&length_delimited_field(
                        BUILDS_FIELD,
                        &reference_payload(identifier, sentinel),
                    ));
                    build_index += 1;
                },
                BUILD_CHUNKS_FIELD => {
                    let identifier = chunk_ids[chunk_index];
                    let sentinel = options
                        .unknown_reference_fields
                        .then_some(0xdef_u64.saturating_add(chunk_index as u64));
                    output.extend_from_slice(&length_delimited_field(
                        BUILD_CHUNKS_FIELD,
                        &reference_payload(identifier, sentinel),
                    ));
                    chunk_index += 1;
                },
                _ => output.extend_from_slice(field.raw()),
            }
        }
        if options.deprecated_chunk_list {
            output.extend_from_slice(&length_delimited_field(
                DEPRECATED_BUILD_CHUNKS_FIELD,
                &kn::BuildChunkArchive {
                    build: Some(reference(BUILDS[0])),
                    ..Default::default()
                }
                .encode_to_vec(),
            ));
        }
        output
    }

    #[allow(deprecated)]
    fn package_bytes(options: FixtureOptions) -> TestResult<Vec<u8>> {
        let document = kn::DocumentArchive {
            super_: tsa::DocumentArchive {
                super_: tsk::DocumentArchive::default(),
                ..Default::default()
            },
            show: reference(2),
            ..Default::default()
        };
        let show = kn::ShowArchive {
            theme: reference(80),
            slide_tree: kn::SlideTreeArchive {
                slides: vec![reference(3)],
                ..Default::default()
            },
            size: tsp::Size {
                width: 1_024.0,
                height: 768.0,
            },
            stylesheet: reference(81),
            mode: Some(-1),
            ..Default::default()
        };
        let node = kn::SlideNodeArchive {
            slide: Some(reference(4)),
            depth: Some(1),
            is_skipped: false,
            has_builds: true,
            has_transition: true,
            ..Default::default()
        };
        let document_component = component(vec![
            object(1, 1, document.encode_to_vec())?,
            object(2, SHOW_MESSAGE_TYPE, show.encode_to_vec())?,
            object(3, SLIDE_NODE_MESSAGE_TYPE, node.encode_to_vec())?,
        ])?;

        let mut slide_objects = vec![object(4, SLIDE_MESSAGE_TYPE, slide_payload(options))?];
        for (index, build) in BUILDS.into_iter().enumerate() {
            slide_objects.push(object(
                build,
                BUILD_MESSAGE_TYPE,
                build_payload(
                    DRAWABLES[index],
                    if index == 0 { "appear" } else { "dissolve" },
                ),
            )?);
            slide_objects.push(object(
                DRAWABLES[index],
                999,
                b"opaque drawable sentinel".to_vec(),
            )?);
        }
        let mut chunk_builds = vec![(CHUNKS[0], BUILDS[0])];
        chunk_builds.push((CHUNKS[1], BUILDS[1]));
        if options.interleaved_chunks {
            chunk_builds.push((32, BUILDS[0]));
        }
        for (index, (chunk, build)) in chunk_builds.into_iter().enumerate() {
            let automatic = options.first_chunk_with_previous && index == 0;
            let referent = !automatic;
            slide_objects.push(object(
                chunk,
                BUILD_CHUNK_MESSAGE_TYPE,
                chunk_payload(
                    build,
                    u64::try_from(index)
                        .expect("fixture index fits")
                        .saturating_add(1),
                    Some(automatic),
                    Some(referent),
                ),
            )?);
        }
        let slide_component = component(slide_objects)?;

        Ok(package::to_bytes(
            [
                ("Data/unrelated.bin", b"unrelated opaque bytes".as_slice()),
                (DOCUMENT_MEMBER, document_component.as_slice()),
                (SLIDE_MEMBER, slide_component.as_slice()),
            ],
            super::super::Limits::default(),
        )?)
    }

    fn exact_bytes(package: &Package) -> Vec<u8> {
        let mut bytes = Vec::new();
        package
            .write_to(&mut bytes)
            .expect("an in-memory Vec accepts every package byte");
        bytes
    }

    fn build_identifiers(package: &Package) -> TestResult<Vec<String>> {
        Ok(package
            .show()?
            .slides()
            .first()
            .ok_or_else(|| io::Error::other("fixture slide missing"))?
            .builds()
            .iter()
            .map(|build| build.animation_type().identifier().to_owned())
            .collect())
    }

    fn slide_reference_records(package_bytes: &[u8], number: u32) -> TestResult<Vec<Vec<u8>>> {
        let catalog = Catalog::from_bytes(package_bytes)?;
        let entry = catalog
            .iter()
            .find(|entry| entry.name() == SLIDE_MEMBER)
            .ok_or_else(|| io::Error::other("fixture slide member missing"))?;
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        let slide = archive
            .object(4)
            .and_then(|object| object.messages.iter().find(|message| message.type_ == 5))
            .ok_or_else(|| io::Error::other("fixture slide object missing"))?;
        Ok(WireView::parse(&slide.data)?
            .fields()
            .filter(|field| field.number() == number)
            .map(|field| field.raw().to_vec())
            .collect())
    }

    #[test]
    fn selector_first_move_reorders_builds_and_chunk_groups_losslessly() -> TestResult<()> {
        let bytes = package_bytes(FixtureOptions::VALID)?;
        let package = Package::from_bytes(&bytes)?;
        assert_eq!(build_identifiers(&package)?, ["appear", "dissolve"]);

        let before_builds = slide_reference_records(&bytes, BUILDS_FIELD)?;
        let before_chunks = slide_reference_records(&bytes, BUILD_CHUNKS_FIELD)?;
        let mut edit = package.edit_slide_build_order();
        edit.move_build("Build fixture", Position::new(0), Position::new(1))?;
        let commit = edit.commit()?;

        assert_eq!(build_identifiers(commit.package())?, ["dissolve", "appear"]);
        assert_eq!(
            read_build_order(commit.package(), Position::new(0))?.chunk_indexes_by_build,
            [vec![0], vec![1]]
        );
        assert!(commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 1);
        assert!(commit.diagnostics().full_reparse_performed());
        assert_eq!(
            slide_reference_records(&exact_bytes(commit.package()), BUILDS_FIELD)?,
            [before_builds[1].clone(), before_builds[0].clone()]
        );
        assert_eq!(
            slide_reference_records(&exact_bytes(commit.package()), BUILD_CHUNKS_FIELD)?,
            [before_chunks[1].clone(), before_chunks[0].clone()]
        );
        assert_eq!(exact_bytes(&package), bytes);
        Ok(())
    }

    #[test]
    fn noop_inverse_and_exact_source_conflict_are_typed() -> TestResult<()> {
        let bytes = package_bytes(FixtureOptions::VALID)?;
        let package = Package::from_bytes(&bytes)?;

        let mut noop = package.edit_slide_build_order();
        noop.move_build(0usize, Position::new(1), Position::new(1))?;
        let noop_commit = noop.commit()?;
        assert!(noop_commit.patch().is_noop());
        assert!(!noop_commit.diagnostics().changed());
        assert_eq!(exact_bytes(noop_commit.package()), bytes);

        let mut equal_distinct = noop_commit.patch().clone();
        equal_distinct.target_bytes = Arc::from(bytes.clone());
        assert!(equal_distinct.is_noop());
        let mut forged = noop_commit.patch().clone();
        let mut different_target = bytes.clone();
        different_target.push(0);
        forged.target_fingerprint = fingerprint(&different_target);
        forged.target_bytes = Arc::from(different_target);
        assert!(!forged.is_noop());

        let mut edit = package.edit_slide_build_order();
        edit.move_build(0usize, Position::new(0), Position::new(1))?;
        let commit = edit.commit()?;
        let applied = package.apply_slide_build_order(commit.patch())?;
        assert_eq!(
            exact_bytes(applied.package()),
            exact_bytes(commit.package())
        );
        assert_eq!(commit.patch().inverse().inverse(), commit.patch().clone());
        let restored = commit
            .package()
            .apply_slide_build_order(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package()), bytes);

        let changed_source = applied.package();
        assert!(matches!(
            changed_source.apply_slide_build_order(commit.patch()),
            Err(SlideBuildOrderError::PatchConflict)
        ));
        let debug = format!("{:?}", commit.patch());
        assert!(!debug.contains("Index/"));
        assert!(!debug.contains("Build fixture"));
        assert!(!debug.contains("bytes"));
        Ok(())
    }

    #[test]
    fn strict_admission_rejects_invalid_group_start_and_legacy_cache_without_mutation()
    -> TestResult<()> {
        for options in [
            FixtureOptions {
                interleaved_chunks: true,
                ..FixtureOptions::VALID
            },
            FixtureOptions {
                first_chunk_with_previous: true,
                ..FixtureOptions::VALID
            },
            FixtureOptions {
                deprecated_chunk_list: true,
                ..FixtureOptions::VALID
            },
        ] {
            let bytes = package_bytes(options)?;
            let package = Package::from_bytes(&bytes)?;
            let source = exact_bytes(&package);
            let result = (|| -> Result<(), SlideBuildOrderError> {
                let mut edit = package.edit_slide_build_order();
                edit.move_build(0usize, Position::new(0), Position::new(1))?;
                let _commit = edit.commit()?;
                Ok(())
            })();
            assert!(matches!(
                result,
                Err(SlideBuildOrderError::UnsupportedTopology)
            ));
            assert_eq!(exact_bytes(&package), source);
        }
        Ok(())
    }

    #[test]
    fn order_list_limits_classify_build_chunk_and_aggregate_reference_pressure() {
        let build = length_delimited_field(BUILDS_FIELD, &reference_payload(20, None));
        let chunk = length_delimited_field(BUILD_CHUNKS_FIELD, &reference_payload(30, None));

        let mut two_builds = build.clone();
        two_builds.extend_from_slice(&build);
        assert!(matches!(
            parse_slide_order_payload(&two_builds, WireLimits::default(), 1),
            Err(SlideBuildOrderError::LimitExceeded {
                kind: SlideBuildOrderLimitKind::Builds,
                observed: 2,
                maximum: 1,
            })
        ));

        let mut two_chunks = chunk.clone();
        two_chunks.extend_from_slice(&chunk);
        assert!(matches!(
            parse_slide_order_payload(&two_chunks, WireLimits::default(), 1),
            Err(SlideBuildOrderError::LimitExceeded {
                kind: SlideBuildOrderLimitKind::BuildChunks,
                observed: 2,
                maximum: 1,
            })
        ));

        let mut aggregate = build;
        aggregate.extend_from_slice(&chunk);
        assert!(matches!(
            parse_slide_order_payload(&aggregate, WireLimits::default(), 1),
            Err(SlideBuildOrderError::LimitExceeded {
                kind: SlideBuildOrderLimitKind::References,
                observed: 2,
                maximum: 1,
            })
        ));
    }
}
